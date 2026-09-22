/**
 * Making things, rather than only reading them.
 *
 * The guarantee under pressure here is the one the whole vault exists for:
 * a document can gain a link, a picture and a page break, and every part of
 * the package it did not need to touch still comes back byte for byte. Only
 * document.xml and the relationships part may differ, and the relationships
 * part only by addition.
 *
 *   npm run test:insert
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
class Q extends DOMParser { constructor() { super({ onError: () => {} }); } }
globalThis.DOMParser = Q; globalThis.XMLSerializer = XMLSerializer;
const { importDocx, exportDocx, unzip, addMedia } = await import('./build/harness.mjs');
const HERE = dirname(fileURLToPath(import.meta.url));
const dec = new TextDecoder();

let failures = 0;
function check(name, ok, detail = '') {
  if (ok) { console.log('  ok    ' + name); return; }
  failures++; console.log('  FAIL  ' + name + (detail ? ' - ' + detail : ''));
}

const bytes = await readFile(join(HERE, 'corpus', 'resume.docx'));
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const original = await unzip(buf);
const RELS = 'word/_rels/document.xml.rels';

async function exported(mutate) {
  const { doc, vault } = await importDocx(buf, 'resume.docx');
  mutate(doc, vault);
  const out = await unzip(await (await exportDocx(doc, vault)).arrayBuffer());
  return {
    out,
    xml: dec.decode(out.get('word/document.xml')),
    rels: dec.decode(out.get(RELS) ?? new Uint8Array()),
    changed: [...original.keys()].filter(
      (k) => dec.decode(original.get(k)) !== dec.decode(out.get(k) ?? new Uint8Array())
    ),
  };
}

/* ------------------------------------------------------------------ */
console.log('\na link added in the editor becomes a real hyperlink');
{
  const { xml, rels, changed } = await exported((doc) => {
    doc.blocks[1].html = 'See <a href="https://example.com/cv">the notes</a>.';
  });
  check('a w:hyperlink is written', /<w:hyperlink r:id="rId\d+">/.test(xml), '');
  check('not plain text', xml.includes('the notes'));

  const before = dec.decode(original.get(RELS));
  const added = /Target="https:\/\/example\.com\/cv"[^>]*TargetMode="External"/.test(rels);
  check('a relationship is appended', added, '');
  check('with an external target mode', rels.includes('TargetMode="External"'));
  check(
    'every relationship it already had is still there',
    [...before.matchAll(/Id="(rId\d+)"/g)].every((m) => rels.includes(`Id="${m[1]}"`))
  );
  const newId = /Id="(rId\d+)"[^>]*example\.com/.exec(rels)?.[1];
  check('the new id did not collide', !!newId && !before.includes(`Id="${newId}"`), newId);
  check(
    'only document.xml and the rels changed',
    changed.length === 2 && changed.includes(RELS),
    changed.join(', ')
  );
}

/* ------------------------------------------------------------------ */
console.log('\nan inserted image gets a part, a relationship and a drawing');
{
  // A one-pixel PNG is enough: what is under test is the plumbing.
  const png = Uint8Array.from(
    atob(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
    ),
    (c) => c.charCodeAt(0)
  );
  const path = 'word/media/added-test.png';

  const { out, xml, rels, changed } = await exported((doc) => {
    // After the import, not before: opening a document clears the media
    // registry, because the pictures in it belong to the package it came
    // from. Adding one first would be adding it to the previous document.
    addMedia(path, png, 'image/png');
    doc.blocks[1].html += `<img data-media="${path}" width="96" height="96">`;
  });

  check('the media part is in the package', out.has(path), [...out.keys()].join(' ').slice(0, 80));
  check('with the bytes we gave it', (out.get(path)?.length ?? 0) === png.length);
  check('a drawing is written', xml.includes('<w:drawing>'));
  check('with an inline extent', /<wp:extent cx="\d+" cy="\d+"\/>/.test(xml));
  // 96px at 96dpi is one inch, which is 914400 EMU.
  check('sized in EMU', xml.includes('cx="914400"'), /cx="(\d+)"/.exec(xml)?.[1]);
  check('the blip points at a relationship', /<a:blip r:embed="rId\d+"\/>/.test(xml));
  check('the namespaces are declared on the elements that use them',
    xml.includes('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"') &&
    xml.includes('xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'));

  const rel = /Id="(rId\d+)" Type="[^"]*\/image" Target="media\/added-test\.png"\/>/.exec(rels);
  check('the image relationship is relative to word/', !!rel, rels.slice(-220));
  check(
    'only document.xml and the rels changed, plus the new part',
    changed.length === 2 && changed.includes(RELS),
    changed.join(', ')
  );
}

/* ------------------------------------------------------------------ */
console.log('\na page break on a paragraph is written to its properties');
{
  const { xml, changed } = await exported((doc) => {
    doc.blocks[2].fmt = { ...(doc.blocks[2].fmt ?? {}), pageBreakBefore: true };
  });
  check('w:pageBreakBefore is written', xml.includes('<w:pageBreakBefore/>'));
  check('exactly once', (xml.match(/<w:pageBreakBefore\/>/g) || []).length === 1);
  check(
    'it comes before the style-level properties it must',
    /<w:pPr>[\s\S]{0,400}?<w:pageBreakBefore\/>[\s\S]*?<\/w:pPr>/.test(xml)
  );
  check('only document.xml changed', changed.length === 1, changed.join(', '));
}

/* ------------------------------------------------------------------ */
console.log('\nan untouched document is still untouched');
{
  const { changed } = await exported(() => {});
  check('nothing at all changed', changed.length === 0, changed.join(', '));
}

console.log(failures === 0 ? '\nall insert checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
