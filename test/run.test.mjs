/**
 * Character formatting.
 *
 * The bug this exists to pin down: the vault kept ONE base w:rPr per
 * paragraph and wrote it onto every regenerated run, so editing a line with
 * a red word in it made the whole line red. Untouched paragraphs were always
 * fine, which is exactly why the round-trip harness never saw it.
 *
 *   npm run test:run
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
class Q extends DOMParser { constructor() { super({ onError: () => {} }); } }
globalThis.DOMParser = Q; globalThis.XMLSerializer = XMLSerializer;
const { importDocx, exportDocx, unzip } = await import('./build/harness.mjs');
const HERE = dirname(fileURLToPath(import.meta.url));
const dec = new TextDecoder();

let failures = 0;
function check(name, ok, detail = '') {
  if (ok) { console.log('  ok    ' + name); return; }
  failures++; console.log('  FAIL  ' + name + (detail ? ' - ' + detail : ''));
}

const bytes = await readFile(join(HERE, 'corpus', 'run-format.docx'));
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const original = await unzip(buf);
const { doc, vault } = await importDocx(buf, 'run-format.docx');

/* ------------------------------------------------------------------ */
console.log('\nmixed runs are read');
{
  const html = doc.blocks[0].html;
  check('a colour becomes a span', /data-color="C00000"/.test(html), html.slice(0, 200));
  check('a size becomes a span', /data-sz="18"/.test(html), html.slice(0, 260));
  check('a font becomes a span', /data-font="Courier New"/.test(html));
  check('strikethrough is read', /data-strike="1"/.test(html));
  check('superscript is read', /data-vert="super"/.test(html));
  check('subscript is read', /data-vert="sub"/.test(html));
  check(
    'the plain text between them carries no span',
    html.includes('Plain, then <span') && html.includes('</span>, then <span'),
    html.slice(0, 160)
  );
}

/* ------------------------------------------------------------------ */
console.log('\na uniform paragraph becomes one span, not one per run');
{
  // Three 14pt runs in the file. Adjacent pieces that agree are merged
  // before the markup is written, so this is one span rather than three.
  const html = doc.blocks[1].html;
  const spans = (html.match(/<span/g) || []).length;
  check('exactly one span', spans === 1, String(spans));
  check('at fourteen points', html.includes('data-sz="14"'), html.slice(0, 120));
}

/* ------------------------------------------------------------------ */
console.log('\nan untouched document still exports byte-identically');
{
  const out = await unzip(await (await exportDocx(doc, vault)).arrayBuffer());
  check(
    'document.xml unchanged',
    dec.decode(original.get('word/document.xml')) === dec.decode(out.get('word/document.xml'))
  );
}

/* ------------------------------------------------------------------ */
console.log('\nediting a mixed paragraph no longer flattens it');
{
  const edited = structuredClone(doc);
  // Change one word, the way typing does, and leave the formatting alone.
  edited.blocks[0].html = edited.blocks[0].html.replace('Plain, then ', 'Edited, then ');
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));
  const first = /<w:p>[\s\S]*?<\/w:p>/.exec(xml)?.[0] ?? '';

  check('the edit is in the file', first.includes('Edited, then'));
  check('the red word is still red', first.includes('<w:color w:val="C00000"/>'), '');
  check('the large word is still large', first.includes('<w:sz w:val="36"/>'));
  check('and carries szCs too', first.includes('<w:szCs w:val="36"/>'));
  check('the courier word keeps its font', first.includes('w:ascii="Courier New"'));
  check('strike survives', first.includes('<w:strike/>'));
  check('superscript survives', first.includes('<w:vertAlign w:val="superscript"/>'));
  check('subscript survives', first.includes('<w:vertAlign w:val="subscript"/>'));

  // The flattening bug: every run taking the FIRST run's properties.
  const colours = (first.match(/<w:color w:val="[0-9A-Fa-f]{6}"\/>/g) || []).length;
  check('the colour is on one run, not all of them', colours === 1, String(colours));

  const others = [...original.keys()].filter(
    (k) => k !== 'word/document.xml' &&
      dec.decode(original.get(k)) !== dec.decode(out.get(k) ?? new Uint8Array())
  );
  check('no other part changed', others.length === 0, others.join(', '));
}

/* ------------------------------------------------------------------ */
console.log('\nbold and colour on the same run both survive');
{
  const edited = structuredClone(doc);
  edited.blocks[2].html = edited.blocks[2].html.replace('beside plain', 'beside ordinary');
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));
  const runs = xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
  const both = runs.filter((r) => r.includes('<w:b/>') && r.includes('w:val="1F4E79"'));
  check('one run has both', both.length === 1, String(both.length));
  check(
    'rPr children are in schema order',
    both[0] ? both[0].indexOf('<w:b/>') < both[0].indexOf('<w:color') : false,
    both[0]?.slice(0, 160)
  );
}

/* ------------------------------------------------------------------ */
console.log('\nsetting a size writes it over the paragraph base');
{
  const edited = structuredClone(doc);
  // The uniform 14pt paragraph: make three words 24pt.
  edited.blocks[1].html = edited.blocks[1].html.replace(
    'A whole paragraph',
    '<span data-sz="24">A whole paragraph</span>'
  );
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));
  check('the override is written', xml.includes('<w:sz w:val="48"/>'), '');
  check('and its szCs matches', xml.includes('<w:szCs w:val="48"/>'));
  const runs = xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
  const doubled = runs.filter(
    (r) => (r.match(/<w:sz w:val=/g) || []).length > 1
  );
  check('no run has two sizes', doubled.length === 0, String(doubled.length));
  check('the rest of the paragraph keeps 14pt', xml.includes('<w:sz w:val="28"/>'));
}

console.log(failures === 0 ? '\nall character-format checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
