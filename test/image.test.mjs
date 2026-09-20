/**
 * Images have to survive being edited around.
 *
 * A w:drawing carries cropping, effects and positioning that cannot be
 * rebuilt from an <img>, so the risk is a paragraph edit quietly degrading
 * or dropping the picture in it.
 *
 *   npm run test:image
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
const check = (name, ok, detail = '') => {
  if (ok) { console.log('  ok    ' + name); return; }
  failures++; console.log('  FAIL  ' + name + (detail ? ' - ' + detail : ''));
};
const count = (s, re) => (s.match(re) || []).length;

const bytes = await readFile(join(HERE, 'corpus', 'images.docx'));
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const original = await unzip(buf);
const origDoc = dec.decode(original.get('word/document.xml'));
const { doc, vault } = await importDocx(buf, 'images.docx');

console.log('\nimport');
const html = doc.blocks.filter((b) => b.kind !== 'table').map((b) => b.html).join('');
check('every drawing became an image', count(html, /<img /g) === count(origDoc, /<w:drawing[ >]/g),
  `${count(html, /<img /g)} of ${count(origDoc, /<w:drawing[ >]/g)}`);
check('images name a media part', /data-media="word\/media\//.test(html));
check('images name a preserved run', /data-run="img\d+"/.test(html));
check('media parts are present', [...original.keys()].some((k) => k.startsWith('word/media/')));
check('each image kept its own run', vault.runXml.size === count(origDoc, /<w:drawing[ >]/g));
check('an image sits inline with text',
  doc.blocks.some((b) => b.kind !== 'table' && /\w<img |<img [^>]*>\w/.test(b.html) ||
    (b.kind !== 'table' && b.html.includes('<img ') && b.html.replace(/<[^>]*>/g, '').trim().length > 0)));

console.log('\nediting the paragraph an image sits in');
{
  const target = doc.blocks.findIndex((b) => b.kind !== 'table' && b.html.includes('<img '));
  const edited = {
    ...doc,
    blocks: doc.blocks.map((b, i) =>
      i === target ? { ...b, html: b.html + ' EDITED' } : b
    ),
  };
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));

  check('the edit is present', xml.includes('EDITED'));
  check('no drawing was lost', count(xml, /<w:drawing[ >]/g) === count(origDoc, /<w:drawing[ >]/g),
    `${count(xml, /<w:drawing[ >]/g)} of ${count(origDoc, /<w:drawing[ >]/g)}`);
  check('the relationship id survived', count(xml, /r:embed=/g) === count(origDoc, /r:embed=/g));
  check('the extent survived', count(xml, /<wp:extent/g) === count(origDoc, /<wp:extent/g));
  const sameBytes = (a, b) => a && b && a.length === b.length && a.every((v, i) => v === b[i]);
  const media = [...original.keys()].filter((k) => k.startsWith('word/media/'));
  check('media bytes untouched', media.every((k) => sameBytes(original.get(k), out.get(k))),
    media.join(', '));
  check('no image warning was raised',
    !vault.warnings.some((w) => w.kind === 'images'),
    vault.warnings.map((w) => w.kind).join(', '));
}

console.log(failures === 0 ? '\nall image checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
