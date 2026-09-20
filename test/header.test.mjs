/**
 * Headers, footers and the fields in them.
 *
 *   npm run test:header
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

const bytes = await readFile(join(HERE, 'corpus', 'letterhead.docx'));
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const original = await unzip(buf);
const { doc, vault } = await importDocx(buf, 'letterhead.docx');

console.log('\nimport');
check('a title page is recognized', doc.titlePage === true);
check('the first-page header is separate', !!doc.headers?.first);
check('a default header exists', !!doc.headers?.default);
check('a footer exists', !!doc.footers?.default);
check('header parts were located', vault.hfParts.size >= 2);
check('header shells were captured', vault.hfShell.size >= 2);

const footerHtml = (doc.footers?.default ?? []).map((b) => b.html).join('');
console.log('\nfields');
check('PAGE is a field, not a number', footerHtml.includes('data-field="PAGE"'));
check('NUMPAGES is a field', footerHtml.includes('data-field="NUMPAGES"'));
check('no page number was baked into the text',
  !/Page\s*\d+\s*of\s*\d+/.test(footerHtml.replace(/<[^>]*>/g, '')));

console.log('\nuntouched export');
{
  const out = await unzip(await (await exportDocx(doc, vault)).arrayBuffer());
  const same = (k) => {
    const a = original.get(k), b = out.get(k);
    return a && b && a.length === b.length && a.every((v, i) => v === b[i]);
  };
  const parts = [...original.keys()].filter((k) => /header\d*\.xml|footer\d*\.xml/.test(k));
  check('header and footer parts are byte-identical', parts.length > 0 && parts.every(same),
    parts.join(', '));
  check('the body is byte-identical',
    dec.decode(original.get('word/document.xml')) === dec.decode(out.get('word/document.xml')));
}

console.log('\nediting the header');
{
  const edited = structuredClone(doc);
  edited.headers.first[0].html = 'NORTHWIND LIMITED &#8212; REVISED';
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const hits = [...out.keys()].filter((k) => /header\d*\.xml$/.test(k))
    .map((k) => dec.decode(out.get(k))).join('');
  check('the edit reached a header part', hits.includes('REVISED'));
  check('the header part is still a header', count(hits, /<w:hdr[ >]/g) >= 1);
  // The footer was not touched, so it must not have been rewritten.
  const sameFooter = [...original.keys()].filter((k) => /footer\d*\.xml$/.test(k))
    .every((k) => {
      const a = original.get(k), b = out.get(k);
      return a && b && a.length === b.length && a.every((v, i) => v === b[i]);
    });
  check('an untouched footer is left alone', sameFooter);
  const outFooters = [...out.keys()].filter((k) => /footer\d*\.xml$/.test(k))
    .map((k) => dec.decode(out.get(k))).join('');
  check('the page field survived as a field', /PAGE/.test(outFooters));
}

console.log(failures === 0 ? '\nall header checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
