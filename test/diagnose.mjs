/**
 * Structural diff for one document, to find where a round trip diverges.
 *
 * Prints element structure only: every text node is redacted to a run of
 * dots of the same length, so a failing document can be diagnosed without
 * its contents being shown or logged.
 *
 *   node test/diagnose.mjs "/path/to/file.docx"
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFile } from 'node:fs/promises';
import { basename } from 'node:path';

class QuietDOMParser extends DOMParser {
  constructor() {
    super({ onError: () => {} });
  }
}
globalThis.DOMParser = QuietDOMParser;
globalThis.XMLSerializer = XMLSerializer;

const { importDocx, exportDocx, unzip } = await import('./build/harness.mjs');

/** Replace text content with dots, keeping every tag and length intact. */
function redact(xml) {
  return xml.replace(/>([^<]+)</g, (_m, t) => '>' + '.'.repeat(t.length) + '<');
}

const path = process.argv[2];
if (!path) {
  console.error('usage: node test/diagnose.mjs <file.docx>');
  process.exit(1);
}

const bytes = await readFile(path);
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const dec = new TextDecoder();

const original = await unzip(buf);
const origDoc = dec.decode(original.get('word/document.xml'));

const { doc, vault } = await importDocx(buf, basename(path));
const exported = await unzip(await (await exportDocx(doc, vault)).arrayBuffer());
const outDoc = dec.decode(exported.get('word/document.xml'));

console.log('file           ', basename(path));
console.log('blocks         ', doc.blocks.length);
console.log('opaque entries ', vault.opaque.length);
const byKind = {};
for (const o of vault.opaque) byKind[o.kind] = (byKind[o.kind] ?? 0) + 1;
console.log('opaque by kind ', JSON.stringify(byKind));
console.log('doc.xml length ', origDoc.length, '->', outDoc.length);

const tags = (s) => {
  const m = {};
  for (const t of s.match(/<w:[a-zA-Z]+[ />]/g) || []) {
    const name = t.slice(0, -1);
    m[name] = (m[name] ?? 0) + 1;
  }
  return m;
};
const a = tags(origDoc);
const b = tags(outDoc);
const diffs = Object.keys({ ...a, ...b })
  .filter((k) => (a[k] ?? 0) !== (b[k] ?? 0))
  .map((k) => `${k}: ${a[k] ?? 0} -> ${b[k] ?? 0}`);
console.log('tag count diffs', diffs.length ? diffs.join(', ') : 'none');

const textOf = (xml) =>
  (xml.match(/<w:t[^>]*>[^<]*<\/w:t>/g) || [])
    .map((t) => t.replace(/<[^>]*>/g, ''))
    .join('');
const ta = textOf(origDoc);
const tb = textOf(outDoc);
console.log('text length    ', ta.length, '->', tb.length);
if (ta !== tb) {
  let i = 0;
  while (i < ta.length && i < tb.length && ta[i] === tb[i]) i++;
  console.log('text diverges at char', i, 'of', ta.length);
}

if (origDoc !== outDoc) {
  let i = 0;
  while (i < origDoc.length && i < outDoc.length && origDoc[i] === outDoc[i]) i++;
  console.log('\nxml diverges at char', i);
  console.log('  original: ' + redact(origDoc.slice(Math.max(0, i - 90), i + 130)));
  console.log('  exported: ' + redact(outDoc.slice(Math.max(0, i - 90), i + 130)));
} else {
  console.log('\nxml identical');
}
