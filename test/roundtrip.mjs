/**
 * Fidelity harness: the automated round trip from the spec.
 *
 * For every .docx in test/corpus/ it asserts that importing and re-exporting
 * without edits gives back the same document, and that editing one paragraph
 * changes only that paragraph. The byte-identity assertion on non-document
 * parts is what enforces the preservation vault: it fails loudly the first
 * time someone regenerates the package from scratch, which is the point.
 *
 *   npm run test:roundtrip
 *
 * Runs in plain Node against the same bundled code the browser uses, with a
 * DOM shim standing in for the browser's parser.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readdir, readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const CORPUS = join(HERE, 'corpus');

// xmldom reports malformed input by calling the error handler rather than by
// inserting a <parsererror> node, so give the code the shape it expects.
class QuietDOMParser extends DOMParser {
  constructor() {
    super({ onError: () => {} });
  }
}
globalThis.DOMParser = QuietDOMParser;
globalThis.XMLSerializer = XMLSerializer;

const { importDocx, exportDocx, unzip } = await import('./build/harness.mjs');

const dec = new TextDecoder();
let failures = 0;
let checks = 0;

function check(file, name, ok, detail = '') {
  checks++;
  if (!ok) {
    failures++;
    console.log(`  FAIL  ${name}${detail ? ' - ' + detail : ''}`);
  }
  return ok;
}

function sameBytes(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function textOf(xml) {
  return (xml.match(/<w:t[^>]*>[^<]*<\/w:t>/g) || [])
    .map((t) => t.replace(/<[^>]*>/g, ''))
    .join('');
}

const countOf = (xml, re) => (xml.match(re) || []).length;

async function zipOf(blob) {
  return unzip(await blob.arrayBuffer());
}

async function run(file) {
  const bytes = await readFile(join(CORPUS, file));
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const original = await unzip(buf);
  const origDoc = dec.decode(original.get('word/document.xml'));

  const { doc, vault } = await importDocx(buf, file);
  const exported = await zipOf(await exportDocx(doc, vault));
  const outDoc = dec.decode(exported.get('word/document.xml'));

  console.log(`\n${file}  (${original.size} parts, ${doc.blocks.length} blocks)`);

  // every part present
  const missing = [...original.keys()].filter((k) => !exported.has(k));
  check(file, 'every original part is present', missing.length === 0, missing.join(', '));
  const added = [...exported.keys()].filter((k) => !original.has(k));
  check(file, 'no parts invented', added.length === 0, added.join(', '));

  // non-document parts byte-identical -- this is what enforces the vault
  const differing = [...original.keys()].filter(
    (k) => k !== 'word/document.xml' && !sameBytes(original.get(k), exported.get(k))
  );
  check(
    file,
    'non-document parts are byte-identical',
    differing.length === 0,
    differing.join(', ')
  );

  // text content
  check(file, 'text content matches', textOf(origDoc) === textOf(outDoc));

  // structural counts
  for (const [label, re] of [
    ['paragraphs', /<w:p[ >]/g],
    ['tables', /<w:tbl[ >]/g],
    ['numbered paragraphs', /<w:numPr[ >]/g],
    ['hyperlinks', /<w:hyperlink[ >]/g],
    ['content controls', /<w:sdt[ >]/g],
    ['insertions', /<w:ins[ >]/g],
    ['deletions', /<w:del[ >]/g],
  ]) {
    const a = countOf(origDoc, re);
    const b = countOf(outDoc, re);
    check(file, `${label} count matches`, a === b, `${a} -> ${b}`);
  }

  // an untouched round trip should not change the document at all
  check(file, 'untouched document.xml is byte-identical', origDoc === outDoc);

  // editing one paragraph must not disturb anything else
  const target = doc.blocks.findIndex((b) => b.html.length > 10);
  if (target >= 0) {
    const edited = {
      ...doc,
      blocks: doc.blocks.map((b, i) =>
        i === target ? { ...b, html: b.html + ' EDITED' } : b
      ),
    };
    const editedZip = await zipOf(await exportDocx(edited, vault));
    const editedDoc = dec.decode(editedZip.get('word/document.xml'));
    const editedDiffering = [...original.keys()].filter(
      (k) => k !== 'word/document.xml' && !sameBytes(original.get(k), editedZip.get(k))
    );
    check(file, 'an edit leaves other parts untouched', editedDiffering.length === 0);
    check(file, 'the edit is present', editedDoc.includes('EDITED'));

    let pre = 0;
    while (pre < outDoc.length && outDoc[pre] === editedDoc[pre]) pre++;
    let suf = 0;
    while (
      suf < outDoc.length - pre &&
      outDoc[outDoc.length - 1 - suf] === editedDoc[editedDoc.length - 1 - suf]
    ) {
      suf++;
    }
    const changed = editedDoc.length - pre - suf;
    check(
      file,
      'an edit rewrites only its own paragraph',
      changed < 600,
      `${changed} chars changed`
    );
    check(
      file,
      'paragraph count unchanged by an edit',
      countOf(outDoc, /<w:p[ >]/g) === countOf(editedDoc, /<w:p[ >]/g)
    );
  }

  if (vault.warnings.length) {
    console.log(
      '  preserved but not shown: ' +
        vault.warnings.map((w) => `${w.kind} x${w.count}`).join(', ')
    );
  }
}

const all = (await readdir(CORPUS)).filter((f) => f.toLowerCase().endsWith('.docx'));
if (all.length === 0) {
  console.error('No .docx files in test/corpus. Run: node test/make-corpus.mjs');
  process.exit(1);
}
console.log(`Round-tripping ${all.length} documents`);
for (const f of all) {
  try {
    await run(f);
  } catch (err) {
    failures++;
    console.log(`\n${f}\n  FAIL  threw: ${err && err.message}`);
  }
}

console.log(
  `\n${checks - failures}/${checks} checks passed across ${all.length} documents`
);
process.exit(failures === 0 ? 0 : 1);
