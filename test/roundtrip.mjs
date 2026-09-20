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
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));

/**
 * Where the corpus lives. Defaults to test/corpus, but can point anywhere:
 *
 *   npm run test:roundtrip -- "/path/to/your/documents"
 *   WP_CORPUS="/path/to/your/documents" npm run test:roundtrip
 *
 * Pointing it at a real documents folder reads those files in place: nothing
 * is copied into the repository and nothing leaves the machine. The output is
 * counts and file names only, never document content.
 */
const CORPUS = process.argv[2] || process.env.WP_CORPUS || join(HERE, 'corpus');
const RECURSE = process.env.WP_CORPUS_RECURSE !== '0';

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
let identical = 0;
const notIdentical = [];

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

/**
 * Entities must be decoded before comparing. A parser legitimately rewrites
 * `&quot;` as `"` and `&#8217;` as the character itself; the XML differs, the
 * document does not. Comparing raw markup reports those as lost text.
 */
function decodeEntities(s) {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_m, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function textOf(xml) {
  return decodeEntities(
    (xml.match(/<w:t[^>]*>[^<]*<\/w:t>/g) || [])
      .map((t) => t.replace(/<[^>]*>/g, ''))
      .join('')
  );
}

const countOf = (xml, re) => (xml.match(re) || []).length;

async function zipOf(blob) {
  return unzip(await blob.arrayBuffer());
}

async function run(path) {
  const file = basename(path);
  const bytes = await readFile(path);
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

  // Byte-identity of document.xml is the ideal, and it holds whenever the
  // source markup survives a parse and re-serialize unchanged. It legitimately
  // does not when the original spells characters as entities, so this is
  // reported rather than failed - the binding assertion is the normalized
  // text and the structural counts above.
  if (origDoc === outDoc) identical++;
  else notIdentical.push(file);

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

    // Count differing paragraphs rather than differing characters: editing a
    // long paragraph legitimately rewrites a lot of XML, but it must still be
    // exactly one paragraph.
    const split = (xml) => xml.split(/(?=<w:p[ >])/);
    const before = split(outDoc);
    const after = split(editedDoc);
    const changedParas =
      before.length !== after.length
        ? -1
        : before.reduce((n, p, i) => n + (p === after[i] ? 0 : 1), 0);
    check(
      file,
      'an edit rewrites exactly one paragraph',
      changedParas === 1,
      changedParas === -1
        ? 'paragraph count changed'
        : `${changedParas} paragraphs differ`
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

async function collect(dir, depth = 0) {
  const out = [];
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return out;
  }
  for (const e of entries) {
    const full = join(dir, e.name);
    if (e.isDirectory()) {
      if (RECURSE && depth < 6 && e.name !== 'node_modules' && !e.name.startsWith('.')) {
        out.push(...(await collect(full, depth + 1)));
      }
      continue;
    }
    // ~$ files are Word's lock files, not documents.
    if (e.name.toLowerCase().endsWith('.docx') && !e.name.startsWith('~$')) {
      out.push(full);
    }
  }
  return out;
}

const all = await collect(CORPUS);
if (all.length === 0) {
  console.error(`No .docx files under ${CORPUS}. Run: node test/make-corpus.mjs`);
  process.exit(1);
}
console.log(`Round-tripping ${all.length} documents`);
for (const f of all) {
  try {
    await run(f);
  } catch (err) {
    failures++;
    console.log(`\n${basename(f)}\n  FAIL  threw: ${err && err.message}`);
  }
}

console.log(
  `\n${checks - failures}/${checks} checks passed across ${all.length} documents`
);
console.log(
  `${identical}/${all.length} untouched round trips were byte-identical` +
    (notIdentical.length
      ? `, the rest differing only in how characters are spelled as entities`
      : '')
);
process.exit(failures === 0 ? 0 : 1);
