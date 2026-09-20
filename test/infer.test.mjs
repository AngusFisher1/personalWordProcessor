/**
 * What the importer should make of a document that has no named styles.
 *
 * resume-direct.docx is built the way real documents are: headings made out
 * of bold, capitals, size and a rule, with no w:pStyle anywhere. If the
 * inference regresses, this is where it shows.
 *
 *   npm run test:infer
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

class QuietDOMParser extends DOMParser {
  constructor() {
    super({ onError: () => {} });
  }
}
globalThis.DOMParser = QuietDOMParser;
globalThis.XMLSerializer = XMLSerializer;

const { importDocx } = await import('./build/harness.mjs');
const HERE = dirname(fileURLToPath(import.meta.url));

let failures = 0;
function expect(name, actual, wanted) {
  const ok = JSON.stringify(actual) === JSON.stringify(wanted);
  if (!ok) {
    failures++;
    console.log(`  FAIL  ${name}`);
    console.log(`        wanted ${JSON.stringify(wanted)}`);
    console.log(`        got    ${JSON.stringify(actual)}`);
  } else {
    console.log(`  ok    ${name}`);
  }
}

async function importFile(name) {
  const bytes = await readFile(join(HERE, 'corpus', name));
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  return importDocx(buf, name);
}

console.log('\ndirectly formatted resume (no styles at all)');
{
  const { doc } = await importFile('resume-direct.docx');
  expect(
    'every paragraph is recognized for what it is',
    doc.blocks.map((b) => b.styleId),
    [
      'Name', // 20pt bold, the largest thing in the document
      'Contact', // small, sits under the title, has an email and a phone
      'SectionHeading', // caps, bold, ruled
      'JobTitle', // partly bold, body size
      'Bullet',
      'Bullet',
      'SectionHeading',
      'Body',
      'SectionHeading',
      'Body',
      'Body', // too long to be a heading, whatever it looks like
    ]
  );
}

console.log('\nnamed styles still win over inference');
{
  const { doc } = await importFile('resume.docx');
  expect(
    'Title / Heading 1 / Heading 3 / list',
    doc.blocks.map((b) => b.styleId),
    [
      'Name',
      'Body',
      'SectionHeading',
      'JobTitle',
      'Bullet',
      'Bullet',
      'Bullet',
      'SectionHeading',
      'Body',
    ]
  );
}

console.log('\na contract is not a resume');
{
  const { doc } = await importFile('contract.docx');
  const styles = doc.blocks.map((b) => b.styleId);
  expect('the title is the title', styles[0], 'Name');
  expect(
    'no contact line is invented under it',
    styles.filter((s) => s === 'Contact').length,
    0
  );
  expect(
    'the clauses stay a list',
    styles.slice(1).every((s) => s === 'Bullet'),
    true
  );
}

console.log('\nprose stays prose');
{
  const { doc } = await importFile('cover-letter.docx');
  const headings = doc.blocks.filter(
    (b) => b.styleId === 'SectionHeading' || b.styleId === 'JobTitle'
  ).length;
  expect('a cover letter has no headings', headings, 0);
}

console.log(failures === 0 ? '\nall inference checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
