/**
 * Sections.
 *
 * A Word document is a sequence of sections, each with its own page size,
 * margins and headers. Reading only the last one lays the whole document out
 * with the geometry of its final few paragraphs - which on real resumes is
 * the difference between a half-inch margin and an inch.
 *
 *   npm run test:section
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
class Q extends DOMParser { constructor() { super({ onError: () => {} }); } }
globalThis.DOMParser = Q; globalThis.XMLSerializer = XMLSerializer;
const { importDocx, exportDocx, unzip, sectionsOf, sectionIndexByBlock } =
  await import('./build/harness.mjs');
const HERE = dirname(fileURLToPath(import.meta.url));
const dec = new TextDecoder();

let failures = 0;
function check(name, ok, detail = '') {
  if (ok) { console.log('  ok    ' + name); return; }
  failures++; console.log('  FAIL  ' + name + (detail ? ' - ' + detail : ''));
}

async function open(name) {
  const bytes = await readFile(join(HERE, 'corpus', name));
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  return { buf, ...(await importDocx(buf, name)) };
}

/* ------------------------------------------------------------------ */
console.log('\ntwo sections are both read');
const { buf, doc, vault } = await open('two-sections.docx');
{
  check('the document has two sections', doc.sections?.length === 2, String(doc.sections?.length));
  const [a, b] = doc.sections ?? [];
  // 540 twips is 0.375in is 36px; 1440 twips is one inch is 96px.
  check('the first section keeps its own margins', a?.page.margins.left === 36, String(a?.page.margins.left));
  check('the second section keeps its own', b?.page.margins.left === 96, String(b?.page.margins.left));
  check('the top margins differ too', a?.page.margins.top !== b?.page.margins.top);
  check(
    'doc.page is the FIRST section, not the last',
    doc.page.margins.left === 36,
    String(doc.page.margins.left)
  );
  check('the first section starts at the top', a?.startId === null);
  check('the second starts at a real block', typeof b?.startId === 'string' && b.startId !== '');
  check(
    'that block exists',
    doc.blocks.some((x) => x.id === b?.startId),
    String(b?.startId)
  );
}

/* ------------------------------------------------------------------ */
console.log('\nevery block lands in the right section');
{
  const byBlock = sectionIndexByBlock(doc);
  const counts = [0, 0];
  for (const b of doc.blocks) counts[byBlock.get(b.id) ?? 0]++;
  check('no block is unassigned', byBlock.size === doc.blocks.length);
  check('both sections have content', counts[0] > 0 && counts[1] > 0, counts.join('/'));
  const startAt = doc.blocks.findIndex((b) => b.id === doc.sections[1].startId);
  check(
    'the split is exactly at the section start',
    doc.blocks.every((b, i) => (byBlock.get(b.id) ?? 0) === (i < startAt ? 0 : 1)),
    'start index ' + startAt
  );
}

/* ------------------------------------------------------------------ */
console.log('\neach section gets its own header');
{
  const [a, b] = doc.sections;
  const text = (s) => (s?.headers?.default ?? []).map((p) => p.html.replace(/<[^>]*>/g, '')).join('');
  check('the first section has a header', text(a).includes('SECTION ONE'), text(a));
  check('the second has a different one', text(b).includes('SECTION TWO'), text(b));
  check('their block ids do not collide', text(a) !== text(b) &&
    (a.headers.default[0].id !== b.headers.default[0].id));
}

/* ------------------------------------------------------------------ */
console.log('\na single-section document is unchanged');
{
  const one = await open('resume.docx');
  check('no sections array is written', one.doc.sections === undefined);
  check('sectionsOf still gives one', sectionsOf(one.doc).length === 1);
  check(
    'its page setup is the document page setup',
    sectionsOf(one.doc)[0].page.margins.left === one.doc.page.margins.left
  );
}

/* ------------------------------------------------------------------ */
console.log('\nboth section breaks survive export');
{
  const original = await unzip(buf);
  const out = await unzip(await (await exportDocx(doc, vault)).arrayBuffer());
  const before = dec.decode(original.get('word/document.xml'));
  const after = dec.decode(out.get('word/document.xml'));
  check('document.xml is byte-identical', before === after);
  check(
    'both sectPr are still there',
    (after.match(/<w:sectPr[ >]/g) || []).length === 2,
    String((after.match(/<w:sectPr[ >]/g) || []).length)
  );

  // And after an edit to a paragraph in the first section.
  const edited = structuredClone(doc);
  edited.blocks[1].html = 'Rewritten opening line.';
  const out2 = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml2 = dec.decode(out2.get('word/document.xml'));
  check('the edit is written', xml2.includes('Rewritten opening line'));
  check(
    'both sectPr survive the edit',
    (xml2.match(/<w:sectPr[ >]/g) || []).length === 2,
    String((xml2.match(/<w:sectPr[ >]/g) || []).length)
  );
  check(
    'the inline sectPr is still inside a pPr',
    /<w:pPr>[\s\S]*?<w:sectPr[ >]/.test(xml2)
  );
  const others = [...original.keys()].filter(
    (k) => k !== 'word/document.xml' &&
      dec.decode(original.get(k)) !== dec.decode(out2.get(k) ?? new Uint8Array())
  );
  check('no other part changed', others.length === 0, others.join(', '));
}

/* ------------------------------------------------------------------ */
console.log("\nediting the second section's header writes back to its own part");
{
  const original = await unzip(buf);
  const edited = structuredClone(doc);
  edited.sections[1].headers.default[0].html = 'REVISED SECOND HEADER';
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());

  const changed = [...original.keys()].filter(
    (k) => dec.decode(original.get(k)) !== dec.decode(out.get(k) ?? new Uint8Array())
  );
  check('exactly one part changed', changed.length === 1, changed.join(', '));
  check('and it is a header part', /^word\/header\d+\.xml$/.test(changed[0] ?? ''), changed[0]);
  const part = dec.decode(out.get(changed[0]));
  check('the new text is in it', part.includes('REVISED SECOND HEADER'));
  check('the old text is gone', !part.includes('SECTION TWO'));
  check(
    "the first section's header is untouched",
    [...original.keys()]
      .filter((k) => /header\d+\.xml$/.test(k) && k !== changed[0])
      .every((k) => dec.decode(original.get(k)) === dec.decode(out.get(k)))
  );
}

console.log(failures === 0 ? '\nall section checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
