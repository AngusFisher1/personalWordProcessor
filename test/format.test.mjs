/**
 * Direct paragraph formatting.
 *
 * Alignment, indents and spacing were preserved on export long before they
 * were drawn, so the risk here is not losing them - it is the opposite: now
 * that they round-trip through a model and back, a paragraph nobody touched
 * has to come out byte-identical anyway, and a paragraph whose alignment
 * changed has to come out as valid OOXML in the order Word demands.
 *
 *   npm run test:format
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

const bytes = await readFile(join(HERE, 'corpus', 'direct-format.docx'));
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const original = await unzip(buf);
const { doc, vault } = await importDocx(buf, 'direct-format.docx');
const at = (i) => doc.blocks[i];

/* ------------------------------------------------------------------ */
console.log('\nwhat the document says is read');
{
  check('a centred paragraph is centred', at(0).fmt?.align === 'center', JSON.stringify(at(0).fmt));
  check('a right-aligned one is right', at(1).fmt?.align === 'right');
  check('a justified one is justified', at(2).fmt?.align === 'justify');
  // 720 twips is half an inch is 36 points.
  check('a left indent reads in points', at(3).fmt?.indentLeft === 36, String(at(3).fmt?.indentLeft));
  check('a right indent reads too', at(3).fmt?.indentRight === 36, String(at(3).fmt?.indentRight));
  check('a hanging indent is negative', at(4).fmt?.firstLine === -18, String(at(4).fmt?.firstLine));
  check('a first-line indent is positive', at(5).fmt?.firstLine === 18, String(at(5).fmt?.firstLine));
  check('space before reads', at(2).fmt?.spaceBefore === 6, String(at(2).fmt?.spaceBefore));
  check('space after reads', at(0).fmt?.spaceAfter === 12, String(at(0).fmt?.spaceAfter));
  check(
    'auto line spacing is a multiplier',
    at(6).fmt?.lineHeight === 1.5 && at(6).fmt?.lineRule === 'auto',
    JSON.stringify(at(6).fmt)
  );
  check(
    'exact line spacing stays in points',
    at(7).fmt?.lineHeight === 14 && at(7).fmt?.lineRule === 'exact',
    JSON.stringify(at(7).fmt)
  );
  check('a plain paragraph has no format at all', at(8).fmt === undefined, JSON.stringify(at(8).fmt));
  check('left is not stored as an override', !('align' in (at(3).fmt ?? {})));
}

/* ------------------------------------------------------------------ */
console.log('\nan untouched document still exports byte-identically');
{
  const out = await unzip(await (await exportDocx(doc, vault)).arrayBuffer());
  check(
    'document.xml unchanged',
    dec.decode(original.get('word/document.xml')) === dec.decode(out.get('word/document.xml'))
  );
  const others = [...original.keys()].filter(
    (k) => dec.decode(original.get(k)) !== dec.decode(out.get(k) ?? new Uint8Array())
  );
  check('no part changed at all', others.length === 0, others.join(', '));
}

/* ------------------------------------------------------------------ */
console.log('\nchanging alignment rewrites only that paragraph');
{
  const edited = structuredClone(doc);
  edited.blocks[3].fmt = { ...edited.blocks[3].fmt, align: 'center' };
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));

  check('the new alignment is written', xml.includes('<w:jc w:val="center"/>'));
  check(
    'the indent it already had survives',
    /<w:ind w:left="720"[^>]*w:right="720"/.test(xml),
    (xml.match(/<w:ind[^>]*>/g) || []).join(' ')
  );
  check('the other paragraphs are untouched', xml.includes('<w:jc w:val="both"/>'));
  const others = [...original.keys()].filter(
    (k) => k !== 'word/document.xml' &&
      dec.decode(original.get(k)) !== dec.decode(out.get(k) ?? new Uint8Array())
  );
  check('no other part changed', others.length === 0, others.join(', '));
}

/* ------------------------------------------------------------------ */
console.log('\nw:pPr children stay in schema order');
{
  // Word offers to repair a file whose paragraph properties are out of
  // order, which is the failure this whole project exists to avoid.
  const ORDER = ['pStyle', 'numPr', 'pBdr', 'spacing', 'ind', 'jc', 'outlineLvl', 'rPr', 'sectPr'];
  const edited = structuredClone(doc);
  // Touch every paragraph, so every w:pPr is regenerated rather than kept.
  for (const b of edited.blocks) {
    b.fmt = { ...(b.fmt ?? {}), align: 'right', spaceBefore: 6, indentLeft: 18 };
  }
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));

  const pPrs = xml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/g) ?? [];
  check('every paragraph has properties', pPrs.length >= 9, String(pPrs.length));

  let bad = 0;
  for (const pPr of pPrs) {
    const names = [...pPr.matchAll(/<w:([A-Za-z0-9]+)[ />]/g)]
      .map((m) => m[1])
      .filter((n) => ORDER.includes(n));
    const ranks = names.map((n) => ORDER.indexOf(n));
    for (let i = 1; i < ranks.length; i++) if (ranks[i] < ranks[i - 1]) bad++;
  }
  check('no property is out of order', bad === 0, bad + ' inversions');
  check('all three were written', /<w:spacing/.test(xml) && /<w:ind/.test(xml) && /<w:jc/.test(xml));
  check('the document still parses as XML', !!new Q().parseFromString(xml, 'application/xml'));
}

/* ------------------------------------------------------------------ */
console.log('\nclearing formatting removes it from the file');
{
  const edited = structuredClone(doc);
  delete edited.blocks[0].fmt; // was centred with space after
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));
  const first = /<w:p>[\s\S]*?<\/w:p>/.exec(xml)?.[0] ?? '';
  check('the title is no longer centred', !first.includes('<w:jc'), first.slice(0, 160));
  check('and its spacing is gone', !first.includes('<w:spacing'), first.slice(0, 160));
  check('other paragraphs keep theirs', xml.includes('<w:jc w:val="right"/>'));
}

console.log(failures === 0 ? '\nall direct-format checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
