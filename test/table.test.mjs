/**
 * Editing inside a table cell has to survive export.
 *
 * The risk this guards is silent loss: a table that exports from the vault
 * verbatim would quietly discard an edit made in one of its cells.
 *
 *   npm run test:table
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

const bytes = await readFile(join(HERE, 'corpus', 'report-tables.docx'));
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const original = await unzip(buf);
const { doc, vault } = await importDocx(buf, 'report-tables.docx');

const table = doc.blocks.find((b) => b.kind === 'table');
console.log('\ntable import');
check('a table block exists', !!table);
check('rows read', table.rows.length === 4, String(table?.rows.length));
check('header row flagged', table.rows[0].headerRow === true);
check('columns read', table.cols.length === 3, String(table?.cols.length));
check('column widths are real', table.cols.every((c) => c.width > 8));
check('cell text read', table.rows[1].cells[0][0].html.includes('North'));

console.log('\nuntouched table exports verbatim');
{
  const out = await unzip(await (await exportDocx(doc, vault)).arrayBuffer());
  check(
    'document.xml unchanged',
    dec.decode(original.get('word/document.xml')) === dec.decode(out.get('word/document.xml'))
  );
}

console.log('\nediting a cell survives export');
{
  const edited = structuredClone(doc);
  const t = edited.blocks.find((b) => b.kind === 'table');
  t.rows[1].cells[1][0].html = '9,999';
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));
  check('the edit is in the file', xml.includes('9,999'));
  check('the old value is gone', !xml.includes('1,240'));
  check('the table is still a table', (xml.match(/<w:tbl[ >]/g) || []).length === 1);
  check('row count preserved', (xml.match(/<w:tr[ >]/g) || []).length === 4);
  check('cell count preserved', (xml.match(/<w:tc[ >]/g) || []).length === 12);
  check('table properties preserved', xml.includes('<w:tblPr>'));
  check('grid preserved', xml.includes('<w:tblGrid>'));
  check('header flag preserved', xml.includes('<w:tblHeader'));
  const others = [...original.keys()].filter(
    (k) => k !== 'word/document.xml' &&
      dec.decode(original.get(k)) !== dec.decode(out.get(k) ?? new Uint8Array())
  );
  check('other parts untouched', others.length === 0, others.join(', '));
}

console.log(failures === 0 ? '\nall table checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
