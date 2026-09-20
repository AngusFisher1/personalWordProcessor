/**
 * What the importer makes of a folder of documents.
 *
 * Reports the style each paragraph was mapped to, and which original Word
 * style ids drove that mapping, so gaps in the mapping table are visible
 * rather than guessed at. Prints style names and counts only, never document
 * text.
 *
 *   node test/style-report.mjs "/path/to/your/documents"
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readdir, readFile } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
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
const ROOT = process.argv[2] || join(HERE, 'corpus');

const styleTally = {};
const pStyleTally = {};
/** w:pStyle values that produced a Body block, i.e. that we did not recognize. */
const unmapped = {};
let docs = 0;
let failed = 0;

function bump(obj, key) {
  obj[key] = (obj[key] ?? 0) + 1;
}

async function walk(dir, depth = 0) {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const e of entries) {
    const full = join(dir, e.name);
    if (e.isDirectory()) {
      if (depth < 5 && e.name !== 'node_modules' && !e.name.startsWith('.')) {
        await walk(full, depth + 1);
      }
      continue;
    }
    if (!e.name.toLowerCase().endsWith('.docx') || e.name.startsWith('~$')) continue;
    let bytes;
    try {
      bytes = await readFile(full);
    } catch {
      continue;
    }
    try {
      const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
      const { doc, vault } = await importDocx(buf, basename(full));
      docs++;
      for (const b of doc.blocks) bump(styleTally, b.styleId);
      // Which w:pStyle each block carried, read back out of the preserved XML.
      for (const b of doc.blocks) {
        const xml = vault.blockXml.get(b.id) ?? '';
        const m = xml.match(/<w:pStyle w:val="([^"]*)"/);
        const p = m ? m[1] : '(none)';
        bump(pStyleTally, p);
        if (b.styleId === 'Body') bump(unmapped, p);
      }
    } catch {
      failed++;
    }
  }
}

await walk(ROOT);

const show = (title, obj, limit = 25) => {
  console.log('\n' + title);
  const rows = Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, limit);
  const w = Math.max(...rows.map(([k]) => k.length), 4);
  for (const [k, v] of rows) console.log('  ' + k.padEnd(w) + '  ' + v);
};

console.log(`${docs} documents imported, ${failed} failed`);
show('mapped to', styleTally);
show('original w:pStyle', pStyleTally);
show('w:pStyle that fell through to Body', unmapped);
