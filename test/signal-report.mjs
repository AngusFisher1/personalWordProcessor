/**
 * Why each paragraph was classified the way it was.
 *
 * Groups paragraphs by their signal profile so a misfiring rule is visible as
 * a population rather than found by reading documents. Signals and counts
 * only, never text.
 *
 *   node test/signal-report.mjs "/path/to/documents"
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readdir, readFile } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
class Q extends DOMParser { constructor() { super({ onError: () => {} }); } }
globalThis.DOMParser = Q; globalThis.XMLSerializer = XMLSerializer;
const { importDocx } = await import('./build/harness.mjs');
const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = process.argv[2] || join(HERE, 'corpus');

const profile = {};
async function walk(dir, depth = 0) {
  let es; try { es = await readdir(dir, { withFileTypes: true }); } catch { return; }
  for (const e of es) {
    const f = join(dir, e.name);
    if (e.isDirectory()) { if (depth < 5 && !e.name.startsWith('.')) await walk(f, depth + 1); continue; }
    if (!e.name.toLowerCase().endsWith('.docx') || e.name.startsWith('~$')) continue;
    let bytes; try { bytes = await readFile(f); } catch { continue; }
    try {
      const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
      const { doc, vault } = await importDocx(buf, basename(f));
      for (const b of doc.blocks) {
        const xml = vault.blockXml.get(b.id) ?? '';
        const short = (b.html.replace(/<[^>]*>/g, '').trim().length) <= 80;
        if (!short) continue;
        const bold = /<w:b\/>|<w:b /.test(xml);
        const caps = /<w:caps/.test(xml);
        const sz = (xml.match(/<w:sz w:val="(\d+)"/) || [])[1] ?? '-';
        const ruled = /<w:pBdr>/.test(xml);
        const before = /<w:spacing[^>]*w:before="([1-9]\d*)"/.test(xml);
        const key = [b.styleId, bold ? 'bold' : '----', caps ? 'caps' : '----',
                     ruled ? 'rule' : '----', before ? 'spB' : '---', 'sz=' + sz].join(' ');
        profile[key] = (profile[key] ?? 0) + 1;
      }
    } catch {}
  }
}
await walk(ROOT);
const rows = Object.entries(profile).sort((a, b) => b[1] - a[1]).slice(0, 30);
console.log('short paragraphs by signal profile (style, bold, caps, rule, spaceBefore, size)\n');
for (const [k, v] of rows) console.log('  ' + String(v).padStart(4) + '  ' + k);
