/**
 * Does our pagination agree with Word's?
 *
 * Every .docx Word saves records the page count it had at that moment in
 * docProps/app.xml. 52 of 63 documents in a real corpus carry one. That is
 * ground truth from the program we are trying to replace, sitting inside
 * files we already have, needing no Word and no LibreOffice to read.
 *
 * The app is driven in a real browser, because pagination is measurement:
 * there is no way to ask this question in Node, where nothing has a height.
 *
 *   npm run dev                 # in another terminal
 *   npm i -D playwright-core
 *   node test/pagination.mjs [dir] [baseUrl]
 *
 * Output is counts and file names. No document text is read or printed.
 */
import { chromium } from 'playwright-core';
import { readdir, readFile } from 'node:fs/promises';
import { join, extname, basename, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const HERE = dirname(fileURLToPath(import.meta.url));
const DIR = process.argv[2] ?? join(HERE, 'corpus');
const BASE = process.argv[3] ?? 'http://localhost:5178';
const CHROME =
  process.env.CHROME_PATH ??
  (process.platform === 'darwin'
    ? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    : process.platform === 'win32'
      ? 'C:/Program Files/Google/Chrome/Application/chrome.exe'
      : '/usr/bin/google-chrome');

const files = [];
async function walk(dir, depth = 0) {
  if (depth > 5) return;
  let entries;
  try { entries = await readdir(dir, { withFileTypes: true }); } catch { return; }
  for (const e of entries) {
    const p = join(dir, e.name);
    if (e.isDirectory()) await walk(p, depth + 1);
    else if (extname(e.name).toLowerCase() === '.docx' && !e.name.startsWith('~$')) files.push(p);
  }
}
await walk(DIR);

/** The page count Word itself recorded, or null when it did not. */
async function wordPages(bytes) {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const entry = zip.file('docProps/app.xml');
    if (!entry) return null;
    const xml = await entry.async('string');
    const app = /<Application>([^<]*)<\/Application>/.exec(xml)?.[1] ?? '';
    // Only Word's own number is ground truth. Other producers copy the
    // field across without recomputing it.
    if (!/Microsoft.*Word/i.test(app)) return null;
    const n = Number(/<Pages>(\d+)<\/Pages>/.exec(xml)?.[1] ?? 0);
    return n > 0 ? n : null;
  } catch {
    return null;
  }
}

const browser = await chromium.launch({ executablePath: CHROME });
const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
const page = await context.newPage();
await page.goto(BASE, { waitUntil: 'networkidle' });
await page.evaluate(() => document.fonts.ready);

const rows = [];
let skipped = 0;

for (const f of files) {
  const bytes = await readFile(f);
  const expected = await wordPages(bytes);
  if (expected === null) {
    skipped++;
    continue;
  }
  const b64 = bytes.toString('base64');
  let got;
  try {
    got = await page.evaluate(async (data) => {
      const bin = atob(data);
      const arr = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      const file = new File([arr], 'doc.docx');
      await window.wp.openDocx(file);
      // Pagination settles over a couple of frames: fonts, then the second
      // pass that resolves NUMPAGES in a footer.
      const settle = async () => {
        let last = -1;
        for (let i = 0; i < 40; i++) {
          await new Promise((r) => setTimeout(r, 60));
          const n = window.wp.pageCount();
          if (n === last) return n;
          last = n;
        }
        return last;
      };
      const n = await settle();
      const blk = document.querySelector('#doc .blk');
      return {
        pages: n,
        asked: window.wp.doc().defaultFont ?? null,
        // The first family the browser actually had. If it is not the one
        // the document asked for, our line breaks cannot match Word's.
        used: blk
          ? getComputedStyle(blk).fontFamily.split(',')[0].replace(/"/g, '').trim()
          : null,
      };
    }, b64);
  } catch (err) {
    rows.push({ name: basename(f), expected, got: null, err: String(err).slice(0, 60) });
    continue;
  }
  rows.push({ name: basename(f), expected, got: got.pages, asked: got.asked, used: got.used });
}

await browser.close();

/* ------------------------------------------------------------------ */

const usable = rows.filter((r) => r.got !== null);
const exact = usable.filter((r) => r.got === r.expected);
const near = usable.filter((r) => r.got !== r.expected && Math.abs(r.got - r.expected) === 1);
const off = usable.filter((r) => Math.abs(r.got - r.expected) > 1);
const pct = (n) => ((n / Math.max(1, usable.length)) * 100).toFixed(0) + '%';

console.log(`\n${files.length} documents, ${usable.length} with a page count from Word` +
  (skipped ? ` (${skipped} without one)` : ''));
console.log(`  exact match      ${String(exact.length).padStart(3)}  ${pct(exact.length)}`);
console.log(`  within one page  ${String(near.length).padStart(3)}  ${pct(near.length)}`);
console.log(`  further out      ${String(off.length).padStart(3)}  ${pct(off.length)}`);

if (off.length > 0) {
  console.log('\nfurther than one page from Word:');
  for (const r of off.sort((a, b) => Math.abs(b.got - b.expected) - Math.abs(a.got - a.expected))) {
    console.log(
      `  ${r.name.slice(0, 48).padEnd(50)} Word ${String(r.expected).padStart(3)}  ours ${String(r.got).padStart(3)}`
    );
  }
}
const substituted = usable.filter(
  (r) => r.asked && r.used && r.asked.toLowerCase() !== r.used.toLowerCase()
);
if (substituted.length > 0) {
  const by = new Map();
  for (const r of substituted) {
    const k = r.asked + ' → ' + r.used;
    by.set(k, (by.get(k) ?? 0) + 1);
  }
  console.log(
    `
font substituted on ${substituted.length} of ${usable.length} documents ` +
      `(${pct(substituted.length)}); their line breaks cannot match Word's:`
  );
  for (const [k, v] of [...by].sort((a, b) => b[1] - a[1])) {
    console.log(`  ${String(v).padStart(3)}  ${k}`);
  }
}

const failed = rows.filter((r) => r.got === null);
if (failed.length > 0) {
  console.log('\nfailed to open:');
  for (const r of failed) console.log(`  ${r.name.slice(0, 48).padEnd(50)} ${r.err}`);
}

// Agreement within one page is the bar: Word's own count moves by a page
// between printer drivers, and ours is measured in a browser.
const rate = (exact.length + near.length) / Math.max(1, usable.length);
console.log(`\nwithin one page of Word on ${(rate * 100).toFixed(0)}% of documents`);
process.exit(off.length === 0 ? 0 : 1);
