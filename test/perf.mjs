/**
 * How the editor behaves on a document nobody would call small.
 *
 * The acceptance tests measured a ten-page document. The library has a
 * thirty-one page one in it, and "the only word processor I need" has to
 * survive the longest thing you will ever write, not the shortest.
 *
 * Measured in a real browser, because every number here is a layout number.
 *
 *   npm run dev                 # in another terminal
 *   npm i -D playwright-core
 *   node test/perf.mjs [pages] [baseUrl]
 */
import { chromium } from 'playwright-core';

const PAGES = Number(process.argv[2] ?? 50);
const BASE = process.argv[3] ?? 'http://localhost:5178';
const CHROME =
  process.env.CHROME_PATH ??
  (process.platform === 'win32'
    ? 'C:/Program Files/Google/Chrome/Application/chrome.exe'
    : process.platform === 'darwin'
      ? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
      : '/usr/bin/google-chrome');

/** Roughly nine blocks to a page at these styles. */
const BLOCKS = PAGES * 9;

const browser = await chromium.launch({ executablePath: CHROME });
const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });

await context.addInitScript(
  ({ blocks }) => {
    const lorem =
      'This paragraph exists to occupy a realistic amount of space on the page ' +
      'so that pagination has something to do. ';
    const doc = {
      id: 'perf',
      title: 'Performance',
      page: { width: 816, height: 1056, margins: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [],
    };
    for (let i = 0; i < blocks; i++) {
      const kind = i % 9;
      if (kind === 0) {
        doc.blocks.push({ id: 'b' + i, styleId: 'SectionHeading', html: 'Section ' + i });
      } else if (kind === 1) {
        doc.blocks.push({ id: 'b' + i, styleId: 'JobTitle', html: 'A subheading at ' + i });
      } else if (kind === 8) {
        doc.blocks.push({
          id: 'b' + i,
          styleId: 'Bullet',
          html: 'A bullet with <b>bold</b> and <i>italic</i> and a ' +
            '<span data-color="C00000">coloured</span> word. ' + lorem,
        });
      } else {
        doc.blocks.push({ id: 'b' + i, styleId: 'Body', html: lorem.repeat(2) });
      }
    }
    localStorage.setItem('wp:doc:perf', JSON.stringify(doc));
    localStorage.setItem(
      'wp:docs',
      JSON.stringify([{ id: 'perf', title: 'Performance', updatedAt: Date.now(), words: 0, pages: 1 }])
    );
    localStorage.setItem('wp:last', 'perf');
  },
  { blocks: BLOCKS }
);

const page = await context.newPage();
const t0 = Date.now();
await page.goto(BASE, { waitUntil: 'networkidle' });
await page.evaluate(() => document.fonts.ready);
await page.waitForFunction(() => window.wp && window.wp.pageCount() > 1, null, { timeout: 60000 });
const bootMs = Date.now() - t0;

const result = await page.evaluate(async () => {
  const q = (n) => Math.round(n * 100) / 100;
  const pct = (xs, p) => {
    const s = [...xs].sort((a, b) => a - b);
    return q(s[Math.min(s.length - 1, Math.floor(s.length * p))] ?? 0);
  };

  const pages = window.wp.pageCount();
  const blocks = document.querySelectorAll('#doc .blk').length;

  /** Type a character and measure the synchronous work it causes. */
  const typeInto = (blk, samples) => {
    const times = [];
    const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
    const node = walker.nextNode();
    if (!node) return times;
    const sel = getSelection();
    for (let i = 0; i < samples; i++) {
      const r = document.createRange();
      r.setStart(node, Math.min(5, node.data.length));
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      const t = performance.now();
      node.insertData(Math.min(5, node.data.length), 'x');
      document.getElementById('doc').dispatchEvent(new InputEvent('input', { bubbles: true }));
      times.push(performance.now() - t);
    }
    return times;
  };

  const blks = [...document.querySelectorAll('#doc .blk')];
  const first = blks[3];
  const last = blks[blks.length - 4];

  const early = typeInto(first, 60);
  const late = typeInto(last, 60);

  // A full reflow from the top, which is what a margin or style change costs.
  const fullStart = performance.now();
  window.wp.paginate();
  const fullMs = performance.now() - fullStart;

  const mem = performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : null;

  return {
    pages,
    blocks,
    earlyMedian: pct(early, 0.5),
    earlyP99: pct(early, 0.99),
    lateMedian: pct(late, 0.5),
    encoreP99: pct(late, 0.99),
    fullMs: q(fullMs),
    mem,
  };
});

await browser.close();

const line = (label, value, unit = 'ms') =>
  console.log('  ' + label.padEnd(34) + String(value).padStart(8) + ' ' + unit);

console.log(`\n${result.pages} pages, ${result.blocks} blocks`);
line('boot to first full layout', bootMs);
line('keystroke on page 1, median', result.earlyMedian);
line('keystroke on page 1, p99', result.earlyP99);
line('keystroke on the last page, median', result.lateMedian);
line('keystroke on the last page, p99', result.encoreP99);
line('full reflow from the top', result.fullMs);
if (result.mem !== null) line('JS heap', result.mem, 'MB');

// The bar the acceptance tests set at ten pages: a keystroke must stay
// under a frame wherever in the document it lands.
const worst = Math.max(result.earlyP99, result.encoreP99);
console.log(
  worst <= 16
    ? `\nevery keystroke inside one frame (worst p99 ${worst}ms)`
    : `\nSLOW: worst keystroke p99 is ${worst}ms, over a 16ms frame`
);
process.exit(worst <= 16 ? 0 : 1);
