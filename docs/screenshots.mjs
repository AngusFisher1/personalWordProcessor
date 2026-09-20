/**
 * The README's screenshots, captured from the running application.
 *
 * They were mockups once, and mockups drift: the first set showed a library
 * with folders and a preview pane that this program has never had. These are
 * driven through the real app in a real browser, so a screenshot that stops
 * being true is a screenshot that stops matching the next run of this file.
 *
 * The document in them is invented. Nothing here reads a real one.
 *
 *   npm run dev                       # in another terminal
 *   npm i -D playwright-core          # not a project dependency
 *   node docs/screenshots.mjs [baseUrl] [outDir]
 *
 * Uses the Chrome already installed on the machine rather than downloading
 * one; set CHROME_PATH if it is somewhere unusual.
 */
import { chromium } from 'playwright-core';
import { mkdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const BASE = process.argv[2] ?? 'http://localhost:5178';
const OUT = process.argv[3] ?? join(HERE, 'screenshots');
const CHROME =
  process.env.CHROME_PATH ??
  (process.platform === 'darwin'
    ? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    : process.platform === 'win32'
      ? 'C:/Program Files/Google/Chrome/Application/chrome.exe'
      : '/usr/bin/google-chrome');

const VIEWPORT = { width: 1560, height: 940 };
const SCALE = 2;

/* ------------------------------------------------------------------ *
 * The document
 * ------------------------------------------------------------------ */

let n = 0;
const id = () => 'b' + ++n;
const p = (styleId, html, extra = {}) => ({ id: id(), styleId, html, ...extra });

const CV = [
  p('Name', 'Miriam Okonkwo'),
  p('Contact', 'Brooklyn, New York · miriam@okonkwo.studio · okonkwo.studio'),

  p('SectionHeading', 'Summary'),
  p(
    'Body',
    'Product designer of fifteen years, most of them spent on tools that people ' +
      'are obliged to use rather than choose. I work close to the engineering, ' +
      'write the interface guidelines myself, and measure the result in the field ' +
      'rather than in a review.'
  ),

  p('SectionHeading', 'Experience'),
  p('JobTitle', 'Northline Instruments — Director of Product Design'),
  p('Body', '<i>Field-service console and hardware companion apps for 40,000 technicians across eleven countries.</i>'),
  p('Bullet', 'Led the ground-up redesign of the Northline console, reducing median task completion from 4.2 minutes to 1.6 and cutting first-week training time in half.'),
  p('Bullet', 'Built and now manage a team of nine designers and two researchers split between New York and Lisbon, with a written critique practice and no standing meetings.'),
  p('Bullet', 'Established the company&rsquo;s first type system, replacing four inherited typefaces with a single commissioned family and a documented scale.'),
  p('Bullet', 'Shipped offline-first editing to the entire technician fleet, which reduced support tickets attributed to data loss by 91% in the first two quarters.'),
  p('Bullet', 'Wrote the product&rsquo;s interface guidelines, now maintained by engineering as a versioned dependency rather than a slide deck.'),

  p('JobTitle', 'Kestrel &amp; Moss — Senior Product Designer'),
  p('Body', '<i>Manuscript and print-production tooling for independent publishers.</i>'),
  p('Bullet', 'Designed the manuscript editor behind Kestrel&rsquo;s publishing platform, in daily use at roughly 1,200 independent presses.'),
  p('Bullet', 'Owned pagination, footnote handling and print export from first prototype through general availability.'),
  p('Bullet', 'Rebuilt the footnote engine so that a note stays on the page that references it, even when the paragraph above it reflows.'),
  p('Bullet', 'Wrote the internal documentation standard still in use across the company&rsquo;s four product teams.'),

  p('JobTitle', 'Ashby Type Foundry — Interface Designer'),
  p('Body', '<i>Specimen tools, licensing flows and the foundry&rsquo;s first web proofing environment.</i>'),
  p('Bullet', 'Designed a proofing environment that let customers set real copy in a trial licence without downloading a file.'),
  p('Bullet', 'Rebuilt the specimen system around variable axes, which shortened release preparation from three weeks to four days.'),
  p('Bullet', 'Drew and shipped two retail families as a secondary responsibility.'),

  p('SectionHeading', 'Education'),
  p('Body', '<b>Rhode Island School of Design</b> — BFA Graphic Design, 2014'),
  p('Body', '<b>Cooper Union</b> — Type@Cooper, Extended Program, 2016'),

  p('SectionHeading', 'Selected Writing'),
  p('Body', '&ldquo;The Widow and the Orphan&rdquo; — <i>Works That Work</i>, 2024'),
  p('Body', '&ldquo;Against the Ribbon&rdquo; — <i>Brand New</i>, 2023'),
  p('Body', '&ldquo;Typesetting for people who are in a hurry&rdquo; — <i>Fonts in Use</i>, 2022'),

  p('SectionHeading', 'Skills &amp; Tools'),
  p('Body', 'Interface design, typography, design systems, research operations, print production, pagination, accessibility auditing.'),
  p('Body', 'Figma, Sketch, Glyphs, InDesign, HTML and CSS, TypeScript to the point of being useful in a code review.'),
];

const PAGE = {
  width: 816,
  height: 1056,
  margins: { top: 96, right: 96, bottom: 96, left: 96 },
};

const MAIN = {
  id: 'doc-cv',
  title: 'Okonkwo — Curriculum Vitae',
  page: PAGE,
  blocks: CV,
  headers: {
    default: [{ id: 'hdr1', styleId: 'Contact', html: 'OKONKWO — CURRICULUM VITAE' }],
  },
  footers: {
    default: [
      {
        id: 'ftr1',
        styleId: 'Contact',
        html:
          'Updated 14 March 2026 · page <span data-field="PAGE"></span> of ' +
          '<span data-field="NUMPAGES"></span>',
      },
    ],
  },
};

const EMPTY = { id: 'doc-empty', title: 'Untitled', page: PAGE, blocks: [p('Body', '')] };

/** A few more entries so the library looks like a library. */
const OTHERS = [
  ['doc-cover', 'Okonkwo — Cover letter, Fieldwork', 312, 1, 26],
  ['doc-refs', 'Okonkwo — References', 148, 1, 60 * 24 * 8],
  ['doc-reflow', 'Notes on Reflow — draft 4', 4118, 9, 60 * 24 * 11],
  ['doc-guides', 'Northline — Interface guidelines v6', 12904, 31, 60 * 24 * 18],
  ['doc-ribbon', 'Against the Ribbon — final', 2740, 6, 60 * 24 * 32],
];

function seed(palette, openId) {
  const now = Date.now();
  const store = {};
  const index = [];

  const put = (docObj, words, pages, minutesAgo) => {
    store['wp:doc:' + docObj.id] = JSON.stringify(docObj);
    index.push({
      id: docObj.id,
      title: docObj.title,
      updatedAt: now - minutesAgo * 60000,
      words,
      pages,
    });
  };

  put(MAIN, 641, 2, 4);
  for (const [oid, title, words, pages, minutesAgo] of OTHERS) {
    put(
      { id: oid, title, page: PAGE, blocks: [{ id: oid + '-b', styleId: 'Body', html: title }] },
      words,
      pages,
      minutesAgo
    );
  }
  put(EMPTY, 0, 1, 60 * 24 * 40);
  // Newest first, the way the app writes it.
  index.sort((a, b) => b.updatedAt - a.updatedAt);

  store['wp:docs'] = JSON.stringify(index);
  store['wp:last'] = openId;
  store['wp:palette'] = palette;
  return store;
}

/* ------------------------------------------------------------------ *
 * Driving
 * ------------------------------------------------------------------ */

async function fresh(browser, { palette = 'night', open = 'doc-cv' } = {}) {
  const context = await browser.newContext({
    viewport: VIEWPORT,
    deviceScaleFactor: SCALE,
    colorScheme: 'dark',
  });
  const store = seed(palette, open);
  await context.addInitScript((s) => {
    for (const [k, v] of Object.entries(s)) window.localStorage.setItem(k, v);
  }, store);
  const page = await context.newPage();
  await page.goto(BASE, { waitUntil: 'networkidle' });
  await page.evaluate(() => document.fonts.ready);
  await page.waitForTimeout(700);
  return { context, page };
}

/** Put the caret in a block, by its index among the rendered blocks. */
async function caretIn(page, blockIndex, offset = 0) {
  await page.evaluate(
    ({ blockIndex, offset }) => {
      const blk = document.querySelectorAll('#doc .page-content .blk')[blockIndex];
      const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
      const node = walker.nextNode() ?? blk;
      const r = document.createRange();
      r.setStart(node, Math.min(offset, node.length ?? 0));
      r.collapse(true);
      const s = getSelection();
      s.removeAllRanges();
      s.addRange(r);
      document.getElementById('doc').focus();
    },
    { blockIndex, offset }
  );
}

async function shot(page, name) {
  await page.waitForTimeout(350);
  await page.screenshot({ path: `${OUT}/${name}.png` });
  console.log('  wrote ' + name + '.png');
}

/* ------------------------------------------------------------------ *
 * The nine
 * ------------------------------------------------------------------ */

async function main() {
  await mkdir(OUT, { recursive: true });
  const browser = await chromium.launch({ executablePath: CHROME });

  /* 01 — the editing view, with a page break in frame */
  {
    const { context, page } = await fresh(browser);
    await caretIn(page, 8, 40);
    await page.evaluate(() => {
      // Frame the page break: the bottom of page one, the labelled rule, and
      // the top of page two. Pagination is the thesis, so it goes in the hero.
      const pages = document.querySelectorAll('#doc .page');
      const doc = document.getElementById('doc');
      const gap = pages[0].getBoundingClientRect().bottom - doc.getBoundingClientRect().top;
      doc.scrollTop += gap - doc.clientHeight * 0.66;
    });
    await shot(page, '01-editing');
    await context.close();
  }

  /* 02 — a selection, with the formatting bar over it */
  {
    const { context, page } = await fresh(browser);
    await caretIn(page, 9, 0);
    for (let i = 0; i < 58; i++) await page.keyboard.press('Shift+ArrowRight');
    await page.waitForTimeout(250);
    await shot(page, '02-selection');
    await context.close();
  }

  /* 03 — the command palette */
  {
    const { context, page } = await fresh(browser);
    await caretIn(page, 6, 0);
    await page.keyboard.press('Control+k');
    await page.waitForTimeout(250);
    await page.keyboard.type('ex', { delay: 40 });
    await shot(page, '03-command-palette');
    await context.close();
  }

  /* 04 — the library */
  {
    const { context, page } = await fresh(browser);
    await page.evaluate(() => {
      document.querySelectorAll('.rail-tab').forEach((t) => {
        if (t.textContent.trim() === 'FILES') t.click();
      });
    });
    await shot(page, '04-library');
    await context.close();
  }

  /* 05 — find and replace, with matches lit and the map ticked */
  {
    const { context, page } = await fresh(browser);
    await caretIn(page, 3, 0);
    await page.keyboard.press('Control+f');
    await page.waitForTimeout(250);
    await page.keyboard.type('design', { delay: 45 });
    await page.waitForTimeout(600);
    await shot(page, '05-find-replace');
    await context.close();
  }

  /* 06 — header and footer editing */
  {
    const { context, page } = await fresh(browser);
    await page.evaluate(() => {
      const b = [...document.querySelectorAll('#rail-actions button')].find((x) =>
        x.textContent.startsWith('Header & footer')
      );
      b?.click();
    });
    await page.waitForTimeout(500);
    await page.evaluate(() => {
      document.querySelector('#doc .page-header .blk')?.scrollIntoView({ block: 'center' });
      document.getElementById('doc').scrollTop = 0;
    });
    await shot(page, '06-header-footer');
    await context.close();
  }

  /* 07 — the export sheet */
  {
    const { context, page } = await fresh(browser);
    await caretIn(page, 4, 0);
    await page.keyboard.press('Control+e');
    await shot(page, '07-export');
    await context.close();
  }

  /* 08 — a new, empty document */
  {
    const { context, page } = await fresh(browser, { open: 'doc-empty' });
    await caretIn(page, 0, 0);
    await shot(page, '08-empty');
    await context.close();
  }

  /* 09 — a light palette */
  {
    const { context, page } = await fresh(browser, { palette: 'mossStone' });
    await caretIn(page, 5, 0);
    await page.evaluate(() => {
      document.querySelectorAll('#doc .page-content .blk')[4]?.scrollIntoView({ block: 'start' });
      document.getElementById('doc').scrollBy(0, -60);
    });
    await shot(page, '09-light');
    await context.close();
  }

  await browser.close();
  console.log('done');
}

await main();
