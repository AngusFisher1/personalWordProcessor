import type { Doc } from './model';
import { contentHeight, contentWidth, isTable } from './model';
import { docEl, pages } from './render';
import { styleOf } from './styles';
import { currentPageSetup, pageCount } from './paginate';

/**
 * The Recto workspace chrome.
 *
 * The rule the whole layout obeys: chrome never crosses the paper edge.
 * Page numbers, the document map and the status readout all live in the
 * gutter, so the page is the only thing that looks like a document and the
 * only light in the room.
 */

const ui = {
  outline: null as HTMLElement | null,
  spine: null as HTMLElement | null,
  readout: null as HTMLElement | null,
  title: null as HTMLElement | null,
  meta: null as HTMLElement | null,
  gutter: null as HTMLElement | null,
  setup: null as HTMLElement | null,
};

function el(tag: string, cls?: string, text?: string): HTMLElement {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
}

export function railEl(): HTMLElement {
  const r = document.getElementById('rail');
  if (!r) throw new Error('#rail missing');
  return r;
}

export function canvasEl(): HTMLElement {
  const c = document.getElementById('canvas');
  if (!c) throw new Error('#canvas missing');
  return c;
}

/* ------------------------------------------------------------------ *
 * Rail
 * ------------------------------------------------------------------ */

export interface RailHooks {
  onOutlineClick(blockId: string): void;
}

let hooks: RailHooks | null = null;

export function buildRail(rail: HTMLElement, h: RailHooks): void {
  hooks = h;
  rail.textContent = '';

  const head = el('div', 'rail-head');
  const mark = el('div', 'rail-mark');
  mark.appendChild(el('span', 'rail-dot'));
  mark.appendChild(el('span', 'rail-name', 'RECTO'));
  head.appendChild(mark);
  rail.appendChild(head);

  const docInfo = el('div', 'rail-doc');
  ui.title = el('div', 'rail-title', 'Untitled');
  ui.meta = el('div', 'rail-meta', '');
  docInfo.append(ui.title, ui.meta);
  rail.appendChild(docInfo);

  // What you act with comes first, then what the document is, then how the
  // page is set up. The outline sits in the middle because it is the only
  // part that grows.
  const actions = el('div', 'rail-actions');
  actions.id = 'rail-actions';
  rail.appendChild(actions);

  const tabs = el('div', 'rail-tabs');
  tabs.append(el('div', 'rail-tab on', 'OUTLINE'));
  rail.appendChild(tabs);

  ui.outline = el('div', 'rail-outline');
  rail.appendChild(ui.outline);

  ui.setup = el('div', 'rail-setup');
  rail.appendChild(ui.setup);
}

/** A row of key/value metadata, as used by PAGE SETUP. */
function setupRow(label: string, value: string): HTMLElement {
  const row = el('div', 'setup-row');
  row.append(el('span', '', label), el('span', 'setup-val', value));
  return row;
}

export function updateRail(
  doc: Doc,
  words: number,
  saved: string,
  failed = false
): void {
  if (ui.title) ui.title.textContent = doc.title || 'Untitled';
  const n = pageCount();
  if (ui.meta) {
    // The meta line already reports the save state, so there is no separate
    // label for it; failure turns this line itself into the warning.
    ui.meta.classList.toggle('err', failed);
    ui.meta.textContent =
      `${n} PP · ${words} W · ${saved}`.toUpperCase();
  }

  const setup = currentPageSetup();
  if (ui.setup) {
    ui.setup.textContent = '';
    ui.setup.appendChild(el('div', 'rail-label', 'PAGE SETUP'));
    const inches = (n2: number) => (n2 / 96).toFixed(2).replace(/\.00$/, '');
    ui.setup.appendChild(
      setupRow(
        setup.width === 816 && setup.height === 1056 ? 'US Letter' : 'Custom',
        `${inches(setup.width)} × ${inches(setup.height)} in`
      )
    );
    ui.setup.appendChild(
      setupRow('Margins', `${(setup.margins.top / 96).toFixed(2)} in`)
    );
    ui.setup.appendChild(
      setupRow('Body', `Source Serif 11/16`)
    );
    ui.setup.appendChild(
      setupRow(
        'Text box',
        `${Math.round(contentWidth(setup))} × ${Math.round(contentHeight(setup))}`
      )
    );
  }
}

/** Headings, in document order, as a clickable outline. */
export function updateOutline(doc: Doc, activeId: string | null): void {
  const box = ui.outline;
  if (!box) return;
  box.textContent = '';

  const entries: { id: string; text: string; sub: boolean }[] = [];
  for (const b of doc.blocks) {
    if (isTable(b)) continue;
    if (b.styleId !== 'SectionHeading' && b.styleId !== 'JobTitle') continue;
    const text = b.html.replace(/<[^>]*>/g, '').trim();
    if (!text) continue;
    entries.push({ id: b.id, text, sub: b.styleId === 'JobTitle' });
  }

  if (entries.length === 0) {
    const empty = el('div', 'outline-empty');
    empty.innerHTML =
      'NO HEADINGS YET<br>HEADINGS APPEAR HERE<br>AS YOU WRITE';
    box.appendChild(empty);
    return;
  }

  for (const e of entries) {
    const row = el('div', 'outline-row' + (e.sub ? ' sub' : ''));
    if (e.id === activeId) row.classList.add('on');
    const tick = el('span', 'outline-tick');
    row.append(tick, el('span', 'outline-text', e.text));
    row.addEventListener('mousedown', (ev) => ev.preventDefault());
    row.addEventListener('click', () => hooks?.onOutlineClick(e.id));
    box.appendChild(row);
  }
}

/* ------------------------------------------------------------------ *
 * Gutter: page numbers and page-break rules
 * ------------------------------------------------------------------ */

export function ensureGutter(): HTMLElement {
  if (ui.gutter && ui.gutter.isConnected) return ui.gutter;
  const g = el('div', 'gutter');
  g.id = 'gutter';
  canvasEl().appendChild(g);
  ui.gutter = g;
  return g;
}

/**
 * Page numbers sit beside each page with a short leader line, and a labelled
 * rule marks each break. Drawn as an overlay rather than inside the pages, so
 * nothing here can end up on the paper or in the printed output.
 */
export function updateGutter(currentPage: number): void {
  const g = ensureGutter();
  const canvas = canvasEl();
  const list = pages();
  g.textContent = '';
  const cRect = canvas.getBoundingClientRect();

  list.forEach((page, i) => {
    const r = page.getBoundingClientRect();
    const top = r.top - cRect.top + canvas.scrollTop;
    const left = r.left - cRect.left;

    const num = el('div', 'gutter-num', String(i + 1).padStart(2, '0'));
    if (i === currentPage) num.classList.add('on');
    num.style.top = top - 1 + 'px';
    num.style.left = left - 76 + 'px';
    g.appendChild(num);

    const lead = el('div', 'gutter-lead');
    if (i === currentPage) lead.classList.add('on');
    lead.style.top = top + 6 + 'px';
    lead.style.left = left - 26 + 'px';
    g.appendChild(lead);

    if (i < list.length - 1) {
      const rule = el('div', 'gutter-break');
      rule.style.top = top + r.height + 18 + 'px';
      rule.appendChild(el('span', 'gutter-break-line'));
      rule.appendChild(el('span', 'gutter-break-label', 'PAGE BREAK'));
      rule.appendChild(el('span', 'gutter-break-line'));
      g.appendChild(rule);
    }
  });
}

/* ------------------------------------------------------------------ *
 * Spine: the document map
 * ------------------------------------------------------------------ */

export function updateSpine(currentPage: number): void {
  let s = ui.spine;
  if (!s || !s.isConnected) {
    s = el('div', 'spine');
    s.id = 'spine';
    canvasEl().appendChild(s);
    ui.spine = s;
  }
  const list = pages();
  s.textContent = '';
  if (list.length === 0) return;

  const doc = docEl();
  const gap = 8;
  const avail = s.clientHeight - 24;
  const barH = Math.max(6, (avail - gap * (list.length - 1)) / list.length);

  list.forEach((_page, i) => {
    const bar = el('div', 'spine-bar');
    if (i === currentPage) bar.classList.add('on');
    bar.style.top = 12 + i * (barH + gap) + 'px';
    bar.style.height = barH + 'px';
    bar.title = `Page ${i + 1}`;
    bar.addEventListener('mousedown', (e) => e.preventDefault());
    bar.addEventListener('click', () => {
      const target = pages()[i];
      if (target) target.scrollIntoView({ block: 'start', behavior: 'smooth' });
    });
    s!.appendChild(bar);
  });

  // The slice of the document actually on screen.
  const total = doc.scrollHeight;
  if (total > 0) {
    const view = el('div', 'spine-view');
    const trackTop = 12;
    const trackH = list.length * barH + (list.length - 1) * gap;
    view.style.top = trackTop + (doc.scrollTop / total) * trackH + 'px';
    view.style.height =
      Math.max(14, (doc.clientHeight / total) * trackH) + 'px';
    s.appendChild(view);
  }
}

/* ------------------------------------------------------------------ *
 * Status readout
 * ------------------------------------------------------------------ */

export interface Readout {
  page: number;
  pages: number;
  line: number;
  col: number;
  words: number;
}

export function updateReadout(r: Readout): void {
  let n = ui.readout;
  if (!n || !n.isConnected) {
    n = el('div', 'readout');
    n.id = 'readout';
    canvasEl().appendChild(n);
    ui.readout = n;
  }
  const strong = (v: string | number) => `<b>${v}</b>`;
  n.innerHTML =
    `PAGE ${strong(String(r.page).padStart(2, '0'))} / ${String(r.pages).padStart(2, '0')}` +
    ` · LN ${strong(r.line)} COL ${strong(r.col)}` +
    ` · ${strong(r.words)} W`;
}

/** Which page index the caret is on, and the line/column within its block. */
export function caretReadout(words: number): Readout {
  const list = pages();
  const sel = window.getSelection();
  let page = 0;
  let line = 1;
  let col = 1;

  if (sel && sel.focusNode) {
    let node: Node | null = sel.focusNode;
    let blk: HTMLElement | null = null;
    while (node) {
      if (
        node.nodeType === Node.ELEMENT_NODE &&
        (node as HTMLElement).classList.contains('blk')
      ) {
        blk = node as HTMLElement;
        break;
      }
      node = node.parentNode;
    }
    if (blk) {
      const pageEl = blk.closest('.page') as HTMLElement | null;
      if (pageEl) page = Math.max(0, list.indexOf(pageEl));
      const r = document.createRange();
      r.setStart(blk, 0);
      try {
        r.setEnd(sel.focusNode, sel.focusOffset);
        const text = r.toString();
        const nl = text.lastIndexOf('\n');
        col = text.length - nl;
        line = text.split('\n').length;
      } catch {
        /* selection moved underneath us */
      }
    }
  }
  return { page: page + 1, pages: Math.max(1, list.length), line, col, words };
}

/** The page the caret is on, for the gutter and spine highlights. */
export function currentPageIndex(): number {
  const sel = window.getSelection();
  if (!sel || !sel.focusNode) return 0;
  let node: Node | null = sel.focusNode;
  while (node) {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const p = (node as HTMLElement).closest?.('.page') as HTMLElement | null;
      if (p) return Math.max(0, pages().indexOf(p));
    }
    node = node.parentNode;
  }
  return 0;
}

export function activeHeadingId(doc: Doc): string | null {
  const sel = window.getSelection();
  if (!sel || !sel.focusNode) return null;
  let node: Node | null = sel.focusNode;
  let blk: HTMLElement | null = null;
  while (node) {
    if (
      node.nodeType === Node.ELEMENT_NODE &&
      (node as HTMLElement).classList.contains('blk')
    ) {
      blk = node as HTMLElement;
      break;
    }
    node = node.parentNode;
  }
  if (!blk) return null;
  const id = blk.dataset.continuesFrom || blk.dataset.blockId;
  if (!id) return null;
  // The heading this block sits under.
  let last: string | null = null;
  for (const b of doc.blocks) {
    if (isTable(b)) continue;
    if (b.styleId === 'SectionHeading' || b.styleId === 'JobTitle') last = b.id;
    if (b.id === id) return last;
  }
  return last;
}

export { styleOf };
