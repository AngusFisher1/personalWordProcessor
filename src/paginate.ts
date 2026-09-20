import type { Block, MarginKey } from './model';
import { MARGINS, contentHeight, contentWidth, newId } from './model';
import { getCaret, nearestBlock, setCaret } from './caret';
import {
  blocksIn,
  docEl,
  makeBlockEl,
  measureEl,
  newPage,
  pageContent,
  pages,
} from './render';
import { styleClass, styleOf } from './styles';

/* ------------------------------------------------------------------ *
 * Layout geometry
 * ------------------------------------------------------------------ */

let margin: MarginKey = 'narrow';

export function setMargin(m: MarginKey): void {
  margin = m;
  document.documentElement.style.setProperty('--pad', MARGINS[m] + 'px');
  clearHeightCache();
}

export function currentMargin(): MarginKey {
  return margin;
}

function limitH(): number {
  return contentHeight(margin);
}

function limitW(): number {
  return contentWidth(margin);
}

/* ------------------------------------------------------------------ *
 * Height cache
 *
 * Keyed by (content width, styleId, html) rather than by block id: identical
 * blocks then share an entry, and editing one block never invalidates another.
 * Cleared when the font loads or the margin changes.
 * ------------------------------------------------------------------ */

const heights = new Map<string, number>();
const CACHE_CAP = 4000;

export function clearHeightCache(): void {
  heights.clear();
}

function keyFor(el: HTMLElement): string {
  return limitW() + '|' + styleOf(el) + '|' + el.innerHTML;
}

function heightOf(el: HTMLElement): number {
  const k = keyFor(el);
  const hit = heights.get(k);
  if (hit !== undefined) return hit;
  // The element is already laid out at the content width, so its own
  // offsetHeight is measured under exactly the render conditions. The
  // #measure sandbox is only needed when the element is not in the page.
  let h = el.offsetHeight;
  if (h === 0) h = measureClone(el);
  if (heights.size > CACHE_CAP) heights.clear();
  heights.set(k, h);
  return h;
}

function measureClone(el: HTMLElement): number {
  const m = measureEl();
  m.style.width = limitW() + 'px';
  const clone = el.cloneNode(true) as HTMLElement;
  m.appendChild(clone);
  const h = clone.offsetHeight;
  clone.remove();
  return h;
}

/** Measure model blocks without rendering them into the document. */
export function measureHeights(blocks: Block[], contentW: number): number[] {
  const m = measureEl();
  m.style.width = contentW + 'px';
  m.textContent = '';
  const els = blocks.map(makeBlockEl);
  for (const e of els) m.appendChild(e);
  const out = els.map((e) => e.offsetHeight); // append all, then read: one layout
  m.textContent = '';
  return out;
}

/* ------------------------------------------------------------------ *
 * Normalization
 *
 * contenteditable will periodically drop a bare text node, a stray div or a
 * br straight into .page-content, usually at a boundary. Repair it before
 * measuring anything.
 * ------------------------------------------------------------------ */

function wrapAsBody(nodes: Node[]): HTMLElement {
  const el = document.createElement('div');
  el.className = 'blk ' + styleClass('Body');
  el.dataset.blockId = newId();
  for (const n of nodes) el.appendChild(n);
  if (!el.firstChild) el.innerHTML = '<br>';
  return el;
}

function isBlk(n: Node): n is HTMLElement {
  return (
    n.nodeType === Node.ELEMENT_NODE &&
    (n as HTMLElement).classList.contains('blk')
  );
}

export function normalize(): void {
  const root = docEl();

  // Stray nodes directly under #doc: fold them into the first page.
  for (const n of Array.from(root.childNodes)) {
    if (
      n.nodeType === Node.ELEMENT_NODE &&
      (n as HTMLElement).classList.contains('page')
    ) {
      continue;
    }
    if (n.nodeType === Node.TEXT_NODE && (n.textContent ?? '').trim() === '') {
      root.removeChild(n);
      continue;
    }
    let first = pages(root)[0];
    if (!first) {
      first = newPage();
      root.insertBefore(first, root.firstChild);
    }
    const c = pageContent(first);
    if (
      n.nodeType === Node.ELEMENT_NODE &&
      (n as HTMLElement).querySelector('.blk')
    ) {
      // A page worth of blocks got hoisted out of its container; put them back.
      const el = n as HTMLElement;
      for (const b of Array.from(el.querySelectorAll('.blk'))) c.appendChild(b);
      el.remove();
    } else {
      c.insertBefore(wrapAsBody([n]), c.firstChild);
    }
  }

  if (pages(root).length === 0) root.appendChild(newPage());

  for (const page of pages(root)) {
    const c = pageContent(page);

    // Hoist blocks the browser nested inside other blocks.
    for (const nested of Array.from(c.querySelectorAll('.blk'))) {
      if (nested.parentElement === c) continue;
      let top: Element = nested;
      while (top.parentElement && top.parentElement !== c) {
        top = top.parentElement;
      }
      c.insertBefore(nested, top.nextSibling);
    }

    // Anything that is not a .blk becomes one, or goes away.
    let run: Node[] = [];
    const flush = (before: Node | null) => {
      if (run.length === 0) return;
      const el = wrapAsBody(run);
      c.insertBefore(el, before);
      run = [];
    };
    for (const n of Array.from(c.childNodes)) {
      if (isBlk(n)) {
        flush(n);
        continue;
      }
      if (n.nodeType === Node.TEXT_NODE) {
        if ((n.textContent ?? '').trim() === '') c.removeChild(n);
        else run.push(n);
        continue;
      }
      if (n.nodeType === Node.ELEMENT_NODE) {
        const el = n as HTMLElement;
        if (el.tagName === 'BR' || el.textContent === '') {
          el.remove();
        } else {
          run.push(...Array.from(el.childNodes)); // unwrap its text
          el.remove();
        }
        continue;
      }
      c.removeChild(n);
    }
    flush(null);

    // Mint an id for any block a browser clone stripped one from.
    for (const b of Array.from(c.children) as HTMLElement[]) {
      if (!b.dataset.blockId) b.dataset.blockId = newId();
      if (b.childNodes.length === 0) b.innerHTML = '<br>';
    }
  }
}

/** Keep one trailing empty Body block so clicking below the last line works. */
export function ensureTrailingBlock(): boolean {
  const root = docEl();
  const all = blocksIn(root);
  const last = all[all.length - 1];
  if (last && (last.textContent ?? '').trim() === '') return false;
  const ps = pages(root);
  const page = ps[ps.length - 1];
  if (!page) return false;
  const el = document.createElement('div');
  el.className = 'blk ' + styleClass('Body');
  el.dataset.blockId = newId();
  el.innerHTML = '<br>';
  pageContent(page).appendChild(el);
  return true;
}

/* ------------------------------------------------------------------ *
 * Pagination
 * ------------------------------------------------------------------ */

function topBlocks(page: Element): HTMLElement[] {
  return Array.from(pageContent(page).children).filter((c) =>
    c.classList.contains('blk')
  ) as HTMLElement[];
}

function firstBlock(page: Element): HTMLElement | null {
  return topBlocks(page)[0] ?? null;
}

/**
 * Height actually used by a page.
 *
 * NOT scrollHeight: .page-content has a fixed height, so scrollHeight never
 * reports less than the full content box and the "room reopened" test could
 * never fire. Summing heightOf() also keeps this identical to the number the
 * reflow accumulator uses, so the cheap check and the full reflow can never
 * disagree about whether a block fits.
 */
function usedHeight(content: HTMLElement): number {
  let h = 0;
  for (const c of Array.from(content.children) as HTMLElement[]) h += heightOf(c);
  return h;
}

function pageIndexOf(page: Element): number {
  return pages().indexOf(page as HTMLElement);
}

/**
 * Reconcile a page against the blocks assigned to it. Blocks are MOVED with
 * insertBefore, never re-created from the model: moving preserves focus,
 * avoids flicker and keeps IME composition alive.
 */
function place(content: HTMLElement, want: HTMLElement[]): void {
  let node = content.firstElementChild;
  for (let i = 0; i < want.length; i++) {
    if (node === want[i]) {
      node = node.nextElementSibling;
    } else {
      content.insertBefore(want[i], node);
    }
  }
  // Anything left over belongs to a later page and is re-parented by that
  // page's own reconcile pass.
}

/**
 * Full reflow. `fromPage` leaves earlier pages untouched and starts the
 * accumulator at the first block of that page: typing on page 1 of a 10-page
 * document must not churn pages 2-10.
 */
export function paginate(opts?: { fromPage?: number }): void {
  const root = docEl();
  normalize();
  const caret = getCaret();

  const before = pages(root);
  const from = Math.max(0, Math.min(opts?.fromPage ?? 0, before.length - 1));

  const seq: HTMLElement[] = [];
  for (let i = from; i < before.length; i++) seq.push(...topBlocks(before[i]));

  // Measure everything up front: one forced layout, then only moves.
  const h = seq.map(heightOf);
  const limit = limitH();

  const assign: HTMLElement[][] = [[]];
  let running = 0;
  for (let i = 0; i < seq.length; i++) {
    let cur = assign[assign.length - 1];
    if (running + h[i] > limit && cur.length >= 1) {
      // A block taller than the content box still gets its own page and is
      // clipped. Line-level splitting is out of scope for v1; the v2 approach
      // is Range.getClientRects() per line box plus a binary search for the
      // split offset, with widow/orphan control.
      cur = [];
      assign.push(cur);
      running = 0;
    }
    cur.push(seq[i]);
    running += h[i];
  }

  let idx = from;
  for (const want of assign) {
    let page = pages(root)[idx];
    if (!page) {
      page = newPage();
      root.appendChild(page);
    }
    place(pageContent(page), want);
    idx++;
  }

  for (const extra of pages(root).slice(idx)) extra.remove();
  if (pages(root).length === 0) root.appendChild(newPage());

  if (caret) setCaret(caret);
}

/** Cheap synchronous check, run on every input before the browser paints. */
export function paginateIfNeeded(): void {
  const root = docEl();
  const sel = window.getSelection();
  const blk = sel?.focusNode ? nearestBlock(sel.focusNode) : null;
  const list = pages(root);
  let page = (blk?.closest('.page') as HTMLElement | null) ?? null;
  if (!page) page = list[list.length - 1] ?? null;
  if (!page) return;

  const content = pageContent(page);
  const limit = limitH();
  const used = usedHeight(content);

  if (used > limit) {
    paginate({ fromPage: pageIndexOf(page) }); // overflowing
    return;
  }
  const next = page.nextElementSibling;
  if (next && next.classList.contains('page')) {
    const fb = firstBlock(next);
    if (fb && used + heightOf(fb) <= limit) {
      paginate({ fromPage: pageIndexOf(page) }); // room reopened
    }
  }
}

export function pageCount(): number {
  return pages().length;
}
