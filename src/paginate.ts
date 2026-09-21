import type { ParagraphBlock, PageSetup } from './model';
import { contentHeight, contentWidth, newId, pageSetup } from './model';
import { getCaret, setCaret } from './caret';
import type { Group } from './render';
import {
  docEl,
  flowChildren,
  isContinuation,
  isTableEl,
  logicalGroups,
  logicalIdOf,
  makeBlockEl,
  measureEl,
  mergeGroup,
  newPage,
  pageContent,
  pages,
} from './render';
import { STYLES, styleClass, styleOf } from './styles';

/* ------------------------------------------------------------------ *
 * Layout geometry
 * ------------------------------------------------------------------ */

let page: PageSetup = pageSetup('narrow');

/** Write a page geometry onto an element's custom properties. */
function writeGeometry(style: CSSStyleDeclaration, p: PageSetup): void {
  style.setProperty('--page-w', p.width + 'px');
  style.setProperty('--page-h', p.height + 'px');
  style.setProperty('--pad-t', p.margins.top + 'px');
  style.setProperty('--pad-r', p.margins.right + 'px');
  style.setProperty('--pad-b', p.margins.bottom + 'px');
  style.setProperty('--pad-l', p.margins.left + 'px');
  // Substitution happens where a custom property is DECLARED, so the derived
  // pair has to be written here too rather than inherited from :root.
  style.setProperty('--content-w', contentWidth(p) + 'px');
  style.setProperty('--content-h', contentHeight(p) + 'px');
}

/** Apply a document's page geometry to the CSS variables the layout reads. */
export function setPageSetup(p: PageSetup): void {
  page = p;
  writeGeometry(document.documentElement.style, p);
  clearHeightCache();
}

export function currentPageSetup(): PageSetup {
  return page;
}

/* ------------------------------------------------------------------ *
 * Sections
 *
 * A document can have more than one page geometry. Which one applies is a
 * property of position in the document, not of the document as a whole, so
 * every page carries its own.
 *
 * A single-section document - nearly all of them - short-circuits every
 * function here before it touches the DOM. The indexing walk is O(blocks),
 * and the per-keystroke check must not pay for a feature the open document
 * does not use.
 * ------------------------------------------------------------------ */

/** Geometry per section index, in document order. */
let sectionSetups: PageSetup[] = [];
/** Logical block ids that open a section on a fresh page. */
let sectionStarts = new Map<string, number>();
/** Resolved on the last indexing walk. */
let sectionOfEl = new WeakMap<HTMLElement, number>();

export function hasSections(): boolean {
  return sectionSetups.length > 1;
}

/**
 * Tell the paginator about the document's sections.
 *
 * `starts` maps the id of each section's first block to its index. A
 * continuous section is not listed, because it does not begin a page and
 * therefore cannot carry a geometry of its own.
 */
export function setSections(setups: PageSetup[], starts: Map<string, number>): void {
  sectionSetups = setups;
  sectionStarts = starts;
  sectionOfEl = new WeakMap();
  clearHeightCache();
}

/**
 * Walk the blocks in order and record which section each is in.
 *
 * Positional rather than looked up from the model, so a paragraph typed into
 * the middle of section two is in section two - the model does not know
 * about it yet, and will not until the next save.
 */
function indexSections(root: HTMLElement): void {
  if (!hasSections()) return;
  sectionOfEl = new WeakMap();
  let current = 0;
  for (const g of logicalGroups(root)) {
    const at = sectionStarts.get(logicalIdOf(g.head));
    if (at !== undefined) current = at;
    sectionOfEl.set(g.head, current);
    for (const t of g.tails) sectionOfEl.set(t, current);
  }
}

function setupOfSection(i: number): PageSetup {
  return sectionSetups[i] ?? page;
}

function sectionOfBlock(el: HTMLElement | null | undefined): number {
  if (!hasSections() || !el) return 0;
  const known = sectionOfEl.get(el);
  if (known !== undefined) return known;
  // A block created since the last walk: it belongs where its page does.
  const pageEl = el.closest('.page') as HTMLElement | null;
  return Number(pageEl?.dataset.section ?? 0) || 0;
}

function setupOfBlock(el: HTMLElement | null | undefined): PageSetup {
  if (!hasSections()) return page;
  return setupOfSection(sectionOfBlock(el));
}

/** The geometry a page on screen was laid out with. */
function setupOfPage(pageEl: Element | null | undefined): PageSetup {
  if (!hasSections()) return page;
  const raw = (pageEl as HTMLElement | null)?.dataset.section;
  return setupOfSection(Number(raw) || 0);
}

/**
 * Space available for body text on one page.
 *
 * NOT a constant. A header and footer take their height out of it, which ones
 * a page uses depends on its number, and the content box itself depends on
 * the section - so every comparison has to go through here. A stray
 * hardcoded value produces a one-page-off error that only shows up in long
 * documents.
 */
function limitFor(pageIndex: number, setup: PageSetup = page): number {
  return contentHeight(setup) - hfTaken(pageIndex);
}

/** Set by the caller, since the paginator does not own the document. */
let hfTaken: (pageIndex: number) => number = () => 0;

export function setHeaderFooterSpace(fn: (pageIndex: number) => number): void {
  hfTaken = fn;
}

/** How far a keep-with-next run may cascade before we let it break. */
const KEEP_CASCADE_MAX = 3;

/* ------------------------------------------------------------------ *
 * Measurement
 *
 * Cached by (content width, style, split state, html) rather than by block id:
 * identical blocks share an entry, and editing one never invalidates another.
 * The split state is part of the key because a split piece suppresses padding
 * and would otherwise collide with the same text rendered whole.
 * ------------------------------------------------------------------ */

interface LineInfo {
  /** Border-box height, sub-pixel. */
  height: number;
  /** Splittable units: lines for a paragraph, rows for a table. */
  lineCount: number;
  /** Used line box height. Uniform, so paragraphs need no per-unit array. */
  lineH: number;
  /** Per-unit heights, for tables, whose rows are all different. */
  unitHeights?: number[];
  /** Leading rows to repeat at the top of each continuation. */
  headerRows?: number;
  headerHeight?: number;
  isTable?: boolean;
  /** Top padding plus top border. */
  padTop: number;
  /** Bottom padding plus bottom border. */
  padBottom: number;
  /** Top of each line's ink box, relative to the element's border-box top. */
  inkTops: number[];
  /** Rendered character count. */
  textLen: number;
  /** Character offset where each line begins; filled in on demand. */
  starts: Map<number, number>;
}

const infos = new Map<string, LineInfo>();
const CACHE_CAP = 4000;

export function clearHeightCache(): void {
  infos.clear();
}

function keyFor(el: HTMLElement): string {
  const split =
    (el.classList.contains('split-cont') ? 'c' : '') +
    (el.classList.contains('split-more') ? 'm' : '');
  // Keyed on the width the block is actually laid out at, which is its
  // section's, not the document's: two sections with different margins wrap
  // the same sentence at different points. The inline style is in the key
  // too, because direct formatting - an indent, a paragraph spacing - changes
  // the height of text that is otherwise identical.
  return (
    contentWidth(setupOfBlock(el)) +
    '|' + styleOf(el) +
    '|' + split +
    '|' + (el.dataset.fmt ?? '') +
    '|' + el.innerHTML
  );
}

/**
 * Line geometry for a block.
 *
 * getClientRects() gives one rect per text fragment, so `<b>bold</b> rest`
 * yields several rects on one line; they are merged back by vertical overlap.
 * Those rects are INK boxes - the glyph extent - not line boxes, so they are
 * used only to count lines and to locate them for the binary search. The
 * heights come from the used line box height, derived by dividing the block's
 * content height by its line count, because that is what actually stacks.
 *
 * This assumes a uniform line height within a block, which holds for every
 * named style: inline b/i/u do not change the font size.
 */
function measureLines(el: HTMLElement): LineInfo {
  const box = el.getBoundingClientRect();
  if (isTableEl(el)) {
    // A table's units are its rows, and unlike lines they are all different
    // heights, so they are measured individually.
    const h = box.height || measureClone(el);
    const cs0 = getComputedStyle(el);
    const rows = (
      Array.from(el.querySelectorAll('tbody > tr')) as HTMLElement[]
    ).filter((r) => !r.dataset.repeat);
    const unitHeights = rows.map((r) => r.getBoundingClientRect().height);
    // Only the leading run of header rows repeats.
    let headerRows = 0;
    while (headerRows < rows.length && rows[headerRows].classList.contains('hdr')) {
      headerRows++;
    }
    let headerHeight = 0;
    for (let i = 0; i < headerRows; i++) headerHeight += unitHeights[i];
    return {
      height: h,
      lineCount: Math.max(1, rows.length),
      lineH: rows.length ? h / rows.length : h,
      unitHeights,
      headerRows,
      headerHeight,
      isTable: true,
      padTop: parseFloat(cs0.paddingTop) || 0,
      padBottom: parseFloat(cs0.paddingBottom) || 0,
      inkTops: [],
      textLen: 0,
      starts: new Map(),
    };
  }
  const height = box.height || measureClone(el);
  const cs = getComputedStyle(el);
  const padTop = parseFloat(cs.paddingTop) + parseFloat(cs.borderTopWidth);
  const padBottom =
    parseFloat(cs.paddingBottom) + parseFloat(cs.borderBottomWidth);

  const r = document.createRange();
  r.selectNodeContents(el);
  const textLen = r.toString().length;

  const lines: { top: number; bottom: number }[] = [];
  for (const rc of Array.from(r.getClientRects())) {
    if (rc.height <= 0.5) continue;
    const last = lines[lines.length - 1];
    if (last && rc.top < last.bottom - 1) {
      last.bottom = Math.max(last.bottom, rc.bottom);
    } else {
      lines.push({ top: rc.top, bottom: rc.bottom });
    }
  }

  const lineCount = Math.max(1, lines.length);
  const inner = Math.max(0, height - padTop - padBottom);
  return {
    height,
    lineCount,
    lineH: inner / lineCount,
    padTop,
    padBottom,
    inkTops: lines.map((l) => l.top - box.top),
    textLen,
    starts: new Map(),
  };
}

function infoOf(el: HTMLElement): LineInfo {
  const k = keyFor(el);
  const hit = infos.get(k);
  if (hit !== undefined) return hit;
  const info = measureLines(el);
  if (infos.size > CACHE_CAP) infos.clear();
  infos.set(k, info);
  return info;
}

function heightOf(el: HTMLElement): number {
  return infoOf(el).height;
}

function measureClone(el: HTMLElement): number {
  const m = measureEl();
  m.style.width = contentWidth(setupOfBlock(el)) + 'px';
  const clone = el.cloneNode(true) as HTMLElement;
  m.appendChild(clone);
  const h = clone.offsetHeight;
  clone.remove();
  return h;
}

/** Measure model blocks without rendering them into the document. */
export function measureHeights(blocks: ParagraphBlock[], contentW: number): number[] {
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
 * Character offset of a line start
 * ------------------------------------------------------------------ */

function locate(el: HTMLElement, offset: number): { node: Node; off: number } {
  const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let seen = 0;
  let n: Node | null;
  while ((n = w.nextNode())) {
    const len = n.textContent?.length ?? 0;
    if (seen + len >= offset) return { node: n, off: offset - seen };
    seen += len;
  }
  return { node: el, off: el.childNodes.length };
}

function prefixRange(el: HTMLElement, o: number): Range {
  const r = document.createRange();
  r.setStart(el, 0);
  const pos = locate(el, o);
  if (pos.node === el) r.setEnd(el, el.childNodes.length);
  else r.setEnd(pos.node, pos.off);
  return r;
}

/** Which line the last character of el[0..o) sits on. */
function lineOfPrefix(
  el: HTMLElement,
  o: number,
  info: LineInfo,
  elTop: number
): number {
  const rects = Array.from(prefixRange(el, o).getClientRects()).filter(
    (x) => x.height > 0.5
  );
  if (rects.length === 0) return 0;
  const y = rects[rects.length - 1].top - elTop;
  for (let i = info.inkTops.length - 1; i >= 1; i--) {
    if (y >= info.inkTops[i] - 1) return i;
  }
  return 0;
}

/**
 * Character offset at which line `L` begins, by binary search: lineOfPrefix is
 * non-decreasing in the offset, so the first offset whose prefix reaches line L
 * is one past the first character of that line.
 */
function lineStartOffset(
  el: HTMLElement,
  info: LineInfo,
  L: number,
  elTop: number
): number {
  const cached = info.starts.get(L);
  if (cached !== undefined) return cached;

  let lo = 1;
  let hi = info.textLen;
  let ans = info.textLen;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (lineOfPrefix(el, mid, info, elTop) >= L) {
      ans = mid;
      hi = mid - 1;
    } else {
      lo = mid + 1;
    }
  }
  const out = Math.max(0, ans - 1);
  info.starts.set(L, out);
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

/** A legitimate direct child of .page-content: a paragraph or a table. */
function isFlow(n: Node): n is HTMLElement {
  if (n.nodeType !== Node.ELEMENT_NODE) return false;
  const el = n as HTMLElement;
  return el.classList.contains('blk') || el.classList.contains('blk-table');
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
      // Paragraphs inside a table cell belong there.
      if (nested.closest('.blk-table')) continue;
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
      if (isFlow(n)) {
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
      if (b.classList.contains('blk') && b.childNodes.length === 0) {
        b.innerHTML = '<br>';
      }
    }
  }
}

/** Keep one trailing empty Body block so clicking below the last line works. */
export function ensureTrailingBlock(): boolean {
  const root = docEl();
  const groups = logicalGroups(root);
  const last = groups[groups.length - 1];
  if (last) {
    const text = [last.head, ...last.tails]
      .map((e) => e.textContent ?? '')
      .join('');
    if (text.trim() === '') return false;
  }
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
 * Page assignment
 * ------------------------------------------------------------------ */

interface Piece {
  g: Group;
  info: LineInfo;
  /** Line range [from, to) of the logical block that this piece renders. */
  from: number;
  to: number;
  el?: HTMLElement;
}

function lineCountOf(info: LineInfo): number {
  return info.lineCount;
}

/** Rendered height of the piece covering units [from, to). */
function pieceHeight(info: LineInfo, from: number, to: number): number {
  const end = Math.min(to, info.lineCount);
  let h = 0;
  if (info.unitHeights) {
    for (let i = from; i < end; i++) h += info.unitHeights[i];
    // A continuation carries a copy of the header rows.
    if (from > 0 && from >= (info.headerRows ?? 0)) h += info.headerHeight ?? 0;
  } else {
    h = Math.max(0, end - from) * info.lineH;
  }
  if (from === 0) h += info.padTop; // only the first piece keeps top padding
  if (to >= info.lineCount) h += info.padBottom; // only the last keeps bottom
  return h;
}

/**
 * Smallest height at which this block could legally put anything on the
 * current page: its first orphanMin lines if it may be split, otherwise the
 * whole thing. Recorded per page as `needh` so the cheap check knows exactly
 * how much room would have to reopen before the layout could change.
 */
function minPlaceable(
  info: LineInfo,
  from: number,
  def: { keepLines: boolean; orphanMin: number; widowMin: number }
): number {
  const lines = lineCountOf(info);
  const whole = pieceHeight(info, from, lines);
  if (def.keepLines) return whole;
  if (lines - (from + def.orphanMin) < def.widowMin) return whole;
  return pieceHeight(info, from, from + def.orphanMin);
}

/** Largest line index (exclusive) that still fits in `remaining`. */
function linesThatFit(info: LineInfo, from: number, remaining: number): number {
  if (info.unitHeights) {
    let used =
      from === 0
        ? info.padTop
        : from >= (info.headerRows ?? 0)
          ? info.headerHeight ?? 0
          : 0;
    let n = from;
    while (n < info.lineCount && used + info.unitHeights[n] <= remaining) {
      used += info.unitHeights[n];
      n++;
    }
    return n;
  }
  const lead = from === 0 ? info.padTop : 0;
  const fits = Math.floor((remaining - lead) / info.lineH + 1e-6);
  return Math.max(from, Math.min(info.lineCount, from + Math.max(0, fits)));
}

function topBlocks(page: Element): HTMLElement[] {
  return flowChildren(pageContent(page));
}

function firstBlock(page: Element): HTMLElement | null {
  return topBlocks(page)[0] ?? null;
}

/**
 * Height actually used by a page.
 *
 * NOT scrollHeight: .page-content has a fixed height, so scrollHeight never
 * reports less than the full content box and the "room reopened" test could
 * never fire.
 */
function usedHeight(content: HTMLElement): number {
  let h = 0;
  for (const c of Array.from(content.children) as HTMLElement[]) h += heightOf(c);
  return h;
}

function assign(
  groups: Group[],
  firstPage: number
): { pages: Piece[][]; need: number[] } {
  const out: Piece[][] = [[]];
  const need: number[] = [];
  let running = 0;
  let gi = 0;
  let from = 0;
  let pulledFor = -1;

  const cur = () => out[out.length - 1];
  /**
   * The section this page renders: the one its first block belongs to, or,
   * on a page with nothing on it yet, the one the next block belongs to.
   * A page cannot straddle two geometries, so the first block on it decides.
   */
  const setupNow = () =>
    setupOfBlock(cur()[0]?.g.head ?? groups[gi]?.head);
  // The limit follows the page being filled, not the document.
  const limitNow = () => limitFor(firstPage + out.length - 1, setupNow());
  const breakPage = (needed: number) => {
    need[out.length - 1] = needed;
    out.push([]);
    running = 0;
  };

  while (gi < groups.length) {
    const g = groups[gi];
    const def = STYLES[styleOf(g.head)];
    const info = infoOf(g.head);
    const lines = lineCountOf(info);

    const opensSection =
      from === 0 && hasSections() && sectionStarts.has(logicalIdOf(g.head));
    if (from === 0 && (def.pageBreakBefore || opensSection) && cur().length > 0) {
      breakPage(0);
      continue;
    }

    const limit = limitNow();
    const restH = pieceHeight(info, from, lines);
    if (restH <= limit - running) {
      cur().push({ g, info, from, to: lines });
      running += restH;
      gi++;
      from = 0;
      continue;
    }

    // Does not fit. Try to split it.
    const maxFit = linesThatFit(info, from, limit - running);
    // Leave at least one unit for the next page either way.
    const widow = info.isTable ? 1 : def.widowMin;
    // If taking every unit that fits would leave a widow, back off to the
    // latest split that does not, rather than abandoning the split entirely.
    const fit = Math.min(maxFit, lines - widow);

    let splittable: boolean;
    if (info.isTable) {
      // A repeated header is not content, so a page holding only the header
      // and one row has stranded that row. Two real rows, or move it whole.
      const headers = info.headerRows ?? 0;
      const rowsHere = fit - from - (from === 0 ? headers : 0);
      splittable = info.lineCount > 1 && rowsHere >= 2 && lines - fit >= 1;
    } else {
      splittable =
        !def.keepLines &&
        info.lineCount > 1 &&
        info.textLen > 0 &&
        fit - from >= def.orphanMin;
    }

    if (splittable) {
      cur().push({ g, info, from, to: fit });
      from = fit;
      // If the split was capped by widow control, one more line still cannot
      // go here: only room for the whole remainder would change the layout.
      breakPage(fit < maxFit ? pieceHeight(info, fit, lines) : info.lineH);
      continue;
    }

    if (cur().length === 0) {
      // An empty page and it still will not fit: place it and let it clip
      // rather than spin. A block taller than a page is a v1 known limit.
      cur().push({ g, info, from, to: lines });
      running += restH;
      gi++;
      from = 0;
      continue;
    }

    // keep-with-next: drag the trailing run of headings down with this block.
    if (pulledFor !== gi) {
      pulledFor = gi;
      const pulled: Piece[] = [];
      while (cur().length > 1 && pulled.length < KEEP_CASCADE_MAX) {
        const last = cur()[cur().length - 1];
        if (last.to < lineCountOf(last.info)) break; // a split piece cannot move
        if (!STYLES[styleOf(last.g.head)].keepWithNext) break;
        pulled.unshift(cur().pop() as Piece);
      }
      // The page we are closing can only take content again if there is room
      // for the whole run we just pulled off it AND for something of the block
      // that displaced it - anything less rebuilds the identical layout.
      const pulledH = pulled.reduce(
        (a, p) => a + pieceHeight(p.info, p.from, p.to),
        0
      );
      breakPage(pulledH + minPlaceable(info, from, def));
      for (const p of pulled) {
        cur().push(p);
        running += pieceHeight(p.info, p.from, p.to);
      }
      continue;
    }

    breakPage(minPlaceable(info, from, def));
  }

  need[out.length - 1] = 0;
  return { pages: out, need };
}

/* ------------------------------------------------------------------ *
 * Applying an assignment to the DOM
 * ------------------------------------------------------------------ */

/** Split `src` at a character offset, returning the new continuation element. */
function splitOffElement(src: HTMLElement, offset: number, id: string): HTMLElement {
  const pos = locate(src, offset);
  const r = document.createRange();
  if (pos.node === src) r.setStart(src, pos.off);
  else r.setStart(pos.node, pos.off);
  r.setEnd(src, src.childNodes.length);

  const tail = document.createElement('div');
  tail.className = src.className;
  tail.dataset.blockId = newId();
  tail.dataset.continuesFrom = id;
  tail.appendChild(r.extractContents());
  if (!tail.firstChild) tail.innerHTML = '<br>';
  src.parentElement?.insertBefore(tail, src.nextSibling);
  return tail;
}

/**
 * Split a table after `fromRow`, returning the continuation element.
 *
 * Rows are MOVED into the new table, and the leading header rows are cloned
 * on top of it. The clones are marked so measurement, readModel and the caret
 * all know to skip them: they are a render artifact, not extra rows.
 */
function splitTableElement(
  src: HTMLElement,
  fromRow: number,
  headerEls: HTMLElement[],
  id: string
): HTMLElement | null {
  const srcTable = src.querySelector('table');
  const srcBody = src.querySelector('tbody');
  if (!srcTable || !srcBody) return null;

  const wrap = document.createElement('div');
  wrap.className = src.className;
  wrap.dataset.blockId = newId();
  wrap.dataset.continuesFrom = id;
  wrap.classList.add('split-cont');

  const table = document.createElement('table');
  table.style.width = srcTable.style.width;
  const colgroup = srcTable.querySelector('colgroup');
  if (colgroup) table.appendChild(colgroup.cloneNode(true));
  const body = document.createElement('tbody');

  const rows = (Array.from(srcBody.children) as HTMLElement[]).filter(
    (r) => !r.dataset.repeat
  );
  if (fromRow <= 0 || fromRow >= rows.length) return null;

  // The rows to repeat come from the ORIGINAL table. Taking them from the
  // piece being split would copy its first data row on the third page,
  // because a continuation's own header is already a marked copy.
  for (const header of headerEls) {
    const clone = header.cloneNode(true) as HTMLElement;
    clone.dataset.repeat = '1';
    clone.setAttribute('contenteditable', 'false');
    // Ids must stay unique; the original row keeps them.
    for (const el of Array.from(clone.querySelectorAll('[data-block-id]'))) {
      (el as HTMLElement).removeAttribute('data-block-id');
    }
    body.appendChild(clone);
  }
  for (let i = fromRow; i < rows.length; i++) body.appendChild(rows[i]);

  table.appendChild(body);
  wrap.appendChild(table);
  src.parentElement?.insertBefore(wrap, src.nextSibling);
  return wrap;
}

/**
 * Give every logical block the pieces its assignment calls for. The head keeps
 * its element identity - only continuations are created and destroyed - so the
 * caret's own element usually survives a reflow untouched.
 */
function materialize(assigned: Piece[][]): void {
  const byGroup = new Map<Group, Piece[]>();
  for (const page of assigned) {
    for (const p of page) {
      const list = byGroup.get(p.g);
      if (list) list.push(p);
      else byGroup.set(p.g, [p]);
    }
  }

  for (const [g, pieces] of byGroup) {
    const head = g.head;
    if (pieces.length === 1) {
      pieces[0].el = head;
      continue;
    }

    const info = pieces[0].info;
    const id = head.dataset.blockId as string;

    if (info.isTable) {
      head.classList.add('split-more');
      pieces[0].el = head;
      const headerEls = (
        Array.from(head.querySelectorAll('tbody > tr')) as HTMLElement[]
      )
        .filter((r) => !r.dataset.repeat)
        .slice(0, info.headerRows ?? 0);
      let src = head;
      let ok = true;
      for (let i = 1; i < pieces.length; i++) {
        // Each continuation splits what is left, so the row index is relative
        // to the piece being cut, not to the original table.
        const cutAt = pieces[i].from - pieces[i - 1].from;
        const tail = splitTableElement(src, cutAt, headerEls, id);
        if (!tail) {
          ok = false;
          break;
        }
        tail.dataset.base = String(pieces[i].from);
        if (i < pieces.length - 1) tail.classList.add('split-more');
        pieces[i].el = tail;
        src = tail;
      }
      if (!ok) for (const p of pieces) p.el = p.el ?? head;
      continue;
    }

    const elTop = head.getBoundingClientRect().top;

    // Compute every boundary against the merged element before cutting it.
    const offsets: number[] = [];
    let ok = true;
    let prev = 0;
    for (let i = 1; i < pieces.length; i++) {
      const abs = lineStartOffset(head, info, pieces[i].from, elTop);
      if (abs <= prev || abs >= info.textLen) {
        ok = false;
        break;
      }
      offsets.push(abs);
      prev = abs;
    }
    if (!ok) {
      // Degenerate split point; render the block whole and let it overflow
      // rather than produce an empty piece.
      for (const p of pieces) p.el = head;
      continue;
    }

    head.classList.add('split-more');
    pieces[0].el = head;
    let src = head;
    let consumed = 0;
    for (let i = 0; i < offsets.length; i++) {
      const tail = splitOffElement(src, offsets[i] - consumed, id);
      tail.dataset.base = String(offsets[i]);
      tail.classList.add('split-cont');
      if (i < offsets.length - 1) tail.classList.add('split-more');
      else tail.classList.remove('split-more');
      consumed = offsets[i];
      pieces[i + 1].el = tail;
      src = tail;
    }
  }
}

/**
 * Reconcile a page's children against the elements assigned to it. Blocks are
 * MOVED with insertBefore, never re-created: moving preserves focus, avoids
 * flicker and keeps IME composition alive.
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

/* ------------------------------------------------------------------ *
 * Reflow
 * ------------------------------------------------------------------ */

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
  let from = Math.max(0, Math.min(opts?.fromPage ?? 0, before.length - 1));
  // A page that opens with a continuation cannot be reflowed on its own: the
  // paragraph it continues starts earlier.
  while (from > 0) {
    const fb = firstBlock(before[from]);
    if (fb && isContinuation(fb)) from--;
    else break;
  }

  indexSections(root);
  const all = logicalGroups(root);
  const anchor = firstBlock(before[from]);
  let start = anchor ? all.findIndex((g) => g.head === anchor) : 0;
  if (start < 0) start = 0;
  const groups = all.slice(start);

  // Work from one element per paragraph, always.
  for (const g of groups) mergeGroup(g);

  const assigned = assign(groups, from);
  materialize(assigned.pages);

  let idx = from;
  for (const pieces of assigned.pages) {
    let page = pages(root)[idx];
    if (!page) {
      page = newPage();
      root.appendChild(page);
    }
    place(
      pageContent(page),
      pieces.map((p) => p.el).filter((e): e is HTMLElement => !!e)
    );
    // What the next content needs in order to also fit here. The cheap check
    // reads this instead of guessing.
    page.dataset.needh = String(Math.max(0, assigned.need[idx - from] ?? 0));
    // A page carries its own geometry, so a document with two sections
    // renders two page sizes without the layout consulting anything global.
    const si = sectionOfBlock(pieces[0]?.g.head);
    if (String(si) !== page.dataset.section) {
      page.dataset.section = String(si);
      writeGeometry(page.style, setupOfSection(si));
    }
    idx++;
  }

  for (const extra of pages(root).slice(idx)) extra.remove();
  if (pages(root).length === 0) root.appendChild(newPage());

  if (caret) setCaret(caret);
}

/** Cheap synchronous check, run on every input event before the browser paints. */
export function paginateIfNeeded(): void {
  const root = docEl();
  const sel = window.getSelection();
  const blk = sel?.focusNode ? nearestBlockEl(sel.focusNode) : null;
  const list = pages(root);
  let page = (blk?.closest('.page') as HTMLElement | null) ?? null;
  if (!page) page = list[list.length - 1] ?? null;
  if (!page) return;

  const content = pageContent(page);
  const limit = limitFor(pages(root).indexOf(page), setupOfPage(page));
  const used = usedHeight(content);

  if (used > limit) {
    paginate({ fromPage: pages(root).indexOf(page) }); // overflowing
    return;
  }
  const need = Number(page.dataset.needh);
  if (Number.isFinite(need) && need > 0 && used + need <= limit) {
    paginate({ fromPage: pages(root).indexOf(page) }); // room reopened
  }
}

function nearestBlockEl(node: Node): HTMLElement | null {
  let n: Node | null = node;
  while (n) {
    if (
      n.nodeType === Node.ELEMENT_NODE &&
      (n as HTMLElement).classList.contains('blk')
    ) {
      return n as HTMLElement;
    }
    n = n.parentNode;
  }
  return null;
}

export function pageCount(): number {
  return pages().length;
}

/** Diagnostics for the acceptance checks: slack left at the foot of each page. */
export function pageSlack(): number[] {
  return pages().map((p, i) => limitFor(i, setupOfPage(p)) - usedHeight(pageContent(p)));
}

