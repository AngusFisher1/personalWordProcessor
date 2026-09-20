import type {
  Block,
  Doc,
  ParagraphBlock,
  StyleId,
  TableBlock,
  TableCell,
} from './model';
import { isTable, newId } from './model';
import { mediaUrl } from './media';
import { styleClass, styleOf } from './styles';

/* ------------------------------------------------------------------ *
 * Element construction
 * ------------------------------------------------------------------ */

export function docEl(): HTMLElement {
  const el = document.getElementById('doc');
  if (!el) throw new Error('#doc missing');
  return el;
}

export function measureEl(): HTMLElement {
  const el = document.getElementById('measure');
  if (!el) throw new Error('#measure missing');
  return el;
}

export function newPage(): HTMLElement {
  const page = document.createElement('div');
  page.className = 'page';
  const content = document.createElement('div');
  content.className = 'page-content';
  page.appendChild(content);
  return page;
}

export function pageContent(page: Element): HTMLElement {
  let c = page.querySelector(':scope > .page-content') as HTMLElement | null;
  if (!c) {
    c = document.createElement('div');
    c.className = 'page-content';
    page.appendChild(c);
  }
  return c;
}

export function makeBlockEl(b: ParagraphBlock): HTMLElement {
  const el = document.createElement('div');
  el.className = 'blk ' + styleClass(b.styleId);
  el.dataset.blockId = b.id;
  if (b.listMarker !== undefined) el.dataset.marker = b.listMarker;
  if (b.listLevel) el.dataset.level = String(b.listLevel);
  setBlockHtml(el, b.html);
  return el;
}

/** Empty blocks carry a <br> in the DOM (zero-height divs are uneditable) but
 *  are stored as '' in the model. */
export function setBlockHtml(el: HTMLElement, html: string): void {
  el.innerHTML = html.trim() === '' ? '<br>' : html;
  resolveImages(el);
}

/**
 * Give every image a src from the media registry. The model stores the part
 * path, not the bytes, so this is what turns a reference into a picture.
 */
export function resolveImages(root: HTMLElement): void {
  for (const img of Array.from(root.querySelectorAll('img[data-media]'))) {
    const el = img as HTMLImageElement;
    const url = mediaUrl(el.dataset.media as string);
    if (url) {
      el.src = url;
    } else {
      // The part is missing or is a format the browser cannot draw. Leave a
      // box of the right size so the layout does not jump.
      el.removeAttribute('src');
      el.classList.add('img-missing');
    }
    el.draggable = false;
  }
}

export function applyStyle(el: HTMLElement, styleId: StyleId): void {
  const keep = Array.from(el.classList).filter(
    (c) => c !== 'blk' && !c.startsWith('s-')
  );
  el.className = ['blk', styleClass(styleId), ...keep].join(' ');
}

/* ------------------------------------------------------------------ *
 * Tables
 *
 * A cell holds ordinary .blk paragraphs, so styles, the caret and the inline
 * sanitizer all work inside one without knowing tables exist. What the table
 * element adds is the grid around them.
 * ------------------------------------------------------------------ */

export function makeTableEl(t: TableBlock): HTMLElement {
  const wrap = document.createElement('div');
  wrap.className = 'blk-table';
  wrap.dataset.blockId = t.id;

  const table = document.createElement('table');
  // Lay the table out at the width the document says, rather than letting the
  // browser's automatic layout disagree with what Word measured.
  const total = t.cols.reduce((a, c) => a + c.width, 0);
  table.style.width = total > 0 ? total + 'px' : '100%';

  const colgroup = document.createElement('colgroup');
  for (const c of t.cols) {
    const col = document.createElement('col');
    if (c.width > 0) col.style.width = c.width + 'px';
    colgroup.appendChild(col);
  }
  table.appendChild(colgroup);

  const tbody = document.createElement('tbody');
  for (const row of t.rows) {
    tbody.appendChild(makeRowEl(row));
  }
  table.appendChild(tbody);
  wrap.appendChild(table);
  return wrap;
}

export function makeRowEl(row: TableBlock['rows'][number]): HTMLElement {
  const tr = document.createElement('tr');
  tr.dataset.rowId = row.id;
  if (row.headerRow) tr.classList.add('hdr');
  for (let i = 0; i < row.cells.length; i++) {
    const cell = row.cells[i];
    // An empty cell is one that a gridSpan absorbed into the cell before it.
    if (cell.length === 0) continue;
    let span = 1;
    while (i + span < row.cells.length && row.cells[i + span].length === 0) span++;

    const td = document.createElement('td');
    if (span > 1) td.colSpan = span;
    for (const p of cell) td.appendChild(makeBlockEl(p));
    tr.appendChild(td);
  }
  return tr;
}

export function isTableEl(el: Element): boolean {
  return el.classList.contains('blk-table');
}

/** The paragraph or table elements laid out directly on a page. */
export function flowChildren(content: HTMLElement): HTMLElement[] {
  return (Array.from(content.children) as HTMLElement[]).filter(
    (c) => c.classList.contains('blk') || c.classList.contains('blk-table')
  );
}

export function pages(root: HTMLElement = docEl()): HTMLElement[] {
  return Array.from(root.querySelectorAll(':scope > .page')) as HTMLElement[];
}

/**
 * Top-level paragraph elements. Paragraphs inside table cells are deliberately
 * excluded: they are part of a table block, not items in the page flow, and
 * pagination and readModel would double-count them.
 */
export function blocksIn(root: HTMLElement): HTMLElement[] {
  return (Array.from(root.querySelectorAll('.blk')) as HTMLElement[]).filter(
    (el) => !el.closest('.blk-table, .page-header, .page-footer')
  );
}

/** Every paragraph element, including the ones inside table cells. */
export function allParagraphEls(root: HTMLElement): HTMLElement[] {
  return Array.from(root.querySelectorAll('.blk')) as HTMLElement[];
}

export function blockEl(id: string): HTMLElement | null {
  const sel = `[data-block-id="${CSS.escape(id)}"]`;
  return docEl().querySelector(sel) as HTMLElement | null;
}

/* ------------------------------------------------------------------ *
 * Logical blocks
 *
 * Pagination may split one paragraph across several page containers. Those
 * pieces are a render artifact: the first piece keeps the logical block's id,
 * and every continuation carries data-continues-from pointing at it plus
 * data-base, the character offset where that piece starts in the merged text.
 *
 * Everything outside paginate.ts should work in logical blocks, not pieces.
 * ------------------------------------------------------------------ */

export interface Group {
  head: HTMLElement;
  tails: HTMLElement[];
}

export function logicalIdOf(el: HTMLElement): string {
  return el.dataset.continuesFrom || el.dataset.blockId || '';
}

export function isContinuation(el: HTMLElement): boolean {
  return !!el.dataset.continuesFrom;
}

export function pieceBase(el: HTMLElement): number {
  const n = Number(el.dataset.base);
  return Number.isFinite(n) ? n : 0;
}

/** Every piece of one logical block, in document order. */
export function piecesOf(id: string): HTMLElement[] {
  const e = CSS.escape(id);
  return Array.from(
    docEl().querySelectorAll(
      `[data-block-id="${e}"], [data-continues-from="${e}"]`
    )
  ) as HTMLElement[];
}

export function clearSplitMarks(el: HTMLElement): void {
  el.classList.remove('split-cont', 'split-more');
  delete el.dataset.continuesFrom;
  delete el.dataset.base;
}

/**
 * Group the piece elements into logical blocks, in document order. A
 * continuation whose head has gone away - which editing across a page break
 * can produce - is promoted to a block of its own rather than dropped.
 */
export function logicalGroups(root: HTMLElement = docEl()): Group[] {
  const out: Group[] = [];
  const flow: HTMLElement[] = [];
  for (const page of pages(root)) flow.push(...flowChildren(pageContent(page)));
  for (const el of flow) {
    const from = el.dataset.continuesFrom;
    const last = out[out.length - 1];
    if (from && last && last.head.dataset.blockId === from) {
      last.tails.push(el);
      continue;
    }
    if (from) clearSplitMarks(el); // orphaned continuation
    if (!el.dataset.blockId) el.dataset.blockId = newId();
    out.push({ head: el, tails: [] });
  }
  return out;
}

export function logicalHeads(root: HTMLElement = docEl()): HTMLElement[] {
  return logicalGroups(root).map((g) => g.head);
}

/** The logical block a piece belongs to. */
export function headOf(el: HTMLElement): HTMLElement {
  if (!el.dataset.continuesFrom) return el;
  return (blockEl(el.dataset.continuesFrom) as HTMLElement | null) ?? el;
}

function isBrOnly(el: HTMLElement): boolean {
  return el.childNodes.length === 1 && el.firstChild?.nodeName === 'BR';
}

/**
 * Undo a split: fold every continuation back into the head. Pagination does
 * this before measuring so it always starts from one element per paragraph.
 */
export function mergeGroup(g: Group): void {
  if (isTableEl(g.head)) {
    mergeTableGroup(g);
    return;
  }
  for (const t of g.tails) {
    // A piece the user emptied contributes only its placeholder <br>, which
    // would otherwise show up as a spurious line break in the merged text.
    if (!isBrOnly(t)) {
      while (t.firstChild) g.head.appendChild(t.firstChild);
    }
    t.remove();
  }
  g.tails.length = 0;
  clearSplitMarks(g.head);
  if (!g.head.firstChild) g.head.innerHTML = '<br>';
}

/** Fold a split table back together, dropping the repeated header copies. */
function mergeTableGroup(g: Group): void {
  const body = g.head.querySelector('tbody');
  for (const t of g.tails) {
    const tb = t.querySelector('tbody');
    if (body && tb) {
      for (const row of Array.from(tb.children) as HTMLElement[]) {
        if (row.dataset.repeat) continue; // a copy, not a row
        body.appendChild(row);
      }
    }
    t.remove();
  }
  g.tails.length = 0;
  clearSplitMarks(g.head);
}

export function mergeAllGroups(root: HTMLElement = docEl()): void {
  for (const g of logicalGroups(root)) mergeGroup(g);
}

/* ------------------------------------------------------------------ *
 * Render / read back
 * ------------------------------------------------------------------ */

/** Lay every block out on a single page. The caller runs paginate() next. */
export function renderAll(doc: Doc, root: HTMLElement): void {
  renderBlocks(doc.blocks, root);
}

export function renderBlocks(blocks: Block[], root: HTMLElement): void {
  root.textContent = '';
  const page = newPage();
  root.appendChild(page);
  const c = pageContent(page);
  for (const b of blocks) {
    c.appendChild(isTable(b) ? makeTableEl(b) : makeBlockEl(b));
  }
}

/**
 * The DOM is the live source of truth while editing; the model object is a
 * serialization target refreshed from the DOM on save, export and snapshot.
 */
export function readModel(root: HTMLElement): Block[] {
  // Continuation pieces are merged back here, so the stored document always
  // has exactly one block per logical paragraph.
  return logicalGroups(root).map((g) =>
    isTableEl(g.head) ? readTableEl(g) : readParagraphGroup(g)
  );
}

function readParagraphGroup(g: Group): ParagraphBlock {
  let id = g.head.dataset.blockId;
  if (!id) {
    id = newId();
    g.head.dataset.blockId = id;
  }
  const marker = g.head.dataset.marker;
  const level = Number(g.head.dataset.level);
  return {
    id,
    styleId: styleOf(g.head),
    html: mergedHtml(g),
    ...(marker !== undefined ? { listMarker: marker } : {}),
    ...(Number.isFinite(level) && level > 0 ? { listLevel: level } : {}),
  };
}

function readParagraphEl(el: HTMLElement): ParagraphBlock {
  return readParagraphGroup({ head: el, tails: [] });
}

/**
 * Read a table back, folding any continuation pieces of it into one. Rows are
 * keyed by id so a header repeated on a later page is not read twice.
 */
function readTableEl(g: Group): TableBlock {
  const id = g.head.dataset.blockId ?? newId();
  const cols = Array.from(
    g.head.querySelectorAll('col')
  ).map((c) => ({ width: parseFloat((c as HTMLElement).style.width) || 0 }));

  const rows: TableBlock['rows'] = [];
  const seen = new Set<string>();
  for (const part of [g.head, ...g.tails]) {
    for (const tr of Array.from(part.querySelectorAll('tr'))) {
      const rowId = (tr as HTMLElement).dataset.rowId ?? newId();
      if (seen.has(rowId)) continue; // a repeated header
      seen.add(rowId);
      const cells: TableCell[] = [];
      for (const td of Array.from(tr.children) as HTMLElement[]) {
        cells.push(
          (Array.from(td.querySelectorAll(':scope > .blk')) as HTMLElement[]).map(
            readParagraphEl
          )
        );
        // Put back the placeholder entries a colspan stands for, so the row
        // keeps one entry per grid column.
        const span = (td as HTMLTableCellElement).colSpan || 1;
        for (let k = 1; k < span; k++) cells.push([]);
      }
      rows.push({
        id: rowId,
        headerRow: tr.classList.contains('hdr'),
        cells,
      });
    }
  }
  return { kind: 'table', id, cols, rows };
}

function mergedHtml(g: Group): string {
  if (g.tails.length === 0) return readBlockHtml(g.head);
  const joined = document.createElement('div');
  for (const el of [g.head, ...g.tails]) {
    if (el !== g.head && isBrOnly(el)) continue;
    for (const n of Array.from(el.childNodes)) joined.appendChild(n.cloneNode(true));
  }
  return readBlockHtml(joined);
}

export function readBlockHtml(el: HTMLElement): string {
  const html = sanitizeInlineHtml(el.innerHTML, { allowImages: true });
  return html === '<br>' ? '' : html;
}

/* ------------------------------------------------------------------ *
 * Inline sanitizer - shared by readModel and the paste handler
 * ------------------------------------------------------------------ */

/** Removed outright, contents and all. */
const DROP = new Set([
  'SCRIPT', 'STYLE', 'META', 'LINK', 'TITLE', 'NOSCRIPT', 'IFRAME', 'OBJECT',
  'EMBED', 'IMG', 'SVG', 'CANVAS', 'VIDEO', 'AUDIO', 'INPUT', 'BUTTON',
  'SELECT', 'TEXTAREA', 'TABLE',
]);

function safeHref(v: string): boolean {
  const s = v.trim().toLowerCase();
  if (s.startsWith('javascript:') || s.startsWith('data:') || s.startsWith('vbscript:')) {
    return false;
  }
  return true;
}

type Emph = 'b' | 'i' | 'u';

/**
 * What an element says about emphasis, reading its inline style before we
 * throw the style away. Two real-world cases make this necessary:
 *
 *  - Google Docs wraps the whole clipboard payload in
 *    `<b style="font-weight:normal">`. Keeping the <b> and dropping the style
 *    would bold the entire pasted document.
 *  - Google Docs marks actual bold as `<span style="font-weight:700">`.
 *    Unwrapping the span and dropping the style would lose the emphasis.
 *
 * null means "says nothing", so the tag name decides.
 */
function styledEmphasis(el: HTMLElement): Record<Emph, boolean | null> {
  const s = el.style;
  const out: Record<Emph, boolean | null> = { b: null, i: null, u: null };

  const w = s.fontWeight;
  if (w) {
    const n = parseInt(w, 10);
    if (!Number.isNaN(n)) out.b = n >= 600;
    else if (w === 'bold' || w === 'bolder') out.b = true;
    else if (w === 'normal' || w === 'lighter') out.b = false;
  }

  const fs = s.fontStyle;
  if (fs) out.i = fs === 'italic' || fs === 'oblique';

  const td = s.textDecorationLine || s.textDecoration;
  if (td) out.u = td.includes('underline');

  return out;
}

const TAG_EMPHASIS: Record<string, Emph> = {
  B: 'b', STRONG: 'b', I: 'i', EM: 'i', CITE: 'i', VAR: 'i',
  U: 'u', INS: 'u',
};

export interface CleanOptions {
  /**
   * Keep <img> elements. True when reading our own content back, false on
   * paste - a pasted image would arrive as a remote URL or a data blob we
   * have nowhere to put.
   */
  allowImages?: boolean;
}

/** Normalizes a subtree in place down to b/i/u/a/br plus text. */
export function cleanInline(parent: Node, opts: CleanOptions = {}): void {
  for (const n of Array.from(parent.childNodes)) {
    if (n.nodeType === Node.TEXT_NODE) continue;
    if (n.nodeType !== Node.ELEMENT_NODE) {
      parent.removeChild(n);
      continue;
    }
    const el = n as HTMLElement;
    const tag = el.tagName;

    if (tag === 'SPAN' && el.hasAttribute('data-field') && opts.allowImages) {
      // A field token. Its text is a render-time value, so it is emptied on
      // the way to the model - storing "Page 3 of 12" would freeze it.
      for (const a of Array.from(el.attributes)) {
        if (a.name !== 'data-field') el.removeAttribute(a.name);
      }
      el.textContent = '';
      continue;
    }
    if (tag === 'IMG' && opts.allowImages) {
      // An image is described by which part it draws and which preserved run
      // writes it back. The resolved src is a render-time detail.
      for (const a of Array.from(el.attributes)) {
        const keep =
          a.name === 'data-media' ||
          a.name === 'data-run' ||
          a.name === 'width' ||
          a.name === 'height' ||
          a.name === 'alt';
        if (!keep) el.removeAttribute(a.name);
      }
      continue;
    }
    if (DROP.has(tag)) {
      parent.removeChild(el);
      continue;
    }
    if (tag === 'BR') {
      for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
      continue;
    }

    // Decide which emphasis this element carries, from its tag and its style.
    const styled = styledEmphasis(el);
    const wanted: Emph[] = [];
    for (const e of ['b', 'i', 'u'] as Emph[]) {
      const fromTag = TAG_EMPHASIS[tag] === e;
      const explicit = styled[e];
      if (explicit === true || (fromTag && explicit !== false)) wanted.push(e);
    }

    const href =
      tag === 'A' && safeHref(el.getAttribute('href') ?? '')
        ? el.getAttribute('href')
        : null;

    cleanInline(el, opts);

    // Rebuild as nested b/i/u (inside an <a> when there is a link), so that
    // emphasis survives whether it arrived as a tag or as a style.
    let inner: Node = document.createDocumentFragment();
    while (el.firstChild) inner.appendChild(el.firstChild);
    for (const e of wanted) {
      const w = document.createElement(e);
      w.appendChild(inner);
      inner = w;
    }
    if (href) {
      const a = document.createElement('a');
      a.setAttribute('href', href);
      a.appendChild(inner);
      inner = a;
    }
    parent.insertBefore(inner, el);
    parent.removeChild(el);
  }
}

export function sanitizeInlineHtml(html: string, opts: CleanOptions = {}): string {
  const t = document.createElement('template');
  t.innerHTML = html;
  cleanInline(t.content, opts);
  return t.innerHTML;
}
