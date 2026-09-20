import { newId } from './model';
import { caretAtEnd, caretAtStart, nearestBlock } from './caret';
import { docEl } from './render';
import { styleClass } from './styles';
import { openMenuAt } from './ui';

/**
 * Editing inside a table.
 *
 * Cells already hold ordinary paragraphs, so typing, styles and the caret
 * need nothing here. What is missing is the grid itself: moving between
 * cells, resizing columns, and adding or removing rows and columns.
 *
 * Every operation works on the DOM and lets readModel pick the result up,
 * the same way editing text does - there is no second copy of the table to
 * keep in step.
 */

export interface TableHooks {
  /** Called after a structural change, to reflow, snapshot and save. */
  onChanged(): void;
}

let hooks: TableHooks | null = null;

/** How close to a column border counts as grabbing it. */
const GRAB = 4;

function cellOf(node: Node | null): HTMLTableCellElement | null {
  let n: Node | null = node;
  while (n) {
    if (n.nodeType === Node.ELEMENT_NODE && (n as Element).tagName === 'TD') {
      return n as HTMLTableCellElement;
    }
    n = n.parentNode;
  }
  return null;
}

function caretCell(): HTMLTableCellElement | null {
  const sel = window.getSelection();
  return sel?.focusNode ? cellOf(sel.focusNode) : null;
}

function cellsOf(table: HTMLElement): HTMLTableCellElement[] {
  return Array.from(table.querySelectorAll('tbody > tr > td'));
}

function firstBlockIn(td: HTMLElement): HTMLElement | null {
  return td.querySelector(':scope > .blk');
}

function emptyBlock(): HTMLElement {
  const el = document.createElement('div');
  el.className = 'blk ' + styleClass('Body');
  el.dataset.blockId = newId();
  el.innerHTML = '<br>';
  return el;
}

/** The whole table element a cell belongs to, including split continuations. */
function wrapOf(td: HTMLElement): HTMLElement | null {
  return td.closest('.blk-table');
}

/* ------------------------------------------------------------------ *
 * Moving between cells
 * ------------------------------------------------------------------ */

function appendRow(after: HTMLTableRowElement): HTMLTableRowElement {
  const tr = document.createElement('tr');
  tr.dataset.rowId = newId();
  for (let i = 0; i < after.cells.length; i++) {
    const td = document.createElement('td');
    if (after.cells[i].colSpan > 1) td.colSpan = after.cells[i].colSpan;
    td.appendChild(emptyBlock());
    tr.appendChild(td);
  }
  after.parentElement?.insertBefore(tr, after.nextSibling);
  return tr;
}

/**
 * Tab moves to the next cell, and from the last cell it adds a row - which
 * is how every table in every word processor grows.
 */
export function tabInTable(back: boolean): boolean {
  const td = caretCell();
  if (!td) return false;
  const wrap = wrapOf(td);
  if (!wrap) return false;

  const cells = cellsOf(wrap);
  const i = cells.indexOf(td);
  if (i < 0) return false;

  let target = cells[i + (back ? -1 : 1)];
  if (!target && !back) {
    const row = td.closest('tr') as HTMLTableRowElement | null;
    if (!row) return false;
    target = appendRow(row).cells[0];
    hooks?.onChanged();
  }
  if (!target) return false;

  const blk = firstBlockIn(target) ?? target.appendChild(emptyBlock());
  docEl().focus();
  if (back) caretAtEnd(blk as HTMLElement);
  else caretAtStart(blk as HTMLElement);
  return true;
}

/**
 * Up and down leave a cell only when the caret is already on the first or
 * last line of it, so arrowing through a paragraph inside a cell still
 * works normally.
 */
export function arrowInTable(dir: -1 | 1): boolean {
  const td = caretCell();
  if (!td) return false;
  const blk = nearestBlock(window.getSelection()?.focusNode ?? null);
  if (!blk) return false;

  const sel = window.getSelection();
  if (!sel || !sel.isCollapsed) return false;

  // Compare against the cell's own content, not its border box. A caret on
  // the first line still sits below the cell padding and the line's half
  // leading, so measuring to the border never reports the edge.
  const rows2 = Array.from(td.querySelectorAll(':scope > .blk')) as HTMLElement[];
  if (rows2.length === 0) return false;
  const firstR = rows2[0].getBoundingClientRect();
  const lastR = rows2[rows2.length - 1].getBoundingClientRect();
  const lineH = parseFloat(getComputedStyle(blk).lineHeight) || 16;

  const caret = sel.getRangeAt(0).getBoundingClientRect();
  // A collapsed range often has no rect at all; fall back to the block.
  const mid =
    caret.height > 0
      ? (caret.top + caret.bottom) / 2
      : blk.getBoundingClientRect().top + lineH / 2;

  const atTop = mid <= firstR.top + lineH;
  const atBottom = mid >= lastR.bottom - lineH;
  if (dir < 0 && !atTop) return false;
  if (dir > 0 && !atBottom) return false;

  const row = td.closest('tr') as HTMLTableRowElement | null;
  const body = row?.parentElement;
  if (!row || !body) return false;
  const rows = Array.from(body.children) as HTMLTableRowElement[];
  const ri = rows.indexOf(row);
  const ci = Array.from(row.cells).indexOf(td);
  const next = rows[ri + dir];
  if (!next) return false;
  const target = next.cells[Math.min(ci, next.cells.length - 1)];
  if (!target) return false;

  const blk2 = firstBlockIn(target);
  if (!blk2) return false;
  docEl().focus();
  if (dir < 0) caretAtEnd(blk2);
  else caretAtStart(blk2);
  return true;
}

/* ------------------------------------------------------------------ *
 * Column widths
 * ------------------------------------------------------------------ */

function colsOf(wrap: HTMLElement): HTMLTableColElement[] {
  return Array.from(wrap.querySelectorAll('colgroup > col'));
}

function widths(wrap: HTMLElement): number[] {
  const table = wrap.querySelector('table') as HTMLElement | null;
  const cols = colsOf(wrap);
  const total = table?.getBoundingClientRect().width ?? 0;
  const declared = cols.map((c) => parseFloat(c.style.width) || 0);
  const sum = declared.reduce((a, b) => a + b, 0);
  if (sum > 0) return declared;
  // No declared widths: fall back to what the browser worked out.
  const cells = (wrap.querySelector('tbody > tr') as HTMLTableRowElement | null)?.cells;
  if (!cells) return declared;
  return Array.from(cells).map((c) => c.getBoundingClientRect().width || total / cells.length);
}

/** Which column border, if any, the pointer is over. */
function borderAt(wrap: HTMLElement, clientX: number): number {
  const table = wrap.querySelector('table');
  if (!table) return -1;
  const left = table.getBoundingClientRect().left;
  const w = widths(wrap);
  let x = left;
  for (let i = 0; i < w.length - 1; i++) {
    x += w[i];
    if (Math.abs(clientX - x) <= GRAB) return i;
  }
  return -1;
}

function applyWidths(wrap: HTMLElement, w: number[]): void {
  const cols = colsOf(wrap);
  // Whole pixels, so the columns keep summing to the table width.
  const rounded = w.map((n) => Math.max(16, Math.round(n)));
  cols.forEach((c, i) => {
    if (rounded[i] !== undefined) c.style.width = rounded[i] + 'px';
  });
  const table = wrap.querySelector('table') as HTMLElement | null;
  if (table) table.style.width = rounded.reduce((a, b) => a + b, 0) + 'px';
}

let dragging: {
  wrap: HTMLElement;
  index: number;
  startX: number;
  start: number[];
} | null = null;

function onMove(e: MouseEvent): void {
  if (!dragging) return;
  const { wrap, index, startX, start } = dragging;
  const dx = e.clientX - startX;
  const next = start.slice();
  // Take from the column on the right, so the table's total width holds.
  const min = 16;
  const delta = Math.max(min - start[index], Math.min(start[index + 1] - min, dx));
  next[index] = start[index] + delta;
  next[index + 1] = start[index + 1] - delta;
  applyWidths(wrap, next);
  e.preventDefault();
}

function onUp(): void {
  if (!dragging) return;
  dragging = null;
  document.removeEventListener('mousemove', onMove, true);
  document.removeEventListener('mouseup', onUp, true);
  document.body.classList.remove('col-resizing');
  hooks?.onChanged();
}

/* ------------------------------------------------------------------ *
 * Rows and columns
 * ------------------------------------------------------------------ */

function hasMerges(wrap: HTMLElement): boolean {
  return cellsOf(wrap).some((c) => c.colSpan > 1 || c.rowSpan > 1);
}

function insertRow(td: HTMLTableCellElement, below: boolean): void {
  const row = td.closest('tr') as HTMLTableRowElement | null;
  if (!row) return;
  const tr = document.createElement('tr');
  tr.dataset.rowId = newId();
  for (let i = 0; i < row.cells.length; i++) {
    const cell = document.createElement('td');
    if (row.cells[i].colSpan > 1) cell.colSpan = row.cells[i].colSpan;
    cell.appendChild(emptyBlock());
    tr.appendChild(cell);
  }
  row.parentElement?.insertBefore(tr, below ? row.nextSibling : row);
  hooks?.onChanged();
}

function deleteRow(td: HTMLTableCellElement): void {
  const row = td.closest('tr') as HTMLTableRowElement | null;
  const body = row?.parentElement;
  if (!row || !body) return;
  if (body.children.length <= 1) return; // a table needs a row
  row.remove();
  hooks?.onChanged();
}

function insertColumn(td: HTMLTableCellElement, after: boolean): void {
  const wrap = wrapOf(td);
  if (!wrap || hasMerges(wrap)) return;
  const row = td.closest('tr') as HTMLTableRowElement | null;
  if (!row) return;
  const at = Array.from(row.cells).indexOf(td) + (after ? 1 : 0);

  const w = widths(wrap);
  const take = Math.max(32, Math.round((w[Math.min(at, w.length - 1)] ?? 80) / 2));
  for (const tr of Array.from(wrap.querySelectorAll('tbody > tr'))) {
    const cell = document.createElement('td');
    cell.appendChild(emptyBlock());
    const ref = (tr as HTMLTableRowElement).cells[at] ?? null;
    tr.insertBefore(cell, ref);
  }
  const colgroup = wrap.querySelector('colgroup');
  if (colgroup) {
    const col = document.createElement('col');
    col.style.width = take + 'px';
    colgroup.insertBefore(col, colgroup.children[at] ?? null);
  }
  const next = widths(wrap);
  const src = Math.min(at + 1, next.length - 1);
  if (next[src] !== undefined) next[src] = Math.max(16, next[src] - take);
  applyWidths(wrap, next);
  hooks?.onChanged();
}

function deleteColumn(td: HTMLTableCellElement): void {
  const wrap = wrapOf(td);
  if (!wrap || hasMerges(wrap)) return;
  const row = td.closest('tr') as HTMLTableRowElement | null;
  if (!row || row.cells.length <= 1) return;
  const at = Array.from(row.cells).indexOf(td);

  const w = widths(wrap);
  const freed = w[at] ?? 0;
  for (const tr of Array.from(wrap.querySelectorAll('tbody > tr'))) {
    (tr as HTMLTableRowElement).cells[at]?.remove();
  }
  wrap.querySelector('colgroup')?.children[at]?.remove();

  const next = widths(wrap);
  const give = Math.min(at, next.length - 1);
  if (next[give] !== undefined) next[give] += freed;
  applyWidths(wrap, next);
  hooks?.onChanged();
}

function deleteTable(td: HTMLTableCellElement): void {
  wrapOf(td)?.remove();
  hooks?.onChanged();
}

/* ------------------------------------------------------------------ *
 * Wiring
 * ------------------------------------------------------------------ */

export function bindTables(root: HTMLElement, h: TableHooks): void {
  hooks = h;

  root.addEventListener('mousemove', (e) => {
    const wrap = (e.target as HTMLElement)?.closest?.('.blk-table') as
      | HTMLElement
      | null;
    if (!wrap || dragging) return;
    wrap.classList.toggle('col-grab', borderAt(wrap, e.clientX) >= 0);
  });

  root.addEventListener(
    'mousedown',
    (e) => {
      const wrap = (e.target as HTMLElement)?.closest?.('.blk-table') as
        | HTMLElement
        | null;
      if (!wrap) return;
      const index = borderAt(wrap, e.clientX);
      if (index < 0) return;
      e.preventDefault();
      dragging = { wrap, index, startX: e.clientX, start: widths(wrap) };
      document.body.classList.add('col-resizing');
      document.addEventListener('mousemove', onMove, true);
      document.addEventListener('mouseup', onUp, true);
    },
    true
  );

  root.addEventListener('contextmenu', (e) => {
    const td = cellOf(e.target as Node);
    if (!td) return;
    e.preventDefault();
    const wrap = wrapOf(td);
    const merged = !!wrap && hasMerges(wrap);
    const anchor = document.createElement('div');
    anchor.style.cssText = `position:fixed;left:${e.clientX}px;top:${e.clientY}px;width:1px;height:1px`;
    document.body.appendChild(anchor);
    openMenuAt(anchor, [
      { heading: 'Row' },
      { label: 'Insert row above', onSelect: () => insertRow(td, false) },
      { label: 'Insert row below', onSelect: () => insertRow(td, true) },
      { label: 'Delete row', onSelect: () => deleteRow(td) },
      { separator: true },
      { heading: 'Column' },
      {
        label: 'Insert column left',
        disabled: merged,
        onSelect: () => insertColumn(td, false),
      },
      {
        label: 'Insert column right',
        disabled: merged,
        onSelect: () => insertColumn(td, true),
      },
      {
        label: 'Delete column',
        disabled: merged,
        onSelect: () => deleteColumn(td),
      },
      { separator: true },
      { label: 'Delete table', onSelect: () => deleteTable(td) },
      ...(merged
        ? [
            { separator: true as const },
            {
              label: 'Merged cells: column edits disabled',
              disabled: true,
            },
          ]
        : []),
    ]);
    setTimeout(() => anchor.remove(), 0);
  });
}

export { caretCell };
