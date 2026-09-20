import type { StyleId } from './model';
import { newId } from './model';
import { STYLES, styleClass, styleOf } from './styles';
import {
  applyStyle,
  docEl,
  isContinuation,
  logicalHeads,
  logicalIdOf,
  mergeGroup,
  pieceBase,
  pageContent,
  piecesOf,
} from './render';
import {
  blockTextLength,
  getCaret,
  localCaretOffset,
  nearestBlock,
  placeCaret,
} from './caret';
import { paginate, paginateIfNeeded } from './paginate';
import { flushTyping, redo, snapshot, undo } from './history';
import { arrowInTable, tabInTable } from './tables';

/** Broadcast so main.ts can autosave and refresh the toolbar. */
export function notifyChanged(): void {
  document.dispatchEvent(new CustomEvent('wp:changed'));
}

function isMod(e: KeyboardEvent): boolean {
  return e.metaKey || e.ctrlKey;
}

function sel(): Selection | null {
  const s = window.getSelection();
  return s && s.rangeCount > 0 ? s : null;
}

function caretBlock(): HTMLElement | null {
  const s = sel();
  return s?.focusNode ? nearestBlock(s.focusNode) : null;
}

/** The head element of the logical block a rendered piece belongs to. */
function headOfPiece(el: HTMLElement): HTMLElement {
  return piecesOf(logicalIdOf(el))[0] ?? el;
}

function makeBlockAfter(ref: HTMLElement, styleId: StyleId): HTMLElement {
  const el = document.createElement('div');
  el.className = 'blk ' + styleClass(styleId);
  el.dataset.blockId = newId();
  el.innerHTML = '<br>';
  ref.parentElement?.insertBefore(el, ref.nextSibling);
  return el;
}

function isEmptyBlock(el: HTMLElement): boolean {
  return (el.textContent ?? '') === '';
}

/* ------------------------------------------------------------------ *
 * Splits
 *
 * Pagination may have split the caret's paragraph across pages. Commands that
 * restructure a paragraph - Enter, and a merge across a block boundary - must
 * see one element, so they fold the pieces back first and let the reflow they
 * end with re-split. Plain typing needs none of this: a character inserted
 * into a piece is already in the right logical place.
 * ------------------------------------------------------------------ */

/** Fold a block's continuations back into its head, keeping the caret put. */
function unsplitBlock(el: HTMLElement): boolean {
  const pieces = piecesOf(logicalIdOf(el));
  if (pieces.length <= 1) return false;
  const pos = getCaret();
  mergeGroup({ head: pieces[0], tails: pieces.slice(1) });
  if (pos) {
    docEl().focus();
    placeCaret(pos);
  }
  return true;
}

function unsplitCaretBlock(): void {
  const blk = caretBlock();
  if (blk) unsplitBlock(blk);
}

function fillIfEmpty(el: HTMLElement): void {
  if (!el.firstChild || (el.textContent ?? '') === '') el.innerHTML = '<br>';
}

/* ------------------------------------------------------------------ *
 * Inline formatting
 * ------------------------------------------------------------------ */

export function toggleInline(cmd: 'bold' | 'italic' | 'underline'): void {
  flushTyping();
  docEl().focus();
  document.execCommand(cmd, false);
  // Bold and italic change glyph widths, so lines can rewrap and heights can
  // change above the caret: this needs a full reflow, not the sync check.
  paginate();
  snapshot('structural');
  notifyChanged();
}

export function inlineState(): Record<'bold' | 'italic' | 'underline', boolean> {
  const q = (c: string) => {
    try {
      return document.queryCommandState(c);
    } catch {
      return false;
    }
  };
  return { bold: q('bold'), italic: q('italic'), underline: q('underline') };
}

/* ------------------------------------------------------------------ *
 * Block styles
 * ------------------------------------------------------------------ */

export function selectedBlocks(): HTMLElement[] {
  const s = sel();
  if (!s) return [];
  if (s.isCollapsed) {
    const b = caretBlock();
    return b ? [headOfPiece(b)] : [];
  }
  const r = s.getRangeAt(0);
  const hit = logicalHeads(docEl()).filter((head) =>
    piecesOf(logicalIdOf(head)).some((p) => r.intersectsNode(p))
  );
  return hit;
}

export function setBlockStyle(styleId: StyleId): void {
  flushTyping();
  const targets = selectedBlocks();
  if (targets.length === 0) return;
  const caret = getCaret();
  for (const el of targets) {
    for (const piece of piecesOf(logicalIdOf(el))) applyStyle(piece, styleId);
  }
  paginate(); // heights change above the caret
  if (caret) {
    docEl().focus();
    placeCaret(caret);
  }
  snapshot('structural');
  notifyChanged();
}

export function currentStyle(): StyleId | null {
  const blocks = selectedBlocks();
  if (blocks.length === 0) return null;
  const first = styleOf(blocks[0]);
  return blocks.every((b) => styleOf(b) === first) ? first : null;
}

/* ------------------------------------------------------------------ *
 * Enter / Backspace / Delete
 * ------------------------------------------------------------------ */

function deleteSelectionIfAny(): void {
  const s = sel();
  if (s && !s.isCollapsed) document.execCommand('delete', false);
}

function splitAtCaret(): void {
  flushTyping();
  deleteSelectionIfAny();
  unsplitCaretBlock();
  const s = sel();
  const blk = caretBlock();
  if (!s || !blk || !s.focusNode) return;

  const atEnd = (getCaret()?.offset ?? 0) >= blockTextLength(blk);
  const styleId = styleOf(blk);
  const nextStyle = atEnd ? STYLES[styleId].enterTo : styleId;

  const r = document.createRange();
  r.setStart(s.focusNode, s.focusOffset);
  r.setEnd(blk, blk.childNodes.length);
  const tail = r.extractContents();

  const el = makeBlockAfter(blk, nextStyle);
  el.innerHTML = '';
  el.appendChild(tail);
  fillIfEmpty(el);
  fillIfEmpty(blk);

  paginate();
  docEl().focus();
  placeCaret({ blockId: el.dataset.blockId as string, offset: 0 });
  snapshot('structural');
  notifyChanged();
}

function orderedBlocks(): HTMLElement[] {
  return logicalHeads(docEl());
}

function mergeInto(prev: HTMLElement, cur: HTMLElement): void {
  flushTyping();
  unsplitBlock(prev);
  unsplitBlock(cur);
  const joinAt = blockTextLength(prev);
  if (isEmptyBlock(prev)) prev.innerHTML = '';
  if (!isEmptyBlock(cur)) {
    while (cur.firstChild) prev.appendChild(cur.firstChild);
  }
  cur.remove();
  fillIfEmpty(prev);
  paginate();
  docEl().focus();
  placeCaret({ blockId: prev.dataset.blockId as string, offset: joinAt });
  snapshot('structural');
  notifyChanged();
}

/** Backspace at offset 0: merge into the previous block, adopting its style. */
function backspaceAtStart(): boolean {
  const s = sel();
  if (!s || !s.isCollapsed) return false;
  const blk = caretBlock();
  if (!blk) return false;

  // At the top of a continuation the caret looks like offset 0 but is in the
  // middle of the paragraph. Fold the pieces together and let the browser do
  // an ordinary deletion inside the one element that results.
  if (isContinuation(blk) && localCaretOffset() === 0) {
    unsplitBlock(blk);
    return false;
  }
  if ((getCaret()?.offset ?? -1) !== 0) return false;

  const all = orderedBlocks();
  const i = all.indexOf(headOfPiece(blk));
  if (i <= 0) return true; // at the very start: swallow, nothing to merge into
  mergeInto(all[i - 1], blk);
  return true;
}

/** Delete at end of block: pull the next block up into this one. */
function deleteAtEnd(): boolean {
  const s = sel();
  if (!s || !s.isCollapsed) return false;
  const blk = caretBlock();
  if (!blk) return false;

  const pieces = piecesOf(logicalIdOf(blk));
  const atPieceEnd = localCaretOffset() === blockTextLength(blk);
  const isLastPiece = pieces[pieces.length - 1] === blk;
  if (atPieceEnd && !isLastPiece) {
    // End of a piece, but the paragraph carries on across the page break.
    unsplitBlock(blk);
    return false;
  }

  const head = headOfPiece(blk);
  const logicalLen = pieces.reduce((n, p) => n + blockTextLength(p), 0);
  if ((getCaret()?.offset ?? -1) < logicalLen) return false;

  const all = orderedBlocks();
  const i = all.indexOf(head);
  if (i < 0 || i >= all.length - 1) return true;
  mergeInto(head, all[i + 1]);
  return true;
}

/* ------------------------------------------------------------------ *
 * Key bindings
 * ------------------------------------------------------------------ */

export function bindShortcuts(root: HTMLElement): void {
  root.addEventListener('keydown', (e: KeyboardEvent) => {
    if (e.isComposing) return;

    if (isMod(e)) {
      const k = e.key.toLowerCase();
      if (k === 'b' || k === 'i' || k === 'u') {
        e.preventDefault();
        toggleInline(k === 'b' ? 'bold' : k === 'i' ? 'italic' : 'underline');
        return;
      }
      if (k === 'z') {
        e.preventDefault(); // before anything else, or native undo fires too
        if (e.shiftKey) redo();
        else undo();
        notifyChanged();
        return;
      }
      if (k === 'y') {
        e.preventDefault();
        redo();
        notifyChanged();
        return;
      }
    }

    if (e.key === 'Enter') {
      if (e.shiftKey) return; // shift+enter inserts a <br>, browser handles it
      e.preventDefault();
      splitAtCaret();
      return;
    }

    if (e.key === 'Backspace') {
      const blk = caretBlock();
      if (blk && backspaceAtStart()) {
        e.preventDefault();
      }
      return;
    }

    if (e.key === 'Delete') {
      const blk = caretBlock();
      if (blk && deleteAtEnd()) {
        e.preventDefault();
      }
      return;
    }

    if (e.key === 'Tab') {
      // Inside a table, Tab walks the cells and grows the table at the end.
      if (tabInTable(e.shiftKey)) {
        e.preventDefault();
        return;
      }
      const blk = caretBlock();
      // No-op in a Bullet in v1; never insert a literal tab.
      if (blk && styleOf(blk) === 'Bullet') e.preventDefault();
      return;
    }

    if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      // Only intercepted at a cell's edge; inside one, arrows behave.
      if (arrowInTable(e.key === 'ArrowUp' ? -1 : 1)) e.preventDefault();
      return;
    }
  });

  // Clicking in the margin below the last line should still place the caret.
  root.addEventListener('mousedown', (e) => {
    const t = e.target as HTMLElement;
    if (t.closest('.blk')) return;
    const page = t.closest('.page') as HTMLElement | null;
    const all = orderedBlocks();
    let target: HTMLElement | undefined;
    if (page) {
      const own = Array.from(pageContent(page).children) as HTMLElement[];
      target = own[own.length - 1];
    }
    if (!target) target = all[all.length - 1];
    if (!target) return;
    e.preventDefault();
    root.focus();
    // target may be a continuation piece, so convert to logical coordinates.
    placeCaret({
      blockId: logicalIdOf(target),
      offset: pieceBase(target) + blockTextLength(target),
    });
  });
}

export { paginateIfNeeded };
