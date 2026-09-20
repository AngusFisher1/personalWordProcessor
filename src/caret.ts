import {
  blocksIn,
  docEl,
  logicalIdOf,
  pieceBase,
  piecesOf,
} from './render';

/**
 * Every repagination moves DOM nodes, which invalidates any Range held against
 * them. So the caret is expressed as data that survives node movement: a block
 * id plus a character offset within that block's rendered text.
 *
 * Both are LOGICAL: when pagination has split a paragraph across pages, the id
 * is the whole paragraph's and the offset counts into its merged text. A
 * position therefore survives being re-split at a different point, which is
 * what makes line-level pagination usable while typing.
 */
export interface CaretPos {
  blockId: string;
  offset: number;
}

export function nearestBlock(node: Node | null): HTMLElement | null {
  let n: Node | null = node;
  while (n) {
    if (n.nodeType === Node.ELEMENT_NODE) {
      const el = n as HTMLElement;
      if (el.classList.contains('blk')) return el;
      if (el.id === 'doc') return null;
    }
    n = n.parentNode;
  }
  return null;
}

export function getCaret(): CaretPos | null {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || !sel.focusNode) return null;
  const piece = nearestBlock(sel.focusNode);
  if (!piece) return null;
  const id = logicalIdOf(piece);
  if (!id) return null;
  const r = document.createRange();
  try {
    r.setStart(piece, 0);
    r.setEnd(sel.focusNode, sel.focusOffset);
  } catch {
    return null;
  }
  // toString() counts rendered characters, so the offset stays correct across
  // <b>/<i> boundaries where node-based offsets do not. pieceBase shifts it
  // into the logical paragraph's coordinate space.
  return { blockId: id, offset: pieceBase(piece) + r.toString().length };
}

/** Where the caret is within its own rendered piece, ignoring any split. */
export function localCaretOffset(): number | null {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || !sel.focusNode) return null;
  const piece = nearestBlock(sel.focusNode);
  if (!piece) return null;
  const r = document.createRange();
  try {
    r.setStart(piece, 0);
    r.setEnd(sel.focusNode, sel.focusOffset);
  } catch {
    return null;
  }
  return r.toString().length;
}

export function setCaret(pos: CaretPos): void {
  // Never steal focus back to the document the user has moved away from.
  if (document.activeElement !== docEl()) return;
  placeCaret(pos);
}

/** setCaret without the focus guard, for commands that know they own focus. */
export function placeCaret(pos: CaretPos): void {
  const pieces = piecesOf(pos.blockId);
  if (pieces.length === 0) {
    const all = blocksIn(docEl());
    const last = all[all.length - 1];
    if (last) caretAtEnd(last);
    return;
  }
  // Find the piece holding this logical offset. At an exact piece boundary the
  // earlier piece wins, which puts the caret at the end of a line rather than
  // at the start of the next page.
  let blk = pieces[0];
  for (const p of pieces) {
    blk = p;
    if (pos.offset <= pieceBase(p) + blockTextLength(p)) break;
  }
  const offset = pos.offset - pieceBase(blk);
  const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
  let seen = 0;
  let n: Node | null;
  while ((n = walker.nextNode())) {
    const len = n.textContent?.length ?? 0;
    if (seen + len >= offset) {
      select(n, offset - seen);
      return;
    }
    seen += len;
  }
  // Offset past the end, e.g. the text shrank.
  caretAtEnd(blk);
}

export function caretAtEnd(blk: HTMLElement): void {
  const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
  let last: Node | null = null;
  let n: Node | null;
  while ((n = walker.nextNode())) last = n;
  if (last) select(last, last.textContent?.length ?? 0);
  else select(blk, 0);
}

export function caretAtStart(blk: HTMLElement): void {
  const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
  const first = walker.nextNode();
  if (first) select(first, 0);
  else select(blk, 0);
}

function select(node: Node, offset: number): void {
  const sel = window.getSelection();
  if (!sel) return;
  const r = document.createRange();
  try {
    r.setStart(node, offset);
  } catch {
    r.selectNodeContents(node);
  }
  r.collapse(true);
  sel.removeAllRanges();
  sel.addRange(r);
}

/** Text length of a block as the caret counts it. */
export function blockTextLength(blk: HTMLElement): number {
  const r = document.createRange();
  r.selectNodeContents(blk);
  return r.toString().length;
}
