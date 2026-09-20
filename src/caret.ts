import { blockEl, blocksIn, docEl } from './render';

/**
 * Every repagination moves DOM nodes, which invalidates any Range held against
 * them. So the caret is expressed as data that survives node movement: a block
 * id plus a character offset within that block's rendered text.
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
  const blk = nearestBlock(sel.focusNode);
  if (!blk) return null;
  const id = blk.dataset.blockId;
  if (!id) return null;
  const r = document.createRange();
  try {
    r.setStart(blk, 0);
    r.setEnd(sel.focusNode, sel.focusOffset);
  } catch {
    return null;
  }
  // toString() counts rendered characters, so the offset stays correct across
  // <b>/<i> boundaries where node-based offsets do not.
  return { blockId: id, offset: r.toString().length };
}

export function setCaret(pos: CaretPos): void {
  // Never steal focus back to the document the user has moved away from.
  if (document.activeElement !== docEl()) return;
  placeCaret(pos);
}

/** setCaret without the focus guard, for commands that know they own focus. */
export function placeCaret(pos: CaretPos): void {
  const blk = blockEl(pos.blockId);
  if (!blk) {
    const all = blocksIn(docEl());
    const last = all[all.length - 1];
    if (last) caretAtEnd(last);
    return;
  }
  const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
  let seen = 0;
  let n: Node | null;
  while ((n = walker.nextNode())) {
    const len = n.textContent?.length ?? 0;
    if (seen + len >= pos.offset) {
      select(n, pos.offset - seen);
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
