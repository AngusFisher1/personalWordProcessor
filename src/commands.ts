import type { StyleId } from './model';
import { newId } from './model';
import { STYLES, styleClass, styleOf } from './styles';
import {
  applyStyle,
  blocksIn,
  docEl,
  pageContent,
} from './render';
import {
  blockTextLength,
  getCaret,
  nearestBlock,
  placeCaret,
} from './caret';
import { paginate, paginateIfNeeded } from './paginate';
import { flushTyping, redo, snapshot, undo } from './history';

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
    return b ? [b] : [];
  }
  const r = s.getRangeAt(0);
  return blocksIn(docEl()).filter((b) => r.intersectsNode(b));
}

export function setBlockStyle(styleId: StyleId): void {
  flushTyping();
  const targets = selectedBlocks();
  if (targets.length === 0) return;
  const caret = getCaret();
  for (const el of targets) applyStyle(el, styleId);
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
  return blocksIn(docEl());
}

function mergeInto(prev: HTMLElement, cur: HTMLElement): void {
  flushTyping();
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
  if ((getCaret()?.offset ?? -1) !== 0) return false;

  const all = orderedBlocks();
  const i = all.indexOf(blk);
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
  if ((getCaret()?.offset ?? -1) < blockTextLength(blk)) return false;

  const all = orderedBlocks();
  const i = all.indexOf(blk);
  if (i < 0 || i >= all.length - 1) return true;
  mergeInto(blk, all[i + 1]);
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
      const blk = caretBlock();
      // No-op in a Bullet in v1; never insert a literal tab.
      if (blk && styleOf(blk) === 'Bullet') e.preventDefault();
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
    placeCaret({
      blockId: target.dataset.blockId as string,
      offset: blockTextLength(target),
    });
  });
}

export { paginateIfNeeded };
