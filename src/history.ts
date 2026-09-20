import type { Block } from './model';
import type { CaretPos } from './caret';
import { caretAtStart, getCaret, placeCaret } from './caret';
import { blocksIn, docEl, readModel, renderBlocks } from './render';
import { paginate } from './paginate';

/**
 * Native undo stops working the moment pagination moves nodes, so we keep our
 * own stack. Invariant: the top of `undoStack` always mirrors the current
 * document, which is what makes undo() a simple "pop and apply the new top".
 * snapshot() is therefore called AFTER a change, not before.
 */
interface Entry {
  blocks: Block[];
  caret: CaretPos | null;
  t: number;
  reason: 'structural' | 'typing' | 'base';
}

const CAP = 100;
const COALESCE_MS = 500;

let undoStack: Entry[] = [];
let redoStack: Entry[] = [];

function current(reason: Entry['reason']): Entry {
  return { blocks: readModel(docEl()), caret: getCaret(), t: Date.now(), reason };
}

/** Seed the stack with the document as loaded. */
export function resetHistory(): void {
  undoStack = [current('base')];
  redoStack = [];
}

let typingTimer = 0;

/**
 * Called on every input event. The snapshot itself is deferred until 500ms of
 * typing inactivity, because current() re-reads every block in the document
 * and must not sit on the keystroke path.
 */
export function noteTyping(): void {
  clearTimeout(typingTimer);
  typingTimer = window.setTimeout(() => {
    typingTimer = 0;
    snapshot('typing');
  }, COALESCE_MS);
}

/** Take a pending typing snapshot now, so undo never skips over it. */
export function flushTyping(): void {
  if (!typingTimer) return;
  clearTimeout(typingTimer);
  typingTimer = 0;
  snapshot('typing');
}

export function snapshot(reason: 'structural' | 'typing'): void {
  const entry = current(reason);
  const last = undoStack[undoStack.length - 1];
  // Only ever merge typing into typing. Merging into the baseline or into a
  // structural entry would silently discard a state the user can still see.
  if (
    reason === 'typing' &&
    last &&
    last.reason === 'typing' &&
    entry.t - last.t < COALESCE_MS
  ) {
    undoStack[undoStack.length - 1] = entry; // extend the current typing run
  } else {
    undoStack.push(entry);
    if (undoStack.length > CAP) undoStack.shift();
  }
  redoStack = [];
}

function apply(entry: Entry): void {
  const root = docEl();
  renderBlocks(entry.blocks, root);
  paginate();
  root.focus();
  if (entry.caret) {
    placeCaret(entry.caret);
  } else {
    // The baseline snapshot predates the first focus, so it has no caret.
    const first = blocksIn(root)[0];
    if (first) caretAtStart(first);
  }
}

export function undo(): void {
  flushTyping();
  if (undoStack.length < 2) return;
  const cur = undoStack.pop() as Entry;
  redoStack.push(cur);
  apply(undoStack[undoStack.length - 1]);
}

export function redo(): void {
  flushTyping();
  const entry = redoStack.pop();
  if (!entry) return;
  undoStack.push(entry);
  apply(entry);
}

export function canUndo(): boolean {
  return undoStack.length > 1 || typingTimer !== 0;
}

export function canRedo(): boolean {
  return redoStack.length > 0;
}
