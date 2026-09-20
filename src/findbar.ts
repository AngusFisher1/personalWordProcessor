import type { FindOptions, Hit } from './find';
import {
  DEFAULT_OPTIONS,
  clearHighlights,
  highlightAll,
  hitPositions,
  replaceAll,
  replaceHit,
  scrollToHit,
  search,
} from './find';
import { docEl } from './render';
import { canvasEl, setSpineMarks } from './shell';

/**
 * The find bar floats in the gutter, off the paper, like the rest of the
 * chrome. It owns no document state: it searches, highlights, and hands the
 * caller a callback when the document actually changed.
 */

export interface FindHost {
  /** Called after a replacement, so the document can reflow and save once. */
  onReplaced(count: number): void;
  /** Called when the bar closes, to put the caret back in the document. */
  onClose(): void;
}

let host: FindHost | null = null;
let panel: HTMLElement | null = null;
let queryInput: HTMLInputElement | null = null;
let replaceInput: HTMLInputElement | null = null;
let countEl: HTMLElement | null = null;
let hits: Hit[] = [];
let index = 0;
let options: FindOptions = { ...DEFAULT_OPTIONS };
let searchTimer = 0;

export function isFindOpen(): boolean {
  return !!panel;
}

function el(tag: string, cls?: string, text?: string): HTMLElement {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
}

function optionButton(label: string, title: string, key: keyof FindOptions): HTMLElement {
  const b = document.createElement('button');
  b.className = 'find-opt' + (options[key] ? ' on' : '');
  b.textContent = label;
  b.title = title;
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', () => {
    options = { ...options, [key]: !options[key] };
    b.classList.toggle('on', options[key]);
    runSearch();
  });
  return b;
}

export function openFind(h: FindHost, seed = ''): void {
  host = h;
  if (panel) {
    queryInput?.focus();
    queryInput?.select();
    return;
  }

  panel = el('div', 'findbar');

  const row1 = el('div', 'find-row');
  queryInput = document.createElement('input');
  queryInput.type = 'text';
  queryInput.className = 'find-input';
  queryInput.placeholder = 'Find';
  queryInput.spellcheck = false;
  queryInput.addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = window.setTimeout(runSearch, 140);
  });
  queryInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      step(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      closeFind();
    }
  });

  countEl = el('span', 'find-count', '');
  const prev = el('button', 'find-nav', '↑');
  const next = el('button', 'find-nav', '↓');
  prev.title = 'Previous match (Shift+Enter)';
  next.title = 'Next match (Enter)';
  for (const [b, dir] of [
    [prev, -1],
    [next, 1],
  ] as [HTMLElement, number][]) {
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => step(dir));
  }
  const close = el('button', 'find-nav', '×');
  close.title = 'Close (Esc)';
  close.addEventListener('mousedown', (e) => e.preventDefault());
  close.addEventListener('click', () => closeFind());

  row1.append(queryInput, countEl, prev, next, close);

  const row2 = el('div', 'find-row');
  replaceInput = document.createElement('input');
  replaceInput.type = 'text';
  replaceInput.className = 'find-input';
  replaceInput.placeholder = 'Replace with';
  replaceInput.spellcheck = false;
  replaceInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      doReplace(e.shiftKey);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      closeFind();
    }
  });
  const one = el('button', 'find-act', 'Replace');
  const all = el('button', 'find-act', 'All');
  one.addEventListener('mousedown', (e) => e.preventDefault());
  all.addEventListener('mousedown', (e) => e.preventDefault());
  one.addEventListener('click', () => doReplace(false));
  all.addEventListener('click', () => doReplace(true));
  row2.append(replaceInput, one, all);

  const row3 = el('div', 'find-row find-opts');
  row3.append(
    optionButton('Aa', 'Match case', 'caseSensitive'),
    optionButton('W', 'Whole word', 'wholeWord'),
    optionButton('.*', 'Regular expression', 'regex')
  );
  row3.appendChild(el('span', 'find-hint', 'ENTER NEXT · ESC CLOSE'));

  panel.append(row1, row2, row3);
  canvasEl().appendChild(panel);

  if (seed) queryInput.value = seed;
  queryInput.focus();
  queryInput.select();
  runSearch();
}

export function closeFind(): void {
  clearTimeout(searchTimer);
  clearHighlights();
  setSpineMarks([]);
  panel?.remove();
  panel = null;
  queryInput = null;
  replaceInput = null;
  countEl = null;
  hits = [];
  index = 0;
  host?.onClose();
}

function runSearch(): void {
  const q = queryInput?.value ?? '';
  hits = q ? search(q, options) : [];
  index = 0;
  paint();
  if (hits.length > 0) scrollToHit(hits[0]);
}

function paint(): void {
  highlightAll(hits, index);
  setSpineMarks(hitPositions(hits));
  if (countEl) {
    countEl.textContent = hits.length
      ? `${index + 1} / ${hits.length}`
      : (queryInput?.value ? 'none' : '');
    countEl.classList.toggle('none', hits.length === 0 && !!queryInput?.value);
  }
  queryInput?.classList.toggle(
    'bad',
    !!queryInput?.value && hits.length === 0
  );
}

function step(dir: number): void {
  if (hits.length === 0) return;
  index = (index + dir + hits.length) % hits.length;
  paint();
  scrollToHit(hits[index]);
}

function doReplace(all: boolean): void {
  if (hits.length === 0) return;
  const text = replaceInput?.value ?? '';

  // One reflow for the whole operation: reflowing per replacement turns a
  // 200-hit replace into minutes of pagination.
  const n = all ? replaceAll(hits, text) : (replaceHit(hits[index], text), 1);

  clearHighlights();
  host?.onReplaced(n);

  // The document moved underneath the hit list, so find them again.
  runSearch();
  if (!all && hits.length > 0) {
    index = Math.min(index, hits.length - 1);
    paint();
    scrollToHit(hits[index]);
  }
}

/** Re-run the current search after the document changed underneath it. */
export function refreshFind(): void {
  if (!panel) return;
  const keep = index;
  runSearch();
  if (hits.length > 0) {
    index = Math.min(keep, hits.length - 1);
    paint();
  }
}

/** The word under the caret, to seed the box the way every editor does. */
export function selectedText(): string {
  const sel = window.getSelection();
  if (!sel || sel.isCollapsed || sel.rangeCount === 0) return '';
  const t = sel.toString().trim();
  if (!t || t.length > 120 || t.includes('\n')) return '';
  // Only seed from a selection inside the document.
  const node = sel.focusNode;
  return node && docEl().contains(node) ? t : '';
}
