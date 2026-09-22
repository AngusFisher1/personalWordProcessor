import { docEl } from './render';
import { canvasEl } from './shell';

/**
 * The selection bubble.
 *
 * Four buttons over a phrase you just dragged across: bold, italic,
 * underline, and take the formatting off. Everything else is in the format
 * bar above the page, which is always there - a bubble that grew a style
 * picker and an alignment group was a second format bar in a worse place.
 *
 * MOUSE ONLY. Someone selecting with Shift+arrows has their hands on the
 * keys already and does not need a thing to appear under them; it would
 * only cover the next line they were about to read.
 *
 * It lives in the canvas overlay rather than inside a page, for the same
 * reason the gutter does: anything inside `.page` is inside the
 * contenteditable and would end up in the printed output.
 */

export interface BubbleHost {
  bold(): void;
  italic(): void;
  underline(): void;
  clearFormatting(): void;
  inlineState(): { bold: boolean; italic: boolean; underline: boolean };
  /** Suppressed while another surface owns the selection. */
  suppressed(): boolean;
}

let host: BubbleHost | null = null;
let bar: HTMLElement | null = null;
let marks: Record<'bold' | 'italic' | 'underline', HTMLElement> | null = null;
let timer = 0;
/** Set by a mouse drag, cleared by a key. The bubble only exists for the first. */
let fromMouse = false;

function button(
  label: string,
  title: string,
  cls: string,
  onClick: () => void
): HTMLElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'fb-btn ' + cls;
  b.textContent = label;
  b.title = title;
  b.setAttribute('aria-label', title);
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', (e) => {
    e.preventDefault();
    onClick();
    sync();
  });
  return b;
}

function build(): HTMLElement {
  const el = document.createElement('div');
  el.className = 'bubble';
  el.setAttribute('role', 'toolbar');
  el.setAttribute('aria-label', 'Formatting');
  el.addEventListener('mousedown', (e) => e.preventDefault());

  const b = button('B', 'Bold', 'fb-b', () => host?.bold());
  const i = button('I', 'Italic', 'fb-i', () => host?.italic());
  const u = button('U', 'Underline', 'fb-u', () => host?.underline());
  el.append(b, i, u);
  marks = { bold: b, italic: i, underline: u };

  const sep = document.createElement('span');
  sep.className = 'fb-sep';
  el.appendChild(sep);
  el.appendChild(
    button('⌧', 'Clear formatting', 'fb-clear', () => host?.clearFormatting())
  );
  return el;
}

function sync(): void {
  if (!bar || !host) return;
  const on = host.inlineState();
  marks?.bold.classList.toggle('on', on.bold);
  marks?.italic.classList.toggle('on', on.italic);
  marks?.underline.classList.toggle('on', on.underline);
}

function hide(): void {
  bar?.remove();
  bar = null;
  marks = null;
}

/**
 * Above the selection, or below it when there is no room.
 *
 * It must never cover the words it is about to change, which is the whole
 * reason it is positioned against the FIRST rect of the selection rather
 * than the union: on a multi-line selection the union's top is the top of
 * the block, and centring on it would put the bubble over the text.
 */
function place(rect: DOMRect): void {
  if (!bar) return;
  const canvas = canvasEl();
  const c = canvas.getBoundingClientRect();
  const w = bar.offsetWidth || 150;
  const h = bar.offsetHeight || 32;
  const gap = 9;

  let left = rect.left - c.left + rect.width / 2 - w / 2;
  left = Math.max(8, Math.min(left, c.width - w - 8));

  const above = rect.top - c.top - h - gap;
  const below = rect.bottom - c.top + gap;
  const top = above >= 4 ? above : below;

  bar.style.left = Math.round(left) + 'px';
  bar.style.top = Math.round(top) + 'px';
  bar.classList.toggle('below', above < 4);
}

/** The selection's first line, or null when there is nothing to format. */
function selectionRect(): DOMRect | null {
  const sel = window.getSelection();
  if (!sel || sel.isCollapsed || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (range.toString().trim() === '') return null;
  const root = docEl();
  if (!root.contains(range.commonAncestorContainer)) return null;
  const rects = Array.from(range.getClientRects()).filter((r) => r.height > 0.5);
  const first = rects[0] ?? range.getBoundingClientRect();
  if (first.width === 0 && first.height === 0) return null;

  // Gone from view: a bubble pinned to the top of the canvas over a
  // selection three pages up is pointing at nothing.
  const c = canvasEl().getBoundingClientRect();
  if (first.bottom < c.top || first.top > c.bottom) return null;
  return first as DOMRect;
}

function update(): void {
  if (!host || host.suppressed() || !fromMouse) {
    hide();
    return;
  }
  const rect = selectionRect();
  if (!rect) {
    hide();
    return;
  }
  if (!bar) {
    bar = build();
    canvasEl().appendChild(bar);
  }
  sync();
  place(rect);
}

function schedule(): void {
  // Coalesced on a timer rather than a frame. selectionchange fires for
  // every character of a drag and each update measures, so the batching is
  // necessary - but requestAnimationFrame does not run at all when the page
  // is considered hidden, and a bubble that never appears in an embedded or
  // backgrounded view is worse than one that measures a millisecond late.
  clearTimeout(timer);
  timer = window.setTimeout(update, 0);
}

export function refreshBubble(): void {
  schedule();
}

export function hideBubble(): void {
  fromMouse = false;
  hide();
}

export function bindBubble(h: BubbleHost): void {
  host = h;
  const root = docEl();

  // A drag arms the bubble; any key disarms it. Tracking the INPUT rather
  // than guessing from the selection is what keeps Shift+arrow silent.
  root.addEventListener('pointerdown', (e) => {
    if (e.pointerType === 'mouse' || e.pointerType === 'pen') fromMouse = true;
  });
  root.addEventListener('keydown', () => {
    fromMouse = false;
    hide();
  });

  document.addEventListener('selectionchange', schedule);
  canvasEl().addEventListener('scroll', schedule, { passive: true });
  root.addEventListener('scroll', schedule, { passive: true });
  window.addEventListener('resize', schedule);
}
