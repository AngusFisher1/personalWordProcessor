import type { BlockAlign, BlockFormat, StyleId } from './model';
import { STYLE_IDS } from './model';
import { DOC_FONT, STYLES } from './styles';
import { docEl } from './render';
import { canvasEl } from './shell';
import { openMenuAt } from './ui';

/**
 * The formatting bar that appears over a selection.
 *
 * The rail is where you go to find out what exists; this is where you go
 * when you already know. Selecting a phrase and reaching 250px left to bold
 * it is the single most repeated gesture in the program, and the bar exists
 * to make that gesture local.
 *
 * It lives in the canvas overlay rather than inside a page, for the same
 * reason the gutter does: anything inside `.page` is inside the
 * contenteditable and would end up in the printed output.
 */

export interface FormatHost {
  bold(): void;
  italic(): void;
  underline(): void;
  setStyle(id: StyleId): void;
  currentStyle(): StyleId | null;
  inlineState(): { bold: boolean; italic: boolean; underline: boolean };
  setAlign(align: BlockAlign): void;
  currentFormat(): BlockFormat;
  /** Suppressed while another surface owns the selection. */
  suppressed(): boolean;
}

/**
 * Alignment, drawn as the lines it produces.
 *
 * Three bars of uneven width, laid out the way the paragraph would be. A
 * glyph would need an icon font, and the four alignments have no distinct
 * characters in any font we can rely on - they would all come out as ≡.
 */
const ALIGNS: { id: BlockAlign; title: string; bars: number[] }[] = [
  { id: 'left', title: 'Align left', bars: [100, 62, 84] },
  { id: 'center', title: 'Centre', bars: [100, 62, 84] },
  { id: 'right', title: 'Align right', bars: [100, 62, 84] },
  { id: 'justify', title: 'Justify', bars: [100, 100, 100] },
];

let host: FormatHost | null = null;
let bar: HTMLElement | null = null;
let styleLabel: HTMLElement | null = null;
let marks: Record<'bold' | 'italic' | 'underline', HTMLElement> | null = null;
let aligns: { id: BlockAlign; el: HTMLElement }[] = [];
let timer = 0;
/** A menu opened from the bar must not make the bar hide itself. */
let pinned = false;

function button(
  label: string,
  title: string,
  cls: string,
  onClick: () => void
): HTMLElement {
  const b = document.createElement('button');
  b.className = 'fb-btn ' + cls;
  b.textContent = label;
  b.title = title;
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
  el.className = 'formatbar';
  el.addEventListener('mousedown', (e) => e.preventDefault());

  const style = document.createElement('button');
  style.className = 'fb-btn fb-style';
  style.title = 'Paragraph style';
  styleLabel = document.createElement('span');
  style.appendChild(styleLabel);
  const caret = document.createElement('span');
  caret.className = 'fb-caret';
  style.appendChild(caret);
  style.addEventListener('mousedown', (e) => e.preventDefault());
  style.addEventListener('click', () => {
    const current = host?.currentStyle() ?? null;
    pinned = true;
    openMenuAt(
      style,
      STYLE_IDS.map((id) => {
        const d = STYLES[id];
        return {
          label: d.label,
          checked: id === current,
          preview: {
            fontFamily: DOC_FONT,
            fontSize: Math.min(19, Math.max(12, d.size * 1.2)) + 'px',
            fontWeight: d.bold ? '700' : '400',
            textTransform: d.uppercase ? 'uppercase' : 'none',
            letterSpacing: d.letterSpacing ? d.letterSpacing + 'px' : 'normal',
          },
          onSelect: () => {
            pinned = false;
            host?.setStyle(id);
            sync();
          },
        };
      })
    );
    // A menu dismissed without choosing must not leave the bar stuck open.
    setTimeout(() => {
      pinned = false;
    }, 0);
  });
  el.appendChild(style);

  const sep = document.createElement('span');
  sep.className = 'fb-sep';
  el.appendChild(sep);

  const b = button('B', 'Bold', 'fb-b', () => host?.bold());
  const i = button('I', 'Italic', 'fb-i', () => host?.italic());
  const u = button('U', 'Underline', 'fb-u', () => host?.underline());
  el.append(b, i, u);
  marks = { bold: b, italic: i, underline: u };

  const sep2 = document.createElement('span');
  sep2.className = 'fb-sep';
  el.appendChild(sep2);

  aligns = ALIGNS.map((a) => {
    const btn = button('', a.title, 'fb-align fb-align-' + a.id, () => host?.setAlign(a.id));
    for (const w of a.bars) {
      const bar = document.createElement('i');
      bar.style.width = w + '%';
      btn.appendChild(bar);
    }
    el.appendChild(btn);
    return { id: a.id, el: btn };
  });

  return el;
}

/** Update the bar's own state from the selection it is sitting over. */
function sync(): void {
  if (!bar || !host) return;
  const st = host.currentStyle();
  if (styleLabel) styleLabel.textContent = st ? STYLES[st].label : 'Mixed';
  const on = host.inlineState();
  marks?.bold.classList.toggle('on', on.bold);
  marks?.italic.classList.toggle('on', on.italic);
  marks?.underline.classList.toggle('on', on.underline);
  const align = host.currentFormat().align ?? 'left';
  for (const a of aligns) a.el.classList.toggle('on', a.id === align);
}

function hide(): void {
  bar?.remove();
  bar = null;
  styleLabel = null;
  marks = null;
  aligns = [];
}

/**
 * Where the bar goes: centred over the selection, above it if there is room
 * and below it if there is not, and never off the side of the canvas.
 */
function place(rect: DOMRect): void {
  if (!bar) return;
  const canvas = canvasEl();
  const c = canvas.getBoundingClientRect();
  const w = bar.offsetWidth || 220;
  const h = bar.offsetHeight || 32;
  const gap = 9;

  let left = rect.left - c.left + rect.width / 2 - w / 2;
  left = Math.max(8, Math.min(left, c.width - w - 8));

  let top = rect.top - c.top - h - gap;
  if (top < 4) top = rect.bottom - c.top + gap;

  bar.style.left = Math.round(left) + 'px';
  bar.style.top = Math.round(top) + 'px';
}

/**
 * The selection's rectangle, or null if there is nothing to format.
 *
 * getBoundingClientRect on a range spanning several lines gives the union,
 * whose centre is the middle of the paragraph. The FIRST rect is the first
 * line, which is where the eye is, so the bar goes over that.
 */
function selectionRect(): DOMRect | null {
  const sel = window.getSelection();
  if (!sel || sel.isCollapsed || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (range.toString().trim() === '') return null;
  const root = docEl();
  if (!root.contains(range.commonAncestorContainer)) return null;
  const rects = Array.from(range.getClientRects()).filter((r) => r.height > 0.5);
  const first = rects[0] ?? range.getBoundingClientRect();
  return first.width === 0 && first.height === 0 ? null : (first as DOMRect);
}

function update(): void {
  if (!host || host.suppressed() || pinned) {
    if (!pinned) hide();
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
  // is considered hidden, and a bar that never appears in an embedded or
  // backgrounded view is worse than one that measures a millisecond late.
  clearTimeout(timer);
  timer = window.setTimeout(update, 0);
}

export function refreshFormatBar(): void {
  schedule();
}

export function hideFormatBar(): void {
  pinned = false;
  hide();
}

export function bindFormatBar(h: FormatHost): void {
  host = h;
  document.addEventListener('selectionchange', schedule);
  // Scrolling moves the selection under the bar, so the bar follows it.
  canvasEl().addEventListener('scroll', schedule, { passive: true });
  docEl().addEventListener('scroll', schedule, { passive: true });
  window.addEventListener('resize', schedule);
  // Typing replaces the selection; the bar should be gone before the
  // character lands, not a frame later.
  docEl().addEventListener('keydown', (e) => {
    if (e.key.length === 1 || e.key === 'Backspace' || e.key === 'Delete') hide();
  });
}
