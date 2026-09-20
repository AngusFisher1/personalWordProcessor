import type { StyleId } from './model';
import { newBlock } from './model';
import {
  applyStyle,
  blocksIn,
  cleanInline,
  docEl,
  makeBlockEl,
} from './render';
import { blockTextLength, nearestBlock, placeCaret } from './caret';
import { paginate } from './paginate';
import { flushTyping, snapshot } from './history';
import { notifyChanged } from './commands';

/**
 * Pasting an existing resume is the first real action a user takes, so the
 * sanitizer is load-bearing rather than polish: nothing from Word or Google
 * Docs may bring its own fonts, colors or spacing into the document.
 */

interface Item {
  styleId: StyleId;
  html: string;
}

const BLOCK_TAGS = new Set([
  'P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'UL', 'OL',
  'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'MAIN', 'ASIDE', 'NAV',
  'BLOCKQUOTE', 'PRE', 'DL', 'DT', 'DD', 'FIGURE', 'FIGCAPTION', 'ADDRESS',
  'HR', 'TR',
]);

const DROP_SELECTOR =
  'script,style,meta,link,title,noscript,iframe,object,embed,img,svg,' +
  'canvas,video,audio,input,button,select,textarea,table';

interface Ctx {
  nameUsed: boolean;
  defaultStyle: StyleId;
}

function hasBlockDescendant(el: Element): boolean {
  for (const c of Array.from(el.children)) {
    if (BLOCK_TAGS.has(c.tagName)) return true;
    if (hasBlockDescendant(c)) return true;
  }
  return false;
}

function styleForTag(el: Element, ctx: Ctx): StyleId {
  // Word desktop does not paste <ul><li>; it pastes paragraphs classed
  // MsoListParagraph with the bullet glyph inlined as text.
  const cls = el.getAttribute('class') ?? '';
  if (/MsoList/i.test(cls)) return 'Bullet';

  switch (el.tagName) {
    case 'H1':
      if (!ctx.nameUsed) {
        ctx.nameUsed = true;
        return 'Name';
      }
      return 'SectionHeading';
    case 'H2':
    case 'H3':
    case 'H4':
    case 'H5':
    case 'H6':
      return 'SectionHeading';
    case 'LI':
      return 'Bullet';
    default:
      return ctx.defaultStyle;
  }
}

/** Sanitize a run of nodes down to inline markup and collapse its whitespace. */
function inlineHtmlOf(nodes: Node[]): string {
  const t = document.createElement('template');
  for (const n of nodes) t.content.appendChild(n.cloneNode(true));
  cleanInline(t.content);
  // Word and Google Docs both paste runs of spaces and non-breaking spaces
  // liberally, and they break the bullet hanging indent.
  const walker = document.createTreeWalker(t.content, NodeFilter.SHOW_TEXT);
  let n: Node | null;
  while ((n = walker.nextNode())) {
    n.textContent = (n.textContent ?? '').replace(/[\s ]+/g, ' ');
  }
  return t.innerHTML.replace(/^(\s|&nbsp;)+|(\s|&nbsp;)+$/g, '');
}

/** Leading bullet glyphs that arrive as literal text rather than as markup. */
const LEADING_BULLET = /^\s*(?:[•·▪●◦⁃∙]|o\s)\s*/;

function push(out: Item[], styleId: StyleId, nodes: Node[]): void {
  let html = inlineHtmlOf(nodes);
  let text = html.replace(/<[^>]*>/g, '').trim();
  if (text === '') return; // drop the blank paragraphs Word sprays everywhere

  let style = styleId;
  if (LEADING_BULLET.test(text)) {
    // A pasted bullet glyph becomes a real Bullet block, not a Body block
    // that happens to start with a dot.
    html = html.replace(/(^|>)([^<]*)/, (m, lt: string, body: string) =>
      LEADING_BULLET.test(body) ? lt + body.replace(LEADING_BULLET, '') : m
    );
    style = 'Bullet';
    text = html.replace(/<[^>]*>/g, '').trim();
    if (text === '') return;
  }
  out.push({ styleId: style, html });
}

function walk(parent: Node, out: Item[], ctx: Ctx): void {
  let run: Node[] = [];
  const flush = () => {
    if (run.length === 0) return;
    push(out, ctx.defaultStyle, run);
    run = [];
  };

  for (const n of Array.from(parent.childNodes)) {
    if (n.nodeType === Node.ELEMENT_NODE) {
      const el = n as Element;
      const isBlock = BLOCK_TAGS.has(el.tagName);
      const nested = hasBlockDescendant(el);

      if (!isBlock && nested) {
        // An inline wrapper that contains block elements. Google Docs wraps
        // its entire clipboard payload in one <b>, so treating this as inline
        // would flatten the whole paste into a single paragraph.
        flush();
        walk(el, out, ctx);
        continue;
      }
      if (isBlock) {
        flush();
        const styleId = styleForTag(el, ctx);
        if (nested) {
          const prev = ctx.defaultStyle;
          ctx.defaultStyle = styleId;
          walk(el, out, ctx);
          ctx.defaultStyle = prev;
        } else {
          push(out, styleId, Array.from(el.childNodes));
        }
        continue;
      }
    }
    run.push(n);
  }
  flush();
}

export function parseHtml(html: string, docIsEmpty: boolean): Item[] {
  const parsed = new DOMParser().parseFromString(html, 'text/html');
  for (const el of Array.from(parsed.body.querySelectorAll(DROP_SELECTOR))) {
    el.remove();
  }
  const out: Item[] = [];
  walk(parsed.body, out, { nameUsed: !docIsEmpty, defaultStyle: 'Body' });
  return out;
}

export function parseText(text: string): Item[] {
  return text
    .split(/\r?\n/)
    .map((line) => line.replace(/[\s ]+/g, ' ').trim())
    .filter((line) => line !== '')
    .map((line) => ({ styleId: 'Body' as StyleId, html: escapeHtml(line) }));
}

function escapeHtml(s: string): string {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

/* ------------------------------------------------------------------ *
 * Insertion
 * ------------------------------------------------------------------ */

function isEmpty(el: HTMLElement): boolean {
  return (el.textContent ?? '') === '';
}

function insertItems(items: Item[]): void {
  flushTyping();
  const s = window.getSelection();
  if (!s || s.rangeCount === 0) return;
  if (!s.isCollapsed) document.execCommand('delete', false);

  const sel2 = window.getSelection();
  if (!sel2 || !sel2.focusNode) return;
  const blk = nearestBlock(sel2.focusNode);
  if (!blk) return;

  // Everything after the caret gets re-attached to the last pasted block.
  const r = document.createRange();
  r.setStart(sel2.focusNode, sel2.focusOffset);
  r.setEnd(blk, blk.childNodes.length);
  const tail = r.extractContents();

  if (isEmpty(blk)) {
    applyStyle(blk, items[0].styleId);
    blk.innerHTML = items[0].html || '<br>';
  } else {
    const t = document.createElement('template');
    t.innerHTML = items[0].html;
    blk.appendChild(t.content);
  }

  let last = blk;
  let ref: HTMLElement = blk;
  for (let i = 1; i < items.length; i++) {
    const el = makeBlockEl(newBlock(items[i].styleId, items[i].html));
    ref.parentElement?.insertBefore(el, ref.nextSibling);
    ref = el;
    last = el;
  }

  const caretOffset = blockTextLength(last);
  if ((tail.textContent ?? '') !== '') {
    if (isEmpty(last)) last.innerHTML = '';
    last.appendChild(tail);
  }
  if (!last.firstChild) last.innerHTML = '<br>';

  paginate();
  docEl().focus();
  placeCaret({ blockId: last.dataset.blockId as string, offset: caretOffset });
}

export function bindPaste(root: HTMLElement): void {
  root.addEventListener('paste', (e: ClipboardEvent) => {
    e.preventDefault();
    const cd = e.clipboardData;
    if (!cd) return;

    const docIsEmpty = blocksIn(docEl()).every(
      (b) => (b.textContent ?? '').trim() === ''
    );

    const html = cd.getData('text/html');
    const plain = cd.getData('text/plain');
    let items = html ? parseHtml(html, docIsEmpty) : [];
    if (items.length === 0) items = parseText(plain);
    if (items.length === 0) return;

    insertItems(items);
    snapshot('structural');
    notifyChanged();
  });
}
