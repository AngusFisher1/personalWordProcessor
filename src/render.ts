import type { Block, Doc, StyleId } from './model';
import { newId } from './model';
import { styleClass, styleOf } from './styles';

/* ------------------------------------------------------------------ *
 * Element construction
 * ------------------------------------------------------------------ */

export function docEl(): HTMLElement {
  const el = document.getElementById('doc');
  if (!el) throw new Error('#doc missing');
  return el;
}

export function measureEl(): HTMLElement {
  const el = document.getElementById('measure');
  if (!el) throw new Error('#measure missing');
  return el;
}

export function newPage(): HTMLElement {
  const page = document.createElement('div');
  page.className = 'page';
  const content = document.createElement('div');
  content.className = 'page-content';
  page.appendChild(content);
  return page;
}

export function pageContent(page: Element): HTMLElement {
  let c = page.querySelector(':scope > .page-content') as HTMLElement | null;
  if (!c) {
    c = document.createElement('div');
    c.className = 'page-content';
    page.appendChild(c);
  }
  return c;
}

export function makeBlockEl(b: Block): HTMLElement {
  const el = document.createElement('div');
  el.className = 'blk ' + styleClass(b.styleId);
  el.dataset.blockId = b.id;
  setBlockHtml(el, b.html);
  return el;
}

/** Empty blocks carry a <br> in the DOM (zero-height divs are uneditable) but
 *  are stored as '' in the model. */
export function setBlockHtml(el: HTMLElement, html: string): void {
  el.innerHTML = html.trim() === '' ? '<br>' : html;
}

export function applyStyle(el: HTMLElement, styleId: StyleId): void {
  const keep = Array.from(el.classList).filter(
    (c) => c !== 'blk' && !c.startsWith('s-')
  );
  el.className = ['blk', styleClass(styleId), ...keep].join(' ');
}

export function pages(root: HTMLElement = docEl()): HTMLElement[] {
  return Array.from(root.querySelectorAll(':scope > .page')) as HTMLElement[];
}

export function blocksIn(root: HTMLElement): HTMLElement[] {
  return Array.from(root.querySelectorAll('.blk')) as HTMLElement[];
}

export function blockEl(id: string): HTMLElement | null {
  const sel = `[data-block-id="${CSS.escape(id)}"]`;
  return docEl().querySelector(sel) as HTMLElement | null;
}

/* ------------------------------------------------------------------ *
 * Render / read back
 * ------------------------------------------------------------------ */

/** Lay every block out on a single page. The caller runs paginate() next. */
export function renderAll(doc: Doc, root: HTMLElement): void {
  renderBlocks(doc.blocks, root);
}

export function renderBlocks(blocks: Block[], root: HTMLElement): void {
  root.textContent = '';
  const page = newPage();
  root.appendChild(page);
  const c = pageContent(page);
  for (const b of blocks) c.appendChild(makeBlockEl(b));
}

/**
 * The DOM is the live source of truth while editing; the model object is a
 * serialization target refreshed from the DOM on save, export and snapshot.
 */
export function readModel(root: HTMLElement): Block[] {
  return blocksIn(root).map((el) => {
    let id = el.dataset.blockId;
    if (!id) {
      id = newId();
      el.dataset.blockId = id;
    }
    return { id, styleId: styleOf(el), html: readBlockHtml(el) };
  });
}

export function readBlockHtml(el: HTMLElement): string {
  const html = sanitizeInlineHtml(el.innerHTML);
  return html === '<br>' ? '' : html;
}

/* ------------------------------------------------------------------ *
 * Inline sanitizer - shared by readModel and the paste handler
 * ------------------------------------------------------------------ */

/** Removed outright, contents and all. */
const DROP = new Set([
  'SCRIPT', 'STYLE', 'META', 'LINK', 'TITLE', 'NOSCRIPT', 'IFRAME', 'OBJECT',
  'EMBED', 'IMG', 'SVG', 'CANVAS', 'VIDEO', 'AUDIO', 'INPUT', 'BUTTON',
  'SELECT', 'TEXTAREA', 'TABLE',
]);

function safeHref(v: string): boolean {
  const s = v.trim().toLowerCase();
  if (s.startsWith('javascript:') || s.startsWith('data:') || s.startsWith('vbscript:')) {
    return false;
  }
  return true;
}

type Emph = 'b' | 'i' | 'u';

/**
 * What an element says about emphasis, reading its inline style before we
 * throw the style away. Two real-world cases make this necessary:
 *
 *  - Google Docs wraps the whole clipboard payload in
 *    `<b style="font-weight:normal">`. Keeping the <b> and dropping the style
 *    would bold the entire pasted document.
 *  - Google Docs marks actual bold as `<span style="font-weight:700">`.
 *    Unwrapping the span and dropping the style would lose the emphasis.
 *
 * null means "says nothing", so the tag name decides.
 */
function styledEmphasis(el: HTMLElement): Record<Emph, boolean | null> {
  const s = el.style;
  const out: Record<Emph, boolean | null> = { b: null, i: null, u: null };

  const w = s.fontWeight;
  if (w) {
    const n = parseInt(w, 10);
    if (!Number.isNaN(n)) out.b = n >= 600;
    else if (w === 'bold' || w === 'bolder') out.b = true;
    else if (w === 'normal' || w === 'lighter') out.b = false;
  }

  const fs = s.fontStyle;
  if (fs) out.i = fs === 'italic' || fs === 'oblique';

  const td = s.textDecorationLine || s.textDecoration;
  if (td) out.u = td.includes('underline');

  return out;
}

const TAG_EMPHASIS: Record<string, Emph> = {
  B: 'b', STRONG: 'b', I: 'i', EM: 'i', CITE: 'i', VAR: 'i',
  U: 'u', INS: 'u',
};

/** Normalizes a subtree in place down to b/i/u/a/br plus text. */
export function cleanInline(parent: Node): void {
  for (const n of Array.from(parent.childNodes)) {
    if (n.nodeType === Node.TEXT_NODE) continue;
    if (n.nodeType !== Node.ELEMENT_NODE) {
      parent.removeChild(n);
      continue;
    }
    const el = n as HTMLElement;
    const tag = el.tagName;

    if (DROP.has(tag)) {
      parent.removeChild(el);
      continue;
    }
    if (tag === 'BR') {
      for (const a of Array.from(el.attributes)) el.removeAttribute(a.name);
      continue;
    }

    // Decide which emphasis this element carries, from its tag and its style.
    const styled = styledEmphasis(el);
    const wanted: Emph[] = [];
    for (const e of ['b', 'i', 'u'] as Emph[]) {
      const fromTag = TAG_EMPHASIS[tag] === e;
      const explicit = styled[e];
      if (explicit === true || (fromTag && explicit !== false)) wanted.push(e);
    }

    const href =
      tag === 'A' && safeHref(el.getAttribute('href') ?? '')
        ? el.getAttribute('href')
        : null;

    cleanInline(el);

    // Rebuild as nested b/i/u (inside an <a> when there is a link), so that
    // emphasis survives whether it arrived as a tag or as a style.
    let inner: Node = document.createDocumentFragment();
    while (el.firstChild) inner.appendChild(el.firstChild);
    for (const e of wanted) {
      const w = document.createElement(e);
      w.appendChild(inner);
      inner = w;
    }
    if (href) {
      const a = document.createElement('a');
      a.setAttribute('href', href);
      a.appendChild(inner);
      inner = a;
    }
    parent.insertBefore(inner, el);
    parent.removeChild(el);
  }
}

export function sanitizeInlineHtml(html: string): string {
  const t = document.createElement('template');
  t.innerHTML = html;
  cleanInline(t.content);
  return t.innerHTML;
}
