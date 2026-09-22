import type { TableBlock } from './model';
import { newBlock, newId } from './model';
import {
  blockEl,
  docEl,
  formatOf,
  logicalIdOf,
  makeTableEl,
  setFormatOf,
} from './render';
import { addMedia, extensionFor } from './media';

/**
 * Making things that were previously only ever read.
 *
 * Tables, images and hyperlinks all arrived in the program as things an
 * imported .docx might contain. Each could be rendered, edited and written
 * back, and none could be created - which is the difference between a
 * viewer with an editing mode and a word processor.
 */

/** The block the caret is in, as a rendered element. */
export function caretBlockEl(): HTMLElement | null {
  const sel = window.getSelection();
  let n: Node | null = sel?.focusNode ?? null;
  while (n) {
    if (
      n.nodeType === Node.ELEMENT_NODE &&
      (n as HTMLElement).classList.contains('blk')
    ) {
      return n as HTMLElement;
    }
    n = n.parentNode;
  }
  return null;
}

/* ------------------------------------------------------------------ *
 * Tables
 * ------------------------------------------------------------------ */

export function newTable(rows: number, cols: number, contentWidth: number): TableBlock {
  const width = Math.floor(contentWidth / cols);
  return {
    kind: 'table',
    id: newId(),
    // The last column takes the rounding, so the widths still sum to the
    // content box: a table a pixel wide of the page is a table that lays
    // out differently from the one that was measured.
    cols: Array.from({ length: cols }, (_v, c) => ({
      width: c === cols - 1 ? contentWidth - width * (cols - 1) : width,
    })),
    rows: Array.from({ length: rows }, (_v, r) => ({
      id: newId(),
      headerRow: r === 0,
      cells: Array.from({ length: cols }, () => [newBlock('Body', '')]),
    })),
  };
}

/**
 * Put a table after the paragraph the caret is in.
 *
 * Followed by an empty paragraph when it would otherwise be the last thing
 * in the document: a table with nothing after it leaves nowhere to put the
 * caret, and no way to type your way out of it.
 */
export function insertTableAtCaret(rows: number, cols: number, contentWidth: number): boolean {
  const at = caretBlockEl();
  if (!at) return false;
  const table = newTable(rows, cols, contentWidth);
  const el = makeTableEl(table);
  at.parentElement?.insertBefore(el, at.nextSibling);

  const after = el.nextElementSibling;
  if (!after || !(after as HTMLElement).classList.contains('blk')) {
    const trailing = document.createElement('div');
    trailing.className = at.className.replace(/\bs-\w+/, 's-Body');
    trailing.dataset.blockId = newId();
    trailing.innerHTML = '<br>';
    el.parentElement?.insertBefore(trailing, el.nextSibling);
  }
  return true;
}

/* ------------------------------------------------------------------ *
 * Page breaks
 * ------------------------------------------------------------------ */

/** Toggle a break above the paragraph the caret is in. */
export function togglePageBreak(): boolean {
  const at = caretBlockEl();
  if (!at) return false;
  const id = logicalIdOf(at);
  const head = blockEl(id) ?? at;
  const now = formatOf(head)?.pageBreakBefore === true;
  setFormatOf(head, { ...(formatOf(head) ?? {}), pageBreakBefore: now ? undefined : true });
  return true;
}

export function pageBreakHere(): boolean {
  const at = caretBlockEl();
  if (!at) return false;
  return formatOf(blockEl(logicalIdOf(at)) ?? at)?.pageBreakBefore === true;
}

/* ------------------------------------------------------------------ *
 * Links
 * ------------------------------------------------------------------ */

/** http, https and mailto only; anything else is not a link we will write. */
export function safeLink(href: string): string | null {
  const t = href.trim();
  if (t === '') return null;
  if (/^(https?:\/\/|mailto:)/i.test(t)) return t;
  // A bare domain is what people actually type.
  if (/^[\w-]+(\.[\w-]+)+(\/|$)/.test(t)) return 'https://' + t;
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t)) return 'mailto:' + t;
  return null;
}

/** Wrap the selection in an anchor, replacing any it already sits inside. */
export function linkSelection(href: string): boolean {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return false;
  const range = sel.getRangeAt(0);
  if (!docEl().contains(range.commonAncestorContainer)) return false;

  // An existing link around the whole selection is retargeted rather than
  // nested: an <a> inside an <a> is not a thing the exporter can write.
  const parent = range.commonAncestorContainer.parentElement;
  const owning = parent?.closest('a');
  if (owning && range.toString() === owning.textContent) {
    owning.setAttribute('href', href);
    return true;
  }

  const a = document.createElement('a');
  a.setAttribute('href', href);
  try {
    a.appendChild(range.extractContents());
  } catch {
    return false;
  }
  // Any anchors that came along inside the selection are flattened.
  for (const nested of Array.from(a.querySelectorAll('a'))) {
    const p = nested.parentNode;
    while (nested.firstChild) p?.insertBefore(nested.firstChild, nested);
    p?.removeChild(nested);
  }
  range.insertNode(a);
  return true;
}

export function unlinkSelection(): boolean {
  const sel = window.getSelection();
  const node = sel?.focusNode;
  const a = (node?.nodeType === Node.ELEMENT_NODE
    ? (node as Element)
    : node?.parentElement
  )?.closest('a');
  if (!a) return false;
  const p = a.parentNode;
  while (a.firstChild) p?.insertBefore(a.firstChild, a);
  p?.removeChild(a);
  return true;
}

/** The link the caret is inside, if any. */
export function linkAtCaret(): string | null {
  const sel = window.getSelection();
  const node = sel?.focusNode;
  const a = (node?.nodeType === Node.ELEMENT_NODE
    ? (node as Element)
    : node?.parentElement
  )?.closest('a');
  return a?.getAttribute('href') ?? null;
}

/* ------------------------------------------------------------------ *
 * Images
 * ------------------------------------------------------------------ */

export interface InsertedImage {
  path: string;
  bytes: Uint8Array;
  width: number;
  height: number;
}

/**
 * Read a picked file and put it in the media registry.
 *
 * Scaled down to the content width if it is wider, because an image larger
 * than the page has nowhere to go: the paginator would measure a block
 * taller than a page and place it clipped.
 */
export async function readImageFile(
  file: File,
  contentWidth: number
): Promise<InsertedImage | null> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const mime = file.type || 'image/png';
  if (!/^image\//.test(mime)) return null;
  const path = 'word/media/added-' + newId().slice(0, 8) + '.' + extensionFor(mime);

  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
  const size = await new Promise<{ w: number; h: number }>((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 0, h: 0 });
    img.src = url;
  });
  URL.revokeObjectURL(url);
  if (size.w === 0) return null;

  const scale = size.w > contentWidth ? contentWidth / size.w : 1;
  addMedia(path, bytes, mime);
  return {
    path,
    bytes,
    width: Math.round(size.w * scale),
    height: Math.round(size.h * scale),
  };
}

/** The markup an inserted image is stored as. */
export function imageHtml(img: InsertedImage): string {
  return (
    `<img data-media="${img.path}" width="${img.width}" height="${img.height}">`
  );
}
