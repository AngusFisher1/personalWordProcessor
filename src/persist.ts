import type {
  Block,
  Doc,
  MarginKey,
  PageSetup,
  ParagraphBlock,
  TableBlock,
} from './model';
import { MARGINS, isStyleId, newId, pageSetup } from './model';

const INDEX_KEY = 'wp:docs';
const LAST_KEY = 'wp:last';
const docKey = (id: string) => 'wp:doc:' + id;

export interface IndexEntry {
  id: string;
  title: string;
  updatedAt: number;
  /** Counts kept in the index so the library lists without loading anything. */
  words: number;
  pages: number;
}

export interface DocMeta {
  words: number;
  pages: number;
}

/** localStorage throws when full; callers learn about it through this event. */
function announce(ok: boolean, error?: unknown): void {
  document.dispatchEvent(
    new CustomEvent('wp:saved', { detail: { ok, error } })
  );
}

export function save(doc: Doc, meta?: DocMeta): void {
  try {
    localStorage.setItem(docKey(doc.id), JSON.stringify(doc));
    const prev = readIndex().find((e) => e.id === doc.id);
    const idx = readIndex().filter((e) => e.id !== doc.id);
    idx.unshift({
      id: doc.id,
      title: doc.title,
      updatedAt: Date.now(),
      words: meta?.words ?? prev?.words ?? 0,
      pages: meta?.pages ?? prev?.pages ?? 1,
    });
    localStorage.setItem(INDEX_KEY, JSON.stringify(idx));
    localStorage.setItem(LAST_KEY, doc.id);
    announce(true);
  } catch (error) {
    announce(false, error);
  }
}

export function readIndex(): IndexEntry[] {
  try {
    const raw = localStorage.getItem(INDEX_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? (parsed as IndexEntry[]) : [];
  } catch {
    return [];
  }
}

export function load(): Doc | null {
  const id = localStorage.getItem(LAST_KEY) ?? readIndex()[0]?.id;
  return id ? loadById(id) : null;
}

export function loadById(id: string): Doc | null {
  try {
    const raw = localStorage.getItem(docKey(id));
    if (!raw) return null;
    return coerce(JSON.parse(raw));
  } catch {
    return null;
  }
}

/** Forget a document entirely: its content, its index entry, its original. */
export function removeDoc(id: string): void {
  try {
    localStorage.removeItem(docKey(id));
    localStorage.setItem(
      INDEX_KEY,
      JSON.stringify(readIndex().filter((e) => e.id !== id))
    );
    if (localStorage.getItem(LAST_KEY) === id) localStorage.removeItem(LAST_KEY);
  } catch {
    /* nothing recoverable to do */
  }
}

/** A copy with fresh ids, so editing it cannot disturb the original. */
export function duplicateDoc(doc: Doc): Doc {
  const copy: Doc = JSON.parse(JSON.stringify(doc));
  copy.id = newId();
  copy.title = doc.title.replace(/\s*\(copy( \d+)?\)$/, '') + ' (copy)';
  // New block ids: the originals are keys into the source document's vault,
  // and two documents must never claim the same preserved XML.
  for (const b of copy.blocks) {
    if ((b as { kind?: string }).kind === 'table') {
      const t = b as { id: string; rows: { id: string; cells: { id: string }[][] }[] };
      t.id = newId();
      for (const row of t.rows) {
        row.id = newId();
        for (const cell of row.cells) for (const p of cell) p.id = newId();
      }
    } else {
      (b as { id: string }).id = newId();
    }
  }
  return copy;
}

/** Bytes of localStorage the documents occupy, for the library footer. */
export function storageBytes(): number {
  let total = 0;
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (!k || !k.startsWith('wp:')) continue;
      total += (localStorage.getItem(k) ?? '').length;
    }
  } catch {
    return 0;
  }
  return total;
}

export function exportJson(doc: Doc): string {
  return JSON.stringify(doc, null, 2);
}

export function importJson(json: string): Doc {
  return coerce(JSON.parse(json));
}

/** Validates the shape and rejects anything we cannot render. */
function coerce(raw: unknown): Doc {
  if (!raw || typeof raw !== 'object') throw new Error('Not a document');
  const o = raw as Record<string, unknown>;
  if (!Array.isArray(o.blocks)) throw new Error('blocks is not an array');

  const blocks: Block[] = o.blocks.map((b, i) => {
    if (!b || typeof b !== 'object') throw new Error(`Block ${i} is not an object`);
    const rb = b as Record<string, unknown>;
    if (rb.kind === 'table') return coerceTable(rb, i);
    if (!isStyleId(rb.styleId)) {
      throw new Error(`Block ${i} has unknown styleId ${String(rb.styleId)}`);
    }
    return {
      id: typeof rb.id === 'string' && rb.id ? rb.id : newId(),
      styleId: rb.styleId,
      html: typeof rb.html === 'string' ? rb.html : '',
      // Markers resolved from an imported numbering definition are part of
      // the document, not of the render, so they have to survive a reload.
      ...(typeof rb.listMarker === 'string' ? { listMarker: rb.listMarker } : {}),
      ...(typeof rb.listLevel === 'number' && rb.listLevel > 0
        ? { listLevel: rb.listLevel }
        : {}),
    };
  });

  return {
    id: typeof o.id === 'string' && o.id ? o.id : newId(),
    title: typeof o.title === 'string' && o.title ? o.title : 'Untitled',
    page: coercePage(o),
    blocks: blocks.length ? blocks : [{ id: newId(), styleId: 'Body', html: '' }],
  };
}

function coerceParagraph(raw: unknown, where: string): ParagraphBlock {
  const rb = (raw ?? {}) as Record<string, unknown>;
  if (!isStyleId(rb.styleId)) {
    throw new Error(`${where} has unknown styleId ${String(rb.styleId)}`);
  }
  return {
    id: typeof rb.id === 'string' && rb.id ? rb.id : newId(),
    styleId: rb.styleId,
    html: typeof rb.html === 'string' ? rb.html : '',
    ...(typeof rb.listMarker === 'string' ? { listMarker: rb.listMarker } : {}),
    ...(typeof rb.listLevel === 'number' && rb.listLevel > 0
      ? { listLevel: rb.listLevel }
      : {}),
  };
}

function coerceTable(rb: Record<string, unknown>, i: number): TableBlock {
  if (!Array.isArray(rb.rows)) throw new Error(`Table ${i} has no rows`);
  const cols = Array.isArray(rb.cols)
    ? rb.cols.map((c) => ({
        width: typeof (c as { width?: unknown })?.width === 'number'
          ? (c as { width: number }).width
          : 0,
      }))
    : [];
  const rows = rb.rows.map((r, ri) => {
    const rr = (r ?? {}) as Record<string, unknown>;
    if (!Array.isArray(rr.cells)) throw new Error(`Table ${i} row ${ri} has no cells`);
    return {
      id: typeof rr.id === 'string' && rr.id ? rr.id : newId(),
      headerRow: rr.headerRow === true,
      cells: rr.cells.map((cell) =>
        Array.isArray(cell)
          ? cell.map((pb, pi) => coerceParagraph(pb, `Table ${i} row ${ri} cell ${pi}`))
          : []
      ),
    };
  });
  return {
    kind: 'table',
    id: typeof rb.id === 'string' && rb.id ? rb.id : newId(),
    cols,
    rows,
  };
}

/** Accepts the current shape and the phase-1 `margin: 'narrow' | 'normal'`. */
function coercePage(o: Record<string, unknown>): PageSetup {
  const p = o.page as Partial<PageSetup> | undefined;
  if (p && typeof p.width === 'number' && typeof p.height === 'number') {
    const m = (p.margins ?? {}) as Partial<PageSetup['margins']>;
    const num = (v: unknown, d: number) => (typeof v === 'number' ? v : d);
    return {
      width: p.width,
      height: p.height,
      margins: {
        top: num(m.top, 48),
        right: num(m.right, 48),
        bottom: num(m.bottom, 48),
        left: num(m.left, 48),
      },
    };
  }
  if (typeof o.margin === 'string' && o.margin in MARGINS) {
    return pageSetup(o.margin as MarginKey);
  }
  return pageSetup('narrow');
}

/**
 * A filename that will survive every filesystem we might land on.
 *
 * Kept deliberately permissive: an accented or non-Latin title should come
 * out as itself, not as a row of dropped letters, so only the characters
 * Windows actually refuses are replaced.
 */
export function safeFileName(title: string, fallback = 'document'): string {
  const cleaned = title
    .replace(/[\/:*?"<>|]/g, '-')
    .replace(/[\x00-\x1f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^\.+/, '')
    .replace(/[. ]+$/, '')
    .slice(0, 80)
    .trim();
  return cleaned === '' ? fallback : cleaned;
}

export function download(filename: string, blob: Blob): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
