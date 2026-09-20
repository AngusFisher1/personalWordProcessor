import type { Block, Doc, MarginKey, PageSetup } from './model';
import { MARGINS, isStyleId, newId, pageSetup } from './model';

const INDEX_KEY = 'wp:docs';
const LAST_KEY = 'wp:last';
const docKey = (id: string) => 'wp:doc:' + id;

export interface IndexEntry {
  id: string;
  title: string;
  updatedAt: number;
}

/** localStorage throws when full; callers learn about it through this event. */
function announce(ok: boolean, error?: unknown): void {
  document.dispatchEvent(
    new CustomEvent('wp:saved', { detail: { ok, error } })
  );
}

export function save(doc: Doc): void {
  try {
    localStorage.setItem(docKey(doc.id), JSON.stringify(doc));
    const idx = readIndex().filter((e) => e.id !== doc.id);
    idx.unshift({ id: doc.id, title: doc.title, updatedAt: Date.now() });
    localStorage.setItem(INDEX_KEY, JSON.stringify(idx.slice(0, 50)));
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
  try {
    const id = localStorage.getItem(LAST_KEY) ?? readIndex()[0]?.id;
    if (!id) return null;
    const raw = localStorage.getItem(docKey(id));
    if (!raw) return null;
    return coerce(JSON.parse(raw));
  } catch {
    return null;
  }
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
    if (!isStyleId(rb.styleId)) {
      throw new Error(`Block ${i} has unknown styleId ${String(rb.styleId)}`);
    }
    return {
      id: typeof rb.id === 'string' && rb.id ? rb.id : newId(),
      styleId: rb.styleId,
      html: typeof rb.html === 'string' ? rb.html : '',
    };
  });

  return {
    id: typeof o.id === 'string' && o.id ? o.id : newId(),
    title: typeof o.title === 'string' && o.title ? o.title : 'Untitled',
    page: coercePage(o),
    blocks: blocks.length ? blocks : [{ id: newId(), styleId: 'Body', html: '' }],
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
