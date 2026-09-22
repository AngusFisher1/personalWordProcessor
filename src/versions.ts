import type { Doc } from './model';
import { plainText } from './model';
import { isTable } from './model';

/**
 * Version history.
 *
 * The whole library is around 600 KB of model. Twenty-five versions of every
 * document costs a few megabytes, which IndexedDB gives away - so the
 * question is not whether it is affordable but why autosave ever meant
 * "overwrite the only copy".
 *
 * Deliberately NOT in localStorage: that is where the documents themselves
 * live, and filling it with history would make the thing it protects fail to
 * save. Losing history is survivable; failing to save is not.
 */

const DB_NAME = 'wp-versions';
const STORE = 'versions';
const KEEP = 25;

/** A version is only worth keeping if it differs from the one before it. */
export interface Version {
  /** `${docId}:${savedAt}` - the key, and the sort order. */
  key: string;
  docId: string;
  savedAt: number;
  words: number;
  blocks: number;
  /** The whole document, as stored. */
  json: string;
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(STORE)) {
        const store = req.result.createObjectStore(STORE, { keyPath: 'key' });
        store.createIndex('docId', 'docId', { unique: false });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function withStore<T>(
  mode: IDBTransactionMode,
  fn: (store: IDBObjectStore) => IDBRequest<T>
): Promise<T> {
  return openDb().then(
    (db) =>
      new Promise<T>((resolve, reject) => {
        const t = db.transaction(STORE, mode);
        const req = fn(t.objectStore(STORE));
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
        t.oncomplete = () => db.close();
      })
  );
}

function wordCount(doc: Doc): number {
  let text = '';
  for (const b of doc.blocks) {
    if (isTable(b)) {
      for (const row of b.rows) {
        for (const cell of row.cells) for (const p of cell) text += ' ' + p.html;
      }
    } else {
      text += ' ' + b.html;
    }
  }
  const t = plainText(text).trim();
  return t === '' ? 0 : t.split(/\s+/).length;
}

/**
 * Keep a version of this document, unless nothing has changed since the last
 * one. Failure is silent by design: history is a luxury on top of the save,
 * and a full or unavailable IndexedDB must never stop the document saving.
 */
export async function keepVersion(doc: Doc): Promise<void> {
  try {
    const json = JSON.stringify(doc);
    const existing = await listVersions(doc.id);
    if (existing[0]?.json === json) return;

    const savedAt = Date.now();
    await withStore('readwrite', (s) =>
      s.put({
        key: doc.id + ':' + savedAt,
        docId: doc.id,
        savedAt,
        words: wordCount(doc),
        blocks: doc.blocks.length,
        json,
      } satisfies Version)
    );

    // Oldest first out. A document edited all day should not push its own
    // morning out of history and then fill the store with the afternoon.
    const all = await listVersions(doc.id);
    for (const old of all.slice(KEEP)) {
      await withStore('readwrite', (s) => s.delete(old.key));
    }
  } catch {
    /* no history this time */
  }
}

/** Newest first. */
export async function listVersions(docId: string): Promise<Version[]> {
  try {
    const all = await withStore<Version[]>('readonly', (s) =>
      s.index('docId').getAll(docId)
    );
    return all.sort((a, b) => b.savedAt - a.savedAt);
  } catch {
    return [];
  }
}

export async function forgetVersions(docId: string): Promise<void> {
  try {
    for (const v of await listVersions(docId)) {
      await withStore('readwrite', (s) => s.delete(v.key));
    }
  } catch {
    /* nothing to do */
  }
}

/** How much history is on disk, for the library footer. */
export async function versionBytes(): Promise<number> {
  try {
    const all = await withStore<Version[]>('readonly', (s) => s.getAll());
    return all.reduce((n, v) => n + v.json.length, 0);
  } catch {
    return 0;
  }
}

/** "4 min ago", "yesterday", "12 Mar". */
export function whenLabel(ts: number): string {
  const secs = Math.max(0, Date.now() - ts) / 1000;
  if (secs < 90) return 'just now';
  const mins = secs / 60;
  if (mins < 60) return Math.round(mins) + ' min ago';
  const hrs = mins / 60;
  if (hrs < 24) return Math.round(hrs) + ' hr ago';
  const days = hrs / 24;
  if (days < 2) return 'yesterday';
  if (days < 7) return Math.round(days) + ' days ago';
  return new Date(ts).toLocaleDateString(undefined, { day: 'numeric', month: 'short' });
}
