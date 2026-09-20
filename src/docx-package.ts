import JSZip from 'jszip';

/**
 * The preservation vault.
 *
 * An imported .docx is kept in memory exactly as it arrived. On export we
 * rewrite only the body of word/document.xml and drop it back into the
 * original package, rather than generating a fresh one. Everything we never
 * understood - theme, fonts, custom XML, document properties, comments,
 * headers, footnotes - survives untouched.
 *
 * Without this, exporting a colleague's document silently strips their theme
 * and metadata, and the round-trip criterion can never pass.
 */

export const DOC_XML = 'word/document.xml';
export const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export interface Warning {
  kind: string;
  detail: string;
  count: number;
}

/** A body child we preserve verbatim but do not render or allow editing. */
export interface OpaqueEntry {
  xml: string;
  /** Re-emitted immediately after this block; null means at the top. */
  afterBlockId: string | null;
  kind: string;
}

/** One w:rPr child, kept by name so it can be re-inserted in schema order. */
export interface RunProp {
  name: string;
  xml: string;
}

export interface Vault {
  /** Every part of the original package, as stored. */
  parts: Map<string, Uint8Array>;
  /** word/document.xml up to and including the <w:body> open tag. */
  docXmlPrefix: string;
  /** word/document.xml from </w:body> onwards. */
  docXmlSuffix: string;
  /** Original paragraph XML per block id, for blocks that were not edited. */
  blockXml: Map<string, string>;
  /** Original <w:pPr> per block id, so an edited block keeps its properties. */
  blockPPr: Map<string, string>;
  /**
   * The run properties of each block's first run, minus the bold/italic/
   * underline we manage ourselves. Real documents carry their heading
   * formatting on the runs rather than in a style, so regenerating an edited
   * paragraph without this resets its size and font to the document default.
   */
  blockRPr: Map<string, RunProp[]>;
  /** The html each block had at import, to detect whether it was edited. */
  blockHtml: Map<string, string>;
  /** The StyleId each block had at import. */
  blockStyle: Map<string, string>;
  /** Our StyleId back to a w:styleId that exists in this package. */
  styleBack: Map<string, string>;
  /** Hyperlink target back to its relationship id, so links survive an edit. */
  relByTarget: Map<string, string>;
  /**
   * Whole runs kept verbatim under a token, for content we render but cannot
   * regenerate - an image, whose w:drawing carries far more than a src.
   */
  runXml: Map<string, string>;
  /** Per table: its w:tblPr and w:tblGrid, so an edited table keeps them. */
  tablePr: Map<string, { tblPr: string; tblGrid: string }>;
  /** Per row id: its w:trPr. */
  rowPr: Map<string, string>;
  /** Per "<rowId>c<index>": that cell's w:tcPr, which carries any gridSpan. */
  cellPr: Map<string, string>;
  opaque: OpaqueEntry[];
  /** Body-level <w:sectPr>, which must stay last in the body. */
  sectPrXml: string | null;
  warnings: Warning[];
}

export function emptyVault(): Vault {
  return {
    parts: new Map(),
    docXmlPrefix: '',
    docXmlSuffix: '',
    blockXml: new Map(),
    blockPPr: new Map(),
    blockRPr: new Map(),
    blockHtml: new Map(),
    blockStyle: new Map(),
    styleBack: new Map(),
    relByTarget: new Map(),
    runXml: new Map(),
    tablePr: new Map(),
    rowPr: new Map(),
    cellPr: new Map(),
    opaque: [],
    sectPrXml: null,
    warnings: [],
  };
}

export function addWarning(v: Vault, kind: string, detail: string): void {
  const hit = v.warnings.find((w) => w.kind === kind);
  if (hit) {
    hit.count++;
    return;
  }
  v.warnings.push({ kind, detail, count: 1 });
}

/* ------------------------------------------------------------------ *
 * Zip
 * ------------------------------------------------------------------ */

export async function unzip(data: ArrayBuffer): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(data);
  const out = new Map<string, Uint8Array>();
  const names = Object.keys(zip.files);
  for (const name of names) {
    const entry = zip.files[name];
    if (entry.dir) continue;
    out.set(name, await entry.async('uint8array'));
  }
  return out;
}

export function partText(parts: Map<string, Uint8Array>, name: string): string | null {
  const bytes = parts.get(name);
  if (!bytes) return null;
  return new TextDecoder('utf-8').decode(bytes);
}

/**
 * Rebuild the package with a new body. Every other part is written back from
 * the vault; only word/document.xml is regenerated, and even that keeps its
 * original prologue so the namespace declarations are exactly as they were.
 */
export async function repack(vault: Vault, bodyInner: string): Promise<Blob> {
  const zip = new JSZip();
  // [Content_Types].xml must be the first entry for some consumers.
  const order = Array.from(vault.parts.keys()).sort((a, b) => {
    if (a === '[Content_Types].xml') return -1;
    if (b === '[Content_Types].xml') return 1;
    return 0;
  });
  for (const name of order) {
    if (name === DOC_XML) continue;
    zip.file(name, vault.parts.get(name) as Uint8Array, { binary: true });
  }
  zip.file(DOC_XML, vault.docXmlPrefix + bodyInner + vault.docXmlSuffix);
  return zip.generateAsync({
    type: 'blob',
    mimeType: DOCX_MIME,
    compression: 'DEFLATE',
  });
}

/* ------------------------------------------------------------------ *
 * Keeping the original across a reload
 *
 * The vault is the original bytes, which are far too big for localStorage, so
 * they live in IndexedDB. Imported block ids are derived from the body index,
 * which means re-parsing the stored original on load rebuilds a vault whose
 * keys still match the edited blocks that came back from localStorage.
 * ------------------------------------------------------------------ */

const DB_NAME = 'wp-docx';
const STORE = 'originals';

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(STORE)) {
        req.result.createObjectStore(STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function tx<T>(
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

export async function saveOriginal(docId: string, bytes: ArrayBuffer): Promise<void> {
  try {
    await tx('readwrite', (s) => s.put(bytes, docId));
  } catch {
    /* no IndexedDB: the vault simply will not survive a reload */
  }
}

export async function loadOriginal(docId: string): Promise<ArrayBuffer | null> {
  try {
    const v = await tx<ArrayBuffer | undefined>('readonly', (s) => s.get(docId));
    return v ?? null;
  } catch {
    return null;
  }
}

export async function deleteOriginal(docId: string): Promise<void> {
  try {
    await tx('readwrite', (s) => s.delete(docId));
  } catch {
    /* nothing to clean up */
  }
}
