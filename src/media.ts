/**
 * Image bytes, held outside the document model.
 *
 * An imported image already exists in the package, so the model refers to it
 * by its part path rather than carrying a copy. Putting data URIs in the
 * model instead would double the storage and exhaust localStorage on the
 * second document with a photograph in it.
 *
 * Object URLs are minted lazily and revoked when the document changes, so a
 * long session does not accumulate one per image per reload.
 */

const urls = new Map<string, string>();
/** Bytes for images added in the editor, which have no package to read. */
const added = new Map<string, Uint8Array>();

const TYPES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/emf',
  wmf: 'image/wmf',
};

function mimeFor(path: string): string {
  const ext = path.split('.').pop()?.toLowerCase() ?? '';
  return TYPES[ext] ?? 'application/octet-stream';
}

/** Formats a browser cannot draw; the document keeps them, we just cannot show them. */
export function isDrawable(path: string): boolean {
  const ext = path.split('.').pop()?.toLowerCase() ?? '';
  return ext !== 'emf' && ext !== 'wmf' && ext !== 'tif' && ext !== 'tiff';
}

/** Point the registry at a freshly imported package's parts. */
export function registerMedia(parts: Map<string, Uint8Array>): void {
  clearMedia();
  for (const [path, bytes] of parts) {
    if (!path.startsWith('word/media/')) continue;
    if (!isDrawable(path)) continue;
    const blob = new Blob([bytes as BlobPart], { type: mimeFor(path) });
    urls.set(path, URL.createObjectURL(blob));
  }
}

export function mediaUrl(path: string): string | null {
  return urls.get(path) ?? null;
}

/**
 * Add one image the user inserted.
 *
 * Kept beside the imported parts under the same `word/media/` prefix, so a
 * document that came from a package and one that did not are drawn, stored
 * and exported by exactly the same code.
 */
export function addMedia(path: string, bytes: Uint8Array, mime?: string): string {
  added.set(path, bytes);
  try {
    const blob = new Blob([bytes as BlobPart], { type: mime ?? mimeFor(path) });
    const old = urls.get(path);
    if (old) URL.revokeObjectURL(old);
    urls.set(path, URL.createObjectURL(blob));
  } catch {
    // No object URLs outside a browser. The bytes are what the export
    // needs; the URL is only ever used to draw the picture on screen.
  }
  return path;
}

/** The extension OOXML expects for a media part of this type. */
export function extensionFor(mime: string): string {
  for (const [ext, type] of Object.entries(TYPES)) {
    if (type === mime) return ext;
  }
  return 'png';
}

export function clearMedia(): void {
  for (const url of urls.values()) URL.revokeObjectURL(url);
  urls.clear();
  added.clear();
}

/** The bytes of an image added in this session, if it was added here. */
export function mediaBytes(path: string): Uint8Array | null {
  return added.get(path) ?? null;
}

export function addedMedia(): Map<string, Uint8Array> {
  return added;
}
