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

export function clearMedia(): void {
  for (const url of urls.values()) URL.revokeObjectURL(url);
  urls.clear();
}
