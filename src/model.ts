// Geometry. CSS pixels at 96px/inch. These are exact; do not approximate.
export const PX_PER_INCH = 96;
export const PAGE_W = 816; // 8.5in
export const PAGE_H = 1056; // 11in

export const MARGINS = {
  narrow: 48, // 0.5in - default
  normal: 96, // 1in
} as const;

export type MarginKey = keyof typeof MARGINS;

// content box, derived:
//   narrow -> 720 x 960
//   normal -> 624 x 864
export function contentWidth(margin: MarginKey): number {
  return PAGE_W - 2 * MARGINS[margin];
}
export function contentHeight(margin: MarginKey): number {
  return PAGE_H - 2 * MARGINS[margin];
}

export type StyleId =
  | 'Name'
  | 'Contact'
  | 'SectionHeading'
  | 'JobTitle'
  | 'Body'
  | 'Bullet';

export const STYLE_IDS: readonly StyleId[] = [
  'Name',
  'Contact',
  'SectionHeading',
  'JobTitle',
  'Body',
  'Bullet',
];

export function isStyleId(x: unknown): x is StyleId {
  return typeof x === 'string' && (STYLE_IDS as readonly string[]).includes(x);
}

export interface Block {
  id: string;
  styleId: StyleId;
  /** Inline markup only: b, i, u, a, br. Anything else is stripped on the way in. */
  html: string;
  /**
   * Set on the continuation half of a block that pagination split across a
   * page boundary, pointing at the id of the block it continues.
   *
   * A pure render artifact: readModel merges continuations back before save,
   * export or an undo snapshot, so the stored document always has exactly one
   * block per logical paragraph and nothing downstream has to know that
   * splitting exists.
   */
  continuesFrom?: string;
}

export interface Doc {
  id: string;
  title: string;
  margin: MarginKey;
  blocks: Block[];
}

// The model is flat and ordered. There is deliberately no `page` field on a
// Block and there must never be one: pages are a rendering outcome, recomputed
// from measurement, never stored.

export function newId(): string {
  const c = globalThis.crypto as Crypto | undefined;
  if (c && typeof c.randomUUID === 'function') return c.randomUUID();
  // Fallback for non-secure contexts.
  return 'id-' + Math.random().toString(36).slice(2) + Date.now().toString(36);
}

export function newBlock(styleId: StyleId, html = ''): Block {
  return { id: newId(), styleId, html };
}

export function emptyDoc(): Doc {
  return {
    id: newId(),
    title: 'Untitled',
    margin: 'narrow',
    blocks: [newBlock('Body')],
  };
}
