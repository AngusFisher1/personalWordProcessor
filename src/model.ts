// Geometry. CSS pixels at 96px/inch. These are exact; do not approximate.
export const PX_PER_INCH = 96;
export const PAGE_W = 816; // 8.5in
export const PAGE_H = 1056; // 11in

export const MARGINS = {
  narrow: 48, // 0.5in - default
  normal: 96, // 1in
} as const;

export type MarginKey = keyof typeof MARGINS;

export interface Margins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

/**
 * Page geometry is per document, not a global constant: an imported .docx
 * carries its own size and margins in w:sectPr, and a letter-only importer
 * fails on the first legal-size or landscape document.
 */
export interface PageSetup {
  width: number;
  height: number;
  margins: Margins;
}

export function uniformMargins(px: number): Margins {
  return { top: px, right: px, bottom: px, left: px };
}

export function pageSetup(margin: MarginKey): PageSetup {
  return {
    width: PAGE_W,
    height: PAGE_H,
    margins: uniformMargins(MARGINS[margin]),
  };
}

// content box, derived:
//   narrow -> 720 x 960
//   normal -> 624 x 864
export function contentWidth(p: PageSetup): number {
  return p.width - p.margins.left - p.margins.right;
}
export function contentHeight(p: PageSetup): number {
  return p.height - p.margins.top - p.margins.bottom;
}

export function isLandscape(p: PageSetup): boolean {
  return p.width > p.height;
}

/** Which preset, if the margins happen to match one. */
export function marginPreset(p: PageSetup): MarginKey | null {
  const m = p.margins;
  if (m.top !== m.right || m.right !== m.bottom || m.bottom !== m.left) return null;
  for (const k of Object.keys(MARGINS) as MarginKey[]) {
    if (MARGINS[k] === m.top) return k;
  }
  return null;
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

  /**
   * List marker text resolved from numbering.xml on import ("2.", "iv.", a
   * bullet glyph). Held out of `html` so it never lands in the text the user
   * edits, and ignored on .docx export, where the paragraph's own w:numPr
   * makes Word renumber the list itself.
   */
  listMarker?: string;
  /** Indent level of a list paragraph, 0-based. */
  listLevel?: number;
}

export interface Doc {
  id: string;
  title: string;
  page: PageSetup;
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
    page: pageSetup('narrow'),
    blocks: [newBlock('Body')],
  };
}
