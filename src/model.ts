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

/* ------------------------------------------------------------------ *
 * Direct character formatting
 *
 * 14,865 of 18,906 runs in a 63-document corpus carry a size, a font or a
 * colour of their own - 79%, against 47% for the paragraph-level equivalent.
 * Worse than not drawing it: the vault kept ONE base w:rPr per paragraph and
 * wrote it onto every regenerated run, so editing a heading with one
 * coloured word in it flattened the whole line to the first run's colour.
 *
 * A run style is stored as the DIFFERENCE from the paragraph's base run
 * properties, which is what keeps the markup lean when a document sets the
 * same font on all 400 of its runs.
 * ------------------------------------------------------------------ */

export interface RunStyle {
  /** Points. */
  size?: number;
  /** Family name as the document spells it. */
  font?: string;
  /** Six hex digits, no leading hash. */
  color?: string;
  strike?: boolean;
  vert?: 'super' | 'sub';
}

export const RUN_STYLE_KEYS: (keyof RunStyle)[] = [
  'size',
  'font',
  'color',
  'strike',
  'vert',
];

export function hasRunStyle(r: RunStyle | undefined): r is RunStyle {
  return !!r && RUN_STYLE_KEYS.some((k) => r[k] !== undefined);
}

export function sameRunStyle(a: RunStyle | undefined, b: RunStyle | undefined): boolean {
  if (!hasRunStyle(a) && !hasRunStyle(b)) return true;
  if (!hasRunStyle(a) || !hasRunStyle(b)) return false;
  return RUN_STYLE_KEYS.every((k) => a[k] === b[k]);
}

export function tidyRunStyle(r: RunStyle): RunStyle | undefined {
  const out: RunStyle = {};
  for (const k of RUN_STYLE_KEYS) {
    const v = r[k];
    if (v !== undefined && v !== null) (out as Record<string, unknown>)[k] = v;
  }
  return hasRunStyle(out) ? out : undefined;
}

/** What `b` says over and above `a`; undefined when they agree. */
export function runStyleDiff(
  base: RunStyle | undefined,
  run: RunStyle | undefined
): RunStyle | undefined {
  const out: RunStyle = {};
  for (const k of RUN_STYLE_KEYS) {
    const want = run?.[k];
    const have = base?.[k];
    if (want !== have && want !== undefined) (out as Record<string, unknown>)[k] = want;
  }
  return hasRunStyle(out) ? out : undefined;
}

/* ------------------------------------------------------------------ *
 * Direct paragraph formatting
 *
 * The six named styles are the vocabulary this program writes in, but they
 * are not the vocabulary the world sends. Across 63 real documents, 1,362 of
 * 2,913 paragraphs carry an alignment, an indent or a paragraph spacing of
 * their own. We preserved all of it on export and drew none of it, so an
 * imported document was faithful in the file and wrong on the screen - and
 * our page breaks landed in different places than Word's.
 *
 * Held in POINTS, like the style table, for the same reason: points are the
 * unit print uses, and the conversion to CSS pixels happens once.
 * ------------------------------------------------------------------ */

export type BlockAlign = 'left' | 'center' | 'right' | 'justify';

export interface BlockFormat {
  align?: BlockAlign;
  /** Points, from the left and right margins. */
  indentLeft?: number;
  indentRight?: number;
  /** Points, signed: negative is a hanging indent, as Word's w:hanging is. */
  firstLine?: number;
  /** Points of space above and below, added to the style's own padding. */
  spaceBefore?: number;
  spaceAfter?: number;
  /** A multiplier when lineRule is 'auto', otherwise points. */
  lineHeight?: number;
  lineRule?: 'auto' | 'exact' | 'atLeast';
}

const FORMAT_KEYS: (keyof BlockFormat)[] = [
  'align',
  'indentLeft',
  'indentRight',
  'firstLine',
  'spaceBefore',
  'spaceAfter',
  'lineHeight',
  'lineRule',
];

/** False for undefined and for an object that says nothing. */
export function hasFormat(f: BlockFormat | undefined): f is BlockFormat {
  return !!f && FORMAT_KEYS.some((k) => f[k] !== undefined);
}

export function sameFormat(a: BlockFormat | undefined, b: BlockFormat | undefined): boolean {
  if (!hasFormat(a) && !hasFormat(b)) return true;
  if (!hasFormat(a) || !hasFormat(b)) return false;
  return FORMAT_KEYS.every((k) => a[k] === b[k]);
}

/** Drops the keys that say nothing, so an empty format is never stored. */
export function tidyFormat(f: BlockFormat): BlockFormat | undefined {
  const out: BlockFormat = {};
  for (const k of FORMAT_KEYS) {
    const v = f[k];
    if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  }
  // A line rule with no measurement is noise.
  if (out.lineHeight === undefined) delete out.lineRule;
  return hasFormat(out) ? out : undefined;
}

export interface ParagraphBlock {
  /** Optional so every paragraph written before tables existed still parses. */
  kind?: 'para';
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

  /**
   * Alignment, indents and spacing the document asked for directly, rather
   * than through its style. Absent on anything this program wrote itself:
   * a paragraph with no direct formatting is the named style, unqualified.
   */
  fmt?: BlockFormat;
}

/**
 * A cell holds an array of paragraphs, so the existing renderer, styles and
 * caret code work inside a cell unchanged. Nested tables are out of scope.
 */
export type TableCell = ParagraphBlock[];

export interface TableRow {
  id: string;
  /** Repeated at the top of each page the table continues onto. */
  headerRow: boolean;
  cells: TableCell[];
}

export interface TableBlock {
  kind: 'table';
  id: string;
  /** px, summing to the content width. */
  cols: { width: number }[];
  rows: TableRow[];
}

export type Block = ParagraphBlock | TableBlock;

/* ------------------------------------------------------------------ *
 * Headers and footers
 *
 * Word keeps up to three of each: one for the first page when the document
 * has a title page, one for even pages when odd and even differ, and one for
 * everything else.
 * ------------------------------------------------------------------ */

export type HFVariant = 'default' | 'first' | 'even';
export type HeaderFooterSet = Partial<Record<HFVariant, ParagraphBlock[]>>;

export interface Doc2Extras {
  headers?: HeaderFooterSet;
  footers?: HeaderFooterSet;
  /** w:titlePg - the first page uses its own header and footer. */
  titlePage?: boolean;
  /** w:evenAndOddHeaders - even pages use their own. */
  evenOdd?: boolean;
  /** px from the page edge to the header and footer, from w:pgMar. */
  headerDistance?: number;
  footerDistance?: number;
  /**
   * The family an imported document sets for text that names none of its
   * own. Absent on documents this program wrote, which are set in its own
   * serif - the point of carrying it is that an imported document should
   * look like itself, and break its lines where Word breaks them.
   */
  defaultFont?: string;
  /**
   * Sections, when the document has more than one. Absent means one section
   * with `page`, `headers` and `footers`, which is what most documents are
   * and what everything written before sections existed still deserializes as.
   */
  sections?: Section[];
}

/* ------------------------------------------------------------------ *
 * Sections
 *
 * A Word document is a sequence of sections, each with its own page size,
 * margins and headers. Reading only the last one - which is what a single
 * `page` field amounts to - lays the whole document out with the geometry of
 * its final few paragraphs. On real resumes that is a visible error: a body
 * set to a half-inch margin rendered at the trailing section's inch.
 * ------------------------------------------------------------------ */

export interface Section {
  id: string;
  /**
   * The block this section starts at. Null for the first section, which
   * starts at the top of the document whatever the first block turns out to
   * be. An id rather than an index, because the blocks around it are edited.
   */
  startId: string | null;
  page: PageSetup;
  headers?: HeaderFooterSet;
  footers?: HeaderFooterSet;
  titlePage?: boolean;
  /** w:type continuous - shares a page with the section before it. */
  continuous?: boolean;
}

/** Always at least one, so callers never branch on whether sections exist. */
export function sectionsOf(doc: Doc): Section[] {
  if (doc.sections && doc.sections.length > 0) return doc.sections;
  return [
    {
      id: 'only',
      startId: null,
      page: doc.page,
      headers: doc.headers,
      footers: doc.footers,
      titlePage: doc.titlePage,
    },
  ];
}

/**
 * Which section each block belongs to, by walking the blocks in order.
 *
 * Resolved by walking rather than stored, so a section whose first block was
 * deleted simply merges into the one before it instead of leaving the
 * document pointing at a block that is not there.
 */
export function sectionIndexByBlock(doc: Doc): Map<string, number> {
  const sections = sectionsOf(doc);
  const starts = new Map<string, number>();
  sections.forEach((s, i) => {
    if (s.startId) starts.set(s.startId, i);
  });
  const out = new Map<string, number>();
  let current = 0;
  for (const b of doc.blocks) {
    const at = starts.get(b.id);
    if (at !== undefined) current = at;
    out.set(b.id, current);
  }
  return out;
}

/** Which header and footer a given page uses. */
export function hfVariant(doc: Doc, pageIndex: number, section?: Section): HFVariant {
  const title = section ? section.titlePage : doc.titlePage;
  if (pageIndex === 0 && title) return 'first';
  if (doc.evenOdd && (pageIndex + 1) % 2 === 0) return 'even';
  return 'default';
}

/** Falls back the way Word does: a missing variant uses the default one. */
export function hfFor(
  set: HeaderFooterSet | undefined,
  variant: HFVariant
): ParagraphBlock[] {
  if (!set) return [];
  return set[variant] ?? set.default ?? [];
}

/**
 * The readable text of a block's inline markup.
 *
 * Stripping the tags is only half of it: the markup holds `&amp;` and
 * `&rsquo;`, and anything that shows this text to a person - the outline, a
 * document title, a word count - has to decode them too. Leaving that out
 * puts a literal "&amp;" in the sidebar next to a paragraph that renders a
 * perfectly good ampersand.
 */
const ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
  rsquo: '’',
  lsquo: '‘',
  rdquo: '”',
  ldquo: '“',
  ndash: '–',
  mdash: '—',
  hellip: '…',
};

export function decodeEntities(s: string): string {
  return s.replace(/&(#x?[0-9a-f]+|[a-z]+);/gi, (whole, body: string) => {
    if (body[0] === '#') {
      const code =
        body[1] === 'x' || body[1] === 'X'
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      return Number.isFinite(code) && code > 0 ? String.fromCodePoint(code) : whole;
    }
    // Ampersand last, so "&amp;lt;" decodes to "&lt;" and not to "<".
    return ENTITIES[body.toLowerCase()] ?? whole;
  });
}

export function plainText(html: string): string {
  return decodeEntities(html.replace(/<[^>]*>/g, ''));
}

export function isTable(b: Block): b is TableBlock {
  return (b as TableBlock).kind === 'table';
}

export function isParagraph(b: Block): b is ParagraphBlock {
  return !isTable(b);
}

/** Every paragraph in a document, including the ones inside table cells. */
export function paragraphsOf(blocks: Block[]): ParagraphBlock[] {
  const out: ParagraphBlock[] = [];
  for (const b of blocks) {
    if (isTable(b)) {
      for (const row of b.rows) for (const cell of row.cells) out.push(...cell);
    } else {
      out.push(b);
    }
  }
  return out;
}

export interface Doc extends Doc2Extras {
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

export function newBlock(styleId: StyleId, html = ''): ParagraphBlock {
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
