import type { StyleId } from './model';
import { STYLE_IDS } from './model';

/**
 * Named styles, defined once here and emitted as `.s-<StyleId>` CSS classes at
 * startup. Everything else (the style dropdown, the Enter key, the .docx
 * exporter) reads this table, so there is a single source of truth for what a
 * style means.
 */
export interface StyleDef {
  id: StyleId;
  label: string;
  /** Style of the block created when Enter is pressed at the end of this one. */
  enterTo: StyleId;
  /** Points. The document is specified in points, as print is. */
  size: number;
  bold: boolean;
  uppercase: boolean;
  /** px */
  letterSpacing: number;
  lineHeight: number;
  /** Points, [top, right, bottom, left]. Padding only - never margin. */
  padding: [number, number, number, number];
  /** Bottom rule, as on SectionHeading. */
  rule: boolean;
  /** Hanging indent in points; 0 for none. */
  hanging: number;
  bullet: boolean;

  /* --- paragraph properties, mirroring Word's --- */

  /** Never separated from the block below it by a page break. */
  keepWithNext: boolean;
  /** Never split across a page boundary at all. */
  keepLines: boolean;
  /** Minimum lines left at the bottom of a page when splitting. */
  orphanMin: number;
  /** Minimum lines carried to the next page when splitting. */
  widowMin: number;
  /** Always starts a new page. */
  pageBreakBefore: boolean;
}

export type ParagraphProps = Pick<
  StyleDef,
  'keepWithNext' | 'keepLines' | 'orphanMin' | 'widowMin' | 'pageBreakBefore'
>;

/** Applied to every style unless it overrides them. */
const PARAGRAPH_DEFAULTS: ParagraphProps = {
  keepWithNext: false,
  keepLines: false,
  orphanMin: 2,
  widowMin: 2,
  pageBreakBefore: false,
} as const;

/**
 * A point is 1/72in and a CSS pixel is 1/96in, so the document's sizes are
 * held in points - the unit print actually uses - and converted once, here.
 */
export const PT = 96 / 72;
export const px = (pt: number) => Math.round(pt * PT * 100) / 100;

/**
 * The document is set in a serif, per the Recto specification. Georgia is the
 * fallback because it is on every machine, so a font that has not arrived
 * yet never leaves the page unmeasurable.
 */
export const DOC_FONT = '"Source Serif 4", "Source Serif Pro", Georgia, Cambria, serif';
/** Font name written into the .docx. */
export const DOCX_FONT = 'Source Serif 4';
export const INK = '#111111';
export const RULE_COLOR = '#999999';

const BASE: Record<StyleId, Omit<StyleDef, keyof ParagraphProps>> = {
  Name: {
    id: 'Name',
    label: 'Name',
    enterTo: 'Contact',
    size: 24,
    bold: true,
    uppercase: false,
    letterSpacing: 0,
    lineHeight: 1.15,
    padding: [0, 0, 2, 0],
    rule: false,
    hanging: 0,
    bullet: false,
  },
  Contact: {
    id: 'Contact',
    label: 'Contact',
    enterTo: 'Body',
    size: 10,
    bold: false,
    uppercase: false,
    letterSpacing: 0,
    lineHeight: 1.3,
    padding: [0, 0, 12, 0],
    rule: false,
    hanging: 0,
    bullet: false,
  },
  SectionHeading: {
    id: 'SectionHeading',
    label: 'Section heading',
    enterTo: 'Body',
    size: 12,
    bold: true,
    uppercase: true,
    letterSpacing: 0.5,
    lineHeight: 1.2,
    padding: [10, 0, 4, 0],
    rule: true,
    hanging: 0,
    bullet: false,
  },
  JobTitle: {
    id: 'JobTitle',
    label: 'Job title',
    enterTo: 'Body',
    size: 11,
    bold: true,
    uppercase: false,
    letterSpacing: 0,
    lineHeight: 1.3,
    padding: [6, 0, 1, 0],
    rule: false,
    hanging: 0,
    bullet: false,
  },
  Body: {
    id: 'Body',
    label: 'Body',
    enterTo: 'Body',
    size: 11,
    bold: false,
    uppercase: false,
    letterSpacing: 0,
    lineHeight: 16 / 11, // Source Serif 11/16, per the specification
    padding: [0, 0, 4, 0],
    rule: false,
    hanging: 0,
    bullet: false,
  },
  Bullet: {
    id: 'Bullet',
    label: 'Bullet',
    enterTo: 'Bullet',
    size: 11,
    bold: false,
    uppercase: false,
    letterSpacing: 0,
    lineHeight: 16 / 11,
    padding: [0, 0, 2, 0],
    rule: false,
    hanging: 12,
    bullet: true,
  },
};

/** Per-style overrides of the paragraph defaults. */
const PARAGRAPH_OVERRIDES: Partial<Record<StyleId, Partial<ParagraphProps>>> = {
  // A heading left as the last line on a page is the most visible pagination
  // failure there is, so headings travel with the block beneath them.
  SectionHeading: { keepWithNext: true, keepLines: true },
  JobTitle: { keepWithNext: true, keepLines: true },
  // A name or a contact line split across a page break is never right.
  Name: { keepWithNext: true, keepLines: true },
  Contact: { keepLines: true },
};

export const STYLES: Record<StyleId, StyleDef> = Object.fromEntries(
  STYLE_IDS.map((id) => [
    id,
    { ...PARAGRAPH_DEFAULTS, ...BASE[id], ...(PARAGRAPH_OVERRIDES[id] ?? {}) },
  ])
) as Record<StyleId, StyleDef>;

export function styleDef(id: StyleId): StyleDef {
  return STYLES[id];
}

export function styleClass(id: StyleId): string {
  return 's-' + id;
}

/** Read the StyleId off a rendered block element, defaulting to Body. */
export function styleOf(el: Element): StyleId {
  for (const id of STYLE_IDS) {
    if (el.classList.contains(styleClass(id))) return id;
  }
  return 'Body';
}

function css(d: StyleDef): string {
  const [t, r, b, l] = d.padding;
  const lines: string[] = [];
  lines.push(`  font-size: ${px(d.size)}px;`);
  lines.push(`  font-weight: ${d.bold ? 700 : 400};`);
  lines.push(`  line-height: ${d.lineHeight};`);
  if (d.uppercase) lines.push('  text-transform: uppercase;');
  if (d.letterSpacing) lines.push(`  letter-spacing: ${px(d.letterSpacing)}px;`);
  // Spacing via padding only. Sibling margins collapse, and summed heights
  // would then disagree with the container height.
  lines.push(
    `  padding: ${px(t)}px ${px(r)}px ${px(b)}px ${px(l + d.hanging)}px;`
  );
  if (d.hanging) lines.push(`  text-indent: -${px(d.hanging)}px;`);
  if (d.rule) lines.push(`  border-bottom: 1px solid ${RULE_COLOR};`);
  let out = `.blk.${styleClass(d.id)} {\n${lines.join('\n')}\n}\n`;
  if (d.bullet) {
    out +=
      `.blk.${styleClass(d.id)}::before {\n` +
      `  content: "•";\n` +
      `  display: inline-block;\n` +
      `  width: ${d.hanging}px;\n` +
      `  text-indent: 0;\n` +
      `}\n`;
  }
  return out;
}

/**
 * A paragraph split across a page boundary renders as several elements. The
 * continuation must not repeat the top padding, the hanging indent or the list
 * marker, and every piece but the last must drop its bottom padding and rule -
 * otherwise a paragraph grows a little taller each time it is split.
 *
 * These come last in the sheet so they beat the `.s-*` rules at equal
 * specificity.
 */
/**
 * An imported list paragraph carries its resolved marker in data-marker, so
 * numbered clauses read as "1." rather than as bullets. Levels indent by
 * 18px, on top of the hanging indent.
 */
const LIST_CSS =
  '.blk[data-marker]::before {\n' +
  '  content: attr(data-marker);\n' +
  '  display: inline-block;\n' +
  '  min-width: 14px;\n' +
  '  padding-right: 4px;\n' +
  '  text-indent: 0;\n' +
  '}\n' +
  '.blk.split-cont[data-marker]::before { content: none; }\n' +
  [1, 2, 3, 4, 5, 6, 7, 8]
    .map(
      (n) =>
        `.blk[data-level="${n}"] { padding-left: ${14 + n * 18}px; }\n`
    )
    .join('');

/**
 * Tables. Widths come from the document's own w:tblGrid, so the layout is
 * fixed rather than content-driven - a browser's automatic table layout would
 * silently disagree with what Word measured and change where rows break.
 */
const TABLE_CSS =
  '.blk-table { margin: 0; padding: 0 0 4px; }\n' +
  '.blk-table table {\n' +
  '  border-collapse: collapse;\n' +
  '  table-layout: fixed;\n' +
  '  width: 100%;\n' +
  '}\n' +
  '.blk-table td {\n' +
  '  border: 1px solid #bbbbbb;\n' +
  '  padding: 3px 5px;\n' +
  '  vertical-align: top;\n' +
  '  overflow-wrap: break-word;\n' +
  '}\n' +
  '.blk-table tr.hdr td { background: #f3f3f3; }\n' +
  '.blk-table td.merged { border-left: 0; }\n' +
  '.blk-table .blk:last-child { padding-bottom: 0; }\n';

/**
 * Inline images. Capped to the text column so an oversized picture cannot
 * push past the margin, and vertical-align keeps the line box honest so
 * measurement sees the height the image actually occupies.
 */
const IMAGE_CSS =
  '.blk img {\n' +
  '  max-width: 100%;\n' +
  '  height: auto;\n' +
  '  vertical-align: bottom;\n' +
  '}\n' +
  '.blk img.img-missing {\n' +
  '  display: inline-block;\n' +
  '  min-width: 32px;\n' +
  '  min-height: 32px;\n' +
  '  background: repeating-linear-gradient(45deg, #eee, #eee 6px, #e0e0e0 6px, #e0e0e0 12px);\n' +
  '  outline: 1px solid #ccc;\n' +
  '}\n';

const SPLIT_CSS =
  '.blk.split-cont { padding-top: 0; text-indent: 0; }\n' +
  '.blk.split-cont::before { content: none; }\n' +
  '.blk.split-more { padding-bottom: 0; border-bottom: none; }\n';

/** Inject the `.s-*` rules. Called once at startup, before first measurement. */
export function injectStyleSheet(): void {
  const id = 'wp-named-styles';
  if (document.getElementById(id)) return;
  const el = document.createElement('style');
  el.id = id;
  el.textContent =
    `.blk { font-family: ${DOC_FONT}; color: ${INK}; }\n` +
    STYLE_IDS.map((s) => css(STYLES[s])).join('') +
    LIST_CSS +
    TABLE_CSS +
    IMAGE_CSS +
    SPLIT_CSS;
  document.head.appendChild(el);
}
