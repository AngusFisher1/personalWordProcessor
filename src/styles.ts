import type { BlockFormat, StyleId } from './model';
import { STYLE_IDS, hasFormat } from './model';

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

/** The table as shipped, before any document says otherwise. */
const SHIPPED: Record<StyleId, StyleDef> = Object.fromEntries(
  STYLE_IDS.map((id) => [
    id,
    { ...PARAGRAPH_DEFAULTS, ...BASE[id], ...(PARAGRAPH_OVERRIDES[id] ?? {}) },
  ])
) as Record<StyleId, StyleDef>;

/**
 * The style table in force.
 *
 * Mutated in place rather than replaced, because everything in the program
 * reads `STYLES[id]` directly - the CSS emitter, the style menu, the Enter
 * key, the paginator's keep-with-next rules and both exporters. One table
 * that changes is the whole feature; a second table would be a second
 * source of truth to get out of step.
 */
export const STYLES: Record<StyleId, StyleDef> = Object.fromEntries(
  STYLE_IDS.map((id) => [id, { ...SHIPPED[id] }])
) as Record<StyleId, StyleDef>;

/** What a document changes about the shipped styles. */
export type StyleOverrides = Partial<Record<StyleId, Partial<StyleDef>>>;

/** The fields a document may override; the rest are structural. */
export const EDITABLE_STYLE_FIELDS = [
  'size',
  'bold',
  'uppercase',
  'letterSpacing',
  'lineHeight',
  'rule',
  'hanging',
  'padding',
  'keepWithNext',
  'keepLines',
  'pageBreakBefore',
  'widowMin',
  'orphanMin',
] as const;

export type EditableStyleField = (typeof EDITABLE_STYLE_FIELDS)[number];

export function shippedStyle(id: StyleId): StyleDef {
  return SHIPPED[id];
}

/**
 * Put a document's style overrides in force and redraw the stylesheet.
 *
 * Every cached height was measured against the old table, so the caller
 * clears the height cache and reflows - which is why this returns rather
 * than doing it: styles.ts must not know the paginator exists.
 */
export function setStyleOverrides(o: StyleOverrides | undefined): void {
  for (const id of STYLE_IDS) {
    const over = o?.[id] ?? {};
    // Assigned field by field into the existing object so that anything
    // holding a reference to STYLES[id] sees the change.
    Object.assign(STYLES[id], SHIPPED[id], over);
  }
  emitStyleSheet();
}

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

/**
 * Find highlights. The current match is solid and the rest are tinted, so
 * the eye can tell where it is without losing sight of the others. These are
 * plain spans the sanitizer unwraps, so a highlight can never be saved.
 */
const FIND_CSS =
  '.blk span.find-hit { background: rgba(217, 123, 60, .28); }\n' +
  '.blk span.find-hit.on { background: var(--acc); color: var(--accfg); }\n';

/**
 * Headers and footers sit inside the page's margin, above and below the body
 * box. They are part of the paper, so unlike the rest of the chrome they do
 * print - which is the whole point of them.
 */
const HF_CSS =
  '.page-header, .page-footer {\n' +
  '  overflow: hidden;\n' +
  '  color: #111111;\n' +
  '}\n' +
  '.page-header .blk, .page-footer .blk { padding-top: 0; padding-bottom: 0; }\n' +
  '.page-header .blk:last-child { padding-bottom: 2pt; }\n' +
  '.page-footer .blk:first-child { padding-top: 2pt; }\n';

const SPLIT_CSS =
  '.blk.split-cont { padding-top: 0; text-indent: 0; }\n' +
  '.blk.split-cont::before { content: none; }\n' +
  '.blk.split-more { padding-bottom: 0; border-bottom: none; }\n';

/** Inject the `.s-*` rules. Called once at startup, before first measurement. */
/**
 * What Word means by one line.
 *
 * `w:lineRule="auto"` counts in multiples of a single line, and Word derives
 * a single line from the font's own ascent, descent and line gap - about
 * 1.2 times the point size for a text face. CSS `line-height: 1.5` means 1.5
 * times the FONT SIZE, which is a fifth tighter than the same document in
 * Word. Applying the factor here keeps a document that says "1.5 lines"
 * looking like itself, while the model goes on storing Word's own number so
 * the round trip stays exact.
 */
/**
 * Open faces drawn to the same metrics as the Microsoft ones.
 *
 * Substituting a font of different metrics moves every line break, which is
 * the difference between a page count that matches Word's and one that does
 * not. These four are designed as drop-in metric equivalents; where one is
 * installed it is a far better second choice than our own serif.
 *
 * Aptos, Word's current default, has no open twin - a document set in it on
 * a machine without it cannot be paginated the way Word paginates it, and
 * the harness reports that rather than hiding it.
 */
const METRIC_TWINS: Record<string, string> = {
  calibri: 'Carlito',
  cambria: 'Caladea',
  'times new roman': 'Liberation Serif',
  arial: 'Liberation Sans',
  helvetica: 'Liberation Sans',
  'courier new': 'Liberation Mono',
};

/**
 * Set the family an imported document asks for.
 *
 * Documents this program writes set nothing here and stay in its own serif.
 */
export function setDocumentFont(family: string | undefined): void {
  const st = document.documentElement.style;
  if (!family) {
    st.removeProperty('--doc-font');
    return;
  }
  const twin = METRIC_TWINS[family.toLowerCase()];
  const chain = twin ? `"${family}", "${twin}", ${DOC_FONT}` : `"${family}", ${DOC_FONT}`;
  st.setProperty('--doc-font', chain);
}

export const WORD_SINGLE_LINE = 1.2;

/**
 * Draw a paragraph's direct formatting on top of its named style.
 *
 * Written as inline styles rather than extra classes because the values are
 * per paragraph and continuous. Two rules it must not break:
 *
 * - Spacing is PADDING, never margin. Sibling margins collapse, and the
 *   paginator sums block heights against a fixed content box; a collapsed
 *   margin makes that sum disagree with the container and pages overflow.
 * - An indent from the document REPLACES the style's own hanging indent
 *   rather than adding to it, or an imported bullet ends up indented twice.
 */
export function applyBlockFormat(
  el: HTMLElement,
  id: StyleId,
  f: BlockFormat | undefined
): void {
  const st = el.style;
  for (const prop of [
    'text-align',
    'padding-top',
    'padding-bottom',
    'padding-left',
    'padding-right',
    'text-indent',
    'line-height',
  ]) {
    st.removeProperty(prop);
  }
  if (!hasFormat(f)) {
    delete el.dataset.fmt;
    return;
  }
  // Kept on the element so readModel can give it back without consulting
  // the model, which is not the source of truth while editing.
  el.dataset.fmt = JSON.stringify(f);

  const d = STYLES[id];
  const [padT, padR, padB, padL] = d.padding;

  if (f.align) st.textAlign = f.align;
  if (f.spaceBefore !== undefined) st.paddingTop = px(padT + f.spaceBefore) + 'px';
  if (f.spaceAfter !== undefined) st.paddingBottom = px(padB + f.spaceAfter) + 'px';

  const indented =
    f.indentLeft !== undefined || f.firstLine !== undefined || f.indentRight !== undefined;
  if (indented) {
    const left = f.indentLeft ?? padL;
    const first = f.firstLine ?? 0;
    // A hanging indent is a negative first line against a padded box, which
    // is exactly how Word's w:ind left + w:hanging compose.
    st.paddingLeft = px(left + Math.max(0, -first)) + 'px';
    st.textIndent = px(first) + 'px';
    if (f.indentRight !== undefined) st.paddingRight = px(padR + f.indentRight) + 'px';
  }

  if (f.lineHeight !== undefined) {
    st.lineHeight =
      f.lineRule === 'exact' || f.lineRule === 'atLeast'
        ? px(f.lineHeight) + 'px'
        : String(f.lineHeight * WORD_SINGLE_LINE);
  }
}

export function injectStyleSheet(): void {
  emitStyleSheet();
}

/** Write (or rewrite) the named-style stylesheet from the table in force. */
function emitStyleSheet(): void {
  const id = 'wp-named-styles';
  let el = document.getElementById(id) as HTMLStyleElement | null;
  if (!el) {
    el = document.createElement('style');
    el.id = id;
    document.head.appendChild(el);
  }
  el.textContent =
    `.blk { font-family: var(--doc-font, ${DOC_FONT}); color: ${INK}; }\n` +
    STYLE_IDS.map((s) => css(STYLES[s])).join('') +
    LIST_CSS +
    TABLE_CSS +
    IMAGE_CSS +
    FIND_CSS +
    HF_CSS +
    SPLIT_CSS;
}
