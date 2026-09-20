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
  /** px */
  size: number;
  bold: boolean;
  uppercase: boolean;
  /** px */
  letterSpacing: number;
  lineHeight: number;
  /** px, [top, right, bottom, left]. Spacing is padding only - never margin. */
  padding: [number, number, number, number];
  /** Bottom rule, as on SectionHeading. */
  rule: boolean;
  /** px hanging indent; 0 for none. */
  hanging: number;
  bullet: boolean;
}

// Deliberately not a web font: a local stack means there is no font-load race
// to correct pagination for, and print matches screen on the first paint.
export const DOC_FONT =
  'Calibri, Carlito, "Segoe UI", "Helvetica Neue", Helvetica, Arial, sans-serif';
/** Font name written into the .docx. Matches the head of DOC_FONT. */
export const DOCX_FONT = 'Calibri';
export const INK = '#111111';
export const RULE_COLOR = '#999999';

export const STYLES: Record<StyleId, StyleDef> = {
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
    lineHeight: 1.35,
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
    lineHeight: 1.35,
    padding: [0, 0, 2, 0],
    rule: false,
    hanging: 14,
    bullet: true,
  },
};

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
  lines.push(`  font-size: ${d.size}px;`);
  lines.push(`  font-weight: ${d.bold ? 700 : 400};`);
  lines.push(`  line-height: ${d.lineHeight};`);
  if (d.uppercase) lines.push('  text-transform: uppercase;');
  if (d.letterSpacing) lines.push(`  letter-spacing: ${d.letterSpacing}px;`);
  // Spacing via padding only. Sibling margins collapse, and summed heights
  // would then disagree with the container height.
  lines.push(`  padding: ${t}px ${r}px ${b}px ${l + d.hanging}px;`);
  if (d.hanging) lines.push(`  text-indent: -${d.hanging}px;`);
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

/** Inject the `.s-*` rules. Called once at startup, before first measurement. */
export function injectStyleSheet(): void {
  const id = 'wp-named-styles';
  if (document.getElementById(id)) return;
  const el = document.createElement('style');
  el.id = id;
  el.textContent =
    `.blk { font-family: ${DOC_FONT}; color: ${INK}; }\n` +
    STYLE_IDS.map((s) => css(STYLES[s])).join('');
  document.head.appendChild(el);
}
