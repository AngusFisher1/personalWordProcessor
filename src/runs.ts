import type { RunStyle } from './model';
import { RUN_STYLE_KEYS, hasRunStyle, tidyRunStyle } from './model';
import { px } from './styles';

/**
 * Character formatting: size, font, colour, strike, super- and subscript.
 *
 * Stored in the markup as data attributes on a span rather than as a style
 * attribute, for the same reason an image stores `data-media` and not a
 * `src`: the stored form is what the document MEANS, and the CSS is a
 * render-time detail. It also means the inline sanitizer can keep a closed
 * list of attributes and strip everything else, so nothing that arrives on
 * the clipboard can smuggle a style into a saved document.
 */

const ATTR: Record<keyof RunStyle, string> = {
  size: 'data-sz',
  font: 'data-font',
  color: 'data-color',
  strike: 'data-strike',
  vert: 'data-vert',
};

/** Every attribute a run span is allowed to carry. */
export const RUN_ATTRS = new Set(Object.values(ATTR));

export function isRunSpan(el: Element): boolean {
  // Case-insensitive: the browser's HTML parser reports SPAN and the XML
  // parser the export path uses reports span.
  if (el.tagName.toUpperCase() !== 'SPAN') return false;
  for (const a of RUN_ATTRS) if (el.hasAttribute(a)) return true;
  return false;
}

export function readRunStyle(el: Element): RunStyle | undefined {
  const out: RunStyle = {};
  const size = el.getAttribute(ATTR.size);
  if (size !== null && Number.isFinite(Number(size))) out.size = Number(size);
  const font = el.getAttribute(ATTR.font);
  if (font) out.font = font;
  const color = el.getAttribute(ATTR.color);
  if (color) out.color = color;
  if (el.hasAttribute(ATTR.strike)) out.strike = true;
  const vert = el.getAttribute(ATTR.vert);
  if (vert === 'super' || vert === 'sub') out.vert = vert;
  return tidyRunStyle(out);
}

export function writeRunStyle(el: Element, r: RunStyle | undefined): void {
  for (const a of RUN_ATTRS) el.removeAttribute(a);
  if (!hasRunStyle(r)) return;
  if (r.size !== undefined) el.setAttribute(ATTR.size, String(r.size));
  if (r.font !== undefined) el.setAttribute(ATTR.font, r.font);
  if (r.color !== undefined) el.setAttribute(ATTR.color, r.color);
  if (r.strike) el.setAttribute(ATTR.strike, '1');
  if (r.vert !== undefined) el.setAttribute(ATTR.vert, r.vert);
}

/** The markup a run span is stored as. */
export function runSpanHtml(r: RunStyle, inner: string): string {
  const bits: string[] = [];
  if (r.size !== undefined) bits.push(`${ATTR.size}="${r.size}"`);
  if (r.font !== undefined) bits.push(`${ATTR.font}="${escapeAttr(r.font)}"`);
  if (r.color !== undefined) bits.push(`${ATTR.color}="${escapeAttr(r.color)}"`);
  if (r.strike) bits.push(`${ATTR.strike}="1"`);
  if (r.vert !== undefined) bits.push(`${ATTR.vert}="${r.vert}"`);
  return bits.length === 0 ? inner : `<span ${bits.join(' ')}>${inner}</span>`;
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}

/* ------------------------------------------------------------------ *
 * Drawing
 * ------------------------------------------------------------------ */

/**
 * Turn the stored attributes into CSS.
 *
 * Superscript is drawn with a smaller font and a raised baseline rather than
 * <sup>, because <sup> changes the line box height and a raised footnote
 * marker would silently make its whole line taller than the paginator
 * measured it.
 */
export function styleRunSpan(el: HTMLElement): void {
  const r = readRunStyle(el);
  const st = el.style;
  for (const prop of [
    'font-size',
    'font-family',
    'color',
    'text-decoration-line',
    'vertical-align',
  ]) {
    st.removeProperty(prop);
  }
  if (!r) return;
  if (r.size !== undefined) st.fontSize = px(r.size) + 'px';
  if (r.font !== undefined) st.fontFamily = `"${r.font}", ${'serif'}`;
  if (r.color !== undefined) st.color = '#' + r.color;
  if (r.strike) st.textDecorationLine = 'line-through';
  if (r.vert !== undefined) {
    st.fontSize = px((r.size ?? 11) * 0.72) + 'px';
    st.verticalAlign = r.vert === 'super' ? '0.35em' : '-0.2em';
  }
}

/** Give every run span in a subtree its CSS. */
export function resolveRuns(root: HTMLElement): void {
  for (const el of Array.from(root.querySelectorAll('span'))) {
    if (isRunSpan(el)) styleRunSpan(el as HTMLElement);
  }
}

/* ------------------------------------------------------------------ *
 * Editing
 * ------------------------------------------------------------------ */

/** Every text node a range touches, with the slice of each that it covers. */
function textSlices(range: Range): { node: Text; from: number; to: number }[] {
  const out: { node: Text; from: number; to: number }[] = [];
  const root = range.commonAncestorContainer;
  const walker = document.createTreeWalker(
    root.nodeType === Node.TEXT_NODE ? (root.parentNode as Node) : root,
    NodeFilter.SHOW_TEXT
  );
  let n: Node | null;
  while ((n = walker.nextNode())) {
    const t = n as Text;
    if (!range.intersectsNode(t)) continue;
    const from = t === range.startContainer ? range.startOffset : 0;
    const to = t === range.endContainer ? range.endOffset : t.data.length;
    if (to > from) out.push({ node: t, from, to });
  }
  return out;
}

/** The run span that wraps this node exactly, if there is one. */
function owningSpan(node: Node): HTMLElement | null {
  const parent = node.parentElement;
  if (!parent || !isRunSpan(parent)) return null;
  return parent;
}

/**
 * Apply a change to every character of a range.
 *
 * A key set to null is REMOVED, which is how "no colour" differs from "a
 * colour I have not set": the first has to survive back into the export as
 * an absent w:color, the second inherits the paragraph's.
 */
export function applyRunStyle(
  range: Range,
  change: Partial<Record<keyof RunStyle, unknown>>
): void {
  const slices = textSlices(range);
  // Later slices first, so splitting one does not move the offsets of the
  // ones before it.
  for (let i = slices.length - 1; i >= 0; i--) {
    const { node, from, to } = slices[i];
    let target = node;
    if (to < target.data.length) target.splitText(to);
    if (from > 0) target = target.splitText(from);

    const existing = owningSpan(target);
    const base = existing ? readRunStyle(existing) ?? {} : {};
    const next: RunStyle = { ...base };
    for (const k of RUN_STYLE_KEYS) {
      if (!(k in change)) continue;
      const v = change[k];
      if (v === null || v === undefined || v === false) delete next[k];
      else (next as Record<string, unknown>)[k] = v;
    }
    const tidy = tidyRunStyle(next);

    // The span already wraps exactly this text: rewrite it in place.
    if (existing && existing.childNodes.length === 1) {
      if (tidy) {
        writeRunStyle(existing, tidy);
        styleRunSpan(existing);
      } else {
        const parent = existing.parentNode;
        while (existing.firstChild) parent?.insertBefore(existing.firstChild, existing);
        parent?.removeChild(existing);
      }
      continue;
    }
    if (!tidy) continue; // nothing to add and nothing to strip

    const span = document.createElement('span');
    writeRunStyle(span, tidy);
    styleRunSpan(span);
    target.parentNode?.insertBefore(span, target);
    span.appendChild(target);
  }
}

/**
 * What the selection has in common, for ticking the menus.
 *
 * A property the selected characters disagree on comes back undefined, the
 * same answer a paragraph-level mixed selection gives.
 */
export function runStyleOfRange(range: Range): RunStyle {
  const slices = textSlices(range);
  if (slices.length === 0) return {};
  const styles = slices.map((s) => owningSpan(s.node) ?? null).map((el) =>
    el ? readRunStyle(el) ?? {} : {}
  );
  const out: RunStyle = { ...styles[0] };
  for (const s of styles.slice(1)) {
    for (const k of RUN_STYLE_KEYS) {
      if (s[k] !== out[k]) delete out[k];
    }
  }
  return out;
}
