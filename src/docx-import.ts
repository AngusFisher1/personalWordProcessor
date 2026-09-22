import type {
  Block,
  BlockAlign,
  BlockFormat,
  Doc,
  HFVariant,
  HeaderFooterSet,
  PageSetup,
  ParagraphBlock,
  RunStyle,
  Section,
  StyleId,
  TableBlock,
  TableCell,
  TableRow,
} from './model';
import {
  contentWidth as contentWidthOf,
  newId,
  sameRunStyle,
  tidyFormat,
  tidyRunStyle,
} from './model';
import type { RunProp, Vault } from './docx-package';
import {
  DOC_XML,
  addWarning,
  emptyVault,
  partText,
  unzip,
} from './docx-package';
import type { ParaSignals } from './docx-infer';
import { inferStyles, looksLikeContact } from './docx-infer';
import { registerMedia } from './media';
import { runSpanHtml } from './runs';

/**
 * Parse OOXML directly rather than converting through HTML.
 *
 * mammoth and friends produce HTML and throw away the style definitions,
 * numbering and section properties that a faithful round trip needs. Reading
 * the XML is more work up front and the only way to write the file back.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/* ------------------------------------------------------------------ *
 * Units. Converted once, here.
 *   1px at 96dpi = 15 twips; 1pt = 2 half-points = 20 twips
 * ------------------------------------------------------------------ */

const TWIP_PER_PX = 15;
export const twipToPx = (t: number) => t / TWIP_PER_PX;

/* ------------------------------------------------------------------ *
 * XML helpers
 *
 * Matched on localName so a package that uses an unusual namespace prefix
 * still parses.
 * ------------------------------------------------------------------ */

function elementChildren(el: Element): Element[] {
  // childNodes rather than children: the latter is not available on every
  // DOM implementation this code runs under.
  return Array.from(el.childNodes).filter((n) => n.nodeType === 1) as Element[];
}

function kids(el: Element | null, name: string): Element[] {
  if (!el) return [];
  return elementChildren(el).filter((c) => c.localName === name);
}

function kid(el: Element | null, name: string): Element | null {
  return kids(el, name)[0] ?? null;
}

function wAttr(el: Element | null, name: string): string | null {
  if (!el) return null;
  return (
    el.getAttributeNS(W, name) ??
    el.getAttribute('w:' + name) ??
    el.getAttribute(name)
  );
}

function rAttr(el: Element | null, name: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(R_NS, name) ?? el.getAttribute('r:' + name);
}

/** <w:b/>, <w:b w:val="1"/> are on; w:val="0"/"false"/"none" is off. */
function onOff(el: Element | null): boolean {
  if (!el) return false;
  const v = wAttr(el, 'val');
  if (v === null) return true;
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

function intOf(v: string | null, fallback = 0): number {
  if (v === null) return fallback;
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : fallback;
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(s: string): string {
  return escapeHtml(s).replace(/"/g, '&quot;');
}

/* ------------------------------------------------------------------ *
 * Inline content
 * ------------------------------------------------------------------ */

/**
 * A run's own size, font, colour and decoration.
 *
 * Absolute, not relative: the difference against the paragraph's base is
 * taken later, once the base is known, because a paragraph's base is itself
 * read from its first run.
 */
function runStyleOf(rPr: Element | null): RunStyle | undefined {
  if (!rPr) return undefined;
  const out: RunStyle = {};
  const sz = kid(rPr, 'sz');
  if (sz) {
    const half = intOf(wAttr(sz, 'val'), 0);
    // Half-points, as every font size in OOXML is.
    if (half > 0) out.size = Math.round((half / 2) * 100) / 100;
  }
  const fonts = kid(rPr, 'rFonts');
  const ascii = fonts ? wAttr(fonts, 'ascii') ?? wAttr(fonts, 'hAnsi') : null;
  if (ascii) out.font = ascii;
  const color = wAttr(kid(rPr, 'color'), 'val');
  // "auto" means "whatever the theme says", which is what we draw anyway.
  if (color && color.toLowerCase() !== 'auto' && /^[0-9a-f]{6}$/i.test(color)) {
    out.color = color.toUpperCase();
  }
  if (onOff(kid(rPr, 'strike'))) out.strike = true;
  const vert = wAttr(kid(rPr, 'vertAlign'), 'val');
  if (vert === 'superscript') out.vert = 'super';
  else if (vert === 'subscript') out.vert = 'sub';
  return tidyRunStyle(out);
}

interface Fmt {
  b: boolean;
  i: boolean;
  u: boolean;
}

type Piece =
  | ({ t: 'text'; text: string; href: string | null; rs?: RunStyle } & Fmt)
  | { t: 'br' }
  | { t: 'img'; media: string; run: string; w: number; h: number }
  | { t: 'field'; name: string };

interface Ctx {
  rels: Map<string, string>;
  vault: Vault;
  /** Serializes an element without repeating the root's namespaces. */
  serialize: (el: Element) => string;
  imageCount: number;
  /**
   * Complex-field state. Word writes PAGE as begin / instrText / separate /
   * cached result / end, spread across sibling runs, so reading one means
   * carrying state between them.
   */
  fieldDepth: number;
  fieldInstr: string;
  fieldSkip: boolean;
}

/** " PAGE  \* MERGEFORMAT " -> "PAGE". */
function fieldName(instr: string): string {
  return (instr.trim().split(/\s+/)[0] ?? '').toUpperCase();
}

const KNOWN_FIELDS = new Set(['PAGE', 'NUMPAGES', 'TITLE', 'FILENAME', 'DATE']);

/** 914400 EMU to the inch, 96 CSS pixels to the inch. */
const EMU_PER_PX = 9525;

function descendant(el: Element, name: string): Element | null {
  if (el.localName === name) return el;
  const all = el.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    if ((all[i] as Element).localName === name) return all[i] as Element;
  }
  return null;
}

/**
 * An inline picture. Both the modern w:drawing and the older w:pict carry a
 * relationship id pointing at a part under word/media, plus a size - in EMU
 * for a drawing, in CSS-ish units in a VML style attribute for a pict.
 */
function readImage(run: Element, holder: Element, ctx: Ctx): Piece | null {
  const blip = descendant(holder, 'blip') ?? descendant(holder, 'imagedata');
  if (!blip) return null;
  const rid = rAttr(blip, 'embed') ?? rAttr(blip, 'id');
  if (!rid) return null;
  const target = ctx.rels.get(rid);
  if (!target) return null;

  // Relationship targets are relative to the part's own folder.
  const media = target.startsWith('/')
    ? target.slice(1)
    : 'word/' + target.replace(/^\.\//, '');

  let w = 0;
  let h = 0;
  const extent = descendant(holder, 'extent');
  if (extent) {
    w = Math.round(intOf(wAttr(extent, 'cx'), 0) / EMU_PER_PX);
    h = Math.round(intOf(wAttr(extent, 'cy'), 0) / EMU_PER_PX);
  } else {
    const shape = descendant(holder, 'shape');
    const style = shape?.getAttribute('style') ?? '';
    const num = (prop: string) => {
      const m = style.match(new RegExp(prop + ':([\d.]+)pt'));
      return m ? Math.round(parseFloat(m[1]) * (96 / 72)) : 0;
    };
    w = num('width');
    h = num('height');
  }

  const token = 'img' + ctx.imageCount++;
  // The whole run goes into the vault: a w:drawing carries cropping, effects
  // and positioning we render none of but must not throw away.
  ctx.vault.runXml.set(token, ctx.serialize(run));
  return { t: 'img', media, run: token, w, h };
}

function uOn(el: Element | null): boolean {
  if (!el) return false;
  const v = wAttr(el, 'val');
  return v !== null && v !== 'none';
}

function emitRun(r: Element, fmt: Fmt, href: string | null, out: Piece[], ctx: Ctx): void {
  const rPr = kid(r, 'rPr');
  const f: Fmt = {
    b: fmt.b || onOff(kid(rPr, 'b')),
    i: fmt.i || onOff(kid(rPr, 'i')),
    u: fmt.u || uOn(kid(rPr, 'u')),
  };
  const rs = runStyleOf(rPr);
  for (const c of elementChildren(r)) {
    switch (c.localName) {
      case 'fldChar': {
        const kind = wAttr(c, 'fldCharType');
        if (kind === 'begin') {
          ctx.fieldDepth++;
          ctx.fieldInstr = '';
          ctx.fieldSkip = false;
        } else if (kind === 'separate') {
          ctx.fieldSkip = true; // what follows is the cached result
        } else if (kind === 'end' && ctx.fieldDepth > 0) {
          ctx.fieldDepth--;
          const name = fieldName(ctx.fieldInstr);
          if (KNOWN_FIELDS.has(name)) out.push({ t: 'field', name });
          ctx.fieldInstr = '';
          ctx.fieldSkip = false;
        }
        break;
      }
      case 'instrText':
        if (ctx.fieldDepth > 0) ctx.fieldInstr += c.textContent ?? '';
        break;
      case 't':
        // Inside a field this is the cached result, which we recompute.
        if (ctx.fieldDepth > 0) break;
        out.push({ t: 'text', text: c.textContent ?? '', href, ...f, rs });
        break;
      case 'br':
        out.push({ t: 'br' });
        break;
      case 'tab':
        // We have no tab stops; a space keeps the words apart.
        out.push({ t: 'text', text: ' ', href, ...f, rs });
        break;
      case 'noBreakHyphen':
        out.push({ t: 'text', text: '-', href, ...f, rs });
        break;
      case 'drawing':
      case 'pict':
      case 'object': {
        const img = readImage(r, c, ctx);
        if (img) out.push(img);
        else {
          addWarning(ctx.vault, 'images', 'Some images are preserved but not shown');
        }
        break;
      }
      case 'delText':
        break; // deleted text is not part of the current document
      default:
        break;
    }
  }
}

function walkInline(
  el: Element,
  fmt: Fmt,
  href: string | null,
  out: Piece[],
  ctx: Ctx
): void {
  for (const c of elementChildren(el)) {
    switch (c.localName) {
      case 'pPr':
        break;
      case 'r':
        emitRun(c, fmt, href, out, ctx);
        break;
      case 'hyperlink': {
        const id = rAttr(c, 'id');
        const target = id ? ctx.rels.get(id) ?? null : null;
        const anchor = wAttr(c, 'anchor');
        walkInline(c, fmt, target ?? (anchor ? '#' + anchor : href), out, ctx);
        break;
      }
      case 'ins':
        addWarning(ctx.vault, 'revisions', 'Tracked changes are shown as accepted');
        walkInline(c, fmt, href, out, ctx);
        break;
      case 'del':
        addWarning(ctx.vault, 'revisions', 'Tracked changes are shown as accepted');
        break;
      case 'sdt':
        addWarning(ctx.vault, 'contentControls', 'Content controls are read-only');
        walkInline(c, fmt, href, out, ctx);
        break;
      case 'fldSimple': {
        const name = fieldName(wAttr(c, 'instr') ?? '');
        if (KNOWN_FIELDS.has(name)) out.push({ t: 'field', name });
        else walkInline(c, fmt, href, out, ctx); // keep its cached text
        break;
      }
      case 'sdtContent':
      case 'smartTag':
        walkInline(c, fmt, href, out, ctx);
        break;
      default:
        break;
    }
  }
}

/**
 * Adjacent pieces that agree are merged first, so a paragraph whose every
 * run is the same size becomes one span rather than one per run.
 */
function piecesToHtml(pieces: Piece[]): string {
  const merged: Piece[] = [];
  for (const p of pieces) {
    const last = merged[merged.length - 1];
    if (
      p.t === 'text' &&
      last &&
      last.t === 'text' &&
      last.b === p.b &&
      last.i === p.i &&
      last.u === p.u &&
      last.href === p.href &&
      sameRunStyle(last.rs, p.rs)
    ) {
      last.text += p.text;
    } else {
      merged.push({ ...p });
    }
  }
  let html = '';
  for (const p of merged) {
    if (p.t === 'br') {
      html += '<br>';
      continue;
    }
    if (p.t === 'field') {
      // Empty on purpose: the value is filled in at render, per page.
      html += `<span data-field="${escapeAttr(p.name)}"></span>`;
      continue;
    }
    if (p.t === 'img') {
      // data-media says which part to draw; data-run says which preserved
      // run to write back. No src: a blob URL would not survive a reload.
      html +=
        `<img data-media="${escapeAttr(p.media)}" data-run="${escapeAttr(p.run)}"` +
        (p.w ? ` width="${p.w}"` : '') +
        (p.h ? ` height="${p.h}"` : '') +
        '>';
      continue;
    }
    if (p.text === '') continue;
    let t = escapeHtml(p.text);
    if (p.rs) t = runSpanHtml(p.rs, t);
    if (p.u) t = '<u>' + t + '</u>';
    if (p.i) t = '<i>' + t + '</i>';
    if (p.b) t = '<b>' + t + '</b>';
    if (p.href) t = '<a href="' + escapeAttr(p.href) + '">' + t + '</a>';
    html += t;
  }
  return html;
}

/* ------------------------------------------------------------------ *
 * Styles
 * ------------------------------------------------------------------ */

function normalizeName(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9]/g, '');
}

const STYLE_BY_NAME: Record<string, StyleId> = {
  title: 'Name',
  name: 'Name',
  subtitle: 'Contact',
  contact: 'Contact',
  heading1: 'SectionHeading',
  heading2: 'SectionHeading',
  sectionheading: 'SectionHeading',
  heading3: 'JobTitle',
  heading4: 'JobTitle',
  heading5: 'JobTitle',
  heading6: 'JobTitle',
  heading7: 'JobTitle',
  heading8: 'JobTitle',
  heading9: 'JobTitle',
  jobtitle: 'JobTitle',
  listparagraph: 'Bullet',
  listbullet: 'Bullet',
  bullet: 'Bullet',
  normal: 'Body',
  bodytext: 'Body',
  body: 'Body',
  defaultparagraphfont: 'Body',
  nospacing: 'Body',
};

interface StyleInfo {
  names: Map<string, string>;
  /** w:docDefaults run size in half-points, for judging what counts as big. */
  defaultSizeHalfPt: number | null;
}

function readStyles(xml: string | null): StyleInfo {
  const names = new Map<string, string>();
  if (!xml) return { names, defaultSizeHalfPt: null };
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const root = doc.documentElement;

  for (const st of kids(root, 'style')) {
    const id = wAttr(st, 'styleId');
    const name = wAttr(kid(st, 'name'), 'val');
    if (id) names.set(id, name ?? id);
  }

  let size: number | null = null;
  const fromDefaults = kid(
    kid(kid(root, 'docDefaults'), 'rPrDefault'),
    'rPr'
  );
  const dsz = intOf(wAttr(kid(fromDefaults, 'sz'), 'val'), 0);
  if (dsz > 0) size = dsz;
  if (size === null) {
    // Fall back to whichever paragraph style is marked as the default.
    for (const st of kids(root, 'style')) {
      if (wAttr(st, 'default') !== '1' && wAttr(st, 'default') !== 'true') continue;
      const sz = intOf(wAttr(kid(kid(st, 'rPr'), 'sz'), 'val'), 0);
      if (sz > 0) {
        size = sz;
        break;
      }
    }
  }
  return { names, defaultSizeHalfPt: size };
}

/**
 * A style we do not recognize returns null so the paragraph goes to the
 * inference pass instead of silently becoming Body. Either way the original
 * w:pStyle goes back out on export, so nothing is lost.
 */
function recognizedStyle(
  styleId: string | null,
  names: Map<string, string>
): StyleId | null {
  if (!styleId) return null;
  const byId = STYLE_BY_NAME[normalizeName(styleId)];
  if (byId) return byId;
  const name = names.get(styleId);
  if (name) {
    const byName = STYLE_BY_NAME[normalizeName(name)];
    if (byName) return byName;
  }
  // Unrecognized: fall through to inference rather than silently to Body.
  return null;
}

/* ------------------------------------------------------------------ *
 * Paragraph statistics, for inferring a style when none is given
 * ------------------------------------------------------------------ */

/**
 * rPr children we manage from the markup instead of preserving.
 *
 * Size, font and colour joined this list when they became editable. They
 * used to be preserved as part of a single per-paragraph base that was then
 * written onto every regenerated run - which is precisely how editing a line
 * with one red word in it turned the whole line red. Each run now carries
 * its own, absolutely, so there is no inheritance to get wrong.
 */
const OWNED_RUN_PROPS = new Set([
  'b', 'bCs', 'i', 'iCs', 'u',
  'sz', 'szCs', 'rFonts', 'color', 'strike', 'vertAlign',
]);

interface Stats {
  boldShare: number;
  caps: boolean;
  sizeHalfPt: number | null;
  baseRPr: RunProp[];
  text: string;
}

/**
 * Measure a paragraph by how many CHARACTERS carry each property, not by how
 * many runs do. Word splits a line into runs for reasons of its own - a spell
 * check boundary, an edit session - so counting runs weights a one-character
 * fragment the same as the rest of the sentence.
 */
function paragraphStats(p: Element, serialize: (el: Element) => string): Stats {
  let boldChars = 0;
  let capsChars = 0;
  let baseRPr: RunProp[] = [];
  let text = '';
  const bySize = new Map<number, number>();

  const all = p.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    const el = all[i] as Element;
    if (el.localName !== 'r') continue;

    let t = '';
    for (const c of elementChildren(el)) {
      if (c.localName === 't') t += c.textContent ?? '';
      else if (c.localName === 'tab') t += ' ';
    }
    if (t === '') continue;

    text += t;
    const rPr = kid(el, 'rPr');
    if (onOff(kid(rPr, 'b'))) boldChars += t.length;
    if (onOff(kid(rPr, 'caps'))) capsChars += t.length;
    const sz = intOf(wAttr(kid(rPr, 'sz'), 'val'), 0);
    if (sz > 0) bySize.set(sz, (bySize.get(sz) ?? 0) + t.length);
    if (baseRPr.length === 0 && rPr) {
      baseRPr = elementChildren(rPr)
        .filter((c) => !OWNED_RUN_PROPS.has(c.localName))
        .map((c) => ({ name: c.localName, xml: serialize(c) }));
    }
  }

  let size: number | null = null;
  let bestChars = 0;
  for (const [sz, chars] of bySize) {
    if (chars > bestChars) {
      size = sz;
      bestChars = chars;
    }
  }

  const letters = text.replace(/[^A-Za-z]/g, '');
  const caps =
    (text.length > 0 && capsChars === text.length) ||
    (letters.length >= 4 && text === text.toUpperCase());

  return {
    boldShare: text.length > 0 ? boldChars / text.length : 0,
    caps,
    sizeHalfPt: size,
    baseRPr,
    text,
  };
}

/* ------------------------------------------------------------------ *
 * Numbering
 * ------------------------------------------------------------------ */

interface Lvl {
  fmt: string;
  text: string;
  start: number;
}

interface Numbering {
  /** numId -> ilvl -> level definition */
  levels: Map<string, Map<number, Lvl>>;
}

function readNumbering(xml: string | null): Numbering {
  const levels = new Map<string, Map<number, Lvl>>();
  if (!xml) return { levels };
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const root = doc.documentElement;

  const abstract = new Map<string, Map<number, Lvl>>();
  for (const an of kids(root, 'abstractNum')) {
    const id = wAttr(an, 'abstractNumId');
    if (!id) continue;
    const m = new Map<number, Lvl>();
    for (const lvl of kids(an, 'lvl')) {
      const ilvl = intOf(wAttr(lvl, 'ilvl'));
      m.set(ilvl, {
        fmt: wAttr(kid(lvl, 'numFmt'), 'val') ?? 'bullet',
        text: wAttr(kid(lvl, 'lvlText'), 'val') ?? '•',
        start: intOf(wAttr(kid(lvl, 'start'), 'val'), 1),
      });
    }
    abstract.set(id, m);
  }
  for (const n of kids(root, 'num')) {
    const numId = wAttr(n, 'numId');
    const absId = wAttr(kid(n, 'abstractNumId'), 'val');
    if (numId && absId && abstract.has(absId)) {
      levels.set(numId, abstract.get(absId) as Map<number, Lvl>);
    }
  }
  return { levels };
}

/** Symbol/Wingdings bullet code points, mapped to characters that render. */
const SYMBOL_BULLETS: Record<string, string> = {
  '': '•',
  '': '▪',
  '': '●',
  '': '➢',
  o: '◦',
};

function roman(n: number): string {
  const table: [number, string][] = [
    [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
    [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
  ];
  let out = '';
  let v = n;
  for (const [k, s] of table) {
    while (v >= k) {
      out += s;
      v -= k;
    }
  }
  return out;
}

function letters(n: number): string {
  let out = '';
  let v = n;
  while (v > 0) {
    const r = (v - 1) % 26;
    out = String.fromCharCode(97 + r) + out;
    v = Math.floor((v - 1) / 26);
  }
  return out;
}

function formatCounter(n: number, fmt: string): string {
  switch (fmt) {
    case 'decimalZero':
      return n < 10 ? '0' + n : String(n);
    case 'lowerLetter':
      return letters(n);
    case 'upperLetter':
      return letters(n).toUpperCase();
    case 'lowerRoman':
      return roman(n);
    case 'upperRoman':
      return roman(n).toUpperCase();
    case 'none':
      return '';
    default:
      return String(n);
  }
}

class ListCounters {
  private counts = new Map<string, number[]>();

  marker(num: Numbering, numId: string, ilvl: number): string {
    const lvl = num.levels.get(numId)?.get(ilvl);
    if (!lvl) return '•';

    if (lvl.fmt === 'bullet') {
      const ch = lvl.text.trim();
      return SYMBOL_BULLETS[ch] ?? (ch || '•');
    }

    const arr = this.counts.get(numId) ?? [];
    while (arr.length <= ilvl) arr.push(0);
    arr[ilvl] = arr[ilvl] === 0 ? lvl.start : arr[ilvl] + 1;
    for (let i = ilvl + 1; i < arr.length; i++) arr[i] = 0;
    this.counts.set(numId, arr);

    // lvlText is a pattern like "%1." or "%1.%2"
    return lvl.text.replace(/%(\d)/g, (_m, d: string) => {
      const idx = parseInt(d, 10) - 1;
      const lv = num.levels.get(numId)?.get(idx);
      const value = arr[idx] || lv?.start || 1;
      return formatCounter(value, lv?.fmt ?? 'decimal');
    });
  }
}

/* ------------------------------------------------------------------ *
 * Direct paragraph formatting
 * ------------------------------------------------------------------ */

/** Twips to points. A point is twenty twips; both are exact. */
const twipToPt = (t: number) => Math.round((t / 20) * 100) / 100;

const ALIGN: Record<string, BlockAlign> = {
  left: 'left',
  start: 'left',
  center: 'center',
  centre: 'center',
  right: 'right',
  end: 'right',
  both: 'justify',
  justify: 'justify',
  distribute: 'justify',
};

/**
 * Read w:jc, w:ind and w:spacing off a paragraph.
 *
 * All three were preserved and written back long before they were drawn, so
 * nothing here changes what a round trip produces. What it changes is what
 * the screen shows - which, on 47% of the paragraphs in a real corpus, was
 * not what the document said.
 */
export function readParagraphFormat(pPr: Element | null): BlockFormat | undefined {
  if (!pPr) return undefined;
  const f: BlockFormat = {};

  const jc = wAttr(kid(pPr, 'jc'), 'val');
  if (jc && ALIGN[jc] && ALIGN[jc] !== 'left') f.align = ALIGN[jc];

  const ind = kid(pPr, 'ind');
  if (ind) {
    // w:start and w:end are the newer spellings of w:left and w:right.
    const left = wAttr(ind, 'left') ?? wAttr(ind, 'start');
    const right = wAttr(ind, 'right') ?? wAttr(ind, 'end');
    const firstLine = wAttr(ind, 'firstLine');
    const hanging = wAttr(ind, 'hanging');
    if (left !== null) f.indentLeft = twipToPt(intOf(left, 0));
    if (right !== null) f.indentRight = twipToPt(intOf(right, 0));
    // Hanging wins over firstLine when a document sets both, as Word does.
    if (hanging !== null) f.firstLine = -twipToPt(intOf(hanging, 0));
    else if (firstLine !== null) f.firstLine = twipToPt(intOf(firstLine, 0));
  }

  const spacing = kid(pPr, 'spacing');
  if (spacing) {
    const before = wAttr(spacing, 'before');
    const after = wAttr(spacing, 'after');
    const line = wAttr(spacing, 'line');
    const rule = wAttr(spacing, 'lineRule');
    if (before !== null) f.spaceBefore = twipToPt(intOf(before, 0));
    if (after !== null) f.spaceAfter = twipToPt(intOf(after, 0));
    if (line !== null) {
      const n = intOf(line, 240);
      if (rule === 'exact' || rule === 'atLeast') {
        f.lineRule = rule;
        f.lineHeight = twipToPt(n);
      } else {
        // "auto" counts in 240ths of a line, so 360 is one and a half.
        f.lineRule = 'auto';
        f.lineHeight = Math.round((n / 240) * 1000) / 1000;
      }
    }
  }

  return tidyFormat(f);
}

/* ------------------------------------------------------------------ *
 * Section properties
 * ------------------------------------------------------------------ */

const DEFAULT_PAGE: PageSetup = {
  width: 816,
  height: 1056,
  margins: { top: 96, right: 96, bottom: 96, left: 96 },
};

export function readSectPr(sectPr: Element | null): PageSetup {
  if (!sectPr) return DEFAULT_PAGE;
  const sz = kid(sectPr, 'pgSz');
  const mar = kid(sectPr, 'pgMar');
  const round = (t: number) => Math.round(twipToPx(t));
  const width = sz ? round(intOf(wAttr(sz, 'w'), 12240)) : DEFAULT_PAGE.width;
  const height = sz ? round(intOf(wAttr(sz, 'h'), 15840)) : DEFAULT_PAGE.height;
  const m = (name: string, d: number) =>
    mar ? Math.max(0, round(intOf(wAttr(mar, name), d))) : d;
  return {
    width,
    height,
    margins: {
      top: m('top', 1440),
      right: m('right', 1440),
      bottom: m('bottom', 1440),
      left: m('left', 1440),
    },
  };
}

/* ------------------------------------------------------------------ *
 * Import
 * ------------------------------------------------------------------ */

export interface ImportResult {
  doc: Doc;
  vault: Vault;
}

function readRels(xml: string | null): Map<string, string> {
  const out = new Map<string, string>();
  if (!xml) return out;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  for (const rel of elementChildren(doc.documentElement)) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) out.set(id, target);
  }
  return out;
}

/** Body children are keyed by position so a re-parse reproduces the same ids. */
function bodyBlockId(index: number): string {
  return 'w' + index;
}

/**
 * Serialize a body child without repeating the namespace declarations the
 * root element already carries. XMLSerializer adds them to every fragment it
 * is handed, which both bloats the file and stops an untouched document from
 * coming back out identical.
 */
function makeSerializer(root: Element): (el: Element) => string {
  const declared: string[] = [];
  for (const a of Array.from(root.attributes)) {
    if (a.name === 'xmlns' || a.name.startsWith('xmlns:')) {
      declared.push(` ${a.name}="${a.value}"`);
    }
  }
  const ser = new XMLSerializer();
  return (el: Element) => {
    let out = ser.serializeToString(el);
    for (const decl of declared) out = out.split(decl).join('');
    return out;
  };
}

export async function importDocx(
  data: ArrayBuffer,
  fileName: string
): Promise<ImportResult> {
  const parts = await unzip(data);
  const docXml = partText(parts, DOC_XML);
  if (!docXml) throw new Error('Not a Word document: word/document.xml is missing');

  const vault = emptyVault();
  vault.parts = parts;

  // Keep everything outside <w:body> verbatim, namespace declarations and all.
  const openIdx = docXml.search(/<w:body(\s[^>]*)?>/);
  const closeIdx = docXml.lastIndexOf('</w:body>');
  if (openIdx < 0 || closeIdx < 0) {
    throw new Error('Not a Word document: no body element');
  }
  const openTag = (docXml.match(/<w:body(\s[^>]*)?>/) as RegExpMatchArray)[0];
  vault.docXmlPrefix = docXml.slice(0, openIdx) + openTag;
  vault.docXmlSuffix = docXml.slice(closeIdx);

  const xml = new DOMParser().parseFromString(docXml, 'application/xml');
  if (xml.getElementsByTagName('parsererror').length > 0) {
    throw new Error('word/document.xml is not valid XML');
  }
  const body = kid(xml.documentElement, 'body');
  if (!body) throw new Error('Not a Word document: no body element');

  const serialize = makeSerializer(xml.documentElement);
  const { names: styleNames, defaultSizeHalfPt } = readStyles(
    partText(parts, 'word/styles.xml')
  );
  const numbering = readNumbering(partText(parts, 'word/numbering.xml'));
  const rels = readRels(partText(parts, 'word/_rels/document.xml.rels'));
  for (const [id, target] of rels) {
    if (!vault.relByTarget.has(target)) vault.relByTarget.set(target, id);
  }
  const ctx: Ctx = {
    rels,
    vault,
    serialize,
    imageCount: 0,
    fieldDepth: 0,
    fieldInstr: '',
    fieldSkip: false,
  };
  registerMedia(parts);
  const counters = new ListCounters();
  // The FIRST section's geometry, found before the walk because table
  // widths are measured against it. Using the body-level sectPr here lays
  // tables out against the margins of the document's last few paragraphs.
  const firstSectPr =
    Array.from(body.childNodes)
      .filter((n): n is Element => n.nodeType === 1 && (n as Element).localName === 'p')
      .map((pEl) => kid(kid(pEl, 'pPr'), 'sectPr'))
      .find((x): x is Element => !!x) ?? kid(body, 'sectPr');
  const page = readSectPr(firstSectPr);

  const blocks: Block[] = [];
  /** Inline w:sectPr found during the walk, in document order. */
  const sectBreaks: { sectPr: Element; afterBlockId: string }[] = [];
  const signals: ParaSignals[] = [];
  const pStyles: (string | null)[] = [];
  /** Parallel to `signals`, so inference can write its answer back. */
  const paragraphs: ParagraphBlock[] = [];
  let lastBlockId: string | null = null;
  let index = -1;

  /** Read one w:p. Used for body paragraphs and for the ones inside cells. */
  const readParagraph = (el: Element, id: string): ParagraphBlock => {
    const pPr = kid(el, 'pPr');
    const pStyle = wAttr(kid(pPr, 'pStyle'), 'val');
    const numPr = kid(pPr, 'numPr');

    let listMarker: string | undefined;
    let listLevel: number | undefined;
    let isList = false;
    if (numPr) {
      const numId = wAttr(kid(numPr, 'numId'), 'val');
      const ilvl = intOf(wAttr(kid(numPr, 'ilvl'), 'val'), 0);
      if (numId && numId !== '0') {
        isList = true;
        listMarker = counters.marker(numbering, numId, ilvl);
        listLevel = ilvl;
      }
    }

    const stats = paragraphStats(el, serialize);
    const pieces: Piece[] = [];
    walkInline(el, { b: false, i: false, u: false }, null, pieces, ctx);
    const html = piecesToHtml(pieces);
    const outline = kid(pPr, 'outlineLvl');
    const bdr = kid(pPr, 'pBdr');

    signals.push({
      explicit: recognizedStyle(pStyle, styleNames),
      outlineLvl: outline ? intOf(wAttr(outline, 'val'), 0) : null,
      sizeHalfPt: stats.sizeHalfPt,
      boldShare: stats.boldShare,
      caps: stats.caps,
      ruled: !!kid(bdr, 'bottom'),
      centered: wAttr(kid(pPr, 'jc'), 'val') === 'center',
      textLen: stats.text.trim().length,
      isList,
      contactish: looksLikeContact(stats.text),
    });
    pStyles.push(pStyle);

    vault.blockXml.set(id, serialize(el));
    if (pPr) vault.blockPPr.set(id, serialize(pPr));
    if (stats.baseRPr.length > 0) vault.blockRPr.set(id, stats.baseRPr);
    vault.blockHtml.set(id, html);

    const fmt = readParagraphFormat(pPr);
    if (fmt) vault.blockFmt.set(id, fmt);

    const block: ParagraphBlock = {
      id,
      styleId: 'Body', // replaced once the whole document has been measured
      html,
      ...(listMarker !== undefined ? { listMarker } : {}),
      ...(listLevel ? { listLevel } : {}),
      ...(fmt ? { fmt } : {}),
    };
    paragraphs.push(block);
    return block;
  };

  /**
   * w:tbl maps straight across. Column widths come from w:tblGrid in twips,
   * and w:tblHeader on a row marks it as one to repeat. The whole table XML
   * also goes into the vault, so a table nobody edited exports verbatim.
   */
  const readTable = (tbl: Element, id: string): TableBlock => {
    const contentW = contentWidthOf(page);
    // Raw first: clamping each column before summing turns a grid of zeroes
    // into a table 8px wide instead of one we should lay out ourselves.
    const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) =>
      Math.round(twipToPx(intOf(wAttr(g, 'w'), 0)))
    );
    const gridSum = grid.reduce((a, b) => a + b, 0);

    const rows: TableRow[] = [];
    let ri = -1;
    for (const tr of kids(tbl, 'tr')) {
      ri++;
      const rowId = `${id}r${ri}`;
      const trPr = kid(tr, 'trPr');
      if (trPr) vault.rowPr.set(rowId, serialize(trPr));

      const cells: TableCell[] = [];
      for (const tc of kids(tr, 'tc')) {
        // Index into `cells`, which includes the placeholders a gridSpan
        // leaves behind, so import and export agree on cell keys.
        const slot = cells.length;
        const tcPr = kid(tc, 'tcPr');
        if (tcPr) vault.cellPr.set(`${rowId}c${slot}`, serialize(tcPr));

        const paras = kids(tc, 'p').map((pEl, pi) =>
          readParagraph(pEl, `${rowId}c${slot}p${pi}`)
        );
        if (paras.length === 0) {
          paras.push({ id: `${rowId}c${slot}p0`, styleId: 'Body', html: '' });
        }
        cells.push(paras);
        // A cell merged across columns is followed by that many empty cells,
        // which the renderer turns back into a colspan. Keeping them means
        // the row still has one entry per grid column.
        const span = intOf(wAttr(kid(tcPr, 'gridSpan'), 'val'), 1);
        for (let k = 1; k < span; k++) cells.push([]);
      }
      rows.push({
        id: rowId,
        headerRow: !!kid(trPr, 'tblHeader'),
        cells,
      });
    }

    const tblPr = kid(tbl, 'tblPr');
    const tblGrid = kid(tbl, 'tblGrid');
    vault.tablePr.set(id, {
      tblPr: tblPr ? serialize(tblPr) : '',
      tblGrid: tblGrid ? serialize(tblGrid) : '',
    });

    const columnCount = Math.max(
      1,
      grid.length,
      ...rows.map((r) => r.cells.length)
    );
    /**
     * A grid is only usable if it describes a table of roughly the right
     * size. Some writers emit a w:tblGrid of nominal widths and size the
     * table by percentage instead, which would otherwise produce a table a
     * few pixels wide. Anything wider than the page is scaled down to fit.
     */
    let widths: number[];
    if (gridSum >= contentW * 0.5) {
      const scale = gridSum > contentW ? contentW / gridSum : 1;
      widths = grid.map((w) => Math.max(8, Math.round(w * scale)));
    } else {
      widths = new Array(columnCount).fill(Math.floor(contentW / columnCount));
    }

    vault.blockXml.set(id, serialize(tbl));
    return { kind: 'table', id, cols: widths.map((w) => ({ width: w })), rows };
  };

  for (const node of Array.from(body.childNodes)) {
    index++;
    if (node.nodeType !== 1) {
      // Whitespace between body elements. Word ignores it, but keeping it
      // means an untouched document comes back byte for byte.
      const text = node.textContent ?? '';
      if (text !== '') {
        vault.opaque.push({
          xml: text.replace(/&/g, '&amp;').replace(/</g, '&lt;'),
          afterBlockId: lastBlockId,
          kind: 'whitespace',
        });
      }
      continue;
    }
    const child = node as Element;
    const name = child.localName;

    if (name === 'sectPr') {
      vault.sectPrXml = serialize(child);
      continue;
    }

    if (name === 'p') {
      const id = bodyBlockId(index);
      blocks.push(readParagraph(child, id));
      lastBlockId = id;
      // A w:sectPr inside a paragraph's properties ends a section AT that
      // paragraph: the properties describe the section just closed, and the
      // next block begins the next one.
      const inlineSect = kid(kid(child, 'pPr'), 'sectPr');
      if (inlineSect) sectBreaks.push({ sectPr: inlineSect, afterBlockId: id });
      continue;
    }

    if (name === 'tbl') {
      const id = bodyBlockId(index);
      blocks.push(readTable(child, id));
      lastBlockId = id;
      continue;
    }

    // Everything else is preserved but not rendered.
    const kindName = name === 'sdt' ? 'contentControls' : 'otherContent';
    addWarning(
      vault,
      kindName,
      kindName === 'contentControls'
        ? 'Content controls are preserved but not shown'
        : 'Some content is preserved but not shown'
    );
    vault.opaque.push({
      xml: serialize(child),
      afterBlockId: lastBlockId,
      kind: kindName,
    });
  }

  // Styles are decided once the whole document has been seen: what counts as
  // a heading depends on how big the body text is, which is not knowable
  // until every paragraph has been measured.
  const resolved = inferStyles(signals, defaultSizeHalfPt);
  paragraphs.forEach((b, i) => {
    b.styleId = resolved[i];
    vault.blockStyle.set(b.id, b.styleId);
    // Remember a real style id for each of ours, so blocks added later can
    // be written with a style this package actually defines.
    const ps = pStyles[i];
    if (ps && !vault.styleBack.has(b.styleId)) vault.styleBack.set(b.styleId, ps);
  });

  if (blocks.length === 0) {
    blocks.push({ id: newId(), styleId: 'Body', html: '' });
  }

  /* ---- headers and footers ---- */

  const sectPr = kid(body, 'sectPr');
  const headers: HeaderFooterSet = {};
  const footers: HeaderFooterSet = {};

  const readPart = (partPath: string, keyPrefix: string): ParagraphBlock[] => {
    const xmlText = partText(parts, partPath);
    if (!xmlText) return [];
    const parsedPart = new DOMParser().parseFromString(xmlText, 'application/xml');
    const root = parsedPart.documentElement;
    if (!root) return [];
    // A header part has its own relationships, for its own images.
    const partRels = readRels(
      partText(
        parts,
        partPath.replace(/^word\/(.*)$/, 'word/_rels/$1.rels')
      )
    );
    // Keep what surrounds the content, so an edited part can be rebuilt with
    // its own namespaces intact rather than generated from scratch.
    const open = xmlText.match(/<w:(?:hdr|ftr)(?:\s[^>]*)?>/);
    const closeAt = Math.max(
      xmlText.lastIndexOf('</w:hdr>'),
      xmlText.lastIndexOf('</w:ftr>')
    );
    if (open && closeAt > 0) {
      vault.hfShell.set(partPath, {
        prefix: xmlText.slice(0, (open.index ?? 0) + open[0].length),
        suffix: xmlText.slice(closeAt),
      });
    }
    const partSerialize = makeSerializer(root);
    const partCtx: Ctx = {
      rels: partRels,
      vault,
      serialize: partSerialize,
      imageCount: 1000,
      fieldDepth: 0,
      fieldInstr: '',
      fieldSkip: false,
    };
    const out: ParagraphBlock[] = [];
    kids(root, 'p').forEach((pEl, i) => {
      const pieces: Piece[] = [];
      walkInline(pEl, { b: false, i: false, u: false }, null, pieces, partCtx);
      const pPr = kid(pEl, 'pPr');
      const pStyle = wAttr(kid(pPr, 'pStyle'), 'val');
      const id = `${keyPrefix}p${i}`;
      const html = piecesToHtml(pieces);
      vault.blockXml.set(id, partSerialize(pEl));
      if (pPr) vault.blockPPr.set(id, partSerialize(pPr));
      vault.blockHtml.set(id, html);
      const styleId = recognizedStyle(pStyle, styleNames) ?? 'Body';
      vault.blockStyle.set(id, styleId);
      out.push({ id, styleId, html });
    });
    return out;
  };

  /**
   * Read one section's header and footer references.
   *
   * Sections after the first get their own key suffix, both in the vault -
   * where the part path is what an edit is written back to - and in the block
   * ids, which have to stay unique across the whole document or the caret and
   * the finder would see two blocks claiming the same id.
   */
  const readHFRefs = (
    from: Element | null,
    sectionIndex: number
  ): { headers: HeaderFooterSet; footers: HeaderFooterSet } => {
    const h: HeaderFooterSet = {};
    const f: HeaderFooterSet = {};
    if (!from) return { headers: h, footers: f };
    const tag = sectionIndex === 0 ? '' : '@' + sectionIndex;
    for (const [which, set] of [
      ['headerReference', h],
      ['footerReference', f],
    ] as [string, HeaderFooterSet][]) {
      for (const ref of kids(from, which)) {
        const type = (wAttr(ref, 'type') ?? 'default') as HFVariant;
        const rid = rAttr(ref, 'id');
        const target = rid ? rels.get(rid) : null;
        if (!target) continue;
        const partPath = target.startsWith('/')
          ? target.slice(1)
          : 'word/' + target.replace(/^\.\//, '');
        vault.hfParts.set(`${which}:${type}${tag}`, partPath);
        const read = readPart(partPath, `${type}-${which[0]}${tag}-`);
        if (read.length > 0) {
          set[type === 'even' ? 'even' : type === 'first' ? 'first' : 'default'] = read;
        }
      }
    }
    return { headers: h, footers: f };
  };

  /* ---- sections ---- */

  // One section per w:sectPr. An inline one describes the section it closes;
  // the body-level one describes the last. A section whose first block does
  // not exist - the break was on the final paragraph - is dropped, because a
  // section with nothing in it has no page to lay out.
  const blockIndex = new Map(blocks.map((b, i) => [b.id, i]));
  const sectionParts: { sectPr: Element | null; startId: string | null }[] = [];
  let nextStart: string | null = null;
  for (const brk of sectBreaks) {
    sectionParts.push({ sectPr: brk.sectPr, startId: nextStart });
    const at = blockIndex.get(brk.afterBlockId);
    nextStart = at === undefined ? null : (blocks[at + 1]?.id ?? null);
    if (nextStart === null) break; // nothing follows the break
  }
  if (nextStart !== null || sectionParts.length === 0) {
    sectionParts.push({ sectPr: kid(body, 'sectPr'), startId: nextStart });
  }

  const sections: Section[] = sectionParts.map((sp, i) => {
    const hf = readHFRefs(sp.sectPr, i);
    const type = wAttr(kid(sp.sectPr, 'type'), 'val');
    return {
      id: 'sect' + i,
      startId: sp.startId,
      page: readSectPr(sp.sectPr),
      ...(Object.keys(hf.headers).length ? { headers: hf.headers } : {}),
      ...(Object.keys(hf.footers).length ? { footers: hf.footers } : {}),
      ...(kid(sp.sectPr, 'titlePg') ? { titlePage: true } : {}),
      ...(type === 'continuous' ? { continuous: true } : {}),
    };
  });

  // The first section is what a reader sees first, and it is what everything
  // that is not section-aware falls back to.
  const first = sections[0];
  Object.assign(headers, first.headers ?? {});
  Object.assign(footers, first.footers ?? {});
  const titlePage = !!first.titlePage;
  const settings = partText(parts, 'word/settings.xml') ?? '';
  const evenOdd = /<w:evenAndOddHeaders/.test(settings);

  const pgMar = kid(sectPr, 'pgMar');
  const headerDistance = pgMar
    ? Math.round(twipToPx(intOf(wAttr(pgMar, 'header'), 720)))
    : 48;
  const footerDistance = pgMar
    ? Math.round(twipToPx(intOf(wAttr(pgMar, 'footer'), 720)))
    : 48;

  const doc: Doc = {
    id: newId(),
    title: fileName.replace(/\.docx$/i, '') || 'Untitled',
    // The first section's geometry, not the last. `page` is what a document
    // with one section is laid out with, and the document opens in its
    // first section either way.
    page: sections[0].page,
    blocks,
    ...(sections.length > 1 ? { sections } : {}),
    ...(Object.keys(headers).length ? { headers } : {}),
    ...(Object.keys(footers).length ? { footers } : {}),
    ...(titlePage ? { titlePage } : {}),
    ...(evenOdd ? { evenOdd } : {}),
    headerDistance,
    footerDistance,
  };

  return { doc, vault };
}
