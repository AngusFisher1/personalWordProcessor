import type { Block, Doc, PageSetup, StyleId } from './model';
import { newId } from './model';
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

interface Fmt {
  b: boolean;
  i: boolean;
  u: boolean;
}

type Piece =
  | ({ t: 'text'; text: string; href: string | null } & Fmt)
  | { t: 'br' };

interface Ctx {
  rels: Map<string, string>;
  vault: Vault;
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
  for (const c of elementChildren(r)) {
    switch (c.localName) {
      case 't':
        out.push({ t: 'text', text: c.textContent ?? '', href, ...f });
        break;
      case 'br':
        out.push({ t: 'br' });
        break;
      case 'tab':
        // We have no tab stops; a space keeps the words apart.
        out.push({ t: 'text', text: ' ', href, ...f });
        break;
      case 'noBreakHyphen':
        out.push({ t: 'text', text: '-', href, ...f });
        break;
      case 'drawing':
      case 'pict':
      case 'object':
        addWarning(ctx.vault, 'images', 'Images are preserved but not shown');
        break;
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
      case 'sdtContent':
      case 'smartTag':
      case 'fldSimple':
        walkInline(c, fmt, href, out, ctx);
        break;
      default:
        break;
    }
  }
}

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
      last.href === p.href
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
    if (p.text === '') continue;
    let t = escapeHtml(p.text);
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

/** rPr children we manage from the markup instead of preserving. */
const OWNED_RUN_PROPS = new Set(['b', 'bCs', 'i', 'iCs', 'u']);

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
  const ctx: Ctx = { rels, vault };
  const counters = new ListCounters();

  const blocks: Block[] = [];
  const signals: ParaSignals[] = [];
  const pStyles: (string | null)[] = [];
  let lastBlockId: string | null = null;
  let index = -1;

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
      const pPr = kid(child, 'pPr');
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

      const pieces: Piece[] = [];
      walkInline(child, { b: false, i: false, u: false }, null, pieces, ctx);
      const html = piecesToHtml(pieces);
      const stats = paragraphStats(child, serialize);
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

      vault.blockXml.set(id, serialize(child));
      if (pPr) vault.blockPPr.set(id, serialize(pPr));
      if (stats.baseRPr.length > 0) vault.blockRPr.set(id, stats.baseRPr);
      vault.blockHtml.set(id, html);

      blocks.push({
        id,
        styleId: 'Body', // replaced below, once the whole document is known
        html,
        ...(listMarker !== undefined ? { listMarker } : {}),
        ...(listLevel ? { listLevel } : {}),
      });
      lastBlockId = id;
      continue;
    }

    // Everything else is preserved but not rendered.
    const kindName =
      name === 'tbl'
        ? 'tables'
        : name === 'sdt'
          ? 'contentControls'
          : 'otherContent';
    addWarning(
      vault,
      kindName,
      kindName === 'tables'
        ? 'Tables are preserved but not shown'
        : kindName === 'contentControls'
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
  blocks.forEach((b, i) => {
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

  const doc: Doc = {
    id: newId(),
    title: fileName.replace(/\.docx$/i, '') || 'Untitled',
    page: readSectPr(kid(body, 'sectPr')),
    blocks,
  };

  return { doc, vault };
}
