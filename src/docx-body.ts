import type { Block, BlockFormat, Doc, ParagraphBlock, TableBlock } from './model';
import { isTable, plainText, sameFormat, sectionsOf } from './model';
import type { RunProp, Vault } from './docx-package';
import { addWarning } from './docx-package';

/**
 * Regenerate the body of word/document.xml for a document that came from a
 * .docx, so it can be dropped back into its original package.
 *
 * The governing rule is that anything the user did not touch goes back out
 * byte-identical. A paragraph is only regenerated when its text or its style
 * actually changed, and even then it keeps its original w:pPr, so indentation,
 * spacing, numbering and direct formatting survive an edit to the words.
 */

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escAttr(s: string): string {
  return esc(s).replace(/"/g, '&quot;');
}

interface Flags {
  b: boolean;
  i: boolean;
  u: boolean;
}

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

/**
 * Parse our inline markup as XML rather than HTML.
 *
 * The markup is ours and tiny (b/i/u/a/br), so XML parsing is safe once <br>
 * is closed and &nbsp; is resolved - and unlike text/html it works outside a
 * browser, which is what lets the round-trip harness run in plain Node.
 */
export function parseInline(html: string): Element | null {
  // Void elements have to be closed for an XML parser. Missing one here does
  // not fail loudly: the parse fails, the fallback emits the text alone, and
  // the element - an image, say - is silently gone from the export.
  const xml =
    '<x>' +
    html
      .replace(/<br\s*\/?>/gi, '<br/>')
      .replace(/<img\b([^>]*?)\/?>/gi, '<img$1/>')
      .replace(/&nbsp;/g, '\u00a0') +
    '</x>';
  try {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const el = doc.documentElement;
    if (!el || el.getElementsByTagName('parsererror').length > 0) return null;
    return el;
  } catch {
    return null;
  }
}

/**
 * Text with the markup and the entities both gone.
 *
 * One shared implementation, because decoding in several passes gets the
 * order wrong: unescaping `&amp;` before `&lt;` turns the literal text
 * "&lt;" into a "<" that was never written.
 */
export function stripTags(html: string): string {
  return plainText(html);
}

/**
 * w:rPr children have a required order, and Word treats a run whose
 * properties are out of order as a repair-worthy error. Preserved properties
 * are merged with the bold/italic/underline we manage and then sorted back
 * into this order.
 */
const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
  'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
  'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz',
  'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign',
  'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath',
];

function rPrXml(f: Flags, base: RunProp[] | undefined): string {
  const props: RunProp[] = base ? base.slice() : [];
  if (f.b) props.push({ name: 'b', xml: '<w:b/>' });
  if (f.i) props.push({ name: 'i', xml: '<w:i/>' });
  if (f.u) props.push({ name: 'u', xml: '<w:u w:val="single"/>' });
  if (props.length === 0) return '';
  props.sort((a, b) => {
    const ia = RPR_ORDER.indexOf(a.name);
    const ib = RPR_ORDER.indexOf(b.name);
    return (ia < 0 ? RPR_ORDER.length : ia) - (ib < 0 ? RPR_ORDER.length : ib);
  });
  return '<w:rPr>' + props.map((p) => p.xml).join('') + '</w:rPr>';
}

function runXml(text: string, f: Flags, base?: RunProp[]): string {
  if (text === '') return '';
  return (
    '<w:r>' +
    rPrXml(f, base) +
    '<w:t xml:space="preserve">' +
    esc(text) +
    '</w:t></w:r>'
  );
}

/** Walk the block's inline markup into runs. Never regex the HTML. */
function runsXml(html: string, vault: Vault, base?: RunProp[]): string {
  const root = parseInline(html);
  // Unparseable markup still has to export as its words rather than vanish.
  if (!root) return runXml(stripTags(html), { b: false, i: false, u: false }, base);

  let out = '';

  const walk = (node: Node, f: Flags, href: string | null): void => {
    for (const n of Array.from(node.childNodes)) {
      if (n.nodeType === TEXT_NODE) {
        out += runXml(n.textContent ?? '', f, base);
        continue;
      }
      if (n.nodeType !== ELEMENT_NODE) continue;
      const el = n as Element;
      switch (el.tagName.toUpperCase()) {
        case 'BR':
          out += '<w:r><w:br/></w:r>';
          break;
        case 'IMG': {
          // Write the original run back untouched. A w:drawing carries
          // cropping, effects and positioning that cannot be rebuilt from
          // an <img>, so regenerating one would quietly degrade it.
          const tok = el.getAttribute('data-run');
          const kept = tok ? vault.runXml.get(tok) : undefined;
          if (kept) out += kept;
          else addWarning(vault, 'images', 'An image could not be written back');
          break;
        }
        case 'B':
          walk(el, { ...f, b: true }, href);
          break;
        case 'I':
          walk(el, { ...f, i: true }, href);
          break;
        case 'U':
          walk(el, { ...f, u: true }, href);
          break;
        case 'A': {
          const target = el.getAttribute('href') ?? href;
          // Only links the package already has a relationship for can be
          // written as hyperlinks; inventing one would mean rewriting
          // document.xml.rels, and then the "untouched parts are identical"
          // guarantee no longer holds.
          const rid = target ? vault.relByTarget.get(target) : undefined;
          if (rid) {
            out += '<w:hyperlink r:id="' + escAttr(rid) + '">';
            walk(el, f, target);
            out += '</w:hyperlink>';
          } else {
            if (target) {
              addWarning(
                vault,
                'newLinks',
                'Links added here export as plain text'
              );
            }
            walk(el, f, target);
          }
          break;
        }
        default:
          walk(el, f, href);
      }
    }
  };

  walk(root, { b: false, i: false, u: false }, null);
  return out;
}

/**
 * w:pPr children have a required order, and Word offers to repair a file
 * whose paragraph properties are out of it. This is the CT_PPr sequence, as
 * far as anything we write reaches.
 */
const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
  'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
  'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
  'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
  'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
  'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr',
  'pPrChange',
];

interface PPrChild {
  name: string;
  start: number;
  end: number;
}

/**
 * The top-level children of a w:pPr, with their offsets.
 *
 * Depth-aware rather than a flat regex: w:numPr, w:rPr and w:sectPr have
 * children of their own, and a pattern that ignored nesting would report
 * w:sectPr's own w:jc as a sibling and write the section's alignment onto
 * the paragraph.
 */
function pPrChildren(pPr: string): PPrChild[] {
  const out: PPrChild[] = [];
  const tag = /<(\/?)w:([A-Za-z0-9]+)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  let depth = 0;
  let open: { name: string; start: number } | null = null;
  let m: RegExpExecArray | null;
  while ((m = tag.exec(pPr))) {
    const [whole, closing, name, , selfClosing] = m;
    if (closing) {
      depth--;
      if (depth === 1 && open && open.name === name) {
        out.push({ name, start: open.start, end: m.index + whole.length });
        open = null;
      }
      continue;
    }
    if (selfClosing) {
      if (depth === 1) out.push({ name, start: m.index, end: m.index + whole.length });
      continue;
    }
    depth++;
    if (depth === 2 && !open) open = { name, start: m.index };
  }
  return out;
}

/**
 * Set or remove one empty w:pPr child, keeping the schema order.
 *
 * Only ever called for w:jc, w:ind and w:spacing, all of which are empty
 * elements, so nothing here has to preserve inner content.
 */
function setPPrChild(pPr: string, name: string, xml: string | null): string {
  const hasWrapper = /^<w:pPr(\s[^>]*)?>/.test(pPr);
  let body = hasWrapper
    ? pPr.replace(/^<w:pPr(\s[^>]*)?>/, '').replace(/<\/w:pPr>$/, '')
    : pPr;
  const openTag = hasWrapper ? (pPr.match(/^<w:pPr(\s[^>]*)?>/) as RegExpMatchArray)[0] : '<w:pPr>';

  const existing = pPrChildren('<w:pPr>' + body + '</w:pPr>').find((c) => c.name === name);
  if (existing) {
    // Offsets are into the wrapped string, so shift by the opening tag.
    const shift = '<w:pPr>'.length;
    body = body.slice(0, existing.start - shift) + body.slice(existing.end - shift);
  }
  if (xml) {
    const rank = PPR_ORDER.indexOf(name);
    const children = pPrChildren('<w:pPr>' + body + '</w:pPr>');
    const shift = '<w:pPr>'.length;
    const after = children.find((c) => {
      const r = PPR_ORDER.indexOf(c.name);
      return r < 0 ? false : r > rank;
    });
    const at = after ? after.start - shift : body.length;
    body = body.slice(0, at) + xml + body.slice(at);
  }
  return body === '' && !xml && !hasWrapper ? '' : openTag + body + '</w:pPr>';
}

/** Points back to twips, the unit w:pPr measures in. */
const ptToTwip = (pt: number) => Math.round(pt * 20);

/**
 * Write a paragraph's direct formatting into its w:pPr.
 *
 * Absent values REMOVE the element rather than leaving the imported one, so
 * setting a centred paragraph back to left actually un-centres it in Word
 * instead of silently keeping w:jc.
 */
function withFormat(pPr: string, f: BlockFormat | undefined): string {
  let out = pPr || '<w:pPr></w:pPr>';

  const align = f?.align;
  out = setPPrChild(
    out,
    'jc',
    align && align !== 'left'
      ? '<w:jc w:val="' + (align === 'justify' ? 'both' : align) + '"/>'
      : null
  );

  const hasInd =
    f?.indentLeft !== undefined || f?.indentRight !== undefined || f?.firstLine !== undefined;
  out = setPPrChild(
    out,
    'ind',
    hasInd
      ? '<w:ind' +
          (f?.indentLeft !== undefined ? ' w:left="' + ptToTwip(f.indentLeft) + '"' : '') +
          (f?.indentRight !== undefined ? ' w:right="' + ptToTwip(f.indentRight) + '"' : '') +
          (f?.firstLine !== undefined && f.firstLine < 0
            ? ' w:hanging="' + ptToTwip(-f.firstLine) + '"'
            : f?.firstLine !== undefined && f.firstLine > 0
              ? ' w:firstLine="' + ptToTwip(f.firstLine) + '"'
              : '') +
          '/>'
      : null
  );

  const hasSpacing =
    f?.spaceBefore !== undefined || f?.spaceAfter !== undefined || f?.lineHeight !== undefined;
  out = setPPrChild(
    out,
    'spacing',
    hasSpacing
      ? '<w:spacing' +
          (f?.spaceBefore !== undefined ? ' w:before="' + ptToTwip(f.spaceBefore) + '"' : '') +
          (f?.spaceAfter !== undefined ? ' w:after="' + ptToTwip(f.spaceAfter) + '"' : '') +
          (f?.lineHeight !== undefined
            ? ' w:line="' +
              (f.lineRule === 'exact' || f.lineRule === 'atLeast'
                ? ptToTwip(f.lineHeight)
                : Math.round(f.lineHeight * 240)) +
              '" w:lineRule="' +
              (f.lineRule ?? 'auto') +
              '"'
            : '') +
          '/>'
      : null
  );

  // An empty w:pPr is legal but noise; drop it if nothing is left.
  return /^<w:pPr(\s[^>]*)?><\/w:pPr>$/.test(out) ? '' : out;
}

/** Put our style id into a preserved w:pPr, replacing any pStyle already there. */
function withStyle(pPr: string, styleId: string | undefined): string {
  if (!styleId) return pPr;
  const tag = '<w:pStyle w:val="' + escAttr(styleId) + '"/>';
  if (/<w:pStyle\b[^>]*\/>/.test(pPr)) {
    return pPr.replace(/<w:pStyle\b[^>]*\/>/, tag);
  }
  // pStyle must come first inside w:pPr.
  const m = pPr.match(/^<w:pPr(\s[^>]*)?>/);
  if (!m) return pPr;
  return pPr.slice(0, m[0].length) + tag + pPr.slice(m[0].length);
}

function blockXml(b: Block, vault: Vault): string {
  if (isTable(b)) return tableXml(b, vault);
  return paragraphXml(b, vault);
}

/** True when nothing inside the table has changed since it was imported. */
function tableUnchanged(t: TableBlock, vault: Vault): boolean {
  if (!vault.blockXml.has(t.id)) return false;
  for (const row of t.rows) {
    for (const cell of row.cells) {
      for (const p of cell) {
        if (!vault.blockXml.has(p.id)) return false;
        if (vault.blockHtml.get(p.id) !== p.html) return false;
        if (vault.blockStyle.get(p.id) !== p.styleId) return false;
      }
    }
  }
  return true;
}

/**
 * An untouched table goes back verbatim. An edited one is rebuilt keeping its
 * w:tblPr, w:tblGrid and every w:trPr and w:tcPr, so borders, widths, shading
 * and merges survive an edit to the words inside a cell.
 */
function tableXml(t: TableBlock, vault: Vault): string {
  const original = vault.blockXml.get(t.id);
  if (original && tableUnchanged(t, vault)) return original;

  const pr = vault.tablePr.get(t.id);
  // The preserved grid describes the columns the table HAD. Adding or
  // removing one makes it disagree with the rows, which is the kind of
  // mismatch Word offers to repair, so regenerate it when the count moved.
  const keptCols = (pr?.tblGrid.match(/<w:gridCol[ />]/g) ?? []).length;
  const grid =
    keptCols === t.cols.length && pr?.tblGrid
      ? pr.tblGrid
      : '<w:tblGrid>' +
        t.cols
          .map((c) => `<w:gridCol w:w="${Math.max(1, Math.round(c.width * 15))}"/>`)
          .join('') +
        '</w:tblGrid>';
  let out = '<w:tbl>' + (pr?.tblPr ?? '') + grid;
  for (const row of t.rows) {
    out += '<w:tr>' + (vault.rowPr.get(row.id) ?? '');
    for (let i = 0; i < row.cells.length; i++) {
      const cell = row.cells[i];
      // Placeholders stand for columns a gridSpan absorbed; the real cell's
      // preserved w:tcPr already carries that span.
      if (cell.length === 0) continue;
      out += '<w:tc>' + (vault.cellPr.get(`${row.id}c${i}`) ?? '');
      for (const p of cell) out += paragraphXml(p, vault);
      out += '</w:tc>';
    }
    out += '</w:tr>';
  }
  return out + '</w:tbl>';
}

function paragraphXml(b: ParagraphBlock, vault: Vault): string {
  const original = vault.blockXml.get(b.id);
  const sameText = vault.blockHtml.get(b.id) === b.html;
  const sameStyle = vault.blockStyle.get(b.id) === b.styleId;
  const sameFmt = sameFormat(vault.blockFmt.get(b.id), b.fmt);
  if (original && sameText && sameStyle && sameFmt) return original;

  const backing = vault.styleBack.get(b.styleId);
  let pPr = vault.blockPPr.get(b.id) ?? '';
  if (!sameStyle) {
    pPr = pPr
      ? withStyle(pPr, backing)
      : backing
        ? '<w:pPr><w:pStyle w:val="' + escAttr(backing) + '"/></w:pPr>'
        : '';
  }
  if (!original && !pPr && backing) {
    pPr = '<w:pPr><w:pStyle w:val="' + escAttr(backing) + '"/></w:pPr>';
  }
  // Only rewritten when it actually changed: an untouched paragraph keeps
  // the w:pPr it arrived with, byte for byte, including the parts of it we
  // have never understood.
  if (!sameFmt) pPr = withFormat(pPr, b.fmt);
  // Carry the paragraph's own run formatting - size, font, colour - into the
  // regenerated runs. Real documents keep their heading appearance there, so
  // without this, editing a heading quietly resets it to the body font.
  return '<w:p>' + pPr + runsXml(b.html, vault, vault.blockRPr.get(b.id)) + '</w:p>';
}

/**
 * Rebuild any header or footer part whose text changed. Untouched parts are
 * left alone entirely, so a document whose letterhead was not edited still
 * exports byte-identically.
 */
export function buildHeaderParts(doc: Doc, vault: Vault): Map<string, string> {
  const out = new Map<string, string>();
  const sections = sectionsOf(doc);
  for (const [key, path] of vault.hfParts) {
    // "headerReference:first" for the first section, "…:first@2" for the
    // third: each section references its own parts, and an edit has to be
    // written back to the part it came from.
    const [which, rest] = key.split(':');
    const [variant, tag] = rest.split('@');
    const si = tag ? Number(tag) : 0;
    const section = sections[si];
    if (!section) continue;
    const set =
      which === 'headerReference'
        ? (section.headers ?? (si === 0 ? doc.headers : undefined))
        : (section.footers ?? (si === 0 ? doc.footers : undefined));
    const blocks = set?.[variant as 'default' | 'first' | 'even'];
    if (!blocks || blocks.length === 0) continue;
    const dirty = blocks.some((b) => vault.blockHtml.get(b.id) !== b.html);
    if (!dirty) continue;
    const shell = vault.hfShell.get(path);
    if (!shell) continue;
    out.set(path, shell.prefix + blocks.map((b) => paragraphXml(b, vault)).join('') + shell.suffix);
  }
  return out;
}

export function buildBody(doc: Doc, vault: Vault): string {
  let out = '';
  const emitted = new Set<number>();

  const emitAfter = (id: string | null) => {
    vault.opaque.forEach((o, i) => {
      if (o.afterBlockId === id && !emitted.has(i)) {
        out += o.xml;
        emitted.add(i);
      }
    });
  };

  emitAfter(null); // anything that preceded the first paragraph

  // Drop the trailing empty block the editor keeps for clicking below the
  // last line - it is ours, not the document's.
  const blocks = doc.blocks.slice();
  for (;;) {
    const last = blocks[blocks.length - 1];
    if (
      blocks.length > 1 &&
      last &&
      !isTable(last) &&
      last.html.trim() === '' &&
      !vault.blockXml.has(last.id)
    ) {
      blocks.pop();
    } else {
      break;
    }
  }

  const present = new Set(blocks.map((b) => b.id));
  for (const b of blocks) {
    out += blockXml(b, vault);
    emitAfter(b.id);
  }

  // Preserved content whose anchor paragraph was deleted still has to go
  // somewhere rather than be dropped.
  vault.opaque.forEach((o, i) => {
    if (emitted.has(i)) return;
    if (o.afterBlockId !== null && !present.has(o.afterBlockId)) {
      out += o.xml;
      emitted.add(i);
      addWarning(vault, 'movedContent', 'Preserved content moved to the end');
    }
  });

  if (vault.sectPrXml) out += vault.sectPrXml; // must stay last in the body
  return out;
}
