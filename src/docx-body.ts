import type { Block, Doc } from './model';
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
  const xml =
    '<x>' +
    html.replace(/<br\s*\/?>/gi, '<br/>').replace(/&nbsp;/g, '\u00a0') +
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

export function stripTags(html: string): string {
  return html
    .replace(/<[^>]*>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, '\u00a0')
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&');
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

function paragraphXml(b: Block, vault: Vault): string {
  const original = vault.blockXml.get(b.id);
  const sameText = vault.blockHtml.get(b.id) === b.html;
  const sameStyle = vault.blockStyle.get(b.id) === b.styleId;
  if (original && sameText && sameStyle) return original;

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
  // Carry the paragraph's own run formatting - size, font, colour - into the
  // regenerated runs. Real documents keep their heading appearance there, so
  // without this, editing a heading quietly resets it to the body font.
  return '<w:p>' + pPr + runsXml(b.html, vault, vault.blockRPr.get(b.id)) + '</w:p>';
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
  while (
    blocks.length > 1 &&
    blocks[blocks.length - 1].html.trim() === '' &&
    !vault.blockXml.has(blocks[blocks.length - 1].id)
  ) {
    blocks.pop();
  }

  const present = new Set(blocks.map((b) => b.id));
  for (const b of blocks) {
    out += paragraphXml(b, vault);
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
