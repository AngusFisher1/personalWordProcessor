import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  LevelFormat,
  Packer,
  PageOrientation,
  Paragraph,
  TextRun,
  UnderlineType,
} from 'docx';
import type { IParagraphStyleOptions, ParagraphChild } from 'docx';
import type { Block, Doc, Margins, ParagraphBlock, StyleId } from './model';
import { STYLE_IDS, isTable } from './model';
import { DOCX_FONT, INK, RULE_COLOR, STYLES } from './styles';
import type { Vault } from './docx-package';
import { repack } from './docx-package';
import { buildBody, parseInline, stripTags } from './docx-body';

/**
 * The model maps almost directly: Block -> Paragraph, styleId -> a named
 * paragraph style, inline tags -> TextRun flags.
 *
 * Page breaks are deliberately not exported. Word repaginates on its own and
 * forcing our breaks in would fight it.
 */

/** 1px at 96dpi = 15 twips; 1pt = 20 twips. */
const TWIP_PER_PX = 15;
const px = (n: number) => Math.round(n * TWIP_PER_PX);
/** The style table is already in points, so this is just doubling. */
const halfPt = (pt: number) => Math.round(pt * 2);
/** Points to twips, for the spacing values in the style table. */
const ptTwip = (pt: number) => Math.round(pt * 20);

const BULLET_REF = 'wp-bullets';

function paragraphStyle(id: StyleId): IParagraphStyleOptions {
  const d = STYLES[id];
  const [top, , bottom] = d.padding;
  return {
    id,
    name: d.label,
    basedOn: 'Normal',
    next: d.enterTo,
    quickFormat: true,
    run: {
      font: DOCX_FONT,
      size: halfPt(d.size),
      bold: d.bold,
      allCaps: d.uppercase,
      color: INK.replace('#', ''),
      ...(d.letterSpacing
        ? { characterSpacing: Math.round(d.letterSpacing * 20) }
        : {}),
    },
    paragraph: {
      spacing: {
        before: ptTwip(top),
        after: ptTwip(bottom),
        line: Math.round(240 * d.lineHeight),
        lineRule: 'auto' as const,
      },
      ...(d.rule
        ? {
            border: {
              bottom: {
                style: BorderStyle.SINGLE,
                size: 6, // eighths of a point
                space: 1,
                color: RULE_COLOR.replace('#', ''),
              },
            },
          }
        : {}),
      ...(d.hanging
        ? { indent: { left: ptTwip(d.hanging), hanging: ptTwip(d.hanging) } }
        : {}),
    },
  };
}

interface RunFlags {
  bold: boolean;
  italics: boolean;
  underline: boolean;
}

/** Walk the sanitized inline markup into runs. Never regex the HTML. */
function runsFrom(html: string): ParagraphChild[] {
  const root = parseInline(html);
  const out: ParagraphChild[] = [];
  if (!root) {
    const text = stripTags(html);
    return text ? [new TextRun({ text })] : out;
  }

  let pendingBreak = false;

  const walk = (
    node: Node,
    flags: RunFlags,
    href: string | null
  ): void => {
    for (const n of Array.from(node.childNodes)) {
      if (n.nodeType === 3) {
        const text = n.textContent ?? '';
        if (text === '') continue;
        const run = new TextRun({
          text,
          bold: flags.bold || undefined,
          italics: flags.italics || undefined,
          underline:
            flags.underline || href
              ? { type: UnderlineType.SINGLE }
              : undefined,
          break: pendingBreak ? 1 : undefined,
        });
        pendingBreak = false;
        out.push(href ? new ExternalHyperlink({ children: [run], link: href }) : run);
        continue;
      }
      if (n.nodeType !== 1) continue;
      const el = n as Element;
      switch (el.tagName.toUpperCase()) {
        case 'BR':
          pendingBreak = true;
          break;
        case 'B':
          walk(el, { ...flags, bold: true }, href);
          break;
        case 'I':
          walk(el, { ...flags, italics: true }, href);
          break;
        case 'U':
          walk(el, { ...flags, underline: true }, href);
          break;
        case 'A':
          walk(el, flags, el.getAttribute('href') || href);
          break;
        default:
          walk(el, flags, href);
      }
    }
  };

  walk(root, { bold: false, italics: false, underline: false }, null);
  if (out.length === 0 && pendingBreak) out.push(new TextRun({ text: '' }));
  return out;
}

function toParagraph(b: ParagraphBlock): Paragraph {
  const isBullet = STYLES[b.styleId].bullet;
  return new Paragraph({
    style: b.styleId,
    children: runsFrom(b.html),
    ...(isBullet ? { numbering: { reference: BULLET_REF, level: 0 } } : {}),
  });
}

/**
 * The fresh-export path has no table support yet, so a table is flattened to
 * its cell paragraphs rather than dropped. Imported documents never take this
 * path: they export through the vault, which preserves the table whole.
 */
function flatten(blocks: Block[]): ParagraphBlock[] {
  const out: ParagraphBlock[] = [];
  for (const b of blocks) {
    if (isTable(b)) {
      for (const row of b.rows) for (const cell of row.cells) out.push(...cell);
    } else {
      out.push(b);
    }
  }
  while (out.length > 1 && out[out.length - 1].html.trim() === '') out.pop();
  return out;
}

function pageMargin(m: Margins) {
  return { top: px(m.top), right: px(m.right), bottom: px(m.bottom), left: px(m.left) };
}

/**
 * A document that came from a .docx goes back into its own package: only the
 * body of word/document.xml is rewritten and everything else is returned
 * untouched. A document we created ourselves is generated from scratch, with
 * the six named styles defined so it behaves like a real Word file.
 */
export async function exportDocx(doc: Doc, vault?: Vault | null): Promise<Blob> {
  if (vault && vault.parts.size > 0) {
    return repack(vault, buildBody(doc, vault));
  }
  return exportFresh(doc);
}

async function exportFresh(doc: Doc): Promise<Blob> {
  const file = new Document({
    creator: 'Personal Word Processor',
    title: doc.title,
    styles: {
      default: {
        document: {
          run: { font: DOCX_FONT, size: halfPt(STYLES.Body.size) },
        },
      },
      paragraphStyles: STYLE_IDS.map(paragraphStyle),
    },
    numbering: {
      config: [
        {
          reference: BULLET_REF,
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: ptTwip(STYLES.Bullet.hanging),
                    hanging: ptTwip(STYLES.Bullet.hanging),
                  },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: px(doc.page.width),
              height: px(doc.page.height),
              orientation:
                doc.page.width > doc.page.height
                  ? PageOrientation.LANDSCAPE
                  : PageOrientation.PORTRAIT,
            },
            margin: pageMargin(doc.page.margins),
          },
        },
        children: flatten(doc.blocks).map(toParagraph),
      },
    ],
  });

  return Packer.toBlob(file);
}
