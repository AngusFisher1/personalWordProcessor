import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  LevelFormat,
  Packer,
  Paragraph,
  TextRun,
  UnderlineType,
} from 'docx';
import type { IParagraphStyleOptions, ParagraphChild } from 'docx';
import type { Block, Doc, MarginKey, StyleId } from './model';
import { MARGINS, PAGE_H, PAGE_W, STYLE_IDS } from './model';
import { DOCX_FONT, INK, RULE_COLOR, STYLES } from './styles';

/**
 * The model maps almost directly: Block -> Paragraph, styleId -> a named
 * paragraph style, inline tags -> TextRun flags.
 *
 * Page breaks are deliberately not exported. Word repaginates on its own and
 * forcing our breaks in would fight it.
 */

/** 1px at 96dpi = 15 twips. */
const TWIP_PER_PX = 15;
const px = (n: number) => Math.round(n * TWIP_PER_PX);
/** px -> half-points (px * 72/96 * 2), rounded to a size Word can store. */
const halfPt = (n: number) => Math.round(n * 1.5);

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
        ? { characterSpacing: Math.round(d.letterSpacing * 0.75 * 20) }
        : {}),
    },
    paragraph: {
      spacing: {
        before: px(top),
        after: px(bottom),
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
        ? { indent: { left: px(d.hanging), hanging: px(d.hanging) } }
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
  const parsed = new DOMParser().parseFromString(
    '<div>' + html + '</div>',
    'text/html'
  );
  const root = parsed.body.firstElementChild;
  const out: ParagraphChild[] = [];
  if (!root) return out;

  let pendingBreak = false;

  const walk = (
    node: Node,
    flags: RunFlags,
    href: string | null
  ): void => {
    for (const n of Array.from(node.childNodes)) {
      if (n.nodeType === Node.TEXT_NODE) {
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
      if (n.nodeType !== Node.ELEMENT_NODE) continue;
      const el = n as HTMLElement;
      switch (el.tagName) {
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

function toParagraph(b: Block): Paragraph {
  const isBullet = STYLES[b.styleId].bullet;
  return new Paragraph({
    style: b.styleId,
    children: runsFrom(b.html),
    ...(isBullet ? { numbering: { reference: BULLET_REF, level: 0 } } : {}),
  });
}

function trimTrailingEmpty(blocks: Block[]): Block[] {
  const out = blocks.slice();
  while (out.length > 1 && out[out.length - 1].html.trim() === '') out.pop();
  return out;
}

function pageMargin(margin: MarginKey) {
  const t = px(MARGINS[margin]);
  return { top: t, right: t, bottom: t, left: t };
}

export async function exportDocx(doc: Doc): Promise<Blob> {
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
                    left: px(STYLES.Bullet.hanging),
                    hanging: px(STYLES.Bullet.hanging),
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
            size: { width: px(PAGE_W), height: px(PAGE_H) },
            margin: pageMargin(doc.margin),
          },
        },
        children: trimTrailingEmpty(doc.blocks).map(toParagraph),
      },
    ],
  });

  return Packer.toBlob(file);
}
