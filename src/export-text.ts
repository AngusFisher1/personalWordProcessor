import type { Block, Doc, ParagraphBlock, StyleId, TableBlock } from './model';
import { PX_PER_INCH, isTable } from './model';
import { DOC_FONT, INK, RULE_COLOR, STYLES, px } from './styles';
import { parseInline, stripTags } from './docx-body';

/**
 * Plain formats.
 *
 * Everything else in this program deepens an investment in Microsoft's file
 * format, which is the right trade because .docx is what the world sends.
 * These exist so the documents are not trapped in it either: Markdown is
 * diffable and readable in fifty years with no software at all, and a single
 * self-contained HTML file opens anywhere.
 */

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

/* ------------------------------------------------------------------ *
 * Markdown
 * ------------------------------------------------------------------ */

/**
 * How an image's part path becomes something a reader can open. Markdown and
 * HTML are text files with no package around them, so the caller decides:
 * a data URI, a path next to the file, or nothing at all.
 */
export type MediaResolver = (part: string) => string | null;

/**
 * Characters that would otherwise be read as markup.
 *
 * Only the ones that bite mid-line. Escaping every `-` and `.` as well is
 * safe but turns readable prose into a thicket of backslashes, which defeats
 * the point of exporting Markdown; line-leading markers are handled where
 * the line is assembled, which is the only place they mean anything.
 */
function escapeMd(s: string): string {
  return s.replace(/([\\`*_[\]<])/g, '\\$1');
}

/** A line of body text that would accidentally open a list, heading or quote. */
function escapeMdLine(s: string): string {
  return s
    .replace(/^(\s*)([#>+-])(\s)/, (_m, lead: string, mark: string, sp: string) =>
      lead + '\\' + mark + sp
    )
    .replace(/^(\s*)(\d+)([.)])(\s)/, (_m, lead: string, n: string, dot: string, sp: string) =>
      lead + n + '\\' + dot + sp
    );
}

function inlineToMarkdown(html: string, media?: MediaResolver): string {
  const root = parseInline(html);
  if (!root) return escapeMd(stripTags(html));
  let out = '';

  const walk = (node: Node, bold: boolean, italic: boolean, href: string | null): void => {
    for (const n of Array.from(node.childNodes)) {
      if (n.nodeType === TEXT_NODE) {
        let t = escapeMd(n.textContent ?? '');
        if (t === '') continue;
        if (bold) t = '**' + t + '**';
        if (italic) t = '*' + t + '*';
        if (href) t = '[' + t + '](' + href + ')';
        out += t;
        continue;
      }
      if (n.nodeType !== ELEMENT_NODE) continue;
      const el = n as Element;
      switch (el.tagName.toUpperCase()) {
        case 'BR':
          out += '  \n'; // two spaces is Markdown's hard line break
          break;
        case 'B':
          walk(el, true, italic, href);
          break;
        case 'I':
          walk(el, bold, true, href);
          break;
        case 'U':
          // Markdown has no underline, and inline HTML is the honest answer.
          out += '<u>';
          walk(el, bold, italic, href);
          out += '</u>';
          break;
        case 'A':
          walk(el, bold, italic, el.getAttribute('href') || href);
          break;
        case 'IMG': {
          const part = el.getAttribute('data-media') ?? '';
          const src = media ? media(part) : part;
          const alt = el.getAttribute('alt') || part.split('/').pop() || 'image';
          // No resolver, no file: say a picture was here rather than link
          // at a package the reader does not have.
          out += src ? '![' + alt + '](' + src + ')' : '`[image: ' + alt + ']`';
          break;
        }
        default:
          walk(el, bold, italic, href);
      }
    }
  };

  walk(root, false, false, null);
  return out;
}

const MD_PREFIX: Record<StyleId, string> = {
  Name: '# ',
  Contact: '',
  SectionHeading: '## ',
  JobTitle: '### ',
  Body: '',
  Bullet: '',
};

function markdownParagraph(b: ParagraphBlock, media?: MediaResolver): string {
  const text = inlineToMarkdown(b.html, media);
  if (text.trim() === '') return '';
  if (b.styleId === 'Bullet') {
    const indent = '  '.repeat(b.listLevel ?? 0);
    // An imported numbered list keeps its number; a plain bullet gets a dash.
    const marker =
      b.listMarker && /\d|[ivxlc]+[.)]/i.test(b.listMarker)
        ? b.listMarker.trim() + ' '
        : '- ';
    return indent + marker + text;
  }
  const prefix = MD_PREFIX[b.styleId];
  return prefix + (prefix ? text : escapeMdLine(text));
}

function markdownTable(t: TableBlock, media?: MediaResolver): string {
  const rows = t.rows.map((row) =>
    row.cells.map((cell) =>
      cell
        .map((p) => inlineToMarkdown(p.html, media))
        .join(' ')
        .replace(/\|/g, '\\|')
        .replace(/\n/g, ' ')
        .trim()
    )
  );
  if (rows.length === 0) return '';
  const width = Math.max(...rows.map((r) => r.length));
  const pad = (r: string[]) => {
    const copy = r.slice();
    while (copy.length < width) copy.push('');
    return copy;
  };
  const out: string[] = [];
  out.push('| ' + pad(rows[0]).join(' | ') + ' |');
  out.push('|' + ' --- |'.repeat(width));
  for (const r of rows.slice(1)) out.push('| ' + pad(r).join(' | ') + ' |');
  return out.join('\n');
}

export function toMarkdown(doc: Doc, media?: MediaResolver): string {
  const parts: string[] = [];
  for (const b of doc.blocks) {
    const text = isTable(b) ? markdownTable(b, media) : markdownParagraph(b, media);
    if (text !== '') parts.push(text);
  }
  // Consecutive list items belong together; everything else gets a blank line.
  let out = '';
  parts.forEach((p, i) => {
    const prev = parts[i - 1];
    const bothList = prev !== undefined && isListLine(prev) && isListLine(p);
    out += (i === 0 ? '' : bothList ? '\n' : '\n\n') + p;
  });
  return out.trimEnd() + '\n';
}

function isListLine(s: string): boolean {
  return /^\s*(-|\d+[.)]|[ivxlc]+[.)])\s/i.test(s);
}

/* ------------------------------------------------------------------ *
 * HTML
 * ------------------------------------------------------------------ */

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** The named styles, as a stylesheet the exported file carries with it. */
function styleSheet(doc: Doc): string {
  const p = doc.page;
  const rule = (id: StyleId): string => {
    const d = STYLES[id];
    const [t, r, b, l] = d.padding;
    const bits = [
      `font-size:${px(d.size)}px`,
      `font-weight:${d.bold ? 700 : 400}`,
      `line-height:${d.lineHeight}`,
      `margin:0`,
      `padding:${px(t)}px ${px(r)}px ${px(b)}px ${px(l + d.hanging)}px`,
    ];
    if (d.uppercase) bits.push('text-transform:uppercase');
    if (d.letterSpacing) bits.push(`letter-spacing:${px(d.letterSpacing)}px`);
    if (d.rule) bits.push(`border-bottom:1px solid ${RULE_COLOR}`);
    if (d.hanging) bits.push(`text-indent:-${px(d.hanging)}px`);
    return `.s-${id}{${bits.join(';')}}`;
  };
  return (
    `body{margin:0;background:#f4f4f5;color:${INK};` +
    `font-family:${DOC_FONT};}` +
    `.page{width:${p.width}px;min-height:${p.height}px;box-sizing:border-box;` +
    `margin:0 auto 24px;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.2);` +
    `padding:${p.margins.top}px ${p.margins.right}px ${p.margins.bottom}px ${p.margins.left}px;}` +
    `table{border-collapse:collapse;width:100%;}` +
    `td{border:1px solid #bbb;padding:3px 5px;vertical-align:top;}` +
    `tr.hdr td{background:#f3f3f3;}` +
    `img{max-width:100%;height:auto;}` +
    `.missing-image{color:#888;font-style:italic;}` +
    (Object.keys(STYLES) as StyleId[]).map(rule).join('') +
    `@media print{body{background:#fff}` +
    `.page{width:auto;min-height:0;margin:0;box-shadow:none}}` +
    `@page{size:${p.width / PX_PER_INCH}in ${p.height / PX_PER_INCH}in;margin:0}`
  );
}

/**
 * Inline markup, near enough verbatim. Only images need a decision, because
 * their source is a path into a package this file is not carrying: with no
 * resolver the picture becomes a note saying what was there, rather than a
 * broken image icon pointing at nothing.
 */
function inlineToHtml(html: string, media?: MediaResolver): string {
  const root = parseInline(html);
  if (!root) return escapeHtml(stripTags(html));
  let out = '';
  const walk = (node: Node): void => {
    for (const n of Array.from(node.childNodes)) {
      if (n.nodeType === TEXT_NODE) {
        out += escapeHtml(n.textContent ?? '');
        continue;
      }
      if (n.nodeType !== ELEMENT_NODE) continue;
      const el = n as Element;
      const tag = el.tagName.toUpperCase();
      if (tag === 'BR') {
        out += '<br>';
      } else if (tag === 'IMG') {
        const part = el.getAttribute('data-media') ?? '';
        const src = media ? media(part) : part;
        const alt = el.getAttribute('alt') || part.split('/').pop() || 'image';
        out += src
          ? `<img alt="${escapeHtml(alt)}" src="${escapeHtml(src)}">`
          : `<span class="missing-image">[image: ${escapeHtml(alt)}]</span>`;
      } else if (tag === 'A') {
        const href = el.getAttribute('href') ?? '';
        out += `<a href="${escapeHtml(href)}">`;
        walk(el);
        out += '</a>';
      } else if (tag === 'B' || tag === 'I' || tag === 'U') {
        out += '<' + tag.toLowerCase() + '>';
        walk(el);
        out += '</' + tag.toLowerCase() + '>';
      } else {
        walk(el);
      }
    }
  };
  walk(root);
  return out;
}

function htmlBlock(b: Block, media?: MediaResolver): string {
  if (isTable(b)) {
    const rows = b.rows
      .map((row) => {
        const cells = row.cells
          .map((cell, i) => {
            if (cell.length === 0) return '';
            let span = 1;
            while (i + span < row.cells.length && row.cells[i + span].length === 0) span++;
            const inner = cell
              .map((p) => `<div class="s-${p.styleId}">${inlineToHtml(p.html, media)}</div>`)
              .join('');
            return `<td${span > 1 ? ` colspan="${span}"` : ''}>${inner}</td>`;
          })
          .join('');
        return `<tr${row.headerRow ? ' class="hdr"' : ''}>${cells}</tr>`;
      })
      .join('');
    return `<table><tbody>${rows}</tbody></table>`;
  }
  const marker = b.listMarker
    ? `<span style="display:inline-block;min-width:14px">${escapeHtml(b.listMarker)}</span>`
    : '';
  return `<div class="s-${b.styleId}">${marker}${inlineToHtml(b.html, media) || '<br>'}</div>`;
}

export function toHtml(doc: Doc, media?: MediaResolver): string {
  return (
    '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">\n' +
    `<title>${escapeHtml(doc.title)}</title>\n` +
    `<style>${styleSheet(doc)}</style>\n</head>\n<body>\n<div class="page">\n` +
    doc.blocks.map((b) => htmlBlock(b, media)).join('\n') +
    '\n</div>\n</body>\n</html>\n'
  );
}
