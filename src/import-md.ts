import type { Block, Doc, ParagraphBlock, StyleId, TableBlock } from './model';
import { isTable, newBlock, newId, pageSetup } from './model';

/**
 * Markdown in.
 *
 * The counterpart to export-text, and deliberately not a full CommonMark
 * implementation: it reads what this program writes, plus the subset of
 * Markdown people actually type. Nothing here is preserved on the way back
 * out, because a Markdown file has no vault - so unlike the .docx path, this
 * is a conversion, not a round trip, and it says so.
 */

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ------------------------------------------------------------------ *
 * Inline
 * ------------------------------------------------------------------ */

/**
 * Inline spans, innermost last.
 *
 * Applied to escaped text so the tags we insert are the only markup in the
 * result, and ordered so that `**bold**` is consumed before `*italic*` can
 * mistake its delimiters for a pair of its own.
 */
const INLINE: [RegExp, (m: RegExpMatchArray) => string][] = [
  [/!\[([^\]]*)\]\(([^)\s]*)[^)]*\)/g, (m) => '[' + (m[1] || 'image') + ']'],
  [/\[([^\]]+)\]\(([^)\s]*)[^)]*\)/g, (m) => `<a href="${m[2]}">${m[1]}</a>`],
  [/\*\*([^*]+)\*\*/g, (m) => '<b>' + m[1] + '</b>'],
  [/__([^_]+)__/g, (m) => '<b>' + m[1] + '</b>'],
  [/(?<![*\w])\*([^*]+)\*(?!\*)/g, (m) => '<i>' + m[1] + '</i>'],
  [/(?<![_\w])_([^_]+)_(?![_\w])/g, (m) => '<i>' + m[1] + '</i>'],
];

/**
 * Hide escaped punctuation from the span patterns.
 *
 * A literal `\*` has to survive a pass that is looking for `*`, and the
 * cheapest way to guarantee that is to make it temporarily not be one.
 */
const SHIELD = String.fromCharCode(0);

/** Every character of a run, hidden from the span patterns. */
function shield(s: string): string {
  return Array.from(s)
    .map((ch) => SHIELD + ch.charCodeAt(0).toString(16).padStart(4, '0'))
    .join('');
}

export function inlineToHtml(md: string): string {
  let s = escapeHtml(md).replace(/\\([\\`*_{}[\]()#+\-.!>|])/g, (_m, ch: string) =>
    shield(ch)
  );
  // Code spans are literal by definition, so their content is hidden before
  // anything goes looking for emphasis inside it.
  s = s.replace(/`([^`]+)`/g, (_m, code: string) => shield(code));
  // Inline HTML we ourselves emit, put back after escaping flattened it.
  s = s.replace(/&lt;(\/?)(u|b|i|br)\s*\/?&gt;/gi, '<$1$2>');
  for (const [re, fn] of INLINE) s = s.replace(re, (...a) => fn(a as unknown as RegExpMatchArray));
  return s.replace(new RegExp(SHIELD + '([0-9a-f]{4})', 'g'), (_m, hex: string) =>
    escapeHtml(String.fromCharCode(parseInt(hex, 16)))
  );
}

/* ------------------------------------------------------------------ *
 * Blocks
 * ------------------------------------------------------------------ */

const HEADING: Record<number, StyleId> = {
  1: 'Name',
  2: 'SectionHeading',
  3: 'JobTitle',
  4: 'JobTitle',
  5: 'JobTitle',
  6: 'JobTitle',
};

const ATX = /^(#{1,6})\s+(.*?)\s*#*\s*$/;
const BULLET = /^(\s*)([-*+])\s+(.*)$/;
const ORDERED = /^(\s*)(\d+[.)]|[ivxlc]+[.)])\s+(.*)$/i;
const TABLE_ROW = /^\s*\|(.*)\|\s*$/;
const DIVIDER = /^\s*\|?[\s:-]*-[\s|:-]*\|?\s*$/;
const RULE = /^\s*([-*_])(\s*\1){2,}\s*$/;

function para(styleId: StyleId, html: string, extra?: Partial<ParagraphBlock>): ParagraphBlock {
  return { ...newBlock(styleId, html), ...extra };
}

/** `| a | b |` split on unescaped pipes. */
function splitRow(line: string): string[] {
  const inner = line.replace(TABLE_ROW, '$1');
  const out: string[] = [];
  let cur = '';
  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i];
    if (ch === '\\' && inner[i + 1] === '|') {
      cur += '|';
      i++;
    } else if (ch === '|') {
      out.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out.map((c) => c.trim());
}

function readTable(lines: string[], at: number, contentW: number): [TableBlock, number] {
  const rows: string[][] = [];
  let i = at;
  let headerRow = false;
  while (i < lines.length && TABLE_ROW.test(lines[i])) {
    if (rows.length === 1 && DIVIDER.test(lines[i])) {
      headerRow = true; // the dashed line is a marker, not a row
      i++;
      continue;
    }
    rows.push(splitRow(lines[i]));
    i++;
  }
  const cols = Math.max(...rows.map((r) => r.length));
  const width = Math.floor(contentW / cols);
  const table: TableBlock = {
    kind: 'table',
    id: newId(),
    cols: Array.from({ length: cols }, (_v, c) => ({
      // The last column absorbs the rounding, so the widths still sum.
      width: c === cols - 1 ? contentW - width * (cols - 1) : width,
    })),
    rows: rows.map((cells, r) => ({
      id: newId(),
      headerRow: r === 0 && headerRow,
      cells: Array.from({ length: cols }, (_v, c) => [
        para('Body', inlineToHtml(cells[c] ?? '')),
      ]),
    })),
  };
  return [table, i];
}

export interface MarkdownOptions {
  title?: string;
  contentWidth?: number;
}

export function fromMarkdown(text: string, opts: MarkdownOptions = {}): Doc {
  const page = pageSetup('narrow');
  const contentW = opts.contentWidth ?? page.width - page.margins.left - page.margins.right;
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const blocks: Block[] = [];

  // A line ending in two spaces asked for a break; otherwise the lines of a
  // paragraph are one paragraph, as Markdown means them.
  const sep = (line: string): string => (/ {2}$/.test(line) ? '<br>' : ' ');

  /** Body lines waiting to be flushed as one paragraph. */
  let buffer: string[] = [];
  const flush = (): void => {
    if (buffer.length === 0) return;
    const html = buffer
      .map((l, i) => inlineToHtml(l.trimEnd()) + (i < buffer.length - 1 ? sep(buffer[i]) : ''))
      .join('');
    buffer = [];
    // The line under the title is the contact line, the same guess the .docx
    // importer makes, and for the same reason: it is nearly always right on
    // the documents this program exists to edit.
    const prev = blocks[blocks.length - 1];
    const contact =
      blocks.length === 1 && !!prev && !isTable(prev) && prev.styleId === 'Name';
    blocks.push(para(contact ? 'Contact' : 'Body', html));
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (line.trim() === '') {
      flush();
      continue;
    }

    if (RULE.test(line)) {
      flush();
      continue; // a horizontal rule has nowhere to go in this model
    }

    const atx = line.match(ATX);
    if (atx) {
      flush();
      blocks.push(para(HEADING[atx[1].length], inlineToHtml(atx[2])));
      continue;
    }

    if (TABLE_ROW.test(line) && !DIVIDER.test(line)) {
      flush();
      const [table, end] = readTable(lines, i, contentW);
      blocks.push(table);
      i = end - 1;
      continue;
    }

    const bullet = line.match(BULLET);
    if (bullet) {
      flush();
      blocks.push(
        para('Bullet', inlineToHtml(bullet[3]), {
          listLevel: Math.floor(bullet[1].length / 2),
        })
      );
      continue;
    }

    const ordered = line.match(ORDERED);
    if (ordered) {
      flush();
      blocks.push(
        para('Bullet', inlineToHtml(ordered[3]), {
          listLevel: Math.floor(ordered[1].length / 2),
          listMarker: ordered[2],
        })
      );
      continue;
    }

    // Setext: the underline decides, so it is read one line ahead. Checked
    // after the list patterns, or a dashed rule under a bullet would turn
    // the bullet into a heading.
    const next = lines[i + 1];
    if (next !== undefined && buffer.length === 0 && /^\s*(=+|-{2,})\s*$/.test(next)) {
      blocks.push(
        para(next.trim().startsWith('=') ? 'Name' : 'SectionHeading', inlineToHtml(line.trim()))
      );
      i++;
      continue;
    }

    if (/^\s*>\s?/.test(line)) {
      // A quote keeps its text and loses its bar; there is no quote style.
      buffer.push(line.replace(/^\s*>\s?/, ''));
      continue;
    }

    buffer.push(line);
  }
  flush();

  if (blocks.length === 0) blocks.push(newBlock('Body'));

  return {
    id: newId(),
    title: opts.title ?? firstHeading(blocks) ?? 'Untitled',
    page,
    blocks,
  };
}

function firstHeading(blocks: Block[]): string | null {
  const first = blocks[0];
  if (!first || isTable(first) || first.styleId !== 'Name') return null;
  const text = first.html.replace(/<[^>]*>/g, '').trim();
  return text === '' ? null : text;
}

