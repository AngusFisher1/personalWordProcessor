import {
  blockEl,
  docEl,
  logicalGroups,
  logicalIdOf,
  pieceBase,
  piecesOf,
} from './render';

/**
 * Find and replace.
 *
 * Searching runs over the MERGED LOGICAL TEXT of each paragraph, not over the
 * rendered DOM. A paragraph split across a page break is two elements, and a
 * bold word is a third; searching what is rendered would miss any match that
 * crosses either boundary, which is most of the interesting ones.
 *
 * Offsets are the same logical offsets the caret uses - block id plus a count
 * of rendered characters - so a hit can be turned back into a position no
 * matter how the paragraph is currently broken up.
 */

export interface Hit {
  blockId: string;
  start: number;
  end: number;
}

export interface FindOptions {
  caseSensitive: boolean;
  wholeWord: boolean;
  regex: boolean;
}

export const DEFAULT_OPTIONS: FindOptions = {
  caseSensitive: false,
  wholeWord: false,
  regex: false,
};

/** Highlighting every match in a huge document is not worth the DOM churn. */
const MAX_HIGHLIGHTS = 600;

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function buildPattern(query: string, o: FindOptions): RegExp | null {
  if (query === '') return null;
  let body = o.regex ? query : escapeRegex(query);
  if (o.wholeWord) body = '\\b(?:' + body + ')\\b';
  try {
    return new RegExp(body, o.caseSensitive ? 'g' : 'gi');
  } catch {
    return null; // an unfinished regex as the user types
  }
}

/** The merged text of one logical paragraph, as the caret counts it. */
function logicalText(pieces: HTMLElement[]): string {
  return pieces.map((p) => p.textContent ?? '').join('');
}

export function search(query: string, o: FindOptions): Hit[] {
  const re = buildPattern(query, o);
  if (!re) return [];
  const out: Hit[] = [];

  for (const g of logicalGroups(docEl())) {
    const id = logicalIdOf(g.head);
    const text = logicalText([g.head, ...g.tails]);
    if (!text) continue;
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text))) {
      if (m[0] === '') {
        re.lastIndex++; // a pattern that can match nothing would spin
        continue;
      }
      out.push({ blockId: id, start: m.index, end: m.index + m[0].length });
      if (out.length > 20000) return out; // pathological pattern guard
    }
  }
  return out;
}

/* ------------------------------------------------------------------ *
 * Mapping a logical range onto the DOM
 * ------------------------------------------------------------------ */

interface Segment {
  node: Text;
  from: number;
  to: number;
}

/**
 * The text-node slices a logical range covers.
 *
 * A range can span several pieces of a split paragraph and several inline
 * elements inside each, so this returns one slice per text node rather than
 * a single range - which is also why highlighting wraps several spans.
 */
function segmentsFor(hit: Hit): Segment[] {
  const out: Segment[] = [];
  for (const piece of piecesOf(hit.blockId)) {
    const base = pieceBase(piece);
    const walker = document.createTreeWalker(piece, NodeFilter.SHOW_TEXT);
    let seen = base;
    let n: Node | null;
    while ((n = walker.nextNode())) {
      const text = n as Text;
      const len = text.data.length;
      const nodeStart = seen;
      const nodeEnd = seen + len;
      seen = nodeEnd;
      if (nodeEnd <= hit.start || nodeStart >= hit.end) continue;
      out.push({
        node: text,
        from: Math.max(0, hit.start - nodeStart),
        to: Math.min(len, hit.end - nodeStart),
      });
    }
  }
  return out;
}

/* ------------------------------------------------------------------ *
 * Highlighting
 * ------------------------------------------------------------------ */

export function clearHighlights(): void {
  const marks = Array.from(docEl().querySelectorAll('span.find-hit'));
  for (const m of marks) {
    const parent = m.parentNode;
    if (!parent) continue;
    while (m.firstChild) parent.insertBefore(m.firstChild, m);
    parent.removeChild(m);
  }
  // Re-join the text nodes splitText left behind, so offsets stay simple.
  for (const blk of Array.from(docEl().querySelectorAll('.blk'))) blk.normalize();
}

/**
 * Wrap each match. The marker is a plain span, which the inline sanitizer
 * unwraps on the way back to the model, so a highlight can never be saved or
 * exported.
 */
export function highlightAll(hits: Hit[], currentIndex: number): void {
  clearHighlights();
  const limit = Math.min(hits.length, MAX_HIGHLIGHTS);
  for (let i = 0; i < limit; i++) {
    const segments = segmentsFor(hits[i]);
    for (const seg of segments) {
      let node = seg.node;
      if (seg.to < node.data.length) node.splitText(seg.to);
      if (seg.from > 0) node = node.splitText(seg.from);
      const span = document.createElement('span');
      span.className = 'find-hit' + (i === currentIndex ? ' on' : '');
      node.parentNode?.insertBefore(span, node);
      span.appendChild(node);
    }
  }
}

/**
 * Where each hit sits down the document, 0 to 1, for the document map.
 *
 * Deduplicated and capped BEFORE any DOM lookup. Resolving an element per
 * hit means a document-wide query per hit, which on a few thousand matches
 * costs more than the search and the replace put together - and the map is
 * a few hundred pixels tall, so the extra ticks would land on each other
 * anyway.
 */
export function hitPositions(hits: Hit[], max = 200): number[] {
  const seen = new Set<string>();
  const ids: string[] = [];
  for (const h of hits) {
    if (seen.has(h.blockId)) continue;
    seen.add(h.blockId);
    ids.push(h.blockId);
    if (ids.length >= max) break;
  }

  const doc = docEl();
  const total = doc.scrollHeight || 1;
  const docTop = doc.getBoundingClientRect().top;
  const out: number[] = [];
  for (const id of ids) {
    const head = blockEl(id);
    if (!head) continue;
    const top = head.getBoundingClientRect().top - docTop + doc.scrollTop;
    out.push(Math.max(0, Math.min(1, top / total)));
  }
  return out;
}

export function scrollToHit(hit: Hit): void {
  const segments = segmentsFor(hit);
  const node = segments[0]?.node;
  const target = (node?.parentElement ?? piecesOf(hit.blockId)[0]) as
    | HTMLElement
    | undefined;
  if (!target) return;
  const doc = docEl();
  const r = target.getBoundingClientRect();
  const d = doc.getBoundingClientRect();
  if (r.top < d.top + 60 || r.bottom > d.bottom - 60) {
    doc.scrollTop += r.top - d.top - doc.clientHeight / 3;
  }
}

/* ------------------------------------------------------------------ *
 * Replacing
 * ------------------------------------------------------------------ */

/**
 * Replace one hit in the DOM.
 *
 * Done in the DOM rather than on the model's html because the offsets are in
 * text space: mapping them into an html string means counting past tags and
 * entities, and getting that subtly wrong corrupts markup rather than text.
 */
export function replaceHit(hit: Hit, text: string, prune = true): void {
  const segments = segmentsFor(hit);
  if (segments.length === 0) return;
  // Later slices first, so earlier offsets stay valid.
  for (let i = segments.length - 1; i >= 1; i--) {
    const s = segments[i];
    s.node.deleteData(s.from, s.to - s.from);
  }
  const first = segments[0];
  first.node.replaceData(first.from, first.to - first.from, text);
  if (prune) pruneEmptyInline(hit.blockId);
}

/**
 * A match that straddled a <b> leaves the emptied half behind, because the
 * replacement text all goes into the first slice. An empty <b></b> is
 * harmless but it is litter, and it accumulates over a replace-all.
 */
export function pruneEmptyInline(blockId: string): void {
  for (const piece of piecesOf(blockId)) {
    for (const el of Array.from(piece.querySelectorAll('b, i, u, a'))) {
      if ((el.textContent ?? '') === '' && !el.querySelector('img, br')) {
        el.remove();
      }
    }
  }
}

/**
 * Replace every hit, then let the caller reflow once. Hits are applied from
 * the end backwards so that each replacement leaves the offsets of the ones
 * before it untouched.
 */
export function replaceAll(hits: Hit[], text: string): number {
  const ordered = hits.slice().sort((a, b) => {
    if (a.blockId === b.blockId) return b.start - a.start;
    return 0;
  });
  // Group by block and apply each block's hits back to front.
  const byBlock = new Map<string, Hit[]>();
  for (const h of ordered) {
    const list = byBlock.get(h.blockId);
    if (list) list.push(h);
    else byBlock.set(h.blockId, [h]);
  }
  let n = 0;
  for (const [blockId, list] of byBlock) {
    for (const h of list.sort((a, b) => b.start - a.start)) {
      replaceHit(h, text, false);
      n++;
    }
    pruneEmptyInline(blockId); // once per block, not once per hit
  }
  return n;
}
