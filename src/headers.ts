import type { Doc, HFVariant, ParagraphBlock, Section } from './model';
import { contentWidth, hfFor, hfVariant, sectionsOf } from './model';
import { makeBlockEl, measureEl, pageContent, pages, readBlockHtml } from './render';

/**
 * Headers and footers.
 *
 * These are why the paginator's content height stops being a constant.
 * A page with a two-line letterhead holds less than one without, the first
 * page can differ from the rest, and even pages can differ from odd - so the
 * space available for body text is a function of the page number.
 *
 * Field values are resolved at render and never written into the model. Bake
 * "Page 3 of 12" into the stored text and every page's footer becomes a
 * different string, and the export writes literals where Word expects a
 * field that renumbers itself.
 */

export interface FieldContext {
  page: number;
  pages: number;
  title: string;
}

/** A field token, as stored: <span data-field="PAGE"></span>. */
const FIELD_RE = /<span([^>]*?)data-field="([A-Z]+)"([^>]*?)><\/span>/g;

function fieldValue(name: string, ctx: FieldContext): string {
  switch (name) {
    case 'PAGE':
      return String(ctx.page);
    case 'NUMPAGES':
      return String(ctx.pages);
    case 'TITLE':
      return ctx.title;
    case 'FILENAME':
      return ctx.title + '.docx';
    case 'DATE':
      return new Date().toLocaleDateString();
    default:
      return '';
  }
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Fill in field tokens for one page. The token itself stays in the markup. */
export function resolveFields(html: string, ctx: FieldContext): string {
  FIELD_RE.lastIndex = 0;
  return html.replace(
    FIELD_RE,
    (_m, pre: string, name: string, post: string) =>
      `<span${pre}data-field="${name}"${post}>${escapeHtml(fieldValue(name, ctx))}</span>`
  );
}

function buildBlocks(
  blocks: ParagraphBlock[],
  ctx: FieldContext,
  stamp?: { page: number; which: 'header' | 'footer'; variant: HFVariant }
): DocumentFragment {
  const frag = document.createDocumentFragment();
  blocks.forEach((b, i) => {
    const el = makeBlockEl({ ...b, html: resolveFields(b.html, ctx) });
    if (stamp) {
      // Every page renders its own copy of the same header, so the ids have
      // to be made unique or the caret and the finder would see duplicates.
      el.dataset.blockId = b.id + '@' + stamp.page;
      el.dataset.hf = `${stamp.variant}:${stamp.which}:${i}`;
    }
    frag.appendChild(el);
  });
  return frag;
}

function slot(page: HTMLElement, cls: string): HTMLElement {
  let el = page.querySelector(':scope > .' + cls) as HTMLElement | null;
  if (!el) {
    el = document.createElement('div');
    el.className = cls;
    // Not editable until the user asks: otherwise the caret wanders out of
    // the body and typing lands in the letterhead.
    el.setAttribute('contenteditable', 'false');
    if (cls === 'page-header') page.insertBefore(el, page.firstChild);
    else page.appendChild(el);
  }
  return el;
}

export function hasHeaderOrFooter(doc: Doc): boolean {
  if (anyIn(doc.headers) || anyIn(doc.footers)) return true;
  return sectionsOf(doc).some((s) => anyIn(s.headers) || anyIn(s.footers));
}

/* ------------------------------------------------------------------ *
 * Measurement
 * ------------------------------------------------------------------ */

export interface HFHeights {
  header: Record<HFVariant, number>;
  footer: Record<HFVariant, number>;
}

/** One entry per section: a section can have its own header and its own width. */
export type HFTable = HFHeights[];

const EMPTY: HFHeights = {
  header: { default: 0, first: 0, even: 0 },
  footer: { default: 0, first: 0, even: 0 },
};

export function emptyHFTable(): HFTable {
  return [{ header: { ...EMPTY.header }, footer: { ...EMPTY.footer } }];
}

/** The header and footer a section uses, falling back to the document's. */
function setsFor(doc: Doc, section: Section): {
  headers: Doc['headers'];
  footers: Doc['footers'];
} {
  return {
    headers: section.headers ?? (section.startId === null ? doc.headers : undefined),
    footers: section.footers ?? (section.startId === null ? doc.footers : undefined),
  };
}

function anyIn(s?: Record<string, unknown>): boolean {
  return !!s && Object.keys(s).length > 0;
}

/**
 * Measure each variant once, off-screen. Three measurements rather than one
 * per page: the content is the same on every page that uses a variant, and
 * the only thing that varies is a field's text, which is handled by the
 * second pagination pass.
 */
export function measureHF(doc: Doc, totalPages: number): HFTable {
  const sections = sectionsOf(doc);
  const table: HFTable = sections.map(() => ({
    header: { default: 0, first: 0, even: 0 },
    footer: { default: 0, first: 0, even: 0 },
  }));
  if (!hasHeaderOrFooter(doc)) return table;

  const m = measureEl();
  const prevWidth = m.style.width;

  sections.forEach((section, si) => {
    const sets = setsFor(doc, section);
    // Measured at the section's own width: a header that wraps to two lines
    // in a narrow section takes twice the space out of that section's body.
    m.style.width = contentWidth(section.page) + 'px';
    for (const variant of ['default', 'first', 'even'] as HFVariant[]) {
      for (const which of ['header', 'footer'] as const) {
        const blocks = hfFor(which === 'header' ? sets.headers : sets.footers, variant);
        if (blocks.length === 0) continue;
        const box = document.createElement('div');
        box.className = which === 'header' ? 'page-header' : 'page-footer';
        box.style.position = 'static';
        box.appendChild(
          buildBlocks(blocks, { page: totalPages, pages: totalPages, title: doc.title })
        );
        m.appendChild(box);
        table[si][which][variant] = box.getBoundingClientRect().height;
        box.remove();
      }
    }
  });

  m.style.width = prevWidth;
  return table;
}

/** Space the header and footer take from the body on a given page. */
export function hfSpace(
  doc: Doc,
  table: HFTable,
  pageIndex: number,
  sectionIndex = 0
): number {
  const sections = sectionsOf(doc);
  const section = sections[sectionIndex] ?? sections[0];
  const heights = table[sectionIndex] ?? table[0] ?? EMPTY;
  const v = hfVariant(doc, pageIndex, section);
  return heights.header[v] + heights.footer[v];
}

/* ------------------------------------------------------------------ *
 * Rendering
 * ------------------------------------------------------------------ */

/**
 * Draw the header and footer onto every page. Called after pagination, once
 * the page count - and therefore NUMPAGES - is known.
 */
let editing = false;

export function isEditingHF(): boolean {
  return editing;
}

/**
 * Enter or leave header editing. The body dims rather than locking, so the
 * page still reads as a page while the letterhead is being worked on.
 */
export function setEditingHF(on: boolean): void {
  editing = on;
  document.getElementById('app')?.classList.toggle('hf-edit', on);
  for (const page of pages()) {
    for (const cls of ['page-header', 'page-footer']) {
      const box = page.querySelector(':scope > .' + cls);
      box?.setAttribute('contenteditable', on ? 'true' : 'false');
    }
  }
}

/** The first header or footer block on screen, to put the caret in. */
export function firstHFBlock(): HTMLElement | null {
  for (const page of pages()) {
    const el = page.querySelector(
      ':scope > .page-header .blk, :scope > .page-footer .blk'
    );
    if (el) return el as HTMLElement;
  }
  return null;
}

/**
 * Read an edited header or footer back into the document.
 *
 * Every page shows a copy of the same content, so an edit to one has to be
 * written to the variant and mirrored onto the others - which renderHF does
 * on the next pass.
 */
export function readHF(doc: Doc): boolean {
  let changed = false;
  for (const page of pages()) {
    for (const cls of ['page-header', 'page-footer'] as const) {
      const box = page.querySelector(':scope > .' + cls) as HTMLElement | null;
      if (!box) continue;
      const rows = Array.from(box.querySelectorAll(':scope > .blk')) as HTMLElement[];
      if (rows.length === 0) continue;
      const first = rows[0].dataset.hf;
      if (!first) continue;
      const [variant, which] = first.split(':') as [HFVariant, 'header' | 'footer'];
      const section = sectionsOf(doc)[sectionOfPage(page)];
      const sets = section ? setsFor(doc, section) : { headers: doc.headers, footers: doc.footers };
      const set = which === 'header' ? sets.headers : sets.footers;
      if (!set) continue;
      const target = set[variant] ?? set.default;
      if (!target) continue;

      rows.forEach((row, i) => {
        const block = target[i];
        if (!block) return;
        const html = readBlockHtml(row);
        if (html !== block.html) {
          block.html = html;
          changed = true;
        }
      });
    }
  }
  return changed;
}

/** The section a page on screen was laid out in, as the paginator tagged it. */
export function sectionOfPage(pageEl: HTMLElement | null | undefined): number {
  return Number(pageEl?.dataset.section ?? 0) || 0;
}

export function renderHF(doc: Doc, table: HFTable): void {
  const list = pages();
  const total = list.length;
  const on = hasHeaderOrFooter(doc);
  const sections = sectionsOf(doc);

  list.forEach((page, i) => {
    if (!on) {
      page.querySelector(':scope > .page-header')?.remove();
      page.querySelector(':scope > .page-footer')?.remove();
      return;
    }
    const si = sectionOfPage(page);
    const section = sections[si] ?? sections[0];
    const heights = table[si] ?? table[0] ?? EMPTY;
    const sets = setsFor(doc, section);
    const v = hfVariant(doc, i, section);
    const ctx: FieldContext = { page: i + 1, pages: total, title: doc.title };

    for (const which of ['header', 'footer'] as const) {
      const blocks = hfFor(which === 'header' ? sets.headers : sets.footers, v);
      const cls = which === 'header' ? 'page-header' : 'page-footer';
      if (blocks.length === 0) {
        page.querySelector(':scope > .' + cls)?.remove();
        continue;
      }
      const box = slot(page, cls);
      box.textContent = '';
      box.appendChild(buildBlocks(blocks, ctx, { page: i, which, variant: v }));
      box.style.height = heights[which][v] + 'px';
      box.setAttribute('contenteditable', editing ? 'true' : 'false');
    }

    // The body box shrinks by whatever the two take.
    pageContent(page).style.height =
      'calc(var(--content-h) - ' + hfSpace(doc, table, i, si) + 'px)';
  });

  if (!on) {
    for (const page of list) pageContent(page).style.removeProperty('height');
  }
}
