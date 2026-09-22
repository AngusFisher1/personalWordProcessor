import type { Doc } from './model';
import type { PageSetup } from './model';
import { isTable, plainText } from './model';
import { docEl, pages } from './render';
import { styleOf } from './styles';
import { currentPageSetup } from './paginate';
import type { IndexEntry } from './persist';
import { MATCH_CAP, searchLibrary } from './persist';
import { setUiState, uiState } from './uistate';
import { openMenuAt } from './ui';

/**
 * The Recto workspace chrome.
 *
 * The rule the whole layout obeys: chrome never crosses the paper edge.
 * Page numbers, the document map and the status readout all live in the
 * gutter, so the page is the only thing that looks like a document and the
 * only light in the room.
 */

const ui = {
  outline: null as HTMLElement | null,
  files: null as HTMLElement | null,
  fileList: null as HTMLElement | null,
  filter: null as HTMLInputElement | null,
  tabs: null as HTMLElement | null,
  spine: null as HTMLElement | null,
  readout: null as HTMLElement | null,
  title: null as HTMLElement | null,
  meta: null as HTMLElement | null,
  gutter: null as HTMLElement | null,
};

function el(tag: string, cls?: string, text?: string): HTMLElement {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
}

export function railEl(): HTMLElement {
  const r = document.getElementById('rail');
  if (!r) throw new Error('#rail missing');
  return r;
}

export function canvasEl(): HTMLElement {
  const c = document.getElementById('canvas');
  if (!c) throw new Error('#canvas missing');
  return c;
}

/* ------------------------------------------------------------------ *
 * Rail
 * ------------------------------------------------------------------ */

export interface RailHooks {
  onOutlineClick(blockId: string): void;
  onOpenDoc(id: string): void;
  onNewDoc(): void;
  onDuplicateDoc(id: string): void;
  onDeleteDoc(id: string): void;
  onExportLibrary(): void;
  /** The File menu, opened from the document title. */
  onTitleMenu(anchor: HTMLElement): void;
  /** Inline rename, from a double-click or F2. */
  onRename(next: string): void;
  onCollapse(): void;
}

export type RailTab = 'outline' | 'files';
let activeTab: RailTab = 'outline';

let hooks: RailHooks | null = null;

/**
 * The navigation panel.
 *
 * NAVIGATION ONLY. The title, the save state, and the two lists - nothing
 * else, ever. This panel used to hold the formatting controls, the file
 * browser, the theme picker and the page setup at once, and every feature
 * added another full-width row that pushed the file list further down.
 *
 * The rule is enforced structurally rather than by good intentions: there
 * is no longer a container for controls to be appended to. A new feature
 * has to find one of the other three zones, or the command palette.
 */
export function buildRail(rail: HTMLElement, h: RailHooks): void {
  hooks = h;
  rail.textContent = '';

  // The brand row is gone; the mark rides along with the title and gives
  // its line back to the list underneath.
  const head = el('div', 'rail-head');
  const titleRow = el('div', 'rail-title-row');
  titleRow.appendChild(el('span', 'rail-dot'));

  ui.title = el('button', 'rail-title', 'Untitled') as HTMLElement;
  ui.title.setAttribute('type', 'button');
  ui.title.setAttribute('aria-haspopup', 'menu');
  ui.title.title = 'Document menu · double-click to rename';
  ui.title.addEventListener('mousedown', (e) => e.preventDefault());
  ui.title.addEventListener('click', () => hooks?.onTitleMenu(ui.title as HTMLElement));
  ui.title.addEventListener('dblclick', (e) => {
    e.preventDefault();
    startRename();
  });
  titleRow.appendChild(ui.title);

  const caret = el('span', 'rail-title-caret');
  titleRow.appendChild(caret);

  const collapse = document.createElement('button');
  collapse.type = 'button';
  collapse.className = 'nav-collapse';
  collapse.textContent = '«';
  collapse.setAttribute('aria-label', 'Hide the side panel');
  collapse.title = 'Hide the side panel';
  collapse.addEventListener('mousedown', (e) => e.preventDefault());
  collapse.addEventListener('click', () => hooks?.onCollapse());
  titleRow.appendChild(collapse);

  head.appendChild(titleRow);

  ui.meta = el('div', 'rail-meta', '');
  head.appendChild(ui.meta);
  rail.appendChild(head);

  const tabs = el('div', 'rail-tabs');
  for (const [id, label] of [
    ['outline', 'OUTLINE'],
    ['files', 'FILES'],
  ] as [RailTab, string][]) {
    const t = el('div', 'rail-tab', label);
    t.dataset.tab = id;
    t.addEventListener('mousedown', (e) => e.preventDefault());
    t.addEventListener('click', () => setRailTab(id));
    tabs.appendChild(t);
  }
  ui.tabs = tabs;
  rail.appendChild(tabs);

  ui.outline = el('div', 'rail-outline');
  rail.appendChild(ui.outline);

  ui.files = el('div', 'rail-files');
  const filter = document.createElement('input');
  filter.type = 'text';
  filter.className = 'files-filter';
  filter.placeholder = 'Filter…';
  filter.spellcheck = false;
  filter.addEventListener('input', () => renderFiles());
  ui.filter = filter;
  ui.files.appendChild(filter);

  ui.fileList = el('div', 'files-list');
  ui.files.appendChild(ui.fileList);

  const foot = el('div', 'files-foot');
  foot.id = 'files-foot';
  ui.files.appendChild(foot);

  rail.appendChild(ui.files);

  setRailTab(uiState().navTab);
}

/**
 * Rename in place.
 *
 * A field where the title is, rather than a dialog somewhere else: the
 * document's name is the one thing in this panel you are allowed to change,
 * and it should be changed where it is written.
 */
export function startRename(): void {
  const title = ui.title;
  if (!title || title.tagName === 'INPUT') return;
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'rail-title rail-title-input';
  input.value = title.textContent ?? '';
  input.spellcheck = false;
  input.setAttribute('aria-label', 'Document name');

  const finish = (commit: boolean): void => {
    const next = input.value.trim();
    input.replaceWith(title);
    if (commit && next !== '' && next !== title.textContent) hooks?.onRename(next);
  };
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      finish(true);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      finish(false);
    }
  });
  input.addEventListener('blur', () => finish(true));

  title.replaceWith(input);
  input.focus();
  input.select();
}

/** Move the keyboard into one of the two lists. */
export function focusRailTab(tab: RailTab): void {
  setRailTab(tab);
  const first = (
    tab === 'files'
      ? ui.filter ?? ui.fileList?.querySelector('.file-row')
      : ui.outline?.querySelector('.outline-row')
  ) as HTMLElement | null;
  first?.focus();
}

export function setRailTab(tab: RailTab): void {
  activeTab = tab;
  setUiState({ navTab: tab });
  for (const t of Array.from(ui.tabs?.children ?? [])) {
    (t as HTMLElement).classList.toggle(
      'on',
      (t as HTMLElement).dataset.tab === tab
    );
  }
  if (ui.outline) ui.outline.hidden = tab !== 'outline';
  if (ui.files) ui.files.hidden = tab !== 'files';
  if (tab === 'files') {
    renderFiles();
    ui.filter?.focus();
  }
}

export function currentRailTab(): RailTab {
  return activeTab;
}

/* ------------------------------------------------------------------ *
 * Files
 * ------------------------------------------------------------------ */

let library: IndexEntry[] = [];
let currentDocId = '';
let libraryBytes = 0;

export function setLibrary(
  entries: IndexEntry[],
  currentId: string,
  bytes: number
): void {
  library = entries;
  currentDocId = currentId;
  libraryBytes = bytes;
  if (activeTab === 'files') renderFiles();
}

function ago(ts: number): string {
  const s = Math.max(0, Date.now() - ts) / 1000;
  if (s < 90) return 'JUST NOW';
  const m = s / 60;
  if (m < 60) return Math.round(m) + ' MIN AGO';
  const h = m / 60;
  if (h < 24) return Math.round(h) + ' HR AGO';
  const d = h / 24;
  if (d < 7) return Math.round(d) + ' DAYS AGO';
  return new Date(ts)
    .toLocaleDateString(undefined, { day: 'numeric', month: 'short' })
    .toUpperCase();
}

function renderFiles(): void {
  const list = ui.fileList;
  if (!list) return;
  const q = (ui.filter?.value ?? '').trim().toLowerCase();

  // Two letters is where searching the text of every document starts being
  // worth the read. Below that the filter is a filter, on titles only.
  const hits = q.length >= 2 ? searchLibrary(q) : null;
  const byId = new Map(library.map((e) => [e.id, e]));
  const shown = hits
    ? hits.map((h) => byId.get(h.id)).filter((e): e is IndexEntry => !!e)
    : q
      ? library.filter((e) => e.title.toLowerCase().includes(q))
      : library;
  const hitById = new Map((hits ?? []).map((h) => [h.id, h]));

  list.textContent = '';

  if (shown.length === 0) {
    const empty = el('div', 'outline-empty');
    empty.textContent = q ? 'NO MATCHES' : 'NO DOCUMENTS YET';
    list.appendChild(empty);
  }

  shown.forEach((e, i) => {
    const row = el('div', 'file-row');
    if (e.id === currentDocId) row.classList.add('on');
    // Recency fade: the further down the list, the quieter the entry.
    row.style.opacity = String(Math.max(0.5, 1 - i * 0.05));

    const tick = el('span', 'outline-tick');
    const body = el('div', 'file-body');
    const hit = hitById.get(e.id);
    body.append(
      el('div', 'file-title', e.title || 'Untitled'),
      el(
        'div',
        'file-meta',
        // When it is a text search the match count earns its place; the
        // rest of the time the row is a name and when you last touched it.
        hit && hit.count > 0
          ? `${hit.count >= MATCH_CAP ? MATCH_CAP + '+' : hit.count} ` +
            `MATCH${hit.count === 1 ? '' : 'ES'} · ${ago(e.updatedAt)}`
          : ago(e.updatedAt)
      )
    );
    // The line the phrase is on, so you can tell which document this is
    // without opening it.
    if (hit?.snippet) {
      const snip = el('div', 'file-snippet');
      const at = hit.snippet.toLowerCase().indexOf(q);
      if (at < 0) snip.textContent = hit.snippet;
      else {
        snip.append(
          document.createTextNode(hit.snippet.slice(0, at)),
          el('mark', 'file-mark', hit.snippet.slice(at, at + q.length)),
          document.createTextNode(hit.snippet.slice(at + q.length))
        );
      }
      body.appendChild(snip);
    }

    const more = document.createElement('button');
    more.className = 'file-more';
    more.textContent = '⋯';
    more.title = 'Document actions';
    more.addEventListener('mousedown', (ev) => ev.preventDefault());
    more.addEventListener('click', (ev) => {
      ev.stopPropagation();
      openMenuAt(more, [
        { label: 'Open', onSelect: () => hooks?.onOpenDoc(e.id) },
        { label: 'Duplicate', onSelect: () => hooks?.onDuplicateDoc(e.id) },
        { separator: true },
        { label: 'Delete…', onSelect: () => hooks?.onDeleteDoc(e.id) },
      ]);
    });

    row.append(tick, body, more);
    row.tabIndex = 0;
    row.setAttribute('role', 'option');
    row.setAttribute('aria-selected', String(e.id === currentDocId));
    row.addEventListener('mousedown', (ev) => ev.preventDefault());
    row.addEventListener('click', () => {
      if (e.id !== currentDocId) hooks?.onOpenDoc(e.id);
    });
    row.addEventListener('keydown', (ev) => listKeys(ev, '.file-row', () => {
      if (e.id !== currentDocId) hooks?.onOpenDoc(e.id);
    }));
    list.appendChild(row);
  });

  const foot = document.getElementById('files-foot');
  if (foot) {
    const kb = Math.max(1, Math.round(libraryBytes / 1024));
    foot.textContent = '';
    foot.appendChild(el('div', 'rail-label', 'ON DISK'));
    foot.appendChild(
      el(
        'div',
        'files-foot-line',
        `${library.length} DOCUMENT${library.length === 1 ? '' : 'S'} · ${kb} KB`
      )
    );
    foot.appendChild(el('div', 'files-foot-line', 'NO ACCOUNT · NO SYNC'));

    // The one command that gets everything out. It sits under the storage
    // figure on purpose: that line is where you find yourself when you start
    // wondering what happens to all this if the browser forgets it.
    const out = document.createElement('button');
    out.className = 'files-foot-btn';
    out.textContent = 'EXPORT EVERYTHING';
    out.title = 'Write every document to Markdown, JSON and HTML in one archive';
    out.addEventListener('mousedown', (ev) => ev.preventDefault());
    out.addEventListener('click', () => hooks?.onExportLibrary());
    foot.appendChild(out);
  }
}

/**
 * A page's own geometry, read back from the custom properties the paginator
 * wrote on it. Reading the element rather than the model means the panel
 * cannot disagree with what is on screen.
 */
function geometryOf(pageEl: HTMLElement | undefined): PageSetup | null {
  if (!pageEl) return null;
  const cs = getComputedStyle(pageEl);
  const px = (name: string) => parseFloat(cs.getPropertyValue(name));
  const width = px('--page-w');
  const height = px('--page-h');
  if (!Number.isFinite(width) || !Number.isFinite(height)) return null;
  return {
    width,
    height,
    margins: {
      top: px('--pad-t') || 0,
      right: px('--pad-r') || 0,
      bottom: px('--pad-b') || 0,
      left: px('--pad-l') || 0,
    },
  };
}

/**
 * A passing message in the meta line.
 *
 * Commands that produce a file - a bulk export, an import - finish silently
 * otherwise, and a download that may or may not have happened is worse than
 * no feedback at all. The last arguments are kept so the line can be put
 * back exactly as it was once the message expires.
 */
let flash: { text: string; err: boolean } | null = null;
let flashTimer = 0;
let lastRail: [Doc, number, string, boolean] | null = null;

export function flashRail(text: string, err = false, ms = 5000): void {
  flash = { text, err };
  window.clearTimeout(flashTimer);
  flashTimer = window.setTimeout(() => {
    flash = null;
    if (lastRail) updateRail(...lastRail);
  }, ms);
  if (lastRail) updateRail(...lastRail);
}

export function updateRail(
  doc: Doc,
  words: number,
  saved: string,
  failed = false
): void {
  lastRail = [doc, words, saved, failed];
  if (ui.title) ui.title.textContent = doc.title || 'Untitled';
  if (ui.meta) {
    // Save state and nothing else. The page and word counts moved to the
    // status bar, which is now the only place either of them appears.
    ui.meta.classList.toggle('err', failed || !!flash?.err);
    ui.meta.textContent = (flash ? flash.text : saved).toUpperCase();
  }

}

/**
 * Comments, under the outline.
 *
 * Not a third tab: a document with comments on it is being reviewed, and
 * while it is, the comments and the headings are the same question - where
 * in this document do I need to be.
 */
function renderComments(doc: Doc, box: HTMLElement): void {
  const comments = doc.comments ?? [];
  if (comments.length === 0) return;
  box.appendChild(el('div', 'rail-label cmt-label', comments.length + ' COMMENTS'));
  for (const c of comments) {
    const row = el('div', 'cmt-row');
    const who = el('div', 'cmt-who');
    who.append(
      el('span', 'cmt-badge', c.initials || String(c.id)),
      el('span', 'cmt-author', c.author || 'Unknown'),
      el('span', 'cmt-when', c.date ? c.date.slice(0, 10) : '')
    );
    row.append(who, el('div', 'cmt-text', c.text));
    row.addEventListener('mousedown', (ev) => ev.preventDefault());
    row.addEventListener('click', () => {
      const anchor = document.querySelector(
        `#doc [data-cmt-id="${CSS.escape(c.id)}"]`
      ) as HTMLElement | null;
      anchor?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    });
    box.appendChild(row);
  }
}

/**
 * Arrow keys move through a list, Enter opens.
 *
 * Roving focus rather than a selected index: the rows are real buttons in
 * the tab order, so moving focus IS moving the selection and there is no
 * second piece of state to keep in step with it.
 */
function listKeys(e: KeyboardEvent, selector: string, open: () => void): void {
  const rows = Array.from(
    (e.currentTarget as HTMLElement).parentElement?.querySelectorAll(selector) ?? []
  ) as HTMLElement[];
  const at = rows.indexOf(e.currentTarget as HTMLElement);
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault();
    const next = rows[at + (e.key === 'ArrowDown' ? 1 : -1)];
    next?.focus();
  } else if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    open();
  }
}

/** Headings, in document order, as a clickable outline. */
export function updateOutline(doc: Doc, activeId: string | null): void {
  const box = ui.outline;
  if (!box) return;
  box.textContent = '';

  const entries: { id: string; text: string; sub: boolean }[] = [];
  for (const b of doc.blocks) {
    if (isTable(b)) continue;
    if (b.styleId !== 'SectionHeading' && b.styleId !== 'JobTitle') continue;
    const text = plainText(b.html).trim();
    if (!text) continue;
    entries.push({ id: b.id, text, sub: b.styleId === 'JobTitle' });
  }

  if (entries.length === 0) {
    const empty = el('div', 'outline-empty');
    empty.innerHTML =
      'NO HEADINGS YET<br>HEADINGS APPEAR HERE<br>AS YOU WRITE';
    box.appendChild(empty);
    renderComments(doc, box);
    return;
  }

  for (const e of entries) {
    const row = el('div', 'outline-row' + (e.sub ? ' sub' : ''));
    if (e.id === activeId) row.classList.add('on');
    const tick = el('span', 'outline-tick');
    row.append(tick, el('span', 'outline-text', e.text));
    row.tabIndex = 0;
    row.setAttribute('role', 'option');
    row.addEventListener('mousedown', (ev) => ev.preventDefault());
    row.addEventListener('click', () => hooks?.onOutlineClick(e.id));
    row.addEventListener('keydown', (ev) =>
      listKeys(ev, '.outline-row', () => hooks?.onOutlineClick(e.id))
    );
    box.appendChild(row);
  }
  renderComments(doc, box);
}

/* ------------------------------------------------------------------ *
 * Gutter: page numbers and page-break rules
 * ------------------------------------------------------------------ */

export function ensureGutter(): HTMLElement {
  if (ui.gutter && ui.gutter.isConnected) return ui.gutter;
  const g = el('div', 'gutter');
  g.id = 'gutter';
  canvasEl().appendChild(g);
  ui.gutter = g;
  return g;
}

/**
 * Page numbers sit beside each page with a short leader line, and a labelled
 * rule marks each break. Drawn as an overlay rather than inside the pages, so
 * nothing here can end up on the paper or in the printed output.
 */
/**
 * What the gutter is drawn against.
 *
 * Every keystroke used to rebuild the whole gutter, one getBoundingClientRect
 * per page. On a document of eighty pages that made a keystroke cost 15ms -
 * O(document) work on the one path that has to be O(page). The positions
 * only move when pagination does, so this signature decides whether to
 * rebuild or merely to move the highlight.
 */
let gutterKey = '';

function paginationKey(list: HTMLElement[]): string {
  if (list.length === 0) return '0';
  // Two rects rather than one per page: the count and the extent together
  // change whenever any page has moved.
  const first = list[0].getBoundingClientRect();
  const last = list[list.length - 1].getBoundingClientRect();
  return list.length + ':' + Math.round(first.top) + ':' + Math.round(last.bottom);
}

/** Move the highlight without touching anything else. */
function markCurrent(root: HTMLElement, selector: string, currentPage: number): void {
  const items = Array.from(root.querySelectorAll(selector)) as HTMLElement[];
  items.forEach((n) => {
    const i = Number(n.dataset.page);
    n.classList.toggle('on', i === currentPage);
  });
}

export function updateGutter(currentPage: number): void {
  const g = ensureGutter();
  const canvas = canvasEl();
  const list = pages();

  const key = paginationKey(list) + '|' + canvas.scrollTop;
  if (key === gutterKey) {
    markCurrent(g, '.gutter-num, .gutter-lead', currentPage);
    return;
  }
  gutterKey = key;

  g.textContent = '';
  const cRect = canvas.getBoundingClientRect();

  list.forEach((page, i) => {
    const r = page.getBoundingClientRect();
    const top = r.top - cRect.top + canvas.scrollTop;
    const left = r.left - cRect.left;

    const num = el('div', 'gutter-num', String(i + 1).padStart(2, '0'));
    num.dataset.page = String(i);
    if (i === currentPage) num.classList.add('on');
    num.style.top = top - 1 + 'px';
    num.style.left = left - 76 + 'px';
    g.appendChild(num);

    const lead = el('div', 'gutter-lead');
    lead.dataset.page = String(i);
    if (i === currentPage) lead.classList.add('on');
    lead.style.top = top + 6 + 'px';
    lead.style.left = left - 26 + 'px';
    g.appendChild(lead);

    if (i < list.length - 1) {
      // A break where the section changes is a different kind of break, and
      // saying so is the only way the reader learns why the margins moved.
      const next = list[i + 1] as HTMLElement;
      const crosses = (next.dataset.section ?? '0') !== (page.dataset.section ?? '0');
      const rule = el('div', 'gutter-break' + (crosses ? ' section' : ''));
      rule.style.top = top + r.height + 18 + 'px';
      rule.appendChild(el('span', 'gutter-break-line'));
      rule.appendChild(
        el('span', 'gutter-break-label', crosses ? 'SECTION BREAK' : 'PAGE BREAK')
      );
      rule.appendChild(el('span', 'gutter-break-line'));
      g.appendChild(rule);
    }
  });
}

/* ------------------------------------------------------------------ *
 * Spine: the document map
 * ------------------------------------------------------------------ */

/** Fractions down the document at which to tick the map, for find matches. */
let spineMarks: number[] = [];
let spineKey = '';

export function setSpineMarks(marks: number[]): void {
  spineMarks = marks;
  spineKey = ''; // the ticks are part of what is drawn
  updateSpine(currentPageIndex());
}

export function updateSpine(currentPage: number): void {
  let s = ui.spine;
  if (!s || !s.isConnected) {
    s = el('div', 'spine');
    s.id = 'spine';
    canvasEl().appendChild(s);
    ui.spine = s;
  }
  const list = pages();
  if (list.length === 0) {
    s.textContent = '';
    return;
  }

  const doc = docEl();
  // The bars only move when the page count or the marks do; the viewport
  // rectangle moves on every scroll and is cheap to reposition alone.
  const key = list.length + ':' + spineMarks.length + ':' + Math.round(s.clientHeight);
  if (key === spineKey) {
    markCurrent(s, '.spine-bar', currentPage);
    positionSpineView(s, doc, list.length);
    return;
  }
  spineKey = key;
  s.textContent = '';

  const gap = 8;
  const avail = s.clientHeight - 24;
  const barH = Math.max(6, (avail - gap * (list.length - 1)) / list.length);

  list.forEach((_page, i) => {
    const bar = el('div', 'spine-bar');
    bar.dataset.page = String(i);
    if (i === currentPage) bar.classList.add('on');
    bar.style.top = 12 + i * (barH + gap) + 'px';
    bar.style.height = barH + 'px';
    bar.title = `Page ${i + 1}`;
    bar.addEventListener('mousedown', (e) => e.preventDefault());
    bar.addEventListener('click', () => {
      const target = pages()[i];
      if (target) target.scrollIntoView({ block: 'start', behavior: 'smooth' });
    });
    s!.appendChild(bar);
  });

  // Find matches, ticked against the map.
  if (spineMarks.length > 0) {
    const trackTop = 12;
    const trackH = list.length * barH + (list.length - 1) * gap;
    for (const f of spineMarks) {
      const tick = el('div', 'spine-mark');
      tick.style.top = trackTop + f * trackH + 'px';
      s.appendChild(tick);
    }
  }

  positionSpineView(s, doc, list.length);
}

/**
 * The slice of the document on screen.
 *
 * Kept as one element that is moved rather than rebuilt: it is the only part
 * of the map that changes on a scroll, and scrolling is the one thing that
 * happens more often than typing.
 */
function positionSpineView(s: HTMLElement, doc: HTMLElement, pageCount: number): void {
  const total = doc.scrollHeight;
  if (total <= 0) return;
  let view = s.querySelector(':scope > .spine-view') as HTMLElement | null;
  if (!view) {
    view = el('div', 'spine-view');
    s.appendChild(view);
  }
  const gap = 8;
  const avail = s.clientHeight - 24;
  const barH = Math.max(6, (avail - gap * (pageCount - 1)) / pageCount);
  const trackTop = 12;
  const trackH = pageCount * barH + (pageCount - 1) * gap;
  view.style.top = trackTop + (doc.scrollTop / total) * trackH + 'px';
  view.style.height = Math.max(14, (doc.clientHeight / total) * trackH) + 'px';
}

/* ------------------------------------------------------------------ *
 * Status readout
 * ------------------------------------------------------------------ */

export interface Readout {
  page: number;
  pages: number;
  line: number;
  col: number;
  words: number;
  /** "Letter · 0.5 in margins", for the right-hand summary. */
  setup: string;
  /** "Saved", "Saving…", "Unsaved changes". */
  saved: string;
  savedFailed?: boolean;
}

/** Opens the inspector at Page setup; set by main.ts. */
let onPageSetup: (() => void) | null = null;
let onSettings: (() => void) | null = null;

export function setStatusHooks(pageSetup: () => void, settings: () => void): void {
  onPageSetup = pageSetup;
  onSettings = settings;
}

/**
 * The status bar: one line, the width of the window.
 *
 * The only place a page count or a word count appears. They used to be here
 * AND in the rail's metadata line AND on every row of the file list, three
 * copies computed three ways, which is three chances to disagree.
 */
export function updateReadout(r: Readout): void {
  let n = ui.readout;
  if (!n || !n.isConnected) {
    n = document.getElementById('statusbar');
    if (!n) return;
    ui.readout = n;
  }
  n.textContent = '';

  const strong = (v: string | number) => `<b>${v}</b>`;
  const left = el('div', 'status-left');
  left.innerHTML =
    `Page ${strong(r.page)} of ${strong(r.pages)}` +
    ` · Ln ${strong(r.line)}, Col ${strong(r.col)}` +
    ` · ${strong(r.words)} ${r.words === 1 ? 'word' : 'words'}`;
  n.appendChild(left);

  const right = el('div', 'status-right');
  const setup = document.createElement('button');
  setup.className = 'status-btn';
  setup.type = 'button';
  setup.textContent = r.setup;
  setup.title = 'Page setup';
  setup.setAttribute('aria-label', 'Page setup: ' + r.setup);
  setup.addEventListener('mousedown', (e) => e.preventDefault());
  setup.addEventListener('click', () => onPageSetup?.());
  right.appendChild(setup);

  const saved = el('span', 'status-saved', r.saved);
  if (r.savedFailed) saved.classList.add('err');
  right.appendChild(saved);

  // The gear lives here rather than in the nav panel, which is allowed to
  // hold navigation and nothing else.
  const gear = document.createElement('button');
  gear.className = 'status-btn status-gear';
  gear.type = 'button';
  gear.textContent = '⚙';
  gear.title = 'Settings';
  gear.setAttribute('aria-label', 'Settings');
  gear.addEventListener('mousedown', (e) => e.preventDefault());
  gear.addEventListener('click', () => onSettings?.());
  right.appendChild(gear);

  n.appendChild(right);
}

/** "Letter · 0.5 in margins", the compact summary in the status bar. */
export function pageSetupSummary(): string {
  const setup = geometryOf(pages()[currentPageIndex()]) ?? currentPageSetup();
  const letter = setup.width === 816 && setup.height === 1056;
  const legal = setup.width === 816 && setup.height === 1344;
  const a4 = Math.abs(setup.width - 794) < 3;
  const paper = letter ? 'Letter' : legal ? 'Legal' : a4 ? 'A4' : 'Custom';
  const inches = (setup.margins.top / 96).toFixed(2).replace(/0$/, '').replace(/\.$/, '');
  return `${paper} · ${inches} in margins`;
}

/** Which page index the caret is on, and the line/column within its block. */
export function caretReadout(
  words: number,
  saved = '',
  savedFailed = false
): Readout {
  const list = pages();
  const sel = window.getSelection();
  let page = 0;
  let line = 1;
  let col = 1;

  if (sel && sel.focusNode) {
    let node: Node | null = sel.focusNode;
    let blk: HTMLElement | null = null;
    while (node) {
      if (
        node.nodeType === Node.ELEMENT_NODE &&
        (node as HTMLElement).classList.contains('blk')
      ) {
        blk = node as HTMLElement;
        break;
      }
      node = node.parentNode;
    }
    if (blk) {
      const pageEl = blk.closest('.page') as HTMLElement | null;
      if (pageEl) page = Math.max(0, list.indexOf(pageEl));
      const r = document.createRange();
      r.setStart(blk, 0);
      try {
        r.setEnd(sel.focusNode, sel.focusOffset);
        const text = r.toString();
        const nl = text.lastIndexOf('\n');
        col = text.length - nl;
        line = text.split('\n').length;
      } catch {
        /* selection moved underneath us */
      }
    }
  }
  return {
    page: page + 1,
    pages: Math.max(1, list.length),
    line,
    col,
    words,
    setup: pageSetupSummary(),
    saved,
    savedFailed,
  };
}

/** The page the caret is on, for the gutter and spine highlights. */
export function currentPageIndex(): number {
  const sel = window.getSelection();
  if (!sel || !sel.focusNode) return 0;
  let node: Node | null = sel.focusNode;
  while (node) {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const p = (node as HTMLElement).closest?.('.page') as HTMLElement | null;
      if (p) return Math.max(0, pages().indexOf(p));
    }
    node = node.parentNode;
  }
  return 0;
}

export function activeHeadingId(doc: Doc): string | null {
  const sel = window.getSelection();
  if (!sel || !sel.focusNode) return null;
  let node: Node | null = sel.focusNode;
  let blk: HTMLElement | null = null;
  while (node) {
    if (
      node.nodeType === Node.ELEMENT_NODE &&
      (node as HTMLElement).classList.contains('blk')
    ) {
      blk = node as HTMLElement;
      break;
    }
    node = node.parentNode;
  }
  if (!blk) return null;
  const id = blk.dataset.continuesFrom || blk.dataset.blockId;
  if (!id) return null;
  // The heading this block sits under.
  let last: string | null = null;
  for (const b of doc.blocks) {
    if (isTable(b)) continue;
    if (b.styleId === 'SectionHeading' || b.styleId === 'JobTitle') last = b.id;
    if (b.id === id) return last;
  }
  return last;
}

export { styleOf };
