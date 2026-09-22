// styles.css is linked from index.html so the page geometry is applied
// before this module runs: the first measurement must not happen unstyled.
import type { BlockAlign, Doc, MarginKey, StyleId } from './model';
import {
  MARGINS,
  STYLE_IDS,
  emptyDoc,
  marginPreset,
  newBlock,
  newId,
  uniformMargins,
  sectionsOf,
} from './model';
import { DOC_FONT, STYLES, injectStyleSheet, setDocumentFont } from './styles';
import { blockEl, blocksIn, docEl, pages, readModel, renderAll } from './render';
import {
  clearHeightCache,
  currentPageSetup,
  ensureTrailingBlock,
  normalize,
  pageCount,
  paginate,
  paginateIfNeeded,
  setHeaderFooterSpace,
  setPageSetup,
  setSections,
} from './paginate';
import {
  bindShortcuts,
  clearBlockFormat,
  currentFormat,
  currentRunStyle,
  currentStyle,
  hasSelection,
  inlineState,
  setBlockFormat,
  setBlockStyle,
  setRunStyle,
  toggleInline,
} from './commands';
import { caretAtStart, getCaret, placeCaret } from './caret';
import type { HFTable } from './headers';
import {
  emptyHFTable,
  firstHFBlock,
  hasHeaderOrFooter,
  hfSpace,
  isEditingHF,
  measureHF,
  readHF,
  renderHF,
  sectionOfPage,
  setEditingHF,
} from './headers';
import { closeFind, isFindOpen, openFind, refreshFind, selectedText } from './findbar';
import { bindPaste } from './paste';
import { bindTables } from './tables';
import {
  canRedo,
  canUndo,
  noteTyping,
  redo,
  resetHistory,
  snapshot,
  undo,
} from './history';
import {
  download,
  duplicateDoc,
  exportJson,
  safeFileName,
  importJson,
  load,
  loadById,
  readIndex,
  removeDoc,
  save,
  storageBytes,
} from './persist';
import type { Vault } from './docx-package';
import { deleteOriginal, loadOriginal, saveOriginal } from './docx-package';
import { importDocx } from './docx-import';
import { closeMenu, iconButton, menuButton, textButton } from './ui';
import type { Command } from './commandbar';
import { closeCommands, isCommandsOpen, openCommands } from './commandbar';
import { bindFormatBar, hideFormatBar } from './formatbar';
import { closeExportSheet, isExportSheetOpen, openExportSheet } from './exportsheet';
import { toHtml, toMarkdown } from './export-text';
import { fromMarkdown } from './import-md';
import { exportLibrary, libraryFileName } from './export-library';
import { PALETTES, applyPalette, currentPaletteId } from './theme';
import {
  activeHeadingId,
  buildRail,
  caretReadout,
  currentPageIndex,
  railEl,
  updateGutter,
  updateOutline,
  setLibrary,
  setRailTab,
  flashRail,
  updateRail,
  updateReadout,
  updateSpine,
} from './shell';

let doc: Doc = emptyDoc();
type SaveState = 'saved' | 'pending' | 'failed';
let saveState: SaveState = 'saved';
/** The original .docx package, when this document came from one. */
let vault: Vault | null = null;

/* ------------------------------------------------------------------ *
 * Toolbar
 * ------------------------------------------------------------------ */

const ui = {
  style: null as ReturnType<typeof menuButton> | null,
  page: null as ReturnType<typeof menuButton> | null,
  para: null as ReturnType<typeof menuButton> | null,
  text: null as ReturnType<typeof menuButton> | null,
  hf: null as HTMLButtonElement | null,
  palette: null as ReturnType<typeof menuButton> | null,
  title: null as HTMLInputElement | null,
  bold: null as HTMLButtonElement | null,
  italic: null as HTMLButtonElement | null,
  underline: null as HTMLButtonElement | null,
  undo: null as HTMLButtonElement | null,
  redo: null as HTMLButtonElement | null,
  counts: null as HTMLSpanElement | null,
};

const IS_MAC = navigator.platform.toLowerCase().includes('mac');
const MOD = IS_MAC ? '⌘' : 'Ctrl+';

/**
 * The direct-formatting presets the menus offer.
 *
 * Deliberately a short list. The point is not to reproduce Word's paragraph
 * dialog, it is to be able to say what an imported document already says,
 * and to centre a title without leaving the keyboard.
 */
const ALIGNMENTS: { id: BlockAlign; label: string; key: string }[] = [
  { id: 'left', label: 'Left', key: 'l' },
  { id: 'center', label: 'Centre', key: 'e' },
  { id: 'right', label: 'Right', key: 'r' },
  { id: 'justify', label: 'Justified', key: 'j' },
];

/** Points. The sizes a document actually uses, not a continuous spinner. */
const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 24, 36];

/**
 * A short palette of ink colours.
 *
 * Named rather than a picker: a document with six arbitrary hex colours in
 * it is a document nobody can restyle later, and every one of these reads on
 * white paper in print.
 */
const INK_COLORS: { label: string; value: string | null }[] = [
  { label: 'Default', value: null },
  { label: 'Black', value: '000000' },
  { label: 'Grey', value: '595959' },
  { label: 'Red', value: 'C00000' },
  { label: 'Orange', value: 'B45309' },
  { label: 'Green', value: '2E6B33' },
  { label: 'Blue', value: '1F4E79' },
  { label: 'Purple', value: '5B2D8E' },
];

const LINE_SPACINGS = [
  { label: 'Single', value: 1 },
  { label: '1.15', value: 1.15 },
  { label: 'One and a half', value: 1.5 },
  { label: 'Double', value: 2 },
];

/** Points, the unit the rest of the style table is in. */
const SPACES = [
  { label: 'None', value: 0 },
  { label: '6 pt', value: 6 },
  { label: '12 pt', value: 12 },
];
const SHIFT = IS_MAC ? '⇧' : 'Shift+';

function inlineGroup(...kids: HTMLElement[]): HTMLElement {
  const d = document.createElement('div');
  d.className = 'tb-inline';
  d.append(...kids);
  return d;
}

/**
 * Controls live in the rail, never over the paper. The order is the order
 * they are reached for: what the file is, then undo, then what the text is.
 */
function buildToolbar(host: HTMLElement): void {
  host.textContent = '';

  host.appendChild(
    textButton(
      'Commands',
      `Every command, by name (${MOD}K)`,
      () => toggleCommands(),
      'tb-cmd',
      MOD + 'K'
    )
  );
  host.appendChild(
    textButton('Find', `Find and replace (${MOD}F)`, () => showFind(), '', MOD + 'F')
  );
  ui.hf = textButton(
    'Header & footer',
    'Edit the header and footer (Esc to return)',
    () => toggleHF(),
    'tb-hf'
  );
  host.appendChild(ui.hf);

  const file = menuButton('File', 'Documents, open, save and export', () => [
    { label: 'New document', hint: MOD + 'N', onSelect: () => newDocument() },
    { label: 'Browse documents', hint: MOD + 'O', onSelect: () => setRailTab('files') },
    { separator: true },
    { label: 'Open Word document…', onSelect: () => void pickDocx() },
    { separator: true },
    { label: 'Export…', hint: MOD + SHIFT + 'E', onSelect: () => showExportSheet() },
    { label: 'Save as Word (.docx)', onSelect: () => void exportWord() },
    { label: 'Print / save as PDF', hint: MOD + 'P', onSelect: () => void printDocument() },
    { separator: true },
    { label: 'Import Markdown…', onSelect: () => void pickText('md') },
    { label: 'Import JSON…', onSelect: () => void pickText('json') },
  ]);
  host.appendChild(file.el);

  ui.undo = iconButton('undo', `Undo (${MOD}Z)`, () => {
    undo();
    afterChange();
  });
  ui.redo = iconButton('redo', `Redo (${MOD}${SHIFT}Z)`, () => {
    redo();
    afterChange();
  });
  host.appendChild(inlineGroup(ui.undo, ui.redo));

  // Each entry previews itself in its own style, so the list shows what the
  // styles look like rather than only what they are called.
  ui.style = menuButton(
    'Body',
    'Paragraph style',
    () => {
      const current = currentStyle();
      return STYLE_IDS.map((id) => {
        const d = STYLES[id];
        return {
          label: d.label,
          checked: id === current,
          preview: {
            fontFamily: DOC_FONT,
            fontSize: Math.min(19, Math.max(12, d.size * 1.2)) + 'px',
            fontWeight: d.bold ? '700' : '400',
            textTransform: d.uppercase ? 'uppercase' : 'none',
            letterSpacing: d.letterSpacing ? d.letterSpacing + 'px' : 'normal',
          },
          onSelect: () => setBlockStyle(id),
        };
      });
    },
    'tb-style'
  );
  host.appendChild(ui.style.el);

  // Character formatting. Disabled with no selection, because every entry
  // applies to a range: there is nothing to size or colour without one.
  ui.text = menuButton('Text', 'Size, font, colour and decoration', () => {
    const r = currentRunStyle();
    const none = !hasSelection();
    const row = (label: string, on: boolean, run: () => void, extra = {}) => ({
      label,
      checked: on,
      disabled: none,
      onSelect: run,
      ...extra,
    });
    return [
      { heading: none ? 'Select text first' : 'Size' },
      ...(none
        ? []
        : FONT_SIZES.map((v) =>
            row(v + ' pt', r.size === v, () => setRunStyle({ size: v }))
          )),
      ...(none ? [] : [{ separator: true as const }, { heading: 'Colour' }]),
      ...(none
        ? []
        : INK_COLORS.map((c) =>
            row(
              c.label,
              (r.color ?? null) === c.value,
              () => setRunStyle({ color: c.value }),
              c.value ? { swatch: '#' + c.value, swatchBg: '#' + c.value } : {}
            )
          )),
      ...(none ? [] : [{ separator: true as const }]),
      ...(none
        ? []
        : [
            row('Strikethrough', !!r.strike, () => setRunStyle({ strike: !r.strike })),
            row('Superscript', r.vert === 'super', () =>
              setRunStyle({ vert: r.vert === 'super' ? null : 'super' })
            ),
            row('Subscript', r.vert === 'sub', () =>
              setRunStyle({ vert: r.vert === 'sub' ? null : 'sub' })
            ),
            { separator: true as const },
            row('Clear character formatting', false, () =>
              setRunStyle({ size: null, font: null, color: null, strike: null, vert: null })
            ),
          ]),
    ];
  });
  host.appendChild(ui.text.el);

  // Direct formatting: what the document asks for over and above its style.
  ui.para = menuButton('Paragraph', 'Alignment, indents and spacing', () => {
    const f = currentFormat();
    const align = f.align ?? 'left';
    const row = (label: string, on: boolean, run: () => void) => ({
      label,
      checked: on,
      onSelect: run,
    });
    return [
      { heading: 'Alignment' },
      ...ALIGNMENTS.map((a) =>
        row(a.label, align === a.id, () => setBlockFormat({ align: a.id }))
      ),
      { separator: true },
      { heading: 'Line spacing' },
      ...LINE_SPACINGS.map((l) =>
        row(l.label, (f.lineRule ?? 'auto') === 'auto' && f.lineHeight === l.value, () =>
          setBlockFormat({ lineHeight: l.value, lineRule: 'auto' })
        )
      ),
      { separator: true },
      { heading: 'Space before' },
      ...SPACES.map((sp) =>
        row(sp.label, f.spaceBefore === sp.value, () =>
          setBlockFormat({ spaceBefore: sp.value })
        )
      ),
      { heading: 'Space after' },
      ...SPACES.map((sp) =>
        row(sp.label, f.spaceAfter === sp.value, () =>
          setBlockFormat({ spaceAfter: sp.value })
        )
      ),
      { separator: true },
      { label: 'Clear direct formatting', onSelect: () => clearBlockFormat() },
    ];
  });
  host.appendChild(ui.para.el);

  ui.bold = textButton('B', `Bold (${MOD}B)`, () => toggleInline('bold'), 'tb-b');
  ui.italic = textButton('I', `Italic (${MOD}I)`, () => toggleInline('italic'), 'tb-i');
  ui.underline = textButton(
    'U',
    `Underline (${MOD}U)`,
    () => toggleInline('underline'),
    'tb-u'
  );
  host.appendChild(inlineGroup(ui.bold, ui.italic, ui.underline));


  ui.page = menuButton('Margins', 'Page margins', () => {
    const preset = marginPreset(doc.page);
    return [
      { heading: 'Margins' },
      ...(Object.keys(MARGINS) as MarginKey[]).map((key) => ({
        label: key === 'narrow' ? 'Narrow — 0.5 in' : 'Normal — 1 in',
        checked: preset === key,
        onSelect: () => setMarginPreset(key),
      })),
    ];
  });
  host.appendChild(ui.page.el);

  // The palettes, each swatched in its own accent.
  ui.palette = menuButton('Palette', 'Workspace palette', () => {
    const now = currentPaletteId();
    const rows = [];
    let lastDark: boolean | null = null;
    for (const p of PALETTES) {
      if (p.dark !== lastDark) {
        rows.push({ heading: p.dark ? 'Dark' : 'Light' });
        lastDark = p.dark;
      }
      rows.push({
        label: p.label,
        note: p.note,
        swatch: p.acc,
        swatchBg: p.bg,
        checked: p.id === now,
        onSelect: () => setPalette(p.id),
      });
    }
    return rows;
  });
  host.appendChild(ui.palette.el);

  host.appendChild(
    textButton(
      'Print',
      `Print or save as PDF (${MOD}P)`,
      () => void printDocument(),
      '',
      MOD + 'P'
    )
  );
}

/* ------------------------------------------------------------------ *
 * The library
 *
 * A document is only ever edited in the DOM, so switching away has to flush
 * whatever is pending before the model is replaced - otherwise the last few
 * seconds of typing are lost to the swap rather than to a crash.
 * ------------------------------------------------------------------ */

function flushSave(): void {
  clearTimeout(saveTimer);
  saveTimer = 0;
  syncModel();
  save(doc, docMeta());
}

function docMeta(): { words: number; pages: number } {
  return {
    words: (docEl().textContent ?? '').match(/\S+/g)?.length ?? 0,
    pages: pageCount(),
  };
}

function refreshLibrary(): void {
  setLibrary(readIndex(), doc.id, storageBytes());
}

function switchTo(id: string): void {
  if (id === doc.id) return;
  flushSave();
  const next = loadById(id);
  if (!next) return;
  vault = null;
  openDoc(next);
  void restoreVault(next);
  refreshLibrary();
}

function newDocument(): void {
  flushSave();
  vault = null;
  const blank = emptyDoc();
  blank.title = 'Untitled';
  openDoc(blank);
  save(blank, { words: 0, pages: 1 });
  refreshLibrary();
  ui.title?.focus();
  ui.title?.select();
}

function duplicate(id: string): void {
  if (id === doc.id) flushSave();
  const source = loadById(id);
  if (!source) return;
  const copy = duplicateDoc(source);
  save(copy, { words: 0, pages: 1 });
  // The copy has fresh block ids, so it cannot share the original's vault:
  // it exports as a new document rather than back into someone else's file.
  switchTo(copy.id);
}

function remove(id: string): void {
  const entry = readIndex().find((e) => e.id === id);
  const name = entry?.title || 'this document';
  if (!confirm(`Delete "${name}"? This cannot be undone.`)) return;

  removeDoc(id);
  void deleteOriginal(id);

  if (id === doc.id) {
    // Land somewhere real rather than on a document that no longer exists.
    const next = readIndex()[0];
    const replacement = next ? loadById(next.id) : null;
    vault = null;
    if (replacement) {
      openDoc(replacement);
      void restoreVault(replacement);
    } else {
      const blank = emptyDoc();
      openDoc(blank);
      save(blank, { words: 0, pages: 1 });
    }
  }
  refreshLibrary();
}

/** The document name, editable in place at the top of the rail. */
function mountTitle(): void {
  const host = document.querySelector('.rail-title');
  if (!host) return;
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'rail-title';
  input.value = doc.title;
  input.setAttribute('aria-label', 'Document name');
  input.addEventListener('input', () => {
    doc.title = input.value || 'Untitled';
    scheduleSave();
  });
  host.replaceWith(input);
  ui.title = input;
}

/**
 * Header editing is a separate context: the body dims, the header slots
 * become editable, and Escape hands the caret back to the body.
 */
function toggleHF(): void {
  if (isEditingHF()) {
    exitHF();
    return;
  }
  if (!hasHeaderOrFooter(doc)) {
    // Nothing to edit yet; give the document an empty header to type into.
    doc.headers = { default: [{ id: newId(), styleId: 'Body', html: '' }] };
    layout();
  }
  setEditingHF(true);
  ui.hf?.classList.add('on');
  const first = firstHFBlock();
  if (first) {
    docEl().focus();
    caretAtStart(first);
  }
  updateToolbar();
}

function exitHF(): void {
  setEditingHF(false);
  ui.hf?.classList.remove('on');
  const first = blocksIn(docEl())[0];
  if (first) {
    docEl().focus();
    caretAtStart(first);
  }
  scheduleSave();
  updateToolbar();
}

function showFind(): void {
  openFind(
    {
      onReplaced: (n) => {
        if (n === 0) return;
        // One reflow for the whole replace, not one per hit.
        normalize();
        paginate();
        snapshot('structural');
        updateToolbar();
        scheduleSave();
      },
      onClose: () => {
        docEl().focus();
        updateToolbar();
      },
    },
    selectedText()
  );
}

function setPalette(id: string): void {
  const p = applyPalette(id);
  ui.palette?.setLabel(p.label);
  // Page shadows and rules changed, but nothing about the text did, so no
  // reflow is needed - only the chrome that depends on the accent.
  refreshChrome();
}

function setMarginPreset(key: MarginKey): void {
  // Only the margins change; an imported page size is left alone.
  const margins = uniformMargins(MARGINS[key]);
  doc.page = { ...doc.page, margins };
  // Every section, not only the first. Choosing a margin from the toolbar is
  // a statement about the document, and leaving a trailing imported section
  // on its old margins would make the change look like it half worked.
  if (doc.sections) {
    doc.sections = doc.sections.map((sct) => ({
      ...sct,
      page: { ...sct.page, margins },
    }));
  }
  setPageSetup(doc.page);
  installSections(doc);
  reflowNow();
  scheduleSave();
}

/* ------------------------------------------------------------------ *
 * State sync
 * ------------------------------------------------------------------ */

function syncModel(): void {
  doc.blocks = readModel(docEl());
}

function updateToolbar(): void {
  const st = currentStyle();
  ui.style?.setLabel(st ? STYLES[st].label : 'Mixed');

  const inline = inlineState();
  ui.bold?.classList.toggle('on', inline.bold);
  ui.italic?.classList.toggle('on', inline.italic);
  ui.underline?.classList.toggle('on', inline.underline);

  if (ui.undo) ui.undo.disabled = !canUndo();
  if (ui.redo) ui.redo.disabled = !canRedo();

  ui.page?.setLabel(
    marginPreset(doc.page) === 'narrow'
      ? 'Narrow margins'
      : marginPreset(doc.page) === 'normal'
        ? 'Normal margins'
        : 'Custom margins'
  );

  updateCounts();
  refreshChrome();
}

/**
 * The rail, gutter, spine and readout all describe where the caret is, so
 * they are refreshed together whenever anything moves.
 */
function refreshChrome(): void {
  const page = currentPageIndex();
  updateGutter(page);
  updateSpine(page);
  updateReadout(caretReadout(lastWords < 0 ? 0 : lastWords));
  updateRail(
    doc,
    lastWords < 0 ? 0 : lastWords,
    saveState === 'failed'
      ? 'NOT SAVED — STORAGE FULL'
      : saveState === 'pending'
        ? 'SAVING'
        : 'SAVED',
    saveState === 'failed'
  );
  scheduleOutline();
}

let findTimer = 0;
/** The document moved, so the hit list is stale. */
function scheduleFindRefresh(): void {
  clearTimeout(findTimer);
  findTimer = window.setTimeout(() => refreshFind(), 300);
}

let outlineTimer = 0;
/** The outline reads the whole model, so it follows the text rather than
 *  every keystroke. */
function scheduleOutline(): void {
  clearTimeout(outlineTimer);
  outlineTimer = window.setTimeout(() => {
    syncModel();
    updateOutline(doc, activeHeadingId(doc));
  }, 350);
}

let countsTimer = 0;
let lastWords = -1;
/** The word count reads the whole document, so it is debounced. */
function updateCounts(): void {
  clearTimeout(countsTimer);
  countsTimer = window.setTimeout(() => {
    const words = (docEl().textContent ?? '').match(/\S+/g)?.length ?? 0;
    if (words !== lastWords) {
      lastWords = words;
      refreshChrome();
    }
  }, 400);
}

let saveTimer = 0;
function scheduleSave(): void {
  clearTimeout(saveTimer);
  if (saveState !== 'failed') {
    saveState = 'pending';
    updateToolbar();
  }
  saveTimer = window.setTimeout(() => {
    syncModel();
    save(doc, docMeta());
  }, 1000);
}

let reflowTimer = 0;
/** Debounced full reflow: paste, style change, margin change, font load, undo. */
function scheduleReflow(): void {
  clearTimeout(reflowTimer);
  reflowTimer = window.setTimeout(() => {
    layout();
    updateToolbar();
  }, 150);
}

function reflowNow(): void {
  clearTimeout(reflowTimer);
  layout();
  updateToolbar();
}

/* ------------------------------------------------------------------ *
 * Layout
 *
 * NUMPAGES is circular: the page count depends on pagination, and a header
 * carrying the count can change height when the number widens, which
 * changes pagination. Resolved by paginating with the count we have,
 * substituting, and repaginating once if a header actually changed height.
 * Two passes, then the second result stands.
 * ------------------------------------------------------------------ */

let hfHeights: HFTable = emptyHFTable();

function sameHeights(a: HFTable, b: HFTable): boolean {
  if (a.length !== b.length) return false;
  return a.every((sa, i) =>
    (['default', 'first', 'even'] as const).every(
      (v) =>
        Math.abs(sa.header[v] - b[i].header[v]) < 0.5 &&
        Math.abs(sa.footer[v] - b[i].footer[v]) < 0.5
    )
  );
}

/**
 * Tell the paginator about this document's sections.
 *
 * A continuous section is deliberately not registered as a page start: it
 * shares a page with what came before, and a page can only have one
 * geometry, so it folds into the section above it rather than being drawn
 * with margins that cannot apply.
 */
function installSections(d: Doc): void {
  const list = sectionsOf(d);
  if (list.length <= 1) {
    setSections([], new Map());
    return;
  }
  const starts = new Map<string, number>();
  list.forEach((sct, i) => {
    if (sct.startId && !sct.continuous) starts.set(sct.startId, i);
  });
  setSections(list.map((sct) => sct.page), starts);
}

function layout(opts?: { fromPage?: number }): void {
  hfHeights = measureHF(doc, Math.max(1, pageCount()));
  paginate(opts);
  renderHF(doc, hfHeights);

  const after = measureHF(doc, pageCount());
  if (!sameHeights(hfHeights, after)) {
    hfHeights = after;
    paginate();
    renderHF(doc, hfHeights);
  }
}



function afterChange(): void {
  updateToolbar();
  scheduleSave();
}

/* ------------------------------------------------------------------ *
 * Export
 * ------------------------------------------------------------------ */

async function printDocument(): Promise<void> {
  // The print dialog can open before a pending debounced reflow runs, which
  // produces a PDF that differs from the screen.
  await document.fonts.ready;
  clearHeightCache();
  reflowNow();
  window.print();
}

async function exportWord(): Promise<void> {
  syncModel();
  const { exportDocx } = await import('./export-docx');
  const blob = await exportDocx(doc, vault);
  download(safeFileName(doc.title) + '.docx', blob);
  if (vault) showWarnings(vault);
}

/* ------------------------------------------------------------------ *
 * .docx
 * ------------------------------------------------------------------ */

async function pickDocx(): Promise<void> {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept =
    '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    void openDocxFile(file);
  });
  input.click();
}

export async function openDocxFile(file: File): Promise<void> {
  const bytes = await file.arrayBuffer();
  try {
    const { doc: imported, vault: v } = await importDocx(bytes, file.name);
    vault = v;
    openDoc(imported);
    // Kept out of localStorage: too big, and only the bytes can rebuild the
    // vault that makes the round trip faithful.
    void saveOriginal(imported.id, bytes);
    save(doc, docMeta());
    refreshLibrary();
    showWarnings(v);
  } catch (err) {
    alert('Could not open that file: ' + (err as Error).message);
  }
}

/** A banner beats silent data loss: say what was preserved but not shown. */
function showWarnings(v: Vault): void {
  const host = document.getElementById('banner');
  if (!host) return;
  if (v.warnings.length === 0) {
    host.hidden = true;
    host.textContent = '';
    return;
  }
  host.hidden = false;
  host.textContent = '';
  const text = document.createElement('span');
  text.textContent =
    'Kept but not shown: ' +
    v.warnings
      .map((w) => `${w.detail.toLowerCase()} (${w.count})`)
      .join('; ') +
    '. All of it is written back on export.';
  host.appendChild(text);
  const close = document.createElement('button');
  close.textContent = 'Dismiss';
  close.addEventListener('click', () => {
    host.hidden = true;
  });
  host.appendChild(close);
}

/**
 * Write the current document out as text.
 *
 * The model is synced first: an export that quietly omits the sentence
 * being typed is worse than no export, and the save debounce means the
 * model is up to half a second behind the screen at any moment.
 */
function saveText(ext: string, mime: string, make: () => string): void {
  syncModel();
  download(safeFileName(doc.title) + ext, new Blob([make()], { type: mime + ';charset=utf-8' }));
  flashRail('EXPORTED ' + ext.slice(1));
}

/**
 * Read a Markdown or JSON file in as a new document.
 *
 * Neither carries a vault, so `vault` is cleared: a document that came from
 * text has no original package to preserve, and keeping the previous one
 * would write the new document's words into the old document's XML.
 */
async function pickText(kind: 'md' | 'json'): Promise<void> {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = kind === 'json' ? 'application/json,.json' : '.md,.markdown,text/markdown';
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result);
      try {
        doc =
          kind === 'json'
            ? importJson(text)
            : fromMarkdown(text, { title: file.name.replace(/[.](md|markdown)$/i, '') });
        vault = null;
        openDoc(doc);
        snapshot('structural');
        save(doc, docMeta());
        refreshLibrary();
        flashRail('IMPORTED ' + doc.blocks.length + ' BLOCKS');
      } catch (err) {
        alert('Could not import that file: ' + (err as Error).message);
      }
    };
    reader.readAsText(file);
  });
  input.click();
}

/**
 * The acceptance test for owning these documents: one command that writes
 * every one of them out in formats this program did not invent.
 */
async function exportWholeLibrary(): Promise<void> {
  syncModel();
  save(doc, docMeta()); // so the open document is exported as it stands
  flashRail('EXPORTING LIBRARY…');
  try {
    const { blob, documents, missing } = await exportLibrary();
    download(libraryFileName(), blob);
    flashRail(
      documents + ' DOCUMENT' + (documents === 1 ? '' : 'S') + ' EXPORTED' +
        (missing.length ? ' · ' + missing.length + ' MISSING' : ''),
      missing.length > 0
    );
  } catch (err) {
    flashRail('EXPORT FAILED', true);
    alert('Could not export the library: ' + (err as Error).message);
  }
}

/**
 * The export sheet: every format, with what each one costs.
 *
 * The .docx line changes wording depending on whether this document came
 * from a package, because the guarantee is completely different. Promising
 * a byte-identical round trip for a document that has no original to go
 * back into would be a lie told at exactly the wrong moment.
 */
function showExportSheet(): void {
  closeMenu();
  syncModel();
  openExportSheet(
    [
      {
        label: 'Word document',
        tag: '.DOCX',
        note: vault
          ? 'Back into the original package. Every part you did not touch comes ' +
            'back byte for byte, pictures and all.'
          : 'A new Word document with the six named styles. This document has no ' +
            'original package to write back into.',
        run: () => void exportWord(),
      },
      {
        label: 'Print or save as PDF',
        tag: MOD + 'P',
        note:
          'Exactly what is on screen. Set margins to None and turn page headers ' +
          'off in the print dialog.',
        run: () => void printDocument(),
      },
      {
        label: 'Markdown',
        tag: '.MD',
        note:
          'Headings, lists, tables and emphasis as plain text, readable in fifty ' +
          'years with no software. Pictures are named, not carried.',
        run: () => saveText('.md', 'text/markdown', () => toMarkdown(doc)),
      },
      {
        label: 'HTML page',
        tag: '.HTML',
        note:
          'One self-contained file with the styles and page setup inlined. Opens ' +
          'in any browser; references nothing.',
        run: () => saveText('.html', 'text/html', () => toHtml(doc)),
      },
      {
        label: 'JSON',
        tag: '.JSON',
        note:
          'The document model verbatim. The only lossless copy, and the only one ' +
          'this program reads back.',
        run: () => saveText('.json', 'application/json', () => exportJson(doc)),
      },
      {
        label: 'The whole library',
        tag: '.ZIP',
        secondary: true,
        note:
          'Every document, as Markdown, JSON and HTML, in one archive. The answer ' +
          'to what happens if this program stops existing.',
        run: () => void exportWholeLibrary(),
      },
    ],
    doc.title || 'Untitled'
  );
}

/* ------------------------------------------------------------------ *
 * The command palette
 *
 * Built fresh each time it opens, so the ticks next to the current style
 * and the current palette are the current ones. Everything the toolbar can
 * do is here; the toolbar is the shortcut, not the other way round.
 * ------------------------------------------------------------------ */

function commands(): Command[] {
  const out: Command[] = [];
  const add = (
    group: string,
    label: string,
    run: () => void,
    extra: Partial<Command> = {}
  ): void => {
    out.push({ id: group + ':' + label, group, label, run, ...extra });
  };

  const style = currentStyle();
  for (const id of STYLE_IDS) {
    add('STYLE', STYLES[id].label, () => setBlockStyle(id), {
      checked: id === style,
      keywords: 'paragraph style ' + id,
    });
  }

  const inline = inlineState();
  add('FORMAT', 'Bold', () => toggleInline('bold'), {
    hint: MOD + 'B',
    checked: inline.bold,
  });
  add('FORMAT', 'Italic', () => toggleInline('italic'), {
    hint: MOD + 'I',
    checked: inline.italic,
  });
  add('FORMAT', 'Underline', () => toggleInline('underline'), {
    hint: MOD + 'U',
    checked: inline.underline,
  });

  const run = currentRunStyle();
  if (hasSelection()) {
    for (const v of FONT_SIZES) {
      add('TEXT', v + ' pt', () => setRunStyle({ size: v }), {
        checked: run.size === v,
        keywords: 'font size point text',
      });
    }
    for (const c of INK_COLORS) {
      add('TEXT', 'Colour — ' + c.label, () => setRunStyle({ color: c.value }), {
        checked: (run.color ?? null) === c.value,
        keywords: 'colour color ink text',
      });
    }
    add('TEXT', 'Strikethrough', () => setRunStyle({ strike: !run.strike }), {
      checked: !!run.strike,
    });
    add('TEXT', 'Superscript', () =>
      setRunStyle({ vert: run.vert === 'super' ? null : 'super' }), {
      checked: run.vert === 'super',
    });
    add('TEXT', 'Subscript', () =>
      setRunStyle({ vert: run.vert === 'sub' ? null : 'sub' }), {
      checked: run.vert === 'sub',
    });
    add('TEXT', 'Clear character formatting', () =>
      setRunStyle({ size: null, font: null, color: null, strike: null, vert: null }), {
      keywords: 'reset size colour font',
    });
  }

  const fmt = currentFormat();
  for (const a of ALIGNMENTS) {
    add('ALIGN', 'Align ' + a.label.toLowerCase(), () => setBlockFormat({ align: a.id }), {
      hint: MOD + a.key.toUpperCase(),
      checked: (fmt.align ?? 'left') === a.id,
      keywords: 'paragraph alignment justify centre center',
    });
  }
  for (const l of LINE_SPACINGS) {
    add('SPACING', 'Line spacing — ' + l.label, () =>
      setBlockFormat({ lineHeight: l.value, lineRule: 'auto' }), {
      checked: (fmt.lineRule ?? 'auto') === 'auto' && fmt.lineHeight === l.value,
      keywords: 'leading line height paragraph',
    });
  }
  for (const sp of SPACES) {
    add('SPACING', 'Space after — ' + sp.label, () => setBlockFormat({ spaceAfter: sp.value }), {
      checked: fmt.spaceAfter === sp.value,
      keywords: 'paragraph spacing below',
    });
    add('SPACING', 'Space before — ' + sp.label, () => setBlockFormat({ spaceBefore: sp.value }), {
      checked: fmt.spaceBefore === sp.value,
      keywords: 'paragraph spacing above',
    });
  }
  add('SPACING', 'Clear direct formatting', () => clearBlockFormat(), {
    keywords: 'reset alignment indent spacing to the style',
  });

  add('EDIT', 'Undo', () => {
    undo();
    afterChange();
  }, { hint: MOD + 'Z' });
  add('EDIT', 'Redo', () => {
    redo();
    afterChange();
  }, { hint: MOD + SHIFT + 'Z' });
  add('EDIT', 'Find and replace', () => showFind(), { hint: MOD + 'F' });
  add('EDIT', 'Edit header and footer', () => toggleHF(), {
    keywords: 'letterhead page number',
  });

  add('FILE', 'New document', () => newDocument(), { hint: MOD + 'N' });
  add('FILE', 'Browse documents', () => setRailTab('files'), {
    hint: MOD + 'O',
    keywords: 'library open recent',
  });
  add('FILE', 'Open Word document…', () => void pickDocx(), { keywords: 'docx import' });
  add('FILE', 'Print / save as PDF', () => void printDocument(), { hint: MOD + 'P' });

  add('EXPORT', 'Export…', () => showExportSheet(), {
    hint: MOD + SHIFT + 'E',
    keywords: 'save as formats sheet',
  });
  add('EXPORT', 'Save as Word (.docx)', () => void exportWord(), { keywords: 'docx' });
  add('EXPORT', 'Export Markdown', () =>
    saveText('.md', 'text/markdown', () => toMarkdown(doc))
  );
  add('EXPORT', 'Export HTML', () => saveText('.html', 'text/html', () => toHtml(doc)));
  add('EXPORT', 'Export JSON', () =>
    saveText('.json', 'application/json', () => exportJson(doc))
  );
  add('EXPORT', 'Export whole library…', () => void exportWholeLibrary(), {
    keywords: 'backup everything zip archive independence',
  });
  add('IMPORT', 'Import Markdown…', () => void pickText('md'));
  add('IMPORT', 'Import JSON…', () => void pickText('json'));

  const preset = marginPreset(doc.page);
  for (const key of Object.keys(MARGINS) as MarginKey[]) {
    add(
      'PAGE',
      key === 'narrow' ? 'Narrow margins — 0.5 in' : 'Normal margins — 1 in',
      () => setMarginPreset(key),
      { checked: preset === key, keywords: 'margin page setup' }
    );
  }

  const nowPalette = currentPaletteId();
  for (const p of PALETTES) {
    add('PALETTE', p.label + ' — ' + p.note, () => setPalette(p.id), {
      checked: p.id === nowPalette,
      keywords: (p.dark ? 'dark' : 'light') + ' theme colour color',
    });
  }

  return out;
}

function toggleCommands(): void {
  closeMenu();
  openCommands(commands());
}

/* ------------------------------------------------------------------ *
 * Boot
 * ------------------------------------------------------------------ */

function openDoc(d: Doc): void {
  if (isFindOpen()) closeFind();
  if (isCommandsOpen()) closeCommands();
  if (isExportSheetOpen()) closeExportSheet();
  hideFormatBar();
  doc = d;
  // Font before geometry: setPageSetup clears the height cache, and every
  // cached height was measured in whatever family was set at the time.
  setDocumentFont(d.defaultFont);
  setPageSetup(d.page);
  if (ui.title) ui.title.value = d.title;
  lastWords = -1;
  setHeaderFooterSpace((i) => hfSpace(doc, hfHeights, i, sectionOfPage(pages()[i])));
  installSections(doc);
  renderAll(doc, docEl());
  normalize();
  ensureTrailingBlock();
  layout();
  resetHistory();
  // Put the caret at the top so the document is ready to type into, and so
  // the toolbar has a paragraph to report the style of.
  const first = blocksIn(docEl())[0];
  if (first) {
    docEl().focus();
    caretAtStart(first);
  }
  updateToolbar();
}

/**
 * Rebuild the preservation vault after a reload by re-parsing the original
 * bytes. Imported block ids are derived from the body index, so the vault's
 * keys still line up with the edited blocks that came back from localStorage.
 */
async function restoreVault(d: Doc): Promise<void> {
  const bytes = await loadOriginal(d.id);
  if (!bytes) return;
  try {
    const { vault: v } = await importDocx(bytes, d.title);
    vault = v;
    showWarnings(v);
  } catch {
    /* the original is unreadable; export will regenerate instead */
  }
}

function sampleDoc(): Doc {
  const b = (styleId: StyleId, html: string) => newBlock(styleId, html);
  return {
    ...emptyDoc(),
    title: 'Resume',
    blocks: [
      b('Name', 'Your Name'),
      b('Contact', 'city, state &middot; you@example.com &middot; (555) 555-5555 &middot; linkedin.com/in/you'),
      b('SectionHeading', 'Experience'),
      b('JobTitle', 'Senior Engineer, Example Corp &mdash; 2022 to present'),
      b('Bullet', 'Led the rewrite of the billing pipeline, cutting invoice errors by 40%.'),
      b('Bullet', 'Mentored four engineers; two were promoted within a year.'),
      b('JobTitle', 'Engineer, Another Company &mdash; 2019 to 2022'),
      b('Bullet', 'Shipped the customer dashboard used by 12,000 accounts daily.'),
      b('SectionHeading', 'Education'),
      b('Body', 'B.S. Computer Science, State University, 2019'),
      b('SectionHeading', 'Skills'),
      b('Body', 'TypeScript, Go, Postgres, distributed systems, technical writing'),
      b('Body', ''),
    ],
  };
}

function boot(): void {
  injectStyleSheet();
  const root = docEl();

  applyPalette(currentPaletteId());
  buildRail(railEl(), {
    onOutlineClick: (blockId) => {
      const target = blockEl(blockId);
      if (!target) return;
      target.scrollIntoView({ block: 'center', behavior: 'smooth' });
      root.focus();
      caretAtStart(target);
      updateToolbar();
    },
    onOpenDoc: (id) => switchTo(id),
    onNewDoc: () => newDocument(),
    onDuplicateDoc: (id) => duplicate(id),
    onDeleteDoc: (id) => remove(id),
    onExportLibrary: () => void exportWholeLibrary(),
  });
  mountTitle();
  const bar = document.getElementById('rail-actions');
  if (!bar) throw new Error('#rail-actions missing');
  buildToolbar(bar);

  // Keep execCommand emitting <b>/<i>/<u> rather than styled spans.
  try {
    document.execCommand('styleWithCSS', false, 'false');
  } catch {
    /* not supported, not fatal */
  }

  const restored = load();
  openDoc(restored ?? sampleDoc());
  if (restored) void restoreVault(restored);
  else save(doc, docMeta());
  refreshLibrary();

  bindShortcuts(root);
  bindPaste(root);
  bindTables(root, {
    onChanged: () => {
      normalize();
      layout();
      snapshot('structural');
      updateToolbar();
      scheduleSave();
    },
  });

  root.addEventListener('input', (e) => {
    closeMenu();
    // An edit inside a header changes how much room the body has, so it has
    // to be read back and re-measured rather than treated as body input.
    const target = e.target as HTMLElement | null;
    if (isEditingHF() && (target?.closest?.('.page-header, .page-footer'))) {
      const caret = getCaret();
      if (readHF(doc)) {
        layout();
        if (caret) placeCaret(caret);
      }
      noteTyping();
      updateToolbar();
      scheduleSave();
      return;
    }
    normalize();
    if (ensureTrailingBlock()) paginate();
    paginateIfNeeded(); // synchronous, before paint
    noteTyping();
    updateToolbar();
    scheduleSave();
    if (isFindOpen()) scheduleFindRefresh();
  });

  document.addEventListener('selectionchange', () => {
    if (document.activeElement === root) updateToolbar();
  });

  // The gutter and the document map are positioned against the canvas, so
  // they follow the scroll rather than being redrawn by the editor.
  root.addEventListener('scroll', () => {
    const page = currentPageIndex();
    updateGutter(page);
    updateSpine(page);
  });
  window.addEventListener('resize', () => refreshChrome());

  window.addEventListener('keydown', (e) => {
    // The platform's own modifier, not either one. Ctrl+K and Ctrl+E are
    // emacs kill-line and end-of-line on a Mac, and a text field that
    // opened a dialog instead would be maddening.
    const mod = IS_MAC ? e.metaKey && !e.ctrlKey : e.ctrlKey;
    if (mod && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      toggleCommands();
      return;
    }
    if (mod && e.shiftKey && e.key.toLowerCase() === 'e') {
      e.preventDefault();
      showExportSheet();
      return;
    }
    // Word's alignment shortcuts, which is why the export sheet is on
    // Shift+E rather than E. Ctrl+R is reload in the browser and
    // right-align in every word processor; in a document window the
    // document wins, and F5 still reloads.
    if (mod && !e.shiftKey) {
      const a = ALIGNMENTS.find((x) => x.key === e.key.toLowerCase());
      if (a) {
        e.preventDefault();
        setBlockFormat({ align: a.id });
        return;
      }
    }
    if (isCommandsOpen()) return; // the palette owns the keyboard while open
    if (e.key === 'Escape' && isFindOpen()) {
      e.preventDefault();
      closeFind();
      return;
    }
    if (e.key === 'Escape' && isEditingHF()) {
      e.preventDefault();
      exitHF();
      return;
    }
    if (!(e.metaKey || e.ctrlKey)) return;
    const k = e.key.toLowerCase();
    if (k === 'n') {
      e.preventDefault();
      newDocument();
    } else if (k === 'o') {
      e.preventDefault();
      setRailTab('files');
    } else if (k === 'f') {
      e.preventDefault();
      showFind();
    }
  });

  bindFormatBar({
    bold: () => toggleInline('bold'),
    italic: () => toggleInline('italic'),
    underline: () => toggleInline('underline'),
    setStyle: (id) => setBlockStyle(id),
    currentStyle,
    inlineState,
    setAlign: (align) => setBlockFormat({ align }),
    currentFormat,
    // Header editing and the palette both move the selection somewhere the
    // bar has no business formatting.
    suppressed: () => isEditingHF() || isCommandsOpen(),
  });

  document.addEventListener('wp:changed', () => {
    updateToolbar();
    scheduleSave();
  });

  document.addEventListener('wp:saved', (e) => {
    saveState = (e as CustomEvent).detail?.ok ? 'saved' : 'failed';
    updateToolbar();
    refreshLibrary();
  });

  // A late font swap would silently invalidate every cached height.
  void document.fonts.ready.then(() => {
    clearHeightCache();
    scheduleReflow();
  });

  // Dropping a .docx onto the page opens it.
  root.addEventListener('dragover', (e) => {
    if (e.dataTransfer?.types.includes('Files')) e.preventDefault();
  });
  root.addEventListener('drop', (e) => {
    const file = e.dataTransfer?.files?.[0];
    if (!file || !/\.docx$/i.test(file.name)) return;
    e.preventDefault();
    void openDocxFile(file);
  });

  window.addEventListener('beforeprint', () => {
    clearHeightCache();
    paginate();
  });

  window.addEventListener('beforeunload', () => flushSave());

  root.focus();
  updateToolbar();
}

boot();

// Exposed for manual poking in the console during the acceptance tests.
Object.assign(window as unknown as Record<string, unknown>, {
  wp: {
    doc: () => doc,
    model: () => readModel(docEl()),
    pageCount,
    paginate,
    page: currentPageSetup,
    vault: () => vault,
    openDocx: openDocxFile,
    markdown: () => {
      syncModel();
      return toMarkdown(doc);
    },
    html: () => {
      syncModel();
      return toHtml(doc);
    },
    exportLibrary,
  },
});
