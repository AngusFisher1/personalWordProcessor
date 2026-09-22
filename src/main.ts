// styles.css is linked from index.html so the page geometry is applied
// before this module runs: the first measurement must not happen unstyled.
import type { Doc, MarginKey, StyleId } from './model';
import {
  MARGINS,
  emptyDoc,
  marginPreset,
  newBlock,
  newId,
  uniformMargins,
  contentWidth,
  sectionsOf,
} from './model';
import type { StyleOverrides } from './styles';
import {
  STYLES,
  injectStyleSheet,
  setDocumentFont,
  setStyleOverrides,
  shippedStyle,
} from './styles';
import { blockEl, blocksIn, docEl, pages, readModel, renderAll, resolveImages } from './render';
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
import type { MenuItem } from './ui';
import { closeMenu, openMenuAt } from './ui';
import { ask, isAskOpen } from './ask';
import { bindMobile, closeDrawer } from './mobile';
import { bindInspector, renderInspector as drawInspector } from './inspector';
import type { Version } from './versions';
import { forgetVersions, keepVersion, listVersions, whenLabel } from './versions';
import {
  imageHtml,
  insertTableAtCaret,
  linkAtCaret,
  linkSelection,
  pageBreakHere,
  readImageFile,
  safeLink,
  togglePageBreak,
  unlinkSelection,
} from './insert';
import {
  closeCommands,
  isCommandsOpen,
  openCommands,
  setCommandsHooks,
} from './commandbar';
import type { CommandActions } from './command-list';
import { buildCommands } from './command-list';
import {
  command,
  commandForEvent,
  formatShortcut,
  isEnabled,
  labelOf,
  recentIds,
  register,
  runCommand,
  setPlatform,
  setRecent,
} from './registry';
import { loadUiState, setUiState, uiState } from './uistate';
import { bindFormatBar, renderFormatBar, syncFormatBar } from './formatbar';
import { bindBubble, hideBubble } from './bubble';
import { closeExportSheet, isExportSheetOpen, openExportSheet } from './exportsheet';
import { toHtml, toMarkdown } from './export-text';
import { fromMarkdown } from './import-md';
import { exportLibrary, libraryFileName } from './export-library';
import { applyPalette, currentPaletteId } from './theme';
import {
  activeHeadingId,
  buildRail,
  caretReadout,
  currentPageIndex,
  railEl,
  updateGutter,
  updateOutline,
  setLibrary,
  flashRail,
  focusRailTab,
  setStatusHooks,
  startRename,
  updateRail,
  updateReadout,
  updateSpine,
} from './shell';

let doc: Doc = emptyDoc();
let vault: Vault | null = null;
type SaveState = 'saved' | 'pending' | 'failed';
let saveState: SaveState = 'saved';

const IS_MAC = navigator.platform.toLowerCase().includes('mac');

/**
 * The handful of elements that have to be reached for by name.
 *
 * It used to hold sixteen: every control in the rail. The rail's controls
 * are gone, so what is left is the title, which the rename replaces in
 * place, and the format bar, which redraws itself from the registry.
 */
const ui = {
  title: null as HTMLElement | null,
};


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
  // History is a luxury on top of the save, never a condition of it: this
  // is deliberately not awaited and its failures are its own. The list is
  // refreshed behind it so the menu is populated before it is opened -
  // a menu that builds its own rows cannot wait for a promise.
  void keepVersion(doc).then(refreshVersions);
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
  // A new document opens with its name selected, because naming it is the
  // first thing you do.
  startRename();
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
  void forgetVersions(id);
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
/**
 * The File menu, opened from the document title.
 *
 * Everything in it is also a command, so this is a shortcut to the palette
 * rather than a second place actions are defined: each row runs a
 * registered id and nothing else.
 */
function openFileMenu(anchor: HTMLElement): void {
  const item = (id: string, extra: Partial<MenuItem> = {}): MenuItem => {
    const c = command(id);
    return {
      label: c ? labelOf(c) : id,
      hint: formatShortcut(c?.shortcut),
      disabled: c ? !isEnabled(c) : true,
      onSelect: () => void runCommand(id),
      ...extra,
    };
  };
  openMenuAt(anchor, [
    item('file.new'),
    item('view.files'),
    item('file.rename'),
    item('file.duplicate'),
    { separator: true },
    item('document.versions'),
    { separator: true },
    item('file.openWord'),
    item('file.importMarkdown'),
    item('file.importJson'),
    { separator: true },
    item('file.export'),
    item('file.exportLibrary'),
    item('file.print'),
    { separator: true },
    item('file.delete'),
  ]);
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
  const first = firstHFBlock();
  if (first) {
    docEl().focus();
    caretAtStart(first);
  }
  updateToolbar();
}

function exitHF(): void {
  setEditingHF(false);
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
  applyPalette(id);
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
  if (uiState().inspectorOpen) renderInspector();
}

/* ------------------------------------------------------------------ *
 * State sync
 * ------------------------------------------------------------------ */

function syncModel(): void {
  doc.blocks = readModel(docEl());
}

/**
 * Everything that reflects where the caret is.
 *
 * The rail's controls used to be refreshed here by name, one line each.
 * The format bar redraws itself from the registry instead, so this is down
 * to the counts and the chrome.
 */
function updateToolbar(): void {
  syncFormatBar();
  updateCounts();
  refreshChrome();
}

/**
 * The rail, gutter, spine and readout all describe where the caret is, so
 * they are refreshed together whenever anything moves.
 */
/** One phrase, used by the nav panel and the status bar alike. */
function savedLabel(): string {
  return saveState === 'failed'
    ? 'Not saved — storage full'
    : saveState === 'pending'
      ? 'Saving…'
      : 'Saved';
}

function refreshChrome(): void {
  const page = currentPageIndex();
  const words = lastWords < 0 ? 0 : lastWords;
  updateGutter(page);
  updateSpine(page);
  updateReadout(caretReadout(words, savedLabel(), saveState === 'failed'));
  updateRail(doc, words, savedLabel(), saveState === 'failed');
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
    void keepVersion(doc).then(refreshVersions);
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
        tag: formatShortcut('Mod+P'),
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
 * History
 * ------------------------------------------------------------------ */

let versionCache: Version[] = [];

async function refreshVersions(): Promise<void> {
  versionCache = await listVersions(doc.id);
}

/**
 * Go back to an earlier save.
 *
 * Restoring is itself an edit rather than a rewind: the version you were on
 * is kept, so changing your mind about changing your mind costs nothing,
 * and one Undo puts it back.
 */
function restoreVersion(v: Version): void {
  try {
    const restored = importJson(v.json);
    // The id stays, or it becomes a different document in the library.
    restored.id = doc.id;
    doc = restored;
    openDoc(doc);
    snapshot('structural');
    save(doc, docMeta());
    refreshLibrary();
    flashRail('RESTORED ' + whenLabel(v.savedAt).toUpperCase());
  } catch (err) {
    flashRail('COULD NOT RESTORE', true);
    alert('That version could not be read: ' + (err as Error).message);
  }
}

/* ------------------------------------------------------------------ *
 * Inserting
 *
 * Tables, images and links all arrived as things an imported .docx might
 * contain: readable, editable, and impossible to create. That is the
 * difference between a viewer with an editing mode and a word processor.
 * ------------------------------------------------------------------ */


function afterInsert(message: string): void {
  closeDrawer();
  syncModel();
  reflowNow();
  snapshot('structural');
  scheduleSave();
  flashRail(message);
}

function doInsertTable(rows: number, cols: number): void {
  if (!insertTableAtCaret(rows, cols, contentWidth(doc.page))) {
    flashRail('PUT THE CARET IN THE DOCUMENT FIRST', true);
    return;
  }
  afterInsert(rows + ' × ' + cols + ' TABLE INSERTED');
}

function doPageBreak(): void {
  if (!togglePageBreak()) {
    flashRail('PUT THE CARET IN THE DOCUMENT FIRST', true);
    return;
  }
  afterInsert(pageBreakHere() ? 'PAGE BREAK ADDED' : 'PAGE BREAK REMOVED');
}

async function doLink(): Promise<void> {
  const existing = linkAtCaret();
  const answer = await ask({
    title: existing ? 'CHANGE LINK' : 'ADD LINK',
    placeholder: 'example.com  ·  https://…  ·  name@example.com',
    value: existing ?? '',
    validate: (v) =>
      v === '' || safeLink(v) ? null : 'Not a web address or an email address',
  });
  if (answer === null) return;
  if (answer === '') {
    if (unlinkSelection()) afterInsert('LINK REMOVED');
    return;
  }
  const href = safeLink(answer) as string;
  if (!linkSelection(href)) {
    flashRail('SELECT THE TEXT TO LINK FIRST', true);
    return;
  }
  afterInsert('LINKED');
}

async function doInsertImage(): Promise<void> {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/png,image/jpeg,image/gif,image/webp';
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    void (async () => {
      const img = await readImageFile(file, contentWidth(doc.page));
      if (!img) {
        flashRail('COULD NOT READ THAT IMAGE', true);
        return;
      }
      const at = blocksIn(docEl()).find((b) => b.contains(
        window.getSelection()?.focusNode ?? document.createElement('i')
      ));
      const target = at ?? blocksIn(docEl())[0];
      if (!target) return;
      target.insertAdjacentHTML('beforeend', imageHtml(img));
      resolveImages(target);
      afterInsert('IMAGE INSERTED');
    })();
  });
  input.click();
}

/* ------------------------------------------------------------------ *
 * The styles editor
 *
 * Six styles are the whole vocabulary this program writes in, and they were
 * hardcoded. A resume set in 10pt and a report set in 12pt are both right,
 * so the overrides live on the DOCUMENT - a global preference would be
 * wrong in one of them.
 * ------------------------------------------------------------------ */

const STYLE_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36];
const STYLE_LEADING = [1, 1.15, 1.3, 1.45, 1.6, 2];
const STYLE_SPACE = [0, 2, 4, 6, 8, 12, 18];

function styleOverride(id: StyleId, field: string, value: unknown): void {
  syncModel();
  const all: Record<string, Record<string, unknown>> = { ...(doc.styles ?? {}) };
  const over = { ...(all[id] ?? {}) };
  // A value equal to the shipped one is not an override, it is agreement.
  const shipped = shippedStyle(id) as unknown as Record<string, unknown>;
  if (JSON.stringify(shipped[field]) === JSON.stringify(value)) delete over[field];
  else over[field] = value;
  if (Object.keys(over).length === 0) delete all[id];
  else all[id] = over;
  doc.styles = Object.keys(all).length > 0 ? all : undefined;

  setStyleOverrides(doc.styles as StyleOverrides | undefined);
  // Every cached height was measured against the old table.
  clearHeightCache();
  reflowNow();
  snapshot('structural');
  scheduleSave();
  flashRail('STYLE UPDATED');
}

export function resetStyles(): void {
  syncModel();
  doc.styles = undefined;
  setStyleOverrides(undefined);
  clearHeightCache();
  reflowNow();
  snapshot('structural');
  scheduleSave();
  flashRail('STYLES RESET');
}

/** The menu for one named style. */
export function styleEditorRows(id: StyleId): MenuItem[] {
  const d = STYLES[id];
  const shipped = shippedStyle(id);
  const changed = JSON.stringify(d) !== JSON.stringify(shipped);
  const pad = d.padding;
  return [
    { heading: d.label + (changed ? ' — changed' : '') },
    ...STYLE_SIZES.map((v) => ({
      label: v + ' pt',
      checked: d.size === v,
      onSelect: () => styleOverride(id, 'size', v),
    })),
    { separator: true },
    { heading: 'Weight and case' },
    { label: 'Bold', checked: d.bold, onSelect: () => styleOverride(id, 'bold', !d.bold) },
    {
      label: 'Small capitals',
      checked: d.uppercase,
      onSelect: () => styleOverride(id, 'uppercase', !d.uppercase),
    },
    { label: 'Rule beneath', checked: d.rule, onSelect: () => styleOverride(id, 'rule', !d.rule) },
    { separator: true },
    { heading: 'Leading' },
    ...STYLE_LEADING.map((v) => ({
      label: v.toFixed(2).replace(/0$/, ''),
      checked: Math.abs(d.lineHeight - v) < 0.01,
      onSelect: () => styleOverride(id, 'lineHeight', v),
    })),
    { separator: true },
    { heading: 'Space above' },
    ...STYLE_SPACE.map((v) => ({
      label: v + ' pt',
      checked: pad[0] === v,
      onSelect: () => styleOverride(id, 'padding', [v, pad[1], pad[2], pad[3]]),
    })),
    { heading: 'Space below' },
    ...STYLE_SPACE.map((v) => ({
      label: v + ' pt',
      checked: pad[2] === v,
      onSelect: () => styleOverride(id, 'padding', [pad[0], pad[1], v, pad[3]]),
    })),
    { separator: true },
    {
      label: 'Keep with next paragraph',
      checked: d.keepWithNext,
      onSelect: () => styleOverride(id, 'keepWithNext', !d.keepWithNext),
    },
    {
      label: 'Never split across pages',
      checked: d.keepLines,
      onSelect: () => styleOverride(id, 'keepLines', !d.keepLines),
    },
    {
      label: 'Always start a new page',
      checked: d.pageBreakBefore,
      onSelect: () => styleOverride(id, 'pageBreakBefore', !d.pageBreakBefore),
    },
  ];
}

/* ------------------------------------------------------------------ *
 * The panels
 *
 * Both side panels are the same idea: a boolean in one place, written to
 * localStorage, read by one function that sets a class. Everything else -
 * the canvas re-centring, the spine getting out of the way - falls out of
 * CSS keyed on that class.
 * ------------------------------------------------------------------ */

function applyPanels(): void {
  const app = document.getElementById('app');
  const inspector = document.getElementById('inspector');
  const state = uiState();
  app?.classList.toggle('nav-collapsed', state.navCollapsed);
  app?.classList.toggle('inspector-open', state.inspectorOpen);
  if (inspector) inspector.hidden = !state.inspectorOpen;
  // The panels change how wide the canvas is, and the page is centred in it.
  scheduleReflow();
}

function toggleNav(): void {
  setUiState({ navCollapsed: !uiState().navCollapsed });
  applyPanels();
}

function toggleInspector(): void {
  setUiState({ inspectorOpen: !uiState().inspectorOpen });
  applyPanels();
  if (uiState().inspectorOpen) renderInspector();
}

function openInspector(section: 'page' | 'headerFooter' | 'print'): void {
  setUiState({ inspectorOpen: true, inspectorSection: section });
  applyPanels();
  renderInspector();
}

function renderInspector(): void {
  drawInspector();
}

/** Page setup changed: re-measure and redraw everything that reads it. */
function afterPageChange(): void {
  setPageSetup(doc.page);
  installSections(doc);
  clearHeightCache();
  reflowNow();
  snapshot('structural');
  scheduleSave();
  renderInspector();
}

/** Filled in by the settings commit. */
function openSettings(): void {
  flashRail('SETTINGS COMING');
}

/** The version list, as a menu off whatever opened it. */
function openVersionMenu(): void {
  const anchor = (ui.title ?? document.getElementById('rail')) as HTMLElement;
  if (versionCache.length === 0) {
    openMenuAt(anchor, [{ heading: 'No earlier saves yet' }]);
    return;
  }
  openMenuAt(anchor, [
    { heading: versionCache.length + ' saved versions' },
    ...versionCache.slice(0, 20).map((v) => ({
      label: whenLabel(v.savedAt),
      note: v.words + ' w',
      onSelect: () => restoreVersion(v),
    })),
    { separator: true },
    {
      label: 'Forget this history',
      onSelect: () => {
        void forgetVersions(doc.id).then(() => {
          versionCache = [];
          flashRail('HISTORY CLEARED');
        });
      },
    },
  ]);
}

/* ------------------------------------------------------------------ *
 * The command registry, wired to this document
 *
 * Every action in the program is defined once, in command-list.ts, and
 * handed the thunks below. The palette, the File menu, the format bar and
 * the keyboard all read that one list, so an action cannot be reachable
 * from one of them and missing from another.
 * ------------------------------------------------------------------ */

function commandActions(): CommandActions {
  return {
    setStyle: (id) => setBlockStyle(id),
    currentStyleId: () => currentStyle(),
    toggleInline: (which) => toggleInline(which),
    inlineState,
    setRunStyle: (change) => setRunStyle(change),
    currentRunStyle,
    hasSelection,
    setBlockFormat: (f) => setBlockFormat(f),
    currentFormat,
    clearBlockFormat: () => clearBlockFormat(),

    insertTable: (rows, cols) => doInsertTable(rows, cols),
    insertImage: () => void doInsertImage(),
    insertLink: () => void doLink(),
    linkAtCaret,
    togglePageBreak: () => doPageBreak(),
    pageBreakHere,

    undo: () => {
      undo();
      afterChange();
    },
    redo: () => {
      redo();
      afterChange();
    },
    find: () => showFind(),
    toggleHeaderFooter: () => toggleHF(),
    setMarginPreset: (key) => setMarginPreset(key),
    marginPreset: () => marginPreset(doc.page),

    newDocument: () => newDocument(),
    renameDocument: () => startRename(),
    duplicateDocument: () => duplicate(doc.id),
    deleteDocument: () => remove(doc.id),
    openWordFile: () => void pickDocx(),
    importMarkdown: () => void pickText('md'),
    importJson: () => void pickText('json'),
    print: () => void printDocument(),
    showExportSheet: () => showExportSheet(),
    exportWord: () => void exportWord(),
    exportMarkdown: () => saveText('.md', 'text/markdown', () => toMarkdown(doc)),
    exportHtml: () => saveText('.html', 'text/html', () => toHtml(doc)),
    exportJson: () => saveText('.json', 'application/json', () => exportJson(doc)),
    exportLibrary: () => void exportWholeLibrary(),
    versionHistory: () => openVersionMenu(),

    toggleNav: () => toggleNav(),
    navCollapsed: () => uiState().navCollapsed,
    focusOutline: () => focusRailTab('outline'),
    focusFiles: () => focusRailTab('files'),
    toggleInspector: () => toggleInspector(),
    inspectorOpen: () => uiState().inspectorOpen,
    openPageSetup: () => openInspector('page'),

    openPalette: () => toggleCommands(),
    openSettings: () => openSettings(),
    setPalette: (id) => setPalette(id),
    currentPaletteId: () => currentPaletteId(),
  };
}

function toggleCommands(): void {
  closeMenu();
  openCommands();
}

/* ------------------------------------------------------------------ *
 * Boot
 * ------------------------------------------------------------------ */

function openDoc(d: Doc): void {
  if (isFindOpen()) closeFind();
  if (isCommandsOpen()) closeCommands();
  if (isExportSheetOpen()) closeExportSheet();
  hideBubble();
  doc = d;
  // Font before geometry: setPageSetup clears the height cache, and every
  // cached height was measured in whatever family was set at the time.
  setDocumentFont(d.defaultFont);
  setStyleOverrides(d.styles as StyleOverrides | undefined);
  setPageSetup(d.page);
  if (ui.title) ui.title.textContent = d.title;
  lastWords = -1;
  void refreshVersions();
  setHeaderFooterSpace((i) => hfSpace(doc, hfHeights, i, sectionOfPage(pages()[i])));
  if (uiState().inspectorOpen) renderInspector();
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
  bindMobile();
  setPlatform(IS_MAC);
  loadUiState();
  setRecent(uiState().recent);
  register(buildCommands(commandActions()));
  setCommandsHooks(() => setUiState({ recent: recentIds() }));
  applyPanels();
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
    onTitleMenu: (anchor) => openFileMenu(anchor),
    onCollapse: () => toggleNav(),
    onRename: (next) => {
      doc.title = next;
      scheduleSave();
      refreshChrome();
      refreshLibrary();
    },
  });

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

  /**
   * One keyboard handler, reading one registry.
   *
   * Shortcuts used to live in three places - here, in commands.ts, and in
   * the find bar - and a binding could exist in one and be missing from
   * the palette. Now a command carries its own chord and this is the only
   * thing that dispatches one.
   *
   * commands.ts still owns Enter, Backspace, Delete, Tab and the arrows:
   * those are editing, not commands, and they have to run inside the
   * contenteditable rather than over it.
   */
  window.addEventListener('keydown', (e) => {
    if (e.isComposing) return;

    // Escape belongs to whatever is open, before anything else looks.
    if (e.key === 'Escape') {
      if (isCommandsOpen()) return; // the palette closes itself
      if (isFindOpen()) {
        e.preventDefault();
        closeFind();
        return;
      }
      if (isEditingHF()) {
        e.preventDefault();
        exitHF();
        return;
      }
    }
    if (isCommandsOpen() || isAskOpen()) return; // they own the keyboard

    const hit = commandForEvent(e);
    if (!hit || !isEnabled(hit)) return;
    e.preventDefault();
    runCommand(hit.id);
    setUiState({ recent: recentIds() });
  });

  setStatusHooks(() => openInspector('page'));
  bindInspector({
    doc: () => doc,
    setPaper: (width, height) => {
      doc.page = { ...doc.page, width, height };
      afterPageChange();
    },
    setMargins: (margins) => {
      doc.page = { ...doc.page, margins };
      if (doc.sections) {
        doc.sections = doc.sections.map((sct) => ({ ...sct, page: { ...sct.page, margins } }));
      }
      afterPageChange();
    },
    setMarginPreset: (key) => setMarginPreset(key),
    setDefaultFont: (family) => {
      if (family) doc.defaultFont = family;
      else delete doc.defaultFont;
      setDocumentFont(doc.defaultFont);
      afterPageChange();
    },
    setBodySize: (pt) => styleOverride('Body', 'size', pt),
    setBodyLeading: (mult) => styleOverride('Body', 'lineHeight', mult),
    setTitlePage: (on) => {
      if (on) doc.titlePage = true;
      else delete doc.titlePage;
      layout();
      scheduleSave();
      renderInspector();
    },
    setEvenOdd: (on) => {
      if (on) doc.evenOdd = true;
      else delete doc.evenOdd;
      layout();
      scheduleSave();
      renderInspector();
    },
    editHeaderFooter: () => toggleHF(),
    print: () => void printDocument(),
  });

  bindFormatBar(() => updateToolbar());
  renderFormatBar();
  // The inspector may have been left open: it is bound after the document
  // opens, so the first draw has to happen here rather than in openDoc.
  if (uiState().inspectorOpen) renderInspector();
  bindBubble({
    bold: () => toggleInline('bold'),
    italic: () => toggleInline('italic'),
    underline: () => toggleInline('underline'),
    clearFormatting: () => {
      setRunStyle({ size: null, font: null, color: null, strike: null, vert: null });
      clearBlockFormat();
    },
    inlineState,
    // Header editing and the palette both move the selection somewhere the
    // bubble has no business formatting.
    suppressed: () => isEditingHF() || isCommandsOpen() || isAskOpen(),
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
