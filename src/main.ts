// styles.css is linked from index.html so the page geometry is applied
// before this module runs: the first measurement must not happen unstyled.
import type { Doc, MarginKey, StyleId } from './model';
import {
  MARGINS,
  STYLE_IDS,
  emptyDoc,
  marginPreset,
  newBlock,
  uniformMargins,
} from './model';
import { DOC_FONT, STYLES, injectStyleSheet } from './styles';
import { blocksIn, docEl, readModel, renderAll } from './render';
import {
  clearHeightCache,
  currentPageSetup,
  ensureTrailingBlock,
  normalize,
  pageCount,
  paginate,
  paginateIfNeeded,
  setPageSetup,
} from './paginate';
import {
  bindShortcuts,
  currentStyle,
  inlineState,
  setBlockStyle,
  toggleInline,
} from './commands';
import { caretAtStart } from './caret';
import { bindPaste } from './paste';
import {
  canRedo,
  canUndo,
  noteTyping,
  redo,
  resetHistory,
  snapshot,
  undo,
} from './history';
import { download, exportJson, importJson, load, save } from './persist';
import type { Vault } from './docx-package';
import { loadOriginal, saveOriginal } from './docx-package';
import { importDocx } from './docx-import';
import {
  closeMenu,
  iconButton,
  menuButton,
  separator,
  textButton,
} from './ui';

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
  title: null as HTMLInputElement | null,
  bold: null as HTMLButtonElement | null,
  italic: null as HTMLButtonElement | null,
  underline: null as HTMLButtonElement | null,
  undo: null as HTMLButtonElement | null,
  redo: null as HTMLButtonElement | null,
  save: null as HTMLSpanElement | null,
  counts: null as HTMLSpanElement | null,
};

const IS_MAC = navigator.platform.toLowerCase().includes('mac');
const MOD = IS_MAC ? '⌘' : 'Ctrl+';
const SHIFT = IS_MAC ? '⇧' : 'Shift+';

function row(cls: string): HTMLDivElement {
  const d = document.createElement('div');
  d.className = 'tb-row ' + cls;
  return d;
}

function spacer(): HTMLSpanElement {
  const s = document.createElement('span');
  s.className = 'tb-spacer';
  return s;
}

/** Grow the title field with its text instead of sitting in a fixed box. */
function sizeTitle(input: HTMLInputElement): void {
  const chars = Math.max(8, (input.value || 'Untitled').length);
  input.style.width = Math.min(440, chars * 8.2 + 26) + 'px';
}

function buildToolbar(bar: HTMLElement): void {
  bar.textContent = '';
  const docRow = row('tb-doc');
  const fmtRow = row('tb-format');
  bar.append(docRow, fmtRow);

  /* ---- document row: what the file is, and whether it is safe ---- */

  const title = document.createElement('input');
  title.type = 'text';
  title.id = 'doctitle';
  title.value = doc.title;
  title.title = 'Document name';
  title.setAttribute('aria-label', 'Document name');
  title.addEventListener('input', () => {
    doc.title = title.value || 'Untitled';
    sizeTitle(title);
    scheduleSave();
  });
  ui.title = title;
  docRow.appendChild(title);
  sizeTitle(title);

  const saveState = document.createElement('span');
  saveState.className = 'tb-save';
  ui.save = saveState;
  docRow.appendChild(saveState);

  docRow.appendChild(spacer());

  const counts = document.createElement('span');
  counts.className = 'tb-counts';
  ui.counts = counts;
  docRow.appendChild(counts);

  /* ---- format row ---- */

  const file = menuButton('File', 'Open, save and export', () => [
    { label: 'Open Word document…', onSelect: () => void pickDocx() },
    { separator: true },
    { label: 'Save as Word (.docx)', onSelect: () => void exportWord() },
    { label: 'Print / save as PDF', hint: MOD + 'P', onSelect: () => void printDocument() },
    { separator: true },
    { heading: 'Plain formats' },
    {
      label: 'Export JSON',
      onSelect: () => {
        syncModel();
        download(
          safeName(doc.title) + '.json',
          new Blob([exportJson(doc)], { type: 'application/json' })
        );
      },
    },
    { label: 'Import JSON…', onSelect: () => void pickJson() },
  ]);
  fmtRow.appendChild(file.el);

  fmtRow.appendChild(separator());

  ui.undo = iconButton('undo', `Undo (${MOD}Z)`, () => {
    undo();
    afterChange();
  });
  ui.redo = iconButton('redo', `Redo (${MOD}${SHIFT}Z)`, () => {
    redo();
    afterChange();
  });
  fmtRow.append(ui.undo, ui.redo);

  fmtRow.appendChild(separator());

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
            fontSize: Math.min(20, Math.max(12, d.size)) + 'px',
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
  fmtRow.appendChild(ui.style.el);

  fmtRow.appendChild(separator());

  ui.bold = textButton('B', `Bold (${MOD}B)`, () => toggleInline('bold'), 'tb-b');
  ui.italic = textButton('I', `Italic (${MOD}I)`, () => toggleInline('italic'), 'tb-i');
  ui.underline = textButton(
    'U',
    `Underline (${MOD}U)`,
    () => toggleInline('underline'),
    'tb-u'
  );
  fmtRow.append(ui.bold, ui.italic, ui.underline);

  fmtRow.appendChild(separator());

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
  fmtRow.appendChild(ui.page.el);

  fmtRow.appendChild(separator());

  fmtRow.appendChild(
    textButton('Print', `Print or save as PDF (${MOD}P)`, () => void printDocument())
  );

  fmtRow.appendChild(spacer());
}

function setMarginPreset(key: MarginKey): void {
  // Only the margins change; an imported page size is left alone.
  doc.page = { ...doc.page, margins: uniformMargins(MARGINS[key]) };
  setPageSetup(doc.page);
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

  if (ui.save) {
    ui.save.classList.toggle('err', saveState === 'failed');
    ui.save.textContent =
      saveState === 'failed'
        ? 'Not saved — storage full. Export to keep your work.'
        : saveState === 'pending'
          ? 'Saving…'
          : 'All changes saved';
  }

  updateCounts();
}

let countsTimer = 0;
/**
 * Page count is free, but the word count reads the whole document, so it is
 * debounced rather than run on every keystroke.
 */
function updateCounts(): void {
  if (!ui.counts) return;
  const pages = pageCount();
  const pageText = `${pages} page${pages === 1 ? '' : 's'}`;
  ui.counts.textContent = pageText + (lastWords >= 0 ? ` · ${lastWords} words` : '');
  clearTimeout(countsTimer);
  countsTimer = window.setTimeout(() => {
    const words = (docEl().textContent ?? '').match(/\S+/g)?.length ?? 0;
    if (words !== lastWords) {
      lastWords = words;
      if (ui.counts) {
        ui.counts.textContent = `${pageCount()} page${pageCount() === 1 ? '' : 's'} · ${words} words`;
      }
    }
  }, 400);
}
let lastWords = -1;

let saveTimer = 0;
function scheduleSave(): void {
  clearTimeout(saveTimer);
  if (saveState !== 'failed') {
    saveState = 'pending';
    updateToolbar();
  }
  saveTimer = window.setTimeout(() => {
    syncModel();
    save(doc);
  }, 1000);
}

let reflowTimer = 0;
/** Debounced full reflow: paste, style change, margin change, font load, undo. */
function scheduleReflow(): void {
  clearTimeout(reflowTimer);
  reflowTimer = window.setTimeout(() => {
    paginate();
    updateToolbar();
  }, 150);
}

function reflowNow(): void {
  clearTimeout(reflowTimer);
  paginate();
  updateToolbar();
}



function afterChange(): void {
  updateToolbar();
  scheduleSave();
}

/* ------------------------------------------------------------------ *
 * Export
 * ------------------------------------------------------------------ */

function safeName(s: string): string {
  return (s || 'document').replace(/[^a-z0-9\-_ ]/gi, '').trim() || 'document';
}

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
  download(safeName(doc.title) + '.docx', blob);
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
    save(doc);
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

async function pickJson(): Promise<void> {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'application/json,.json';
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        doc = importJson(String(reader.result));
        vault = null;
        openDoc(doc);
        snapshot('structural');
        save(doc);
      } catch (err) {
        alert('Could not import that file: ' + (err as Error).message);
      }
    };
    reader.readAsText(file);
  });
  input.click();
}

/* ------------------------------------------------------------------ *
 * Boot
 * ------------------------------------------------------------------ */

function openDoc(d: Doc): void {
  doc = d;
  setPageSetup(d.page);
  if (ui.title) {
    ui.title.value = d.title;
    sizeTitle(ui.title);
  }
  lastWords = -1;
  renderAll(doc, docEl());
  normalize();
  ensureTrailingBlock();
  paginate();
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
  const bar = document.getElementById('toolbar');
  if (!bar) throw new Error('#toolbar missing');

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

  bindShortcuts(root);
  bindPaste(root);

  root.addEventListener('input', () => {
    closeMenu();
    normalize();
    if (ensureTrailingBlock()) paginate();
    paginateIfNeeded(); // synchronous, before paint
    noteTyping();
    updateToolbar();
    scheduleSave();
  });

  document.addEventListener('selectionchange', () => {
    if (document.activeElement === root) updateToolbar();
  });

  document.addEventListener('wp:changed', () => {
    updateToolbar();
    scheduleSave();
  });

  document.addEventListener('wp:saved', (e) => {
    saveState = (e as CustomEvent).detail?.ok ? 'saved' : 'failed';
    updateToolbar();
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

  window.addEventListener('beforeunload', () => {
    clearTimeout(saveTimer);
    syncModel();
    save(doc);
  });

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
  },
});
