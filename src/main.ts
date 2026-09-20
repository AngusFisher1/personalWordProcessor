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
import { blockEl, blocksIn, docEl, readModel, renderAll } from './render';
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
import { closeMenu, iconButton, menuButton, textButton } from './ui';
import { PALETTES, applyPalette, currentPaletteId } from './theme';
import {
  activeHeadingId,
  buildRail,
  caretReadout,
  currentPageIndex,
  railEl,
  updateGutter,
  updateOutline,
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
    textButton('Print', `Print or save as PDF (${MOD}P)`, () => void printDocument())
  );
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

function setPalette(id: string): void {
  const p = applyPalette(id);
  ui.palette?.setLabel(p.label);
  // Page shadows and rules changed, but nothing about the text did, so no
  // reflow is needed - only the chrome that depends on the accent.
  refreshChrome();
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
  if (ui.title) ui.title.value = d.title;
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

  // The gutter and the document map are positioned against the canvas, so
  // they follow the scroll rather than being redrawn by the editor.
  root.addEventListener('scroll', () => {
    const page = currentPageIndex();
    updateGutter(page);
    updateSpine(page);
  });
  window.addEventListener('resize', () => refreshChrome());

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
