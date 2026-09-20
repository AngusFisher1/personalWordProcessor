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
import { STYLES, injectStyleSheet } from './styles';
import { docEl, readModel, renderAll } from './render';
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

let doc: Doc = emptyDoc();
let saveFailed = false;
/** The original .docx package, when this document came from one. */
let vault: Vault | null = null;

/* ------------------------------------------------------------------ *
 * Toolbar
 * ------------------------------------------------------------------ */

const ui = {
  style: null as HTMLSelectElement | null,
  margin: null as HTMLSelectElement | null,
  title: null as HTMLInputElement | null,
  bold: null as HTMLButtonElement | null,
  italic: null as HTMLButtonElement | null,
  underline: null as HTMLButtonElement | null,
  undo: null as HTMLButtonElement | null,
  redo: null as HTMLButtonElement | null,
  status: null as HTMLSpanElement | null,
};

function button(
  label: string,
  title: string,
  cls: string,
  onClick: () => void
): HTMLButtonElement {
  const b = document.createElement('button');
  b.textContent = label;
  b.title = title;
  if (cls) b.className = cls;
  // mousedown default would blur the document and drop the selection.
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', onClick);
  return b;
}

function sep(): HTMLSpanElement {
  const s = document.createElement('span');
  s.className = 'sep';
  return s;
}

function buildToolbar(bar: HTMLElement): void {
  const mod = navigator.platform.toLowerCase().includes('mac') ? 'Cmd' : 'Ctrl';

  const title = document.createElement('input');
  title.type = 'text';
  title.value = doc.title;
  title.title = 'Document title';
  title.style.cssText =
    'font:13px var(--ui);height:28px;border:1px solid transparent;' +
    'border-radius:4px;padding:0 8px;width:180px;';
  title.addEventListener('focus', () => (title.style.borderColor = '#d8d8dc'));
  title.addEventListener('blur', () => (title.style.borderColor = 'transparent'));
  title.addEventListener('input', () => {
    doc.title = title.value || 'Untitled';
    scheduleSave();
  });
  ui.title = title;
  bar.appendChild(title);

  bar.appendChild(sep());

  ui.undo = button('↶', `Undo (${mod}+Z)`, '', () => {
    undo();
    afterChange();
  });
  ui.redo = button('↷', `Redo (${mod}+Shift+Z)`, '', () => {
    redo();
    afterChange();
  });
  bar.appendChild(ui.undo);
  bar.appendChild(ui.redo);

  bar.appendChild(sep());

  const style = document.createElement('select');
  style.title = 'Paragraph style';
  for (const id of STYLE_IDS) {
    const o = document.createElement('option');
    o.value = id;
    o.textContent = STYLES[id].label;
    style.appendChild(o);
  }
  style.addEventListener('mousedown', () => saveSelection());
  style.addEventListener('change', () => {
    restoreSelection();
    setBlockStyle(style.value as StyleId);
  });
  ui.style = style;
  bar.appendChild(style);

  bar.appendChild(sep());

  ui.bold = button('B', `Bold (${mod}+B)`, 'b', () => toggleInline('bold'));
  ui.italic = button('I', `Italic (${mod}+I)`, 'i', () => toggleInline('italic'));
  ui.underline = button('U', `Underline (${mod}+U)`, 'u', () =>
    toggleInline('underline')
  );
  bar.appendChild(ui.bold);
  bar.appendChild(ui.italic);
  bar.appendChild(ui.underline);

  bar.appendChild(sep());

  const margin = document.createElement('select');
  margin.title = 'Page margins';
  for (const key of Object.keys(MARGINS) as MarginKey[]) {
    const o = document.createElement('option');
    o.value = key;
    o.textContent =
      key === 'narrow' ? 'Narrow margins (0.5in)' : 'Normal margins (1in)';
    margin.appendChild(o);
  }
  margin.addEventListener('change', () => {
    // Only the margins change; an imported page size is left alone.
    doc.page = { ...doc.page, margins: uniformMargins(MARGINS[margin.value as MarginKey]) };
    setPageSetup(doc.page);
    reflowNow();
    scheduleSave();
  });
  ui.margin = margin;
  bar.appendChild(margin);

  bar.appendChild(sep());

  bar.appendChild(
    button('Open .docx', 'Open a Word document', '', () => {
      void pickDocx();
    })
  );
  bar.appendChild(
    button('Print / PDF', 'Print to PDF', '', () => {
      void printDocument();
    })
  );
  bar.appendChild(
    button('.docx', 'Export as Word document', '', () => {
      void exportWord();
    })
  );
  bar.appendChild(
    button('Export JSON', 'Download the document as JSON', '', () => {
      syncModel();
      download(
        safeName(doc.title) + '.json',
        new Blob([exportJson(doc)], { type: 'application/json' })
      );
    })
  );
  bar.appendChild(
    button('Import JSON', 'Replace the document from a JSON file', '', () => {
      void pickJson();
    })
  );

  const spacer = document.createElement('span');
  spacer.className = 'spacer';
  bar.appendChild(spacer);

  const status = document.createElement('span');
  status.className = 'status';
  ui.status = status;
  bar.appendChild(status);
}

/* ------------------------------------------------------------------ *
 * Selection bookkeeping for toolbar widgets that steal focus
 * ------------------------------------------------------------------ */

let stashed: Range | null = null;

function saveSelection(): void {
  const s = window.getSelection();
  stashed = s && s.rangeCount ? s.getRangeAt(0).cloneRange() : null;
}

function restoreSelection(): void {
  if (!stashed) return;
  const s = window.getSelection();
  if (!s) return;
  docEl().focus();
  s.removeAllRanges();
  s.addRange(stashed);
}

/* ------------------------------------------------------------------ *
 * State sync
 * ------------------------------------------------------------------ */

function syncModel(): void {
  doc.blocks = readModel(docEl());
}

function updateToolbar(): void {
  const st = currentStyle();
  if (ui.style) ui.style.value = st ?? '';
  const inline = inlineState();
  ui.bold?.classList.toggle('on', inline.bold);
  ui.italic?.classList.toggle('on', inline.italic);
  ui.underline?.classList.toggle('on', inline.underline);
  if (ui.undo) ui.undo.disabled = !canUndo();
  if (ui.redo) ui.redo.disabled = !canRedo();
  if (ui.status) {
    ui.status.classList.toggle('err', saveFailed);
    ui.status.textContent = saveFailed
      ? 'Save failed - storage full. Export JSON to keep your work.'
      : `${pageCount()} page${pageCount() === 1 ? '' : 's'}`;
  }
}

let saveTimer = 0;
function scheduleSave(): void {
  clearTimeout(saveTimer);
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
  if (ui.margin) ui.margin.value = marginPreset(d.page) ?? '';
  if (ui.title) ui.title.value = d.title;
  renderAll(doc, docEl());
  normalize();
  ensureTrailingBlock();
  paginate();
  resetHistory();
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
    saveFailed = !(e as CustomEvent).detail?.ok;
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
