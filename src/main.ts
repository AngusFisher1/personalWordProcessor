// styles.css is linked from index.html so the page geometry is applied
// before this module runs: the first measurement must not happen unstyled.
import type { Doc, MarginKey, StyleId } from './model';
import { MARGINS, STYLE_IDS, emptyDoc, newBlock } from './model';
import { STYLES, injectStyleSheet } from './styles';
import { docEl, readModel, renderAll } from './render';
import {
  clearHeightCache,
  currentMargin,
  ensureTrailingBlock,
  normalize,
  pageCount,
  paginate,
  paginateIfNeeded,
  setMargin,
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

let doc: Doc = emptyDoc();
let saveFailed = false;

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
    doc.margin = margin.value as MarginKey;
    setMargin(doc.margin);
    reflowNow();
    scheduleSave();
  });
  ui.margin = margin;
  bar.appendChild(margin);

  bar.appendChild(sep());

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
  const blob = await exportDocx(doc);
  download(safeName(doc.title) + '.docx', blob);
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
  setMargin(d.margin);
  if (ui.margin) ui.margin.value = d.margin;
  if (ui.title) ui.title.value = d.title;
  renderAll(doc, docEl());
  normalize();
  ensureTrailingBlock();
  paginate();
  resetHistory();
  updateToolbar();
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

  openDoc(load() ?? sampleDoc());

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
    margin: currentMargin,
  },
});
