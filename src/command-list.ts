import type { BlockAlign, BlockFormat, MarginKey, RunStyle, StyleId } from './model';
import { MARGINS, STYLE_IDS } from './model';
import { STYLES } from './styles';
import { PALETTES } from './theme';
import type { CommandDef } from './registry';

/**
 * Every command in the program, in one list.
 *
 * The actions are injected rather than imported so this module can be built
 * and audited outside a browser - which is what lets a test assert that no
 * two commands share a shortcut without booting the editor to find out.
 */

export interface CommandActions {
  /* --- format --- */
  setStyle(id: StyleId): void;
  currentStyleId(): StyleId | null;
  toggleInline(which: 'bold' | 'italic' | 'underline'): void;
  inlineState(): { bold: boolean; italic: boolean; underline: boolean };
  setRunStyle(change: Partial<Record<keyof RunStyle, unknown>>): void;
  currentRunStyle(): RunStyle;
  hasSelection(): boolean;
  setBlockFormat(f: Partial<BlockFormat>): void;
  currentFormat(): BlockFormat;
  clearBlockFormat(): void;

  /* --- insert --- */
  insertTable(rows: number, cols: number): void;
  insertImage(): void;
  insertLink(): void;
  linkAtCaret(): string | null;
  togglePageBreak(): void;
  pageBreakHere(): boolean;

  /* --- document --- */
  undo(): void;
  redo(): void;
  find(): void;
  toggleHeaderFooter(): void;
  setMarginPreset(key: MarginKey): void;
  marginPreset(): MarginKey | null;

  /* --- file --- */
  newDocument(): void;
  renameDocument(): void;
  duplicateDocument(): void;
  deleteDocument(): void;
  openWordFile(): void;
  importMarkdown(): void;
  importJson(): void;
  print(): void;
  showExportSheet(): void;
  exportWord(): void;
  exportMarkdown(): void;
  exportHtml(): void;
  exportJson(): void;
  exportLibrary(): void;
  versionHistory(): void;

  /* --- view --- */
  toggleNav(): void;
  navCollapsed(): boolean;
  focusOutline(): void;
  focusFiles(): void;
  toggleInspector(): void;
  inspectorOpen(): boolean;
  openPageSetup(): void;

  /* --- app --- */
  openPalette(): void;
  openSettings(): void;
  setPalette(id: string): void;
  currentPaletteId(): string;
}

/** Points. The sizes a document actually uses, not a continuous spinner. */
export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 24, 36];

/**
 * A short palette of ink colours, named rather than picked: a document with
 * six arbitrary hex colours in it is one nobody can restyle later, and every
 * one of these reads on white paper in print.
 */
export const INK_COLORS: { label: string; value: string | null }[] = [
  { label: 'Default', value: null },
  { label: 'Black', value: '000000' },
  { label: 'Grey', value: '595959' },
  { label: 'Red', value: 'C00000' },
  { label: 'Orange', value: 'B45309' },
  { label: 'Green', value: '2E6B33' },
  { label: 'Blue', value: '1F4E79' },
  { label: 'Purple', value: '5B2D8E' },
];

export const ALIGNMENTS: { id: BlockAlign; label: string; shortcut: string }[] = [
  { id: 'left', label: 'Left', shortcut: 'Mod+L' },
  { id: 'center', label: 'Centre', shortcut: 'Mod+E' },
  { id: 'right', label: 'Right', shortcut: 'Mod+R' },
  { id: 'justify', label: 'Justified', shortcut: 'Mod+J' },
];

export const LINE_SPACINGS = [
  { label: 'Single', value: 1 },
  { label: '1.15', value: 1.15 },
  { label: 'One and a half', value: 1.5 },
  { label: 'Double', value: 2 },
];

/** Points, the unit the rest of the style table is in. */
export const SPACES = [
  { label: 'None', value: 0 },
  { label: '6 pt', value: 6 },
  { label: '12 pt', value: 12 },
];

export const TABLE_SIZES = [
  { rows: 2, cols: 2 },
  { rows: 3, cols: 2 },
  { rows: 3, cols: 3 },
  { rows: 4, cols: 3 },
  { rows: 5, cols: 2 },
  { rows: 6, cols: 4 },
];

export function buildCommands(a: CommandActions): CommandDef[] {
  const out: CommandDef[] = [];
  const push = (c: CommandDef): void => {
    out.push(c);
  };
  /** Every character-level command needs a range to act on. */
  const sel = () => a.hasSelection();

  /* ---------------- Format ---------------- */

  for (const id of STYLE_IDS) {
    push({
      id: 'format.style.' + id,
      label: STYLES[id].label,
      category: 'Format',
      keywords: 'paragraph style ' + id,
      checked: () => a.currentStyleId() === id,
      run: () => a.setStyle(id),
    });
  }

  for (const [which, key] of [
    ['bold', 'B'],
    ['italic', 'I'],
    ['underline', 'U'],
  ] as const) {
    push({
      id: 'format.' + which,
      label: which[0].toUpperCase() + which.slice(1),
      category: 'Format',
      shortcut: 'Mod+' + key,
      checked: () => a.inlineState()[which],
      run: () => a.toggleInline(which),
    });
  }

  for (const size of FONT_SIZES) {
    push({
      id: 'format.size.' + size,
      label: size + ' pt',
      category: 'Format',
      keywords: 'font size point text',
      enabled: sel,
      checked: () => a.currentRunStyle().size === size,
      run: () => a.setRunStyle({ size }),
    });
  }

  for (const c of INK_COLORS) {
    push({
      id: 'format.colour.' + (c.value ?? 'default'),
      label: 'Colour — ' + c.label,
      category: 'Format',
      keywords: 'colour color ink text',
      enabled: sel,
      checked: () => (a.currentRunStyle().color ?? null) === c.value,
      run: () => a.setRunStyle({ color: c.value }),
    });
  }

  push({
    id: 'format.strike',
    label: 'Strikethrough',
    category: 'Format',
    enabled: sel,
    checked: () => !!a.currentRunStyle().strike,
    run: () => a.setRunStyle({ strike: !a.currentRunStyle().strike }),
  });
  for (const vert of ['super', 'sub'] as const) {
    push({
      id: 'format.' + vert,
      label: vert === 'super' ? 'Superscript' : 'Subscript',
      category: 'Format',
      enabled: sel,
      checked: () => a.currentRunStyle().vert === vert,
      run: () =>
        a.setRunStyle({ vert: a.currentRunStyle().vert === vert ? null : vert }),
    });
  }
  push({
    id: 'format.clearCharacter',
    label: 'Clear character formatting',
    category: 'Format',
    keywords: 'reset size colour font',
    enabled: sel,
    run: () =>
      a.setRunStyle({ size: null, font: null, color: null, strike: null, vert: null }),
  });

  for (const al of ALIGNMENTS) {
    push({
      id: 'format.align.' + al.id,
      label: 'Align ' + al.label.toLowerCase(),
      category: 'Format',
      shortcut: al.shortcut,
      keywords: 'paragraph alignment justify centre center',
      checked: () => (a.currentFormat().align ?? 'left') === al.id,
      run: () => a.setBlockFormat({ align: al.id }),
    });
  }

  for (const l of LINE_SPACINGS) {
    push({
      id: 'format.leading.' + l.value,
      label: 'Line spacing — ' + l.label,
      category: 'Format',
      keywords: 'leading line height paragraph',
      checked: () => {
        const f = a.currentFormat();
        return (f.lineRule ?? 'auto') === 'auto' && f.lineHeight === l.value;
      },
      run: () => a.setBlockFormat({ lineHeight: l.value, lineRule: 'auto' }),
    });
  }

  for (const sp of SPACES) {
    push({
      id: 'format.spaceBefore.' + sp.value,
      label: 'Space before — ' + sp.label,
      category: 'Format',
      keywords: 'paragraph spacing above',
      checked: () => a.currentFormat().spaceBefore === sp.value,
      run: () => a.setBlockFormat({ spaceBefore: sp.value }),
    });
    push({
      id: 'format.spaceAfter.' + sp.value,
      label: 'Space after — ' + sp.label,
      category: 'Format',
      keywords: 'paragraph spacing below',
      checked: () => a.currentFormat().spaceAfter === sp.value,
      run: () => a.setBlockFormat({ spaceAfter: sp.value }),
    });
  }

  push({
    id: 'format.clearParagraph',
    label: 'Clear direct formatting',
    category: 'Format',
    keywords: 'reset alignment indent spacing to the style',
    run: () => a.clearBlockFormat(),
  });

  /* ---------------- Insert ---------------- */

  for (const t of TABLE_SIZES) {
    push({
      id: `insert.table.${t.rows}x${t.cols}`,
      label: `Table ${t.rows} × ${t.cols}`,
      category: 'Insert',
      keywords: 'insert table grid rows columns',
      run: () => a.insertTable(t.rows, t.cols),
    });
  }
  push({
    id: 'insert.image',
    label: 'Image…',
    category: 'Insert',
    keywords: 'picture photo insert',
    run: () => a.insertImage(),
  });
  push({
    id: 'insert.link',
    label: () => (a.linkAtCaret() ? 'Change link…' : 'Link…'),
    category: 'Insert',
    shortcut: 'Mod+Shift+K',
    keywords: 'hyperlink url insert',
    run: () => a.insertLink(),
  });
  push({
    id: 'insert.pageBreak',
    label: 'Page break above',
    category: 'Insert',
    keywords: 'insert page break',
    checked: () => a.pageBreakHere(),
    run: () => a.togglePageBreak(),
  });

  /* ---------------- Document ---------------- */

  push({
    id: 'document.undo',
    label: 'Undo',
    category: 'Document',
    shortcut: 'Mod+Z',
    run: () => a.undo(),
  });
  push({
    id: 'document.redo',
    label: 'Redo',
    category: 'Document',
    shortcut: 'Mod+Shift+Z',
    aliases: ['Mod+Y'],
    run: () => a.redo(),
  });
  push({
    id: 'document.find',
    label: 'Find and replace',
    category: 'Document',
    shortcut: 'Mod+F',
    run: () => a.find(),
  });
  push({
    id: 'document.headerFooter',
    label: 'Edit header and footer',
    category: 'Document',
    keywords: 'letterhead page number',
    run: () => a.toggleHeaderFooter(),
  });
  for (const key of Object.keys(MARGINS) as MarginKey[]) {
    push({
      id: 'document.margins.' + key,
      label: key === 'narrow' ? 'Narrow margins — 0.5 in' : 'Normal margins — 1 in',
      category: 'Document',
      keywords: 'margin page setup',
      checked: () => a.marginPreset() === key,
      run: () => a.setMarginPreset(key),
    });
  }
  push({
    id: 'document.versions',
    label: 'Version history',
    category: 'Document',
    keywords: 'versions earlier saves restore',
    run: () => a.versionHistory(),
  });

  /* ---------------- File ---------------- */

  push({
    id: 'file.new',
    label: 'New document',
    category: 'File',
    shortcut: 'Mod+N',
    run: () => a.newDocument(),
  });
  push({
    id: 'file.rename',
    label: 'Rename…',
    category: 'File',
    shortcut: 'F2',
    run: () => a.renameDocument(),
  });
  push({
    id: 'file.duplicate',
    label: 'Duplicate',
    category: 'File',
    run: () => a.duplicateDocument(),
  });
  push({
    id: 'file.delete',
    label: 'Delete…',
    category: 'File',
    keywords: 'remove discard',
    run: () => a.deleteDocument(),
  });
  push({
    id: 'file.openWord',
    label: 'Open Word document…',
    category: 'File',
    keywords: 'docx import',
    run: () => a.openWordFile(),
  });
  push({
    id: 'file.importMarkdown',
    label: 'Import Markdown…',
    category: 'File',
    run: () => a.importMarkdown(),
  });
  push({
    id: 'file.importJson',
    label: 'Import JSON…',
    category: 'File',
    run: () => a.importJson(),
  });
  push({
    id: 'file.print',
    label: 'Print / save as PDF',
    category: 'File',
    shortcut: 'Mod+P',
    run: () => a.print(),
  });
  push({
    id: 'file.export',
    label: 'Export…',
    category: 'File',
    shortcut: 'Mod+Shift+E',
    keywords: 'save as formats sheet',
    run: () => a.showExportSheet(),
  });
  push({
    id: 'file.exportWord',
    label: 'Save as Word (.docx)',
    category: 'File',
    keywords: 'docx',
    run: () => a.exportWord(),
  });
  push({
    id: 'file.exportMarkdown',
    label: 'Export Markdown',
    category: 'File',
    run: () => a.exportMarkdown(),
  });
  push({
    id: 'file.exportHtml',
    label: 'Export HTML',
    category: 'File',
    run: () => a.exportHtml(),
  });
  push({
    id: 'file.exportJson',
    label: 'Export JSON',
    category: 'File',
    run: () => a.exportJson(),
  });
  push({
    id: 'file.exportLibrary',
    label: 'Export everything…',
    category: 'File',
    keywords: 'backup zip archive independence library',
    run: () => a.exportLibrary(),
  });

  /* ---------------- View ---------------- */

  push({
    id: 'view.nav',
    label: () => (a.navCollapsed() ? 'Show the side panel' : 'Hide the side panel'),
    category: 'View',
    shortcut: 'Mod+\\',
    keywords: 'sidebar nav collapse expand',
    run: () => a.toggleNav(),
  });
  push({
    id: 'view.outline',
    label: 'Outline',
    category: 'View',
    shortcut: 'Mod+Shift+O',
    keywords: 'headings navigate',
    run: () => a.focusOutline(),
  });
  push({
    id: 'view.files',
    label: 'Browse documents',
    category: 'View',
    shortcut: 'Mod+O',
    keywords: 'library open recent files',
    run: () => a.focusFiles(),
  });
  push({
    id: 'view.inspector',
    label: () => (a.inspectorOpen() ? 'Hide document settings' : 'Document settings'),
    category: 'View',
    shortcut: 'Mod+.',
    keywords: 'inspector page setup margins panel',
    run: () => a.toggleInspector(),
  });
  push({
    id: 'view.pageSetup',
    label: 'Page setup',
    category: 'View',
    keywords: 'margins paper size font leading',
    run: () => a.openPageSetup(),
  });

  /* ---------------- App ---------------- */

  push({
    id: 'app.palette',
    label: 'Commands',
    category: 'App',
    shortcut: 'Mod+K',
    keywords: 'command palette everything',
    run: () => a.openPalette(),
  });
  push({
    id: 'app.settings',
    label: 'Settings',
    category: 'App',
    shortcut: 'Mod+,',
    keywords: 'preferences theme storage styles',
    run: () => a.openSettings(),
  });
  for (const p of PALETTES) {
    push({
      id: 'app.theme.' + p.id,
      label: 'Theme — ' + p.label,
      category: 'App',
      keywords: (p.dark ? 'dark' : 'light') + ' theme colour color palette ' + p.note,
      checked: () => a.currentPaletteId() === p.id,
      run: () => a.setPalette(p.id),
    });
  }

  return out;
}
