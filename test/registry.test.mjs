/**
 * The command registry.
 *
 * Everything in the interface reads from one list, so the list is the one
 * place a mistake can hide: two commands on the same shortcut, a command
 * with no label, an id that collides with another and silently replaces it.
 * None of that needs a browser to find, which is why the list is built from
 * injected actions rather than imported ones.
 *
 *   npm run test:registry
 */
const {
  buildCommands,
  auditCommands,
  parseChord,
  formatShortcut,
  setPlatform,
  register,
  clearRegistry,
  allCommands,
  visibleCommands,
  runCommand,
  commandForEvent,
  recentCommands,
  noteUsed,
  CATEGORIES,
} = await import('./build/harness.mjs');

let failures = 0;
function check(name, ok, detail = '') {
  if (ok) { console.log('  ok    ' + name); return; }
  failures++; console.log('  FAIL  ' + name + (detail ? ' - ' + detail : ''));
}

/** Every action, recording what was called. */
function stubActions() {
  const calls = [];
  const note = (name) => (...args) => calls.push(name + (args.length ? ':' + args.join(',') : ''));
  return {
    calls,
    setStyle: note('setStyle'),
    currentStyleId: () => 'Body',
    toggleInline: note('toggleInline'),
    inlineState: () => ({ bold: false, italic: false, underline: false }),
    setRunStyle: note('setRunStyle'),
    currentRunStyle: () => ({}),
    hasSelection: () => true,
    setBlockFormat: note('setBlockFormat'),
    currentFormat: () => ({}),
    clearBlockFormat: note('clearBlockFormat'),
    insertTable: note('insertTable'),
    insertImage: note('insertImage'),
    insertLink: note('insertLink'),
    linkAtCaret: () => null,
    togglePageBreak: note('togglePageBreak'),
    pageBreakHere: () => false,
    undo: note('undo'),
    redo: note('redo'),
    find: note('find'),
    toggleHeaderFooter: note('toggleHeaderFooter'),
    setMarginPreset: note('setMarginPreset'),
    marginPreset: () => 'narrow',
    newDocument: note('newDocument'),
    renameDocument: note('renameDocument'),
    duplicateDocument: note('duplicateDocument'),
    deleteDocument: note('deleteDocument'),
    openWordFile: note('openWordFile'),
    importMarkdown: note('importMarkdown'),
    importJson: note('importJson'),
    print: note('print'),
    showExportSheet: note('showExportSheet'),
    exportWord: note('exportWord'),
    exportMarkdown: note('exportMarkdown'),
    exportHtml: note('exportHtml'),
    exportJson: note('exportJson'),
    exportLibrary: note('exportLibrary'),
    versionHistory: note('versionHistory'),
    toggleNav: note('toggleNav'),
    navCollapsed: () => false,
    focusOutline: note('focusOutline'),
    focusFiles: note('focusFiles'),
    toggleInspector: note('toggleInspector'),
    inspectorOpen: () => false,
    openPageSetup: note('openPageSetup'),
    openPalette: note('openPalette'),
    openSettings: note('openSettings'),
    setPalette: note('setPalette'),
    currentPaletteId: () => 'mossChalk',
  };
}

const actions = stubActions();
const defs = buildCommands(actions);

/* ------------------------------------------------------------------ */
console.log('\nthe list is well formed');
{
  const problems = auditCommands(defs);
  const of = (kind) => problems.filter((p) => p.kind === kind);

  check('there are commands at all', defs.length > 40, String(defs.length));
  check('no duplicate ids', of('duplicate-id').length === 0,
    of('duplicate-id').map((p) => p.detail).join('; '));
  check('no duplicate shortcuts', of('duplicate-shortcut').length === 0,
    of('duplicate-shortcut').map((p) => p.detail).join('; '));
  check('every command has a label', of('missing-label').length === 0,
    of('missing-label').map((p) => p.detail).join('; '));
  check('every shortcut parses', of('bad-shortcut').length === 0,
    of('bad-shortcut').map((p) => p.detail).join('; '));
  check('nothing else is wrong', problems.length === 0, String(problems.length));
}

/* ------------------------------------------------------------------ */
console.log('\nevery command is categorised and dotted');
{
  const uncategorised = defs.filter((c) => !CATEGORIES.includes(c.category));
  check('every category is a known one', uncategorised.length === 0,
    uncategorised.map((c) => c.id).join(', '));

  const undotted = defs.filter((c) => !/^[a-z]+\.[A-Za-z0-9.×x-]+$/.test(c.id));
  check('every id is namespaced', undotted.length === 0, undotted.map((c) => c.id).join(', '));

  const perCategory = new Map();
  for (const c of defs) perCategory.set(c.category, (perCategory.get(c.category) ?? 0) + 1);
  check('every category is used', perCategory.size === CATEGORIES.length,
    [...perCategory].map(([k, v]) => k + '=' + v).join(' '));
}

/* ------------------------------------------------------------------ */
console.log('\nthe things a button used to do are all still commands');
{
  // The Step 0 inventory, as a checklist. Every control that was removed
  // from the sidebar has to be reachable from the palette instead.
  const must = [
    'app.palette', 'app.settings', 'document.find', 'document.headerFooter',
    'document.undo', 'document.redo', 'document.versions',
    'file.new', 'file.rename', 'file.duplicate', 'file.delete',
    'file.openWord', 'file.importMarkdown', 'file.importJson',
    'file.print', 'file.export', 'file.exportWord', 'file.exportMarkdown',
    'file.exportHtml', 'file.exportJson', 'file.exportLibrary',
    'insert.image', 'insert.link', 'insert.pageBreak',
    'format.bold', 'format.italic', 'format.underline',
    'format.clearCharacter', 'format.clearParagraph',
    'view.nav', 'view.outline', 'view.files', 'view.inspector', 'view.pageSetup',
  ];
  const ids = new Set(defs.map((c) => c.id));
  const missing = must.filter((id) => !ids.has(id));
  check('nothing from the inventory is missing', missing.length === 0, missing.join(', '));

  check('every theme is a command', defs.filter((c) => c.id.startsWith('app.theme.')).length === 12);
  check('every style is a command', defs.filter((c) => c.id.startsWith('format.style.')).length === 6);
  check('both margin presets are commands',
    defs.filter((c) => c.id.startsWith('document.margins.')).length === 2);
}

/* ------------------------------------------------------------------ */
console.log('\nchords parse and print');
{
  setPlatform(false);
  check('Mod+B', formatShortcut('Mod+B') === 'Ctrl+B', formatShortcut('Mod+B'));
  check('Mod+Shift+E', formatShortcut('Mod+Shift+E') === 'Ctrl+Shift+E', formatShortcut('Mod+Shift+E'));
  check('F2 needs no modifier', formatShortcut('F2') === 'F2', formatShortcut('F2'));
  check('Mod+\\ survives', formatShortcut('Mod+\\') === 'Ctrl+\\', formatShortcut('Mod+\\'));
  setPlatform(true);
  check('on a Mac it is a command symbol', formatShortcut('Mod+B') === '⌘B', formatShortcut('Mod+B'));
  setPlatform(false);

  check('a chord round-trips', parseChord('Mod+Shift+K')?.key === 'k');
  check('and carries its modifiers',
    parseChord('Mod+Shift+K')?.mod === true && parseChord('Mod+Shift+K')?.shift === true);
  check('an unknown modifier is rejected', parseChord('Hyper+K') === null);
}

/* ------------------------------------------------------------------ */
console.log('\nthe register runs and remembers');
{
  clearRegistry();
  register(defs);
  check('everything registered', allCommands().length === defs.length);
  check('registering twice does not duplicate',
    (register(defs), allCommands().length === defs.length), String(allCommands().length));

  actions.calls.length = 0;
  check('running by id works', runCommand('format.bold') === true);
  check('and called the action', actions.calls.includes('toggleInline:bold'), actions.calls.join(','));
  check('an unknown id does nothing', runCommand('nope.nothing') === false);

  check('the last thing run is the first one remembered',
    recentCommands()[0]?.id === 'format.bold', recentCommands()[0]?.id);
  noteUsed('file.print');
  check('and the one before it is next',
    recentCommands()[0]?.id === 'file.print' && recentCommands()[1]?.id === 'format.bold');
}

/* ------------------------------------------------------------------ */
console.log('\nkey events find their command');
{
  setPlatform(false);
  const ev = (key, o = {}) => ({
    key, metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, ...o,
  });
  check('Ctrl+B is bold', commandForEvent(ev('b', { ctrlKey: true }))?.id === 'format.bold');
  check('plain b is not', commandForEvent(ev('b')) === null);
  check('Ctrl+Shift+Z is redo',
    commandForEvent(ev('z', { ctrlKey: true, shiftKey: true }))?.id === 'document.redo');
  check('Ctrl+Z is undo, not redo',
    commandForEvent(ev('z', { ctrlKey: true }))?.id === 'document.undo');
  check('the Ctrl+Y alias also redoes',
    commandForEvent(ev('y', { ctrlKey: true }))?.id === 'document.redo');
  check('Ctrl+E centres, it does not export',
    commandForEvent(ev('e', { ctrlKey: true }))?.id === 'format.align.center');
  check('Ctrl+Shift+E exports',
    commandForEvent(ev('e', { ctrlKey: true, shiftKey: true }))?.id === 'file.export');
  check('F2 renames', commandForEvent(ev('F2'))?.id === 'file.rename');

  setPlatform(true);
  check('on a Mac, Ctrl+K is not the palette',
    commandForEvent(ev('k', { ctrlKey: true })) === null);
  check('but Cmd+K is', commandForEvent(ev('k', { metaKey: true }))?.id === 'app.palette');
  setPlatform(false);
}

/* ------------------------------------------------------------------ */
console.log('\npredicates gate what is offered');
{
  const noSelection = { ...stubActions(), hasSelection: () => false };
  const gated = buildCommands(noSelection);
  const size = gated.find((c) => c.id === 'format.size.12');
  check('a size needs a selection', size?.enabled?.() === false);
  const bold = gated.find((c) => c.id === 'format.bold');
  check('bold does not', (bold?.enabled?.() ?? true) === true);

  clearRegistry();
  register(gated);
  check('a disabled command refuses to run', runCommand('format.size.12') === false);
  check('but is still visible', visibleCommands().some((c) => c.id === 'format.size.12'));
}

console.log(failures === 0 ? '\nall registry checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
