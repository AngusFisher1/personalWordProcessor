/**
 * The command registry.
 *
 * One place where an action is defined, and five places that read it: the
 * palette, the menus, the format bar, the selection bubble and the keyboard.
 * Before this, a command like Bold was wired three times - a rail button, a
 * palette entry and a keydown branch - and the three drifted. Adding a
 * command here makes it reachable everywhere at once, which is the only way
 * the rule "every action is in the palette" can stay true without anybody
 * remembering to keep it true.
 *
 * Nothing here touches the DOM, so the registry can be built and checked
 * outside a browser.
 */

export type CommandCategory =
  | 'Format'
  | 'Insert'
  | 'Document'
  | 'File'
  | 'View'
  | 'App';

export const CATEGORIES: CommandCategory[] = [
  'Format',
  'Insert',
  'Document',
  'File',
  'View',
  'App',
];

export interface CommandDef {
  /** Stable, dotted, and never shown: `format.bold`, `file.export`. */
  id: string;
  /** What the palette and the menus call it. A function when it varies. */
  label: string | (() => string);
  category: CommandCategory;
  /**
   * Canonical shortcut: `Mod+B`, `Mod+Shift+E`, `F2`. `Mod` is the
   * platform's own - Command on a Mac, Control elsewhere - because
   * accepting either turns Ctrl+E into a dialog for a Mac user who meant
   * to move to the end of the line.
   */
  shortcut?: string;
  /**
   * Second binding for the same command, for a key people already have in
   * their fingers: Ctrl+Y for redo, which Word has had for thirty years.
   */
  aliases?: string[];
  /** Extra words to match on that are not worth showing. */
  keywords?: string;
  /** Radio or checkbox state, read when the command is displayed. */
  checked?: () => boolean;
  /** Shown but not runnable. Defaults to true. */
  enabled?: () => boolean;
  /** Not shown at all. Defaults to true. */
  visible?: () => boolean;
  run(): void;
}

/* ------------------------------------------------------------------ *
 * The register
 * ------------------------------------------------------------------ */

const byId = new Map<string, CommandDef>();
const order: string[] = [];

export function register(defs: CommandDef[]): void {
  for (const def of defs) {
    if (!byId.has(def.id)) order.push(def.id);
    byId.set(def.id, def);
  }
}

export function clearRegistry(): void {
  byId.clear();
  order.length = 0;
}

export function allCommands(): CommandDef[] {
  return order.map((id) => byId.get(id) as CommandDef);
}

export function command(id: string): CommandDef | undefined {
  return byId.get(id);
}

/** Every command that wants to be seen right now, in registration order. */
export function visibleCommands(): CommandDef[] {
  return allCommands().filter((c) => c.visible?.() ?? true);
}

export function labelOf(c: CommandDef): string {
  return typeof c.label === 'function' ? c.label() : c.label;
}

export function isEnabled(c: CommandDef): boolean {
  return c.enabled?.() ?? true;
}

/** Run by id, unless it is disabled. Returns whether it ran. */
export function runCommand(id: string): boolean {
  const c = byId.get(id);
  if (!c || !isEnabled(c)) return false;
  c.run();
  noteUsed(id);
  return true;
}

/* ------------------------------------------------------------------ *
 * Recently used
 *
 * An empty palette should open on what you actually do, not on whatever
 * happened to be registered first.
 * ------------------------------------------------------------------ */

const RECENT_MAX = 8;
let recent: string[] = [];

export function noteUsed(id: string): void {
  recent = [id, ...recent.filter((r) => r !== id)].slice(0, RECENT_MAX);
}

export function recentCommands(): CommandDef[] {
  return recent
    .map((id) => byId.get(id))
    .filter((c): c is CommandDef => !!c && (c.visible?.() ?? true));
}

export function setRecent(ids: string[]): void {
  recent = ids.slice(0, RECENT_MAX);
}

export function recentIds(): string[] {
  return recent.slice();
}

/* ------------------------------------------------------------------ *
 * Shortcuts
 * ------------------------------------------------------------------ */

export interface Chord {
  mod: boolean;
  shift: boolean;
  alt: boolean;
  /** Lower case for letters; as written for named keys. */
  key: string;
}

export function parseChord(shortcut: string): Chord | null {
  const parts = shortcut.split('+').map((p) => p.trim());
  const key = parts.pop();
  if (!key) return null;
  const chord: Chord = { mod: false, shift: false, alt: false, key: key.toLowerCase() };
  for (const p of parts) {
    const n = p.toLowerCase();
    if (n === 'mod') chord.mod = true;
    else if (n === 'shift') chord.shift = true;
    else if (n === 'alt') chord.alt = true;
    else return null;
  }
  return chord;
}

let isMac = false;

/** Set once at boot; kept out of module scope so Node can import this. */
export function setPlatform(mac: boolean): void {
  isMac = mac;
}

export function modLabel(): string {
  return isMac ? '⌘' : 'Ctrl+';
}

export function shiftLabel(): string {
  return isMac ? '⇧' : 'Shift+';
}

/** `Mod+Shift+E` as a person reads it on this machine. */
export function formatShortcut(shortcut: string | undefined): string {
  if (!shortcut) return '';
  const chord = parseChord(shortcut);
  if (!chord) return shortcut;
  const pretty: Record<string, string> = {
    '\\': '\\',
    '.': '.',
    ',': ',',
    arrowup: '↑',
    arrowdown: '↓',
    enter: '↵',
    escape: 'Esc',
  };
  const key = pretty[chord.key] ?? chord.key.toUpperCase();
  return (
    (chord.mod ? modLabel() : '') +
    (chord.shift ? shiftLabel() : '') +
    (chord.alt ? (isMac ? '⌥' : 'Alt+') : '') +
    key
  );
}

/**
 * The command a key event should run, if any.
 *
 * Only the platform's own modifier counts. Accepting either means a Mac
 * user's Ctrl+K - kill to end of line, in every text field they have ever
 * used - opens a dialog instead.
 */
export function commandForEvent(e: {
  key: string;
  metaKey: boolean;
  ctrlKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}): CommandDef | null {
  const mod = isMac ? e.metaKey && !e.ctrlKey : e.ctrlKey && !e.metaKey;
  const key = e.key.toLowerCase();
  for (const c of allCommands()) {
    for (const binding of [c.shortcut, ...(c.aliases ?? [])]) {
      if (!binding) continue;
      const chord = parseChord(binding);
      if (!chord) continue;
      if (chord.key !== key) continue;
      if (chord.mod !== mod) continue;
      if (chord.shift !== e.shiftKey) continue;
      if (chord.alt !== e.altKey) continue;
      if (!(c.visible?.() ?? true)) continue;
      return c;
    }
  }
  return null;
}

/* ------------------------------------------------------------------ *
 * Self-checks, for the test suite
 * ------------------------------------------------------------------ */

export interface RegistryProblem {
  kind: 'duplicate-id' | 'duplicate-shortcut' | 'missing-label' | 'bad-shortcut';
  detail: string;
}

/**
 * Everything that can be wrong with a set of commands without running them.
 *
 * Takes the list rather than reading the register, so a test can check a
 * freshly built set without a browser to register it in.
 */
export function auditCommands(defs: CommandDef[]): RegistryProblem[] {
  const problems: RegistryProblem[] = [];
  const ids = new Set<string>();
  const chords = new Map<string, string>();

  for (const c of defs) {
    if (ids.has(c.id)) problems.push({ kind: 'duplicate-id', detail: c.id });
    ids.add(c.id);

    const label = typeof c.label === 'function' ? c.label() : c.label;
    if (!label || label.trim() === '') {
      problems.push({ kind: 'missing-label', detail: c.id });
    }
    if (!CATEGORIES.includes(c.category)) {
      problems.push({ kind: 'missing-label', detail: c.id + ' has category ' + c.category });
    }
    for (const binding of [c.shortcut, ...(c.aliases ?? [])]) {
      if (!binding) continue;
      const chord = parseChord(binding);
      if (!chord) {
        problems.push({ kind: 'bad-shortcut', detail: c.id + ': ' + binding });
        continue;
      }
      const key = [chord.mod && 'mod', chord.shift && 'shift', chord.alt && 'alt', chord.key]
        .filter(Boolean)
        .join('+');
      const owner = chords.get(key);
      if (owner) {
        problems.push({
          kind: 'duplicate-shortcut',
          detail: binding + ' on both ' + owner + ' and ' + c.id,
        });
      }
      chords.set(key, c.id);
    }
  }
  return problems;
}
