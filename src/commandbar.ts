/**
 * The command palette.
 *
 * Every command in the program, reachable by typing part of its name. It
 * exists because the rail cannot grow forever: twelve palettes, six styles,
 * two margin presets, four export formats and a dozen actions do not fit in
 * 248px, and the ones that do fit are found by hunting rather than by
 * knowing what you want.
 *
 * Unlike every other control in this program, the palette DOES take focus:
 * it has a text field, so it must. That is why it captures the selection on
 * the way in and restores it before running anything - a command that
 * applies to the selection has to find the selection still there.
 */

import type { CommandDef } from './registry';
import {
  formatShortcut,
  isEnabled,
  labelOf,
  noteUsed,
  recentCommands,
  visibleCommands,
} from './registry';

let host: HTMLElement | null = null;
let input: HTMLInputElement | null = null;
let list: HTMLElement | null = null;
let shown: CommandDef[] = [];
let active = 0;
let savedRange: Range | null = null;

export function isCommandsOpen(): boolean {
  return !!host;
}

/* ------------------------------------------------------------------ *
 * Matching
 * ------------------------------------------------------------------ */

/**
 * Subsequence match, scored.
 *
 * Not a fuzzy library: "eh" should find "Export HTML", and the cheapest rule
 * that does is "every letter in order". Consecutive letters and letters at
 * the start of a word score higher, so "exh" ranks Export HTML above
 * "Export... and then an h somewhere".
 */
function score(text: string, query: string): number {
  if (query === '') return 0;
  const hay = text.toLowerCase();
  let i = 0;
  let points = 0;
  let streak = 0;
  for (const ch of query.toLowerCase()) {
    if (ch === ' ') continue;
    const at = hay.indexOf(ch, i);
    if (at < 0) return -1;
    const boundary = at === 0 || /[\s—·(/]/.test(hay[at - 1]);
    streak = at === i ? streak + 1 : 0;
    points += 1 + streak * 2 + (boundary ? 3 : 0);
    i = at + 1;
  }
  // A short label that matched is a better answer than a long one.
  return points * 100 - text.length;
}

function filter(query: string): CommandDef[] {
  const all = visibleCommands();
  if (query.trim() === '') {
    // What you actually do, then everything else. An empty palette that
    // opens on whatever happened to be registered first is a list you have
    // to read; one that opens on your last eight is a list you recognise.
    const recent = recentCommands();
    const seen = new Set(recent.map((c) => c.id));
    return [...recent, ...all.filter((c) => !seen.has(c.id))];
  }
  const scored: { c: CommandDef; s: number }[] = [];
  for (const c of all) {
    const label = labelOf(c);
    const s = Math.max(
      score(label, query),
      score(c.category + ' ' + label, query) - 1,
      c.keywords ? score(c.keywords, query) - 2 : -1
    );
    if (s >= 0) scored.push({ c, s });
  }
  scored.sort((a, b) => b.s - a.s);
  return scored.map((x) => x.c);
}

/* ------------------------------------------------------------------ *
 * Rendering
 * ------------------------------------------------------------------ */

function row(c: CommandDef, i: number): HTMLElement {
  const el = document.createElement('div');
  const off = !isEnabled(c);
  el.className = 'cmd-row' + (i === active ? ' on' : '') + (off ? ' off' : '');
  el.dataset.index = String(i);
  el.setAttribute('role', 'option');
  el.setAttribute('aria-selected', String(i === active));
  if (off) el.setAttribute('aria-disabled', 'true');

  const tick = document.createElement('span');
  tick.className = 'cmd-tick';
  tick.textContent = c.checked?.() ? '✓' : '';
  el.appendChild(tick);

  const label = document.createElement('span');
  label.className = 'cmd-label';
  label.textContent = labelOf(c);
  el.appendChild(label);

  const group = document.createElement('span');
  group.className = 'cmd-group';
  group.textContent = c.category;
  el.appendChild(group);

  const hint = document.createElement('span');
  hint.className = 'cmd-hint';
  hint.textContent = formatShortcut(c.shortcut);
  el.appendChild(hint);

  el.addEventListener('mousedown', (e) => e.preventDefault());
  el.addEventListener('click', () => {
    active = i;
    accept();
  });
  el.addEventListener('mousemove', () => {
    if (active === i) return;
    active = i;
    draw();
  });
  return el;
}

function draw(): void {
  if (!list) return;
  list.textContent = '';
  if (shown.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'cmd-empty';
    empty.textContent = 'NO COMMAND MATCHES';
    list.appendChild(empty);
    return;
  }
  shown.forEach((c, i) => list?.appendChild(row(c, i)));
  const on = list.querySelector('.cmd-row.on') as HTMLElement | null;
  on?.scrollIntoView({ block: 'nearest' });
}

function refilter(): void {
  shown = filter(input?.value ?? '');
  active = 0;
  draw();
}

/* ------------------------------------------------------------------ *
 * Open, close, run
 * ------------------------------------------------------------------ */

function restoreSelection(): void {
  if (!savedRange) return;
  const sel = window.getSelection();
  sel?.removeAllRanges();
  sel?.addRange(savedRange);
}

function accept(): void {
  const chosen = shown[active];
  if (!chosen || !isEnabled(chosen)) return;
  closeCommands();
  // The caret has to be back before the command runs: Bold applied to a
  // selection that the palette's own text field stole is applied to nothing.
  restoreSelection();
  noteUsed(chosen.id);
  chosen.run();
  onRan?.();
}

/** Told after a command runs, so the recent list can be persisted. */
let onRan: (() => void) | null = null;

export function setCommandsHooks(ran: () => void): void {
  onRan = ran;
}

export function closeCommands(): void {
  host?.remove();
  host = null;
  input = null;
  list = null;
  shown = [];
}

export function openCommands(): void {
  if (host) {
    closeCommands();
    return; // a second press closes it, the way every palette does
  }
  const sel = window.getSelection();
  savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0).cloneRange() : null;

  host = document.createElement('div');
  host.className = 'cmd-scrim';
  host.addEventListener('mousedown', (e) => {
    if (e.target === host) {
      e.preventDefault();
      closeCommands();
      restoreSelection();
    }
  });

  const panel = document.createElement('div');
  panel.className = 'cmd-panel';

  const head = document.createElement('div');
  head.className = 'cmd-head';
  const mark = document.createElement('span');
  mark.className = 'cmd-mark';
  head.appendChild(mark);

  input = document.createElement('input');
  input.className = 'cmd-input';
  input.type = 'text';
  input.placeholder = 'Type a command';
  input.spellcheck = false;
  input.autocomplete = 'off';
  head.appendChild(input);
  panel.appendChild(head);

  list = document.createElement('div');
  list.className = 'cmd-list';
  panel.appendChild(list);

  host.appendChild(panel);
  document.body.appendChild(host);

  input.addEventListener('input', refilter);
  input.addEventListener('keydown', (e) => {
    switch (e.key) {
      case 'Escape':
        e.preventDefault();
        closeCommands();
        restoreSelection();
        break;
      case 'Enter':
        e.preventDefault();
        accept();
        break;
      case 'ArrowDown':
        e.preventDefault();
        active = shown.length ? (active + 1) % shown.length : 0;
        draw();
        break;
      case 'ArrowUp':
        e.preventDefault();
        active = shown.length ? (active - 1 + shown.length) % shown.length : 0;
        draw();
        break;
      case 'Home':
        if (input?.value === '') {
          e.preventDefault();
          active = 0;
          draw();
        }
        break;
      case 'End':
        if (input?.value === '') {
          e.preventDefault();
          active = Math.max(0, shown.length - 1);
          draw();
        }
        break;
    }
  });

  refilter();
  input.focus();
}
