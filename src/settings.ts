import { PALETTES } from './theme';
import { STYLES, shippedStyle } from './styles';
import { STYLE_IDS } from './model';
import type { StyleId } from './model';
import { openMenuAt } from './ui';
import type { MenuItem } from './ui';

/**
 * Settings: the things that belong to the program rather than to a document.
 *
 * A sheet rather than a panel, because none of it is consulted while you
 * write. The theme is chosen once a season, the named styles once a
 * document, and the storage line is there to be reassuring rather than
 * useful - it is a feature worth being reminded of, just not on every
 * screen.
 */

export interface SettingsHost {
  currentPaletteId(): string;
  setPalette(id: string): void;
  /** Rows for one named style, built by the styles editor. */
  styleRows(id: StyleId): MenuItem[];
  resetStyles(): void;
  stylesChanged(): boolean;
  storage(): { documents: number; bytes: number };
}

let host: SettingsHost | null = null;
let sheet: HTMLElement | null = null;

export function bindSettings(h: SettingsHost): void {
  host = h;
}

export function isSettingsOpen(): boolean {
  return !!sheet;
}

export function closeSettings(): void {
  sheet?.remove();
  sheet = null;
}

function el(tag: string, cls?: string, text?: string): HTMLElement {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
}

function section(title: string): HTMLElement {
  const s = el('section', 'set-section');
  s.appendChild(el('h2', 'set-title', title));
  return s;
}

function themeSection(): HTMLElement {
  const s = section('Theme');
  const grid = el('div', 'set-themes');
  const now = host?.currentPaletteId();
  for (const p of PALETTES) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'set-theme' + (p.id === now ? ' on' : '');
    b.setAttribute('aria-pressed', String(p.id === now));
    b.title = p.label + ' · ' + p.note;
    // Each swatch is painted in its own palette, so the list shows what
    // the choice looks like rather than only what it is called.
    b.style.background = p.bg;
    b.style.borderColor = p.line;
    const dot = el('span', 'set-theme-dot');
    dot.style.background = p.acc;
    const name = el('span', 'set-theme-name', p.label);
    name.style.color = p.hi;
    b.append(dot, name);
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => {
      host?.setPalette(p.id);
      redraw();
    });
    grid.appendChild(b);
  }
  s.appendChild(grid);
  return s;
}

function stylesSection(): HTMLElement {
  const s = section('Named styles');
  s.appendChild(
    el(
      'p',
      'set-note',
      'The six styles this document is written in. Changes belong to this ' +
        'document alone — a resume set in 10pt and a report set in 12pt ' +
        'are both right.'
    )
  );
  const list = el('div', 'set-styles');
  for (const id of STYLE_IDS) {
    const d = STYLES[id];
    const changed = JSON.stringify(d) !== JSON.stringify(shippedStyle(id));
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'set-style';
    b.setAttribute('aria-haspopup', 'menu');
    const name = el('span', 'set-style-name', d.label);
    name.style.fontWeight = d.bold ? '700' : '400';
    name.style.textTransform = d.uppercase ? 'uppercase' : 'none';
    const size = el('span', 'set-style-size', d.size + ' pt');
    b.append(name, size);
    if (changed) b.appendChild(el('span', 'set-style-dot', '·'));
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => {
      if (host) openMenuAt(b, host.styleRows(id));
    });
    list.appendChild(b);
  }
  s.appendChild(list);

  const reset = document.createElement('button');
  reset.type = 'button';
  reset.className = 'set-action';
  reset.textContent = 'Reset every style';
  reset.disabled = !host?.stylesChanged();
  reset.addEventListener('mousedown', (e) => e.preventDefault());
  reset.addEventListener('click', () => {
    host?.resetStyles();
    redraw();
  });
  s.appendChild(reset);
  return s;
}

function storageSection(): HTMLElement {
  const s = section('Storage');
  const info = host?.storage() ?? { documents: 0, bytes: 0 };
  const kb = Math.max(1, Math.round(info.bytes / 1024));
  s.appendChild(
    el(
      'div',
      'set-storage',
      `On disk · ${info.documents} document${info.documents === 1 ? '' : 's'} ` +
        `· ${kb} KB · No account · No sync`
    )
  );
  return s;
}

function redraw(): void {
  if (!sheet) return;
  const panel = sheet.querySelector('.set-panel');
  if (!panel) return;
  panel.textContent = '';
  panel.appendChild(head());
  panel.appendChild(themeSection());
  panel.appendChild(stylesSection());
  panel.appendChild(storageSection());
}

function head(): HTMLElement {
  const h = el('div', 'set-head');
  h.append(el('div', 'set-heading', 'SETTINGS'));
  return h;
}

export function openSettingsSheet(): void {
  closeSettings();
  sheet = el('div', 'set-scrim');
  sheet.addEventListener('mousedown', (e) => {
    if (e.target === sheet) {
      e.preventDefault();
      closeSettings();
    }
  });
  const panel = el('div', 'set-panel');
  panel.addEventListener('mousedown', (e) => e.preventDefault());
  sheet.appendChild(panel);
  document.body.appendChild(sheet);
  redraw();

  const onKey = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape') return;
    e.preventDefault();
    closeSettings();
    document.removeEventListener('keydown', onKey, true);
  };
  document.addEventListener('keydown', onKey, true);
}
