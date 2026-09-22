import {
  command,
  formatShortcut,
  isEnabled,
  labelOf,
  runCommand,
} from './registry';
import { openMenuAt } from './ui';
import type { MenuItem } from './ui';
import { DOC_FONT, STYLES } from './styles';
import { STYLE_IDS } from './model';

/**
 * The format bar, pinned above the page.
 *
 * Exactly the page's width and aligned with it, because it acts on what is
 * on the page and a bar wider than the paper reads as part of the window
 * instead. Quieter than the page too: no boxed outlines until hover, so the
 * only thing in the room with weight is still the document.
 *
 * Every control here runs a registered command. Nothing in this file knows
 * what Bold does - it knows that `format.bold` exists, what it is called,
 * whether it is on, and what its shortcut is.
 */

let bar: HTMLElement | null = null;
let onChanged: (() => void) | null = null;

/** Lowest priority first: these are the ones the overflow eats. */
const PRIORITY = ['fb-leading', 'fb-size', 'fb-insert', 'fb-align'];

function run(id: string): void {
  runCommand(id);
  onChanged?.();
}

/** A menu row that runs a command, reading everything about it from one place. */
function itemFor(id: string, extra: Partial<MenuItem> = {}): MenuItem {
  const c = command(id);
  if (!c) return { label: id, disabled: true };
  return {
    label: labelOf(c),
    hint: formatShortcut(c.shortcut),
    checked: c.checked?.() ?? false,
    disabled: !isEnabled(c),
    onSelect: () => run(id),
    ...extra,
  };
}

function iconButton(
  id: string,
  cls: string,
  content: (b: HTMLElement) => void
): HTMLElement {
  const c = command(id);
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'fb-btn ' + cls;
  // The button remembers which command it runs, so a state refresh can
  // find it again without rebuilding the bar.
  b.dataset.cmd = id;
  const label = c ? labelOf(c) : id;
  const hint = formatShortcut(c?.shortcut);
  // "Bold · Ctrl+B", everywhere, rather than a shortcut printed on some
  // buttons and missing from others.
  b.title = hint ? label + ' · ' + hint : label;
  b.setAttribute('aria-label', b.title);
  b.setAttribute('aria-pressed', String(c?.checked?.() ?? false));
  if (c?.checked?.()) b.classList.add('on');
  if (c && !isEnabled(c)) b.disabled = true;
  content(b);
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', (e) => {
    e.preventDefault();
    run(id);
  });
  return b;
}

function menuButton(
  label: string,
  title: string,
  cls: string,
  rows: () => MenuItem[]
): HTMLElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'fb-btn ' + cls;
  b.title = title;
  b.setAttribute('aria-label', title);
  b.setAttribute('aria-haspopup', 'menu');
  const text = document.createElement('span');
  text.textContent = label;
  b.appendChild(text);
  const caret = document.createElement('span');
  caret.className = 'fb-caret';
  b.appendChild(caret);
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', () => openMenuAt(b, rows()));
  return b;
}

function divider(): HTMLElement {
  const d = document.createElement('span');
  d.className = 'fb-sep';
  return d;
}

/** Three bars laid out the way the paragraph would be. */
function alignIcon(which: string): (b: HTMLElement) => void {
  const widths = which === 'justify' ? [100, 100, 100] : [100, 62, 84];
  return (b) => {
    for (const w of widths) {
      const i = document.createElement('i');
      i.style.width = w + '%';
      b.appendChild(i);
    }
  };
}

/* ------------------------------------------------------------------ *
 * Drawing
 * ------------------------------------------------------------------ */

function currentStyleLabel(): string {
  for (const id of STYLE_IDS) {
    if (command('format.style.' + id)?.checked?.()) return STYLES[id].label;
  }
  return 'Mixed';
}

function currentSizeLabel(): string {
  const hit = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 24, 36].find((v) =>
    command('format.size.' + v)?.checked?.()
  );
  return hit === undefined ? 'Size' : hit + ' pt';
}

function stylePreview(id: (typeof STYLE_IDS)[number]): Partial<CSSStyleDeclaration> {
  const d = STYLES[id];
  return {
    fontFamily: DOC_FONT,
    fontSize: Math.min(19, Math.max(12, d.size * 1.2)) + 'px',
    fontWeight: d.bold ? '700' : '400',
    textTransform: d.uppercase ? 'uppercase' : 'none',
    letterSpacing: d.letterSpacing ? d.letterSpacing + 'px' : 'normal',
  };
}

function build(): HTMLElement {
  const el = document.createElement('div');
  el.className = 'formatbar';
  el.setAttribute('role', 'toolbar');
  el.setAttribute('aria-label', 'Formatting');
  el.addEventListener('mousedown', (e) => e.preventDefault());

  // A picker shows the value it holds, so this reads "Body" or "Job title".
  el.appendChild(
    menuButton(currentStyleLabel(), 'Paragraph style', 'fb-style', () =>
      STYLE_IDS.map((id) =>
        itemFor('format.style.' + id, { preview: stylePreview(id) })
      )
    )
  );
  el.appendChild(divider());

  el.appendChild(iconButton('format.bold', 'fb-b', (b) => (b.textContent = 'B')));
  el.appendChild(iconButton('format.italic', 'fb-i', (b) => (b.textContent = 'I')));
  el.appendChild(iconButton('format.underline', 'fb-u', (b) => (b.textContent = 'U')));
  el.appendChild(divider());

  const aligns = document.createElement('span');
  aligns.className = 'fb-group fb-align';
  for (const which of ['left', 'center', 'right', 'justify']) {
    aligns.appendChild(
      iconButton('format.align.' + which, 'fb-alignbtn fb-align-' + which, alignIcon(which))
    );
  }
  el.appendChild(aligns);

  el.appendChild(
    menuButton(currentSizeLabel(), 'Text size and colour', 'fb-size', () => [
      { heading: 'Size' },
      ...[8, 9, 10, 10.5, 11, 12, 14, 16, 18, 24, 36].map((v) =>
        itemFor('format.size.' + v)
      ),
      { separator: true },
      { heading: 'Colour' },
      ...['default', '000000', '595959', 'C00000', 'B45309', '2E6B33', '1F4E79', '5B2D8E'].map(
        (c) =>
          itemFor(
            'format.colour.' + c,
            c === 'default' ? {} : { swatch: '#' + c, swatchBg: '#' + c }
          )
      ),
    ])
  );

  el.appendChild(
    menuButton('Spacing', 'Line spacing', 'fb-leading', () => [
      { heading: 'Line spacing' },
      ...[1, 1.15, 1.5, 2].map((v) => itemFor('format.leading.' + v)),
      { separator: true },
      itemFor('format.clearParagraph'),
    ])
  );

  el.appendChild(divider());
  el.appendChild(
    menuButton('Insert', 'Insert', 'fb-insert', () => [
      { heading: 'Table' },
      ...[
        [2, 2],
        [3, 2],
        [3, 3],
        [4, 3],
        [5, 2],
        [6, 4],
      ].map(([r, c]) => itemFor(`insert.table.${r}x${c}`, { label: `${r} × ${c}` })),
      { separator: true },
      itemFor('insert.image'),
      itemFor('insert.link'),
      itemFor('insert.pageBreak'),
    ])
  );

  const spacer = document.createElement('span');
  spacer.className = 'fb-spacer';
  el.appendChild(spacer);

  el.appendChild(
    iconButton('app.palette', 'fb-more', (b) => (b.textContent = '⋯'))
  );
  return el;
}

/**
 * Move low-priority groups into the overflow until the bar fits the page.
 *
 * The bar is exactly the page's width by definition, so "too wide" means
 * the contents, not the bar. Nothing is hidden without somewhere to go:
 * what leaves is always still in the palette behind the same button.
 */
function applyOverflow(): void {
  if (!bar) return;
  for (const cls of PRIORITY) {
    bar.querySelector('.' + cls)?.classList.remove('fb-hidden');
  }
  for (const cls of PRIORITY) {
    if (bar.scrollWidth <= bar.clientWidth) break;
    bar.querySelector('.' + cls)?.classList.add('fb-hidden');
  }
}

export function renderFormatBar(): void {
  const hostEl = document.getElementById('formatbar');
  if (!hostEl) return;
  hostEl.textContent = '';
  bar = build();
  hostEl.appendChild(bar);
  applyOverflow();
}

export function bindFormatBar(changed: () => void): void {
  onChanged = changed;
  window.addEventListener('resize', applyOverflow);
}

/** Re-read every control's state without rebuilding the bar. */
export function syncFormatBar(): void {
  if (!bar) return;
  for (const b of Array.from(bar.querySelectorAll('button[data-cmd]'))) {
    const el = b as HTMLButtonElement;
    const c = command(el.dataset.cmd as string);
    if (!c) continue;
    const on = c.checked?.() ?? false;
    el.classList.toggle('on', on);
    el.setAttribute('aria-pressed', String(on));
    el.disabled = !isEnabled(c);
  }
  const style = bar.querySelector('.fb-style span') as HTMLElement | null;
  if (style) style.textContent = currentStyleLabel();
  const size = bar.querySelector('.fb-size span') as HTMLElement | null;
  if (size) size.textContent = currentSizeLabel();
}
