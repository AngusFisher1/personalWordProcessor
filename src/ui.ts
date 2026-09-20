/**
 * Small toolbar widgets: buttons and dropdown menus.
 *
 * The one rule every control here obeys is to preventDefault on mousedown.
 * A toolbar that takes focus destroys the document selection, and then Bold
 * has nothing to apply itself to. Nothing in here ever focuses itself, so the
 * caret stays exactly where the user left it.
 */

export interface MenuItem {
  label?: string;
  /** Right-aligned shortcut or note. */
  hint?: string;
  checked?: boolean;
  disabled?: boolean;
  separator?: true;
  /** A small uppercase group heading; not selectable. */
  heading?: string;
  /** Inline styles for the label, used to preview a paragraph style. */
  preview?: Partial<CSSStyleDeclaration>;
  onSelect?: () => void;
}

function noFocus(el: HTMLElement): void {
  el.addEventListener('mousedown', (e) => e.preventDefault());
}

export function textButton(
  label: string,
  title: string,
  onClick: () => void,
  cls = ''
): HTMLButtonElement {
  const b = document.createElement('button');
  b.className = ('tb-btn ' + cls).trim();
  b.textContent = label;
  b.title = title;
  noFocus(b);
  b.addEventListener('click', onClick);
  return b;
}

const ICONS: Record<string, string> = {
  undo: '<path d="M4.5 9.5h8a4 4 0 0 1 0 8H9"/><path d="M7.5 5.5l-3 4 3 4"/>',
  redo: '<path d="M15.5 9.5h-8a4 4 0 0 0 0 8H11"/><path d="M12.5 5.5l3 4-3 4"/>',
  print:
    '<path d="M6 8V3.5h8V8"/><rect x="3.5" y="8" width="13" height="6" rx="1"/>' +
    '<path d="M6 12.5h8v4H6z"/>',
};

export function iconButton(
  name: keyof typeof ICONS | string,
  title: string,
  onClick: () => void
): HTMLButtonElement {
  const b = document.createElement('button');
  b.className = 'tb-btn tb-icon';
  b.title = title;
  b.innerHTML =
    '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" ' +
    'stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">' +
    (ICONS[name] ?? '') +
    '</svg>';
  b.setAttribute('aria-label', title);
  noFocus(b);
  b.addEventListener('click', onClick);
  return b;
}

export function separator(): HTMLSpanElement {
  const s = document.createElement('span');
  s.className = 'tb-sep';
  return s;
}

/* ------------------------------------------------------------------ *
 * Dropdown menus
 * ------------------------------------------------------------------ */

let closeOpenMenu: (() => void) | null = null;

export function closeMenu(): void {
  if (closeOpenMenu) closeOpenMenu();
}

function openPanel(anchor: HTMLElement, items: MenuItem[]): void {
  const panel = document.createElement('div');
  panel.className = 'menu';
  noFocus(panel);

  for (const it of items) {
    if (it.separator) {
      const s = document.createElement('div');
      s.className = 'menu-sep';
      panel.appendChild(s);
      continue;
    }
    if (it.heading) {
      const h = document.createElement('div');
      h.className = 'menu-heading';
      h.textContent = it.heading;
      panel.appendChild(h);
      continue;
    }
    const row = document.createElement('button');
    row.className = 'menu-item';
    row.disabled = !!it.disabled;
    noFocus(row);

    const tick = document.createElement('span');
    tick.className = 'menu-tick';
    tick.textContent = it.checked ? '✓' : '';
    row.appendChild(tick);

    const label = document.createElement('span');
    label.className = 'menu-label';
    label.textContent = it.label ?? '';
    if (it.preview) Object.assign(label.style, it.preview);
    row.appendChild(label);

    if (it.hint) {
      const hint = document.createElement('span');
      hint.className = 'menu-hint';
      hint.textContent = it.hint;
      row.appendChild(hint);
    }

    row.addEventListener('click', () => {
      closeMenu();
      it.onSelect?.();
    });
    panel.appendChild(row);
  }

  document.body.appendChild(panel);

  // Positioned after insertion so the measured width can keep it on screen.
  const r = anchor.getBoundingClientRect();
  const w = panel.offsetWidth;
  panel.style.left = Math.max(8, Math.min(r.left, window.innerWidth - w - 8)) + 'px';
  panel.style.top = r.bottom + 4 + 'px';
  anchor.classList.add('open');

  const close = () => {
    panel.remove();
    anchor.classList.remove('open');
    document.removeEventListener('mousedown', onDown, true);
    document.removeEventListener('keydown', onKey, true);
    window.removeEventListener('resize', close);
    closeOpenMenu = null;
  };
  const onDown = (e: MouseEvent) => {
    if (!panel.contains(e.target as Node) && !anchor.contains(e.target as Node)) {
      close();
    }
  };
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  };
  document.addEventListener('mousedown', onDown, true);
  document.addEventListener('keydown', onKey, true);
  window.addEventListener('resize', close);
  closeOpenMenu = close;
}

export interface MenuButton {
  el: HTMLButtonElement;
  setLabel(text: string): void;
}

/** A button that opens a menu built fresh on each click, so it reflects state. */
export function menuButton(
  label: string,
  title: string,
  build: () => MenuItem[],
  cls = ''
): MenuButton {
  const b = document.createElement('button');
  b.className = ('tb-btn tb-menubtn ' + cls).trim();
  b.title = title;

  const text = document.createElement('span');
  text.className = 'tb-menubtn-label';
  text.textContent = label;
  const caret = document.createElement('span');
  caret.className = 'tb-caret';
  caret.textContent = '▾';
  b.append(text, caret);

  noFocus(b);
  b.addEventListener('click', () => {
    const wasOpen = b.classList.contains('open');
    closeMenu();
    if (!wasOpen) openPanel(b, build());
  });

  return {
    el: b,
    setLabel(t: string) {
      text.textContent = t;
    },
  };
}
