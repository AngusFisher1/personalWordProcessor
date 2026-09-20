/**
 * The export sheet.
 *
 * A menu of file formats tells you what you can produce. It does not tell
 * you what each one costs, and every format here costs something different:
 * one comes back byte for byte, one keeps the pictures, one keeps nothing
 * but the words and is the only one still readable in fifty years.
 *
 * Saying so at the moment of choosing is the whole point of the sheet. It
 * is the same information the README carries, put where the decision is
 * actually made.
 */

export interface ExportOption {
  label: string;
  /** One line on what this format keeps, and what it does not. */
  note: string;
  /** Right-aligned extension or shortcut. */
  tag?: string;
  /** Drawn quieter and below the rule. */
  secondary?: boolean;
  run(): void;
}

let host: HTMLElement | null = null;

export function isExportSheetOpen(): boolean {
  return !!host;
}

export function closeExportSheet(): void {
  host?.remove();
  host = null;
}

function rowFor(o: ExportOption): HTMLElement {
  const row = document.createElement('button');
  row.className = 'xs-row' + (o.secondary ? ' quiet' : '');
  row.addEventListener('mousedown', (e) => e.preventDefault());
  row.addEventListener('click', () => {
    closeExportSheet();
    o.run();
  });

  const main = document.createElement('div');
  main.className = 'xs-main';

  const top = document.createElement('div');
  top.className = 'xs-top';
  const label = document.createElement('span');
  label.className = 'xs-label';
  label.textContent = o.label;
  top.appendChild(label);
  if (o.tag) {
    const tag = document.createElement('span');
    tag.className = 'xs-tag';
    tag.textContent = o.tag;
    top.appendChild(tag);
  }
  main.appendChild(top);

  const note = document.createElement('div');
  note.className = 'xs-note';
  note.textContent = o.note;
  main.appendChild(note);

  row.appendChild(main);
  return row;
}

export function openExportSheet(options: ExportOption[], title: string): void {
  closeExportSheet();

  host = document.createElement('div');
  host.className = 'xs-scrim';
  host.addEventListener('mousedown', (e) => {
    if (e.target === host) {
      e.preventDefault();
      closeExportSheet();
    }
  });

  const panel = document.createElement('div');
  panel.className = 'xs-panel';
  panel.addEventListener('mousedown', (e) => e.preventDefault());

  const head = document.createElement('div');
  head.className = 'xs-head';
  const h = document.createElement('div');
  h.className = 'xs-title';
  h.textContent = 'EXPORT';
  const sub = document.createElement('div');
  sub.className = 'xs-sub';
  sub.textContent = title;
  head.append(h, sub);
  panel.appendChild(head);

  let ruled = false;
  for (const o of options) {
    if (o.secondary && !ruled) {
      ruled = true;
      const sep = document.createElement('div');
      sep.className = 'xs-sep';
      panel.appendChild(sep);
    }
    panel.appendChild(rowFor(o));
  }

  host.appendChild(panel);
  document.body.appendChild(host);

  // Esc closes, like every other overlay. Bound on the document because the
  // sheet takes no focus - nothing in it is a text field.
  const onKey = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape') return;
    e.preventDefault();
    closeExportSheet();
    document.removeEventListener('keydown', onKey, true);
  };
  document.addEventListener('keydown', onKey, true);
}
