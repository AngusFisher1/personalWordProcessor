import type { Doc, MarginKey, PageSetup } from './model';
import { MARGINS, PAGE_H, PAGE_W, marginPreset } from './model';
import { STYLES } from './styles';
import type { InspectorSection } from './uistate';
import { setUiState, uiState } from './uistate';

/**
 * Document settings, on the right, hidden by default.
 *
 * Everything here describes the DOCUMENT rather than the selection: the
 * paper, the margins, the body face, the running head. That is the line
 * that decides what belongs - if it changes with the cursor it goes in the
 * format bar, and if it changes with the file it goes here.
 *
 * This replaces the margins dropdown and the read-only page setup readout
 * that used to sit at the bottom of the left panel. One place, editable in
 * place, no second copy to disagree with.
 */

export interface InspectorHost {
  doc(): Doc;
  setPaper(width: number, height: number): void;
  setMargins(m: { top: number; right: number; bottom: number; left: number }): void;
  setMarginPreset(key: MarginKey): void;
  setDefaultFont(family: string | null): void;
  setBodySize(pt: number): void;
  setBodyLeading(mult: number): void;
  /** w:titlePg and w:evenAndOddHeaders. */
  setTitlePage(on: boolean): void;
  setEvenOdd(on: boolean): void;
  editHeaderFooter(): void;
  print(): void;
}

let host: InspectorHost | null = null;

export function bindInspector(h: InspectorHost): void {
  host = h;
}

/* ------------------------------------------------------------------ *
 * Fields
 * ------------------------------------------------------------------ */

function el(tag: string, cls?: string, text?: string): HTMLElement {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
}

function section(title: string, id: InspectorSection): HTMLElement {
  const s = el('section', 'insp-section');
  s.dataset.section = id;
  s.appendChild(el('h2', 'insp-title', title));
  return s;
}

function field(label: string, control: HTMLElement): HTMLElement {
  const row = el('div', 'insp-field');
  const name = el('label', 'insp-label', label);
  const id = 'insp-' + label.toLowerCase().replace(/[^a-z]+/g, '-');
  control.id = id;
  name.setAttribute('for', id);
  row.append(name, control);
  return row;
}

/**
 * A segmented control.
 *
 * The presets a document actually uses, as one row of choices rather than a
 * dropdown: there are three of them and the current one should be visible
 * without opening anything.
 */
function segmented(
  options: { label: string; value: string }[],
  current: string,
  onPick: (value: string) => void
): HTMLElement {
  const wrap = el('div', 'insp-seg');
  wrap.setAttribute('role', 'radiogroup');
  for (const o of options) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'insp-seg-btn' + (o.value === current ? ' on' : '');
    b.textContent = o.label;
    b.setAttribute('role', 'radio');
    b.setAttribute('aria-checked', String(o.value === current));
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => onPick(o.value));
    wrap.appendChild(b);
  }
  return wrap;
}

/** Inches, to two places, committed on Enter or blur. */
function inches(value: number, onSet: (px: number) => void): HTMLInputElement {
  const i = document.createElement('input');
  i.type = 'number';
  i.className = 'insp-num';
  i.step = '0.05';
  i.min = '0';
  i.max = '4';
  i.value = (value / 96).toFixed(2);
  const commit = (): void => {
    const n = Number(i.value);
    if (!Number.isFinite(n) || n < 0 || n > 4) {
      i.value = (value / 96).toFixed(2);
      return;
    }
    onSet(Math.round(n * 96));
  };
  i.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commit();
    }
  });
  i.addEventListener('blur', commit);
  return i;
}

function toggle(label: string, on: boolean, onSet: (v: boolean) => void): HTMLElement {
  const row = el('div', 'insp-field');
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'insp-toggle' + (on ? ' on' : '');
  b.setAttribute('role', 'switch');
  b.setAttribute('aria-checked', String(on));
  b.textContent = on ? 'On' : 'Off';
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', () => onSet(!on));
  row.append(el('label', 'insp-label', label), b);
  return row;
}

function action(label: string, hint: string, onClick: () => void): HTMLElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'insp-action';
  b.textContent = label;
  b.title = hint || label;
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', onClick);
  return b;
}

/* ------------------------------------------------------------------ *
 * Sections
 * ------------------------------------------------------------------ */

const PAPERS = [
  { label: 'Letter', value: 'letter', w: PAGE_W, h: PAGE_H },
  { label: 'Legal', value: 'legal', w: 816, h: 1344 },
  { label: 'A4', value: 'a4', w: 794, h: 1123 },
];

function paperOf(p: PageSetup): string {
  const hit = PAPERS.find((x) => Math.abs(x.w - p.width) < 3 && Math.abs(x.h - p.height) < 3);
  return hit?.value ?? 'custom';
}

function pageSection(doc: Doc): HTMLElement {
  const s = section('Page setup', 'page');
  const p = doc.page;

  s.appendChild(
    field(
      'Paper',
      segmented(
        [...PAPERS.map((x) => ({ label: x.label, value: x.value })), { label: 'Custom', value: 'custom' }],
        paperOf(p),
        (v) => {
          const hit = PAPERS.find((x) => x.value === v);
          if (hit) host?.setPaper(hit.w, hit.h);
        }
      )
    )
  );

  s.appendChild(
    field(
      'Margins',
      segmented(
        [
          ...(Object.keys(MARGINS) as MarginKey[]).map((k) => ({
            label: k === 'narrow' ? 'Narrow' : 'Normal',
            value: k as string,
          })),
          { label: 'Custom', value: 'custom' },
        ],
        marginPreset(p) ?? 'custom',
        (v) => {
          if (v !== 'custom') host?.setMarginPreset(v as MarginKey);
        }
      )
    )
  );

  // The four edges, editable. A preset is a shortcut to a set of these,
  // not a different kind of thing.
  const edges = el('div', 'insp-edges');
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const cell = el('div', 'insp-edge');
    cell.appendChild(el('span', 'insp-edge-label', side[0].toUpperCase() + side.slice(1)));
    cell.appendChild(
      inches(p.margins[side], (px) => host?.setMargins({ ...p.margins, [side]: px }))
    );
    edges.appendChild(cell);
  }
  s.appendChild(edges);

  const font = document.createElement('input');
  font.type = 'text';
  font.className = 'insp-text';
  font.value = doc.defaultFont ?? '';
  font.placeholder = 'Source Serif 4';
  font.spellcheck = false;
  const commitFont = (): void => host?.setDefaultFont(font.value.trim() || null);
  font.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitFont();
    }
  });
  font.addEventListener('blur', commitFont);
  s.appendChild(field('Body font', font));

  const body = STYLES.Body;
  s.appendChild(
    field(
      'Body size',
      segmented(
        [9, 10, 10.5, 11, 12].map((v) => ({ label: String(v), value: String(v) })),
        String(body.size),
        (v) => host?.setBodySize(Number(v))
      )
    )
  );
  s.appendChild(
    field(
      'Leading',
      segmented(
        [1.2, 1.35, 1.45, 1.6].map((v) => ({ label: v.toFixed(2), value: String(v) })),
        String(Math.round(body.lineHeight * 100) / 100),
        (v) => host?.setBodyLeading(Number(v))
      )
    )
  );
  return s;
}

function headerSection(doc: Doc): HTMLElement {
  const s = section('Header & footer', 'headerFooter');
  s.appendChild(
    toggle('First page differs', doc.titlePage === true, (v) => host?.setTitlePage(v))
  );
  s.appendChild(
    toggle('Odd and even differ', doc.evenOdd === true, (v) => host?.setEvenOdd(v))
  );
  // Editing happens ON the page - the body dims and the slots become
  // editable - so this opens that mode rather than pretending a panel can
  // hold a thing that is typed into the paper.
  s.appendChild(
    action('Edit header & footer', 'The body dims; Esc returns', () =>
      host?.editHeaderFooter()
    )
  );
  return s;
}

function printSection(): HTMLElement {
  const s = section('Print', 'print');
  s.appendChild(
    el(
      'p',
      'insp-note',
      'In the browser’s print dialog set Margins to None and turn page ' +
        'headers off. With those set, the output matches the screen exactly.'
    )
  );
  s.appendChild(action('Print / save as PDF', 'Ctrl+P', () => host?.print()));
  return s;
}

/* ------------------------------------------------------------------ *
 * Drawing
 * ------------------------------------------------------------------ */

export function renderInspector(): void {
  const root = document.getElementById('inspector');
  if (!root || !host) return;
  const doc = host.doc();
  root.textContent = '';

  const head = el('div', 'insp-head');
  head.appendChild(el('span', 'insp-heading', 'DOCUMENT'));
  root.appendChild(head);

  root.appendChild(pageSection(doc));
  root.appendChild(headerSection(doc));
  root.appendChild(printSection());

  // Opened from the status bar or from a command that names a section:
  // scroll to it rather than making the reader hunt.
  const want = uiState().inspectorSection;
  const target = root.querySelector(`[data-section="${want}"]`);
  if (want !== 'page') target?.scrollIntoView({ block: 'start' });
  setUiState({ inspectorSection: want });
}
