/**
 * Recto palettes.
 *
 * The workspace is a darkroom: the only light in it is the page. Every
 * surface, rule and label is derived from ten tokens, so a palette is a data
 * change rather than a stylesheet fork, and the page itself stays white in
 * all of them - it is paper.
 */

export interface Palette {
  id: string;
  /** Shown in the picker. */
  label: string;
  /** One-line description of the colour idea. */
  note: string;
  dark: boolean;
  /** Workspace behind the pages. */
  bg: string;
  /** The rail. */
  rail: string;
  /** Floating surfaces: menus, the format bar, the command palette. */
  panel: string;
  /** Hairlines and borders. */
  line: string;
  /** Quietest text: metadata, shortcut hints. */
  dim: string;
  /** Secondary text. */
  mid: string;
  /** Primary text. */
  hi: string;
  /** Accent. */
  acc: string;
  /** Text drawn on top of the accent. */
  accfg: string;
  /** Page drop shadow, retuned per palette. */
  pgsh: string;
}

const DARK_SHADOW = '0 1px 2px rgba(0,0,0,.6), 0 34px 80px -30px rgba(0,0,0,.9)';

export const PALETTES: Palette[] = [
  {
    id: 'night', label: 'Night', note: 'amber on near-black', dark: true,
    bg: '#0A0A0B', rail: '#0D0D0F', panel: '#121215', line: '#23232A',
    dim: '#7C7C84', mid: '#8D8D98', hi: '#EAEAEC', acc: '#D97B3C',
    accfg: '#15100C', pgsh: DARK_SHADOW,
  },
  {
    id: 'oxide', label: 'Oxide', note: 'iron red on warm black', dark: true,
    bg: '#0C0A09', rail: '#100D0C', panel: '#16120F', line: '#2A211D',
    dim: '#857B74', mid: '#9A8B81', hi: '#EEE9E4', acc: '#C65E48',
    accfg: '#190E0B', pgsh: DARK_SHADOW,
  },
  {
    id: 'cobalt', label: 'Cobalt', note: 'signal blue on neutral black', dark: true,
    bg: '#09090A', rail: '#0C0C0E', panel: '#111113', line: '#212127',
    dim: '#7C7C84', mid: '#8A8A95', hi: '#E9E9EC', acc: '#5375E3',
    accfg: '#111112', pgsh: DARK_SHADOW,
  },
  {
    id: 'sulphur', label: 'Sulphur', note: 'brass on olive black', dark: true,
    bg: '#0A0A09', rail: '#0D0D0B', panel: '#131310', line: '#26251E',
    dim: '#7F7E72', mid: '#8F8E7F', hi: '#EBEAE3', acc: '#C9A227',
    accfg: '#14110A', pgsh: DARK_SHADOW,
  },
  {
    id: 'moss', label: 'Moss', note: 'sage on green black', dark: true,
    bg: '#080A09', rail: '#0B0D0C', panel: '#101312', line: '#1F2622',
    dim: '#757F78', mid: '#85928A', hi: '#E6EAE7', acc: '#6E8F4E',
    accfg: '#0A0F07', pgsh: DARK_SHADOW,
  },
  {
    id: 'plum', label: 'Plum', note: 'dusty rose on violet black', dark: true,
    bg: '#0A0809', rail: '#0D0B0D', panel: '#131013', line: '#271F26',
    dim: '#837983', mid: '#948994', hi: '#EBE7EA', acc: '#B3637F',
    accfg: '#111111', pgsh: DARK_SHADOW,
  },
  {
    id: 'noir', label: 'Noir', note: 'white on pure black', dark: true,
    bg: '#000000', rail: '#040404', panel: '#0B0B0B', line: '#1E1E1E',
    dim: '#797979', mid: '#909090', hi: '#FFFFFF', acc: '#F2F2F2',
    accfg: '#000000',
    pgsh: '0 1px 2px rgba(0,0,0,.9), 0 34px 80px -30px rgba(0,0,0,1)',
  },
  {
    id: 'day', label: 'Daylight', note: 'burnt orange on warm grey', dark: false,
    bg: '#DBD7D0', rail: '#D3CEC6', panel: '#EDEAE4', line: '#C1BBB1',
    dim: '#5D5851', mid: '#565149', hi: '#191713', acc: '#85491F',
    accfg: '#FFF8F2',
    pgsh: '0 1px 2px rgba(35,30,20,.20), 0 30px 62px -28px rgba(35,30,20,.55)',
  },
  {
    id: 'mossStone', label: 'Moss on stone', note: 'sage on warm light grey', dark: false,
    bg: '#DBD8CF', rail: '#D3CFC5', panel: '#EDEBE4', line: '#C0BCB0',
    dim: '#5C5950', mid: '#55524A', hi: '#181713', acc: '#486032',
    accfg: '#F4F7EF',
    pgsh: '0 1px 2px rgba(30,32,22,.18), 0 30px 62px -28px rgba(30,32,22,.5)',
  },
  {
    id: 'mossChalk', label: 'Moss on chalk', note: 'sage on cool light grey', dark: false,
    bg: '#D6D9D6', rail: '#CED2CE', panel: '#EBEEEA', line: '#B5BAB5',
    dim: '#575B56', mid: '#4D534C', hi: '#111412', acc: '#486132',
    accfg: '#F4F7EF',
    pgsh: '0 1px 2px rgba(18,26,18,.18), 0 30px 62px -28px rgba(18,26,18,.5)',
  },
  {
    id: 'bone', label: 'Bone', note: 'ink blue on parchment', dark: false,
    bg: '#E8E4DA', rail: '#E0DBCF', panel: '#F5F2EB', line: '#C8C2B3',
    dim: '#666056', mid: '#5A5449', hi: '#161410', acc: '#2F4B8C',
    accfg: '#F4F6FF',
    pgsh: '0 1px 2px rgba(35,30,20,.18), 0 30px 62px -28px rgba(35,30,20,.5)',
  },
  {
    id: 'chalk', label: 'Chalk', note: 'oxblood on cool grey', dark: false,
    bg: '#D6D8DA', rail: '#CED1D4', panel: '#EBEDEF', line: '#B6BABE',
    dim: '#575A5E', mid: '#4E5357', hi: '#121416', acc: '#8C3A2E',
    accfg: '#FFF4F1',
    pgsh: '0 1px 2px rgba(20,25,30,.18), 0 30px 62px -28px rgba(20,25,30,.5)',
  },
];

const STORAGE_KEY = 'wp:palette';
const DEFAULT_ID = 'night';

export function paletteById(id: string): Palette {
  return PALETTES.find((p) => p.id === id) ?? PALETTES[0];
}

export function currentPaletteId(): string {
  try {
    return localStorage.getItem(STORAGE_KEY) ?? DEFAULT_ID;
  } catch {
    return DEFAULT_ID;
  }
}

/** Write a palette onto the root element as custom properties. */
export function applyPalette(id: string): Palette {
  const p = paletteById(id);
  const s = document.documentElement.style;
  s.setProperty('--bg', p.bg);
  s.setProperty('--rail', p.rail);
  s.setProperty('--panel', p.panel);
  s.setProperty('--line', p.line);
  s.setProperty('--dim', p.dim);
  s.setProperty('--mid', p.mid);
  s.setProperty('--hi', p.hi);
  s.setProperty('--acc', p.acc);
  s.setProperty('--accfg', p.accfg);
  s.setProperty('--pgsh', p.pgsh);
  document.documentElement.dataset.palette = p.id;
  document.documentElement.dataset.paletteDark = String(p.dark);
  try {
    localStorage.setItem(STORAGE_KEY, p.id);
  } catch {
    /* a palette that cannot be remembered is not worth failing over */
  }
  return p;
}
