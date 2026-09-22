/**
 * What the interface looked like when you left it.
 *
 * Kept apart from the documents, under its own key, because it is not part
 * of any of them: a collapsed panel follows the person, not the file. It
 * uses the same localStorage the palette choice already used, with the same
 * `wp:` prefix, so there is one storage convention rather than two.
 *
 * Every read is defensive. This is the first thing that runs at boot, and a
 * half-written value here must not be the reason the editor fails to open.
 */

const KEY = 'wp:ui';

export type NavTab = 'outline' | 'files';
export type InspectorSection = 'page' | 'headerFooter' | 'print';

export interface UiState {
  navCollapsed: boolean;
  inspectorOpen: boolean;
  inspectorSection: InspectorSection;
  navTab: NavTab;
  /** Command ids, most recent first. */
  recent: string[];
}

const DEFAULTS: UiState = {
  navCollapsed: false,
  inspectorOpen: false,
  inspectorSection: 'page',
  navTab: 'outline',
  recent: [],
};

let state: UiState = { ...DEFAULTS };

const SECTIONS: InspectorSection[] = ['page', 'headerFooter', 'print'];

export function loadUiState(): UiState {
  try {
    const raw = localStorage.getItem(KEY);
    const o = raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
    state = {
      navCollapsed: o.navCollapsed === true,
      inspectorOpen: o.inspectorOpen === true,
      inspectorSection: SECTIONS.includes(o.inspectorSection as InspectorSection)
        ? (o.inspectorSection as InspectorSection)
        : DEFAULTS.inspectorSection,
      navTab: o.navTab === 'files' ? 'files' : 'outline',
      recent: Array.isArray(o.recent)
        ? o.recent.filter((x): x is string => typeof x === 'string').slice(0, 8)
        : [],
    };
  } catch {
    state = { ...DEFAULTS };
  }
  return state;
}

export function uiState(): UiState {
  return state;
}

/**
 * Change part of the state and write it.
 *
 * Failure is silent: a full or blocked localStorage should cost you a
 * remembered panel, never the panel itself.
 */
export function setUiState(patch: Partial<UiState>): UiState {
  state = { ...state, ...patch };
  try {
    localStorage.setItem(KEY, JSON.stringify(state));
  } catch {
    /* the panel still works, it just will not be remembered */
  }
  return state;
}
