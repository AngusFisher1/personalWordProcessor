/**
 * Panel state.
 *
 * Small, but it is the thing that decides what the window looks like when
 * you open it, and it reads from storage that a person or another tab can
 * have written anything into. Every field is validated; a half-written
 * value must cost a remembered panel, never the editor.
 *
 *   npm run test:uistate
 */
const store = new Map();
globalThis.localStorage = {
  getItem: (k) => (store.has(k) ? store.get(k) : null),
  setItem: (k, v) => store.set(k, String(v)),
  removeItem: (k) => store.delete(k),
};

const { loadUiState, uiState, setUiState } = await import('./build/harness.mjs');

let failures = 0;
function check(name, ok, detail = '') {
  if (ok) { console.log('  ok    ' + name); return; }
  failures++; console.log('  FAIL  ' + name + (detail ? ' - ' + detail : ''));
}

console.log('\nan empty store gives the defaults');
{
  store.clear();
  const s = loadUiState();
  check('the nav is open', s.navCollapsed === false);
  check('the inspector is closed', s.inspectorOpen === false);
  check('the outline is the tab', s.navTab === 'outline');
  check('page is the section', s.inspectorSection === 'page');
  check('nothing is recent', s.recent.length === 0);
}

console.log('\nwhat is set is remembered');
{
  store.clear();
  loadUiState();
  setUiState({ navCollapsed: true, inspectorOpen: true, navTab: 'files' });
  const raw = JSON.parse(store.get('wp:ui'));
  check('it is written under wp:ui', !!raw, [...store.keys()].join(','));
  check('and written whole', raw.navCollapsed === true && raw.navTab === 'files');

  const s = loadUiState();
  check('and read back', s.navCollapsed === true && s.inspectorOpen === true);
  check('with the tab', s.navTab === 'files');
  check('uiState agrees with the load', uiState().navTab === 'files');
}

console.log('\nrubbish in storage does not break the boot');
{
  for (const bad of ['', '{', 'null', '[]', '"a string"', '{"navTab":42}']) {
    store.clear();
    store.set('wp:ui', bad);
    const s = loadUiState();
    check(
      'survives ' + JSON.stringify(bad),
      s.navTab === 'outline' && typeof s.navCollapsed === 'boolean' && Array.isArray(s.recent)
    );
  }
}

console.log('\nfields are validated, not trusted');
{
  store.clear();
  store.set('wp:ui', JSON.stringify({
    navCollapsed: 'yes',
    inspectorOpen: 1,
    navTab: 'nonsense',
    inspectorSection: 'nonsense',
    recent: ['a', 5, 'b', null, 'c'],
  }));
  const s = loadUiState();
  check('a string is not true', s.navCollapsed === false);
  check('a one is not true either', s.inspectorOpen === false);
  check('an unknown tab falls back', s.navTab === 'outline');
  check('an unknown section falls back', s.inspectorSection === 'page');
  check('only strings survive in recent', s.recent.join(',') === 'a,b,c', s.recent.join(','));
}

console.log('\nthe recent list is capped');
{
  store.clear();
  loadUiState();
  setUiState({ recent: Array.from({ length: 30 }, (_v, i) => 'cmd.' + i) });
  check('at eight', loadUiState().recent.length === 8, String(loadUiState().recent.length));
}

console.log('\na storage that refuses to write does not throw');
{
  store.clear();
  loadUiState();
  const real = globalThis.localStorage.setItem;
  globalThis.localStorage.setItem = () => { throw new Error('quota'); };
  let threw = false;
  try { setUiState({ navCollapsed: true }); } catch { threw = true; }
  globalThis.localStorage.setItem = real;
  check('the write is swallowed', threw === false);
  check('and the panel still moved in memory', uiState().navCollapsed === true);
}

console.log(failures === 0 ? '\nall panel state checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
