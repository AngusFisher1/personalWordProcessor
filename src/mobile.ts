/**
 * A phone.
 *
 * The page is 816 CSS pixels wide and must stay that way: every height the
 * paginator compares against is a CSS pixel, and a layout that shrank the
 * page would shrink the measurements with it and break the one thing this
 * program exists to do.
 *
 * So nothing is scaled in CSS. The VIEWPORT is widened instead - the browser
 * lays the app out at a fixed width and scales the rendered result down to
 * the device, exactly as it does for any fixed-width page. Measurement is
 * untouched, the caret lands where it looks like it should, and print still
 * matches the screen because the screen is still 816 pixels of paper.
 *
 * What does change is the chrome. At this width the rail would take a third
 * of the paper, so it becomes a drawer, and the document map goes away.
 */

/** Page, plus enough gutter to see the paper edge. */
const LAYOUT_WIDTH = 880;

let on = false;

export function isMobile(): boolean {
  return on;
}

function viewportTag(): HTMLMetaElement {
  let tag = document.querySelector('meta[name="viewport"]') as HTMLMetaElement | null;
  if (!tag) {
    tag = document.createElement('meta');
    tag.name = 'viewport';
    document.head.appendChild(tag);
  }
  return tag;
}

/**
 * A phone or a small tablet: a coarse pointer AND a screen too narrow for
 * the page. Either alone is the wrong test - a touchscreen laptop has a
 * coarse pointer and plenty of room, and a narrow window on a desktop is
 * someone who can simply widen it.
 */
function shouldBeMobile(): boolean {
  const coarse = window.matchMedia('(pointer: coarse)').matches;
  const narrow = window.matchMedia('(max-width: 900px)').matches;
  return coarse && narrow;
}

/**
 * On a phone the drawer IS the nav panel's collapsed state.
 *
 * Two booleans for one panel is how a panel ends up open according to one
 * of them and closed according to the other, so the phone drives the same
 * state the Ctrl+\ toggle does.
 */
let toggleNav: (() => void) | null = null;
let navCollapsed: (() => boolean) | null = null;

export function setNavHooks(toggle: () => void, collapsed: () => boolean): void {
  toggleNav = toggle;
  navCollapsed = collapsed;
}

function setDrawer(open: boolean): void {
  if (!navCollapsed || !toggleNav) return;
  if (navCollapsed() === !open) return;
  toggleNav();
}

function buildDrawerControls(): void {
  if (document.getElementById('drawer-toggle')) return;

  const toggle = document.createElement('button');
  toggle.id = 'drawer-toggle';
  toggle.className = 'drawer-toggle';
  toggle.setAttribute('aria-label', 'Show controls');
  toggle.textContent = '≡';
  toggle.addEventListener('mousedown', (e) => e.preventDefault());
  toggle.addEventListener('click', () => setDrawer(navCollapsed?.() ?? true));
  document.body.appendChild(toggle);

  // Tapping the page closes the drawer, which is what every drawer does and
  // what a reader will try first.
  const scrim = document.createElement('div');
  scrim.className = 'drawer-scrim';
  scrim.addEventListener('pointerdown', () => setDrawer(false));
  document.body.appendChild(scrim);
}

function apply(): void {
  const want = shouldBeMobile();
  if (want === on) return;
  on = want;
  document.body.classList.toggle('is-mobile', on);
  viewportTag().setAttribute(
    'content',
    on
      ? `width=${LAYOUT_WIDTH}, initial-scale=${(screen.width / LAYOUT_WIDTH).toFixed(3)}`
      : 'width=device-width, initial-scale=1.0'
  );
  if (on) buildDrawerControls();
  else setDrawer(false);
}

export function bindMobile(): void {
  apply();
  // Rotation changes which of the two this is.
  window.addEventListener('orientationchange', () => setTimeout(apply, 120));
  window.matchMedia('(max-width: 900px)').addEventListener('change', apply);
}

/** Close the drawer after a command runs, or it covers what it just did. */
export function closeDrawer(): void {
  if (on) setDrawer(false);
}
