# Paged web word processor

A browser-based word processor that renders documents as fixed-size US Letter
pages, the way Word and Google Docs do. Text flows from page to page as you
type, and printing produces a PDF that matches the screen one to one. Built for
resume editing.

The browser has no native concept of a page. Pagination is implemented here by
measuring content and moving DOM nodes between fixed-height page containers.

```bash
npm install
npm run dev
```

`npm run build` type-checks and bundles to `dist/`.

## Stack

Vanilla TypeScript, Vite, and exactly one runtime dependency (`docx`, loaded
lazily and only when you export). No framework: React's virtual DOM fights the
direct node manipulation that is the core of this app. No rich-text library —
writing the pagination layer is the point of the project.

## Geometry

All geometry is in CSS pixels at 96px per inch, and the numbers are exact.

| | |
|---|---|
| Page | 816 × 1056 (8.5in × 11in) |
| Narrow margin (default) | 48px (0.5in) → content box 720 × 960 |
| Normal margin | 96px (1in) → content box 624 × 864 |

## How it works

`src/`

| File | Role |
|---|---|
| `main.ts` | entry, toolbar, event wiring, autosave |
| `model.ts` | `Block`, `Doc`, geometry constants, id generation |
| `styles.ts` | the six named styles; emitted as `.s-<StyleId>` CSS at startup |
| `render.ts` | model → DOM, DOM → model, the inline sanitizer |
| `paginate.ts` | measurement, height cache, page assignment, normalization |
| `caret.ts` | serialize and restore the caret across node moves |
| `commands.ts` | toolbar actions, Enter/Backspace/Delete, shortcuts |
| `paste.ts` | clipboard sanitizer and block mapping |
| `persist.ts` | localStorage, JSON import/export |
| `history.ts` | undo/redo stack |
| `export-docx.ts` | .docx generation |
| `styles.css` | page geometry, print rules |

A few decisions are load-bearing:

**One `contenteditable`, on `#doc`** — not one per page. The browser then owns
selection, caret movement, arrow keys, shift-selection and text input across
page boundaries for free. One editable per page breaks the first time someone
backspaces at the top of page 2.

**The model is flat and ordered, and a `Block` has no `page` field.** Pages are
a rendering outcome, recomputed from measurement, never stored. Persisting page
membership turns every later bug into a stale-state bug.

**The DOM is the source of truth while editing.** The `Doc` object is a
serialization target, refreshed from the DOM on save, export and undo snapshot.
Keeping a JS model in sync with `contenteditable` on every keystroke means
fighting the browser over every input event.

**Pagination moves existing elements, never re-creates them.** `insertBefore`
preserves focus, avoids flicker and keeps IME composition alive. Re-rendering
from the model on every reflow makes typing feel broken in a way that is hard
to trace back to the line that caused it.

**Every function that moves nodes captures the caret first and restores it
last.** The caret is stored as a block id plus a character offset counted with
`Range.toString().length`, so it survives both node movement and `<b>`/`<i>`
boundaries where node-based offsets do not.

### The two pagination paths

`paginateIfNeeded()` runs synchronously on every input event, before paint. It
looks only at the page containing the caret: if that page overflows, or if room
has reopened for the first block of the next page, it triggers a reflow scoped
to that page onward. Everything else is untouched.

`paginate({ fromPage })` is the full reflow: measure every block from that page
forward, accumulate heights against the content box, assign blocks to pages,
reconcile each page's children, drop trailing empty pages, restore the caret.
It also runs debounced (150ms) after paste, style changes, margin changes and
undo.

Heights are cached by `(content width, styleId, innerHTML)` rather than by
block id, so identical blocks share an entry and editing one block never
invalidates another. The cache is cleared when the margin changes or fonts
settle.

Two details worth knowing if you touch this code:

- `.page-content` has a **fixed** height, so `scrollHeight` never reports less
  than the full content box. Used height is the sum of the block heights, which
  also guarantees the cheap check and the full reflow never disagree about
  whether a block fits.
- Spacing is **padding only**, never margin. Sibling margins collapse, and
  summed block heights would then disagree with the container height.

`normalize()` runs before every measurement. `contenteditable` periodically
drops a bare text node, a stray `<div>` or a `<br>` straight into
`.page-content`, usually at a page boundary; normalize wraps, unwraps or
removes it, hoists blocks the browser nested inside other blocks, and mints ids
for blocks a clone stripped them from. Select-all-and-delete destroys the page
elements outright, and normalize rebuilds them.

### Line-level splitting

A paragraph that does not fit the remaining space is split at a line boundary
rather than moved whole, so pages fill to the bottom margin.

The split point comes from `Range.getClientRects()`, which yields one rect per
text fragment; those are merged by vertical overlap into lines. Those rects are
**ink** boxes - the glyph extent, about 13px for 11px text - not line boxes, so
they are used only to count and locate lines. Heights come from the used line
box height, derived by dividing the block's content height by its line count.
Measuring from the ink boxes instead makes every split piece render a couple of
pixels taller than predicted, which is enough to overflow a page.

The character offset where a line begins is found by binary search:
`lineOfPrefix` is non-decreasing in the offset, so the first offset whose prefix
reaches line L is one past that line's first character. About ten range
measurements per split.

The rendered pieces are a **render artifact**. The first piece keeps the logical
block's id; continuations carry `data-continues-from` and `data-base`, the
character offset where the piece starts in the merged text. `readModel` folds
them back, so the stored document always has one block per paragraph and
persistence, export, undo and .docx know nothing about splitting. The caret is
likewise logical - block id plus offset into the merged text - which is what
lets a position survive being re-split at a different point while typing.

Pagination folds every paragraph back into one element before measuring and
re-splits afterwards. Because the head keeps its element identity and only
continuations are created and destroyed, the caret's own element normally
survives a reflow untouched, so focus and IME composition are not disturbed.

Per-style paragraph properties mirror Word's:

| Property | Default | Meaning |
|---|---|---|
| `keepWithNext` | on for Name, SectionHeading, JobTitle | never separated from the block below |
| `keepLines` | off (on for the heading styles) | never split at all |
| `orphanMin` / `widowMin` | 2 | minimum lines either side of a split |
| `pageBreakBefore` | off | always starts a page |

Two details that are easy to get wrong:

- When the maximum number of lines that fit would leave a widow, the split
  **backs off** to the latest legal line rather than abandoning the split. Not
  doing that leaves several lines of avoidable whitespace at the foot of a page.
- Each page records `data-needh`: exactly how much room would have to reopen
  before the layout could change. That is what the cheap per-keystroke check
  compares against. Getting it wrong in the keep-with-next case - where blocks
  are dragged off a page after its budget was computed - makes every keystroke
  trigger a reflow that rebuilds the identical layout.

### Styles

Six named styles, defined once in `styles.ts` and read by the CSS emitter, the
style dropdown, the Enter key and the .docx exporter.

| Style | Size / weight | Enter produces |
|---|---|---|
| Name | 24px bold | Contact |
| Contact | 10px | Body |
| SectionHeading | 12px bold, uppercase, bottom rule | Body |
| JobTitle | 11px bold | Body |
| Body | 11px, line-height 1.35 | Body |
| Bullet | 11px, hanging indent 14px | Bullet |

Pressing Enter at the end of a block creates the next style in that table;
pressing it mid-block splits and both halves keep the current style, as in Word.
Backspace at offset 0 merges into the previous block and adopts its style.

## Printing to PDF

Use the **Print / PDF** button, which forces a reflow and waits for fonts
before opening the dialog.

In the browser's print dialog you must set:

- **Margins: None**
- **Headers and footers: off**

There is no way to set these programmatically. With them set, output matches
the screen exactly, because each `.page` is already an exact letter box and the
print stylesheet supplies `@page { size: letter; margin: 0 }`.

## .docx export

`Block` → `Paragraph`, `styleId` → a named paragraph style, `<b>/<i>/<u>` →
`TextRun` flags, `<a href>` → `ExternalHyperlink`, Bullet → a paragraph with a
numbering reference. All six styles are defined in `styles.paragraphStyles`, so
they show up in Word's style gallery and the file behaves like a real Word
document rather than a dump of formatted text.

Notes:

- **Page breaks are not exported.** Word repaginates on its own and forcing our
  breaks in would fight it. Expect the .docx to break pages in different places
  than the PDF.
- Font sizes are rounded to the nearest half point (Word's smallest unit), so
  11px becomes 8.5pt rather than exactly 8.25pt.
- The document font is Calibri, matching the on-screen stack.

## Storage

Autosaves to `localStorage` on a 1s debounce under `wp:doc:<id>`, with an index
at `wp:docs`. If storage is full the save fails visibly in the toolbar rather
than silently — export JSON at that point to keep your work. JSON export writes
the `Doc` verbatim; import validates that `blocks` is an array and every
`styleId` is known before replacing the document.

## Not built, on purpose

- **Splitting a single line.** A block whose one line is taller than a whole
  page is placed and clipped.

- Tables, images, columns, text boxes. Tables are **dropped** on paste, text
  and all — worth knowing if you paste a resume that uses a table for layout.
- .docx or PDF import; headers, footers, page numbers; collaboration; spell
  check beyond the browser's own; mobile layout; any backend.

Undo granularity is coarser than Word's: typing coalesces into one entry per
500ms of activity, and each structural change is one entry.

## Acceptance tests

Run these in order; all pass as of the last change.

0. **Line-level pagination** - on randomized multi-page documents: no page
   overflows, no heading is ever the last line on a page, and every page's
   trailing whitespace is provably explained (its slack is strictly less
   than its recorded `needh`). Text survives split and re-merge intact.
1. **Geometry** — page is 816×1056, content box 720×960; the last line on a
   full page sits within one line height of the bottom margin.
2. **Flow down** — hold Enter to page 3 and beyond. No page ever exceeds the
   content box mid-keystroke; the caret stays with the cursor.
3. **Flow up** — hold Backspace back to page 1. Blocks pull up, empty pages
   disappear, the caret lands at the join point every time.
4. **Boundary typing** — type on the last line of page 1 until the block wraps
   to a line that no longer fits. It moves to page 2 mid-word without losing
   focus or dropping a character.
5. **Mid-document insert** — paste into page 2 of a 5-page document. Later
   pages reflow, the scroll position does not jump.
6. **Style change reflow** — change a block on page 1 to Name; everything below
   reflows and the caret is preserved.
7. **Paste hostile input** — paste from Google Docs and from Word. No foreign
   fonts, colors, spacing or classes survive; bullets become Bullet blocks
   (including Word's `MsoListParagraph` paragraphs with the glyph inlined as
   text); Google Docs' `<b style="font-weight:normal">` wrapper does not bold
   the document; emphasis carried as `<span style="font-weight:700">` is kept.
8. **Font race** — no web fonts are loaded, so pagination is correct on first
   paint rather than corrected a moment later. The stylesheet is a
   render-blocking `<link>` so the first measurement is never unstyled.
9. **Undo depth** — type, restyle, paste, then undo four times. Each step
   reverses cleanly with the caret restored, and undo stops at the baseline.
10. **Reload** — refresh mid-edit; the document restores identically, page
    breaks included.
11. **Export match** — the .docx carries all six named styles, bold/italic/
    underline, hyperlinks, line breaks, bullet numbering, 8.5×11in page size
    and the right margins, with no forced page breaks.
12. **Performance** — on a 10-page document the input handler runs at ~0.5ms
    median and under 2ms at p99, whether typing on page 1 or page 10.
