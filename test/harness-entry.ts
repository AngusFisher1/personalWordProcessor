// Entry point bundled for the Node-side harness. The app itself never imports
// this; it exists so the round-trip suite exercises exactly the code the
// browser runs, rather than a reimplementation of it.
export { importDocx } from '../src/docx-import';
export { exportDocx } from '../src/export-docx';
export { unzip, partText } from '../src/docx-package';
export type { Vault } from '../src/docx-package';
export { toMarkdown, toHtml } from '../src/export-text';
export { fromMarkdown } from '../src/import-md';
export { sectionsOf, sectionIndexByBlock } from '../src/model';
