import JSZip from 'jszip';
import type { Doc } from './model';
import { isTable } from './model';
import { exportJson, loadById, readIndex, safeFileName } from './persist';
import { toHtml, toMarkdown } from './export-text';

/**
 * The whole library, in one archive, in formats nothing owns.
 *
 * This is the escape hatch the rest of the program is measured against: if
 * it works, no document here is hostage to this program continuing to exist,
 * or to its localStorage surviving a cleared browser. It deliberately does
 * not go through the .docx path - that path needs each document's original
 * package, which lives in IndexedDB and may be gone - so it writes what can
 * always be written from the model alone.
 */

export interface BulkResult {
  blob: Blob;
  documents: number;
  /** Documents the index listed but storage could not produce. */
  missing: string[];
}

function wordCount(doc: Doc): number {
  let text = '';
  for (const b of doc.blocks) {
    if (isTable(b)) {
      for (const row of b.rows) for (const cell of row.cells) for (const p of cell) text += ' ' + p.html;
    } else {
      text += ' ' + b.html;
    }
  }
  const words = text.replace(/<[^>]*>/g, ' ').trim();
  return words === '' ? 0 : words.split(/\s+/).length;
}

function stamp(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return (
    d.getFullYear() +
    '-' +
    pad(d.getMonth() + 1) +
    '-' +
    pad(d.getDate()) +
    ' ' +
    pad(d.getHours()) +
    pad(d.getMinutes())
  );
}

/**
 * Every document, as Markdown, JSON and HTML.
 *
 * Markdown is for a human with a text editor, JSON is a lossless copy of the
 * model this program can read back, and HTML is what you send someone who
 * only has a browser. Three formats rather than one because the point is not
 * to have a backup, it is to have no single thing that has to still work.
 */
export async function exportLibrary(): Promise<BulkResult> {
  const zip = new JSZip();
  const index = readIndex();
  const missing: string[] = [];
  const rows: string[] = [];
  const manifest: Record<string, unknown>[] = [];
  const used = new Set<string>();
  let n = 0;

  for (const entry of index) {
    const doc = loadById(entry.id);
    if (!doc) {
      missing.push(entry.title || entry.id);
      continue;
    }
    // Two documents may share a title; the archive cannot.
    let name = safeFileName(doc.title, 'document');
    if (used.has(name.toLowerCase())) {
      let i = 2;
      while (used.has((name + ' (' + i + ')').toLowerCase())) i++;
      name = name + ' (' + i + ')';
    }
    used.add(name.toLowerCase());

    zip.file('markdown/' + name + '.md', toMarkdown(doc));
    zip.file('json/' + name + '.json', exportJson(doc));
    zip.file('html/' + name + '.html', toHtml(doc));

    const words = wordCount(doc);
    const when = new Date(entry.updatedAt || Date.now());
    rows.push(
      '| [' + name + '](markdown/' + encodeURI(name) + '.md) | ' +
        words.toLocaleString() +
        ' | ' +
        when.toISOString().slice(0, 10) +
        ' |'
    );
    manifest.push({
      id: doc.id,
      title: doc.title,
      file: name,
      words,
      blocks: doc.blocks.length,
      updatedAt: entry.updatedAt,
    });
    n++;
  }

  const now = new Date();
  zip.file(
    'README.md',
    [
      '# Library export',
      '',
      stamp(now) + ' — ' + n + ' document' + (n === 1 ? '' : 's'),
      '',
      '- `markdown/` — one .md per document, readable anywhere.',
      '- `json/` — the full model, the only lossless copy; re-importable.',
      '- `html/` — one self-contained page per document.',
      '',
      'Images are not in this archive: they live in each document’s original',
      '.docx, which is where they stay lossless. Export a document as Word to',
      'get its pictures.',
      '',
      '| Document | Words | Updated |',
      '| --- | --- | --- |',
      ...rows,
      ...(missing.length
        ? ['', '## Not exported', '', ...missing.map((m) => '- ' + m + ' (not found in storage)')]
        : []),
      '',
    ].join('\n')
  );
  zip.file(
    'manifest.json',
    JSON.stringify({ exportedAt: now.toISOString(), documents: manifest }, null, 2)
  );

  const blob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  return { blob, documents: n, missing };
}

export function libraryFileName(): string {
  return 'library ' + stamp(new Date()) + '.zip';
}
