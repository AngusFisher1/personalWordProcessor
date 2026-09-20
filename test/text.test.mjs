/**
 * The plain formats.
 *
 * The .docx path is guarded by byte-identity, which is a strong check but
 * only tells you the file came back unchanged. These formats are conversions,
 * so what matters instead is that nothing the reader can see goes missing:
 * every heading keeps its rank, every bullet stays a bullet, and the words
 * survive a trip out to Markdown and back.
 *
 *   npm run test:text
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
class Q extends DOMParser { constructor() { super({ onError: () => {} }); } }
globalThis.DOMParser = Q; globalThis.XMLSerializer = XMLSerializer;
const { importDocx, toMarkdown, toHtml, fromMarkdown, plainText } =
  await import('./build/harness.mjs');
const HERE = dirname(fileURLToPath(import.meta.url));

let failures = 0;
function check(name, ok, detail = '') {
  if (ok) { console.log('  ok    ' + name); return; }
  failures++; console.log('  FAIL  ' + name + (detail ? ' - ' + detail : ''));
}

const text = (html) => html.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').trim();
const para = (styleId, html, extra = {}) => ({ id: 'b' + (++seq), styleId, html, ...extra });
let seq = 0;

/* ------------------------------------------------------------------ */
console.log('\nMarkdown export');
{
  const doc = {
    id: 'd1',
    title: 'Sample',
    page: { width: 816, height: 1056, margins: { top: 48, right: 48, bottom: 48, left: 48 } },
    blocks: [
      para('Name', 'Ada Lovelace'),
      para('Contact', 'ada@example.com'),
      para('SectionHeading', 'Experience'),
      para('JobTitle', 'Analyst'),
      para('Body', 'Wrote the <b>first</b> program and <i>said so</i>.'),
      para('Bullet', 'Designed the engine'),
      para('Bullet', 'Nested point', { listLevel: 1 }),
      para('Bullet', 'Second step', { listMarker: '2.' }),
      para('Body', 'See <a href="https://example.com">the notes</a>.'),
    ],
  };
  const md = toMarkdown(doc);
  check('the title is an h1', md.includes('# Ada Lovelace'));
  check('a section is an h2', md.includes('## Experience'));
  check('a job title is an h3', md.includes('### Analyst'));
  check('bold survives', md.includes('**first**'));
  check('italic survives', md.includes('*said so*'));
  check('a link survives', md.includes('[the notes](https://example.com)'));
  check('a bullet is a dash', md.includes('- Designed the engine'));
  check('a nested bullet is indented', md.includes('  - Nested point'));
  check('a numbered item keeps its number', md.includes('2. Second step'));
  check(
    'consecutive bullets are not separated by blank lines',
    !/- Designed the engine\n\n/.test(md)
  );
}

/* ------------------------------------------------------------------ */
console.log('\nMarkdown escaping');
{
  const doc = {
    id: 'd2', title: 'Escapes',
    page: { width: 816, height: 1056, margins: { top: 48, right: 48, bottom: 48, left: 48 } },
    blocks: [
      para('Body', 'A * literal asterisk and an _underscore_ char'),
      para('Body', '- not a bullet'),
      para('Body', '1. not a list'),
      para('Body', '# not a heading'),
    ],
  };
  const md = toMarkdown(doc);
  const back = fromMarkdown(md);
  const styles = back.blocks.map((b) => b.styleId);
  check('an escaped dash stays body text', styles[0] === 'Body' && styles[1] === 'Body');
  check('no line became a list', !styles.includes('Bullet'), styles.join(','));
  check('no line became a heading', !styles.includes('Name'), styles.join(','));
  check(
    'the literal text comes back',
    back.blocks.some((b) => text(b.html) === '- not a bullet'),
    back.blocks.map((b) => text(b.html)).join(' | ')
  );
}

/* ------------------------------------------------------------------ */
console.log('\nMarkdown import');
{
  const doc = fromMarkdown(
    [
      '# Ada Lovelace',
      '',
      'ada@example.com',
      '',
      'Experience',
      '----------',
      '',
      '### Analyst',
      '',
      'Wrote the **first** program and *said so*.',
      '',
      '- One',
      '  - Two',
      '3) Three',
      '',
      '| A | B |',
      '| --- | --- |',
      '| 1 | 2 |',
      '',
    ].join('\n')
  );
  const styles = doc.blocks.map((b) => b.kind === 'table' ? 'table' : b.styleId);
  check('the title becomes Name', styles[0] === 'Name');
  check('the line under it becomes Contact', styles[1] === 'Contact', styles.join(','));
  check('a setext underline becomes a section', styles[2] === 'SectionHeading', styles.join(','));
  check('an h3 becomes a job title', styles[3] === 'JobTitle', styles.join(','));
  check('bold parses', doc.blocks[4].html.includes('<b>first</b>'), doc.blocks[4].html);
  check('italic parses', doc.blocks[4].html.includes('<i>said so</i>'), doc.blocks[4].html);
  check('bullets parse', styles[5] === 'Bullet' && styles[6] === 'Bullet');
  check('nesting is read', doc.blocks[6].listLevel === 1, String(doc.blocks[6].listLevel));
  check('a numbered item keeps its marker', doc.blocks[7].listMarker === '3)', doc.blocks[7].listMarker);
  const table = doc.blocks.find((b) => b.kind === 'table');
  check('a pipe table becomes a table', !!table);
  check('its header row is flagged', table?.rows[0].headerRow === true);
  check('the dashed line is not a row', table?.rows.length === 2, String(table?.rows.length));
  check('its cells are read', text(table?.rows[1].cells[1][0].html) === '2');
  check('the widths sum to the content box', table?.cols.reduce((a, c) => a + c.width, 0) === 720);
  check('the title comes from the h1', doc.title === 'Ada Lovelace', doc.title);
}

/* ------------------------------------------------------------------ */
console.log('\na real document survives the Markdown round trip');
{
  const bytes = await readFile(join(HERE, 'corpus', 'resume.docx'));
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const { doc } = await importDocx(buf, 'resume.docx');
  const back = fromMarkdown(toMarkdown(doc));

  const words = (d) =>
    d.blocks
      .flatMap((b) =>
        b.kind === 'table'
          ? b.rows.flatMap((r) => r.cells.flatMap((c) => c.map((p) => p.html)))
          : [b.html]
      )
      .join(' ')
      .replace(/<[^>]*>/g, ' ')
      .replace(/&amp;/g, '&')
      .split(/\s+/)
      .filter(Boolean);

  const before = words(doc);
  const after = words(back);
  const lost = before.filter((w, i) => after[i] !== w).length;
  check(
    'every word comes back in order',
    lost === 0,
    lost + ' of ' + before.length + ' differ'
  );

  const rank = (d) =>
    d.blocks.filter((b) => b.kind !== 'table' && b.styleId !== 'Body' && b.styleId !== 'Bullet')
      .length;
  check(
    'the heading structure survives',
    rank(back) >= rank(doc),
    rank(doc) + ' out, ' + rank(back) + ' back'
  );
  const bullets = (d) => d.blocks.filter((b) => b.styleId === 'Bullet').length;
  check(
    'every bullet is still a bullet',
    bullets(back) === bullets(doc),
    bullets(doc) + ' out, ' + bullets(back) + ' back'
  );
}

/* ------------------------------------------------------------------ */
console.log('\nHTML export');
{
  const bytes = await readFile(join(HERE, 'corpus', 'report-tables.docx'));
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const { doc } = await importDocx(buf, 'report-tables.docx');
  const html = toHtml(doc);

  check('it is a whole document', html.startsWith('<!doctype html>'));
  check('it carries its own stylesheet', html.includes('<style>') && html.includes('.s-Body{'));
  check('it names the document', html.includes('<title>'));
  check('the page box matches the page setup', html.includes(doc.page.width + 'px'));
  check('a table is a table', html.includes('<table>') && html.includes('<td>'));
  check('there are no unresolved model attributes', !html.includes('data-block-id'));
  check('nothing external is referenced', !/<(script|link)\b/i.test(html));

  // Every paragraph's text should appear in the output.
  const missing = doc.blocks
    .filter((b) => b.kind !== 'table')
    .map((b) => text(b.html))
    .filter((t) => t.length > 3 && !html.includes(t.slice(0, 24).replace(/&/g, '&amp;')));
  check('every paragraph is present', missing.length === 0, String(missing.length) + ' missing');
}

/* ------------------------------------------------------------------ */
console.log('\nan image with no resolver is reported, not linked');
{
  const doc = {
    id: 'd3', title: 'Pictures',
    page: { width: 816, height: 1056, margins: { top: 48, right: 48, bottom: 48, left: 48 } },
    blocks: [para('Body', 'Before <img data-media="word/media/image1.png" data-run="r1"/> after')],
  };
  const md = toMarkdown(doc, () => null);
  const html = toHtml(doc, () => null);
  check('Markdown says a picture was here', md.includes('[image: image1.png]'), md.trim());
  check('Markdown does not link a missing file', !md.includes(']('), md.trim());
  check('HTML says a picture was here', html.includes('[image: image1.png]'));
  check('HTML emits no broken img', !html.includes('<img'));

  const resolved = toMarkdown(doc, (p) => 'media/' + p.split('/').pop());
  check('a resolver is used when given', resolved.includes('](media/image1.png)'), resolved.trim());
}

/* ------------------------------------------------------------------ */
console.log('\nplain text decodes entities as well as stripping tags');
{
  // The outline, a document title and a word count all show this text to a
  // person. Stripping the tags without decoding puts a literal "&amp;" in
  // the sidebar beside a paragraph that renders a perfectly good ampersand.
  check('an ampersand decodes', plainText('Kestrel &amp; Moss') === 'Kestrel & Moss');
  check('a curly quote decodes', plainText('the company&rsquo;s') === 'the company’s');
  check('tags go first', plainText('<b>Skills</b> &amp; <i>Tools</i>') === 'Skills & Tools');
  check('a numeric entity decodes', plainText('caf&#233; &#x2014; bar') === 'café — bar');
  check(
    'ampersand is not decoded twice',
    plainText('a &amp;lt; b') === 'a &lt; b',
    plainText('a &amp;lt; b')
  );
  check('an unknown entity is left alone', plainText('&notareal; x') === '&notareal; x');
  check('no markup and no entities is a no-op', plainText('plain words') === 'plain words');
}

console.log(failures === 0 ? '\nall text-format checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
