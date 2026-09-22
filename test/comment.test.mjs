/**
 * Comments.
 *
 * Not one document in a 63-file corpus has a comment in it, which is exactly
 * why this test exists: the anchors used to be dropped on the first edit,
 * and nothing in that corpus would ever have caught it. A comment whose
 * range has vanished points at nothing, which is silent loss of somebody
 * else's words rather than your own.
 *
 *   npm run test:comment
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
class Q extends DOMParser { constructor() { super({ onError: () => {} }); } }
globalThis.DOMParser = Q; globalThis.XMLSerializer = XMLSerializer;
const { importDocx, exportDocx, unzip } = await import('./build/harness.mjs');
const HERE = dirname(fileURLToPath(import.meta.url));
const dec = new TextDecoder();

let failures = 0;
function check(name, ok, detail = '') {
  if (ok) { console.log('  ok    ' + name); return; }
  failures++; console.log('  FAIL  ' + name + (detail ? ' - ' + detail : ''));
}

const bytes = await readFile(join(HERE, 'corpus', 'comments.docx'));
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
const original = await unzip(buf);
const { doc, vault } = await importDocx(buf, 'comments.docx');

console.log('\ncomments are read');
{
  check('both are found', doc.comments?.length === 2, String(doc.comments?.length));
  check('with their authors', doc.comments?.[0].author === 'Dana Whitfield', doc.comments?.[0].author);
  check('and initials', doc.comments?.[0].initials === 'DW');
  check('and their text', doc.comments?.[0].text.includes('tighten this opening'));
  check('and a date', (doc.comments?.[0].date ?? '').startsWith('2026-03-14'));
}

console.log('\nthe anchors survive in the markup');
{
  const withCmt = doc.blocks.filter((b) => b.html && b.html.includes('data-cmt'));
  check('two paragraphs carry anchors', withCmt.length === 2, String(withCmt.length));
  check('a range start', withCmt[0].html.includes('data-cmt="start"'));
  check('a range end', withCmt[0].html.includes('data-cmt="end"'));
  check('a reference', withCmt[0].html.includes('data-cmt="ref"'));
  check('carrying the id', withCmt[0].html.includes('data-cmt-id="1"'), withCmt[0].html.slice(0, 90));
  check('the anchors hold no text', !/data-cmt="start"[^>]*>[^<]/.test(withCmt[0].html));
}

console.log('\nan untouched document still exports byte-identically');
{
  const out = await unzip(await (await exportDocx(doc, vault)).arrayBuffer());
  check(
    'document.xml unchanged',
    dec.decode(original.get('word/document.xml')) === dec.decode(out.get('word/document.xml'))
  );
  check('comments.xml unchanged',
    dec.decode(original.get('word/comments.xml')) === dec.decode(out.get('word/comments.xml')));
}

console.log('\nediting a commented paragraph keeps the comment anchored');
{
  const edited = structuredClone(doc);
  const i = edited.blocks.findIndex((b) => b.html && b.html.includes('data-cmt'));
  edited.blocks[i].html = edited.blocks[i].html.replace('The first paragraph', 'The FIRST paragraph');
  const out = await unzip(await (await exportDocx(edited, vault)).arrayBuffer());
  const xml = dec.decode(out.get('word/document.xml'));

  check('the edit is in the file', xml.includes('The FIRST paragraph'));
  check('the range start survives', xml.includes('<w:commentRangeStart w:id="1"/>'));
  check('the range end survives', xml.includes('<w:commentRangeEnd w:id="1"/>'));
  check('the reference survives', xml.includes('<w:commentReference w:id="1"/>'));
  check('and is inside a run, as the schema wants',
    /<w:r><w:commentReference w:id="1"\/><\/w:r>/.test(xml));
  check('the other comment is untouched too', xml.includes('<w:commentRangeStart w:id="2"/>'));
  check('comments.xml is still untouched',
    dec.decode(original.get('word/comments.xml')) === dec.decode(out.get('word/comments.xml')));
  const others = [...original.keys()].filter(
    (k) => k !== 'word/document.xml' &&
      dec.decode(original.get(k)) !== dec.decode(out.get(k) ?? new Uint8Array())
  );
  check('no other part changed', others.length === 0, others.join(', '));
}

console.log(failures === 0 ? '\nall comment checks passed' : `\n${failures} failed`);
process.exit(failures === 0 ? 0 : 1);
