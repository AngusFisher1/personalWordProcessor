/**
 * Generates a SEED corpus of .docx files into test/corpus/.
 *
 * These are synthetic. The spec is right that synthetic documents only
 * exercise the paths someone already thought of, so this is a floor, not the
 * corpus: drop real .docx files into test/corpus/ and the harness picks them
 * up automatically. What these do cover is structure this app does not itself
 * generate - tables, headers and footers, multi-level numbering, landscape and
 * legal page sizes, tracked changes, content controls - which is exactly what
 * the importer has to survive.
 *
 *   node test/make-corpus.mjs
 */
import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  ExternalHyperlink,
  Footer,
  Header,
  HeadingLevel,
  LevelFormat,
  PageNumber,
  PageOrientation,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import JSZip from 'jszip';
import { deflateSync } from 'node:zlib';
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const OUT = join(HERE, 'corpus');

/* A real PNG, built here so the corpus needs no binary fixtures in git. */
function crc32(buf) {
  let c, table = [];
  for (let n = 0; n < 256; n++) {
    c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  let crc = 0xffffffff;
  for (const b of buf) crc = table[(crc ^ b) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function makePng(w, h, rgb) {
  const raw = Buffer.alloc((w * 3 + 1) * h);
  let o = 0;
  for (let y = 0; y < h; y++) {
    raw[o++] = 0; // filter: none
    for (let x = 0; x < w; x++) {
      // A visible diagonal so the picture is obviously a picture.
      const on = (x + y) % 24 < 12;
      raw[o++] = on ? rgb[0] : 255;
      raw[o++] = on ? rgb[1] : 255;
      raw[o++] = on ? rgb[2] : 255;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 2;  // colour type: truecolour
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

const lorem =
  'This paragraph exists to occupy a realistic amount of space on the page so that pagination has something to do. ';

async function save(name, doc) {
  const buf = await Packer.toBuffer(doc);
  await writeFile(join(OUT, name), buf);
  console.log('  ' + name + '  ' + buf.length + ' bytes');
}

function para(text, opts = {}) {
  return new Paragraph({ text, ...opts });
}

/* ---------------- 1. resume ---------------- */
async function resume() {
  await save(
    'resume.docx',
    new Document({
      sections: [
        {
          children: [
            new Paragraph({ text: 'Dana Whitfield', heading: HeadingLevel.TITLE }),
            para('Bristol · dana@example.com · (555) 010-0100'),
            new Paragraph({ text: 'Experience', heading: HeadingLevel.HEADING_1 }),
            new Paragraph({ text: 'Staff Engineer, Northwind', heading: HeadingLevel.HEADING_3 }),
            new Paragraph({
              bullet: { level: 0 },
              children: [
                new TextRun('Rebuilt the '),
                new TextRun({ text: 'billing pipeline', bold: true }),
                new TextRun(' and cut invoice errors by '),
                new TextRun({ text: '40%', italics: true }),
                new TextRun('.'),
              ],
            }),
            new Paragraph({ bullet: { level: 0 }, text: 'Mentored four engineers.' }),
            new Paragraph({ bullet: { level: 1 }, text: 'Two were promoted within a year.' }),
            new Paragraph({ text: 'Education', heading: HeadingLevel.HEADING_1 }),
            para('B.S. Computer Science, State University'),
          ],
        },
      ],
    })
  );
}

/* ---------------- 2. cover letter ---------------- */
async function coverLetter() {
  await save(
    'cover-letter.docx',
    new Document({
      sections: [
        {
          children: [
            para('12 March 2026'),
            para('Dear hiring manager,'),
            ...Array.from({ length: 6 }, (_, i) =>
              para('Paragraph ' + (i + 1) + '. ' + lorem.repeat(3))
            ),
            para('Yours sincerely,'),
            para('Dana Whitfield'),
          ],
        },
      ],
    })
  );
}

/* ---------------- 3. contract with numbered clauses ---------------- */
async function contract() {
  await save(
    'contract.docx',
    new Document({
      numbering: {
        config: [
          {
            reference: 'clauses',
            levels: [
              { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START },
              { level: 1, format: LevelFormat.DECIMAL, text: '%1.%2', alignment: AlignmentType.START },
              { level: 2, format: LevelFormat.LOWER_ROMAN, text: '(%3)', alignment: AlignmentType.START },
            ],
          },
        ],
      },
      sections: [
        {
          children: [
            new Paragraph({ text: 'Services Agreement', heading: HeadingLevel.TITLE }),
            ...[
              [0, 'Definitions. In this agreement the following terms apply.'],
              [1, 'Services means the work described in Schedule A.'],
              [1, 'Fees means the amounts set out in Schedule B.'],
              [2, 'Fees exclude value added tax.'],
              [2, 'Fees are payable within thirty days.'],
              [0, 'Term. This agreement runs for twelve months.'],
              [1, 'Either party may terminate on notice.'],
              [0, 'Governing law. The laws of England and Wales apply.'],
            ].map(
              ([level, text]) =>
                new Paragraph({ text, numbering: { reference: 'clauses', level } })
            ),
          ],
        },
      ],
    })
  );
}

/* ---------------- 4. report with tables ---------------- */
async function report() {
  const row = (cells, header = false) =>
    new TableRow({
      tableHeader: header,
      children: cells.map(
        (t) =>
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: t, bold: header })] })],
          })
      ),
    });
  await save(
    'report-tables.docx',
    new Document({
      sections: [
        {
          children: [
            new Paragraph({ text: 'Quarterly Report', heading: HeadingLevel.TITLE }),
            para(lorem.repeat(4)),
            new Paragraph({ text: 'Results', heading: HeadingLevel.HEADING_1 }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                row(['Region', 'Revenue', 'Change'], true),
                row(['North', '1,240', '+8%']),
                row(['South', '980', '-3%']),
                row(['East', '1,510', '+21%']),
              ],
            }),
            para(lorem.repeat(3)),
            new Paragraph({ text: 'Outlook', heading: HeadingLevel.HEADING_1 }),
            para(lorem.repeat(5)),
          ],
        },
      ],
    })
  );
}

/* ---------------- 5/6. landscape and legal ---------------- */
async function landscape() {
  await save(
    'landscape.docx',
    new Document({
      sections: [
        {
          properties: { page: { size: { orientation: PageOrientation.LANDSCAPE } } },
          children: [
            new Paragraph({ text: 'Wide View', heading: HeadingLevel.HEADING_1 }),
            ...Array.from({ length: 8 }, (_, i) => para('Row ' + i + '. ' + lorem.repeat(2))),
          ],
        },
      ],
    })
  );
}

async function legal() {
  await save(
    'legal-size.docx',
    new Document({
      sections: [
        {
          // 8.5in x 14in in twips
          properties: { page: { size: { width: 12240, height: 20160 } } },
          children: [
            new Paragraph({ text: 'Legal Size', heading: HeadingLevel.HEADING_1 }),
            ...Array.from({ length: 12 }, (_, i) => para('Clause ' + i + '. ' + lorem.repeat(2))),
          ],
        },
      ],
    })
  );
}

/* ---------------- 7. letterhead with header/footer ---------------- */
async function letterhead() {
  await save(
    'letterhead.docx',
    new Document({
      sections: [
        {
          properties: { titlePage: true },
          headers: {
            first: new Header({ children: [para('NORTHWIND LIMITED')] }),
            default: new Header({ children: [para('Northwind — continued')] }),
          },
          footers: {
            default: new Footer({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun('Page '),
                    new TextRun({ children: [PageNumber.CURRENT] }),
                    new TextRun(' of '),
                    new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
                  ],
                }),
              ],
            }),
          },
          children: Array.from({ length: 20 }, (_, i) =>
            para('Letterhead paragraph ' + i + '. ' + lorem.repeat(3))
          ),
        },
      ],
    })
  );
}

/* ---------------- 7b. two sections, different margins ---------------- */
/**
 * The shape real resumes turn out to have: a body set tight, and a trailing
 * section on Word's own margins. Reading only the last sectPr - which is
 * what a single page setup amounts to - lays the whole document out with
 * the geometry of its last few paragraphs.
 */
async function twoSections() {
  await save(
    'two-sections.docx',
    new Document({
      sections: [
        {
          properties: {
            page: { margin: { top: 360, right: 540, bottom: 360, left: 540 } },
          },
          headers: { default: new Header({ children: [para('SECTION ONE')] }) },
          children: [
            new Paragraph({ text: 'Tight Section', heading: HeadingLevel.HEADING_1 }),
            ...Array.from({ length: 10 }, (_, i) =>
              para('Narrow margin paragraph ' + i + '. ' + lorem.repeat(2))
            ),
          ],
        },
        {
          properties: {
            page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
          },
          headers: { default: new Header({ children: [para('SECTION TWO')] }) },
          children: [
            new Paragraph({ text: 'Wide Section', heading: HeadingLevel.HEADING_1 }),
            ...Array.from({ length: 6 }, (_, i) =>
              para('Inch margin paragraph ' + i + '. ' + lorem.repeat(2))
            ),
          ],
        },
      ],
    })
  );
}

/* ---------------- 7c. direct paragraph formatting ---------------- */
/**
 * Alignment, indents and spacing set on the paragraph rather than through a
 * style. 47% of the paragraphs in a real corpus carry at least one of these.
 */
async function directFormat() {
  await save(
    'direct-format.docx',
    new Document({
      sections: [
        {
          children: [
            new Paragraph({
              text: 'Centred Title',
              alignment: AlignmentType.CENTER,
              spacing: { after: 240 },
            }),
            new Paragraph({
              text: 'Right aligned dateline, 14 March 2026',
              alignment: AlignmentType.RIGHT,
            }),
            new Paragraph({
              text: 'Justified body. ' + lorem.repeat(3),
              alignment: AlignmentType.JUSTIFIED,
              spacing: { before: 120, after: 120 },
            }),
            new Paragraph({
              text: 'An indented block quote, set in from both margins. ' + lorem,
              indent: { left: 720, right: 720 },
            }),
            new Paragraph({
              text: 'A hanging indent, the shape a definition list wants. ' + lorem,
              indent: { left: 720, hanging: 360 },
            }),
            new Paragraph({
              text: 'First-line indent, the shape running prose wants. ' + lorem,
              indent: { firstLine: 360 },
            }),
            new Paragraph({
              text: 'One and a half line spacing. ' + lorem.repeat(2),
              spacing: { line: 360, lineRule: 'auto' },
            }),
            new Paragraph({
              text: 'Exactly fourteen points of leading. ' + lorem.repeat(2),
              spacing: { line: 280, lineRule: 'exact' },
            }),
            para('Nothing set on this one at all.'),
          ],
        },
      ],
    })
  );
}

/* ---------------- 7d. character formatting ---------------- */
/**
 * Mixed sizes, fonts and colours WITHIN paragraphs. 79% of the runs in a
 * real corpus carry at least one of these, and the vault used to flatten
 * them all to the first run's properties on any edit.
 */
async function runFormat() {
  await save(
    'run-format.docx',
    new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'Plain, then ' }),
                new TextRun({ text: 'red', color: 'C00000' }),
                new TextRun({ text: ', then ' }),
                new TextRun({ text: 'large', size: 36 }),
                new TextRun({ text: ', then ' }),
                new TextRun({ text: 'Courier', font: 'Courier New' }),
                new TextRun({ text: ', then ' }),
                new TextRun({ text: 'struck', strike: true }),
                new TextRun({ text: ', then x' }),
                new TextRun({ text: '2', superScript: true }),
                new TextRun({ text: ' and H' }),
                new TextRun({ text: '2', subScript: true }),
                new TextRun({ text: 'O.' }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'A whole paragraph at one size. ', size: 28 }),
                new TextRun({ text: 'Still the same size here. ', size: 28 }),
                new TextRun({ text: lorem, size: 28 }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'Bold and coloured together', bold: true, color: '1F4E79' }),
                new TextRun({ text: ' beside plain text. ' }),
                new TextRun({ text: lorem }),
              ],
            }),
          ],
        },
      ],
    })
  );
}

/* ---------------- 8. hyperlinks and unusual font ---------------- */
async function links() {
  await save(
    'links-and-fonts.docx',
    new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [
                new TextRun('Visit '),
                new ExternalHyperlink({
                  children: [new TextRun({ text: 'the site', style: 'Hyperlink' })],
                  link: 'https://example.com/docs',
                }),
                new TextRun(' for details.'),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'Set in Garamond. ', font: 'Garamond', size: 28 }),
                new TextRun({ text: 'And in Consolas.', font: 'Consolas', size: 20 }),
              ],
            }),
            ...Array.from({ length: 4 }, (_, i) => para('Body ' + i + '. ' + lorem.repeat(2))),
          ],
        },
      ],
    })
  );
}

/* ---------------- 8b. a document with pictures ---------------- */
async function withImages() {
  const logo = makePng(160, 90, [217, 123, 60]);
  const figure = makePng(320, 180, [76, 111, 227]);
  await save(
    'images.docx',
    new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [new ImageRun({ data: logo, type: 'png', transformation: { width: 120, height: 68 } })],
            }),
            new Paragraph({ text: 'Illustrated Report', heading: HeadingLevel.TITLE }),
            para(lorem.repeat(3)),
            new Paragraph({
              children: [new ImageRun({ data: figure, type: 'png', transformation: { width: 320, height: 180 } })],
            }),
            para('Figure 1. A caption beneath the figure.'),
            para(lorem.repeat(4)),
            new Paragraph({
              children: [
                new TextRun('An image can also sit '),
                new ImageRun({ data: logo, type: 'png', transformation: { width: 36, height: 20 } }),
                new TextRun(' inside a sentence.'),
              ],
            }),
          ],
        },
      ],
    })
  );
}

/* ---------------- 9. a directly formatted resume ----------------------------
 * The shape real documents actually have: no named styles anywhere, headings
 * made out of bold, capitals, size and a rule. This is the case the style
 * mapping table cannot see, so it is what the inference has to earn.
 * ---------------------------------------------------------------------- */
async function directResume() {
  const rule = {
    bottom: { style: BorderStyle.SINGLE, size: 6, space: 1, color: '999999' },
  };
  const line = (text, o = {}) =>
    new Paragraph({
      children: [new TextRun({ text, font: 'Calibri', size: o.size ?? 22, bold: o.bold })],
      ...(o.border ? { border: rule } : {}),
      ...(o.align ? { alignment: o.align } : {}),
    });

  await save(
    'resume-direct.docx',
    new Document({
      numbering: {
        config: [
          {
            reference: 'plain-bullets',
            levels: [
              { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT },
            ],
          },
        ],
      },
      sections: [
        {
          children: [
            line('DANA WHITFIELD', { size: 40, bold: true, align: AlignmentType.CENTER }),
            line('Bristol · dana@example.com · (555) 010-0100', { size: 18 }),
            line('EXPERIENCE', { size: 24, bold: true, border: true }),
            // Partly bold, the usual "Title, Company - dates" idiom.
            new Paragraph({
              children: [
                new TextRun({ text: 'Staff Engineer, Northwind', font: 'Calibri', size: 22, bold: true }),
                new TextRun({ text: ' — 2022 to present', font: 'Calibri', size: 22 }),
              ],
            }),
            new Paragraph({
              numbering: { reference: 'plain-bullets', level: 0 },
              children: [new TextRun({ text: 'Rebuilt the billing pipeline.', font: 'Calibri', size: 22 })],
            }),
            new Paragraph({
              numbering: { reference: 'plain-bullets', level: 0 },
              children: [new TextRun({ text: 'Mentored four engineers.', font: 'Calibri', size: 22 })],
            }),
            line('EDUCATION', { size: 24, bold: true, border: true }),
            line('B.S. Computer Science, State University, 2019'),
            line('SKILLS', { size: 24, bold: true, border: true }),
            line('TypeScript, Go, Postgres, distributed systems, technical writing'),
            line(
              'A closing paragraph long enough that nothing could mistake it for a heading, running past the length where a line stops being a label and starts being prose.'
            ),
          ],
        },
      ],
    })
  );
}

/* ---------------- 10. hand-built: tracked changes + content control ---------
 * The docx library has no API for revision marks, and they are exactly the
 * kind of thing the preservation vault exists for, so this one is written as
 * raw OOXML.
 * ---------------------------------------------------------------------- */
async function trackedChanges() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const body = `
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Reviewed Document</w:t></w:r></w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">The quick </w:t></w:r>
    <w:ins w:id="1" w:author="Reviewer" w:date="2026-01-02T10:00:00Z"><w:r><w:t xml:space="preserve">and nimble </w:t></w:r></w:ins>
    <w:del w:id="2" w:author="Reviewer" w:date="2026-01-02T10:00:00Z"><w:r><w:delText xml:space="preserve">lazy </w:delText></w:r></w:del>
    <w:r><w:t>brown fox.</w:t></w:r>
  </w:p>
  <w:sdt>
    <w:sdtPr><w:alias w:val="Client"/><w:tag w:val="client"/><w:id w:val="99"/><w:text/></w:sdtPr>
    <w:sdtContent><w:p><w:r><w:t>Acme Corporation</w:t></w:r></w:p></w:sdtContent>
  </w:sdt>
  <w:p><w:r><w:t>${'A closing paragraph. '.repeat(10)}</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;

  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
</w:styles>`
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`
  );
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await writeFile(join(OUT, 'tracked-changes.docx'), buf);
  console.log('  tracked-changes.docx  ' + buf.length + ' bytes');
}

await mkdir(OUT, { recursive: true });
console.log('Writing seed corpus to test/corpus/');
await resume();
await coverLetter();
await contract();
await report();
await landscape();
await legal();
await letterhead();
await twoSections();
await directFormat();
await runFormat();
await links();
await withImages();
await directResume();
await trackedChanges();
console.log('Done. Add real .docx files to test/corpus/ - the harness picks them up.');
