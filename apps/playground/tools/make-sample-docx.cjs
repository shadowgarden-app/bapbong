// Generates apps/playground/public/sample.docx — a small document that
// exercises the import + layout pipeline end to end: marks, alignment,
// first-line indent, numbered list, table (incl. colspan), inline image.
//
//   node apps/playground/tools/make-sample-docx.cjs
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

// 1x1 red PNG; drawn at 64x64 via wp:extent (64px × 9525 EMU/px = 609600).
const PNG_RED =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const LOREM =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure ' +
  'dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.';

const listP = (numId, ilvl, text) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

const td = (text, extraTcPr = '') =>
  `<w:tc><w:tcPr>${extraTcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>bapbong sample document</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="both"/></w:pPr>
      <w:r><w:t>${LOREM}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="720"/></w:pPr>
      <w:r><w:t xml:space="preserve">This paragraph has a first-line indent, plus </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:t xml:space="preserve">, </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:r><w:rPr><w:color w:val="C0392B"/></w:rPr><w:t>colored</w:t></w:r>
      <w:r><w:t xml:space="preserve"> runs.</w:t></w:r>
    </w:p>
    ${listP('1', 0, 'First numbered item')}
    ${listP('1', 0, 'Second numbered item wraps onto a continuation line when it gets long enough to exceed the content width')}
    ${listP('1', 0, 'Third numbered item')}
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="4515"/><w:gridCol w:w="4515"/></w:tblGrid>
      <w:tr>${td('Cell A1')}${td('Cell B1 with a bit more text so the row grows taller than its neighbour')}</w:tr>
      <w:tr>${td('A merged cell spanning both columns', '<w:gridSpan w:val="2"/>')}</w:tr>
    </w:tbl>
    <w:p>
      <w:r><w:t xml:space="preserve">Inline image: </w:t></w:r>
      <w:r><w:drawing><wp:inline>
        <wp:extent cx="609600" cy="609600"/>
        <wp:docPr id="1" name="Picture 1" descr="red square"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r>
      <w:r><w:t xml:space="preserve"> sits on the baseline.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const NUMBERING_XML = `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:start w:val="1"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

const RELS_XML = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/></Relationships>`;

async function main() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/_rels/document.xml.rels', RELS_XML);
  zip.file('word/media/image1.png', PNG_RED, { base64: true });

  const bytes = await zip.generateAsync({ type: 'nodebuffer' });
  const out = path.join(__dirname, '..', 'public', 'sample.docx');
  fs.writeFileSync(out, bytes);
  console.log(`wrote ${out} (${bytes.length} bytes)`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
