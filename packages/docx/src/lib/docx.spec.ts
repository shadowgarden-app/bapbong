import JSZip from 'jszip';
import { Mark } from 'prosemirror-model';
import { importDocx } from './docx';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
// 1x1 transparent PNG.
const PNG_1x1 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}">
  <w:body>
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Hello </w:t></w:r>
      <w:r><w:t>world</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr><w:t>Italic underlined</w:t></w:r>
    </w:p>
    <w:p><w:pPr/></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

/** Build a minimal in-memory .docx (only the parts the importer reads). */
async function makeDocx(
  documentXml: string,
  stylesXml?: string,
  numberingXml?: string,
  relsXml?: string,
  media?: Record<string, string>,
  themeXml?: string,
  parts?: Record<string, string>, // extra zip entries by full path, e.g. word/header1.xml
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file('word/document.xml', documentXml);
  if (stylesXml) zip.file('word/styles.xml', stylesXml);
  if (numberingXml) zip.file('word/numbering.xml', numberingXml);
  if (relsXml) zip.file('word/_rels/document.xml.rels', relsXml);
  if (media) {
    for (const [name, base64] of Object.entries(media)) {
      zip.file(`word/media/${name}`, base64, { base64: true });
    }
  }
  if (themeXml) zip.file('word/theme/theme1.xml', themeXml);
  if (parts) for (const [path, content] of Object.entries(parts)) zip.file(path, content);
  return zip.generateAsync({ type: 'uint8array' });
}

/** A w:p that belongs to list `numId` at indent `ilvl` with the given text. */
function listP(numId: string, ilvl: number, text: string): string {
  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

/** Map a node's marks to `{ name: attrs }` for order-independent assertions. */
function markMap(marks: readonly Mark[]): Record<string, Record<string, unknown>> {
  return Object.fromEntries(marks.map((m) => [m.type.name, m.attrs]));
}

describe('importDocx', () => {
  it('maps paragraphs and runs into the bapbong schema', async () => {
    const { doc, rawDocumentXml } = await importDocx(await makeDocx(DOCUMENT_XML));

    // Two text paragraphs + one empty (<w:p><w:pPr/></w:p>); sectPr is ignored.
    expect(doc.childCount).toBe(3);

    const p0 = doc.child(0);
    expect(p0.textContent).toBe('Hello world');
    expect(p0.child(0).text).toBe('Hello ');
    expect(p0.child(0).marks.map((m) => m.type.name)).toContain('strong');
    expect(p0.child(1).text).toBe('world');
    expect(p0.child(1).marks).toHaveLength(0);

    const p1 = doc.child(1);
    expect(p1.textContent).toBe('Italic underlined');
    expect(p1.child(0).marks.map((m) => m.type.name).sort()).toEqual(['em', 'underline']);

    expect(doc.child(2).childCount).toBe(0);
    expect(rawDocumentXml).toContain('<w:body>');
  });

  it('treats a toggle disabled via w:val="false" as off', async () => {
    const xml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:rPr><w:b w:val="false"/></w:rPr><w:t>not bold</w:t></w:r></w:p></w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(xml));
    expect(doc.child(0).child(0).marks).toHaveLength(0);
  });

  it('resolves the run-property cascade (docDefaults -> style -> inline)', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="48"/></w:rPr></w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>body</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));

    // Title: docDefaults(12pt, Calibri) + Heading1(bold, #FF0000, 24pt).
    const title = markMap(doc.child(0).child(0).marks);
    expect(Object.keys(title).sort()).toEqual(['fontFamily', 'fontSize', 'strong', 'textColor']);
    expect(title.textColor.color).toBe('#FF0000');
    expect(title.fontSize.size).toBe(24);
    expect(title.fontFamily.family).toBe('Calibri');

    // body: docDefaults(12pt, Calibri) + inline italic.
    const body = markMap(doc.child(1).child(0).marks);
    expect(Object.keys(body).sort()).toEqual(['em', 'fontFamily', 'fontSize']);
    expect(body.fontSize.size).toBe(12);
    expect(body.fontFamily.family).toBe('Calibri');
  });

  it('lets inline run properties override the paragraph style', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:styleId="S"><w:rPr><w:b/></w:rPr></w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="S"/></w:pPr><w:r><w:rPr><w:b w:val="false"/></w:rPr><w:t>x</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    expect(doc.child(0).child(0).marks).toHaveLength(0);
  });

  it('numbers list paragraphs with multilevel counters', async () => {
    const numberingXml = `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}">
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:start w:val="1"/></w:lvl>
        <w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2"/><w:start w:val="1"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    </w:numbering>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${listP('1', 0, 'first')}
      ${listP('1', 0, 'second')}
      ${listP('1', 1, 'sub a')}
      ${listP('1', 1, 'sub b')}
      ${listP('1', 0, 'third')}
      <w:p><w:r><w:t>plain</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml, undefined, numberingXml));

    const markers: (string | null)[] = [];
    doc.forEach((node) => markers.push((node.attrs.list as { marker: string } | null)?.marker ?? null));
    expect(markers).toEqual(['1.', '2.', '2.a', '2.b', '3.', null]);
    expect(doc.child(0).attrs.list).toMatchObject({ numId: '1', level: 0 });
    expect(doc.child(5).attrs.list).toBeNull();
  });

  it('formats bullet and roman markers, with independent counters per numId', async () => {
    const numberingXml = `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}">
      <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>
      <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/><w:start w:val="1"/></w:lvl></w:abstractNum>
      <w:num w:numId="10"><w:abstractNumId w:val="0"/></w:num>
      <w:num w:numId="11"><w:abstractNumId w:val="1"/></w:num>
    </w:numbering>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${listP('10', 0, 'bullet')}
      ${listP('11', 0, 'roman one')}
      ${listP('11', 0, 'roman two')}
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml, undefined, numberingXml));
    const marker = (i: number) => (doc.child(i).attrs.list as { marker: string }).marker;
    expect(marker(0)).toBe('•');
    expect(marker(1)).toBe('I)');
    expect(marker(2)).toBe('II)');
  });

  it('imports tables in document order with colspan and column widths', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>before</w:t></w:r></w:p>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="1440"/></w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>spanned</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
      <w:p><w:r><w:t>after</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));

    // Block order (paragraph, table, paragraph) is preserved.
    expect(doc.childCount).toBe(3);
    expect(doc.child(0).textContent).toBe('before');
    expect(doc.child(2).textContent).toBe('after');

    const table = doc.child(1);
    expect(table.type.name).toBe('table');
    expect(table.childCount).toBe(2);

    const row0 = table.child(0);
    expect(row0.childCount).toBe(2);
    expect(row0.child(0).type.name).toBe('table_cell');
    expect(row0.child(0).textContent).toBe('A1');
    expect(row0.child(0).attrs.colspan).toBe(1);
    expect(row0.child(0).attrs.colwidth).toEqual([192]); // 2880 twips / 15
    expect(row0.child(1).attrs.colwidth).toEqual([96]); // 1440 / 15

    const merged = table.child(1).child(0);
    expect(merged.attrs.colspan).toBe(2);
    expect(merged.attrs.colwidth).toEqual([192, 96]);
    expect(merged.textContent).toBe('spanned');
  });

  it('collapses vertical merges (w:vMerge) into rowspan', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>L</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>R1</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
          <w:tc><w:p><w:r><w:t>R2</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p/></w:tc>
          <w:tc><w:p><w:r><w:t>R3</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    const table = doc.child(0);
    expect(table.childCount).toBe(3); // three rows

    // Left column: one cell spanning all three rows.
    const topLeft = table.child(0).child(0);
    expect(topLeft.textContent).toBe('L');
    expect(topLeft.attrs.rowspan).toBe(3);
    expect(table.child(0).childCount).toBe(2);

    // Rows 2 and 3 only keep their right-column cell (continue cells dropped).
    expect(table.child(1).childCount).toBe(1);
    expect(table.child(1).child(0).textContent).toBe('R2');
    expect(table.child(1).child(0).attrs.rowspan).toBe(1);
    expect(table.child(2).childCount).toBe(1);
    expect(table.child(2).child(0).textContent).toBe('R3');
  });

  it('indents list items per numbering level (w:lvl/w:pPr/w:ind)', async () => {
    const numberingXml = `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}">
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:start w:val="1"/></w:lvl>
        <w:lvl w:ilvl="1"><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2."/><w:start w:val="1"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    </w:numbering>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${listP('1', 0, 'top')}
      ${listP('1', 1, 'nested')}
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="2160"/></w:pPr><w:r><w:t>override</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, undefined, numberingXml));

    // The lvl pPr indents flow through the cascade: 720tw=48px, 1440tw=96px.
    expect(doc.child(0).attrs.indent).toEqual({ left: 48, hanging: 24 });
    expect(doc.child(1).attrs.indent).toEqual({ left: 96, hanging: 24 });
    // Inline w:ind overrides the numbering layer per attribute.
    expect(doc.child(2).attrs.indent).toEqual({ left: 144, hanging: 24 });
  });

  it('cascades w:jc and w:ind from paragraph styles (pPr)', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:docDefaults><w:pPrDefault><w:pPr><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
      <w:style w:styleId="Base"><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
      <w:style w:styleId="Quote"><w:basedOn w:val="Base"/><w:pPr><w:jc w:val="center"/><w:ind w:firstLine="360"/></w:pPr></w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Quote"/><w:jc w:val="right"/><w:ind w:hanging="240"/></w:pPr><w:r><w:t>inline wins</w:t></w:r></w:p>
      <w:p><w:r><w:t>defaults only</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));

    // p0: style chain — Base gives ind.left (720tw=48px), Quote adds center + firstLine (360tw=24px).
    expect(doc.child(0).attrs.align).toBe('center');
    expect(doc.child(0).attrs.indent).toEqual({ left: 48, firstLine: 24 });
    // p1: inline overrides — jc right; hanging replaces the style's firstLine.
    expect(doc.child(1).attrs.align).toBe('right');
    expect(doc.child(1).attrs.indent).toEqual({ left: 48, hanging: 16 });
    // p2: docDefaults pPrDefault applies when nothing else does.
    expect(doc.child(2).attrs.align).toBe('justify');
  });

  it('maps PAGE/NUMPAGES fields to page_field atoms, keeps other fields as text', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:t xml:space="preserve">Trang </w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:rPr><w:b/></w:rPr><w:t>3</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:t xml:space="preserve"> / </w:t></w:r>
        <w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple>
      </w:p>
      <w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> DATE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>11/06/2026</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));

    const p0 = doc.child(0);
    expect(p0.child(0).text).toBe('Trang ');
    expect(p0.child(1).type.name).toBe('page_field');
    expect(p0.child(1).attrs.kind).toBe('page');
    expect(markMap(p0.child(1).marks).strong).toBeDefined(); // result-run formatting kept
    expect(p0.child(2).text).toBe(' / ');
    expect(p0.child(3).type.name).toBe('page_field');
    expect(p0.child(3).attrs.kind).toBe('pages');
    // Unknown instruction (DATE) falls back to its cached result text.
    expect(doc.child(1).textContent).toBe('11/06/2026');
  });

  it('parses w:tab run elements and w:tabs stop definitions', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:styleId="TOC"><w:pPr><w:tabs>
        <w:tab w:val="right" w:pos="9000" w:leader="dot"/>
        <w:tab w:val="bar" w:pos="450"/>
        <w:tab w:val="clear" w:pos="900"/>
      </w:tabs></w:pPr></w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="TOC"/></w:pPr>
        <w:r><w:t>Chương 1</w:t><w:tab/><w:t>5</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const p = doc.child(0);
    expect(p.textContent).toBe('Chương 1\t5'); // <w:tab/> → \t, in order
    expect(p.attrs.tabs).toEqual([{ pos: 600, val: 'right', leader: 'dot' }]); // bar/clear dropped
  });

  it('imports wp:anchor drawings as floating images', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"><w:body>
      <w:p><w:r><w:drawing><wp:anchor distL="114300" distR="114300" distT="0" distB="0">
        <wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV>
        <wp:extent cx="952500" cy="476250"/>
        <wp:wrapSquare wrapText="bothSides"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:anchor></w:drawing></w:r></w:p>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/></Relationships>`;
    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, { 'image1.png': PNG_1x1 }),
    );
    const img = doc.child(0).child(0);
    expect(img.type.name).toBe('image');
    expect(img.attrs.width).toBe(100);
    expect(img.attrs.float).toEqual({
      wrap: 'square',
      hAlign: 'right',
      hRel: 'margin',
      vOffset: 20, // 190500 EMU
      vRel: 'paragraph',
      distL: 12,
      distR: 12,
      distT: 0,
      distB: 0,
    });
  });

  it('parses table border visibility (direct w:tblBorders and via table style)', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:styleId="TableGrid"><w:tblPr><w:tblBorders>
        <w:top w:val="single"/><w:bottom w:val="single"/><w:left w:val="single"/>
        <w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/>
      </w:tblBorders></w:tblPr></w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblPr><w:tblBorders><w:top w:val="single"/><w:insideH w:val="nil"/></w:tblBorders></w:tblPr>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>direct</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
      <w:tbl>
        <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>styled</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>bare</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    expect(doc.child(0).attrs.borders).toEqual({ top: true, insideH: false });
    expect(doc.child(1).attrs.borders).toEqual({
      top: true,
      bottom: true,
      left: true,
      right: true,
      insideH: true,
      insideV: true,
    });
    expect(doc.child(2).attrs.borders).toBeNull(); // borderless by default
  });

  it('parses per-table cell margins (w:tblCellMar)', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblPr><w:tblCellMar>
          <w:left w:w="300" w:type="dxa"/>
          <w:right w:w="150" w:type="dxa"/>
          <w:top w:w="0" w:type="nil"/>
        </w:tblCellMar></w:tblPr>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    expect(doc.child(0).attrs.cellPadding).toEqual({ left: 20, right: 10, top: 0 });
    expect(doc.child(1).attrs.cellPadding).toBeNull(); // no override → defaults
  });

  it('marks w:tblHeader rows as header rows', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>head</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:trPr><w:tblHeader w:val="false"/></w:trPr><w:tc><w:p><w:r><w:t>off</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p><w:r><w:t>body</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const table = doc.child(0);
    expect(table.child(0).attrs.header).toBe(true);
    expect(table.child(1).attrs.header).toBe(false); // explicit w:val="false"
    expect(table.child(2).attrs.header).toBe(false);
  });

  it('applies link marks from hyperlinks resolved via relationships', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
      <w:p>
        <w:hyperlink r:id="rId5"><w:r><w:t>click here</w:t></w:r></w:hyperlink>
        <w:r><w:t xml:space="preserve"> and plain</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId5" Type="${R_NS}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`;

    const { doc } = await importDocx(await makeDocx(documentXml, undefined, undefined, relsXml));
    const p = doc.child(0);
    expect(p.child(0).text).toBe('click here');
    expect(markMap(p.child(0).marks).link.href).toBe('https://example.com/');
    expect(p.child(1).text).toBe(' and plain');
    expect(p.child(1).marks).toHaveLength(0);
  });

  it('imports inline images as image nodes with data-URL src and px size', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"><w:body>
      <w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="952500" cy="476250"/>
        <wp:docPr id="1" name="Picture 1" descr="a cat"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/></Relationships>`;

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, { 'image1.png': PNG_1x1 }),
    );
    const img = doc.child(0).child(0);
    expect(img.type.name).toBe('image');
    expect(img.attrs.src).toBe(`data:image/png;base64,${PNG_1x1}`);
    expect(img.attrs.width).toBe(100); // 952500 EMU / 9525
    expect(img.attrs.height).toBe(50); // 476250 / 9525
    expect(img.attrs.alt).toBe('a cat');
  });

  it('resolves theme colors (w:themeColor) via theme1.xml, incl. shade', async () => {
    const themeXml = `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}"><a:themeElements><a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme></a:themeElements></a:theme>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:rPr><w:color w:themeColor="accent1"/></w:rPr><w:t>accent</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:color w:themeColor="accent1" w:themeShade="80"/></w:rPr><w:t>shaded</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, undefined, undefined, themeXml),
    );
    expect(markMap(doc.child(0).child(0).marks).textColor.color).toBe('#4472C4');
    // shade 0x80/255 ≈ 0.502 → 4472C4 darkened ≈ 223962
    expect(markMap(doc.child(1).child(0).marks).textColor.color).toBe('#223962');
  });

  it('parses header/footer parts referenced by sectPr', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
      <w:p><w:r><w:t>body</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference w:type="default" r:id="rIdH"/>
        <w:footerReference w:type="default" r:id="rIdF"/>
      </w:sectPr>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}">
      <Relationship Id="rIdH" Type="${R_NS}/header" Target="header1.xml"/>
      <Relationship Id="rIdF" Type="${R_NS}/footer" Target="footer1.xml"/>
    </Relationships>`;
    const headerXml = `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Header text</w:t></w:r></w:p></w:hdr>`;
    const footerXml = `<?xml version="1.0"?><w:ftr xmlns:w="${W_NS}"><w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>`;

    const { doc, headers, footers } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, undefined, undefined, {
        'word/header1.xml': headerXml,
        'word/footer1.xml': footerXml,
      }),
    );

    expect(doc.textContent).toBe('body');
    expect(headers.default.textContent).toBe('Header text');
    expect(headers.default.child(0).child(0).marks.map((m) => m.type.name)).toContain('strong');
    expect(footers.default.textContent).toBe('Footer text');
  });

  it('maps w:jc to paragraph alignment', async () => {
    const xml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>c</w:t></w:r></w:p>
      <w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>j</w:t></w:r></w:p>
      <w:p><w:pPr><w:jc w:val="end"/></w:pPr><w:r><w:t>r</w:t></w:r></w:p>
      <w:p><w:r><w:t>default</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(xml));
    expect(doc.child(0).attrs['align']).toBe('center');
    expect(doc.child(1).attrs['align']).toBe('justify');
    expect(doc.child(2).attrs['align']).toBe('right');
    expect(doc.child(3).attrs['align']).toBeNull();
  });

  it('maps w:ind (twips) to indentation in px, hanging winning over firstLine', async () => {
    const xml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:ind w:left="720" w:right="360" w:hanging="240" w:firstLine="480"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
      <w:p><w:pPr><w:ind w:start="1440" w:firstLine="240"/></w:pPr><w:r><w:t>y</w:t></w:r></w:p>
      <w:p><w:r><w:t>none</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(xml));
    // 720/15=48, 360/15=24, 240/15=16; hanging present so firstLine dropped.
    expect(doc.child(0).attrs['indent']).toEqual({ left: 48, right: 24, hanging: 16 });
    // w:start aliases w:left; 1440/15=96, 240/15=16.
    expect(doc.child(1).attrs['indent']).toEqual({ left: 96, firstLine: 16 });
    expect(doc.child(2).attrs['indent']).toBeNull();
  });

  it('throws when word/document.xml is missing', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'nope');
    await expect(importDocx(await zip.generateAsync({ type: 'uint8array' }))).rejects.toThrow(
      /document\.xml/,
    );
  });
});
