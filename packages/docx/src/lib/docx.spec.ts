import JSZip from 'jszip';
import { Mark, Schema } from 'prosemirror-model';
import {
  bookmarkLabel,
  createNumberingCounter,
  fieldAt,
  findBookmark,
  schema,
  type NumberingDefs,
} from '@shadow-garden/bapbong-model';
import { importDocx } from './docx';
import { DocxImportError, IMPORT_ERROR_MESSAGES, sniffDocx } from './sniff';
import { buildEncryptedDocx, importFailure } from './crypto-docx.spec-helper';

// The comment mark lives in the comment plugin, not the base schema. Comment
// import tests compose a schema carrying it (minimal local spec — no docx→plugin
// dependency) and import against it; otherwise comment values are filtered out.
const withComments = new Schema({
  nodes: schema.spec.nodes,
  marks: schema.spec.marks.append({ comment: { attrs: { ids: {} } } }),
});

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const R_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const WP_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
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
  if (parts)
    for (const [path, content] of Object.entries(parts))
      zip.file(path, content);
  return zip.generateAsync({ type: 'uint8array' });
}

/** A w:p that belongs to list `numId` at indent `ilvl` with the given text. */
function listP(numId: string, ilvl: number, text: string): string {
  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

/** Map a node's marks to `{ name: attrs }` for order-independent assertions. */
function markMap(
  marks: readonly Mark[],
): Record<string, Record<string, unknown>> {
  return Object.fromEntries(marks.map((m) => [m.type.name, m.attrs]));
}

describe('importDocx', () => {
  it('maps paragraphs and runs into the bapbong schema', async () => {
    const { doc, rawDocumentXml } = await importDocx(
      await makeDocx(DOCUMENT_XML),
    );

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
    expect(
      p1
        .child(0)
        .marks.map((m) => m.type.name)
        .sort(),
    ).toEqual(['em', 'underline']);

    expect(doc.child(2).childCount).toBe(0);
    expect(rawDocumentXml).toContain('<w:body>');
  });

  it('keeps numeric-looking text verbatim (no strnum mangling)', async () => {
    // Word splits "1.500.000" across runs (rsid); each fragment must survive
    // untouched — parseTagValue would turn "1." → 1, "00" → 0, "100.000" → 100.
    const runs = ['1.', '5', '00', '.000']
      .map((t) => `<w:r><w:t>${t}</w:t></w:r>`)
      .join('');
    const xml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>${runs}</w:p><w:p><w:r><w:t>100.000</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(xml));
    expect(doc.child(0).textContent).toBe('1.500.000');
    expect(doc.child(1).textContent).toBe('100.000');
  });

  it('unwraps w:sdt content controls (block + inline checkbox)', async () => {
    const W14_CHK = 'http://schemas.microsoft.com/office/word/2010/wordml';
    // Block-level control (how Word wraps cover pages) holding a paragraph
    // and a table; inline control with a w14:checkbox whose glyph run is the
    // usual MS-Gothic ☒; nested control inside the block one.
    const xml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:w14="${W14_CHK}"><w:body>
      <w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>
        <w:p><w:r><w:t>Cover title</w:t></w:r></w:p>
        <w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>
          <w:p><w:r><w:t>Nested subtitle</w:t></w:r></w:p>
        </w:sdtContent></w:sdt>
        <w:tbl><w:tr><w:tc><w:p><w:r><w:t>In table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      </w:sdtContent></w:sdt>
      <w:p>
        <w:sdt><w:sdtPr><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr><w:sdtContent>
          <w:r><w:rPr><w:rFonts w:ascii="MS Gothic"/></w:rPr><w:t>☒</w:t></w:r>
        </w:sdtContent></w:sdt>
        <w:r><w:t xml:space="preserve"> Đồng ý</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(xml));

    expect(doc.child(0).textContent).toBe('Cover title');
    expect(doc.child(1).textContent).toBe('Nested subtitle');
    expect(doc.child(2).type.name).toBe('table');
    expect(doc.child(2).textContent).toBe('In table');
    expect(doc.child(3).textContent).toBe('☒ Đồng ý');
  });

  it('flattens OMML equations to readable text runs', async () => {
    const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
    const mr = (t: string, rPr = '') => `<m:r>${rPr}<m:t>${t}</m:t></m:r>`;
    const xml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:m="${M_NS}"><w:body>
      <w:p><m:oMathPara><m:oMath>
        <m:sSub><m:e>${mr('t', '<w:rPr><w:i/></w:rPr>')}</m:e><m:sub>${mr('1')}</m:sub></m:sSub>
      </m:oMath></m:oMathPara></w:p>
      <w:p>
        <w:r><w:t xml:space="preserve">S = </w:t></w:r>
        <m:oMath>
          <m:sSup><m:e>${mr('x')}</m:e><m:sup>${mr('2')}</m:sup></m:sSup>
          ${mr('+')}
          <m:f><m:num>${mr('a')}</m:num><m:den>${mr('b')}</m:den></m:f>
          ${mr('+')}
          <m:f><m:num><m:r><m:t>a+b</m:t></m:r></m:num><m:den>${mr('c')}</m:den></m:f>
          ${mr('+')}
          <m:rad><m:deg/><m:e>${mr('y')}</m:e></m:rad>
          ${mr('+')}
          <m:d><m:e>${mr('u+v')}</m:e></m:d>
        </m:oMath>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(xml));

    // Subscripted digits become Unicode subscripts; the run is formatted like
    // the equation's first math run (italic here).
    expect(doc.child(0).textContent).toBe('t₁');
    expect(
      doc
        .child(0)
        .child(0)
        .marks.map((m) => m.type.name),
    ).toContain('em');

    // Inline equation after plain text: sup, fractions (multi-term numerator
    // gets parens), radical, delimiter.
    expect(doc.child(1).textContent).toBe('S = x²+a/b+(a+b)/c+√(y)+(u+v)');
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
    expect(Object.keys(title).sort()).toEqual([
      'fontFamily',
      'fontSize',
      'strong',
      'textColor',
    ]);
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

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, numberingXml),
    );

    // Markers are no longer frozen at import — they're recounted from the
    // defs riding the doc (this is what lets edits renumber live).
    const counter = createNumberingCounter(
      doc.attrs.numbering as NumberingDefs,
    );
    const markers: (string | null)[] = [];
    doc.forEach((node) => {
      const list = node.attrs.list as { numId: string; level: number } | null;
      markers.push(list ? counter.next(list.numId, list.level) : null);
    });
    expect(markers).toEqual(['1.', '2.', '2.a', '2.b', '3.', null]);
    expect(doc.child(0).attrs.list).toEqual({ numId: '1', level: 0 });
    expect(doc.child(5).attrs.list).toBeNull();
  });

  it('applies w:lvlOverride / w:startOverride per numId', async () => {
    const numberingXml = `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}">
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:start w:val="1"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      <w:num w:numId="2"><w:abstractNumId w:val="0"/>
        <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>
      </w:num>
    </w:numbering>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${listP('1', 0, 'one')}
      ${listP('2', 0, 'five')}
    </w:body></w:document>`;

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, numberingXml),
    );
    const defs = doc.attrs.numbering as NumberingDefs;
    // The overridden num restarts at 5 and counts independently of numId 1.
    expect(defs['2'].levels[0].start).toBe(5);
    expect(defs['2'].key).not.toBe(defs['1'].key);
    const counter = createNumberingCounter(defs);
    expect(counter.next('1', 0)).toBe('1.');
    expect(counter.next('2', 0)).toBe('5.');
  });

  it('links paragraph styles to their numbering level via lvl w:pStyle', async () => {
    // Numbered heading styles: the style's numPr names only the numId; the
    // LEVEL comes from whichever lvl claims the style with w:pStyle.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:styleId="Heading2">
        <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>
      </w:style>
    </w:styles>`;
    const numberingXml = `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}">
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:start w:val="1"/><w:pStyle w:val="Heading1"/></w:lvl>
        <w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:start w:val="1"/><w:pStyle w:val="Heading2"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    </w:numbering>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>linked heading</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading2"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>explicit ilvl wins</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(
      await makeDocx(documentXml, stylesXml, numberingXml),
    );
    // No written ilvl → the lvl>pStyle link picks level 1 (not the default 0).
    expect(doc.child(0).attrs.list).toEqual({ numId: '1', level: 1 });
    // A written w:ilvl is direct intent and beats the link.
    expect(doc.child(1).attrs.list).toEqual({ numId: '1', level: 0 });
  });

  it('parses label styling (w:lvlJc / w:suff / w:isLgl / lvl rPr) into the defs', async () => {
    const numberingXml = `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}">
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0">
          <w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/><w:start w:val="1"/>
          <w:lvlJc w:val="right"/><w:suff w:val="space"/>
          <w:rPr><w:b/><w:color w:val="C00000"/><w:rFonts w:ascii="Georgia"/><w:sz w:val="20"/></w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="1">
          <w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:start w:val="1"/><w:isLgl/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    </w:numbering>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${listP('1', 0, 'one')}
      ${listP('1', 1, 'sub')}
    </w:body></w:document>`;

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, numberingXml),
    );
    const defs = doc.attrs.numbering as NumberingDefs;
    expect(defs['1'].levels[0]).toMatchObject({
      jc: 'right',
      suff: 'space',
      rPr: { bold: true, color: '#C00000', family: 'Georgia', sizePt: 10 },
    });
    expect(defs['1'].levels[1].isLgl).toBe(true);
    // Legal numbering: the level-1 placeholder %1 renders decimal even
    // though level 0 is upperRoman.
    const counter = createNumberingCounter(defs);
    expect(counter.next('1', 0)).toBe('I.');
    expect(counter.next('1', 1)).toBe('1.1');
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

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, numberingXml),
    );
    const counter = createNumberingCounter(
      doc.attrs.numbering as NumberingDefs,
    );
    const marker = (i: number) => {
      const list = doc.child(i).attrs.list as { numId: string; level: number };
      return counter.next(list.numId, list.level);
    };
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

  it('resolves percentage table widths against the page content width', async () => {
    // Seen in the wild: w:tblW/w:tcW type="pct" with a PLACEHOLDER tblGrid
    // (100 twips ≈ 7px per column) — taking the grid literally stacked one
    // character per line. A4 default content width = 794 − 96 − 96 = 602px.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblPr><w:tblW w:type="pct" w:w="100%"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:tcW w:type="pct" w:w="30%"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>
          <w:tc><w:tcPr><w:tcW w:type="pct" w:w="70%"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const row = doc.child(0).child(0);
    expect(row.child(0).attrs.colwidth).toEqual([181]); // 30% of 602
    expect(row.child(1).attrs.colwidth).toEqual([421]); // 70% of 602
  });

  it('imports w:pBdr paragraph borders and w:tcMar cell margins', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:pPr><w:pBdr><w:top w:val="single" w:color="cccccc" w:sz="12"/></w:pBdr></w:pPr>
        <w:r><w:t>ruled</w:t></w:r>
      </w:p>
      <w:tbl><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid><w:tr><w:tc>
        <w:tcPr><w:tcMar><w:top w:type="dxa" w:w="60"/><w:left w:type="dxa" w:w="75"/></w:tcMar></w:tcPr>
        <w:p><w:r><w:t>c</w:t></w:r></w:p>
      </w:tc></w:tr></w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    expect(doc.child(0).attrs.borders).toEqual({
      top: { width: 2, style: 'solid', color: '#CCCCCC' }, // sz 12 → 1.5pt → 2px
    });
    const cell = doc.child(1).child(0).child(0);
    expect(cell.attrs.padding).toEqual({ top: 4, left: 5 }); // 60tw, 75tw
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
    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, numberingXml),
    );

    // The lvl pPr indents flow through the cascade: 720tw=48px, 1440tw=96px.
    expect(doc.child(0).attrs.indent).toEqual({ left: 48, hanging: 24 });
    expect(doc.child(1).attrs.indent).toEqual({ left: 96, hanging: 24 });
    // Inline w:ind overrides the numbering layer per attribute.
    expect(doc.child(2).attrs.indent).toEqual({ left: 144, hanging: 24 });
  });

  it('parses w:vertAlign as super/subscript marks', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:t>E=mc</w:t></w:r>
        <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r>
        <w:r><w:t>, H</w:t></w:r>
        <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r>
        <w:r><w:t>O</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const p = doc.child(0);
    expect(markMap(p.child(1).marks).vertAlign.value).toBe('super');
    expect(markMap(p.child(3).marks).vertAlign.value).toBe('sub');
  });

  it('parses highlight/shading on runs and w:shd cell fills', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>hi</w:t></w:r>
        <w:r><w:rPr><w:shd w:fill="C0E0FF"/></w:rPr><w:t>shaded</w:t></w:r>
      </w:p>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:tcPr><w:shd w:fill="D9E2F3"/></w:tcPr><w:p><w:r><w:t>head</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const p = doc.child(0);
    expect(markMap(p.child(0).marks).highlight.color).toBe('#FFFF00'); // named → hex
    expect(markMap(p.child(1).marks).highlight.color).toBe('#C0E0FF'); // w:shd fill
    expect(doc.child(1).child(0).child(0).attrs.background).toBe('#D9E2F3'); // cell fill
  });

  it('maps w:br to hard_break nodes and page breaks to pageBreakBefore', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>line1</w:t><w:br/><w:t>line2</w:t></w:r></w:p>
      <w:p><w:r><w:br w:type="page"/><w:t>next page</w:t></w:r></w:p>
      <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>also new</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const p0 = doc.child(0);
    expect(p0.childCount).toBe(3); // text, hard_break, text
    expect(p0.child(1).type.name).toBe('hard_break');
    expect(p0.child(0).text).toBe('line1');
    expect(p0.child(2).text).toBe('line2');
    expect(doc.child(1).attrs.pageBreakBefore).toBe(true); // w:br type=page
    expect(doc.child(2).attrs.pageBreakBefore).toBe(true); // w:pageBreakBefore
  });

  it('honors w:pageBreakBefore toggle values (val="false"/"0" = no break)', async () => {
    // An inline w:val="false" cancels the break — including one inherited
    // from a style layer (the override is the LAST cascade layer with the
    // element, and its value must be read, not just its presence).
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:styleId="Breaky"><w:pPr><w:pageBreakBefore/></w:pPr></w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pageBreakBefore w:val="false"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>
      <w:p><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>
      <w:p><w:pPr><w:pageBreakBefore w:val="true"/></w:pPr><w:r><w:t>c</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Breaky"/><w:pageBreakBefore w:val="false"/></w:pPr><w:r><w:t>d</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Breaky"/></w:pPr><w:r><w:t>e</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    expect(doc.child(0).attrs.pageBreakBefore).toBeFalsy(); // val=false
    expect(doc.child(1).attrs.pageBreakBefore).toBeFalsy(); // val=0
    expect(doc.child(2).attrs.pageBreakBefore).toBe(true); // val=true
    expect(doc.child(3).attrs.pageBreakBefore).toBeFalsy(); // inline false beats style
    expect(doc.child(4).attrs.pageBreakBefore).toBe(true); // style alone
  });

  it('imports page size and margins from w:sectPr', async () => {
    // US Letter (12240×15840 twips) landscape, 0.5in (720tw) margins.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="12240" w:h="15840" w:orient="landscape"/>
        <w:pgMar w:top="720" w:right="1080" w:bottom="720" w:left="1080"/>
      </w:sectPr>
    </w:body></w:document>`;
    const { page, doc } = await importDocx(await makeDocx(documentXml));
    expect(page).toEqual({
      width: 1056, // 15840tw → 1056px (landscape swap)
      height: 816, // 12240tw → 816px
      margin: { top: 48, right: 72, bottom: 48, left: 72 },
    });
    // The same geometry rides the doc (doc.attrs.page) so page-setup
    // commands can edit it through dispatch/undo.
    expect(doc.attrs.page).toEqual(page);
  });

  it('parses unit-suffixed pgMar values and never yields NaN margins', async () => {
    // Seen in the wild from non-Word producers: w:pgMar w:top="20pt". A plain
    // Number() made these NaN, and NaN margins hung pagination forever.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="20pt" w:right="40pt" w:bottom="20pt" w:left="1in"/>
      </w:sectPr>
    </w:body></w:document>`;
    const { page } = await importDocx(await makeDocx(documentXml));
    expect(page.margin).toEqual({ top: 27, right: 53, bottom: 27, left: 96 });
    // Garbage falls back to the default instead of NaN.
    const bad = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr><w:pgMar w:top="wide" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>
    </w:body></w:document>`;
    const { page: page2 } = await importDocx(await makeDocx(bad));
    expect(page2.margin.top).toBe(96);
    expect(page2.margin.right).toBe(48);
  });

  it('imports per-section page geometry overrides (mixed portrait/landscape)', async () => {
    // Section 1 portrait Letter (differs from body → override), body section
    // landscape Letter (the document default — no override).
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:sectPr>
        <w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      </w:sectPr></w:pPr><w:r><w:t>portrait part</w:t></w:r></w:p>
      <w:p><w:r><w:t>landscape part</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      </w:sectPr>
    </w:body></w:document>`;
    const { doc, page } = await importDocx(await makeDocx(documentXml));
    expect(page).toMatchObject({ width: 1056, height: 816 }); // body: landscape
    const sections = doc.attrs.sections as {
      blockCount: number;
      page?: { width: number; height: number };
    }[];
    expect(sections).toHaveLength(2);
    expect(sections[0].page).toMatchObject({ width: 816, height: 1056 }); // override
    expect(sections[1].page).toBeUndefined(); // body geometry — no override
  });

  it('drops per-section geometry matching the document default', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:cols w:num="2" w:space="425"/>
      </w:sectPr></w:pPr><w:r><w:t>two cols</w:t></w:r></w:p>
      <w:p><w:r><w:t>one col</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const sections = doc.attrs.sections as { page?: unknown }[];
    expect(sections).toHaveLength(2);
    expect(sections[0].page).toBeUndefined(); // same geometry → no override
    expect(sections[1].page).toBeUndefined();
  });

  it('survives a PAGE field packed into a single run (Google Docs shape)', async () => {
    // begin + instrText + separate + end all in ONE w:r, the visible text in
    // the NEXT run. The old per-run state machine left the field open and
    // swallowed the rest of the paragraph.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve">PAGE</w:instrText><w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:t xml:space="preserve"> after the field</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const p = doc.child(0);
    expect(p.textContent).toBe(' after the field');
    // The PAGE field itself is materialized as a page_field node.
    let fields = 0;
    p.forEach((n) => {
      if (n.type.name === 'page_field') fields++;
    });
    expect(fields).toBe(1);
  });

  it('collects per-section headers with Link-to-Previous inheritance', async () => {
    const hdr = (t: string) =>
      `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:hdr>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdH2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/></Relationships>`;
    // Section 1: own header + titlePg. Section 2: NOTHING (inherits section
    // 1's — Word's "Link to Previous"). Section 3 (body): its own header.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
      <w:p><w:pPr><w:sectPr>
        <w:headerReference w:type="default" r:id="rIdH1"/><w:titlePg/>
      </w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p>
      <w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>
      <w:p><w:r><w:t>three</w:t></w:r></w:p>
      <w:sectPr><w:headerReference w:type="default" r:id="rIdH2"/></w:sectPr>
    </w:body></w:document>`;
    const { sectionChrome, headers } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        relsXml,
        undefined,
        undefined,
        {
          'word/header1.xml': hdr('CHAPTER ONE'),
          'word/header2.xml': hdr('CHAPTER TWO'),
        },
      ),
    );
    expect(sectionChrome).toHaveLength(3);
    expect(sectionChrome?.[0].headers['default']?.textContent).toBe(
      'CHAPTER ONE',
    );
    expect(sectionChrome?.[0].titlePg).toBe(true);
    // Link to Previous: section 2 shows section 1's header, but not titlePg.
    expect(sectionChrome?.[1].headers['default']?.textContent).toBe(
      'CHAPTER ONE',
    );
    expect(sectionChrome?.[1].titlePg).toBe(false);
    // Section 3 overrides with its own.
    expect(sectionChrome?.[2].headers['default']?.textContent).toBe(
      'CHAPTER TWO',
    );
    // Flat fields = the last section's resolved chrome (legacy shape).
    expect(headers['default']?.textContent).toBe('CHAPTER TWO');
  });

  it('omits sectionChrome when only the body sectPr declares chrome', async () => {
    const hdr = `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
      <w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>one</w:t></w:r></w:p>
      <w:p><w:r><w:t>two</w:t></w:r></w:p>
      <w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>
    </w:body></w:document>`;
    const { sectionChrome, headers } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        relsXml,
        undefined,
        undefined,
        {
          'word/header1.xml': hdr,
        },
      ),
    );
    // Uniform chrome — the flat fields cover every page; no per-section set.
    expect(sectionChrome).toBeUndefined();
    expect(headers['default']?.textContent).toBe('H');
  });

  it('defaults page geometry to A4 when sectPr omits it', async () => {
    const { page } = await importDocx(await makeDocx(DOCUMENT_XML));
    expect(page).toEqual({
      width: 794,
      height: 1123,
      margin: { top: 96, right: 96, bottom: 96, left: 96 },
    });
  });

  it('parses w:spacing (before/after + line rule)', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>
      <w:p><w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    // 240tw=16px before, 120tw=8px after, line 360/240 = 1.5× (auto).
    expect(doc.child(0).attrs.spacing).toEqual({
      before: 16,
      after: 8,
      line: 1.5,
      lineRule: 'auto',
    });
    // exact: 480tw → 32px.
    expect(doc.child(1).attrs.spacing).toEqual({ line: 32, lineRule: 'exact' });
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
      await makeDocx(documentXml, undefined, undefined, relsXml, {
        'image1.png': PNG_1x1,
      }),
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
    // w:val=single, no sz/color → 0.5pt clamped to 0.75px, solid, auto→grey.
    const single = { width: 0.75, style: 'solid', color: '#b0b0b0' };
    expect(doc.child(0).attrs.borders).toEqual({ top: single, insideH: false });
    expect(doc.child(1).attrs.borders).toEqual({
      top: single,
      bottom: single,
      left: single,
      right: single,
      insideH: single,
      insideV: single,
    });
    expect(doc.child(2).attrs.borders).toBeNull(); // borderless by default
  });

  it('numbers footnote references and exposes notes as page-bottom bodies', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>First</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
      <w:p><w:r><w:t>Second</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>
    </w:body></w:document>`;
    const footnotesXml = `<?xml version="1.0"?><w:footnotes xmlns:w="${W_NS}">
      <w:footnote w:id="0" w:type="separator"><w:p/></w:footnote>
      <w:footnote w:id="2"><w:p><w:r><w:t>Note alpha</w:t></w:r></w:p></w:footnote>
      <w:footnote w:id="3"><w:p><w:r><w:t>Note beta</w:t></w:r></w:p></w:footnote>
    </w:footnotes>`;
    const { doc, footnotes } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        {
          'word/footnotes.xml': footnotesXml,
        },
      ),
    );
    // references render as superscript 1, 2 in document order, carrying the
    // footnote mark so the layout engine can match them to their bodies.
    const p0 = doc.child(0);
    const ref = p0.child(p0.childCount - 1);
    expect(ref.text).toBe('1');
    expect(markMap(ref.marks).vertAlign.value).toBe('super');
    expect(markMap(ref.marks).footnote.num).toBe(1);
    // bodies live in the footnotes map keyed by display number, NOT appended.
    expect(footnotes[1].textContent).toMatch(/1\. Note alpha/);
    expect(footnotes[2].textContent).toMatch(/2\. Note beta/);
    const tail: string[] = [];
    doc.forEach((n) => tail.push(n.textContent));
    expect(tail.join('\n')).not.toMatch(/Note alpha/);
  });

  it('appends endnotes at the document end (not page-bottom)', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>Body</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p>
    </w:body></w:document>`;
    const endnotesXml = `<?xml version="1.0"?><w:endnotes xmlns:w="${W_NS}">
      <w:endnote w:id="0" w:type="separator"><w:p/></w:endnote>
      <w:endnote w:id="2"><w:p><w:r><w:t>End gamma</w:t></w:r></w:p></w:endnote>
    </w:endnotes>`;
    const { doc, footnotes } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        {
          'word/endnotes.xml': endnotesXml,
        },
      ),
    );
    expect(Object.keys(footnotes)).toHaveLength(0); // endnotes aren't footnotes
    const tail: string[] = [];
    doc.forEach((n) => tail.push(n.textContent));
    expect(tail).toContain('Ghi chú cuối');
    expect(tail.join('\n')).toMatch(/1\. End gamma/);
  });

  it('parses w:cols + section breaks into doc.sections', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:cols w:num="1"/></w:sectPr></w:pPr><w:r><w:t>Intro</w:t></w:r></w:p>
      <w:p><w:r><w:t>Col A</w:t></w:r></w:p>
      <w:p><w:r><w:t>Col B</w:t></w:r></w:p>
      <w:sectPr><w:cols w:num="2" w:space="425"/></w:sectPr>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const sections = doc.attrs['sections'] as
      | {
          blockCount: number;
          columns: { count: number; gap: number };
          newPage: boolean;
        }[]
      | null;
    expect(sections).toHaveLength(2);
    // First section ended by the in-paragraph sectPr: 1 block, continuous, 1 col.
    expect(sections?.[0]).toMatchObject({ blockCount: 1, newPage: false });
    expect(sections?.[0].columns.count).toBe(1);
    // Final body sectPr: the remaining 2 blocks, 2 columns, 425 twips ≈ 28px gap.
    expect(sections?.[1].blockCount).toBe(2);
    expect(sections?.[1].columns).toEqual({
      count: 2,
      gap: Math.round(425 / 15),
    });
  });

  it('leaves doc.sections null for a plain single-column document', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>Plain</w:t></w:r></w:p>
      <w:sectPr><w:cols w:num="1"/></w:sectPr>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    expect(doc.attrs['sections']).toBeNull();
  });

  it('accepts tracked changes: keeps w:ins text, drops w:del', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:t>Keep </w:t></w:r>
        <w:ins><w:r><w:t>inserted </w:t></w:r></w:ins>
        <w:del><w:r><w:delText>removed </w:delText></w:r></w:del>
        <w:r><w:t>tail</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    // inserted text is preserved (was dropped before); deleted text is gone.
    expect(doc.child(0).textContent).toBe('Keep inserted tail');
  });

  it('parses per-cell borders (w:tcBorders) and w:sym symbols', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:sym w:font="Wingdings" w:char="F0B7"/><w:t> item</w:t></w:r></w:p>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc>
          <w:tcPr><w:tcBorders><w:bottom w:val="single"/><w:top w:val="nil"/></w:tcBorders></w:tcPr>
          <w:p><w:r><w:t>x</w:t></w:r></w:p>
        </w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    expect(doc.child(0).textContent).toBe('• item'); // F0B7 → bullet
    expect(doc.child(1).child(0).child(0).attrs.borders).toEqual({
      bottom: { width: 0.75, style: 'solid', color: '#b0b0b0' },
      top: false,
    });
  });

  it('parses table alignment, row height and cell vAlign', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblPr><w:jc w:val="center"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:trPr><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr>
          <w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const table = doc.child(0);
    expect(table.attrs.align).toBe('center');
    expect(table.child(0).attrs.height).toEqual({ value: 40, exact: false }); // 600tw → 40px
    expect(table.child(0).child(0).attrs.vAlign).toBe('center');
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
    expect(doc.child(0).attrs.cellPadding).toEqual({
      left: 20,
      right: 10,
      top: 0,
    });
    expect(doc.child(1).attrs.cellPadding).toBeNull(); // no override → defaults
  });

  it('a table naming no w:tblStyle still inherits the default table style', async () => {
    // Word applies w:default="1" per style TYPE, not just to paragraphs. Stock
    // documents park the 108-twip cell margins and any table borders on
    // "TableNormal" and then never name it, so ignoring the default table
    // style dropped both.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="table" w:default="1" w:styleId="TableNormal">
        <w:tblPr>
          <w:tblCellMar>
            <w:left w:w="300" w:type="dxa"/><w:right w:w="150" w:type="dxa"/>
          </w:tblCellMar>
          <w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/></w:tblBorders>
        </w:tblPr>
      </w:style>
      <w:style w:type="table" w:styleId="Fancy">
        <w:tblPr><w:tblCellMar><w:left w:w="600" w:type="dxa"/></w:tblCellMar></w:tblPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
      <w:tbl>
        <w:tblPr><w:tblStyle w:val="Fancy"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    expect(doc.child(0).attrs.cellPadding).toEqual({ left: 20, right: 10 });
    expect(doc.child(0).attrs.borders).toMatchObject({
      top: { style: 'solid' },
    });
    // A table that DOES name a style resolves through that one, not the
    // default — the default only fills in for tables naming nothing.
    expect(doc.child(1).attrs.cellPadding).toEqual({ left: 40 });
  });

  it('the default style of one type never leaks into another', async () => {
    // Style ids are unique document-wide, so an untyped lookup mostly works;
    // what must not happen is the PARAGRAPH default being picked for a table
    // (or vice versa) just because it is the first w:default="1" in the part.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:spacing w:after="240"/></w:pPr>
      </w:style>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal">
        <w:tblPr><w:tblCellMar><w:left w:w="300" w:type="dxa"/></w:tblCellMar></w:tblPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>body</w:t></w:r></w:p>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    expect(doc.child(0).attrs.spacing).toMatchObject({ after: 16 }); // 240tw
    expect(doc.child(1).attrs.cellPadding).toEqual({ left: 20 });
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

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml),
    );
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
      await makeDocx(documentXml, undefined, undefined, relsXml, {
        'image1.png': PNG_1x1,
      }),
    );
    const img = doc.child(0).child(0);
    expect(img.type.name).toBe('image');
    expect(img.attrs.src).toBe(`data:image/png;base64,${PNG_1x1}`);
    expect(img.attrs.width).toBe(100); // 952500 EMU / 9525
    expect(img.attrs.height).toBe(50); // 476250 / 9525
    expect(img.attrs.alt).toBe('a cat');
  });

  it('finds a w:drawing wrapped in mc:AlternateContent (Choice branch)', async () => {
    // Word wraps some drawings in Choice/Fallback pairs; the image must be
    // rescued from the mc:Choice branch instead of being dropped.
    const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" xmlns:mc="${MC_NS}"><w:body>
      <w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline>
        <wp:extent cx="952500" cy="476250"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r></w:p>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/></Relationships>`;

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, {
        'image1.png': PNG_1x1,
      }),
    );
    const img = doc.child(0).child(0);
    expect(img.type.name).toBe('image');
    expect(img.attrs.src).toBe(`data:image/png;base64,${PNG_1x1}`);
    expect(img.attrs.width).toBe(100);
  });

  it('imports rPr w:spacing as tracking, without disturbing pPr w:spacing', async () => {
    // The two elements share a tag name and nothing else: one is character
    // tracking on a run, the other is before/after/line on a paragraph.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:pPr><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/></w:pPr>
        <w:r><w:rPr><w:spacing w:val="26"/></w:rPr><w:t>tracked</w:t></w:r>
        <w:r><w:rPr><w:spacing w:val="0"/></w:rPr><w:t>reset</w:t></w:r>
        <w:r><w:t>plain</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const para = doc.child(0);
    // Paragraph spacing survived untouched.
    expect(para.attrs.spacing).toMatchObject({ before: 16, after: 8 });
    const track = (i: number) =>
      para.child(i).marks.find((m) => m.type.name === 'letterSpacing')?.attrs[
        'twips'
      ];
    expect(track(0)).toBe(26);
    // 0 is an explicit "back to normal" override of a style, not an absence.
    expect(track(1)).toBe(0);
    expect(track(2)).toBeUndefined();
  });

  it('imports w:position as a baseline shift independent of vertAlign', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:rPr><w:position w:val="-2"/></w:rPr><w:t>lowered</w:t></w:r>
        <w:r><w:rPr><w:position w:val="6"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>both</w:t></w:r>
        <w:r><w:rPr><w:position w:val="0"/></w:rPr><w:t>reset</w:t></w:r>
        <w:r><w:t>plain</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const marksOf = (i: number) =>
      Object.fromEntries(
        doc
          .child(0)
          .child(i)
          .marks.map((m) => [m.type.name, m.attrs]),
      );
    expect(marksOf(0)['position']).toEqual({ halfPoints: -2 });
    // A run can carry both: w:position moves the baseline, w:vertAlign also
    // shrinks the glyphs.
    expect(marksOf(1)['position']).toEqual({ halfPoints: 6 });
    expect(marksOf(1)['vertAlign']).toEqual({ value: 'super' });
    // 0 is an explicit "back to the baseline" override, not an absent value.
    expect(marksOf(2)['position']).toEqual({ halfPoints: 0 });
    expect(marksOf(3)['position']).toBeUndefined();
  });

  it('resolves the full DrawingML colour union and its transforms', async () => {
    const WPS_NS =
      'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
    // accent1 at 60% luminance with +40% offset is how Word writes
    // "Accent 1, Lighter 40%" — reading only @val loses the adjustment and
    // paints the shape in the base accent.
    const themeXml = `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}"><a:themeElements>
      <a:clrScheme name="Office"><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme>
      <a:fmtScheme><a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst></a:fmtScheme>
    </a:themeElements></a:theme>`;
    const shape = (spPr: string, style: string) =>
      `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>
        <a:graphic><a:graphicData><wps:wsp xmlns:wps="${WPS_NS}">
          <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPr}</wps:spPr>
          ${style}
        </wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const documentXml =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}"><w:body><w:p>` +
      shape(
        `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill>`,
        '',
      ) +
      shape(
        `<a:solidFill><a:sysClr val="window" lastClr="FFFFFF"/></a:solidFill>`,
        '',
      ) +
      shape(`<a:solidFill><a:prstClr val="red"/></a:solidFill>`, '') +
      // No direct fill at all: the shape style's fillRef points at the theme
      // format scheme, with its own colour standing in for phClr.
      shape(
        '',
        `<wps:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></wps:style>`,
      ) +
      // An explicit a:noFill must beat the style reference.
      shape(
        '<a:noFill/>',
        `<wps:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></wps:style>`,
      ) +
      `</w:p></w:body></w:document>`;

    const { doc } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        themeXml,
      ),
    );
    const fills = doc
      .child(0)
      .children.map((n) => (n.attrs.shape as { fill?: string }).fill);
    // Word's own swatch for "Accent 1, Lighter 40%" on 4472C4.
    expect(fills[0]).toBe('#8FAADC');
    expect(fills[1]).toBe('#FFFFFF');
    expect(fills[2]).toBe('#FF0000');
    expect(fills[3]).toBe('#4472C4'); // phClr substituted into the scheme
    expect(fills[4]).toBeUndefined();
  });

  it('imports wps shapes (rect/line) as shape-carrying image nodes', async () => {
    const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const WPS_NS =
      'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
    const themeXml = `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}"><a:themeElements><a:clrScheme name="Office">
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme></a:themeElements></a:theme>`;
    // Checkbox-style rect: anchored, outline width via a:ln, color via the
    // style's lnRef (theme accent1) — the shape real quote documents draw.
    const rect = `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300">
      <wp:positionH relativeFrom="column"><wp:posOffset>3041015</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>-67945</wp:posOffset></wp:positionV>
      <wp:extent cx="170815" cy="150495"/><wp:wrapThrough wrapText="bothSides"/><wp:docPr id="5" name="Rectangle 5"/>
      <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp xmlns:wps="${WPS_NS}"><wps:cNvSpPr/>
        <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="170815" cy="150495"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="19050"/></wps:spPr>
        <wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></wps:style>
      <wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>`;
    // Horizontal-rule style straight connector with a direct outline color.
    const line = `<w:drawing><wp:inline><wp:extent cx="952500" cy="0"/><wp:docPr id="6" name="Straight Connector 6"/>
      <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp xmlns:wps="${WPS_NS}">
        <wps:spPr><a:xfrm flipV="1"/><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="9525"><a:solidFill><a:srgbClr val="C45911"/></a:solidFill></a:ln></wps:spPr>
      </wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:mc="${MC_NS}"><w:body>
      <w:p><w:r>${rect}</w:r><w:r>${line}</w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        themeXml,
      ),
    );
    const para = doc.child(0);
    expect(para.childCount).toBe(2);

    const rectNode = para.child(0);
    expect(rectNode.type.name).toBe('image');
    expect(rectNode.attrs.src).toBe('');
    expect(rectNode.attrs.width).toBe(18); // 170815 EMU
    expect(rectNode.attrs.height).toBe(16); // 150495 EMU
    expect(rectNode.attrs.shape).toEqual({
      kind: 'rect',
      strokeWidth: 2,
      stroke: '#4472C4',
    });
    expect(rectNode.attrs.float).toMatchObject({
      wrap: 'square',
      hOffset: 319,
      vRel: 'paragraph',
    });

    const lineNode = para.child(1);
    expect(lineNode.attrs.shape).toEqual({
      kind: 'line',
      strokeWidth: 1,
      stroke: '#C45911',
      flipV: true,
    });
    expect(lineNode.attrs.width).toBe(100);
    expect(lineNode.attrs.height).toBe(0);
    expect(lineNode.attrs.float).toBeNull();
  });

  it('flattens wpg group pictures into per-member floats', async () => {
    const WPG_NS =
      'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
    // Group extent 200×100 px (1905000×952500 EMU), child space 2× that with
    // chOff (100000, -50000) → scale 0.5. Two member pictures.
    const pic = (rid: string, x: number, y: number, cx: number, cy: number) =>
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></pic:spPr></pic:pic>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" xmlns:wpg="${WPG_NS}"><w:body>
      <w:p><w:r><w:drawing><wp:anchor>
        <wp:positionH relativeFrom="column"><wp:posOffset>95250</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV>
        <wp:extent cx="1905000" cy="952500"/><wp:wrapNone/><wp:docPr id="9" name="Group 9"/>
        <a:graphic><a:graphicData uri="${WPG_NS}"><wpg:wgp>
          <wpg:grpSpPr><a:xfrm>
            <a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/>
            <a:chOff x="100000" y="-50000"/><a:chExt cx="3810000" cy="1905000"/>
          </a:xfrm></wpg:grpSpPr>
          ${pic('rId9', 100000, -50000, 1905000, 952500)}
          ${pic('rId9', 2005000, 902500, 1905000, 952500)}
        </wpg:wgp></a:graphicData></a:graphic>
      </wp:anchor></w:drawing></w:r></w:p>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}">
      <Relationship Id="rId9" Type="${R_NS}/image" Target="media/image9.png"/>
    </Relationships>`;

    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, {
        'image9.png': PNG_1x1,
      }),
    );
    const para = doc.child(0);
    expect(para.childCount).toBe(2);
    // Member 1 sits at the group origin: float offset = anchor offset (10, 20).
    const m1 = para.child(0);
    expect(m1.type.name).toBe('image');
    expect(m1.attrs.src).toMatch(/^data:image\/png/);
    expect(m1.attrs.width).toBe(100); // 1905000 EMU × 0.5 scale
    expect(m1.attrs.height).toBe(50);
    expect(m1.attrs.float).toMatchObject({
      wrap: 'none',
      hOffset: 10,
      vOffset: 20,
    });
    // Member 2 offset (1905000, 952500) in child space → +100, +50 px on page.
    const m2 = para.child(1);
    expect(m2.attrs.float).toMatchObject({ hOffset: 110, vOffset: 70 });
  });

  it('imports textbox (wps:txbx) paragraphs onto the shape node', async () => {
    const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const WPS_NS =
      'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
    const box = `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300">
      <wp:positionH relativeFrom="column"><wp:posOffset>62865</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>137453</wp:posOffset></wp:positionV>
      <wp:extent cx="6124575" cy="2607310"/><wp:wrapNone/><wp:docPr id="65" name="Text Box 65"/>
      <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp xmlns:wps="${WPS_NS}"><wps:cNvSpPr txBox="1"/>
        <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6124575" cy="2607310"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>
        <wps:txbx><w:txbxContent>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Phiếu học tập: </w:t></w:r><w:r><w:t>Học sinh trả lời.</w:t></w:r></w:p>
          <w:p><w:r><w:t>Câu 1</w:t></w:r></w:p>
        </w:txbxContent></wps:txbx>
      <wps:bodyPr lIns="182880" tIns="91440"/></wps:wsp></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:mc="${MC_NS}"><w:body>
      <w:p><w:r>${box}</w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    const node = doc.child(0).child(0);
    expect(node.type.name).toBe('image');
    expect(node.attrs.shape).toMatchObject({ kind: 'rect', fill: '#FFFFFF' });

    const tb = node.attrs.textbox as { paragraphs: unknown[]; inset?: unknown };
    expect(tb).toBeTruthy();
    expect(tb.paragraphs).toHaveLength(2);
    // Formatting survives: paragraphs are real PM JSON with marks.
    const p0 = schema.nodeFromJSON(tb.paragraphs[0] as never);
    expect(p0.textContent).toBe('Phiếu học tập: Học sinh trả lời.');
    expect(p0.child(0).marks.map((m) => m.type.name)).toContain('strong');
    // Explicit bodyPr insets in EMU → px (missing sides get Word defaults).
    expect(tb.inset).toEqual({ l: 19, t: 10, r: 10, b: 5 });
  });

  it('imports legacy VML images (w:object + v:imagedata) with pt sizes', async () => {
    const V_NS = 'urn:schemas-microsoft-com:vml';
    const O_NS = 'urn:schemas-microsoft-com:office:office';
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:v="${V_NS}" xmlns:o="${O_NS}"><w:body>
      <w:p><w:r><w:object w:dxaOrig="4240" w:dyaOrig="2290">
        <v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:108.3pt;height:61.35pt">
          <v:imagedata r:id="rId9" o:title="hinh 1"/>
        </v:shape>
      </w:object></w:r></w:p>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId9" Type="${R_NS}/image" Target="media/image2.png"/></Relationships>`;
    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, {
        'image2.png': PNG_1x1,
      }),
    );
    const img = doc.child(0).child(0);
    expect(img.type.name).toBe('image');
    expect(img.attrs.src).toBe(`data:image/png;base64,${PNG_1x1}`);
    expect(img.attrs.width).toBe(144); // 108.3pt × 96/72
    expect(img.attrs.height).toBe(82); // 61.35pt × 96/72
    expect(img.attrs.alt).toBe('hinh 1');
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
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        themeXml,
      ),
    );
    expect(markMap(doc.child(0).child(0).marks).textColor.color).toBe(
      '#4472C4',
    );
    // shade 0x80/255 ≈ 0.502 → 4472C4 darkened ≈ 223962
    expect(markMap(doc.child(1).child(0).marks).textColor.color).toBe(
      '#223962',
    );
  });

  it('runs the field machine inside w:hyperlink (TOC PAGEREF entries)', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:hyperlink w:anchor="_Toc1" w:history="1">
          <w:r><w:t>Chapter one</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>7</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:hyperlink>
      </w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    const p = doc.child(0);
    // Instruction plumbing dropped, entry text + cached page number kept —
    // and the result keeps the hyperlink (same marks ⇒ PM merges the nodes).
    expect(p.textContent).toBe('Chapter one7');
    expect(markMap(p.child(0).marks).link.href).toBe('#_Toc1');
  });

  it('applies the default paragraph style to content that names none', async () => {
    // Word's "Normal" (w:default="1") is where a document keeps its line
    // spacing and space-after. Skipping it collapses every unstyled
    // paragraph — empty ones especially, which is what pushes a heading up
    // the page against a floating logo.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr>
        <w:rPr><w:sz w:val="22"/></w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Tight">
        <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>unstyled</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Tight"/></w:pPr><w:r><w:t>own style</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    // Normal's spacing reaches the unstyled paragraph: 200 twips ≈ 13px
    // after, and w:line 276 is a 1.15 multiple.
    expect(doc.child(0).attrs.spacing).toMatchObject({
      after: 13,
      line: 1.15,
      lineRule: 'auto',
    });
    // A paragraph naming its OWN style does not fall back to Normal — its
    // basedOn chain decides, exactly as Word resolves it.
    expect(doc.child(1).attrs.spacing).toMatchObject({ after: 0, line: 1 });
  });

  it('resolves pagination keeps through the style cascade (toggle-aware)', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:styleId="Glue">
        <w:pPr><w:keepNext/><w:keepLines/></w:pPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Glue"/></w:pPr><w:r><w:t>from style</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Glue"/><w:keepNext w:val="0"/></w:pPr><w:r><w:t>inline off wins</w:t></w:r></w:p>
      <w:p><w:pPr><w:widowControl w:val="0"/></w:pPr><w:r><w:t>widows allowed</w:t></w:r></w:p>
      <w:p><w:r><w:t>defaults</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    expect(doc.child(0).attrs).toMatchObject({
      keepNext: true,
      keepLines: true,
    });
    expect(doc.child(1).attrs.keepNext).toBe(false); // inline w:val="0" wins
    expect(doc.child(1).attrs.keepLines).toBe(true); // untouched by the override
    expect(doc.child(2).attrs.widowControl).toBe(false);
    expect(doc.child(3).attrs).toMatchObject({
      keepNext: false,
      keepLines: false,
      widowControl: true,
    });
  });

  it('imports w:smallCaps and w:dstrike as marks (toggle-aware)', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:rPr><w:smallCaps/></w:rPr><w:t>caps</w:t></w:r>
        <w:r><w:rPr><w:dstrike/></w:rPr><w:t>gone</w:t></w:r>
        <w:r><w:rPr><w:smallCaps w:val="0"/></w:rPr><w:t>off</w:t></w:r>
      </w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    const p = doc.child(0);
    expect(markMap(p.child(0).marks).smallCaps).toBeTruthy();
    expect(markMap(p.child(1).marks).dstrike).toBeTruthy();
    expect(markMap(p.child(2).marks).smallCaps).toBeUndefined();
  });

  it('translates ordinary text set in a symbol font (Wingdings checkbox)', async () => {
    // A ticked checkbox in a real HR form is the letter "x" in Wingdings —
    // not a w:sym. Without translation it renders as a literal "x" wherever
    // the font is missing.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr><w:t>x</w:t></w:r>
        <w:r><w:t xml:space="preserve"> Nữ</w:t></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr><w:t>A</w:t></w:r>
      </w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    const p = doc.child(0);
    // Translated → font-independent, so the font mark goes away and the run
    // merges with the plain text beside it.
    expect(p.child(0).text).toBe('☒ Nữ');
    expect(markMap(p.child(0).marks).fontFamily).toBeUndefined();
    // An untranslatable char keeps the font so its glyph still resolves.
    const last = p.child(p.childCount - 1);
    expect(last.text).toBe('A');
    expect(markMap(last.marks).fontFamily.family).toBe('Wingdings');
  });

  it('maps known w:sym codes to Unicode, tags unknown ones with the symbol font', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r>
        <w:sym w:font="Wingdings" w:char="F0FC"/>
        <w:sym w:font="Wingdings" w:char="F0CF"/>
      </w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    const p = doc.child(0);
    expect(p.child(0).text).toBe('✔'); // known → font-independent Unicode
    const unknown = p.child(1);
    expect(unknown.text).toBe(String.fromCodePoint(0xf0cf)); // PUA kept…
    expect(markMap(unknown.marks).fontFamily.family).toBe('Wingdings'); // …in its font
  });

  it('anchors bookmarks on their paragraph and spans the TOC field across them', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:hyperlink w:anchor="_Toc1"><w:r><w:t>Entry one</w:t></w:r></w:hyperlink>
      </w:p>
      <w:p>
        <w:hyperlink w:anchor="_Toc2"><w:r><w:t>Entry two</w:t></w:r></w:hyperlink>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
      <w:p><w:bookmarkStart w:id="1" w:name="_Toc1"/><w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="1"/>
        <w:r><w:t>Chapter one</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    // Both TOC entry paragraphs — the one that OPENS the field included —
    // carry the same field object; the heading after it does not.
    const f0 = doc.child(0).attrs.field;
    expect(f0).toMatchObject({ kind: 'toc' });
    expect(f0.instr).toContain('TOC');
    expect(doc.child(1).attrs.field).toBe(f0); // identity marks the span
    expect(doc.child(2).attrs.field).toBeNull();
    // The span covers BOTH entry paragraphs from anywhere inside it — the
    // caret sitting in the first must not shrink it to that paragraph.
    const span = { from: 0, to: doc.child(0).nodeSize + doc.child(1).nodeSize };
    expect(fieldAt(doc, 1)).toMatchObject(span);
    expect(fieldAt(doc, span.to - 1)).toMatchObject(span);
    expect(fieldAt(doc, span.to + 2)).toBeNull(); // the heading is outside
    // The heading anchors its bookmark; Word's cursor bookmark is dropped.
    expect(doc.child(2).attrs.bookmarks).toEqual(['_Toc1']);
    expect(findBookmark(doc, '_Toc1')).toBe(
      2 + doc.child(0).nodeSize + doc.child(1).nodeSize - 1,
    );
    expect(bookmarkLabel(doc, '_Toc1')).toBe('Chapter one');
    expect(findBookmark(doc, '_Toc2')).toBeNull(); // no paragraph claims it
  });

  it('keeps cached results of nested and cross-paragraph fields (TOC)', async () => {
    // Word's real TOC shape: the TOC field OPENS in the first entry paragraph
    // and its end lives paragraphs later; each entry nests a PAGEREF field.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:hyperlink w:anchor="_Toc1">
          <w:r><w:t>Entry one</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText> PAGEREF _Toc1 </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>3</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:hyperlink>
      </w:p>
      <w:p>
        <w:hyperlink w:anchor="_Toc2"><w:r><w:t>Entry two</w:t></w:r></w:hyperlink>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    expect(doc.child(0).textContent).toBe('Entry one3');
    expect(doc.child(1).textContent).toBe('Entry two');
    expect(markMap(doc.child(0).child(0).marks).link.href).toBe('#_Toc1');
  });

  it('falls back to first-row dxa w:tcW when w:tblGrid is missing', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tr>
          <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>
          <w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml));
    const row = doc.child(0).child(0);
    expect(row.child(0).attrs.colwidth).toEqual([200]); // 3000 twips
    expect(row.child(1).attrs.colwidth).toEqual([400]); // 6000 twips
  });

  it('resolves w:shd solid patterns and theme fills', async () => {
    const themeXml = `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}"><a:themeElements><a:clrScheme name="Office">
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme></a:themeElements></a:theme>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:rPr><w:shd w:val="solid" w:color="FF0000" w:fill="auto"/></w:rPr><w:t>solid</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:themeFill="accent1"/></w:rPr><w:t>themed</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        themeXml,
      ),
    );
    expect(markMap(doc.child(0).child(0).marks).highlight.color).toBe(
      '#FF0000',
    );
    expect(markMap(doc.child(1).child(0).marks).highlight.color).toBe(
      '#4472C4',
    );
  });

  it('resolves theme fonts (w:asciiTheme) via a:fontScheme', async () => {
    const themeXml = `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}"><a:themeElements>
      <a:fontScheme name="Office">
        <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
        <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
      </a:fontScheme>
    </a:themeElements></a:theme>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/></w:rPr><w:t>heading</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr><w:t>body</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:asciiTheme="minorHAnsi"/></w:rPr><w:t>literal wins</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        themeXml,
      ),
    );
    expect(markMap(doc.child(0).child(0).marks).fontFamily.family).toBe(
      'Calibri Light',
    );
    expect(markMap(doc.child(1).child(0).marks).fontFamily.family).toBe(
      'Calibri',
    );
    expect(markMap(doc.child(2).child(0).marks).fontFamily.family).toBe(
      'Arial',
    );
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
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        relsXml,
        undefined,
        undefined,
        {
          'word/header1.xml': headerXml,
          'word/footer1.xml': footerXml,
        },
      ),
    );

    expect(doc.textContent).toBe('body');
    expect(headers.default.textContent).toBe('Header text');
    expect(
      headers.default
        .child(0)
        .child(0)
        .marks.map((m) => m.type.name),
    ).toContain('strong');
    expect(footers.default.textContent).toBe('Footer text');
  });

  it('parses first/even header variants + titlePg + evenAndOddHeaders', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
      <w:p><w:r><w:t>body</w:t></w:r></w:p>
      <w:sectPr>
        <w:titlePg/>
        <w:headerReference w:type="default" r:id="rIdH"/>
        <w:headerReference w:type="first" r:id="rIdHF"/>
        <w:headerReference w:type="even" r:id="rIdHE"/>
      </w:sectPr>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}">
      <Relationship Id="rIdH" Type="${R_NS}/header" Target="header1.xml"/>
      <Relationship Id="rIdHF" Type="${R_NS}/header" Target="header2.xml"/>
      <Relationship Id="rIdHE" Type="${R_NS}/header" Target="header3.xml"/>
    </Relationships>`;
    const hdr = (t: string) =>
      `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:hdr>`;
    const settingsXml = `<?xml version="1.0"?><w:settings xmlns:w="${W_NS}"><w:evenAndOddHeaders/></w:settings>`;

    const { headers, titlePg, evenAndOdd } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        relsXml,
        undefined,
        undefined,
        {
          'word/header1.xml': hdr('Default hdr'),
          'word/header2.xml': hdr('First hdr'),
          'word/header3.xml': hdr('Even hdr'),
          'word/settings.xml': settingsXml,
        },
      ),
    );
    expect(headers.default.textContent).toBe('Default hdr');
    expect(headers.first.textContent).toBe('First hdr');
    expect(headers.even.textContent).toBe('Even hdr');
    expect(titlePg).toBe(true);
    expect(evenAndOdd).toBe(true);
  });

  it('imports comments + marks the commented range', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:t>Before </w:t></w:r>
        <w:commentRangeStart w:id="0"/>
        <w:r><w:t>commented</w:t></w:r>
        <w:commentRangeEnd w:id="0"/>
        <w:r><w:commentReference w:id="0"/></w:r>
        <w:r><w:t> after</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const commentsXml = `<?xml version="1.0"?><w:comments xmlns:w="${W_NS}">
      <w:comment w:id="0" w:author="Reviewer" w:date="2026-01-02"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:comment>
    </w:comments>`;
    const { doc, comments } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        {
          'word/comments.xml': commentsXml,
        },
      ),
      { schema: withComments },
    );
    expect(comments).toEqual([
      { id: 0, author: 'Reviewer', date: '2026-01-02', text: 'Note body' },
    ]);
    // Only the "commented" run carries the comment mark.
    const p0 = doc.child(0);
    const byText = new Map<string, readonly Mark[]>();
    p0.forEach((n) => {
      if (n.isText) byText.set(n.text ?? '', n.marks);
    });
    expect(markMap(byText.get('commented') ?? []).comment.ids).toEqual([0]);
    expect(
      byText.get('Before ')?.some((m) => m.type.name === 'comment'),
    ).toBeFalsy();
    expect(
      byText.get(' after')?.some((m) => m.type.name === 'comment'),
    ).toBeFalsy();
  });

  it('imports threaded + resolved comments from commentsExtended.xml', async () => {
    // Only the root (id 0) has a range in the body; the reply (id 1) lives only
    // in comments.xml and is linked to id 0 by paraIdParent in commentsExtended.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:commentRangeStart w:id="0"/><w:r><w:t>here</w:t></w:r><w:commentRangeEnd w:id="0"/>
        <w:r><w:commentReference w:id="0"/></w:r>
      </w:p>
    </w:body></w:document>`;
    const commentsXml = `<?xml version="1.0"?><w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
      <w:comment w:id="0" w:author="A" w:date="d0"><w:p w14:paraId="P0"><w:r><w:t>root</w:t></w:r></w:p></w:comment>
      <w:comment w:id="1" w:author="B" w:date="d1"><w:p w14:paraId="P1"><w:r><w:t>reply</w:t></w:r></w:p></w:comment>
    </w:comments>`;
    const extXml = `<?xml version="1.0"?><w15:commentsEx xmlns:w15="${W15_NS}">
      <w15:commentEx w15:paraId="P0" w15:done="1"/>
      <w15:commentEx w15:paraId="P1" w15:paraIdParent="P0" w15:done="0"/>
    </w15:commentsEx>`;
    const { doc } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        {
          'word/comments.xml': commentsXml,
          'word/commentsExtended.xml': extXml,
        },
      ),
      { schema: withComments },
    );
    const nodes = doc.attrs['comments'] as {
      id: number;
      parentId: number | null;
      resolved: boolean;
    }[];
    const byId = new Map(nodes.map((n) => [n.id, n]));
    expect(nodes).toHaveLength(2); // reply included even though it has no body range
    expect(byId.get(0)).toMatchObject({ parentId: null, resolved: true });
    expect(byId.get(1)).toMatchObject({ parentId: 0, resolved: false });
  });

  it('filters comments out when the schema lacks the comment mark (no plugin)', async () => {
    // Same comment-bearing docx, but imported with the DEFAULT (base) schema —
    // i.e. the comment plugin is absent. No comment mark + no threads survive.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:commentRangeStart w:id="0"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="0"/>
        <w:r><w:commentReference w:id="0"/></w:r>
      </w:p>
    </w:body></w:document>`;
    const commentsXml = `<?xml version="1.0"?><w:comments xmlns:w="${W_NS}">
      <w:comment w:id="0" w:author="Reviewer" w:date="2026-01-02"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment>
    </w:comments>`;
    const { doc, comments } = await importDocx(
      await makeDocx(
        documentXml,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        {
          'word/comments.xml': commentsXml,
        },
      ),
    ); // ← no { schema } → base schema has no comment mark
    expect(comments).toEqual([]); // flat list filtered
    expect(doc.attrs['comments']).toBeNull(); // no thread data on the doc
    let anyCommentMark = false;
    doc.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'comment'))
        anyCommentMark = true;
    });
    expect(anyCommentMark).toBe(false); // text carries no comment mark
  });

  it('leaves titlePg/evenAndOdd false when unset', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>body</w:t></w:r></w:p>
      <w:sectPr/>
    </w:body></w:document>`;
    const { titlePg, evenAndOdd } = await importDocx(
      await makeDocx(documentXml),
    );
    expect(titlePg).toBe(false);
    expect(evenAndOdd).toBe(false);
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
    expect(doc.child(0).attrs['indent']).toEqual({
      left: 48,
      right: 24,
      hanging: 16,
    });
    // w:start aliases w:left; 1440/15=96, 240/15=16.
    expect(doc.child(1).attrs['indent']).toEqual({ left: 96, firstLine: 16 });
    expect(doc.child(2).attrs['indent']).toBeNull();
  });

  it('throws when word/document.xml is missing', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'nope');
    await expect(
      importDocx(await zip.generateAsync({ type: 'uint8array' })),
    ).rejects.toThrow(/no document content/);
  });
});

describe('importDocx with a password', () => {
  /** A real .docx, encrypted — the whole path a protected file takes. */
  const protectedDocx = async (password: string) => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>bí mật quý 3</w:t></w:r></w:p>
    </w:body></w:document>`;
    return buildEncryptedDocx(await makeDocx(documentXml), password);
  };

  it('opens a protected document when the password is right', async () => {
    const { doc } = await importDocx(await protectedDocx('mở ra'), {
      password: 'mở ra',
    });
    expect(doc.child(0).textContent).toBe('bí mật quý 3');
  });

  it('without a password it is still the encrypted-kind failure', async () => {
    const err = await importFailure(importDocx(await protectedDocx('x')));
    expect(err.kind).toBe('encrypted');
  });

  it('a wrong password is its own kind, so the prompt can stay open', async () => {
    const err = await importFailure(
      importDocx(await protectedDocx('đúng'), { password: 'sai' }),
    );
    expect(err.kind).toBe('wrong-password');
  });

  it('an unopenable scheme is a dead end, not a retry loop', async () => {
    // A Standard-encryption (v3.2) container: real OLE, wrong scheme.
    const info = new Uint8Array(16);
    new DataView(info.buffer).setUint16(0, 3, true);
    new DataView(info.buffer).setUint16(2, 2, true);
    const ole = new Uint8Array(1024);
    ole.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    const name = 'EncryptionInfo';
    for (let i = 0; i < name.length; i++) ole[512 + i * 2] = name.charCodeAt(i);
    const err = await importFailure(importDocx(ole, { password: 'anything' }));
    expect(err.kind).toBe('unsupported-encryption');
  });
});

describe('sniffDocx + classified import errors', () => {
  const bytes = (s: string) => new TextEncoder().encode(s);
  /** A minimal OLE header followed by a UTF-16LE stream name somewhere after. */
  const ole = (streamName?: string) => {
    const head = new Uint8Array(1024);
    head.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    if (streamName) {
      for (let i = 0; i < streamName.length; i++)
        head[512 + i * 2] = streamName.charCodeAt(i);
    }
    return head;
  };
  const kindOf = async (input: Uint8Array) =>
    importDocx(input).then(
      () => 'OPENED',
      (e) => (e instanceof DocxImportError ? e.kind : `untyped: ${e}`),
    );

  it('classifies the common impostors by magic bytes', () => {
    expect(sniffDocx(bytes('%PDF-1.7 x'))).toBe('pdf');
    expect(sniffDocx(bytes('{\\rtf1 hello'))).toBe('rtf');
    expect(sniffDocx(bytes('  <!DOCTYPE html><html>'))).toBe('html');
    expect(sniffDocx(bytes(''))).toBe('empty');
    expect(sniffDocx(bytes('just some plain text'))).toBe('unknown');
    expect(sniffDocx(bytes('PKrest-of-zip'))).toBe('zip');
    expect(sniffDocx(ole('WordDocument'))).toBe('legacy-doc');
    expect(sniffDocx(ole('EncryptionInfo'))).toBe('encrypted');
    expect(sniffDocx(ole())).toBe('ole');
  });

  it('importDocx throws typed errors for certain non-docx inputs', async () => {
    expect(await kindOf(bytes('%PDF-1.7'))).toBe('pdf');
    expect(await kindOf(ole('EncryptionInfo'))).toBe('encrypted');
    expect(await kindOf(ole('WordDocument'))).toBe('legacy-doc');
    expect(await kindOf(bytes(''))).toBe('empty');
    expect(await kindOf(bytes('plain text impostor'))).toBe('unknown');
  });

  it('a PK header with a broken archive is corrupt-zip (library detail kept)', async () => {
    expect(await kindOf(bytes('PK then garbage, no central directory'))).toBe(
      'corrupt-zip',
    );
    const err = await importFailure(importDocx(bytes('PKgarbage')));
    expect(err.detail).toBeTruthy();
  });

  it('a real zip that is not a docx names the sibling format', async () => {
    const mk = async (path: string) => {
      const zip = new JSZip();
      zip.file(path, '<x/>');
      return zip.generateAsync({ type: 'uint8array' as const });
    };
    expect(await kindOf(await mk('xl/workbook.xml'))).toBe('xlsx');
    expect(await kindOf(await mk('ppt/presentation.xml'))).toBe('pptx');
    expect(await kindOf(await mk('unrelated.txt'))).toBe('no-document');
  });

  it('every kind carries a human-readable message', async () => {
    const err = await importFailure(importDocx(bytes('%PDF-')));
    expect(err.message).toBe(IMPORT_ERROR_MESSAGES.pdf);
    expect(err.message).toContain('PDF');
  });
});
