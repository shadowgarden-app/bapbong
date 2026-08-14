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
import { exportDocx } from './export';
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
const WPS_NS_G =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
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

  it('ranks headings by w:outlineLvl, falling back to the style name', async () => {
    // Word ranks paragraphs by outline level, not by what the style is
    // called. The style name is only a fallback for documents that declare
    // no outline level at all.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading2">
        <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="ChuongMuc">
        <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="TOCHeading">
        <w:basedOn w:val="Heading1"/>
        <w:pPr><w:outlineLvl w:val="9"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading3"/>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="ChuongMuc"/></w:pPr><w:r><w:t>non-English style name</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading2"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>promoted inline</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Table of Contents</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>name only</w:t></w:r></w:p>
    </w:body></w:document>`;

    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    // A style Word never named "HeadingN" is still a heading — level 2.
    expect(doc.child(0).attrs.heading).toBe(2);
    // Direct formatting outranks the style, and outlineLvl outranks the name.
    expect(doc.child(1).attrs.heading).toBe(1);
    // outlineLvl 9 is body text: basedOn Heading1 does NOT make this one.
    expect(doc.child(2).attrs.heading).toBeFalsy();
    // Nothing in the cascade declares a level → the name decides.
    expect(doc.child(3).attrs.heading).toBe(3);
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

  it('reads w:shd at the paragraph layer, through the style cascade', async () => {
    // "The background color behind the paragraph, then the pattern color
    // using the mask supplied by the pattern over that background"
    // (ECMA-376). We paint the background; a shd that carries ONLY a pattern
    // colour has no fill to resolve and stays out of the model.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:styleId="Boxed">
        <w:pPr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:pPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Boxed"/></w:pPr><w:r><w:t>from the style</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Boxed"/><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/></w:pPr><w:r><w:t>inline wins</w:t></w:r></w:p>
      <w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:pPr><w:r><w:t>auto is no fill</w:t></w:r></w:p>
      <w:p><w:pPr><w:shd w:val="pct25" w:color="FF0000"/></w:pPr><w:r><w:t>pattern only</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    expect(doc.child(0).attrs.shading).toBe('#FFFF00');
    // The cascade resolves like every other pPr property: most derived wins.
    expect(doc.child(1).attrs.shading).toBe('#00FF00');
    expect(doc.child(2).attrs.shading).toBeNull();
    expect(doc.child(3).attrs.shading).toBeNull();
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

  it('imports w:pgNumType (page-number restart + format) per section', async () => {
    // Front matter numbered i, ii… then the body restarting at 1 — the shape
    // real reports use (roman TOC pages before a decimal body).
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgNumType w:fmt="lowerRoman" w:start="1"/>
      </w:sectPr></w:pPr><w:r><w:t>front matter</w:t></w:r></w:p>
      <w:p><w:r><w:t>body</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgNumType w:start="1"/>
      </w:sectPr>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const sections = doc.attrs.sections as {
      pageNumbers?: { start?: number; fmt?: string };
    }[];
    expect(sections).toHaveLength(2);
    expect(sections[0].pageNumbers).toEqual({ start: 1, fmt: 'lowerRoman' });
    expect(sections[1].pageNumbers).toEqual({ start: 1 });
  });

  it('rides sections for a single-section doc whose pgNumType restarts', async () => {
    // One section, but numbering starts at 5 — the sections attr must ride
    // (the layout can't know the restart otherwise). An EMPTY pgNumType is a
    // no-op and must not force the attr.
    const numbered = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType w:start="5"/></w:sectPr>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(numbered));
    const sections = doc.attrs.sections as {
      pageNumbers?: { start?: number };
    }[];
    expect(sections).toHaveLength(1);
    expect(sections[0].pageNumbers).toEqual({ start: 5 });

    const plain = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType/></w:sectPr>
    </w:body></w:document>`;
    const back = await importDocx(await makeDocx(plain));
    expect(back.doc.attrs.sections).toBeNull();
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

  it('anchors a body-level w:bookmarkStart to the block that follows', async () => {
    // EG_BlockLevelElts admits bookmarkStart, so this is ordinary markup —
    // Word writes it for TOC anchors ahead of a content control, which is
    // exactly the shape large_sample has. parseParagraph only saw the ones
    // inside a w:p, so a link to these resolved to nothing.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:bookmarkStart w:id="0" w:name="_Toc1" w:displacedByCustomXml="next"/>
      <w:bookmarkStart w:id="1" w:name="_Toc2"/>
      <w:bookmarkStart w:id="2" w:name="_GoBack"/>
      <w:p><w:pPr/><w:bookmarkStart w:id="3" w:name="own"/><w:r><w:t>first</w:t></w:r></w:p>
      <w:bookmarkStart w:id="4" w:name="beforeTable"/>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
      <w:p><w:r><w:t>after</w:t></w:r></w:p>
      <w:bookmarkStart w:id="5" w:name="trailing"/>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    // Both land on the next paragraph, ahead of its own name, in document
    // order. Word's cursor bookmark stays noise.
    expect(doc.child(0).attrs.bookmarks).toEqual(['_Toc1', '_Toc2', 'own']);
    // A table carries no bookmarks attr, so one before it waits for the
    // paragraph after it.
    expect(doc.child(2).attrs.bookmarks).toEqual(['beforeTable', 'trailing']);
  });

  it('accumulates tab stops across the cascade; w:val=clear removes one', async () => {
    // Word's footers are the everyday case: the Footer style sets a centre
    // stop and a right stop, and a paragraph clears one of them.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:styleId="Base"><w:pPr><w:tabs>
        <w:tab w:val="center" w:pos="4320"/>
      </w:tabs></w:pPr></w:style>
      <w:style w:styleId="Footer"><w:basedOn w:val="Base"/><w:pPr><w:tabs>
        <w:tab w:val="right" w:pos="8640" w:leader="underscore"/>
      </w:tabs></w:pPr></w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>inherits both</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Footer"/><w:tabs>
        <w:tab w:val="clear" w:pos="4320"/>
        <w:tab w:val="left" w:pos="1440"/>
      </w:tabs></w:pPr><w:r><w:t>clears the inherited centre</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Footer"/><w:tabs>
        <w:tab w:val="clear" w:pos="4320"/>
        <w:tab w:val="clear" w:pos="8640"/>
      </w:tabs></w:pPr><w:r><w:t>clears everything</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Footer"/><w:tabs>
        <w:tab w:val="decimal" w:pos="8640"/>
      </w:tabs></w:pPr><w:r><w:t>redefines a position</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    // Two layers of the style chain add up — the derived one does not replace.
    expect(doc.child(0).attrs.tabs).toEqual([
      { pos: 288, val: 'center' },
      { pos: 576, val: 'right', leader: 'underscore' },
    ]);
    // A clear takes the inherited stop out; the paragraph's own stop stays.
    expect(doc.child(1).attrs.tabs).toEqual([
      { pos: 96, val: 'left' },
      { pos: 576, val: 'right', leader: 'underscore' },
    ]);
    // Clearing them all falls back to Word's default tab grid (null).
    expect(doc.child(2).attrs.tabs).toBeNull();
    // Same position, different alignment: the derived layer wins that slot.
    expect(doc.child(3).attrs.tabs).toEqual([
      { pos: 288, val: 'center' },
      { pos: 576, val: 'decimal' },
    ]);
  });

  it('reads a:srcRect as a crop, keeping outsets and ignoring the empty one', async () => {
    // ST_Percentage in thousandths of a percent. Positive insets, negative
    // outsets (the source rectangle reaches past the bitmap). The empty
    // <a:srcRect/> Word stamps on nearly every picture must stay a no-op —
    // the audit classifies that shape as inert and it has to remain so.
    const pic = (rid: string, srcRect: string) =>
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${rid}"/>${srcRect}</pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"><w:body>
      ${pic('rId7', '<a:srcRect/>')}
      ${pic('rId7', '<a:srcRect r="18507"/>')}
      ${pic('rId7', '<a:srcRect l="52083" t="36813" r="10126" b="19557"/>')}
      ${pic('rId7', '<a:srcRect l="-25000"/>')}
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/></Relationships>`;
    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, {
        'image1.png': PNG_1x1,
      }),
    );
    const crop = (i: number) => doc.child(i).child(0).attrs.crop;
    expect(crop(0)).toBeNull();
    expect(crop(1)).toEqual({ l: 0, t: 0, r: 0.18507, b: 0 });
    expect(crop(2)).toEqual({
      l: 0.52083,
      t: 0.36813,
      r: 0.10126,
      b: 0.19557,
    });
    // An outset stays negative — clamping it to 0 would silently crop
    // padding away.
    expect(crop(3)).toEqual({ l: -0.25, t: 0, r: 0, b: 0 });
    // The box is untouched: a crop changes what shows, not how big it is.
    expect(doc.child(2).child(0).attrs.width).toBe(100);
  });

  it('reads a picture border (a:ln on pic:spPr)', async () => {
    // Word's "picture border". Shape outlines have always been read; a
    // framed photo lost its frame because nothing looked at the picture's
    // own spPr.
    const pic = (rid: string, spPr: string) =>
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/>
        <a:graphic><a:graphicData><pic:pic>
          <pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill>
          <pic:spPr>${spPr}</pic:spPr>
        </pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"><w:body>
      ${pic('rId7', '<a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>')}
      ${pic('rId7', '<a:ln><a:noFill/></a:ln>')}
      ${pic('rId7', '')}
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/></Relationships>`;
    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, {
        'image1.png': PNG_1x1,
      }),
    );
    // 9525 EMU = 1px.
    expect(doc.child(0).child(0).attrs.outline).toEqual({
      width: 1,
      style: 'solid',
      color: '#FF0000',
    });
    // The spec's explicit "no outline" stays no outline, and so does silence.
    expect(doc.child(1).child(0).attrs.outline).toBeNull();
    expect(doc.child(2).child(0).attrs.outline).toBeNull();
  });

  it('takes a shape outline width from the theme line style its lnRef names', async () => {
    // a:lnRef @idx is one-based into the theme's a:lnStyleLst. A gallery
    // shape carries only the ref, so without resolving it every such shape
    // drew at the 1px fallback whatever the theme said.
    const themeXml = `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}"><a:themeElements>
      <a:clrScheme name="Office"><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme>
      <a:fmtScheme><a:lnStyleLst>
        <a:ln w="9525"><a:prstDash val="solid"/></a:ln>
        <a:ln w="28575"><a:prstDash val="dash"/></a:ln>
      </a:lnStyleLst></a:fmtScheme>
    </a:themeElements></a:theme>`;
    const shape = (idx: string, ln: string) =>
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/>
        <a:graphic><a:graphicData><wps:wsp>
          <wps:spPr><a:prstGeom prst="rect"/>${ln}</wps:spPr>
          <wps:style><a:lnRef idx="${idx}"><a:schemeClr val="accent1"/></a:lnRef></wps:style>
        </wps:wsp></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:wps="${WPS_NS_G}"><w:body>
      ${shape('2', '')}
      ${shape('2', '<a:ln w="9525"/>')}
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
    const shapeOf = (i: number) =>
      doc.child(i).child(0).attrs.shape as Record<string, unknown>;
    // Entry 2 of the list: 28575 EMU = 3px, dashed.
    expect(shapeOf(0).strokeWidth).toBe(3);
    expect(shapeOf(0).dash).toBeDefined();
    // The shape's own a:ln overrides the width it states; the theme's dash,
    // which the shape says nothing about, survives.
    expect(shapeOf(1).strokeWidth).toBe(1);
    expect(shapeOf(1).dash).toBeDefined();
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

  it('takes the mc:Choice it understands, else the Fallback', async () => {
    // ISO/IEC 29500-3: walk the Choice branches in order and take the first
    // whose @Requires namespaces the consumer understands; otherwise the
    // Fallback. khbd's inked pages are the case that made this matter — the
    // "aink" Choice hands over a w:drawing whose graphicData holds a pen
    // stroke we cannot draw, while the Fallback holds it as a picture.
    const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" xmlns:mc="${MC_NS}" xmlns:w14="${W14_NS}"><w:body>
      <w:p><w:r><mc:AlternateContent>
        <mc:Choice Requires="aink"><w:drawing><wp:inline>
          <wp:extent cx="360" cy="360"/><wp:docPr descr="the stroke itself"/>
          <a:graphic><a:graphicData><w14:contentPart r:id="rId8"/></a:graphicData></a:graphic>
        </wp:inline></w:drawing></mc:Choice>
        <mc:Fallback><w:drawing><wp:inline>
          <wp:extent cx="952500" cy="476250"/><wp:docPr descr="rasterised ink"/>
          <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
        </wp:inline></w:drawing></mc:Fallback>
      </mc:AlternateContent></w:r></w:p>
    </w:body></w:document>`;
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/></Relationships>`;
    const { doc } = await importDocx(
      await makeDocx(documentXml, undefined, undefined, relsXml, {
        'image1.png': PNG_1x1,
      }),
    );
    const img = doc.child(0).child(0);
    // Before the MCE rule this paragraph rendered nothing at all: the Choice
    // won on "has a w:drawing", then the ink was dropped.
    expect(img.type.name).toBe('image');
    expect(img.attrs.alt).toBe('rasterised ink');
    expect(img.attrs.width).toBe(100); // the Fallback's extent, not 360 EMU
  });

  // ── conditional table formatting (w:tblStylePr) ───────────────────
  // One style carrying every branch, each with a DIFFERENT left indent, so a
  // cell's indent names the branch that won. twipsToPx divides by 15.
  const CONDITIONAL_STYLES = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
    <w:style w:type="table" w:default="1" w:styleId="TableNormal"/>
    <w:style w:type="table" w:styleId="Grid">
      <w:pPr><w:ind w:left="150"/></w:pPr>
      <w:tblPr>
        <w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>
      </w:tblPr>
      <w:tblStylePr w:type="wholeTable"><w:pPr><w:ind w:left="3000"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="band1Horz"><w:pPr><w:ind w:left="300"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="band2Horz"><w:pPr><w:ind w:left="450"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="band1Vert"><w:pPr><w:ind w:left="600"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="band2Vert"><w:pPr><w:ind w:left="750"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="firstCol"><w:pPr><w:ind w:left="900"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="lastCol"><w:pPr><w:ind w:left="1050"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="firstRow"><w:pPr><w:ind w:left="1200"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="lastRow"><w:pPr><w:ind w:left="1350"/></w:pPr></w:tblStylePr>
      <w:tblStylePr w:type="nwCell"><w:pPr><w:ind w:left="1500"/></w:pPr></w:tblStylePr>
    </w:style>
    <w:style w:type="table" w:styleId="NoBands">
      <w:pPr><w:ind w:left="150"/></w:pPr>
      <w:tblStylePr w:type="band1Horz"><w:pPr><w:ind w:left="300"/></w:pPr></w:tblStylePr>
    </w:style>
  </w:styles>`;

  /** A table of `rows` × `cols` cells, each holding one empty paragraph. */
  const condTable = (
    styleId: string,
    look: string,
    rows: number,
    cols: number,
    trPrOf: (r: number) => string = () => '',
  ) =>
    `<w:tbl><w:tblPr><w:tblStyle w:val="${styleId}"/>${look}</w:tblPr>` +
    `<w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(cols)}</w:tblGrid>` +
    Array.from({ length: rows }, (_, r) => {
      const cells = Array.from(
        { length: cols },
        (_, c) => `<w:tc><w:p><w:r><w:t>r${r}c${c}</w:t></w:r></w:p></w:tc>`,
      ).join('');
      return `<w:tr>${trPrOf(r)}${cells}</w:tr>`;
    }).join('') +
    `</w:tbl>`;

  it("applies w:tblStylePr branches in Word's order, gated by w:tblLook", async () => {
    // 0x01E0 = firstRow|lastRow|firstColumn|lastColumn, neither noHBand nor
    // noVBand — table 0 adds noVBand (0x0400 → 0x05E0) to isolate ROW banding,
    // table 1 adds noHBand (0x0200 → 0x03E0) to isolate COLUMN banding.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${condTable('Grid', '<w:tblLook w:val="05E0"/>', 4, 3)}
      ${condTable('Grid', '<w:tblLook w:val="03E0"/>', 3, 4)}
    </w:body></w:document>`;
    const { doc } = await importDocx(
      await makeDocx(documentXml, CONDITIONAL_STYLES),
    );
    const at = (t: number, r: number, c: number) =>
      (
        doc.child(t).child(r).child(c).child(0).attrs['indent'] as {
          left?: number;
        } | null
      )?.left ?? 0;

    // Row banding, and the order between branches.
    expect(at(0, 0, 1)).toBe(80); // firstRow, and NOT banded
    expect(at(0, 0, 0)).toBe(100); // firstRow+firstCol → nwCell last
    expect(at(0, 0, 2)).toBe(80); // lastCol then firstRow → firstRow wins
    expect(at(0, 1, 1)).toBe(20); // first BODY row → band1Horz
    expect(at(0, 2, 1)).toBe(30); // second body row → band2Horz
    expect(at(0, 1, 0)).toBe(60); // band1Horz then firstCol → firstCol wins
    expect(at(0, 3, 1)).toBe(90); // lastRow, and NOT banded
    expect(at(0, 3, 0)).toBe(90); // firstCol then lastRow → lastRow wins

    // Column banding: the first column drops out of the count, so the column
    // beside it is band1, the next band2.
    expect(at(1, 1, 1)).toBe(40);
    expect(at(1, 1, 2)).toBe(50);
    expect(at(1, 1, 0)).toBe(60); // firstCol
    expect(at(1, 1, 3)).toBe(70); // lastCol

    // "Word does not apply and discards on save any properties within the
    // tblStylePr element when the type attribute has a value of wholeTable"
    // (MS-OI29500 §17.18.89(a)) — 200px must appear nowhere.
    for (let t = 0; t < 2; t++)
      for (let r = 0; r < 3; r++)
        for (let c = 0; c < 3; c++) expect(at(t, r, c)).not.toBe(200);
  });

  it('reads a missing w:tblLook as 0x04A0, and lets attributes beat w:val', async () => {
    // "In Word, when the tblLook element is omitted, the bitmask … is assumed
    // to be 0x04A0" — firstRow + firstColumn + no vertical banding, with
    // horizontal banding ALLOWED. Table 1 says 0x0000 in w:val but sets the
    // attributes, which Word reads instead ("Word reads the val attribute … if,
    // and only if, none of the attributes … are present").
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${condTable('Grid', '', 3, 2)}
      ${condTable('Grid', '<w:tblLook w:val="0000" w:firstRow="1" w:noHBand="1" w:noVBand="1"/>', 3, 2)}
      ${condTable('NoBands', '<w:tblLook w:val="01E0"/>', 3, 2)}
    </w:body></w:document>`;
    const { doc } = await importDocx(
      await makeDocx(documentXml, CONDITIONAL_STYLES),
    );
    const at = (t: number, r: number, c: number) =>
      (
        doc.child(t).child(r).child(c).child(0).attrs['indent'] as {
          left?: number;
        } | null
      )?.left ?? 0;

    expect(at(0, 0, 0)).toBe(100); // firstRow+firstColumn → nwCell
    expect(at(0, 1, 1)).toBe(20); // banding is on: first body row
    // lastRow is OFF in 0x04A0, so the last row is still a banded body row —
    // the rule Word's own w:cnfStyle confirmed on a 39-row Light Grid table.
    expect(at(0, 2, 1)).toBe(30);

    // Attributes present ⇒ w:val ignored: firstRow applies, banding does not.
    // Both no*Band flags have to be spelled out — an ABSENT noVBand means
    // vertical banding is allowed, which is the schema default, not "off".
    expect(at(1, 0, 1)).toBe(80);
    expect(at(1, 1, 1)).toBe(10); // no branch → the table style's own pPr
    expect(at(1, 1, 0)).toBe(10); // firstColumn absent from the attributes

    // "Word uses 0 as the default for tblStyleRowBandSize [and] does not apply
    // any banded row conditional formatting" (MS-OI29500 §2.1.251) — NoBands
    // declares band1Horz but no band size, so it never fires.
    expect(at(2, 1, 1)).toBe(10);
  });

  it("takes a table's alignment from its style when the table declares none", async () => {
    // Same fallback shape as w:tblBorders and w:tblCellMar: the table's own
    // w:jc wins, the style's applies otherwise, and ST_JcTable's "end" maps to
    // right. Row-level w:jc is NOT read — see the note at the call site.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal"/>
      <w:style w:type="table" w:styleId="Base"><w:tblPr><w:jc w:val="center"/></w:tblPr></w:style>
      <w:style w:type="table" w:styleId="Derived"><w:basedOn w:val="Base"/></w:style>
      <w:style w:type="table" w:styleId="Ends"><w:tblPr><w:jc w:val="end"/></w:tblPr></w:style>
    </w:styles>`;
    const tbl = (pr: string) =>
      `<w:tbl><w:tblPr>${pr}</w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${tbl('<w:tblStyle w:val="Base"/>')}
      ${tbl('<w:tblStyle w:val="Derived"/>')}
      ${tbl('<w:tblStyle w:val="Ends"/>')}
      ${tbl('<w:tblStyle w:val="Base"/><w:jc w:val="left"/>')}
      ${tbl('')}
      <w:p><w:r><w:t>tail</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const align = (i: number) => doc.child(i).attrs['align'];
    expect(align(0)).toBe('center'); // straight from the style
    expect(align(1)).toBe('center'); // …and through basedOn
    expect(align(2)).toBe('right'); // ST_JcTable "end"
    expect(align(3)).toBe(null); // the table's own w:jc wins
    expect(align(4)).toBe(null); // no style alignment, no table alignment
  });

  it("layers a table style's w:tcPr under the conditional branches", async () => {
    // Cell properties stack the same way the paragraph ones do: the style's
    // own w:tcPr → the branches that reach the cell → the cell's own w:tcPr.
    // A later layer's w:shd with fill="auto" means NO shading and clears an
    // inherited fill; w:tcMar merges per side.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal"/>
      <w:style w:type="table" w:styleId="Banded">
        <w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>
        <w:tcPr>
          <w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>
          <w:vAlign w:val="center"/>
          <w:tcMar><w:left w:w="150" w:type="dxa"/><w:top w:w="300" w:type="dxa"/></w:tcMar>
        </w:tcPr>
        <w:tblStylePr w:type="firstRow">
          <w:tcPr>
            <w:shd w:val="clear" w:color="auto" w:fill="336699"/>
            <w:vAlign w:val="bottom"/>
            <w:tcMar><w:left w:w="450" w:type="dxa"/></w:tcMar>
          </w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="band1Horz">
          <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr>
        </w:tblStylePr>
      </w:style>
    </w:styles>`;
    const cell = (inner = '') =>
      `<w:tc><w:tcPr>${inner}</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblPr><w:tblStyle w:val="Banded"/><w:tblLook w:val="0420"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
        <w:tr>${cell()}${cell('<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>')}</w:tr>
        <w:tr>${cell()}${cell()}</w:tr>
        <w:tr>${cell()}${cell()}</w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const cellAt = (r: number, c: number) => doc.child(0).child(r).child(c);

    // Row 0 is firstRow: its branch beats the style's own w:tcPr…
    expect(cellAt(0, 0).attrs['background']).toBe('#336699');
    expect(cellAt(0, 0).attrs['vAlign']).toBe('bottom');
    // …and w:tcMar merges: left comes from the branch, top from the style.
    expect(cellAt(0, 0).attrs['padding']).toMatchObject({ left: 30, top: 20 });
    // …while the cell's own w:shd still beats the branch.
    expect(cellAt(0, 1).attrs['background']).toBe('#FF0000');

    // Row 1 is the first body row → band1Horz, whose fill="auto" CLEARS the
    // style's EEEEEE rather than leaving it in place.
    expect(cellAt(1, 0).attrs['background']).toBe(null);
    expect(cellAt(1, 0).attrs['vAlign']).toBe('center'); // still the style's
    // Row 2 is band2Horz, which this style does not declare → style's own fill.
    expect(cellAt(2, 0).attrs['background']).toBe('#EEEEEE');
  });

  it("maps a branch's insideH/insideV to its own region, not the table", async () => {
    // Inside a w:tblStylePr branch, insideH/insideV mean the edges BETWEEN
    // cells OF THAT REGION. For a firstRow branch the region is one row: every
    // cell is on its top and bottom edge, so top/bottom come from the branch's
    // own w:top/w:bottom while the vertical edges between the row's cells come
    // from insideV — except the row's outer left/right, which take w:left and
    // w:right. A firstCol branch is the same rectangle turned 90°.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal"/>
      <w:style w:type="table" w:styleId="Edges">
        <w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders>
          <w:top w:val="single" w:sz="24" w:space="0" w:color="FF0000"/>
          <w:bottom w:val="single" w:sz="16" w:space="0" w:color="00FF00"/>
          <w:left w:val="single" w:sz="8" w:space="0" w:color="0000FF"/>
          <w:right w:val="single" w:sz="8" w:space="0" w:color="0000FF"/>
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="FFFF00"/>
          <w:insideH w:val="none"/>
        </w:tcBorders></w:tcPr></w:tblStylePr>
        <w:tblStylePr w:type="firstCol"><w:tcPr><w:tcBorders>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="00FFFF"/>
        </w:tcBorders></w:tcPr></w:tblStylePr>
      </w:style>
    </w:styles>`;
    const cell = '<w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>';
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblPr><w:tblStyle w:val="Edges"/><w:tblLook w:val="06A0"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="900"/><w:gridCol w:w="900"/><w:gridCol w:w="900"/></w:tblGrid>
        <w:tr>${cell.repeat(3)}</w:tr>
        <w:tr>${cell.repeat(3)}</w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const borders = (r: number, c: number) =>
      doc.child(0).child(r).child(c).attrs['borders'] as Record<
        string,
        { color?: string } | false
      > | null;

    // First row, middle cell: the region's top and bottom edges are this
    // cell's own, and both vertical edges are interior → insideV.
    expect(borders(0, 1)?.['top']).toMatchObject({ color: '#FF0000' });
    expect(borders(0, 1)?.['bottom']).toMatchObject({ color: '#00FF00' });
    expect(borders(0, 1)?.['left']).toMatchObject({ color: '#FFFF00' });
    expect(borders(0, 1)?.['right']).toMatchObject({ color: '#FFFF00' });
    // First row, first cell: its left edge is the region's left edge.
    expect(borders(0, 0)?.['left']).toMatchObject({ color: '#0000FF' });
    expect(borders(0, 2)?.['right']).toMatchObject({ color: '#0000FF' });

    // The firstCol branch's region is a column, so ITS insideH lands on the
    // horizontal edge between the two rows of column 0 — the bottom of (0,0)
    // is the firstRow branch's (applied later, hence winning) and the top of
    // (1,0) is the column branch's.
    expect(borders(1, 0)?.['top']).toMatchObject({ color: '#00FFFF' });
    // Nothing reaches row 1's other cells.
    expect(borders(1, 1)).toBe(null);
  });

  it('gives a w:tblHeader row the firstRow branch', async () => {
    // "In addition, if the cell is in a row with the w:tblHeader element, then
    // add the run style property from w:tblStylePr[@w:type = 'firstRow']"
    // — Eric White. Such a row is also out of the banding count.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${condTable('Grid', '<w:tblLook w:val="05E0"/>', 4, 3, (r) =>
        r === 1 ? '<w:trPr><w:tblHeader/></w:trPr>' : '',
      )}
    </w:body></w:document>`;
    const { doc } = await importDocx(
      await makeDocx(documentXml, CONDITIONAL_STYLES),
    );
    const at = (r: number, c: number) =>
      (
        doc.child(0).child(r).child(c).child(0).attrs['indent'] as {
          left?: number;
        } | null
      )?.left ?? 0;
    expect(at(0, 1)).toBe(80); // the real first row
    expect(at(1, 1)).toBe(80); // repeated heading row, not band1Horz
    expect(at(2, 1)).toBe(30); // body index 1 → band2Horz, unshifted by row 1
  });

  it("applies a table style's w:rPr to the runs inside, toggles included", async () => {
    // Same slot as the paragraph layer: docDefaults → table style → paragraph
    // style → character style → direct. The toggle properties come through
    // as plain values, because Word "resets the value of the toggle property
    // to the value specified by the paragraph style if a value is present"
    // (MS-OI29500 §2.1.258) instead of flipping it.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:docDefaults><w:rPrDefault><w:rPr>
        <w:sz w:val="20"/><w:color w:val="111111"/>
      </w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal"/>
      <w:style w:type="table" w:styleId="Grid">
        <w:rPr>
          <w:sz w:val="28"/><w:color w:val="FF0000"/>
          <w:b/><w:i/><w:smallCaps/><w:strike/>
          <w:u w:val="single"/>
        </w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Big"><w:rPr><w:sz w:val="40"/></w:rPr></w:style>
    </w:styles>`;
    const row = (inner: string) => `<w:tr><w:tc>${inner}</w:tc></w:tr>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>outside</w:t></w:r></w:p>
      <w:tbl>
        <w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        ${row('<w:p><w:r><w:t>plain</w:t></w:r></w:p>')}
        ${row('<w:p><w:pPr><w:pStyle w:val="Big"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>')}
        ${row('<w:p><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>direct</w:t></w:r></w:p>')}
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    // row → cell → paragraph → run, inside the table at block 1.
    const cellRun = (r: number) =>
      doc.child(1).child(r).child(0).child(0).child(0);
    const markOf = (node: { marks: readonly Mark[] }, name: string) =>
      node.marks.find((m) => m.type.name === name);

    // Outside: docDefaults only.
    expect(markOf(doc.child(0).child(0), 'fontSize')?.attrs['size']).toBe(10);

    // Inside: the table style beats docDefaults, for size and colour.
    expect(markOf(cellRun(0), 'fontSize')?.attrs['size']).toBe(14);
    expect(markOf(cellRun(0), 'textColor')?.attrs['color']).toBe('#FF0000');
    // Non-toggle character formatting from a table style does apply.
    expect(markOf(cellRun(0), 'underline')).toBeDefined();
    // …and so do the toggles.
    expect(markOf(cellRun(0), 'strong')).toBeDefined();
    expect(markOf(cellRun(0), 'em')).toBeDefined();
    expect(markOf(cellRun(0), 'smallCaps')).toBeDefined();
    expect(markOf(cellRun(0), 'strike')).toBeDefined();
    // A paragraph style beats the table style; direct beats everything.
    expect(markOf(cellRun(1), 'fontSize')?.attrs['size']).toBe(20);
    expect(markOf(cellRun(2), 'fontSize')?.attrs['size']).toBe(9);
  });

  it('does not cancel bold when two style layers both set it', async () => {
    // ISO 29500 makes w:b a toggle property: a style setting it flips whatever
    // the hierarchy had. Word does not — §2.1.258 says it RESETS to the value
    // given. So a paragraph style and a character style that both say w:b
    // leave the run bold, and w:val="0" turns bold off rather than being the
    // spec's no-op ("setting it to false … shall result in the current setting
    // remaining unchanged"). 259 runs in one 249-page lesson plan sit on this
    // exact shape, and Word shows them bold.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal"/>
      <w:style w:type="paragraph" w:styleId="BoldPara"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="BoldChar"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="UnboldChar"><w:rPr><w:b w:val="0"/></w:rPr></w:style>
    </w:styles>`;
    const p = (rStyle: string) =>
      `<w:p><w:pPr><w:pStyle w:val="BoldPara"/></w:pPr><w:r>` +
      `<w:rPr>${rStyle}</w:rPr><w:t>x</w:t></w:r></w:p>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${p('')}
      ${p('<w:rStyle w:val="BoldChar"/>')}
      ${p('<w:rStyle w:val="UnboldChar"/>')}
      ${p('<w:b w:val="0"/>')}
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const bold = (i: number) =>
      doc
        .child(i)
        .child(0)
        .marks.some((m) => m.type.name === 'strong');
    expect(bold(0)).toBe(true); // paragraph style alone
    expect(bold(1)).toBe(true); // + character style: still bold, NOT flipped
    expect(bold(2)).toBe(false); // a style's w:val="0" turns it off
    expect(bold(3)).toBe(false); // direct formatting is absolute
  });

  it('computes auto paragraph spacing the way Word does', async () => {
    // "If this attribute is specified, then any value in the before or
    // beforeLines attributes is ignored" — the 100 twips Word writes beside the
    // flag is its own cached guess, not an instruction. What replaces it is
    // 14pt, and only at a boundary between two paragraphs: nothing before the
    // first paragraph of a container, nothing after the last, nothing between
    // items of the same list.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal"/>
      <w:style w:type="paragraph" w:styleId="NormalWeb">
        <w:pPr><w:spacing w:before="100" w:beforeAutospacing="1"
                          w:after="100" w:afterAutospacing="1"/></w:pPr>
      </w:style>
    </w:styles>`;
    const web = (inner = '', extra = '') =>
      `<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>${extra}</w:pPr>` +
      `<w:r><w:t>${inner || 'x'}</w:t></w:r></w:p>`;
    const listed = (numId: string) =>
      web(
        '',
        `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`,
      );
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${web('first')}
      ${web('middle')}
      ${listed('1')}
      ${listed('1')}
      ${web('after the list')}
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>${web('alone in its cell')}</w:tc>
          <w:tc>${web('cell top')}${web('cell bottom')}</w:tc>
        </w:tr>
      </w:tbl>
      ${web('last')}
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const sp = (n: { attrs: Record<string, unknown> }) =>
      n.attrs['spacing'] as {
        before?: number;
        after?: number;
        beforeAuto?: boolean;
        afterAuto?: boolean;
      } | null;
    const AUTO = 19; // 14pt = 280 twips

    // First body paragraph: no predecessor, so no space before; the next
    // paragraph earns the space after. The literal 100tw (≈7px) is gone.
    expect(sp(doc.child(0))).toMatchObject({ before: 0, after: AUTO });
    // …and the flags survive for the exporter.
    expect(sp(doc.child(0))).toMatchObject({
      beforeAuto: true,
      afterAuto: true,
    });
    expect(sp(doc.child(1))).toMatchObject({ before: AUTO, after: AUTO });
    // Two items of one list: nothing BETWEEN them, but the boundaries with the
    // paragraphs around the list are ordinary ones — "spacing is added only
    // after the last item in the list".
    expect(sp(doc.child(2))).toMatchObject({ before: AUTO, after: 0 });
    expect(sp(doc.child(3))).toMatchObject({ before: 0, after: AUTO });
    // A table is not a paragraph, so the paragraph before it gets no after.
    expect(sp(doc.child(4))).toMatchObject({ before: AUTO, after: 0 });

    // In a cell the container is the cell: alone means no neighbour on either
    // side, and a pair only meets in the middle.
    const cell = (c: number) => doc.child(5).child(0).child(c);
    expect(sp(cell(0).child(0))).toMatchObject({ before: 0, after: 0 });
    expect(sp(cell(1).child(0))).toMatchObject({ before: 0, after: AUTO });
    expect(sp(cell(1).child(1))).toMatchObject({ before: AUTO, after: 0 });
  });

  it('honours w:doNotUseHTMLParagraphAutoSpacing and an inline "off"', async () => {
    // The compat setting swaps the HTML emulation for a fixed pair: "5 points
    // of spacing before and 10 points of spacing after". No file in the repo
    // sets it, so this test is the only thing holding the branch up.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="paragraph" w:styleId="NormalWeb">
        <w:pPr><w:spacing w:before="100" w:beforeAutospacing="1"
                          w:after="100" w:afterAutospacing="1"/></w:pPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="NormalWeb"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="NormalWeb"/>
        <w:spacing w:before="240" w:beforeAutospacing="0" w:after="240" w:afterAutospacing="0"/>
      </w:pPr><w:r><w:t>b</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(
      await makeDocx(
        documentXml,
        stylesXml,
        undefined,
        undefined,
        undefined,
        undefined,
        {
          'word/settings.xml':
            `<?xml version="1.0"?><w:settings xmlns:w="${W_NS}">` +
            `<w:doNotUseHTMLParagraphAutoSpacing/></w:settings>`,
        },
      ),
    );
    const sp = (i: number) =>
      doc.child(i).attrs['spacing'] as {
        before?: number;
        after?: number;
      } | null;
    // 5pt → 7px, 10pt → 13px, and the neighbour rules do not apply: these are
    // ordinary fixed values once the compat flag is on.
    expect(sp(0)).toMatchObject({ before: 7, after: 13 });
    // An inline "0" turns the flag off, so this paragraph's own 240tw applies.
    expect(sp(1)).toMatchObject({ before: 16, after: 16 });
  });

  it("applies a table style's w:pPr to the paragraphs inside the table", async () => {
    // "The global default paragraph properties · The table style paragraph
    // properties · The paragraph properties applied directly to a paragraph"
    // — so a table style beats docDefaults and loses to a paragraph style,
    // and w:spacing merges per ATTRIBUTE across those layers.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:docDefaults><w:pPrDefault><w:pPr>
        <w:spacing w:before="240" w:after="240" w:line="360" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr></w:pPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal">
        <w:pPr><w:jc w:val="right"/></w:pPr>
      </w:style>
      <w:style w:type="table" w:styleId="Grid">
        <w:basedOn w:val="TableNormal"/>
        <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Loud">
        <w:pPr><w:jc w:val="center"/></w:pPr>
      </w:style>
    </w:styles>`;
    const cell = (inner: string) => `<w:tr><w:tc>${inner}</w:tc></w:tr>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>outside</w:t></w:r></w:p>
      <w:tbl>
        <w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        ${cell('<w:p><w:r><w:t>plain</w:t></w:r></w:p>')}
        ${cell('<w:p><w:pPr><w:pStyle w:val="Loud"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>')}
        ${cell('<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>direct</w:t></w:r></w:p>')}
      </w:tbl>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
        ${cell('<w:p><w:r><w:t>no tblStyle</w:t></w:r></w:p>')}
      </w:tbl>
      <w:p><w:r><w:t>after</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    // block → row → cell → paragraph
    const para = (b: number, r: number) =>
      doc.child(b).child(r).child(0).child(0);

    // Outside any table: docDefaults only.
    expect(doc.child(0).attrs.align).toBe('justify');
    expect(doc.child(0).attrs.spacing).toMatchObject({ after: 16, line: 1.5 });

    // Inside: the table style wins over docDefaults. `before` survives from
    // docDefaults because the table style's w:spacing names only after/line —
    // per-attribute merging, not wholesale replacement.
    expect(para(1, 0).attrs.align).toBe('left');
    expect(para(1, 0).attrs.spacing).toMatchObject({
      before: 16,
      after: 0,
      line: 1,
    });
    // A paragraph style beats the table style.
    expect(para(1, 1).attrs.align).toBe('center');
    // Direct formatting beats everything.
    expect(para(1, 2).attrs.align).toBe('right');
    // No w:tblStyle → the DEFAULT table style still applies (jc=right).
    expect(para(2, 0).attrs.align).toBe('right');
    // The stack is popped: the paragraph after the table is unaffected.
    expect(doc.child(3).attrs.align).toBe('justify');
  });

  it('takes the innermost table style, and never leaks into a textbox', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal"/>
      <w:style w:type="table" w:styleId="Outer"><w:pPr><w:jc w:val="right"/></w:pPr></w:style>
      <w:style w:type="table" w:styleId="Inner"><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
    </w:styles>`;
    const textbox =
      `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>` +
      `<a:graphic><a:graphicData><wps:wsp>` +
      `<wps:spPr><a:prstGeom prst="rect"/></wps:spPr>` +
      `<wps:txbx><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
      `</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:wps="${WPS_NS_G}"><w:body>
      <w:tbl>
        <w:tblPr><w:tblStyle w:val="Outer"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
        <w:tr><w:tc>
          <w:p><w:r><w:t>outer cell</w:t></w:r></w:p>
          <w:tbl>
            <w:tblPr><w:tblStyle w:val="Inner"/></w:tblPr>
            <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>inner cell</w:t></w:r></w:p></w:tc></w:tr>
          </w:tbl>
          <w:p><w:r><w:t>after inner</w:t></w:r></w:p>
          <w:p>${textbox}</w:p>
        </w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const outerCell = doc.child(0).child(0).child(0);
    expect(outerCell.child(0).attrs.align).toBe('right'); // Outer
    // Nested table: its own style, not the enclosing one.
    expect(outerCell.child(1).child(0).child(0).child(0).attrs.align).toBe(
      'center',
    );
    // Back to Outer once the nested table closes.
    expect(outerCell.child(2).attrs.align).toBe('right');
    // A textbox is its own story: the table style must not reach its text.
    const box = outerCell.child(3).child(0).attrs.textbox as {
      blocks: { attrs: { align: string | null } }[];
    };
    expect(box.blocks[0].attrs.align).toBeNull();
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

  it('reads unequal column widths, by Word’s rules not the standard’s', async () => {
    // A trailing second section keeps doc.attrs.sections populated: a document
    // that is one single-column section carries none at all.
    const colsOf = async (cols: string) => {
      const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
        <w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${cols}</w:sectPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>
        <w:p><w:r><w:t>b</w:t></w:r></w:p>
        <w:sectPr><w:cols w:num="1"/></w:sectPr>
      </w:body></w:document>`;
      const { doc } = await importDocx(await makeDocx(documentXml));
      const sections = doc.attrs['sections'] as
        | { columns: { count: number; gap: number; cols?: unknown } }[]
        | null;
      return sections?.[0].columns;
    };
    const COLS = '<w:col w:w="8096" w:space="652"/><w:col w:w="3022"/>';

    // equalWidth="0": each column carries its own width and the space AFTER
    // it, and cols/@w:space is disregarded.
    expect(
      await colsOf(
        `<w:cols w:num="2" w:space="720" w:equalWidth="0">${COLS}</w:cols>`,
      ),
    ).toMatchObject({
      count: 2,
      cols: [
        { width: Math.round(8096 / 15), space: Math.round(652 / 15) },
        { width: Math.round(3022 / 15), space: 0 },
      ],
    });

    // Absent equalWidth is TRUE — the standard declares no default, "Word uses
    // a default value of true" (MS-OI29500 §17.6.4 b) — and when it is on the
    // col elements are ignored.
    expect(await colsOf(`<w:cols w:num="2">${COLS}</w:cols>`)).toEqual({
      count: 2,
      gap: 48,
    });
    expect(
      await colsOf(`<w:cols w:num="2" w:equalWidth="1">${COLS}</w:cols>`),
    ).toEqual({ count: 2, gap: 48 });

    // The count comes from @w:num, NOT from the number of children. The
    // standard says num is "ignored in favor of the number of child col
    // elements"; Word instead "requires that the value of the num attribute
    // matches" them, and assumes 1 when num is absent (§17.6.4 c). A file that
    // breaks that requirement gets equal columns rather than a made-up grid.
    expect(
      await colsOf(`<w:cols w:num="3" w:equalWidth="0">${COLS}</w:cols>`),
    ).toEqual({ count: 3, gap: 48 });
    expect(await colsOf(`<w:cols w:equalWidth="0">${COLS}</w:cols>`)).toEqual({
      count: 1,
      gap: 48,
    });
    // …and 45 is Word's ceiling on num (§17.6.4 a).
    expect((await colsOf('<w:cols w:num="90"/>'))?.count).toBe(45);
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

  it('keeps text wrapped in w:smartTag / w:customXml', async () => {
    // CT_SmartTagRun is a transparent wrapper: Word writes it when it
    // recognises a place or a date, and everything inside is ordinary text.
    // Real case from bc_rieng: "Việt Nam" with "Nam" inside nested tags.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:t xml:space="preserve">Việt </w:t></w:r>
        <w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">
          <w:smartTagPr><w:attr w:name="ProductID" w:val="x"/></w:smartTagPr>
          <w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="country-region">
            <w:r><w:t>Nam</w:t></w:r>
          </w:smartTag>
        </w:smartTag>
        <w:r><w:t>, tail</w:t></w:r>
      </w:p>
      <w:p>
        <w:hyperlink w:anchor="here">
          <w:smartTag w:element="place"><w:r><w:t>linked place</w:t></w:r></w:smartTag>
        </w:hyperlink>
      </w:p>
      <w:customXml w:element="block">
        <w:customXmlPr/>
        <w:p><w:r><w:t>block inside customXml</w:t></w:r></w:p>
      </w:customXml>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    // Nested smart tags unwrap all the way down.
    expect(doc.child(0).textContent).toBe('Việt Nam, tail');
    // A link's runs reach the same unwrapper, and keep the link mark.
    expect(doc.child(1).textContent).toBe('linked place');
    expect(
      doc
        .child(1)
        .child(0)
        .marks.map((m) => m.type.name),
    ).toContain('link');
    // Block-level w:customXml unwraps to its paragraphs.
    expect(doc.child(2).textContent).toBe('block inside customXml');
  });

  it('parses w:tl2br / w:br2tl as cell diagonals, not sides', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:tcBorders><w:tl2br w:val="single" w:sz="4" w:color="auto"/></w:tcBorders></w:tcPr><w:p/></w:tc>
          <w:tc><w:tcPr><w:tcBorders>
            <w:tl2br w:val="single" w:sz="4" w:color="auto"/>
            <w:br2tl w:val="single" w:sz="4" w:color="auto"/>
          </w:tcBorders></w:tcPr><w:p/></w:tc>
        </w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const row = doc.child(0).child(0);
    const struck = row.child(0).attrs.diagonals as Record<string, unknown>;
    expect(Object.keys(struck)).toEqual(['tl2br']);
    // A diagonal is not one of the four sides — it must not leak into them.
    expect(row.child(0).attrs.borders).toBeNull();
    // Both together are the X Vietnamese school tables use for "no data".
    expect(
      Object.keys(row.child(1).attrs.diagonals as Record<string, unknown>),
    ).toEqual(['tl2br', 'br2tl']);
  });

  it('applies a row w:tblPrEx in place of the table borders', async () => {
    // "Properties which shall be applied to the contents of this row in place
    // of the table properties" — Word writes it when two tables are merged.
    // Row 1 has the exception; row 0 keeps the table's own (absent) borders.
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
        <w:tr>
          <w:tc><w:p/></w:tc><w:tc><w:p/></w:tc>
        </w:tr>
        <w:tr>
          <w:tblPrEx><w:tblBorders>
            <w:top w:val="double" w:sz="8" w:color="FF0000"/>
            <w:bottom w:val="double" w:sz="8" w:color="FF0000"/>
            <w:left w:val="double" w:sz="8" w:color="FF0000"/>
            <w:right w:val="double" w:sz="8" w:color="FF0000"/>
            <w:insideH w:val="dashed" w:sz="4" w:color="0000FF"/>
            <w:insideV w:val="single" w:sz="4" w:color="00FF00"/>
          </w:tblBorders></w:tblPrEx>
          <w:tc><w:tcPr><w:tcBorders><w:left w:val="nil"/></w:tcBorders></w:tcPr><w:p/></w:tc>
          <w:tc><w:p/></w:tc>
        </w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const rows = doc.child(0);
    // The row without an exception is untouched.
    expect(rows.child(0).child(0).attrs.borders).toBeNull();
    const left = rows.child(1).child(0).attrs.borders as Record<
      string,
      unknown
    >;
    // Which of the exception's six sides reaches an edge depends on where
    // the cell sits, exactly as it does for the table's own set. This is the
    // LAST row's FIRST column: the bottom is the table's outline (so it takes
    // the exception's `bottom`) while the top is shared with row 0 (so it
    // takes `insideH`, NOT the exception's `top`).
    expect(left.top).toMatchObject({ style: 'dashed', color: '#0000FF' });
    expect(left.bottom).toMatchObject({ style: 'double', color: '#FF0000' });
    expect(left.right).toMatchObject({ style: 'solid', color: '#00FF00' });
    // The cell's own w:tcBorders still win — the exception replaces the
    // TABLE's edges, not the cell's.
    expect(left.left).toBe(false);
    const right = rows.child(1).child(1).attrs.borders as Record<
      string,
      unknown
    >;
    expect(right.left).toMatchObject({ color: '#00FF00' });
    expect(right.right).toMatchObject({ style: 'double' });
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

  it('imports w:kern as a threshold, resolved against the run size', async () => {
    // "The smallest font size which shall have its kerning automatically
    // adjusted; if the font size is smaller than this value, then no font
    // kerning shall be performed." Both in half-points. The style declares
    // the threshold, each run declares its own size — so the same threshold
    // bites one run and not the next.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:rPr><w:kern w:val="32"/><w:sz w:val="26"/></w:rPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>13pt body, below the threshold</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:t>18pt, above it</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    // The model keeps the DECLARED threshold on both runs — it is the layout
    // that compares, because only there is the run's size finally settled.
    const kern = (i: number) =>
      doc
        .child(i)
        .child(0)
        .marks.find((m) => m.type.name === 'kern')?.attrs['halfPoints'];
    expect(kern(0)).toBe(32);
    expect(kern(1)).toBe(32);
    expect(
      doc
        .child(1)
        .child(0)
        .marks.find((m) => m.type.name === 'fontSize')?.attrs['size'],
    ).toBe(18);
  });

  it('reads w:kern from docDefaults, a style and the run alike', async () => {
    // Word's default is NOT to kern, so the threshold has to survive every
    // layer it can be declared at — otherwise flipping the default would
    // silently un-kern documents that asked for kerning in styles.xml.
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:docDefaults><w:rPrDefault><w:rPr>
        <w:sz w:val="24"/><w:kern w:val="2"/>
      </w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
      <w:style w:type="paragraph" w:styleId="Big">
        <w:rPr><w:kern w:val="32"/></w:rPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>from docDefaults</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Big"/></w:pPr><w:r><w:t>from a style</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:kern w:val="24"/></w:rPr><w:t>from the run</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>inherits 2</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const halfPoints = (i: number) =>
      doc
        .child(i)
        .child(0)
        .marks.find((m) => m.type.name === 'kern')?.attrs['halfPoints'];
    expect(halfPoints(0)).toBe(2);
    expect(halfPoints(1)).toBe(32); // the style's threshold beats docDefaults
    expect(halfPoints(2)).toBe(24); // the run's beats both
    expect(halfPoints(3)).toBe(2); // still inherited when the run says nothing
  });

  it('imports w:w as a character scale, bare or percent-suffixed', async () => {
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p>
        <w:r><w:rPr><w:w w:val="80"/></w:rPr><w:t>bare</w:t></w:r>
        <w:r><w:rPr><w:w w:val="70%"/></w:rPr><w:t>suffixed</w:t></w:r>
        <w:r><w:rPr><w:w w:val="100"/></w:rPr><w:t>normal</w:t></w:r>
        <w:r><w:rPr><w:w w:val="80"/><w:spacing w:val="26"/></w:rPr><w:t>both</w:t></w:r>
        <w:r><w:t>plain</w:t></w:r>
      </w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml));
    const para = doc.child(0);
    const marks = (i: number) =>
      Object.fromEntries(
        para.child(i).marks.map((m) => [m.type.name, m.attrs]),
      );
    expect(marks(0)['charScale']).toEqual({ percent: 80 });
    expect(marks(1)['charScale']).toEqual({ percent: 70 });
    // 100 is Word's default but still an explicit override of a style.
    expect(marks(2)['charScale']).toEqual({ percent: 100 });
    // The two are independent and 17 runs in khbd carry both.
    expect(marks(3)['charScale']).toEqual({ percent: 80 });
    expect(marks(3)['letterSpacing']).toEqual({ twips: 26 });
    expect(marks(4)['charScale']).toBeUndefined();
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
      <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp xmlns:wps="${WPS_NS_G}"><wps:cNvSpPr/>
        <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="170815" cy="150495"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="19050"/></wps:spPr>
        <wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></wps:style>
      <wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>`;
    // Horizontal-rule style straight connector with a direct outline color.
    const line = `<w:drawing><wp:inline><wp:extent cx="952500" cy="0"/><wp:docPr id="6" name="Straight Connector 6"/>
      <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp xmlns:wps="${WPS_NS_G}">
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
      <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp xmlns:wps="${WPS_NS_G}"><wps:cNvSpPr txBox="1"/>
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

    const tb = node.attrs.textbox as { blocks: unknown[]; inset?: unknown };
    expect(tb).toBeTruthy();
    expect(tb.blocks).toHaveLength(2);
    // Formatting survives: paragraphs are real PM JSON with marks.
    const p0 = schema.nodeFromJSON(tb.blocks[0] as never);
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

  it('lets a border w:themeColor supersede its cached w:color', async () => {
    // "If the border specifies the use of a theme color via the themeColor
    // attribute, this value is superseded by the theme color value" — w:color
    // is the last-computed rendering, kept for consumers with no theme part.
    const themeXml = `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}"><a:themeElements><a:clrScheme name="Office">
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme></a:themeElements></a:theme>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pBdr>
        <w:top w:val="single" w:sz="6" w:color="STALE0" w:themeColor="accent1"/>
        <w:bottom w:val="single" w:sz="6" w:color="00FF00"/>
      </w:pBdr></w:pPr><w:r><w:t>ruled</w:t></w:r></w:p>
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
    const b = doc.child(0).attrs.borders as Record<
      string,
      { color: string } | false
    >;
    // The theme wins even though w:color says something else.
    expect((b.top as { color: string }).color).toBe('#4472C4');
    // No themeColor → the literal colour still applies.
    expect((b.bottom as { color: string }).color).toBe('#00FF00');
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

  it('resolves contextualSpacing against same-styled neighbours', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:spacing w:after="200"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="ListParagraph">
        <w:basedOn w:val="Normal"/><w:pPr><w:contextualSpacing/></w:pPr>
      </w:style>
    </w:styles>`;
    const p = (style: string | null, text: string) =>
      `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      ${p('ListParagraph', 'one')}
      ${p('ListParagraph', 'two')}
      ${p('ListParagraph', 'three')}
      ${p(null, 'body')}
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    const cs = (i: number) => doc.child(i).attrs['contextualSpacing'];
    // First item: nothing above it, a sibling below.
    expect(cs(0)).toEqual({ before: false, after: true });
    // Middle item: collapses on both sides.
    expect(cs(1)).toEqual({ before: true, after: true });
    // Last item: the paragraph below is a DIFFERENT style, so its space-after
    // survives — that is the gap Word keeps at the end of a list.
    expect(cs(2)).toEqual({ before: true, after: false });
    // The plain paragraph never carries the flag at all.
    expect(cs(3)).toBeNull();
    // The declared spacing is untouched — only the layout applies the collapse,
    // so export still writes what the document said.
    expect(doc.child(1).attrs['spacing']).toMatchObject({ after: 13 });
  });

  it('does not collapse between two different styles that both set the flag', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:spacing w:after="200"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="ListA">
        <w:basedOn w:val="Normal"/><w:pPr><w:contextualSpacing/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="ListB">
        <w:basedOn w:val="Normal"/><w:pPr><w:contextualSpacing/></w:pPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="ListA"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="ListB"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    // "Identical styles" means the same style, not merely the same flag.
    expect(doc.child(0).attrs['contextualSpacing']).toEqual({
      before: false,
      after: false,
    });
  });

  it('honours w:contextualSpacing w:val="false" over the style', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
      <w:style w:type="paragraph" w:styleId="L">
        <w:pPr><w:contextualSpacing/></w:pPr>
      </w:style>
    </w:styles>`;
    const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="L"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="L"/><w:contextualSpacing w:val="false"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>
    </w:body></w:document>`;
    const { doc } = await importDocx(await makeDocx(documentXml, stylesXml));
    expect(doc.child(0).attrs['contextualSpacing']).toEqual({
      before: false,
      after: true,
    });
    expect(doc.child(1).attrs['contextualSpacing']).toBeNull();
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

describe('DrawingML shape stroke + geometry adjust', () => {
  const dmlDoc = (spPr: string) =>
    `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:wps="${WPS_NS_G}"><w:body>
      <w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/>
        <a:graphic><a:graphicData><wps:wsp><wps:spPr>${spPr}</wps:spPr></wps:wsp></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>
    </w:body></w:document>`;
  const shapeOf = async (spPr: string) => {
    const { doc } = await importDocx(
      (await makeDocx(dmlDoc(spPr))).buffer as ArrayBuffer,
      { schema },
    );
    const imgs: import('prosemirror-model').Node[] = [];
    doc.descendants((n) => {
      if (n.type.name === 'image') imgs.push(n);
      return true;
    });
    return (imgs[0]?.attrs['shape'] ?? null) as Record<string, unknown> | null;
  };

  it('reads a:custDash exactly and a:ln@cap by its real tokens', async () => {
    // d/sp are thousandths of a percent of the LINE WIDTH: 100000 = 1×.
    const s = await shapeOf(
      `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>` +
        `<a:ln w="12700" cap="rnd"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>` +
        `<a:custDash><a:ds d="300000" sp="100000"/></a:custDash></a:ln>`,
    );
    expect(s).toMatchObject({ kind: 'line', dash: [3, 1], cap: 'round' });
  });

  it('takes roundRect adj as the corner ratio but ignores another preset adj', async () => {
    const round = await shapeOf(
      `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 12500"/></a:avLst></a:prstGeom>` +
        `<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`,
    );
    expect(round).toMatchObject({ kind: 'roundRect', cornerRatio: 0.125 });
    // Same attribute on horizontalScroll means the size of its curl — reading
    // it as roundness would deform the shape (khbd.docx has exactly this).
    const scroll = await shapeOf(
      `<a:prstGeom prst="horizontalScroll"><a:avLst><a:gd name="adj" fmla="val 12500"/></a:avLst></a:prstGeom>` +
        `<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`,
    );
    expect(scroll?.['cornerRatio']).toBeUndefined();
  });
});

describe('legacy VML shapes', () => {
  // Old Word draws flowcharts as w:pict + v:* — Word-verified against a real
  // report whose 5-gap диаграм was entirely roundrects + t32 connectors.
  const vmlDoc = (body: string) =>
    `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"><w:body>${body}</w:body></w:document>`;

  it('imports a positioned roundrect with its textbox content', async () => {
    const bytes = await makeDocx(
      vmlDoc(`<w:p><w:r><w:pict>
        <v:roundrect id="_x0000_s1041" style="position:absolute;margin-left:24.15pt;margin-top:5.3pt;width:81.75pt;height:23.25pt;z-index:251659264" arcsize="10923f">
          <v:textbox inset="1.55pt,.65pt,1.55pt,.65pt"><w:txbxContent>
            <w:p><w:r><w:t>Dịch vụ kì vọng</w:t></w:r></w:p>
          </w:txbxContent></v:textbox>
        </v:roundrect>
      </w:pict></w:r></w:p>`),
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    let img: import('prosemirror-model').Node | null = null;
    doc.descendants((n) => {
      if (n.type.name === 'image') img = n;
      return true;
    });
    expect(img).not.toBeNull();
    const a = img!.attrs;
    expect(a['shape']).toMatchObject({
      kind: 'roundRect',
      stroke: '#000000',
      fill: '#ffffff',
    });
    expect(a['width']).toBe(109); // 81.75pt
    expect(a['height']).toBe(31);
    expect(a['float']).toMatchObject({
      wrap: 'none',
      hOffset: 32, // 24.15pt
      vOffset: 7,
      hRel: 'margin',
      vRel: 'paragraph',
    });
    const tb = a['textbox'] as { blocks: unknown[]; inset?: object };
    expect(tb.blocks).toHaveLength(1);
    expect(JSON.stringify(tb.blocks)).toContain('Dịch vụ kì vọng');
    expect(tb.inset).toMatchObject({ l: 2, t: 1, r: 2, b: 1 });
  });

  it('keeps a table inside a textbox — txbxContent holds block content', async () => {
    // CT_TxbxContent is EG_BlockLevelElts: the same group w:body uses. Reading
    // only w:p dropped two whole tables from a factsheet whose textboxes are
    // invisible frames holding nothing else.
    const bytes = await makeDocx(
      vmlDoc(`<w:p><w:r><w:pict>
        <v:shape id="_x0000_s1032" type="#_x0000_t202" style="position:absolute;margin-left:24pt;margin-top:56pt;width:270pt;height:135pt" filled="f" stroked="f">
          <v:textbox inset="0,0,0,0"><w:txbxContent>
            <w:tbl>
              <w:tblGrid><w:gridCol w:w="3737"/><w:gridCol w:w="1729"/></w:tblGrid>
              <w:tr><w:tc><w:p><w:r><w:t>Điểm đến</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>Khoảng cách</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            <w:p><w:r><w:t>chú thích</w:t></w:r></w:p>
          </w:txbxContent></v:textbox>
        </v:shape>
      </w:pict></w:r></w:p>`),
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    let img: import('prosemirror-model').Node | null = null;
    doc.descendants((n) => {
      if (n.type.name === 'image') img = n;
      return true;
    });
    const tb = img!.attrs['textbox'] as { blocks: { type: string }[] };
    expect(tb.blocks.map((b) => b.type)).toEqual(['table', 'paragraph']);
    expect(JSON.stringify(tb.blocks[0])).toContain('Khoảng cách');

    // …and it must survive a save: emitting only the paragraphs would delete
    // the table the first time the customer pressed Ctrl+S.
    const out = await exportDocx(doc);
    const zip = await JSZip.loadAsync(out);
    const xml = (await zip.file('word/document.xml')?.async('string')) ?? '';
    const box = xml.slice(
      xml.indexOf('<w:txbxContent>'),
      xml.indexOf('</w:txbxContent>'),
    );
    expect(box).toContain('<w:tbl>');
    expect(box).toContain('Khoảng cách');
    expect(box).toContain('chú thích');
  });

  it('measures a VML float from the anchor its wrap names', async () => {
    const floatOf = async (shapeAttrs: string, wrap: string) => {
      const bytes = await makeDocx(
        vmlDoc(`<w:p><w:r><w:pict>
          <v:shape type="#_x0000_t202" style="position:absolute;margin-left:24pt;margin-top:5pt;width:60pt;height:30pt${shapeAttrs}">
            <v:textbox><w:txbxContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:txbxContent></v:textbox>
            ${wrap}
          </v:shape>
        </w:pict></w:r></w:p>`),
      );
      const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
      let img: import('prosemirror-model').Node | null = null;
      doc.descendants((n) => {
        if (n.type.name === 'image') img = n;
        return true;
      });
      return img!.attrs['float'] as Record<string, unknown>;
    };
    // w10:wrap/@anchorx — the factsheet's two table boxes carry exactly this,
    // and reading it as "margin" put them a left margin too far right.
    expect(await floatOf('', '<w10:wrap anchorx="page"/>')).toMatchObject({
      hRel: 'page',
      vRel: 'paragraph',
    });
    expect(
      await floatOf('', '<w10:wrap anchorx="margin" anchory="margin"/>'),
    ).toMatchObject({ hRel: 'margin', vRel: 'margin' });
    // An omitted anchorx must not beat a style that speaks: a file in the
    // corpus pairs <w10:wrap type="through"/> with mso-position-*-relative:text
    // on 15 shapes, and the spec's "assume page" default would move them all.
    expect(
      await floatOf(
        ';mso-position-horizontal-relative:text;mso-position-vertical-relative:text',
        '<w10:wrap type="through"/>',
      ),
    ).toMatchObject({ hRel: 'margin', vRel: 'paragraph' });
    // The style alone is enough — no w10:wrap element at all.
    expect(
      await floatOf(';mso-position-horizontal-relative:page', ''),
    ).toMatchObject({ hRel: 'page' });
  });

  it('reads arcsize in each encoding, halving it against the shorter side', async () => {
    const ratioOf = async (attr: string) => {
      const bytes = await makeDocx(
        vmlDoc(`<w:p><w:r><w:pict>
          <v:roundrect style="position:absolute;width:60pt;height:30pt"${attr}/>
        </w:pict></w:r></w:p>`),
      );
      const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
      let img: import('prosemirror-model').Node | null = null;
      doc.descendants((n) => {
        if (n.type.name === 'image') img = n;
        return true;
      });
      return (img!.attrs['shape'] as Record<string, number>)['cornerRatio'];
    };
    // VML states roundness against HALF the shorter side, the model against
    // the whole one — hence every value here is the file's number ÷ 2.
    expect(await ratioOf(' arcsize="10923f"')).toBeCloseTo(0.08333, 4); // 10923/65536/2
    expect(await ratioOf(' arcsize="25%"')).toBeCloseTo(0.125, 4);
    expect(await ratioOf(' arcsize="0.5"')).toBeCloseTo(0.25, 4);
    expect(await ratioOf(' arcsize="200%"')).toBeCloseTo(0.5, 4); // clamped
    expect(await ratioOf('')).toBeCloseTo(0.1, 4); // spec default 0.2 ÷ 2
  });

  it('takes alt text from the shape, never from its id', async () => {
    const bytes = await makeDocx(
      vmlDoc(`<w:p><w:r><w:pict>
        <v:roundrect id="_x0000_s1026" o:spid="_x0000_s1026" alt="Kỳ vọng của khách"
                     style="position:absolute;width:60pt;height:30pt"/>
      </w:pict></w:r></w:p>
      <w:p><w:r><w:pict>
        <v:roundrect id="_x0000_s1027" style="position:absolute;width:60pt;height:30pt"/>
      </w:pict></w:r></w:p>`),
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    const alts: string[] = [];
    doc.descendants((n) => {
      if (n.type.name === 'image') alts.push(n.attrs['alt'] as string);
      return true;
    });
    // The described shape says what it is; the undescribed one says nothing —
    // an identifier read aloud ("_x0000_s1027") is worse than silence.
    expect(alts).toEqual(['Kỳ vọng của khách', '']);
  });

  it('a shapetype declared in a PICTURE pict still resolves for a later shape', async () => {
    // The picture path returns before the shape parser, so a type declared
    // beside an image used to vanish — and the connector below it, whose kind
    // depends on that type, was dropped.
    const relsXml = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId9" Type="${R_NS}/image" Target="media/i.png"/></Relationships>`;
    const bytes = await makeDocx(
      vmlDoc(`<w:p><w:r><w:pict>
        <v:shapetype id="_x0000_t32" o:spt="32" coordsize="21600,21600"/>
        <v:shape id="Picture 2" type="#_x0000_t75" alt="Ảnh minh hoạ" style="width:75pt;height:75pt">
          <v:imagedata r:id="rId9" o:title=""/>
        </v:shape>
      </w:pict></w:r></w:p>
      <w:p><w:r><w:pict>
        <v:shape id="_x0000_s1030" type="#_x0000_t32"
                 style="position:absolute;width:120pt;height:0"/>
      </w:pict></w:r></w:p>`),
      undefined,
      undefined,
      relsXml,
      { 'i.png': PNG_1x1 },
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    const imgs: import('prosemirror-model').Node[] = [];
    doc.descendants((n) => {
      if (n.type.name === 'image') imgs.push(n);
      return true;
    });
    expect(imgs).toHaveLength(2);
    expect(imgs[0].attrs['alt']).toBe('Ảnh minh hoạ'); // from v:shape @alt
    expect((imgs[1].attrs['shape'] as { kind: string }).kind).toBe('line');
  });

  it('imports a t32 connector as an arrowed line via the shapetype registry', async () => {
    const bytes = await makeDocx(
      vmlDoc(`<w:p><w:r><w:pict>
        <v:shapetype id="_x0000_t32" coordsize="21600,21600" o:spt="32" o:oned="t" path="m,l21600,21600e" filled="f">
          <v:path arrowok="t" fillok="f" o:connecttype="none"/>
        </v:shapetype>
        <v:shape id="_x0000_s1028" type="#_x0000_t32" style="position:absolute;margin-left:28.05pt;margin-top:16.45pt;width:18.15pt;height:12pt;flip:y;z-index:251680256" o:connectortype="straight" strokeweight="1.5pt">
          <v:stroke endarrow="block"/>
        </v:shape>
      </w:pict></w:r></w:p>`),
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    let img: import('prosemirror-model').Node | null = null;
    doc.descendants((n) => {
      if (n.type.name === 'image') img = n;
      return true;
    });
    expect(img).not.toBeNull();
    const a = img!.attrs;
    expect(a['shape']).toMatchObject({
      kind: 'line',
      stroke: '#000000',
      strokeWidth: 2, // 1.5pt
      flipV: true,
      arrowEnd: true,
    });
    expect((a['shape'] as Record<string, unknown>)['fill']).toBeUndefined();
    expect(a['float']).toMatchObject({ wrap: 'none' });
  });

  it('a second pict resolves the type from a shapetype defined earlier', async () => {
    const bytes = await makeDocx(
      vmlDoc(`<w:p><w:r><w:pict>
          <v:shapetype id="_x0000_t32" o:spt="32" path="m,l21600,21600e"/>
          <v:shape type="#_x0000_t32" style="position:absolute;width:10pt;height:0"/>
        </w:pict></w:r><w:r><w:pict>
          <v:shape type="#_x0000_t32" style="position:absolute;width:20pt;height:0" strokecolor="white"/>
        </w:pict></w:r></w:p>`),
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    const shapes: Record<string, unknown>[] = [];
    doc.descendants((n) => {
      if (n.type.name === 'image')
        shapes.push(n.attrs['shape'] as Record<string, unknown>);
      return true;
    });
    expect(shapes).toHaveLength(2);
    expect(shapes[1]).toMatchObject({ kind: 'line', stroke: '#ffffff' });
  });

  it('unmapped VML stays dropped (no phantom nodes)', async () => {
    const bytes = await makeDocx(
      vmlDoc(
        `<w:p><w:r><w:pict><v:curve style="width:10pt;height:10pt"/></w:pict></w:r><w:r><w:t>sau</w:t></w:r></w:p>`,
      ),
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    let images = 0;
    doc.descendants((n) => {
      if (n.type.name === 'image') images++;
      return true;
    });
    expect(images).toBe(0);
    expect(doc.textContent).toBe('sau');
  });
});

describe('legacy VML: dash and vertical text', () => {
  const vmlDoc2 = (body: string) =>
    `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>${body}</w:body></w:document>`;
  const firstImage = (doc: import('prosemirror-model').Node) => {
    let img: import('prosemirror-model').Node | null = null;
    doc.descendants((n) => {
      if (n.type.name === 'image') img = n;
      return true;
    });
    return img!;
  };

  it('a dashstyle connector keeps the document dash lengths verbatim', async () => {
    const bytes = await makeDocx(
      vmlDoc2(`<w:p><w:r><w:pict>
        <v:shape style="position:absolute;width:120pt;height:0" o:connectortype="straight">
          <v:stroke dashstyle="1 1" startarrow="block" endarrow="block" endcap="round"/>
        </v:shape>
      </w:pict></w:r></w:p>`),
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    // "1 1" is dash 1 × stroke width, gap 1 × stroke width — kept verbatim
    // rather than flattened to a boolean the painter would have to guess at.
    expect(firstImage(doc).attrs['shape']).toMatchObject({
      kind: 'line',
      dash: [1, 1],
      cap: 'round',
      arrowStart: true,
      arrowEnd: true,
    });
  });

  it('a named dashstyle falls back to the generic pattern; solid stays solid', async () => {
    const one = async (stroke: string) => {
      const bytes = await makeDocx(
        vmlDoc2(`<w:p><w:r><w:pict>
          <v:shape style="position:absolute;width:120pt;height:0" o:connectortype="straight">
            ${stroke}
          </v:shape>
        </w:pict></w:r></w:p>`),
      );
      const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
      return firstImage(doc).attrs['shape'] as Record<string, unknown>;
    };
    // The spec names the styles but not their lengths, so a named one keeps
    // the generic dash instead of a table invented here.
    expect((await one('<v:stroke dashstyle="longdashdot"/>'))['dash']).toEqual([
      4, 3,
    ]);
    expect(
      (await one('<v:stroke dashstyle="solid"/>'))['dash'],
    ).toBeUndefined();
    // flat is the spec default for endcap — nothing to model.
    expect((await one('<v:stroke endcap="flat"/>'))['cap']).toBeUndefined();
    expect((await one('<v:stroke endcap="square"/>'))['cap']).toBe('square');
  });

  it('layout-flow:vertical rotates the box, swapping dims around its center', async () => {
    const bytes = await makeDocx(
      vmlDoc2(`<w:p><w:r><w:pict>
        <v:roundrect style="position:absolute;margin-left:100pt;margin-top:60pt;width:30pt;height:90pt">
          <v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent>
            <w:p><w:r><w:t>Khách hàng</w:t></w:r></w:p>
          </w:txbxContent></v:textbox>
        </v:roundrect>
      </w:pict></w:r></w:p>`),
    );
    const { doc } = await importDocx(bytes.buffer as ArrayBuffer, { schema });
    const a = firstImage(doc).attrs;
    // 30x90pt = 40x120px box → laid out sideways 120x40, rotated -90 (CCW,
    // bottom-to-top), re-centered on the original rect.
    expect(a['rotation']).toBe(-90);
    expect(a['width']).toBe(120);
    expect(a['height']).toBe(40);
    // center preserved: left 100pt=133.33px + (40-120)/2 → 93; top 60pt=80px + (120-40)/2 → 120
    expect(a['float']).toMatchObject({ hOffset: 93, vOffset: 120 });
    expect(JSON.stringify(a['textbox'])).toContain('Khách hàng');
  });
});
