import JSZip from 'jszip';
import { Mark } from 'prosemirror-model';
import { importDocx } from './docx';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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
async function makeDocx(documentXml: string, stylesXml?: string): Promise<Uint8Array> {
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
  return zip.generateAsync({ type: 'uint8array' });
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

  it('throws when word/document.xml is missing', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'nope');
    await expect(importDocx(await zip.generateAsync({ type: 'uint8array' }))).rejects.toThrow(
      /document\.xml/,
    );
  });
});
