/**
 * `w:color w:val="auto"` — Word's Automatic colour — REPLACES a colour the
 * style cascade brings in (D-2609-LSTQ: headings styled blue, their runs
 * saying auto, painted black by Word and blue by us). Measured in Word 365
 * (probe K1/K2): auto under a blue paragraph style and under a red
 * character style both paint black.
 *
 * Runs carry their resolved formatting as marks, so Automatic is simply no
 * colour mark — and a save must say so explicitly wherever the paragraph's
 * style would otherwise colour the run in Word.
 */
import JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';
import { importDocx } from './docx';
import { exportDocx } from './export';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W_NS}">` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>` +
  `<w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="4F81BD"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="RedChar"><w:name w:val="Red Char"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>` +
  `</w:styles>`;

const NUMBERING_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W_NS}">` +
  `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
  `<w:lvlText w:val="CHƯƠNG %1."/><w:rPr><w:b/></w:rPr></w:lvl></w:abstractNum>` +
  `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

const para = (pPr: string, rPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}"><w:body>` +
  // The report's heading: style blue, mark and run both Automatic.
  para(
    '<w:pStyle w:val="Heading2"/><w:rPr><w:color w:val="auto"/></w:rPr>',
    '<w:color w:val="auto"/>',
    'auto heading',
  ) +
  para('<w:pStyle w:val="Heading2"/>', '', 'styled heading') +
  para('', '<w:rStyle w:val="RedChar"/><w:color w:val="auto"/>', 'auto char') +
  para('', '', 'plain') +
  // The report's chapter label: a red paragraph mark, the level bold.
  para(
    '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:rPr><w:color w:val="FF0000"/><w:sz w:val="26"/></w:rPr>',
    '<w:color w:val="FF0000"/>',
    'chapter',
  ) +
  `</w:body></w:document>`;

async function makeDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

const colorOf = (p: PMNode): string | null => {
  let c: string | null = null;
  p.forEach((n) => {
    const m = n.marks.find((k) => k.type.name === 'textColor');
    if (m) c = m.attrs['color'] as string;
  });
  return c;
};

/** Each w:p of the saved body, as XML. */
async function savedParagraphs(doc: PMNode, carry?: JSZip) {
  const zip = await JSZip.loadAsync(await exportDocx(doc, { carry }));
  const xml = await zip.file('word/document.xml')!.async('string');
  return xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [];
}

describe('Automatic colour (w:color="auto")', () => {
  it('replaces the colour of the paragraph style', async () => {
    const { doc } = await importDocx(await makeDocx());
    expect(doc.child(0).textContent).toBe('auto heading');
    expect(colorOf(doc.child(0))).toBeNull();
    // Without it the heading keeps the style's blue.
    expect(colorOf(doc.child(1))).toBe('#4F81BD');
  });

  it('replaces the colour of the character style', async () => {
    const { doc } = await importDocx(await makeDocx());
    expect(colorOf(doc.child(2))).toBeNull();
  });

  it('leaves the paragraph mark without a colour', async () => {
    const { doc } = await importDocx(await makeDocx());
    expect(doc.child(0).attrs['markFont']?.color).toBeUndefined();
  });

  it('is written out where the style would colour the run', async () => {
    const { doc, raw } = await importDocx(await makeDocx());
    const [autoHeading, styled, autoChar, plain] = await savedParagraphs(
      doc,
      raw,
    );
    // Under Heading2 (blue): the run AND the mark say auto.
    expect(autoHeading).toContain('<w:pStyle w:val="Heading2"/>');
    expect(autoHeading?.match(/<w:color w:val="auto"\/>/g)).toHaveLength(2);
    expect(styled).toContain('<w:color w:val="4F81BD"/>');
    // Normal colours nothing: no colour written at all.
    expect(autoChar).not.toContain('<w:color');
    expect(plain).not.toContain('<w:color');
  });

  it('reads back as Automatic after a save', async () => {
    const { doc, raw } = await importDocx(await makeDocx());
    const again = await importDocx(await exportDocx(doc, { carry: raw }));
    expect(colorOf(again.doc.child(0))).toBeNull();
    expect(colorOf(again.doc.child(1))).toBe('#4F81BD');
  });
});

describe('the paragraph mark colour', () => {
  it('is read into markFont — the list label is drawn in it', async () => {
    const { doc } = await importDocx(await makeDocx());
    expect(doc.child(4).attrs['markFont']).toMatchObject({
      color: '#FF0000',
      sizePt: 13,
    });
  });

  it('is written back once, from markFont', async () => {
    const { doc, raw } = await importDocx(await makeDocx());
    const chapter = (await savedParagraphs(doc, raw))[4];
    const mark = /<w:pPr>[\s\S]*?<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(chapter)?.[1];
    expect(mark?.match(/<w:color\b[^>]*\/>/g)).toEqual([
      '<w:color w:val="FF0000"/>',
    ]);
  });
});
