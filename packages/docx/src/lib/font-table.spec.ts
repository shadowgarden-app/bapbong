import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { importDocx } from './docx.js';
import { exportDocx } from './export.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
  <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial MT" w:hAnsi="Arial MT"/></w:rPr><w:t>3500</w:t></w:r></w:p>
  <w:sectPr/>
</w:body></w:document>`;

// The shapes Word writes: a PDF converter's font with an altName, a symbol
// font calling itself "roman", an East Asian font with a localized altName,
// a font whose altName is itself, and one that says nothing usable.
const FONT_TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="${W_NS}">
  <w:font w:name="Arial MT"><w:altName w:val="Arial"/><w:charset w:val="01"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>
  <w:font w:name="Wingdings 2"><w:charset w:val="02"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
  <w:font w:name="MS Mincho"><w:altName w:val="ＭＳ 明朝"/><w:charset w:val="80"/><w:family w:val="modern"/><w:pitch w:val="fixed"/></w:font>
  <w:font w:name="Trebuchet MS"><w:altName w:val="Trebuchet MS"/><w:charset w:val="00"/><w:family w:val="swiss"/></w:font>
  <w:font w:name="Lucida Typewriter"><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="fixed"/></w:font>
  <w:font w:name="VNI-Times"><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>
</w:fonts>`;

async function makeDocx(fontTable?: string): Promise<ArrayBuffer> {
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
  if (fontTable) zip.file('word/fontTable.xml', fontTable);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return bytes.slice().buffer as ArrayBuffer;
}

describe('word/fontTable.xml', () => {
  it('records each font’s substitutes on the doc', async () => {
    const { doc } = await importDocx(await makeDocx(FONT_TABLE_XML));
    expect(doc.attrs['fontSubstitutes']).toEqual({
      'Arial MT': { altNames: ['Arial'], generic: 'swiss' },
      // Symbol and East Asian charsets take no class stand-in.
      'MS Mincho': { altNames: ['ＭＳ 明朝'] },
      // An altName naming the font itself is no alternative.
      'Trebuchet MS': { generic: 'swiss' },
      'Lucida Typewriter': { generic: 'roman', fixedPitch: true },
    });
  });

  it('is absent when the package has no font table', async () => {
    const { doc } = await importDocx(await makeDocx());
    expect(doc.attrs['fontSubstitutes']).toBeNull();
  });

  it('leaves the runs, and so the saved file, on the font they name', async () => {
    const opened = await importDocx(await makeDocx(FONT_TABLE_XML));
    const run = opened.doc.firstChild!.firstChild!;
    expect(
      run.marks.find((m) => m.type.name === 'fontFamily')?.attrs['family'],
    ).toBe('Arial MT');
    const saved = await JSZip.loadAsync(
      await exportDocx(opened.doc, { carry: opened.raw }),
    );
    const xml = await saved.file('word/document.xml')!.async('string');
    expect(xml).toContain('w:ascii="Arial MT"');
    expect(xml).not.toMatch(/w:ascii="Arial"/);
  });
});
