import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  beginParagraphSpacingShadow,
  endParagraphSpacingShadow,
  importDocx,
} from './docx';
import JSZip from 'jszip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * The gate in front of un-baking paragraph spacing: for every paragraph the
 * importer sees, the full baked cascade must equal what the layout will
 * re-stack from the pieces the model now carries — the doc's floor
 * (paragraphDefaults), the table style's slot (the sheet), and the
 * paragraph's own attr. While the importer still computes both, any
 * disagreement is cheap to see and attribute.
 */

async function docxOf(documentXml: string, stylesXml: string) {
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
  zip.file('word/styles.xml', stylesXml);
  return zip.generateAsync({ type: 'uint8array' });
}

async function shadowImport(bytes: Uint8Array) {
  beginParagraphSpacingShadow();
  let result;
  try {
    result = await importDocx(bytes);
  } catch (e) {
    endParagraphSpacingShadow();
    throw e;
  }
  return { ...endParagraphSpacingShadow(), doc: result.doc };
}

// Word's stock shape: docDefaults with spacing, a Normal that says nothing,
// a Light-Grid-like table style whose own pPr single-spaces and whose lastRow
// branch adds a gap above.
const STYLES = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:pPrDefault><w:pPr>
    <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
  </w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Tight"><w:name w:val="Tight"/>
    <w:pPr><w:spacing w:after="40"/></w:pPr>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>
  <w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/>
    <w:basedOn w:val="TableNormal"/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:tblStylePr w:type="lastRow"><w:pPr><w:spacing w:before="120"/></w:pPr></w:tblStylePr>
  </w:style></w:styles>`;

const p = (inner: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${inner}</w:t></w:r></w:p>`;
const row = (...cells: string[]) =>
  `<w:tr>${cells.map((c) => `<w:tc>${c}</w:tc>`).join('')}</w:tr>`;
const DOCUMENT = `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
  ${p('body')}
  ${p('tight', '<w:pStyle w:val="Tight"/>')}
  <w:tbl>
    <w:tblPr><w:tblStyle w:val="Grid"/><w:tblLook w:val="0040" w:lastRow="1"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
    ${row(p('a1'), p('a2', '<w:spacing w:after="100"/>'))}
    ${row(p('b1', '<w:pStyle w:val="Tight"/>'), p('b2'))}
  </w:tbl>
</w:body></w:document>`;

describe('paragraph-spacing un-bake', () => {
  it('leaves the paragraph attr with only what sits above the table slot', async () => {
    const { doc, paragraphs, mismatches } = await shadowImport(
      await docxOf(DOCUMENT, STYLES),
    );
    expect(mismatches).toEqual([]);
    expect(paragraphs).toBe(6);
    // The floor moved to the doc (200tw ≈ 13px, 276/240 = 1.15).
    expect(doc.attrs['paragraphDefaults']).toEqual({
      spacing: { after: 13, line: 1.15, lineRule: 'auto' },
    });
    // A body paragraph on the floor alone: nothing of its own.
    expect(doc.child(0).attrs['spacing']).toBeNull();
    // A paragraph style is ABOVE the slot: it stays in the attr.
    expect(doc.child(1).attrs['spacing']).toEqual({ after: 3 });
    const cell = (r: number, c: number) =>
      doc.child(2).child(r).child(c).child(0);
    // In the styled table the slot is the sheet's, not the paragraph's…
    expect(cell(0, 0).attrs['spacing']).toBeNull();
    expect(doc.attrs['tableStyles']['Grid']).toMatchObject({
      paragraph: { spacing: { after: 0, line: 1, lineRule: 'auto' } },
      cond: { lastRow: { paragraph: { spacing: { before: 8 } } } },
    });
    // …while direct formatting and paragraph styles still ride the attr.
    expect(cell(0, 1).attrs['spacing']).toEqual({ after: 7 });
    expect(cell(1, 0).attrs['spacing']).toEqual({ after: 3 });
  });

  it('promotes an auto-spacing flag set below the slot into the attr', async () => {
    // docDefaults asks for automatic space-after: the flag lives under the
    // table slot, but which side is automatic is decided at import from the
    // whole cascade, so it must reach the attr (with the number
    // resolveAutoSpacing works out) for the layout and exporter to see.
    const styles = STYLES.replace(
      '<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>',
      '<w:spacing w:after="200" w:afterAutospacing="1"/>',
    );
    const { doc, mismatches } = await shadowImport(
      await docxOf(DOCUMENT, styles),
    );
    expect(mismatches).toEqual([]);
    expect(doc.child(0).attrs['spacing']).toMatchObject({ afterAuto: true });
    expect(typeof doc.child(0).attrs['spacing'].after).toBe('number');
    // A layer above that names a literal after does not cancel the flag
    // (ECMA-376 §17.3.1.33: the literal is ignored while the flag is on).
    expect(doc.child(1).attrs['spacing']).toMatchObject({ afterAuto: true });
  });
});

describe('paragraph-spacing shadow: real corpus', () => {
  const dir = path.resolve(
    path.dirname(new URL(import.meta.url).pathname),
    '../../../../apps/playground/public',
  );
  const files = fs.existsSync(dir)
    ? fs
        .readdirSync(dir)
        .filter((f) => f.endsWith('.docx'))
        .filter((f) => fs.statSync(path.join(dir, f)).size < 3_000_000)
    : [];
  let corpusParagraphs = 0;
  it.skipIf(files.length === 0).each(files)(
    '%s',
    async (f) => {
      const bytes = new Uint8Array(fs.readFileSync(path.join(dir, f)));
      const { paragraphs, mismatches } = await shadowImport(bytes);
      corpusParagraphs += paragraphs;
      expect(mismatches).toEqual([]);
    },
    30000,
  );
  it.skipIf(files.length === 0)('compared real paragraphs, not nothing', () => {
    expect(corpusParagraphs).toBeGreaterThan(1000);
  });
});
