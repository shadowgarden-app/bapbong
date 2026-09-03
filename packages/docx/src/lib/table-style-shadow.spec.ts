import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { beginTableStyleShadow, endTableStyleShadow, importDocx } from './docx';
import JSZip from 'jszip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * The shadow check: import with the sheet path computing every cell's style
 * layer ALONGSIDE the baked path, and demand they agree. This is the gate in
 * front of the flip that stops baking — while both paths are live and the
 * importer still knows which formatting is direct and which came from the
 * style, any disagreement is cheap to see and attribute.
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

async function shadowImport(
  bytes: Uint8Array,
): Promise<{ cells: number; mismatches: string[] }> {
  beginTableStyleShadow();
  try {
    await importDocx(bytes);
  } catch (e) {
    endTableStyleShadow();
    throw e;
  }
  return endTableStyleShadow();
}

const G = '9BBB59';
const side = (s: string, val = 'single', sz = '8') =>
  `<w:${s} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${G}"/>`;
// The probe-F2 style again — every branch kind in one place.
const STYLES = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>
    <w:tblPr><w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="table" w:styleId="Probe"><w:name w:val="Probe"/>
    <w:basedOn w:val="TableNormal"/>
    <w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>
      <w:tblBorders>${side('top')}${side('left')}${side('bottom')}${side('right')}${side('insideH')}${side('insideV')}</w:tblBorders>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr>
      <w:tcBorders>${side('top')}${side('bottom', 'single', '18')}${side('insideH', 'nil', '0')}${side('insideV')}</w:tcBorders>
      <w:shd w:val="clear" w:color="auto" w:fill="1F4E79"/>
    </w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="lastRow"><w:rPr><w:b/><w:i/></w:rPr><w:tcPr>
      <w:tcBorders>${side('top', 'double', '6')}</w:tcBorders>
    </w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>
    <w:tblStylePr w:type="lastCol"><w:tcPr><w:vAlign w:val="center"/></w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D6E3BC"/></w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="band1Vert"><w:tcPr><w:tcMar><w:left w:w="240" w:type="dxa"/></w:tcMar><w:shd w:val="clear" w:color="auto" w:fill="EAF1DD"/></w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="band2Horz"><w:tcPr><w:tcBorders>${side('insideV', 'nil', '0')}</w:tcBorders></w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="nwCell"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr></w:tblStylePr>
  </w:style></w:styles>`;

function probeTable(look: string, rows = 5, cols = 5): string {
  const grid = Array.from(
    { length: cols },
    () => '<w:gridCol w:w="1700"/>',
  ).join('');
  const trs = Array.from(
    { length: rows },
    (_, r) =>
      `<w:tr>${Array.from(
        { length: cols },
        (_, c) => `<w:tc><w:p><w:r><w:t>r${r}c${c}</w:t></w:r></w:p></w:tc>`,
      ).join('')}</w:tr>`,
  ).join('');
  return `<w:tbl><w:tblPr><w:tblStyle w:val="Probe"/>${look}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${trs}</w:tbl>`;
}

describe('table-style shadow: sheet application ≡ baked application', () => {
  it('agrees on every cell across the tblLook matrix', async () => {
    const looks = [
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
      '<w:tblLook w:val="01E0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/>',
      '<w:tblLook w:val="0740" w:firstRow="0" w:lastRow="1" w:firstColumn="0" w:lastColumn="1" w:noHBand="1" w:noVBand="0"/>',
      '',
      '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>',
    ];
    const body = looks.map((l) => probeTable(l)).join('<w:p/>');
    const bytes = await docxOf(
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`,
      STYLES,
    );
    const { cells, mismatches } = await shadowImport(bytes);
    expect(mismatches).toEqual([]);
    expect(cells).toBe(5 * 5 * 5);
  });

  it('a branch vAlign reset ("top") agrees between the paths', async () => {
    // Word treats vAlign values other than center/bottom as clearing an
    // inherited centring; the sheet models that as an explicit null. The
    // shadow caught this as a mismatch before the sheet could express it.
    const styles = STYLES.replace(
      '<w:tblStylePr w:type="lastCol"><w:tcPr><w:vAlign w:val="center"/></w:tcPr></w:tblStylePr>',
      '<w:tblStylePr w:type="lastCol"><w:tcPr><w:vAlign w:val="top"/></w:tcPr></w:tblStylePr>',
    );
    const table = probeTable(
      '<w:tblLook w:val="01E0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/>',
      3,
      3,
    );
    const bytes = await docxOf(
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>${table}</w:body></w:document>`,
      styles,
    );
    const { cells, mismatches } = await shadowImport(bytes);
    expect(mismatches).toEqual([]);
    expect(cells).toBe(9);
  });

  it('direct cell formatting does not disturb the comparison', async () => {
    // The shadow compares the STYLE layer only; a cell's own shading, borders
    // and margins ride above it on both paths.
    const table = probeTable(
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
      3,
      3,
    ).replace(
      '<w:tc><w:p><w:r><w:t>r1c1</w:t></w:r></w:p></w:tc>',
      `<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/><w:tcBorders>${side('top', 'single', '24')}</w:tcBorders><w:tcMar><w:left w:w="500" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t>r1c1</w:t></w:r></w:p></w:tc>`,
    );
    const bytes = await docxOf(
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>${table}</w:body></w:document>`,
      STYLES,
    );
    expect((await shadowImport(bytes)).mismatches).toEqual([]);
  });

  it('merged cells (gridSpan) keep the two paths aligned', async () => {
    const table = `<w:tbl><w:tblPr><w:tblStyle w:val="Probe"/><w:tblLook w:val="01E0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="1700"/><w:gridCol w:w="1700"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>`;
    const bytes = await docxOf(
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>${table}</w:body></w:document>`,
      STYLES,
    );
    expect((await shadowImport(bytes)).mismatches).toEqual([]);
  });
});

describe('the import carries the live-theming data', () => {
  it('styled tables get styleId/look attrs and the doc gets the sheet', async () => {
    const bytes = await docxOf(
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>${probeTable(
        '<w:tblLook w:val="0740" w:firstRow="0" w:lastRow="1" w:firstColumn="0" w:lastColumn="1" w:noHBand="1" w:noVBand="0"/>',
        3,
        3,
      )}<w:tbl><w:tblGrid><w:gridCol w:w="1700"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`,
      STYLES,
    );
    const { doc } = await importDocx(bytes);
    const styled = doc.child(0);
    expect(styled.attrs['styleId']).toBe('Probe');
    expect(styled.attrs['look']).toEqual({
      firstRow: false,
      lastRow: true,
      firstCol: false,
      lastCol: true,
      hBand: false,
      vBand: true,
    });
    // A table naming no style stays on the fully-baked path: no attrs.
    expect(doc.child(1).attrs['styleId']).toBeNull();
    expect(doc.child(1).attrs['look']).toBeNull();
    const sheet = doc.attrs['tableStyles'] as Record<string, unknown>;
    expect(Object.keys(sheet)).toEqual(['Probe']);
    expect(sheet['Probe']).toMatchObject({ bands: { row: 1, col: 1 } });
  });
});

describe('the explicit clear survives the model and the save', () => {
  it('shd fill="auto" in a styled table imports as false and exports as auto', async () => {
    const table = probeTable(
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
      2,
      2,
    ).replace(
      '<w:tc><w:p><w:r><w:t>r0c0</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:r><w:t>r0c0</w:t></w:r></w:p></w:tc>',
    );
    const bytes = await docxOf(
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>${table}</w:body></w:document>`,
      STYLES,
    );
    const { doc } = await importDocx(bytes);
    const cell = doc.child(0).child(0).child(0);
    expect(cell.attrs['background']).toBe(false);
    expect(cell.attrs['vAlign']).toBe(false);
    // And the plain cell next to it declares nothing at all.
    expect(doc.child(0).child(0).child(1).attrs['background']).toBeNull();
  });
});

describe('table-style shadow: real corpus', () => {
  // The playground's sample corpus, when this checkout has it (the package's
  // own __fixtures__ stay self-contained). Files over 3MB are skipped for
  // spec-runtime, not correctness — the styles they carry also appear in the
  // smaller files.
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
  let corpusCells = 0;
  it.skipIf(files.length === 0).each(files)(
    '%s',
    async (f) => {
      const bytes = new Uint8Array(fs.readFileSync(path.join(dir, f)));
      const { cells, mismatches } = await shadowImport(bytes);
      corpusCells += cells;
      expect(mismatches).toEqual([]);
    },
    30000,
  );
  it.skipIf(files.length === 0)('compared real cells, not nothing', () => {
    expect(corpusCells).toBeGreaterThan(100);
  });
});
