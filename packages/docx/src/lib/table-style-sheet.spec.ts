import { describe, expect, it } from 'vitest';
import { cellStyleLayer } from '@shadow-garden/bapbong-contracts';
import { buildTableStyleSheet } from './docx';
import { parseXml } from './ooxml';
import { buildStyleRegistry } from './styles';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// The probe-F2 style, verbatim shape: Word's rendering of exactly this XML is
// what the PDF measured, so the sheet built from it is pinned to ground truth.
const G = '9BBB59';
const side = (s: string, val = 'single', sz = '8') =>
  `<w:${s} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${G}"/>`;
const STYLES = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>
    <w:tblPr><w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="table" w:styleId="LightGridProbe"><w:name w:val="Light Grid Probe"/>
    <w:basedOn w:val="TableNormal"/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>
      <w:tblBorders>${side('top')}${side('left')}${side('bottom')}${side('right')}${side('insideH')}${side('insideV')}</w:tblBorders>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr>
      <w:tcBorders>${side('top')}${side('bottom', 'single', '18')}${side('insideH', 'nil', '0')}${side('insideV')}</w:tcBorders>
    </w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="lastRow"><w:pPr><w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:b/></w:rPr><w:tcPr>
      <w:tcBorders>${side('top', 'double', '6')}</w:tcBorders>
    </w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="band1Horz"><w:tcPr>
      <w:shd w:val="clear" w:color="auto" w:fill="D6E3BC"/>
    </w:tcPr></w:tblStylePr>
  </w:style></w:styles>`;

function sheetFor(ids: string[]) {
  const root = parseXml(STYLES);
  return buildTableStyleSheet(ids, buildStyleRegistry(root));
}

/** Word's default gates (w:tblLook absent): firstRow + firstCol + hBand. */
const WORD_LOOK = {
  firstRow: true,
  lastRow: false,
  firstCol: true,
  lastCol: false,
  hBand: true,
  vBand: false,
};

describe('buildTableStyleSheet', () => {
  it('resolves a Light-Grid-shaped style into plain values', () => {
    const style = sheetFor(['LightGridProbe'])['LightGridProbe'];
    expect(style.bands).toEqual({ row: 1, col: 1 });
    // Table borders: six #9BBB59 sides; cell margins from TableNormal via
    // basedOn (108tw ≈ 7px after twipsToPx rounding).
    expect(style.table.borders?.top).toMatchObject({
      style: 'solid',
      color: '#9BBB59',
    });
    expect(style.table.borders?.insideV).toMatchObject({ color: '#9BBB59' });
    expect(style.table.cellPadding).toEqual({
      top: 0,
      left: 7,
      bottom: 0,
      right: 7,
    });
  });

  it('conditional branches keep region borders, fills, and fonts', () => {
    const style = sheetFor(['LightGridProbe'])['LightGridProbe'];
    const fr = style.cond.firstRow;
    expect(fr?.font).toEqual({ bold: true });
    // sz is eighths of a point: 18/8 pt vs 8/8 pt — the header's bottom edge
    // is heavier than its top, and insideH val="nil" is an EXPLICIT false.
    const top = fr?.borders?.top;
    const bottom = fr?.borders?.bottom;
    if (!top || !bottom) throw new Error('header edges missing from sheet');
    expect(bottom.width).toBeGreaterThan(top.width);
    expect(fr?.borders?.insideH).toBe(false);
    expect(style.cond.band1Horz).toEqual({ background: '#D6E3BC' });
  });

  it('builds EVERY branch, gated or not — look changes at runtime', () => {
    // The old baked path only ever read the branches the document's tblLook
    // admitted; the sheet must carry lastRow too, because the editor can flip
    // that bit long after import.
    const style = sheetFor(['LightGridProbe'])['LightGridProbe'];
    expect(style.cond.lastRow?.font).toEqual({ bold: true });
    expect(style.cond.lastRow?.borders?.top).toMatchObject({
      style: 'double',
    });
  });

  it('feeds cellStyleLayer to the same answers Word gave in probe F2', () => {
    const style = sheetFor(['LightGridProbe'])['LightGridProbe'];
    const look = {
      firstRow: true,
      lastRow: false,
      firstCol: true,
      lastCol: false,
      hBand: true,
      vBand: false,
    };
    const pos = (row: number, col: number) => ({
      row,
      rowCount: 5,
      col,
      colspan: 1,
      colCount: 5,
      header: false,
    });
    // Header cell: bold, heavy bottom, no fill (PDF table A, r1).
    const header = cellStyleLayer(style, look, pos(0, 1));
    expect(header.font?.bold).toBe(true);
    expect(header.background).toBeUndefined();
    // Body band row: fill, no bold (PDF table A, r2/r4).
    const band = cellStyleLayer(style, look, pos(1, 2));
    expect(band.font?.bold).toBeUndefined();
    expect(band.background).toBe('#D6E3BC');
    // lastRow branch exists in the sheet but the look gates it off (r5).
    const last = cellStyleLayer(style, look, pos(4, 2));
    expect(last.font?.bold).toBeUndefined();
  });

  it('carries the pPr spacing of the style and of every branch', () => {
    // The paragraph slot of Word's cascade, as plain px/multipliers: the
    // style single-spaces its cells (after 0, line ×1) and lastRow adds a
    // 120tw = 8px gap above. Reading the branch here is also what clears
    // the audit's last UNKNOWN on the Kpop corpus file — a gated branch's
    // w:pPr/w:spacing that the baked path never had a reason to visit.
    const style = sheetFor(['LightGridProbe'])['LightGridProbe'];
    expect(style.paragraph).toEqual({
      spacing: { after: 0, line: 1, lineRule: 'auto' },
    });
    expect(style.cond.lastRow?.paragraph).toEqual({
      spacing: { before: 8, after: 0, line: 1, lineRule: 'auto' },
    });
    expect(style.cond.firstRow?.paragraph).toBeUndefined();
    // cellStyleLayer stacks them: body cells get the style's, the last row
    // (gate on) the branch's on top.
    const on = { ...WORD_LOOK, lastRow: true };
    const pos = (row: number) => ({
      row,
      rowCount: 5,
      col: 1,
      colspan: 1,
      colCount: 5,
      header: false,
    });
    expect(cellStyleLayer(style, on, pos(2)).spacing).toEqual({
      after: 0,
      line: 1,
      lineRule: 'auto',
    });
    expect(cellStyleLayer(style, on, pos(4)).spacing?.before).toBe(8);
    expect(cellStyleLayer(style, WORD_LOOK, pos(4)).spacing?.before).toBe(
      undefined,
    );
  });
});
