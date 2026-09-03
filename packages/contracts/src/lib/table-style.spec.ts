import { describe, expect, it } from 'vitest';
import {
  branchRegion,
  cellStyleLayer,
  condTypesFor,
  type ResolvedTableStyle,
  type TableBandSizes,
  type TableCellPos,
  type TableLook,
} from './table-style.js';

const look = (on: Partial<TableLook>): TableLook => ({
  firstRow: false,
  lastRow: false,
  firstCol: false,
  lastCol: false,
  hBand: false,
  vBand: false,
  ...on,
});
const BANDS: TableBandSizes = { row: 1, col: 1 };
const at = (row: number, col: number): TableCellPos => ({
  row,
  rowCount: 5,
  col,
  colspan: 1,
  colCount: 5,
  header: false,
});

// Ground truth: probe F2 — one Light-Grid-shaped style rendered by Word under
// four tblLook combinations, read back cell-by-cell from Word's own PDF
// (bold = rPr branches, fills = shd branches). The expectations below are
// those measurements, not our own reasoning.
describe('condTypesFor (probe F2 matrices)', () => {
  it('table A/D — Word default 04A0: firstRow + firstCol + hBand', () => {
    const L = look({ firstRow: true, firstCol: true, hBand: true });
    expect(condTypesFor(at(0, 0), L, BANDS)).toEqual([
      'firstCol',
      'firstRow',
      'nwCell',
    ]);
    expect(condTypesFor(at(0, 1), L, BANDS)).toEqual(['firstRow']);
    // Banding counts the body: row 1 is body row 0 → band1 (shaded in the
    // PDF), row 2 → band2 (not shaded), row 4 → body 3 → band2.
    expect(condTypesFor(at(1, 1), L, BANDS)).toEqual(['band1Horz']);
    expect(condTypesFor(at(2, 1), L, BANDS)).toEqual(['band2Horz']);
    expect(condTypesFor(at(4, 1), L, BANDS)).toEqual(['band2Horz']);
    // Row banding runs INTO the first column (PDF: r2c1 is shaded).
    expect(condTypesFor(at(1, 0), L, BANDS)).toEqual(['band1Horz', 'firstCol']);
  });

  it('table B — everything on: column banding under firstRow/lastRow', () => {
    const L = look({
      firstRow: true,
      lastRow: true,
      firstCol: true,
      lastCol: true,
      hBand: true,
      vBand: true,
    });
    // PDF: r1c2 keeps the vertical band fill UNDER firstRow's bold.
    expect(condTypesFor(at(0, 1), L, BANDS)).toEqual(['band1Vert', 'firstRow']);
    expect(condTypesFor(at(4, 3), L, BANDS)).toEqual(['band1Vert', 'lastRow']);
    expect(condTypesFor(at(4, 4), L, BANDS)).toEqual([
      'lastCol',
      'lastRow',
      'seCell',
    ]);
    // Interior cell: both bands, even phases (PDF: r3c3 unshaded — band2
    // branches carry no fill in the probe style).
    expect(condTypesFor(at(2, 2), L, BANDS)).toEqual([
      'band2Horz',
      'band2Vert',
    ]);
  });

  it('table C — only lastRow + lastCol + vBand', () => {
    const L = look({ lastRow: true, lastCol: true, vBand: true });
    // No firstCol gate → banding starts at column 0 (PDF: c1 and c3 shaded).
    expect(condTypesFor(at(0, 0), L, BANDS)).toEqual(['band1Vert']);
    expect(condTypesFor(at(0, 2), L, BANDS)).toEqual(['band1Vert']);
    expect(condTypesFor(at(0, 1), L, BANDS)).toEqual(['band2Vert']);
    // lastCol is excluded from its own axis's banding (PDF: c5 never shaded).
    expect(condTypesFor(at(0, 4), L, BANDS)).toEqual(['lastCol']);
    // Vertical banding runs INTO the last row (PDF: r5c1 shaded AND bold).
    expect(condTypesFor(at(4, 0), L, BANDS)).toEqual(['band1Vert', 'lastRow']);
  });

  it('a w:tblHeader row takes firstRow formatting wherever it sits', () => {
    const L = look({ firstRow: true });
    expect(condTypesFor({ ...at(2, 1), header: true }, L, BANDS)).toEqual([
      'firstRow',
    ]);
  });

  it('band size 0 disables that axis (Word default, MS-OI29500 2.1.251)', () => {
    const L = look({ hBand: true, vBand: true });
    expect(condTypesFor(at(1, 1), L, { row: 0, col: 0 })).toEqual([]);
  });
});

describe('branchRegion', () => {
  it('bands are one region per stripe, bounded by the body', () => {
    const L = look({ firstRow: true, hBand: true });
    expect(branchRegion('band1Horz', at(1, 2), L, BANDS)).toEqual({
      rowStart: 1,
      rowEnd: 1,
      colStart: 0,
      colEnd: 4,
    });
    expect(branchRegion('firstRow', at(0, 2), L, BANDS)).toEqual({
      rowStart: 0,
      rowEnd: 0,
      colStart: 0,
      colEnd: 4,
    });
  });
});

describe('cellStyleLayer', () => {
  // A Light-Grid-shaped style: bold header with a heavy bottom edge and
  // banded fills — the same shape probe F2 verified against Word.
  const green = { width: 1, style: 'solid' as const, color: '#9BBB59' };
  const heavy = { ...green, width: 3 };
  const style: ResolvedTableStyle = {
    table: { borders: { top: green, bottom: green } },
    cond: {
      firstRow: {
        font: { bold: true },
        borders: { bottom: heavy, insideV: green, insideH: false },
      },
      band1Horz: { background: '#D6E3BC' },
    },
    bands: { row: 1, col: 1 },
  };
  const L = look({ firstRow: true, firstCol: true, hBand: true });

  it('header cell: bold + the branch bottom edge, no band fill', () => {
    const layer = cellStyleLayer(style, L, at(0, 1));
    expect(layer.font).toEqual({ bold: true });
    expect(layer.background).toBeUndefined();
    // Cell is on every edge of the one-row region except left/right interior:
    // bottom comes from the branch's bottom, left/right fall to insideV.
    expect(layer.borders).toEqual({
      bottom: heavy,
      left: green,
      right: green,
    });
  });

  it('body band cell: fill without bold', () => {
    const layer = cellStyleLayer(style, L, at(1, 1));
    expect(layer.font).toBeUndefined();
    expect(layer.background).toBe('#D6E3BC');
    expect(layer.borders ?? null).toBeNull();
  });

  it('gated branches contribute nothing', () => {
    const withLast: ResolvedTableStyle = {
      ...style,
      cond: { ...style.cond, lastRow: { font: { bold: true } } },
    };
    const layer = cellStyleLayer(withLast, L, at(4, 1));
    expect(layer.font).toBeUndefined();
  });

  it('paragraph spacing stacks per field: style, then the admitted branches', () => {
    // Word's stock Light Grid: the style itself single-spaces every cell
    // (after 0, line 1) and the lastRow branch only re-states the same —
    // but a branch that sets ONE field must keep the style's others.
    const spaced: ResolvedTableStyle = {
      ...style,
      paragraph: { spacing: { after: 0, line: 1, lineRule: 'auto' } },
      cond: {
        ...style.cond,
        lastRow: { paragraph: { spacing: { before: 8 } } },
      },
    };
    const body = cellStyleLayer(spaced, L, at(2, 1));
    expect(body.spacing).toEqual({ after: 0, line: 1, lineRule: 'auto' });
    // lastRow gated off by L: the branch's before does not reach row 4.
    expect(cellStyleLayer(spaced, L, at(4, 1)).spacing).toEqual(body.spacing);
    const last = cellStyleLayer(spaced, look({ lastRow: true }), at(4, 1));
    expect(last.spacing).toEqual({
      before: 8,
      after: 0,
      line: 1,
      lineRule: 'auto',
    });
    // No paragraph delta anywhere → no field at all, not an empty object.
    expect(cellStyleLayer(style, L, at(2, 1)).spacing).toBeUndefined();
  });
});
