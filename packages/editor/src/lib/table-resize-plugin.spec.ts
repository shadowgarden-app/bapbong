import type {
  ResolvedCell,
  ResolvedLayout,
  ResolvedTable,
} from '@shadow-garden/bapbong-contracts';
import { borderAt } from './table-resize-plugin';

// Synthetic geometry: an outer 1×2 table (border at x=110); the left cell
// hosts a nested 1×2 table (border at x=60) whose right edge (x=100) is 10px
// from the outer border.
const cell = (
  x: number,
  y: number,
  w: number,
  h: number,
  tables?: ResolvedTable[],
): ResolvedCell =>
  ({
    x,
    y,
    width: w,
    height: h,
    colspan: 1,
    rowspan: 1,
    lines: [],
    ...(tables ? { tables } : {}),
  }) as ResolvedCell;

const nested = {
  x: 20,
  y: 20,
  width: 80,
  height: 40,
  cells: [cell(20, 20, 40, 40), cell(60, 20, 40, 40)],
} as ResolvedTable;
const outer = {
  x: 10,
  y: 10,
  width: 200,
  height: 100,
  cells: [cell(10, 10, 100, 100, [nested]), cell(110, 10, 100, 100)],
} as ResolvedTable;
const lay = {
  pages: [{ index: 0, width: 400, height: 200, lines: [], tables: [outer] }],
} as unknown as ResolvedLayout;
const at = (x: number, y: number) => ({ pageIndex: 0, x, y });

describe('borderAt', () => {
  it('finds the nested interior border (innermost wins)', () => {
    const hit = borderAt(lay, at(61, 30));
    expect(hit?.borderX).toBe(60);
    expect(hit?.leftCells).toEqual([nested.cells[0]]);
    expect(hit?.rightCells).toEqual([nested.cells[1]]);
    expect(hit?.tableY).toBe(20); // the NESTED table's band, not the outer's
  });

  it('still finds the outer border, including from inside the host cell', () => {
    expect(borderAt(lay, at(111, 80))?.borderX).toBe(110);
    // Point vertically inside the host cell but below the nested table.
    expect(borderAt(lay, at(108, 80))?.borderX).toBe(110);
  });

  it("a nested table's outer edge is not resizable and falls back", () => {
    // x=100 is the nested right edge — not resizable in the nested table —
    // and 10px from the outer border, beyond EDGE_TOL: no hit at all.
    expect(borderAt(lay, at(100, 30))).toBeNull();
    // The nested LEFT edge (x=20) likewise resolves to nothing.
    expect(borderAt(lay, at(20, 30))).toBeNull();
  });

  it('a miss on one table no longer hides its siblings', () => {
    // Two tables side by side on one page; the point rests on the FIRST
    // table's outer-right edge (not resizable) which is within EDGE_TOL of
    // the SECOND table's interior border.
    const t1 = {
      x: 10,
      y: 10,
      width: 100,
      height: 50,
      cells: [cell(10, 10, 100, 50)],
    } as ResolvedTable;
    const t2 = {
      x: 112,
      y: 10,
      width: 100,
      height: 50,
      cells: [cell(112, 10, 2, 50), cell(114, 10, 98, 50)],
    } as ResolvedTable;
    const two = {
      pages: [
        { index: 0, width: 400, height: 200, lines: [], tables: [t1, t2] },
      ],
    } as unknown as ResolvedLayout;
    // 110 = t1 right edge (unresizable) AND within tol of t2's border at 114.
    expect(borderAt(two, at(111, 30))?.borderX).toBe(114);
  });
});
