import type { EqNode, EqSlotRect } from '@shadow-garden/bapbong-contracts';
import {
  autoCorrectAt,
  eqChar,
  insertAt,
  horizontalStep,
  removeBefore,
  rowAt,
  verticalStep,
  withRow,
} from './equation-edit';

const chr = (ch: string): EqNode => ({ t: 'chr', ch });

describe('equation slot edits', () => {
  const ast: EqNode[] = [
    chr('𝑥'),
    { t: 'frac', num: [chr('𝑎')], den: [chr('𝑏'), chr('2')] },
  ];

  it('addresses rows by slot path and edits immutably', () => {
    expect(
      rowAt(ast, [1, 'den'])!.map((n) => (n as { ch: string }).ch),
    ).toEqual(['𝑏', '2']);
    const next = withRow(ast, [1, 'den'], (r) => r.slice(1));
    expect(rowAt(next, [1, 'den'])).toHaveLength(1);
    expect(rowAt(ast, [1, 'den'])).toHaveLength(2); // original untouched
    expect(rowAt(ast, [9, 'nope'])).toBeNull();
  });

  it('inserts and removes at caret stops', () => {
    const ins = insertAt(ast, [1, 'num'], 1, eqChar('y'));
    expect(
      rowAt(ins, [1, 'num'])!.map((n) => (n as { ch: string }).ch),
    ).toEqual(
      ['𝑎', '𝑦'], // typed letters take the math-italic letterform
    );
    const del = removeBefore(ast, [1, 'den'], 2);
    expect(
      rowAt(del, [1, 'den'])!.map((n) => (n as { ch: string }).ch),
    ).toEqual(['𝑏']);
    expect(removeBefore(ast, [1, 'den'], 0)).toBe(ast); // no-op at 0
  });

  it('completes \\name on the trigger space, consuming the backslash run', () => {
    const typed = ['\\', 'o', 'm', 'e', 'g', 'a'].map(chr);
    const withCmd = withRow(ast, [1, 'num'], (r) => [...r, ...typed]);
    const fixed = autoCorrectAt(withCmd, [1, 'num'], 7)!;
    expect(
      rowAt(fixed.ast, [1, 'num'])!.map((n) => (n as { ch: string }).ch),
    ).toEqual(['𝑎', 'ω']);
    expect(fixed.caret).toBe(2);
    // No match → null (the space inserts normally).
    expect(autoCorrectAt(ast, [1, 'num'], 1)).toBeNull();
  });
});

describe('walking the caret through a fraction', () => {
  // 𝑥 = 𝑎/𝑏𝑐, laid out the way the engine emits slots: the outer row first,
  // then each child row it drew, each carrying its path and caret stops.
  const ast: EqNode[] = [
    chr('𝑥'),
    chr('='),
    { t: 'frac', num: [chr('𝑎')], den: [chr('𝑏'), chr('𝑐')] },
  ];
  const slots: EqSlotRect[] = [
    {
      path: [],
      x: 0,
      y: -14,
      width: 40,
      height: 22,
      caretXs: [0, 8, 16, 40],
      em: 16,
    },
    {
      path: [2, 'num'],
      x: 20,
      y: -14,
      width: 12,
      height: 9,
      caretXs: [0, 12],
      em: 11,
    },
    {
      path: [2, 'den'],
      x: 20,
      y: -2,
      width: 18,
      height: 9,
      caretXs: [0, 9, 18],
      em: 11,
    },
  ];
  const OUTER = 0;
  const NUM = 1;
  const DEN = 2;

  it('steps right INTO the numerator rather than over the fraction', () => {
    // Caret in the outer row, just before the fraction (after "𝑥=").
    expect(horizontalStep(ast, slots, OUTER, 2, 1)).toEqual({
      slot: NUM,
      caret: 0,
    });
  });

  it('falls from the end of the numerator into the denominator', () => {
    expect(horizontalStep(ast, slots, NUM, 1, 1)).toEqual({
      slot: DEN,
      caret: 0,
    });
  });

  it('climbs out past the fraction at the end of the denominator', () => {
    expect(horizontalStep(ast, slots, DEN, 2, 1)).toEqual({
      slot: OUTER,
      caret: 3,
    });
  });

  it('mirrors going left: into the denominator, then the numerator', () => {
    expect(horizontalStep(ast, slots, OUTER, 3, -1)).toEqual({
      slot: DEN,
      caret: 2,
    });
    expect(horizontalStep(ast, slots, DEN, 0, -1)).toEqual({
      slot: NUM,
      caret: 1,
    });
  });

  it('leaves the equation at its edges', () => {
    expect(horizontalStep(ast, slots, OUTER, 3, 1)).toBe('out-right');
    expect(horizontalStep(ast, slots, OUTER, 0, -1)).toBe('out-left');
  });

  it('takes the short way down and up, keeping the caret x', () => {
    // End of the numerator sits at x = 32; the denominator stops are at
    // 20 / 29 / 38, so the caret lands on 29 — the same place across, not
    // the start of the row.
    expect(verticalStep(ast, slots, NUM, 1, 1)).toEqual({
      slot: DEN,
      caret: 1,
    });
    expect(verticalStep(ast, slots, DEN, 0, -1)).toEqual({
      slot: NUM,
      caret: 0,
    });
  });

  it('does not offer the row it is drawn inside as the row below', () => {
    // The outer row encloses both; from the denominator there is nothing
    // below, so the document's own line motion should take over.
    expect(verticalStep(ast, slots, DEN, 0, 1)).toBeNull();
  });
});

describe('walking a radical and a script', () => {
  // ⁿ√𝑥 followed by 𝑎 with a superscript: the two shapes whose rows are not
  // side by side, so the geometry has to carry the up/down moves.
  const ast: EqNode[] = [
    { t: 'rad', deg: [chr('𝑛')], body: [chr('𝑥')], showDeg: true },
    { t: 'scr', base: [chr('𝑎')], sub: [], sup: [chr('2')], slots: 'sup' },
  ];
  const slots: EqSlotRect[] = [
    {
      path: [],
      x: 0,
      y: -16,
      width: 60,
      height: 24,
      caretXs: [0, 30, 60],
      em: 16,
    },
    {
      path: [0, 'deg'],
      x: 0,
      y: -16,
      width: 6,
      height: 7,
      caretXs: [0, 6],
      em: 9,
    },
    {
      path: [0, 'body'],
      x: 10,
      y: -10,
      width: 12,
      height: 12,
      caretXs: [0, 12],
      em: 16,
    },
    {
      path: [1, 'base'],
      x: 34,
      y: -10,
      width: 10,
      height: 12,
      caretXs: [0, 10],
      em: 16,
    },
    {
      path: [1, 'sup'],
      x: 44,
      y: -16,
      width: 6,
      height: 7,
      caretXs: [0, 6],
      em: 11,
    },
  ];
  const [OUTER, DEG, BODY, BASE, SUP] = [0, 1, 2, 3, 4];

  it('reads a radical degree-first, then its body', () => {
    expect(horizontalStep(ast, slots, OUTER, 0, 1)).toEqual({
      slot: DEG,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, DEG, 1, 1)).toEqual({
      slot: BODY,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, BODY, 1, 1)).toEqual({
      slot: OUTER,
      caret: 1,
    });
  });

  it('reads a script base-first, then the script it carries', () => {
    expect(horizontalStep(ast, slots, OUTER, 1, 1)).toEqual({
      slot: BASE,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, BASE, 1, 1)).toEqual({
      slot: SUP,
      caret: 0,
    });
  });

  it('offers no subscript to step into when the script has none', () => {
    // A superscript-only script must not strand the caret in a row the
    // layout never drew.
    expect(horizontalStep(ast, slots, SUP, 1, 1)).toEqual({
      slot: OUTER,
      caret: 2,
    });
  });

  it('takes the degree and the exponent with up, the row under with down', () => {
    expect(verticalStep(ast, slots, BODY, 0, -1)).toEqual({
      slot: DEG,
      caret: 1,
    });
    expect(verticalStep(ast, slots, DEG, 0, 1)).toEqual({
      slot: BODY,
      caret: 0,
    });
    expect(verticalStep(ast, slots, BASE, 1, -1)).toEqual({
      slot: SUP,
      caret: 0,
    });
    expect(verticalStep(ast, slots, SUP, 0, 1)).toEqual({
      slot: BASE,
      caret: 1,
    });
  });
});

describe('walking a bracket and a big operator', () => {
  // (𝑎) followed by a ∑ carrying both limits: one row and three rows.
  const ast: EqNode[] = [
    { t: 'fence', l: '(', r: ')', body: [chr('𝑎')] },
    {
      t: 'big',
      op: '∑',
      lo: [chr('0')],
      hi: [chr('𝑛')],
      body: [chr('𝑘')],
      showLo: true,
      showHi: true,
    },
  ];
  const slots: EqSlotRect[] = [
    {
      path: [],
      x: 0,
      y: -20,
      width: 80,
      height: 34,
      caretXs: [0, 24, 80],
      em: 16,
    },
    {
      path: [0, 'body'],
      x: 6,
      y: -10,
      width: 12,
      height: 12,
      caretXs: [0, 12],
      em: 16,
    },
    {
      path: [1, 'lo'],
      x: 30,
      y: 4,
      width: 8,
      height: 8,
      caretXs: [0, 8],
      em: 11,
    },
    {
      path: [1, 'hi'],
      x: 30,
      y: -20,
      width: 8,
      height: 8,
      caretXs: [0, 8],
      em: 11,
    },
    {
      path: [1, 'body'],
      x: 46,
      y: -10,
      width: 12,
      height: 12,
      caretXs: [0, 12],
      em: 16,
    },
  ];
  const [OUTER, FENCE, LO, HI, BODY] = [0, 1, 2, 3, 4];

  it('steps inside a bracket and back out past it', () => {
    expect(horizontalStep(ast, slots, OUTER, 0, 1)).toEqual({
      slot: FENCE,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, FENCE, 1, 1)).toEqual({
      slot: OUTER,
      caret: 1,
    });
  });

  it('reads a big operator lower limit, upper limit, then operand', () => {
    expect(horizontalStep(ast, slots, OUTER, 1, 1)).toEqual({
      slot: LO,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, LO, 1, 1)).toEqual({
      slot: HI,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, HI, 1, 1)).toEqual({
      slot: BODY,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, BODY, 1, 1)).toEqual({
      slot: OUTER,
      caret: 2,
    });
  });

  it('reaches the limits stacked on the operator from the operand', () => {
    expect(verticalStep(ast, slots, BODY, 0, -1)).toEqual({
      slot: HI,
      caret: 1,
    });
    expect(verticalStep(ast, slots, BODY, 0, 1)).toEqual({
      slot: LO,
      caret: 1,
    });
  });

  it('steps between the two limits, not through the operand', () => {
    // The operand is the nearest row below the upper limit, but the limits
    // are one stack on the operator sign: down means the other limit.
    expect(verticalStep(ast, slots, HI, 0, 1)).toEqual({ slot: LO, caret: 0 });
    expect(verticalStep(ast, slots, LO, 0, -1)).toEqual({ slot: HI, caret: 0 });
  });
});

describe('walking a matrix', () => {
  // A 2×2 grid. Cells live in a flat store, so the path names them c0..c3.
  const ast: EqNode[] = [
    {
      t: 'mat',
      cols: 2,
      cells: [[chr('𝑎')], [chr('𝑏')], [chr('𝑐')], [chr('𝑑')]],
    },
  ];
  const cell = (i: number, x: number, y: number): EqSlotRect => ({
    path: [0, `c${i}`],
    x,
    y,
    width: 10,
    height: 10,
    caretXs: [0, 10],
    em: 16,
  });
  const slots: EqSlotRect[] = [
    { path: [], x: 0, y: -14, width: 30, height: 26, caretXs: [0, 30], em: 16 },
    cell(0, 0, -12),
    cell(1, 16, -12),
    cell(2, 0, 2),
    cell(3, 16, 2),
  ];
  const [OUTER, C0, C1, C2, C3] = [0, 1, 2, 3, 4];

  it('addresses a cell by path and edits it immutably', () => {
    expect(rowAt(ast, [0, 'c3'])).toEqual([chr('𝑑')]);
    const next = withRow(ast, [0, 'c3'], (r) => [...r, chr('!')]);
    expect(rowAt(next, [0, 'c3'])).toHaveLength(2);
    expect(rowAt(ast, [0, 'c3'])).toHaveLength(1); // original untouched
    expect(rowAt(next, [0, 'c0'])).toEqual([chr('𝑎')]); // siblings untouched
  });

  it('reads the cells in reading order, then leaves the grid', () => {
    expect(horizontalStep(ast, slots, OUTER, 0, 1)).toEqual({
      slot: C0,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, C0, 1, 1)).toEqual({
      slot: C1,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, C1, 1, 1)).toEqual({
      slot: C2,
      caret: 0,
    });
    expect(horizontalStep(ast, slots, C3, 1, 1)).toEqual({
      slot: OUTER,
      caret: 1,
    });
  });

  it('moves down a column, not to the next cell across', () => {
    // A grid declares no vertical stack: geometry is the right answer here,
    // and it keeps the column.
    expect(verticalStep(ast, slots, C0, 0, 1)).toEqual({ slot: C2, caret: 0 });
    expect(verticalStep(ast, slots, C1, 0, 1)).toEqual({ slot: C3, caret: 0 });
    expect(verticalStep(ast, slots, C3, 0, -1)).toEqual({ slot: C1, caret: 0 });
  });
});
