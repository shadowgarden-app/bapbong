import type { EqNode } from '@shadow-garden/bapbong-contracts';
import {
  autoCorrectAt,
  eqChar,
  insertAt,
  removeBefore,
  rowAt,
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
