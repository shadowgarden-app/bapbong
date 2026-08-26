import type { EqNode, VectorOp } from '@shadow-garden/bapbong-contracts';
import { layoutEquation } from './equation-layout.js';

// Deterministic measurer: every char is half an em wide.
const measure = (text: string, font: { sizePt: number }) =>
  [...text].length * ((font.sizePt * 96) / 72) * 0.5;

const chr = (ch: string): EqNode => ({ t: 'chr', ch });

describe('layoutEquation', () => {
  it('lays a plain row on one baseline with cumulative caret stops', () => {
    const r = layoutEquation([chr('x'), chr('='), chr('2')], 12, measure);
    const em = 16;
    expect(r.width).toBeCloseTo(3 * em * 0.5 + 2, 1);
    // One text op (chars coalesce), baseline at the ascent line.
    const texts = r.ops.filter((o) => o.kind === 'text');
    expect(texts).toHaveLength(1);
    expect((texts[0] as { y: number }).y).toBeCloseTo(r.ascent, 5);
    // Root slot: caret stops 0..3.
    expect(r.slots[0].path).toEqual([]);
    expect(r.slots[0].caretXs).toHaveLength(4);
    expect(r.slots[0].caretXs[0]).toBe(0);
  });

  it('stacks a fraction around the math axis and records both slots', () => {
    const frac: EqNode = { t: 'frac', num: [chr('a')], den: [chr('b')] };
    const r = layoutEquation([frac], 12, measure);
    const paths = r.slots.map((s) => s.path.join('.'));
    expect(paths).toContain('0.num');
    expect(paths).toContain('0.den');
    // Numerator renders above the denominator.
    const num = r.slots.find((s) => s.path.join('.') === '0.num')!;
    const den = r.slots.find((s) => s.path.join('.') === '0.den')!;
    expect(num.y + num.height).toBeLessThanOrEqual(den.y + 0.01);
    // The bar is a line op between them.
    const bar = r.ops.find((o) => o.kind === 'line') as VectorOp & {
      y1: number;
    };
    expect(bar).toBeTruthy();
    expect(bar.y1).toBeGreaterThan(num.y);
    expect(bar.y1).toBeLessThan(den.y + den.height);
    // Taller than a plain row: real 2D stacking.
    const plain = layoutEquation([chr('a')], 12, measure);
    expect(r.height).toBeGreaterThan(plain.height * 1.3);
  });

  it('raises superscripts and keeps subscripts low', () => {
    const scr: EqNode = {
      t: 'scr',
      base: [chr('x')],
      sub: [],
      sup: [chr('2')],
    };
    const r = layoutEquation([scr], 12, measure);
    const sup = r.slots.find((s) => s.path.join('.') === '0.sup')!;
    const base = r.slots.find((s) => s.path.join('.') === '0.base')!;
    expect(sup.y).toBeLessThan(base.y); // higher on the canvas
    expect(sup.em).toBeCloseTo(16 * 0.66, 1);
  });

  it('renders an empty slot as a placeholder box with one caret stop', () => {
    const frac: EqNode = { t: 'frac', num: [], den: [chr('b')] };
    const r = layoutEquation([frac], 12, measure);
    const num = r.slots.find((s) => s.path.join('.') === '0.num')!;
    expect(num.caretXs).toHaveLength(1);
    expect(num.width).toBeGreaterThan(0);
  });

  it('keeps every op inside the reported box', () => {
    const ast: EqNode[] = [
      chr('f'),
      { t: 'fence', l: '(', r: ')', body: [chr('x')] },
      chr('='),
      {
        t: 'frac',
        num: [chr('m')],
        den: [
          { t: 'scr', base: [chr('x')], sub: [], sup: [chr('2')] },
          chr('+'),
          { t: 'rad', deg: [], body: [chr('y')] },
        ],
      },
    ];
    const r = layoutEquation(ast, 12, measure);
    for (const op of r.ops) {
      if (op.kind === 'text') {
        expect(op.x).toBeGreaterThanOrEqual(-0.01);
        expect(op.y).toBeGreaterThanOrEqual(-0.01);
        expect(op.y).toBeLessThanOrEqual(r.height + 0.01);
      } else if (op.kind === 'line') {
        for (const v of [op.y1, op.y2]) {
          expect(v).toBeGreaterThanOrEqual(-0.01);
          expect(v).toBeLessThanOrEqual(r.height + 0.01);
        }
      }
    }
    // Baseline seat: raise = -(height - ascent) puts the internal baseline
    // exactly on the line's.
    expect(r.ascent).toBeGreaterThan(0);
    expect(r.ascent).toBeLessThan(r.height);
  });
});
