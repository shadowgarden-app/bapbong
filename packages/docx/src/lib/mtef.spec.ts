import { mtefFromWmf, mtefToLinear } from './mtef.js';

// Builders for a minimal MTEF v5 stream (records verified against the spec
// and 183 real MathType equations — see mtef.ts).
const HEADER = [
  5,
  1,
  0,
  6,
  9,
  ...'DSMT6'.split('').map((c) => c.charCodeAt(0)),
  0,
  1,
];
const chr = (tf: number, ch: string, opts = 0) => [
  2,
  opts,
  tf + 128,
  ch.charCodeAt(0) & 0xff,
  ch.charCodeAt(0) >> 8,
];
const line = (...objs: number[][]) => [1, 0, ...objs.flat(), 0];
const nullLine = () => [1, 1];
const stream = (...objs: number[][]) =>
  new Uint8Array([...HEADER, ...line(...objs).slice(0, -1), 0, 0]);

describe('mtefToLinear', () => {
  it('renders chars with math-italic variables and upright functions', () => {
    // y = f(x): variables via typeface 3, operator via 6.
    const m = stream(
      chr(3, 'y'),
      chr(6, '='),
      chr(3, 'f'),
      chr(6, '('),
      chr(3, 'x'),
      chr(6, ')'),
    );
    expect(mtefToLinear(m)).toBe('𝑦=𝑓(𝑥)');
  });

  it('renders a fraction template with parenthesized sides', () => {
    // TMPL 11 (fraction), two LINE slots: (x+1) / 2.
    const m = stream([
      3,
      0,
      11,
      0,
      0,
      ...line(chr(3, 'x'), chr(6, '+'), chr(8, '1')),
      ...line(chr(8, '2')),
      0,
    ]);
    expect(mtefToLinear(m)).toBe('(𝑥+1)/2');
  });

  it('renders scripts from the fixed [sub, sup] slots', () => {
    // x then TMPL 28 (sup): NULL sub line + "4".
    const m = stream(chr(3, 'x'), [
      3,
      0,
      28,
      0,
      0,
      ...nullLine(),
      ...line(chr(8, '4')),
      0,
    ]);
    expect(mtefToLinear(m)).toBe('𝑥⁴');
  });

  it('renders an embellished char (vector arrow)', () => {
    // CHAR with embellishment list: B + EMBELL type 11 (right arrow over).
    const m = stream([...chr(3, 'B', 0x01), 6, 0, 11, 0]);
    expect(mtefToLinear(m)).toBe('𝐵⃗');
  });

  it('rejects other versions and unknown records', () => {
    expect(mtefToLinear(new Uint8Array([4, 1, 0, 3, 0, 0]))).toBeNull();
    expect(mtefToLinear(stream([99, 0]))).toBeNull();
  });
});

describe('mtefFromWmf', () => {
  it('lifts the payload after the Design Science comment marker', () => {
    const payload = stream(chr(3, 'x'));
    const wmf = new Uint8Array([
      1,
      2,
      3,
      ...'AppsMFCC'.split('').map((c) => c.charCodeAt(0)),
      9,
      9,
      ...'Design Science, Inc.\0'.split('').map((c) => c.charCodeAt(0)),
      ...payload,
    ]);
    expect(mtefToLinear(mtefFromWmf(wmf)!)).toBe('𝑥');
    expect(mtefFromWmf(new Uint8Array([1, 2, 3]))).toBeNull();
  });
});
