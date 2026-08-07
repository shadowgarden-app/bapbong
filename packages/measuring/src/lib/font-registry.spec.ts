import * as opentype from 'opentype.js';
import { createApproxMeasurer, createApproxMetrics } from './measuring.js';
import {
  FontRegistry,
  createFontRegistryMeasurer,
  createFontRegistryMetrics,
} from './font-registry.js';
import type { FontSpec } from '@shadow-garden/bapbong-contracts';

/** A synthetic font covering unicode [from, to] with fixed advance widths, so
 *  every assertion is exact and platform-independent (the point of measuring
 *  from font files). Ranges let us fake subset files (latin vs vietnamese). */
function makeFont(
  opts: {
    from?: number;
    to?: number;
    advance?: number;
    em?: number;
    ascender?: number;
    descender?: number;
  } = {},
) {
  const {
    from = 65,
    to = 90,
    advance = 500,
    em = 1000,
    ascender = 800,
    descender = -200,
  } = opts;
  const glyphs = [
    new opentype.Glyph({
      name: '.notdef',
      unicode: 0,
      advanceWidth: advance,
      path: new opentype.Path(),
    }),
  ];
  for (let c = from; c <= to; c++) {
    glyphs.push(
      new opentype.Glyph({
        name: `u${c}`,
        unicode: c,
        advanceWidth: advance,
        path: new opentype.Path(),
      }),
    );
  }
  return new opentype.Font({
    familyName: 'Test',
    styleName: 'Regular',
    unitsPerEm: em,
    ascender,
    descender,
    glyphs,
  });
}

const spec = (over: Partial<FontSpec> = {}): FontSpec => ({
  family: 'Test',
  sizePt: 12,
  bold: false,
  italic: false,
  ...over,
});

describe('FontRegistry', () => {
  it('resolves exact variant, falls back to regular, else null', () => {
    const reg = new FontRegistry();
    const regular = makeFont();
    const boldFace = makeFont();
    reg.register('Test', {}, regular);
    reg.register('Test', { bold: true }, boldFace);

    expect(reg.primary(spec({ bold: true }))).toBe(boldFace); // exact
    expect(reg.primary(spec({ italic: true }))).toBe(regular); // variant → regular
    expect(reg.primary(spec({ family: 'Nope' }))).toBeNull();
    expect(reg.has(spec())).toBe(true);
    expect(reg.has(spec({ family: 'Nope' }))).toBe(false);
  });

  it('is case-insensitive on family name', () => {
    const reg = new FontRegistry();
    reg.register('Test', {}, makeFont());
    expect(reg.has(spec({ family: 'tEsT' }))).toBe(true);
  });

  it('registerBytes round-trips a serialized font', () => {
    const reg = new FontRegistry();
    reg.registerBytes('Test', {}, makeFont().toArrayBuffer());
    expect(reg.has(spec())).toBe(true);
  });
});

describe('font-registry measurer', () => {
  it('sums advance widths for a registered face (deterministic)', () => {
    const reg = new FontRegistry();
    // Parse via bytes: opentype's cmap lookup only works on a parsed font, which
    // is also the production path (registerBytes).
    reg.registerBytes(
      'Test',
      {},
      makeFont({ advance: 500, em: 1000 }).toArrayBuffer(),
    );
    const measure = createFontRegistryMeasurer(reg, createApproxMeasurer(0.5));
    // 12pt → 16px; each glyph 500/1000·16 = 8px; "AB" = 16px.
    expect(measure('AB', spec())).toBeCloseTo(16);
    // 24pt → 32px; "ABC" = 3·16 = 48px.
    expect(measure('ABC', spec({ sizePt: 24 }))).toBeCloseTo(48);
  });

  it('routes each char to the subset file that has its glyph', () => {
    const reg = new FontRegistry();
    reg.registerBytes(
      'Test',
      {},
      makeFont({ from: 65, to: 77, advance: 500 }).toArrayBuffer(),
    ); // A–M @ 500
    reg.registerBytes(
      'Test',
      {},
      makeFont({ from: 78, to: 90, advance: 300 }).toArrayBuffer(),
    ); // N–Z @ 300
    const measure = createFontRegistryMeasurer(reg, createApproxMeasurer(0.5));
    // "A" (500/1000·16=8) + "N" (300/1000·16=4.8) = 12.8.
    expect(measure('AN', spec())).toBeCloseTo(12.8);
    // char in no file → first file's .notdef advance (500 → 8).
    expect(measure('z', spec())).toBeCloseTo(8);
  });

  it('falls back for families absent from the registry', () => {
    const reg = new FontRegistry();
    const measure = createFontRegistryMeasurer(reg, createApproxMeasurer(0.5));
    // approx: length·sizePt·0.5 = 4·10·0.5 = 20.
    expect(measure('abcd', spec({ family: 'Missing', sizePt: 10 }))).toBe(20);
  });
});

describe('font-registry metrics', () => {
  it('reads ascent/descent from a registered face (deterministic)', () => {
    const reg = new FontRegistry();
    reg.register(
      'Test',
      {},
      makeFont({ em: 1000, ascender: 800, descender: -200 }),
    );
    const metrics = createFontRegistryMetrics(reg, createApproxMetrics());
    // 12pt → 16px; ascent 800/1000·16 = 12.8, descent 200/1000·16 = 3.2.
    const { ascent, descent } = metrics(spec());
    expect(ascent).toBeCloseTo(12.8);
    expect(descent).toBeCloseTo(3.2);
  });

  it('falls back for unknown families', () => {
    const reg = new FontRegistry();
    const metrics = createFontRegistryMetrics(reg, createApproxMetrics());
    const { ascent, descent } = metrics(spec({ family: 'Missing' }));
    expect(ascent).toBeCloseTo(12.8); // approx 0.8·em
    expect(descent).toBeCloseTo(3.2);
  });
});
