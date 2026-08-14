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
    lineGap?: number;
    /** OS/2 fields the Word line-height rule reads. */
    os2?: {
      fsSelection?: number;
      usWinAscent?: number;
      usWinDescent?: number;
      sTypoAscender?: number;
      sTypoDescender?: number;
      sTypoLineGap?: number;
    };
  } = {},
) {
  const {
    from = 65,
    to = 90,
    advance = 500,
    em = 1000,
    ascender = 800,
    descender = -200,
    lineGap = 0,
    os2,
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
  const font = new opentype.Font({
    familyName: 'Test',
    styleName: 'Regular',
    unitsPerEm: em,
    ascender,
    descender,
    glyphs,
  });
  // opentype.Font's constructor takes no hhea lineGap or OS/2 values; the
  // metrics reader goes to the tables, so set them there.
  font.tables.hhea = { ...font.tables.hhea, lineGap };
  if (os2) font.tables.os2 = { ...font.tables.os2, ...os2 };
  return font;
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

  it("lends a family's vertical metrics to its weight siblings, not its widths", () => {
    // Word names "Calibri Light" as a family of its own, so a document written
    // in it asks for something no registry has. Its vertical metrics belong to
    // the family — "every font in a family must share the same vertical metric
    // values" — but its advance widths are the very thing the weight changes,
    // so those keep falling back.
    const reg = new FontRegistry();
    const base = makeFont({ advance: 500, ascender: 800, descender: -200 });
    reg.register('Test', {}, base);

    expect(reg.primary(spec({ family: 'Test Light' }))).toBe(base);
    expect(reg.primary(spec({ family: 'Test Semibold' }))).toBe(base);
    expect(reg.primary(spec({ family: 'Test Black Italic' }))).toBe(base);
    // …while the width path still says "I do not have this face".
    expect(reg.has(spec({ family: 'Test Light' }))).toBe(false);

    // A WIDTH word is not a weight word: Narrow and Condensed faces really are
    // drawn to different advances, and borrowing anything from the base family
    // would mis-break every line.
    expect(reg.primary(spec({ family: 'Test Narrow' }))).toBeNull();
    expect(reg.primary(spec({ family: 'Test Condensed' }))).toBeNull();
    // An unrelated family is still unknown, weight word or not.
    expect(reg.primary(spec({ family: 'Nope Light' }))).toBeNull();
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

  it("reads Word's line-height rule: win metrics plus external leading", () => {
    // GDI: the cell is usWinAscent+usWinDescent, and tmExternalLeading is
    // MAX(0, (hheaAsc − hheaDesc + lineGap) − (winAsc + winDesc)). Here the
    // hhea total (800+200+120 = 1120) exceeds the win total (850+250 = 1100)
    // by 20 units, so 20/1000 of the em is leading.
    const reg = new FontRegistry();
    reg.register(
      'Test',
      {},
      makeFont({
        em: 1000,
        ascender: 800,
        descender: -200,
        lineGap: 120,
        os2: { usWinAscent: 850, usWinDescent: 250 },
      }),
    );
    const metrics = createFontRegistryMetrics(reg, createApproxMetrics());
    const m = metrics(spec()); // 12pt → 16px
    expect(m.ascent).toBeCloseTo(13.6); // 850/1000·16
    expect(m.descent).toBeCloseTo(4.0); // 250/1000·16
    expect(m.leading).toBeCloseTo(0.32); // 20/1000·16
  });

  it('clamps the leading at zero when the win box already covers hhea', () => {
    const reg = new FontRegistry();
    reg.register(
      'Test',
      {},
      makeFont({
        em: 1000,
        ascender: 800,
        descender: -200,
        lineGap: 0,
        os2: { usWinAscent: 900, usWinDescent: 300 },
      }),
    );
    const metrics = createFontRegistryMetrics(reg, createApproxMetrics());
    expect(metrics(spec()).leading).toBe(0);
  });

  it('switches to the typographic metrics when USE_TYPO_METRICS is set', () => {
    // fsSelection bit 7. Not a corner case: Arimo (the metric-compatible stand-
    // in for Arial) sets it, and its usWin values are far larger than its
    // typographic ones — reading them would make every line 25% too tall.
    const reg = new FontRegistry();
    reg.register(
      'Test',
      {},
      makeFont({
        em: 1000,
        ascender: 800,
        descender: -200,
        lineGap: 120,
        os2: {
          fsSelection: 0x80,
          usWinAscent: 1100,
          usWinDescent: 400,
          sTypoAscender: 750,
          sTypoDescender: -250,
          sTypoLineGap: 100,
        },
      }),
    );
    const metrics = createFontRegistryMetrics(reg, createApproxMetrics());
    const m = metrics(spec()); // 12pt → 16px
    expect(m.ascent).toBeCloseTo(12.0); // 750/1000·16, not the 1100 win value
    expect(m.descent).toBeCloseTo(4.0); // 250/1000·16
    expect(m.leading).toBeCloseTo(1.6); // 100/1000·16
  });

  it('falls back for unknown families', () => {
    const reg = new FontRegistry();
    const metrics = createFontRegistryMetrics(reg, createApproxMetrics());
    const { ascent, descent } = metrics(spec({ family: 'Missing' }));
    expect(ascent).toBeCloseTo(12.8); // approx 0.8·em
    expect(descent).toBeCloseTo(3.2);
  });
});
