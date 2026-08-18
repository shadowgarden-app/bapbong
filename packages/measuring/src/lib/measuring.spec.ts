import {
  createApproxMeasurer,
  createApproxMetrics,
  fontToCss,
} from './measuring.js';

describe('measuring', () => {
  it('builds CSS font shorthand', () => {
    // The symbol fallback rides every shorthand (see SYMBOL_FALLBACK_FAMILY).
    expect(
      fontToCss({ family: 'Arial', sizePt: 11, bold: true, italic: false }),
    ).toBe('700 11pt Arial, "Noto Sans Symbols 2"');
    expect(
      fontToCss({ family: 'Times', sizePt: 12, bold: false, italic: true }),
    ).toBe('italic 400 12pt Times, "Noto Sans Symbols 2"');
  });

  it('approx measurer scales with text length and font size', () => {
    const measure = createApproxMeasurer(0.5);
    expect(
      measure('abcd', { family: 'x', sizePt: 10, bold: false, italic: false }),
    ).toBe(20);
    expect(
      measure('abcd', { family: 'x', sizePt: 20, bold: false, italic: false }),
    ).toBe(40);
  });

  it('approx metrics split the em into ascent/descent', () => {
    const metrics = createApproxMetrics();
    // 12pt → 16px em; ascent 0.8·em = 12.8, descent 0.2·em = 3.2.
    const { ascent, descent } = metrics({
      family: 'x',
      sizePt: 12,
      bold: false,
      italic: false,
    });
    expect(ascent).toBeCloseTo(12.8);
    expect(descent).toBeCloseTo(3.2);
  });
});
