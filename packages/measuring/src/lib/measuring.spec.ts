import { createApproxMeasurer, fontToCss } from './measuring.js';

describe('measuring', () => {
  it('builds CSS font shorthand', () => {
    expect(fontToCss({ family: 'Arial', sizePt: 11, bold: true, italic: false })).toBe('700 11pt Arial');
    expect(fontToCss({ family: 'Times', sizePt: 12, bold: false, italic: true })).toBe(
      'italic 400 12pt Times',
    );
  });

  it('approx measurer scales with text length and font size', () => {
    const measure = createApproxMeasurer(0.5);
    expect(measure('abcd', { family: 'x', sizePt: 10, bold: false, italic: false })).toBe(20);
    expect(measure('abcd', { family: 'x', sizePt: 20, bold: false, italic: false })).toBe(40);
  });
});
