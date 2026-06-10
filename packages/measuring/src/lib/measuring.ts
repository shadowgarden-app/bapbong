import type { FontSpec, MeasureText } from '@shadow-garden/bapbong-contracts';

/** CSS font shorthand for a FontSpec, e.g. "italic 700 11pt Arial". */
export function fontToCss(font: FontSpec): string {
  const style = font.italic ? 'italic ' : '';
  const weight = font.bold ? '700 ' : '400 ';
  return `${style}${weight}${font.sizePt}pt ${font.family}`;
}

/**
 * Canvas-backed text measurer (browser). Reuses a single offscreen 2D context,
 * which is the accurate way to size text before painting it on canvas.
 */
export function createCanvasMeasurer(): MeasureText {
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('bapbong-measuring: 2D canvas context unavailable');
  return (text, font) => {
    ctx.font = fontToCss(font);
    return ctx.measureText(text).width;
  };
}

/**
 * Headless approximate measurer (no DOM): width ≈ chars × sizePt × factor.
 * For SSR/tests, or as a first-paint estimate before fonts load.
 */
export function createApproxMeasurer(factor = 0.5): MeasureText {
  return (text, font) => text.length * font.sizePt * factor;
}
