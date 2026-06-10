import type { FontSpec, MeasureMetrics, MeasureText } from '@shadow-garden/bapbong-contracts';

const PT_TO_PX = 96 / 72;

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

/**
 * Canvas-backed vertical metrics (browser). Reads the font bounding box so the
 * layout engine can build baseline-accurate line boxes.
 */
export function createCanvasMetrics(): MeasureMetrics {
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('bapbong-measuring: 2D canvas context unavailable');
  return (font) => {
    ctx.font = fontToCss(font);
    const m = ctx.measureText('Mg');
    const ascent = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent;
    const descent = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent;
    return { ascent, descent };
  };
}

/**
 * Headless approximate metrics (no DOM): ascent ≈ 0.8·em, descent ≈ 0.2·em.
 * For SSR/tests, or before web fonts load.
 */
export function createApproxMetrics(): MeasureMetrics {
  return (font) => {
    const em = font.sizePt * PT_TO_PX;
    return { ascent: em * 0.8, descent: em * 0.2 };
  };
}
