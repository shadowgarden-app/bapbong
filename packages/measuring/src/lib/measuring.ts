import {
  applyGlyphSpec,
  fontShorthand,
  glyphCount,
  type FontSpec,
  type MeasureMetrics,
  type MeasureText,
} from '@shadow-garden/bapbong-contracts';

const PT_TO_PX = 96 / 72;

/** CSS font shorthand for a FontSpec, e.g. "italic 700 11pt Arial".
 *
 *  @deprecated Re-exported for compatibility. It only says what the shorthand
 *  can say, so a context configured with it draws tracked text untracked while
 *  the measurer still charges for the tracking. Use `applyGlyphSpec`. */
export function fontToCss(font: FontSpec): string {
  return fontShorthand(font);
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
    // Configure and measure through the same helper the painter uses, so the
    // browser does the tracking arithmetic once for both of us.
    applyGlyphSpec(ctx, font);
    return ctx.measureText(text).width;
  };
}

/**
 * Headless approximate measurer (no DOM): width ≈ chars × sizePt × factor.
 * For SSR/tests, or as a first-paint estimate before fonts load.
 */
export function createApproxMeasurer(factor = 0.5): MeasureText {
  return (text, font) =>
    text.length * font.sizePt * factor +
    (font.letterSpacing ?? 0) * glyphCount(text);
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
    // Plain shorthand on purpose: vertical metrics take a FontFace, so there
    // is no glyph adjustment here to apply — and none could move a baseline.
    ctx.font = fontShorthand({ ...font });
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

/**
 * Wait for the given font families to be usable before measuring — otherwise
 * `measureText` runs against fallback fonts and every wrap/pagination
 * coordinate is wrong once the real font arrives. Loads the four common
 * variants per family. Resolves immediately where the CSS Font Loading API is
 * unavailable (headless/tests) or for fonts the browser can't load.
 */
export async function ensureFontsLoaded(
  families: string[],
  sizePt = 12,
): Promise<void> {
  const fonts = (globalThis as { document?: { fonts?: FontFaceSet } }).document
    ?.fonts;
  if (!fonts?.load) return;
  const variants = ['', 'italic ', '700 ', 'italic 700 '];
  const loads: Promise<unknown>[] = [];
  for (const family of new Set(families)) {
    for (const variant of variants) {
      // load() rejects for unparsable specs; unknown fonts just resolve empty.
      loads.push(
        fonts.load(`${variant}${sizePt}pt "${family}"`).catch(() => []),
      );
    }
  }
  await Promise.all(loads);
}
