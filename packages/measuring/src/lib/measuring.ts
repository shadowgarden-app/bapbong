import {
  applyGlyphSpec,
  fontShorthand,
  glyphCount,
  type FontAvailable,
  type FontSpec,
  type MeasureMetrics,
  type MeasureText,
} from '@shadow-garden/bapbong-contracts';
import type { FontRegistry } from './font-registry.js';

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
    // browser does the tracking arithmetic once for both of us. measureText
    // ignores the transform, so the horizontal scale is applied here — the
    // painter puts the very same number on its transform instead.
    const scale = applyGlyphSpec(ctx, font);
    return ctx.measureText(text).width * scale;
  };
}

/**
 * Headless approximate measurer (no DOM): width ≈ chars × sizePt × factor.
 * For SSR/tests, or as a first-paint estimate before fonts load.
 */
export function createApproxMeasurer(factor = 0.5): MeasureText {
  return (text, font) =>
    text.length * font.sizePt * factor * (font.scaleX ?? 1) +
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
  const leadingOf = canvasLeading(ctx);
  return (font) => {
    // Plain shorthand on purpose: vertical metrics take a FontFace, so there
    // is no glyph adjustment here to apply — and none could move a baseline.
    ctx.font = fontShorthand({ ...font });
    const m = ctx.measureText('Mg');
    const ascent = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent;
    const descent = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent;
    // Two reasons this path is the fallback and not the goal: the values come
    // back rounded to whole pixels (Times New Roman at 13pt measures 15+4
    // where the font says 15.15+3.68), and they describe the face the ENGINE
    // picked, which need not be the one Word used. createFontRegistryMetrics
    // reads the tables instead.
    return { ascent, descent, leading: leadingOf(font, ascent + descent) };
  };
}

/** Reference size for the leading probe: big enough that the browser's
 *  sub-pixel rounding is noise once the result is scaled back down. */
const LEADING_PROBE_PX = 400;

/**
 * The external leading of a font, asked of the browser rather than guessed.
 *
 * A 2D context exposes the font's box but not its line gap, so a line measured
 * from canvas alone is short by whatever gap the font declares — 4.2% of the em
 * for Times New Roman, 22% for Calibri, 0 for Verdana and Georgia. Over a long
 * document that is pages: a teaching plan whose font we could not read
 * paginated 31 pages short of Word, and the missing gap was most of it.
 *
 * But the browser DOES know the gap: `line-height: normal` is defined to be the
 * font's own line height, cell plus gap. So one hidden element per face answers
 * it exactly, and the answer is the same number Word calls single spacing.
 *
 * Measured once per family+weight+slant at a large size and scaled, because the
 * gap is a fraction of the em; without a DOM (SSR) it reports none rather than
 * inventing one.
 */
function canvasLeading(
  ctx: CanvasRenderingContext2D,
): (font: FontSpec, cell: number) => number {
  const doc = (globalThis as { document?: Document }).document;
  if (!doc?.body) return () => 0;
  const cache = new Map<string, number>();
  return (font, cell) => {
    const key = `${font.family}|${font.bold ? 'b' : ''}|${font.italic ? 'i' : ''}`;
    let ratio = cache.get(key);
    if (ratio === undefined) {
      const probe = doc.createElement('div');
      probe.style.cssText =
        'position:absolute;visibility:hidden;white-space:nowrap;' +
        'line-height:normal;padding:0;border:0;margin:0';
      probe.style.font = fontShorthand({
        ...font,
        sizePt: LEADING_PROBE_PX * (72 / 96),
      });
      probe.textContent = 'Mg';
      doc.body.appendChild(probe);
      const natural = probe.getBoundingClientRect().height;
      doc.body.removeChild(probe);
      // The same face measured through canvas at the probe size — subtracting
      // the box the caller will use is what leaves the GAP, and doing it at
      // one size keeps the two measurements commensurable.
      ctx.font = fontShorthand({
        ...font,
        sizePt: LEADING_PROBE_PX * (72 / 96),
      });
      const pm = ctx.measureText('Mg');
      const probeCell =
        (pm.fontBoundingBoxAscent ?? pm.actualBoundingBoxAscent) +
        (pm.fontBoundingBoxDescent ?? pm.actualBoundingBoxDescent);
      ratio =
        probeCell > 0 ? Math.max(0, (natural - probeCell) / probeCell) : 0;
      cache.set(key, ratio);
    }
    return cell * ratio;
  };
}

/**
 * Headless approximate metrics (no DOM): ascent ≈ 0.8·em, descent ≈ 0.2·em,
 * and no external leading — the shape has no font to read one from.
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

/** The generic families a missing family falls back to. */
const PROBE_GENERICS = ['monospace', 'serif', 'sans-serif'];
/** Wide and narrow glyphs side by side, so any two faces differ in width. */
const PROBE_TEXT = 'mmmmmmmmmmlli1WQ@';

/**
 * Which font families this host can really draw: a face the registry holds,
 * or one the canvas paints as itself — asked of the canvas the painter uses,
 * by the classic test: list the family ahead of a generic one and see whether
 * the width moves. Unknown families fall straight through to the generic, so
 * it never does for them; checking three generics catches the family that
 * happens to BE one of them (Courier as the monospace default).
 *
 * This is the answer to "will text in this family come out in this family?",
 * the question font substitution needs (see the layout's `fontAvailable`) —
 * not `document.fonts.check()`, which reports true for any name with no
 * @font-face to wait on, installed or not.
 *
 * Only positive answers are remembered: a face still loading answers no now
 * and yes on the relayout its arrival triggers. Without a DOM the registry
 * alone answers.
 */
export function createFontAvailability(registry?: FontRegistry): FontAvailable {
  const drawable = new Set<string>();
  let ctx: CanvasRenderingContext2D | null | undefined;
  return (family) => {
    if (registry?.hasFamily(family)) return true;
    const key = family.toLowerCase();
    if (drawable.has(key)) return true;
    if (ctx === undefined) {
      const doc = (globalThis as { document?: Document }).document;
      ctx = doc?.createElement('canvas').getContext('2d') ?? null;
    }
    const c = ctx;
    if (!c) return false;
    const quoted = `"${family.replace(/["\\]/g, '\\$&')}"`;
    const drawn = PROBE_GENERICS.some((generic) => {
      c.font = `72px ${generic}`;
      const fallback = c.measureText(PROBE_TEXT).width;
      c.font = `72px ${quoted}, ${generic}`;
      return c.measureText(PROBE_TEXT).width !== fallback;
    });
    if (drawn) drawable.add(key);
    return drawn;
  };
}
