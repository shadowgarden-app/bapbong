import * as opentype from 'opentype.js';
import type { FontSpec, MeasureMetrics, MeasureText } from '@shadow-garden/bapbong-contracts';

const PT_TO_PX = 96 / 72;

/** One requestable font variant. Missing flags default to the regular face. */
export interface FontVariant {
  bold?: boolean;
  italic?: boolean;
}

/** Map key for a face: family (case-insensitive) + weight + slant. */
function faceKey(family: string, bold: boolean, italic: boolean): string {
  return `${family.toLowerCase()}|${bold ? 'b' : ''}|${italic ? 'i' : ''}`;
}

/**
 * Holds parsed font faces keyed by family + bold/italic and resolves the best
 * match for a {@link FontSpec}. Widths and vertical metrics derived from a
 * registered face are byte-for-byte identical on every platform and WebView
 * engine — that determinism is the whole reason layout is measured from font
 * files instead of the browser's `measureText`.
 */
export class FontRegistry {
  private readonly faces = new Map<string, opentype.Font>();

  /** Register an already-parsed opentype font for a family + variant. */
  register(family: string, variant: FontVariant, font: opentype.Font): void {
    this.faces.set(faceKey(family, !!variant.bold, !!variant.italic), font);
  }

  /** Parse font-file bytes (TTF/OTF/WOFF) and register the resulting face. */
  registerBytes(family: string, variant: FontVariant, bytes: ArrayBuffer): void {
    this.register(family, variant, opentype.parse(bytes));
  }

  /** Whether a face (exact or regular-variant fallback) can serve this spec. */
  has(spec: FontSpec): boolean {
    return this.resolve(spec) !== null;
  }

  /**
   * The best registered face for a spec: the exact family+bold+italic face if
   * present, else the family's regular face, else null. The regular-variant
   * fallback lets a document lay out from real metrics even when only one weight
   * was bundled (bold/italic then approximated by the same outlines).
   */
  resolve(spec: FontSpec): opentype.Font | null {
    return (
      this.faces.get(faceKey(spec.family, spec.bold, spec.italic)) ??
      this.faces.get(faceKey(spec.family, false, false)) ??
      null
    );
  }
}

/**
 * A {@link MeasureText} that sums glyph advance widths from a registered face —
 * engine-independent, unlike canvas `measureText`. Families absent from the
 * registry defer to `fallback` (e.g. a canvas measurer). Kerning is disabled to
 * match Word's default line-breaking (kerning-for-fonts is off by default).
 */
export function createFontRegistryMeasurer(registry: FontRegistry, fallback: MeasureText): MeasureText {
  return (text, font) => {
    const face = registry.resolve(font);
    if (!face) return fallback(text, font);
    return face.getAdvanceWidth(text, font.sizePt * PT_TO_PX, { kerning: false });
  };
}

/**
 * A {@link MeasureMetrics} reading a registered face's hhea ascent/descent
 * (scaled to px), deferring to `fallback` for unknown families.
 */
export function createFontRegistryMetrics(registry: FontRegistry, fallback: MeasureMetrics): MeasureMetrics {
  return (font) => {
    const face = registry.resolve(font);
    if (!face) return fallback(font);
    const scale = (font.sizePt * PT_TO_PX) / face.unitsPerEm;
    return { ascent: face.ascender * scale, descent: Math.abs(face.descender) * scale };
  };
}
