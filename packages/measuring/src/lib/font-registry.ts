import * as opentype from 'opentype.js';
import {
  glyphCount,
  type FontSpec,
  type MeasureMetrics,
  type MeasureText,
} from '@shadow-garden/bapbong-contracts';

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
 * Holds parsed font faces keyed by family + bold/italic and resolves per-glyph
 * advance widths + vertical metrics from the font files. Widths derived here are
 * byte-for-byte identical on every platform and WebView engine — that
 * determinism is the whole reason layout is measured from fonts rather than the
 * browser's `measureText`.
 *
 * A single face may be backed by **several files** because webfonts are commonly
 * shipped subsetted by `unicode-range` (e.g. `@fontsource` splits latin /
 * latin-ext / vietnamese). Each requested character is routed to the first file
 * in the face that actually contains its glyph.
 */
export class FontRegistry {
  private readonly faces = new Map<string, opentype.Font[]>();
  /** Per-face memo: code point → the file that serves it (or its .notdef fallback). */
  private readonly glyphMemo = new Map<string, Map<number, opentype.Font>>();

  /** Register a parsed opentype font for a family + variant (appends to the
   *  face's file list; register subset files in coverage-priority order). */
  register(family: string, variant: FontVariant, font: opentype.Font): void {
    const key = faceKey(family, !!variant.bold, !!variant.italic);
    const list = this.faces.get(key);
    if (list) list.push(font);
    else this.faces.set(key, [font]);
    this.glyphMemo.clear(); // a new file may now cover code points cached elsewhere
  }

  /** Parse font-file bytes (TTF/OTF/WOFF) and register the resulting face. */
  registerBytes(
    family: string,
    variant: FontVariant,
    bytes: ArrayBuffer,
  ): void {
    this.register(family, variant, opentype.parse(bytes));
  }

  /** The files backing a spec: the exact family+bold+italic face, else the
   *  family's regular face, else empty. */
  private files(spec: FontSpec): opentype.Font[] {
    return (
      this.faces.get(faceKey(spec.family, spec.bold, spec.italic)) ??
      this.faces.get(faceKey(spec.family, false, false)) ??
      []
    );
  }

  /** Whether any registered file can serve this spec. */
  has(spec: FontSpec): boolean {
    return this.files(spec).length > 0;
  }

  /** The face's first file, for shared vertical metrics (all subset files of one
   *  face carry the same hhea/head metrics), or null. */
  primary(spec: FontSpec): opentype.Font | null {
    return this.files(spec)[0] ?? null;
  }

  /** The file in the face that has a glyph for `codePoint`, else the first file
   *  (whose .notdef advance is used), else null. Memoised per face. */
  fileForCodePoint(spec: FontSpec, codePoint: number): opentype.Font | null {
    const list = this.files(spec);
    if (list.length === 0) return null;
    if (list.length === 1) return list[0];
    const key = faceKey(spec.family, spec.bold, spec.italic);
    let memo = this.glyphMemo.get(key);
    if (!memo) {
      memo = new Map();
      this.glyphMemo.set(key, memo);
    }
    const cached = memo.get(codePoint);
    if (cached) return cached;
    const ch = String.fromCodePoint(codePoint);
    const found = list.find((f) => f.charToGlyph(ch).index !== 0) ?? list[0];
    memo.set(codePoint, found);
    return found;
  }
}

/**
 * A {@link MeasureText} that sums glyph advance widths from a registered face —
 * engine-independent, unlike canvas `measureText`. Each character is routed to
 * the subset file that contains its glyph. Families absent from the registry
 * defer to `fallback` (e.g. a canvas measurer). Kerning is not applied, matching
 * Word's default line-breaking (kerning-for-fonts is off by default).
 */
export function createFontRegistryMeasurer(
  registry: FontRegistry,
  fallback: MeasureText,
): MeasureText {
  return (text, font) => {
    if (!registry.has(font)) return fallback(text, font);
    const px = font.sizePt * PT_TO_PX;
    let total = 0;
    for (const ch of text) {
      const file = registry.fileForCodePoint(font, ch.codePointAt(0) ?? 0);
      if (!file) continue;
      total +=
        ((file.charToGlyph(ch).advanceWidth ?? 0) / file.unitsPerEm) * px;
    }
    // Tracking is charged arithmetically here, not by the browser. This
    // measurer already diverges from canvas by more than that (no kerning, no
    // shaping) — it exists to be DETERMINISTIC across environments, not to
    // match canvas pixel for pixel.
    return (
      total * (font.scaleX ?? 1) + (font.letterSpacing ?? 0) * glyphCount(text)
    );
  };
}

/**
 * A {@link MeasureMetrics} reading a registered face's hhea ascent/descent
 * (scaled to px), deferring to `fallback` for unknown families.
 */
export function createFontRegistryMetrics(
  registry: FontRegistry,
  fallback: MeasureMetrics,
): MeasureMetrics {
  return (font) => {
    const face = registry.primary(font);
    if (!face) return fallback(font);
    const scale = (font.sizePt * PT_TO_PX) / face.unitsPerEm;
    return {
      ascent: face.ascender * scale,
      descent: Math.abs(face.descender) * scale,
    };
  };
}
