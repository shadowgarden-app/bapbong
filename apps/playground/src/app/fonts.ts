import { FontRegistry } from '@shadow-garden/bapbong-measuring';

/**
 * Bundled metric-compatible fonts for engine-independent layout. Each stands in
 * for a font Word documents actually ask for, so measuring from its outlines
 * reproduces Word's line-breaking and pagination on every platform — unlike
 * canvas `measureText`, which reports whatever face the engine picked, rounded
 * to whole pixels, and with no external leading at all.
 *
 * @fontsource ships each face split into `unicode-range` subsets; we load the
 * three a Latin/Vietnamese document needs and register them as one multi-file
 * face (FontRegistry routes each character to the subset that has its glyph).
 * The `.css` imports in styles.css load the same faces for canvas rendering.
 */

/** Each bundled family, and the document family names it answers to. Verified
 *  before adding: measured character by character against the fonts installed
 *  on macOS, Tinos and Arimo return byte-identical advance widths to Times New
 *  Roman and Arial across 39 ASCII and 13 Vietnamese characters — so bundling
 *  them changes line HEIGHT (the point) without touching line breaking.
 *
 *  Their vertical metrics reach Word's numbers by different routes, which is
 *  why createFontRegistryMetrics has to implement the whole rule: Tinos leaves
 *  USE_TYPO_METRICS clear and its usWin/hhea values already match Times New
 *  Roman, while Arimo SETS the flag — read its usWin values instead of its
 *  sTypo ones and every Arial line comes out 25% too tall. */
const FAMILIES = [
  { file: 'carlito', names: ['Carlito', 'Calibri'] },
  { file: 'tinos', names: ['Tinos', 'Times New Roman'] },
  { file: 'arimo', names: ['Arimo', 'Arial'] },
] as const;

/** Subsets covering Latin + Vietnamese (cyrillic/greek omitted — not needed). */
const SUBSETS = ['latin', 'latin-ext', 'vietnamese'] as const;

const VARIANTS = [
  { bold: false, italic: false, weight: '400', style: 'normal' },
  { bold: true, italic: false, weight: '700', style: 'normal' },
  { bold: false, italic: true, weight: '400', style: 'italic' },
  { bold: true, italic: true, weight: '700', style: 'italic' },
] as const;

/**
 * Fetch and parse the bundled woff faces into a {@link FontRegistry}. Served at
 * `/fonts/<family>/*.woff` (see project.json assets). Individual fetch failures
 * are tolerated — a missing subset just falls through to canvas at measure time.
 */
export async function loadBundledFonts(): Promise<FontRegistry> {
  const registry = new FontRegistry();
  await Promise.all(
    FAMILIES.flatMap(({ file, names }) =>
      VARIANTS.flatMap((v) =>
        SUBSETS.map(async (subset) => {
          const url = `fonts/${file}/${file}-${subset}-${v.weight}-${v.style}.woff`;
          try {
            const res = await fetch(url);
            if (!res.ok) return;
            const bytes = await res.arrayBuffer();
            for (const family of names) {
              registry.registerBytes(
                family,
                { bold: v.bold, italic: v.italic },
                bytes.slice(0),
              );
            }
          } catch {
            /* offline / missing subset — canvas fallback covers it */
          }
        }),
      ),
    ),
  );
  return registry;
}
