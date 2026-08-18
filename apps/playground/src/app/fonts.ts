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

/** The unicode ranges @fontsource gives each subset (copied from its css), so
 *  the alias faces below compose per character exactly like the originals. */
const SUBSET_RANGES: Record<(typeof SUBSETS)[number], string> = {
  latin:
    'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD',
  'latin-ext':
    'U+0100-02BA,U+02BD-02C5,U+02C7-02CC,U+02CE-02D7,U+02DD-02FF,U+0304,U+0308,U+0329,U+1D00-1DBF,U+1E00-1E9F,U+1EF2-1EFF,U+2020,U+20A0-20AB,U+20AD-20C0,U+2113,U+2C60-2C7F,U+A720-A7FF',
  vietnamese:
    'U+0102-0103,U+0110-0111,U+0128-0129,U+0168-0169,U+01A0-01A1,U+01AF-01B0,U+0300-0301,U+0303-0304,U+0308-0309,U+0323,U+0329,U+1EA0-1EF9,U+20AB',
};

/**
 * Canvas paints by the DOCUMENT's family name ("Calibri"), while the bundled
 * @font-face (styles.css) is named after the bundled font ("Carlito"). On a
 * machine without Calibri the canvas fell back to whatever the engine picked —
 * a serif, wider than the Carlito the layout measured — and words painted over
 * each other. Register the bundled face under the document's name as well,
 * system font FIRST: `local()` keeps the real Calibri where it is installed
 * (same metrics, its own glyphs) and the metric twin fills in elsewhere. Loaded
 * up front, with the registry, so the first paint already has it.
 */
function aliasFace(
  alias: string,
  url: string,
  v: (typeof VARIANTS)[number],
  subset: (typeof SUBSETS)[number],
): Promise<unknown> {
  if (typeof FontFace === 'undefined' || !document.fonts)
    return Promise.resolve();
  // local() names a FACE, not a family: local('Times New Roman') is the
  // Regular face, and a weight-700 FontFace built on it paints bold text
  // with regular glyphs. Name the face the variant really is — its full name
  // ("Times New Roman Bold Italic") and the PostScript spellings macOS and
  // Windows use for it.
  const style =
    v.bold && v.italic
      ? 'Bold Italic'
      : v.bold
        ? 'Bold'
        : v.italic
          ? 'Italic'
          : '';
  const compact = alias.replace(/\s+/g, '');
  const ps = style.replace(/\s+/g, '');
  const locals = [
    style ? `${alias} ${style}` : alias,
    style ? `${compact}-${ps}` : compact,
    style ? `${compact}-${ps}MT` : `${compact}MT`,
    style ? `${compact}PS-${ps}MT` : `${compact}PSMT`,
  ]
    .map((n) => `local('${n}')`)
    .join(', ');
  const face = new FontFace(alias, `${locals}, url(${url})`, {
    weight: v.weight,
    style: v.style,
    unicodeRange: SUBSET_RANGES[subset],
  });
  document.fonts.add(face);
  return face.load().catch(() => undefined); // offline: canvas fallback covers it
}

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
          // The alias faces for painting; the first name IS the bundled font
          // and styles.css already declares it.
          const aliases = Promise.all(
            names.slice(1).map((n) => aliasFace(n, url, v, subset)),
          );
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
          await aliases;
        }),
      ),
    ),
  );
  return registry;
}
