import { FontRegistry } from '@shadow-garden/bapbong-measuring';

/**
 * Bundled metric-compatible fonts for engine-independent layout. Carlito is
 * metric-compatible with Calibri (Word's default since 2007), so measuring from
 * its outlines reproduces Word's line-breaking/pagination on every platform —
 * unlike canvas `measureText`, which varies by WebView engine.
 *
 * @fontsource ships each face split into `unicode-range` subsets; we load the
 * three a Latin/Vietnamese document needs and register them as one multi-file
 * face (FontRegistry routes each character to the subset that has its glyph).
 * The `.css` imports in styles.css load the same faces for canvas rendering.
 */

/** Register the Carlito faces under these family names, so a document's own
 *  "Calibri" runs and an explicit "Carlito" both resolve to real metrics. */
const FAMILY_NAMES = ['Carlito', 'Calibri'];

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
 * `/fonts/carlito/*.woff` (see project.json assets). Individual fetch failures
 * are tolerated — a missing subset just falls through to canvas at measure time.
 */
export async function loadBundledFonts(): Promise<FontRegistry> {
  const registry = new FontRegistry();
  await Promise.all(
    VARIANTS.flatMap((v) =>
      SUBSETS.map(async (subset) => {
        const url = `fonts/carlito/carlito-${subset}-${v.weight}-${v.style}.woff`;
        try {
          const res = await fetch(url);
          if (!res.ok) return;
          const bytes = await res.arrayBuffer();
          for (const family of FAMILY_NAMES) {
            registry.registerBytes(family, { bold: v.bold, italic: v.italic }, bytes.slice(0));
          }
        } catch {
          /* offline / missing subset — canvas fallback covers it */
        }
      }),
    ),
  );
  return registry;
}
