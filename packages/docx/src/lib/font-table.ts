import type {
  FontSubstitute,
  FontSubstitutes,
} from '@shadow-garden/bapbong-contracts';
import { attrOf, child, children, type OoxmlNode } from './ooxml.js';

/** Charsets whose text no Western class face can stand in for: 02 SYMBOL
 *  (the codes are pictures, not letters) and the East Asian ones — SHIFTJIS,
 *  HANGUL, JOHAB, GB2312, CHINESEBIG5. A "modern" MS Mincho drawn in Courier
 *  New would be no closer to Word than the browser's own CJK fallback. */
const NO_GENERIC_CHARSETS = new Set(['02', '80', '81', '82', '86', '88']);

/** The w:family classes a substitute can be drawn from. `script`,
 *  `decorative` and `auto` say nothing a stand-in could honour. */
const GENERIC = new Set(['roman', 'swiss', 'modern']);

/**
 * Read word/fontTable.xml (its `w:fonts` root) into what the layout needs to
 * draw a font this machine lacks: Word's own substitutes for it.
 *
 * Documents converted from PDF name fonts by their PostScript-ish names —
 * "Arial MT", "MyriadPro-Regular" — which no system installs, and Word draws
 * them in the font table's `w:altName` (else a face of the same `w:family`).
 * Only entries that offer one of the two are kept; null when none does.
 */
export function parseFontTable(
  fonts: OoxmlNode | undefined,
): FontSubstitutes | null {
  const out: FontSubstitutes = {};
  for (const font of children(fonts, 'w:font')) {
    const name = attrOf(font, 'w:name');
    if (!name) continue;
    const entry: FontSubstitute = {};
    // One w:altName per font, but its value may list several (comma-
    // separated, as East Asian fonts' localized names sometimes are).
    const altNames = (attrOf(child(font, 'w:altName'), 'w:val') ?? '')
      .split(',')
      .map((s) => s.trim())
      .filter((s) => s && s.toLowerCase() !== name.toLowerCase());
    if (altNames.length > 0) entry.altNames = altNames;
    const charset = attrOf(child(font, 'w:charset'), 'w:val');
    const family = attrOf(child(font, 'w:family'), 'w:val');
    if (
      family &&
      GENERIC.has(family) &&
      !NO_GENERIC_CHARSETS.has(charset ?? '')
    ) {
      entry.generic = family as FontSubstitute['generic'];
      if (attrOf(child(font, 'w:pitch'), 'w:val') === 'fixed')
        entry.fixedPitch = true;
    }
    if (entry.altNames || entry.generic) out[name] = entry;
  }
  return Object.keys(out).length > 0 ? out : null;
}
