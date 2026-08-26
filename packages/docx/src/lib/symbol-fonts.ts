/**
 * Symbol fonts: fonts whose bytes are PICTURES, not letters.
 *
 * The same glyph reaches the importer three ways — `w:sym w:char="F06F"`
 * (PUA-offset), an ordinary run whose rFonts names the font (a form's ticked
 * checkbox is often just the letter "x" in Wingdings), and a numbering
 * level's `w:lvlText` (Word's default bullet is U+F0B7 in Symbol). All three
 * go through these tables, keyed by the LOW byte, so a document renders its
 * pictures anywhere instead of showing tofu where the font is missing.
 */
/** Wingdings code point → the Unicode character Word draws for it. Keyed by
 *  the LOW byte, because the same glyph reaches us two ways: as `w:sym
 *  w:char="F06F"` (PUA-offset) and as an ordinary run of text whose rFonts
 *  says Wingdings — a form's ticked checkbox is often just the letter "x"
 *  in that font. Both go through symbolChar. */
export const WINGDINGS: Record<number, string> = {
  0x4a: '☺',
  0x4b: '😐',
  0x4c: '☹',
  0x4d: '💣',
  0x4e: '☠',
  0x51: '✈',
  0x6c: '●',
  0x6e: '■',
  // The checkbox family: Word draws these at text size, so the Unicode
  // BALLOT BOX forms match it far better than the tiny ▫/□ geometric ones.
  0x6f: '☐',
  0x70: '☐',
  0x71: '☐',
  0x72: '☐',
  0x73: '☐',
  0x75: '◆',
  0x78: '☒', // "x" — a ticked box, the usual mark in Vietnamese HR forms
  0xa7: '▪',
  0xa8: '☐',
  0xb7: '•',
  0xe0: '→',
  0xfb: '✗',
  0xfc: '✔',
  0xfd: '☒',
  0xfe: '☑',
};

/** Wingdings 2 is a DIFFERENT font with its own layout — not a superset of
 *  Wingdings, though it shares the name's stem: 0x52 is the ticked box here
 *  and a pointing hand there. Word's HR forms tick their boxes with it
 *  (`w:sym w:font="Wingdings 2" w:char="F052"`). Codes per Microsoft's chart
 *  as tabulated at alanwood.net/demos/wingdings-2.html; the Unicode 7 astral
 *  equivalents (🗴, 🗵, ⯾) are traded for BMP look-alikes that every font
 *  set has (✗, ☒, ⊗). */
export const WINGDINGS_2: Record<number, string> = {
  0x4f: '✗', // 🗴 ballot script X
  0x50: '✓',
  0x51: '☒', // 🗵 ballot box with script X
  0x52: '☑',
  0x53: '☒',
  0x54: '☒',
  0x55: '⊗', // ⯾ circled X
  0x56: '⊗',
  0x95: '•',
  0x96: '●',
  0x97: '●',
  0x98: '●',
  0x99: '○',
  0x9a: '○',
  0x9b: '○',
  0x9c: '○',
  0x9d: '◉',
  0x9e: '⦿',
  0x9f: '◾',
  0xa0: '■',
  0xa1: '◼',
  0xa2: '■',
  // The empty boxes: BALLOT BOX, like WINGDINGS' 0x6F–0x73, so an unticked
  // one sits beside a ticked one at the same size.
  0xa3: '☐',
  0xa4: '☐',
  0xa5: '☐',
};

/** Monotype Sorts — Monotype's clone of ITC Zapf Dingbats, same encoding.
 *  Only the box family, which forms lean on: 0x6F–0x72 are Zapf's four
 *  squares. 0x7F is unassigned in Zapf Dingbats yet Word draws an EMPTY
 *  BALLOT BOX for it (an application form pairs "Monotype Sorts F07F Not
 *  yet" with "Wingdings 2 F052 Yes", and Word shows ☐ beside ☑). */
export const MONOTYPE_SORTS: Record<number, string> = {
  0x6f: '❏',
  0x70: '❐',
  0x71: '❑',
  0x72: '❒',
  0x7f: '☐',
};

/** Symbol font: only the few glyphs documents actually lean on (its letters
 *  are Greek, which we leave to the font). */
export const SYMBOL_FONT: Record<number, string> = {
  0xb7: '•',
  0xd7: '×',
  0xb0: '°',
  0xa0: '€',
};

/** Fonts whose bytes are pictures, not letters. */
export function symbolTable(
  font: string | undefined,
): Record<number, string> | null {
  if (!font) return null;
  const f = font.toLowerCase().trim();
  // Exact names: "Wingdings 2" and "Wingdings 3" are different fonts, and a
  // prefix match once sent Wingdings 2's ticked box through Wingdings'
  // table (0x52 there is a pointing hand — the box came out as tofu).
  if (f === 'wingdings') return WINGDINGS;
  if (f === 'wingdings 2') return WINGDINGS_2;
  if (f === 'monotype sorts') return MONOTYPE_SORTS;
  if (f === 'symbol') return SYMBOL_FONT;
  return null;
}

/** One character of a symbol font → the Unicode Word shows. Mapped chars come
 *  back font-independent (no `font`), so they render anywhere; unmapped ones
 *  keep their code point AND the font name, and the caller tags them with a
 *  fontFamily mark so the glyph still appears where the font is installed
 *  (and survives a save either way). */
export function symbolChar(
  code: string | undefined,
  font: string | undefined,
): { text: string; font?: string } {
  if (!code) return { text: '' };
  const n = parseInt(code, 16);
  if (Number.isNaN(n)) return { text: '' };
  // w:sym codes sit in the PUA (F0xx); a plain run's character is the byte.
  const mapped = symbolTable(font)?.[n & 0xff];
  if (mapped) return { text: mapped };
  return { text: String.fromCodePoint(n), ...(font && { font }) };
}

/** A run of ordinary text set in a symbol font, translated to Unicode. Null
 *  when the font isn't one (the overwhelmingly common case — one cheap map
 *  lookup) or when nothing in the text maps, so the run passes through
 *  untouched and keeps its font. */
export function symbolFontText(
  text: string,
  font: string | undefined,
): string | null {
  const table = symbolTable(font);
  if (!table) return null;
  let changed = false;
  let out = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    // The tables are keyed by the font's BYTE. A character reaches us either
    // as that byte (run text) or PUA-offset (w:sym, and a numbering level's
    // w:lvlText — Word writes its default bullet as U+F0B7). Only the PUA
    // window is folded down; masking every code point would map unrelated
    // Unicode letters that happen to share a low byte.
    const mapped = table[cp >= 0xf000 && cp <= 0xf0ff ? cp & 0xff : cp];
    if (mapped) {
      out += mapped;
      changed = true;
    } else {
      out += ch;
    }
  }
  return changed ? out : null;
}
