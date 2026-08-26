/**
 * Unicode Mathematical Alphanumeric letterforms — how equations spell their
 * letters in plain text (𝑥, 𝒫, ℝ).
 *
 * Word renders OMML letters in math style — italic by default, or the
 * variant `m:scr`/`m:sty` names. This model has no math typesetting yet, so
 * the letterform IS the text: the importer maps OMML letters through these
 * alphabets, and the equation-insert command maps what the user types the
 * same way. One table, one behaviour.
 *
 * Each block gives base codepoints for A and a plus the letters Unicode had
 * already encoded elsewhere (the "holes" the block skips — ℎ ℬ ℤ …).
 */
export interface MathAlphabet {
  upper: number;
  lower: number;
  holes?: Record<string, number>;
}

/* prettier-ignore */
export const MATH_ALPHABETS: Record<string, MathAlphabet> = {
  italic: { upper: 0x1d434, lower: 0x1d44e, holes: { h: 0x210e } },
  bold: { upper: 0x1d400, lower: 0x1d41a },
  'bold-italic': { upper: 0x1d468, lower: 0x1d482 },
  script: {
    upper: 0x1d49c,
    lower: 0x1d4b6,
    holes: {
      B: 0x212c, E: 0x2130, F: 0x2131, H: 0x210b, I: 0x2110, L: 0x2112,
      M: 0x2133, R: 0x211b, e: 0x212f, g: 0x210a, o: 0x2134,
    },
  },
  fraktur: {
    upper: 0x1d504,
    lower: 0x1d51e,
    holes: { C: 0x212d, H: 0x210c, I: 0x2111, R: 0x211c, Z: 0x2128 },
  },
  'double-struck': {
    upper: 0x1d538,
    lower: 0x1d552,
    holes: {
      C: 0x2102, H: 0x210d, N: 0x2115, P: 0x2119, Q: 0x211a, R: 0x211d,
      Z: 0x2124,
    },
  },
  'sans-serif': { upper: 0x1d5a0, lower: 0x1d5ba },
  monospace: { upper: 0x1d670, lower: 0x1d68a },
};

/** ASCII letters of `text` restyled through one math alphabet; everything
 *  else (digits, operators, punctuation) passes through, matching Word —
 *  math style only reshapes letters. */
export function mathLetters(text: string, alphabet: MathAlphabet): string {
  return [...text]
    .map((ch) => {
      const hole = alphabet.holes?.[ch];
      if (hole) return String.fromCodePoint(hole);
      if (ch >= 'A' && ch <= 'Z')
        return String.fromCodePoint(alphabet.upper + ch.charCodeAt(0) - 65);
      if (ch >= 'a' && ch <= 'z')
        return String.fromCodePoint(alphabet.lower + ch.charCodeAt(0) - 97);
      return ch;
    })
    .join('');
}
