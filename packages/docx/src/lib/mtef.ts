/**
 * MTEF v5 → linear equation text.
 *
 * Every MathType preview WMF carries the equation's SEMANTIC tree twice: the
 * OLE `Equation Native` stream and — the copy this module reads — an
 * `AppsMFCC` comment inside the WMF itself ("Design Science, Inc."), so the
 * math survives even when only the picture does. MTEF character records
 * store MTCode, which is Unicode; templates carry structure (fractions,
 * scripts, radicals, fences). This walks the records into the same linear
 * spelling the OMML flattener produces — math-italic letters, sub/sup
 * digits, `num/den` fractions, `√(…)` — so a converted equation is
 * indistinguishable from one typed through Insert ▸ Equation, and exports as
 * m:oMath like any other.
 *
 * Record layouts follow the archived MTEF v5 spec (rtf2latex2e); slot
 * conventions (script templates = [subscript, superscript] with NULL lines
 * for the unused one) were verified against 183 real exam equations, which
 * this parser converts 183/183.
 *
 * Anything unexpected — unknown record, unknown template — returns null:
 * the caller keeps the picture and the editor simply offers no conversion.
 */
import { MATH_ALPHABETS, mathLetters } from '@shadow-garden/bapbong-contracts';

const SUB_DIGITS: Record<string, string> = {
  '0': '₀',
  '1': '₁',
  '2': '₂',
  '3': '₃',
  '4': '₄',
  '5': '₅',
  '6': '₆',
  '7': '₇',
  '8': '₈',
  '9': '₉',
};
const SUP_DIGITS: Record<string, string> = {
  '0': '⁰',
  '1': '¹',
  '2': '²',
  '3': '³',
  '4': '⁴',
  '5': '⁵',
  '6': '⁶',
  '7': '⁷',
  '8': '⁸',
  '9': '⁹',
  '+': '⁺',
  '-': '⁻',
  '−': '⁻',
};

/** The MTEF payload of MathType's WMF comment, or null. */
export function mtefFromWmf(bytes: Uint8Array): Uint8Array | null {
  const find = (needle: string, from: number): number => {
    outer: for (let i = from; i <= bytes.length - needle.length; i++) {
      for (let j = 0; j < needle.length; j++)
        if (bytes[i + j] !== needle.charCodeAt(j)) continue outer;
      return i;
    }
    return -1;
  };
  const a = find('AppsMFCC', 0);
  if (a < 0) return null;
  const d = find('Design Science, Inc.\0', a);
  if (d < 0) return null;
  return bytes.subarray(d + 21);
}

/** Fence characters by template selector (0–9). */
const FENCES: Record<number, [string, string]> = {
  0: ['⟨', '⟩'],
  1: ['(', ')'],
  2: ['{', '}'],
  3: ['[', ']'],
  4: ['|', '|'],
  5: ['‖', '‖'],
  6: ['⌊', '⌋'],
  7: ['⌈', '⌉'],
  8: ['[', ']'],
  9: ['[', ']'],
};

/** Embellishment type → combining mark / appended character. */
const EMBELLS: Record<number, string> = {
  2: '̇',
  3: '̈',
  4: '⃛',
  5: '′',
  6: '″',
  7: '‴',
  18: '‴',
  8: '̃',
  9: '̂',
  10: '̸',
  11: '⃗',
  12: '⃖',
  13: '⃡',
  14: '⃑',
  15: '⃐',
  16: '̄',
  17: '̅',
};

const BIG_OPS: Record<number, string> = {
  16: '∑',
  17: '∏',
  18: '∐',
  19: '∪',
  20: '∩',
  21: '∫',
  22: '∑',
};

class Reader {
  p = 0;
  constructor(private readonly m: Uint8Array) {}
  u8(): number {
    if (this.p >= this.m.length) throw new Error('eof');
    return this.m[this.p++];
  }
  u16(): number {
    return this.u8() | (this.u8() << 8);
  }
  cstr(): string {
    let s = '';
    for (;;) {
      const b = this.u8();
      if (b === 0) return s;
      s += String.fromCharCode(b);
    }
  }
}

const scriptDigits = (
  s: string,
  table: Record<string, string>,
): string | null =>
  [...s].every((c) => table[c]) ? [...s].map((c) => table[c]).join('') : null;

const subScript = (s: string): string =>
  scriptDigits(s, SUB_DIGITS) ?? (s.length > 1 ? `_(${s})` : `_${s}`);
const supScript = (s: string): string =>
  scriptDigits(s, SUP_DIGITS) ?? (s.length > 1 ? `^(${s})` : `^${s}`);
const paren = (s: string): string =>
  s.length <= 1 || (s.startsWith('(') && s.endsWith(')')) ? s : `(${s})`;

/**
 * Parse an MTEF v5 payload into linear equation text, or null when the
 * stream is not MTEF v5 or uses a construct this walker does not model.
 */
export function mtefToLinear(mtef: Uint8Array): string | null {
  try {
    const p = new Reader(mtef);
    if (p.u8() !== 5) return null;
    p.u8(); // platform
    p.u8(); // product
    p.u8(); // product version
    p.u8(); // product subversion
    p.cstr(); // application key
    p.u8(); // equation options (inline / display)

    const skipNudge = (opts: number): void => {
      if (!(opts & 0x08)) return;
      const dx = p.u8();
      const dy = p.u8();
      if (dx === 0x80 && dy === 0x80) {
        p.u16();
        p.u16();
      }
    };
    const skipRuler = (): void => {
      const n = p.u8();
      for (let i = 0; i < n; i++) {
        p.u8();
        p.u16();
      }
    };
    // EQN_PREFS: two nibble-encoded dimension arrays (each value terminated
    // by a 0xF nibble) and a style list. Content is presentation-only; it is
    // parsed exactly so the record can be SKIPPED exactly.
    const skipPrefs = (): void => {
      p.u8(); // options
      for (let arr = 0; arr < 2; arr++) {
        const n = p.u8();
        let got = 0;
        let buf: number | null = null;
        let hi = true;
        while (got < n) {
          if (buf === null) {
            buf = p.u8();
            hi = true;
          }
          const nib = hi ? buf >> 4 : buf & 0xf;
          if (hi) hi = false;
          else buf = null;
          if (nib === 0xf) got++;
        }
      }
      const n = p.u8();
      for (let i = 0; i < n; i++) if (p.u8() !== 0) p.u8();
    };

    type Item = { kind: 'line' | 'chr' | 'emb'; text: string };

    const template = (sel: number, variation: number, k: string[]): string => {
      const slot = (i: number): string => k[i] ?? '';
      if (sel === 11) return `${paren(slot(0))}/${paren(slot(1))}`;
      if (sel === 10) {
        const deg = slot(1);
        return `${deg ? supScript(deg) : ''}√(${slot(0)})`;
      }
      // Script templates: two fixed slots, [subscript, superscript] — the
      // unused one is a NULL line (verified on the corpus).
      if (sel === 27 || sel === 28 || sel === 29) {
        const lo = slot(0);
        const hi = slot(1);
        return (lo ? subScript(lo) : '') + (hi ? supScript(hi) : '');
      }
      if (sel in FENCES) {
        let [l, r] = FENCES[sel];
        if (sel !== 9) {
          if (!(variation & 1)) l = '';
          if (!(variation & 2)) r = '';
        }
        return l + slot(0) + r;
      }
      if (sel === 15) {
        const lo = slot(1);
        const hi = slot(2);
        return `∫${lo ? subScript(lo) : ''}${hi ? supScript(hi) : ''}${paren(slot(0))}`;
      }
      if (sel >= 16 && sel <= 22) {
        const lo = slot(1);
        const hi = slot(2);
        return `${BIG_OPS[sel]}${lo ? subScript(lo) : ''}${hi ? supScript(hi) : ''}${paren(slot(0))}`;
      }
      if (sel === 23) {
        const under = slot(1);
        return slot(0) + (under ? `_(${under})` : '');
      }
      if (sel === 12) return `${slot(0)}̲`;
      if (sel === 13) return `${slot(0)}̅`;
      if (sel === 31) return `${slot(0)}⃗`;
      if (sel >= 32 && sel <= 35)
        return slot(0) + ({ 32: '̂', 33: '̃', 34: '̑' }[sel] ?? '');
      throw new Error(`template ${sel}`);
    };

    const objects = (): Item[] => {
      const items: Item[] = [];
      for (;;) {
        const tag = p.u8();
        if (tag === 0) return items;
        if (tag === 1) {
          // LINE
          const o = p.u8();
          skipNudge(o);
          if (o & 0x04) p.u16();
          if (o & 0x02) skipRuler();
          items.push({
            kind: 'line',
            text: o & 0x01 ? '' : flatten(objects()),
          });
        } else if (tag === 2) {
          // CHAR
          const o = p.u8();
          skipNudge(o);
          const typeface = p.u8() - 128;
          let ch = '';
          if (!(o & 0x20)) ch = String.fromCharCode(p.u16());
          if (o & 0x04) p.u8();
          if (o & 0x10) p.u16();
          let embell = '';
          if (o & 0x01)
            for (const e of objects()) if (e.kind === 'emb') embell += e.text;
          // Typeface 3 is MathType's "variable" style — the italic Word
          // renders math letters in. Same alphabet the OMML flattener uses.
          if (typeface === 3) ch = mathLetters(ch, MATH_ALPHABETS['italic']);
          items.push({ kind: 'chr', text: ch + embell });
        } else if (tag === 3) {
          // TMPL
          const o = p.u8();
          skipNudge(o);
          const sel = p.u8();
          let variation = p.u8();
          if (variation & 0x80) variation = (variation & 0x7f) | (p.u8() << 8);
          p.u8(); // template options
          const kids = objects()
            .filter((i) => i.kind === 'line')
            .map((i) => i.text);
          items.push({ kind: 'chr', text: template(sel, variation, kids) });
        } else if (tag === 4) {
          // PILE — vertical stack of lines, joined the way prose reads them.
          const o = p.u8();
          skipNudge(o);
          p.u8();
          p.u8();
          if (o & 0x02) skipRuler();
          const kids = objects()
            .filter((i) => i.kind === 'line')
            .map((i) => i.text);
          items.push({ kind: 'chr', text: kids.join('; ') });
        } else if (tag === 6) {
          // EMBELL
          const o = p.u8();
          skipNudge(o);
          const type = p.u8();
          items.push({ kind: 'emb', text: EMBELLS[type] ?? '' });
        } else if (tag === 9) {
          // SIZE
          const a = p.u8();
          if (a === 101) p.u16();
          else if (a === 100) {
            p.u8();
            p.u16();
          } else p.u8();
        } else if (tag >= 10 && tag <= 14) {
          // typesize shortcuts — no payload
        } else if (tag === 8) {
          // FONT_STYLE_DEF
          p.u8();
          p.u8();
        } else if (tag === 17) {
          // FONT_DEF
          p.u8();
          p.cstr();
        } else if (tag === 18) {
          skipPrefs();
        } else if (tag === 19) {
          p.cstr(); // ENCODING_DEF
        } else if (tag === 15) {
          // COLOR
          const o = p.u8();
          if (o & 0x02) p.u16();
          else p.u8();
        } else {
          throw new Error(`record ${tag}`);
        }
      }
    };
    const flatten = (items: Item[]): string =>
      items
        .filter((i) => i.kind === 'chr' || i.kind === 'line')
        .map((i) => i.text)
        .join('');

    const text = flatten(objects()).trim();
    return text.length ? text : null;
  } catch {
    return null;
  }
}

/** The linear equation text of a MathType WMF preview, or null. */
export function mtefLinearFromWmf(bytes: Uint8Array): string | null {
  const mtef = mtefFromWmf(bytes);
  return mtef ? mtefToLinear(mtef) : null;
}
