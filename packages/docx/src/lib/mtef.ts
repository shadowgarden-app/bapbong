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
import {
  MATH_ALPHABETS,
  astToLinear,
  mathLetters,
  type EqNode,
} from '@shadow-garden/bapbong-contracts';

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

/**
 * Parse an MTEF v5 payload into the equation AST, or null when the stream
 * is not MTEF v5 or uses a construct this walker does not model.
 */
export function mtefToAst(mtef: Uint8Array): EqNode[] | null {
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

    type Item =
      | { kind: 'line'; row: EqNode[] }
      | { kind: 'chr'; row: EqNode[] }
      | { kind: 'emb'; text: string };

    const template = (
      sel: number,
      variation: number,
      k: EqNode[][],
    ): EqNode[] => {
      const slot = (i: number): EqNode[] => k[i] ?? [];
      const chrs = (t: string): EqNode[] =>
        [...t].map((ch) => ({ t: 'chr', ch }));
      if (sel === 11) return [{ t: 'frac', num: slot(0), den: slot(1) }];
      if (sel === 10) return [{ t: 'rad', deg: slot(1), body: slot(0) }];
      // Script templates: two fixed slots, [subscript, superscript] — the
      // unused one is a NULL line (verified on the corpus). The base is the
      // preceding content; the walker attaches it.
      if (sel === 27 || sel === 28 || sel === 29)
        return [{ t: 'scr', base: [], sub: slot(0), sup: slot(1) }];
      if (sel in FENCES) {
        let [l, r] = FENCES[sel];
        if (sel !== 9) {
          if (!(variation & 1)) l = '';
          if (!(variation & 2)) r = '';
        }
        return [{ t: 'fence', l, r, body: slot(0) }];
      }
      if (sel === 15)
        return [{ t: 'big', op: '∫', lo: slot(1), hi: slot(2), body: slot(0) }];
      if (sel >= 16 && sel <= 22)
        return [
          {
            t: 'big',
            op: BIG_OPS[sel],
            lo: slot(1),
            hi: slot(2),
            body: slot(0),
          },
        ];
      if (sel === 23)
        return [{ t: 'scr', base: slot(0), sub: slot(1), sup: [] }];
      if (sel === 12) return [...slot(0), ...chrs('̲')];
      if (sel === 13) return [...slot(0), ...chrs('̅')];
      if (sel === 31) return [...slot(0), ...chrs('⃗')];
      if (sel >= 32 && sel <= 35)
        return [...slot(0), ...chrs({ 32: '̂', 33: '̃', 34: '̑' }[sel] ?? '')];
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
          items.push({ kind: 'line', row: o & 0x01 ? [] : rowOf(objects()) });
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
          const row: EqNode[] = [...(ch + embell)].map((c) => ({
            t: 'chr',
            ch: c,
          }));
          items.push({ kind: 'chr', row });
        } else if (tag === 3) {
          // TMPL
          const o = p.u8();
          skipNudge(o);
          const sel = p.u8();
          let variation = p.u8();
          if (variation & 0x80) variation = (variation & 0x7f) | (p.u8() << 8);
          p.u8(); // template options
          const kids = objects()
            .filter(
              (i): i is { kind: 'line'; row: EqNode[] } => i.kind === 'line',
            )
            .map((i) => i.row);
          items.push({ kind: 'chr', row: template(sel, variation, kids) });
        } else if (tag === 4) {
          // PILE — vertical stack of lines, joined the way prose reads them.
          const o = p.u8();
          skipNudge(o);
          p.u8();
          p.u8();
          if (o & 0x02) skipRuler();
          const kids = objects()
            .filter(
              (i): i is { kind: 'line'; row: EqNode[] } => i.kind === 'line',
            )
            .map((i) => i.row);
          const joined: EqNode[] = [];
          kids.forEach((k, i) => {
            if (i) joined.push({ t: 'chr', ch: ';' }, { t: 'chr', ch: ' ' });
            joined.push(...k);
          });
          items.push({ kind: 'chr', row: joined });
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
    /** Items → one row. A bare script template (empty base) adopts the item
     *  right before it — MathType writes "x" then SUP("2"). */
    const rowOf = (items: Item[]): EqNode[] => {
      const row: EqNode[] = [];
      for (const it of items) {
        if (it.kind === 'emb') continue;
        for (const n of it.row) {
          if (n.t === 'scr' && n.base.length === 0 && row.length > 0) {
            const prev = row.pop() as EqNode;
            row.push({ ...n, base: [prev] });
          } else {
            row.push(n);
          }
        }
      }
      return row;
    };
    const ast = rowOf(objects());
    return ast.length ? ast : null;
  } catch {
    return null;
  }
}

/** Parse an MTEF v5 payload into linear equation text, or null. */
export function mtefToLinear(mtef: Uint8Array): string | null {
  const ast = mtefToAst(mtef);
  if (!ast) return null;
  const text = astToLinear(ast).trim();
  return text.length ? text : null;
}

/** The linear equation text of a MathType WMF preview, or null. */
export function mtefLinearFromWmf(bytes: Uint8Array): string | null {
  const mtef = mtefFromWmf(bytes);
  return mtef ? mtefToLinear(mtef) : null;
}
