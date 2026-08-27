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

/**
 * Word's Math AutoCorrect, the symbol part: inside an equation, `\name`
 * becomes its character the moment a space completes it (the space is the
 * trigger and is consumed — Word and Google Docs both work this way).
 * Names and mappings follow Word's list — `\epsilon` is the lunate ϵ with
 * `\varepsilon` for ε, `\phi` is ϕ with `\varphi` for φ — so habits carry
 * over unchanged.
 */
/* prettier-ignore */
export const MATH_AUTOCORRECT: Record<string, string> = {
  // Greek, lowercase.
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ϵ',
  varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ',
  iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ',
  omicron: 'ο', pi: 'π', varpi: 'ϖ', rho: 'ρ', varrho: 'ϱ',
  sigma: 'σ', varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'ϕ',
  varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  // Greek, uppercase.
  Alpha: 'Α', Beta: 'Β', Gamma: 'Γ', Delta: 'Δ', Epsilon: 'Ε', Zeta: 'Ζ',
  Eta: 'Η', Theta: 'Θ', Iota: 'Ι', Kappa: 'Κ', Lambda: 'Λ', Mu: 'Μ',
  Nu: 'Ν', Xi: 'Ξ', Omicron: 'Ο', Pi: 'Π', Rho: 'Ρ', Sigma: 'Σ',
  Tau: 'Τ', Upsilon: 'Υ', Phi: 'Φ', Chi: 'Χ', Psi: 'Ψ', Omega: 'Ω',
  // Operators and relations.
  pm: '±', mp: '∓', times: '×', div: '÷', cdot: '⋅', ast: '∗',
  leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠',
  approx: '≈', equiv: '≡', sim: '∼', simeq: '≃', cong: '≅',
  propto: '∝', ll: '≪', gg: '≫', perp: '⊥', parallel: '∥',
  // Calculus and big operators.
  infty: '∞', partial: '∂', nabla: '∇', sqrt: '√', cbrt: '∛',
  int: '∫', iint: '∬', iiint: '∭', oint: '∮', sum: '∑', prod: '∏',
  // Sets and logic.
  in: '∈', notin: '∉', ni: '∋', subset: '⊂', supset: '⊃',
  subseteq: '⊆', supseteq: '⊇', cup: '∪', cap: '∩', emptyset: '∅',
  forall: '∀', exists: '∃', nexists: '∄', wedge: '∧', vee: '∨',
  neg: '¬', therefore: '∴', because: '∵',
  // Arrows.
  to: '→', rightarrow: '→', leftarrow: '←', uparrow: '↑', downarrow: '↓',
  leftrightarrow: '↔', mapsto: '↦', Rightarrow: '⇒', Leftarrow: '⇐',
  Leftrightarrow: '⇔',
  // Letterlike and blackboard.
  ell: 'ℓ', hbar: 'ℏ', Re: 'ℜ', Im: 'ℑ', aleph: 'ℵ', wp: '℘',
  doubleN: 'ℕ', doubleZ: 'ℤ', doubleQ: 'ℚ', doubleR: 'ℝ', doubleC: 'ℂ',
  // Miscellany the exam corpus leans on.
  degree: '°', prime: '′', pprime: '″', angle: '∠', triangle: '△',
  ldots: '…', cdots: '⋯', vdots: '⋮', ddots: '⋱', bullet: '∙',
  oplus: '⊕', otimes: '⊗', circ: '∘',
};

/** The `\name` immediately before the caret that a typed space completes —
 *  `{ length, to }` (length = characters to replace, backslash included) or
 *  null. Only for text inside an equation; the caller checks the mark. */
export function mathAutoCorrectMatch(
  before: string,
): { length: number; to: string } | null {
  const m = /\\([A-Za-z]+)$/.exec(before);
  if (!m) return null;
  const to = MATH_AUTOCORRECT[m[1]];
  return to ? { length: m[0].length, to } : null;
}

// ── Equation AST (2D typesetting) ───────────────────────────────────
//
// The structured form of an equation — what OMML and MTEF both map onto.
// Deliberately small: every node is JSON-serializable (it rides a
// ProseMirror node attr), rows are plain arrays, and characters carry their
// letterform IN the character (math-italic applied when the tree is built),
// so a renderer never needs style context.

/** One character of an equation. The glyph is already letterformed. */
export interface EqChr {
  t: 'chr';
  ch: string;
}
/** A fraction: numerator over denominator. */
export interface EqFrac {
  t: 'frac';
  num: EqNode[];
  den: EqNode[];
}
/** A radical; `deg` is empty for a plain square root. */
export interface EqRad {
  t: 'rad';
  deg: EqNode[];
  body: EqNode[];
  /** Whether this radical HAS a degree, as OOXML's m:degHide records. Same
   *  reason EqScr carries `slots`: emptiness cannot tell a fresh nth-root
   *  template (waiting for its index) from a plain square root. Optional so
   *  equations stored before this still read correctly. */
  showDeg?: boolean;
}

/** Whether a radical shows a degree row — its own answer when it has one,
 *  read off emptiness otherwise. */
export function radShowDeg(n: EqRad): boolean {
  return n.showDeg ?? n.deg.length > 0;
}
/** Scripts on a base: either list may be empty. */
export interface EqScr {
  t: 'scr';
  base: EqNode[];
  sub: EqNode[];
  sup: EqNode[];
  /** Which scripts this node HAS, as OOXML distinguishes m:sSub / m:sSup /
   *  m:sSubSup. Without it the shape is read off emptiness, which is right
   *  for a finished equation but cannot express a FRESH one: a template with
   *  both rows still empty has to know whether it is waiting for a subscript,
   *  a superscript, or both. Optional so equations stored before this stay
   *  valid. */
  slots?: 'sub' | 'sup' | 'both';
}

/** Which script rows a node carries — its own answer when it has one, read
 *  off emptiness otherwise. */
export function scrSlots(n: EqScr): 'sub' | 'sup' | 'both' {
  if (n.slots) return n.slots;
  if (n.sup.length) return n.sub.length ? 'both' : 'sup';
  return 'sub';
}
/** A fenced group; `l`/`r` may be '' for a one-sided fence. */
export interface EqFence {
  t: 'fence';
  l: string;
  r: string;
  body: EqNode[];
}
/** A big operator (∑ ∫ ∏ …) with optional limits, then its operand. */
export interface EqBig {
  t: 'big';
  op: string;
  lo: EqNode[];
  hi: EqNode[];
  body: EqNode[];
  /** Whether this operator SHOWS a lower / upper limit, as OOXML's
   *  m:subHide / m:supHide record. Same reason EqScr carries `slots` and
   *  EqRad carries `showDeg`: a fresh "sum with limits" has both rows empty
   *  and still needs them, while a plain ∑ must not offer rows that are not
   *  there. Optional so equations stored before this still read. */
  showLo?: boolean;
  showHi?: boolean;
}

/** Which limit rows a big operator shows — its own answer when it has one,
 *  read off emptiness otherwise. */
export function bigLimits(n: EqBig): { lo: boolean; hi: boolean } {
  return {
    lo: n.showLo ?? n.lo.length > 0,
    hi: n.showHi ?? n.hi.length > 0,
  };
}

/** A named function applied to an argument (m:func): `sin 𝑥`. The name is a
 *  row of its own so `log` can carry a base and the caret can edit it. */
export interface EqFunc {
  t: 'func';
  name: EqNode[];
  body: EqNode[];
}

/** A mark over a base (m:acc) — a vector arrow, a bar, a hat. `chr` is the
 *  COMBINING codepoint OOXML stores; the layout paints a spacing lookalike,
 *  since a lone combining mark draws unreliably on a canvas. */
export interface EqAcc {
  t: 'acc';
  chr: string;
  body: EqNode[];
}

/** An operator with a limit written under (m:limLow) or over (m:limUpp) it —
 *  `lim` with `𝑛→∞` beneath, `max` over a set. */
export interface EqLim {
  t: 'lim';
  base: EqNode[];
  lim: EqNode[];
  /** The limit sits under the base (m:limLow). False puts it over. */
  below: boolean;
}

/** A grid (m:m). Cells are stored FLAT in reading order — cell (r, c) is at
 *  `r * cols + c` — because a slot path addresses one row of nodes at a time
 *  and a two-index address would not fit that scheme. */
export interface EqMat {
  t: 'mat';
  cols: number;
  cells: EqNode[][];
}

export type EqNode =
  | EqChr
  | EqFrac
  | EqRad
  | EqScr
  | EqFence
  | EqBig
  | EqFunc
  | EqAcc
  | EqLim
  | EqMat;

/** The row name addressing a matrix cell in a slot path. */
export const eqCellName = (index: number): string => `c${index}`;
/** The cell index a row name addresses, or null when it is not a cell. */
export function eqCellIndex(name: string): number | null {
  const m = /^c(\d+)$/.exec(name);
  return m ? Number(m[1]) : null;
}

/**
 * The rows of a node that form a VERTICAL stack, top to bottom — what up and
 * down step through inside that structure.
 *
 * Not every row belongs: a big operator stacks its two limits on the operator
 * sign, while the operand sits beside it, so down from the upper limit is the
 * LOWER limit even though the operand is nearer. Rows left out here are still
 * reachable by up and down; they are just found by geometry instead of by
 * this order.
 */
export function eqVerticalRows(n: EqNode): readonly string[] {
  switch (n.t) {
    case 'frac':
      return ['num', 'den'];
    case 'rad':
      return ['deg', 'body'];
    case 'scr':
      return ['sup', 'base', 'sub'];
    case 'big':
      return ['hi', 'lo'];
    case 'lim':
      return n.below ? ['base', 'lim'] : ['lim', 'base'];
    default:
      return [];
  }
}

/**
 * A node's editable rows in CARET order — the order the caret walks them
 * stepping right, which is reading order: a fraction's numerator before its
 * denominator, a radical's degree before its body.
 *
 * Not the same as the order the layout emits slot rects in (that is a drawing
 * order), and not derivable from the node's field order either. Callers
 * filter this against the rows the layout actually drew, so a script that
 * carries only a subscript never offers a superscript to step into.
 */
export function eqRowNames(n: EqNode): readonly string[] {
  switch (n.t) {
    case 'frac':
      return ['num', 'den'];
    case 'rad':
      return ['deg', 'body'];
    case 'scr':
      return ['base', 'sub', 'sup'];
    case 'fence':
      return ['body'];
    case 'big':
      return ['lo', 'hi', 'body'];
    case 'func':
      return ['name', 'body'];
    case 'acc':
      return ['body'];
    case 'lim':
      return ['base', 'lim'];
    case 'mat':
      // Reading order, row by row — the same order Tab walks a table.
      return n.cells.map((_, i) => eqCellName(i));
    default:
      return [];
  }
}

const SUB_DIGIT = '₀₁₂₃₄₅₆₇₈₉';
const SUP_DIGIT = '⁰¹²³⁴⁵⁶⁷⁸⁹';
const digitsVia = (s: string, alphabet: string): string | null =>
  /^[0-9]+$/.test(s) ? [...s].map((d) => alphabet[Number(d)]).join('') : null;

/** The linear spelling of an equation tree — the same shape the OMML
 *  flattener and Insert ▸ Equation type: it is the equation's plain-text
 *  identity (a11y mirror, search, the linear editing mode). */
export function astToLinear(row: EqNode[]): string {
  // Char counts, not UTF-16 lengths: a math-italic letter is a surrogate
  // pair and must still count as ONE character.
  const count = (s: string): number => [...s].length;
  const paren = (s: string): string =>
    count(s) <= 1 || (s.startsWith('(') && s.endsWith(')')) ? s : `(${s})`;
  const sub = (s: string): string =>
    digitsVia(s, SUB_DIGIT) ?? (count(s) > 1 ? `_(${s})` : `_${s}`);
  const sup = (s: string): string =>
    digitsVia(s, SUP_DIGIT) ?? (count(s) > 1 ? `^(${s})` : `^${s}`);
  const one = (n: EqNode): string => {
    switch (n.t) {
      case 'chr':
        return n.ch;
      case 'frac':
        return `${paren(astToLinear(n.num))}/${paren(astToLinear(n.den))}`;
      case 'rad': {
        const deg = astToLinear(n.deg);
        return `${deg ? sup(deg) : ''}√(${astToLinear(n.body)})`;
      }
      case 'scr': {
        const lo = astToLinear(n.sub);
        const hi = astToLinear(n.sup);
        return astToLinear(n.base) + (lo ? sub(lo) : '') + (hi ? sup(hi) : '');
      }
      case 'fence':
        return n.l + astToLinear(n.body) + n.r;
      case 'big': {
        const lo = astToLinear(n.lo);
        const hi = astToLinear(n.hi);
        return (
          n.op +
          (lo ? sub(lo) : '') +
          (hi ? sup(hi) : '') +
          paren(astToLinear(n.body))
        );
      }
      case 'func':
        return `${astToLinear(n.name)}\u2061${paren(astToLinear(n.body))}`;
      case 'acc':
        return astToLinear(n.body) + n.chr;
      case 'lim': {
        const l = astToLinear(n.lim);
        return astToLinear(n.base) + (l ? (n.below ? sub(l) : sup(l)) : '');
      }
      case 'mat': {
        // Word's linear matrix spelling: cells joined by &, rows by @.
        const lines: string[] = [];
        for (let i = 0; i < n.cells.length; i += n.cols)
          lines.push(
            n.cells
              .slice(i, i + n.cols)
              .map((c) => astToLinear(c))
              .join('&'),
          );
        return `(${lines.join('@')})`;
      }
    }
  };
  return row.map(one).join('');
}

/** A structural guard for AST data deserialized from a node attr — the attr
 *  is JSON from a document, not a trusted value. */
export function isEqRow(v: unknown): v is EqNode[] {
  if (!Array.isArray(v)) return false;
  return v.every((n) => {
    if (typeof n !== 'object' || n === null) return false;
    const t = (n as { t?: unknown }).t;
    if (t === 'chr') return typeof (n as EqChr).ch === 'string';
    if (t === 'frac')
      return isEqRow((n as EqFrac).num) && isEqRow((n as EqFrac).den);
    if (t === 'rad')
      return isEqRow((n as EqRad).deg) && isEqRow((n as EqRad).body);
    if (t === 'scr')
      return (
        isEqRow((n as EqScr).base) &&
        isEqRow((n as EqScr).sub) &&
        isEqRow((n as EqScr).sup)
      );
    if (t === 'fence')
      return (
        typeof (n as EqFence).l === 'string' &&
        typeof (n as EqFence).r === 'string' &&
        isEqRow((n as EqFence).body)
      );
    if (t === 'big')
      return (
        typeof (n as EqBig).op === 'string' &&
        isEqRow((n as EqBig).lo) &&
        isEqRow((n as EqBig).hi) &&
        isEqRow((n as EqBig).body)
      );
    if (t === 'func')
      return isEqRow((n as EqFunc).name) && isEqRow((n as EqFunc).body);
    if (t === 'acc')
      return typeof (n as EqAcc).chr === 'string' && isEqRow((n as EqAcc).body);
    if (t === 'lim')
      return (
        typeof (n as EqLim).below === 'boolean' &&
        isEqRow((n as EqLim).base) &&
        isEqRow((n as EqLim).lim)
      );
    if (t === 'mat') {
      const m = n as EqMat;
      return (
        typeof m.cols === 'number' &&
        m.cols > 0 &&
        Array.isArray(m.cells) &&
        m.cells.every((c) => isEqRow(c))
      );
    }
    return false;
  });
}
