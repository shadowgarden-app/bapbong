/**
 * What the equation palette offers: the symbol sets Word groups its Symbols
 * gallery into, and the structures this engine can actually draw.
 *
 * Data only — no DOM, no layout. The panel (bapbong-ui) renders it and the
 * slot editor (bapbong-editor) inserts from it, so both agree on one list.
 */
import type { EqNode } from './math.js';

/** One named page of the symbol gallery. */
export interface MathSymbolSet {
  name: string;
  /** The glyphs, in reading order. Iterate with `[...chars]` — several are
   *  outside the BMP, so `.length` and `charAt` would split them. */
  chars: string;
}

/**
 * Word's eight sets, trimmed to glyphs that render in the maths faces we
 * ship. Exotic codepoints (combining negations, half-supported geometry
 * marks) are deliberately absent: a tofu box in a palette teaches the user
 * the character is broken, which is worse than not offering it.
 */
export const MATH_SYMBOL_SETS: readonly MathSymbolSet[] = [
  {
    name: 'Basic math',
    chars: '±∞=≠≈≡∼≅∝<>≤≥∓+−×÷·∗∘!∀∃∄∈∉⊂⊃⊆⊇∪∩∅∧∨¬∴∵√∛∜∫∑∏∂∇°%‰′″…⋯',
  },
  {
    name: 'Greek letters',
    chars: 'αβγδεζηθικλμνξοπρστυφχψωϑϕϖϱςΓΔΘΛΞΠΣΥΦΨΩ',
  },
  {
    name: 'Letter-like symbols',
    chars: 'ℂℍℕℙℚℝℤℵℶℷℸℊℋℐℒℓ℘ℛℯℰℱℳℴℏÅ℧∁∅∞',
  },
  {
    name: 'Operators',
    chars: '∀∁∂∃∄∅∆∇∈∉∋∌∏∐∑∓∔∖∗∘∙√∛∜∝∞∠∡∢∣∤∥∦∧∨∩∪∫∬∭∮∯∰∴∵∶∷',
  },
  {
    name: 'Arrows',
    chars: '←↑→↓↔↕↖↗↘↙↚↛↞↠↢↣↦↩↪↭↮⇐⇑⇒⇓⇔⇕⇖⇗⇘⇙⇚⇛⇤⇥⟵⟶⟷⟸⟹⟺',
  },
  {
    name: 'Negated relations',
    chars: '≠≁≄≇≉≢∤∦≮≯≰≱⊀⊁⊄⊅⊈⊉⊬⊭⊮⊯∉∌⋠⋡⋢⋣⋪⋫⋬⋭',
  },
  {
    name: 'Scripts',
    chars: '𝒜ℬ𝒞𝒟ℰℱ𝒢ℋℐ𝒥𝒦ℒℳ𝒩𝒪𝒫𝒬ℛ𝒮𝒯𝒰𝒱𝒲𝒳𝒴𝒵𝔄𝔅ℭ𝔇𝔈𝔉𝔊ℌℑ𝔍𝔎𝔏𝔐𝔑𝔒𝔓𝔔ℜ𝔖𝔗𝔘𝔙𝔚𝔛𝔜ℨ',
  },
  {
    name: 'Geometry',
    chars: '∠∡∢∟⊥∥∦△▱▭○◊⌒°′″≅∼∴∵→↔⇒⇔',
  },
];

/** A template the Structures tab inserts, with every slot left empty. */
export interface EqStructure {
  /** Heading this template sits under in the panel. */
  group: string;
  /** Label under its preview. */
  name: string;
  /** The node to insert. Every row inside is empty, so the layout draws its
   *  dotted placeholder boxes and the user sees where to type. */
  node: EqNode;
  /** Where the caret lands inside the new node: the path from it to the row
   *  that takes the caret. Usually one row name, but a template can nest —
   *  a bracketed matrix puts the caret in the grid's first cell, two levels
   *  down from the fence that was inserted. */
  focus: readonly (string | number)[];
}

const frac = (): EqNode => ({ t: 'frac', num: [], den: [] });
const scr = (slots: 'sub' | 'sup' | 'both'): EqNode => ({
  t: 'scr',
  base: [],
  sub: [],
  sup: [],
  slots,
});
const rad = (deg: string | null): EqNode => ({
  t: 'rad',
  // A degree the template fills in — the 3 of a cube root — is content like
  // any other: the caret can reach it and type over it.
  deg: deg ? upright(deg) : [],
  body: [],
  showDeg: deg !== null,
});
const fence = (l: string, r: string): EqNode => ({
  t: 'fence',
  l,
  r,
  body: [],
});
/** Plain, upright characters — a function name is not a variable, so it must
 *  not be letterformed into math italic the way typed letters are. */
const upright = (s: string): EqNode[] =>
  [...s].map((ch) => ({ t: 'chr', ch }) as EqNode);
const func = (name: string): EqNode => ({
  t: 'func',
  name: upright(name),
  body: [],
});
const acc = (chr: string): EqNode => ({ t: 'acc', chr, body: [] });
/** A function whose NAME carries a subscript — log with a base. */
const funcSub = (name: string): EqNode => ({
  t: 'func',
  name: [{ t: 'scr', base: upright(name), sub: [], sup: [], slots: 'sub' }],
  body: [],
});
const mat = (rows: number, cols: number): EqNode => ({
  t: 'mat',
  cols,
  cells: Array.from({ length: rows * cols }, () => [] as EqNode[]),
});
/** A grid inside a fence — a bracketed matrix, or the brace a system of
 *  equations is written with. */
const braced = (l: string, r: string, rows: number, cols: number): EqNode => ({
  t: 'fence',
  l,
  r,
  body: [mat(rows, cols)],
});
const lim = (name: string, below = true): EqNode => ({
  t: 'lim',
  base: upright(name),
  lim: [],
  below,
});

const big = (op: string, limits = false): EqNode => ({
  t: 'big',
  op,
  lo: [],
  hi: [],
  body: [],
  showLo: limits,
  showHi: limits,
});

/**
 * The structures the palette offers.
 *
 * This is every shape the AST has a node for. Word's gallery also carries
 * Function, Accent, Limit-and-log, Operator and Matrix — each needs a node
 * kind that does not exist yet (an upright function name with limits, a mark
 * over a base, a grid), so offering them here would mean offering something
 * the layout cannot draw.
 */
export const EQ_STRUCTURES: readonly EqStructure[] = [
  { group: 'Fraction', name: 'Stacked', node: frac(), focus: ['num'] },

  { group: 'Script', name: 'Superscript', node: scr('sup'), focus: ['base'] },
  { group: 'Script', name: 'Subscript', node: scr('sub'), focus: ['base'] },
  {
    group: 'Script',
    name: 'Sub and super',
    node: scr('both'),
    focus: ['base'],
  },

  { group: 'Radical', name: 'Square root', node: rad(null), focus: ['body'] },
  { group: 'Radical', name: 'Cube root', node: rad('3'), focus: ['body'] },
  { group: 'Radical', name: 'Nth root', node: rad(''), focus: ['deg'] },

  { group: 'Integral', name: 'Integral', node: big('∫'), focus: ['body'] },
  { group: 'Integral', name: 'Definite', node: big('∫', true), focus: ['lo'] },
  { group: 'Integral', name: 'Double', node: big('∬'), focus: ['body'] },
  { group: 'Integral', name: 'Contour', node: big('∮'), focus: ['body'] },

  {
    group: 'Large operator',
    name: 'Summation',
    node: big('∑'),
    focus: ['body'],
  },
  {
    group: 'Large operator',
    name: 'Sum with limits',
    node: big('∑', true),
    focus: ['lo'],
  },
  { group: 'Large operator', name: 'Product', node: big('∏'), focus: ['body'] },
  {
    group: 'Large operator',
    name: 'Product with limits',
    node: big('∏', true),
    focus: ['lo'],
  },
  { group: 'Large operator', name: 'Union', node: big('⋃'), focus: ['body'] },
  {
    group: 'Large operator',
    name: 'Intersection',
    node: big('⋂'),
    focus: ['body'],
  },

  {
    group: 'Bracket',
    name: 'Parentheses',
    node: fence('(', ')'),
    focus: ['body'],
  },
  {
    group: 'Bracket',
    name: 'Brackets',
    node: fence('[', ']'),
    focus: ['body'],
  },
  { group: 'Bracket', name: 'Braces', node: fence('{', '}'), focus: ['body'] },
  {
    group: 'Bracket',
    name: 'Absolute value',
    node: fence('|', '|'),
    focus: ['body'],
  },
  { group: 'Bracket', name: 'Norm', node: fence('‖', '‖'), focus: ['body'] },
  { group: 'Bracket', name: 'Angle', node: fence('⟨', '⟩'), focus: ['body'] },
  { group: 'Bracket', name: 'Floor', node: fence('⌊', '⌋'), focus: ['body'] },
  { group: 'Bracket', name: 'Ceiling', node: fence('⌈', '⌉'), focus: ['body'] },

  { group: 'Function', name: 'sin', node: func('sin'), focus: ['body'] },
  { group: 'Function', name: 'cos', node: func('cos'), focus: ['body'] },
  { group: 'Function', name: 'tan', node: func('tan'), focus: ['body'] },
  { group: 'Function', name: 'cot', node: func('cot'), focus: ['body'] },
  { group: 'Function', name: 'sec', node: func('sec'), focus: ['body'] },
  { group: 'Function', name: 'csc', node: func('csc'), focus: ['body'] },
  { group: 'Function', name: 'arcsin', node: func('arcsin'), focus: ['body'] },
  { group: 'Function', name: 'arccos', node: func('arccos'), focus: ['body'] },
  { group: 'Function', name: 'arctan', node: func('arctan'), focus: ['body'] },
  { group: 'Function', name: 'sinh', node: func('sinh'), focus: ['body'] },
  { group: 'Function', name: 'cosh', node: func('cosh'), focus: ['body'] },
  { group: 'Function', name: 'tanh', node: func('tanh'), focus: ['body'] },
  { group: 'Function', name: 'ln', node: func('ln'), focus: ['body'] },
  { group: 'Function', name: 'log', node: func('log'), focus: ['body'] },
  // The base rides the NAME, not the argument, and the caret opens on it —
  // it is the part you came to this entry to type.
  {
    group: 'Function',
    name: 'log base',
    node: funcSub('log'),
    focus: ['name', 0, 'sub'],
  },
  { group: 'Function', name: 'exp', node: func('exp'), focus: ['body'] },
  // The name is empty and editable — for anything the list does not carry.
  { group: 'Function', name: 'Custom', node: func(''), focus: ['name'] },

  { group: 'Accent', name: 'Bar', node: acc('\u0305'), focus: ['body'] },
  { group: 'Accent', name: 'Vector', node: acc('\u20d7'), focus: ['body'] },
  { group: 'Accent', name: 'Hat', node: acc('\u0302'), focus: ['body'] },
  { group: 'Accent', name: 'Tilde', node: acc('\u0303'), focus: ['body'] },
  { group: 'Accent', name: 'Dot', node: acc('\u0307'), focus: ['body'] },
  { group: 'Accent', name: 'Double dot', node: acc('\u0308'), focus: ['body'] },

  { group: 'Limit', name: 'Limit', node: lim('lim'), focus: ['lim'] },
  { group: 'Limit', name: 'Maximum', node: lim('max'), focus: ['lim'] },
  { group: 'Limit', name: 'Minimum', node: lim('min'), focus: ['lim'] },
  { group: 'Limit', name: 'Over', node: lim('', false), focus: ['base'] },

  { group: 'Matrix', name: '1×2', node: mat(1, 2), focus: ['c0'] },
  { group: 'Matrix', name: '2×1', node: mat(2, 1), focus: ['c0'] },
  { group: 'Matrix', name: '2×2', node: mat(2, 2), focus: ['c0'] },
  { group: 'Matrix', name: '3×3', node: mat(3, 3), focus: ['c0'] },
  {
    group: 'Matrix',
    name: 'Bracketed 2×2',
    node: braced('[', ']', 2, 2),
    focus: ['body', 0, 'c0'],
  },
  {
    group: 'Matrix',
    name: 'Determinant 2×2',
    node: braced('|', '|', 2, 2),
    focus: ['body', 0, 'c0'],
  },
  // Two and three equations under one brace — how a system is set in
  // Vietnamese exam papers, and awkward to build out of the pieces.
  {
    group: 'Matrix',
    name: 'System of 2',
    node: braced('{', '', 2, 1),
    focus: ['body', 0, 'c0'],
  },
  {
    group: 'Matrix',
    name: 'System of 3',
    node: braced('{', '', 3, 1),
    focus: ['body', 0, 'c0'],
  },
];

/** The groups in palette order, each with its templates. */
export function eqStructureGroups(): { name: string; items: EqStructure[] }[] {
  const out: { name: string; items: EqStructure[] }[] = [];
  for (const s of EQ_STRUCTURES) {
    const last = out[out.length - 1];
    if (last && last.name === s.group) last.items.push(s);
    else out.push({ name: s.group, items: [s] });
  }
  return out;
}
