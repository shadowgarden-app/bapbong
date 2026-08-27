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
  /** The slot of the new node that takes the caret — a row name on `node`. */
  focus: string;
}

const frac = (): EqNode => ({ t: 'frac', num: [], den: [] });
const scr = (slots: 'sub' | 'sup' | 'both'): EqNode => ({
  t: 'scr',
  base: [],
  sub: [],
  sup: [],
  slots,
});
const rad = (showDeg: boolean): EqNode => ({
  t: 'rad',
  deg: [],
  body: [],
  showDeg,
});
const fence = (l: string, r: string): EqNode => ({
  t: 'fence',
  l,
  r,
  body: [],
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
  { group: 'Fraction', name: 'Stacked', node: frac(), focus: 'num' },

  { group: 'Script', name: 'Superscript', node: scr('sup'), focus: 'sup' },
  { group: 'Script', name: 'Subscript', node: scr('sub'), focus: 'sub' },
  { group: 'Script', name: 'Sub and super', node: scr('both'), focus: 'sub' },

  { group: 'Radical', name: 'Square root', node: rad(false), focus: 'body' },
  { group: 'Radical', name: 'Nth root', node: rad(true), focus: 'deg' },

  { group: 'Integral', name: 'Integral', node: big('∫'), focus: 'body' },
  { group: 'Integral', name: 'Definite', node: big('∫', true), focus: 'lo' },
  { group: 'Integral', name: 'Double', node: big('∬'), focus: 'body' },
  { group: 'Integral', name: 'Contour', node: big('∮'), focus: 'body' },

  { group: 'Large operator', name: 'Summation', node: big('∑'), focus: 'body' },
  {
    group: 'Large operator',
    name: 'Sum with limits',
    node: big('∑', true),
    focus: 'lo',
  },
  { group: 'Large operator', name: 'Product', node: big('∏'), focus: 'body' },
  {
    group: 'Large operator',
    name: 'Product with limits',
    node: big('∏', true),
    focus: 'lo',
  },
  { group: 'Large operator', name: 'Union', node: big('⋃'), focus: 'body' },
  {
    group: 'Large operator',
    name: 'Intersection',
    node: big('⋂'),
    focus: 'body',
  },

  {
    group: 'Bracket',
    name: 'Parentheses',
    node: fence('(', ')'),
    focus: 'body',
  },
  { group: 'Bracket', name: 'Brackets', node: fence('[', ']'), focus: 'body' },
  { group: 'Bracket', name: 'Braces', node: fence('{', '}'), focus: 'body' },
  {
    group: 'Bracket',
    name: 'Absolute value',
    node: fence('|', '|'),
    focus: 'body',
  },
  { group: 'Bracket', name: 'Norm', node: fence('‖', '‖'), focus: 'body' },
  { group: 'Bracket', name: 'Angle', node: fence('⟨', '⟩'), focus: 'body' },
  { group: 'Bracket', name: 'Floor', node: fence('⌊', '⌋'), focus: 'body' },
  { group: 'Bracket', name: 'Ceiling', node: fence('⌈', '⌉'), focus: 'body' },
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
