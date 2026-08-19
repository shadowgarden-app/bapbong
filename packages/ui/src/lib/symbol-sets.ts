/**
 * The characters the Symbol dialog offers, and the pure logic around them
 * (recently-used list, code-point parsing) — kept free of the DOM so it is
 * unit-tested where the dialog itself, per repo convention, is checked in the
 * browser.
 *
 * Every entry is a real Unicode character, never a symbol-font PUA code: the
 * document gets `w:t` text that any Word renders (with its own font
 * substitution), and the editor draws it with the run's font backed by the
 * bundled symbol face (SYMBOL_FALLBACK_FAMILY). Curated rather than the whole
 * of Unicode — Word's dialog is a font grid, this is a task grid: the things
 * people reach for in a form, a report, a contract.
 */

export interface SymbolEntry {
  /** The character (one code point; may be astral). */
  char: string;
  /** Short English name, as the Unicode name reads in a tooltip. */
  name: string;
}

export interface SymbolGroup {
  id: string;
  label: string;
  entries: SymbolEntry[];
}

const g = (
  id: string,
  label: string,
  pairs: [string, string][],
): SymbolGroup => ({
  id,
  label,
  entries: pairs.map(([char, name]) => ({ char, name })),
});

/** The groups, in dialog order. Ids are stable (persist recents / tests). */
export const SYMBOL_GROUPS: readonly SymbolGroup[] = [
  g('checks', 'Checkboxes & ticks', [
    ['☐', 'Ballot box'],
    ['☑', 'Ballot box with check'],
    ['☒', 'Ballot box with X'],
    ['✓', 'Check mark'],
    ['✔', 'Heavy check mark'],
    ['✗', 'Ballot X'],
    ['✘', 'Heavy ballot X'],
    ['✕', 'Multiplication X'],
    ['✖', 'Heavy multiplication X'],
    ['❏', 'Lower right drop-shadowed white square'],
    ['❐', 'Upper right drop-shadowed white square'],
    ['❑', 'Lower right shadowed white square'],
    ['❒', 'Upper right shadowed white square'],
    ['▢', 'White square with rounded corners'],
    ['▣', 'White square containing black small square'],
    ['⊠', 'Squared times'],
    ['⊗', 'Circled times'],
    ['⦿', 'Circled bullet'],
    ['◉', 'Fisheye'],
    ['○', 'White circle'],
    ['●', 'Black circle'],
    ['◯', 'Large circle'],
    ['◻', 'White medium square'],
    ['◼', 'Black medium square'],
  ]),
  g('bullets', 'Bullets & shapes', [
    ['•', 'Bullet'],
    ['◦', 'White bullet'],
    ['▪', 'Black small square'],
    ['▫', 'White small square'],
    ['■', 'Black square'],
    ['□', 'White square'],
    ['▬', 'Black rectangle'],
    ['◆', 'Black diamond'],
    ['◇', 'White diamond'],
    ['◊', 'Lozenge'],
    ['★', 'Black star'],
    ['☆', 'White star'],
    ['✦', 'Black four pointed star'],
    ['✧', 'White four pointed star'],
    ['►', 'Black right-pointing pointer'],
    ['▸', 'Black right-pointing small triangle'],
    ['▹', 'White right-pointing small triangle'],
    ['▲', 'Black up-pointing triangle'],
    ['△', 'White up-pointing triangle'],
    ['▼', 'Black down-pointing triangle'],
    ['▽', 'White down-pointing triangle'],
    ['◄', 'Black left-pointing pointer'],
    ['❖', 'Black diamond minus white X'],
    ['✽', 'Heavy teardrop-spoked asterisk'],
    ['✿', 'Black florette'],
    ['❀', 'White florette'],
    ['➢', 'Three-D top-lighted rightwards arrowhead'],
    ['➤', 'Black rightwards arrowhead'],
    ['➣', 'Three-D bottom-lighted rightwards arrowhead'],
    ['✓', 'Check mark'],
  ]),
  g('arrows', 'Arrows', [
    ['←', 'Leftwards arrow'],
    ['↑', 'Upwards arrow'],
    ['→', 'Rightwards arrow'],
    ['↓', 'Downwards arrow'],
    ['↔', 'Left right arrow'],
    ['↕', 'Up down arrow'],
    ['↖', 'North west arrow'],
    ['↗', 'North east arrow'],
    ['↘', 'South east arrow'],
    ['↙', 'South west arrow'],
    ['⇐', 'Leftwards double arrow'],
    ['⇑', 'Upwards double arrow'],
    ['⇒', 'Rightwards double arrow'],
    ['⇓', 'Downwards double arrow'],
    ['⇔', 'Left right double arrow'],
    ['⇕', 'Up down double arrow'],
    ['↩', 'Leftwards arrow with hook'],
    ['↪', 'Rightwards arrow with hook'],
    ['↵', 'Downwards arrow with corner leftwards'],
    ['⇄', 'Rightwards arrow over leftwards arrow'],
    ['⇆', 'Leftwards arrow over rightwards arrow'],
    ['➔', 'Heavy wide-headed rightwards arrow'],
    ['➜', 'Heavy round-tipped rightwards arrow'],
    ['➝', 'Triangle-headed rightwards arrow'],
    ['➞', 'Heavy triangle-headed rightwards arrow'],
    ['➡', 'Black rightwards arrow'],
    ['⬅', 'Leftwards black arrow'],
    ['⬆', 'Upwards black arrow'],
    ['⬇', 'Downwards black arrow'],
    ['↻', 'Clockwise open circle arrow'],
    ['↺', 'Anticlockwise open circle arrow'],
    ['⤴', 'Arrow pointing rightwards then curving upwards'],
    ['⤵', 'Arrow pointing rightwards then curving downwards'],
  ]),
  g('math', 'Math', [
    ['±', 'Plus-minus sign'],
    ['×', 'Multiplication sign'],
    ['÷', 'Division sign'],
    ['−', 'Minus sign'],
    ['≤', 'Less-than or equal to'],
    ['≥', 'Greater-than or equal to'],
    ['≠', 'Not equal to'],
    ['≈', 'Almost equal to'],
    ['≡', 'Identical to'],
    ['∞', 'Infinity'],
    ['√', 'Square root'],
    ['∑', 'N-ary summation'],
    ['∏', 'N-ary product'],
    ['∫', 'Integral'],
    ['∂', 'Partial differential'],
    ['∆', 'Increment'],
    ['∇', 'Nabla'],
    ['∈', 'Element of'],
    ['∉', 'Not an element of'],
    ['∩', 'Intersection'],
    ['∪', 'Union'],
    ['⊂', 'Subset of'],
    ['⊃', 'Superset of'],
    ['∅', 'Empty set'],
    ['∀', 'For all'],
    ['∃', 'There exists'],
    ['¬', 'Not sign'],
    ['∧', 'Logical and'],
    ['∨', 'Logical or'],
    ['°', 'Degree sign'],
    ['′', 'Prime'],
    ['″', 'Double prime'],
    ['‰', 'Per mille sign'],
    ['µ', 'Micro sign'],
    ['π', 'Greek small letter pi'],
    ['∝', 'Proportional to'],
    ['∠', 'Angle'],
    ['⊥', 'Up tack'],
    ['∥', 'Parallel to'],
    ['¹', 'Superscript one'],
    ['²', 'Superscript two'],
    ['³', 'Superscript three'],
    ['½', 'Vulgar fraction one half'],
    ['⅓', 'Vulgar fraction one third'],
    ['¼', 'Vulgar fraction one quarter'],
    ['¾', 'Vulgar fraction three quarters'],
  ]),
  g('currency', 'Currency', [
    ['₫', 'Dong sign'],
    ['$', 'Dollar sign'],
    ['€', 'Euro sign'],
    ['£', 'Pound sign'],
    ['¥', 'Yen sign'],
    ['₩', 'Won sign'],
    ['₹', 'Indian rupee sign'],
    ['₽', 'Ruble sign'],
    ['฿', 'Thai currency symbol baht'],
    ['₱', 'Peso sign'],
    ['¢', 'Cent sign'],
    ['¤', 'Currency sign'],
    ['₿', 'Bitcoin sign'],
  ]),
  g('punct', 'Punctuation & typography', [
    ['–', 'En dash'],
    ['—', 'Em dash'],
    ['…', 'Horizontal ellipsis'],
    ['‘', 'Left single quotation mark'],
    ['’', 'Right single quotation mark'],
    ['“', 'Left double quotation mark'],
    ['”', 'Right double quotation mark'],
    ['‚', 'Single low-9 quotation mark'],
    ['„', 'Double low-9 quotation mark'],
    ['«', 'Left-pointing double angle quotation mark'],
    ['»', 'Right-pointing double angle quotation mark'],
    ['‹', 'Single left-pointing angle quotation mark'],
    ['›', 'Single right-pointing angle quotation mark'],
    ['§', 'Section sign'],
    ['¶', 'Pilcrow sign'],
    ['†', 'Dagger'],
    ['‡', 'Double dagger'],
    ['•', 'Bullet'],
    ['·', 'Middle dot'],
    ['№', 'Numero sign'],
    ['©', 'Copyright sign'],
    ['®', 'Registered sign'],
    ['™', 'Trade mark sign'],
    ['℠', 'Service mark'],
    ['¡', 'Inverted exclamation mark'],
    ['¿', 'Inverted question mark'],
    ['‽', 'Interrobang'],
    ['ˆ', 'Modifier letter circumflex accent'],
    ['˜', 'Small tilde'],
    ['¦', 'Broken bar'],
  ]),
  g('greek', 'Greek', [
    ['α', 'Alpha'],
    ['β', 'Beta'],
    ['γ', 'Gamma'],
    ['δ', 'Delta'],
    ['ε', 'Epsilon'],
    ['ζ', 'Zeta'],
    ['η', 'Eta'],
    ['θ', 'Theta'],
    ['ι', 'Iota'],
    ['κ', 'Kappa'],
    ['λ', 'Lambda'],
    ['μ', 'Mu'],
    ['ν', 'Nu'],
    ['ξ', 'Xi'],
    ['ο', 'Omicron'],
    ['π', 'Pi'],
    ['ρ', 'Rho'],
    ['σ', 'Sigma'],
    ['ς', 'Final sigma'],
    ['τ', 'Tau'],
    ['υ', 'Upsilon'],
    ['φ', 'Phi'],
    ['χ', 'Chi'],
    ['ψ', 'Psi'],
    ['ω', 'Omega'],
    ['Α', 'Capital alpha'],
    ['Β', 'Capital beta'],
    ['Γ', 'Capital gamma'],
    ['Δ', 'Capital delta'],
    ['Ε', 'Capital epsilon'],
    ['Ζ', 'Capital zeta'],
    ['Η', 'Capital eta'],
    ['Θ', 'Capital theta'],
    ['Ι', 'Capital iota'],
    ['Κ', 'Capital kappa'],
    ['Λ', 'Capital lambda'],
    ['Μ', 'Capital mu'],
    ['Ν', 'Capital nu'],
    ['Ξ', 'Capital xi'],
    ['Ο', 'Capital omicron'],
    ['Π', 'Capital pi'],
    ['Ρ', 'Capital rho'],
    ['Σ', 'Capital sigma'],
    ['Τ', 'Capital tau'],
    ['Υ', 'Capital upsilon'],
    ['Φ', 'Capital phi'],
    ['Χ', 'Capital chi'],
    ['Ψ', 'Capital psi'],
    ['Ω', 'Capital omega'],
  ]),
  g('latin', 'Latin (Vietnamese & accents)', [
    ['Đ', 'Capital D with stroke'],
    ['đ', 'Small d with stroke'],
    ['Ă', 'Capital A with breve'],
    ['ă', 'Small a with breve'],
    ['Â', 'Capital A with circumflex'],
    ['â', 'Small a with circumflex'],
    ['Ê', 'Capital E with circumflex'],
    ['ê', 'Small e with circumflex'],
    ['Ô', 'Capital O with circumflex'],
    ['ô', 'Small o with circumflex'],
    ['Ơ', 'Capital O with horn'],
    ['ơ', 'Small o with horn'],
    ['Ư', 'Capital U with horn'],
    ['ư', 'Small u with horn'],
    ['À', 'Capital A with grave'],
    ['à', 'Small a with grave'],
    ['Á', 'Capital A with acute'],
    ['á', 'Small a with acute'],
    ['Ã', 'Capital A with tilde'],
    ['ã', 'Small a with tilde'],
    ['Ả', 'Capital A with hook above'],
    ['ả', 'Small a with hook above'],
    ['Ạ', 'Capital A with dot below'],
    ['ạ', 'Small a with dot below'],
    ['È', 'Capital E with grave'],
    ['é', 'Small e with acute'],
    ['Ì', 'Capital I with grave'],
    ['í', 'Small i with acute'],
    ['Ò', 'Capital O with grave'],
    ['ó', 'Small o with acute'],
    ['Ù', 'Capital U with grave'],
    ['ú', 'Small u with acute'],
    ['Ý', 'Capital Y with acute'],
    ['ỹ', 'Small y with tilde'],
    ['Ç', 'Capital C with cedilla'],
    ['ç', 'Small c with cedilla'],
    ['Ñ', 'Capital N with tilde'],
    ['ñ', 'Small n with tilde'],
    ['Ö', 'Capital O with diaeresis'],
    ['ö', 'Small o with diaeresis'],
    ['Ü', 'Capital U with diaeresis'],
    ['ü', 'Small u with diaeresis'],
    ['ß', 'Small sharp s'],
    ['Æ', 'Capital AE'],
    ['æ', 'Small ae'],
    ['Ø', 'Capital O with stroke'],
    ['ø', 'Small o with stroke'],
    ['Œ', 'Capital ligature OE'],
    ['œ', 'Small ligature oe'],
  ]),
  g('misc', 'Misc', [
    ['☎', 'Black telephone'],
    ['✆', 'Telephone location sign'],
    ['✉', 'Envelope'],
    ['✂', 'Black scissors'],
    ['✎', 'Lower right pencil'],
    ['✏', 'Pencil'],
    ['✍', 'Writing hand'],
    ['☺', 'White smiling face'],
    ['☹', 'White frowning face'],
    ['♠', 'Black spade suit'],
    ['♣', 'Black club suit'],
    ['♥', 'Black heart suit'],
    ['♦', 'Black diamond suit'],
    ['♪', 'Eighth note'],
    ['♫', 'Beamed eighth notes'],
    ['☀', 'Black sun with rays'],
    ['☁', 'Cloud'],
    ['☂', 'Umbrella'],
    ['⚑', 'Black flag'],
    ['⚐', 'White flag'],
    ['⚠', 'Warning sign'],
    ['⛔', 'No entry'],
    ['♀', 'Female sign'],
    ['♂', 'Male sign'],
    ['⌘', 'Place of interest sign'],
    ['⌥', 'Option key'],
    ['⇧', 'Upwards white arrow'],
    ['⏎', 'Return symbol'],
    ['⌫', 'Erase to the left'],
    ['⏏', 'Eject symbol'],
    ['♻', 'Black universal recycling symbol'],
    ['✈', 'Airplane'],
    ['⚖', 'Scales'],
    ['⚕', 'Staff of aesculapius'],
  ]),
];

/** Word's "Special Characters" tab: typographic characters people cannot
 *  type, each with what to type instead where the editor offers a way (an
 *  AutoCorrect sequence or a shortcut) — informational until those ship. */
export interface SpecialCharacter extends SymbolEntry {
  /** How to get it without the dialog, shown beside the name (may be ''). */
  hint: string;
}
export const SPECIAL_CHARACTERS: readonly SpecialCharacter[] = [
  { char: '—', name: 'Em dash', hint: '' },
  { char: '–', name: 'En dash', hint: '' },
  { char: '‑', name: 'Nonbreaking hyphen', hint: '' },
  { char: '­', name: 'Optional hyphen', hint: '' },
  { char: ' ', name: 'Nonbreaking space', hint: '' },
  { char: ' ', name: 'Em space', hint: '' },
  { char: ' ', name: 'En space', hint: '' },
  { char: ' ', name: 'Thin space', hint: '' },
  { char: '©', name: 'Copyright', hint: '' },
  { char: '®', name: 'Registered', hint: '' },
  { char: '™', name: 'Trademark', hint: '' },
  { char: '§', name: 'Section', hint: '' },
  { char: '¶', name: 'Paragraph', hint: '' },
  { char: '…', name: 'Ellipsis', hint: '' },
  { char: '‘', name: 'Single opening quote', hint: '' },
  { char: '’', name: 'Single closing quote', hint: '' },
  { char: '“', name: 'Double opening quote', hint: '' },
  { char: '”', name: 'Double closing quote', hint: '' },
  { char: '₫', name: 'Dong sign', hint: '' },
  { char: '№', name: 'Numero sign', hint: '' },
];

/** Every entry the dialog knows, char → name (first group wins on repeats,
 *  which only happen where a character legitimately belongs to two groups —
 *  the bullet, the check mark). */
export const SYMBOL_NAMES: ReadonlyMap<string, string> = (() => {
  const m = new Map<string, string>();
  for (const grp of SYMBOL_GROUPS)
    for (const e of grp.entries) if (!m.has(e.char)) m.set(e.char, e.name);
  for (const e of SPECIAL_CHARACTERS) if (!m.has(e.char)) m.set(e.char, e.name);
  return m;
})();

/** `U+2611` for ☑ — the code point, upper-case hex, at least four digits. */
export function codePointLabel(char: string): string {
  const cp = char.codePointAt(0) ?? 0;
  return `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`;
}

/** The character a user's "Character code" input names: bare hex (`2611`),
 *  `U+2611`, `0x2611`; case-insensitive; surrounding space ignored. Null for
 *  anything else, for surrogates, and for code points beyond Unicode. */
export function parseCodePoint(input: string): string | null {
  const m = /^\s*(?:u\+|0x)?([0-9a-f]{1,6})\s*$/i.exec(input);
  if (!m) return null;
  const cp = parseInt(m[1], 16);
  if (cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) return null;
  // Control characters are not something to insert by number.
  if (cp < 0x20 || (cp >= 0x7f && cp < 0xa0)) return null;
  return String.fromCodePoint(cp);
}

/** Word keeps a row of the last sixteen. Most-recent first, no repeats. */
export const RECENT_MAX = 16;
export function pushRecent(
  recent: readonly string[],
  char: string,
  max = RECENT_MAX,
): string[] {
  return [char, ...recent.filter((c) => c !== char)].slice(0, max);
}

/** Case-insensitive name search across every group; the group order and each
 *  group's entry order are kept, so results read like the grid does. */
export function searchSymbols(query: string): SymbolEntry[] {
  const q = query.trim().toLowerCase();
  if (!q) return [];
  const seen = new Set<string>();
  // Ranked, then stable: a group whose LABEL matches comes first ("arrow"
  // lists the Arrows group before the arrowhead bullets), then names that
  // start with the query, then the rest — each in grid order.
  const ranked: { e: SymbolEntry; rank: number }[] = [];
  for (const grp of SYMBOL_GROUPS) {
    const groupHit = grp.label.toLowerCase().includes(q);
    for (const e of grp.entries) {
      const name = e.name.toLowerCase();
      if (seen.has(e.char) || !name.includes(q)) continue;
      seen.add(e.char);
      ranked.push({ e, rank: groupHit ? 0 : name.startsWith(q) ? 1 : 2 });
    }
  }
  return ranked.sort((a, b) => a.rank - b.rank).map((r) => r.e);
}
