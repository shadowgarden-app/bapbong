import { attrOf, child, children, OoxmlNode } from './ooxml';

interface LevelDef {
  numFmt: string;
  lvlText: string;
  start: number;
  /** The level's w:pPr — carries the per-level indent (w:ind) that gives
   *  nested lists their staircase. */
  pPr?: OoxmlNode;
}

/** Stateful numbering counter built from `word/numbering.xml`. Call `next`
 *  once per list paragraph, in document order, to get its marker string. */
export interface NumberingResolver {
  next(numId: string, level: number): string;
  /** The lvl's paragraph properties (indent layer in the pPr cascade). */
  levelPPr(numId: string, level: number): OoxmlNode | undefined;
}

function toLetters(n: number): string {
  // bijective base-26: 1->A, 26->Z, 27->AA
  let s = '';
  let x = n;
  while (x > 0) {
    const rem = (x - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s || 'A';
}

function toRoman(n: number): string {
  if (n <= 0) return String(n);
  const table: [number, string][] = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let x = n;
  let out = '';
  for (const [v, sym] of table) {
    while (x >= v) {
      out += sym;
      x -= v;
    }
  }
  return out;
}

function formatCounter(n: number, fmt: string): string {
  switch (fmt) {
    case 'lowerLetter':
      return toLetters(n).toLowerCase();
    case 'upperLetter':
      return toLetters(n);
    case 'lowerRoman':
      return toRoman(n).toLowerCase();
    case 'upperRoman':
      return toRoman(n);
    default: // decimal, decimalZero, and unknown formats
      return String(n);
  }
}

export function buildNumbering(numberingRoot: OoxmlNode | undefined): NumberingResolver {
  const numberingEl = child(numberingRoot, 'w:numbering');

  const abstract = new Map<string, Map<number, LevelDef>>();
  for (const abstractNum of children(numberingEl, 'w:abstractNum')) {
    const id = attrOf(abstractNum, 'w:abstractNumId');
    if (id === undefined) continue;
    const levels = new Map<number, LevelDef>();
    for (const lvl of children(abstractNum, 'w:lvl')) {
      const ilvl = Number(attrOf(lvl, 'w:ilvl') ?? '0');
      levels.set(Number.isNaN(ilvl) ? 0 : ilvl, {
        numFmt: attrOf(child(lvl, 'w:numFmt'), 'w:val') ?? 'decimal',
        lvlText: attrOf(child(lvl, 'w:lvlText'), 'w:val') ?? '',
        start: Number(attrOf(child(lvl, 'w:start'), 'w:val') ?? '1') || 1,
        pPr: child(lvl, 'w:pPr'),
      });
    }
    abstract.set(id, levels);
  }

  const numToAbstract = new Map<string, string>();
  for (const num of children(numberingEl, 'w:num')) {
    const numId = attrOf(num, 'w:numId');
    const absId = attrOf(child(num, 'w:abstractNumId'), 'w:val');
    if (numId !== undefined && absId !== undefined) numToAbstract.set(numId, absId);
  }

  // Running counters per abstractNumId, indexed by level (undefined = unstarted).
  const counters = new Map<string, (number | undefined)[]>();
  const startOf = (absId: string, level: number) => abstract.get(absId)?.get(level)?.start ?? 1;

  function next(numId: string, level: number): string {
    const absId = numToAbstract.get(numId);
    const levels = absId === undefined ? undefined : abstract.get(absId);
    if (absId === undefined || !levels) return '';

    const arr = counters.get(absId) ?? [];
    arr[level] = (arr[level] ?? startOf(absId, level) - 1) + 1;
    for (let l = level + 1; l < arr.length; l++) arr[l] = undefined; // reset deeper levels
    counters.set(absId, arr);

    const def = levels.get(level);
    if (!def || def.numFmt === 'none') return '';
    if (def.numFmt === 'bullet') return def.lvlText || '•';

    // Substitute %1..%9 with the formatted counter of that 1-based level.
    return def.lvlText.replace(/%(\d)/g, (_match, digit: string) => {
      const lvl = Number(digit) - 1;
      const value = arr[lvl] ?? startOf(absId, lvl);
      return formatCounter(value, levels.get(lvl)?.numFmt ?? 'decimal');
    });
  }

  function levelPPr(numId: string, level: number): OoxmlNode | undefined {
    const absId = numToAbstract.get(numId);
    return absId === undefined ? undefined : abstract.get(absId)?.get(level)?.pPr;
  }

  return { next, levelPPr };
}
