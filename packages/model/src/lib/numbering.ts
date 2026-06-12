/**
 * Live list numbering. The document stores WHAT a paragraph is (numId +
 * level) and the numbering definitions (on the doc node's `numbering` attr);
 * marker strings ("1.", "2.a.", "•") are recomputed by walking the document —
 * the layout engine does this every pass, so inserting/deleting/reordering
 * list items renumbers everything, like Word.
 */

/** One w:lvl of an abstract numbering definition (plain data — attr-safe). */
export interface NumberingLevelDef {
  numFmt: string;
  lvlText: string;
  start: number;
}

/** Definitions keyed by numId. `key` groups numIds that share one abstract
 *  definition (their counters advance together, mirroring w:abstractNumId). */
export interface NumberingDefs {
  [numId: string]: {
    key: string;
    levels: Record<number, NumberingLevelDef>;
  };
}

/** Stateful counter: call `next` once per list paragraph, in document order. */
export interface NumberingCounter {
  next(numId: string, level: number): string;
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

export function createNumberingCounter(defs: NumberingDefs | null | undefined): NumberingCounter {
  // Running counters per shared definition key, indexed by level.
  const counters = new Map<string, (number | undefined)[]>();

  function next(numId: string, level: number): string {
    const def = defs?.[numId];
    if (!def) return '';
    const startOf = (lvl: number) => def.levels[lvl]?.start ?? 1;

    const arr = counters.get(def.key) ?? [];
    arr[level] = (arr[level] ?? startOf(level) - 1) + 1;
    for (let l = level + 1; l < arr.length; l++) arr[l] = undefined; // reset deeper levels
    counters.set(def.key, arr);

    const lvlDef = def.levels[level];
    if (!lvlDef || lvlDef.numFmt === 'none') return '';
    if (lvlDef.numFmt === 'bullet') return lvlDef.lvlText || '•';

    // Substitute %1..%9 with the formatted counter of that 1-based level.
    return lvlDef.lvlText.replace(/%(\d)/g, (_match, digit: string) => {
      const lvl = Number(digit) - 1;
      const value = arr[lvl] ?? startOf(lvl);
      return formatCounter(value, def.levels[lvl]?.numFmt ?? 'decimal');
    });
  }

  return { next };
}
