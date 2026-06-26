import type { NumberingDefs } from '@shadow-garden/bapbong-model';
import { attrOf, child, children, OoxmlNode } from './ooxml.js';

/** Parsed view of `word/numbering.xml`. Markers are NOT computed at import —
 *  the defs ride the document (doc attr) and the layout engine recounts them
 *  every pass, so edits renumber live. */
export interface NumberingResolver {
  /** Plain-data definitions for the doc node's `numbering` attr. */
  defs: NumberingDefs | null;
  /** The lvl's paragraph properties (indent layer in the pPr cascade). */
  levelPPr(numId: string, level: number): OoxmlNode | undefined;
}

export function buildNumbering(numberingRoot: OoxmlNode | undefined): NumberingResolver {
  const numberingEl = child(numberingRoot, 'w:numbering');

  interface LevelEntry {
    numFmt: string;
    lvlText: string;
    start: number;
    pPr?: OoxmlNode;
  }
  const abstract = new Map<string, Map<number, LevelEntry>>();
  for (const abstractNum of children(numberingEl, 'w:abstractNum')) {
    const id = attrOf(abstractNum, 'w:abstractNumId');
    if (id === undefined) continue;
    const levels = new Map<number, LevelEntry>();
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

  let defs: NumberingDefs | null = null;
  for (const [numId, absId] of numToAbstract) {
    const levels = abstract.get(absId);
    if (!levels) continue;
    const plain: Record<number, { numFmt: string; lvlText: string; start: number }> = {};
    for (const [ilvl, def] of levels) {
      plain[ilvl] = { numFmt: def.numFmt, lvlText: def.lvlText, start: def.start };
    }
    (defs ??= {})[numId] = { key: absId, levels: plain };
  }

  function levelPPr(numId: string, level: number): OoxmlNode | undefined {
    const absId = numToAbstract.get(numId);
    return absId === undefined ? undefined : abstract.get(absId)?.get(level)?.pPr;
  }

  return { defs, levelPPr };
}
