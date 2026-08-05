import type { NumberingDefs } from '@shadow-garden/bapbong-model';
import { audit } from './audit.js';
import { attrOf, child, children, OoxmlNode } from './ooxml.js';

/** Parsed view of `word/numbering.xml`. Markers are NOT computed at import —
 *  the defs ride the document (doc attr) and the layout engine recounts them
 *  every pass, so edits renumber live. */
export interface NumberingResolver {
  /** Plain-data definitions for the doc node's `numbering` attr. */
  defs: NumberingDefs | null;
  /** The lvl's paragraph properties (indent layer in the pPr cascade). */
  levelPPr(numId: string, level: number): OoxmlNode | undefined;
  /** XML-audit hook, call once after every story is parsed: the lvl pPr of
   *  levels NO paragraph referenced are marked as consumed subtrees — like
   *  an unused style, an unreferenced level can't lose this document's data
   *  (numbering.xml survives export as a carried part), so reporting its
   *  indent/tabs as UNKNOWN would only bury the real gaps. */
  auditMarkUnusedLevels(): void;
}

interface LevelEntry {
  numFmt: string;
  lvlText: string;
  start: number;
  pPr?: OoxmlNode;
}

function parseLvl(lvl: OoxmlNode): LevelEntry {
  return {
    numFmt: attrOf(child(lvl, 'w:numFmt'), 'w:val') ?? 'decimal',
    lvlText: attrOf(child(lvl, 'w:lvlText'), 'w:val') ?? '',
    start: Number(attrOf(child(lvl, 'w:start'), 'w:val') ?? '1') || 1,
    pPr: child(lvl, 'w:pPr'),
  };
}

export function buildNumbering(
  numberingRoot: OoxmlNode | undefined,
): NumberingResolver {
  const numberingEl = child(numberingRoot, 'w:numbering');

  const abstract = new Map<string, Map<number, LevelEntry>>();
  for (const abstractNum of children(numberingEl, 'w:abstractNum')) {
    const id = attrOf(abstractNum, 'w:abstractNumId');
    if (id === undefined) continue;
    const levels = new Map<number, LevelEntry>();
    for (const lvl of children(abstractNum, 'w:lvl')) {
      const ilvl = Number(attrOf(lvl, 'w:ilvl') ?? '0');
      levels.set(Number.isNaN(ilvl) ? 0 : ilvl, parseLvl(lvl));
    }
    abstract.set(id, levels);
  }

  // w:num maps a numId to its abstract definition, optionally with per-level
  // overrides: w:startOverride restarts the count (the "Restart numbering"
  // command), a full w:lvl child redefines the level outright.
  const numToAbstract = new Map<string, string>();
  const numOverrides = new Map<string, Map<number, LevelEntry>>();
  for (const num of children(numberingEl, 'w:num')) {
    const numId = attrOf(num, 'w:numId');
    const absId = attrOf(child(num, 'w:abstractNumId'), 'w:val');
    if (numId === undefined || absId === undefined) continue;
    numToAbstract.set(numId, absId);
    for (const o of children(num, 'w:lvlOverride')) {
      const ilvl = Number(attrOf(o, 'w:ilvl') ?? '0') || 0;
      const startOverride = attrOf(child(o, 'w:startOverride'), 'w:val');
      const fullLvl = child(o, 'w:lvl');
      const base = fullLvl
        ? parseLvl(fullLvl)
        : abstract.get(absId)?.get(ilvl);
      if (!base) continue;
      const entry: LevelEntry = { ...base };
      if (startOverride !== undefined)
        entry.start = Number(startOverride) || 1;
      let ov = numOverrides.get(numId);
      if (!ov) numOverrides.set(numId, (ov = new Map()));
      ov.set(ilvl, entry);
    }
  }

  let defs: NumberingDefs | null = null;
  for (const [numId, absId] of numToAbstract) {
    const levels = abstract.get(absId);
    if (!levels) continue;
    const ov = numOverrides.get(numId);
    const plain: Record<number, { numFmt: string; lvlText: string; start: number }> = {};
    for (const [ilvl, def] of levels) {
      const eff = ov?.get(ilvl) ?? def;
      plain[ilvl] = { numFmt: eff.numFmt, lvlText: eff.lvlText, start: eff.start };
    }
    for (const [ilvl, def] of ov ?? []) {
      // Overrides for levels the abstract never declared.
      plain[ilvl] ??= { numFmt: def.numFmt, lvlText: def.lvlText, start: def.start };
    }
    // An overridden num counts independently from its siblings on the same
    // abstract (that is what startOverride is FOR) — give it its own key.
    (defs ??= {})[numId] = { key: ov ? `${absId}#${numId}` : absId, levels: plain };
  }

  const usedPPr = new Set<OoxmlNode>();
  function levelPPr(numId: string, level: number): OoxmlNode | undefined {
    const pPr =
      numOverrides.get(numId)?.get(level)?.pPr ??
      (() => {
        const absId = numToAbstract.get(numId);
        return absId === undefined
          ? undefined
          : abstract.get(absId)?.get(level)?.pPr;
      })();
    if (pPr) usedPPr.add(pPr);
    return pPr;
  }

  return {
    defs,
    levelPPr,
    auditMarkUnusedLevels: () => {
      if (!audit.enabled) return;
      for (const levels of abstract.values())
        for (const entry of levels.values())
          if (entry.pPr && !usedPPr.has(entry.pPr)) audit.markSubtree(entry.pPr);
      for (const levels of numOverrides.values())
        for (const entry of levels.values())
          if (entry.pPr && !usedPPr.has(entry.pPr)) audit.markSubtree(entry.pPr);
    },
  };
}
