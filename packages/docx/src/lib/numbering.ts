import type {
  MarkerRunProps,
  NumberingDefs,
  NumberingLevelDef,
} from '@shadow-garden/bapbong-model';
import { audit } from './audit.js';
import { symbolFontText } from './symbol-fonts.js';
import {
  attrOf,
  child,
  children,
  OoxmlNode,
  parseRunProps,
  ThemeColorResolver,
  ThemeFontResolver,
} from './ooxml.js';

/** Parsed view of `word/numbering.xml`. Markers are NOT computed at import —
 *  the defs ride the document (doc attr) and the layout engine recounts them
 *  every pass, so edits renumber live. */
export interface NumberingResolver {
  /** Plain-data definitions for the doc node's `numbering` attr. */
  defs: NumberingDefs | null;
  /** The lvl's paragraph properties (indent layer in the pPr cascade). */
  levelPPr(numId: string, level: number): OoxmlNode | undefined;
  /** The level a paragraph STYLE is linked to (lvl > w:pStyle) inside the
   *  numId's definition — numbered heading styles pick their level this
   *  way, with only a numId (no ilvl) in the style's numPr. */
  levelForStyle(numId: string, styleId: string): number | undefined;
  /** XML-audit hook, call once after every story is parsed: the lvl pPr of
   *  levels NO paragraph referenced are marked as consumed subtrees — like
   *  an unused style, an unreferenced level can't lose this document's data
   *  (numbering.xml survives export as a carried part), so reporting its
   *  indent/tabs as UNKNOWN would only bury the real gaps. */
  auditMarkUnusedLevels(): void;
}

/** A parsed w:lvl: the plain-data level def plus its pPr (cascade layer). */
type LevelEntry = NumberingLevelDef & { pPr?: OoxmlNode };

/** Strip the OoxmlNode so the entry is attr-safe plain data. */
function plainDef(entry: LevelEntry): NumberingLevelDef {
  const { pPr: _pPr, ...plain } = entry;
  return plain;
}

function parseLvl(
  lvl: OoxmlNode,
  resolveTheme?: ThemeColorResolver,
  resolveFont?: ThemeFontResolver,
): LevelEntry {
  const entry: LevelEntry = {
    numFmt: attrOf(child(lvl, 'w:numFmt'), 'w:val') ?? 'decimal',
    lvlText: attrOf(child(lvl, 'w:lvlText'), 'w:val') ?? '',
    start: Number(attrOf(child(lvl, 'w:start'), 'w:val') ?? '1') || 1,
    pPr: child(lvl, 'w:pPr'),
  };
  // A bullet's glyph is a PICTURE in a symbol font: Word's default bullet is
  // U+F0B7 in Symbol, the second level U+F0A7 in Wingdings. Those code points
  // are private-use — without the font installed they render as tofu, which
  // is exactly what every bulleted list came out as. Translate the label to
  // real Unicode here, the same way run text and w:sym already are (the
  // level's own w:rPr names the font). The marker's family stays on the
  // level so an untranslated glyph still uses it, and export writes the
  // original XML back from the carried numbering part.
  const lvlFont = attrOf(child(child(lvl, 'w:rPr'), 'w:rFonts'), 'w:ascii');
  const mapped = symbolFontText(entry.lvlText, lvlFont);
  if (mapped !== null) entry.lvlText = mapped;
  // Label alignment (w:lvlJc). left/start is the default — omitted.
  const jc = attrOf(child(lvl, 'w:lvlJc'), 'w:val');
  if (jc === 'right' || jc === 'end') entry.jc = 'right';
  else if (jc === 'center') entry.jc = 'center';
  // Label→text separator (w:suff). tab is the default — omitted.
  const suff = attrOf(child(lvl, 'w:suff'), 'w:val');
  if (suff === 'space' || suff === 'nothing') entry.suff = suff;
  // Legal numbering (w:isLgl).
  const isLgl = child(lvl, 'w:isLgl');
  if (isLgl && attrOf(isLgl, 'w:val') !== '0') entry.isLgl = true;
  // Label formatting (w:lvl > w:rPr) — the number/bullet's own font.
  const rPrEl = child(lvl, 'w:rPr');
  if (rPrEl) {
    const props = parseRunProps(rPrEl, resolveTheme, resolveFont);
    const rPr: MarkerRunProps = {
      ...(props.bold !== undefined && { bold: props.bold }),
      ...(props.italic !== undefined && { italic: props.italic }),
      ...(props.sizePt !== undefined && { sizePt: props.sizePt }),
      ...(props.fontFamily !== undefined && { family: props.fontFamily }),
      ...(props.color !== undefined && { color: props.color }),
    };
    if (Object.keys(rPr).length > 0) entry.rPr = rPr;
    // Underline/strike/highlight on labels are not painted — a decision;
    // the properties survive in the carried numbering.xml either way.
    audit.markSubtree(rPrEl);
  }
  return entry;
}

export function buildNumbering(
  numberingRoot: OoxmlNode | undefined,
  resolveTheme?: ThemeColorResolver,
  resolveFont?: ThemeFontResolver,
): NumberingResolver {
  const numberingEl = child(numberingRoot, 'w:numbering');

  const abstract = new Map<string, Map<number, LevelEntry>>();
  // lvl > w:pStyle links: styleId → ilvl, per abstract definition.
  const styleLinks = new Map<string, Map<string, number>>();
  for (const abstractNum of children(numberingEl, 'w:abstractNum')) {
    const id = attrOf(abstractNum, 'w:abstractNumId');
    if (id === undefined) continue;
    const levels = new Map<number, LevelEntry>();
    const links = new Map<string, number>();
    for (const lvl of children(abstractNum, 'w:lvl')) {
      const rawIlvl = Number(attrOf(lvl, 'w:ilvl') ?? '0');
      const ilvl = Number.isNaN(rawIlvl) ? 0 : rawIlvl;
      levels.set(ilvl, parseLvl(lvl, resolveTheme, resolveFont));
      const linkedStyle = attrOf(child(lvl, 'w:pStyle'), 'w:val');
      if (linkedStyle !== undefined && !links.has(linkedStyle))
        links.set(linkedStyle, ilvl);
    }
    abstract.set(id, levels);
    if (links.size > 0) styleLinks.set(id, links);
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
        ? parseLvl(fullLvl, resolveTheme, resolveFont)
        : abstract.get(absId)?.get(ilvl);
      if (!base) continue;
      const entry: LevelEntry = { ...base };
      if (startOverride !== undefined) entry.start = Number(startOverride) || 1;
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
    const plain: Record<number, NumberingLevelDef> = {};
    for (const [ilvl, def] of levels) {
      plain[ilvl] = plainDef(ov?.get(ilvl) ?? def);
    }
    for (const [ilvl, def] of ov ?? []) {
      // Overrides for levels the abstract never declared.
      plain[ilvl] ??= plainDef(def);
    }
    // An overridden num counts independently from its siblings on the same
    // abstract (that is what startOverride is FOR) — give it its own key.
    (defs ??= {})[numId] = {
      key: ov ? `${absId}#${numId}` : absId,
      levels: plain,
    };
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

  function levelForStyle(numId: string, styleId: string): number | undefined {
    const absId = numToAbstract.get(numId);
    return absId === undefined
      ? undefined
      : styleLinks.get(absId)?.get(styleId);
  }

  return {
    defs,
    levelPPr,
    levelForStyle,
    auditMarkUnusedLevels: () => {
      if (!audit.enabled) return;
      for (const levels of abstract.values())
        for (const entry of levels.values())
          if (entry.pPr && !usedPPr.has(entry.pPr))
            audit.markSubtree(entry.pPr);
      for (const levels of numOverrides.values())
        for (const entry of levels.values())
          if (entry.pPr && !usedPPr.has(entry.pPr))
            audit.markSubtree(entry.pPr);
    },
  };
}
