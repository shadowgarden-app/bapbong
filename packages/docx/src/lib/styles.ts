import {
  attrOf,
  child,
  children,
  mergeRunProps,
  OoxmlNode,
  parseRunProps,
  RunProps,
  ThemeColorResolver,
} from './ooxml.js';

interface StyleDef {
  basedOn?: string;
  rPr: RunProps;
  pPr?: OoxmlNode;
  /** w:tblPr/w:tblBorders of a table style, if any. */
  tblBorders?: OoxmlNode;
}

/** Resolved view of `word/styles.xml`: document defaults plus named styles,
 *  with `w:basedOn` inheritance flattened on demand. */
export interface StyleRegistry {
  /** docDefaults → rPrDefault, the lowest layer of the run cascade. */
  docDefaults: RunProps;
  /** docDefaults → pPrDefault, the lowest layer of the paragraph cascade. */
  docDefaultsPPr: OoxmlNode | undefined;
  /** Effective run properties contributed by a styleId (incl. basedOn chain). */
  resolveStyle(styleId: string | undefined): RunProps;
  /** The styleId's w:pPr nodes, base-most first (basedOn ancestors → style).
   *  Callers append the inline pPr and resolve "later wins" per property. */
  resolveStylePPr(styleId: string | undefined): OoxmlNode[];
  /** The most-derived w:tblBorders a table style (chain) contributes, if any. */
  resolveTableBorders(styleId: string | undefined): OoxmlNode | undefined;
}

const EMPTY: RunProps = {};

export function buildStyleRegistry(
  stylesRoot: OoxmlNode | undefined,
  resolveTheme?: ThemeColorResolver,
): StyleRegistry {
  const stylesEl = child(stylesRoot, 'w:styles');

  const rPrDefault = child(child(child(stylesEl, 'w:docDefaults'), 'w:rPrDefault'), 'w:rPr');
  const docDefaults = parseRunProps(rPrDefault, resolveTheme);
  const docDefaultsPPr = child(child(child(stylesEl, 'w:docDefaults'), 'w:pPrDefault'), 'w:pPr');

  const defs = new Map<string, StyleDef>();
  for (const style of children(stylesEl, 'w:style')) {
    const id = attrOf(style, 'w:styleId');
    if (id === undefined) continue;
    defs.set(id, {
      basedOn: attrOf(child(style, 'w:basedOn'), 'w:val'),
      rPr: parseRunProps(child(style, 'w:rPr'), resolveTheme),
      pPr: child(style, 'w:pPr'),
      tblBorders: child(child(style, 'w:tblPr'), 'w:tblBorders'),
    });
  }

  function resolve(styleId: string | undefined, seen: Set<string>): RunProps {
    if (!styleId || seen.has(styleId)) return EMPTY;
    const def = defs.get(styleId);
    if (!def) return EMPTY;
    seen.add(styleId);
    const base = def.basedOn ? resolve(def.basedOn, seen) : EMPTY;
    return mergeRunProps(base, def.rPr);
  }

  function resolvePPr(styleId: string | undefined, seen: Set<string>): OoxmlNode[] {
    if (!styleId || seen.has(styleId)) return [];
    const def = defs.get(styleId);
    if (!def) return [];
    seen.add(styleId);
    const base = def.basedOn ? resolvePPr(def.basedOn, seen) : [];
    return def.pPr ? [...base, def.pPr] : base;
  }

  function resolveTblBorders(styleId: string | undefined, seen: Set<string>): OoxmlNode | undefined {
    if (!styleId || seen.has(styleId)) return undefined;
    const def = defs.get(styleId);
    if (!def) return undefined;
    seen.add(styleId);
    return def.tblBorders ?? resolveTblBorders(def.basedOn, seen);
  }

  return {
    docDefaults,
    docDefaultsPPr,
    resolveStyle: (styleId) => resolve(styleId, new Set<string>()),
    resolveStylePPr: (styleId) => resolvePPr(styleId, new Set<string>()),
    resolveTableBorders: (styleId) => resolveTblBorders(styleId, new Set<string>()),
  };
}
