import {
  attrOf,
  child,
  children,
  mergeRunProps,
  OoxmlNode,
  parseRunProps,
  RunProps,
  ThemeColorResolver,
} from './ooxml';

interface StyleDef {
  basedOn?: string;
  rPr: RunProps;
}

/** Resolved view of `word/styles.xml`: document defaults plus named styles,
 *  with `w:basedOn` inheritance flattened on demand. */
export interface StyleRegistry {
  /** docDefaults → rPrDefault, the lowest layer of the run cascade. */
  docDefaults: RunProps;
  /** Effective run properties contributed by a styleId (incl. basedOn chain). */
  resolveStyle(styleId: string | undefined): RunProps;
}

const EMPTY: RunProps = {};

export function buildStyleRegistry(
  stylesRoot: OoxmlNode | undefined,
  resolveTheme?: ThemeColorResolver,
): StyleRegistry {
  const stylesEl = child(stylesRoot, 'w:styles');

  const rPrDefault = child(child(child(stylesEl, 'w:docDefaults'), 'w:rPrDefault'), 'w:rPr');
  const docDefaults = parseRunProps(rPrDefault, resolveTheme);

  const defs = new Map<string, StyleDef>();
  for (const style of children(stylesEl, 'w:style')) {
    const id = attrOf(style, 'w:styleId');
    if (id === undefined) continue;
    defs.set(id, {
      basedOn: attrOf(child(style, 'w:basedOn'), 'w:val'),
      rPr: parseRunProps(child(style, 'w:rPr'), resolveTheme),
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

  return {
    docDefaults,
    resolveStyle: (styleId) => resolve(styleId, new Set<string>()),
  };
}
