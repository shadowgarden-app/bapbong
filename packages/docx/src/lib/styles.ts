import { asRecord, attr, mergeRunProps, parseRunProps, RunProps, toArray } from './ooxml';

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
  stylesTree: Record<string, unknown> | undefined,
): StyleRegistry {
  const stylesEl = asRecord(stylesTree?.['w:styles']);

  const rPrDefault = asRecord(
    asRecord(asRecord(stylesEl?.['w:docDefaults'])?.['w:rPrDefault'])?.['w:rPr'],
  );
  const docDefaults = parseRunProps(rPrDefault);

  const defs = new Map<string, StyleDef>();
  for (const styleUnknown of toArray(stylesEl?.['w:style'])) {
    const style = asRecord(styleUnknown);
    const id = attr(style, '@_w:styleId');
    if (!style || !id) continue;
    defs.set(id, {
      basedOn: attr(style['w:basedOn'], '@_w:val'),
      rPr: parseRunProps(asRecord(style['w:rPr'])),
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
