import { audit } from './audit.js';
import {
  attrOf,
  child,
  children,
  mergeRunProps,
  OoxmlNode,
  parseRunProps,
  RunProps,
  ThemeColorResolver,
  ThemeFontResolver,
} from './ooxml.js';

interface StyleDef {
  basedOn?: string;
  rPr: RunProps;
  pPr?: OoxmlNode;
  /** w:tblPr/w:tblBorders of a table style, if any. */
  tblBorders?: OoxmlNode;
  /** w:tblPr/w:tblCellMar of a table style (Word's defaults carry the
   *  108-twip side margins here), if any. */
  tblCellMar?: OoxmlNode;
  /** The w:style element itself — for the audit's unused-style sweep. */
  el: OoxmlNode;
  /** w:default="1": Word applies this style to unstyled content, so it is
   *  implicitly in use even when nothing names it. */
  isDefault: boolean;
}

/** Resolved view of `word/styles.xml`: document defaults plus named styles,
 *  with `w:basedOn` inheritance flattened on demand. */
export interface StyleRegistry {
  /** docDefaults → rPrDefault, the lowest layer of the run cascade. */
  docDefaults: RunProps;
  /** docDefaults → pPrDefault, the lowest layer of the paragraph cascade. */
  docDefaultsPPr: OoxmlNode | undefined;
  /** The paragraph style Word applies to content that names none —
   *  `w:style w:type="paragraph" w:default="1"`, almost always "Normal".
   *  It routinely carries the document's line spacing and space-after, so
   *  skipping it collapses every unstyled paragraph. */
  defaultParagraphStyleId: string | undefined;
  /** Effective run properties contributed by a styleId (incl. basedOn chain). */
  resolveStyle(styleId: string | undefined): RunProps;
  /** The styleId's w:pPr nodes, base-most first (basedOn ancestors → style).
   *  Callers append the inline pPr and resolve "later wins" per property. */
  resolveStylePPr(styleId: string | undefined): OoxmlNode[];
  /** The most-derived w:tblBorders a table style (chain) contributes, if any. */
  resolveTableBorders(styleId: string | undefined): OoxmlNode | undefined;
  /** The most-derived w:tblCellMar a table style (chain) contributes, if any. */
  resolveTableCellMar(styleId: string | undefined): OoxmlNode | undefined;
  /** XML-audit hook, call once after every story is parsed: styles NOTHING
   *  referenced (directly, via basedOn, or as a w:default) are marked as
   *  consumed subtrees — an unused style can't lose this document's data
   *  (styles.xml itself survives export via the carry package), so reporting
   *  its properties as UNKNOWN would only bury the real gaps. */
  auditMarkUnusedStyles(): void;
}

const EMPTY: RunProps = {};

export function buildStyleRegistry(
  stylesRoot: OoxmlNode | undefined,
  resolveTheme?: ThemeColorResolver,
  resolveFont?: ThemeFontResolver,
): StyleRegistry {
  const stylesEl = child(stylesRoot, 'w:styles');

  const rPrDefault = child(
    child(child(stylesEl, 'w:docDefaults'), 'w:rPrDefault'),
    'w:rPr',
  );
  const docDefaults = parseRunProps(rPrDefault, resolveTheme, resolveFont);
  const docDefaultsPPr = child(
    child(child(stylesEl, 'w:docDefaults'), 'w:pPrDefault'),
    'w:pPr',
  );

  const defs = new Map<string, StyleDef>();
  let defaultParagraphStyleId: string | undefined;
  for (const style of children(stylesEl, 'w:style')) {
    const id = attrOf(style, 'w:styleId');
    if (id === undefined) continue;
    const isDefault =
      attrOf(style, 'w:default') === '1' ||
      attrOf(style, 'w:default') === 'true';
    if (isDefault && attrOf(style, 'w:type') === 'paragraph')
      defaultParagraphStyleId ??= id;
    defs.set(id, {
      basedOn: attrOf(child(style, 'w:basedOn'), 'w:val'),
      rPr: parseRunProps(child(style, 'w:rPr'), resolveTheme, resolveFont),
      pPr: child(style, 'w:pPr'),
      tblBorders: child(child(style, 'w:tblPr'), 'w:tblBorders'),
      tblCellMar: child(child(style, 'w:tblPr'), 'w:tblCellMar'),
      el: style,
      isDefault,
    });
  }

  // Every id the document actually pulled through the cascade (directly or
  // as a basedOn ancestor) — feeds the audit's unused-style sweep.
  const usedIds = new Set<string>();

  function resolve(styleId: string | undefined, seen: Set<string>): RunProps {
    if (!styleId || seen.has(styleId)) return EMPTY;
    const def = defs.get(styleId);
    if (!def) return EMPTY;
    seen.add(styleId);
    usedIds.add(styleId);
    const base = def.basedOn ? resolve(def.basedOn, seen) : EMPTY;
    return mergeRunProps(base, def.rPr);
  }

  function resolvePPr(
    styleId: string | undefined,
    seen: Set<string>,
  ): OoxmlNode[] {
    if (!styleId || seen.has(styleId)) return [];
    const def = defs.get(styleId);
    if (!def) return [];
    seen.add(styleId);
    usedIds.add(styleId);
    const base = def.basedOn ? resolvePPr(def.basedOn, seen) : [];
    return def.pPr ? [...base, def.pPr] : base;
  }

  function resolveTblBorders(
    styleId: string | undefined,
    seen: Set<string>,
  ): OoxmlNode | undefined {
    if (!styleId || seen.has(styleId)) return undefined;
    const def = defs.get(styleId);
    if (!def) return undefined;
    seen.add(styleId);
    usedIds.add(styleId);
    return def.tblBorders ?? resolveTblBorders(def.basedOn, seen);
  }

  function resolveTblCellMar(
    styleId: string | undefined,
    seen: Set<string>,
  ): OoxmlNode | undefined {
    if (!styleId || seen.has(styleId)) return undefined;
    const def = defs.get(styleId);
    if (!def) return undefined;
    seen.add(styleId);
    usedIds.add(styleId);
    return def.tblCellMar ?? resolveTblCellMar(def.basedOn, seen);
  }

  return {
    docDefaults,
    docDefaultsPPr,
    defaultParagraphStyleId,
    resolveStyle: (styleId) => resolve(styleId, new Set<string>()),
    resolveStylePPr: (styleId) => resolvePPr(styleId, new Set<string>()),
    resolveTableBorders: (styleId) =>
      resolveTblBorders(styleId, new Set<string>()),
    resolveTableCellMar: (styleId) =>
      resolveTblCellMar(styleId, new Set<string>()),
    auditMarkUnusedStyles: () => {
      if (!audit.enabled) return;
      for (const [id, def] of defs) {
        // w:default styles are implicitly applied to unstyled content — an
        // unread property THERE is a real gap, never "unused".
        if (!usedIds.has(id) && !def.isDefault) audit.markSubtree(def.el);
      }
    },
  };
}
