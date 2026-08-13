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

/** The four kinds a `w:style` can declare. Word applies the `w:default="1"`
 *  style OF A KIND to content of that kind naming no style — which is why the
 *  kind has to be read at all, not just glanced at for the paragraph case. */
export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

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
  /** w:style/@w:type. Optional in the schema; absent reads as "paragraph". */
  type: string;
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
  /** The style Word applies to content of `type` that names none — the
   *  `w:default="1"` one, "Normal" / "Default Paragraph Font" / "Table
   *  Normal" / "No List" in a stock document. These are not inert defaults:
   *  Normal routinely carries the document's line spacing and space-after
   *  (skipping it collapsed every unstyled paragraph until 49862f1), and
   *  Table Normal carries the 108-twip cell margins.
   *
   *  Type-specific resolvers below fall back to this on their own; the
   *  polymorphic `resolveStyle` cannot, since it serves both paragraph and
   *  character lookups — its callers pass the right default in. */
  defaultStyleIdFor(type: StyleType): string | undefined;
  /** Effective run properties contributed by a styleId (incl. basedOn chain). */
  resolveStyle(styleId: string | undefined): RunProps;
  /** The styleId's w:pPr nodes, base-most first (basedOn ancestors → style).
   *  Callers append the inline pPr and resolve "later wins" per property. */
  resolveStylePPr(styleId: string | undefined): OoxmlNode[];
  /**
   * A TABLE style's w:pPr / w:rPr — the defaults it gives the paragraphs and
   * runs INSIDE the table (CT_Style: "Style Paragraph Properties — formats
   * paragraphs within the table").
   *
   * Same shape as resolveStylePPr/resolveStyle, and deliberately the same two
   * functions underneath: rolling up a basedOn chain does not care what type
   * of style it is walking. What differs is only where the caller puts the
   * result — a table style sits between the document defaults and the
   * paragraph style, not after it.
   *
   * A table naming no `w:tblStyle` resolves through the default table style,
   * exactly as Word does.
   */
  resolveTableStylePPr(styleId: string | undefined): OoxmlNode[];
  resolveTableStyleRPr(styleId: string | undefined): RunProps;
  /** The most-derived w:tblBorders a table style (chain) contributes, if any.
   *  A table naming no `w:tblStyle` resolves through the default table style,
   *  exactly as Word does. */
  resolveTableBorders(styleId: string | undefined): OoxmlNode | undefined;
  /** The most-derived w:tblCellMar a table style (chain) contributes, if any.
   *  Same default-style fallback as resolveTableBorders. */
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
  const defaultIds = new Map<string, string>();
  for (const style of children(stylesEl, 'w:style')) {
    const id = attrOf(style, 'w:styleId');
    if (id === undefined) continue;
    const isDefault =
      attrOf(style, 'w:default') === '1' ||
      attrOf(style, 'w:default') === 'true';
    // Read unconditionally: asking only inside the isDefault branch left the
    // kind of every ordinary style unread (63 of them in one lesson plan).
    const type = attrOf(style, 'w:type') ?? 'paragraph';
    if (isDefault && !defaultIds.has(type)) defaultIds.set(type, id);
    defs.set(id, {
      type,
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

  const defaultStyleIdFor = (type: StyleType) => defaultIds.get(type);
  // A table with no w:tblStyle still inherits the default table style, the
  // same way an unstyled paragraph inherits Normal.
  const tableStyle = (styleId: string | undefined) =>
    styleId ?? defaultStyleIdFor('table');

  return {
    docDefaults,
    docDefaultsPPr,
    defaultStyleIdFor,
    resolveStyle: (styleId) => resolve(styleId, new Set<string>()),
    resolveStylePPr: (styleId) => resolvePPr(styleId, new Set<string>()),
    resolveTableStylePPr: (styleId) =>
      resolvePPr(tableStyle(styleId), new Set<string>()),
    resolveTableStyleRPr: (styleId) =>
      resolve(tableStyle(styleId), new Set<string>()),
    resolveTableBorders: (styleId) =>
      resolveTblBorders(tableStyle(styleId), new Set<string>()),
    resolveTableCellMar: (styleId) =>
      resolveTblCellMar(tableStyle(styleId), new Set<string>()),
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
