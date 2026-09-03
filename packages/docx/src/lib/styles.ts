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

/**
 * One `w:tblStylePr` branch of a table style, rolled up through the basedOn
 * chain: the paragraph/run/cell defaults it contributes to the region it names
 * (first row, a banded column, the top-left cell…).
 *
 * `pPr`/`tcPr` stay as node LISTS (base-most first) because their children
 * merge per-property — the same shape `resolveStylePPr` returns; `rPr` is
 * already merged since RunProps is a flat record.
 */
export interface CondLayer {
  pPr: OoxmlNode[];
  rPr: RunProps;
  tcPr: OoxmlNode[];
}

interface StyleDef {
  basedOn?: string;
  rPr: RunProps;
  pPr?: OoxmlNode;
  /** `w:tblStylePr` branches, document order. Collected WITHOUT the audit
   *  accessors: a wholeTable branch must stay unread (see resolveCond). */
  cond: OoxmlNode[];
  /** w:tblPr/w:tblStyleRowBandSize and …ColBandSize, unparsed. */
  bandRow?: string;
  bandCol?: string;
  /** w:style/w:tcPr of a table style — cell defaults for the whole table. */
  tcPr?: OoxmlNode;
  /** w:tblPr/w:tblBorders of a table style, if any. */
  tblBorders?: OoxmlNode;
  /** w:tblPr/w:tblCellMar of a table style (Word's defaults carry the
   *  108-twip side margins here), if any. */
  tblCellMar?: OoxmlNode;
  /** w:tblPr/w:jc of a table style — the alignment it gives the whole table. */
  tblJc?: string;
  /** The style's `w:tblPr`, for properties resolved LAZILY (tblInd): read
   *  only on the layer the cascade actually lands on, so a base style's
   *  value that a derived style overrides stays unread — and the audit's
   *  own inert/overridden reasoning, not a stray visit, decides its fate. */
  tblPr?: OoxmlNode;
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
  /** A TABLE style's own `w:tcPr` nodes, base-most first — the cell defaults
   *  it gives every cell of the table, under the conditional branches and the
   *  cell's own w:tcPr. Same default-table-style fallback as the others. */
  resolveTableStyleTcPr(styleId: string | undefined): OoxmlNode[];
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
  /**
   * The table style's CONDITIONAL branches (`w:tblStylePr`), keyed by
   * `@w:type` — "firstRow", "band1Horz", "seCell"… Which of them reach a given
   * cell is decided per cell from its position and the table's `w:tblLook`;
   * this only rolls the definitions up through the basedOn chain.
   *
   * `wholeTable` is deliberately absent: *"Word does not apply and discards on
   * save any properties within the tblStylePr element when the type attribute
   * has a value of wholeTable"* (MS-OI29500 §17.18.89(a)). Whole-table
   * formatting lives at the style's own w:pPr/w:rPr/w:tblPr, which the
   * resolveTableStylePPr / resolveTableStyleRPr / resolveTableBorders /
   * resolveTableCellMar resolvers already read.
   */
  resolveTableStyleCond(styleId: string | undefined): Map<string, CondLayer>;
  /**
   * `w:tblStyleRowBandSize` / `w:tblStyleColBandSize` — how many rows/columns
   * make one band. **Word's default is 0, not the standard's 1**, and *"if
   * tblStyleRowBandSize is set to 0, Word does not apply any banded row
   * conditional formatting"* (MS-OI29500 §2.1.251) — so a style that declares
   * neither gets no banding at all.
   */
  resolveTableBandSizes(styleId: string | undefined): {
    row: number;
    col: number;
  };
  /** The most-derived w:tblBorders a table style (chain) contributes, if any.
   *  A table naming no `w:tblStyle` resolves through the default table style,
   *  exactly as Word does. */
  resolveTableBorders(styleId: string | undefined): OoxmlNode | undefined;
  /** The most-derived w:tblCellMar a table style (chain) contributes, if any.
   *  Same default-style fallback as resolveTableBorders. */
  resolveTableCellMar(styleId: string | undefined): OoxmlNode | undefined;
  /** The most-derived `w:tblPr/w:jc` a table style (chain) contributes — the
   *  table's alignment when the table itself declares none. Raw ST_JcTable
   *  value ("center", "end", …); the caller maps it. */
  resolveTableJc(styleId: string | undefined): string | undefined;
  /** The most-derived `w:tblPr/w:tblInd` a table style (chain) contributes —
   *  the table's indent when the table itself declares none. */
  resolveTableInd(styleId: string | undefined): OoxmlNode | undefined;
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
    // Both alignment spellings are read EAGERLY (`??` would leave the loser
    // unread and the audit noisy): w:tblPr/w:jc, and STYLE-LEVEL w:trPr/w:jc
    // — Word honours the latter (probe T7: a style whose only alignment is
    // <w:trPr><w:jc w:val="center"/></w:trPr> centres the table, measured at
    // exactly (content−grid)/2 in its PDF).
    const tblPrJc = attrOf(child(child(style, 'w:tblPr'), 'w:jc'), 'w:val');
    const trPrJc = attrOf(child(child(style, 'w:trPr'), 'w:jc'), 'w:val');
    // A whole-table w:shd in a table STYLE is a measured no-op: Word paints
    // none of it (probe T7's FDE9D9 fill is absent from the PDF in both
    // compat 12 and 15). Consumed so the audit reports a decision, not a gap.
    const styleTblShd = child(child(style, 'w:tblPr'), 'w:shd');
    if (styleTblShd) audit.markSubtree(styleTblShd);
    defs.set(id, {
      type,
      basedOn: attrOf(child(style, 'w:basedOn'), 'w:val'),
      rPr: parseRunProps(child(style, 'w:rPr'), resolveTheme, resolveFont),
      pPr: child(style, 'w:pPr'),
      tcPr: child(style, 'w:tcPr'),
      tblBorders: child(child(style, 'w:tblPr'), 'w:tblBorders'),
      tblCellMar: child(child(style, 'w:tblPr'), 'w:tblCellMar'),
      tblJc: tblPrJc ?? trPrJc,
      tblPr: child(style, 'w:tblPr'),
      // Filtered off node.children, not via children(): the accessor would
      // mark every branch as read, including the wholeTable one the audit is
      // supposed to keep reporting.
      cond: style.children.filter((c) => c.name === 'w:tblStylePr'),
      bandRow: attrOf(
        child(child(style, 'w:tblPr'), 'w:tblStyleRowBandSize'),
        'w:val',
      ),
      bandCol: attrOf(
        child(child(style, 'w:tblPr'), 'w:tblStyleColBandSize'),
        'w:val',
      ),
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

  function resolveTcPr(
    styleId: string | undefined,
    seen: Set<string>,
  ): OoxmlNode[] {
    if (!styleId || seen.has(styleId)) return [];
    const def = defs.get(styleId);
    if (!def) return [];
    seen.add(styleId);
    usedIds.add(styleId);
    const base = def.basedOn ? resolveTcPr(def.basedOn, seen) : [];
    return def.tcPr ? [...base, def.tcPr] : base;
  }

  const BORDER_SIDES = [
    'w:top',
    'w:left',
    'w:bottom',
    'w:right',
    'w:insideH',
    'w:insideV',
  ];

  /**
   * Per-side overlay of a basedOn chain (collected derived-first): walking
   * base→derived, each layer's declaration of a side replaces the one
   * beneath it, so the result is the nearest declaration per side. `names`
   * are the spellings of ONE logical side (w:left/w:start), so a derived
   * w:start beats a base w:left instead of sitting next to it.
   *
   * A replaced declaration is consumed on purpose: the merge read it and a
   * nearer layer beat it, so nothing on the page can come from it — the
   * audit must not report a shadowed side as a gap (TableNormal's cell
   * margins under a TableGrid that redeclares all four used to show as
   * four UNKNOWN lines per file).
   */
  function overlaySides(
    chain: OoxmlNode[],
    sides: readonly (readonly string[])[],
  ): OoxmlNode[] {
    const out: OoxmlNode[] = [];
    for (const names of sides) {
      let win: OoxmlNode | undefined;
      for (let i = chain.length - 1; i >= 0; i--) {
        const el = names.map((n) => child(chain[i], n)).find(Boolean);
        if (!el) continue;
        if (win) audit.markSubtree(win);
        win = el;
      }
      if (win) out.push(win);
    }
    return out;
  }

  function resolveTblBorders(
    styleId: string | undefined,
    seen: Set<string>,
  ): OoxmlNode | undefined {
    // Word inherits table-style borders PER SIDE through basedOn — a derived
    // style overriding only w:bottom keeps the base's other five sides. With
    // a single declaration the element passes through as-is.
    const chain: OoxmlNode[] = [];
    let id = styleId;
    while (id && !seen.has(id)) {
      const def = defs.get(id);
      if (!def) break;
      seen.add(id);
      usedIds.add(id);
      if (def.tblBorders) chain.push(def.tblBorders);
      id = def.basedOn;
    }
    if (chain.length === 0) return undefined;
    if (chain.length === 1) return chain[0];
    return {
      name: 'w:tblBorders',
      attrs: {},
      children: overlaySides(
        chain,
        BORDER_SIDES.map((s) => [s]),
      ),
      text: '',
    };
  }

  function resolveTblJc(
    styleId: string | undefined,
    seen: Set<string>,
  ): string | undefined {
    if (!styleId || seen.has(styleId)) return undefined;
    const def = defs.get(styleId);
    if (!def) return undefined;
    seen.add(styleId);
    usedIds.add(styleId);
    return def.tblJc ?? resolveTblJc(def.basedOn, seen);
  }

  function resolveTblInd(
    styleId: string | undefined,
    seen: Set<string>,
  ): OoxmlNode | undefined {
    if (!styleId || seen.has(styleId)) return undefined;
    const def = defs.get(styleId);
    if (!def) return undefined;
    seen.add(styleId);
    usedIds.add(styleId);
    return child(def.tblPr, 'w:tblInd') ?? resolveTblInd(def.basedOn, seen);
  }

  /** Logical margin sides; left/right also answer to the ST_ThemeColor-era
   *  w:start/w:end spellings, so a derived style writing w:start must beat a
   *  base's w:left, not sit beside it. */
  const MAR_SIDES: readonly (readonly string[])[] = [
    ['w:top'],
    ['w:left', 'w:start'],
    ['w:bottom'],
    ['w:right', 'w:end'],
  ];

  function resolveTblCellMar(
    styleId: string | undefined,
    seen: Set<string>,
  ): OoxmlNode | undefined {
    // Word inherits cell margins PER SIDE through basedOn, exactly like
    // tblBorders above — measured, not guessed (probe F1, Word PDF): a
    // derived style declaring ONLY w:top keeps the base's left/right 720
    // twips (text inset 36.24pt = 56.64 − 20.4 in the PDF, all three
    // tables), and an EMPTY w:tblCellMar erases nothing. The previous
    // `def.tblCellMar ?? recurse` was wrong both ways: a partial override
    // dropped the base's other sides, and an empty element cut the chain.
    const chain: OoxmlNode[] = [];
    let id = styleId;
    while (id && !seen.has(id)) {
      const def = defs.get(id);
      if (!def) break;
      seen.add(id);
      usedIds.add(id);
      if (def.tblCellMar) chain.push(def.tblCellMar);
      id = def.basedOn;
    }
    if (chain.length === 0) return undefined;
    if (chain.length === 1) return chain[0];
    return {
      name: 'w:tblCellMar',
      attrs: {},
      children: overlaySides(chain, MAR_SIDES),
      text: '',
    };
  }

  /** Roll the conditional branches up the chain: a derived style's branch of
   *  the same type layers ON TOP of the base's, per-property, so a style that
   *  only overrides firstRow's shading keeps the base's firstRow fonts. */
  function resolveCond(
    styleId: string | undefined,
    seen: Set<string>,
  ): Map<string, CondLayer> {
    if (!styleId || seen.has(styleId)) return new Map();
    const def = defs.get(styleId);
    if (!def) return new Map();
    seen.add(styleId);
    usedIds.add(styleId);
    const out = def.basedOn ? resolveCond(def.basedOn, seen) : new Map();
    for (const branch of def.cond) {
      // attrOf records the ATTRIBUTE without marking the element, so a
      // wholeTable branch leaves this loop entirely unread — which is the
      // point: the audit keeps reporting it (as inert, see audit.ts) instead
      // of us claiming to honour something Word itself discards.
      const type = attrOf(branch, 'w:type');
      if (type === undefined || type === 'wholeTable') continue;
      audit.mark(branch);
      const base = out.get(type);
      const pPr = child(branch, 'w:pPr');
      const tcPr = child(branch, 'w:tcPr');
      out.set(type, {
        pPr: pPr ? [...(base?.pPr ?? []), pPr] : (base?.pPr ?? []),
        rPr: mergeRunProps(
          base?.rPr ?? EMPTY,
          parseRunProps(child(branch, 'w:rPr'), resolveTheme, resolveFont),
        ),
        tcPr: tcPr ? [...(base?.tcPr ?? []), tcPr] : (base?.tcPr ?? []),
      });
    }
    return out;
  }

  function resolveBands(
    styleId: string | undefined,
    seen: Set<string>,
  ): { row: number; col: number } {
    if (!styleId || seen.has(styleId)) return { row: 0, col: 0 };
    const def = defs.get(styleId);
    if (!def) return { row: 0, col: 0 };
    seen.add(styleId);
    usedIds.add(styleId);
    const base = resolveBands(def.basedOn, seen);
    const num = (v: string | undefined, dflt: number) => {
      const n = Number(v);
      return v === undefined || Number.isNaN(n) ? dflt : n;
    };
    return {
      row: num(def.bandRow, base.row),
      col: num(def.bandCol, base.col),
    };
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
    resolveTableStyleTcPr: (styleId) =>
      resolveTcPr(tableStyle(styleId), new Set<string>()),
    resolveTableStyleCond: (styleId) =>
      resolveCond(tableStyle(styleId), new Set<string>()),
    resolveTableBandSizes: (styleId) =>
      resolveBands(tableStyle(styleId), new Set<string>()),
    resolveTableBorders: (styleId) =>
      resolveTblBorders(tableStyle(styleId), new Set<string>()),
    resolveTableCellMar: (styleId) =>
      resolveTblCellMar(tableStyle(styleId), new Set<string>()),
    resolveTableJc: (styleId) =>
      resolveTblJc(tableStyle(styleId), new Set<string>()),
    resolveTableInd: (styleId) =>
      resolveTblInd(tableStyle(styleId), new Set<string>()),
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
