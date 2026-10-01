/**
 * A DrawingML chart part (c:chartSpace) → a {@link ChartSpec}: the data from
 * the part's caches, and every colour, line and font DECIDED here, so the
 * layout engine only has to lay it out.
 *
 * Two generations of chart write very different XML, and this is where they
 * meet:
 *
 *  - Word 2007/2010 charts say almost nothing about their look — a
 *    `c:style` number (1…48) and the theme. Everything else is Word's
 *    automatic formatting for that style: series colours cycling the accents
 *    (or fading one accent), lines from the theme's line styles, grey axes,
 *    10pt text, an 18pt bold title. {@link styleDefaults} is that table.
 *  - Word 2013+ charts write the look out in full (`c:spPr`, `c:txPr` on
 *    every element) — what the file says is what paints.
 *
 * So the rule is per element: explicit formatting wins, the c:style default
 * fills whatever the element leaves unsaid. The default table follows the
 * Excel 2007 auto-format tables as LibreOffice's OOXML import reconstructed
 * them (oox/source/drawingml/chart/objectformatter.cxx), corrected where Word
 * was measured to differ (chart-area border).
 *
 * Returns null for what is not rendered (3-D, stock, bubble, surface,
 * of-pie, chartex): the caller keeps the placeholder.
 */
import type {
  ChartAxis,
  ChartBox,
  ChartDataLabels,
  ChartFont,
  ChartGroup,
  ChartLegend,
  ChartLine,
  ChartManualLayout,
  ChartMarker,
  ChartMarkerSymbol,
  ChartPoint,
  ChartSeries,
  ChartSpec,
  ChartText,
} from '@shadow-garden/bapbong-contracts';
import { formatNumber } from '@shadow-garden/bapbong-contracts';
import {
  attrOf,
  child,
  children,
  findDescendant,
  parseXml,
  type OoxmlNode,
} from './ooxml.js';
import {
  buildThemeFontResolver,
  buildThemeResolver,
  drawingColor,
  type ThemeResolver,
} from './theme.js';

const EMU_PER_PT = 12700;

/** What the chart's look is resolved against: the document theme, or the
 *  chart's own c:themeOverride part. */
export interface ChartTheme {
  color: ThemeResolver;
  /** 'minor' / 'major' latin typeface. */
  font: (slot: 'minor' | 'major') => string | undefined;
  /** The fmtScheme's a:lnStyleLst entry (1-based) — width and look of the
   *  themed lines the chart styles name. */
  lineStyle: (idx: number) => OoxmlNode | undefined;
  /** The fmtScheme's a:fillStyleLst entry (1-based) filled with `phClr`. */
  fillStyle: (idx: number, phClr: string) => string;
}

export function chartTheme(themeRoot: OoxmlNode | undefined): ChartTheme {
  const color = buildThemeResolver(themeRoot);
  const fonts = buildThemeFontResolver(themeRoot);
  const lines = findDescendant(themeRoot, 'a:lnStyleLst');
  const fills = findDescendant(themeRoot, 'a:fillStyleLst');
  return {
    color,
    font: (slot) => fonts(slot === 'major' ? 'majorHAnsi' : 'minorHAnsi'),
    lineStyle: (idx) => lines?.children[idx - 1],
    fillStyle: (idx, phClr) => {
      const entry = fills?.children[idx - 1];
      if (!entry) return phClr;
      if (entry.name === 'a:solidFill')
        return drawingColor(entry, color, phClr) ?? phClr;
      // A gradient entry paints as its first stop — the box model has a
      // single colour. Close for the stock themes' subtle gradients.
      const stop = child(child(entry, 'a:gsLst'), 'a:gs');
      return drawingColor(stop, color, phClr) ?? phClr;
    },
  };
}

// ── colours ─────────────────────────────────────────────────────────────────

type Mod = [
  tag: 'tint' | 'shade' | 'lumMod' | 'lumOff' | 'satMod',
  val: number,
];

/** A synthetic a:solidFill holding `clr` + transforms — so every colour, the
 *  defaults' and the file's alike, goes through the one DrawingML resolver. */
function clrNode(clr: OoxmlNode): OoxmlNode {
  return { name: 'a:solidFill', attrs: {}, children: [clr], text: '' };
}
function schemeNode(val: string, mods: Mod[] = []): OoxmlNode {
  return {
    name: 'a:schemeClr',
    attrs: { val },
    children: mods.map(([t, v]) => ({
      name: `a:${t}`,
      attrs: { val: String(v) },
      children: [],
      text: '',
    })),
    text: '',
  };
}
function scheme(theme: ChartTheme, val: string, mods: Mod[] = []): string {
  return drawingColor(clrNode(schemeNode(val, mods)), theme.color) ?? '#000000';
}
function hexWith(theme: ChartTheme, hex: string, mods: Mod[]): string {
  if (mods.length === 0) return hex;
  const node: OoxmlNode = {
    name: 'a:srgbClr',
    attrs: { val: hex.replace('#', '') },
    children: schemeNode('x', mods).children,
    text: '',
  };
  return drawingColor(clrNode(node), theme.color) ?? hex;
}

// ── the c:style default table ───────────────────────────────────────────────

/** The 48 styles are a 6×8 grid: the row picks the effects (outline,
 *  intensity, dark background), the column the colours — 1 greys, 2 the
 *  accents in turn, 3…8 one accent faded across the series. */
function styleCell(style: number): { row: number; col: number } {
  const s = Math.min(48, Math.max(1, Math.round(style)));
  return { row: Math.floor((s - 1) / 8), col: ((s - 1) % 8) + 1 };
}

const GREYS: Mod[][] = [88500, 55000, 78000, 92500, 70000, 30000].map((v) => [
  ['tint', v],
]);
const DARK_GREYS: Mod[][] = [5000, 55000, 78000, 15000, 70000, 30000].map(
  (v) => [['tint', v]],
);

/**
 * A series' automatic colour. Colours cycle through the column's pattern;
 * each further cycle — or, for a single faded accent, each further series —
 * moves from shaded to tinted across −70%…+70% (Excel's chart tint).
 */
export function seriesColor(
  theme: ChartTheme,
  style: number,
  index: number,
  count: number,
): string {
  const { row, col } = styleCell(style);
  let pattern: string[];
  if (col === 1)
    pattern = (row === 5 ? DARK_GREYS : GREYS).map((m) =>
      scheme(theme, 'dk1', m),
    );
  else if (col === 2)
    pattern = [1, 2, 3, 4, 5, 6].map((n) => scheme(theme, `accent${n}`));
  else pattern = [scheme(theme, `accent${col - 2}`)];
  const base = pattern[index % pattern.length];
  const cycle = Math.floor(index / pattern.length);
  const maxCycle = Math.floor(Math.max(0, count - 1) / pattern.length);
  const f = ((cycle + 1) / (maxCycle + 2)) * 1.4 - 0.7;
  const v = Math.round(f * 100000);
  if (v < 0) return hexWith(theme, base, [['shade', v + 100000]]);
  if (v > 0) return hexWith(theme, base, [['tint', 100000 - v]]);
  return base;
}

/** Width of the theme's line style `idx`, in points (the Office default
 *  9525 EMU = 0.75pt when the theme says nothing). */
function themeLineWidth(theme: ChartTheme, idx: number): number {
  const w = Number(attrOf(theme.lineStyle(idx), 'w'));
  return Number.isFinite(w) && w > 0 ? w / EMU_PER_PT : 0.75;
}

/** Line style `idx` of the theme, coloured with `phClr`. */
function themedLine(
  theme: ChartTheme,
  idx: number,
  phClr: string,
  widthScale = 1,
): ChartLine {
  const entry = theme.lineStyle(idx);
  return {
    color:
      drawingColor(child(entry, 'a:solidFill'), theme.color, phClr) ?? phClr,
    widthPt: themeLineWidth(theme, idx) * widthScale,
  };
}

/** Everything the auto-formatting of one c:style decides. */
export interface StyleDefaults {
  text: string;
  chartFill: string;
  chartLine: ChartLine | null;
  plotFill: string | null;
  axisLine: ChartLine;
  majorGrid: ChartLine;
  minorGrid: ChartLine;
  /** Line / scatter / radar series. */
  seriesLine: (color: string) => ChartLine;
  /** Bar / area / pie series. */
  seriesFill: (color: string) => string;
  seriesOutline: (color: string, accentIndex: number) => ChartLine | null;
  marker: (color: string) => { fill: string; line: ChartLine };
}

export function styleDefaults(theme: ChartTheme, style: number): StyleDefaults {
  const { row, col } = styleCell(style);
  const s = (row * 8 + col) as number;
  const thin = themeLineWidth(theme, 1);
  const tx = s <= 40 ? scheme(theme, 'tx1') : scheme(theme, 'lt1');
  const greyLine = (
    mods: Mod[],
    base = s <= 32 ? 'tx1' : 'dk1',
  ): ChartLine => ({
    color: scheme(theme, base, mods),
    widthPt: thin,
  });
  const intense = row === 2 || row === 3 || row === 5;
  const lineScale = row === 0 ? 3 : row === 3 ? 7 : 5;
  return {
    text: tx,
    chartFill:
      s <= 32
        ? scheme(theme, 'bg1')
        : s <= 40
          ? scheme(theme, 'lt1')
          : scheme(theme, 'dk1'),
    // Measured: Word frames a 2007-style chart with the axis grey (0.75pt
    // #868686 on the Office theme). LibreOffice's table leaves it unframed.
    chartLine: greyLine([['tint', 75000]]),
    plotFill:
      s <= 32
        ? scheme(theme, 'bg1')
        : s <= 34
          ? scheme(theme, 'dk1', [['tint', 20000]])
          : s <= 40
            ? scheme(theme, `accent${s - 34}`, [['tint', 20000]])
            : scheme(theme, 'dk1', [['tint', 95000]]),
    axisLine: greyLine([['tint', 75000]]),
    majorGrid: greyLine([['tint', 75000]]),
    minorGrid: {
      color: scheme(theme, 'tx1', [['tint', s <= 40 ? 50000 : 90000]]),
      widthPt: thin,
    },
    seriesLine: (c) => themedLine(theme, 1, c, lineScale),
    seriesFill: (c) => theme.fillStyle(intense ? 3 : 1, c),
    seriesOutline: (c, accent) => {
      if (row === 1) return { color: scheme(theme, 'lt1'), widthPt: thin };
      if (row !== 4) return null;
      if (s === 33)
        return {
          color: scheme(theme, 'dk1', [['shade', 50000]]),
          widthPt: thin,
        };
      if (s === 34)
        return {
          color: scheme(theme, `accent${(accent % 6) + 1}`, [['shade', 50000]]),
          widthPt: thin,
        };
      return {
        color: scheme(theme, `accent${s - 34}`, [['shade', 50000]]),
        widthPt: thin,
      };
    },
    marker: (c) => ({
      fill: theme.fillStyle(1, c),
      line: themedLine(theme, 1, c),
    }),
  };
}

/** The automatic marker of the n-th series (Excel's cycle). */
const MARKER_CYCLE: ChartMarkerSymbol[] = [
  'diamond',
  'square',
  'triangle',
  'x',
  'star',
  'circle',
  'plus',
  'dot',
  'dash',
];

// ── explicit formatting ─────────────────────────────────────────────────────

/** c:spPr's fill: a colour, null for a:noFill, undefined when unsaid. */
function spFill(
  spPr: OoxmlNode | undefined,
  theme: ChartTheme,
): string | null | undefined {
  if (!spPr) return undefined;
  if (child(spPr, 'a:noFill')) return null;
  const solid = child(spPr, 'a:solidFill');
  if (solid) return drawingColor(solid, theme.color);
  const grad = child(spPr, 'a:gradFill');
  if (grad)
    return drawingColor(child(child(grad, 'a:gsLst'), 'a:gs'), theme.color);
  const patt = child(spPr, 'a:pattFill');
  if (patt) return drawingColor(child(patt, 'a:fgClr'), theme.color);
  return undefined;
}

const PRESET_DASH: Record<string, number[]> = {
  dash: [4, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  dot: [1, 3],
  dashDot: [4, 3, 1, 3],
  lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** c:spPr's line over a default: null for a:noFill, the default where the
 *  file is silent, the default's colour with the file's width where only
 *  the width is given. */
function spLine(
  spPr: OoxmlNode | undefined,
  theme: ChartTheme,
  fallback: ChartLine | null,
): ChartLine | null {
  const ln = child(spPr, 'a:ln');
  if (!ln) return fallback;
  if (child(ln, 'a:noFill')) return null;
  const w = Number(attrOf(ln, 'w'));
  const solid = child(ln, 'a:solidFill');
  const grad = child(ln, 'a:gradFill');
  const color =
    (solid && drawingColor(solid, theme.color)) ??
    (grad &&
      drawingColor(child(child(grad, 'a:gsLst'), 'a:gs'), theme.color)) ??
    fallback?.color;
  if (!color) return null;
  const line: ChartLine = {
    color,
    widthPt:
      Number.isFinite(w) && w > 0
        ? w / EMU_PER_PT
        : (fallback?.widthPt ?? 0.75),
  };
  const dash = attrOf(child(ln, 'a:prstDash'), 'val');
  if (dash && PRESET_DASH[dash]) line.dash = PRESET_DASH[dash];
  return line;
}

function spBox(
  spPr: OoxmlNode | undefined,
  theme: ChartTheme,
  fallback: ChartBox,
): ChartBox {
  const fill = spFill(spPr, theme);
  return {
    fill: fill === undefined ? fallback.fill : fill,
    line: spLine(spPr, theme, fallback.line),
  };
}

/** a:defRPr / a:rPr over a font. */
function runFont(
  rPr: OoxmlNode | undefined,
  theme: ChartTheme,
  base: ChartFont,
): ChartFont {
  if (!rPr) return base;
  const out = { ...base };
  const sz = Number(attrOf(rPr, 'sz'));
  if (Number.isFinite(sz) && sz > 0) out.sizePt = sz / 100;
  const b = attrOf(rPr, 'b');
  if (b !== undefined) out.bold = b === '1' || b === 'true';
  const i = attrOf(rPr, 'i');
  if (i !== undefined) out.italic = i === '1' || i === 'true';
  const color = drawingColor(child(rPr, 'a:solidFill'), theme.color);
  if (color) out.color = color;
  const face = attrOf(child(rPr, 'a:latin'), 'typeface');
  if (face) {
    const resolved =
      face === '+mn-lt'
        ? theme.font('minor')
        : face === '+mj-lt'
          ? theme.font('major')
          : face;
    if (resolved) out.family = resolved;
  }
  return out;
}

/** The first a:p's a:pPr/a:defRPr of a c:txPr or c:rich. */
function defRPr(body: OoxmlNode | undefined): OoxmlNode | undefined {
  return child(child(child(body, 'a:p'), 'a:pPr'), 'a:defRPr');
}

/** a:bodyPr@rot in degrees; the "-60000000" sentinel (and absence) = auto. */
function bodyRotation(body: OoxmlNode | undefined): number | undefined {
  const rot = Number(attrOf(child(body, 'a:bodyPr'), 'rot'));
  if (!Number.isFinite(rot) || rot === -60000000) return undefined;
  return rot / 60000;
}

function txFont(
  el: OoxmlNode | undefined,
  theme: ChartTheme,
  base: ChartFont,
): ChartFont {
  return runFont(defRPr(child(el, 'c:txPr')), theme, base);
}

function manualLayout(
  el: OoxmlNode | undefined,
): ChartManualLayout | undefined {
  const ml = child(child(el, 'c:layout'), 'c:manualLayout');
  if (!ml) return undefined;
  const num = (tag: string) => {
    const v = Number(attrOf(child(ml, tag), 'val'));
    return Number.isFinite(v) ? v : undefined;
  };
  const x = num('c:x');
  const y = num('c:y');
  if (x === undefined || y === undefined) return undefined;
  const mode = (tag: string) =>
    attrOf(child(ml, tag), 'val') === 'edge' ? 'edge' : 'factor';
  const target = attrOf(child(ml, 'c:layoutTarget'), 'val');
  return {
    ...(target === 'inner' || target === 'outer' ? { target } : {}),
    xMode: mode('c:xMode'),
    yMode: mode('c:yMode'),
    x,
    y,
    ...(num('c:w') !== undefined && { w: num('c:w') }),
    ...(num('c:h') !== undefined && { h: num('c:h') }),
  };
}

const flag = (el: OoxmlNode | undefined, fallback = false): boolean => {
  if (!el) return fallback;
  const v = attrOf(el, 'val');
  return v === undefined || v === '1' || v === 'true';
};
const valOf = (el: OoxmlNode | undefined, tag: string): string | undefined =>
  attrOf(child(el, tag), 'val');
const numOf = (el: OoxmlNode | undefined, tag: string): number | undefined => {
  const v = Number(valOf(el, tag));
  return Number.isFinite(v) ? v : undefined;
};

// ── data ────────────────────────────────────────────────────────────────────

/** A cache's points by index: c:strCache / c:numCache / c:strLit / c:numLit. */
function cachePoints(cache: OoxmlNode | undefined): {
  count: number;
  pts: Map<number, string>;
} {
  const pts = new Map<number, string>();
  for (const pt of children(cache, 'c:pt')) {
    const idx = Number(attrOf(pt, 'idx'));
    const v = child(pt, 'c:v')?.text;
    if (Number.isFinite(idx) && v !== undefined) pts.set(idx, String(v));
  }
  const declared = numOf(cache, 'c:ptCount');
  const count = declared ?? (pts.size ? Math.max(...pts.keys()) + 1 : 0);
  return { count, pts };
}

/** The cache under a data reference (…Ref → …Cache, or a literal). */
function dataCache(src: OoxmlNode | undefined): {
  cache?: OoxmlNode;
  numeric: boolean;
} {
  const numRef = child(src, 'c:numRef');
  if (numRef) return { cache: child(numRef, 'c:numCache'), numeric: true };
  const strRef = child(src, 'c:strRef');
  if (strRef) return { cache: child(strRef, 'c:strCache'), numeric: false };
  const multi = child(src, 'c:multiLvlStrRef');
  if (multi) {
    // The innermost level labels each point; outer levels group them.
    const lvl = child(child(multi, 'c:multiLvlStrCache'), 'c:lvl');
    return { cache: lvl, numeric: false };
  }
  const numLit = child(src, 'c:numLit');
  if (numLit) return { cache: numLit, numeric: true };
  return { cache: child(src, 'c:strLit'), numeric: false };
}

function numbers(src: OoxmlNode | undefined): {
  values: (number | null)[];
  formatCode: string;
  pointFormats?: Record<number, string>;
} {
  const { cache } = dataCache(src);
  const { count, pts } = cachePoints(cache);
  const values: (number | null)[] = [];
  for (let i = 0; i < count; i++) {
    const v = pts.get(i);
    const n = v === undefined ? NaN : Number(v);
    values.push(Number.isFinite(n) ? n : null);
  }
  // A cell formatted apart from its column keeps its own code on the point
  // (D-2609-LSTQ: a 0 in a "0.0%" series labelled "0%").
  const pointFormats: Record<number, string> = {};
  for (const pt of children(cache, 'c:pt')) {
    const code = attrOf(pt, 'formatCode');
    const idx = Number(attrOf(pt, 'idx'));
    if (code && Number.isFinite(idx)) pointFormats[idx] = code;
  }
  return {
    values,
    formatCode: child(cache, 'c:formatCode')?.text || 'General',
    ...(Object.keys(pointFormats).length > 0 && { pointFormats }),
  };
}

function labels(src: OoxmlNode | undefined): string[] {
  const { cache, numeric } = dataCache(src);
  const { count, pts } = cachePoints(cache);
  const fmt = child(cache, 'c:formatCode')?.text || 'General';
  const out: string[] = [];
  for (let i = 0; i < count; i++) {
    const v = pts.get(i) ?? '';
    const n = Number(v);
    out.push(
      numeric && v !== '' && Number.isFinite(n) ? formatNumber(n, fmt) : v,
    );
  }
  return out;
}

function seriesName(ser: OoxmlNode, fallbackIdx: number): string {
  const tx = child(ser, 'c:tx');
  const direct = child(tx, 'c:v')?.text;
  if (direct !== undefined) return String(direct);
  const { cache } = dataCache(tx);
  const { pts } = cachePoints(cache);
  return pts.get(0) ?? `Series ${fallbackIdx + 1}`;
}

// ── the parse ───────────────────────────────────────────────────────────────

type GroupType = ChartGroup['type'];
const GROUP_TAGS: Record<string, GroupType> = {
  'c:barChart': 'bar',
  'c:lineChart': 'line',
  'c:areaChart': 'area',
  'c:pieChart': 'pie',
  'c:doughnutChart': 'doughnut',
  'c:scatterChart': 'scatter',
  'c:radarChart': 'radar',
};
/** Plot types this renderer does not draw: the whole chart keeps the
 *  placeholder rather than painting part of it. */
const UNSUPPORTED = new Set([
  'c:bar3DChart',
  'c:line3DChart',
  'c:area3DChart',
  'c:pie3DChart',
  'c:surfaceChart',
  'c:surface3DChart',
  'c:bubbleChart',
  'c:stockChart',
  'c:ofPieChart',
]);

/** c:style, preferring the Office 2010 c14:style in an AlternateContent
 *  Choice (100 + the 2007 number) over its Fallback. Default 2. */
function chartStyleNumber(space: OoxmlNode | undefined): number {
  const direct = numOf(space, 'c:style');
  if (direct !== undefined) return direct;
  for (const alt of children(space, 'mc:AlternateContent')) {
    const c14 = numOf(child(alt, 'mc:Choice'), 'c14:style');
    if (c14 !== undefined) return c14 > 100 ? c14 - 100 : c14;
    const fb = numOf(child(alt, 'mc:Fallback'), 'c:style');
    if (fb !== undefined) return fb;
  }
  return 2;
}

function parseLabels(
  el: OoxmlNode | undefined,
  theme: ChartTheme,
  font: ChartFont,
  inherited: ChartDataLabels | null,
): ChartDataLabels | null {
  if (!el) return inherited;
  if (flag(child(el, 'c:delete'))) return null;
  const pick = (tag: string, from?: boolean) => {
    const node = child(el, tag);
    return node ? flag(node) : (from ?? false);
  };
  const deleted = children(el, 'c:dLbl')
    .filter((d) => flag(child(d, 'c:delete')))
    .map((d) => numOf(d, 'c:idx'))
    .filter((n): n is number => n !== undefined);
  const out: ChartDataLabels = {
    showVal: pick('c:showVal', inherited?.showVal),
    showPercent: pick('c:showPercent', inherited?.showPercent),
    showCatName: pick('c:showCatName', inherited?.showCatName),
    showSerName: pick('c:showSerName', inherited?.showSerName),
    separator: child(el, 'c:separator')?.text ?? inherited?.separator ?? ', ',
    font: txFont(el, theme, inherited?.font ?? font),
  };
  const pos = valOf(el, 'c:dLblPos') ?? inherited?.position;
  if (pos) out.position = pos as ChartDataLabels['position'];
  const numFmt = child(el, 'c:numFmt');
  if (numFmt && attrOf(numFmt, 'sourceLinked') !== '1')
    out.formatCode = attrOf(numFmt, 'formatCode');
  else if (inherited?.formatCode) out.formatCode = inherited.formatCode;
  if (deleted.length) out.deleted = deleted;
  if (!out.showVal && !out.showPercent && !out.showCatName && !out.showSerName)
    return deleted.length ? out : null;
  return out;
}

function parseMarker(
  el: OoxmlNode | undefined,
  theme: ChartTheme,
  auto: { symbol: ChartMarkerSymbol; fill: string; line: ChartLine } | null,
): ChartMarker | null {
  const symbol = valOf(el, 'c:symbol');
  if (symbol === 'none') return null;
  if (!el && !auto) return null;
  const spPr = child(el, 'c:spPr');
  const resolved = (symbol && symbol !== 'auto' ? symbol : auto?.symbol) as
    | ChartMarkerSymbol
    | undefined;
  if (!resolved) return null;
  const fill = spFill(spPr, theme);
  return {
    symbol: resolved,
    size: numOf(el, 'c:size') ?? 7,
    fill: fill === undefined ? (auto?.fill ?? null) : fill,
    line: spLine(spPr, theme, auto?.line ?? null),
  };
}

export function parseChartSpec(
  xml: string,
  theme: ChartTheme,
): ChartSpec | null {
  const space = child(parseXml(xml), 'c:chartSpace');
  const chart = child(space, 'c:chart');
  const plot = child(chart, 'c:plotArea');
  if (!plot) return null;
  if (plot.children.some((c) => UNSUPPORTED.has(c.name))) return null;

  const style = chartStyleNumber(space);
  const d = styleDefaults(theme, style);
  const minor = theme.font('minor') ?? 'Calibri';
  // The chart-wide text, then each element's default on top of it.
  const baseFont = txFont(space, theme, {
    family: minor,
    sizePt: 10,
    bold: false,
    italic: false,
    color: d.text,
  });
  const spaceSize = defRPr(child(space, 'c:txPr'))
    ? baseFont.sizePt
    : undefined;

  // Series colours cycle by c:idx across the WHOLE chart, not per group.
  const allSeries = plot.children
    .filter((c) => GROUP_TAGS[c.name])
    .flatMap((g) => children(g, 'c:ser'));
  const seriesCount =
    Math.max(-1, ...allSeries.map((s, i) => numOf(s, 'c:idx') ?? i)) + 1;

  const groups: ChartGroup[] = [];
  let order = 0;
  for (const g of plot.children) {
    const type = GROUP_TAGS[g.name];
    if (!type) continue;
    const grouping = (valOf(g, 'c:grouping') ??
      (type === 'bar' ? 'clustered' : 'standard')) as ChartGroup['grouping'];
    const varyColors = flag(
      child(g, 'c:varyColors'),
      type === 'pie' || type === 'doughnut',
    );
    const filled =
      type === 'bar' ||
      type === 'area' ||
      type === 'pie' ||
      type === 'doughnut';
    const groupMarkers =
      type === 'line'
        ? flag(child(g, 'c:marker'), true)
        : type === 'scatter' || type === 'radar';
    const scatterStyle = valOf(g, 'c:scatterStyle');
    const radarStyle = valOf(g, 'c:radarStyle');
    const groupLabels = parseLabels(child(g, 'c:dLbls'), theme, baseFont, null);
    const series: ChartSeries[] = [];
    for (const ser of children(g, 'c:ser')) {
      const idx = numOf(ser, 'c:idx') ?? order;
      order++;
      const auto = seriesColor(theme, style, idx, seriesCount);
      const spPr = child(ser, 'c:spPr');
      const xy = type === 'scatter';
      const { values, formatCode, pointFormats } = numbers(
        child(ser, xy ? 'c:yVal' : 'c:val'),
      );
      const cats = xy ? [] : labels(child(ser, 'c:cat'));
      const sFill = spFill(spPr, theme);
      const fill = filled
        ? sFill === undefined
          ? d.seriesFill(auto)
          : sFill
        : null;
      const defaultLine = filled
        ? d.seriesOutline(auto, idx)
        : d.seriesLine(auto);
      let line = spLine(spPr, theme, defaultLine);
      // A scatter "marker only" style, and a filled radar, draw no line
      // unless the series says so.
      if (
        xy &&
        (scatterStyle === 'marker' || scatterStyle === 'none') &&
        !child(spPr, 'a:ln')
      )
        line = null;
      if (type === 'radar' && radarStyle === 'filled' && !child(spPr, 'a:ln'))
        line = null;
      const markerAuto =
        groupMarkers && !(type === 'radar' && radarStyle !== 'marker')
          ? {
              symbol: MARKER_CYCLE[idx % MARKER_CYCLE.length],
              ...d.marker(auto),
            }
          : null;
      const marker = filled
        ? null
        : parseMarker(child(ser, 'c:marker'), theme, markerAuto);
      const points: ChartPoint[] = [];
      for (const dPt of children(ser, 'c:dPt')) {
        const pIdx = numOf(dPt, 'c:idx');
        if (pIdx === undefined) continue;
        const pSp = child(dPt, 'c:spPr');
        const pFill = spFill(pSp, theme);
        const p: ChartPoint = { idx: pIdx };
        if (pFill !== undefined) p.fill = pFill;
        if (child(pSp, 'a:ln')) p.line = spLine(pSp, theme, line);
        const ex = numOf(dPt, 'c:explosion');
        if (ex !== undefined) p.explosion = ex;
        points.push(p);
      }
      // A pie (or any varyColors group of one series) colours every point.
      if (varyColors && filled && children(g, 'c:ser').length === 1) {
        const n = values.length;
        for (let i = 0; i < n; i++) {
          if (points.some((p) => p.idx === i && p.fill !== undefined)) continue;
          const c = seriesColor(theme, style, i, n);
          const existing = points.find((p) => p.idx === i);
          if (existing) existing.fill = d.seriesFill(c);
          else
            points.push({
              idx: i,
              fill: d.seriesFill(c),
              line: d.seriesOutline(c, i),
            });
        }
        points.sort((a, b) => a.idx - b.idx);
      }
      const s: ChartSeries = {
        name: seriesName(ser, idx),
        values,
        categories: cats,
        formatCode,
        fill,
        line,
        marker,
        smooth: flag(child(ser, 'c:smooth')),
        labels: parseLabels(
          child(ser, 'c:dLbls'),
          theme,
          baseFont,
          groupLabels,
        ),
      };
      if (xy) s.xValues = numbers(child(ser, 'c:xVal')).values;
      if (pointFormats) s.pointFormats = pointFormats;
      if (points.length) s.points = points;
      const ex = numOf(ser, 'c:explosion');
      if (ex !== undefined) s.explosion = ex;
      series.push(s);
    }
    const group: ChartGroup = {
      type,
      grouping,
      varyColors,
      gapWidth: numOf(g, 'c:gapWidth') ?? 150,
      overlap:
        numOf(g, 'c:overlap') ??
        (grouping === 'stacked' || grouping === 'percentStacked' ? 100 : 0),
      axisIds: children(g, 'c:axId')
        .map((a) => Number(attrOf(a, 'val')))
        .filter(Number.isFinite),
      series,
    };
    if (type === 'bar')
      group.barDir = valOf(g, 'c:barDir') === 'bar' ? 'bar' : 'col';
    if (type === 'doughnut') group.holeSize = numOf(g, 'c:holeSize') ?? 50;
    if (type === 'pie' || type === 'doughnut')
      group.firstSliceAng = numOf(g, 'c:firstSliceAng') ?? 0;
    groups.push(group);
  }
  if (groups.length === 0) return null;

  // Text for an automatic title: the one series' name, else Word's default.
  const onlySeries = groups.flatMap((g) => g.series);
  const autoTitle =
    onlySeries.length === 1 ? onlySeries[0].name : 'Chart Title';
  const titleFont: ChartFont = {
    ...baseFont,
    sizePt: spaceSize !== undefined ? spaceSize * 1.2 : 18,
    bold: true,
  };

  const axes: ChartAxis[] = [];
  for (const ax of plot.children) {
    if (!['c:catAx', 'c:valAx', 'c:dateAx'].includes(ax.name)) continue;
    const kind =
      ax.name === 'c:valAx' ? 'val' : ax.name === 'c:dateAx' ? 'date' : 'cat';
    const scaling = child(ax, 'c:scaling');
    const numFmt = child(ax, 'c:numFmt');
    const crossesAt = numOf(ax, 'c:crossesAt');
    const gridOf = (tag: string, fallback: ChartLine) => {
      const g = child(ax, tag);
      return g ? spLine(child(g, 'c:spPr'), theme, fallback) : null;
    };
    const spPr = child(ax, 'c:spPr');
    const axis: ChartAxis = {
      id: numOf(ax, 'c:axId') ?? 0,
      kind,
      position: (valOf(ax, 'c:axPos') ??
        (kind === 'val' ? 'l' : 'b')) as ChartAxis['position'],
      deleted: flag(child(ax, 'c:delete')),
      crossAxisId: numOf(ax, 'c:crossAx') ?? 0,
      crosses:
        crossesAt ??
        ((valOf(ax, 'c:crosses') ?? 'autoZero') as 'autoZero' | 'min' | 'max'),
      crossBetween:
        valOf(ax, 'c:crossBetween') === 'midCat' ? 'midCat' : 'between',
      reversed: valOf(scaling, 'c:orientation') === 'maxMin',
      formatCode: attrOf(numFmt, 'formatCode') ?? 'General',
      sourceLinked: attrOf(numFmt, 'sourceLinked') === '1' || !numFmt,
      majorGridlines: gridOf('c:majorGridlines', d.majorGrid),
      minorGridlines: gridOf('c:minorGridlines', d.minorGrid),
      line: spLine(spPr, theme, d.axisLine),
      majorTickMark: (valOf(ax, 'c:majorTickMark') ??
        'out') as ChartAxis['majorTickMark'],
      minorTickMark: (valOf(ax, 'c:minorTickMark') ??
        'none') as ChartAxis['minorTickMark'],
      tickLabelPosition: (valOf(ax, 'c:tickLblPos') ??
        'nextTo') as ChartAxis['tickLabelPosition'],
      font: txFont(ax, theme, baseFont),
      title: null,
    };
    const min = numOf(scaling, 'c:min');
    const max = numOf(scaling, 'c:max');
    const logBase = numOf(scaling, 'c:logBase');
    const majorUnit = numOf(ax, 'c:majorUnit');
    const minorUnit = numOf(ax, 'c:minorUnit');
    if (min !== undefined) axis.min = min;
    if (max !== undefined) axis.max = max;
    if (logBase !== undefined) axis.logBase = logBase;
    if (majorUnit !== undefined) axis.majorUnit = majorUnit;
    if (minorUnit !== undefined) axis.minorUnit = minorUnit;
    const rot = bodyRotation(child(ax, 'c:txPr'));
    if (rot !== undefined) axis.labelRotation = rot;
    const t = child(ax, 'c:title');
    if (t)
      axis.title = parseTitle(
        t,
        theme,
        { ...baseFont, bold: true },
        'Axis Title',
        {
          fill: null,
          line: null,
        },
      );
    axes.push(axis);
  }

  // A c:title is shown — with its own text, or the automatic one when it
  // has none. c:autoTitleDeleted only says no title was ever wanted.
  const titleEl = child(chart, 'c:title');
  const title = titleEl
    ? parseTitle(titleEl, theme, titleFont, autoTitle, {
        fill: null,
        line: null,
      })
    : null;

  const legendEl = child(chart, 'c:legend');
  let legend: ChartLegend | null = null;
  if (legendEl) {
    legend = {
      position: (valOf(legendEl, 'c:legendPos') ??
        'r') as ChartLegend['position'],
      overlay: flag(child(legendEl, 'c:overlay')),
      font: txFont(legendEl, theme, baseFont),
      box: spBox(child(legendEl, 'c:spPr'), theme, { fill: null, line: null }),
    };
    const lay = manualLayout(legendEl);
    if (lay) legend.layout = lay;
    const deleted = children(legendEl, 'c:legendEntry')
      .filter((e) => flag(child(e, 'c:delete')))
      .map((e) => numOf(e, 'c:idx'))
      .filter((n): n is number => n !== undefined);
    if (deleted.length) legend.deleted = deleted;
  }

  const plotBox = spBox(child(plot, 'c:spPr'), theme, {
    fill: d.plotFill,
    line: null,
  });
  const plotLayout = manualLayout(plot);
  const dispBlanks = valOf(chart, 'c:dispBlanksAs');
  return {
    groups,
    axes,
    title,
    legend,
    chartArea: spBox(child(space, 'c:spPr'), theme, {
      fill: d.chartFill,
      line: d.chartLine,
    }),
    plotArea: plotLayout ? { ...plotBox, layout: plotLayout } : plotBox,
    // The schema default is true; a 2007 chart that wants square corners
    // says roundedCorners val="0".
    roundedCorners: flag(child(space, 'c:roundedCorners'), true),
    dispBlanksAs:
      dispBlanks === 'zero' || dispBlanks === 'span' ? dispBlanks : 'gap',
  };
}

function parseTitle(
  el: OoxmlNode,
  theme: ChartTheme,
  base: ChartFont,
  autoText: string,
  boxDefault: ChartBox,
): ChartText {
  const rich = child(child(el, 'c:tx'), 'c:rich');
  let font = txFont(el, theme, base);
  let text = autoText;
  let rotation = bodyRotation(child(el, 'c:txPr'));
  if (rich) {
    font = runFont(defRPr(rich), theme, font);
    const paras = children(rich, 'a:p');
    const firstRun = child(child(paras[0], 'a:r'), 'a:rPr');
    font = runFont(firstRun, theme, font);
    text = paras
      .map((p) =>
        children(p, 'a:r')
          .map((r) => child(r, 'a:t')?.text ?? '')
          .join(''),
      )
      .join('\n');
    rotation = bodyRotation(rich) ?? rotation;
  } else {
    // A title bound to a cell (c:strRef) shows the cached text.
    const ref = child(child(el, 'c:tx'), 'c:strRef');
    if (ref) text = labels(child(el, 'c:tx')).join(' ') || autoText;
  }
  const out: ChartText = {
    text,
    font,
    overlay: flag(child(el, 'c:overlay')),
    box: spBox(child(el, 'c:spPr'), theme, boxDefault),
  };
  if (rotation !== undefined) out.rotation = rotation;
  const lay = manualLayout(el);
  if (lay) out.layout = lay;
  return out;
}
