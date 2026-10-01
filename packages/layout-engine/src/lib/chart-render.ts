/**
 * A {@link ChartSpec} laid out and drawn as a vector display list.
 *
 * Word lays an automatically-placed chart out itself — nothing in the file
 * says where the plot area, legend or labels go — so the layout here
 * follows what Word 365 was MEASURED to do (PDFs of D-2609-LSTQ and the
 * chart probe, read with PyMuPDF); the constants below carry those numbers.
 * A chart with c:manualLayout gets the file's fractions instead.
 *
 * The display list is in POINTS of the chart frame (the spec's own unit);
 * the painter scales it to the box. Text is measured with the layout's own
 * measurer, so a label sits where the painter will draw it.
 *
 * Pure: spec + size + measurer in, ops out.
 */
import type {
  ChartAxis,
  ChartBox,
  ChartFont,
  ChartGroup,
  ChartLine,
  ChartManualLayout,
  ChartMarker,
  ChartSeries,
  ChartSpec,
  FontSpec,
  MeasureText,
  VectorImageSpec,
  VectorOp,
} from '@shadow-garden/bapbong-contracts';
import { formatNumber } from '@shadow-garden/bapbong-contracts';
import { autoScale, ticks, type AxisScale } from './chart-axis.js';

const PX_PER_PT = 96 / 72;

// ── measured layout constants (points) ─────────────────────────────────────
/** Chart edge → content, left/right and bottom. */
const PAD_X = 6.5;
const PAD_BOTTOM = 6.4;
/** Chart edge → the top value label (or the title). */
const PAD_TOP = 4.2;
/** Title → what is below it. */
const TITLE_GAP = 6;
/** Value tick label → plot edge. */
const VAL_LABEL_GAP = 9.2;
/** Category axis → its labels. */
const CAT_LABEL_GAP = 6.4;
/** Axis title → tick labels. */
const AXIS_TITLE_GAP = 4;
/** Widest a side legend may grow, as a fraction of the chart width. */
const LEGEND_MAX = 0.282;
/** Legend frame inset, and legend → plot gap. */
const LEGEND_PAD = 3;
const LEGEND_GAP = 6;
/** A line series' legend key and the gap before its text. */
const LINE_KEY = 19.2;
const KEY_GAP = 2.2;
/** A filled series' legend square, in ems of the legend font. */
const BOX_KEY_EM = 0.55;
/** Gap between legend entries laid out in a row (top/bottom legends). */
const LEGEND_ROW_GAP = 10;
/** Line box of chart text, in ems (Calibri's ascender+descender+gap). */
const LINE_HEIGHT = 1.22;
/** Major tick mark length. */
const TICK = 3.2;
/** Data label → the point/marker edge it labels. */
const LABEL_GAP = 6;
/** Data label above a column's end. */
const BAR_LABEL_GAP = 2;
/** Corner radius of a rounded chart area. */
const CORNER = 8;

type Box = { x: number; y: number; w: number; h: number };

const fontSpec = (f: ChartFont): FontSpec => ({
  family: f.family,
  sizePt: f.sizePt,
  bold: f.bold,
  italic: f.italic,
});

class Ctx {
  readonly ops: VectorOp[] = [];
  constructor(private readonly measure: MeasureText) {}
  width(text: string, f: ChartFont): number {
    return text ? this.measure(text, fontSpec(f)) / PX_PER_PT : 0;
  }
  lineHeight(f: ChartFont): number {
    return f.sizePt * LINE_HEIGHT;
  }
  /** Greedy word wrap; hard breaks kept. */
  wrap(text: string, f: ChartFont, maxW: number): string[] {
    const out: string[] = [];
    for (const para of text.split('\n')) {
      const words = para.split(/(\s+)/);
      let line = '';
      for (const w of words) {
        const next = line + w;
        if (line.trim() && this.width(next.trimEnd(), f) > maxW && w.trim()) {
          out.push(line.trimEnd());
          line = w.trimStart();
        } else line = next;
      }
      out.push(line.trimEnd());
    }
    return out;
  }
  /** One line of text; `top` is the line box's top. */
  text(
    s: string,
    f: ChartFont,
    x: number,
    top: number,
    align: 'left' | 'center' | 'right' = 'left',
    rotation?: number,
  ): void {
    if (!s) return;
    const w = this.width(s, f);
    const left = align === 'center' ? x - w / 2 : align === 'right' ? x - w : x;
    this.ops.push({
      kind: 'text',
      x: left,
      // Baseline: the line box less the descender (a quarter em).
      y: top + this.lineHeight(f) - f.sizePt * 0.25,
      text: s,
      size: f.sizePt,
      family: f.family,
      ...(f.bold && { bold: true }),
      ...(f.italic && { italic: true }),
      color: f.color,
      ...(rotation && { rotation }),
    });
  }
  line(
    x1: number,
    y1: number,
    x2: number,
    y2: number,
    l: ChartLine | null,
  ): void {
    if (!l) return;
    this.ops.push({
      kind: 'line',
      x1,
      y1,
      x2,
      y2,
      width: l.widthPt,
      color: l.color,
      ...(l.dash && { dash: l.dash }),
    });
  }
  poly(
    points: { x: number; y: number }[],
    fill: string | null,
    line: ChartLine | null,
  ): void {
    if (points.length < 2 || (!fill && !line)) return;
    this.ops.push({
      kind: 'polygon',
      points,
      ...(fill && { fill }),
      ...(line && { stroke: line.color, strokeWidth: line.widthPt }),
    });
  }
  polyline(
    points: { x: number; y: number }[],
    l: ChartLine | null,
    round = true,
  ): void {
    if (!l || points.length < 2) return;
    this.ops.push({
      kind: 'polyline',
      points,
      stroke: l.color,
      strokeWidth: l.widthPt,
      ...(round && { join: 'round' as const, cap: 'round' as const }),
      ...(l.dash && { dash: l.dash }),
    });
  }
  rect(b: Box, box: ChartBox): void {
    this.poly(rectPts(b), box.fill, box.line);
  }
}

const rectPts = (b: Box) => [
  { x: b.x, y: b.y },
  { x: b.x + b.w, y: b.y },
  { x: b.x + b.w, y: b.y + b.h },
  { x: b.x, y: b.y + b.h },
];

function roundRectPts(b: Box, r: number): { x: number; y: number }[] {
  const rr = Math.min(r, b.w / 2, b.h / 2);
  const pts: { x: number; y: number }[] = [];
  const arc = (cx: number, cy: number, a0: number) => {
    for (let i = 0; i <= 6; i++) {
      const a = a0 + (i / 6) * (Math.PI / 2);
      pts.push({ x: cx + rr * Math.cos(a), y: cy + rr * Math.sin(a) });
    }
  };
  arc(b.x + b.w - rr, b.y + rr, -Math.PI / 2);
  arc(b.x + b.w - rr, b.y + b.h - rr, 0);
  arc(b.x + rr, b.y + b.h - rr, Math.PI / 2);
  arc(b.x + rr, b.y + rr, Math.PI);
  return pts;
}

/** A manual-layout box (edge mode = fractions of the chart; factor mode
 *  offsets the automatic box). */
function placed(
  layout: ChartManualLayout,
  auto: Box,
  W: number,
  H: number,
): Box {
  const x = layout.xMode === 'edge' ? layout.x * W : auto.x + layout.x * W;
  const y = layout.yMode === 'edge' ? layout.y * H : auto.y + layout.y * H;
  return {
    x,
    y,
    w: layout.w !== undefined ? layout.w * W : auto.w,
    h: layout.h !== undefined ? layout.h * H : auto.h,
  };
}

// ── legend ──────────────────────────────────────────────────────────────────

interface LegendEntry {
  name: string;
  fill: string | null;
  line: ChartLine | null;
  marker: ChartMarker | null;
  /** Drawn as a line key (line/scatter/radar series) or a square. */
  lineKey: boolean;
}

function legendEntries(spec: ChartSpec): LegendEntry[] {
  const out: LegendEntry[] = [];
  for (const g of spec.groups) {
    const pieLike = g.type === 'pie' || g.type === 'doughnut';
    if (
      pieLike ||
      (g.varyColors && g.series.length === 1 && g.type !== 'line')
    ) {
      const s = g.series[0];
      if (!s) continue;
      s.categories.forEach((c, i) => {
        const p = s.points?.find((pt) => pt.idx === i);
        out.push({
          name: c,
          fill: p?.fill ?? s.fill,
          line: p?.line ?? s.line,
          marker: null,
          lineKey: false,
        });
      });
      continue;
    }
    const lineKey =
      g.type === 'line' || g.type === 'scatter' || g.type === 'radar';
    for (const s of g.series)
      out.push({
        name: s.name,
        fill: s.fill,
        line: s.line,
        marker: s.marker,
        lineKey,
      });
  }
  const del = new Set(spec.legend?.deleted ?? []);
  return out.filter((_, i) => !del.has(i));
}

interface LegendLayout {
  box: Box;
  draw: () => void;
}

function layoutLegend(
  c: Ctx,
  spec: ChartSpec,
  area: Box,
  W: number,
): LegendLayout | null {
  const lg = spec.legend;
  if (!lg) return null;
  const entries = legendEntries(spec);
  if (entries.length === 0) return null;
  const f = lg.font;
  const lh = c.lineHeight(f);
  const keyW = (e: LegendEntry) =>
    e.lineKey ? LINE_KEY : f.sizePt * BOX_KEY_EM;
  const vertical =
    lg.position === 'r' || lg.position === 'l' || lg.position === 'tr';
  type Placed = { e: LegendEntry; lines: string[]; w: number; h: number };
  let placedEntries: Placed[];
  let box: Box;
  if (vertical) {
    const maxW = LEGEND_MAX * W;
    placedEntries = entries.map((e) => {
      const kw = keyW(e) + KEY_GAP;
      const lines = c.wrap(e.name, f, Math.max(10, maxW - kw - 2 * LEGEND_PAD));
      const tw = Math.max(...lines.map((l) => c.width(l, f)));
      return { e, lines, w: kw + tw, h: lines.length * lh };
    });
    const w = Math.max(...placedEntries.map((p) => p.w)) + 2 * LEGEND_PAD;
    const h = placedEntries.reduce((s, p) => s + p.h, 0) + 2 * LEGEND_PAD;
    const x = lg.position === 'l' ? area.x : area.x + area.w - w;
    const y = lg.position === 'tr' ? area.y : area.y + (area.h - h) / 2;
    box = { x, y, w, h };
  } else {
    placedEntries = entries.map((e) => {
      const kw = keyW(e) + KEY_GAP;
      const lines = c.wrap(e.name, f, area.w * 0.8);
      const tw = Math.max(...lines.map((l) => c.width(l, f)));
      return { e, lines, w: kw + tw, h: lines.length * lh };
    });
    // Rows: as many entries as fit side by side.
    const rows: Placed[][] = [[]];
    let rowW = 0;
    for (const p of placedEntries) {
      const add = (rows[rows.length - 1].length ? LEGEND_ROW_GAP : 0) + p.w;
      if (
        rowW + add > area.w - 2 * LEGEND_PAD &&
        rows[rows.length - 1].length
      ) {
        rows.push([p]);
        rowW = p.w;
      } else {
        rows[rows.length - 1].push(p);
        rowW += add;
      }
    }
    const rowWidths = rows.map(
      (r) => r.reduce((s, p) => s + p.w, 0) + LEGEND_ROW_GAP * (r.length - 1),
    );
    const rowHeights = rows.map((r) => Math.max(...r.map((p) => p.h)));
    const w = Math.max(...rowWidths) + 2 * LEGEND_PAD;
    const h = rowHeights.reduce((s, v) => s + v, 0) + 2 * LEGEND_PAD;
    const x = area.x + (area.w - w) / 2;
    const y = lg.position === 't' ? area.y : area.y + area.h - h;
    box = { x, y, w, h };
    const auto = box;
    if (lg.layout) box = placed(lg.layout, auto, W, area.h);
    return {
      box,
      draw: () => {
        c.rect(box, lg.box);
        let ry = box.y + LEGEND_PAD;
        rows.forEach((r, ri) => {
          let rx = box.x + (box.w - rowWidths[ri]) / 2;
          for (const p of r) {
            drawEntry(
              c,
              p.e,
              p.lines,
              f,
              rx,
              ry + (rowHeights[ri] - p.h) / 2,
              keyW(p.e),
            );
            rx += p.w + LEGEND_ROW_GAP;
          }
          ry += rowHeights[ri];
        });
      },
    };
  }
  if (lg.layout) box = placed(lg.layout, box, W, area.h);
  return {
    box,
    draw: () => {
      c.rect(box, lg.box);
      let y = box.y + LEGEND_PAD;
      for (const p of placedEntries) {
        drawEntry(c, p.e, p.lines, f, box.x + LEGEND_PAD, y, keyW(p.e));
        y += p.h;
      }
    },
  };
}

function drawEntry(
  c: Ctx,
  e: LegendEntry,
  lines: string[],
  f: ChartFont,
  x: number,
  top: number,
  kw: number,
): void {
  const lh = c.lineHeight(f);
  const mid = top + (lines.length * lh) / 2;
  if (e.lineKey) {
    c.line(x, mid, x + kw, mid, e.line);
    if (e.marker) drawMarker(c, e.marker, x + kw / 2, mid);
  } else {
    const s = kw;
    c.rect({ x, y: mid - s / 2, w: s, h: s }, { fill: e.fill, line: e.line });
  }
  lines.forEach((l, i) => c.text(l, f, x + kw + KEY_GAP, top + i * lh));
}

// ── markers ─────────────────────────────────────────────────────────────────

function circlePts(cx: number, cy: number, r: number, n = 16) {
  return Array.from({ length: n }, (_, i) => {
    const a = (i / n) * Math.PI * 2;
    return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
  });
}

function drawMarker(c: Ctx, m: ChartMarker, x: number, y: number): void {
  const r = m.size / 2;
  const box = { fill: m.fill, line: m.line };
  switch (m.symbol) {
    case 'square':
      return c.rect({ x: x - r, y: y - r, w: m.size, h: m.size }, box);
    case 'diamond':
      return c.poly(
        [
          { x, y: y - r },
          { x: x + r, y },
          { x, y: y + r },
          { x: x - r, y },
        ],
        box.fill,
        box.line,
      );
    case 'triangle':
      return c.poly(
        [
          { x, y: y - r },
          { x: x + r, y: y + r },
          { x: x - r, y: y + r },
        ],
        box.fill,
        box.line,
      );
    case 'circle':
      return c.poly(circlePts(x, y, r), box.fill, box.line);
    case 'dot':
      return c.poly(
        circlePts(x, y, r / 2.5, 10),
        box.fill ?? box.line?.color ?? null,
        null,
      );
    case 'dash':
      return c.rect({ x: x - r, y: y - r / 4, w: m.size, h: r / 2 }, box);
    case 'x':
    case 'star':
    case 'plus': {
      const l = m.line ?? (m.fill ? { color: m.fill, widthPt: 0.75 } : null);
      if (m.symbol !== 'plus') {
        c.line(x - r, y - r, x + r, y + r, l);
        c.line(x - r, y + r, x + r, y - r, l);
      }
      if (m.symbol !== 'x') {
        c.line(x, y - r, x, y + r, l);
        if (m.symbol === 'plus') c.line(x - r, y, x + r, y, l);
      }
      return;
    }
    default:
      return;
  }
}

// ── data ────────────────────────────────────────────────────────────────────

const isStacked = (g: ChartGroup) =>
  g.grouping === 'stacked' || g.grouping === 'percentStacked';
const nCategories = (g: ChartGroup) =>
  Math.max(
    0,
    ...g.series.map((s) => Math.max(s.values.length, s.categories.length)),
  );

/** The values a group plots, after stacking/percent: [series][cat] → [from, to]. */
function plotted(
  g: ChartGroup,
  blanksZero: boolean,
): ([number, number] | null)[][] {
  const n = nCategories(g);
  const pos = new Array<number>(n).fill(0);
  const neg = new Array<number>(n).fill(0);
  const totals = new Array<number>(n).fill(0);
  if (g.grouping === 'percentStacked')
    for (const s of g.series)
      for (let i = 0; i < n; i++) totals[i] += Math.abs(s.values[i] ?? 0);
  return g.series.map((s) =>
    Array.from({ length: n }, (_, i) => {
      let v = s.values[i];
      if (v == null) {
        if (!blanksZero && !isStacked(g)) return null;
        v = 0;
      }
      if (g.grouping === 'percentStacked') v = totals[i] ? v / totals[i] : 0;
      if (!isStacked(g)) return [0, v] as [number, number];
      if (v >= 0) {
        const from = pos[i];
        pos[i] += v;
        return [from, pos[i]] as [number, number];
      }
      const from = neg[i];
      neg[i] += v;
      return [from, neg[i]] as [number, number];
    }),
  );
}

function dataRange(
  groups: ChartGroup[],
  blanksZero: boolean,
): [number, number] {
  let lo = Infinity;
  let hi = -Infinity;
  for (const g of groups)
    for (const s of plotted(g, blanksZero))
      for (const p of s) {
        if (!p) continue;
        lo = Math.min(lo, p[1], isStacked(g) ? p[0] : p[1]);
        hi = Math.max(hi, p[1], isStacked(g) ? p[0] : p[1]);
      }
  return [lo === Infinity ? 0 : lo, hi === -Infinity ? 1 : hi];
}

function scaleFor(
  axis: ChartAxis | undefined,
  groups: ChartGroup[],
  blanksZero: boolean,
): AxisScale {
  if (groups.some((g) => g.grouping === 'percentStacked')) {
    const hasNeg = groups.some((g) =>
      g.series.some((s) => s.values.some((v) => (v ?? 0) < 0)),
    );
    return autoScale(hasNeg ? -1 : 0, 1, {
      min: axis?.min ?? (hasNeg ? -1 : 0),
      max: axis?.max ?? 1,
      ...(axis?.majorUnit !== undefined && { major: axis.majorUnit }),
    });
  }
  const [lo, hi] = dataRange(groups, blanksZero);
  return autoScale(lo, hi, {
    ...(axis?.min !== undefined && { min: axis.min }),
    ...(axis?.max !== undefined && { max: axis.max }),
    ...(axis?.majorUnit !== undefined && { major: axis.majorUnit }),
  });
}

function xRange(groups: ChartGroup[]): [number, number] {
  let lo = Infinity;
  let hi = -Infinity;
  for (const g of groups)
    for (const s of g.series)
      (s.xValues ?? s.values.map((_, i) => i + 1)).forEach((x) => {
        if (x == null) return;
        lo = Math.min(lo, x);
        hi = Math.max(hi, x);
      });
  return [lo === Infinity ? 0 : lo, hi === -Infinity ? 1 : hi];
}

/** A data label's text for one point. */
function labelText(
  s: ChartSeries,
  i: number,
  value: number,
  percent: number | null,
): string {
  const lb = s.labels;
  if (!lb || lb.deleted?.includes(i)) return '';
  const parts: string[] = [];
  if (lb.showSerName) parts.push(s.name);
  if (lb.showCatName) parts.push(s.categories[i] ?? String(i + 1));
  if (lb.showVal)
    parts.push(
      formatNumber(value, lb.formatCode ?? s.pointFormats?.[i] ?? s.formatCode),
    );
  if (lb.showPercent && percent !== null)
    parts.push(formatNumber(percent, '0%'));
  return parts.join(lb.separator);
}

/** Smooth a polyline (Catmull-Rom through the points), sampled. */
function smoothPts(
  pts: { x: number; y: number }[],
): { x: number; y: number }[] {
  if (pts.length < 3) return pts;
  const out: { x: number; y: number }[] = [];
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] ?? pts[i];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[i + 2] ?? p2;
    for (let t = 0; t < 1; t += 1 / 10) {
      const t2 = t * t;
      const t3 = t2 * t;
      const f = (a: number, b: number, cc: number, d: number) =>
        0.5 *
        (2 * b +
          (-a + cc) * t +
          (2 * a - 5 * b + 4 * cc - d) * t2 +
          (-a + 3 * b - 3 * cc + d) * t3);
      out.push({ x: f(p0.x, p1.x, p2.x, p3.x), y: f(p0.y, p1.y, p2.y, p3.y) });
    }
  }
  out.push(pts[pts.length - 1]);
  return out;
}

// ── the chart ───────────────────────────────────────────────────────────────

export function renderChart(
  spec: ChartSpec,
  widthPx: number,
  heightPx: number,
  measure: MeasureText,
): VectorImageSpec {
  const W = widthPx / PX_PER_PT;
  const H = heightPx / PX_PER_PT;
  const c = new Ctx(measure);
  const frame: Box = { x: 0, y: 0, w: W, h: H };
  // Chart area.
  if (spec.roundedCorners)
    c.poly(
      roundRectPts(frame, CORNER),
      spec.chartArea.fill,
      spec.chartArea.line,
    );
  else c.rect(frame, spec.chartArea);

  let area: Box = {
    x: PAD_X,
    y: PAD_TOP,
    w: W - 2 * PAD_X,
    h: H - PAD_TOP - PAD_BOTTOM,
  };

  // Title.
  let drawTitle: (() => void) | null = null;
  const t = spec.title;
  if (t && t.text) {
    const lines = c.wrap(t.text, t.font, W * 0.8);
    const lh = c.lineHeight(t.font);
    const h = lines.length * lh;
    let top = area.y + 3.9;
    let cx = W / 2;
    if (t.layout) {
      const b = placed(t.layout, { x: 0, y: top, w: W, h }, W, H);
      top = b.y;
      cx = b.x + Math.max(...lines.map((l) => c.width(l, t.font))) / 2;
    }
    drawTitle = () =>
      lines.forEach((l, i) => c.text(l, t.font, cx, top + i * lh, 'center'));
    if (!t.overlay && !t.layout) {
      const used = top + h + TITLE_GAP - area.y;
      area = { ...area, y: area.y + used, h: area.h - used };
    }
  }

  // Legend.
  const legend = layoutLegend(c, spec, area, W);
  if (legend && !spec.legend?.overlay && !spec.legend?.layout) {
    const b = legend.box;
    switch (spec.legend?.position) {
      case 'l':
        area = {
          ...area,
          x: b.x + b.w + LEGEND_GAP,
          w: area.w - b.w - LEGEND_GAP,
        };
        break;
      case 't':
        area = {
          ...area,
          y: b.y + b.h + LEGEND_GAP,
          h: area.h - b.h - LEGEND_GAP,
        };
        break;
      case 'b':
        area = { ...area, h: b.y - LEGEND_GAP - area.y };
        break;
      default:
        area = { ...area, w: b.x - LEGEND_GAP - area.x };
    }
  }

  const pieGroups = spec.groups.filter(
    (g) => g.type === 'pie' || g.type === 'doughnut',
  );
  const radar = spec.groups.filter((g) => g.type === 'radar');
  if (pieGroups.length && pieGroups.length === spec.groups.length) {
    drawPies(c, spec, pieGroups, area);
  } else if (radar.length && radar.length === spec.groups.length) {
    drawRadar(c, spec, radar, area);
  } else {
    drawCartesian(c, spec, area, W, H);
  }
  legend?.draw();
  drawTitle?.();
  return { width: W, height: H, ops: c.ops };
}

// ── cartesian charts (bar, line, area, scatter) ────────────────────────────

function drawCartesian(
  c: Ctx,
  spec: ChartSpec,
  outer: Box,
  W: number,
  H: number,
): void {
  const axisById = new Map(spec.axes.map((a) => [a.id, a]));
  const groups = spec.groups.filter(
    (g) => g.type !== 'pie' && g.type !== 'doughnut' && g.type !== 'radar',
  );
  if (groups.length === 0) return;
  const horizontal = groups.some((g) => g.type === 'bar' && g.barDir === 'bar');
  const blanksZero = spec.dispBlanksAs === 'zero';
  const scatter = groups.every((g) => g.type === 'scatter');

  // Value axes (primary + an optional secondary) and their groups.
  const valIds = [...new Set(groups.map((g) => g.axisIds[1]))];
  const valAxes = valIds.map((id) => ({
    axis: axisById.get(id),
    groups: groups.filter((g) => g.axisIds[1] === id),
  }));
  const primary = valAxes[0];
  const catAxis = axisById.get(groups[0].axisIds[0]);
  const scales = valAxes.map((v) => scaleFor(v.axis, v.groups, blanksZero));
  const xScale = scatter
    ? (() => {
        const [lo, hi] = xRange(groups);
        return autoScale(lo, hi, {
          ...(catAxis?.min !== undefined && { min: catAxis.min }),
          ...(catAxis?.max !== undefined && { max: catAxis.max }),
          ...(catAxis?.majorUnit !== undefined && { major: catAxis.majorUnit }),
        });
      })()
    : null;
  const nCats = Math.max(1, ...groups.map(nCategories));
  const categories =
    groups.flatMap((g) => g.series).find((s) => s.categories.length)
      ?.categories ?? Array.from({ length: nCats }, (_, i) => String(i + 1));

  const valFmt = (axis: ChartAxis | undefined, g: ChartGroup[]) =>
    g.some((x) => x.grouping === 'percentStacked') &&
    (!axis || axis.formatCode === 'General')
      ? '0%'
      : (axis?.formatCode ?? 'General');
  const valLabels = valAxes.map((v, i) =>
    v.axis && !v.axis.deleted && v.axis.tickLabelPosition !== 'none'
      ? ticks(scales[i]).map((tv) => formatNumber(tv, valFmt(v.axis, v.groups)))
      : [],
  );
  const catShown =
    catAxis && !catAxis.deleted && catAxis.tickLabelPosition !== 'none';
  const xLabels =
    scatter && xScale && catShown
      ? ticks(xScale).map((tv) =>
          formatNumber(tv, catAxis?.formatCode ?? 'General'),
        )
      : [];

  // Reserve room for tick labels and axis titles around the inner plot.
  let inner: Box;
  const titleRoom = (a: ChartAxis | undefined) =>
    a?.title && !a.deleted ? c.lineHeight(a.title.font) + AXIS_TITLE_GAP : 0;
  if (spec.plotArea.layout && spec.plotArea.layout.target !== 'outer') {
    inner = placed(spec.plotArea.layout, outer, W, H);
  } else {
    const valFont = primary.axis?.font;
    const lhVal = valFont ? c.lineHeight(valFont) : 0;
    const maxW = (labels: string[], f?: ChartFont) =>
      f ? Math.max(0, ...labels.map((l) => c.width(l, f))) : 0;
    let left = 0;
    let right = 0;
    let top = 0;
    let bottom = 0;
    const catFont = catAxis?.font;
    const lhCat = catFont ? c.lineHeight(catFont) : 0;
    if (!horizontal) {
      // Values on the left (and a secondary on the right); categories below.
      valAxes.forEach((v, i) => {
        if (!v.axis || v.axis.deleted) return;
        const room =
          maxW(valLabels[i], v.axis.font) + VAL_LABEL_GAP + titleRoom(v.axis);
        if (v.axis.position === 'r') right = Math.max(right, room);
        else left = Math.max(left, room);
      });
      top = lhVal / 2;
      if (catShown) bottom = lhCat + CAT_LABEL_GAP + titleRoom(catAxis);
      if (scatter && xLabels.length) {
        bottom = lhCat + CAT_LABEL_GAP + titleRoom(catAxis);
        right = Math.max(
          right,
          maxW([xLabels[xLabels.length - 1]], catFont) / 2,
        );
      }
    } else {
      if (catShown)
        left = maxW(categories, catFont) + VAL_LABEL_GAP + titleRoom(catAxis);
      valAxes.forEach((v, i) => {
        if (!v.axis || v.axis.deleted) return;
        const room = lhVal + CAT_LABEL_GAP + titleRoom(v.axis);
        if (v.axis.position === 't') top = Math.max(top, room);
        else bottom = Math.max(bottom, room);
        right = Math.max(
          right,
          maxW([valLabels[i][valLabels[i].length - 1] ?? ''], v.axis.font) / 2,
        );
      });
    }
    inner = {
      x: outer.x + left,
      y: outer.y + top,
      w: Math.max(10, outer.w - left - right),
      h: Math.max(10, outer.h - top - bottom),
    };
    if (spec.plotArea.layout) {
      // An "outer" manual layout places the box tick labels included.
      const o = placed(spec.plotArea.layout, outer, W, H);
      inner = {
        x: o.x + left,
        y: o.y + top,
        w: Math.max(10, o.w - left - right),
        h: Math.max(10, o.h - top - bottom),
      };
    }
  }

  c.rect(inner, spec.plotArea);

  // Value → position along the value direction.
  const valPos = (s: AxisScale, axis: ChartAxis | undefined) => (v: number) => {
    const f = (v - s.min) / (s.max - s.min || 1);
    const t = axis?.reversed ? 1 - f : f;
    return horizontal ? inner.x + t * inner.w : inner.y + inner.h - t * inner.h;
  };
  const posOf = scales.map((s, i) => valPos(s, valAxes[i].axis));
  const between = catAxis?.crossBetween !== 'midCat';
  const catSpan = horizontal ? inner.h : inner.w;
  const catW = between ? catSpan / nCats : catSpan / Math.max(1, nCats - 1);
  /** Category i's centre along the category direction. */
  const catPos = (i: number) => {
    const k = catAxis?.reversed ? nCats - 1 - i : i;
    // Horizontal bars list the first category at the BOTTOM.
    const kk = horizontal ? nCats - 1 - k : k;
    const offset = between ? (kk + 0.5) * catW : kk * catW;
    return (horizontal ? inner.y : inner.x) + offset;
  };
  const xPos = (x: number) =>
    xScale
      ? inner.x + ((x - xScale.min) / (xScale.max - xScale.min || 1)) * inner.w
      : inner.x;

  // Gridlines (minor under major), value axis first.
  valAxes.forEach((v, i) => {
    const a = v.axis;
    if (!a) return;
    const s = scales[i];
    const at = posOf[i];
    if (a.minorGridlines) {
      const minor = a.minorUnit ?? s.major / 5;
      for (let tv = s.min; tv <= s.max + 1e-9; tv += minor)
        gridAt(c, horizontal, at(tv), inner, a.minorGridlines);
    }
    if (a.majorGridlines)
      for (const tv of ticks(s))
        gridAt(c, horizontal, at(tv), inner, a.majorGridlines);
  });
  if (catAxis?.majorGridlines) {
    if (xScale)
      for (const tv of ticks(xScale))
        gridAt(c, true, xPos(tv), inner, catAxis.majorGridlines);
    else
      for (let i = 0; i <= nCats; i++) {
        const p =
          (horizontal ? inner.y : inner.x) + i * (between ? catW : catW);
        gridAt(c, !horizontal, p, inner, catAxis.majorGridlines);
      }
  }

  // Series: areas first, then bars, then lines/scatter (Word's z-order).
  const order = ['area', 'bar', 'line', 'scatter'];
  const labelsToDraw: (() => void)[] = [];
  const sorted = [...groups].sort(
    (a, b) => order.indexOf(a.type) - order.indexOf(b.type),
  );
  for (const g of sorted) {
    const vi = valIds.indexOf(g.axisIds[1]);
    const at = posOf[vi];
    const s = scales[vi];
    const base = at(Math.min(Math.max(0, s.min), s.max));
    const values = plotted(g, blanksZero);
    if (g.type === 'bar')
      drawBars(c, g, values, at, base, catPos, catW, horizontal, labelsToDraw);
    else if (g.type === 'area')
      drawAreas(c, g, values, at, base, catPos, labelsToDraw);
    else if (g.type === 'line')
      drawLines(c, g, values, at, catPos, spec.dispBlanksAs, labelsToDraw);
    else if (g.type === 'scatter')
      drawScatter(c, g, at, xPos, spec.dispBlanksAs, labelsToDraw);
  }

  // Axes: lines, ticks, labels, titles.
  valAxes.forEach((v, i) => {
    const a = v.axis;
    if (!a || a.deleted) return;
    const s = scales[i];
    const at = posOf[i];
    const onRight = a.position === 'r' || a.position === 't';
    // The value axis stands where the category axis crosses it (its edge).
    const edge = horizontal
      ? onRight
        ? inner.y
        : inner.y + inner.h
      : onRight
        ? inner.x + inner.w
        : inner.x;
    if (horizontal) c.line(inner.x, edge, inner.x + inner.w, edge, a.line);
    else c.line(edge, inner.y, edge, inner.y + inner.h, a.line);
    const dir = onRight ? 1 : -1;
    const tickLen = (m: ChartAxis['majorTickMark']) =>
      m === 'out'
        ? [0, TICK]
        : m === 'in'
          ? [-TICK, 0]
          : m === 'cross'
            ? [-TICK, TICK]
            : null;
    const tl = tickLen(a.majorTickMark);
    ticks(s).forEach((tv, k) => {
      const p = at(tv);
      if (tl) {
        if (horizontal)
          c.line(
            p,
            edge + dir * -tl[0],
            p,
            edge + dir * -tl[1],
            a.line ?? null,
          );
        else
          c.line(edge + dir * tl[0], p, edge + dir * tl[1], p, a.line ?? null);
      }
      const label = valLabels[i][k];
      if (!label) return;
      const lh = c.lineHeight(a.font);
      if (horizontal)
        c.text(
          label,
          a.font,
          p,
          onRight ? edge - CAT_LABEL_GAP - lh : edge + CAT_LABEL_GAP,
          'center',
        );
      else
        c.text(
          label,
          a.font,
          onRight ? edge + VAL_LABEL_GAP : edge - VAL_LABEL_GAP,
          p - lh / 2,
          onRight ? 'left' : 'right',
        );
    });
    if (a.title)
      drawAxisTitle(
        c,
        a,
        horizontal ? 'h' : onRight ? 'r' : 'l',
        inner,
        valLabels[i],
      );
  });

  if (catAxis && !catAxis.deleted) {
    // The category axis crosses the value axis at zero (autoZero), or at the
    // value/end the file names.
    const s = scales[0];
    const cross =
      catAxis.crosses === 'max'
        ? s.max
        : catAxis.crosses === 'min'
          ? s.min
          : typeof catAxis.crosses === 'number'
            ? catAxis.crosses
            : Math.min(Math.max(0, s.min), s.max);
    const p = posOf[0](cross);
    if (horizontal) c.line(p, inner.y, p, inner.y + inner.h, catAxis.line);
    else c.line(inner.x, p, inner.x + inner.w, p, catAxis.line);
    const tl =
      catAxis.majorTickMark === 'out'
        ? [0, TICK]
        : catAxis.majorTickMark === 'in'
          ? [-TICK, 0]
          : catAxis.majorTickMark === 'cross'
            ? [-TICK, TICK]
            : null;
    const labelEdge = horizontal ? inner.x : inner.y + inner.h;
    if (xScale) {
      for (const [k, tv] of ticks(xScale).entries()) {
        const x = xPos(tv);
        if (tl) c.line(x, p + tl[0], x, p + tl[1], catAxis.line);
        if (catShown)
          c.text(
            xLabels[k],
            catAxis.font,
            x,
            labelEdge + CAT_LABEL_GAP,
            'center',
          );
      }
    } else {
      if (tl)
        for (let i = 0; i <= (between ? nCats : nCats - 1); i++) {
          const q = (horizontal ? inner.y : inner.x) + i * catW;
          if (horizontal) c.line(p - tl[0], q, p - tl[1], q, catAxis.line);
          else c.line(q, p + tl[0], q, p + tl[1], catAxis.line);
        }
      if (catShown)
        drawCategoryLabels(
          c,
          catAxis,
          categories,
          catPos,
          catW,
          horizontal,
          labelEdge,
        );
    }
    if (catAxis.title)
      drawAxisTitle(c, catAxis, horizontal ? 'l' : 'h', inner, categories);
  }

  for (const draw of labelsToDraw) draw();
}

function gridAt(
  c: Ctx,
  horizontalAxis: boolean,
  p: number,
  inner: Box,
  l: ChartLine,
): void {
  if (horizontalAxis) c.line(p, inner.y, p, inner.y + inner.h, l);
  else c.line(inner.x, p, inner.x + inner.w, p, l);
}

function drawCategoryLabels(
  c: Ctx,
  axis: ChartAxis,
  categories: string[],
  catPos: (i: number) => number,
  catW: number,
  horizontal: boolean,
  edge: number,
): void {
  const f = axis.font;
  const lh = c.lineHeight(f);
  if (horizontal) {
    categories.forEach((cat, i) =>
      c.text(cat, f, edge - VAL_LABEL_GAP, catPos(i) - lh / 2, 'right'),
    );
    return;
  }
  const widest = Math.max(0, ...categories.map((l) => c.width(l, f)));
  const rotation =
    axis.labelRotation ??
    (widest > catW &&
    categories.some((l) => !l.includes(' ') || c.width(l, f) > catW * 2)
      ? -45
      : 0);
  categories.forEach((cat, i) => {
    const x = catPos(i);
    if (rotation) {
      // Slanted labels hang from the axis, ending at their category.
      const w = c.width(cat, f);
      const a = (rotation * Math.PI) / 180;
      c.text(
        cat,
        f,
        x - w * Math.cos(a) - lh * 0.3,
        edge + CAT_LABEL_GAP - w * Math.sin(a),
        'left',
        rotation,
      );
      return;
    }
    const lines = widest > catW ? c.wrap(cat, f, catW) : [cat];
    lines.forEach((l, k) =>
      c.text(l, f, x, edge + CAT_LABEL_GAP + k * lh, 'center'),
    );
  });
}

function drawAxisTitle(
  c: Ctx,
  a: ChartAxis,
  side: 'l' | 'r' | 'h',
  inner: Box,
  labels: string[],
): void {
  const t = a.title;
  if (!t) return;
  const lh = c.lineHeight(t.font);
  if (side === 'h') {
    const labelLh = c.lineHeight(a.font);
    c.text(
      t.text,
      t.font,
      inner.x + inner.w / 2,
      inner.y + inner.h + CAT_LABEL_GAP + labelLh + AXIS_TITLE_GAP,
      'center',
      t.rotation,
    );
    return;
  }
  const labelW = Math.max(0, ...labels.map((l) => c.width(l, a.font)));
  const w = c.width(t.text, t.font);
  const rot = t.rotation ?? -90;
  const x =
    side === 'l'
      ? inner.x - VAL_LABEL_GAP - labelW - AXIS_TITLE_GAP - lh
      : inner.x + inner.w + VAL_LABEL_GAP + labelW + AXIS_TITLE_GAP;
  if (rot === -90 || rot === 270) {
    // Rotated a quarter turn counter-clockwise: the baseline runs upward.
    const cy = inner.y + inner.h / 2;
    c.ops.push({
      kind: 'text',
      x: x + lh - t.font.sizePt * 0.25,
      y: cy + w / 2,
      text: t.text,
      size: t.font.sizePt,
      family: t.font.family,
      ...(t.font.bold && { bold: true }),
      ...(t.font.italic && { italic: true }),
      color: t.font.color,
      rotation: -90,
    });
  } else {
    c.text(t.text, t.font, x, inner.y + inner.h / 2 - lh / 2, 'left');
  }
}

function drawBars(
  c: Ctx,
  g: ChartGroup,
  values: ([number, number] | null)[][],
  at: (v: number) => number,
  base: number,
  catPos: (i: number) => number,
  catW: number,
  horizontal: boolean,
  labels: (() => void)[],
): void {
  const stacked = isStacked(g);
  const n = stacked ? 1 : g.series.length;
  const overlap = stacked ? 100 : g.overlap;
  const barW = catW / (n - ((n - 1) * overlap) / 100 + g.gapWidth / 100);
  const step = barW * (1 - overlap / 100);
  const groupW = barW + (n - 1) * step;
  g.series.forEach((s, si) => {
    const k = stacked ? 0 : si;
    values[si].forEach((p, i) => {
      if (!p) return;
      const centre = catPos(i);
      const start = centre - groupW / 2 + k * step;
      // Horizontal bars stack their series bottom-up within a category.
      const lane = horizontal ? centre + groupW / 2 - k * step - barW : start;
      const from = stacked ? at(p[0]) : base;
      const to = at(p[1]);
      const pt = s.points?.find((q) => q.idx === i);
      const fill = pt?.fill !== undefined ? pt.fill : s.fill;
      const line = pt?.line !== undefined ? pt.line : s.line;
      const b: Box = horizontal
        ? { x: Math.min(from, to), y: lane, w: Math.abs(to - from), h: barW }
        : { x: lane, y: Math.min(from, to), w: barW, h: Math.abs(to - from) };
      c.rect(b, { fill, line });
      const raw = s.values[i];
      const text = raw == null ? '' : labelText(s, i, raw, null);
      if (!text || !s.labels) return;
      const lb = s.labels;
      const pos = lb.position ?? (stacked ? 'ctr' : 'outEnd');
      labels.push(() => {
        const lh = c.lineHeight(lb.font);
        const w = c.width(text, lb.font);
        const neg = p[1] < p[0];
        if (!horizontal) {
          const cx = b.x + b.w / 2;
          const topEnd = neg ? b.y + b.h : b.y;
          let top: number;
          if (pos === 'ctr') top = b.y + b.h / 2 - lh / 2;
          else if (pos === 'inEnd')
            top = neg ? topEnd - lh - BAR_LABEL_GAP : topEnd + BAR_LABEL_GAP;
          else if (pos === 'inBase')
            top = neg ? b.y + BAR_LABEL_GAP : b.y + b.h - lh - BAR_LABEL_GAP;
          else top = neg ? topEnd + BAR_LABEL_GAP : topEnd - lh - BAR_LABEL_GAP;
          c.text(text, lb.font, cx, top, 'center');
        } else {
          const cy = b.y + b.h / 2 - lh / 2;
          const endX = neg ? b.x : b.x + b.w;
          let x: number;
          if (pos === 'ctr') x = b.x + b.w / 2 - w / 2;
          else if (pos === 'inEnd')
            x = neg ? endX + BAR_LABEL_GAP : endX - w - BAR_LABEL_GAP;
          else if (pos === 'inBase')
            x = neg ? b.x + b.w - w - BAR_LABEL_GAP : b.x + BAR_LABEL_GAP;
          else x = neg ? endX - w - BAR_LABEL_GAP : endX + BAR_LABEL_GAP;
          c.text(text, lb.font, x, cy);
        }
      });
    });
  });
}

function drawAreas(
  c: Ctx,
  g: ChartGroup,
  values: ([number, number] | null)[][],
  at: (v: number) => number,
  base: number,
  catPos: (i: number) => number,
  labels: (() => void)[],
): void {
  const stacked = isStacked(g);
  // Back to front: later series in a standard area sit behind earlier ones.
  const order = stacked
    ? g.series.map((_, i) => i)
    : g.series.map((_, i) => i).reverse();
  for (const si of order) {
    const s = g.series[si];
    const pts = values[si].map((p, i) => ({
      x: catPos(i),
      y: at(p ? p[1] : 0),
      lo: stacked && p ? at(p[0]) : base,
    }));
    if (pts.length < 2) continue;
    const poly = [
      ...pts.map((p) => ({ x: p.x, y: p.y })),
      ...[...pts].reverse().map((p) => ({ x: p.x, y: p.lo })),
    ];
    c.poly(poly, s.fill, s.line);
    pts.forEach((p, i) => {
      const raw = s.values[i];
      const text = raw == null ? '' : labelText(s, i, raw, null);
      if (!text || !s.labels) return;
      const lb = s.labels;
      labels.push(() =>
        c.text(
          text,
          lb.font,
          p.x,
          (p.y + p.lo) / 2 - c.lineHeight(lb.font) / 2,
          'center',
        ),
      );
    });
  }
}

function pointLabel(
  c: Ctx,
  s: ChartSeries,
  text: string,
  x: number,
  y: number,
  labels: (() => void)[],
): void {
  const lb = s.labels;
  if (!lb || !text) return;
  const r = s.marker && s.marker.symbol !== 'none' ? s.marker.size / 2 : 0;
  const pos = lb.position ?? 'r';
  labels.push(() => {
    const lh = c.lineHeight(lb.font);
    const w = c.width(text, lb.font);
    switch (pos) {
      case 'l':
        return c.text(text, lb.font, x - r - LABEL_GAP - w, y - lh / 2);
      case 't':
        return c.text(text, lb.font, x, y - r - BAR_LABEL_GAP - lh, 'center');
      case 'b':
        return c.text(text, lb.font, x, y + r + BAR_LABEL_GAP, 'center');
      case 'ctr':
        return c.text(text, lb.font, x, y - lh / 2, 'center');
      default:
        return c.text(text, lb.font, x + r + LABEL_GAP, y - lh / 2);
    }
  });
}

function drawLines(
  c: Ctx,
  g: ChartGroup,
  values: ([number, number] | null)[][],
  at: (v: number) => number,
  catPos: (i: number) => number,
  blanks: ChartSpec['dispBlanksAs'],
  labels: (() => void)[],
): void {
  g.series.forEach((s, si) => {
    const pts = values[si].map((p, i) =>
      p ? { x: catPos(i), y: at(p[1]), i } : null,
    );
    // A blank breaks the line ("gap") or is stepped over ("span").
    const runs: { x: number; y: number }[][] = [[]];
    for (const p of pts) {
      if (!p) {
        if (blanks !== 'span' && runs[runs.length - 1].length) runs.push([]);
        continue;
      }
      runs[runs.length - 1].push({ x: p.x, y: p.y });
    }
    for (const run of runs) c.polyline(s.smooth ? smoothPts(run) : run, s.line);
    for (const p of pts) {
      if (!p) continue;
      if (s.marker) drawMarker(c, markerFor(s, p.i), p.x, p.y);
      const raw = s.values[p.i];
      if (raw != null)
        pointLabel(c, s, labelText(s, p.i, raw, null), p.x, p.y, labels);
    }
  });
}

function markerFor(s: ChartSeries, i: number): ChartMarker {
  const m = s.marker as ChartMarker;
  const pt = s.points?.find((q) => q.idx === i);
  if (!pt) return m;
  return {
    ...m,
    ...(pt.fill !== undefined && { fill: pt.fill }),
    ...(pt.line !== undefined && { line: pt.line }),
  };
}

function drawScatter(
  c: Ctx,
  g: ChartGroup,
  at: (v: number) => number,
  xPos: (x: number) => number,
  blanks: ChartSpec['dispBlanksAs'],
  labels: (() => void)[],
): void {
  for (const s of g.series) {
    const xs = s.xValues ?? s.values.map((_, i) => i + 1);
    const pts = s.values.map((v, i) => {
      const x = xs[i];
      return v == null || x == null ? null : { x: xPos(x), y: at(v), i };
    });
    const runs: { x: number; y: number }[][] = [[]];
    for (const p of pts) {
      if (!p) {
        if (blanks !== 'span' && runs[runs.length - 1].length) runs.push([]);
        continue;
      }
      runs[runs.length - 1].push({ x: p.x, y: p.y });
    }
    for (const run of runs) c.polyline(s.smooth ? smoothPts(run) : run, s.line);
    for (const p of pts) {
      if (!p) continue;
      if (s.marker) drawMarker(c, markerFor(s, p.i), p.x, p.y);
      const raw = s.values[p.i];
      if (raw != null)
        pointLabel(c, s, labelText(s, p.i, raw, null), p.x, p.y, labels);
    }
  }
}

// ── pie & doughnut ──────────────────────────────────────────────────────────

/** Proportion of the plot box's smaller side a pie's diameter takes. */
const PIE_FILL = 0.9;

function drawPies(
  c: Ctx,
  spec: ChartSpec,
  groups: ChartGroup[],
  area: Box,
): void {
  const box = spec.plotArea.layout
    ? placed(spec.plotArea.layout, area, area.w, area.h)
    : area;
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  const R = (Math.min(box.w, box.h) / 2) * PIE_FILL;
  const g = groups[0];
  const s = g.series[0];
  if (!s) return;
  const vals = s.values.map((v) => Math.max(0, v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0) || 1;
  const hole = g.type === 'doughnut' ? (g.holeSize ?? 50) / 100 : 0;
  // Degrees clockwise from 12 o'clock.
  let a0 = ((g.firstSliceAng ?? 0) - 90) * (Math.PI / 180);
  const labels: (() => void)[] = [];
  vals.forEach((v, i) => {
    const sweep = (v / total) * Math.PI * 2;
    const a1 = a0 + sweep;
    const pt = s.points?.find((q) => q.idx === i);
    const fill = pt?.fill !== undefined ? pt.fill : s.fill;
    const line = pt?.line !== undefined ? pt.line : s.line;
    const ex = ((pt?.explosion ?? s.explosion ?? 0) / 100) * R;
    const mid = (a0 + a1) / 2;
    const ox = cx + ex * Math.cos(mid);
    const oy = cy + ex * Math.sin(mid);
    const steps = Math.max(2, Math.ceil((sweep / (Math.PI * 2)) * 72));
    const outer = Array.from({ length: steps + 1 }, (_, k) => {
      const a = a0 + (sweep * k) / steps;
      return { x: ox + R * Math.cos(a), y: oy + R * Math.sin(a) };
    });
    const innerPts = hole
      ? outer.map((_, k) => {
          const a = a1 - (sweep * k) / steps;
          return {
            x: ox + R * hole * Math.cos(a),
            y: oy + R * hole * Math.sin(a),
          };
        })
      : [{ x: ox, y: oy }];
    if (v > 0) c.poly([...outer, ...innerPts], fill, line);
    const text = labelText(s, i, s.values[i] ?? 0, v / total);
    if (text && s.labels) {
      const lb = s.labels;
      labels.push(() => {
        const lh = c.lineHeight(lb.font);
        const w = c.width(text, lb.font);
        const pos = lb.position ?? 'bestFit';
        // Inside the slice when it fits (bestFit), else just outside.
        const rIn = hole ? (R * (1 + hole)) / 2 : R * 0.65;
        const fits = sweep * rIn > w * 0.9 || pos === 'ctr' || pos === 'inEnd';
        const rr =
          pos === 'outEnd' || !fits ? R + LABEL_GAP + Math.max(w, lh) / 2 : rIn;
        c.text(
          text,
          lb.font,
          ox + rr * Math.cos(mid),
          oy + rr * Math.sin(mid) - lh / 2,
          'center',
        );
      });
    }
    a0 = a1;
  });
  for (const draw of labels) draw();
}

// ── radar ───────────────────────────────────────────────────────────────────

function drawRadar(
  c: Ctx,
  spec: ChartSpec,
  groups: ChartGroup[],
  area: Box,
): void {
  const axisById = new Map(spec.axes.map((a) => [a.id, a]));
  const g0 = groups[0];
  const valAxis = axisById.get(g0.axisIds[1]);
  const catAxis = axisById.get(g0.axisIds[0]);
  const scale = scaleFor(valAxis, groups, spec.dispBlanksAs === 'zero');
  const n = Math.max(1, ...groups.map(nCategories));
  const cats = g0.series[0]?.categories ?? [];
  const f = catAxis?.font;
  const labelRoom = f
    ? c.lineHeight(f) + Math.max(0, ...cats.map((l) => c.width(l, f))) / 2
    : 0;
  const cx = area.x + area.w / 2;
  const cy = area.y + area.h / 2;
  const R = Math.max(10, Math.min(area.w, area.h) / 2 - labelRoom);
  const at = (i: number, v: number) => {
    const a = -Math.PI / 2 + (i / n) * Math.PI * 2;
    const r = ((v - scale.min) / (scale.max - scale.min || 1)) * R;
    return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
  };
  if (valAxis?.majorGridlines)
    for (const tv of ticks(scale)) {
      const ring = Array.from({ length: n }, (_, i) => at(i, tv));
      c.poly(ring, null, valAxis.majorGridlines);
    }
  for (let i = 0; i < n; i++)
    c.line(
      cx,
      cy,
      at(i, scale.max).x,
      at(i, scale.max).y,
      catAxis?.line ?? valAxis?.majorGridlines ?? null,
    );
  for (const g of groups)
    for (const s of g.series) {
      const pts = s.values.map((v, i) => at(i, v ?? scale.min));
      if (s.fill) c.poly(pts, s.fill, s.line);
      else c.polyline([...pts, pts[0]], s.line);
      if (s.marker)
        pts.forEach((p) => drawMarker(c, s.marker as ChartMarker, p.x, p.y));
    }
  if (f && catAxis && !catAxis.deleted)
    cats.forEach((cat, i) => {
      const p = at(i, scale.max);
      const a = -Math.PI / 2 + (i / n) * Math.PI * 2;
      const lh = c.lineHeight(f);
      c.text(
        cat,
        f,
        p.x + Math.cos(a) * (lh / 2 + 4),
        p.y + Math.sin(a) * (lh / 2 + 4) - lh / 2,
        'center',
      );
    });
  if (valAxis && !valAxis.deleted && valAxis.tickLabelPosition !== 'none')
    for (const tv of ticks(scale)) {
      const p = at(0, tv);
      const lh = c.lineHeight(valAxis.font);
      c.text(
        formatNumber(tv, valAxis.formatCode),
        valAxis.font,
        p.x - 4,
        p.y - lh / 2,
        'right',
      );
    }
}
