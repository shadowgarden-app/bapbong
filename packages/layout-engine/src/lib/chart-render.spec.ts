import type {
  ChartAxis,
  ChartFont,
  ChartGroup,
  ChartSeries,
  ChartSpec,
  MeasureText,
  VectorOp,
  VectorPolygonOp,
  VectorTextOp,
} from '@shadow-garden/bapbong-contracts';
import { renderChart } from './chart-render.js';

/** Half an em per character, in px — enough to lay out against. */
const measure: MeasureText = (text, font) =>
  text.length * font.sizePt * (96 / 72) * 0.5;

const font: ChartFont = {
  family: 'Calibri',
  sizePt: 10,
  bold: false,
  italic: false,
  color: '#000000',
};
const grid = { color: '#868686', widthPt: 0.75 };

const axis = (over: Partial<ChartAxis>): ChartAxis => ({
  id: 1,
  kind: 'cat',
  position: 'b',
  deleted: false,
  crossAxisId: 2,
  crosses: 'autoZero',
  crossBetween: 'between',
  reversed: false,
  formatCode: 'General',
  sourceLinked: true,
  majorGridlines: null,
  minorGridlines: null,
  line: grid,
  majorTickMark: 'none',
  minorTickMark: 'none',
  tickLabelPosition: 'nextTo',
  font,
  title: null,
  ...over,
});

const series = (over: Partial<ChartSeries>): ChartSeries => ({
  name: 'S',
  values: [],
  categories: ['2011', '2012', '2013'],
  formatCode: 'General',
  fill: null,
  line: null,
  marker: null,
  smooth: false,
  labels: null,
  ...over,
});

const chart = (
  group: Partial<ChartGroup>,
  valFormat = 'General',
  legend = true,
): ChartSpec => ({
  groups: [
    {
      type: 'line',
      grouping: 'standard',
      varyColors: false,
      gapWidth: 150,
      overlap: 0,
      axisIds: [1, 2],
      series: [],
      ...group,
    },
  ],
  axes: [
    axis({}),
    axis({
      id: 2,
      kind: 'val',
      position: 'l',
      crossAxisId: 1,
      majorGridlines: grid,
      formatCode: valFormat,
    }),
  ],
  title: null,
  legend: legend
    ? { position: 'r', overlay: false, font, box: { fill: null, line: null } }
    : null,
  chartArea: { fill: '#FFFFFF', line: grid },
  plotArea: { fill: '#FFFFFF', line: null },
  roundedCorners: false,
  dispBlanksAs: 'gap',
});

const texts = (ops: VectorOp[]) =>
  (ops.filter((o) => o.kind === 'text') as VectorTextOp[]).map((t) => t.text);
const fills = (ops: VectorOp[], color: string) =>
  ops.filter(
    (o): o is VectorPolygonOp => o.kind === 'polygon' && o.fill === color,
  );

describe('renderChart', () => {
  // D-2609-LSTQ's first chart: one line series, "0.0%" values, a 0 kept "0%".
  const line = chart(
    {
      series: [
        series({
          name: 'Growth',
          values: [0, 0.862, 0.449],
          formatCode: '0.0%',
          pointFormats: { 0: '0%' },
          line: { color: '#4A7EBB', widthPt: 2.25 },
          marker: {
            symbol: 'diamond',
            size: 7,
            fill: '#4F81BD',
            line: { color: '#4A7EBB', widthPt: 0.75 },
          },
          labels: {
            showVal: true,
            showPercent: false,
            showCatName: false,
            showSerName: false,
            separator: ', ',
            font,
          },
        }),
      ],
    },
    '0%',
  );

  it('draws in points of the frame', () => {
    const v = renderChart(line, 528, 345, measure);
    expect(v.width).toBe(396);
    expect(v.height).toBeCloseTo(258.75);
  });

  it('labels the auto-scaled value axis 0%…100% with a gridline per tick', () => {
    const v = renderChart(line, 528, 345, measure);
    const t = texts(v.ops);
    for (let p = 0; p <= 100; p += 10) expect(t).toContain(`${p}%`);
    const gridlines = v.ops.filter(
      (o) => o.kind === 'line' && o.y1 === o.y2 && o.color === grid.color,
    );
    // 11 gridlines plus the category axis on the zero line.
    expect(gridlines.length).toBe(12);
  });

  it('puts the points at the category centres, with their data labels', () => {
    const v = renderChart(line, 528, 345, measure);
    const pl = v.ops.find((o) => o.kind === 'polyline');
    expect(pl && pl.kind === 'polyline' && pl.points).toHaveLength(3);
    if (pl?.kind !== 'polyline') return;
    const [a, b, cc] = pl.points;
    expect(b.x - a.x).toBeCloseTo(cc.x - b.x);
    // Three point markers and the legend key's.
    expect(fills(v.ops, '#4F81BD')).toHaveLength(4);
    expect(texts(v.ops)).toEqual(
      expect.arrayContaining(['0%', '86.2%', '44.9%', 'Growth']),
    );
  });

  it('sizes clustered bars by the gap width', () => {
    const bars = chart(
      {
        type: 'bar',
        barDir: 'col',
        grouping: 'clustered',
        series: [
          series({ name: 'A', values: [888, 1306, 1521], fill: '#4F81BD' }),
          series({ name: 'B', values: [239, 328, 1028], fill: '#C0504D' }),
        ],
      },
      'General',
      false,
    );
    const v = renderChart(bars, 545, 288, measure);
    const a = fills(v.ops, '#4F81BD');
    const b = fills(v.ops, '#C0504D');
    expect(a).toHaveLength(3);
    const w = (p: VectorPolygonOp) => p.points[1].x - p.points[0].x;
    const catW = a[1].points[0].x - a[0].points[0].x;
    // Bar = category / (series + gap/100); the second series sits right beside.
    expect(w(a[0])).toBeCloseTo(catW / 3.5);
    expect(b[0].points[0].x).toBeCloseTo(a[0].points[1].x);
    expect(texts(v.ops)).toEqual(expect.arrayContaining(['0', '200', '1600']));
  });

  it('stacks to 100% on a percent axis', () => {
    const pct = chart(
      {
        type: 'bar',
        barDir: 'col',
        grouping: 'percentStacked',
        overlap: 100,
        series: [
          series({ values: [1, 1, 3], fill: '#111111' }),
          series({ values: [1, 3, 1], fill: '#222222' }),
        ],
      },
      'General',
      false,
    );
    const v = renderChart(pct, 400, 300, measure);
    expect(texts(v.ops)).toEqual(expect.arrayContaining(['0%', '50%', '100%']));
    const [lo] = fills(v.ops, '#111111');
    const [hi] = fills(v.ops, '#222222');
    // First category: half and half — the upper bar starts where the lower ends.
    expect(hi.points[2].y).toBeCloseTo(lo.points[0].y);
  });

  it('cuts a pie into one slice per point', () => {
    const pie = chart({
      type: 'pie',
      varyColors: true,
      series: [
        series({
          values: [3, 2, 1],
          points: [
            { idx: 0, fill: '#4F81BD' },
            { idx: 1, fill: '#C0504D' },
            { idx: 2, fill: '#9BBB59' },
          ],
        }),
      ],
    });
    const v = renderChart(pie, 400, 300, measure);
    for (const c of ['#4F81BD', '#C0504D', '#9BBB59'])
      expect(fills(v.ops, c).length).toBeGreaterThanOrEqual(1);
    // The legend lists the categories, not the series.
    expect(texts(v.ops)).toEqual(
      expect.arrayContaining(['2011', '2012', '2013']),
    );
  });
});
