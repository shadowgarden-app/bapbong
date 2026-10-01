/**
 * A chart part → ChartSpec: the data, and the look Word would give it.
 * Colours are checked against what Word 365 drew for D-2609-LSTQ (PDF,
 * PyMuPDF): a 2007-style line series in accent1 is #4A7EBB at 2.25pt with
 * #4F81BD diamond markers; the bars are the plain accents.
 */
import { parseXml } from './ooxml';
import { chartTheme, parseChartSpec, seriesColor } from './chart-spec';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const C14 = 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart';

/** The Office 2007 theme's colours and the part of its format scheme a
 *  chart style draws on. */
const THEME = parseXml(
  `<a:theme xmlns:a="${A}"><a:themeElements>` +
    `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
    `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2>` +
    `<a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>` +
    `<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3>` +
    `<a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>` +
    `<a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>` +
    `<a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>` +
    `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
    `<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill></a:ln>` +
    `<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` +
    `<a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
    `</a:fmtScheme></a:themeElements></a:theme>`,
);
const theme = chartTheme(THEME);

const strCache = (vals: string[]) =>
  `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache>`;
const numCache = (
  vals: number[],
  fmt = 'General',
  pointFmt: Record<number, string> = {},
) =>
  `<c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${vals.length}"/>${vals
    .map(
      (v, i) =>
        `<c:pt idx="${i}"${pointFmt[i] ? ` formatCode="${pointFmt[i]}"` : ''}><c:v>${v}</c:v></c:pt>`,
    )
    .join('')}</c:numCache>`;
const ser = (
  i: number,
  name: string,
  vals: number[],
  extra = '',
  fmt = 'General',
) =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>S!$B$1</c:f>${strCache([name])}</c:strRef></c:tx>${extra}` +
  `<c:cat><c:strRef><c:f>S!$A$2</c:f>${strCache(['2011', '2012', '2013'].slice(0, vals.length))}</c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>S!$B$2</c:f>${numCache(vals, fmt, i === 0 ? { 0: '0%' } : {})}</c:numRef></c:val></c:ser>`;
const axes =
  `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>` +
  `<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:catAx>` +
  `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>` +
  `<c:numFmt formatCode="0%" sourceLinked="1"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
const space = (
  plot: string,
  style = `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="c14" xmlns:c14="${C14}"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>`,
  chartExtra = '<c:title><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>',
) =>
  `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:roundedCorners val="0"/>${style}` +
  `<c:chart>${chartExtra}<c:plotArea><c:layout/>${plot}</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend></c:chart></c:chartSpace>`;

const LINE = space(
  `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser(0, 'Growth', [0, 0.862, 0.449], '', '0.0%')}` +
    `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>` +
    `<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes}`,
);

describe('parseChartSpec', () => {
  it('reads a 2007-style line chart the way Word draws it', () => {
    const spec = parseChartSpec(LINE, theme)!;
    const s = spec.groups[0].series[0];
    expect(spec.groups[0].type).toBe('line');
    expect(s.values).toEqual([0, 0.862, 0.449]);
    expect(s.categories).toEqual(['2011', '2012', '2013']);
    expect(s.formatCode).toBe('0.0%');
    expect(s.pointFormats).toEqual({ 0: '0%' });
    // Theme line style 1 (shade 95%, satMod 105%) × 300% — measured #4A7EBB.
    expect(s.line).toEqual({ color: '#4A7EBB', widthPt: 2.25 });
    expect(s.marker).toEqual({
      symbol: 'diamond',
      size: 7,
      fill: '#4F81BD',
      line: { color: '#4A7EBB', widthPt: 0.75 },
    });
    expect(s.labels?.showVal).toBe(true);
  });

  it('titles a one-series chart with the series name', () => {
    expect(parseChartSpec(LINE, theme)!.title).toMatchObject({
      text: 'Growth',
      font: { family: 'Calibri', sizePt: 18, bold: true },
    });
  });

  it('frames the chart and greys the axes like Word 2007 styles', () => {
    const spec = parseChartSpec(LINE, theme)!;
    expect(spec.chartArea.fill).toBe('#FFFFFF');
    expect(spec.chartArea.line?.widthPt).toBe(0.75);
    const val = spec.axes.find((a) => a.kind === 'val')!;
    expect(val.majorGridlines?.color).toBe(spec.chartArea.line?.color);
    expect(val.formatCode).toBe('0%');
    expect(spec.roundedCorners).toBe(false);
  });

  it('colours clustered bars with the accents in turn', () => {
    const spec = parseChartSpec(
      space(
        `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser(0, 'A', [1, 2, 3])}${ser(1, 'B', [2, 3, 4])}` +
          `<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}`,
      ),
      theme,
    )!;
    const [a, b] = spec.groups[0].series;
    expect([a.fill, b.fill]).toEqual(['#4F81BD', '#C0504D']);
    expect(a.line).toBeNull();
    expect(spec.groups[0]).toMatchObject({
      barDir: 'col',
      gapWidth: 150,
      overlap: 0,
    });
  });

  it('lets explicit formatting win over the style (Word 2013+ charts)', () => {
    const sp = `<c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>`;
    const spec = parseChartSpec(
      space(
        `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>${ser(0, 'A', [1, 2, 3], sp)}` +
          `<c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}`,
      ),
      theme,
    )!;
    const s = spec.groups[0].series[0];
    expect(s.fill).toBe('#123456');
    expect(s.line).toEqual({ color: '#ABCDEF', widthPt: 1.5, dash: [4, 3] });
    expect(spec.groups[0]).toMatchObject({
      barDir: 'bar',
      grouping: 'stacked',
      overlap: 100,
    });
  });

  it('fades one accent across the series for styles 3…8', () => {
    const c = [0, 1, 2].map((i) => seriesColor(theme, 3, i, 3));
    // Shaded, pure, tinted — the middle series is the accent itself.
    expect(c[1]).toBe('#4F81BD');
    expect(c[0]).not.toBe(c[1]);
    expect(c[2]).not.toBe(c[1]);
  });

  it('gives every slice of a pie its own colour', () => {
    const spec = parseChartSpec(
      space(
        `<c:pieChart><c:varyColors val="1"/>${ser(0, 'Sales', [3, 2, 1])}<c:firstSliceAng val="0"/></c:pieChart>`,
      ),
      theme,
    )!;
    const pts = spec.groups[0].series[0].points!;
    expect(pts.map((p) => p.fill)).toEqual(['#4F81BD', '#C0504D', '#9BBB59']);
  });

  it('draws a chart naming no style with black axes (measured)', () => {
    const spec = parseChartSpec(
      LINE.replace(/<mc:AlternateContent.*?<\/mc:AlternateContent>/, ''),
      theme,
    )!;
    const val = spec.axes.find((a) => a.kind === 'val')!;
    expect(val.line?.color).toBe('#000000');
    expect(val.majorGridlines?.color).toBe('#000000');
    // Everything else is style 2's.
    expect(spec.groups[0].series[0].line?.color).toBe('#4A7EBB');
  });

  it('leaves what it does not draw to the placeholder', () => {
    expect(
      parseChartSpec(
        space('<c:bar3DChart><c:barDir val="col"/></c:bar3DChart>'),
        theme,
      ),
    ).toBeNull();
  });
});
