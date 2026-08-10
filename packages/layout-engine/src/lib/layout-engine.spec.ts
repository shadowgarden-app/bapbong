import { Schema } from 'prosemirror-model';
import type {
  FlowBlock,
  FlowParagraph,
  FlowTableCell,
  FontSpec,
  LayoutConfig,
  MeasureText,
  ResolvedTable,
} from '@shadow-garden/bapbong-contracts';
import {
  createLayoutCache,
  layout,
  layoutBlocks,
  toFlowBlocks,
} from './layout-engine.js';

// 10px per character, font-agnostic — keeps wrapping math predictable.
const measure: MeasureText = (text) => text.length * 10;

const font = (over: Partial<FontSpec> = {}): FontSpec => ({
  family: 'Arial',
  sizePt: 10, // → 10 * 96/72 * 1.2 = 16px line height
  bold: false,
  italic: false,
  ...over,
});

// Typed as FlowParagraph, not FlowBlock: tests spread paragraph-only fields
// (keepNext, widowControl…) onto it, which the block union wouldn't allow.
const para = (text: string, marker?: string): FlowParagraph => ({
  type: 'paragraph',
  runs: [{ text, font: font() }],
  marker,
});

const config = (over: Partial<LayoutConfig['page']> = {}): LayoutConfig => ({
  measureText: measure,
  defaultFont: { sizePt: 10 },
  page: {
    width: 240,
    height: 1000,
    margin: { top: 20, right: 20, bottom: 20, left: 20 },
    ...over,
  },
});

describe('layoutBlocks', () => {
  it('wraps a paragraph at the content width', () => {
    // content width = 240 - 20 - 20 = 200px → 5 words of "aaaa"(40) + spaces.
    const { pages } = layoutBlocks(
      [para('aaaa bbbb cccc dddd eeee')],
      config(),
    );
    expect(pages).toHaveLength(1);
    expect(pages[0].lines).toHaveLength(2);
    expect(pages[0].lines[0].segments[0].text).toBe('aaaa');
    expect(pages[0].lines[0].y).toBe(20);
    expect(pages[0].lines[1].segments[0].text).toBe('eeee');
    expect(pages[0].lines[1].y).toBe(36); // 20 + 16
  });

  it('survives a degenerate page config (NaN margins) without hanging', () => {
    // Regression: a real-world docx shipped w:pgMar values with unit suffixes
    // ("20pt") — Number() made them NaN, and NaN margins turned the page-fill
    // loop infinite (the whole app froze). Sanitized configs must terminate.
    const cfg = config({
      margin: { top: NaN, right: NaN, bottom: NaN, left: NaN },
    });
    const { pages } = layoutBlocks([para('one'), para('two')], cfg);
    expect(pages.length).toBeGreaterThan(0);
    expect(pages[0].lines.length).toBeGreaterThan(0);
    // Margins that leave no content box degrade rather than loop.
    const tight = config({
      height: 50,
      margin: { top: 40, right: 20, bottom: 40, left: 20 },
    });
    expect(layoutBlocks([para('x')], tight).pages.length).toBeGreaterThan(0);
  });

  it('paginates when content exceeds the page height', () => {
    // page height 80, margins 20 → content 20..60 = 40px → 2 lines/page.
    const cfg = config({ height: 80 });
    const { pages } = layoutBlocks(
      [para('one'), para('two'), para('three')],
      cfg,
    );
    expect(pages).toHaveLength(2);
    expect(pages[0].lines).toHaveLength(2);
    expect(pages[1].lines).toHaveLength(1);
    expect(pages[1].lines[0].y).toBe(20); // reset to top on the new page
  });

  it('places a list marker before the indented content', () => {
    const { pages } = layoutBlocks([para('hi', '1.')], config());
    const [markerSeg, textSeg] = pages[0].lines[0].segments;
    expect(markerSeg.text).toBe('1.');
    expect(markerSeg.x).toBe(20);
    expect(textSeg.text).toBe('hi');
    expect(textSeg.x).toBe(50); // 20 + measure("1. ") = 20 + 30
  });

  it('always emits at least one page', () => {
    expect(layoutBlocks([], config()).pages).toHaveLength(1);
  });

  // content box is x ∈ [20, 220] (width 200); "aaaa" measures 40px.
  const aligned = (align: 'center' | 'right'): FlowBlock => ({
    type: 'paragraph',
    runs: [{ text: 'aaaa', font: font() }],
    align,
  });

  it('centers a line within the content box', () => {
    const { pages } = layoutBlocks([aligned('center')], config());
    // offset = (200 - 40) / 2 = 80 → 20 + 80.
    expect(pages[0].lines[0].segments[0].x).toBe(100);
  });

  it('right-aligns a line within the content box', () => {
    const { pages } = layoutBlocks([aligned('right')], config());
    expect(pages[0].lines[0].segments[0].x).toBe(180); // 20 + (200 - 40)
  });

  it('justifies all lines but the last', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'aaaa bbbb cccc dddd eeee', font: font() }],
      align: 'justify',
    };
    const { pages } = layoutBlocks([block], config());
    const [line0, line1] = pages[0].lines;
    // line0 = "aaaa bbbb cccc dddd": content 190, slack 10 over 3 gaps → +3.333/gap.
    expect(line0.segments[0].x).toBe(20); // first word pinned left
    expect(line0.segments[1].text).toBe(' ');
    expect(line0.segments[2].x).toBeCloseTo(73.333, 2); // "bbbb": 20 + 40 + 10 + 3.333
    // last line is not justified — left-aligned at the margin.
    expect(line1.segments[0].text).toBe('eeee');
    expect(line1.segments[0].x).toBe(20);
  });

  it('hangs a list marker and aligns wrapped lines under the text', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'aaaa bbbb cccc dddd', font: font() }],
      marker: '1.',
    };
    const { pages } = layoutBlocks([block], config());
    const [line0, line1] = pages[0].lines;
    expect(line0.segments[0]).toMatchObject({ text: '1.', x: 20 }); // marker hangs at margin
    expect(line0.segments[1]).toMatchObject({ text: 'aaaa', x: 50 }); // 20 + measure("1. ")
    // wrapped line aligns under the text (x = 50), not back at the margin (20).
    expect(line1.segments[0]).toMatchObject({ text: 'dddd', x: 50 });
  });

  it('styles the list label per markerStyle (jc / suff / own font+color)', () => {
    // Hanging indent: text at left=60, marker anchored at 60-30=30.
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hi', font: font() }],
      marker: '9.',
      markerStyle: {
        jc: 'right',
        font: { bold: true, sizePt: 8 },
        color: '#C00000',
      },
      indent: { left: 40, hanging: 30 },
    };
    const { pages } = layoutBlocks([block], config());
    const [markerSeg, textSeg] = pages[0].lines[0].segments;
    // anchor = 20 + 40 - 30 = 30; jc=right → right edge AT the anchor.
    expect(markerSeg).toMatchObject({
      text: '9.',
      x: 10, // 30 - width(20)
      color: '#C00000',
    });
    expect(markerSeg.font).toMatchObject({ bold: true, sizePt: 8 });
    // suff defaults to tab: marker end (30) ≤ text position (60) → jump there.
    expect(textSeg).toMatchObject({ text: 'hi', x: 60 });
  });

  it('renders small caps as case-split segments (reduced uppercase)', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'Ab cD', font: font(), smallCaps: true, pos: 100 }],
    };
    const { pages } = layoutBlocks([block], config());
    const segs = pages[0].lines[0].segments;
    // "A" full size · "B" reduced+uppercased · " " neutral · "C" reduced ·
    // "D" full.
    expect(segs.map((s) => s.text)).toEqual(['A', 'B', ' ', 'C', 'D']);
    expect(segs.map((s) => s.font.sizePt)).toEqual([10, 8, 10, 8, 10]);
    // PM positions still map 1:1 to the ORIGINAL characters.
    expect(segs.map((s) => s.pos)).toEqual([100, 101, 102, 103, 104]);
  });

  it('carries dstrike through to the painted segment', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'gone', font: font(), dstrike: true }],
    };
    const { pages } = layoutBlocks([block], config());
    expect(pages[0].lines[0].segments[0].dstrike).toBe(true);
  });

  it("suff 'nothing' keeps the text tight against the label", () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hi', font: font() }],
      marker: '1.',
      markerStyle: { suff: 'nothing' },
    };
    const { pages } = layoutBlocks([block], config());
    const [markerSeg, textSeg] = pages[0].lines[0].segments;
    expect(markerSeg.x).toBe(20);
    expect(textSeg.x).toBe(40); // straight after "1." (20px), no gap
  });

  // Pagination keeps. Page height 100 → band y ∈ [20, 80], 16px lines: 3 fit.
  // 15-char words wrap one per line, so word count = line count.
  const words = (n: number) =>
    Array.from({ length: n }, () => 'a'.repeat(15)).join(' ');
  const keepCfg = () => config({ height: 100 });

  it('widow/orphan control never strands a lone line on either side', () => {
    // 1 filler line + a 4-line paragraph: 2 lines fit after the filler and 2
    // move — a clean 2/2 split needs no adjustment.
    const even = layoutBlocks([para(words(1)), para(words(4))], keepCfg());
    expect(even.pages[0].lines).toHaveLength(3); // filler + 2
    expect(even.pages[1].lines).toHaveLength(2);
    // 2 filler lines + a 3-line paragraph: only 1 line would fit — an orphan —
    // so the whole paragraph moves.
    const orphan = layoutBlocks([para(words(2)), para(words(3))], keepCfg());
    expect(orphan.pages[0].lines).toHaveLength(2); // filler only
    expect(orphan.pages[1].lines).toHaveLength(3);
    // 1 filler line + a 3-line paragraph: 2 fit but the LAST line would sit
    // alone up top (widow) → give it company → but that strands an orphan →
    // the whole paragraph moves.
    const widow = layoutBlocks([para(words(1)), para(words(3))], keepCfg());
    expect(widow.pages[0].lines).toHaveLength(1);
    expect(widow.pages[1].lines).toHaveLength(3);
  });

  it('w:widowControl off splits wherever the band ends', () => {
    const { pages } = layoutBlocks(
      [para(words(1)), { ...para(words(3)), widowControl: false }],
      keepCfg(),
    );
    expect(pages[0].lines).toHaveLength(3); // filler + 2 — lone tail allowed
    expect(pages[1].lines).toHaveLength(1);
  });

  it('keepLines moves the whole paragraph instead of splitting', () => {
    const { pages } = layoutBlocks(
      [
        para(words(1)),
        { ...para(words(3)), keepLines: true, widowControl: false },
      ],
      keepCfg(),
    );
    expect(pages[0].lines).toHaveLength(1);
    expect(pages[1].lines).toHaveLength(3);
  });

  it('keepNext keeps a heading with the opening of what follows', () => {
    // 2 filler lines + heading + 4-line body: the heading alone still fits,
    // but the body's orphan-legal opening (2 lines) would not — the heading
    // moves and opens the next page with the body.
    const { pages } = layoutBlocks(
      [para(words(2)), { ...para(words(1)), keepNext: true }, para(words(4))],
      keepCfg(),
    );
    expect(pages[0].lines).toHaveLength(2); // filler only
    expect(pages[1].lines).toHaveLength(3); // heading + body's first 2 lines
    expect(pages[2].lines).toHaveLength(2); // body's last 2 lines
  });

  it('wraps lines placed BEFORE a float anchor around it (letterhead)', () => {
    // The letterhead pattern: a margin-positioned logo is ANCHORED in the
    // title paragraph, but its box sits at the top of the page, over the
    // empty paragraphs placed before the anchor. Word wraps the whole page
    // around a float wherever its anchor sits — the page replays with the
    // exclusion seeded, so the earlier lines move aside too.
    const filler = (t: string): FlowBlock => ({
      type: 'paragraph',
      runs: [{ text: t, font: font() }],
    });
    const anchor: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'title', font: font() }],
      floats: [
        {
          src: 'logo.png',
          width: 60,
          height: 60,
          wrap: 'square',
          hAlign: 'left',
          hRel: 'margin',
          vOffset: 0,
          vRel: 'margin', // page-constant: sits at the band top (y=20)
        },
      ],
    };
    const { pages } = layoutBlocks(
      [filler('aa'), filler('bb'), anchor],
      config(),
    );
    // Float occupies x 20–80, y 20–80: every line on the page starts right
    // of it — including the two placed before the anchor.
    const xs = pages[0].lines.map((l) => Math.round(l.x));
    expect(xs).toEqual([80, 80, 80]);
    expect(pages[0].floats?.[0]).toMatchObject({ x: 20, y: 20 });
  });

  it('a float that would push its own anchor off the page settles, not loops', () => {
    // Band [20, 80]. Two fillers reach y=52; the anchor's line fits at
    // 52–68. Its full-width float band (40px at the top) would displace the
    // fillers DOWN past the band, pushing the anchor to page 2 — where the
    // float must follow, un-seeding the exclusion. The sticky-drop rule
    // accepts the original layout instead of oscillating.
    const filler = (t: string): FlowBlock => ({
      type: 'paragraph',
      runs: [{ text: t, font: font() }],
    });
    const anchor: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'anchor', font: font() }],
      floats: [
        {
          src: 'wide.png',
          width: 200,
          height: 40,
          wrap: 'topAndBottom',
          hAlign: 'left',
          hRel: 'margin',
          vOffset: 0,
          vRel: 'margin',
        },
      ],
    };
    const { pages } = layoutBlocks(
      [filler('a'), filler('b'), anchor],
      config({ height: 100 }),
    );
    // Terminates (the test completing is the point) with the pre-trigger
    // layout: fillers unmoved, anchor still on page 1.
    expect(pages).toHaveLength(1);
    expect(pages[0].lines.map((l) => Math.round(l.x))).toEqual([20, 20, 20]);
  });

  it('applies left + firstLine indent to the first line', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'aaaa bbbb cccc dddd', font: font() }],
      indent: { left: 10, firstLine: 15 },
    };
    const { pages } = layoutBlocks([block], config());
    const [line0, line1] = pages[0].lines;
    expect(line0.segments[0].x).toBe(45); // 20 + 10 (left) + 15 (firstLine)
    expect(line1.segments[0].x).toBe(30); // continuation at left indent only (20 + 10)
  });

  it('lays out an inline image and grows the line to fit it', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [
        { text: 'see ', font: font() },
        { src: 'data:img', width: 50, height: 30 },
        { text: ' end', font: font() },
      ],
    };
    const { pages } = layoutBlocks([block], config());
    const [line] = pages[0].lines;
    // see(30) ' '(10) img(50) ' '(10) end(30) all fit on one line from x=20.
    expect(line.images).toHaveLength(1);
    expect(line.images?.[0]).toMatchObject({
      src: 'data:img',
      x: 60,
      width: 50,
      height: 30,
    });
    expect(line.segments.at(-1)).toMatchObject({ text: 'end', x: 120 });
    // image (30) is taller than the text line box (16) → line height = 30.
    expect(line.height).toBe(30);
  });

  it('wraps when an inline image does not fit', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [
        { text: 'aaaaaaaaaaaaaaaa', font: font() }, // 16 chars → 160px
        { src: 'data:img', width: 80, height: 20 }, // 160 + 80 > 200 → wraps
      ],
    };
    const { pages } = layoutBlocks([block], config());
    expect(pages[0].lines).toHaveLength(2);
    expect(pages[0].lines[0].segments[0].text).toBe('aaaaaaaaaaaaaaaa');
    expect(pages[0].lines[1].images?.[0]).toMatchObject({ x: 20, width: 80 });
  });

  // Build a single-cell paragraph cell (1×1) with optional overrides.
  const cell = (
    text: string,
    over: Partial<FlowTableCell> = {},
  ): FlowTableCell => ({
    colspan: 1,
    rowspan: 1,
    colwidth: null,
    content: [{ type: 'paragraph', runs: [{ text, font: font() }] }],
    ...over,
  });
  const table = (rows: FlowTableCell[][]): FlowBlock => ({
    type: 'table',
    rows: rows.map((cells) => ({ cells })),
  });

  it('lays out a grid with column widths and row heights', () => {
    // content box [20,220] (200px). Columns 80 + 120 from row-0 colwidths.
    // Word's implicit indent shifts the grid left by the default cell margin
    // (7.2px) so the first cell's TEXT lands on the margin.
    const t = table([
      [cell('a', { colwidth: [80] }), cell('b', { colwidth: [120] })],
      [cell('cc'), cell('dd')],
    ]);
    const { pages } = layoutBlocks([t], config());
    expect(pages).toHaveLength(1);
    const [resolved] = pages[0].tables ?? [];
    expect(resolved.x).toBeCloseTo(12.8); // 20 − 7.2 implicit indent
    expect(resolved).toMatchObject({ y: 20 });
    expect(resolved.width).toBeCloseTo(200);
    expect(resolved.height).toBeCloseTo(32); // 2 rows × 16px (float)
    expect(resolved.cells).toHaveLength(4);
    // top-left cell sits at the table origin; its text lands on the margin.
    expect(resolved.cells[0].x).toBeCloseTo(12.8);
    expect(resolved.cells[0]).toMatchObject({ y: 20, width: 80 });
    expect(resolved.cells[0].height).toBeCloseTo(16);
    expect(resolved.cells[0].lines[0]).toMatchObject({ y: 20 });
    expect(resolved.cells[0].lines[0].segments[0].text).toBe('a');
    expect(resolved.cells[0].lines[0].segments[0].x).toBeCloseTo(20);
    // second column at 12.8 + 80 = 92.8; second row at 20 + 16 = 36.
    expect(resolved.cells[1].x).toBeCloseTo(92.8);
    expect(resolved.cells[1]).toMatchObject({ y: 20, width: 120 });
    expect(resolved.cells[2].x).toBeCloseTo(12.8);
    expect(resolved.cells[2]).toMatchObject({ y: 36, width: 80 });
  });

  it('honors a tblGrid wider than a single-column flow (overflows, like Word)', () => {
    // content box [20,220] = 200px; the table's grid is 200+200 = 400px wide.
    // Word renders the STORED grid and lets the table run into the right
    // margin (autofit recomputes on edit, not on open) — shrinking it here
    // desynced the columns from every margin-anchored object's x position.
    const t = table([
      [cell('a', { colwidth: [200] }), cell('b', { colwidth: [200] })],
    ]);
    const [resolved] = layoutBlocks([t], config()).pages[0].tables ?? [];
    expect(resolved.width).toBeCloseTo(400); // grid kept, spills past 220
    expect(resolved.cells[0].x).toBeCloseTo(12.8); // implicit indent shift
    expect(resolved.cells[0].width).toBeCloseTo(200);
    expect(resolved.cells[1].x).toBeCloseTo(212.8);
    expect(resolved.cells[1].width).toBeCloseTo(200);
  });

  it('still scales a wide tblGrid down inside a narrow section column', () => {
    // Two 95px columns (200 − 10 gap) / 2: the 400px grid would spill into
    // the neighbouring column's text, so THERE the clamp stays.
    const t = table([
      [cell('a', { colwidth: [200] }), cell('b', { colwidth: [200] })],
    ]);
    const cfg = { ...config(), columns: { count: 2, gap: 10 } };
    const [resolved] = layoutBlocks([t], cfg).pages[0].tables ?? [];
    expect(resolved.width).toBeCloseTo(95); // 400 → scaled to the column
    expect(resolved.cells[0].width).toBeCloseTo(47.5);
    expect(resolved.cells[1].width).toBeCloseTo(47.5);
  });

  it('still scales a nested table down to its cell', () => {
    // Outer cell is 100px wide; the nested grid claims 300px. Spilling would
    // paint over the neighbouring cell, so nested tables keep the clamp.
    const nested: FlowBlock = table([[cell('deep', { colwidth: [300] })]]);
    const outer = table([
      [
        { ...cell('host', { colwidth: [100] }), content: [nested] },
        cell('side', { colwidth: [100] }),
      ],
    ]);
    const [resolved] = layoutBlocks([outer], config()).pages[0].tables ?? [];
    const inner = resolved.cells[0].tables?.[0];
    expect(inner).toBeDefined();
    // 300px grid clamped to the host cell's content width (100 − padding).
    expect(inner && inner.width).toBeLessThanOrEqual(100);
  });

  it('does not cluster adjacent runs that differ only in tracking', () => {
    // Consecutive same-run tokens are measured CUMULATIVELY so kerning across
    // a run boundary is right. Two runs with different tracking are not the
    // same run: clustering them would measure the joined text with one side's
    // tracking. This measurer charges 1px per character so the cumulative
    // path and the per-run path give visibly different totals.
    const tracking: MeasureText = (text, f) =>
      text.length * 10 + (f.letterSpacing ?? 0) * text.length;
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [
        { text: 'ab', font: font({ letterSpacing: 1 }) },
        { text: 'cd', font: font() },
      ],
    };
    const [a, b] = layoutBlocks([block], {
      ...config(),
      measureText: tracking,
    }).pages[0].lines[0].segments;
    expect(a.width).toBe(22); // 2×10 + 2×1
    expect(b.width).toBe(20); // measured on its own, not as "abcd"
    expect(b.x).toBe(a.x + 22);
  });

  it('reduces super/subscript font size and flags the segment', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [
        { text: 'x', font: font() },
        { text: '2', font: font({ sizePt: 10 }), vertAlign: 'super' },
      ],
    };
    const [base, sup] = layoutBlocks([block], config()).pages[0].lines[0]
      .segments;
    expect(base.vertAlign).toBeUndefined();
    expect(sup.vertAlign).toBe('super');
    expect(sup.font.sizePt).toBeCloseTo(6.6); // 10 × 0.66
  });

  it('carries highlight onto segments and cell fill onto resolved cells', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hi', font: font(), background: '#FFFF00' }],
    };
    const seg = layoutBlocks([block], config()).pages[0].lines[0].segments[0];
    expect(seg.background).toBe('#FFFF00');

    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [{ ...cell('h'), background: '#D9E2F3' }] }],
    };
    const resolvedCell = layoutBlocks([t], config()).pages[0].tables?.[0]
      ?.cells[0];
    expect(resolvedCell?.background).toBe('#D9E2F3');
  });

  it('carries underline/strike flags and measured width onto segments', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'ab cd', font: font(), underline: true, strike: true }],
    };
    const { pages } = layoutBlocks([block], config());
    const [s0] = pages[0].lines[0].segments;
    expect(s0).toMatchObject({
      text: 'ab',
      underline: true,
      strike: true,
      width: 20,
    });
  });

  it('pads cell content by the Word default cell margin', () => {
    const t = table([[cell('a', { colwidth: [100] })]]);
    const { pages } = layoutBlocks([t], config());
    const line = pages[0].tables?.[0]?.cells[0].lines[0];
    // Implicit indent (−7.2) + left padding (+7.2): text sits ON the margin.
    expect(line?.segments[0].x).toBeCloseTo(20);
    expect(line?.width).toBeCloseTo(100 - 2 * 7.2);
  });

  it('aligns the table, honors trHeight and cell vAlign', () => {
    const t: FlowBlock = {
      type: 'table',
      align: 'center',
      rows: [
        {
          height: { value: 50, exact: false }, // floor → row grows to 50
          cells: [{ ...cell('x', { colwidth: [100] }), vAlign: 'center' }],
        },
      ],
    };
    const resolved = layoutBlocks([t], config()).pages[0].tables?.[0];
    // content [20,220] (200px), table 100px wide, centered → x = 20 + 50.
    expect(resolved?.x).toBeCloseTo(70);
    expect(resolved?.cells[0]).toMatchObject({ x: 70, y: 20, height: 50 });
    // vAlign center: line centered in the 50px row (content 16); +20 page top.
    expect(resolved?.cells[0].lines[0].y).toBeCloseTo(37);
  });

  it('honors per-table cell margins (w:tblCellMar)', () => {
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [cell('a', { colwidth: [100] })] }],
      cellPadding: { left: 20, right: 4, top: 6 },
    };
    const { pages } = layoutBlocks([t], config());
    const resolved = pages[0].tables?.[0];
    const line = resolved?.cells[0].lines[0];
    // Implicit indent (−20, the custom left margin) + left padding (+20).
    expect(line?.segments[0].x).toBeCloseTo(20);
    expect(line?.width).toBeCloseTo(100 - 20 - 4); // custom left + right
    expect(line?.y).toBeCloseTo(20 + 6); // custom top
    expect(resolved?.cells[0].height).toBeCloseTo(16 + 6); // top + content (+ bottom default 0)
  });

  it('honors per-cell margin overrides (w:tcMar)', () => {
    const t: FlowBlock = {
      type: 'table',
      rows: [
        {
          cells: [
            { ...cell('a', { colwidth: [100] }), padding: { left: 2, top: 3 } },
          ],
        },
      ],
    };
    const { pages } = layoutBlocks([t], config());
    const resolved = pages[0].tables?.[0];
    const line = resolved?.cells[0].lines[0];
    // Table-level pad stays default (7.2 → implicit indent −7.2); the CELL
    // overrides left to 2px, so its text sits at 12.8 + 2.
    expect(line?.segments[0].x).toBeCloseTo(12.8 + 2);
    expect(line?.y).toBeCloseTo(20 + 3); // custom top
    expect(resolved?.cells[0].height).toBeCloseTo(16 + 3);
  });

  it('emits paragraph border boxes (w:pBdr) around the placed lines', () => {
    const side = { width: 1, style: 'solid' as const, color: '#CCCCCC' };
    const bordered: FlowBlock = { ...para('one two'), borders: { top: side } };
    const { pages } = layoutBlocks([para('lead'), bordered], config());
    const boxes = pages[0].paraBorders ?? [];
    expect(boxes).toHaveLength(1);
    expect(boxes[0].borders.top).toEqual(side);
    expect(boxes[0].drawTop).toBe(true);
    expect(boxes[0].drawBottom).toBe(true);
    expect(boxes[0].x).toBe(20); // content left
    expect(boxes[0].width).toBe(200); // content width
    expect(boxes[0].y).toBe(36); // below the 16px lead line
    expect(boxes[0].height).toBe(16); // one line
  });

  it('splits a bordered paragraph across pages: top edge only on the first', () => {
    // page height 80, margins 20 → 2 lines per page; 4 lines span 2 pages.
    const side = { width: 1, style: 'solid' as const, color: '#000000' };
    const bordered: FlowBlock = {
      ...para(
        'aaaa bbbb cccc dddd eeee gggg hhhh iiii jjjj kkkk llll mmmm nnnn oooo pppp qqqq',
      ),
      borders: { top: side, bottom: side },
    };
    const cfg = config({ height: 80 });
    const { pages } = layoutBlocks([bordered], cfg);
    expect(pages.length).toBeGreaterThan(1);
    const first = pages[0].paraBorders?.[0];
    const last = pages[pages.length - 1].paraBorders?.[0];
    expect(first).toMatchObject({ drawTop: true, drawBottom: false });
    expect(last).toMatchObject({ drawTop: false, drawBottom: true });
  });

  it('measures same-font adjacent tokens cumulatively (cross-run kerning)', () => {
    // Kerning-aware fake: every "av" pair tightens the advance by 2px.
    const kerned: MeasureText = (text) => {
      const pairs = (text.match(/av/g) ?? []).length;
      return text.length * 10 - pairs * 2;
    };
    const cfg = { ...config(), measureText: kerned };
    // "a" + "v" split across two same-font runs (e.g. a mark boundary).
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [
        { text: 'a', font: font() },
        { text: 'v', font: font() },
      ],
    };
    const [line] = layoutBlocks([block], cfg).pages[0].lines;
    expect(line.segments[0]).toMatchObject({ x: 20, width: 10 });
    expect(line.segments[1]).toMatchObject({ x: 30, width: 8 }); // 18 − 10: pair kerned

    // A font change breaks the glyph run — no cross-boundary kerning.
    const mixed: FlowBlock = {
      type: 'paragraph',
      runs: [
        { text: 'a', font: font() },
        { text: 'v', font: font({ bold: true }) },
      ],
    };
    const [mline] = layoutBlocks([mixed], cfg).pages[0].lines;
    expect(mline.segments[1]).toMatchObject({ x: 30, width: 10 });
  });

  it('spans columns via colspan', () => {
    const t = table([
      [cell('wide', { colspan: 2, colwidth: [80, 120] })],
      [cell('x', { colwidth: [80] }), cell('y', { colwidth: [120] })],
    ]);
    const { pages } = layoutBlocks([t], config());
    const [resolved] = pages[0].tables ?? [];
    expect(resolved.cells[0].x).toBeCloseTo(12.8); // implicit indent
    expect(resolved.cells[0]).toMatchObject({ width: 200, colspan: 2 });
  });

  it('reserves a rowspan cell’s column so the row below shifts right', () => {
    // 3 cols (80/60/60). A rowspan-2 cell holds col 0 of rows 1-2, so the
    // 2-cell third row must land in cols 1 and 2 — not overlap col 0.
    const t = table([
      [
        cell('a', { colwidth: [80] }),
        cell('b', { colwidth: [60] }),
        cell('c', { colwidth: [60] }),
      ],
      [cell('M', { rowspan: 2, colwidth: [80] }), cell('b1'), cell('c1')],
      [cell('b2'), cell('c2')],
    ]);
    const { pages } = layoutBlocks([t], config());
    const cells = pages[0].tables?.[0].cells ?? [];
    expect(cells).toHaveLength(8); // 3 + 3 + 2
    // colX = [12.8, 92.8, 152.8] (implicit indent); merged cell holds col 0.
    expect(cells[3].x).toBeCloseTo(12.8);
    expect(cells[3]).toMatchObject({ rowspan: 2 });
    // third row: b2 → col 1, c2 → col 2. (Before the fix these overlapped the
    // merge and left col 2 empty.)
    const b2 = cells[6];
    const c2 = cells[7];
    expect(b2.x).toBeCloseTo(92.8);
    expect(b2.lines[0].segments[0].text).toBe('b2');
    expect(c2.x).toBeCloseTo(152.8);
    expect(c2.lines[0].segments[0].text).toBe('c2');
  });

  it('splits an overflowing table at the row boundary (Word-like)', () => {
    // page height 80 → content 20..60 (40px). One line (16) + table.
    const cfg = config({ height: 80 });
    const para: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hi', font: font() }],
    };
    const t = table([[cell('a')], [cell('b')]]); // 2 rows × 16 = 32px tall
    const { pages } = layoutBlocks([para, t], cfg);
    // line at y=20 → free 24px: row 'a' (16) fits, row 'b' flows to page 2.
    expect(pages).toHaveLength(2);
    expect(pages[0].lines[0].segments[0].text).toBe('hi');
    const f1 = pages[0].tables?.[0];
    expect(f1?.cells).toHaveLength(1);
    expect(f1?.cells[0].lines[0].segments[0].text).toBe('a');
    expect(f1?.height).toBeCloseTo(16);
    const f2 = pages[1].tables?.[0];
    expect(f2).toMatchObject({ y: 20 });
    expect(f2?.cells[0].lines[0]).toMatchObject({ y: 20 });
    expect(f2?.cells[0].lines[0].segments[0].text).toBe('b');
  });

  it('splits a long table between rows across pages', () => {
    // content height 60; 5 one-line rows (80px) → 3 rows + 2 rows.
    const t = table([
      [cell('r1')],
      [cell('r2')],
      [cell('r3')],
      [cell('r4')],
      [cell('r5')],
    ]);
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    const f1 = pages[0].tables?.[0];
    expect(f1?.cells.map((c) => c.lines[0].segments[0].text)).toEqual([
      'r1',
      'r2',
      'r3',
    ]);
    expect(f1?.height).toBeCloseTo(48);
    const f2 = pages[1].tables?.[0];
    expect(f2?.cells.map((c) => c.lines[0].segments[0].text)).toEqual([
      'r4',
      'r5',
    ]);
    expect(f2?.cells[0]).toMatchObject({ y: 20 });
  });

  it('repeats header rows on every continuation fragment', () => {
    // content height 60; header (16) + 4 body rows (64) = 80px total.
    // Cells carry PM positions so the ghost-vs-original distinction is real.
    const pcell = (text: string, pos: number): FlowTableCell => ({
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      content: [
        {
          type: 'paragraph',
          runs: [{ text, font: font(), pos }],
          pos,
          end: pos + text.length,
        },
      ],
    });
    const t: FlowBlock = {
      type: 'table',
      rows: [
        { cells: [pcell('head', 1)], header: true },
        { cells: [pcell('r1', 10)] },
        { cells: [pcell('r2', 20)] },
        { cells: [pcell('r3', 30)] },
        { cells: [pcell('r4', 40)] },
      ],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    // page 1: header + r1 + r2 (cut at 48 ≤ 60).
    const f1 = pages[0].tables?.[0];
    expect(f1?.cells.map((c) => c.lines[0].segments[0].text)).toEqual([
      'head',
      'r1',
      'r2',
    ]);
    // page 2: GHOST header (positions stripped) + r3 + r4.
    const f2 = pages[1].tables?.[0];
    expect(f2?.cells.map((c) => c.lines[0].segments[0].text)).toEqual([
      'head',
      'r3',
      'r4',
    ]);
    const ghost = f2?.cells[0].lines[0];
    expect(ghost?.segments[0].pos).toBeUndefined(); // not caret-addressable
    expect(ghost?.from).toBeUndefined();
    expect(f2?.cells[0]).toMatchObject({ y: 20 });
    expect(f2?.cells[1]).toMatchObject({ y: 36 }); // body resumes under the ghost
    // the original header (page 1) keeps its positions
    expect(f1?.cells[0].lines[0].segments[0].pos).toBeDefined();
  });

  it('fills the remaining page space before splitting an oversize row', () => {
    // content 60; a paragraph line (16) leaves 44px. The single row holds
    // 6 lines (96px) — taller than even a FULL page, so Word starts it in the
    // leftover space instead of abandoning it as a blank gap.
    const para6: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hi', font: font() }],
    };
    const tallCell: FlowTableCell = {
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      content: Array.from({ length: 6 }, (_, i) => para(`l${i + 1}`)),
    };
    const t: FlowBlock = { type: 'table', rows: [{ cells: [tallCell] }] };
    const { pages } = layoutBlocks([para6, t], config({ height: 100 }));
    expect(pages).toHaveLength(3);
    // page 1: the paragraph + the row's first 2 lines fill the 44px leftover.
    const f1 = pages[0].tables?.[0];
    expect(f1?.cells[0].lines.map((l) => l.segments[0].text)).toEqual([
      'l1',
      'l2',
    ]);
    expect(f1?.y).toBe(36); // right below the paragraph, no blank gap
    // pages 2–3: the re-stacked remainder continues from each page top.
    expect(
      pages[1].tables?.[0]?.cells[0].lines.map((l) => l.segments[0].text),
    ).toEqual(['l3', 'l4', 'l5']);
    expect(
      pages[2].tables?.[0]?.cells[0].lines.map((l) => l.segments[0].text),
    ).toEqual(['l6']);
  });

  it('splits an ordinary tall row in the leftover space (Word default)', () => {
    // content 60; paragraph (16) leaves 44px; one row of 3 lines (48px) does
    // not fit the leftover. Word's default lets rows break across pages, so
    // the row starts here (2 lines) and continues on page 2 — no blank gap.
    const para1: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hi', font: font() }],
    };
    const rowCell: FlowTableCell = {
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      content: Array.from({ length: 3 }, (_, i) => para(`l${i + 1}`)),
    };
    const t: FlowBlock = { type: 'table', rows: [{ cells: [rowCell] }] };
    const { pages } = layoutBlocks([para1, t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    expect(
      pages[0].tables?.[0]?.cells[0].lines.map((l) => l.segments[0].text),
    ).toEqual(['l1', 'l2']);
    expect(
      pages[1].tables?.[0]?.cells[0].lines.map((l) => l.segments[0].text),
    ).toEqual(['l3']);
  });

  it('moves a table off a sliver of band instead of overflowing the page', () => {
    // Band [20, 80] = 60px. Three filler lines leave a 12px strip — too thin
    // to hold even one row line, so the cut moves nothing. The table must go
    // to the next page: dumping it here painted it straight off the bottom
    // of the sheet (a customer form lost most of a work-history table).
    const filler: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'a b c', font: font() }],
    };
    const rowCell: FlowTableCell = {
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      content: Array.from({ length: 2 }, (_, i) => para(`r${i + 1}`)),
    };
    const t: FlowBlock = { type: 'table', rows: [{ cells: [rowCell] }] };
    const { pages } = layoutBlocks(
      [filler, filler, filler, t],
      config({ height: 100 }),
    );
    expect(pages).toHaveLength(2);
    expect(pages[0].tables ?? []).toHaveLength(0); // nothing crammed in
    const placed = pages[1].tables?.[0];
    expect(placed?.y).toBe(20); // opens the fresh page
    expect(placed && placed.y + placed.height).toBeLessThanOrEqual(80);
  });

  it('suspends vertical centering when a row splits (no empty first fragment)', () => {
    // Band [20, 80]. Two filler lines leave 28px; the row is 64px tall (cell
    // B has 4 lines) so it splits at 28. Cell A's single vAlign=center line
    // sits at y 24–40 — BELOW the cut — so the old split shipped it to page
    // 2 and painted an empty 28px box on page 1 (the customer form's
    // 'Tên công ty' header row showed 3 blank cells). Word suspends the
    // centering while the row is split: every cell's first line stays.
    const filler = para(words(2));
    const rowA: FlowTableCell = {
      ...cell('hub', { colwidth: [100] }),
      vAlign: 'center',
    };
    const rowB: FlowTableCell = {
      colspan: 1,
      rowspan: 1,
      colwidth: [100],
      content: Array.from({ length: 4 }, (_, i) => para(`b${i + 1}`)),
    };
    const t: FlowBlock = { type: 'table', rows: [{ cells: [rowA, rowB] }] };
    const { pages } = layoutBlocks([filler, t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    const [cellA0, cellB0] = pages[0].tables?.[0]?.cells ?? [];
    expect(cellA0.lines.map((l) => l.segments[0]?.text)).toEqual(['hub']);
    expect(cellB0.lines.map((l) => l.segments[0]?.text)).toEqual(['b1']);
    const [cellA1, cellB1] = pages[1].tables?.[0]?.cells ?? [];
    expect(cellA1.lines).toHaveLength(0); // its one line already landed
    expect(cellB1.lines.map((l) => l.segments[0]?.text)).toEqual([
      'b2',
      'b3',
      'b4',
    ]);
  });

  it('a split row keeps its cell borders and fill on BOTH fragments', () => {
    // Word draws a border on both sides of a row that breaks across pages
    // (a known Word behavior users work around, not a bug to imitate away).
    // The continuation used to be built field-by-field, so it silently lost
    // `borders` and `background`: a double-ruled header row came back as the
    // table's plain grid line, and a shaded cell lost its fill.
    const dbl = { style: 'double' as const, width: 0.75, color: '#000000' };
    const rowCell: FlowTableCell = {
      colspan: 1,
      rowspan: 1,
      colwidth: [100],
      background: '#EEEEEE',
      borders: { top: dbl, bottom: dbl, left: dbl, right: dbl },
      content: Array.from({ length: 4 }, (_, i) => para(`r${i + 1}`)),
    };
    const t: FlowBlock = { type: 'table', rows: [{ cells: [rowCell] }] };
    const { pages } = layoutBlocks([para('x'), t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    for (const p of pages) {
      const cell = p.tables?.[0]?.cells[0];
      expect(cell?.background).toBe('#EEEEEE');
      expect(cell?.borders?.left).toMatchObject({ style: 'double' });
      expect(cell?.borders?.top).toMatchObject({ style: 'double' });
    }
  });

  it('carries every cell field across a split unless deliberately dropped', () => {
    // Guards the CLASS of bug above rather than the two fields it hit: the
    // same table laid out with and without a page break must produce cells
    // with the same set of keys. Anything new on ResolvedCell that the
    // splitter forgets shows up here — the drop list is the only exemption,
    // and it has to be written down.
    const DELIBERATELY_DROPPED = ['vShift']; // centering pauses mid-split
    const dbl = { style: 'double' as const, width: 0.75, color: '#000000' };
    const rowCell: FlowTableCell = {
      colspan: 1,
      rowspan: 1,
      colwidth: [100],
      background: '#EEEEEE',
      vAlign: 'center',
      borders: { top: dbl, bottom: dbl, left: dbl, right: dbl },
      content: Array.from({ length: 4 }, (_, i) => para(`r${i + 1}`)),
    };
    const t: FlowBlock = { type: 'table', rows: [{ cells: [rowCell] }] };
    // Tall page: one piece, every field present. Short page: it splits.
    const whole = layoutBlocks([para('x'), t], config({ height: 400 })).pages[0]
      .tables?.[0]?.cells[0];
    const split = layoutBlocks([para('x'), t], config({ height: 100 })).pages[1]
      .tables?.[0]?.cells[0];
    expect(whole && split).toBeTruthy();
    const w = whole as unknown as Record<string, unknown>;
    const s = split as unknown as Record<string, unknown>;
    // Compare fields that actually carry a value: some clone helpers spread
    // `key: undefined` through, which is not a lost field.
    const missing = Object.keys(w).filter(
      (k) =>
        w[k] !== undefined &&
        !DELIBERATELY_DROPPED.includes(k) &&
        s[k] === undefined,
    );
    expect(missing).toEqual([]);
  });

  it('re-bases cantSplit bands onto each fragment', () => {
    // The bands are measured from the table's top, so a fragment that starts
    // mid-table must renumber them — copying them verbatim would apply
    // "never break this row" to whatever row now sits at those coordinates.
    const plain = (n: number): FlowTableCell => ({
      colspan: 1,
      rowspan: 1,
      colwidth: [100],
      content: Array.from({ length: n }, (_, i) => para(`p${i}`)),
    });
    const t: FlowBlock = {
      type: 'table',
      rows: [
        { cells: [plain(3)] }, // 48px, ordinary — gets cut mid-row
        { cells: [plain(3)], cantSplit: true }, // 48px, must stay whole
      ],
    };
    const { pages } = layoutBlocks([para('x'), t], config({ height: 100 }));
    // The cantSplit row ends up alone on the last fragment, and its band is
    // now measured from THAT fragment's top (0..48), not the original 48..96.
    const last = pages[pages.length - 1].tables?.[0];
    expect(last?.cantSplitBands).toHaveLength(1);
    expect(last?.cantSplitBands?.[0].top).toBeCloseTo(0);
    expect(last?.cantSplitBands?.[0].bottom).toBeCloseTo(48);
    // Fragments that no longer contain the row carry no band at all.
    expect(pages[0].tables?.[0]?.cantSplitBands).toBeUndefined();
  });

  it('moves a w:cantSplit row whole when it fits a fresh page', () => {
    // Same geometry, but the row is marked cantSplit → the old behavior:
    // leave the gap, start the intact row on page 2.
    const para1: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hi', font: font() }],
    };
    const rowCell: FlowTableCell = {
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      content: Array.from({ length: 3 }, (_, i) => para(`l${i + 1}`)),
    };
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [rowCell], cantSplit: true }],
    };
    const { pages } = layoutBlocks([para1, t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    expect(pages[0].tables ?? []).toHaveLength(0); // gap left intentionally
    expect(pages[1].tables?.[0]?.cells[0].lines).toHaveLength(3); // intact row
  });

  it('carries border visibility onto split fragments', () => {
    const side = { width: 1, style: 'solid', color: '#000000' } as const;
    const t: FlowBlock = {
      type: 'table',
      rows: [
        [cell('r1')],
        [cell('r2')],
        [cell('r3')],
        [cell('r4')],
        [cell('r5')],
      ].map((cells) => ({ cells })),
      borders: { top: side, insideH: side },
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    expect(pages[0].tables?.[0]?.borders).toEqual({ top: side, insideH: side });
    expect(pages[1].tables?.[0]?.borders).toEqual({ top: side, insideH: side });
  });

  it('splits a row taller than the page mid-row and re-stacks the remainder', () => {
    // content height 60. Row 1: cell with 5 paragraphs (80px) > full page.
    const tallCell: FlowTableCell = {
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      content: Array.from({ length: 5 }, (_, i) => para(`l${i + 1}`)),
    };
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [tallCell] }, { cells: [cell('after')] }],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    // page 1: the first 3 lines fit in the 60px band (cut at the line boundary).
    const f1 = pages[0].tables?.[0];
    expect(f1?.cells[0].lines.map((l) => l.segments[0].text)).toEqual([
      'l1',
      'l2',
      'l3',
    ]);
    // page 2: remaining lines re-stack from the page top; the next row follows.
    const f2 = pages[1].tables?.[0];
    expect(f2?.cells[0].lines.map((l) => l.segments[0].text)).toEqual([
      'l4',
      'l5',
    ]);
    expect(f2?.cells[0].lines[0]).toMatchObject({ y: 20 });
    expect(f2?.cells[1].lines[0].segments[0].text).toBe('after');
    expect(f2?.cells[1].y).toBeCloseTo(52); // 20 + continuation row (32)
  });

  it('carries segment positions and line from/to across wraps', () => {
    // 24 chars × 10px > 200px content width; wraps before "eeee".
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'aaaa bbbb cccc dddd eeee', font: font(), pos: 1 }],
      pos: 1,
      end: 25,
    };
    const { pages } = layoutBlocks([block], config());
    const [line0, line1] = pages[0].lines;
    expect(line0.segments.map((s) => s.pos)).toEqual([1, 5, 6, 10, 11, 15, 16]); // words + spaces
    expect(line0.from).toBe(1);
    expect(line0.to).toBe(20); // after "dddd" (trailing space trimmed)
    expect(line1.from).toBe(21);
    expect(line1.to).toBe(25);
  });

  it('collapses an empty paragraph line to its content position', () => {
    const block: FlowBlock = { type: 'paragraph', runs: [], pos: 7, end: 7 };
    const { pages } = layoutBlocks([block], config());
    expect(pages[0].lines[0]).toMatchObject({ from: 7, to: 7 });
  });

  it('renders typed leading spaces on a first line, drops them after a wrap', () => {
    // Real documents right-position text ("Ký tên") with a run of spaces —
    // those must render. Spaces landing at a soft-wrapped line start don't.
    const first: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: '     x', font: font() }], // 5 leading spaces @10px each
    };
    const { pages } = layoutBlocks([first], config());
    const segs = pages[0].lines[0].segments;
    // The glyph lands 50px past the content left edge (spaces painted).
    expect(segs.at(-1)?.x).toBeCloseTo(20 + 50);

    // A paragraph wide enough to wrap: the continuation line still starts
    // flush at the content edge (no leading space), as before.
    const wrapping: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'aaaaaaaaaa '.repeat(8).trim(), font: font() }],
    };
    const wrapped = layoutBlocks([wrapping], config()).pages[0].lines;
    expect(wrapped.length).toBeGreaterThan(1);
    expect(wrapped[1].segments[0].x).toBeCloseTo(20);
    expect(wrapped[1].segments[0].text?.startsWith(' ')).toBe(false);
  });

  it('breaks a word at character level when the band cannot fit it (narrow cell)', () => {
    // "STT" in a header cell narrower than one word: Word stacks S/T/T.
    const cellPara: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'STT', font: font(), pos: 10 }],
      pos: 9,
    };
    const table: FlowBlock = {
      type: 'table',
      // colwidth 26 − 2×7.2 padding → 11.6px band: one 10px char per line.
      rows: [
        {
          cells: [
            { colspan: 1, rowspan: 1, colwidth: [26], content: [cellPara] },
          ],
        },
      ],
    };
    const { pages } = layoutBlocks([table], config());
    const cell = pages[0].tables?.[0].cells[0];
    expect(cell?.lines.map((l) => l.segments[0]?.text)).toEqual([
      'S',
      'T',
      'T',
    ]);
    // PM positions follow the split characters (caret/selection stay usable).
    expect(cell?.lines.map((l) => l.segments[0]?.pos)).toEqual([10, 11, 12]);
  });

  it('applies paragraph spacing (before/after gaps + line multiplier)', () => {
    const cfg: LayoutConfig = {
      ...config(),
      measureMetrics: () => ({ ascent: 12, descent: 4 }),
    };
    const a: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'a', font: font() }],
    };
    const b: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'b', font: font() }],
      spacing: { before: 10, after: 6, line: 2, lineRule: 'auto' },
    };
    const c: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'c', font: font() }],
    };
    const { pages } = layoutBlocks([a, b, c], cfg);
    const [la, lb, lc] = pages[0].lines;
    expect(la.y).toBe(20); // top margin
    // b: before=10 gap → y 36+10=46; line 2× of 16 → height 32.
    expect(lb.y).toBe(46);
    expect(lb.height).toBe(32);
    // c follows b's line (32) + after gap (6): 46+32+6 = 84.
    expect(lc.y).toBe(84);
  });

  it('forces a new line at a hard break', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [
        { text: 'a', font: font(), pos: 1 },
        { break: true, pos: 2 },
        { text: 'b', font: font(), pos: 3 },
      ],
    };
    const { pages } = layoutBlocks([block], config());
    expect(pages[0].lines).toHaveLength(2);
    expect(pages[0].lines[0].segments[0].text).toBe('a');
    expect(pages[0].lines[1].segments[0].text).toBe('b');
  });

  it('starts a pageBreakBefore paragraph on a new page', () => {
    const a: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'a', font: font() }],
    };
    const b: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'b', font: font() }],
      pageBreakBefore: true,
    };
    const { pages } = layoutBlocks([a, b], config());
    expect(pages).toHaveLength(2);
    expect(pages[1].lines[0].segments[0].text).toBe('b');
  });

  it('uses injected font metrics for the line box and baseline', () => {
    const cfg: LayoutConfig = {
      ...config(),
      measureMetrics: () => ({ ascent: 12, descent: 4 }),
    };
    const { pages } = layoutBlocks([para('hi')], cfg);
    const [line] = pages[0].lines;
    expect(line.height).toBe(16); // ascent + descent
    expect(line.baseline).toBe(12); // baseline sits at the ascent
  });

  it('jumps to custom left tab stops and fills dot leaders', () => {
    // content left = 20; stop at +100 with a dot leader (TOC style).
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'Ch1\t9', font: font(), pos: 1 }],
      tabs: [{ pos: 100, val: 'left', leader: 'dot' }],
    };
    const { pages } = layoutBlocks([block], config());
    const segs = pages[0].lines[0].segments;
    expect(segs[0]).toMatchObject({ text: 'Ch1', x: 20 });
    // tab spans 50..120 (70px) → leader of 6 dots (70/10 − 1), pos stripped.
    expect(segs[1].text).toBe('......');
    expect(segs[1].pos).toBeUndefined();
    expect(segs[2]).toMatchObject({ text: '9', x: 120 }); // lands at the stop
  });

  it('right-aligns and decimal-aligns tab groups at their stops', () => {
    const right: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'a\tend', font: font() }],
      tabs: [{ pos: 150, val: 'right' }],
    };
    const r = layoutBlocks([right], config()).pages[0].lines[0].segments;
    // group "end" (30px) ENDS at 20+150 → starts at 140.
    expect(r.at(-1)).toMatchObject({ text: 'end', x: 140 });

    const decimal: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'x\t123.45', font: font() }],
      tabs: [{ pos: 100, val: 'decimal' }],
    };
    const d = layoutBlocks([decimal], config()).pages[0].lines[0].segments;
    // separator sits at the stop: "123" (30px) before 20+100 → group at 90.
    expect(d.at(-1)).toMatchObject({ text: '123.45', x: 90 });
  });

  it('falls back to the default grid past the last custom stop', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'aaaaaaaaaaaa\tb', font: font() }], // 12 chars → x=140 > stop 100
      tabs: [{ pos: 100, val: 'left' }],
    };
    const { pages } = layoutBlocks([block], { ...config(), tabWidth: 50 });
    // past the stop → default grid: next multiple of 50 from 20 → 170.
    expect(pages[0].lines[0].segments.at(-1)).toMatchObject({
      text: 'b',
      x: 170,
    });
  });

  it('advances a tab to the next tab stop', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'a\tb', font: font() }],
    };
    const { pages } = layoutBlocks([block], { ...config(), tabWidth: 50 });
    const segs = pages[0].lines[0].segments;
    // 'a'@20 (10px) → tab from x=30 to the next stop (20 + 50 = 70) → 'b'@70.
    expect(segs[0]).toMatchObject({ text: 'a', x: 20 });
    expect(segs.at(-1)).toMatchObject({ text: 'b', x: 70 });
  });
});

describe('toFlowBlocks', () => {
  const schema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
      image: {
        inline: true,
        group: 'inline',
        attrs: { src: {}, width: { default: null }, height: { default: null } },
      },
      table: { group: 'block', content: 'table_row+' },
      table_row: { content: 'table_cell+' },
      table_cell: {
        content: 'block+',
        attrs: {
          colspan: { default: 1 },
          rowspan: { default: 1 },
          colwidth: { default: null },
        },
      },
    },
    marks: { strong: {}, fontSize: { attrs: { size: {} } } },
  });

  it('resolves marks into run fonts', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('Hi', [schema.mark('strong')]),
        schema.text('Big', [schema.mark('fontSize', { size: 20 })]),
      ]),
    ]);
    const block = toFlowBlocks(doc, { sizePt: 10 })[0];
    expect(block.type).toBe('paragraph');
    if (block.type !== 'paragraph') return;
    const [r0, r1] = block.runs;
    // both are text runs; narrow off the FlowInline union (text has `text`).
    expect('text' in r0).toBe(true);
    if ('text' in r0) expect(r0.font.bold).toBe(true);
    if ('text' in r1) expect(r1.font.sizePt).toBe(20);
  });

  it('emits an inline image from an image node', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('a'),
        schema.node('image', { src: 'u', width: 40, height: 20 }),
      ]),
    ]);
    const block = toFlowBlocks(doc)[0];
    expect(block.type).toBe('paragraph');
    if (block.type !== 'paragraph') return;
    expect(block.runs).toHaveLength(2);
    const img = block.runs[1];
    expect('src' in img).toBe(true);
    if ('src' in img)
      expect(img).toMatchObject({ src: 'u', width: 40, height: 20 });
  });

  it('carries the list marker', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { list: { marker: '1.' } }, [
        schema.text('item'),
      ]),
    ]);
    const block = toFlowBlocks(doc)[0];
    expect(block.type === 'paragraph' && block.marker).toBe('1.');
  });

  it('threads absolute PM positions through paragraphs, images and cells', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('ab'),
        schema.node('image', { src: 'u', width: 10, height: 10 }),
        schema.text('cd'),
      ]),
      schema.node('table', null, [
        schema.node('table_row', null, [
          schema.node('table_cell', null, [
            schema.node('paragraph', null, [schema.text('x')]),
          ]),
        ]),
      ]),
    ]);
    const [para, table] = toFlowBlocks(doc);
    // doc → paragraph at 0, content starts at 1: "ab"@1, image@3, "cd"@4.
    if (para.type !== 'paragraph') throw new Error('expected paragraph');
    expect(para.pos).toBe(1);
    expect(para.end).toBe(6);
    expect(para.runs.map((r) => r.pos)).toEqual([1, 3, 4]);
    // table at 7 → row at 8 → cell at 9 → inner paragraph at 10, content at 11.
    if (table.type !== 'table') throw new Error('expected table');
    const inner = table.rows[0].cells[0].content[0];
    expect(inner.type === 'paragraph' && inner.pos).toBe(11);
  });

  it('flattens a table into rows and cells', () => {
    const p = (t: string) => schema.node('paragraph', null, [schema.text(t)]);
    const td = (t: string, attrs: Record<string, unknown> = {}) =>
      schema.node('table_cell', attrs, [p(t)]);
    const doc = schema.node('doc', null, [
      schema.node('table', null, [
        schema.node('table_row', null, [
          td('a', { colwidth: [80] }),
          td('b', { colwidth: [120] }),
        ]),
        schema.node('table_row', null, [td('c'), td('d', { colspan: 1 })]),
      ]),
    ]);
    const block = toFlowBlocks(doc)[0];
    expect(block.type).toBe('table');
    if (block.type === 'table') {
      expect(block.rows).toHaveLength(2);
      expect(block.rows[0].cells).toHaveLength(2);
      expect(block.rows[0].cells[0].colwidth).toEqual([80]);
      // cell content flattened recursively into FlowBlocks.
      const inner = block.rows[0].cells[0].content[0];
      expect(inner.type).toBe('paragraph');
    }
  });
});

describe('floats in table cells', () => {
  const schema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
      image: {
        inline: true,
        group: 'inline',
        attrs: {
          src: {},
          width: { default: null },
          height: { default: null },
          float: { default: null },
          shape: { default: null },
        },
      },
      table: { group: 'block', content: 'table_row+' },
      table_row: { content: 'table_cell+' },
      table_cell: {
        content: 'block+',
        attrs: {
          colspan: { default: 1 },
          rowspan: { default: 1 },
          colwidth: { default: null },
        },
      },
    },
  });

  it('keeps anchored floats when flattening cell content (no inline degrade)', () => {
    const doc = schema.node('doc', null, [
      schema.node('table', null, [
        schema.node('table_row', null, [
          schema.node('table_cell', null, [
            schema.node('paragraph', null, [
              schema.text('x'),
              schema.node('image', {
                src: '',
                width: 18,
                height: 16,
                float: {
                  wrap: 'square',
                  hOffset: 30,
                  vOffset: 5,
                  vRel: 'paragraph',
                },
                shape: { kind: 'rect', stroke: '#4472C4', strokeWidth: 2 },
              }),
            ]),
          ]),
        ]),
      ]),
    ]);
    const block = toFlowBlocks(doc)[0];
    if (block.type !== 'table') throw new Error('expected table');
    const para = block.rows[0].cells[0].content[0];
    if (para.type !== 'paragraph') throw new Error('expected paragraph');
    expect(para.floats).toHaveLength(1);
    expect(para.floats?.[0]).toMatchObject({
      width: 18,
      shape: { kind: 'rect' },
    });
    // The carrying image node's PM position rides along (resize hit-testing).
    expect(typeof para.floats?.[0].pos).toBe('number');
    expect(para.runs.filter((r) => 'src' in r)).toHaveLength(0); // not inline
  });

  it('positions the float at its anchor offsets inside the cell box', () => {
    const cellPara: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hello', font: font() }],
      floats: [
        {
          src: '',
          width: 18,
          height: 16,
          wrap: 'square',
          hOffset: 30,
          vOffset: 5,
          vRel: 'paragraph',
          shape: { kind: 'rect', stroke: '#4472C4', strokeWidth: 2 },
        },
      ],
    };
    const table: FlowBlock = {
      type: 'table',
      rows: [
        {
          cells: [
            { colspan: 1, rowspan: 1, colwidth: [120], content: [cellPara] },
          ],
        },
      ],
    };
    const { pages } = layoutBlocks([table], config());
    const cell = pages[0].tables?.[0].cells[0];
    expect(cell?.floats).toHaveLength(1);
    const f = cell?.floats?.[0];
    // x = cell content left (cell.x + 7.2 pad) + hOffset; y = pad top + vOffset.
    expect(f?.x).toBeCloseTo((cell?.x ?? 0) + 7.2 + 30);
    expect(f?.y).toBeCloseTo((cell?.y ?? 0) + 5);
    expect(f?.shape).toMatchObject({ kind: 'rect', stroke: '#4472C4' });
    // v1: the cell text does not wrap around the float — one full-band line.
    expect(cell?.lines).toHaveLength(1);
  });

  it('positions anchored floats in the header band (chrome)', () => {
    // The horizontal rule real headers draw under their contact block.
    const header = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('contact'),
        schema.node('image', {
          src: '',
          width: 200,
          height: 0,
          float: { wrap: 'none', hOffset: -10, vOffset: 30, vRel: 'paragraph' },
          shape: { kind: 'line', stroke: '#000000', strokeWidth: 1 },
        }),
      ]),
    ]);
    const body = schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('body')]),
    ]);
    const resolved = layout(body, config(), undefined, { header });
    const f = resolved.pageHeader?.floats?.[0];
    // Chrome pins at 48: y = 48 + paragraph top (0) + vOffset; x = left 20 − 10.
    expect(f).toMatchObject({ x: 10, y: 78, width: 200 });
    expect(f?.shape).toMatchObject({ kind: 'line' });
  });
});

describe('rotated inline images', () => {
  it('grows the line box to the rotated bounding-box height (width untouched)', () => {
    // 100×20 image rotated 90°: the painted box is 20 wide × 100 tall around
    // the same center, so the line needs ascent 60 (h/2 + rotH/2) and descent
    // 40 — but the token still advances 100 horizontally (no re-wrap).
    const metrics = (f: FontSpec) => ({
      ascent: sizeAscent(f),
      descent: sizeDescent(f),
    });
    const sizeAscent = (f: FontSpec) => f.sizePt * (96 / 72) * 0.8;
    const sizeDescent = (f: FontSpec) => f.sizePt * (96 / 72) * 0.25;
    const cfg: LayoutConfig = { ...config(), measureMetrics: metrics };
    const para: FlowBlock = {
      type: 'paragraph',
      runs: [
        { text: 'x', font: font() },
        { src: 'a', width: 100, height: 20, rotation: 90 },
      ],
    };
    const { pages } = layoutBlocks([para], cfg);
    const line = pages[0].lines[0];
    expect(line.baseline).toBeCloseTo(60); // ascent: 20/2 + 100/2
    expect(line.height).toBeCloseTo(100); // + descent (100−20)/2 = 40
    expect(line.images?.[0].width).toBe(100); // horizontal advance unchanged
    // Unrotated control: the line only needs the image height above baseline.
    const flat: FlowBlock = {
      type: 'paragraph',
      runs: [{ src: 'a', width: 100, height: 20 }],
    };
    const l2 = layoutBlocks([flat], cfg).pages[0].lines[0];
    expect(l2.baseline).toBeCloseTo(20);
  });
});

describe('textbox floats', () => {
  it('flows textbox paragraphs inside the box (box-local lines, stripped positions)', () => {
    // Box 120×60 anchored at hOffset 40 / vOffset 10. Interior: default inset
    // l 10, r 10 → text band [10, 110] = 100px → "aaaa bbbb cccc" (14 chars,
    // 40px words) wraps into 2 lines at y = inset.t (5) and 5 + 16.
    const boxPara: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'aaaa bbbb cccc', font: font(), pos: 123 }],
      pos: 122,
      end: 137,
    };
    const host: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'body', font: font() }],
      floats: [
        {
          src: '',
          width: 120,
          height: 60,
          wrap: 'none',
          hOffset: 40,
          vOffset: 10,
          vRel: 'paragraph',
          shape: { kind: 'rect', stroke: '#000000', strokeWidth: 1 },
          content: [boxPara],
          pos: 42,
        },
      ],
    };
    const { pages } = layoutBlocks([host], config());
    const f = pages[0].floats?.[0];
    expect(f).toMatchObject({ x: 60, y: 30, width: 120, height: 60, pos: 42 });
    // Lines are box-local: x from the interior inset, y from the box top.
    expect(f?.lines).toHaveLength(2);
    expect(f?.lines?.[0].segments.map((s) => s.text).join('')).toBe(
      'aaaa bbbb',
    );
    expect(f?.lines?.[0].segments[0].x).toBe(10);
    expect(f?.lines?.[0].y).toBe(5);
    expect(f?.lines?.[1].segments.map((s) => s.text)).toEqual(['cccc']);
    expect(f?.lines?.[1].y).toBe(21); // 5 + 16px line
    // Never caret-addressable — PM positions are stripped.
    expect(f?.lines?.[0].segments[0].pos).toBeUndefined();
  });
});

describe('floating images', () => {
  const words = (n: number, len = 9) =>
    Array.from({ length: n }, () => 'a'.repeat(len)).join(' ');

  it('narrows lines beside a right-aligned square float, resumes below it', () => {
    // content [20,220]; float 80×40 at right → rect x 140..220, y 20..60.
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: words(8), font: font() }],
      floats: [
        { src: 'f', width: 80, height: 40, wrap: 'square', hAlign: 'right' },
      ],
    };
    const { pages } = layoutBlocks([block], config());
    const lines = pages[0].lines;
    // Lines overlapping the float band [20,60) get the 120px band; the rest 200px.
    const narrowed = lines.filter((l) => l.y < 60);
    const full = lines.filter((l) => l.y >= 60);
    expect(narrowed.length).toBeGreaterThan(0);
    expect(full.length).toBeGreaterThan(0);
    for (const l of narrowed) expect(l.width).toBeCloseTo(120); // 140 − 20
    for (const l of full) expect(l.width).toBeCloseTo(200);
    // The float itself lands on the page for the painter.
    expect(pages[0].floats?.[0]).toMatchObject({
      x: 140,
      y: 20,
      width: 80,
      height: 40,
    });
  });

  it('honors text-to-image gaps (distL)', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: words(6), font: font() }],
      floats: [
        {
          src: 'f',
          width: 80,
          height: 40,
          wrap: 'square',
          hAlign: 'right',
          distL: 10,
        },
      ],
    };
    const { pages } = layoutBlocks([block], config());
    expect(pages[0].lines[0].width).toBeCloseTo(110); // band ends 10px before the image
  });

  it("w:line 'auto' scales text but never an inline image's height", () => {
    // Word/Google Docs size an image line to the image — the line multiple
    // applies to text boxes only. Multiplying the image (319×1.08 ≈ +25px per
    // picture) paginated real documents a page earlier than Word.
    const image = { src: 'shot.png', width: 100, height: 319 };
    const mk = (spacing?: { line: number }): FlowBlock => ({
      type: 'paragraph',
      runs: [image],
      ...(spacing ? { spacing } : {}),
    });
    const plain = layoutBlocks([mk()], config()).pages[0].lines[0];
    const auto = layoutBlocks([mk({ line: 259 / 240 })], config()).pages[0]
      .lines[0];
    expect(auto.height).toBeCloseTo(plain.height); // image box untouched
    // …while a TEXT line still grows by the multiple.
    const t = (spacing?: { line: number }): FlowBlock => ({
      type: 'paragraph',
      runs: [{ text: 'hello', font: font() }],
      ...(spacing ? { spacing } : {}),
    });
    const tPlain = layoutBlocks([t()], config()).pages[0].lines[0];
    const tAuto = layoutBlocks([t({ line: 2 })], config()).pages[0].lines[0];
    expect(tAuto.height).toBeCloseTo(tPlain.height * 2);
  });

  it('moves an inline image below square floats it cannot fit between', () => {
    // Geometry from a real report (CEA template): column 624px wide, two
    // floating text-box labels at +88 and +410 (107px wide, dist 12), and an
    // inline 300×328 screenshot in the same paragraph. The gap between the
    // labels is 191px — over MIN_BAND, so the band was accepted and the image
    // overflowed 97px into the right label. Word drops the image below the
    // labels instead; so do we, via the band's minWidth.
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ src: 'big.png', width: 300, height: 328 }],
      floats: [
        {
          src: 'l1',
          width: 107,
          height: 32,
          wrap: 'square',
          hOffset: 88,
          hRel: 'margin',
          vOffset: 14,
          vRel: 'paragraph',
          distL: 12,
          distR: 12,
          distT: 5,
          distB: 5,
        },
        {
          src: 'l2',
          width: 107,
          height: 32,
          wrap: 'square',
          hOffset: 410,
          hRel: 'margin',
          vOffset: 15,
          vRel: 'paragraph',
          distL: 12,
          distR: 12,
          distT: 5,
          distB: 5,
        },
      ],
    };
    const { pages } = layoutBlocks([block], {
      measureText: measure,
      defaultFont: { sizePt: 10 },
      page: {
        width: 816,
        height: 1056,
        margin: { top: 96, right: 96, bottom: 96, left: 96 },
      },
    });
    const pg = pages[0];
    const line = pg.lines[0];
    const img = line.images![0];
    // Below both labels' exclusion rects (bottom = 96+15+32+5 = 148; the walk
    // stops at the lower blocker bottom that clears the width, 147→148).
    expect(line.y).toBeGreaterThanOrEqual(147);
    // Full-width band → image starts at the margin and clears the labels.
    expect(img.x).toBeCloseTo(96);
    const l2 = pg.floats!.find((f) => f.src === 'l2')!;
    expect(img.x + img.width).toBeLessThanOrEqual(l2.x);
  });

  it('still overflows an image wider than the whole column', () => {
    // No band can ever fit it — the old behavior (place and overflow) must
    // survive, and the band walk must not loop forever looking for one.
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ src: 'huge.png', width: 300, height: 40 }],
      floats: [
        { src: 'f', width: 80, height: 40, wrap: 'square', hAlign: 'right' },
      ],
    };
    const { pages } = layoutBlocks([block], config()); // column 200px wide
    const line = pages[0].lines[0];
    expect(line.images![0].width).toBe(300);
    // It skipped below the float (band there is the full 200px column) rather
    // than squeezing beside it.
    expect(line.y).toBeGreaterThanOrEqual(60);
  });

  it('a paragraph pushed to the next page takes its floats with it', () => {
    // Page content band [20, 980] (height 1000, margins 20). The filler eats
    // most of it; the float-carrying paragraph's first line no longer fits, so
    // the paragraph breaks to page 2 — and the float must register THERE, at
    // page-2 coordinates. Registering before the break decision left it on
    // page 1, positioned in page-1's space (the drag-to-page-2 bounce report).
    const filler: FlowBlock = {
      type: 'paragraph',
      // 4 words/line ('aaaa ' = 50px in the 200px band) × 60 lines × 16px
      // = 960px: exactly the content band, so the NEXT paragraph breaks.
      runs: [
        {
          text: Array(60 * 4)
            .fill('aaaa')
            .join(' '),
          font: font(),
        },
      ],
    };
    const anchored: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hello', font: font() }],
      floats: [
        {
          src: 'f',
          width: 40,
          height: 30,
          wrap: 'none',
          vRel: 'margin',
          vOffset: 10,
          hOffset: 5,
        },
      ],
    };
    const { pages } = layoutBlocks(
      [filler, anchored],
      config({ height: 1000 }),
    );
    expect(pages).toHaveLength(2);
    expect(pages[0].floats ?? []).toHaveLength(0);
    expect(pages[1].floats).toHaveLength(1);
    // vRel margin on the page it actually starts on: y = top(20) + 10.
    expect(pages[1].floats![0]).toMatchObject({ x: 25, y: 30 });
    expect(pages[1].lines.length).toBeGreaterThan(0); // text came along too
  });

  it('topAndBottom floats push the text below them', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'hello', font: font() }],
      floats: [{ src: 'f', width: 50, height: 30, wrap: 'topAndBottom' }],
    };
    const { pages } = layoutBlocks([block], config());
    expect(pages[0].floats?.[0]).toMatchObject({ y: 20 });
    expect(pages[0].lines[0].y).toBeCloseTo(50); // below the float band
  });

  it('affects following paragraphs while the float is still in the way', () => {
    // Para 1: one line + a tall right float (height 100 → y 20..120).
    const p1: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'short', font: font() }],
      floats: [
        { src: 'f', width: 80, height: 100, wrap: 'square', hAlign: 'right' },
      ],
    };
    // Para 2: no floats of its own — would normally take the drafts fast path.
    const p2: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: words(8), font: font() }],
    };
    const { pages } = layoutBlocks([p1, p2], config());
    const p2Lines = pages[0].lines.slice(1);
    const narrowed = p2Lines.filter((l) => l.y < 120);
    expect(narrowed.length).toBeGreaterThan(0);
    for (const l of narrowed) expect(l.width).toBeCloseTo(120);
  });

  it("'none' wrap paints the float without touching the text", () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: words(4), font: font() }],
      floats: [
        { src: 'f', width: 80, height: 40, wrap: 'none', hAlign: 'right' },
      ],
    };
    const { pages } = layoutBlocks([block], config());
    expect(pages[0].floats).toHaveLength(1);
    for (const l of pages[0].lines) expect(l.width).toBeCloseTo(200); // untouched
  });
});

describe('layout with page chrome (header/footer)', () => {
  const schema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const p = (text: string) =>
    schema.node('paragraph', null, [schema.text(text)]);
  const docOf = (...texts: string[]) => schema.node('doc', null, texts.map(p));

  it('pins the header/footer bands and shrinks the body band', () => {
    // page 300 tall, margins 20 — chrome distance 48 dominates the margins.
    const cfg = { ...config({ height: 300 }) };
    const resolved = layout(docOf('body'), cfg, undefined, {
      header: docOf('head'),
      footer: docOf('foot'),
    });
    // header pinned at 48; its PM positions are stripped.
    expect(resolved.pageHeader?.lines[0]).toMatchObject({ y: 48 });
    expect(resolved.pageHeader?.lines[0].segments[0].pos).toBeUndefined();
    expect(resolved.pageHeader?.lines[0].from).toBeUndefined();
    // footer bottom sits at pageHeight − 48.
    const f = resolved.pageFooter?.lines[0];
    expect((f?.y ?? 0) + (f?.height ?? 0)).toBeCloseTo(300 - 48);
    // body starts below the header band (48 + 16 > margin 20).
    expect(resolved.pages[0].lines[0]).toMatchObject({ y: 64 });
  });

  it('lays out page-number fields sized by the page total', () => {
    const fieldSchema = new Schema({
      nodes: {
        doc: { content: 'block+' },
        paragraph: {
          group: 'block',
          content: 'inline*',
          attrs: { list: { default: null } },
        },
        text: { group: 'inline' },
        page_field: {
          inline: true,
          group: 'inline',
          atom: true,
          attrs: { kind: {} },
        },
      },
      marks: {},
    });
    const footer = fieldSchema.node('doc', null, [
      fieldSchema.node('paragraph', null, [
        fieldSchema.text('Trang '),
        fieldSchema.node('page_field', { kind: 'page' }),
      ]),
    ]);
    // 12 body lines → 2 pages (content band shrinks for the footer).
    const body = fieldSchema.node(
      'doc',
      null,
      Array.from({ length: 30 }, (_, i) =>
        fieldSchema.node('paragraph', null, [fieldSchema.text(`p${i}`)]),
      ),
    );
    const resolved = layout(body, { ...config({ height: 300 }) }, undefined, {
      footer,
    });
    const segs = resolved.pageFooter?.lines[0].segments ?? [];
    const fieldSeg = segs.find((s) => s.field);
    expect(fieldSeg?.field).toBe('pageNumber');
    // 2 pages → second pass measures with '2' (1 digit × 10px).
    expect(resolved.pages.length).toBeGreaterThan(1);
    expect(fieldSeg?.width).toBe(10 * String(resolved.pages.length).length);
    expect(fieldSeg?.pos).toBeUndefined(); // chrome stays non-addressable
  });

  it('paginates against the shrunken body band', () => {
    // content band: top 64, bottom 300−48−16 = 236 → 172px ≈ 10 lines of 16.
    const cfg = { ...config({ height: 300 }) };
    const body = schema.node(
      'doc',
      null,
      Array.from({ length: 12 }, (_, i) => p(`p${i}`)),
    );
    const resolved = layout(body, cfg, undefined, {
      header: docOf('head'),
      footer: docOf('foot'),
    });
    expect(resolved.pages).toHaveLength(2);
    expect(resolved.pages[0].lines.length).toBe(10);
    const last = resolved.pages[0].lines.at(-1);
    expect((last?.y ?? 0) + (last?.height ?? 0)).toBeLessThanOrEqual(236.01);
  });

  it('lays out first/even chrome variants and records the selection', () => {
    const cfg = { ...config({ height: 400 }) };
    const body = schema.node(
      'doc',
      null,
      Array.from({ length: 20 }, (_, i) => p(`p${i}`)),
    );
    const resolved = layout(body, cfg, undefined, {
      header: docOf('def'),
      headerFirst: docOf('first'),
      headerEven: docOf('even'),
      titlePg: true,
      evenAndOdd: true,
    });
    expect(resolved.pageHeader?.lines[0].segments[0].text).toBe('def');
    expect(resolved.pageHeaderFirst?.lines[0].segments[0].text).toBe('first');
    expect(resolved.pageHeaderEven?.lines[0].segments[0].text).toBe('even');
    expect(resolved.chromeSelect).toEqual({ titlePg: true, evenAndOdd: true });
  });
});

describe('layout with footnotes', () => {
  const fnSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
    },
    marks: {
      footnote: { attrs: { num: {} } },
      vertAlign: { attrs: { value: {} } },
    },
  });
  const p = (text: string) =>
    fnSchema.node('paragraph', null, [fnSchema.text(text)]);
  // A paragraph ending in a footnote reference (superscript number + mark).
  const refPara = (text: string, num: number) =>
    fnSchema.node('paragraph', null, [
      fnSchema.text(text),
      fnSchema.text(String(num), [
        fnSchema.mark('footnote', { num }),
        fnSchema.mark('vertAlign', { value: 'super' }),
      ]),
    ]);
  const noteDoc = (text: string) => fnSchema.node('doc', null, [p(text)]);

  it('reserves bottom space and lays the note body there', () => {
    const cfg = config({ height: 300 }); // body band [20, 280]
    const body = fnSchema.node('doc', null, [
      refPara('See', 1),
      ...Array.from({ length: 8 }, (_, i) => p(`p${i}`)),
    ]);
    const resolved = layout(body, cfg, undefined, undefined, {
      1: noteDoc('Note one'),
    });

    const fn = resolved.pages[0].footnotes;
    expect(fn).toBeDefined();
    // the note body is painted at the page bottom, below every body line…
    const noteLine = fn?.lines[0];
    const lastBody = resolved.pages[0].lines.at(-1);
    const noteText = (noteLine?.segments ?? []).map((s) => s.text).join('');
    expect(noteText).toContain('Note one');
    expect(noteLine?.y ?? 0).toBeGreaterThan(lastBody?.y ?? 0);
    // …and within the page (above the bottom margin).
    expect((noteLine?.y ?? 0) + (noteLine?.height ?? 0)).toBeLessThanOrEqual(
      280.01,
    );
    // the separator rule sits above the first note line.
    expect(fn?.separatorY ?? Infinity).toBeLessThan(noteLine?.y ?? 0);
    // footnote lines belong to a separate story — never caret-addressable.
    expect(noteLine?.segments[0].pos).toBeUndefined();
    expect(noteLine?.from).toBeUndefined();
  });

  it('shrinks the body band so fewer lines fit the reference page', () => {
    const cfg = config({ height: 300 });
    const body = fnSchema.node('doc', null, [
      refPara('See', 1),
      ...Array.from({ length: 40 }, (_, i) => p(`p${i}`)),
    ]);
    const without = layout(body, cfg);
    const withFn = layout(body, cfg, undefined, undefined, {
      1: noteDoc('Note one'),
    });
    expect(withFn.pages[0].footnotes).toBeDefined();
    expect(without.pages[0].footnotes).toBeUndefined();
    // the reserved footnote area pushes a body line off page 1.
    expect(withFn.pages[0].lines.length).toBeLessThan(
      without.pages[0].lines.length,
    );
  });

  it('leaves pages free of footnotes when none are passed', () => {
    const body = fnSchema.node('doc', null, [refPara('See', 1), p('tail')]);
    const resolved = layout(body, config({ height: 300 }));
    expect(resolved.pages.every((pg) => pg.footnotes === undefined)).toBe(true);
  });

  const tblSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
      table: {
        group: 'block',
        content: 'table_row+',
        attrs: {
          cellPadding: { default: null },
          borders: { default: null },
          align: { default: null },
        },
      },
      table_row: {
        content: 'table_cell+',
        attrs: { header: { default: false }, height: { default: null } },
      },
      table_cell: {
        content: 'block+',
        attrs: {
          colspan: { default: 1 },
          rowspan: { default: 1 },
          colwidth: { default: null },
          background: { default: null },
          vAlign: { default: null },
          borders: { default: null },
        },
      },
    },
    marks: {
      footnote: { attrs: { num: {} } },
      vertAlign: { attrs: { value: {} } },
    },
  });

  it('reserves a footnote referenced inside a table cell', () => {
    const cellPara = tblSchema.node('paragraph', null, [
      tblSchema.text('Cell '),
      tblSchema.text('1', [
        tblSchema.mark('footnote', { num: 1 }),
        tblSchema.mark('vertAlign', { value: 'super' }),
      ]),
    ]);
    const doc = tblSchema.node('doc', null, [
      tblSchema.node('table', null, [
        tblSchema.node('table_row', null, [
          tblSchema.node('table_cell', null, [cellPara]),
        ]),
      ]),
    ]);
    const note = tblSchema.node('doc', null, [
      tblSchema.node('paragraph', null, [tblSchema.text('Table note')]),
    ]);
    const r = layout(doc, config({ height: 300 }), undefined, undefined, {
      1: note,
    });
    const fn = r.pages[0].footnotes;
    expect(fn).toBeDefined();
    expect(
      (fn?.lines ?? [])
        .flatMap((l) => l.segments)
        .map((s) => s.text)
        .join(''),
    ).toContain('Table note');
  });
});

describe('multi-column layout', () => {
  // content width 200; 2 cols, gap 20 → colWidth 90; col0 x=20, col1 x=130.
  it('balances columns evenly on a single-page section', () => {
    const cfg = { ...config({ height: 200 }), columns: { count: 2, gap: 20 } };
    // band [20,180] = 160px → 10 lines fit a column, but 15 lines balance to
    // ~7/8 across the two columns instead of packing column 0 to 10.
    const blocks = Array.from({ length: 15 }, (_, i) => para(`p${i}`));
    const r = layoutBlocks(blocks, cfg);
    expect(r.pages).toHaveLength(1);
    const col0 = r.pages[0].lines.filter((l) => l.x === 20);
    const col1 = r.pages[0].lines.filter((l) => l.x === 130);
    expect(col0.length + col1.length).toBe(15);
    expect(Math.abs(col0.length - col1.length)).toBeLessThanOrEqual(1); // even
    expect(col0.length).toBeLessThan(10); // not greedily packed
    expect(col1[0].y).toBe(20); // column 1 restarts at the band top
  });

  it('fills full-height columns on non-final pages, balances the last', () => {
    const cfg = { ...config({ height: 200 }), columns: { count: 2, gap: 20 } };
    const blocks = Array.from({ length: 25 }, (_, i) => para(`p${i}`));
    const r = layoutBlocks(blocks, cfg);
    expect(r.pages).toHaveLength(2);
    // Page 1 is greedy: column 0 packs to the full 10 lines.
    expect(r.pages[0].lines.filter((l) => l.x === 20)).toHaveLength(10);
    expect(r.pages[0].lines).toHaveLength(20);
    // Page 2 (the final page) balances its 5 lines ~3/2 across the columns.
    const p2col0 = r.pages[1].lines.filter((l) => l.x === 20);
    const p2col1 = r.pages[1].lines.filter((l) => l.x === 130);
    expect(p2col0.length + p2col1.length).toBe(5);
    expect(Math.abs(p2col0.length - p2col1.length)).toBeLessThanOrEqual(1);
  });

  it('single column is unchanged (no x shift)', () => {
    const cfg = { ...config({ height: 200 }), columns: { count: 1, gap: 20 } };
    const r = layoutBlocks([para('a'), para('b')], cfg);
    expect(r.pages[0].lines.every((l) => l.x === 20)).toBe(true);
  });

  const secSchema = new Schema({
    nodes: {
      doc: { content: 'block+', attrs: { sections: { default: null } } },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const secDoc = (sections: unknown, n: number) =>
    secSchema.node(
      'doc',
      { sections },
      Array.from({ length: n }, (_, i) =>
        secSchema.node('paragraph', null, [secSchema.text(`p${i}`)]),
      ),
    );

  it('switches columns at a continuous section break', () => {
    // section A: 2 paras, 1 col; section B: 20 paras, 2 cols, continuous.
    const doc = secDoc(
      [
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: false },
        { blockCount: 20, columns: { count: 2, gap: 20 }, newPage: false },
      ],
      22,
    );
    const r = layout(doc, config({ height: 200 }));
    const all = r.pages.flatMap((p) => p.lines);
    expect(all.some((l) => l.x === 130)).toBe(true); // column 1 of section B
    expect(r.pages.length).toBeGreaterThan(1);
  });

  it('starts a new page at a next-page section break', () => {
    // Both sections are short enough to share a page, but the break forces two.
    const doc = secDoc(
      [
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: true },
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: true },
      ],
      4,
    );
    const r = layout(doc, config({ height: 1000 }));
    expect(r.pages).toHaveLength(2);
    expect(r.pages[1].lines[0].y).toBe(20); // section B at the new page top
  });

  it('lays a per-section page override on its own page geometry', () => {
    // Section B rotates to landscape (500×240) with wider margins.
    const landscape = {
      width: 500,
      height: 240,
      margin: { top: 30, right: 40, bottom: 30, left: 40 },
    };
    const doc = secDoc(
      [
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: true },
        {
          blockCount: 2,
          columns: { count: 1, gap: 0 },
          newPage: true,
          page: landscape,
        },
      ],
      4,
    );
    const r = layout(doc, config({ height: 1000 }));
    expect(r.pages).toHaveLength(2);
    // Page 0: document geometry; page 1: the override's.
    expect(r.pages[0]).toMatchObject({ width: 240, height: 1000 });
    expect(r.pages[1]).toMatchObject({ width: 500, height: 240 });
    // Section B content starts at ITS margins, not the document's.
    expect(r.pages[1].lines[0].y).toBe(30);
    expect(r.pages[1].lines[0].x).toBe(40);
  });

  it('promotes a continuous break to next-page when geometry differs', () => {
    const landscape = {
      width: 500,
      height: 240,
      margin: { top: 20, right: 20, bottom: 20, left: 20 },
    };
    const doc = secDoc(
      [
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: false },
        {
          blockCount: 2,
          columns: { count: 1, gap: 0 },
          newPage: false, // continuous — but geometry can't switch mid-page
          page: landscape,
        },
      ],
      4,
    );
    const r = layout(doc, config({ height: 1000 }));
    expect(r.pages).toHaveLength(2);
    expect(r.pages[1]).toMatchObject({ width: 500, height: 240 });
  });

  it('a first-section override applies from page one', () => {
    const landscape = {
      width: 500,
      height: 240,
      margin: { top: 20, right: 20, bottom: 20, left: 20 },
    };
    const doc = secDoc(
      [
        {
          blockCount: 2,
          columns: { count: 1, gap: 0 },
          newPage: true,
          page: landscape,
        },
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: true },
      ],
      4,
    );
    const r = layout(doc, config({ height: 1000 }));
    expect(r.pages).toHaveLength(2);
    expect(r.pages[0]).toMatchObject({ width: 500, height: 240 });
    expect(r.pages[1]).toMatchObject({ width: 240, height: 1000 }); // back to doc default
  });

  it('stamps pages with their section chrome set and section-first flag', () => {
    const chromeDoc = (t: string) =>
      secSchema.node('doc', null, [
        secSchema.node('paragraph', null, [secSchema.text(t)]),
      ]);
    const doc = secDoc(
      [
        { blockCount: 20, columns: { count: 1, gap: 0 }, newPage: true },
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: true },
      ],
      22,
    );
    const r = layout(doc, config({ height: 200 }), undefined, {
      sections: [
        { header: chromeDoc('CH-ONE'), titlePg: true },
        { header: chromeDoc('CH-TWO') },
      ],
    });
    expect(r.chromeSets).toHaveLength(2);
    expect(r.chromeSets?.[0].titlePg).toBe(true);
    expect(r.pages.length).toBeGreaterThan(2);
    // Section 1 spans several pages: page 0 is its section-first page, the
    // following section-1 pages are not; the last page belongs to section 2
    // and is ITS first page.
    expect(r.pages[0]).toMatchObject({ chromeIndex: 0, sectionFirst: true });
    expect(r.pages[1]).toMatchObject({ chromeIndex: 0, sectionFirst: false });
    const last = r.pages[r.pages.length - 1];
    expect(last).toMatchObject({ chromeIndex: 1, sectionFirst: true });
    // Each set's band is laid out (header content present).
    expect(r.chromeSets?.[0].header?.lines.length).toBeGreaterThan(0);
    expect(r.chromeSets?.[1].header?.lines.length).toBeGreaterThan(0);
  });

  it("a tall section header shrinks only that section's pages", () => {
    // Section 2's header is 5 lines tall; section 1 has none.
    const tallHeader = secSchema.node(
      'doc',
      null,
      Array.from({ length: 5 }, (_, i) =>
        secSchema.node('paragraph', null, [secSchema.text(`h${i}`)]),
      ),
    );
    const doc = secDoc(
      [
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: true },
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: true },
      ],
      4,
    );
    const r = layout(doc, config({ height: 1000 }), undefined, {
      sections: [{}, { header: tallHeader }],
    });
    expect(r.pages).toHaveLength(2);
    const firstLineY = (p: number) => r.pages[p].lines[0]?.y ?? -1;
    expect(firstLineY(0)).toBe(20); // section 1: plain margin top
    expect(firstLineY(1)).toBeGreaterThan(20); // section 2: pushed below its header
  });

  it('sanitizes a degenerate per-section override instead of hanging', () => {
    const doc = secDoc(
      [
        { blockCount: 2, columns: { count: 1, gap: 0 }, newPage: true },
        {
          blockCount: 2,
          columns: { count: 1, gap: 0 },
          newPage: true,
          page: {
            width: NaN,
            height: -5,
            margin: { top: NaN, right: 0, bottom: 0, left: 0 },
          },
        },
      ],
      4,
    );
    const r = layout(doc, config({ height: 1000 })); // must terminate
    expect(r.pages.length).toBeGreaterThanOrEqual(2);
    expect(Number.isFinite(r.pages[1].width)).toBe(true);
  });
});

describe('page cache (resume-from-edited-page)', () => {
  const pSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*' },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const para = (t: string) =>
    pSchema.node('paragraph', null, [pSchema.text(t)]);
  // Nine one-line paragraphs → three pages of three lines each (band 60px,
  // 16px lines). The SAME node instances are reused across docs, exactly as
  // ProseMirror's structural sharing keeps unchanged nodes identical.
  const nodes = 'abcdefghi'.split('').map((ch) => para(ch.repeat(3)));
  const mkDoc = (override?: { at: number; text: string }) =>
    pSchema.node(
      'doc',
      null,
      nodes.map((n, i) =>
        override && i === override.at ? para(override.text) : n,
      ),
    );
  const cfg = () => config({ height: 100 });

  it('reuses every unchanged page of an untouched document by identity', () => {
    const cache = createLayoutCache();
    const doc = mkDoc();
    const r1 = layout(doc, cfg(), cache);
    expect(r1.pages).toHaveLength(3);
    const r2 = layout(doc, cfg(), cache);
    expect(r2.pages).toHaveLength(3);
    // Pages before the boundary page come straight from the cache — same
    // objects. The final page is re-placed (boundary insurance) but equal.
    expect(r2.pages[0]).toBe(r1.pages[0]);
    expect(r2.pages[1]).toBe(r1.pages[1]);
    expect(r2.pages[2].lines.map((l) => l.y)).toEqual(
      r1.pages[2].lines.map((l) => l.y),
    );
  });

  it('an edit on the LAST page keeps the prefix and re-places only the tail', () => {
    const cache = createLayoutCache();
    const r1 = layout(mkDoc(), cfg(), cache);
    const r2 = layout(mkDoc({ at: 8, text: 'zzz' }), cfg(), cache);
    expect(r2.pages[0]).toBe(r1.pages[0]); // prefix page: cached object
    // The edited page is genuinely re-placed with the new text.
    const last = r2.pages[2].lines.map((l) => l.segments[0]?.text);
    expect(last).toContain('zzz');
  });

  it('an edit MID-document re-attaches the unchanged tail by identity', () => {
    const cache = createLayoutCache();
    const r1 = layout(mkDoc(), cfg(), cache);
    // Same-length edit on page 2: page 3's items and carry are unchanged,
    // so the tail splices back — the very object from the previous pass.
    const r2 = layout(mkDoc({ at: 4, text: 'yyy' }), cfg(), cache);
    expect(r2.pages).toHaveLength(3);
    expect(r2.pages[1].lines.map((l) => l.segments[0]?.text)).toContain('yyy');
    expect(r2.pages[2]).toBe(r1.pages[2]);
  });

  it('a size-changing edit splices the tail with shifted PM positions', () => {
    const cache = createLayoutCache();
    const r1 = layout(mkDoc(), cfg(), cache);
    // 'eee' → 'eeeee': same line count, document positions after it +2.
    const r2 = layout(mkDoc({ at: 4, text: 'eeeee' }), cfg(), cache);
    expect(r2.pages).toHaveLength(3);
    const oldLine = r1.pages[2].lines[0];
    const newLine = r2.pages[2].lines[0];
    expect(newLine.from).toBe((oldLine.from ?? 0) + 2);
    expect(newLine.segments[0]?.pos).toBe((oldLine.segments[0]?.pos ?? 0) + 2);
    // Geometry is untouched — only positions moved.
    expect(newLine.x).toBe(oldLine.x);
    expect(newLine.y).toBe(oldLine.y);
  });
});

describe('incremental table re-layout', () => {
  const tSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
      table: {
        group: 'block',
        content: 'table_row+',
        attrs: {
          cellPadding: { default: null },
          borders: { default: null },
          align: { default: null },
        },
      },
      table_row: {
        content: 'table_cell+',
        attrs: { header: { default: false }, height: { default: null } },
      },
      table_cell: {
        content: 'block+',
        attrs: {
          colspan: { default: 1 },
          rowspan: { default: 1 },
          colwidth: { default: null },
          background: { default: null },
          vAlign: { default: null },
          borders: { default: null },
        },
      },
    },
    marks: {},
  });
  const para = (t: string) =>
    tSchema.node('paragraph', null, [tSchema.text(t)]);
  const cell = (t: string) => tSchema.node('table_cell', null, [para(t)]);
  // ONE table node reused across layouts (PM keeps unchanged nodes identical).
  const tableNode = tSchema.node('table', null, [
    tSchema.node('table_row', null, [cell('ZZZ'), cell('ZZZ')]),
  ]);
  const firstCellFrom = (r: ReturnType<typeof layout>) => {
    for (const pg of r.pages)
      for (const t of pg.tables ?? []) {
        const l = t.cells[0]?.lines[0];
        if (l?.from != null) return l.from;
      }
    return -1;
  };

  it('reuses a cached table instead of re-measuring its cells', () => {
    let zzz = 0;
    const measure: MeasureText = (text) => {
      if (text.includes('ZZZ')) zzz++;
      return text.length * 10;
    };
    const cfg = { ...config(), measureText: measure };
    const cache = createLayoutCache();
    layout(tSchema.node('doc', null, [para('a'), tableNode]), cfg, cache);
    const first = zzz;
    expect(first).toBeGreaterThan(0);
    // Edit elsewhere: a new paragraph before the SAME table node.
    layout(
      tSchema.node('doc', null, [para('a'), para('b'), tableNode]),
      cfg,
      cache,
    );
    expect(zzz).toBe(first); // table cells were NOT re-measured
  });

  it("shifts a cached table's PM positions when it moves", () => {
    const cfg = config();
    const cache = createLayoutCache();
    const r1 = layout(
      tSchema.node('doc', null, [para('a'), tableNode]),
      cfg,
      cache,
    );
    const r2 = layout(
      tSchema.node('doc', null, [para('aa'), para('bb'), tableNode]),
      cfg,
      cache,
    );
    expect(firstCellFrom(r2)).toBeGreaterThan(firstCellFrom(r1)); // moved down → positions shifted
  });

  const targetLine = (r: ReturnType<typeof layout>, text: string) => {
    for (const pg of r.pages)
      for (const l of pg.lines)
        if (l.segments.some((s) => s.text === text)) return l;
    return undefined;
  };

  it('reuses a cached paragraph instead of re-measuring it', () => {
    let zzz = 0;
    const measure: MeasureText = (text) => {
      if (text.includes('ZZZ')) zzz++;
      return text.length * 10;
    };
    const cfg = { ...config(), measureText: measure };
    const cache = createLayoutCache();
    const target = para('ZZZ'); // ONE node reused across layouts
    layout(tSchema.node('doc', null, [para('a'), target]), cfg, cache);
    const first = zzz;
    expect(first).toBeGreaterThan(0);
    layout(
      tSchema.node('doc', null, [para('a'), para('b'), target]),
      cfg,
      cache,
    );
    expect(zzz).toBe(first); // cache hit → NOT re-measured
  });

  it("shifts a cached paragraph's PM positions to match a fresh layout", () => {
    const cache = createLayoutCache();
    const target = para('hello'); // reused → cache hit on the 2nd layout
    // Seed the cache with `target` at one position…
    layout(tSchema.node('doc', null, [para('a'), target]), config(), cache);
    // …then grow the preceding paragraph so `target` moves by +4 (cache hit +
    // in-place draft shift). Its positions must equal a from-scratch layout.
    const docFinal = tSchema.node('doc', null, [para('aaaaa'), target]);
    const cached = layout(docFinal, config(), cache);
    const fresh = layout(docFinal, config()); // no cache → oracle
    const cl = targetLine(cached, 'hello');
    const fl = targetLine(fresh, 'hello');
    expect(fl?.from).toBeDefined();
    expect(cl!.from).toBe(fl!.from);
    expect(cl!.to).toBe(fl!.to);
    expect(cl!.segments[0].pos).toBe(fl!.segments[0].pos);
  });

  it('lays a list-bearing table fresh (no stale numbering)', () => {
    let measures = 0;
    const measure: MeasureText = (t) => {
      measures++;
      return t.length * 10;
    };
    const cfg = { ...config(), measureText: measure };
    const cache = createLayoutCache();
    const listCell = tSchema.node('table_cell', null, [
      tSchema.node('paragraph', { list: { numId: '1', level: 0 } }, [
        tSchema.text('item'),
      ]),
    ]);
    const listTable = tSchema.node('table', null, [
      tSchema.node('table_row', null, [listCell]),
    ]);
    const doc = tSchema.node('doc', null, [listTable]);
    layout(doc, cfg, cache);
    const first = measures;
    layout(doc, cfg, cache); // same node, but list tables aren't cached
    expect(measures).toBeGreaterThan(first); // re-measured (not cached)
  });
});

describe('layout with comments', () => {
  const cSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
    },
    marks: { comment: { attrs: { ids: {} } } },
  });

  it('carries comment ids onto the commented segments only', () => {
    const doc = cSchema.node('doc', null, [
      cSchema.node('paragraph', null, [
        cSchema.text('plain '),
        cSchema.text('noted', [cSchema.mark('comment', { ids: [3, 7] })]),
      ]),
    ]);
    const segs = layout(doc, config()).pages[0].lines[0].segments;
    expect(segs.find((s) => s.text === 'noted')?.commentIds).toEqual([3, 7]);
    expect(segs.find((s) => s.text === 'plain')?.commentIds).toBeUndefined();
  });
});

describe('live list numbering', () => {
  const numbering = {
    '1': {
      key: 'a0',
      levels: { 0: { numFmt: 'decimal', lvlText: '%1.', start: 1 } },
    },
  };
  const listSchema = new Schema({
    nodes: {
      doc: { content: 'block+', attrs: { numbering: { default: null } } },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const li = (text: string) =>
    listSchema.node('paragraph', { list: { numId: '1', level: 0 } }, [
      listSchema.text(text),
    ]);
  const markersOf = (resolved: ReturnType<typeof layout>) =>
    resolved.pages[0].lines.map((l) => l.segments[0]?.text);

  it('counts markers at layout time and renumbers cached paragraphs', () => {
    const a = li('one');
    const b = li('two');
    const doc1 = listSchema.node('doc', { numbering }, [a, b]);
    const cache = createLayoutCache();
    expect(markersOf(layout(doc1, config(), cache))).toEqual(['1.', '2.']);

    // Insert a new first item; a and b keep node identity (cache hits) but
    // their markers shift — the cache must not serve the stale '1.'/'2.'.
    const doc2 = listSchema.node('doc', { numbering }, [li('zero'), a, b]);
    expect(markersOf(layout(doc2, config(), cache))).toEqual([
      '1.',
      '2.',
      '3.',
    ]);
  });
});

describe('layout with LayoutCache', () => {
  const schema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const p = (text: string) =>
    schema.node('paragraph', null, [schema.text(text)]);

  /** MeasureText that records every measured string. */
  const counting = () => {
    const calls: string[] = [];
    const fn: MeasureText = (text) => {
      calls.push(text);
      return text.length * 10;
    };
    return { fn, calls };
  };

  it('skips measuring entirely on an identical second layout', () => {
    const doc = schema.node('doc', null, [p('aa bb'), p('cc dd')]);
    const cache = createLayoutCache();
    const m = counting();
    const cfg = { ...config(), measureText: m.fn };
    const first = layout(doc, cfg, cache);
    expect(m.calls.length).toBeGreaterThan(0);
    m.calls.length = 0;
    const second = layout(doc, cfg, cache);
    expect(m.calls).toEqual([]); // every paragraph came from the cache
    expect(second).toEqual(first);
  });

  it('re-measures only the changed paragraph and shifts the rest', () => {
    const pA = p('aa');
    const pC = p('cc');
    const doc1 = schema.node('doc', null, [pA, p('bb'), pC]);
    const cache = createLayoutCache();
    const m = counting();
    const cfg = { ...config(), measureText: m.fn };
    layout(doc1, cfg, cache);

    // Replace the middle paragraph with a longer one; A and C keep identity.
    const doc2 = schema.node('doc', null, [pA, p('bbbb xx'), pC]);
    m.calls.length = 0;
    const result = layout(doc2, cfg, cache);
    expect(m.calls).toEqual(['bbbb', ' ', 'xx']); // only the new paragraph
    // C moved: 'bbbb xx' has nodeSize 9 → C's node sits at 13, content at 14.
    // The cached drafts must be position-shifted without re-measuring.
    const lineC = result.pages[0].lines[2];
    expect(lineC.from).toBe(14);
    expect(lineC.segments[0].pos).toBe(14);
  });

  it('invalidates on content-width change', () => {
    const doc = schema.node('doc', null, [p('aa')]);
    const cache = createLayoutCache();
    const m = counting();
    layout(doc, { ...config(), measureText: m.fn }, cache);
    m.calls.length = 0;
    layout(doc, { ...config({ width: 300 }), measureText: m.fn }, cache);
    expect(m.calls).toEqual(['aa']); // width changed → re-wrapped
  });
});

describe('mid-table tail resync', () => {
  // An edit ABOVE a long table used to re-place every page from the edit to
  // the table's end: mid-table boundaries never matched the previous run.
  // With the remainder cursor they can — same table node, same consume
  // cursor, same page scalars ⇒ the tail splices back.
  const sch = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*' },
      text: { group: 'inline' },
      table: { group: 'block', content: 'table_row+' },
      table_row: {
        content: 'table_cell+',
        attrs: { header: { default: false }, height: { default: null } },
      },
      table_cell: {
        content: 'block+',
        attrs: {
          colspan: { default: 1 },
          rowspan: { default: 1 },
          colwidth: { default: null },
          background: { default: null },
          vAlign: { default: null },
          borders: { default: null },
        },
      },
    },
    marks: {},
  });
  const para = (t: string) => sch.node('paragraph', null, [sch.text(t)]);
  // ONE 12-row table node shared across docs (PM structural sharing): three
  // pages of 60px content hold ~3 one-line 16px rows each, so the table
  // spans several page boundaries.
  const tableNode = sch.node(
    'table',
    null,
    Array.from({ length: 12 }, (_, r) =>
      sch.node('table_row', null, [
        sch.node('table_cell', null, [para(`r${r}`)]),
      ]),
    ),
  );
  const mkDoc = (lead: string) =>
    sch.node('doc', null, [para(lead), tableNode]);
  const cfg = () => config({ height: 100 });

  it('an edit above the table splices its unchanged pages by identity', () => {
    const cache = createLayoutCache();
    const r1 = layout(mkDoc('aaa'), cfg(), cache);
    expect(r1.pages.length).toBeGreaterThan(3); // sanity: table crosses pages
    // Same-length edit: the table's fragments land identically, so every
    // page after the edited one must come back as the SAME object.
    const r2 = layout(mkDoc('bbb'), cfg(), cache);
    expect(r2.pages.length).toBe(r1.pages.length);
    expect(r2.pages[0].lines[0].segments[0]?.text).toBe('bbb');
    for (let p = 1; p < r2.pages.length; p++) {
      expect(r2.pages[p]).toBe(r1.pages[p]);
    }
  });

  it('a size-changing edit splices with shifted PM positions', () => {
    const cache = createLayoutCache();
    const r1 = layout(mkDoc('aaa'), cfg(), cache);
    const r2 = layout(mkDoc('aaaaa'), cfg(), cache); // +2 positions
    expect(r2.pages.length).toBe(r1.pages.length);
    const oldCell = r1.pages[1].tables?.[0]?.cells[0];
    const newCell = r2.pages[1].tables?.[0]?.cells[0];
    expect(newCell?.lines[0]?.from).toBe((oldCell?.lines[0]?.from ?? 0) + 2);
    expect(newCell?.lines[0]?.y).toBe(oldCell?.lines[0]?.y);
    // A THIRD pass resyncs against the REBASED entries (posDelta path).
    const r3 = layout(mkDoc('aaaaaaa'), cfg(), cache); // +2 again
    const c3 = r3.pages[1].tables?.[0]?.cells[0];
    expect(c3?.lines[0]?.from).toBe((oldCell?.lines[0]?.from ?? 0) + 4);
    expect(c3?.lines[0]?.y).toBe(oldCell?.lines[0]?.y);
  });

  it('an edit INSIDE the table still re-places its pages', () => {
    // The table node itself changed → no resync candidate; correctness
    // baseline: output equals a cold layout.
    const cache = createLayoutCache();
    layout(mkDoc('aaa'), cfg(), cache);
    const edited = sch.node('doc', null, [
      para('aaa'),
      sch.node(
        'table',
        null,
        Array.from({ length: 12 }, (_, r) =>
          sch.node('table_row', null, [
            sch.node('table_cell', null, [para(r === 6 ? 'EDIT' : `r${r}`)]),
          ]),
        ),
      ),
    ]);
    const warm = layout(edited, cfg(), cache);
    const cold = layout(edited, cfg());
    expect(warm.pages.length).toBe(cold.pages.length);
    warm.pages.forEach((pg, i) => {
      const flat = (r: typeof cold) =>
        (r.pages[i].tables ?? []).flatMap((t) =>
          t.cells.flatMap((c) =>
            c.lines.map((l) => [l.y, l.segments[0]?.text]),
          ),
        );
      expect(flat(warm)).toEqual(flat(cold));
    });
  });
});

describe('table cells: paragraph rules and page fill', () => {
  // Cell content used to go through a reduced second layout engine: w:spacing
  // was swallowed, keeps were dropped, and the paginator preferred a row
  // boundary over filling the page. All three against Word ground truth
  // (nested_table.docx: Word starts row 2 on page 1 and breaks it mid-row).
  const cellOf = (content: FlowBlock[]): FlowTableCell => ({
    colspan: 1,
    rowspan: 1,
    colwidth: null,
    content,
  });
  const p = (text: string, over: Partial<FlowParagraph> = {}): FlowParagraph =>
    ({
      type: 'paragraph',
      runs: [{ text, font: font() }],
      ...over,
    }) as FlowParagraph;
  const lines1 = (n: number, pre: string) =>
    Array.from({ length: n }, (_, i) => p(`${pre}${i + 1}`));
  const rowsOf = (t: ResolvedTable | undefined) =>
    t?.cells.map((c) => c.lines.map((l) => l.segments[0]?.text));

  it('w:spacing before/after is honored inside a cell', () => {
    const mk = (spaced: boolean) =>
      layoutBlocks(
        [
          {
            type: 'table',
            rows: [
              {
                cells: [
                  cellOf([
                    p('a'),
                    p(
                      'b',
                      spaced ? { spacing: { before: 30, after: 30 } } : {},
                    ),
                  ]),
                ],
              },
            ],
          },
        ],
        config(),
      ).pages[0].tables?.[0];
    const plain = mk(false);
    const spaced = mk(true);
    expect(spaced?.height).toBeGreaterThan((plain?.height ?? 0) + 59);
    // The gap sits between the paragraphs, not after the first line.
    expect(spaced?.cells[0].lines[1].y).toBe(
      (plain?.cells[0].lines[1].y ?? 0) + 30,
    );
  });

  it('fills the page before a tall row instead of leaving a blank gap', () => {
    // Word ground truth (nested_table.docx page 1): a short first row, then
    // the tall second row STARTS in the leftover space and breaks mid-row.
    // Preferring the row boundary left ~90% of the page blank.
    const t: FlowBlock = {
      type: 'table',
      rows: [
        { cells: [cellOf([p('head')])] },
        { cells: [cellOf(lines1(6, 'l'))] },
      ],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(pages).toHaveLength(3);
    expect(rowsOf(pages[0].tables?.[0])).toEqual([['head'], ['l1', 'l2']]);
    expect(rowsOf(pages[1].tables?.[0])).toEqual([['l3', 'l4', 'l5']]);
    expect(rowsOf(pages[2].tables?.[0])).toEqual([['l6']]);
  });

  it('a keepNext heading in a cell moves with its body', () => {
    // budget after 'hi' = 44px = 2 lines + slack. The geometric cut falls
    // right after HEAD; keepNext forbids that boundary, so HEAD moves down
    // with its body instead of ending the fragment alone.
    const t: FlowBlock = {
      type: 'table',
      rows: [
        {
          cells: [
            cellOf([p('b1'), p('HEAD', { keepNext: true }), ...lines1(4, 'c')]),
          ],
        },
      ],
    };
    const { pages } = layoutBlocks([p('hi'), t], config({ height: 100 }));
    expect(rowsOf(pages[0].tables?.[0])).toEqual([['b1']]);
    expect(rowsOf(pages[1].tables?.[0])?.[0]?.slice(0, 2)).toEqual([
      'HEAD',
      'c1',
    ]);
  });

  it('widow/orphan control holds inside a split row', () => {
    // One 5-line paragraph; the geometric cut leaves 1 line above. Orphan
    // rule (≥2 each side) pushes the whole paragraph down.
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [cellOf([p('aaa bbb ccc ddd eee')])] }],
    };
    const { pages } = layoutBlocks(
      [p('hi'), p('hi2'), t],
      config({ height: 100 }),
    );
    const first = rowsOf(pages[0].tables?.[0]) ?? [[]];
    expect(first[0]).toEqual([]); // nothing stranded above
    expect((rowsOf(pages[1].tables?.[0])?.[0] ?? []).length).toBeGreaterThan(1);
  });

  it('widowControl false splits wherever the band ends', () => {
    const t: FlowBlock = {
      type: 'table',
      rows: [
        {
          cells: [cellOf([p('aaa bbb ccc ddd eee', { widowControl: false })])],
        },
      ],
    };
    const { pages } = layoutBlocks(
      [p('hi'), p('hi2'), t],
      config({ height: 100 }),
    );
    expect(rowsOf(pages[0].tables?.[0])?.[0]).toEqual(['aaa']); // free split
  });

  it('keepLines moves the whole paragraph to the next fragment', () => {
    // A narrow column forces the keepLines paragraph to wrap to 4 lines.
    const t: FlowBlock = {
      type: 'table',
      rows: [
        {
          cells: [
            {
              ...cellOf([
                p('b1'),
                p('aaaaaa bbbbbb cccccc dddddd', { keepLines: true }),
              ]),
              colwidth: [80],
            },
          ],
        },
      ],
    };
    const { pages } = layoutBlocks([p('hi'), t], config({ height: 100 }));
    expect(rowsOf(pages[0].tables?.[0])).toEqual([['b1']]);
    expect((rowsOf(pages[1].tables?.[0])?.[0] ?? []).length).toBe(4);
  });
});

describe('nested table pagination', () => {
  // A nested table taller than a full band used to move whole forever and
  // paint straight off the bottom of the sheet (the painter does not clip to
  // the page). It now splits recursively; a page-fitting nested table keeps
  // the whole-move behavior.
  const cellOf = (content: FlowBlock[]): FlowTableCell => ({
    colspan: 1,
    rowspan: 1,
    colwidth: null,
    content,
  });
  const p = (text: string): FlowParagraph => ({
    type: 'paragraph',
    runs: [{ text, font: font() }],
  });
  const nestedOf = (rows: number, pre: string): FlowBlock => ({
    type: 'table',
    rows: Array.from({ length: rows }, (_, r) => ({
      cells: [cellOf([p(`${pre}${r + 1}`)])],
    })),
  });
  const nestedRows = (pg: { tables?: ResolvedTable[] }) =>
    (pg.tables ?? []).flatMap((t) =>
      t.cells.flatMap((c) =>
        (c.tables ?? []).map((nt) =>
          nt.cells.map((nc) => nc.lines[0]?.segments[0]?.text),
        ),
      ),
    );
  const CONTENT_BOTTOM = 80; // height 100, margins 20 — nothing may pass this

  const worstBottom = (pages: { tables?: ResolvedTable[] }[]) =>
    Math.max(
      ...pages.flatMap((pg) => (pg.tables ?? []).map((t) => t.y + t.height)),
    );

  it('a nested table taller than the page splits instead of overflowing', () => {
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [cellOf([nestedOf(8, 'n')])] }],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(pages.length).toBeGreaterThan(1);
    expect(worstBottom(pages)).toBeLessThanOrEqual(CONTENT_BOTTOM);
    // Every nested row survives, in order, across the fragments.
    const all = pages.flatMap((pg) => nestedRows(pg).flat());
    expect(all).toEqual(['n1', 'n2', 'n3', 'n4', 'n5', 'n6', 'n7', 'n8']);
  });

  it('content after the split nested table re-stacks, not translates', () => {
    // [tall nested][paragraph after] — the paragraph must land right under
    // the nested remainder on the last fragment, not at its stale offset.
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [cellOf([nestedOf(8, 'n'), p('after')])] }],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(worstBottom(pages)).toBeLessThanOrEqual(CONTENT_BOTTOM);
    const last = pages[pages.length - 1].tables?.[0]?.cells[0];
    const afterLine = last?.lines.find((l) => l.segments[0]?.text === 'after');
    const lastNested = last?.tables?.[last.tables.length - 1];
    expect(afterLine).toBeDefined();
    expect(lastNested).toBeDefined();
    // Right below the nested remainder (cell padding aside), no stale gap.
    expect(afterLine?.y ?? 0).toBeLessThan(
      (lastNested?.y ?? 0) + (lastNested?.height ?? 0) + 8,
    );
  });

  it('a page-fitting nested table splits when the page edge falls inside it', () => {
    // Word-verified (fixture fx5): a nested table with room to move whole
    // to the next page STILL splits at the page edge. This test used to
    // assert the opposite (whole-move) — that assumption predated the
    // Word ground truth and was wrong.
    // Geometry: content 60px = 3 lines. Cell = [1 line][3-row nested][tail]:
    // the page edge falls after the nested table's second row.
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [cellOf([p('a1'), nestedOf(3, 'n'), p('tail')])] }],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    expect(nestedRows(pages[0])).toEqual([['n1', 'n2']]); // split mid-nested
    expect(nestedRows(pages[1])).toEqual([['n3']]);
    // The tail re-stacks under the nested remainder on page 2.
    const last = pages[1].tables?.[0]?.cells[0];
    expect(last?.lines.map((l) => l.segments[0]?.text)).toEqual(['tail']);
  });

  it('nested-in-nested taller than the page splits at every level', () => {
    const inner = nestedOf(8, 'i');
    const middle: FlowBlock = {
      type: 'table',
      rows: [{ cells: [cellOf([inner])] }],
    };
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [cellOf([middle])] }],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(worstBottom(pages)).toBeLessThanOrEqual(CONTENT_BOTTOM);
    // All inner rows survive across pages.
    const texts: string[] = [];
    const walk = (t2: ResolvedTable) => {
      for (const c of t2.cells) {
        for (const l of c.lines)
          if (l.segments[0]?.text) texts.push(l.segments[0].text);
        for (const nt of c.tables ?? []) walk(nt);
      }
    };
    pages.forEach((pg) => (pg.tables ?? []).forEach(walk));
    expect(texts).toEqual(['i1', 'i2', 'i3', 'i4', 'i5', 'i6', 'i7', 'i8']);
  });
});

describe('row-start veto (keepNext opening the first cell)', () => {
  // Word-verified (fixture fx2 + nested_table.docx row 3): a row whose FIRST
  // cell opens with a keepNext paragraph must not START in a band's
  // leftover — it begins on a fresh band, and only then splits normally.
  // A keepNext opener in a LATER cell does not veto (nested_table row 2).
  const cellOf = (content: FlowBlock[]): FlowTableCell => ({
    colspan: 1,
    rowspan: 1,
    colwidth: null,
    content,
  });
  const p = (text: string, over: Partial<FlowParagraph> = {}): FlowParagraph =>
    ({
      type: 'paragraph',
      runs: [{ text, font: font() }],
      ...over,
    }) as FlowParagraph;
  const lines1 = (n: number, pre: string, start = 1) =>
    Array.from({ length: n }, (_, i) => p(`${pre}${start + i}`));
  const rowsOf = (t: ResolvedTable | undefined) =>
    t?.cells.map((c) => c.lines.map((l) => l.segments[0]?.text));
  // content 60px = 3 lines of 16px per page (config height 100).

  it('pushes the row whole, then splits it normally (fx2 shape)', () => {
    const t: FlowBlock = {
      type: 'table',
      rows: [
        { cells: [cellOf([p('h1')])] },
        {
          cells: [cellOf([p('K1', { keepNext: true }), ...lines1(5, 'K', 2)])],
        },
      ],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    // page 1: header row ONLY — the 6-line row must not start in the
    // 2-line leftover, even though it cannot fit any single page whole.
    expect(rowsOf(pages[0].tables?.[0])).toEqual([['h1']]);
    // page 2: the row starts at the top of the fresh band and splits.
    expect(rowsOf(pages[1].tables?.[0])).toEqual([['K1', 'K2', 'K3']]);
    expect(rowsOf(pages[2].tables?.[0])).toEqual([['K4', 'K5', 'K6']]);
  });

  it('vetoes the whole table when its first row is keep-start mid-band', () => {
    const t: FlowBlock = {
      type: 'table',
      rows: [
        {
          cells: [cellOf([p('K1', { keepNext: true }), ...lines1(4, 'K', 2)])],
        },
      ],
    };
    const { pages } = layoutBlocks([p('intro'), t], config({ height: 100 }));
    expect(pages[0].tables ?? []).toHaveLength(0); // page 1: paragraph only
    expect(rowsOf(pages[1].tables?.[0])).toEqual([['K1', 'K2', 'K3']]);
  });

  it('a keep-start row that FITS the leftover stays (no gratuitous push)', () => {
    const t: FlowBlock = {
      type: 'table',
      rows: [{ cells: [cellOf([p('K1', { keepNext: true }), p('K2')])] }],
    };
    const { pages } = layoutBlocks([p('intro'), t], config({ height: 100 }));
    expect(pages).toHaveLength(1); // 1 + 2 lines = 3 → fits, no split → no veto
    expect(rowsOf(pages[0].tables?.[0])).toEqual([['K1', 'K2']]);
  });

  it('keepNext opening a LATER cell does not veto (nested_table row 2)', () => {
    const t: FlowBlock = {
      type: 'table',
      rows: [
        { cells: [cellOf([p('h1')]), cellOf([p('h2')])] },
        {
          cells: [
            cellOf(lines1(5, 'A')),
            cellOf([p('K1', { keepNext: true }), ...lines1(4, 'K', 2)]),
          ],
        },
      ],
    };
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    // The tall row STARTS in the leftover: page 1 carries its first lines.
    expect(rowsOf(pages[0].tables?.[0])).toEqual([
      ['h1'],
      ['h2'],
      ['A1', 'A2'],
      ['K1', 'K2'],
    ]);
  });
});
