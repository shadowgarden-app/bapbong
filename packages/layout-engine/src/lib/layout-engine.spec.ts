import { Schema } from 'prosemirror-model';
import type {
  FlowBlock,
  FlowTableCell,
  FontSpec,
  LayoutConfig,
  MeasureText,
} from '@shadow-garden/bapbong-contracts';
import { layoutBlocks, toFlowBlocks } from './layout-engine.js';

// 10px per character, font-agnostic — keeps wrapping math predictable.
const measure: MeasureText = (text) => text.length * 10;

const font = (over: Partial<FontSpec> = {}): FontSpec => ({
  family: 'Arial',
  sizePt: 10, // → 10 * 96/72 * 1.2 = 16px line height
  bold: false,
  italic: false,
  ...over,
});

const para = (text: string, marker?: string): FlowBlock => ({
  type: 'paragraph',
  runs: [{ text, font: font() }],
  marker,
});

const config = (over: Partial<LayoutConfig['page']> = {}): LayoutConfig => ({
  measureText: measure,
  defaultFont: { sizePt: 10 },
  page: { width: 240, height: 1000, margin: { top: 20, right: 20, bottom: 20, left: 20 }, ...over },
});

describe('layoutBlocks', () => {
  it('wraps a paragraph at the content width', () => {
    // content width = 240 - 20 - 20 = 200px → 5 words of "aaaa"(40) + spaces.
    const { pages } = layoutBlocks([para('aaaa bbbb cccc dddd eeee')], config());
    expect(pages).toHaveLength(1);
    expect(pages[0].lines).toHaveLength(2);
    expect(pages[0].lines[0].segments[0].text).toBe('aaaa');
    expect(pages[0].lines[0].y).toBe(20);
    expect(pages[0].lines[1].segments[0].text).toBe('eeee');
    expect(pages[0].lines[1].y).toBe(36); // 20 + 16
  });

  it('paginates when content exceeds the page height', () => {
    // page height 80, margins 20 → content 20..60 = 40px → 2 lines/page.
    const cfg = config({ height: 80 });
    const { pages } = layoutBlocks([para('one'), para('two'), para('three')], cfg);
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
    expect(line.images?.[0]).toMatchObject({ src: 'data:img', x: 60, width: 50, height: 30 });
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
  const cell = (text: string, over: Partial<FlowTableCell> = {}): FlowTableCell => ({
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
    const t = table([
      [cell('a', { colwidth: [80] }), cell('b', { colwidth: [120] })],
      [cell('cc'), cell('dd')],
    ]);
    const { pages } = layoutBlocks([t], config());
    expect(pages).toHaveLength(1);
    const [resolved] = pages[0].tables ?? [];
    expect(resolved).toMatchObject({ x: 20, y: 20, width: 200 });
    expect(resolved.height).toBeCloseTo(32); // 2 rows × 16px (float)
    expect(resolved.cells).toHaveLength(4);
    // top-left cell sits at the table origin; its text starts there too.
    expect(resolved.cells[0]).toMatchObject({ x: 20, y: 20, width: 80 });
    expect(resolved.cells[0].height).toBeCloseTo(16);
    expect(resolved.cells[0].lines[0]).toMatchObject({ y: 20 });
    expect(resolved.cells[0].lines[0].segments[0].text).toBe('a');
    // second column starts at 20 + 80 = 100; second row at 20 + 16 = 36.
    expect(resolved.cells[1]).toMatchObject({ x: 100, y: 20, width: 120 });
    expect(resolved.cells[2]).toMatchObject({ x: 20, y: 36, width: 80 });
  });

  it('spans columns via colspan', () => {
    const t = table([
      [cell('wide', { colspan: 2, colwidth: [80, 120] })],
      [cell('x', { colwidth: [80] }), cell('y', { colwidth: [120] })],
    ]);
    const { pages } = layoutBlocks([t], config());
    const [resolved] = pages[0].tables ?? [];
    expect(resolved.cells[0]).toMatchObject({ x: 20, width: 200, colspan: 2 });
  });

  it('flows a paragraph, then a table, then paginates a table that overflows', () => {
    // page height 80 → content 20..60 (40px). One line (16) + table.
    const cfg = config({ height: 80 });
    const para: FlowBlock = { type: 'paragraph', runs: [{ text: 'hi', font: font() }] };
    const t = table([[cell('a')], [cell('b')]]); // 2 rows × 16 = 32px tall
    const { pages } = layoutBlocks([para, t], cfg);
    // line at y=20 (16px) → next free y=36; table (32) would end at 68 > 60 → new page.
    expect(pages).toHaveLength(2);
    expect(pages[0].lines[0].segments[0].text).toBe('hi');
    expect(pages[0].tables ?? []).toHaveLength(0);
    expect(pages[1].tables?.[0]).toMatchObject({ y: 20 });
    expect(pages[1].tables?.[0]?.height).toBeCloseTo(32);
  });

  it('uses injected font metrics for the line box and baseline', () => {
    const cfg: LayoutConfig = { ...config(), measureMetrics: () => ({ ascent: 12, descent: 4 }) };
    const { pages } = layoutBlocks([para('hi')], cfg);
    const [line] = pages[0].lines;
    expect(line.height).toBe(16); // ascent + descent
    expect(line.baseline).toBe(12); // baseline sits at the ascent
  });

  it('advances a tab to the next tab stop', () => {
    const block: FlowBlock = { type: 'paragraph', runs: [{ text: 'a\tb', font: font() }] };
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
      paragraph: { group: 'block', content: 'inline*', attrs: { list: { default: null } } },
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
        attrs: { colspan: { default: 1 }, rowspan: { default: 1 }, colwidth: { default: null } },
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
    // both are text runs; narrow off the InlineRun | InlineImage union.
    expect('src' in r0).toBe(false);
    if (!('src' in r0)) expect(r0.font.bold).toBe(true);
    if (!('src' in r1)) expect(r1.font.sizePt).toBe(20);
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
    if ('src' in img) expect(img).toMatchObject({ src: 'u', width: 40, height: 20 });
  });

  it('carries the list marker', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { list: { marker: '1.' } }, [schema.text('item')]),
    ]);
    const block = toFlowBlocks(doc)[0];
    expect(block.type === 'paragraph' && block.marker).toBe('1.');
  });

  it('flattens a table into rows and cells', () => {
    const p = (t: string) => schema.node('paragraph', null, [schema.text(t)]);
    const td = (t: string, attrs: Record<string, unknown> = {}) =>
      schema.node('table_cell', attrs, [p(t)]);
    const doc = schema.node('doc', null, [
      schema.node('table', null, [
        schema.node('table_row', null, [td('a', { colwidth: [80] }), td('b', { colwidth: [120] })]),
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
