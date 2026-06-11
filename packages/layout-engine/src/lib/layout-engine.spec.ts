import { Schema } from 'prosemirror-model';
import type {
  FlowBlock,
  FlowTableCell,
  FontSpec,
  LayoutConfig,
  MeasureText,
} from '@shadow-garden/bapbong-contracts';
import { createLayoutCache, layout, layoutBlocks, toFlowBlocks } from './layout-engine.js';

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

  it('carries underline/strike flags and measured width onto segments', () => {
    const block: FlowBlock = {
      type: 'paragraph',
      runs: [{ text: 'ab cd', font: font(), underline: true, strike: true }],
    };
    const { pages } = layoutBlocks([block], config());
    const [s0] = pages[0].lines[0].segments;
    expect(s0).toMatchObject({ text: 'ab', underline: true, strike: true, width: 20 });
  });

  it('pads cell content by the Word default cell margin', () => {
    const t = table([[cell('a', { colwidth: [100] })]]);
    const { pages } = layoutBlocks([t], config());
    const line = pages[0].tables?.[0]?.cells[0].lines[0];
    expect(line?.segments[0].x).toBeCloseTo(20 + 7.2); // cell.x + 108 twips
    expect(line?.width).toBeCloseTo(100 - 2 * 7.2);
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
    expect(line?.segments[0].x).toBeCloseTo(20 + 20); // custom left
    expect(line?.width).toBeCloseTo(100 - 20 - 4); // custom left + right
    expect(line?.y).toBeCloseTo(20 + 6); // custom top
    expect(resolved?.cells[0].height).toBeCloseTo(16 + 6); // top + content (+ bottom default 0)
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
    expect(resolved.cells[0]).toMatchObject({ x: 20, width: 200, colspan: 2 });
  });

  it('splits an overflowing table at the row boundary (Word-like)', () => {
    // page height 80 → content 20..60 (40px). One line (16) + table.
    const cfg = config({ height: 80 });
    const para: FlowBlock = { type: 'paragraph', runs: [{ text: 'hi', font: font() }] };
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
    const t = table([[cell('r1')], [cell('r2')], [cell('r3')], [cell('r4')], [cell('r5')]]);
    const { pages } = layoutBlocks([t], config({ height: 100 }));
    expect(pages).toHaveLength(2);
    const f1 = pages[0].tables?.[0];
    expect(f1?.cells.map((c) => c.lines[0].segments[0].text)).toEqual(['r1', 'r2', 'r3']);
    expect(f1?.height).toBeCloseTo(48);
    const f2 = pages[1].tables?.[0];
    expect(f2?.cells.map((c) => c.lines[0].segments[0].text)).toEqual(['r4', 'r5']);
    expect(f2?.cells[0]).toMatchObject({ y: 20 });
  });

  it('repeats header rows on every continuation fragment', () => {
    // content height 60; header (16) + 4 body rows (64) = 80px total.
    // Cells carry PM positions so the ghost-vs-original distinction is real.
    const pcell = (text: string, pos: number): FlowTableCell => ({
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      content: [{ type: 'paragraph', runs: [{ text, font: font(), pos }], pos, end: pos + text.length }],
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
    expect(f1?.cells.map((c) => c.lines[0].segments[0].text)).toEqual(['head', 'r1', 'r2']);
    // page 2: GHOST header (positions stripped) + r3 + r4.
    const f2 = pages[1].tables?.[0];
    expect(f2?.cells.map((c) => c.lines[0].segments[0].text)).toEqual(['head', 'r3', 'r4']);
    const ghost = f2?.cells[0].lines[0];
    expect(ghost?.segments[0].pos).toBeUndefined(); // not caret-addressable
    expect(ghost?.from).toBeUndefined();
    expect(f2?.cells[0]).toMatchObject({ y: 20 });
    expect(f2?.cells[1]).toMatchObject({ y: 36 }); // body resumes under the ghost
    // the original header (page 1) keeps its positions
    expect(f1?.cells[0].lines[0].segments[0].pos).toBeDefined();
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
    expect(f1?.cells[0].lines.map((l) => l.segments[0].text)).toEqual(['l1', 'l2', 'l3']);
    // page 2: remaining lines re-stack from the page top; the next row follows.
    const f2 = pages[1].tables?.[0];
    expect(f2?.cells[0].lines.map((l) => l.segments[0].text)).toEqual(['l4', 'l5']);
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

describe('layout with page chrome (header/footer)', () => {
  const schema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*', attrs: { list: { default: null } } },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const p = (text: string) => schema.node('paragraph', null, [schema.text(text)]);
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
        paragraph: { group: 'block', content: 'inline*', attrs: { list: { default: null } } },
        text: { group: 'inline' },
        page_field: { inline: true, group: 'inline', atom: true, attrs: { kind: {} } },
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
      Array.from({ length: 30 }, (_, i) => fieldSchema.node('paragraph', null, [fieldSchema.text(`p${i}`)])),
    );
    const resolved = layout(body, { ...config({ height: 300 }) }, undefined, { footer });
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
    const body = schema.node('doc', null, Array.from({ length: 12 }, (_, i) => p(`p${i}`)));
    const resolved = layout(body, cfg, undefined, {
      header: docOf('head'),
      footer: docOf('foot'),
    });
    expect(resolved.pages).toHaveLength(2);
    expect(resolved.pages[0].lines.length).toBe(10);
    const last = resolved.pages[0].lines.at(-1);
    expect((last?.y ?? 0) + (last?.height ?? 0)).toBeLessThanOrEqual(236.01);
  });
});

describe('layout with LayoutCache', () => {
  const schema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*', attrs: { list: { default: null } } },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const p = (text: string) => schema.node('paragraph', null, [schema.text(text)]);

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
