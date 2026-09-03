import { describe, expect, it } from 'vitest';
import { schema } from '@shadow-garden/bapbong-model';
import type {
  FlowParagraph,
  FlowTable,
  LayoutConfig,
  MeasureText,
  ResolvedTableStyle,
  TableLook,
  TableStyleSheet,
} from '@shadow-garden/bapbong-contracts';
import { layout, toFlowBlocks } from './layout-engine.js';

/**
 * The live half of table theming: a document whose tables carry styleId/look
 * and whose doc.attrs hold the resolved sheet must come out of flattening
 * with the style APPLIED — fonts in the run base (so measuring sees them),
 * fills/borders/paddings on the cells, always UNDER the cell's own attrs.
 * The sheet values here mirror the probe-F2 style the importer-side specs
 * pin to Word's PDF.
 */

const measure: MeasureText = (text) => text.length * 10;
const green = { width: 1, style: 'solid' as const, color: '#9BBB59' };
const heavy = { ...green, width: 3 };

const STYLE: ResolvedTableStyle = {
  table: {
    borders: { top: green, insideH: green },
    cellPadding: { left: 7, right: 7 },
  },
  font: { family: 'Cambria' },
  paragraph: { spacing: { after: 0, line: 1, lineRule: 'auto' } },
  cond: {
    firstRow: {
      font: { bold: true, color: '#FFFFFF' },
      background: '#1F4E79',
      borders: { bottom: heavy },
    },
    firstCol: { font: { bold: true } },
    band1Horz: { background: '#D6E3BC' },
    lastRow: { font: { italic: true }, paragraph: { spacing: { before: 8 } } },
  },
  bands: { row: 1, col: 1 },
};
/** The floor: what w:docDefaults/w:pPrDefault would give every paragraph. */
const FLOOR = { spacing: { after: 13, line: 1.15, lineRule: 'auto' as const } };
const SHEET: TableStyleSheet = { Probe: STYLE };
const LOOK: TableLook = {
  firstRow: true,
  lastRow: false,
  firstCol: true,
  lastCol: false,
  hBand: true,
  vBand: false,
};

const cellNode = (text: string, attrs: Record<string, unknown> = {}) =>
  schema.nodes['table_cell'].create(attrs, [
    schema.nodes['paragraph'].create(null, text ? [schema.text(text)] : []),
  ]);
const rowNode = (texts: string[]) =>
  schema.nodes['table_row'].create(
    null,
    texts.map((t) => cellNode(t)),
  );

function docWith(
  tableAttrs: Record<string, unknown>,
  rows?: unknown[],
  docAttrs: Record<string, unknown> = {},
) {
  const table = schema.nodes['table'].create(
    tableAttrs,
    (rows as ReturnType<typeof rowNode>[]) ?? [
      rowNode(['h1', 'h2', 'h3']),
      rowNode(['a1', 'a2', 'a3']),
      rowNode(['b1', 'b2', 'b3']),
    ],
  );
  return schema.nodes['doc'].create({ tableStyles: SHEET, ...docAttrs }, [
    table,
  ]);
}
const cellPara = (flow: FlowTable, row: number, col: number) =>
  flow.rows[row].cells[col].content[0] as FlowParagraph;

const flowOf = (doc: ReturnType<typeof docWith>) =>
  toFlowBlocks(doc)[0] as FlowTable;
const runFont = (cellPara: unknown) =>
  (
    (cellPara as FlowParagraph).runs[0] as {
      font: { bold?: boolean; italic?: boolean; family?: string };
    }
  ).font;

describe('tableToFlow applies the live style sheet', () => {
  it('fonts deliberately do NOT flow from the layer', () => {
    // Run fonts are still baked into marks at import, and the mark model has
    // no explicit-off — a live bold under a baked strong mark cannot be
    // toggled away (⌘B in a styled header got stuck). So the layer's font
    // half stays parked until marks learn negation; fonts travel through
    // marks, rewritten by applyTableStyle on a style change.
    const flow = flowOf(docWith({ styleId: 'Probe', look: LOOK }));
    const header = flow.rows[0].cells[1].content[0] as FlowParagraph;
    expect(runFont(header).bold).toBe(false);
    expect(runFont(header).family).toBe('Arial');
    expect(header.runs[0]).toMatchObject({ color: undefined });
  });

  it('fills, borders and table-level fallbacks follow the gates', () => {
    const flow = flowOf(docWith({ styleId: 'Probe', look: LOOK }));
    expect(flow.rows[0].cells[2].background).toBe('#1F4E79');
    expect(flow.rows[0].cells[2].borders).toMatchObject({ bottom: heavy });
    // Banding counts the body: row 1 is band1 (filled), row 2 is band2.
    expect(flow.rows[1].cells[1].background).toBe('#D6E3BC');
    expect(flow.rows[2].cells[1].background).toBeUndefined();
    expect(flow.cellPadding).toEqual({ left: 7, right: 7 });
    expect(flow.borders).toMatchObject({ top: green });
  });

  it('the cell’s own attrs ride over the layer, per property', () => {
    const rows = [
      schema.nodes['table_row'].create(null, [
        cellNode('h1', { background: '#FF0000', borders: { bottom: green } }),
        cellNode('h2'),
      ]),
      rowNode(['a1', 'a2']),
    ];
    const flow = flowOf(docWith({ styleId: 'Probe', look: LOOK }, rows));
    // Direct fill and bottom border win; the style keeps the untouched side.
    expect(flow.rows[0].cells[0].background).toBe('#FF0000');
    expect(flow.rows[0].cells[0].borders).toMatchObject({ bottom: green });
    expect(flow.rows[0].cells[1].background).toBe('#1F4E79');
  });

  it('no styleId, unknown styleId, or a missing sheet change nothing', () => {
    const plain = flowOf(docWith({}));
    expect(plain.rows[0].cells[0].background).toBeUndefined();
    const unknown = flowOf(docWith({ styleId: 'Nope', look: LOOK }));
    expect(unknown.rows[0].cells[0].background).toBeUndefined();
    const table = schema.nodes['table'].create({ styleId: 'Probe' }, [
      rowNode(['x']),
    ]);
    const noSheet = toFlowBlocks(
      schema.nodes['doc'].create(null, [table]),
    )[0] as FlowTable;
    expect(noSheet.rows[0].cells[0].background).toBeUndefined();
  });

  it('a null look reads as Word’s 04A0 default', () => {
    const flow = flowOf(docWith({ styleId: 'Probe' }));
    expect(flow.rows[0].cells[1].background).toBe('#1F4E79');
    expect(flow.rows[1].cells[1].background).toBe('#D6E3BC');
  });

  it('the style reaches the painted cells through a full layout()', () => {
    const cfg: LayoutConfig = {
      measureText: measure,
      page: {
        width: 600,
        height: 400,
        margin: { top: 20, right: 20, bottom: 20, left: 20 },
      },
    };
    const resolved = layout(docWith({ styleId: 'Probe', look: LOOK }), cfg);
    const table = (resolved.pages[0].tables ?? [])[0];
    expect(table).toBeTruthy();
    expect(table.cells[0].background).toBe('#1F4E79');
  });
});

describe('paragraph spacing stacks floor → cell layer → own attr', () => {
  const styled = (
    rows?: unknown[],
    look: TableLook | null = LOOK,
    docAttrs: Record<string, unknown> = { paragraphDefaults: FLOOR },
  ) => flowOf(docWith({ styleId: 'Probe', look }, rows, docAttrs));

  it('a cell paragraph takes the style’s spacing over the floor', () => {
    const flow = styled();
    // Style: after 0, single — beats the floor's after 13 / 1.15 per field.
    expect(cellPara(flow, 1, 1).spacing).toEqual({
      after: 0,
      line: 1,
      lineRule: 'auto',
    });
  });

  it('the lastRow branch reaches the last row only through its gate', () => {
    const off = styled();
    expect(cellPara(off, 2, 1).spacing?.before).toBeUndefined();
    const on = styled(undefined, { ...LOOK, lastRow: true });
    expect(cellPara(on, 2, 1).spacing).toEqual({
      before: 8,
      after: 0,
      line: 1,
      lineRule: 'auto',
    });
    // …and the branch is a delta: the style's own fields still stand.
    expect(cellPara(on, 1, 1).spacing?.before).toBeUndefined();
  });

  it('the paragraph’s own attr wins per field', () => {
    const rows = [
      schema.nodes['table_row'].create(null, [
        schema.nodes['table_cell'].create(null, [
          schema.nodes['paragraph'].create({ spacing: { after: 4 } }, [
            schema.text('x'),
          ]),
        ]),
      ]),
    ];
    const flow = styled(rows);
    expect(cellPara(flow, 0, 0).spacing).toEqual({
      after: 4,
      line: 1,
      lineRule: 'auto',
    });
  });

  it('an unstyled table hands its cells the floor, not a neighbour’s layer', () => {
    const flow = flowOf(docWith({}, undefined, { paragraphDefaults: FLOOR }));
    expect(cellPara(flow, 0, 0).spacing).toEqual(FLOOR.spacing);
    // Without a floor either, a plain paragraph has no spacing at all.
    expect(cellPara(flowOf(docWith({})), 0, 0).spacing).toBeUndefined();
  });

  it('body paragraphs stand on the floor too, in flattening and in layout()', () => {
    const para = (text: string, attrs: Record<string, unknown> | null = null) =>
      schema.nodes['paragraph'].create(attrs, [schema.text(text)]);
    const doc = (floor: boolean) =>
      schema.nodes['doc'].create(
        floor ? { paragraphDefaults: { spacing: { after: 20 } } } : null,
        [para('one'), para('two')],
      );
    const flow = toFlowBlocks(doc(true))[0] as FlowParagraph;
    expect(flow.spacing).toEqual({ after: 20 });
    const cfg: LayoutConfig = {
      measureText: measure,
      page: {
        width: 600,
        height: 400,
        margin: { top: 20, right: 20, bottom: 20, left: 20 },
      },
    };
    const gap = (d: ReturnType<typeof doc>) => {
      const lines = layout(d, cfg).pages[0].lines;
      return lines[1].y - lines[0].y;
    };
    // The placer's before/after come from the same stack: 20px more between
    // the two paragraphs once the floor asks for it.
    expect(gap(doc(true)) - gap(doc(false))).toBe(20);
  });
});
