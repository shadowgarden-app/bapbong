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
  cond: {
    firstRow: {
      font: { bold: true, color: '#FFFFFF' },
      background: '#1F4E79',
      borders: { bottom: heavy },
    },
    firstCol: { font: { bold: true } },
    band1Horz: { background: '#D6E3BC' },
    lastRow: { font: { italic: true } },
  },
  bands: { row: 1, col: 1 },
};
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

function docWith(tableAttrs: Record<string, unknown>, rows?: unknown[]) {
  const table = schema.nodes['table'].create(
    tableAttrs,
    (rows as ReturnType<typeof rowNode>[]) ?? [
      rowNode(['h1', 'h2', 'h3']),
      rowNode(['a1', 'a2', 'a3']),
      rowNode(['b1', 'b2', 'b3']),
    ],
  );
  return schema.nodes['doc'].create({ tableStyles: SHEET }, [table]);
}

const flowOf = (doc: ReturnType<typeof docWith>) =>
  toFlowBlocks(doc)[0] as FlowTable;
const runFont = (cellPara: unknown) =>
  (
    (cellPara as FlowParagraph).runs[0] as {
      font: { bold?: boolean; italic?: boolean; family?: string };
    }
  ).font;

describe('tableToFlow applies the live style sheet', () => {
  it('per-region fonts land in the run base, colour beside it', () => {
    const flow = flowOf(docWith({ styleId: 'Probe', look: LOOK }));
    const header = flow.rows[0].cells[1].content[0] as FlowParagraph;
    expect(runFont(header)).toMatchObject({
      bold: true,
      family: 'Cambria',
    });
    expect(header.runs[0]).toMatchObject({ color: '#FFFFFF' });
    // Body cell: the style's base family without the header's bold.
    const body = flow.rows[2].cells[1].content[0] as FlowParagraph;
    expect(runFont(body)).toMatchObject({ bold: false, family: 'Cambria' });
    // First column stays bold below the header.
    expect(runFont(flow.rows[1].cells[0].content[0])).toMatchObject({
      bold: true,
    });
  });

  it('fills, borders and table-level fallbacks follow the gates', () => {
    const flow = flowOf(docWith({ styleId: 'Probe', look: LOOK }));
    expect(flow.rows[0].cells[2].background).toBe('#1F4E79');
    expect(flow.rows[0].cells[2].borders).toMatchObject({ bottom: heavy });
    // Banding counts the body: row 1 is band1 (filled), row 2 is band2.
    expect(flow.rows[1].cells[1].background).toBe('#D6E3BC');
    expect(flow.rows[2].cells[1].background).toBeUndefined();
    // lastRow is gated off by this look.
    expect(runFont(flow.rows[2].cells[1].content[0]).italic).toBe(false);
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
    expect(runFont(plain.rows[0].cells[0].content[0]).bold).toBe(false);
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
    const headerSeg = table.cells[0].lines[0]?.segments[0];
    expect(headerSeg && 'font' in headerSeg ? headerSeg.font.bold : null).toBe(
      true,
    );
  });
});
