import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';
import { PmDocSession } from './pm-session.js';

const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;

/** The first table of `doc`, as rows of cell texts. */
function rows(doc: PMNode): string[][] {
  let table: PMNode | null = null;
  doc.descendants((n) => {
    if (!table && n.type.name === 'table') table = n;
    return !table;
  });
  const out: string[][] = [];
  (table as PMNode | null)?.forEach((row) => {
    const r: string[] = [];
    row.forEach((cell) => r.push(cell.textContent));
    out.push(r);
  });
  return out;
}

/** Every cell's colwidth, row by row, and each row's total. */
function widths(doc: PMNode): number[][] {
  const out: number[][] = [];
  doc.descendants((n) => {
    if (n.type.name !== 'table_row') return true;
    const r: number[] = [];
    n.forEach((cell) =>
      r.push(...((cell.attrs['colwidth'] as number[]) ?? [])),
    );
    out.push(r);
    return false;
  });
  return out;
}
const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);

async function withTable(): Promise<HeadlessSession> {
  const s = await HeadlessSession.open(
    await exportDocx(
      schema.node('doc', null, [
        schema.node('paragraph', null, [schema.text('Intro')]),
      ]),
    ),
  );
  await s.insertContent(
    [
      {
        table: [
          ['Item', 'Qty', 'Price'],
          ['Pen', '2', '1.00'],
          [{ text: 'Total', colspan: 2 }, '2.00'],
        ],
        widths: ['50%', '20%', '30%'],
      },
    ],
    { position: 'document_end' },
  );
  return s;
}

describe('edit_table: columns', () => {
  it('inserts an empty column, widening a merged cell that spans the insertion point; the table keeps its width', async () => {
    const s = await withTable();
    const before = sum(widths(docOf(s))[0]);
    const r = await s.editTable(0, { insertColumns: { at: 1 } });
    expect(r.cols).toBe(4);
    expect(rows(docOf(s))).toEqual([
      ['Item', '', 'Qty', 'Price'],
      ['Pen', '', '2', '1.00'],
      ['Total', '2.00'], // "Total" spans columns 0-2 now
    ]);
    const w = widths(docOf(s));
    expect(w.every((row) => row.length === 4)).toBe(true);
    expect(Math.abs(sum(w[0]) - before)).toBeLessThanOrEqual(4);
    // Appending: omit at.
    await s.editTable(0, { insertColumns: { count: 2 } });
    expect(rows(docOf(s))[0]).toEqual(['Item', '', 'Qty', 'Price', '', '']);
  });

  it('deletes columns: a merged cell narrows, a plain one goes; the rest widen to keep the table as wide', async () => {
    const s = await withTable();
    const before = sum(widths(docOf(s))[0]);
    await s.editTable(0, { deleteColumns: [1] });
    expect(rows(docOf(s))).toEqual([
      ['Item', 'Price'],
      ['Pen', '1.00'],
      ['Total', '2.00'],
    ]);
    const w = widths(docOf(s));
    expect(w.every((row) => row.length === 2)).toBe(true);
    expect(Math.abs(sum(w[0]) - before)).toBeLessThanOrEqual(2);
    // Keep at least one column; out of range is named.
    await expect(s.editTable(0, { deleteColumns: [0, 1] })).rejects.toThrow(
      /at least one column/,
    );
    await expect(s.editTable(0, { deleteColumns: [5] })).rejects.toThrow(
      /out of range/,
    );
  });

  it('inserted rows in the same call cover the columns the table has by then', async () => {
    const s = await withTable();
    await s.editTable(0, {
      insertColumns: { at: 3 },
      insertRows: { rows: [['a', 'b', 'c', 'd']] },
    });
    expect(rows(docOf(s)).at(-1)).toEqual(['a', 'b', 'c', 'd']);
  });

  it('works on the grid when a cell is merged down from the row above', async () => {
    // A spans rows 0-1 in column 0: row 1 holds only C, in column 1.
    const cell = (t: string, attrs: Record<string, unknown> = {}) =>
      schema.nodes['table_cell'].create(
        { colwidth: [100], ...attrs },
        schema.node('paragraph', null, [schema.text(t)]),
      );
    const table = schema.nodes['table'].create(null, [
      schema.nodes['table_row'].create(null, [
        cell('A', { rowspan: 2 }),
        cell('B'),
      ]),
      schema.nodes['table_row'].create(null, [cell('C')]),
    ]);
    let state = EditorState.create({
      doc: schema.node('doc', null, [table]),
    });
    const s = new PmDocSession({
      getState: () => state,
      apply: (tr: Transaction) => {
        state = state.apply(tr);
      },
      getVersion: () => 'v',
      meta: () => ({}),
      save: async () => undefined,
    });
    await s.editTable(0, { insertColumns: { at: 1 } });
    expect(rows(state.doc)).toEqual([
      ['A', '', 'B'],
      ['', 'C'],
    ]);
    // Row 1's own cells sit in columns 1 and 2, so they take those widths.
    expect(widths(state.doc)).toEqual([
      [67, 67, 67],
      [67, 67],
    ]);
    await s.editTable(0, { deleteColumns: [0] });
    expect(rows(state.doc)).toEqual([
      ['', 'B'],
      ['', 'C'],
    ]);
  });
});

describe('edit_table: delete_table', () => {
  it('removes the table, and only the table', async () => {
    const s = await withTable();
    const r = await s.editTable(0, { deleteTable: true });
    expect(r).toMatchObject({ rows: 0, cols: 0 });
    expect((await s.snapshot()).blocks.map((b) => b.text)).toEqual(['Intro']);
  });

  it('leaves an empty paragraph when the table was all there was', async () => {
    const s = await withTable();
    await s.deleteBlocks(0, 1); // "Intro"
    await s.editTable(0, { deleteTable: true });
    expect(docOf(s).childCount).toBe(1);
    expect(docOf(s).child(0).type.name).toBe('paragraph');
  });

  it('refuses other changes in the same call', async () => {
    const s = await withTable();
    await expect(
      s.editTable(0, { deleteTable: true, header: true }),
    ).rejects.toThrow(ContentError);
    expect(rows(docOf(s))).toHaveLength(3);
  });
});
