import { Schema, type Node as PMNode } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { describe, expect, it } from 'vitest';
import type {
  ResolvedTableStyle,
  TableLook,
} from '@shadow-garden/bapbong-contracts';
import {
  applyTableStyle,
  currentTableStyle,
  setTableLook,
  WORD_DEFAULT_TABLE_LOOK,
} from './table-style.js';
import { insertTable } from './insert.js';

// Minimal schema, same policy as commands.spec: commands key everything by
// NAME, so no model dependency — a backend could drive this identically.
const schema = new Schema({
  nodes: {
    doc: { content: 'block+', attrs: { tableStyles: { default: null } } },
    paragraph: { group: 'block', content: 'inline*', toDOM: () => ['p', 0] },
    text: { group: 'inline' },
    table: {
      group: 'block',
      content: 'table_row+',
      attrs: { styleId: { default: null }, look: { default: null } },
      toDOM: () => ['table', ['tbody', 0]],
    },
    table_row: {
      content: 'table_cell+',
      attrs: { header: { default: false } },
      toDOM: () => ['tr', 0],
    },
    table_cell: {
      content: 'block+',
      attrs: { colspan: { default: 1 } },
      toDOM: () => ['td', 0],
    },
  },
  marks: {
    strong: { toDOM: () => ['strong', 0] },
    em: { toDOM: () => ['em', 0] },
    textColor: { attrs: { color: {} }, toDOM: () => ['span', 0] },
    fontFamily: { attrs: { family: {} }, toDOM: () => ['span', 0] },
    fontSize: { attrs: { size: {} }, toDOM: () => ['span', 0] },
  },
});

const BOLD_HEADER: ResolvedTableStyle = {
  table: {},
  cond: {
    firstRow: { font: { bold: true, color: '#FFFFFF' } },
    lastRow: { font: { italic: true } },
  },
  bands: { row: 1, col: 1 },
};
const QUIET: ResolvedTableStyle = {
  table: {},
  cond: {},
  bands: { row: 0, col: 0 },
};

type Marks = string[];
const cellP = (text: string, marks: Marks = [], attrs = {}) =>
  schema.nodes['table_cell'].create(attrs, [
    schema.nodes['paragraph'].create(
      null,
      text
        ? [
            schema.text(
              text,
              marks.map((m) =>
                m === 'color'
                  ? schema.marks['textColor'].create({ color: '#FFFFFF' })
                  : schema.marks[m].create(),
              ),
            ),
          ]
        : [],
    ),
  ]);

function stateWith(
  tableAttrs: Record<string, unknown>,
  rows: PMNode[],
  docAttrs: Record<string, unknown> = {},
) {
  const doc = schema.nodes['doc'].create(docAttrs, [
    schema.nodes['table'].create(tableAttrs, rows),
  ]);
  // Caret inside the first cell's paragraph.
  const state = EditorState.create({ doc });
  return state.apply(state.tr.setSelection(TextSelection.create(doc, 4)));
}

const run = (state: EditorState, cmd: ReturnType<typeof applyTableStyle>) => {
  let out = state;
  const ok = cmd.run(state, (tr) => (out = state.apply(tr)));
  return { ok, state: out };
};

const marksAt = (doc: PMNode, row: number, col: number): string[] => {
  const text = doc.child(0).child(row).child(col).child(0).firstChild;
  return (text?.marks ?? []).map((m) => m.type.name).sort();
};

describe('applyTableStyle', () => {
  it('sets the pair, injects the sheet, and bakes the new fonts on', () => {
    const state = stateWith({}, [
      schema.nodes['table_row'].create(null, [cellP('h1'), cellP('h2')]),
      schema.nodes['table_row'].create(null, [cellP('a1'), cellP('a2')]),
    ]);
    const { ok, state: next } = run(
      state,
      applyTableStyle({ styleId: 'S', style: BOLD_HEADER }),
    );
    expect(ok).toBe(true);
    const table = next.doc.child(0);
    expect(table.attrs['styleId']).toBe('S');
    expect(table.attrs['look']).toEqual(WORD_DEFAULT_TABLE_LOOK);
    expect(
      (next.doc.attrs['tableStyles'] as Record<string, unknown>)['S'],
    ).toEqual(BOLD_HEADER);
    // Header row gained the style's fonts as MARKS (fonts are mark-borne
    // until the mark model learns negation); the body did not.
    expect(marksAt(next.doc, 0, 0)).toEqual(['strong', 'textColor']);
    expect(marksAt(next.doc, 1, 0)).toEqual([]);
  });

  it('moves style-derived marks on a switch, keeps the user’s', () => {
    // As imported: header baked bold+white, one body cell bolded BY THE USER,
    // one header run un-bolded by the user (diverged both ways).
    const state = stateWith(
      { styleId: 'S', look: WORD_DEFAULT_TABLE_LOOK },
      [
        schema.nodes['table_row'].create(null, [
          cellP('h1', ['strong', 'color']),
          cellP('h2', ['color']), // user removed the bold
        ]),
        schema.nodes['table_row'].create(null, [
          cellP('a1', ['strong']), // user bolded a body cell
          cellP('a2'),
        ]),
      ],
      { tableStyles: { S: BOLD_HEADER, Q: QUIET } },
    );
    const { state: next } = run(state, applyTableStyle({ styleId: 'Q' }));
    expect(marksAt(next.doc, 0, 0)).toEqual([]); // style's bold+color moved off
    expect(marksAt(next.doc, 0, 1)).toEqual([]); // color was the style's too
    expect(marksAt(next.doc, 1, 0)).toEqual(['strong']); // user's bold stays
    // But the user's REMOVED bold in h2 stays removed — its strong was never
    // re-added because the old value there diverged from the style.
    expect(next.doc.child(0).attrs['styleId']).toBe('Q');
  });

  it('styleId null strips the style and its baked fonts', () => {
    const state = stateWith(
      { styleId: 'S', look: WORD_DEFAULT_TABLE_LOOK },
      [
        schema.nodes['table_row'].create(null, [
          cellP('h1', ['strong', 'color']),
        ]),
      ],
      { tableStyles: { S: BOLD_HEADER } },
    );
    const { state: next } = run(state, applyTableStyle({ styleId: null }));
    expect(next.doc.child(0).attrs['styleId']).toBeNull();
    expect(next.doc.child(0).attrs['look']).toBeNull();
    expect(marksAt(next.doc, 0, 0)).toEqual([]);
  });

  it('refuses an id neither the sheet nor the caller defines', () => {
    const state = stateWith({}, [
      schema.nodes['table_row'].create(null, [cellP('x')]),
    ]);
    expect(applyTableStyle({ styleId: 'Nope' }).run(state)).toBe(false);
  });

  it('does nothing outside a table', () => {
    const doc = schema.nodes['doc'].create(null, [
      schema.nodes['paragraph'].create(null, [schema.text('plain')]),
    ]);
    const state = EditorState.create({ doc });
    expect(applyTableStyle({ styleId: 'S', style: QUIET }).run(state)).toBe(
      false,
    );
  });
});

describe('setTableLook', () => {
  it('flipping a gate moves the gated region’s fonts', () => {
    const state = stateWith(
      { styleId: 'S', look: WORD_DEFAULT_TABLE_LOOK },
      [
        schema.nodes['table_row'].create(null, [
          cellP('h1', ['strong', 'color']),
        ]),
        schema.nodes['table_row'].create(null, [cellP('mid')]),
        schema.nodes['table_row'].create(null, [cellP('last')]),
      ],
      { tableStyles: { S: BOLD_HEADER } },
    );
    // lastRow ON: the last row gains the branch's italic.
    const on = run(state, setTableLook({ lastRow: true })).state;
    expect(marksAt(on.doc, 2, 0)).toEqual(['em']);
    expect((on.doc.child(0).attrs['look'] as TableLook).lastRow).toBe(true);
    // …and OFF strips it again.
    const off = run(on, setTableLook({ lastRow: false })).state;
    expect(marksAt(off.doc, 2, 0)).toEqual([]);
    // Header untouched throughout.
    expect(marksAt(off.doc, 0, 0)).toEqual(['strong', 'textColor']);
  });

  it('needs a styled table', () => {
    const state = stateWith({}, [
      schema.nodes['table_row'].create(null, [cellP('x')]),
    ]);
    expect(setTableLook({ lastRow: true }).run(state)).toBe(false);
  });
});

describe('currentTableStyle', () => {
  it('reports the pair with Word’s default look', () => {
    const state = stateWith({ styleId: 'S' }, [
      schema.nodes['table_row'].create(null, [cellP('x')]),
    ]);
    expect(currentTableStyle(state)).toEqual({
      styleId: 'S',
      look: WORD_DEFAULT_TABLE_LOOK,
    });
  });
});

describe('insertTable with a style', () => {
  // Word's new table carries the "Table Grid" STYLE. A direct border grid
  // would outrank whatever style the user picks next — black vertical
  // lines under Medium Shading was the reported bug — so with a style the
  // table holds the style pair only, and the definition joins the sheet.
  const grid: ResolvedTableStyle = {
    table: { borders: { top: { width: 1, style: 'solid', color: '#000' } } },
    cond: {},
    bands: { row: 0, col: 0 },
  };
  const paraState = () => {
    const doc = schema.nodes['doc'].create(null, [
      schema.nodes['paragraph'].create(null, [schema.text('x')]),
    ]);
    const state = EditorState.create({ doc });
    return state.apply(state.tr.setSelection(TextSelection.create(doc, 1)));
  };

  it('is born with the style pair and injects the definition', () => {
    const { ok, state } = run(
      paraState(),
      insertTable(2, 2, { styleId: 'TableGrid', style: grid }),
    );
    expect(ok).toBe(true);
    let table: PMNode | null = null;
    state.doc.descendants((n) => {
      if (n.type.name === 'table') table = n;
      return !table;
    });
    if (!table) throw new Error('expected an inserted table');
    const t: PMNode = table;
    expect(t.attrs['styleId']).toBe('TableGrid');
    expect(t.attrs['look']).toEqual(WORD_DEFAULT_TABLE_LOOK);
    expect(state.doc.attrs['tableStyles']).toEqual({ TableGrid: grid });
  });

  it('without a style the table stays unstyled (the direct-grid fallback)', () => {
    const { state } = run(paraState(), insertTable(1, 1));
    let styleId: unknown = 'unset';
    state.doc.descendants((n) => {
      if (n.type.name === 'table') styleId = n.attrs['styleId'];
      return true;
    });
    expect(styleId).toBeNull();
    expect(state.doc.attrs['tableStyles']).toBeNull();
  });
});
