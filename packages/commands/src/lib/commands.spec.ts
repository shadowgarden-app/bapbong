import { Schema, type Node as PMNode } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Command } from '@shadow-garden/bapbong-contracts';
import {
  toggleMarkCommand,
  isMarkActive,
  setFontSize,
  activeFontSize,
  setFontFamily,
  activeFontFamily,
  setTextColor,
  activeTextColor,
  setHighlight,
  activeHighlight,
  clearMarks,
} from './marks.js';
import { setAlign, activeAlign, toggleHeading } from './paragraph.js';
import { cellAt, setCellBackground, setCellsAttrs } from './table.js';
import { insertImage, insertTable, pageBreakCommand, setLink } from './insert.js';
import { deleteSelectionCommand } from './edit.js';
import { deleteColumn, deleteRow, deleteTable, insertColumn, insertRow, mergeCells } from './table-structure.js';
import { insertSectionBreak, removeSectionBreak, setColumns } from './sections.js';
import { toggleList } from './list.js';
import { defaultCommands } from './registry.js';

// A minimal schema standing in for the real document schema — the commands key
// nodes/marks by *name*, so this is all they need. Building it here (no model
// dependency) is also exactly how a backend would: construct a Schema +
// EditorState in Node and drive commands, with no DOM anywhere.
const schema = new Schema({
  nodes: {
    doc: { content: 'block+', attrs: { numbering: { default: null }, sections: { default: null } } },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: { align: { default: null }, pageBreakBefore: { default: false }, list: { default: null }, heading: { default: null } },
      toDOM: () => ['p', 0],
    },
    text: { group: 'inline' },
    image: { inline: true, group: 'inline', attrs: { src: {}, alt: { default: '' } }, toDOM: () => ['img'] },
    table: { group: 'block', content: 'table_row+', toDOM: () => ['table', ['tbody', 0]] },
    table_row: { content: 'table_cell+', toDOM: () => ['tr', 0] },
    table_cell: {
      content: 'block+',
      attrs: {
        background: { default: null },
        vAlign: { default: null },
        colwidth: { default: null },
        colspan: { default: 1 },
        rowspan: { default: 1 },
        borders: { default: null },
      },
      toDOM: () => ['td', 0],
    },
  },
  marks: {
    strong: { toDOM: () => ['strong', 0] },
    em: { toDOM: () => ['em', 0] },
    underline: { toDOM: () => ['u', 0] },
    strike: { toDOM: () => ['s', 0] },
    vertAlign: { attrs: { value: {} }, toDOM: () => ['sup', 0] },
    fontSize: { attrs: { size: {} }, toDOM: () => ['span', 0] },
    fontFamily: { attrs: { family: {} }, toDOM: () => ['span', 0] },
    textColor: { attrs: { color: {} }, toDOM: () => ['span', 0] },
    highlight: { attrs: { color: {} }, toDOM: () => ['span', 0] },
    link: { attrs: { href: {} }, inclusive: false, toDOM: () => ['a', 0] },
  },
});

const n = (type: string, attrs: object | null, content?: unknown) =>
  schema.node(type, attrs, content as never);

/** Apply a command (capturing its transaction) and return the resulting state. */
function apply(state: EditorState, cmd: Command): EditorState {
  let next = state;
  cmd.run(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

/** First node of the given type in the doc, or null. */
function findNode(state: EditorState, typeName: string): PMNode | null {
  let found: PMNode | null = null;
  state.doc.descendants((node) => {
    if (!found && node.type.name === typeName) found = node;
  });
  return found;
}

/** A 2×2 table doc with the caret in the top-left cell (pos 4). */
function gridState(): EditorState {
  const cell = (t: string) => n('table_cell', null, n('paragraph', null, schema.text(t)));
  const row = (a: string, b: string) => n('table_row', null, [cell(a), cell(b)]);
  const doc = n('doc', null, n('table', null, [row('a', 'b'), row('c', 'd')]));
  const state = EditorState.create({ schema, doc });
  return state.apply(state.tr.setSelection(TextSelection.create(doc, 4)));
}

/** A doc of one paragraph "hello", selection covering the whole word. */
function paraState(): EditorState {
  const doc = n('doc', null, n('paragraph', null, schema.text('hello')));
  const state = EditorState.create({ schema, doc });
  return state.apply(state.tr.setSelection(TextSelection.create(doc, 1, 6)));
}

describe('commands (headless / Node — backend-shaped usage)', () => {
  it('toggleMarkCommand marks the selection and reports active', () => {
    const bold = toggleMarkCommand('bold', 'strong');
    const before = paraState();
    expect(bold.isActive?.(before)).toBe(false);

    const after = apply(before, bold);
    expect(after.doc.rangeHasMark(1, 6, schema.marks['strong'])).toBe(true);
    expect(isMarkActive(after, 'strong')).toBe(true);
  });

  it('setFontSize applies/clears a value mark; activeFontSize reads it', () => {
    const sized = apply(paraState(), setFontSize(18));
    expect(activeFontSize(sized)).toBe(18);
    // re-sizing replaces the value (no duplicate marks)
    const resized = apply(sized, setFontSize(24));
    expect(activeFontSize(resized)).toBe(24);
    // null clears it
    expect(activeFontSize(apply(resized, setFontSize(null)))).toBeNull();
    // mixed selection → null
    expect(activeFontSize(paraState())).toBeNull();
  });

  it('setFontFamily applies/clears the family; activeFontFamily reads it', () => {
    const f = apply(paraState(), setFontFamily('Georgia'));
    expect(activeFontFamily(f)).toBe('Georgia');
    expect(activeFontFamily(apply(f, setFontFamily(null)))).toBeNull();
  });

  it('setTextColor applies/clears the colour; activeTextColor reads it', () => {
    const c = apply(paraState(), setTextColor('#e24b4a'));
    expect(activeTextColor(c)).toBe('#e24b4a');
    expect(activeTextColor(apply(c, setTextColor(null)))).toBeNull();
  });

  it('setHighlight applies/clears the colour; activeHighlight reads it', () => {
    const h = apply(paraState(), setHighlight('#fff59d'));
    expect(activeHighlight(h)).toBe('#fff59d');
    expect(activeHighlight(apply(h, setHighlight(null)))).toBeNull();
  });

  it('clearMarks strips every mark from the selection', () => {
    // pile on a few marks, then clear them all
    let s = apply(paraState(), toggleMarkCommand('bold', 'strong'));
    s = apply(s, setTextColor('#e24b4a'));
    s = apply(s, setFontSize(20));
    expect(isMarkActive(s, 'strong')).toBe(true);
    const cleared = apply(s, clearMarks());
    expect(isMarkActive(cleared, 'strong')).toBe(false);
    expect(activeTextColor(cleared)).toBeNull();
    expect(activeFontSize(cleared)).toBeNull();
  });

  it('run() without dispatch is a probe — returns true but does not mutate', () => {
    const before = paraState();
    const applied = toggleMarkCommand('bold', 'strong').run(before); // no dispatch
    expect(applied).toBe(true);
    expect(before.doc.rangeHasMark(1, 6, schema.marks['strong'])).toBe(false);
  });

  it('a command for a mark absent from the schema stays inert', () => {
    // `comment` isn't in this schema (it's plugin-contributed) → inactive + no-op.
    const cmd = toggleMarkCommand('comment');
    const before = paraState();
    expect(cmd.isEnabled?.(before)).toBe(false);
    expect(cmd.run(before)).toBe(false);
  });

  it('setAlign sets paragraph alignment; activeAlign reads it back', () => {
    const before = paraState();
    expect(activeAlign(before)).toBeNull();

    const after = apply(before, setAlign('center'));
    expect(after.doc.firstChild?.attrs['align']).toBe('center');
    expect(activeAlign(after)).toBe('center');
    expect(setAlign('center').isActive?.(after)).toBe(true);
    expect(setAlign('right').isActive?.(after)).toBe(false);
  });

  it('cellAt locates the containing cell; setCellBackground fills it', () => {
    const doc = n(
      'doc',
      null,
      n('table', null, n('table_row', null, n('table_cell', null, n('paragraph', null, schema.text('x'))))),
    );
    let state = EditorState.create({ schema, doc });
    state = state.apply(state.tr.setSelection(TextSelection.create(doc, 4))); // inside the cell's text

    const cell = cellAt(state);
    expect(cell?.node.type.name).toBe('table_cell');

    const after = apply(state, setCellBackground('#ffd600'));
    expect(cellAt(after)?.node.attrs['background']).toBe('#ffd600');
  });

  it('setCellBackground is disabled outside a table', () => {
    expect(setCellBackground('#fff').isEnabled?.(paraState())).toBe(false);
  });

  it('defaultCommands() is a name-keyed registry of the built-in static commands', () => {
    const commands = defaultCommands();
    expect(commands.has('bold')).toBe(true);
    expect(commands.has('align-center')).toBe(true);
    expect([...commands].map((c) => c.name)).toContain('italic');

    // Driving the editor through the registry (as a toolbar/backend would).
    const bold = commands.get('bold');
    if (!bold) throw new Error('expected a "bold" command in the registry');
    const after = apply(paraState(), bold);
    expect(after.doc.rangeHasMark(1, 6, schema.marks['strong'])).toBe(true);
  });

  it('superscript toggles vertAlign with its attr; sub/super are exclusive', () => {
    const sup = toggleMarkCommand('superscript', 'vertAlign', { value: 'super' });
    const sub = toggleMarkCommand('subscript', 'vertAlign', { value: 'sub' });
    const after = apply(paraState(), sup);
    expect(sup.isActive?.(after)).toBe(true);
    expect(sub.isActive?.(after)).toBe(false); // same mark type → mutually exclusive
    expect(isMarkActive(after, 'vertAlign', { value: 'super' })).toBe(true);
  });

  it('pageBreakCommand toggles pageBreakBefore on the paragraph', () => {
    const cmd = pageBreakCommand();
    const before = paraState();
    expect(cmd.isActive?.(before)).toBe(false);
    const after = apply(before, cmd);
    expect(after.doc.firstChild?.attrs['pageBreakBefore']).toBe(true);
    expect(cmd.isActive?.(after)).toBe(true);
  });

  it('insertTable inserts a rows×cols grid of cells', () => {
    const after = apply(paraState(), insertTable(2, 3));
    const table = findNode(after, 'table');
    if (!table) throw new Error('expected an inserted table');
    expect(table.childCount).toBe(2); // rows
    expect(table.firstChild?.childCount).toBe(3); // cells in the first row
  });

  it('insertImage inserts an inline image; setLink needs a range', () => {
    const img = findNode(apply(paraState(), insertImage('data:img', 'pic')), 'image');
    expect(img?.attrs['src']).toBe('data:img');

    const linked = apply(paraState(), setLink('https://x.test'));
    expect(linked.doc.rangeHasMark(1, 6, schema.marks['link'])).toBe(true);

    // empty selection → link disabled
    const ps = paraState();
    const caret = ps.apply(ps.tr.setSelection(TextSelection.create(ps.doc, 1)));
    expect(setLink('https://x.test').isEnabled?.(caret)).toBe(false);
  });

  it('insertRow adds a row; insertColumn adds a cell to every row', () => {
    const table0 = findNode(gridState(), 'table');
    expect(table0?.childCount).toBe(2);

    const rowsAfter = findNode(apply(gridState(), insertRow(true)), 'table');
    expect(rowsAfter?.childCount).toBe(3); // 2 → 3 rows

    const colsAfter = findNode(apply(gridState(), insertColumn(true)), 'table');
    expect(colsAfter?.childCount).toBe(2); // still 2 rows
    expect(colsAfter?.firstChild?.childCount).toBe(3); // 2 → 3 cells per row
    expect(colsAfter?.lastChild?.childCount).toBe(3);
  });

  it('setCellsAttrs sets the same attrs on every given cell in one tr', () => {
    const state = gridState();
    const positions: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.type.name === 'table_cell') positions.push(pos);
    });
    expect(positions.length).toBe(4);
    const after = apply(state, setCellsAttrs(positions, { background: '#abcdef' }));
    const bgs: Array<string | null> = [];
    after.doc.descendants((node) => {
      if (node.type.name === 'table_cell') bgs.push(node.attrs['background'] as string | null);
    });
    expect(bgs.every((b) => b === '#abcdef')).toBe(true);
  });

  it('insert row/column are disabled outside a table', () => {
    expect(insertRow(true).isEnabled?.(paraState())).toBe(false);
    expect(insertColumn(true).isEnabled?.(gridState())).toBe(true);
  });

  it('deleteRow removes a row; deleteColumn removes a cell from every row', () => {
    const rowsAfter = findNode(apply(gridState(), deleteRow()), 'table');
    expect(rowsAfter?.childCount).toBe(1); // 2 → 1 row

    const colsAfter = findNode(apply(gridState(), deleteColumn()), 'table');
    expect(colsAfter?.childCount).toBe(2); // still 2 rows
    expect(colsAfter?.firstChild?.childCount).toBe(1); // 2 → 1 cell per row
    expect(colsAfter?.lastChild?.childCount).toBe(1);
  });

  it('deleting the last row/column removes the whole table', () => {
    // 1×1 table: caret in its only cell.
    const oneByOne = () => {
      const doc = n('doc', null, n('table', null, n('table_row', null, n('table_cell', null, n('paragraph', null, schema.text('x'))))));
      const s = EditorState.create({ schema, doc });
      return s.apply(s.tr.setSelection(TextSelection.create(doc, 4)));
    };
    expect(findNode(apply(oneByOne(), deleteRow()), 'table')).toBeNull();
    expect(findNode(apply(oneByOne(), deleteColumn()), 'table')).toBeNull();
    expect(findNode(apply(gridState(), deleteTable()), 'table')).toBeNull();
  });

  it('delete row/column/table are disabled outside a table', () => {
    expect(deleteRow().isEnabled?.(paraState())).toBe(false);
    expect(deleteColumn().isEnabled?.(paraState())).toBe(false);
    expect(deleteTable().isEnabled?.(gridState())).toBe(true);
  });

  it('mergeCells merges the top row into one colspan-2 cell, appending content', () => {
    const state = gridState(); // 2×2: row0 [a,b], row1 [c,d]
    const cellPos: number[] = []; // table_cell node positions
    state.doc.descendants((node, pos) => {
      if (node.type.name === 'table_cell') cellPos.push(pos);
    });
    // top-row cells are the first two; block grid 1 row × 2 cols
    const block = [
      { pos: cellPos[0], row: 0, col: 0 },
      { pos: cellPos[1], row: 0, col: 1 },
    ];
    const after = apply(state, mergeCells(block, 1, 2));
    const table = findNode(after, 'table');
    expect(table?.firstChild?.childCount).toBe(1); // row 0: 2 cells → 1
    expect(table?.firstChild?.firstChild?.attrs['colspan']).toBe(2);
    expect(table?.lastChild?.childCount).toBe(2); // row 1 untouched
    expect(table?.firstChild?.firstChild?.textContent).toBe('ab'); // 'a' + appended 'b'
  });

  // A 3-paragraph doc with the caret in the 2nd block (index 1).
  const threePara = () => {
    const doc = n('doc', null, [
      n('paragraph', null, schema.text('aa')),
      n('paragraph', null, schema.text('bb')),
      n('paragraph', null, schema.text('cc')),
    ]);
    const s = EditorState.create({ schema, doc });
    return s.apply(s.tr.setSelection(TextSelection.create(doc, 5)));
  };

  it('insertSectionBreak splits the section at the caret block, keeping columns', () => {
    const after = apply(threePara(), insertSectionBreak({ newPage: true }));
    const sections = after.doc.attrs['sections'] as { blockCount: number; newPage: boolean; columns: { count: number } }[];
    expect(sections.map((s) => s.blockCount)).toEqual([2, 1]); // break after block 1
    expect(sections[0].newPage).toBe(false); // first part inherits original start
    expect(sections[1].newPage).toBe(true); // new part starts on a new page
    expect(sections.every((s) => s.columns.count === 1)).toBe(true); // columns unchanged
  });

  it('setColumns sets the caret section column count (and reports active)', () => {
    const after = apply(threePara(), setColumns(2));
    const sections = after.doc.attrs['sections'] as { blockCount: number; columns: { count: number } }[];
    expect(sections).toHaveLength(1);
    expect(sections[0].columns.count).toBe(2);
    expect(setColumns(2).isActive?.(after)).toBe(true);
    expect(setColumns(1).isActive?.(after)).toBe(false);
  });

  it('insertSectionBreak is disabled in a single-block doc', () => {
    const doc = n('doc', null, n('paragraph', null, schema.text('only')));
    const s = EditorState.create({ schema, doc });
    expect(insertSectionBreak({ newPage: true }).isEnabled?.(s)).toBe(false);
  });

  it('removeSectionBreak merges two sections — lower columns win, upper page-start kept', () => {
    const doc = n(
      'doc',
      {
        sections: [
          { blockCount: 1, columns: { count: 1, gap: 0 }, newPage: false },
          { blockCount: 2, columns: { count: 2, gap: 28 }, newPage: true },
        ],
      },
      [
        n('paragraph', null, schema.text('a')),
        n('paragraph', null, schema.text('b')),
        n('paragraph', null, schema.text('c')),
      ],
    );
    const after = apply(EditorState.create({ schema, doc }), removeSectionBreak(0));
    const sections = after.doc.attrs['sections'] as { blockCount: number; columns: { count: number }; newPage: boolean }[];
    expect(sections).toHaveLength(1);
    expect(sections[0].blockCount).toBe(3);
    expect(sections[0].columns.count).toBe(2); // following section's layout wins
    expect(sections[0].newPage).toBe(false); // merged range keeps the upper's start
  });

  it('removeSectionBreak is a no-op for an out-of-range boundary', () => {
    expect(removeSectionBreak(5).run(threePara(), undefined)).toBe(false);
  });

  it('deleteSelectionCommand removes the selected text', () => {
    const cmd = deleteSelectionCommand();
    expect(cmd.isEnabled?.(paraState())).toBe(true);
    const after = apply(paraState(), cmd);
    expect(after.doc.textContent).not.toContain('hello');
  });

  it('registry includes the new static commands', () => {
    const commands = defaultCommands();
    for (const name of ['superscript', 'subscript', 'undo', 'redo', 'page-break', 'bullet-list', 'ordered-list', 'heading-1', 'heading-2']) {
      expect(commands.has(name)).toBe(true);
    }
  });

  it('toggleHeading sets/clears the paragraph heading level and reports active', () => {
    const h1 = toggleHeading(1);
    expect(h1.isActive?.(paraState())).toBe(false);
    const on = apply(paraState(), h1);
    expect(findNode(on, 'paragraph')?.attrs['heading']).toBe(1);
    expect(h1.isActive?.(on)).toBe(true);
    // toggling the same level reverts to a body paragraph
    const off = apply(on, toggleHeading(1));
    expect(findNode(off, 'paragraph')?.attrs['heading']).toBeNull();
    // a different level switches
    const h2 = apply(on, toggleHeading(2));
    expect(findNode(h2, 'paragraph')?.attrs['heading']).toBe(2);
  });

  it('toggleList sets/clears the paragraph list attr and reports active', () => {
    const bullet = toggleList('bullet');
    expect(bullet.isActive?.(paraState())).toBe(false);
    const on = apply(paraState(), bullet);
    expect((findNode(on, 'paragraph')?.attrs['list'] as { numId: string }).numId).toBe('bb-bullet');
    expect(bullet.isActive?.(on)).toBe(true);
    // ordered injects a numbering def so the counter can number live
    const ordered = apply(paraState(), toggleList('ordered'));
    expect((ordered.doc.attrs['numbering'] as Record<string, unknown>)['bb-ordered']).toBeTruthy();
    // toggling the same kind again clears it
    const off = apply(on, toggleList('bullet'));
    expect(findNode(off, 'paragraph')?.attrs['list']).toBeNull();
  });
});
