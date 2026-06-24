import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Command } from '@shadow-garden/bapbong-contracts';
import { toggleMarkCommand, isMarkActive } from './marks.js';
import { setAlign, activeAlign } from './paragraph.js';
import { cellAt, setCellBackground } from './table.js';
import { defaultCommands } from './registry.js';

// A minimal schema standing in for the real document schema — the commands key
// nodes/marks by *name*, so this is all they need. Building it here (no model
// dependency) is also exactly how a backend would: construct a Schema +
// EditorState in Node and drive commands, with no DOM anywhere.
const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: { align: { default: null } },
      toDOM: () => ['p', 0],
    },
    text: { group: 'inline' },
    table: { group: 'block', content: 'table_row+', toDOM: () => ['table', ['tbody', 0]] },
    table_row: { content: 'table_cell+', toDOM: () => ['tr', 0] },
    table_cell: {
      content: 'block+',
      attrs: {
        background: { default: null },
        vAlign: { default: null },
        colwidth: { default: null },
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
});
