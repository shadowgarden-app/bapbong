import { Schema } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import {
  backspaceOutdent,
  createEditingState,
  moveCaretCommand,
  shiftListLevel,
  splitListItem,
  wordRangeAt,
} from './input-bridge.js';

// EditorView needs a real DOM, so the headless tests cover the state side:
// plugins, typing transactions and history. The view itself is exercised in
// the playground (browser).
const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*' },
    text: { group: 'inline' },
  },
  marks: {},
});

const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('hello')])]);

describe('createEditingState', () => {
  it('applies text insertions', () => {
    let state = createEditingState(doc);
    state = state.apply(state.tr.insertText('!', 6));
    expect(state.doc.textContent).toBe('hello!');
  });

  it('undoes through prosemirror-history', () => {
    let state = createEditingState(doc);
    state = state.apply(state.tr.insertText('!', 6));
    expect(state.doc.textContent).toBe('hello!');
    const ok = undo(state, (tr) => (state = state.apply(tr)));
    expect(ok).toBe(true);
    expect(state.doc.textContent).toBe('hello');
  });

  it('keeps the selection mapped through edits', () => {
    let state = createEditingState(doc);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
    state = state.apply(state.tr.insertText('xy', 1));
    expect(state.selection.head).toBe(5); // shifted by the 2 inserted chars
  });
});

describe('moveCaretCommand', () => {
  it('collapses by default and extends from the anchor with extend=true', () => {
    let state = createEditingState(doc); // 'hello'
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));

    moveCaretCommand(() => 4)(state, (tr) => (state = state.apply(tr)));
    expect(state.selection.empty).toBe(true);
    expect(state.selection.head).toBe(4);

    moveCaretCommand(() => 6, true)(state, (tr) => (state = state.apply(tr)));
    expect(state.selection.anchor).toBe(4); // anchor kept
    expect(state.selection.head).toBe(6); // head moved

    // compute → null leaves the key to the next handler
    expect(moveCaretCommand(() => null)(state, () => undefined)).toBe(false);
  });
});

describe('splitListItem', () => {
  const listSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*', attrs: { list: { default: null } } },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const listAttrs = { list: { numId: '1', level: 0 } };

  it('splits keeping the list attrs (the new item stays in the list)', () => {
    const d = listSchema.node('doc', null, [
      listSchema.node('paragraph', listAttrs, [listSchema.text('item')]),
    ]);
    let state = createEditingState(d);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 5))); // end of "item"
    const handled = splitListItem(state, (tr) => (state = state.apply(tr)));
    expect(handled).toBe(true);
    expect(state.doc.childCount).toBe(2);
    expect(state.doc.child(1).attrs['list']).toEqual(listAttrs.list);
  });

  it('exits the list on an empty item, and defers outside lists', () => {
    const d = listSchema.node('doc', null, [listSchema.node('paragraph', listAttrs)]);
    let state = createEditingState(d);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
    expect(splitListItem(state, (tr) => (state = state.apply(tr)))).toBe(true);
    expect(state.doc.child(0).attrs['list']).toBeNull(); // left the list

    const plain = createEditingState(
      listSchema.node('doc', null, [listSchema.node('paragraph', null, [listSchema.text('x')])]),
    );
    expect(splitListItem(plain, () => undefined)).toBe(false); // base keymap's turn
  });
});

describe('backspaceOutdent', () => {
  const listSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null }, indent: { default: null } },
      },
      text: { group: 'inline' },
    },
    marks: {},
  });
  const listAttrs = { list: { numId: '1', level: 0 }, indent: { left: 48, hanging: 24 } };

  it('outdents in steps: drop marker → clear indent → defer to join', () => {
    const d = listSchema.node('doc', null, [
      listSchema.node('paragraph', listAttrs, [listSchema.text('item')]),
    ]);
    let state = createEditingState(d);
    const caretAtStart = () =>
      (state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1))));

    // Step 1: drops the marker, keeps the indent, text stays on the same line.
    caretAtStart();
    expect(backspaceOutdent(state, (tr) => (state = state.apply(tr)))).toBe(true);
    expect(state.doc.child(0).attrs['list']).toBeNull();
    expect(state.doc.child(0).attrs['indent']).toEqual(listAttrs.indent);
    expect(state.doc.child(0).textContent).toBe('item');

    // Step 2: clears the indent (caret returns to the margin).
    caretAtStart();
    expect(backspaceOutdent(state, (tr) => (state = state.apply(tr)))).toBe(true);
    expect(state.doc.child(0).attrs['indent']).toBeNull();

    // Step 3: nothing left to outdent → base keymap joins backward.
    caretAtStart();
    expect(backspaceOutdent(state, () => undefined)).toBe(false);
  });

  it('Tab/Shift-Tab shift the level within the definition, moving indent 24px', () => {
    const numSchema = new Schema({
      nodes: {
        doc: { content: 'block+', attrs: { numbering: { default: null } } },
        paragraph: {
          group: 'block',
          content: 'inline*',
          attrs: { list: { default: null }, indent: { default: null } },
        },
        text: { group: 'inline' },
      },
      marks: {},
    });
    const numbering = {
      '1': { key: 'k1', levels: { 0: {}, 1: {} } }, // two levels only
    };
    const d = numSchema.node('doc', { numbering }, [
      numSchema.node('paragraph', { list: { numId: '1', level: 0 } }, [
        numSchema.text('item'),
      ]),
    ]);
    let state = createEditingState(d);
    const apply = (cmd: typeof backspaceOutdent) =>
      cmd(state, (tr) => (state = state.apply(tr)));
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));

    // demote: level 0 → 1, indent appears
    expect(apply(shiftListLevel(1))).toBe(true);
    expect(state.doc.child(0).attrs['list']).toEqual({ numId: '1', level: 1 });
    expect(state.doc.child(0).attrs['indent']).toEqual({ left: 24 });

    // capped at the definition's deepest level (1): no-op returns false
    expect(apply(shiftListLevel(1))).toBe(false);

    // promote: back to level 0, indent clears to null
    expect(apply(shiftListLevel(-1))).toBe(true);
    expect(state.doc.child(0).attrs['list']).toEqual({ numId: '1', level: 0 });
    expect(state.doc.child(0).attrs['indent']).toBeNull();

    // promote at level 0: nothing to do
    expect(apply(shiftListLevel(-1))).toBe(false);

    // outside a list: defers (Tab keeps its default behavior)
    const plain = createEditingState(
      numSchema.node('doc', null, [
        numSchema.node('paragraph', null, [numSchema.text('x')]),
      ]),
    );
    expect(shiftListLevel(1)(plain, () => undefined)).toBe(false);
  });

  it('defers mid-block and on a ranged selection', () => {
    const d = listSchema.node('doc', null, [
      listSchema.node('paragraph', listAttrs, [listSchema.text('item')]),
    ]);
    let mid = createEditingState(d);
    mid = mid.apply(mid.tr.setSelection(TextSelection.create(mid.doc, 3))); // inside "item"
    expect(backspaceOutdent(mid, () => undefined)).toBe(false); // normal char delete

    let ranged = createEditingState(d);
    ranged = ranged.apply(ranged.tr.setSelection(TextSelection.create(ranged.doc, 1, 3)));
    expect(backspaceOutdent(ranged, () => undefined)).toBe(false); // deletes the range
  });
});

describe('wordRangeAt', () => {
  const para = (text: string) => schema.node('doc', null, [schema.node('paragraph', null, [schema.text(text)])]);

  it('finds the word around a position', () => {
    const d = para('hello world');
    // "world" spans pos 7..12; pos 9 is inside it.
    expect(wordRangeAt(d, 9)).toEqual({ from: 7, to: 12 });
  });

  it('treats Vietnamese diacritics as word characters', () => {
    const d = para('xin chào bạn');
    // "chào" spans pos 5..9.
    expect(wordRangeAt(d, 7)).toEqual({ from: 5, to: 9 });
  });

  it('returns null on whitespace between words and outside textblocks', () => {
    const d = para('a b');
    expect(wordRangeAt(d, 2)?.from).toBe(1); // boundary touches "a"
    const gap = para('a  b'); // double space: pos 3 touches no word
    expect(wordRangeAt(gap, 3)).toBeNull();
    expect(wordRangeAt(d, 0)).toBeNull(); // doc-level position
  });
});
