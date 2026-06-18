import { Schema } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import {
  addCommentTr,
  createEditingState,
  deleteCommentTr,
  editCommentTr,
  moveCaretCommand,
  replyCommentTr,
  resolveCommentTr,
  splitListItem,
  wordRangeAt,
  type CommentNode,
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

describe('comment authoring commands', () => {
  const cSchema = new Schema({
    nodes: {
      doc: { content: 'block+', attrs: { comments: { default: null } } },
      paragraph: { group: 'block', content: 'inline*' },
      text: { group: 'inline' },
    },
    marks: { comment: { attrs: { ids: {} } } },
  });
  const make = () =>
    createEditingState(cSchema.node('doc', null, [cSchema.node('paragraph', null, [cSchema.text('hello world')])]));
  const getComments = (state: { doc: { attrs: Record<string, unknown> } }) =>
    (state.doc.attrs['comments'] as CommentNode[] | null) ?? [];
  const body = (t: string) => ({ type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: t }] }] });
  const markedIds = (state: { doc: { descendants: (f: (n: { isText: boolean; marks: readonly { type: { name: string }; attrs: Record<string, unknown> }[] }) => void) => void } }) => {
    const ids: number[][] = [];
    state.doc.descendants((n) => {
      if (!n.isText) return;
      const m = n.marks.find((mk) => mk.type.name === 'comment');
      if (m) ids.push(m.attrs['ids'] as number[]);
    });
    return ids;
  };

  it('add tags the range and appends a thread root', () => {
    let state = make();
    // "world" spans pos 7..12.
    state = state.apply(addCommentTr(state, { from: 7, to: 12 }, { author: 'A', date: 'd1', body: body('note') }));
    expect(getComments(state)).toHaveLength(1);
    expect(getComments(state)[0]).toMatchObject({ id: 1, parentId: null, author: 'A', resolved: false });
    expect(markedIds(state)).toEqual([[1]]); // only "world" carries the mark
  });

  it('reply / resolve / edit update the thread', () => {
    let state = make();
    state = state.apply(addCommentTr(state, { from: 7, to: 12 }, { author: 'A', date: 'd1', body: body('note') }));
    state = state.apply(replyCommentTr(state, 1, { author: 'B', date: 'd2', body: body('reply') }));
    expect(getComments(state)[1]).toMatchObject({ id: 2, parentId: 1, author: 'B' });
    state = state.apply(resolveCommentTr(state, 1, true));
    expect(getComments(state)[0].resolved).toBe(true);
    state = state.apply(editCommentTr(state, 1, body('edited')));
    expect(getComments(state)[0].body).toEqual(body('edited'));
  });

  it('delete removes the subtree and strips the mark', () => {
    let state = make();
    state = state.apply(addCommentTr(state, { from: 7, to: 12 }, { author: 'A', date: 'd1', body: body('note') }));
    state = state.apply(replyCommentTr(state, 1, { author: 'B', date: 'd2', body: body('reply') }));
    state = state.apply(deleteCommentTr(state, 1));
    expect(getComments(state)).toHaveLength(0); // root + reply gone
    expect(markedIds(state)).toEqual([]); // mark stripped
  });

  it('add is undoable (mark + thread revert together)', () => {
    let state = make();
    state = state.apply(addCommentTr(state, { from: 7, to: 12 }, { author: 'A', date: 'd1', body: body('note') }));
    const ok = undo(state, (tr) => (state = state.apply(tr)));
    expect(ok).toBe(true);
    expect(getComments(state)).toHaveLength(0);
    expect(markedIds(state)).toEqual([]);
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
