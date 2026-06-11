import { Schema } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import { createEditingState, wordRangeAt } from './input-bridge.js';

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
