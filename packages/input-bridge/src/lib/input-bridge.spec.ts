import { Schema } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import { createEditingState } from './input-bridge.js';

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
