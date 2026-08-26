import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { mathRangeAt } from './equation-plugin';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*' },
    text: { group: 'inline' },
  },
  marks: { math: {}, em: {} },
});
const math = schema.marks['math'].create();
const em = schema.marks['em'].create();

const state = (children: unknown[], caret: number) => {
  const doc = schema.node('doc', null, [
    schema.node('paragraph', null, children as never),
  ]);
  let s = EditorState.create({ schema, doc });
  s = s.apply(s.tr.setSelection(TextSelection.create(s.doc, caret)));
  return s;
};

describe('mathRangeAt (the equation region under the caret)', () => {
  it('finds the contiguous math range across mark-split text nodes', () => {
    // "ab" + math("ω=" + bold-ish "2πf") + "cd" → math range [3, 8).
    const s = state(
      [
        schema.text('ab'),
        schema.text('ω=', [math]),
        schema.text('2πf', [math, em]),
        schema.text('cd'),
      ],
      5,
    );
    expect(mathRangeAt(s as never, 5)).toEqual({ from: 3, to: 8 });
  });

  it('includes the inclusive end boundary, excludes plain text', () => {
    const s = state([schema.text('ab'), schema.text('ω', [math])], 4);
    // Caret right AFTER the equation still shows the region…
    expect(mathRangeAt(s as never, 4)).toEqual({ from: 3, to: 4 });
    // …the plain text before it does not.
    expect(mathRangeAt(s as never, 2)).toBeNull();
    expect(mathRangeAt(s as never, 1)).toBeNull();
  });

  it('returns null in an empty paragraph and out of range', () => {
    const s = state([], 1);
    expect(mathRangeAt(s as never, 1)).toBeNull();
    expect(mathRangeAt(s as never, 99)).toBeNull();
  });
});
