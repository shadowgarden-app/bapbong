import { Schema, type Node as PMNode } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import { fieldAt } from '@shadow-garden/bapbong-model';
import { tocPlugin } from './toc-plugin';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: { field: { default: null }, bookmarks: { default: null } },
    },
    text: { group: 'inline' },
  },
  marks: { link: { attrs: { href: {} } } },
});

const entry = (field: object, text: string, anchor: string) =>
  schema.node('paragraph', { field }, [
    schema.text(text, [schema.marks['link'].create({ href: `#${anchor}` })]),
  ]);
const heading = (text: string, anchor: string) =>
  schema.node('paragraph', { bookmarks: [anchor] }, [schema.text(text)]);

/** The plugin over `doc`, with each bookmarked heading laid out on the page
 *  `pages[anchor]` (null = not laid out yet). */
function run(doc: PMNode, pages: Record<string, number | null>) {
  let state = EditorState.create({ schema, doc });
  const toc = tocPlugin();
  toc.setup?.({
    get state() {
      return state;
    },
    dispatch: (tr: Transaction) => {
      state = state.apply(tr);
    },
    caretRect: (pos: number) => {
      const name = (state.doc.resolve(pos).parent.attrs['bookmarks'] ??
        [])[0] as string | undefined;
      const page = name ? pages[name] : null;
      return page == null ? null : { pageIndex: page };
    },
    layout: null,
  } as never);
  const changed = toc.updatePageNumbers(1);
  return { changed, doc: state.doc };
}

const field = { kind: 'toc', instr: 'TOC \\o "1-3" \\h \\z \\u', dirty: true };
const doc = () =>
  schema.node('doc', null, [
    entry(field, 'Intro\t1', 'a'),
    entry(field, 'Method\t1', 'b'),
    heading('Intro', 'a'),
    heading('Method', 'b'),
  ]);

describe('toc updatePageNumbers', () => {
  it('fills the numbers and clears dirty, keeping the span one field', () => {
    const { changed, doc: out } = run(doc(), { a: 1, b: 3 });
    expect(changed).toBe(2);
    expect(out.child(0).textContent).toBe('Intro\t2');
    expect(out.child(1).textContent).toBe('Method\t4');
    const f0 = out.child(0).attrs['field'];
    expect(f0).toMatchObject({ kind: 'toc', dirty: false });
    // Same object on every entry — fieldAt groups the span by identity.
    expect(out.child(1).attrs['field']).toBe(f0);
    expect(fieldAt(out, 1)).toMatchObject({ from: 0 });
    expect(fieldAt(out, 1)?.to).toBe(
      out.child(0).nodeSize + out.child(1).nodeSize,
    );
  });

  it('clears dirty even when no number changed', () => {
    const { changed, doc: out } = run(doc(), { a: 0, b: 0 });
    expect(changed).toBe(0);
    expect(out.child(0).attrs['field'].dirty).toBe(false);
  });

  it('stays dirty while any entry has no laid-out page', () => {
    const { doc: out } = run(doc(), { a: 1, b: null });
    expect(out.child(0).textContent).toBe('Intro\t2');
    expect(out.child(0).attrs['field'].dirty).toBe(true);
  });

  it('reads a split "\\t12" run and leaves a number with no tab alone', () => {
    const link = (a: string) => [
      schema.marks['link'].create({ href: `#${a}` }),
    ];
    const split = schema.node('doc', null, [
      schema.node('paragraph', { field }, [
        schema.text('Intro', link('a')),
        schema.text('\t1', link('a')),
      ]),
      schema.node('paragraph', { field }, [
        schema.text('Article 12', link('b')),
      ]),
      heading('Intro', 'a'),
      heading('Article 12', 'b'),
    ]);
    const { doc: out } = run(split, { a: 4, b: 4 });
    expect(out.child(0).textContent).toBe('Intro\t5');
    expect(out.child(1).textContent).toBe('Article 12');
  });
});
