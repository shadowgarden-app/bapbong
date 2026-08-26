import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import {
  autoCorrectPlugin,
  backspaceOutdent,
  createEditingState,
  moveCaretCommand,
  shiftListLevel,
  splitListItem,
  splitParagraphKeepFormat,
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

const doc = schema.node('doc', null, [
  schema.node('paragraph', null, [schema.text('hello')]),
]);

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
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 3)),
    );
    state = state.apply(state.tr.insertText('xy', 1));
    expect(state.selection.head).toBe(5); // shifted by the 2 inserted chars
  });
});

describe('moveCaretCommand', () => {
  it('collapses by default and extends from the anchor with extend=true', () => {
    let state = createEditingState(doc); // 'hello'
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 2)),
    );

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
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: { list: { default: null } },
      },
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
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 5)),
    ); // end of "item"
    const handled = splitListItem(state, (tr) => (state = state.apply(tr)));
    expect(handled).toBe(true);
    expect(state.doc.childCount).toBe(2);
    expect(state.doc.child(1).attrs['list']).toEqual(listAttrs.list);
  });

  it('exits the list on an empty item, and defers outside lists', () => {
    const d = listSchema.node('doc', null, [
      listSchema.node('paragraph', listAttrs),
    ]);
    let state = createEditingState(d);
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 1)),
    );
    expect(splitListItem(state, (tr) => (state = state.apply(tr)))).toBe(true);
    expect(state.doc.child(0).attrs['list']).toBeNull(); // left the list

    const plain = createEditingState(
      listSchema.node('doc', null, [
        listSchema.node('paragraph', null, [listSchema.text('x')]),
      ]),
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
  const listAttrs = {
    list: { numId: '1', level: 0 },
    indent: { left: 48, hanging: 24 },
  };

  it('outdents in steps: drop marker → clear indent → defer to join', () => {
    const d = listSchema.node('doc', null, [
      listSchema.node('paragraph', listAttrs, [listSchema.text('item')]),
    ]);
    let state = createEditingState(d);
    const caretAtStart = () =>
      (state = state.apply(
        state.tr.setSelection(TextSelection.create(state.doc, 1)),
      ));

    // Step 1: drops the marker, keeps the indent, text stays on the same line.
    caretAtStart();
    expect(backspaceOutdent(state, (tr) => (state = state.apply(tr)))).toBe(
      true,
    );
    expect(state.doc.child(0).attrs['list']).toBeNull();
    expect(state.doc.child(0).attrs['indent']).toEqual(listAttrs.indent);
    expect(state.doc.child(0).textContent).toBe('item');

    // Step 2: clears the indent (caret returns to the margin).
    caretAtStart();
    expect(backspaceOutdent(state, (tr) => (state = state.apply(tr)))).toBe(
      true,
    );
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
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 2)),
    );

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
    ranged = ranged.apply(
      ranged.tr.setSelection(TextSelection.create(ranged.doc, 1, 3)),
    );
    expect(backspaceOutdent(ranged, () => undefined)).toBe(false); // deletes the range
  });
});

describe('wordRangeAt', () => {
  const para = (text: string) =>
    schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text(text)]),
    ]);

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

// ── The paragraph mark in the editing loop ──────────────────────────
describe('paragraph mark (¶) behaviour', () => {
  const pSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: {
        group: 'block',
        content: 'inline*',
        attrs: {
          align: { default: null },
          heading: { default: null },
          bookmarks: { default: null },
          markFont: { default: null },
        },
      },
      text: { group: 'inline' },
    },
    marks: {
      strong: {},
      em: {},
      fontSize: { attrs: { size: {} } },
      fontFamily: { attrs: { family: {} } },
    },
  });
  const mark = { family: 'Times New Roman', sizePt: 8, bold: true };
  const p = (attrs: object | null, text?: string) =>
    pSchema.node('paragraph', attrs, text ? [pSchema.text(text)] : []);
  const at = (state: ReturnType<typeof createEditingState>, pos: number) =>
    state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));

  it('Enter at the end keeps the formatting and the ¶ font on the new paragraph', () => {
    const d = pSchema.node('doc', null, [
      p({ align: 'center', markFont: mark, bookmarks: ['x'] }, 'ab'),
    ]);
    let state = at(createEditingState(d), 3); // after "ab"
    expect(
      splitParagraphKeepFormat(state, (tr) => (state = state.apply(tr))),
    ).toBe(true);
    expect(state.doc.childCount).toBe(2);
    const next = state.doc.child(1);
    expect(next.attrs['align']).toBe('center');
    expect(next.attrs['markFont']).toEqual(mark);
    expect(next.attrs['bookmarks']).toBeNull(); // anchors stay where they were
    expect(state.doc.child(0).attrs['bookmarks']).toEqual(['x']);
  });

  it('Enter mid-paragraph and inside a heading defer to the base split', () => {
    const d = pSchema.node('doc', null, [p({ markFont: mark }, 'ab')]);
    expect(splitParagraphKeepFormat(at(createEditingState(d), 2))).toBe(false);
    // A heading at its end: the new paragraph is a plain one (heading null).
    const h = pSchema.node('doc', null, [
      p({ heading: 1, markFont: mark }, 'Title'),
    ]);
    let state = at(createEditingState(h), 6);
    splitParagraphKeepFormat(state, (tr) => (state = state.apply(tr)));
    expect(state.doc.child(1).attrs['heading']).toBeNull();
    expect(state.doc.child(1).attrs['markFont']).toEqual(mark);
  });

  it('a caret coming to rest in an EMPTY paragraph is primed with the ¶ font', () => {
    const d = pSchema.node('doc', null, [p({ markFont: mark }), p(null, 'cd')]);
    let state = createEditingState(d);
    // Any transaction that lands the caret in the empty paragraph seeds it.
    state = at(state, 1);
    const names = (state.storedMarks ?? []).map(
      (m) => `${m.type.name}${JSON.stringify(m.attrs)}`,
    );
    expect(names).toEqual([
      'fontFamily{"family":"Times New Roman"}',
      'fontSize{"size":8}',
      'strong{}',
    ]);
    // Typing then carries those marks — 8pt bold Times, as Word would.
    state = state.apply(state.tr.insertText('x'));
    const typed = state.doc.child(0).firstChild!;
    expect(typed.marks.map((m) => m.type.name)).toEqual([
      'fontFamily',
      'fontSize',
      'strong',
    ]);
    // A full paragraph offers its own marks; nothing is seeded there.
    state = at(state, 4);
    expect(state.storedMarks).toBeNull();
    // An explicit (empty) stored-mark set — bold toggled off at the caret —
    // is respected, not overwritten.
    const cleared = createEditingState(d);
    const s2 = cleared.apply(
      cleared.tr
        .setSelection(TextSelection.create(cleared.doc, 1))
        .setStoredMarks([]),
    );
    expect(s2.storedMarks).toEqual([]);
  });
});

describe('autoCorrectPlugin (math branch)', () => {
  const mathSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*' },
      text: { group: 'inline' },
    },
    marks: { math: {} },
  });
  const math = mathSchema.marks['math'].create();

  /** A state whose paragraph holds plain `lead` then math-marked `eq`,
   *  caret at the end, plus a mock view for the plugin prop. */
  const setup = (lead: string, eq: string) => {
    const children = [
      ...(lead ? [mathSchema.text(lead)] : []),
      ...(eq ? [mathSchema.text(eq, [math])] : []),
    ];
    const doc = mathSchema.node('doc', null, [
      mathSchema.node('paragraph', null, children),
    ]);
    let state = EditorState.create({ doc, plugins: [autoCorrectPlugin()] });
    const caret = 1 + lead.length + eq.length;
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, caret)),
    );
    let applied = state;
    const view = {
      state,
      composing: false,
      dispatch(tr: Parameters<typeof state.apply>[0]) {
        applied = state.apply(tr);
      },
    };
    const input = (text: string) =>
      state.plugins[0].props.handleTextInput!.call(
        state.plugins[0],
        view as never,
        caret,
        caret,
        text,
        () => state.tr,
      );
    return { input, text: () => applied.doc.textContent };
  };

  it('converts \\name on the trigger space, consuming the space', () => {
    const s = setup('x = ', '\\omega');
    expect(s.input(' ')).toBe(true);
    expect(s.text()).toBe('x = ω');
  });

  it('handles Chromium delivering the space merged with a neighbour', () => {
    // One keystroke arrives as "a " — the \name straddles doc and input.
    const s = setup('', '\\omeg');
    expect(s.input('a ')).toBe(true);
    expect(s.text()).toBe('ω');
  });

  it('handles a fully-typed name in one input, keeping any prefix', () => {
    const s = setup('', '\\x');
    // "\x" in the doc is dead; typed remainder completes nothing — no-op…
    expect(s.input('y ')).toBe(false);
    // …while a complete name inside the input converts, prefix kept.
    const s2 = setup('', 'k');
    expect(s2.input('\\pi ')).toBe(true);
    expect(s2.text()).toBe('kπ');
  });

  it('stays quiet outside math, on unknown names, and mid-word input', () => {
    const plain = setup('\\omega', '');
    expect(plain.input(' ')).toBe(false);
    const unknown = setup('', '\\banana');
    expect(unknown.input(' ')).toBe(false);
    const midword = setup('', '\\omega');
    expect(midword.input(' q')).toBe(false);
  });
});

describe('autoCorrectPlugin (NBSP trigger)', () => {
  it('accepts the NBSP contenteditable writes for a trailing space', () => {
    const mathSchema = new Schema({
      nodes: {
        doc: { content: 'block+' },
        paragraph: { group: 'block', content: 'inline*' },
        text: { group: 'inline' },
      },
      marks: { math: {} },
    });
    const math = mathSchema.marks['math'].create();
    const doc = mathSchema.node('doc', null, [
      mathSchema.node('paragraph', null, [mathSchema.text('\\tau', [math])]),
    ]);
    let state = EditorState.create({ doc, plugins: [autoCorrectPlugin()] });
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 5)),
    );
    let applied = state;
    const view = {
      state,
      composing: false,
      dispatch(tr: Parameters<typeof state.apply>[0]) {
        applied = state.apply(tr);
      },
    };
    const handled = state.plugins[0].props.handleTextInput!.call(
      state.plugins[0],
      view as never,
      5,
      5,
      ' ',
      () => state.tr,
    );
    expect(handled).toBe(true);
    expect(applied.doc.textContent).toBe('τ');
  });
});
