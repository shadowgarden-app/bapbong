import { EditorState, TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
// Everything below comes through the façade — this is exactly the surface a
// backend installs (`@shadow-garden/bapbong-headless`). If a re-export drops a
// name (e.g. an ambiguous `export *`), this import fails and the test breaks.
import {
  schema,
  importDocx,
  exportDocx,
  toggleMarkCommand,
  setAlign,
  activeAlign,
  isMarkActive,
  defaultCommands,
  type Command,
} from './index.js';

/** Apply a command (capturing its transaction) and return the next state. */
function apply(state: EditorState, cmd: Command): EditorState {
  let next = state;
  cmd.run(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

/** The text run carrying `text` in `doc`, with its mark names. */
function runFor(doc: PMNode, text: string): string[] | undefined {
  let marks: string[] | undefined;
  doc.descendants((node) => {
    if (node.isText && node.text === text) marks = node.marks.map((m) => m.type.name);
  });
  return marks;
}

describe('bapbong-headless — backend round-trip (no DOM)', () => {
  it('exposes the isomorphic surface through the façade', () => {
    expect(typeof importDocx).toBe('function');
    expect(typeof exportDocx).toBe('function');
    expect(typeof toggleMarkCommand).toBe('function');
    expect(typeof setAlign).toBe('function');
    expect(schema.nodes['paragraph']).toBeTruthy();
    expect(defaultCommands().get('bold')).toBeTruthy();
  });

  it('runs entirely headless — no browser globals are touched', () => {
    // The whole tier is isomorphic; under Node these globals simply don't exist.
    // A regression that reaches for the DOM would throw long before here.
    expect(typeof document).toBe('undefined');
    expect(typeof window).toBe('undefined');
  });

  it('import → edit via commands on EditorState → export survives a round-trip', async () => {
    // 1. Author a doc with the real model schema (as a backend would build one).
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('Hello world')]),
      schema.node('paragraph', null, [schema.text('second line')]),
    ]);

    // 2. Export to a real .docx package from scratch (no carry), then re-import —
    //    this is the headless I/O path with zero browser deps.
    const firstBytes = await exportDocx(doc);
    expect(firstBytes).toBeInstanceOf(Uint8Array);
    const imported = await importDocx(firstBytes);
    expect(imported.doc.childCount).toBe(2);

    // 3. Drive the SAME commands the editor UI uses, on a Node EditorState.
    let state = EditorState.create({ doc: imported.doc });

    // Bold the whole first paragraph.
    const p1Start = 1;
    const p1End = p1Start + imported.doc.child(0).content.size;
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, p1Start, p1End)));
    expect(isMarkActive(state, 'strong')).toBe(false);
    state = apply(state, toggleMarkCommand('bold', 'strong'));
    expect(isMarkActive(state, 'strong')).toBe(true);

    // Centre the second paragraph.
    const p2Start = p1End + 2; // skip the paragraph boundary tokens
    const p2End = p2Start + imported.doc.child(1).content.size;
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, p2Start, p2End)));
    state = apply(state, setAlign('center'));
    expect(activeAlign(state)).toBe('center');

    // 4. Export the edited doc (carry the source package so unmodelled parts
    //    survive) and re-import — the edits must persist through the format.
    const editedBytes = await exportDocx(state.doc, { carry: imported.raw });
    const back = await importDocx(editedBytes);

    expect(back.doc.child(0).textContent).toBe('Hello world');
    expect(runFor(back.doc, 'Hello world')).toContain('strong');
    expect(back.doc.child(1).attrs['align']).toBe('center');
  });
});
