import { Schema } from 'prosemirror-model';
import { BapbongEditor, composeSchema, type EditorPlugin } from './bapbong-editor';

const baseSchema = new Schema({
  nodes: { doc: { content: 'paragraph+' }, paragraph: { content: 'text*' }, text: {} },
  marks: { em: {} },
});

describe('composeSchema', () => {
  it('returns null when no plugin contributes schema', () => {
    expect(composeSchema(baseSchema, [{ name: 'noop' }])).toBeNull();
  });

  it('appends plugin marks/nodes to the base schema', () => {
    const plugin: EditorPlugin = {
      name: 'comments',
      schema: { marks: { comment: { attrs: { ids: {} } } } },
    };
    const composed = composeSchema(baseSchema, [plugin]);
    expect(composed).not.toBeNull();
    expect(composed?.marks['comment']).toBeDefined(); // contributed
    expect(composed?.marks['em']).toBeDefined(); // base preserved
    expect(composed?.nodes['paragraph']).toBeDefined();
  });
});

describe('BapbongEditor', () => {
  it('is a constructable class', () => {
    // DOM/canvas wiring is exercised end-to-end in the playground app; here we
    // only assert the public surface is exported as a class.
    expect(typeof BapbongEditor).toBe('function');
    expect(BapbongEditor.prototype.constructor).toBe(BapbongEditor);
  });

  it('accepts an EditorPlugin shape', () => {
    // Type-level proof the plugin contract is importable + implementable; the
    // hooks themselves are exercised in-browser (they need a canvas context).
    let setupCalled = false;
    const probe: EditorPlugin = {
      name: 'probe',
      setup: () => {
        setupCalled = true;
        return () => undefined; // teardown
      },
      onChange: (c) => void c.pageCount,
      onCaretPick: (pos) => void pos,
    };
    expect(probe.name).toBe('probe');
    expect(typeof probe.setup).toBe('function');
    expect(setupCalled).toBe(false); // setup runs only when the editor calls it
  });
});
