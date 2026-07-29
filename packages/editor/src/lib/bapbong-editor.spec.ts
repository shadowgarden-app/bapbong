import { Schema } from 'prosemirror-model';
import {
  BapbongEditor,
  composeSchema,
  orderPluginsByUses,
  pluginLookupFor,
  type EditorPlugin,
} from './bapbong-editor';

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

describe('orderPluginsByUses', () => {
  const P = (name: string, uses?: string[]): EditorPlugin => ({ name, ...(uses ? { uses } : {}) });

  it('puts dependencies before dependents', () => {
    const order = orderPluginsByUses([P('c', ['b']), P('a'), P('b', ['a'])]);
    expect(order.map((p) => p.name)).toEqual(['a', 'b', 'c']);
  });

  it('leaves independent plugins in their given order', () => {
    // Most plugins declare nothing; registration order must stay meaningful
    // (the pointer hook depends on it).
    const order = orderPluginsByUses([P('x'), P('y'), P('z')]);
    expect(order.map((p) => p.name)).toEqual(['x', 'y', 'z']);
  });

  it('rejects a dependency that is not registered — at registration', () => {
    expect(() => orderPluginsByUses([P('a', ['ghost'])])).toThrow(/not registered/);
  });

  it('rejects a dependency cycle', () => {
    expect(() => orderPluginsByUses([P('a', ['b']), P('b', ['a'])])).toThrow(/cycle/);
  });
});

describe('pluginLookupFor (the `uses` gate)', () => {
  const P = (name: string, uses?: string[]): EditorPlugin => ({ name, ...(uses ? { uses } : {}) });
  const find = P('find');
  const table = P('table-selection');
  const registry = new Map([['find', find], ['table-selection', table]]);
  const resolve = (n: string) => registry.get(n) ?? null;

  it('hands back a plugin the caller declared', () => {
    const lookup = pluginLookupFor(P('a', ['find']), resolve);
    expect(lookup('find')).toBe(find);
  });

  it('refuses a plugin that is registered but NOT declared', () => {
    // The whole point: reachable is not the same as allowed. Silent access
    // would let a hidden dependency exist that setup ordering never honours.
    const lookup = pluginLookupFor(P('a', ['find']), resolve);
    expect(() => lookup('table-selection')).toThrow(/without declaring it/);
  });

  it('refuses everything when the caller declares nothing', () => {
    const lookup = pluginLookupFor(P('a'), resolve);
    expect(() => lookup('find')).toThrow(/without declaring it/);
  });

  it('reports a declared-but-missing dependency distinctly', () => {
    const lookup = pluginLookupFor(P('a', ['ghost']), resolve);
    expect(() => lookup('ghost')).toThrow(/not registered/);
  });
});
