import { Collection } from './collection.js';

interface Plugin {
  name: string;
  v: number;
}
interface Widget {
  id: number;
  label: string;
}

describe('Collection', () => {
  it('keys by `id` property by default (no options)', () => {
    const c = new Collection<Widget>([
      { id: 1, label: 'a' },
      { id: 2, label: 'b' },
    ]);
    expect(c.get(2)?.label).toBe('b'); // looked up by the `id` property
    expect(c.size).toBe(2);
  });

  it('keys by a custom idProperty via options (e.g. `name`)', () => {
    const a = { name: 'a', v: 1 };
    const c = new Collection<Plugin>([a, { name: 'b', v: 2 }], {
      idProperty: 'name',
    });
    expect(c.get('a')).toBe(a); // by key
    expect(c.get(a)).toBe(a); // by item
    expect(c.has('b')).toBe(true);
    expect(c.remove(a)).toBe(true); // by item
    expect(c.get('a')).toBeUndefined();
    expect(c.remove('missing')).toBe(false);
  });

  it('throws when an item lacks the key property (default id, item has none)', () => {
    // Plugin has no `id`; with no { idProperty } it defaults to "id" → throws.
    expect(() => new Collection<Plugin>([{ name: 'x', v: 1 }])).toThrow(
      /no "id" key/,
    );
    // Passing the right idProperty is fine.
    expect(
      () =>
        new Collection<Plugin>([{ name: 'x', v: 1 }], { idProperty: 'name' }),
    ).not.toThrow();
  });

  it('add replaces by key and keeps insertion order; iterates values', () => {
    const c = new Collection<Plugin>([], { idProperty: 'name' });
    c.add({ name: 'x', v: 1 })
      .add({ name: 'y', v: 2 })
      .add({ name: 'x', v: 9 }); // replaces x
    expect([...c].map((p) => `${p.name}:${p.v}`)).toEqual(['x:9', 'y:2']);
    expect(c.entries()).toEqual([
      ['x', { name: 'x', v: 9 }],
      ['y', { name: 'y', v: 2 }],
    ]);
  });
});
