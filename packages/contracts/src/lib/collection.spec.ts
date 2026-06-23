import { Collection } from './collection';

interface Plugin {
  name: string;
  v: number;
}
interface Widget {
  id: number;
  label: string;
}

describe('Collection', () => {
  it('keys by a custom idProperty (not hardcoded to "name")', () => {
    const c = new Collection<Widget>('id', [
      { id: 1, label: 'a' },
      { id: 2, label: 'b' },
    ]);
    expect(c.get(2)?.label).toBe('b'); // looked up by the `id` property
    expect(c.size).toBe(2);
  });

  it('looks up / removes by key or by item', () => {
    const a = { name: 'a', v: 1 };
    const c = new Collection<Plugin>('name', [a, { name: 'b', v: 2 }]);
    expect(c.get('a')).toBe(a); // by key
    expect(c.get(a)).toBe(a); // by item
    expect(c.has('b')).toBe(true);
    expect(c.remove(a)).toBe(true); // by item
    expect(c.get('a')).toBeUndefined();
    expect(c.remove('missing')).toBe(false);
  });

  it('add replaces by key and keeps insertion order; iterates values', () => {
    const c = new Collection<Plugin>('name');
    c.add({ name: 'x', v: 1 }).add({ name: 'y', v: 2 }).add({ name: 'x', v: 9 }); // replaces x
    expect([...c].map((p) => `${p.name}:${p.v}`)).toEqual(['x:9', 'y:2']);
    expect(c.entries()).toEqual([
      ['x', { name: 'x', v: 9 }],
      ['y', { name: 'y', v: 2 }],
    ]);
  });
});
