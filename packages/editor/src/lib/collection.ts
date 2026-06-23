/**
 * A small name-keyed, insertion-ordered collection. Items must carry a unique
 * `name`; that name is the key, so lookups/removals accept either the name or
 * the item itself. Iterating yields the items (values) in insertion order.
 *
 * Used by the editor as its plugin registry (built-in + host plugins keyed by
 * `EditorPlugin.name`), but generic over anything with a `name`.
 */
export class Collection<T extends { readonly name: string }> implements Iterable<T> {
  private readonly items = new Map<string, T>();

  constructor(initial: Iterable<T> = []) {
    for (const it of initial) this.add(it);
  }

  private keyOf(keyOrValue: string | T): string {
    return typeof keyOrValue === 'string' ? keyOrValue : keyOrValue.name;
  }

  /** The item named `key` (or matching the given item), or undefined. */
  get(keyOrValue: string | T): T | undefined {
    return this.items.get(this.keyOf(keyOrValue));
  }

  /** Add (or replace, by name) an item. Returns this for chaining. */
  add(value: T): this {
    this.items.set(value.name, value);
    return this;
  }

  /** Remove by name (or by item). Returns whether something was removed. */
  remove(keyOrValue: string | T): boolean {
    return this.items.delete(this.keyOf(keyOrValue));
  }

  has(keyOrValue: string | T): boolean {
    return this.items.has(this.keyOf(keyOrValue));
  }

  get size(): number {
    return this.items.size;
  }

  /** [name, item] pairs, in insertion order. */
  entries(): [string, T][] {
    return [...this.items.entries()];
  }

  [Symbol.iterator](): Iterator<T> {
    return this.items.values();
  }
}
