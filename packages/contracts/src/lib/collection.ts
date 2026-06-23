/** Property names of T whose value can serve as a key (string | number | symbol). */
type IdKeysOf<T> = { [K in keyof T]: T[K] extends PropertyKey ? K : never }[keyof T];

/**
 * A key-property-keyed, insertion-ordered collection. The id property name is
 * supplied at construction — `new Collection('name', items)` keys each item by
 * its `name`, but any primitive-valued property works. Lookups/removals accept
 * either the key or the item itself; iterating yields the items in insertion
 * order.
 *
 * Dependency-free + isomorphic. Used by the editor as its plugin registry
 * (`new Collection<EditorPlugin>('name', …)`), but generic over anything with a
 * primitive id property.
 */
export class Collection<T extends object, IdKey extends IdKeysOf<T> = IdKeysOf<T>>
  implements Iterable<T>
{
  private readonly items = new Map<T[IdKey], T>();

  /** @param idProperty the property whose value is each item's unique key. */
  constructor(private readonly idProperty: IdKey, initial: Iterable<T> = []) {
    for (const it of initial) this.add(it);
  }

  private keyOf(keyOrValue: T[IdKey] | T): T[IdKey] {
    return typeof keyOrValue === 'object' && keyOrValue !== null
      ? (keyOrValue as T)[this.idProperty]
      : (keyOrValue as T[IdKey]);
  }

  /** The item with `key` (or matching the given item), or undefined. */
  get(keyOrValue: T[IdKey] | T): T | undefined {
    return this.items.get(this.keyOf(keyOrValue));
  }

  /** Add (or replace, by key) an item. Returns this for chaining. */
  add(value: T): this {
    this.items.set(value[this.idProperty], value);
    return this;
  }

  /** Remove by key (or by item). Returns whether something was removed. */
  remove(keyOrValue: T[IdKey] | T): boolean {
    return this.items.delete(this.keyOf(keyOrValue));
  }

  has(keyOrValue: T[IdKey] | T): boolean {
    return this.items.has(this.keyOf(keyOrValue));
  }

  get size(): number {
    return this.items.size;
  }

  /** [key, item] pairs, in insertion order. */
  entries(): [T[IdKey], T][] {
    return [...this.items.entries()];
  }

  [Symbol.iterator](): Iterator<T> {
    return this.items.values();
  }
}
