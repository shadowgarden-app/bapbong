/** Property names of T whose value can serve as a key (string | number | symbol). */
type IdKeysOf<T> = { [K in keyof T]: T[K] extends PropertyKey ? K : never }[keyof T];

/** Default key property: `id` when T has one, otherwise any valid id property. */
type DefaultIdKey<T> = 'id' extends IdKeysOf<T> ? 'id' : IdKeysOf<T>;

/**
 * A key-property-keyed, insertion-ordered collection. Items are keyed by a
 * property whose value is a primitive (`string | number | symbol`); it defaults
 * to `id`, or pass `{ idProperty }` to key by another (e.g. plugins by `name`).
 * Lookups/removals accept either the key or the item; iterating yields the
 * items in insertion order.
 *
 * Dependency-free + isomorphic. Used by the editor as its plugin registry
 * (`new Collection<EditorPlugin>([], { idProperty: 'name' })`), but generic.
 *
 * Note: if T has no `id` property you must pass `{ idProperty }` — otherwise
 * items would key by an absent property at runtime.
 */
export class Collection<T extends object, IdKey extends IdKeysOf<T> = DefaultIdKey<T>>
  implements Iterable<T>
{
  private readonly items = new Map<T[IdKey], T>();
  private readonly idProperty: IdKey;
  private readonly defaulted: boolean;

  constructor(initial: T[] = [], options?: { idProperty?: IdKey }) {
    this.defaulted = options?.idProperty === undefined;
    this.idProperty = (options?.idProperty ?? 'id') as IdKey;
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

  /** Add (or replace, by key) an item. Throws if the item lacks its key
   *  property (e.g. defaulting to `id` on items that have none). */
  add(value: T): this {
    const key = value[this.idProperty];
    if ((key as unknown) === undefined || (key as unknown) === null) {
      throw new Error(
        `Collection: item has no "${String(this.idProperty)}" key` +
          (this.defaulted ? ' — pass { idProperty } to key by another property.' : '.'),
      );
    }
    this.items.set(key, value);
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
