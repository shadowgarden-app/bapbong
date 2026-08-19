/**
 * Keyboard shortcuts as DATA — one registry the editor dispatches from, the
 * menubar labels from and the Keyboard-shortcuts dialog lists from, instead of
 * a keymap here, a window listener there and a hand-typed "⌘X" in the menu.
 *
 * Pure: no DOM. The editor (input-bridge) turns a KeyboardEvent into a name
 * and looks it up; this module owns the name grammar and its display form.
 */

/**
 * One binding: a key chord → a {@link Command} by name.
 *
 * `key` uses ProseMirror's keymap grammar: modifiers `Shift-`, `Alt-`,
 * `Ctrl-`, `Meta-` and the portable `Mod-` (⌘ on macOS, Ctrl elsewhere),
 * then a key name — a single character (`b`, `/`, `\\`) or a
 * `KeyboardEvent.key` name (`Enter`, `Tab`, `ArrowUp`, `F2`). Order of the
 * modifiers does not matter; the registry keys by the normalized form (see
 * {@link normalizeKey}), so `Shift-Mod-z` and `Mod-Shift-z` are the same
 * binding and a later `add` replaces an earlier one — that is how a host
 * overrides a core binding, and the dialog then shows the host's.
 */
export interface Keybinding {
  key: string;
  /** The command's `name` in the editor's command registry. */
  command: string;
  /**
   * `editor` (default): the chord is the editor's, handled by its keymap while
   * it has focus — the editor's own registry. `window`: an app-wide chord
   * (⌘S, ⌘F, ⌘W) handled wherever focus is, even with no document open — the
   * HOST's registry, dispatched by a document listener (input-bridge
   * `installWindowKeymap`); such a chord must carry a Mod/Ctrl/Alt modifier,
   * or it would swallow ordinary typing in the app's other inputs.
   */
  scope?: 'editor' | 'window';
  /** When the binding applies — a short phrase for the dialog's "When" column
   *  ("in a list", "hex digits before caret"). Display only. */
  when?: string;
  /** Who registered it: `core`, `plugin:<name>`, or the host app's name.
   *  Shown in the dialog; also what lets a host list or clear its own. */
  source: string;
}

const MODIFIERS = new Set([
  'shift',
  'alt',
  'ctrl',
  'control',
  'meta',
  'cmd',
  'mod',
]);

/**
 * The canonical form of a key chord for the current platform, in the exact
 * shape ProseMirror's keymap produces from a KeyboardEvent:
 * `[Alt-][Ctrl-][Meta-][Shift-]<name>`, `Mod` resolved to Meta (mac) or Ctrl,
 * a one-character name lower-cased when Shift is not part of the chord.
 * `mac` is a parameter so the same code normalizes for either platform (and
 * tests both).
 */
export function normalizeKey(key: string, mac: boolean): string {
  const parts = key.split(/-(?!$)/);
  let name = parts[parts.length - 1];
  if (name === 'Space') name = ' ';
  let alt = false;
  let ctrl = false;
  let shift = false;
  let meta = false;
  for (let i = 0; i < parts.length - 1; i++) {
    const m = parts[i].toLowerCase();
    if (m === 'shift') shift = true;
    else if (m === 'alt') alt = true;
    else if (m === 'ctrl' || m === 'control') ctrl = true;
    else if (m === 'meta' || m === 'cmd') meta = true;
    else if (m === 'mod') {
      if (mac) meta = true;
      else ctrl = true;
    } else throw new RangeError(`Unrecognized modifier name: ${parts[i]}`);
  }
  // A single character is matched by its BASE key, lower-case: without Shift
  // that is what the event reports; with Shift the keymap falls back to the
  // keycode's base character (also lower-case) — so `Shift-Mod-Z` and
  // `Shift-Mod-z` must land on the same entry.
  if (name.length === 1) name = name.toLowerCase();
  return (
    (alt ? 'Alt-' : '') +
    (ctrl ? 'Ctrl-' : '') +
    (meta ? 'Meta-' : '') +
    (shift ? 'Shift-' : '') +
    name
  );
}

/** Whether a chord carries a Mod/Ctrl/Meta/Alt modifier — what a
 *  window-scope binding needs so it cannot swallow plain typing. */
export function hasCommandModifier(key: string): boolean {
  return key
    .split(/-(?!$)/)
    .slice(0, -1)
    .some((m) =>
      ['mod', 'ctrl', 'control', 'meta', 'cmd', 'alt'].includes(
        m.toLowerCase(),
      ),
    );
}

/** Whether every part of a chord is a modifier or a key — a cheap sanity check
 *  for registration (typos like `Mod+b`). */
export function isValidKey(key: string): boolean {
  const parts = key.split(/-(?!$)/);
  const name = parts[parts.length - 1];
  if (!name) return false;
  // A key name is one character, or a KeyboardEvent.key word (Enter, ArrowUp,
  // F2) — anything else ("Mod+z", "ctrl b") is a typo in the chord.
  if (name.length > 1 && !/^[A-Za-z][A-Za-z0-9]*$/.test(name)) return false;
  return parts.slice(0, -1).every((m) => MODIFIERS.has(m.toLowerCase()));
}

/** Display names of keys that are not their `KeyboardEvent.key`. */
const KEY_LABELS: Record<string, { mac: string; other: string }> = {
  ' ': { mac: 'Space', other: 'Space' },
  Space: { mac: 'Space', other: 'Space' },
  Enter: { mac: '↩', other: 'Enter' },
  Backspace: { mac: '⌫', other: 'Backspace' },
  Delete: { mac: '⌦', other: 'Delete' },
  Escape: { mac: 'Esc', other: 'Esc' },
  Tab: { mac: '⇥', other: 'Tab' },
  ArrowUp: { mac: '↑', other: '↑' },
  ArrowDown: { mac: '↓', other: '↓' },
  ArrowLeft: { mac: '←', other: '←' },
  ArrowRight: { mac: '→', other: '→' },
  Home: { mac: '↖', other: 'Home' },
  End: { mac: '↘', other: 'End' },
  PageUp: { mac: '⇞', other: 'PgUp' },
  PageDown: { mac: '⇟', other: 'PgDn' },
};

/**
 * The chord as the pieces a `<kbd>` row shows, in the platform's own order and
 * glyphs: mac `⌃ ⌥ ⇧ ⌘ Z` (Apple's order), elsewhere `Ctrl Alt Shift Z`.
 * `Mod` becomes ⌘ / Ctrl. Letters are upper-cased.
 */
export function formatKey(key: string, mac: boolean): string[] {
  const parts = key.split(/-(?!$)/);
  const name = parts[parts.length - 1];
  const mods = new Set(parts.slice(0, -1).map((m) => m.toLowerCase()));
  const has = (...names: string[]) => names.some((n) => mods.has(n));
  const out: string[] = [];
  if (mac) {
    if (has('ctrl', 'control')) out.push('⌃');
    if (has('alt')) out.push('⌥');
    if (has('shift')) out.push('⇧');
    if (has('mod', 'meta', 'cmd')) out.push('⌘');
  } else {
    if (has('mod', 'ctrl', 'control')) out.push('Ctrl');
    if (has('alt')) out.push('Alt');
    if (has('shift')) out.push('Shift');
    if (has('meta', 'cmd')) out.push('Win');
  }
  const label = KEY_LABELS[name];
  out.push(
    label
      ? mac
        ? label.mac
        : label.other
      : name.length === 1
        ? name.toUpperCase()
        : name,
  );
  return out;
}

/** `formatKey` joined the way a menu label reads: `⇧⌘Z` on mac, `Ctrl+Shift+Z`
 *  elsewhere. */
export function keyLabel(key: string, mac: boolean): string {
  const parts = formatKey(key, mac);
  return mac ? parts.join('') : parts.join('+');
}

/** Rows for a shortcuts table: one per COMMAND, its keys gathered, sorted by
 *  the command's title (falling back to its name), case-insensitively. */
export interface KeybindingRow {
  command: string;
  title: string;
  keys: string[];
  when?: string;
  source: string;
  scope: 'editor' | 'window';
}

export function keybindingRows(
  bindings: Iterable<Keybinding>,
  titleOf: (command: string) => string | undefined,
): KeybindingRow[] {
  const byCommand = new Map<string, KeybindingRow>();
  for (const b of bindings) {
    const row = byCommand.get(b.command);
    if (row) {
      row.keys.push(b.key);
      // The first binding's when/source stand for the row; a second key for
      // the same command normally comes from the same registrar.
      continue;
    }
    byCommand.set(b.command, {
      command: b.command,
      title: titleOf(b.command) ?? b.command,
      keys: [b.key],
      when: b.when,
      source: b.source,
      scope: b.scope ?? 'editor',
    });
  }
  return [...byCommand.values()].sort((a, b) =>
    a.title.localeCompare(b.title, undefined, { sensitivity: 'base' }),
  );
}

/** Rows whose command title/name, key label, when or source contain `query`
 *  (case-insensitive); an empty query keeps every row. */
export function filterKeybindingRows(
  rows: readonly KeybindingRow[],
  query: string,
  mac: boolean,
): KeybindingRow[] {
  const q = query.trim().toLowerCase();
  if (!q) return [...rows];
  return rows.filter((r) =>
    [
      r.title,
      r.command,
      r.when ?? '',
      r.source,
      ...r.keys.map((k) => keyLabel(k, mac)),
      ...r.keys,
    ].some((s) => s.toLowerCase().includes(q)),
  );
}

/**
 * The editor's keybinding registry. Keyed by the NORMALIZED chord for the
 * platform it was built for, so `Shift-Mod-z` and `Mod-Shift-z` are one entry
 * and a later `add` replaces an earlier — a host overriding a core binding is
 * exactly that. `get(name)` takes the normalized name a keydown produces (see
 * input-bridge), which is what makes dispatch a single Map lookup. Iterates in
 * insertion order; `bySource` is how a host or plugin finds its own to remove.
 */
export class KeybindingRegistry implements Iterable<Keybinding> {
  private readonly items = new Map<string, Keybinding>();
  private readonly listeners = new Set<() => void>();

  constructor(private readonly mac: boolean) {}

  /** The normalized (registry) name of a chord. */
  nameOf(key: string): string {
    return normalizeKey(key, this.mac);
  }

  /** Add or replace (by chord). Throws on a malformed chord, and on a
   *  window-scope chord without a command modifier — that one would swallow
   *  plain typing everywhere in the app. */
  add(binding: Keybinding): this {
    if (!isValidKey(binding.key))
      throw new RangeError(`Keybinding: malformed key "${binding.key}"`);
    if (binding.scope === 'window' && !hasCommandModifier(binding.key))
      throw new RangeError(
        `Keybinding: window-scope "${binding.key}" needs Mod/Ctrl/Alt`,
      );
    this.items.set(this.nameOf(binding.key), binding);
    this.emit();
    return this;
  }

  /** By chord (raw or normalized). */
  get(key: string): Keybinding | undefined {
    return this.items.get(this.nameOf(key)) ?? this.items.get(key);
  }

  delete(key: string): boolean {
    const hit = this.items.delete(this.nameOf(key)) || this.items.delete(key);
    if (hit) this.emit();
    return hit;
  }

  /** Every binding a registrar owns (`source`), e.g. to remove them on
   *  teardown. */
  bySource(source: string): Keybinding[] {
    return [...this.items.values()].filter((b) => b.source === source);
  }

  /** Remove every binding of one source. */
  clearSource(source: string): void {
    let hit = false;
    for (const [k, b] of this.items)
      if (b.source === source) {
        this.items.delete(k);
        hit = true;
      }
    if (hit) this.emit();
  }

  get size(): number {
    return this.items.size;
  }

  [Symbol.iterator](): Iterator<Keybinding> {
    return this.items.values();
  }

  /** Subscribe to any change (add/delete); returns an unsubscribe. The menubar
   *  relabels and the shortcuts dialog re-renders on this. */
  onChange(cb: () => void): () => void {
    this.listeners.add(cb);
    return () => this.listeners.delete(cb);
  }

  private emit(): void {
    for (const cb of this.listeners) cb();
  }
}
