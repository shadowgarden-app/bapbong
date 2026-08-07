import type {
  Collection,
  Command,
  Dispatch,
  EditorChange,
} from '@shadow-garden/bapbong-contracts';

/** The editor `run` state type, derived from the {@link Command} contract so
 *  this package needs no direct ProseMirror dependency. */
export type EditorStateOf = Parameters<Command['run']>[0];

/**
 * The minimal editor surface every bapbong-ui widget binds to. `BapbongEditor`
 * satisfies it structurally, so this package never imports the editor — both
 * just speak the `contracts` vocabulary (the same decoupling as the plugin
 * contract). A host framework only supplies the element + this handle.
 */
export interface EditorHandle {
  /** Named command registry the UI renders + dispatches against. */
  readonly commands: Collection<Command>;
  /** Current document + selection (throws before a document is loaded). */
  readonly state: EditorStateOf;
  /** Apply a transaction. */
  dispatch: Dispatch;
  /** Return focus to the editor (after a control takes a click). */
  focus(): void;
  /** Subscribe to editor cycles; returns an unsubscribe. Drives active state. */
  onChange(cb: (c: EditorChange) => void): () => void;
}

/** Inject a stylesheet once, keyed by `id` (idempotent across mounts). */
export function injectStyle(id: string, css: string): void {
  if (document.getElementById(id)) return;
  const el = document.createElement('style');
  el.id = id;
  el.textContent = css;
  document.head.appendChild(el);
}

/** A row of colour swatches plus the OS colour picker.
 *
 *  A dropdown of named colours is the wrong control for a colour: it hides the
 *  choices behind a click and can only ever offer the handful someone thought
 *  to name. A row shows every preset at once, and the trailing
 *  `input[type=color]` hands off to the platform picker for anything else —
 *  the real thing rather than a hex field imitating one.
 *
 *  Styling is the caller's `prefix`, so a widget keeps its own class names.
 *  Returns a `sync` to call whenever the value changes outside the row. */
export interface SwatchRow {
  el: HTMLElement;
  sync(): void;
}

export function swatchRow(opts: {
  prefix: string;
  presets: readonly string[];
  /** Label + title for the "no colour" chip; omit to drop that chip. */
  clearLabel?: string;
  get(): string | null;
  set(color: string | null): void;
}): SwatchRow {
  const { prefix, presets, clearLabel, get, set } = opts;
  const el = document.createElement('div');
  el.className = `${prefix}-swatches`;
  const chips: Array<{ node: HTMLElement; color: string | null }> = [];

  const add = (node: HTMLElement, color: string | null) => {
    node.addEventListener('click', () => {
      set(color);
      sync();
    });
    el.append(node);
    chips.push({ node, color });
  };

  if (clearLabel !== undefined) {
    const none = document.createElement('button');
    none.type = 'button';
    none.className = `${prefix}-swatch ${prefix}-swatch-none`;
    none.title = clearLabel;
    none.textContent = '⦸';
    add(none, null);
  }
  for (const color of presets) {
    const sw = document.createElement('button');
    sw.type = 'button';
    sw.className = `${prefix}-swatch`;
    sw.style.background = color;
    sw.title = color;
    add(sw, color);
  }

  // The platform picker, for everything the presets do not cover.
  const custom = document.createElement('input');
  custom.type = 'color';
  custom.className = `${prefix}-swatch ${prefix}-swatch-custom`;
  custom.title = 'Custom colour…';
  custom.addEventListener('input', () => {
    set(custom.value.toUpperCase());
    sync();
  });
  el.append(custom);

  function sync(): void {
    const value = get();
    for (const { node, color } of chips)
      node.classList.toggle('on', color === value);
    // Keep the picker showing the live colour, but never steal a value it is
    // mid-edit — that would fight the user dragging inside the OS panel.
    if (document.activeElement !== custom && value) custom.value = value;
    custom.classList.toggle('on', !!value && !presets.some((p) => p === value));
  }
  sync();
  return { el, sync };
}
