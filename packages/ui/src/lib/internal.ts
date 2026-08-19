import type {
  Collection,
  Command,
  Dispatch,
  EditorChange,
  KeybindingRegistry,
} from '@shadow-garden/bapbong-contracts';
import { keyLabel } from '@shadow-garden/bapbong-contracts';

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
  /** The editor's keybinding registry — what menu rows and toolbar tooltips
   *  label their shortcuts from. Optional so a minimal handle still works. */
  readonly keybindings?: KeybindingRegistry;
}

/** Whether chords should read as ⌘ (the same test the editor's keymap makes
 *  to resolve `Mod`). Read once; the platform does not change under us. */
export const IS_MAC: boolean =
  typeof navigator !== 'undefined'
    ? /Mac|iP(hone|[oa]d)/.test(navigator.platform)
    : false;

/**
 * The display label of a command's first binding across the given registries
 * (`⇧⌘Z` / `Ctrl+Shift+Z`), or undefined. The editor's registry comes first;
 * a host passes its app registry after it for chords like ⌘F that live there.
 */
export function shortcutLabel(
  command: string,
  registries: readonly (KeybindingRegistry | undefined)[],
): string | undefined {
  for (const r of registries) {
    if (!r) continue;
    for (const b of r)
      if (b.command === command) return keyLabel(b.key, IS_MAC);
  }
  return undefined;
}

/** Inject a stylesheet once, keyed by `id` (idempotent across mounts). */
export function injectStyle(id: string, css: string): void {
  if (document.getElementById(id)) return;
  const el = document.createElement('style');
  el.id = id;
  el.textContent = css;
  document.head.appendChild(el);
}

/** A rectangle to hang a floating panel from — a caret box, a button, or any
 *  point with a height. Viewport coordinates. */
export interface FloatAnchor {
  x: number;
  y: number;
  height: number;
}

/** Place `el` below `anchor`, flipping above when the viewport runs out and
 *  clamping horizontally. Shared by the link panel and the colour picker: both
 *  are body-level fixed panels, which is also how they escape the toolbar's
 *  `overflow: hidden` without any container-relative arithmetic. */
export function placeFloating(el: HTMLElement, anchor: FloatAnchor): void {
  const r = el.getBoundingClientRect();
  const pad = 6;
  const gap = 6;
  const x = Math.max(
    pad,
    Math.min(anchor.x, window.innerWidth - r.width - pad),
  );
  const below = anchor.y + anchor.height + gap;
  const above = anchor.y - gap - r.height;
  const y =
    below + r.height <= window.innerHeight - pad || above < pad
      ? Math.min(below, window.innerHeight - r.height - pad)
      : above;
  el.style.left = `${x}px`;
  el.style.top = `${y}px`;
}

/**
 * Arrow-key navigation over a 2-D grid of focusable cells with a roving
 * tabindex — one cell is tabbable, the arrows move focus, so Tab passes over
 * the whole grid in one stop. Cells get `role="gridcell"`; the caller sets
 * `role="grid"` on the container. Rows may be ragged (the last row of a
 * symbol grid): a move clamps to the target row's last cell. Returns a
 * `focusCell` for the caller to place focus programmatically (initial cell,
 * search results).
 */
export function rovingGrid(cells: HTMLElement[][]): {
  focusCell(r: number, c: number): void;
} {
  let cur: [number, number] = [0, 0];
  const at = (r: number, c: number): HTMLElement | undefined => cells[r]?.[c];
  const focusCell = (r: number, c: number): void => {
    const rows = cells.length;
    if (rows === 0) return;
    const nr = Math.min(rows - 1, Math.max(0, r));
    const nc = Math.min(cells[nr].length - 1, Math.max(0, c));
    const prev = at(cur[0], cur[1]);
    if (prev) prev.tabIndex = -1;
    const next = at(nr, nc);
    if (!next) return;
    next.tabIndex = 0;
    next.focus();
    cur = [nr, nc];
  };
  cells.forEach((row, r) =>
    row.forEach((cell, c) => {
      cell.setAttribute('role', 'gridcell');
      cell.tabIndex = r === 0 && c === 0 ? 0 : -1;
      cell.addEventListener('keydown', (e) => {
        const delta: Record<string, [number, number]> = {
          ArrowRight: [0, 1],
          ArrowLeft: [0, -1],
          ArrowDown: [1, 0],
          ArrowUp: [-1, 0],
          Home: [0, -Infinity],
          End: [0, Infinity],
        };
        const d = delta[e.key];
        if (!d) return;
        e.preventDefault();
        focusCell(r + d[0], c + d[1]);
      });
      cell.addEventListener('focus', () => {
        cur = [r, c];
      });
    }),
  );
  return { focusCell };
}
