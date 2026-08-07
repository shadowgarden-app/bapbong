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
