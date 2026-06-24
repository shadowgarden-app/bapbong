import type { Collection, Command, Dispatch, EditorChange } from '@shadow-garden/bapbong-contracts';

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
