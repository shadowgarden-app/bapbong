import type { EditorState, Transaction } from 'prosemirror-state';

/** Apply a transaction to the editor — ProseMirror's dispatch convention. */
export type Dispatch = (tr: Transaction) => void;

/**
 * A named editor operation: the shared vocabulary the toolbar/menubar, plugins,
 * and a headless (server-side) caller all speak. It operates purely on
 * {@link EditorState} with no DOM, so the *same* command runs in the browser
 * and on the backend.
 *
 * This is just the contract — implementations live in
 * `@shadow-garden/bapbong-commands`. Keeping the type here lets the editor
 * expose a `Collection<Command>` registry, and the menubar read it, without
 * either importing the implementations (same rationale as {@link EditorPlugin}).
 */
export interface Command {
  /** Stable identifier — the registry key and the name a menu/toolbar references. */
  readonly name: string;
  /**
   * Apply the operation. Following ProseMirror's convention, omit `dispatch` to
   * probe whether it *would* apply (returns true/false without mutating).
   */
  run(state: EditorState, dispatch?: Dispatch): boolean;
  /** Whether the operation is currently "on" (e.g. bold over the selection) —
   *  drives toggle/active UI. */
  isActive?(state: EditorState): boolean;
  /** Whether the operation can apply at all (else a button renders disabled).
   *  When omitted, callers fall back to a no-dispatch `run` probe. */
  isEnabled?(state: EditorState): boolean;
}
