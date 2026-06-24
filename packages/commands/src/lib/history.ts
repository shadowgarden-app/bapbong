import { undo, redo, undoDepth, redoDepth } from 'prosemirror-history';
import type { EditorState } from 'prosemirror-state';
import type { Command } from '@shadow-garden/bapbong-contracts';

// `undoDepth`/`redoDepth` read the prosemirror-history plugin state; on a
// headless state without history they throw, so probe defensively (disabled).
const depth = (fn: (s: EditorState) => number, state: EditorState): number => {
  try {
    return fn(state);
  } catch {
    return 0;
  }
};

/** Undo the last change (requires the prosemirror-history plugin in state). */
export function undoCommand(): Command {
  return {
    name: 'undo',
    run: (state, dispatch) => undo(state, dispatch),
    isEnabled: (state) => depth(undoDepth, state) > 0,
  };
}

/** Redo the last undone change. */
export function redoCommand(): Command {
  return {
    name: 'redo',
    run: (state, dispatch) => redo(state, dispatch),
    isEnabled: (state) => depth(redoDepth, state) > 0,
  };
}
