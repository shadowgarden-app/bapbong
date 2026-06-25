import { deleteSelection } from 'prosemirror-commands';
import type { Command } from '@shadow-garden/bapbong-contracts';

/** Delete the current selection (no-op when the selection is empty). */
export function deleteSelectionCommand(): Command {
  return {
    name: 'delete',
    run: (state, dispatch) => deleteSelection(state, dispatch),
    isEnabled: (state) => !state.selection.empty,
  };
}
