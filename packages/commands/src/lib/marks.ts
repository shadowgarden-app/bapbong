import { toggleMark } from 'prosemirror-commands';
import type { EditorState } from 'prosemirror-state';
import type { Command } from '@shadow-garden/bapbong-contracts';

/**
 * Whether `markName` covers the selection (or sits in the stored marks when the
 * selection is empty). Schema-agnostic: an unknown mark is simply inactive, so
 * a command for a mark the document's schema doesn't define stays inert.
 */
export function isMarkActive(state: EditorState, markName: string): boolean {
  const type = state.schema.marks[markName];
  if (!type) return false;
  const { from, to, empty, $from } = state.selection;
  if (empty) return !!type.isInSet(state.storedMarks ?? $from.marks());
  return state.doc.rangeHasMark(from, to, type);
}

/**
 * Toggle a mark over the selection. `name` is the command id (what menus
 * reference, e.g. `bold`); `markName` is the schema mark it toggles (e.g.
 * `strong`) and defaults to `name` when they match (`underline`, `strike`).
 */
export function toggleMarkCommand(name: string, markName: string = name): Command {
  return {
    name,
    run(state, dispatch) {
      const type = state.schema.marks[markName];
      return type ? toggleMark(type)(state, dispatch) : false;
    },
    isActive: (state) => isMarkActive(state, markName),
    isEnabled: (state) => !!state.schema.marks[markName],
  };
}
