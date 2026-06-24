import { toggleMark } from 'prosemirror-commands';
import type { EditorState } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import type { Command } from '@shadow-garden/bapbong-contracts';

const matches = (mark: Mark, attrs?: Record<string, unknown>): boolean =>
  !attrs || Object.entries(attrs).every(([k, v]) => mark.attrs[k] === v);

/**
 * Whether `markName` (optionally with specific `attrs`) covers the selection,
 * or sits in the stored marks when the selection is empty. Schema-agnostic: an
 * unknown mark is simply inactive. With `attrs` (e.g. `{ value: 'super' }` for
 * a superscript) it matches only marks carrying those attribute values.
 */
export function isMarkActive(
  state: EditorState,
  markName: string,
  attrs?: Record<string, unknown>,
): boolean {
  const type = state.schema.marks[markName];
  if (!type) return false;
  const { from, to, empty, $from } = state.selection;
  if (empty) {
    const mark = type.isInSet(state.storedMarks ?? $from.marks());
    return !!mark && matches(mark, attrs);
  }
  if (!attrs) return state.doc.rangeHasMark(from, to, type);
  let found = false;
  state.doc.nodesBetween(from, to, (node) => {
    if (found || !node.isText) return;
    const mark = node.marks.find((m) => m.type === type);
    if (mark && matches(mark, attrs)) found = true;
  });
  return found;
}

/**
 * Toggle a mark over the selection. `name` is the command id (what menus
 * reference, e.g. `bold`); `markName` is the schema mark it toggles (defaults to
 * `name`). Pass `attrs` for attribute-bearing marks (e.g. superscript =
 * `vertAlign` with `{ value: 'super' }`); same-type marks exclude each other,
 * so super/sub stay mutually exclusive.
 */
export function toggleMarkCommand(
  name: string,
  markName: string = name,
  attrs?: Record<string, unknown>,
): Command {
  return {
    name,
    run(state, dispatch) {
      const type = state.schema.marks[markName];
      return type ? toggleMark(type, attrs)(state, dispatch) : false;
    },
    isActive: (state) => isMarkActive(state, markName, attrs),
    isEnabled: (state) => !!state.schema.marks[markName],
  };
}
