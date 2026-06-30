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

/**
 * Set a value-bearing mark (`fontSize`/`textColor`/`highlight`/`fontFamily`) to
 * specific `attrs` over the selection — replacing any existing value — or
 * remove it when `attrs` is null. Empty selection updates the stored marks so
 * the next typed text carries the value. The value-bearing analogue of
 * {@link toggleMarkCommand} (which only flips a mark on/off).
 */
export function setMarkAttr(markName: string, attrs: Record<string, unknown> | null): Command {
  return {
    name: `set-${markName}`,
    run(state, dispatch) {
      const type = state.schema.marks[markName];
      if (!type) return false;
      if (dispatch) {
        const { from, to, empty } = state.selection;
        const tr = state.tr;
        if (empty) {
          tr.removeStoredMark(type);
          if (attrs) tr.addStoredMark(type.create(attrs));
        } else {
          tr.removeMark(from, to, type);
          if (attrs) tr.addMark(from, to, type.create(attrs));
        }
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isEnabled: (state) => !!state.schema.marks[markName],
  };
}

/**
 * The value of `markName`'s `attrKey` over the selection: the stored/cursor
 * mark when empty, the shared value if uniform across a range, else null (mixed
 * or unset). Used to show the current font size / colour in a toolbar control.
 */
export function activeMarkValue(state: EditorState, markName: string, attrKey: string): unknown {
  const type = state.schema.marks[markName];
  if (!type) return null;
  const { from, to, empty, $from } = state.selection;
  const valueOf = (marks: readonly Mark[]): unknown => {
    const mark = type.isInSet(marks);
    return mark ? mark.attrs[attrKey] : null;
  };
  if (empty) return valueOf(state.storedMarks ?? $from.marks());
  let value: unknown;
  let seen = false;
  state.doc.nodesBetween(from, to, (node) => {
    if (!node.isText) return;
    const v = valueOf(node.marks);
    if (!seen) {
      value = v;
      seen = true;
    } else if (value !== v) {
      value = null;
    }
  });
  return seen ? value : null;
}

/** Set the font size (points) over the selection; null clears it. */
export function setFontSize(pt: number | null): Command {
  return setMarkAttr('fontSize', pt == null ? null : { size: pt });
}

/** The font size (pt) at the selection, or null when mixed/unset. */
export function activeFontSize(state: EditorState): number | null {
  const v = activeMarkValue(state, 'fontSize', 'size');
  return typeof v === 'number' ? v : null;
}

/** Set the font family over the selection; null clears it. */
export function setFontFamily(family: string | null): Command {
  return setMarkAttr('fontFamily', family ? { family } : null);
}

/** The font family at the selection, or null when mixed/unset. */
export function activeFontFamily(state: EditorState): string | null {
  const v = activeMarkValue(state, 'fontFamily', 'family');
  return typeof v === 'string' ? v : null;
}
