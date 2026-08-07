import { toggleMark } from 'prosemirror-commands';
import type { EditorState } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import type {
  CharacterFormatting,
  Command,
} from '@shadow-garden/bapbong-contracts';

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
export function setMarkAttr(
  markName: string,
  attrs: Record<string, unknown> | null,
): Command {
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
export function activeMarkValue(
  state: EditorState,
  markName: string,
  attrKey: string,
): unknown {
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

// ── Advanced character formatting ───────────────────────────────────
// Continuous, value-bearing properties. Each keeps the DOCUMENT's own unit,
// the way the marks do: twips for tracking, percent for the glyph scale,
// half-points for the baseline shift. Converting here would put a second
// conversion beside the layout's, and the two would drift.

/** Set character tracking (w:spacing) in TWIPS; null clears it. */
export function setLetterSpacing(twips: number | null): Command {
  return setMarkAttr('letterSpacing', twips == null ? null : { twips });
}

/** Tracking in twips over the selection, or null when unset or mixed. */
export function activeLetterSpacing(state: EditorState): number | null {
  const v = activeMarkValue(state, 'letterSpacing', 'twips');
  return typeof v === 'number' ? v : null;
}

/** Set horizontal glyph scale (w:w) as a PERCENT; null clears it. */
export function setCharScale(percent: number | null): Command {
  return setMarkAttr('charScale', percent == null ? null : { percent });
}

/** Glyph scale percent over the selection, or null when unset or mixed. */
export function activeCharScale(state: EditorState): number | null {
  const v = activeMarkValue(state, 'charScale', 'percent');
  return typeof v === 'number' ? v : null;
}

/** Set the baseline shift (w:position) in HALF-POINTS; null clears it. */
export function setPosition(halfPoints: number | null): Command {
  return setMarkAttr('position', halfPoints == null ? null : { halfPoints });
}

/** Baseline shift in half-points over the selection, or null when unset. */
export function activePosition(state: EditorState): number | null {
  const v = activeMarkValue(state, 'position', 'halfPoints');
  return typeof v === 'number' ? v : null;
}

// ── The Font dialog's read/apply pair ───────────────────────────────

/** A boolean mark over the selection: true/false when uniform, undefined when
 *  mixed — which the dialog shows as an indeterminate box and leaves alone. */
function activeToggle(
  state: EditorState,
  markName: string,
): boolean | undefined {
  const type = state.schema.marks[markName];
  if (!type) return undefined;
  const { from, to, empty, $from } = state.selection;
  if (empty) return !!type.isInSet(state.storedMarks ?? $from.marks());
  let value: boolean | undefined;
  let mixed = false;
  state.doc.nodesBetween(from, to, (node) => {
    if (!node.isText || mixed) return;
    const on = !!type.isInSet(node.marks);
    if (value === undefined) value = on;
    else if (value !== on) mixed = true;
  });
  return mixed ? undefined : (value ?? false);
}

/** Read every property the Font dialog shows, for the current selection. */
export function activeCharacterFormatting(
  state: EditorState,
): CharacterFormatting {
  const va = activeMarkValue(state, 'vertAlign', 'value');
  return {
    family: activeFontFamily(state),
    sizePt: activeFontSize(state),
    bold: activeToggle(state, 'strong'),
    italic: activeToggle(state, 'em'),
    underline: activeToggle(state, 'underline'),
    strike: activeToggle(state, 'strike'),
    doubleStrike: activeToggle(state, 'dstrike'),
    smallCaps: activeToggle(state, 'smallCaps'),
    vertAlign: va === 'super' || va === 'sub' ? va : null,
    color: activeTextColor(state),
    highlight: activeHighlight(state),
    scalePercent: activeCharScale(state) ?? 100,
    letterSpacingTwips: activeLetterSpacing(state) ?? 0,
    positionHalfPoints: activePosition(state) ?? 0,
  };
}

/**
 * Apply the whole dialog in ONE transaction, so the user's visit is a single
 * undo step rather than thirteen. `undefined` fields are skipped — that is how
 * a mixed selection keeps the formatting the user never looked at.
 */
export function applyCharacterFormatting(v: CharacterFormatting): Command {
  return {
    name: 'apply-character-formatting',
    run(state, dispatch) {
      if (!dispatch) return true;
      const { from, to, empty } = state.selection;
      const tr = state.tr;
      const put = (markName: string, attrs: Record<string, unknown> | null) => {
        const type = state.schema.marks[markName];
        if (!type) return;
        if (empty) {
          tr.removeStoredMark(type);
          if (attrs) tr.addStoredMark(type.create(attrs));
        } else {
          tr.removeMark(from, to, type);
          if (attrs) tr.addMark(from, to, type.create(attrs));
        }
      };
      const toggle = (markName: string, on: boolean | undefined) => {
        if (on !== undefined) put(markName, on ? {} : null);
      };
      if (v.family !== undefined)
        put('fontFamily', v.family ? { family: v.family } : null);
      if (v.sizePt !== undefined)
        put('fontSize', v.sizePt ? { size: v.sizePt } : null);
      if (v.color !== undefined)
        put('textColor', v.color ? { color: v.color } : null);
      if (v.highlight !== undefined)
        put('highlight', v.highlight ? { color: v.highlight } : null);
      toggle('strong', v.bold);
      toggle('em', v.italic);
      toggle('underline', v.underline);
      toggle('strike', v.strike);
      toggle('dstrike', v.doubleStrike);
      toggle('smallCaps', v.smallCaps);
      if (v.vertAlign !== undefined)
        put('vertAlign', v.vertAlign ? { value: v.vertAlign } : null);
      // The three below always carry a definite number; the default IS the
      // clear, so a run never keeps a "scale 100%" mark Word would not write.
      put(
        'charScale',
        v.scalePercent === 100 ? null : { percent: v.scalePercent },
      );
      put(
        'letterSpacing',
        v.letterSpacingTwips === 0 ? null : { twips: v.letterSpacingTwips },
      );
      put(
        'position',
        v.positionHalfPoints === 0
          ? null
          : { halfPoints: v.positionHalfPoints },
      );
      dispatch(tr.scrollIntoView());
      return true;
    },
  };
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

/** Set the text colour over the selection; null clears it. */
export function setTextColor(color: string | null): Command {
  return setMarkAttr('textColor', color ? { color } : null);
}

/** The text colour at the selection, or null when mixed/unset. */
export function activeTextColor(state: EditorState): string | null {
  const v = activeMarkValue(state, 'textColor', 'color');
  return typeof v === 'string' ? v : null;
}

/**
 * Remove all inline marks from the selection (Clear formatting). On an empty
 * selection it clears the stored marks so the next typed text is unformatted.
 */
export function clearMarks(): Command {
  return {
    name: 'clear-format',
    run(state, dispatch) {
      const { from, to, empty } = state.selection;
      if (dispatch) {
        const tr = state.tr;
        if (empty) tr.setStoredMarks([]);
        else tr.removeMark(from, to); // no mark arg → strip every mark in the range
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
  };
}

/** Set the highlight (text background) colour over the selection; null clears it. */
export function setHighlight(color: string | null): Command {
  return setMarkAttr('highlight', color ? { color } : null);
}

/** The highlight colour at the selection, or null when mixed/unset. */
export function activeHighlight(state: EditorState): string | null {
  const v = activeMarkValue(state, 'highlight', 'color');
  return typeof v === 'string' ? v : null;
}
