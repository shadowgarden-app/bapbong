import { toggleMark } from 'prosemirror-commands';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import type {
  CharacterFormatting,
  Command,
} from '@shadow-garden/bapbong-contracts';

// ── The paragraph mark ──────────────────────────────────────────────
// Word's ¶ is a character: it sits at the end of the paragraph, carries its
// own font (w:pPr/w:rPr → the paragraph's `markFont` attr) and takes part in
// the LAST line's height. Word re-formats it with the text whenever the
// selection takes it in, so the font commands here do the same — otherwise
// shrinking a paragraph to 8pt would leave its last line 11pt tall, and a
// document saved from here would reopen in Word with the same gap.

/** The `markFont` keys a character mark maps onto, and how. */
type MarkFont = {
  family?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
};
type MarkFontPatch = Partial<Record<keyof MarkFont, unknown>>;

/** What a mark set on the selection means for the paragraph mark's font:
 *  the value to carry (`undefined` = clear that key), or null when the mark
 *  has no bearing on the ¶ (colour, underline, …). */
function markFontPatch(
  markName: string,
  attrs: Record<string, unknown> | null,
): MarkFontPatch | null {
  switch (markName) {
    case 'fontSize':
      return { sizePt: attrs ? attrs['size'] : undefined };
    case 'fontFamily':
      return { family: attrs ? attrs['family'] : undefined };
    case 'strong':
      return { bold: attrs ? true : undefined };
    case 'em':
      return { italic: attrs ? true : undefined };
    default:
      return null;
  }
}

/**
 * The paragraphs whose MARK the selection takes in, Word's way: the ¶ ends
 * the paragraph, so it is selected when the selection runs PAST the
 * paragraph's content — into the next block, or to the end of the document /
 * cell (select-all, a whole-paragraph or cell selection). A selection that
 * stops at the last character leaves it out. A collapsed caret in an EMPTY
 * paragraph counts too: the mark is the only thing there for formatting to
 * land on, and Word re-sizes it (the empty line grows) rather than only
 * priming the next keystroke.
 */
function paragraphsWithMark(
  state: EditorState,
): { pos: number; attrs: Record<string, unknown> }[] {
  const { from, to, empty } = state.selection;
  const hits: { pos: number; attrs: Record<string, unknown> }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return;
    if (!node.type.spec.attrs?.['markFont']) return false; // schema without a ¶
    const contentEnd = pos + 1 + node.content.size;
    if (to > contentEnd || (empty && node.content.size === 0))
      hits.push({ pos, attrs: node.attrs });
    return false;
  });
  return hits;
}

/** Apply `patch` (or, with `reset`, clear the whole font) to the mark of every
 *  paragraph the selection takes in. Positions are pre-mapping: callers add
 *  this to a transaction whose steps so far only add/remove marks. */
function retouchMarks(
  tr: Transaction,
  state: EditorState,
  patch: MarkFontPatch | 'reset',
): void {
  // A doc-changing step drops the transaction's stored marks — the very ones
  // a caret command just set for the next keystroke. Put them back after.
  const stored = tr.storedMarks;
  for (const { pos, attrs } of paragraphsWithMark(state)) {
    const cur = (attrs['markFont'] as MarkFont | null) ?? {};
    let next: MarkFont | null;
    if (patch === 'reset') {
      next = null;
    } else {
      const merged: Record<string, unknown> = { ...cur };
      for (const [k, v] of Object.entries(patch)) {
        if (v === undefined) delete merged[k];
        else merged[k] = v;
      }
      next = Object.keys(merged).length ? (merged as MarkFont) : null;
    }
    // Same font, same object shape → nothing to record.
    if (JSON.stringify(next) === JSON.stringify(attrs['markFont'] ?? null))
      continue;
    tr.setNodeMarkup(pos, null, { ...attrs, markFont: next });
  }
  if (stored && tr.storedMarks !== stored) tr.setStoredMarks(stored);
}

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
      if (!type) return false;
      const patch = markFontPatch(markName, attrs ?? {});
      if (!dispatch || !patch) return toggleMark(type, attrs)(state, dispatch);
      // toggleMark REMOVES when the mark is anywhere in the range (or in the
      // stored marks at a caret) and adds otherwise — the ¶ follows the same
      // verdict.
      const { from, to, empty, $from } = state.selection;
      const on = empty
        ? !type.isInSet(state.storedMarks ?? $from.marks())
        : !state.doc.rangeHasMark(from, to, type);
      return toggleMark(type, attrs)(state, (tr) => {
        retouchMarks(tr, state, on ? patch : markFontPatch(markName, null)!);
        dispatch(tr);
      });
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
        const patch = markFontPatch(markName, attrs);
        if (patch) retouchMarks(tr, state, patch);
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
      const markPatch: MarkFontPatch = {};
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
        Object.assign(markPatch, markFontPatch(markName, attrs) ?? {});
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
      if (Object.keys(markPatch).length) retouchMarks(tr, state, markPatch);
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
        retouchMarks(tr, state, 'reset'); // the ¶ back to the base font too
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
