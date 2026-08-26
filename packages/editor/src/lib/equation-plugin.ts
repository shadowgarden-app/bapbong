import type {
  EditorPlugin,
  PluginContext,
  RangeDecoration,
} from '@shadow-garden/bapbong-contracts';

/** The editor state type, taken from the plugin context (no direct PM dep). */
type State = PluginContext['state'];

/** Word's equation-region tint: light enough to read through, blue enough to
 *  say "you are inside an equation now". Painted BEHIND the text. */
const REGION_TINT = 'rgba(59, 130, 246, 0.12)';
/** A hairline under the region marks its exact extent — the closest the
 *  decoration kinds come to the mockup's dashed frame. */
const REGION_EDGE = '#7fb2ec';

/**
 * The maximal contiguous math-marked text range containing `pos`, or null
 * when `pos` does not touch one. "Touching" counts the boundary the mark is
 * inclusive at (its end): a caret right after an equation still shows the
 * region — that is where continued typing extends it.
 */
export function mathRangeAt(
  state: State,
  pos: number,
): { from: number; to: number } | null {
  if (pos < 0 || pos > state.doc.content.size) return null;
  const $pos = state.doc.resolve(pos);
  const parent = $pos.parent;
  if (!parent.isTextblock) return null;
  const start = $pos.start();
  const isMath = (i: number): boolean => {
    const child = parent.maybeChild(i);
    return (
      (child?.isText ?? false) &&
      (child as { marks: readonly { type: { name: string } }[] }).marks.some(
        (m) => m.type.name === 'math',
      )
    );
  };
  // The child at the caret, or the one just before it (inclusive-end).
  const index = $pos.index();
  let anchor = -1;
  if (isMath(index)) anchor = index;
  else if ($pos.textOffset === 0 && index > 0 && isMath(index - 1))
    anchor = index - 1;
  if (anchor < 0) return null;
  let first = anchor;
  while (first > 0 && isMath(first - 1)) first--;
  let last = anchor;
  while (last + 1 < parent.childCount && isMath(last + 1)) last++;
  let from = start;
  for (let i = 0; i < first; i++) from += parent.child(i).nodeSize;
  let to = from;
  for (let i = first; i <= last; i++) to += parent.child(i).nodeSize;
  return { from, to };
}

/**
 * Equation affordance: while the caret (or selection) sits in a math-marked
 * run, the whole contiguous equation gets a light tint and a hairline under
 * its extent — the "you are editing an equation" frame of the approved UX,
 * expressed through the decoration pipeline (background + underline kinds).
 */
export function equationPlugin(): EditorPlugin {
  let ctx: PluginContext | null = null;
  /** The region painted last frame — content repaints happen only on doc
   *  changes, so a caret moving in or out of an equation must request one. */
  let last: { from: number; to: number } | null = null;
  const rangeFor = (state: State) => {
    const { selection } = state;
    return (
      mathRangeAt(state, selection.from) ??
      (selection.empty ? null : mathRangeAt(state, selection.to))
    );
  };
  return {
    name: 'equation',
    setup(c) {
      ctx = c;
      return () => {
        ctx = null;
      };
    },
    onChange() {
      if (!ctx) return;
      const range = rangeFor(ctx.state);
      const changed =
        (range === null) !== (last === null) ||
        (range !== null &&
          last !== null &&
          (range.from !== last.from || range.to !== last.to));
      if (changed) ctx.requestPaint();
    },
    decorations(c): RangeDecoration[] {
      const range = rangeFor(c.state);
      last = range;
      if (!range) return [];
      return [
        {
          from: range.from,
          to: range.to,
          kind: 'background',
          color: REGION_TINT,
        },
        {
          from: range.from,
          to: range.to,
          kind: 'underline',
          color: REGION_EDGE,
        },
      ];
    },
  };
}
