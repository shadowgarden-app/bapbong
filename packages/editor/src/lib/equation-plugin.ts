import type {
  EditorPlugin,
  EqNode,
  EqSlotRect,
  LayoutImageSegment,
  PluginContext,
  RangeDecoration,
  ResolvedLayout,
  ResolvedTable,
} from '@shadow-garden/bapbong-contracts';
import { isEqRow } from '@shadow-garden/bapbong-contracts';
import {
  autoCorrectAt,
  eqChar,
  insertAt,
  removeBefore,
  rowAt,
  structureFor,
} from './equation-edit';

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

/** A typeset equation found in the layout: page + box origin + its segment. */
interface EqHit {
  pageIndex: number;
  x: number;
  top: number;
  seg: LayoutImageSegment;
}

/** Every equation segment on the layout, by walking lines (tables too). */
function eqSegments(layout: ResolvedLayout | null): EqHit[] {
  if (!layout) return [];
  const out: EqHit[] = [];
  for (const page of layout.pages) {
    const lines = [...page.lines];
    const visit = (t: ResolvedTable): void => {
      for (const cell of t.cells) {
        lines.push(...cell.lines);
        cell.tables?.forEach(visit);
      }
    };
    page.tables?.forEach(visit);
    for (const line of lines) {
      for (const img of line.images ?? []) {
        if (!img.eqSlots || img.pos == null) continue;
        out.push({
          pageIndex: page.index,
          x: img.x,
          top: line.y + line.baseline - img.height - (img.raise ?? 0),
          seg: img,
        });
      }
    }
  }
  return out;
}

/** The slot editor's position inside one equation node. */
interface EqSel {
  pos: number;
  /** Index into the segment's eqSlots. */
  slot: number;
  /** Caret stop index within the slot. */
  caret: number;
  /** After a structural edit: the slot PATH to land in once the layout
   *  rebuilds (slot indexes shift when the tree changes). */
  pendingPath?: string;
}

/**
 * Equation affordances:
 *  - while the caret sits in a LINEAR math run, the region tints (approved
 *    UX frame) — see mathRangeAt;
 *  - a click inside a TYPESET equation node opens the slot editor: a caret
 *    guide in the slot, typing/Backspace edit the AST, `/ ^ _` insert
 *    structures, `\name`+space autocorrects, arrows/Tab walk slots, Esc
 *    leaves. Every edit is one ProseMirror transaction (one undo step).
 */
export function equationPlugin(): EditorPlugin {
  let ctx: PluginContext | null = null;
  let last: { from: number; to: number } | null = null;
  let eq: EqSel | null = null;
  /** The document selection the slot editor was entered on. While it owns the
   *  caret the document's own selection never moves (every key is claimed), so
   *  a selection that HAS moved means something else took over — ⌘A, find, a
   *  host command — and the slot editor must let go of the caret. */
  let enteredOn = '';

  const rangeFor = (state: State) => {
    const { selection } = state;
    return (
      mathRangeAt(state, selection.from) ??
      (selection.empty ? null : mathRangeAt(state, selection.to))
    );
  };

  const nodeAt = (c: PluginContext, pos: number) => {
    if (pos < 0 || pos >= c.state.doc.content.size) return null;
    const n = c.state.doc.nodeAt(pos);
    return n?.type.name === 'equation' ? n : null;
  };

  const hitFor = (c: PluginContext, pos: number): EqHit | null =>
    eqSegments(c.layout).find((h) => h.seg.pos === pos) ?? null;

  const showCaret = (c: PluginContext): void => {
    if (!eq) {
      c.setGuide(null);
      return;
    }
    const hit = hitFor(c, eq.pos);
    const slot = hit?.seg.eqSlots?.[eq.slot];
    if (!hit || !slot) {
      c.setGuide(null);
      return;
    }
    const caret = Math.min(eq.caret, slot.caretXs.length - 1);
    c.setGuide({
      kind: 'caret',
      pageIndex: hit.pageIndex,
      x: hit.x + slot.x + slot.caretXs[caret],
      y: hit.top + slot.y,
      height: slot.height,
    });
  };

  const selKey = (c: PluginContext): string =>
    `${c.state.selection.from}:${c.state.selection.to}`;

  const enter = (c: PluginContext, next: EqSel): void => {
    eq = next;
    enteredOn = selKey(c);
    showCaret(c);
  };

  const leave = (c: PluginContext): void => {
    eq = null;
    c.setGuide(null);
  };

  const astOf = (c: PluginContext): EqNode[] | null => {
    if (!eq) return null;
    const ast = nodeAt(c, eq.pos)?.attrs['ast'];
    return isEqRow(ast) ? ast : null;
  };

  const commit = (
    c: PluginContext,
    ast: EqNode[],
    next: { slot?: number; caret: number; pendingPath?: string },
  ): void => {
    if (!eq) return;
    eq = {
      pos: eq.pos,
      slot: next.slot ?? eq.slot,
      caret: next.caret,
      ...(next.pendingPath !== undefined && { pendingPath: next.pendingPath }),
    };
    c.dispatch(c.state.tr.setNodeAttribute(eq.pos, 'ast', ast));
  };

  const slotOf = (c: PluginContext): EqSlotRect | null => {
    if (!eq) return null;
    return hitFor(c, eq.pos)?.seg.eqSlots?.[eq.slot] ?? null;
  };

  return {
    name: 'equation',
    setup(c) {
      ctx = c;
      return () => {
        ctx = null;
        eq = null;
      };
    },
    onChange() {
      if (!ctx) return;
      // A freshly inserted equation arrives selected as a node (the gallery
      // and Insert ▸ Equation both do this): step straight into its first
      // slot, so typing continues inside the equation instead of replacing
      // it — what Word does when it drops the caret into a new equation.
      const sel = ctx.state.selection as { node?: { type: { name: string } } };
      if (sel.node?.type.name === 'equation') {
        const pos = ctx.state.selection.from;
        if (!eq || eq.pos !== pos) enter(ctx, { pos, slot: 0, caret: 0 });
      } else if (eq && selKey(ctx) !== enteredOn) {
        leave(ctx);
      }
      // The typeset editor: re-anchor after every layout (slots move on each
      // reflow; a structural edit lands in its pending path).
      if (eq) {
        if (!nodeAt(ctx, eq.pos)) {
          leave(ctx);
        } else {
          const hit = hitFor(ctx, eq.pos);
          if (hit?.seg.eqSlots && eq.pendingPath !== undefined) {
            const idx = hit.seg.eqSlots.findIndex(
              (s) => JSON.stringify(s.path) === eq!.pendingPath,
            );
            if (idx >= 0) eq = { pos: eq.pos, slot: idx, caret: 0 };
            else eq = { pos: eq.pos, slot: eq.slot, caret: eq.caret };
          }
          showCaret(ctx);
        }
      }
      // The linear math-region tint.
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
    onPointer(ev) {
      const c = ctx;
      if (!c || ev.type !== 'down' || !ev.point) return false;
      const { pageIndex, x, y } = ev.point;
      for (const hit of eqSegments(c.layout)) {
        if (hit.pageIndex !== pageIndex) continue;
        const lx = x - hit.x;
        const ly = y - hit.top;
        if (lx < 0 || ly < 0 || lx > hit.seg.width || ly > hit.seg.height)
          continue;
        const slots = hit.seg.eqSlots ?? [];
        // Deepest slot under the point (children come after their parents).
        let slot = -1;
        for (let i = 0; i < slots.length; i++) {
          const s = slots[i];
          if (
            lx >= s.x &&
            lx <= s.x + s.width &&
            ly >= s.y &&
            ly <= s.y + s.height
          )
            slot = i;
        }
        if (slot < 0) slot = 0;
        const s = slots[slot];
        if (!s) return false;
        let caret = 0;
        for (let i = 1; i < s.caretXs.length; i++)
          if (
            Math.abs(lx - s.x - s.caretXs[i]) <
            Math.abs(lx - s.x - s.caretXs[caret])
          )
            caret = i;
        enter(c, { pos: hit.seg.pos!, slot, caret });
        return true;
      }
      if (eq) leave(c);
      return false;
    },
    onKey(ev) {
      const c = ctx;
      if (!c || !eq) return false;
      if (ev.metaKey || ev.ctrlKey) return false; // shortcuts stay global
      const ast = astOf(c);
      const slot = slotOf(c);
      if (!ast || !slot) {
        leave(c);
        return false;
      }
      const path = slot.path;
      const caret = Math.min(eq.caret, slot.caretXs.length - 1);
      const slots = hitFor(c, eq.pos)?.seg.eqSlots ?? [];

      if (ev.key === 'Escape') {
        leave(c);
        return true;
      }
      if (ev.key === 'Tab') {
        // Next slot, empty ones first in encounter order.
        const next = (eq.slot + 1) % Math.max(1, slots.length);
        eq = { pos: eq.pos, slot: next, caret: 0 };
        showCaret(c);
        return true;
      }
      if (ev.key === 'ArrowLeft' || ev.key === 'ArrowRight') {
        const d = ev.key === 'ArrowRight' ? 1 : -1;
        const target = caret + d;
        if (target >= 0 && target < slot.caretXs.length) {
          eq = { pos: eq.pos, slot: eq.slot, caret: target };
        } else {
          const next = eq.slot + d;
          if (next < 0 || next >= slots.length) {
            // Out of the equation, the way Word does it: the document caret
            // reappears on the side the arrow was heading.
            const at = d > 0 ? eq.pos + 1 : eq.pos;
            leave(c);
            c.setSelection(at);
            return true;
          }
          eq = {
            pos: eq.pos,
            slot: next,
            caret: d > 0 ? 0 : slots[next].caretXs.length - 1,
          };
        }
        showCaret(c);
        return true;
      }
      if (ev.key === 'Backspace') {
        commit(c, removeBefore(ast, path, caret), {
          caret: Math.max(0, caret - 1),
        });
        return true;
      }
      if (ev.key === ' ') {
        const fixed = autoCorrectAt(ast, path, caret);
        if (fixed) {
          commit(c, fixed.ast, { caret: fixed.caret });
          return true;
        }
        commit(c, insertAt(ast, path, caret, { t: 'chr', ch: ' ' }), {
          caret: caret + 1,
        });
        return true;
      }
      const structure = structureFor(ev.key);
      if (structure) {
        const idx = rowAt(ast, path)?.length ?? 0;
        void idx;
        const target =
          structure.t === 'frac' ? 'num' : ev.key === '_' ? 'sub' : 'sup';
        commit(c, insertAt(ast, path, caret, structure), {
          caret: caret + 1,
          pendingPath: JSON.stringify([...path, caret, target]),
        });
        return true;
      }
      if (ev.key.length === 1 && !ev.altKey) {
        commit(c, insertAt(ast, path, caret, eqChar(ev.key)), {
          caret: caret + 1,
        });
        return true;
      }
      return false;
    },
  };
}
