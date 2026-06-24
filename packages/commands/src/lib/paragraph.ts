import type { EditorState } from 'prosemirror-state';
import type { Command } from '@shadow-garden/bapbong-contracts';

/** Paragraph alignment values (mirrors the `paragraph.align` schema attr / w:jc). */
export const ALIGNMENTS = ['left', 'center', 'right', 'justify'] as const;
export type Align = (typeof ALIGNMENTS)[number];

/**
 * The alignment shared by every paragraph touched by the selection, or null
 * when they differ or use the default. Drives the toolbar's active alignment.
 */
export function activeAlign(state: EditorState): Align | null {
  const { from, to } = state.selection;
  let seen: Align | null | undefined;
  state.doc.nodesBetween(from, to, (node) => {
    if (node.type.name !== 'paragraph') return;
    const a = (node.attrs['align'] ?? null) as Align | null;
    seen = seen === undefined ? a : seen === a ? seen : null; // null = mixed
  });
  return seen ?? null;
}

/**
 * Set (or clear, with `null`) paragraph alignment across the selection. The
 * command id is `align-<value>` (or `align-clear`) so each alignment is its own
 * toolbar button.
 */
export function setAlign(align: Align | null): Command {
  return {
    name: align ? `align-${align}` : 'align-clear',
    run(state, dispatch) {
      const { from, to } = state.selection;
      const tr = state.tr;
      let changed = false;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (node.type.name !== 'paragraph' || node.attrs['align'] === align) return;
        tr.setNodeAttribute(pos, 'align', align);
        changed = true;
      });
      if (changed && dispatch) dispatch(tr);
      return changed;
    },
    isActive: (state) => activeAlign(state) === align,
  };
}
