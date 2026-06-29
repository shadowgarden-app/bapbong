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

/** Every paragraph the selection touches (heading toggle target). */
function selectedParagraphs(state: EditorState): { pos: number; node: import('prosemirror-model').Node }[] {
  const { from, to } = state.selection;
  const out: { pos: number; node: import('prosemirror-model').Node }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === 'paragraph') out.push({ pos, node });
  });
  return out;
}

/**
 * Toggle a heading level (1–6) on the selected paragraphs (`heading-<level>`).
 * If every target is already this level it reverts to a body paragraph;
 * otherwise they all become this level. The layout sizes a heading from its
 * level and the model serialises it as `<h1>`–`<h6>` (semantic for the a11y
 * mirror) / a Word "Heading N" style on export.
 */
export function toggleHeading(level: number): Command {
  return {
    name: `heading-${level}`,
    run(state, dispatch) {
      const paras = selectedParagraphs(state);
      if (paras.length === 0) return false;
      if (dispatch) {
        const allThisLevel = paras.every((p) => p.node.attrs['heading'] === level);
        const next = allThisLevel ? null : level;
        const tr = state.tr;
        for (const p of paras) tr.setNodeAttribute(p.pos, 'heading', next);
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isActive: (state) => {
      const paras = selectedParagraphs(state);
      return paras.length > 0 && paras.every((p) => p.node.attrs['heading'] === level);
    },
    isEnabled: (state) => selectedParagraphs(state).length > 0,
  };
}
