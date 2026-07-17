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

/** The six named paragraph styles the toolbar dropdown offers. */
export type ParagraphStyleKey = 'normal' | 'title' | 'subtitle' | 'h1' | 'h2' | 'h3';

/** heading/styleId attr pair for each dropdown key — the single source of the
 *  "styleId set ⇒ heading null" invariant. */
const STYLE_ATTRS: Record<ParagraphStyleKey, { heading: number | null; styleId: string | null }> = {
  normal: { heading: null, styleId: null },
  title: { heading: null, styleId: 'Title' },
  subtitle: { heading: null, styleId: 'Subtitle' },
  h1: { heading: 1, styleId: null },
  h2: { heading: 2, styleId: null },
  h3: { heading: 3, styleId: null },
};

/**
 * Set (not toggle — dropdowns state absolutes) the named paragraph style on
 * every selected paragraph: Normal / Title / Subtitle / Heading 1–3. The only
 * writer of `styleId`, so heading and styleId can never both be set.
 */
export function setParagraphStyle(key: ParagraphStyleKey): Command {
  const attrs = STYLE_ATTRS[key];
  return {
    name: `paragraph-style-${key}`,
    run(state, dispatch) {
      const paras = selectedParagraphs(state);
      if (paras.length === 0) return false;
      if (dispatch) {
        const tr = state.tr;
        for (const p of paras) {
          tr.setNodeAttribute(p.pos, 'heading', attrs.heading);
          tr.setNodeAttribute(p.pos, 'styleId', attrs.styleId);
        }
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isActive: (state) => activeParagraphStyle(state) === key,
    isEnabled: (state) => selectedParagraphs(state).length > 0,
  };
}

/** The dropdown's current value: the shared style of every selected paragraph,
 *  or null when the selection is empty / mixed / on an unlisted style (H4–H6).
 *  O(selection) — it only walks the selected range (called per transaction). */
export function activeParagraphStyle(state: EditorState): ParagraphStyleKey | null {
  const paras = selectedParagraphs(state);
  if (paras.length === 0) return null;
  const keyOf = (node: (typeof paras)[number]['node']): ParagraphStyleKey | null => {
    const styleId = node.attrs['styleId'];
    if (styleId === 'Title') return 'title';
    if (styleId === 'Subtitle') return 'subtitle';
    const h = node.attrs['heading'];
    if (h === 1 || h === 2 || h === 3) return `h${h}` as ParagraphStyleKey;
    return h == null ? 'normal' : null; // h4–h6: real but not a dropdown entry
  };
  const first = keyOf(paras[0].node);
  return paras.every((p) => keyOf(p.node) === first) ? first : null;
}
