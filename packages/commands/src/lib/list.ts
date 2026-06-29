import type { EditorState, Transaction } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Command } from '@shadow-garden/bapbong-contracts';

export type ListKind = 'bullet' | 'ordered';

/** Reserved numId per kind for editor-authored lists. (Imported lists keep
 *  their own numIds; these are only minted when the user toggles a list on.) */
const NUM_ID: Record<ListKind, string> = { bullet: 'bb-bullet', ordered: 'bb-ordered' };

/** Numbering definition injected into `doc.attrs.numbering` when a list is
 *  first toggled on — three levels each (matches `NumberingDefs` structurally;
 *  built as plain data so this stays in the isomorphic command layer). */
const DEF: Record<ListKind, { key: string; levels: Record<number, { numFmt: string; lvlText: string; start: number }> }> = {
  bullet: {
    key: 'bb-bullet',
    levels: {
      0: { numFmt: 'bullet', lvlText: '•', start: 1 },
      1: { numFmt: 'bullet', lvlText: '◦', start: 1 },
      2: { numFmt: 'bullet', lvlText: '▪', start: 1 },
    },
  },
  ordered: {
    key: 'bb-ordered',
    levels: {
      0: { numFmt: 'decimal', lvlText: '%1.', start: 1 },
      1: { numFmt: 'lowerLetter', lvlText: '%2.', start: 1 },
      2: { numFmt: 'lowerRoman', lvlText: '%3.', start: 1 },
    },
  },
};

/** Top-level-ish paragraphs overlapping the selection (the toggle targets). */
function paragraphsInSelection(state: EditorState): { pos: number; node: ProseMirrorNode }[] {
  const { from, to } = state.selection;
  const out: { pos: number; node: ProseMirrorNode }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === 'paragraph') out.push({ pos, node });
  });
  return out;
}

function listNumId(node: ProseMirrorNode): string | null {
  return (node.attrs['list'] as { numId: string } | null)?.numId ?? null;
}

/**
 * Toggle a bullet / numbered list on the selected paragraphs. If every target
 * is already this kind, it's removed; otherwise they're set to it. Lists are a
 * paragraph attr (`{ numId, level }`) recomputed live by the numbering counter,
 * so markers renumber as you edit. The ordered definition is injected into
 * `doc.attrs.numbering` on first use; the bullet renders from a static marker.
 *
 * (Export writes `w:numPr` for the numId; round-tripping editor-authored lists
 * to a Word-openable marker needs numbering.xml regen on export — a follow-up.)
 */
export function toggleList(kind: ListKind): Command {
  const numId = NUM_ID[kind];
  return {
    name: kind === 'bullet' ? 'bullet-list' : 'ordered-list',
    run(state, dispatch) {
      const paras = paragraphsInSelection(state);
      if (paras.length === 0) return false;
      if (dispatch) {
        const allThisKind = paras.every((p) => listNumId(p.node) === numId);
        let tr: Transaction = state.tr;
        if (allThisKind) {
          for (const p of paras) tr = tr.setNodeAttribute(p.pos, 'list', null);
        } else {
          const defs = { ...((state.doc.attrs['numbering'] as Record<string, unknown> | null) ?? {}) };
          if (!defs[numId]) {
            defs[numId] = DEF[kind];
            tr = tr.setDocAttribute('numbering', defs);
          }
          for (const p of paras) {
            const level = (p.node.attrs['list'] as { level?: number } | null)?.level ?? 0;
            tr = tr.setNodeAttribute(p.pos, 'list', { numId, level });
          }
        }
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isActive(state) {
      const paras = paragraphsInSelection(state);
      return paras.length > 0 && paras.every((p) => listNumId(p.node) === numId);
    },
    isEnabled: (state) => paragraphsInSelection(state).length > 0,
  };
}
