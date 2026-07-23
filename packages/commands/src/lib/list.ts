import type { EditorState, Transaction } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Command } from '@shadow-garden/bapbong-contracts';

export type ListKind = 'bullet' | 'ordered';

/** One nesting level of a list preset (matches `NumberingLevelDef`). */
export interface ListPresetLevel {
  numFmt: string;
  lvlText: string;
  start: number;
}

/** A marker style offered in the list dropdowns. `samples` are the literal
 *  markers shown in the preview card, one per nesting level. */
export interface ListPreset {
  id: string;
  /** numId minted for paragraphs using this preset (also the def key). The
   *  first preset keeps the legacy `bb-bullet`/`bb-ordered` ids so documents
   *  authored before presets existed keep their numbering. */
  numId: string;
  levels: [ListPresetLevel, ListPresetLevel, ListPresetLevel];
  samples: [string, string, string];
}

const bulletPreset = (
  id: string,
  numId: string,
  glyphs: [string, string, string],
): ListPreset => ({
  id,
  numId,
  levels: [
    { numFmt: 'bullet', lvlText: glyphs[0], start: 1 },
    { numFmt: 'bullet', lvlText: glyphs[1], start: 1 },
    { numFmt: 'bullet', lvlText: glyphs[2], start: 1 },
  ],
  samples: glyphs,
});

const lvl = (numFmt: string, lvlText: string): ListPresetLevel => ({
  numFmt,
  lvlText,
  start: 1,
});

const orderedPreset = (
  id: string,
  numId: string,
  levels: [ListPresetLevel, ListPresetLevel, ListPresetLevel],
  samples: [string, string, string],
): ListPreset => ({ id, numId, levels, samples });

/** The marker styles offered per kind (mirrors the Google Docs pickers). */
const PRESETS: Record<ListKind, ListPreset[]> = {
  bullet: [
    bulletPreset('disc', 'bb-bullet', ['•', '◦', '▪']),
    bulletPreset('diamond', 'bb-bullet-diamond', ['❖', '➢', '▪']),
    bulletPreset('square', 'bb-bullet-square', ['■', '□', '▪']),
    bulletPreset('arrow', 'bb-bullet-arrow', ['➜', '◆', '•']),
    bulletPreset('star', 'bb-bullet-star', ['★', '○', '▪']),
    bulletPreset('chevron', 'bb-bullet-chevron', ['➤', '◦', '▪']),
  ],
  ordered: [
    orderedPreset(
      'decimal',
      'bb-ordered',
      [
        lvl('decimal', '%1.'),
        lvl('lowerLetter', '%2.'),
        lvl('lowerRoman', '%3.'),
      ],
      ['1.', 'a.', 'i.'],
    ),
    orderedPreset(
      'paren',
      'bb-ordered-paren',
      [
        lvl('decimal', '%1)'),
        lvl('lowerLetter', '%2)'),
        lvl('lowerRoman', '%3)'),
      ],
      ['1)', 'a)', 'i)'],
    ),
    orderedPreset(
      'multilevel',
      'bb-ordered-multilevel',
      [
        lvl('decimal', '%1.'),
        lvl('decimal', '%1.%2.'),
        lvl('decimal', '%1.%2.%3.'),
      ],
      ['1.', '1.1.', '1.1.1.'],
    ),
    orderedPreset(
      'upper',
      'bb-ordered-upper',
      [
        lvl('upperLetter', '%1.'),
        lvl('lowerLetter', '%2.'),
        lvl('lowerRoman', '%3.'),
      ],
      ['A.', 'a.', 'i.'],
    ),
    orderedPreset(
      'roman',
      'bb-ordered-roman',
      [
        lvl('upperRoman', '%1.'),
        lvl('upperLetter', '%2.'),
        lvl('decimal', '%3.'),
      ],
      ['I.', 'A.', '1.'],
    ),
    orderedPreset(
      'zero',
      'bb-ordered-zero',
      [
        lvl('decimalZero', '%1.'),
        lvl('lowerLetter', '%2.'),
        lvl('lowerRoman', '%3.'),
      ],
      ['01.', 'a.', 'i.'],
    ),
  ],
};

/** The marker presets a list dropdown offers for `kind`. */
export function listPresets(kind: ListKind): readonly ListPreset[] {
  return PRESETS[kind];
}

type Defs = Record<
  string,
  { key: string; levels: Record<number, ListPresetLevel> }
>;

/** The numbering-def entry a preset injects into `doc.attrs.numbering`. */
function defOf(preset: ListPreset): Defs[string] {
  return {
    key: preset.numId,
    levels: { 0: preset.levels[0], 1: preset.levels[1], 2: preset.levels[2] },
  };
}

/** Top-level-ish paragraphs overlapping the selection (the toggle targets). */
function paragraphsInSelection(
  state: EditorState,
): { pos: number; node: ProseMirrorNode }[] {
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

/** Which kind a paragraph's list is: editor presets by numId prefix, imported
 *  numIds by their definition's level-0 format (bullet vs numbered). */
function kindOf(node: ProseMirrorNode, defs: Defs | null): ListKind | null {
  const numId = listNumId(node);
  if (!numId) return null;
  if (numId.startsWith('bb-bullet')) return 'bullet';
  if (numId.startsWith('bb-ordered')) return 'ordered';
  const fmt = defs?.[numId]?.levels?.[0]?.numFmt;
  if (!fmt) return null;
  return fmt === 'bullet' ? 'bullet' : 'ordered';
}

/** Apply `preset` to the given paragraphs (keeping each item's nesting
 *  level), injecting its numbering definition on first use. */
function applyPreset(
  state: EditorState,
  tr: Transaction,
  paras: { pos: number; node: ProseMirrorNode }[],
  preset: ListPreset,
): Transaction {
  const defs = { ...((state.doc.attrs['numbering'] as Defs | null) ?? {}) };
  if (!defs[preset.numId]) {
    defs[preset.numId] = defOf(preset);
    tr = tr.setDocAttribute('numbering', defs);
  }
  for (const p of paras) {
    const level =
      (p.node.attrs['list'] as { level?: number } | null)?.level ?? 0;
    tr = tr.setNodeAttribute(p.pos, 'list', { numId: preset.numId, level });
  }
  return tr;
}

/**
 * Toggle a bullet / numbered list on the selected paragraphs. If every target
 * is already this kind (any preset, or an imported list of this kind), the
 * list is removed; otherwise the kind's default preset is applied. Lists are a
 * paragraph attr (`{ numId, level }`) recomputed live by the numbering counter,
 * so markers renumber as you edit.
 *
 * (Export writes `w:numPr` for the numId; round-tripping editor-authored lists
 * to a Word-openable marker needs numbering.xml regen on export — a follow-up.)
 */
export function toggleList(kind: ListKind): Command {
  return {
    name: kind === 'bullet' ? 'bullet-list' : 'ordered-list',
    run(state, dispatch) {
      const paras = paragraphsInSelection(state);
      if (paras.length === 0) return false;
      if (dispatch) {
        const defs = (state.doc.attrs['numbering'] as Defs | null) ?? null;
        const allThisKind = paras.every((p) => kindOf(p.node, defs) === kind);
        let tr: Transaction = state.tr;
        if (allThisKind) {
          for (const p of paras) tr = tr.setNodeAttribute(p.pos, 'list', null);
        } else {
          tr = applyPreset(state, tr, paras, PRESETS[kind][0]);
        }
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isActive(state) {
      const paras = paragraphsInSelection(state);
      const defs = (state.doc.attrs['numbering'] as Defs | null) ?? null;
      return (
        paras.length > 0 && paras.every((p) => kindOf(p.node, defs) === kind)
      );
    },
    isEnabled: (state) => paragraphsInSelection(state).length > 0,
  };
}

/** Set the selected paragraphs to a specific marker preset (switching style in
 *  place — never toggling off; the plain button does that). */
export function applyListPreset(kind: ListKind, presetId: string): Command {
  const preset =
    PRESETS[kind].find((p) => p.id === presetId) ?? PRESETS[kind][0];
  return {
    name: `${kind === 'bullet' ? 'bullet-list' : 'ordered-list'}:${preset.id}`,
    run(state, dispatch) {
      const paras = paragraphsInSelection(state);
      if (paras.length === 0) return false;
      if (dispatch)
        dispatch(applyPreset(state, state.tr, paras, preset).scrollIntoView());
      return true;
    },
    isEnabled: (state) => paragraphsInSelection(state).length > 0,
  };
}

/** The preset id every selected list paragraph shares, or null (mixed styles,
 *  imported numbering, or no list) — drives the dropdown's selected card. */
export function activeListPresetId(
  state: EditorState,
  kind: ListKind,
): string | null {
  const paras = paragraphsInSelection(state);
  if (paras.length === 0) return null;
  const ids = new Set(paras.map((p) => listNumId(p.node)));
  if (ids.size !== 1) return null;
  const [only] = ids;
  return PRESETS[kind].find((p) => p.numId === only)?.id ?? null;
}
