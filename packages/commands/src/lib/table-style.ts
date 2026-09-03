import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import {
  cellStyleLayer,
  type Command,
  type ResolvedTableStyle,
  type TableLook,
  type TableStyleFont,
  type TableStyleSheet,
} from '@shadow-garden/bapbong-contracts';

/**
 * Live table theming, the command half: applying a style or flipping a
 * tblLook gate is one transaction — the table's styleId/look attrs (which
 * the layout reads against doc.attrs.tableStyles), plus a FONT-MARK rewrite.
 *
 * The rewrite exists because run fonts are still baked into marks (the mark
 * model has no explicit-off, so the layout's font layer is parked — see
 * tableToFlow). Rule, per field per run: when the run's current value IS the
 * old style's contribution, it moves with the style; when the user diverged,
 * their formatting stays — Word's direct-formatting-wins, reconstructed
 * after the fact. A run whose value merely COINCIDES with the old style's is
 * indistinguishable from a baked one and moves too: the known limitation of
 * rewriting where layers were flattened.
 */

/** Word's tblLook when a table declares none — 0x04A0. */
export const WORD_DEFAULT_TABLE_LOOK: TableLook = {
  firstRow: true,
  lastRow: false,
  firstCol: true,
  lastCol: false,
  hBand: true,
  vBand: false,
};

/** The table enclosing the selection head, or null. */
export function tableAt(
  state: EditorState,
): { pos: number; node: ProseMirrorNode } | null {
  const $head = state.selection.$head;
  for (let d = $head.depth; d > 0; d--) {
    const node = $head.node(d);
    if (node.type.name === 'table') return { pos: $head.before(d), node };
  }
  return null;
}

/** What the panel renders: the enclosing table's style pair, with look
 *  defaulted the way Word defaults it. Null outside a table. */
export function currentTableStyle(
  state: EditorState,
): { styleId: string | null; look: TableLook } | null {
  const t = tableAt(state);
  if (!t) return null;
  return {
    styleId: (t.node.attrs['styleId'] as string | null) ?? null,
    look: (t.node.attrs['look'] as TableLook | null) ?? WORD_DEFAULT_TABLE_LOOK,
  };
}

export interface ApplyTableStyleOptions {
  /** The style to apply, or null to remove the table's style. */
  styleId: string | null;
  /** Gates to apply along with it; absent keeps the table's current look. */
  look?: TableLook;
  /** The style's resolved definition, for a style the document's sheet does
   *  not hold yet (a gallery built-in) — injected into doc.attrs.tableStyles
   *  in the same transaction. Ignored when the sheet already has the id. */
  style?: ResolvedTableStyle;
}

/** Apply (or remove) a table style on the enclosing table. */
export function applyTableStyle(opts: ApplyTableStyleOptions): Command {
  return {
    name: 'apply-table-style',
    title: 'Table style',
    run(state, dispatch) {
      const t = tableAt(state);
      if (!t) return false;
      const look =
        opts.look ??
        (t.node.attrs['look'] as TableLook | null) ??
        WORD_DEFAULT_TABLE_LOOK;
      return restyle(
        state,
        dispatch,
        t.pos,
        t.node,
        opts.styleId,
        look,
        opts.style,
      );
    },
    isActive(state) {
      const t = tableAt(state);
      return !!t && t.node.attrs['styleId'] === opts.styleId;
    },
    isEnabled(state) {
      return tableAt(state) !== null;
    },
  };
}

/** Flip tblLook gates on the enclosing styled table (merge semantics: only
 *  the fields in `patch` change). */
export function setTableLook(patch: Partial<TableLook>): Command {
  return {
    name: 'set-table-look',
    title: 'Table style options',
    run(state, dispatch) {
      const t = tableAt(state);
      const styleId = t?.node.attrs['styleId'] as string | null;
      if (!t || !styleId) return false;
      const look = {
        ...((t.node.attrs['look'] as TableLook | null) ??
          WORD_DEFAULT_TABLE_LOOK),
        ...patch,
      };
      return restyle(state, dispatch, t.pos, t.node, styleId, look);
    },
    isEnabled(state) {
      return !!(tableAt(state)?.node.attrs['styleId'] as string | null);
    },
  };
}

function restyle(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  tablePos: number,
  table: ProseMirrorNode,
  newId: string | null,
  newLook: TableLook,
  inject?: ResolvedTableStyle,
): boolean {
  const sheet = (state.doc.attrs['tableStyles'] ?? {}) as TableStyleSheet;
  const oldId = table.attrs['styleId'] as string | null;
  const oldLook =
    (table.attrs['look'] as TableLook | null) ?? WORD_DEFAULT_TABLE_LOOK;
  const oldStyle = oldId ? sheet[oldId] : undefined;
  const newStyle = newId ? (sheet[newId] ?? inject) : undefined;
  // An id neither the sheet nor the caller can define would style nothing.
  if (newId && !newStyle) return false;
  if (!dispatch) return true;
  const tr = state.tr;
  if (newId && !sheet[newId] && inject)
    tr.setDocAttribute('tableStyles', { ...sheet, [newId]: inject });
  tr.setNodeAttribute(tablePos, 'styleId', newId);
  tr.setNodeAttribute(tablePos, 'look', newId ? { ...newLook } : null);
  rewriteFontMarks(
    tr,
    state,
    tablePos,
    table,
    oldStyle,
    oldLook,
    newStyle,
    newLook,
  );
  dispatch(tr.scrollIntoView());
  return true;
}

/** One font field's mark surgery: `read` the run's current value, and when it
 *  equals the OLD style's contribution but not the new one, move it. */
interface FontField {
  markName: string;
  of: (f: TableStyleFont | undefined) => string | number | boolean | undefined;
  read: (
    marks: readonly ProseMirrorNode['marks'][number][],
  ) => string | number | boolean | undefined;
  create: (
    state: EditorState,
    value: string | number | boolean,
  ) => ProseMirrorNode['marks'][number] | null;
}

const FIELDS: FontField[] = [
  {
    markName: 'strong',
    of: (f) => f?.bold === true || undefined,
    read: (marks) => marks.some((m) => m.type.name === 'strong') || undefined,
    create: (state) => state.schema.marks['strong']?.create() ?? null,
  },
  {
    markName: 'em',
    of: (f) => f?.italic === true || undefined,
    read: (marks) => marks.some((m) => m.type.name === 'em') || undefined,
    create: (state) => state.schema.marks['em']?.create() ?? null,
  },
  {
    markName: 'textColor',
    of: (f) => f?.color,
    read: (marks) =>
      marks.find((m) => m.type.name === 'textColor')?.attrs['color'] as
        | string
        | undefined,
    create: (state, v) =>
      state.schema.marks['textColor']?.create({ color: v }) ?? null,
  },
  {
    markName: 'fontFamily',
    of: (f) => f?.family,
    read: (marks) =>
      marks.find((m) => m.type.name === 'fontFamily')?.attrs['family'] as
        | string
        | undefined,
    create: (state, v) =>
      state.schema.marks['fontFamily']?.create({ family: v }) ?? null,
  },
  {
    markName: 'fontSize',
    of: (f) => f?.sizePt,
    read: (marks) =>
      marks.find((m) => m.type.name === 'fontSize')?.attrs['size'] as
        | number
        | undefined,
    create: (state, v) =>
      state.schema.marks['fontSize']?.create({ size: v }) ?? null,
  },
];

function rewriteFontMarks(
  tr: Transaction,
  state: EditorState,
  tablePos: number,
  table: ProseMirrorNode,
  oldStyle: ResolvedTableStyle | undefined,
  oldLook: TableLook,
  newStyle: ResolvedTableStyle | undefined,
  newLook: TableLook,
): void {
  if (!oldStyle && !newStyle) return;
  const rowCount = table.childCount;
  let colCount = 0;
  table.forEach((row) => {
    let n = 0;
    row.forEach((c) => (n += Number(c.attrs['colspan']) || 1));
    colCount = Math.max(colCount, n);
  });
  let rowIdx = 0;
  table.forEach((row, rowOffset) => {
    const rowPos = tablePos + 1 + rowOffset;
    const header = row.attrs['header'] === true;
    let colIdx = 0;
    row.forEach((cell, cellOffset) => {
      const colspan = Number(cell.attrs['colspan']) || 1;
      // The SAME cell-position walk the layout uses (colspan advances the
      // grid column; rowspan is not tracked — consistency with tableToFlow
      // matters more than span perfection).
      const pos = {
        row: rowIdx,
        rowCount,
        col: colIdx,
        colspan,
        colCount,
        header,
      };
      const oldFont = oldStyle
        ? cellStyleLayer(oldStyle, oldLook, pos).font
        : undefined;
      const newFont = newStyle
        ? cellStyleLayer(newStyle, newLook, pos).font
        : undefined;
      colIdx += colspan;
      const cellPos = rowPos + 1 + cellOffset;
      cell.descendants((node, rel) => {
        if (!node.isText) return true;
        const from = cellPos + 1 + rel;
        const to = from + node.nodeSize;
        for (const field of FIELDS) {
          const oldV = field.of(oldFont);
          const newV = field.of(newFont);
          if (oldV === newV) continue;
          if (field.read(node.marks) !== oldV) continue; // user diverged
          const type = state.schema.marks[field.markName];
          if (!type) continue;
          if (oldV !== undefined) tr.removeMark(from, to, type);
          if (newV !== undefined) {
            const mark = field.create(state, newV);
            if (mark) tr.addMark(from, to, mark);
          }
        }
        return true;
      });
    });
    rowIdx++;
  });
}
