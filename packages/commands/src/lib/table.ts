import type { EditorState } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Command } from '@shadow-garden/bapbong-contracts';

/** A located table cell: its node + the doc position just before it. */
export interface CellInfo {
  pos: number;
  node: ProseMirrorNode;
}

/** Cell attributes a caller may set (mirrors the `table_cell` schema attrs). */
export interface CellAttrs {
  background: string | null;
  vAlign: 'center' | 'bottom' | null;
  colwidth: number[] | null;
  borders: unknown;
}

/** The `table_cell` containing the selection head, or null when outside a table.
 *  The cell-properties dialog and context menu use this to know their target. */
export function cellAt(state: EditorState): CellInfo | null {
  const $head = state.selection.$head;
  for (let d = $head.depth; d > 0; d--) {
    const node = $head.node(d);
    if (node.type.name === 'table_cell') return { pos: $head.before(d), node };
  }
  return null;
}

/**
 * Merge `attrs` onto the `table_cell` at `pos` — the generic primitive the cell
 * properties dialog dispatches. No-op (returns false) if `pos` isn't a cell.
 */
export function setCellAttrs(pos: number, attrs: Partial<CellAttrs>): Command {
  return {
    name: 'cell-attrs',
    run(state, dispatch) {
      const cell = state.doc.nodeAt(pos);
      if (!cell || cell.type.name !== 'table_cell') return false;
      if (dispatch) {
        const tr = state.tr;
        for (const [key, value] of Object.entries(attrs)) tr.setNodeAttribute(pos, key, value);
        dispatch(tr);
      }
      return true;
    },
  };
}

/** Set the same attrs on several cells (by doc position) in one transaction —
 *  the cell-properties dialog applied across a selected block. */
export function setCellsAttrs(positions: number[], attrs: Partial<CellAttrs>): Command {
  return {
    name: 'cells-attrs',
    run(state, dispatch) {
      const cells = positions.filter((pos) => state.doc.nodeAt(pos)?.type.name === 'table_cell');
      if (cells.length === 0) return false;
      if (dispatch) {
        const tr = state.tr;
        for (const pos of cells) {
          for (const [key, value] of Object.entries(attrs)) tr.setNodeAttribute(pos, key, value);
        }
        dispatch(tr);
      }
      return true;
    },
  };
}

/** Set the current cell's background fill (`null` clears it). */
export function setCellBackground(color: string | null): Command {
  return {
    name: 'cell-background',
    run: (state, dispatch) => {
      const cell = cellAt(state);
      return cell ? setCellAttrs(cell.pos, { background: color }).run(state, dispatch) : false;
    },
    isEnabled: (state) => cellAt(state) != null,
  };
}

/**
 * Set one cell's column width in px. The L2 column-resize plugin owns the
 * geometry (which cells share a column) and composes this across each of them;
 * on its own it sets a single cell. (colspan>1 cells carry a width-per-spanned-
 * column array — here we set a single width.)
 */
export function setColumnWidth(cellPos: number, width: number): Command {
  return {
    name: 'column-width',
    run: (state, dispatch) => setCellAttrs(cellPos, { colwidth: [width] }).run(state, dispatch),
  };
}
