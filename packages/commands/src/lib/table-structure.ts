import type { EditorState } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Command } from '@shadow-garden/bapbong-contracts';

/** Where the selection sits within a table (for row/column insertion). Cells are
 *  assumed colspan-1 grids — spanning cells are a best-effort edge case. */
interface TableContext {
  tablePos: number;
  table: ProseMirrorNode;
  rowPos: number;
  row: ProseMirrorNode;
  rowIndex: number;
  cellIndex: number;
}

function tableContext(state: EditorState): TableContext | null {
  const $head = state.selection.$head;
  let cellDepth = -1;
  for (let d = $head.depth; d > 0; d--) {
    if ($head.node(d).type.name === 'table_cell') {
      cellDepth = d;
      break;
    }
  }
  if (cellDepth < 2) return null; // need table_cell → table_row → table
  return {
    tablePos: $head.before(cellDepth - 2),
    table: $head.node(cellDepth - 2),
    rowPos: $head.before(cellDepth - 1),
    row: $head.node(cellDepth - 1),
    rowIndex: $head.index(cellDepth - 2),
    cellIndex: $head.index(cellDepth - 1),
  };
}

/** A fresh empty cell carrying the given column width (px) / colspan. */
function emptyCell(
  state: EditorState,
  attrs: { colwidth?: number[] | null; colspan?: number },
): ProseMirrorNode {
  const cell = state.schema.nodes['table_cell'];
  const paragraph = state.schema.nodes['paragraph'];
  return cell.create(attrs, paragraph.create());
}

/** Insert a blank row above/below the current row, mirroring its column widths. */
export function insertRow(below: boolean): Command {
  return {
    name: below ? 'insert-row-below' : 'insert-row-above',
    run(state, dispatch) {
      const cx = tableContext(state);
      if (!cx) return false;
      if (dispatch) {
        const cells: ProseMirrorNode[] = [];
        cx.row.forEach((cell) =>
          cells.push(
            emptyCell(state, {
              colwidth: cell.attrs['colwidth'] as number[] | null,
              colspan: cell.attrs['colspan'] as number,
            }),
          ),
        );
        const newRow = state.schema.nodes['table_row'].create(null, cells);
        const at = below ? cx.rowPos + cx.row.nodeSize : cx.rowPos;
        dispatch(state.tr.insert(at, newRow).scrollIntoView());
      }
      return true;
    },
    isEnabled: (state) => tableContext(state) != null,
  };
}

/** Insert a blank column left/right of the current one (in every row), copying
 *  the current column's width. */
export function insertColumn(right: boolean): Command {
  return {
    name: right ? 'insert-column-right' : 'insert-column-left',
    run(state, dispatch) {
      const cx = tableContext(state);
      if (!cx) return false;
      if (dispatch) {
        const targetIndex = right ? cx.cellIndex + 1 : cx.cellIndex;
        const curCell = cx.row.maybeChild(
          Math.min(cx.cellIndex, cx.row.childCount - 1),
        );
        const colwidth =
          (curCell?.attrs['colwidth'] as number[] | null) ?? null;

        // Doc position where the target cell starts in each row (or the row's
        // content end when the row is shorter), in the original document.
        const inserts: number[] = [];
        let rowStart = cx.tablePos + 1; // first row node start
        cx.table.forEach((row) => {
          let pos = rowStart + 1; // inside the row, before its first cell
          row.forEach((cell, offset, index) => {
            if (index < targetIndex) pos += cell.nodeSize;
          });
          inserts.push(pos);
          rowStart += row.nodeSize;
        });

        let tr = state.tr;
        for (const pos of inserts) {
          tr = tr.insert(tr.mapping.map(pos), emptyCell(state, { colwidth }));
        }
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isEnabled: (state) => tableContext(state) != null,
  };
}

/** Delete the current row. Removing the last row deletes the whole table.
 *  (Spanning cells are best-effort, like {@link insertRow}.) */
export function deleteRow(): Command {
  return {
    name: 'delete-row',
    run(state, dispatch) {
      const cx = tableContext(state);
      if (!cx) return false;
      if (dispatch) {
        const tr =
          cx.table.childCount <= 1
            ? state.tr.delete(cx.tablePos, cx.tablePos + cx.table.nodeSize)
            : state.tr.delete(cx.rowPos, cx.rowPos + cx.row.nodeSize);
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isEnabled: (state) => tableContext(state) != null,
  };
}

/** Delete the current column (the cell at this index in every row). Removing
 *  the last column deletes the whole table. (Spanning cells are best-effort.) */
export function deleteColumn(): Command {
  return {
    name: 'delete-column',
    run(state, dispatch) {
      const cx = tableContext(state);
      if (!cx) return false;
      if (dispatch) {
        if (cx.row.childCount <= 1) {
          dispatch(
            state.tr
              .delete(cx.tablePos, cx.tablePos + cx.table.nodeSize)
              .scrollIntoView(),
          );
          return true;
        }
        // The doc range of the target cell in each row (original positions).
        const dels: { from: number; to: number }[] = [];
        let rowStart = cx.tablePos + 1; // first row node start
        cx.table.forEach((row) => {
          let pos = rowStart + 1; // inside the row, before its first cell
          row.forEach((cell, _offset, index) => {
            if (index < cx.cellIndex) pos += cell.nodeSize;
          });
          const cell = row.maybeChild(cx.cellIndex);
          if (cell) dels.push({ from: pos, to: pos + cell.nodeSize });
          rowStart += row.nodeSize;
        });
        let tr = state.tr;
        for (const d of dels) {
          tr = tr.delete(tr.mapping.map(d.from), tr.mapping.map(d.to));
        }
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isEnabled: (state) => tableContext(state) != null,
  };
}

/** A selected cell within a block grid (matches the editor's `SelectedCell` —
 *  passed structurally so this stays in the isomorphic command layer). */
export interface MergeCell {
  pos: number;
  row: number;
  col: number;
}

/**
 * Merge a rectangular block of cells into the top-left one: it gains
 * `colspan = cols` / `rowspan = rows`, the rest are removed, and their
 * non-empty content is appended (so nothing is lost). The spanned-over cells
 * disappear from their rows — exactly the shape the layout's occupancy grid
 * expects for an imported merged cell. Best-effort if the block already
 * contains spanning cells.
 */
export function mergeCells(
  cells: MergeCell[],
  rows: number,
  cols: number,
): Command {
  return {
    name: 'merge-cells',
    run(state, dispatch) {
      if (cells.length < 2) return false;
      const located = cells
        .map((c) => ({ c, node: state.doc.nodeAt(c.pos) }))
        .filter(
          (x): x is { c: MergeCell; node: ProseMirrorNode } =>
            x.node?.type.name === 'table_cell',
        );
      if (located.length < 2) return false;
      const tl = located.find((x) => x.c.row === 0 && x.c.col === 0);
      if (!tl) return false;
      if (dispatch) {
        const cellType = state.schema.nodes['table_cell'];
        // Merged width = the top-row cells' widths concatenated (best-effort).
        let colwidth: number[] | null = [];
        for (const x of located
          .filter((x) => x.c.row === 0)
          .sort((a, b) => a.c.col - b.c.col)) {
          const cw = x.node.attrs['colwidth'] as number[] | null;
          if (!cw) {
            colwidth = null;
            break;
          }
          colwidth.push(...cw);
        }
        // Top-left keeps its blocks; append non-empty blocks from the others.
        const content: ProseMirrorNode[] = [];
        tl.node.forEach((child) => content.push(child));
        for (const x of located) {
          if (x === tl) continue;
          x.node.forEach((child) => {
            if (!child.isTextblock || child.textContent.trim().length > 0)
              content.push(child);
          });
        }
        const merged = cellType.create(
          { ...tl.node.attrs, colspan: cols, rowspan: rows, colwidth },
          content,
        );
        // Apply back-to-front so the original positions stay valid.
        const ops = located
          .map((x) => ({
            from: x.c.pos,
            to: x.c.pos + x.node.nodeSize,
            node: x === tl ? merged : null,
          }))
          .sort((a, b) => b.from - a.from);
        let tr = state.tr;
        for (const op of ops)
          tr = op.node
            ? tr.replaceWith(op.from, op.to, op.node)
            : tr.delete(op.from, op.to);
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
  };
}

/** Delete the entire table the selection sits in. */
export function deleteTable(): Command {
  return {
    name: 'delete-table',
    run(state, dispatch) {
      const cx = tableContext(state);
      if (!cx) return false;
      if (dispatch)
        dispatch(
          state.tr
            .delete(cx.tablePos, cx.tablePos + cx.table.nodeSize)
            .scrollIntoView(),
        );
      return true;
    },
    isEnabled: (state) => tableContext(state) != null,
  };
}
