import type {
  EditorPlugin,
  EditorPointerEvent,
  PagePoint,
  PluginContext,
  ResolvedCell,
  ResolvedLayout,
} from '@shadow-garden/bapbong-contracts';

/** The editor state type, taken from the plugin context (no direct PM dep). */
type State = PluginContext['state'];

const EDGE_TOL = 5; // px proximity to a column border to start a resize
const MIN_WIDTH = 24; // px minimum column width

/**
 * Find a column border (a cell's right edge) within {@link EDGE_TOL} of `point`,
 * and the colspan-1 cells of that column (those sharing the right edge). Their
 * shared left edge `leftX` anchors the new width during a drag.
 */
function columnBorderAt(
  layout: ResolvedLayout | null,
  point: PagePoint,
): { cells: ResolvedCell[]; leftX: number } | null {
  const page = layout?.pages[point.pageIndex];
  if (!page?.tables) return null;
  for (const table of page.tables) {
    if (point.y < table.y || point.y > table.y + table.height) continue;
    if (point.x < table.x - EDGE_TOL || point.x > table.x + table.width + EDGE_TOL) continue;
    for (const cell of table.cells) {
      if (cell.colspan !== 1 || point.y < cell.y || point.y > cell.y + cell.height) continue;
      const rightEdge = cell.x + cell.width;
      if (Math.abs(point.x - rightEdge) <= EDGE_TOL) {
        const cells = table.cells.filter((c) => c.colspan === 1 && Math.abs(c.x + c.width - rightEdge) < 1);
        return { cells, leftX: cell.x };
      }
    }
  }
  return null;
}

/** The `table_cell` doc position for a resolved cell (via its first line's
 *  caret slot, walking up to the cell node). */
function cellDocPos(state: State, cell: ResolvedCell): number | null {
  const from = cell.lines[0]?.from;
  if (from == null) return null;
  const $pos = state.doc.resolve(from);
  for (let d = $pos.depth; d > 0; d--) {
    if ($pos.node(d).type.name === 'table_cell') return $pos.before(d);
  }
  return null;
}

interface DragState {
  positions: number[]; // table_cell doc positions of the column being resized
  leftX: number; // page-local left edge of the column (stable during drag)
  lastWidth: number | null;
}

/**
 * Internal plugin: drag a table column's right border to resize it. Hovering a
 * border shows a `col-resize` cursor; a drag rewrites the column cells'
 * `colwidth` attr live (the layout engine sizes columns from it). Consecutive
 * drag transactions coalesce into one undo step via history time-grouping.
 *
 * Built on the editor's pointer hook + `ctx.layout`/`ctx.setCursor` — no UI,
 * no schema; purely geometry + transactions.
 */
export function tableResizePlugin(): EditorPlugin {
  let ctx: PluginContext | null = null;
  let drag: DragState | null = null;

  const applyWidth = (c: PluginContext, pointX: number): void => {
    if (!drag) return;
    const width = Math.max(MIN_WIDTH, Math.round(pointX - drag.leftX));
    if (width === drag.lastWidth) return;
    drag.lastWidth = width;
    let tr = c.state.tr;
    for (const pos of drag.positions) {
      if (c.state.doc.nodeAt(pos)?.type.name === 'table_cell') {
        tr = tr.setNodeAttribute(pos, 'colwidth', [width]);
      }
    }
    if (tr.docChanged) c.dispatch(tr);
  };

  return {
    name: 'table-resize',
    setup(c) {
      ctx = c;
    },
    onPointer(ev: EditorPointerEvent): boolean {
      const c = ctx;
      if (!c) return false;

      if (ev.type === 'move') {
        if (drag) {
          if (ev.point) applyWidth(c, ev.point.x);
          return true; // claim: suppress the editor's selection drag
        }
        if (ev.buttons === 0) {
          const onBorder = ev.point ? columnBorderAt(c.layout, ev.point) : null;
          c.setCursor(onBorder ? 'col-resize' : null);
        }
        return false; // hover only sets the cursor; don't claim
      }

      if (ev.type === 'down') {
        if (ev.buttons !== 1 || !ev.point) return false;
        const border = columnBorderAt(c.layout, ev.point);
        if (!border) return false;
        const positions = border.cells
          .map((cell) => cellDocPos(c.state, cell))
          .filter((p): p is number => p != null);
        if (positions.length === 0) return false;
        drag = { positions, leftX: border.leftX, lastWidth: null };
        c.setCursor('col-resize');
        return true; // claim → the editor captures the pointer for us
      }

      if (ev.type === 'up') {
        if (!drag) return false;
        drag = null;
        c.setCursor(null);
        return true;
      }

      return false;
    },
  };
}
