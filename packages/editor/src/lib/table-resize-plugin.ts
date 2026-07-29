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

/** An interior column border (shared by a left column and a right column). The
 *  table's outer edges have no neighbour, so they aren't resizable — dragging
 *  redistributes width between the two adjacent columns, keeping the table (and
 *  its alignment) a fixed total width. */
interface BorderHit {
  pageIndex: number;
  borderX: number; // current border x (page-local)
  leftX: number; // left column's left edge
  rightX: number; // right column's right edge
  leftCells: ResolvedCell[];
  rightCells: ResolvedCell[];
  tableY: number;
  tableHeight: number;
}

const near = (a: number, b: number) => Math.abs(a - b) < 1;

function borderAt(
  layout: ResolvedLayout | null,
  point: PagePoint,
): BorderHit | null {
  const page = layout?.pages[point.pageIndex];
  if (!page?.tables) return null;
  for (const table of page.tables) {
    if (point.y < table.y || point.y > table.y + table.height) continue;
    if (
      point.x < table.x - EDGE_TOL ||
      point.x > table.x + table.width + EDGE_TOL
    )
      continue;
    const simple = table.cells.filter((c) => c.colspan === 1);
    for (const cell of simple) {
      if (point.y < cell.y || point.y > cell.y + cell.height) continue;
      const rightEdge = cell.x + cell.width;
      if (Math.abs(point.x - rightEdge) > EDGE_TOL) continue;
      const leftCells = simple.filter((c) => near(c.x + c.width, rightEdge));
      const rightCells = simple.filter((c) => near(c.x, rightEdge));
      if (rightCells.length === 0) return null; // outer-right edge → not resizable
      const leftX = Math.min(...leftCells.map((c) => c.x));
      const rightX = Math.max(...rightCells.map((c) => c.x + c.width));
      return {
        pageIndex: point.pageIndex,
        borderX: rightEdge,
        leftX,
        rightX,
        leftCells,
        rightCells,
        tableY: table.y,
        tableHeight: table.height,
      };
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
  hit: BorderHit;
  borderX: number; // current (clamped) border position
}

/**
 * Internal plugin: drag a table column's interior border to resize it. Hovering
 * a border shows a `col-resize` cursor; during the drag only a vertical guide
 * follows the cursor (no re-layout); on drop it rewrites just the two adjacent
 * columns' `colwidth` (left +Δ, right −Δ) in one transaction — so the table's
 * total width stays fixed and the change is a single undo step.
 *
 * Built on the editor's pointer hook + `ctx.layout`/`setCursor`/`setGuide` — no
 * UI, no schema; geometry + one transaction.
 */
export function tableResizePlugin(): EditorPlugin {
  let ctx: PluginContext | null = null;
  let drag: DragState | null = null;

  const showGuide = (c: PluginContext, hit: BorderHit, x: number): void =>
    c.setGuide({
      pageIndex: hit.pageIndex,
      x,
      y: hit.tableY,
      height: hit.tableHeight,
    });

  const commit = (c: PluginContext, d: DragState): void => {
    const { hit, borderX } = d;
    if (near(borderX, hit.borderX)) return; // no move → nothing to commit
    const leftWidth = Math.round(borderX - hit.leftX);
    const rightWidth = Math.round(hit.rightX - borderX);
    const state = c.state;
    let tr = state.tr;
    const setCol = (cell: ResolvedCell, width: number) => {
      const pos = cellDocPos(state, cell);
      if (pos != null && state.doc.nodeAt(pos)?.type.name === 'table_cell') {
        tr = tr.setNodeAttribute(pos, 'colwidth', [width]);
      }
    };
    for (const cell of hit.leftCells) setCol(cell, leftWidth);
    for (const cell of hit.rightCells) setCol(cell, rightWidth);
    if (tr.docChanged) c.dispatch(tr);
  };

  return {
    name: 'table-resize',
    setup(c) {
      ctx = c;
      // Clear what this plugin can leave on screen (a column-resize guide
      // mid-drag, the col-resize cursor) — see image-resize for the rule.
      return () => {
        c.setGuide(null);
        c.setCursor(null);
      };
    },
    onPointer(ev: EditorPointerEvent): boolean {
      const c = ctx;
      if (!c) return false;

      if (ev.type === 'move') {
        if (drag) {
          if (ev.point) {
            // clamp so neither column goes below MIN_WIDTH
            const lo = drag.hit.leftX + MIN_WIDTH;
            const hi = drag.hit.rightX - MIN_WIDTH;
            drag.borderX = Math.max(lo, Math.min(hi, ev.point.x));
            showGuide(c, drag.hit, drag.borderX);
          }
          return true; // claim: suppress the editor's selection drag
        }
        if (ev.buttons === 0) {
          const onBorder = ev.point ? borderAt(c.layout, ev.point) : null;
          c.setCursor(onBorder ? 'col-resize' : null);
        }
        return false;
      }

      if (ev.type === 'down') {
        if (ev.buttons !== 1 || !ev.point) return false;
        const hit = borderAt(c.layout, ev.point);
        if (!hit) return false;
        drag = { hit, borderX: hit.borderX };
        c.setCursor('col-resize');
        showGuide(c, hit, hit.borderX);
        return true; // claim → the editor captures the pointer for us
      }

      if (ev.type === 'up') {
        if (!drag) return false;
        const d = drag;
        drag = null;
        c.setGuide(null);
        c.setCursor(null);
        commit(c, d);
        return true;
      }

      return false;
    },
  };
}
