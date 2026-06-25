import type {
  EditorPlugin,
  EditorPointerEvent,
  OverlayRect,
  PagePoint,
  PluginContext,
  ResolvedCell,
  ResolvedLayout,
  ResolvedTable,
} from '@shadow-garden/bapbong-contracts';

interface CellHit {
  table: ResolvedTable;
  cell: ResolvedCell;
}

/** The table cell at a page-local point (top-level tables only), or null. */
function cellAtPoint(layout: ResolvedLayout | null, point: PagePoint): CellHit | null {
  const page = layout?.pages[point.pageIndex];
  if (!page?.tables) return null;
  for (const table of page.tables) {
    if (point.x < table.x || point.x > table.x + table.width) continue;
    if (point.y < table.y || point.y > table.y + table.height) continue;
    for (const cell of table.cells) {
      if (point.x >= cell.x && point.x <= cell.x + cell.width && point.y >= cell.y && point.y <= cell.y + cell.height) {
        return { table, cell };
      }
    }
  }
  return null;
}

/** Cells of `table` overlapping the bounding box of cells `a` and `b` — the
 *  rectangular block spanned by the drag. */
function blockCells(table: ResolvedTable, a: ResolvedCell, b: ResolvedCell): ResolvedCell[] {
  const left = Math.min(a.x, b.x);
  const top = Math.min(a.y, b.y);
  const right = Math.max(a.x + a.width, b.x + b.width);
  const bottom = Math.max(a.y + a.height, b.y + b.height);
  const eps = 0.5;
  return table.cells.filter(
    (c) => c.x + c.width > left + eps && c.x < right - eps && c.y + c.height > top + eps && c.y < bottom - eps,
  );
}

/**
 * Internal plugin: drag across table cells to select a rectangular block. A
 * drag *within* one cell stays text selection; once it crosses into another
 * cell of the same table the plugin claims the drag, collapses the text
 * selection, and paints the block as a translucent highlight. Pointer-event
 * based, so mouse / touch / pen all work.
 *
 * The selection is editor overlay state (bapbong's table schema has no
 * ProseMirror CellSelection); it drives the cell highlight and (next) the
 * floating action icon + cell-properties dialog.
 */
export function tableSelectionPlugin(): EditorPlugin {
  let ctx: PluginContext | null = null;
  let anchor: PagePoint | null = null; // where the press began (page-local)
  let selecting = false; // crossed into another cell → block selection active
  let collapsed = false; // collapsed the text selection once

  const reset = (c: PluginContext) => {
    anchor = null;
    selecting = false;
    collapsed = false;
    c.setHighlight(null);
  };

  return {
    name: 'table-selection',
    setup(c) {
      ctx = c;
    },
    onPointer(ev: EditorPointerEvent): boolean {
      const c = ctx;
      if (!c) return false;

      if (ev.type === 'down') {
        reset(c); // a fresh press clears any prior block selection
        anchor = ev.buttons === 1 ? ev.point : null;
        return false; // let the editor start an in-cell text drag
      }

      if (ev.type === 'move') {
        if (!anchor || !(ev.buttons & 1) || !ev.point) return false;
        const a = cellAtPoint(c.layout, anchor);
        const h = cellAtPoint(c.layout, ev.point);
        if (!a || !h || a.table !== h.table) return selecting; // outside/other table
        if (a.cell === h.cell && !selecting) return false; // same cell → text drag
        selecting = true;
        if (!collapsed) {
          // Collapse the text selection so its overlay doesn't compete. A
          // selection-only change keeps the layout geometry valid.
          c.setSelection(c.state.selection.from);
          collapsed = true;
        }
        const rects: OverlayRect[] = blockCells(a.table, a.cell, h.cell).map((cell) => ({
          pageIndex: ev.point!.pageIndex,
          x: cell.x,
          y: cell.y,
          width: cell.width,
          height: cell.height,
        }));
        c.setHighlight(rects);
        return true; // claim → suppress the editor's text drag
      }

      if (ev.type === 'up') {
        if (selecting) return true; // keep the block highlight; suppress text-selection commit
        anchor = null;
        return false;
      }

      return false;
    },
  };
}
