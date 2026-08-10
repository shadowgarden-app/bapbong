import type {
  EditorPlugin,
  EditorPointerEvent,
  OverlayRect,
  PagePoint,
  PluginContext,
  ResolvedCell,
  ResolvedTable,
} from '@shadow-garden/bapbong-contracts';
import { commonTableAt } from '@shadow-garden/bapbong-selection';

/** The editor state type, from the plugin context (no direct PM dep). */
type State = PluginContext['state'];

/** A selected cell + its position within the block grid (0-based). */
export interface SelectedCell {
  pos: number;
  row: number;
  col: number;
}
/** The selected rectangular cell block as a grid. */
export interface CellBlock {
  cells: SelectedCell[];
  rows: number;
  cols: number;
}

/** Richer handle reached as `editor.plugin('table-selection')`. */
export interface TableSelectionPlugin extends EditorPlugin {
  /** Subscribe to the action trigger (icon tap / menu); receives the selected
   *  block (cells + grid size). Returns an unsubscribe. */
  onAction(cb: (block: CellBlock) => void): () => void;
  /** The current selected block, or null when there's no selection. */
  block(): CellBlock | null;
  /** Clear the block selection + its overlay. */
  clear(): void;
}

/** Register the handle's type so `editor.plugin('table-selection')` is typed
 *  without the core knowing this plugin exists. */
declare module '@shadow-garden/bapbong-contracts' {
  interface EditorPluginHandles {
    'table-selection': TableSelectionPlugin;
  }
}

/** Cells of `table` overlapping the bounding box of cells `a` and `b`. */
function blockCells(
  table: ResolvedTable,
  a: ResolvedCell,
  b: ResolvedCell,
): ResolvedCell[] {
  const left = Math.min(a.x, b.x);
  const top = Math.min(a.y, b.y);
  const right = Math.max(a.x + a.width, b.x + b.width);
  const bottom = Math.max(a.y + a.height, b.y + b.height);
  const eps = 0.5;
  return table.cells.filter(
    (c) =>
      c.x + c.width > left + eps &&
      c.x < right - eps &&
      c.y + c.height > top + eps &&
      c.y < bottom - eps,
  );
}

/** The `table_cell` doc position for a resolved cell (via its first line). */
function cellDocPos(state: State, cell: ResolvedCell): number | null {
  const from = cell.lines[0]?.from;
  if (from == null) return null;
  const $pos = state.doc.resolve(from);
  for (let d = $pos.depth; d > 0; d--) {
    if ($pos.node(d).type.name === 'table_cell') return $pos.before(d);
  }
  return null;
}

/**
 * Internal plugin: drag across table cells to select a rectangular block. A
 * drag within one cell stays text selection; crossing into another cell of the
 * same table claims the drag, collapses the text selection, paints the block as
 * a translucent highlight, and (on release) shows a small action button at the
 * block's top-right — a touch-friendly trigger to open cell properties. The
 * host subscribes via `onAction` and reads the selected block via `block()`.
 *
 * Pointer-event based (mouse/touch/pen). The selection is editor overlay state
 * (bapbong's table schema has no ProseMirror CellSelection).
 */
export function tableSelectionPlugin(): TableSelectionPlugin {
  let ctx: PluginContext | null = null;
  let anchor: PagePoint | null = null; // press origin (page-local)
  let lastHead: PagePoint | null = null; // latest drag point
  let selecting = false;
  let collapsed = false;
  let current: CellBlock | null = null; // the finalized block
  const listeners = new Set<(block: CellBlock) => void>();

  const reset = (c: PluginContext) => {
    anchor = null;
    lastHead = null;
    selecting = false;
    collapsed = false;
    current = null;
    c.setHighlight(null);
    c.setActionButton(null);
  };

  const finalize = (c: PluginContext): void => {
    // The block forms in the DEEPEST table containing both endpoints — a
    // drag inside a nested table selects nested cells; a drag from a nested
    // cell out into another outer cell selects in the OUTER table, the
    // nested table's whole host cell included (Word's semantics).
    // (cellA === cellB here means the drag crossed a border and came back —
    // that finalizes as a one-cell block, as it always has.)
    const common =
      anchor && lastHead && commonTableAt(c.layout, anchor, lastHead);
    if (!common) return;
    const cells = blockCells(common.table, common.cellA, common.cellB);
    // Grid position from distinct cell tops/lefts.
    const ys = [...new Set(cells.map((cell) => Math.round(cell.y)))].sort(
      (p, q) => p - q,
    );
    const xs = [...new Set(cells.map((cell) => Math.round(cell.x)))].sort(
      (p, q) => p - q,
    );
    const sel: SelectedCell[] = [];
    for (const cell of cells) {
      const pos = cellDocPos(c.state, cell);
      if (pos != null)
        sel.push({
          pos,
          row: ys.indexOf(Math.round(cell.y)),
          col: xs.indexOf(Math.round(cell.x)),
        });
    }
    if (sel.length === 0) return;
    current = { cells: sel, rows: ys.length, cols: xs.length };
    const topRight: PagePoint = {
      pageIndex: anchor!.pageIndex,
      x: Math.max(...cells.map((cell) => cell.x + cell.width)),
      y: Math.min(...cells.map((cell) => cell.y)),
    };
    c.setActionButton(
      topRight,
      () => current && listeners.forEach((cb) => cb(current!)),
    );
  };

  return {
    name: 'table-selection',
    setup(c) {
      ctx = c;
      // Without this, a cell-block highlight and its action button survived
      // into the NEXT document — blue rectangles floating over unrelated text.
      return () => {
        c.setHighlight(null);
        c.setActionButton(null);
      };
    },
    onPointer(ev: EditorPointerEvent): boolean {
      const c = ctx;
      if (!c) return false;

      if (ev.type === 'down') {
        // Ignore right/middle clicks — keep the block so the context menu (and
        // its Cell properties item) can act on it. A primary press clears it.
        if (ev.buttons !== 1) return false;
        reset(c);
        anchor = ev.point;
        return false; // let the editor start an in-cell text drag
      }

      if (ev.type === 'move') {
        if (!anchor || !(ev.buttons & 1) || !ev.point) return false;
        const common = commonTableAt(c.layout, anchor, ev.point);
        if (!common) return selecting;
        // Same cell of the deepest common table on both ends: the drag never
        // crossed a cell border at any level — a text drag, not a block.
        if (common.cellA === common.cellB && !selecting) return false;
        selecting = true;
        lastHead = ev.point;
        if (!collapsed) {
          c.setSelection(c.state.selection.from); // collapse text selection (layout unchanged)
          collapsed = true;
        }
        const rects: OverlayRect[] = blockCells(
          common.table,
          common.cellA,
          common.cellB,
        ).map((cell) => ({
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
        if (selecting) {
          finalize(c);
          return true; // keep the block + button; suppress text-selection commit
        }
        anchor = null;
        return false;
      }

      return false;
    },

    onAction(cb) {
      listeners.add(cb);
      return () => listeners.delete(cb);
    },
    block() {
      return current;
    },
    clear() {
      if (ctx) reset(ctx);
    },
  };
}
