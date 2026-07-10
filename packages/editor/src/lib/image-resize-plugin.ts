import type {
  EditorPlugin,
  EditorPointerEvent,
  OverlayFrame,
  PluginContext,
  ResolvedLayout,
  ResolvedTable,
} from '@shadow-garden/bapbong-contracts';
import { imageAtPoint } from '@shadow-garden/bapbong-selection';

/** The editor state type, taken from the plugin context (no direct PM dep). */
type State = PluginContext['state'];

const HANDLE_TOL = 6; // px proximity to a handle center to start a resize
const MIN_SIZE = 16; // px minimum image dimension

/** Plugin-local selection: the image node's PM position. Geometry is always
 *  re-derived from the current layout (it moves on every reflow). */
interface Selected {
  pos: number;
}

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface DragState {
  handle: Handle;
  base: Rect; // frame rect when the drag started (page-local)
  pageIndex: number;
  startX: number;
  startY: number;
  rect: Rect; // current (clamped) preview rect
}

/** The handle under a page point, testing each handle's center on `rect`. */
function handleAt(rect: Rect, x: number, y: number): Handle | null {
  const xs: Record<'w' | 'c' | 'e', number> = { w: rect.x, c: rect.x + rect.width / 2, e: rect.x + rect.width };
  const ys: Record<'n' | 'm' | 's', number> = { n: rect.y, m: rect.y + rect.height / 2, s: rect.y + rect.height };
  const H: [Handle, number, number][] = [
    ['nw', xs.w, ys.n], ['n', xs.c, ys.n], ['ne', xs.e, ys.n],
    ['w', xs.w, ys.m], ['e', xs.e, ys.m],
    ['sw', xs.w, ys.s], ['s', xs.c, ys.s], ['se', xs.e, ys.s],
  ];
  for (const [h, hx, hy] of H) {
    if (Math.abs(x - hx) <= HANDLE_TOL && Math.abs(y - hy) <= HANDLE_TOL) return h;
  }
  return null;
}

/** New rect for dragging `handle` by (dx, dy) from `base`: the opposite
 *  edge/corner stays anchored; corners keep the aspect ratio (edges are free)
 *  unless Shift flips the mode; both dimensions clamp to MIN_SIZE. */
export function resizeRect(base: Rect, handle: Handle, dx: number, dy: number, shift: boolean): Rect {
  const west = handle.includes('w');
  const north = handle.includes('n');
  const horiz = handle !== 'n' && handle !== 's';
  const vert = handle !== 'e' && handle !== 'w';
  let w = horiz ? Math.max(MIN_SIZE, base.width + (west ? -dx : dx)) : base.width;
  let h = vert ? Math.max(MIN_SIZE, base.height + (north ? -dy : dy)) : base.height;
  const corner = horiz && vert;
  if (corner !== shift && base.width > 0 && base.height > 0) {
    // Keep aspect: follow whichever axis moved more (relative to the base).
    const sw = w / base.width;
    const sh = h / base.height;
    const s = Math.max(Math.abs(sw - 1) >= Math.abs(sh - 1) ? sw : sh, MIN_SIZE / base.width, MIN_SIZE / base.height);
    w = base.width * s;
    h = base.height * s;
  }
  return {
    x: west ? base.x + base.width - w : base.x,
    y: north ? base.y + base.height - h : base.y,
    width: w,
    height: h,
  };
}

/** The page-local frame for the image node at `pos`, from the current layout —
 *  or null when the node isn't laid out (page not built, image removed). */
function frameForPos(layout: ResolvedLayout | null, pos: number): OverlayFrame | null {
  if (!layout) return null;
  for (const page of layout.pages) {
    const floats = [...(page.floats ?? [])];
    const visit = (t: ResolvedTable) => {
      for (const cell of t.cells) {
        floats.push(...(cell.floats ?? []));
        cell.tables?.forEach(visit);
      }
    };
    page.tables?.forEach(visit);
    for (const f of floats) {
      if (f.pos === pos) {
        return { pageIndex: page.index, x: f.x, y: f.y, width: f.width, height: f.height };
      }
    }
    const lines = [...page.lines];
    const visitLines = (t: ResolvedTable) => {
      for (const cell of t.cells) {
        lines.push(...cell.lines);
        cell.tables?.forEach(visitLines);
      }
    };
    page.tables?.forEach(visitLines);
    for (const line of lines) {
      for (const img of line.images ?? []) {
        if (img.pos !== pos) continue;
        return {
          pageIndex: page.index,
          x: img.x,
          y: line.y + line.baseline - img.height,
          width: img.width,
          height: img.height,
        };
      }
    }
  }
  return null;
}

/** Whether the doc still has an image node at `pos`. */
function imageNodeAt(state: State, pos: number): boolean {
  return state.doc.nodeAt(pos)?.type.name === 'image';
}

/**
 * Internal plugin: click an image to select it — a frame with 8 resize handles
 * and a rotate knob appears (DOM overlay; the content canvas is untouched).
 * Clicking elsewhere deselects. After a doc change the selection survives only
 * if the same position still holds an image (a resize commit does; edits that
 * shift positions drop it), and the frame re-anchors to the fresh layout.
 *
 * Resize / rotate drags land in later steps (M16 R3/R5); this step is the
 * selection lifecycle the drags build on.
 */
export function imageResizePlugin(): EditorPlugin {
  let ctx: PluginContext | null = null;
  let sel: Selected | null = null;
  let drag: DragState | null = null;

  const refresh = (c: PluginContext): void => {
    if (!sel || !imageNodeAt(c.state, sel.pos)) {
      sel = null;
      c.setFrame(null);
      return;
    }
    c.setFrame(frameForPos(c.layout, sel.pos));
  };

  /** Commit the drag's final size — ONE transaction, the gesture's only touch
   *  on the document (and a single undo step). */
  const commit = (c: PluginContext, d: DragState): void => {
    if (!sel || !imageNodeAt(c.state, sel.pos)) return;
    const w = Math.round(d.rect.width);
    const h = Math.round(d.rect.height);
    if (w === Math.round(d.base.width) && h === Math.round(d.base.height)) return;
    c.dispatch(c.state.tr.setNodeAttribute(sel.pos, 'width', w).setNodeAttribute(sel.pos, 'height', h));
  };

  return {
    name: 'image-resize',
    setup(c) {
      ctx = c;
      return () => c.setFrame(null);
    },
    onChange() {
      // Skip while dragging: the preview frame must keep following the
      // pointer, not snap back to the (unchanged) layout.
      if (ctx && sel && !drag) refresh(ctx);
    },
    onPointer(ev: EditorPointerEvent): boolean {
      const c = ctx;
      if (!c) return false;

      if (ev.type === 'move') {
        if (drag) {
          if (ev.point) {
            drag.rect = resizeRect(
              drag.base,
              drag.handle,
              ev.point.x - drag.startX,
              ev.point.y - drag.startY,
              ev.shiftKey,
            );
            // DOM-only preview: frame + size readout follow the pointer; the
            // document, layout, and content canvas stay untouched.
            c.setFrame({
              pageIndex: drag.pageIndex,
              ...drag.rect,
              label: `${Math.round(drag.rect.width)} × ${Math.round(drag.rect.height)}`,
            });
          }
          return true;
        }
        return false;
      }

      if (ev.type === 'down' && ev.buttons === 1 && ev.point) {
        // A drag starts on a handle of the current selection…
        if (sel && imageNodeAt(c.state, sel.pos)) {
          const frame = frameForPos(c.layout, sel.pos);
          if (frame && frame.pageIndex === ev.point.pageIndex) {
            const handle = handleAt(frame, ev.point.x, ev.point.y);
            if (handle) {
              drag = {
                handle,
                base: { x: frame.x, y: frame.y, width: frame.width, height: frame.height },
                pageIndex: frame.pageIndex,
                startX: ev.point.x,
                startY: ev.point.y,
                rect: { x: frame.x, y: frame.y, width: frame.width, height: frame.height },
              };
              return true; // claim → the editor captures the pointer for us
            }
          }
        }
        // …otherwise a click selects (or deselects) an image.
        const hit = imageAtPoint(c.layout as ResolvedLayout, ev.point);
        if (hit) {
          sel = { pos: hit.pos };
          c.setFrame({ pageIndex: hit.pageIndex, ...hit.rect });
          return true; // claim: keep the caret where it is
        }
        if (sel) {
          sel = null;
          c.setFrame(null);
        }
        return false; // let the editor place the caret
      }

      if (ev.type === 'up') {
        if (!drag) return false;
        const d = drag;
        drag = null;
        commit(c, d);
        // The commit's relayout refresh re-anchors the frame; when nothing
        // changed (click on a handle without moving) re-anchor here.
        refresh(c);
        return true;
      }

      return false;
    },
  };
}
