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
const KNOB_OFFSET = 27; // rotate knob center above the top edge (mirrors setFrame's DOM)
const KNOB_TOL = 8; // px proximity to the rotate knob
const SNAP_TOL = 3; // degrees of pull toward 0/90/180/270

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
  rotation: number; // image rotation during the drag (deltas map to local axes)
  startX: number;
  startY: number;
  rect: Rect; // current (clamped) preview rect
}

interface RotateState {
  base: Rect;
  pageIndex: number;
  /** Image rotation minus the pointer's start angle: the grip stays under
   *  the pointer as it orbits the center. */
  grip: number;
  rotation: number; // current (snapped) preview rotation
}

/** `point` mapped into the frame's unrotated local space (inverse-rotate
 *  around the rect center). Handle/knob hit-tests run in this space. */
export function toLocal(
  x: number,
  y: number,
  rect: Rect,
  rotation: number,
): { x: number; y: number } {
  if (!rotation) return { x, y };
  const rad = (-rotation * Math.PI) / 180;
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const dx = x - cx;
  const dy = y - cy;
  return {
    x: cx + dx * Math.cos(rad) - dy * Math.sin(rad),
    y: cy + dx * Math.sin(rad) + dy * Math.cos(rad),
  };
}

/** Resize cursor for a handle, accounting for the frame's rotation: each
 *  handle points at a compass angle (n = 0°, e = 90°, …); adding the rotation
 *  and quantizing to 45° picks among the four bidirectional resize cursors. */
export function cursorFor(handle: Handle, rotation: number): string {
  const DIR: Record<Handle, number> = {
    n: 0,
    ne: 45,
    e: 90,
    se: 135,
    s: 180,
    sw: 225,
    w: 270,
    nw: 315,
  };
  const idx =
    Math.round(((((DIR[handle] + rotation) % 360) + 360) % 360) / 45) % 4;
  return ['ns-resize', 'nesw-resize', 'ew-resize', 'nwse-resize'][idx];
}

/** Snap: within SNAP_TOL of a cardinal angle → that angle; with Shift → 15°
 *  steps. Result normalized to [0, 360). */
export function snapAngle(deg: number, shift: boolean): number {
  let d = ((deg % 360) + 360) % 360;
  if (shift) return (Math.round(d / 15) * 15) % 360;
  for (const c of [0, 90, 180, 270, 360]) {
    if (Math.abs(d - c) <= SNAP_TOL) {
      d = c % 360;
      break;
    }
  }
  return d;
}

/** The handle under a page point, testing each handle's center on `rect`. */
function handleAt(rect: Rect, x: number, y: number): Handle | null {
  const xs: Record<'w' | 'c' | 'e', number> = {
    w: rect.x,
    c: rect.x + rect.width / 2,
    e: rect.x + rect.width,
  };
  const ys: Record<'n' | 'm' | 's', number> = {
    n: rect.y,
    m: rect.y + rect.height / 2,
    s: rect.y + rect.height,
  };
  const H: [Handle, number, number][] = [
    ['nw', xs.w, ys.n],
    ['n', xs.c, ys.n],
    ['ne', xs.e, ys.n],
    ['w', xs.w, ys.m],
    ['e', xs.e, ys.m],
    ['sw', xs.w, ys.s],
    ['s', xs.c, ys.s],
    ['se', xs.e, ys.s],
  ];
  for (const [h, hx, hy] of H) {
    if (Math.abs(x - hx) <= HANDLE_TOL && Math.abs(y - hy) <= HANDLE_TOL)
      return h;
  }
  return null;
}

/** New rect for dragging `handle` by (dx, dy) from `base`: the opposite
 *  edge/corner stays anchored; corners keep the aspect ratio (edges are free)
 *  unless Shift flips the mode; both dimensions clamp to MIN_SIZE. */
export function resizeRect(
  base: Rect,
  handle: Handle,
  dx: number,
  dy: number,
  shift: boolean,
): Rect {
  const west = handle.includes('w');
  const north = handle.includes('n');
  const horiz = handle !== 'n' && handle !== 's';
  const vert = handle !== 'e' && handle !== 'w';
  let w = horiz
    ? Math.max(MIN_SIZE, base.width + (west ? -dx : dx))
    : base.width;
  let h = vert
    ? Math.max(MIN_SIZE, base.height + (north ? -dy : dy))
    : base.height;
  const corner = horiz && vert;
  if (corner !== shift && base.width > 0 && base.height > 0) {
    // Keep aspect: follow whichever axis moved more (relative to the base).
    const sw = w / base.width;
    const sh = h / base.height;
    const s = Math.max(
      Math.abs(sw - 1) >= Math.abs(sh - 1) ? sw : sh,
      MIN_SIZE / base.width,
      MIN_SIZE / base.height,
    );
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
function frameForPos(
  layout: ResolvedLayout | null,
  pos: number,
): OverlayFrame | null {
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
        return {
          pageIndex: page.index,
          x: f.x,
          y: f.y,
          width: f.width,
          height: f.height,
          ...(f.rotation ? { rotation: f.rotation } : {}),
        };
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
          ...(img.rotation ? { rotation: img.rotation } : {}),
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
  let rot: RotateState | null = null;
  let hoverCursor = false; // we set the canvas cursor (so we may clear it)

  const setCursor = (c: PluginContext, cursor: string | null): void => {
    if (cursor === null && !hoverCursor) return; // don't clobber other plugins
    c.setCursor(cursor);
    hoverCursor = cursor !== null;
  };

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
    if (w === Math.round(d.base.width) && h === Math.round(d.base.height))
      return;
    c.dispatch(
      c.state.tr
        .setNodeAttribute(sel.pos, 'width', w)
        .setNodeAttribute(sel.pos, 'height', h),
    );
  };

  /** Commit the rotate gesture's final angle — likewise one transaction. */
  const commitRotation = (c: PluginContext, r: RotateState): void => {
    if (!sel || !imageNodeAt(c.state, sel.pos)) return;
    const deg = Math.round(r.rotation * 10) / 10;
    if (deg === rotationAt(c.state, sel.pos)) return;
    c.dispatch(c.state.tr.setNodeAttribute(sel.pos, 'rotation', deg));
  };

  const rotationAt = (state: State, pos: number): number =>
    Number(state.doc.nodeAt(pos)?.attrs['rotation']) || 0;

  /** Pointer angle around the rect center, in clockwise degrees where 0 points
   *  up (the knob's rest direction). */
  const angleAround = (rect: Rect, x: number, y: number): number => {
    const cx = rect.x + rect.width / 2;
    const cy = rect.y + rect.height / 2;
    return (Math.atan2(x - cx, -(y - cy)) * 180) / Math.PI;
  };

  return {
    name: 'image-resize',
    setup(c) {
      ctx = c;
      return () => c.setFrame(null);
    },
    onChange() {
      // Skip while dragging/rotating: the preview frame must keep following
      // the pointer, not snap back to the (unchanged) layout.
      if (ctx && sel && !drag && !rot) refresh(ctx);
    },
    onKey(ev) {
      const c = ctx;
      if (!c) return false;
      // Backspace/Delete on a selected image removes the node. The selection
      // lives in THIS plugin (not a PM NodeSelection), so without claiming
      // the key here the hidden editor would backspace at its own caret and
      // leave the image standing.
      if (
        (ev.key === 'Backspace' || ev.key === 'Delete') &&
        sel &&
        !drag &&
        !rot
      ) {
        if (!imageNodeAt(c.state, sel.pos)) return false;
        const node = c.state.doc.nodeAt(sel.pos);
        if (!node) return false;
        c.dispatch(c.state.tr.delete(sel.pos, sel.pos + node.nodeSize));
        sel = null;
        c.setFrame(null);
        setCursor(c, null);
        return true;
      }
      if (ev.key !== 'Escape') return false;
      // Mid-gesture: abandon the preview, snap the frame back to the layout
      // (nothing was dispatched, so there is nothing to undo).
      if (drag || rot) {
        drag = null;
        rot = null;
        setCursor(c, null);
        refresh(c);
        return true;
      }
      if (sel) {
        sel = null;
        c.setFrame(null);
        setCursor(c, null);
        return true;
      }
      return false;
    },
    onPointer(ev: EditorPointerEvent): boolean {
      const c = ctx;
      if (!c) return false;

      if (ev.type === 'move') {
        if (rot) {
          if (ev.point) {
            rot.rotation = snapAngle(
              angleAround(rot.base, ev.point.x, ev.point.y) + rot.grip,
              ev.shiftKey,
            );
            // DOM-only preview: the frame's CSS transform follows the pointer.
            c.setFrame({
              pageIndex: rot.pageIndex,
              ...rot.base,
              rotation: rot.rotation,
              label: `${Math.round(rot.rotation)}°`,
            });
          }
          return true;
        }
        if (drag) {
          if (ev.point) {
            // Deltas map into the image's local axes, so dragging a rotated
            // image's corner still grows it along its own edges.
            const d = toLocal(ev.point.x, ev.point.y, drag.base, drag.rotation);
            const s = toLocal(
              drag.startX,
              drag.startY,
              drag.base,
              drag.rotation,
            );
            drag.rect = resizeRect(
              drag.base,
              drag.handle,
              d.x - s.x,
              d.y - s.y,
              ev.shiftKey,
            );
            // DOM-only preview: frame + size readout follow the pointer; the
            // document, layout, and content canvas stay untouched.
            c.setFrame({
              pageIndex: drag.pageIndex,
              ...drag.rect,
              ...(drag.rotation ? { rotation: drag.rotation } : {}),
              label: `${Math.round(drag.rect.width)} × ${Math.round(drag.rect.height)}`,
            });
          }
          return true;
        }
        // Hover feedback over the selection's handles / rotate knob.
        if (
          ev.buttons === 0 &&
          ev.point &&
          sel &&
          imageNodeAt(c.state, sel.pos)
        ) {
          const frame = frameForPos(c.layout, sel.pos);
          if (frame && frame.pageIndex === ev.point.pageIndex) {
            const rotation = rotationAt(c.state, sel.pos);
            const base = {
              x: frame.x,
              y: frame.y,
              width: frame.width,
              height: frame.height,
            };
            const local = toLocal(ev.point.x, ev.point.y, base, rotation);
            if (
              Math.abs(local.x - (base.x + base.width / 2)) <= KNOB_TOL &&
              Math.abs(local.y - (base.y - KNOB_OFFSET)) <= KNOB_TOL
            ) {
              setCursor(c, 'grab');
              return false;
            }
            const handle = handleAt(base, local.x, local.y);
            setCursor(c, handle ? cursorFor(handle, rotation) : null);
            return false;
          }
          setCursor(c, null);
        }
        return false;
      }

      if (ev.type === 'down' && ev.buttons === 1 && ev.point) {
        // A drag starts on a handle (or the rotate knob) of the selection…
        if (sel && imageNodeAt(c.state, sel.pos)) {
          const frame = frameForPos(c.layout, sel.pos);
          if (frame && frame.pageIndex === ev.point.pageIndex) {
            const rotation = rotationAt(c.state, sel.pos);
            const base = {
              x: frame.x,
              y: frame.y,
              width: frame.width,
              height: frame.height,
            };
            const local = toLocal(ev.point.x, ev.point.y, base, rotation);
            const knobX = base.x + base.width / 2;
            const knobY = base.y - KNOB_OFFSET;
            if (
              Math.abs(local.x - knobX) <= KNOB_TOL &&
              Math.abs(local.y - knobY) <= KNOB_TOL
            ) {
              rot = {
                base,
                pageIndex: frame.pageIndex,
                grip: rotation - angleAround(base, ev.point.x, ev.point.y),
                rotation,
              };
              setCursor(c, 'grabbing');
              return true;
            }
            const handle = handleAt(base, local.x, local.y);
            if (handle) {
              drag = {
                handle,
                base,
                pageIndex: frame.pageIndex,
                rotation,
                startX: ev.point.x,
                startY: ev.point.y,
                rect: { ...base },
              };
              return true; // claim → the editor captures the pointer for us
            }
          }
        }
        // …otherwise a click selects (or deselects) an image.
        const hit = imageAtPoint(c.layout as ResolvedLayout, ev.point);
        if (hit) {
          sel = { pos: hit.pos };
          c.setFrame({
            pageIndex: hit.pageIndex,
            ...hit.rect,
            ...(rotationAt(c.state, hit.pos)
              ? { rotation: rotationAt(c.state, hit.pos) }
              : {}),
          });
          return true; // claim: keep the caret where it is
        }
        if (sel) {
          sel = null;
          c.setFrame(null);
          setCursor(c, null);
        }
        return false; // let the editor place the caret
      }

      if (ev.type === 'up') {
        if (rot) {
          const r = rot;
          rot = null;
          setCursor(c, null);
          commitRotation(c, r);
          refresh(c);
          return true;
        }
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
