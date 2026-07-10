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

/** Plugin-local selection: the image node's PM position. Geometry is always
 *  re-derived from the current layout (it moves on every reflow). */
interface Selected {
  pos: number;
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

  const refresh = (c: PluginContext): void => {
    if (!sel || !imageNodeAt(c.state, sel.pos)) {
      sel = null;
      c.setFrame(null);
      return;
    }
    c.setFrame(frameForPos(c.layout, sel.pos));
  };

  return {
    name: 'image-resize',
    setup(c) {
      ctx = c;
      return () => c.setFrame(null);
    },
    onChange() {
      if (ctx && sel) refresh(ctx);
    },
    onPointer(ev: EditorPointerEvent): boolean {
      const c = ctx;
      if (!c) return false;

      if (ev.type === 'down' && ev.buttons === 1 && ev.point) {
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

      return false;
    },
  };
}
