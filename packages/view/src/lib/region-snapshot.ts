/**
 * A picture of part of the document — the passage a reader picked to talk
 * about, cut out of its page with a little margin, the picked lines tinted.
 *
 * This file is the geometry only: which page bands to cut and where the
 * tint lands inside each cut. Rendering lives in RenderCore.regionSnapshot,
 * which paints the pages off screen and follows this plan.
 */
import type { SelectionRect } from '@shadow-garden/bapbong-contracts';

export interface SnapshotRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** One cut out of one page: `crop` in page px, `marks` relative to it. */
export interface SnapshotPiece {
  pageIndex: number;
  crop: SnapshotRect;
  marks: SnapshotRect[];
}

/**
 * Cut one band per page the rects touch, in page order: the bounding box
 * of that page's rects grown by `pad`, clamped to the page. The marks are
 * the rects themselves, moved into the band's coordinates.
 *
 * `pageSize(i)` answers the page's own size (pages differ in orientation);
 * a rect on a page it cannot size is dropped.
 */
export function planSnapshot(
  rects: readonly SelectionRect[],
  pageSize: (pageIndex: number) => { width: number; height: number } | null,
  pad: number,
): SnapshotPiece[] {
  const byPage = new Map<number, SelectionRect[]>();
  for (const r of rects) {
    if (r.width <= 0 || r.height <= 0) continue;
    const list = byPage.get(r.pageIndex);
    if (list) list.push(r);
    else byPage.set(r.pageIndex, [r]);
  }
  const pieces: SnapshotPiece[] = [];
  for (const pageIndex of [...byPage.keys()].sort((a, b) => a - b)) {
    const size = pageSize(pageIndex);
    if (!size) continue;
    const list = byPage.get(pageIndex)!;
    const left = Math.max(0, Math.min(...list.map((r) => r.x)) - pad);
    const top = Math.max(0, Math.min(...list.map((r) => r.y)) - pad);
    const right = Math.min(
      size.width,
      Math.max(...list.map((r) => r.x + r.width)) + pad,
    );
    const bottom = Math.min(
      size.height,
      Math.max(...list.map((r) => r.y + r.height)) + pad,
    );
    if (right <= left || bottom <= top) continue;
    pieces.push({
      pageIndex,
      crop: { x: left, y: top, width: right - left, height: bottom - top },
      marks: list.map((r) => ({
        x: r.x - left,
        y: r.y - top,
        width: r.width,
        height: r.height,
      })),
    });
  }
  return pieces;
}

/** The scale that fits `width`×`height` inside `maxSide` on its long
 *  side — never above `natural` (no upscaling past the bitmap). */
export function fitScale(
  width: number,
  height: number,
  maxSide: number,
  natural: number,
): number {
  const long = Math.max(width, height);
  if (long <= 0) return natural;
  return Math.min(natural, maxSide / long);
}
