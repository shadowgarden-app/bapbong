/**
 * A wpg group of shapes → a vector display list, in px of the group's box.
 *
 * Word's converter turns a legacy drawing canvas into exactly this: a group
 * (possibly holding groups) of freeform shapes — a:custGeom paths of moveTo
 * / lnTo / close — with a solid fill or none and a solid outline. The whole
 * group paints as one image box (see parseCarriedGroup), so every point is
 * carried from its path space through each level's coordinate system down
 * to px of that box:
 *
 *   path (w × h)  →  shape box (a:off / a:ext in the parent's child space)
 *   child space (a:chOff / a:chExt)  →  group box (a:off / a:ext in ITS
 *   parent's child space)  →  …  →  the top group's box, EMU  →  px
 *
 * Outline widths are NOT carried through the scales: DrawingML states them
 * in absolute EMU, and a scaled group keeps its lines' weight. They leave
 * here as px of the box, the unit the painter scales by zoom.
 *
 * Pure: nodes in, ops out; colours through the caller's resolver.
 */
import type {
  VectorOp,
  VectorPolygonOp,
  VectorPolylineOp,
} from '@shadow-garden/bapbong-contracts';
import { attrOf, child, children, type OoxmlNode } from './ooxml.js';

const EMU_PER_PX = 9525;

type Pt = { x: number; y: number };
type Map2 = (p: Pt) => Pt;

/** a:xfrm's box: offset, extent, and the flips/rotation that act on it. */
interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number; // degrees clockwise
  flipH: boolean;
  flipV: boolean;
}

function boxOf(xfrm: OoxmlNode | undefined): Box {
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  const n = (el: OoxmlNode | undefined, a: string) =>
    Number(attrOf(el, a) ?? '0') || 0;
  return {
    x: n(off, 'x'),
    y: n(off, 'y'),
    w: n(ext, 'cx'),
    h: n(ext, 'cy'),
    rot: (Number(attrOf(xfrm, 'rot') ?? '0') || 0) / 60000,
    flipH: attrOf(xfrm, 'flipH') === '1',
    flipV: attrOf(xfrm, 'flipV') === '1',
  };
}

/** Box-local point (0..w, 0..h) → the parent's space: flip inside the box,
 *  rotate about its centre, then place it at its offset. */
function placeInBox(b: Box): Map2 {
  const cx = b.w / 2;
  const cy = b.h / 2;
  const rad = (b.rot * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  return ({ x, y }) => {
    let lx = b.flipH ? b.w - x : x;
    let ly = b.flipV ? b.h - y : y;
    if (b.rot) {
      const dx = lx - cx;
      const dy = ly - cy;
      lx = cx + dx * cos - dy * sin;
      ly = cy + dx * sin + dy * cos;
    }
    return { x: b.x + lx, y: b.y + ly };
  };
}

/** A group's child space → the group's box-local space. */
function childToBox(xfrm: OoxmlNode | undefined, b: Box): Map2 {
  const chOff = child(xfrm, 'a:chOff');
  const chExt = child(xfrm, 'a:chExt');
  const ox = Number(attrOf(chOff, 'x') ?? '0') || 0;
  const oy = Number(attrOf(chOff, 'y') ?? '0') || 0;
  const cw = Number(attrOf(chExt, 'cx') ?? '0') || b.w || 1;
  const ch = Number(attrOf(chExt, 'cy') ?? '0') || b.h || 1;
  const sx = b.w / cw;
  const sy = b.h / ch;
  return ({ x, y }) => ({ x: (x - ox) * sx, y: (y - oy) * sy });
}

export interface FreeformGroupResult {
  ops: VectorOp[];
  /** Members drawn / members that could not be (unsupported geometry). */
  drawn: number;
  skipped: number;
}

export interface FreeformGroupOptions {
  /** CSS colour of an element's a:solidFill child (or undefined). */
  solidFill(parent: OoxmlNode | undefined): string | undefined;
  /** Called for every node read as understood — the audit's hook. */
  consumed?(node: OoxmlNode | undefined): void;
}

/** Points of a cubic Bézier, flattened: enough segments that a curve the
 *  size of a page stays smooth, few enough for a small arrow. */
const BEZIER_STEPS = 12;

/**
 * The display list of `wgp` (a wpg:wgp). Coordinates in px of the group's
 * box — the top a:ext — so the caller's VectorImageSpec is that box's px.
 */
export function freeformGroupOps(
  wgp: OoxmlNode,
  opts: FreeformGroupOptions,
): FreeformGroupResult {
  const ops: VectorOp[] = [];
  let drawn = 0;
  let skipped = 0;
  const topXfrm = child(child(wgp, 'wpg:grpSpPr'), 'a:xfrm');
  opts.consumed?.(topXfrm);
  const topBox = boxOf(topXfrm);
  // The top group's own offset is where the anchor puts it — not ours.
  const toBox = childToBox(topXfrm, { ...topBox, x: 0, y: 0 });
  const toPx: Map2 = (p) => {
    const q = toBox(p);
    return { x: q.x / EMU_PER_PX, y: q.y / EMU_PER_PX };
  };

  const walk = (grp: OoxmlNode, toTop: Map2): void => {
    for (const m of grp.children) {
      if (m.name === 'wpg:grpSp') {
        const xfrm = child(child(m, 'wpg:grpSpPr'), 'a:xfrm');
        opts.consumed?.(xfrm);
        const b = boxOf(xfrm);
        const inner = childToBox(xfrm, b);
        const place = placeInBox(b);
        walk(m, (p) => toTop(place(inner(p))));
      } else if (m.name === 'wps:wsp') {
        if (shapeOps(m, toTop)) drawn++;
        else skipped++;
      }
    }
  };

  const shapeOps = (wsp: OoxmlNode, toTop: Map2): boolean => {
    const spPr = child(wsp, 'wps:spPr');
    const xfrm = child(spPr, 'a:xfrm');
    const geom = child(spPr, 'a:custGeom');
    if (!geom) return false;
    const b = boxOf(xfrm);
    const place = placeInBox(b);
    const fill = child(spPr, 'a:noFill') ? undefined : opts.solidFill(spPr);
    const ln = child(spPr, 'a:ln');
    const stroke =
      !ln || child(ln, 'a:noFill') ? undefined : opts.solidFill(ln);
    const strokeWidth = Number(attrOf(ln, 'w') ?? '0') / EMU_PER_PX;
    const join = child(ln, 'a:round') ? ('round' as const) : undefined;
    const capAttr = attrOf(ln, 'cap');
    const cap =
      capAttr === 'rnd'
        ? ('round' as const)
        : capAttr === 'sq'
          ? ('square' as const)
          : undefined;
    const shapeOut: VectorOp[] = [];
    for (const path of children(child(geom, 'a:pathLst'), 'a:path')) {
      // A path states its own coordinate space; an absent (or zero) extent
      // means the shape's — its points are then in the shape's units.
      const pw = Number(attrOf(path, 'w') ?? '0') || b.w || 1;
      const ph = Number(attrOf(path, 'h') ?? '0') || b.h || 1;
      const toShape: Map2 = ({ x, y }) => ({
        x: (x * b.w) / pw,
        y: (y * b.h) / ph,
      });
      const map: Map2 = (p) => toPx(toTop(place(toShape(p))));
      const pathFill = attrOf(path, 'fill') === 'none' ? undefined : fill;
      const pathStroke = attrOf(path, 'stroke') === '0' ? undefined : stroke;
      const pathOut = pathOps(path, map, pathFill, pathStroke, {
        width: strokeWidth,
        join,
        cap,
      });
      // One command we cannot draw and the whole shape stays out — a
      // half-drawn shape reads as a different shape.
      if (pathOut === null) return false;
      shapeOut.push(...pathOut);
    }
    if (shapeOut.length === 0) return false;
    ops.push(...shapeOut);
    opts.consumed?.(xfrm);
    opts.consumed?.(geom);
    opts.consumed?.(ln);
    return true;
  };

  /** One a:path → its subpaths as ops. Null when it holds a command we
   *  cannot draw (arcTo): the shape is skipped rather than drawn wrong. */
  const pathOps = (
    path: OoxmlNode,
    map: Map2,
    fill: string | undefined,
    stroke: string | undefined,
    line: { width: number; join?: 'round'; cap?: 'round' | 'square' },
  ): VectorOp[] | null => {
    const { width: strokeWidth, join, cap } = line;
    const out: VectorOp[] = [];
    let pts: Pt[] = [];
    let last: Pt = { x: 0, y: 0 };
    const pt = (el: OoxmlNode | undefined): Pt => ({
      x: Number(attrOf(el, 'x') ?? '0') || 0,
      y: Number(attrOf(el, 'y') ?? '0') || 0,
    });
    const flush = (closed: boolean) => {
      if (pts.length >= 2) {
        const mapped = pts.map(map);
        if (closed) {
          out.push({
            kind: 'polygon',
            points: mapped,
            ...(fill ? { fill } : {}),
            ...(stroke && strokeWidth > 0 ? { stroke, strokeWidth } : {}),
            ...(join ? { join } : {}),
          } satisfies VectorPolygonOp);
        } else {
          // An open subpath is filled as if closed (the spec's rule), and
          // stroked open.
          if (fill)
            out.push({
              kind: 'polygon',
              points: mapped,
              fill,
            } satisfies VectorPolygonOp);
          if (stroke && strokeWidth > 0)
            out.push({
              kind: 'polyline',
              points: mapped,
              stroke,
              strokeWidth,
              ...(join ? { join } : {}),
              ...(cap ? { cap } : {}),
            } satisfies VectorPolylineOp);
        }
      }
      pts = [];
    };
    for (const c of path.children) {
      const ps = children(c, 'a:pt').map(pt);
      switch (c.name) {
        case 'a:moveTo':
          flush(false);
          last = ps[0] ?? last;
          pts = [last];
          break;
        case 'a:lnTo':
          if (ps[0]) {
            if (pts.length === 0) pts = [last];
            last = ps[0];
            pts.push(last);
          }
          break;
        case 'a:cubicBezTo':
        case 'a:quadBezTo': {
          const p0 = last;
          let cps: (Pt | undefined)[] = ps;
          if (c.name === 'a:quadBezTo') {
            // The same curve as a cubic: controls 2/3 of the way from each
            // end towards the quadratic's one control point.
            const [q, e] = ps;
            if (!q || !e) return null;
            cps = [
              {
                x: p0.x + ((q.x - p0.x) * 2) / 3,
                y: p0.y + ((q.y - p0.y) * 2) / 3,
              },
              {
                x: e.x + ((q.x - e.x) * 2) / 3,
                y: e.y + ((q.y - e.y) * 2) / 3,
              },
              e,
            ];
          }
          if (cps.length < 3 || cps.some((p) => !p)) return null;
          if (pts.length === 0) pts = [last];
          const [p1, p2, p3] = cps as [Pt, Pt, Pt];
          for (let i = 1; i <= BEZIER_STEPS; i++) {
            const t = i / BEZIER_STEPS;
            const u = 1 - t;
            pts.push({
              x:
                u * u * u * p0.x +
                3 * u * u * t * p1.x +
                3 * u * t * t * p2.x +
                t * t * t * p3.x,
              y:
                u * u * u * p0.y +
                3 * u * u * t * p1.y +
                3 * u * t * t * p2.y +
                t * t * t * p3.y,
            });
          }
          last = p3;
          break;
        }
        case 'a:close':
          flush(true);
          break;
        default:
          return null; // a:arcTo — not drawn
      }
    }
    flush(false);
    return out;
  };

  walk(wgp, (p) => p);
  return { ops, drawn, skipped };
}
