/**
 * A VML group (`v:group`) → one drawing: its geometry as a vector display
 * list over the group's box, and its text boxes as frames placed in it.
 *
 * Legacy Word documents draw diagrams, framed captions and whole bar charts
 * this way (D-2609-DHQ8: 18 groups — shadowed caption boxes, a marketing-mix
 * wheel, charts of rectangles and text boxes). Members are positioned in the
 * group's own coordinate space (coordorigin / coordsize), groups nest, and a
 * member's `style` lengths are in its parent's units, not points.
 *
 * Geometry: v:rect, v:roundrect, v:oval, v:line, and v:shape — a text frame
 * (shapetype 202) or a freeform `path`. Paths that compute their points
 * (`@n` formula references) are skipped: their shape cannot be known without
 * evaluating the shapetype's formulas.
 *
 * Pure: nodes in, display list out; text boxes are read through the
 * caller's `textbox` reader (it needs the document context).
 */
import type {
  VectorImageSpec,
  VectorOp,
} from '@shadow-garden/bapbong-contracts';
import { attrOf, child, type OoxmlNode } from './ooxml.js';

/** A text box placed in the group's box, px from its top-left. */
export interface GroupFrame {
  x: number;
  y: number;
  width: number;
  height: number;
  textbox: {
    blocks: unknown[];
    inset?: { l: number; t: number; r: number; b: number };
  };
}

export interface VmlGroupDrawing {
  vector: VectorImageSpec;
  frames: GroupFrame[];
}

export interface VmlGroupReaders {
  textbox: (el: OoxmlNode) => GroupFrame['textbox'] | null;
  color: (v: string | undefined) => string | undefined;
  /** CSS length (pt, in, …) → px; bare numbers are px. */
  length: (v: string | undefined) => number | undefined;
  /** v:shapetype id → o:spt, for `type="#id"` references. */
  shapeType: (id: string) => number | undefined;
  /** A picture member's v:imagedata → data URL, or undefined. */
  image: (el: OoxmlNode) => string | undefined;
  /** Mark a node consumed (the audit). */
  read: (n: OoxmlNode) => void;
}

/** Maps a member's coordinates (its parent's units) to group px. */
interface Frame {
  ox: number;
  oy: number;
  sx: number;
  sy: number;
  /** px of the parent's coordinate origin. */
  px: number;
  py: number;
}

const pair = (
  v: string | undefined,
  dx: number,
  dy: number,
): [number, number] => {
  const p = (v ?? '').split(',').map((s) => Number.parseFloat(s));
  return [Number.isFinite(p[0]) ? p[0] : dx, Number.isFinite(p[1]) ? p[1] : dy];
};

/** `left:2050;top:411;width:3402;height:2042` — unitless numbers are the
 *  parent's coordinate units; lengths with units are px (outermost only). */
function memberBox(
  style: string,
  units: (v: string | undefined) => number | undefined,
): {
  x: number;
  y: number;
  w: number;
  h: number;
  flipX: boolean;
  flipY: boolean;
} {
  const get: Record<string, string> = {};
  for (const part of style.split(';')) {
    const i = part.indexOf(':');
    if (i > 0) get[part.slice(0, i).trim()] = part.slice(i + 1).trim();
  }
  const flip = get['flip'] ?? '';
  return {
    x: units(get['left'] ?? get['margin-left']) ?? 0,
    y: units(get['top'] ?? get['margin-top']) ?? 0,
    w: units(get['width']) ?? 0,
    h: units(get['height']) ?? 0,
    flipX: flip.includes('x'),
    flipY: flip.includes('y'),
  };
}

const num = (v: string | undefined): number | undefined => {
  if (v === undefined) return undefined;
  const n = Number.parseFloat(v);
  return Number.isFinite(n) ? n : undefined;
};

type Pt = { x: number; y: number };

/**
 * A VML `path` string → subpaths in the shape's coordinate space.
 * Commands: m l r t c v x e nf ns, the quadrant arcs qx/qy, and the
 * box-arcs at/ar/wa/wr. Null when the path references formulas (`@n`).
 */
export function parseVmlPath(
  path: string,
): { pts: Pt[]; closed: boolean }[] | null {
  if (path.includes('@')) return null;
  // Commands and their argument runs. An empty argument between commas is
  // 0 — Word writes `r,22` for "0 across, 22 down" and `m,l,21600` for
  // the corners of the unit box.
  const steps: { cmd: string; n: number[] }[] = [];
  // Commands run together (`xe` is close, then end), so the two-letter
  // ones are matched by name and anything else is one letter.
  for (const m of path.matchAll(
    /(at|ar|wa|wr|qx|qy|nf|ns|ae|al|h[a-i]|[a-z])([^a-z]*)/gi,
  )) {
    const args = m[2].trim();
    const n = args
      ? args
          .split(',')
          .flatMap((part) =>
            part.trim() === ''
              ? [0]
              : (part.match(/-?\d*\.?\d+/g) ?? []).map(Number),
          )
      : [];
    steps.push({ cmd: m[1].toLowerCase(), n });
  }
  const subs: { pts: Pt[]; closed: boolean }[] = [];
  let cur: Pt[] = [];
  let at: Pt = { x: 0, y: 0 };
  const flush = (closed: boolean) => {
    if (cur.length > 1) subs.push({ pts: cur, closed });
    cur = [];
  };
  const cubic = (p0: Pt, p1: Pt, p2: Pt, p3: Pt) => {
    for (let k = 1; k <= 12; k++) {
      const t = k / 12;
      const u = 1 - t;
      cur.push({
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
  };
  const arc = (
    l: number,
    t: number,
    r: number,
    b: number,
    sx: number,
    sy: number,
    ex: number,
    ey: number,
    move: boolean,
    clockwise: boolean,
  ) => {
    const cx = (l + r) / 2;
    const cy = (t + b) / 2;
    const rx = Math.abs(r - l) / 2 || 1;
    const ry = Math.abs(b - t) / 2 || 1;
    const a0 = Math.atan2((sy - cy) / ry, (sx - cx) / rx);
    let a1 = Math.atan2((ey - cy) / ry, (ex - cx) / rx);
    // at/wa run counter-clockwise, ar/wr clockwise (screen coordinates).
    if (clockwise) while (a1 <= a0) a1 += Math.PI * 2;
    else while (a1 >= a0) a1 -= Math.PI * 2;
    const start = { x: cx + rx * Math.cos(a0), y: cy + ry * Math.sin(a0) };
    if (move) {
      flush(false);
      cur = [start];
    } else cur.push(start);
    const steps = Math.max(
      4,
      Math.ceil((Math.abs(a1 - a0) / (Math.PI * 2)) * 48),
    );
    for (let k = 1; k <= steps; k++) {
      const a = a0 + ((a1 - a0) * k) / steps;
      cur.push({ x: cx + rx * Math.cos(a), y: cy + ry * Math.sin(a) });
    }
    at = cur[cur.length - 1];
  };
  for (const { cmd, n } of steps) {
    switch (cmd) {
      case 'm':
        flush(false);
        if (n.length >= 2) {
          at = { x: n[0], y: n[1] };
          cur = [at];
          for (let k = 2; k + 1 < n.length; k += 2)
            cur.push((at = { x: n[k], y: n[k + 1] }));
        }
        break;
      case 't':
        flush(false);
        if (n.length >= 2) {
          at = { x: at.x + n[0], y: at.y + n[1] };
          cur = [at];
        }
        break;
      case 'l':
        for (let k = 0; k + 1 < n.length; k += 2)
          cur.push((at = { x: n[k], y: n[k + 1] }));
        break;
      case 'r':
        if (cur.length === 0) cur = [at];
        for (let k = 0; k + 1 < n.length; k += 2)
          cur.push((at = { x: at.x + n[k], y: at.y + n[k + 1] }));
        break;
      case 'c':
        if (cur.length === 0) cur = [at];
        for (let k = 0; k + 5 < n.length; k += 6) {
          const p3 = { x: n[k + 4], y: n[k + 5] };
          cubic(at, { x: n[k], y: n[k + 1] }, { x: n[k + 2], y: n[k + 3] }, p3);
          at = p3;
        }
        break;
      case 'v':
        if (cur.length === 0) cur = [at];
        for (let k = 0; k + 5 < n.length; k += 6) {
          const o = at;
          const p3 = { x: o.x + n[k + 4], y: o.y + n[k + 5] };
          cubic(
            o,
            { x: o.x + n[k], y: o.y + n[k + 1] },
            { x: o.x + n[k + 2], y: o.y + n[k + 3] },
            p3,
          );
          at = p3;
        }
        break;
      case 'qx':
      case 'qy': {
        // Elliptical quadrants, alternating axis first.
        let xFirst = cmd === 'qx';
        if (cur.length === 0) cur = [at];
        for (let k = 0; k + 1 < n.length; k += 2) {
          const end = { x: n[k], y: n[k + 1] };
          const corner = xFirst ? { x: end.x, y: at.y } : { x: at.x, y: end.y };
          const c1 = {
            x: at.x + (corner.x - at.x) * 0.5523,
            y: at.y + (corner.y - at.y) * 0.5523,
          };
          const c2 = {
            x: end.x + (corner.x - end.x) * 0.5523,
            y: end.y + (corner.y - end.y) * 0.5523,
          };
          cubic(at, c1, c2, end);
          at = end;
          xFirst = !xFirst;
        }
        break;
      }
      case 'at':
      case 'ar':
      case 'wa':
      case 'wr':
        for (let k = 0; k + 7 < n.length; k += 8)
          arc(
            n[k],
            n[k + 1],
            n[k + 2],
            n[k + 3],
            n[k + 4],
            n[k + 5],
            n[k + 6],
            n[k + 7],
            cmd === 'at' || cmd === 'ar',
            cmd === 'ar' || cmd === 'wr',
          );
        break;
      case 'x':
        if (cur.length) {
          const first = cur[0];
          flush(true);
          at = first;
        }
        break;
      case 'e':
        flush(false);
        break;
      default:
      // nf / ns and the rest: no geometry.
    }
  }
  flush(false);
  return subs;
}

const MEMBER_TAGS = new Set([
  'v:group',
  'v:rect',
  'v:roundrect',
  'v:oval',
  'v:line',
  'v:shape',
]);

/** Draws members into one display list and frame list. */
function drawer(readers: VmlGroupReaders) {
  const ops: VectorOp[] = [];
  const frames: GroupFrame[] = [];
  const toPx = (f: Frame, x: number, y: number): Pt => ({
    x: f.px + (x - f.ox) * f.sx,
    y: f.py + (y - f.oy) * f.sy,
  });

  /** One drawn member at `tl`, `pw` × `ph` px; `f` maps its parent's
   *  coordinates (a line's from / to). */
  const member = (
    el: OoxmlNode,
    f: Frame,
    tl: Pt,
    pw: number,
    ph: number,
    box: { flipX: boolean; flipY: boolean },
  ): void => {
    if (child(el, 'v:imagedata')) {
      const src = readers.image(el);
      if (src)
        ops.push({
          kind: 'image',
          x: tl.x,
          y: tl.y,
          width: pw,
          height: ph,
          src,
        });
      return;
    }
    // Fill and stroke: VML draws both unless told not to.
    const stroked =
      attrOf(el, 'stroked') !== 'f' && attrOf(el, 'stroked') !== 'false';
    const strokeEl = child(el, 'v:stroke');
    const fillEl = child(el, 'v:fill');
    const filledAttr = attrOf(el, 'filled');
    let filled = filledAttr !== 'f' && filledAttr !== 'false';
    if (attrOf(fillEl, 'on') === 'f' || attrOf(fillEl, 'on') === 'false')
      filled = false;
    const fill = filled
      ? (readers.color(attrOf(el, 'fillcolor')) ??
        readers.color(attrOf(fillEl, 'color')) ??
        '#ffffff')
      : null;
    const strokeColor = stroked
      ? (readers.color(attrOf(el, 'strokecolor')) ??
        readers.color(attrOf(strokeEl, 'color')) ??
        '#000000')
      : null;
    const strokeW =
      readers.length(
        attrOf(el, 'strokeweight') ?? attrOf(strokeEl, 'weight'),
      ) ?? 1;
    const dashStyle = attrOf(strokeEl, 'dashstyle');
    const dash =
      dashStyle && dashStyle !== 'solid'
        ? (() => {
            const n = dashStyle.split(/[\s,]+/).map(Number);
            return n.every((v) => Number.isFinite(v) && v >= 0) ? n : [4, 3];
          })()
        : undefined;
    const outline = (pts: Pt[], closed: boolean) => {
      if (!strokeColor || pts.length < 2) return;
      if (closed)
        ops.push({
          kind: 'polygon',
          points: pts,
          stroke: strokeColor,
          strokeWidth: strokeW,
        });
      else
        ops.push({
          kind: 'polyline',
          points: pts,
          stroke: strokeColor,
          strokeWidth: strokeW,
          ...(dash && { dash }),
        });
    };
    const rect = (x: number, y: number, w: number, h: number): Pt[] => [
      { x, y },
      { x: x + w, y },
      { x: x + w, y: y + h },
      { x, y: y + h },
    ];

    if (el.name === 'v:line') {
      const [x1, y1] = pair(attrOf(el, 'from'), 0, 0);
      const [x2, y2] = pair(attrOf(el, 'to'), 0, 0);
      const a = toPx(f, x1, y1);
      const b = toPx(f, x2, y2);
      if (strokeColor)
        ops.push({
          kind: 'line',
          x1: a.x,
          y1: a.y,
          x2: b.x,
          y2: b.y,
          width: strokeW,
          color: strokeColor,
          ...(dash && { dash }),
        });
      return;
    }

    let geometry: { pts: Pt[]; closed: boolean }[] = [];
    if (el.name === 'v:rect')
      geometry = [{ pts: rect(tl.x, tl.y, pw, ph), closed: true }];
    else if (el.name === 'v:roundrect') {
      const r = Math.min(pw, ph) * 0.1;
      const pts: Pt[] = [];
      const corner = (cx: number, cy: number, a0: number) => {
        for (let k = 0; k <= 4; k++) {
          const a = a0 + (k / 4) * (Math.PI / 2);
          pts.push({ x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) });
        }
      };
      corner(tl.x + pw - r, tl.y + r, -Math.PI / 2);
      corner(tl.x + pw - r, tl.y + ph - r, 0);
      corner(tl.x + r, tl.y + ph - r, Math.PI / 2);
      corner(tl.x + r, tl.y + r, Math.PI);
      geometry = [{ pts, closed: true }];
    } else if (el.name === 'v:oval') {
      const pts: Pt[] = [];
      for (let k = 0; k < 48; k++) {
        const a = (k / 48) * Math.PI * 2;
        pts.push({
          x: tl.x + pw / 2 + (pw / 2) * Math.cos(a),
          y: tl.y + ph / 2 + (ph / 2) * Math.sin(a),
        });
      }
      geometry = [{ pts, closed: true }];
    } else {
      // v:shape: a freeform path in its own coordinate space, or a frame.
      const path = attrOf(el, 'path');
      const ref = attrOf(el, 'type');
      const spt = ref?.startsWith('#')
        ? readers.shapeType(ref.slice(1))
        : undefined;
      if (path) {
        const sub = parseVmlPath(path);
        if (sub) {
          const [scw, sch] = pair(attrOf(el, 'coordsize'), 1000, 1000);
          const [sox, soy] = pair(attrOf(el, 'coordorigin'), 0, 0);
          const map = (p: Pt): Pt => {
            let u = (p.x - sox) / (scw || 1);
            let v = (p.y - soy) / (sch || 1);
            if (box.flipX) u = 1 - u;
            if (box.flipY) v = 1 - v;
            return { x: tl.x + u * pw, y: tl.y + v * ph };
          };
          geometry = sub.map((s) => ({
            pts: s.pts.map(map),
            closed: s.closed,
          }));
        }
      } else if (spt === 202 || spt === 1 || child(el, 'v:textbox')) {
        geometry = [{ pts: rect(tl.x, tl.y, pw, ph), closed: true }];
      } else if (spt === 32 || attrOf(el, 'o:connectortype')) {
        const a = box.flipY ? { x: tl.x, y: tl.y + ph } : tl;
        const b = box.flipY
          ? { x: tl.x + pw, y: tl.y }
          : { x: tl.x + pw, y: tl.y + ph };
        if (strokeColor)
          ops.push({
            kind: 'line',
            x1: a.x,
            y1: a.y,
            x2: b.x,
            y2: b.y,
            width: strokeW,
            color: strokeColor,
            ...(dash && { dash }),
          });
      }
    }
    for (const g of geometry) {
      if (fill && g.closed) ops.push({ kind: 'polygon', points: g.pts, fill });
      outline(g.pts, g.closed);
    }
    const tb = child(el, 'v:textbox');
    if (tb) {
      const content = readers.textbox(el);
      if (content)
        frames.push({
          x: tl.x,
          y: tl.y,
          width: pw,
          height: ph,
          textbox: content,
        });
    }
  };

  const walk = (parent: OoxmlNode, f: Frame): void => {
    for (const el of parent.children) {
      if (el.name === 'v:shapetype') {
        readers.read(el);
        continue;
      }
      if (!MEMBER_TAGS.has(el.name)) continue;
      readers.read(el);
      const box = memberBox(attrOf(el, 'style') ?? '', num);
      const tl = toPx(f, box.x, box.y);
      const pw = box.w * f.sx;
      const ph = box.h * f.sy;
      if (el.name === 'v:group') {
        const [ccw, cch] = pair(attrOf(el, 'coordsize'), 1000, 1000);
        const [cox, coy] = pair(attrOf(el, 'coordorigin'), 0, 0);
        walk(el, {
          ox: cox,
          oy: coy,
          sx: pw / (ccw || 1),
          sy: ph / (cch || 1),
          px: tl.x,
          py: tl.y,
        });
      } else member(el, f, tl, pw, ph, box);
    }
  };

  return {
    member,
    walk,
    result: (width: number, height: number): VmlGroupDrawing => ({
      vector: { width, height, ops },
      frames,
    }),
  };
}

export function vmlGroupDrawing(
  group: OoxmlNode,
  width: number,
  height: number,
  readers: VmlGroupReaders,
): VmlGroupDrawing {
  const d = drawer(readers);
  const [cw, ch] = pair(attrOf(group, 'coordsize'), 1000, 1000);
  const [ox, oy] = pair(attrOf(group, 'coordorigin'), 0, 0);
  d.walk(group, {
    ox,
    oy,
    sx: width / (cw || 1),
    sy: height / (ch || 1),
    px: 0,
    py: 0,
  });
  return d.result(width, height);
}

/** A lone freeform v:shape (a `path` of its own — an arrow, a page frame)
 *  drawn the way a group member is, filling a `width` × `height` box. */
export function vmlShapeDrawing(
  shape: OoxmlNode,
  width: number,
  height: number,
  readers: VmlGroupReaders,
): VmlGroupDrawing {
  const d = drawer(readers);
  const box = memberBox(attrOf(shape, 'style') ?? '', () => undefined);
  const identity: Frame = { ox: 0, oy: 0, sx: 1, sy: 1, px: 0, py: 0 };
  d.member(shape, identity, { x: 0, y: 0 }, width, height, box);
  return d.result(width, height);
}
