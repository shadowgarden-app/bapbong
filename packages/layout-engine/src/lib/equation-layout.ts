/**
 * 2D equation typesetting: an EqNode[] AST → a vector display list.
 *
 * The equation node is an inline ATOM. This module measures its tree with
 * the engine's own text measurer and produces the SAME contract the painter
 * already replays for metafile previews (VectorImageSpec ops), plus the
 * box's baseline (so the token seats on the line via `raise`) and the
 * editable slot rectangles the equation plugin steers by. No painter or
 * selection changes — an equation is a vector image that knows its slots.
 *
 * Conventions: one coordinate space per equation, px, origin at the TOP-LEFT
 * of the box, baseline at `ascent`. Internally rows lay out baseline-relative
 * (y = 0 on the baseline, up is negative) and are shifted once at the end.
 */
import type {
  EqNode,
  EqSlotRect,
  FontMetrics,
  FontSpec,
  MeasureMetrics,
  MeasureText,
  VectorOp,
} from '@shadow-garden/bapbong-contracts';
import {
  bigLimits,
  radShowDeg,
  scrSlots,
} from '@shadow-garden/bapbong-contracts';

export interface EquationLayoutResult {
  width: number;
  height: number;
  /** Box top → the equation's main baseline. */
  ascent: number;
  ops: VectorOp[];
  slots: EqSlotRect[];
}

/** Script (sub/sup, fraction parts) size, as Word scales them. */
const SCRIPT = 0.66;
/** The math axis (fraction bar center) above the baseline, in em. */
const AXIS = 0.27;
/** Gap between a fraction bar and its rows, in em. */
const FRAC_GAP = 0.1;
/** Horizontal padding around fraction rows, in em. */
const FRAC_PAD = 0.08;
/** Superscript baseline rise / subscript drop, in em of the BASE size. */
const SUP_RISE = 0.42;
const SUB_DROP = 0.18;
/** Fallback vertical metrics (em fractions) when no metrics source. */
const ASC = 0.9;
const DESC = 0.22;

const FAMILY = 'Times New Roman';

/** A laid-out row: baseline-relative ops, box extents, slot records. */
interface RowBox {
  w: number;
  asc: number;
  desc: number;
  /** Ops with y relative to THIS row's baseline. */
  ops: VectorOp[];
  slots: EqSlotRect[];
  /** Caret x stops for the row's own items. */
  caretXs: number[];
}

const shift = (ops: VectorOp[], dx: number, dy: number): VectorOp[] =>
  ops.map((op) =>
    op.kind === 'text'
      ? { ...op, x: op.x + dx, y: op.y + dy }
      : op.kind === 'line'
        ? {
            ...op,
            x1: op.x1 + dx,
            y1: op.y1 + dy,
            x2: op.x2 + dx,
            y2: op.y2 + dy,
          }
        : {
            ...op,
            points: op.points.map((p) => ({ x: p.x + dx, y: p.y + dy })),
          },
  );

const shiftSlots = (
  slots: EqSlotRect[],
  dx: number,
  dy: number,
): EqSlotRect[] => slots.map((s) => ({ ...s, x: s.x + dx, y: s.y + dy }));

export function layoutEquation(
  ast: EqNode[],
  sizePt: number,
  measure: MeasureText,
  metrics?: MeasureMetrics,
): EquationLayoutResult {
  const font = (pt: number): FontSpec => ({
    family: FAMILY,
    sizePt: pt,
    bold: false,
    italic: false,
  });
  const vert = (pt: number): FontMetrics => {
    const em = (pt * 96) / 72;
    if (metrics) {
      const m = metrics(font(pt));
      if (m && m.ascent > 0) return m;
    }
    return { ascent: ASC * em, descent: DESC * em };
  };

  const row = (
    nodes: EqNode[],
    pt: number,
    path: (number | string)[],
  ): RowBox => {
    const em = (pt * 96) / 72;
    const m = vert(pt);
    const ops: VectorOp[] = [];
    const slots: EqSlotRect[] = [];
    const caretXs: number[] = [0];
    let x = 0;
    let asc = m.ascent * 0.72; // an empty row still has x-height presence
    let desc = m.descent * 0.6;

    // An empty row renders Word's dotted placeholder slot.
    if (nodes.length === 0) {
      const w = em * 0.62;
      const h = em * 0.72;
      const y0 = -h;
      const dash = (
        x1: number,
        y1: number,
        x2: number,
        y2: number,
      ): VectorOp => ({
        kind: 'line',
        x1,
        y1,
        x2,
        y2,
        width: 1,
        color: '#9a9790',
      });
      // Dotted-ish square: four hairlines.
      ops.push(
        dash(0, y0, w, y0),
        dash(w, y0, w, 0),
        dash(w, 0, 0, 0),
        dash(0, 0, 0, y0),
      );
      slots.push({
        path,
        x: 0,
        y: -h,
        width: w,
        height: h + m.descent * 0.3,
        caretXs: [w / 2],
        em,
      });
      return { w, asc: h, desc: m.descent * 0.6, ops, slots, caretXs: [w / 2] };
    }

    let pending = ''; // consecutive chars coalesce into one text op
    const flushText = (): void => {
      if (!pending) return;
      ops.push({
        kind: 'text',
        x: x - measure(pending, font(pt)),
        y: 0,
        text: pending,
        size: em,
        family: FAMILY,
        color: '#000000',
      });
      pending = '';
    };

    nodes.forEach((n, i) => {
      if (n.t === 'chr') {
        pending += n.ch;
        x += measure(n.ch, font(pt));
        asc = Math.max(asc, m.ascent);
        desc = Math.max(desc, m.descent);
        caretXs.push(x);
        return;
      }
      flushText();
      const box = child(n, pt, [...path, i]);
      ops.push(...shift(box.ops, x, 0));
      slots.push(...shiftSlots(box.slots, x, 0));
      x += box.w;
      asc = Math.max(asc, box.asc);
      desc = Math.max(desc, box.desc);
      caretXs.push(x);
    });
    flushText();

    slots.unshift({
      path,
      x: 0,
      y: -asc,
      width: x,
      height: asc + desc,
      caretXs,
      em,
    });
    return { w: x, asc, desc, ops, slots, caretXs };
  };

  const child = (
    n: Exclude<EqNode, { t: 'chr' }>,
    pt: number,
    path: (number | string)[],
  ): RowBox => {
    const em = (pt * 96) / 72;
    const spt = pt * SCRIPT;
    switch (n.t) {
      case 'frac': {
        const num = row(n.num, spt, [...path, 'num']);
        const den = row(n.den, spt, [...path, 'den']);
        const pad = FRAC_PAD * em;
        const w = Math.max(num.w, den.w) + 2 * pad;
        const axis = -AXIS * em;
        const gap = FRAC_GAP * em;
        const bar = Math.max(1, em * 0.055);
        const numBase = axis - gap - num.desc;
        const denBase = axis + gap + den.asc;
        const ops: VectorOp[] = [
          ...shift(num.ops, pad + (w - 2 * pad - num.w) / 2, numBase),
          ...shift(den.ops, pad + (w - 2 * pad - den.w) / 2, denBase),
          {
            kind: 'line',
            x1: 0,
            y1: axis,
            x2: w,
            y2: axis,
            width: bar,
            color: '#000000',
          },
        ];
        const slots = [
          ...shiftSlots(num.slots, pad + (w - 2 * pad - num.w) / 2, numBase),
          ...shiftSlots(den.slots, pad + (w - 2 * pad - den.w) / 2, denBase),
        ];
        return {
          w,
          asc: -(numBase - num.asc),
          desc: denBase + den.desc,
          ops,
          slots,
          caretXs: [0, w],
        };
      }
      case 'scr': {
        const base = row(n.base, pt, [...path, 'base']);
        const sub = row(n.sub, spt, [...path, 'sub']);
        const sup = row(n.sup, spt, [...path, 'sup']);
        // A script shows the rows it HAS, empty or not — an empty one is the
        // placeholder box the user types into. Reading this off emptiness
        // instead would make a fresh superscript template indistinguishable
        // from a fresh subscript one.
        const kind = scrSlots(n);
        const showSub = kind !== 'sup';
        const showSup = kind !== 'sub';
        const supBase = -SUP_RISE * em;
        const subBase = SUB_DROP * em;
        const scriptW = Math.max(showSub ? sub.w : 0, showSup ? sup.w : 0);
        const ops = [...base.ops];
        const slots = [...base.slots];
        if (showSup) {
          ops.push(...shift(sup.ops, base.w, supBase));
          slots.push(...shiftSlots(sup.slots, base.w, supBase));
        }
        if (showSub) {
          ops.push(...shift(sub.ops, base.w, subBase));
          slots.push(...shiftSlots(sub.slots, base.w, subBase));
        }
        return {
          w: base.w + scriptW,
          asc: Math.max(base.asc, showSup ? -(supBase - sup.asc) : 0),
          desc: Math.max(base.desc, showSub ? subBase + sub.desc : 0),
          ops,
          slots,
          caretXs: [0, base.w + scriptW],
        };
      }
      case 'rad': {
        const body = row(n.body, pt, [...path, 'body']);
        const deg = radShowDeg(n)
          ? row(n.deg, pt * 0.55, [...path, 'deg'])
          : null;
        const h = body.asc + body.desc;
        // The radical sign, scaled to wrap the body's height.
        const glyphPt = Math.max(
          pt,
          (h / (1.12 * ((96 / 72) * 1))) * (72 / 96) * 1.0,
        );
        const gem = (glyphPt * 96) / 72;
        const signW = measure('√', font(glyphPt));
        const degW = deg ? Math.max(0, deg.w - signW * 0.4) : 0;
        const top = -body.asc - em * 0.08;
        const ops: VectorOp[] = [];
        const slots: EqSlotRect[] = [];
        if (deg) {
          ops.push(...shift(deg.ops, 0, top + gem * 0.35));
          slots.push(...shiftSlots(deg.slots, 0, top + gem * 0.35));
        }
        ops.push({
          kind: 'text',
          x: degW,
          y: top + gem * 0.98,
          text: '√',
          size: gem * 1.06,
          family: FAMILY,
          color: '#000000',
        });
        ops.push({
          kind: 'line',
          x1: degW + signW,
          y1: top,
          x2: degW + signW + body.w + em * 0.1,
          y2: top,
          width: Math.max(1, em * 0.05),
          color: '#000000',
        });
        ops.push(...shift(body.ops, degW + signW + em * 0.05, 0));
        slots.push(...shiftSlots(body.slots, degW + signW + em * 0.05, 0));
        return {
          w: degW + signW + body.w + em * 0.15,
          asc: -top,
          desc: body.desc,
          ops,
          slots,
          caretXs: [0, degW + signW + body.w + em * 0.15],
        };
      }
      case 'fence': {
        const body = row(n.body, pt, [...path, 'body']);
        const h = body.asc + body.desc;
        const em1 = em;
        const scale = Math.max(1, h / (em1 * 1.12));
        const fpt = pt * scale;
        const fem = (fpt * 96) / 72;
        const fm = vert(fpt);
        // Center the fence on the body's vertical middle.
        const mid = (-body.asc + body.desc) / 2;
        const fenceBase =
          mid + (fm.ascent - fm.descent) / 2 - (fm.ascent - fem * 0.72);
        const lw = n.l ? measure(n.l, font(fpt)) : 0;
        const rw = n.r ? measure(n.r, font(fpt)) : 0;
        const ops: VectorOp[] = [];
        if (n.l)
          ops.push({
            kind: 'text',
            x: 0,
            y: fenceBase,
            text: n.l,
            size: fem,
            family: FAMILY,
            color: '#000000',
          });
        ops.push(...shift(body.ops, lw, 0));
        if (n.r)
          ops.push({
            kind: 'text',
            x: lw + body.w,
            y: fenceBase,
            text: n.r,
            size: fem,
            family: FAMILY,
            color: '#000000',
          });
        return {
          w: lw + body.w + rw,
          asc: Math.max(body.asc, -(fenceBase - fm.ascent * 0.8)),
          desc: Math.max(body.desc, fenceBase + fm.descent),
          ops,
          slots: shiftSlots(body.slots, lw, 0),
          caretXs: [0, lw + body.w + rw],
        };
      }
      case 'big': {
        const opPt = pt * 1.4;
        const opEm = (opPt * 96) / 72;
        const opW = measure(n.op, font(opPt));
        const lo = row(n.lo, spt, [...path, 'lo']);
        const hi = row(n.hi, spt, [...path, 'hi']);
        const body = row(n.body, pt, [...path, 'body']);
        const opBase = em * 0.06; // big glyphs sit slightly low
        // A limit the operator does not HAVE is neither drawn nor offered
        // to the caret; an empty one it does have draws its placeholder box
        // like any other row.
        const lim = bigLimits(n);
        const scriptW = Math.max(lim.lo ? lo.w : 0, lim.hi ? hi.w : 0);
        const supBase = -SUP_RISE * em * 1.25;
        const subBase = SUB_DROP * em * 2.2;
        const ops: VectorOp[] = [
          {
            kind: 'text',
            x: 0,
            y: opBase,
            text: n.op,
            size: opEm,
            family: FAMILY,
            color: '#000000',
          },
        ];
        const slots: EqSlotRect[] = [];
        if (lim.hi) {
          ops.push(...shift(hi.ops, opW, supBase));
          slots.push(...shiftSlots(hi.slots, opW, supBase));
        }
        if (lim.lo) {
          ops.push(...shift(lo.ops, opW, subBase));
          slots.push(...shiftSlots(lo.slots, opW, subBase));
        }
        const bodyX = opW + scriptW + em * 0.08;
        ops.push(...shift(body.ops, bodyX, 0));
        slots.push(...shiftSlots(body.slots, bodyX, 0));
        return {
          w: bodyX + body.w,
          asc: Math.max(
            opEm * 0.85,
            lim.hi ? -(supBase - hi.asc) : 0,
            body.asc,
          ),
          desc: Math.max(opEm * 0.3, lim.lo ? subBase + lo.desc : 0, body.desc),
          ops,
          slots,
          caretXs: [0, bodyX + body.w],
        };
      }
    }
  };

  const root = row(ast, sizePt, []);
  // Shift everything into top-left space; pad hairline so strokes at the
  // edges don't clip.
  const padX = 1;
  const width = Math.max(2, root.w + 2 * padX);
  const ascent = root.asc + 1;
  const height = root.asc + root.desc + 2;
  return {
    width,
    height,
    ascent,
    ops: shift(root.ops, padX, ascent),
    slots: shiftSlots(root.slots, padX, ascent),
  };
}
