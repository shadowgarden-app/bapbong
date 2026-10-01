/**
 * Automatic value-axis scaling, Word/Excel's way.
 *
 * Measured on the five charts of D-2609-LSTQ (Word 365 PDF):
 *   0…0.862 → 0…100% step 10%   0…0.688 → 0…80% step 10%
 *   0…0.40  → 0…45% step 5%     0…1521  → 0…1600 step 200
 *   0…1914  → 0…2500 step 500
 * One rule reproduces all five: pad the data range by 5% beyond its far
 * end, then take the SMALLEST major unit of the 1-2-5 series whose ticks
 * cover the padded range in at most 10 intervals; the ends are the unit
 * multiples that enclose it.
 *
 * The minimum snaps to zero when the data sits far enough from it — the
 * documented Excel rule: all values ≥ 0 and the spread is more than a sixth
 * of the maximum (mirror image for all-negative data).
 */

export interface AxisScale {
  min: number;
  max: number;
  major: number;
}

const MAX_INTERVALS = 10;

/** 1, 2, 5, 10, 20, 50 … starting at or below `x`. */
function unitsFrom(x: number): number[] {
  const p = 10 ** Math.floor(Math.log10(x));
  const out: number[] = [];
  for (let k = 0; k < 6; k++)
    for (const m of [1, 2, 5]) out.push(m * p * 10 ** (k - 1));
  return out;
}

/** Floating-point safe floor/ceil to a multiple. */
const floorTo = (v: number, u: number) => Math.floor(v / u + 1e-9) * u;
const ceilTo = (v: number, u: number) => Math.ceil(v / u - 1e-9) * u;
const clean = (v: number) => Number(v.toPrecision(12));

export function autoScale(
  dataMin: number,
  dataMax: number,
  fixed: { min?: number; max?: number; major?: number } = {},
): AxisScale {
  let lo = Math.min(dataMin, dataMax);
  let hi = Math.max(dataMin, dataMax);
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) {
    lo = 0;
    hi = 1;
  }
  if (lo === hi) {
    // A flat series: Excel opens a range around it from zero.
    if (hi > 0) lo = 0;
    else if (hi < 0) hi = 0;
    else hi = 1;
  }
  // Snap the far end to zero.
  let snapLo = lo;
  let snapHi = hi;
  if (lo >= 0 && (hi - lo) / hi > 1 / 6) snapLo = 0;
  if (hi <= 0 && (hi - lo) / -lo > 1 / 6) snapHi = 0;
  if (fixed.min !== undefined) snapLo = fixed.min;
  if (fixed.max !== undefined) snapHi = fixed.max;
  const span = snapHi - snapLo || Math.abs(snapHi) || 1;
  // The 5% headroom goes beyond whichever end the data reaches past zero.
  const padHi =
    fixed.max !== undefined ? snapHi : hi > 0 ? hi + 0.05 * span : snapHi;
  const padLo =
    fixed.min !== undefined ? snapLo : lo < 0 ? lo - 0.05 * span : snapLo;

  const pick = (u: number): AxisScale => ({
    min: fixed.min !== undefined ? fixed.min : clean(floorTo(padLo, u)),
    max: fixed.max !== undefined ? fixed.max : clean(ceilTo(padHi, u)),
    major: u,
  });
  if (fixed.major !== undefined && fixed.major > 0) return pick(fixed.major);
  const range = padHi - padLo || 1;
  for (const u of unitsFrom(range / MAX_INTERVALS)) {
    const s = pick(u);
    // Data exactly on a tick stays inside: ceil already includes it.
    if ((s.max - s.min) / u <= MAX_INTERVALS + 1e-9) return s;
  }
  const u = 10 ** Math.ceil(Math.log10(range));
  return pick(u);
}

/** The tick values of a scale, min to max inclusive. */
export function ticks(s: AxisScale): number[] {
  const out: number[] = [];
  const n = Math.round((s.max - s.min) / s.major);
  for (let i = 0; i <= n && i < 1000; i++) out.push(clean(s.min + i * s.major));
  return out;
}
