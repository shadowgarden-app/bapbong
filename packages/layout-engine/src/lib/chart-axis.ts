/**
 * Automatic value-axis scaling, Word/Excel's way — measured in Word 365
 * (the chart probe: 14 data sets, 3 chart heights; and D-2609-LSTQ).
 *
 *  - The far end gets 5% headroom beyond the data.
 *  - The near end snaps to zero when the data sits far enough from it —
 *    all values ≥ 0 and the spread is more than a sixth of the maximum
 *    (mirror image for all-negative data). Otherwise it drops by half the
 *    spread: 50…55 → 47…56, 1000…1002 → 999…1002.5.
 *  - The major unit is the SMALLEST of the 1-2-5 series whose ticks cover
 *    that range in at most 10 intervals — and, on a short axis, intervals
 *    no closer than `minSpacing` (888…1521 is 0…1600 by 200 on a 216pt
 *    plot, but 0…2000 by 500 on a 108pt one).
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
  for (let k = 0; k < 8; k++)
    for (const m of [1, 2, 5]) out.push(m * p * 10 ** (k - 1));
  return out;
}

/** Floating-point safe floor/ceil to a multiple. */
const floorTo = (v: number, u: number) => Math.floor(v / u + 1e-9) * u;
const ceilTo = (v: number, u: number) => Math.ceil(v / u - 1e-9) * u;
const clean = (v: number) => Number(v.toPrecision(12));

export interface ScaleOptions {
  min?: number;
  max?: number;
  major?: number;
  /** Axis length and the closest two ticks may sit (same unit). */
  length?: number;
  minSpacing?: number;
}

export function autoScale(
  dataMin: number,
  dataMax: number,
  fixed: ScaleOptions = {},
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
  const spread = hi - lo;
  let padLo: number;
  let padHi: number;
  if (lo >= 0) {
    padLo = spread / hi > 1 / 6 ? 0 : lo - spread / 2;
    padHi = hi + 0.05 * (hi - padLo);
  } else if (hi <= 0) {
    padHi = spread / -lo > 1 / 6 ? 0 : hi + spread / 2;
    padLo = lo - 0.05 * (padHi - lo);
  } else {
    padHi = hi + 0.05 * spread;
    padLo = lo - 0.05 * spread;
  }
  if (fixed.min !== undefined) padLo = fixed.min;
  if (fixed.max !== undefined) padHi = fixed.max;

  const pick = (u: number): AxisScale => ({
    min: fixed.min !== undefined ? fixed.min : clean(floorTo(padLo, u)),
    max: fixed.max !== undefined ? fixed.max : clean(ceilTo(padHi, u)),
    major: u,
  });
  if (fixed.major !== undefined && fixed.major > 0) return pick(fixed.major);
  const fits =
    fixed.length && fixed.minSpacing
      ? Math.max(1, Math.floor(fixed.length / fixed.minSpacing + 1e-9))
      : MAX_INTERVALS;
  const maxIntervals = Math.min(MAX_INTERVALS, fits);
  const range = padHi - padLo || 1;
  for (const u of unitsFrom(range / MAX_INTERVALS)) {
    const s = pick(u);
    if ((s.max - s.min) / u <= maxIntervals + 1e-9) return s;
  }
  return pick(10 ** Math.ceil(Math.log10(range)));
}

/** The tick values of a scale, min to max inclusive. */
export function ticks(s: AxisScale): number[] {
  const out: number[] = [];
  const n = Math.round((s.max - s.min) / s.major);
  for (let i = 0; i <= n && i < 1000; i++) out.push(clean(s.min + i * s.major));
  return out;
}
