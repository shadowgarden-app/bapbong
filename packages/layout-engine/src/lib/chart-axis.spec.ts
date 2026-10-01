import { autoScale, ticks } from './chart-axis.js';

describe('autoScale (Word/Excel automatic value axis)', () => {
  // The five charts of D-2609-LSTQ, as Word 365 drew them.
  it.each([
    [0, 0.862, 0, 1, 0.1],
    [0, 0.688, 0, 0.8, 0.1],
    [0, 0.4, 0, 0.45, 0.05],
    [239.4, 1521.39, 0, 1600, 200],
    [224.65, 1914.16, 0, 2500, 500],
  ])('%s…%s → %s…%s step %s', (lo, hi, min, max, major) => {
    expect(autoScale(lo, hi)).toEqual({ min, max, major });
  });

  // The chart probe (Word 365, 216pt plots).
  it.each([
    [-35, 120, -60, 140, 20],
    [0.003, 0.012, 0, 0.014, 0.002],
    [45678, 123456, 0, 140000, 20000],
    [50, 55, 47, 56, 1],
    [-20, -5, -25, 0, 5],
    [3, 9, 0, 10, 1],
    [1000, 1002, 999, 1002.5, 0.5],
  ])('probe %s…%s → %s…%s step %s', (lo, hi, min, max, major) => {
    expect(autoScale(lo, hi)).toEqual({ min, max, major });
  });

  it('spaces ticks further apart on a short axis', () => {
    // 888…1521 on a 108pt plot: 200 apart would be 13.5pt — Word takes 500.
    expect(autoScale(0, 1521.39, { length: 108, minSpacing: 15 })).toEqual({
      min: 0,
      max: 2000,
      major: 500,
    });
    expect(autoScale(0, 1521.39, { length: 216, minSpacing: 15 }).major).toBe(
      200,
    );
  });

  it('keeps explicit bounds and unit', () => {
    expect(autoScale(0, 7, { min: 0, max: 10, major: 2 })).toEqual({
      min: 0,
      max: 10,
      major: 2,
    });
  });

  it('lists the ticks', () => {
    expect(ticks({ min: 0, max: 1, major: 0.1 })).toHaveLength(11);
    expect(ticks({ min: 0, max: 1, major: 0.1 })[3]).toBe(0.3);
  });
});
