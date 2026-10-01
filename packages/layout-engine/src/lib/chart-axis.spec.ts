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
