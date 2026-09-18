import { fitScale, planSnapshot } from './region-snapshot.js';

const page = () => ({ width: 800, height: 1100 });

describe('planSnapshot', () => {
  it('cuts the bounding box of a page’s rects, grown by the pad', () => {
    const pieces = planSnapshot(
      [
        { pageIndex: 0, x: 100, y: 200, width: 500, height: 20 },
        { pageIndex: 0, x: 80, y: 220, width: 300, height: 20 },
      ],
      page,
      12,
    );
    expect(pieces).toEqual([
      {
        pageIndex: 0,
        crop: { x: 68, y: 188, width: 544, height: 64 },
        marks: [
          { x: 32, y: 12, width: 500, height: 20 },
          { x: 12, y: 32, width: 300, height: 20 },
        ],
      },
    ]);
  });

  it('clamps the band to the page', () => {
    const [piece] = planSnapshot(
      [{ pageIndex: 0, x: 4, y: 1090, width: 796, height: 10 }],
      page,
      12,
    );
    expect(piece.crop).toEqual({ x: 0, y: 1078, width: 800, height: 22 });
    expect(piece.marks[0]).toEqual({ x: 4, y: 12, width: 796, height: 10 });
  });

  it('cuts one band per page, in page order, for a passage across pages', () => {
    const pieces = planSnapshot(
      [
        { pageIndex: 3, x: 100, y: 40, width: 200, height: 20 },
        { pageIndex: 2, x: 100, y: 1000, width: 200, height: 20 },
      ],
      page,
      0,
    );
    expect(pieces.map((p) => p.pageIndex)).toEqual([2, 3]);
  });

  it('drops empty rects and pages it cannot size', () => {
    const pieces = planSnapshot(
      [
        { pageIndex: 0, x: 10, y: 10, width: 0, height: 20 },
        { pageIndex: 9, x: 10, y: 10, width: 50, height: 20 },
      ],
      (i) => (i === 9 ? null : page()),
      4,
    );
    expect(pieces).toEqual([]);
  });
});

describe('fitScale', () => {
  it('shrinks the long side to the cap', () => {
    expect(fitScale(2400, 600, 1200, 2)).toBe(0.5);
  });

  it('never scales past the bitmap’s own density', () => {
    expect(fitScale(300, 100, 1200, 2)).toBe(2);
  });
});
