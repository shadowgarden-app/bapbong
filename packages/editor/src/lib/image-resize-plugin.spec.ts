import { resizeRect } from './image-resize-plugin';

const base = { x: 100, y: 50, width: 200, height: 100 };

describe('resizeRect', () => {
  it('corner drag keeps the aspect ratio, anchored at the opposite corner', () => {
    // SE +40 px on x (dominant): scale 1.2 → 240×120, origin unchanged.
    expect(resizeRect(base, 'se', 40, 10, false)).toEqual({ x: 100, y: 50, width: 240, height: 120 });
    // NW −40 px on x: same growth, but the SE corner stays put.
    const nw = resizeRect(base, 'nw', -40, -10, false);
    expect(nw.width).toBeCloseTo(240);
    expect(nw.height).toBeCloseTo(120);
    expect(nw.x + nw.width).toBeCloseTo(300); // anchored right edge
    expect(nw.y + nw.height).toBeCloseTo(150); // anchored bottom edge
  });

  it('corner drag with Shift resizes freely', () => {
    expect(resizeRect(base, 'se', 40, 10, true)).toEqual({ x: 100, y: 50, width: 240, height: 110 });
  });

  it('edge drags are free on their single axis', () => {
    expect(resizeRect(base, 'e', 30, 999, false)).toEqual({ x: 100, y: 50, width: 230, height: 100 });
    expect(resizeRect(base, 's', 999, 20, false)).toEqual({ x: 100, y: 50, width: 200, height: 120 });
    // West edge anchors the right edge.
    expect(resizeRect(base, 'w', 20, 0, false)).toEqual({ x: 120, y: 50, width: 180, height: 100 });
  });

  it('edge drag with Shift keeps the aspect ratio', () => {
    const r = resizeRect(base, 'e', 100, 0, true); // 200→300 = ×1.5
    expect(r.width).toBeCloseTo(300);
    expect(r.height).toBeCloseTo(150);
  });

  it('clamps both dimensions to the minimum size', () => {
    const r = resizeRect(base, 'se', -500, -500, true);
    expect(r.width).toBe(16);
    expect(r.height).toBe(16);
    // With aspect the clamp applies to the smaller dimension.
    const a = resizeRect(base, 'se', -500, -500, false);
    expect(a.height).toBeCloseTo(16);
    expect(a.width).toBeCloseTo(32); // 2:1 aspect preserved
  });
});
