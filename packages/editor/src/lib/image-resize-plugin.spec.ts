import { resizeRect, snapAngle, toLocal } from './image-resize-plugin';

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

describe('snapAngle', () => {
  it('pulls angles within 3° onto the cardinals and normalizes', () => {
    expect(snapAngle(2, false)).toBe(0);
    expect(snapAngle(-2, false)).toBe(0); // 358 → 360 → 0
    expect(snapAngle(88.5, false)).toBe(90);
    expect(snapAngle(181, false)).toBe(180);
    expect(snapAngle(268, false)).toBe(270);
    // Outside the pull zone the angle is free.
    expect(snapAngle(15, false)).toBe(15);
    expect(snapAngle(84, false)).toBe(84);
  });

  it('Shift steps by 15°', () => {
    expect(snapAngle(22, true)).toBe(15);
    expect(snapAngle(23, true)).toBe(30);
    expect(snapAngle(-7, true)).toBe(0); // 353 → 360 → 0
  });
});

describe('toLocal', () => {
  it('inverse-rotates a point around the rect center', () => {
    const rect = { x: 0, y: 0, width: 100, height: 100 }; // center (50, 50)
    // Frame rotated 90° cw: its unrotated top-center (50, 0) appears at
    // (100, 50) on screen — mapping that screen point back gives (50, 0).
    const p = toLocal(100, 50, rect, 90);
    expect(p.x).toBeCloseTo(50);
    expect(p.y).toBeCloseTo(0);
    // No rotation → identity.
    expect(toLocal(7, 9, rect, 0)).toEqual({ x: 7, y: 9 });
  });
});
