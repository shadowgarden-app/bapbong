import { cursorFor, resizeRect, snapAngle, toLocal } from './image-resize-plugin';

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

describe('cursorFor', () => {
  it('maps handles to bidirectional resize cursors', () => {
    expect(cursorFor('n', 0)).toBe('ns-resize');
    expect(cursorFor('ne', 0)).toBe('nesw-resize');
    expect(cursorFor('e', 0)).toBe('ew-resize');
    expect(cursorFor('se', 0)).toBe('nwse-resize');
  });

  it('rotates the cursor with the frame', () => {
    // At 90° the north handle points east.
    expect(cursorFor('n', 90)).toBe('ew-resize');
    expect(cursorFor('se', 90)).toBe('nesw-resize');
    // Between steps it quantizes to the nearest 45°.
    expect(cursorFor('n', 40)).toBe('nesw-resize');
  });
});

describe('wrap-mode conversions (floatForWrapMode)', () => {
  const keep = { hOffset: 120, vOffset: 340 };

  it('maps float attrs to strip modes', async () => {
    const { wrapModeOf } = await import('./image-resize-plugin');
    expect(wrapModeOf(null)).toBe('inline');
    expect(wrapModeOf({ wrap: 'square' })).toBe('square');
    expect(wrapModeOf({ wrap: 'topAndBottom' })).toBe('topAndBottom');
    expect(wrapModeOf({ wrap: 'none' })).toBe('front');
    expect(wrapModeOf({ wrap: 'none', behind: true })).toBe('behind');
  });

  it('is a no-op when the mode is already in effect', async () => {
    const { floatForWrapMode } = await import('./image-resize-plugin');
    expect(floatForWrapMode(null, 'inline', keep)).toBeUndefined();
    expect(floatForWrapMode({ wrap: 'square' }, 'square', keep)).toBeUndefined();
  });

  it('inline → float pins the rendered position via margin-relative offsets', async () => {
    const { floatForWrapMode } = await import('./image-resize-plugin');
    const f = floatForWrapMode(null, 'square', keep) as Record<string, unknown>;
    expect(f).toMatchObject({
      wrap: 'square',
      hRel: 'margin',
      hOffset: 120,
      vRel: 'margin',
      vOffset: 340,
    });
  });

  it('float → inline returns null; wrap switches keep the position attrs', async () => {
    const { floatForWrapMode } = await import('./image-resize-plugin');
    const cur = { wrap: 'square', hOffset: 50, vOffset: 60, hRel: 'margin', vRel: 'paragraph' };
    expect(floatForWrapMode(cur, 'inline', keep)).toBeNull();
    const toBreak = floatForWrapMode(cur, 'topAndBottom', keep) as Record<string, unknown>;
    expect(toBreak).toMatchObject({ wrap: 'topAndBottom', hOffset: 50, vOffset: 60 });
  });

  it('behind pairs with wrapNone and clears when leaving z-modes', async () => {
    const { floatForWrapMode } = await import('./image-resize-plugin');
    const behind = floatForWrapMode({ wrap: 'square', hOffset: 5 }, 'behind', keep) as Record<string, unknown>;
    expect(behind).toMatchObject({ wrap: 'none', behind: true, hOffset: 5 });
    const front = floatForWrapMode(behind, 'front', keep) as Record<string, unknown>;
    expect(front['behind']).toBeUndefined();
    expect(front).toMatchObject({ wrap: 'none' });
    const square = floatForWrapMode(behind, 'square', keep) as Record<string, unknown>;
    expect(square['behind']).toBeUndefined();
  });
});
