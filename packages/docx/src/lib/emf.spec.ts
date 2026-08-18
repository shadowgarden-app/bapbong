import {
  emfBitmapDataUrl,
  extractSoleDib,
  extractSoleWmfDib,
  wmfBitmapDataUrl,
} from './emf.js';

/** A minimal EMF: HEADER + `extra` records + EOF. */
function emf(...records: Uint8Array[]): Uint8Array {
  const header = new Uint8Array(88);
  const hv = new DataView(header.buffer);
  hv.setUint32(0, 1, true); // EMR_HEADER
  hv.setUint32(4, 88, true);
  hv.setUint32(40, 0x464d4520, true); // ' EMF'
  const eof = new Uint8Array(20);
  const ev = new DataView(eof.buffer);
  ev.setUint32(0, 14, true);
  ev.setUint32(4, 20, true);
  const parts = [header, ...records, eof];
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/** An EMR_STRETCHBLT carrying a 2×1 24bpp DIB (red, blue). */
function stretchBlt(usage = 0): Uint8Array {
  const bmi = new Uint8Array(40);
  const bv = new DataView(bmi.buffer);
  bv.setUint32(0, 40, true); // biSize
  bv.setInt32(4, 2, true); // width
  bv.setInt32(8, 1, true); // height
  bv.setUint16(12, 1, true); // planes
  bv.setUint16(14, 24, true); // bpp
  const bits = new Uint8Array([0, 0, 255, 255, 0, 0, 0, 0]); // BGR BGR pad
  const rec = new Uint8Array(108 + 40 + 8);
  const rv = new DataView(rec.buffer);
  rv.setUint32(0, 77, true); // EMR_STRETCHBLT
  rv.setUint32(4, rec.length, true);
  rv.setUint32(80, usage, true); // UsageSrc
  rv.setUint32(84, 108, true); // offBmiSrc
  rv.setUint32(88, 40, true); // cbBmiSrc
  rv.setUint32(92, 148, true); // offBitsSrc
  rv.setUint32(96, 8, true); // cbBitsSrc
  rec.set(bmi, 108);
  rec.set(bits, 148);
  return rec;
}

/** An EMR_POLYLINE16 with no points — any drawing primitive will do. */
function polyline(): Uint8Array {
  const rec = new Uint8Array(28);
  const rv = new DataView(rec.buffer);
  rv.setUint32(0, 87, true);
  rv.setUint32(4, 28, true);
  return rec;
}

describe('emf: bitmap-only metafiles', () => {
  it('re-frames the sole DIB as a BMP file', () => {
    const bmp = extractSoleDib(emf(stretchBlt()))!;
    expect(bmp).not.toBeNull();
    expect(String.fromCharCode(bmp[0], bmp[1])).toBe('BM');
    const v = new DataView(bmp.buffer);
    expect(v.getUint32(2, true)).toBe(bmp.length); // file size
    expect(v.getUint32(10, true)).toBe(14 + 40); // pixel array offset
    expect(v.getInt32(14 + 4, true)).toBe(2); // width from the DIB header
    expect(bmp[14 + 40 + 2]).toBe(255); // first pixel's R
    expect(emfBitmapDataUrl(emf(stretchBlt()))).toMatch(
      /^data:image\/bmp;base64,Qk/,
    ); // "BM" → "Qk"
  });

  it('leaves anything that is not a plain picture alone', () => {
    // A vector primitive beside the bitmap.
    expect(emfBitmapDataUrl(emf(polyline(), stretchBlt()))).toBeNull();
    // Two bitmaps.
    expect(emfBitmapDataUrl(emf(stretchBlt(), stretchBlt()))).toBeNull();
    // Palette-indexed pixels (DIB_PAL_COLORS) need a palette we don't carry.
    expect(emfBitmapDataUrl(emf(stretchBlt(1)))).toBeNull();
    // Not an EMF at all.
    expect(emfBitmapDataUrl(new Uint8Array(200))).toBeNull();
    // No bitmap.
    expect(emfBitmapDataUrl(emf())).toBeNull();
  });
});

/** A WMF: META_HEADER + records (Size in words, Function) + EOF. */
function wmf(...records: Uint8Array[]): Uint8Array {
  const header = new Uint8Array(18);
  const hv = new DataView(header.buffer);
  hv.setUint16(0, 1, true); // memory metafile
  hv.setUint16(2, 9, true); // header size in words
  const eof = new Uint8Array(6);
  new DataView(eof.buffer).setUint32(0, 3, true); // 3 words, function 0
  const parts = [header, ...records, eof];
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/** META_STRETCHDIB with a 2×1 8bpp paletted DIB (2 colours). */
function stretchDib(usage = 0): Uint8Array {
  const bmi = new Uint8Array(40 + 2 * 4);
  const bv = new DataView(bmi.buffer);
  bv.setUint32(0, 40, true);
  bv.setInt32(4, 2, true);
  bv.setInt32(8, 1, true);
  bv.setUint16(12, 1, true);
  bv.setUint16(14, 8, true); // 8bpp
  bv.setUint32(32, 2, true); // biClrUsed = 2 → 8-byte colour table
  bmi.set([0, 0, 255, 0, 255, 0, 0, 0], 40); // red, blue (BGRx)
  const bits = new Uint8Array([0, 1, 0, 0]); // two indices + pad
  const rec = new Uint8Array(28 + bmi.length + bits.length);
  const rv = new DataView(rec.buffer);
  rv.setUint32(0, rec.length / 2, true); // size in words
  rv.setUint16(4, 0x0f43, true); // META_STRETCHDIB
  rv.setUint16(10, usage, true); // ColorUsage
  rec.set(bmi, 28);
  rec.set(bits, 28 + bmi.length);
  return rec;
}

/** META_POLYGON with no points — a drawing primitive. */
function wmfPolygon(): Uint8Array {
  const rec = new Uint8Array(8);
  const rv = new DataView(rec.buffer);
  rv.setUint32(0, 4, true);
  rv.setUint16(4, 0x0324, true);
  return rec;
}

describe('wmf: bitmap-only metafiles', () => {
  it('re-frames the sole DIB — colour table included — as a BMP file', () => {
    const bmp = extractSoleWmfDib(wmf(stretchDib()))!;
    expect(bmp).not.toBeNull();
    expect(String.fromCharCode(bmp[0], bmp[1])).toBe('BM');
    const v = new DataView(bmp.buffer);
    expect(v.getUint32(2, true)).toBe(bmp.length);
    // Pixel array sits after the 40-byte header AND the 2-entry palette.
    expect(v.getUint32(10, true)).toBe(14 + 40 + 8);
    expect(bmp[14 + 40 + 8]).toBe(0); // first index → red
    expect(bmp[14 + 40 + 8 + 1]).toBe(1); // second → blue
    // A placeable header in front changes nothing.
    const placeable = new Uint8Array(22 + wmf(stretchDib()).length);
    new DataView(placeable.buffer).setUint32(0, 0x9ac6cdd7, true);
    placeable.set(wmf(stretchDib()), 22);
    expect(wmfBitmapDataUrl(placeable)).toMatch(/^data:image\/bmp;base64,Qk/);
  });

  it('leaves vector or palette-indexed WMFs alone', () => {
    expect(wmfBitmapDataUrl(wmf(wmfPolygon(), stretchDib()))).toBeNull();
    expect(wmfBitmapDataUrl(wmf(stretchDib(), stretchDib()))).toBeNull();
    expect(wmfBitmapDataUrl(wmf(stretchDib(1)))).toBeNull();
    expect(wmfBitmapDataUrl(new Uint8Array(30))).toBeNull();
  });
});
