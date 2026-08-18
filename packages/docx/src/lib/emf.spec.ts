import { emfBitmapDataUrl, extractSoleDib } from './emf.js';

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
