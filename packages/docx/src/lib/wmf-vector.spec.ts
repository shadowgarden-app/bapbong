import { wmfVectorImage } from './wmf-vector.js';
import type { WmfLineOp, WmfPolygonOp, WmfTextOp } from './wmf-vector.js';

// Builders producing the MathType shape of WMF: placeable header, window
// org/ext, fonts selected per run, EXTTEXTOUT with per-char advances,
// MOVETO/LINETO strokes. Layouts follow [MS-WMF]; sizes are in 16-bit words.

function rec(fn: number, params: number[]): Uint8Array {
  const out = new Uint8Array(6 + params.length * 2);
  const v = new DataView(out.buffer);
  v.setUint32(0, out.length / 2, true);
  v.setUint16(4, fn, true);
  params.forEach((val, i) => v.setInt16(6 + i * 2, val, true));
  return out;
}

/** A record whose parameters are raw bytes (strings, LogFont). */
function recBytes(fn: number, params: Uint8Array): Uint8Array {
  const padded = params.length + (params.length & 1);
  const out = new Uint8Array(6 + padded);
  const v = new DataView(out.buffer);
  v.setUint32(0, out.length / 2, true);
  v.setUint16(4, fn, true);
  out.set(params, 6);
  return out;
}

function wmf(records: Uint8Array[], { placeable = true } = {}): Uint8Array {
  const eof = rec(0x0000, []);
  const body = [...records, eof];
  const bodyLen = body.reduce((s, r) => s + r.length, 0);
  const header = new Uint8Array(placeable ? 40 : 18);
  const v = new DataView(header.buffer);
  let off = 0;
  if (placeable) {
    v.setUint32(0, 0x9ac6cdd7, true);
    v.setInt16(6, 0, true); // bbox left
    v.setInt16(8, 0, true); // top
    v.setInt16(10, 1200, true); // right
    v.setInt16(12, 400, true); // bottom
    v.setUint16(14, 1440, true); // units per inch
    off = 22;
  }
  v.setUint16(off, 1, true); // memory metafile
  v.setUint16(off + 2, 9, true); // header words
  v.setUint32(off + 6, (header.length - off + bodyLen) / 2, true);
  const out = new Uint8Array(header.length + bodyLen);
  out.set(header);
  let o = header.length;
  for (const r of body) {
    out.set(r, o);
    o += r.length;
  }
  return out;
}

function logFont(
  face: string,
  {
    height = -213,
    weight = 400,
    italic = false,
    underline = false,
    charset = 0,
  } = {},
): Uint8Array {
  const name = new TextEncoder().encode(face);
  const out = new Uint8Array(18 + name.length + 1);
  const v = new DataView(out.buffer);
  v.setInt16(0, height, true);
  v.setInt16(8, weight, true);
  out[10] = italic ? 1 : 0;
  out[11] = underline ? 1 : 0;
  out[13] = charset;
  out.set(name, 18);
  return out;
}

function extTextOut(
  x: number,
  y: number,
  text: string,
  dx: number[] | null,
): Uint8Array {
  const chars = new TextEncoder().encode(text);
  const padded = chars.length + (chars.length & 1);
  const out = new Uint8Array(8 + padded + (dx ? dx.length * 2 : 0));
  const v = new DataView(out.buffer);
  v.setInt16(0, y, true);
  v.setInt16(2, x, true);
  v.setInt16(4, chars.length, true);
  v.setUint16(6, 0, true); // fwOpts
  out.set(chars, 8);
  dx?.forEach((d, i) => v.setInt16(8 + padded + i * 2, d, true));
  return recBytes(0x0a32, out);
}

const SET_WINDOW = [
  rec(0x020b, [-40, -100]), // SETWINDOWORG: y, x
  rec(0x020c, [400, 1200]), // SETWINDOWEXT: y, x
];

describe('wmfVectorImage', () => {
  it('parses the MathType shape: fonts, kerned text, strokes, in z-order', () => {
    const image = wmfVectorImage(
      wmf([
        ...SET_WINDOW,
        rec(0x0209, [0x0000, 0x0000]), // SETTEXTCOLOR black
        rec(0x012e, [24]), // SETTEXTALIGN TA_BASELINE
        recBytes(0x02fb, logFont('Times New Roman', { italic: true })),
        rec(0x012d, [0]), // select the font
        extTextOut(0, 120, 'x', [12]),
        rec(0x02fa, [0, 8, 0, 0, 0]), // CREATEPENINDIRECT width 8, black
        rec(0x012d, [1]), // select the pen
        rec(0x0214, [200, 0]), // MOVETO y=200 x=0
        rec(0x0213, [200, 300]), // LINETO y=200 x=300
        recBytes(0x02fb, logFont('Symbol', { charset: 2 })),
        rec(0x012d, [2]),
        extTextOut(50, 320, 'w', null),
      ]),
    );
    expect(image).not.toBeNull();
    expect(image!.width).toBe(1200);
    expect(image!.height).toBe(400);
    expect(image!.unitsPerInch).toBe(1440);
    expect(image!.ops.map((op) => op.kind)).toEqual(['text', 'line', 'text']);

    const [x, bar, omega] = image!.ops as [WmfTextOp, WmfLineOp, WmfTextOp];
    // The window origin (-100, -40) is subtracted from every coordinate.
    expect([x.x, x.y]).toEqual([100, 160]);
    expect(String.fromCharCode(...x.bytes)).toBe('x');
    expect(x.dx).toEqual([12]);
    expect(x.face).toBe('Times New Roman');
    expect(x.italic).toBe(true);
    expect(x.align).toBe(24);

    expect([bar.x1, bar.y1, bar.x2, bar.y2]).toEqual([100, 240, 400, 240]);
    expect(bar.width).toBe(8);

    expect(omega.face).toBe('Symbol');
    expect(omega.charset).toBe(2);
    expect(omega.dx).toBeNull();
  });

  it('reuses the lowest freed object-table slot', () => {
    const image = wmfVectorImage(
      wmf([
        ...SET_WINDOW,
        recBytes(0x02fb, logFont('Times New Roman')), // slot 0
        recBytes(0x02fb, logFont('Symbol')), // slot 1
        rec(0x01f0, [0]), // free slot 0
        recBytes(0x02fb, logFont('MT Extra')), // must land in slot 0
        rec(0x012d, [0]),
        extTextOut(0, 0, 'a', null),
      ]),
    );
    expect((image!.ops[0] as WmfTextOp).face).toBe('MT Extra');
  });

  it('keeps a non-placeable file, reporting no units-per-inch', () => {
    const image = wmfVectorImage(
      wmf([...SET_WINDOW, extTextOut(0, 0, 'a', null)], { placeable: false }),
    );
    expect(image).not.toBeNull();
    expect(image!.unitsPerInch).toBe(0);
  });

  it('tolerates Escape records (MathType embeds MTEF there)', () => {
    const image = wmfVectorImage(
      wmf([
        ...SET_WINDOW,
        recBytes(0x0626, new TextEncoder().encode('..MathType..')),
        extTextOut(0, 0, 'a', null),
      ]),
    );
    expect(image!.ops).toHaveLength(1);
  });

  it('parses a filled polygon (MathType arrowheads) under pen and brush', () => {
    const image = wmfVectorImage(
      wmf([
        ...SET_WINDOW,
        rec(0x02fa, [0, 4, 0, 0, 0]), // pen width 4, black
        rec(0x02fc, [0, 0xff00, 0x0000, 0]), // BS_SOLID, COLORREF 0x0000FF00
        rec(0x012d, [0]),
        rec(0x012d, [1]),
        // POLYGON: count, then x,y pairs (x first, unlike MOVETO).
        rec(0x0324, [3, 0, 0, 10, 0, 5, 10]),
      ]),
    );
    const poly = image!.ops[0] as WmfPolygonOp;
    expect(poly.kind).toBe('polygon');
    // Window origin (-100, -40) subtracted from every vertex.
    expect(poly.points).toEqual([
      { x: 100, y: 40 },
      { x: 110, y: 40 },
      { x: 105, y: 50 },
    ]);
    expect(poly.fill).toBe(0x00ff00);
    expect(poly.strokeWidth).toBe(4);
  });

  it('keeps a hollow-brush polygon unfilled', () => {
    const image = wmfVectorImage(
      wmf([
        ...SET_WINDOW,
        rec(0x02fc, [1, 0, 0, 0]), // BS_NULL
        rec(0x012d, [0]),
        rec(0x0324, [3, 0, 0, 10, 0, 5, 10]),
      ]),
    );
    expect((image!.ops[0] as WmfPolygonOp).fill).toBeNull();
  });

  it('rejects drawing primitives outside the supported set', () => {
    expect(
      wmfVectorImage(
        wmf([...SET_WINDOW, rec(0x041b, [10, 10, 0, 0])]), // RECTANGLE
      ),
    ).toBeNull();
  });

  it('rejects bitmap records — those belong to the DIB extraction path', () => {
    expect(
      wmfVectorImage(wmf([...SET_WINDOW, rec(0x0f43, [0, 0, 0, 0, 0, 0])])),
    ).toBeNull();
  });

  it('rejects a file with no EOF or no window extent', () => {
    const noExt = wmfVectorImage(wmf([extTextOut(0, 0, 'a', null)]));
    expect(noExt).toBeNull();
    const truncated = wmf([...SET_WINDOW, extTextOut(0, 0, 'a', null)]);
    expect(wmfVectorImage(truncated.subarray(0, truncated.length - 8))).toBe(
      null,
    );
  });

  it('rejects selecting or deleting a slot that was never filled', () => {
    expect(wmfVectorImage(wmf([...SET_WINDOW, rec(0x012d, [3])]))).toBeNull();
    expect(wmfVectorImage(wmf([...SET_WINDOW, rec(0x01f0, [0])]))).toBeNull();
  });
});
