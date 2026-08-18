/**
 * Windows Enhanced Metafiles (`.emf`) in a .docx.
 *
 * Browsers decode no metafile, so an EMF picture used to import as an
 * `application/octet-stream` data URL the painter could not draw — a blank of
 * the right size where Word shows a logo. Rendering EMF in general means a GDI
 * interpreter; this module does not attempt one. It handles the one shape of
 * EMF Word itself produces constantly: a metafile that is nothing but a
 * BITMAP — pasting a picture from the clipboard, or an older Word converting a
 * DIB, wraps the pixels in a single EMR_STRETCHBLT/EMR_STRETCHDIBITS record
 * (a 1.7MB "logo" in one application form is exactly that: HEADER, three
 * viewport records, one STRETCHBLT, EOF). For those the pixels are lifted out
 * of the record and re-framed as a .bmp — the DIB inside a metafile record IS
 * a BMP minus its 14-byte file header — which every browser and Word decode.
 *
 * Anything that draws vectors (paths, polygons, text) or blits more than one
 * bitmap is left alone: `emfBitmapDataUrl` returns null and the caller keeps
 * the original bytes, so a later real renderer (or the export) still has them.
 *
 * Record layouts: [MS-EMF] 2.3.1 (bitmap records) and 2.3.4 (control records).
 * Every offset below is from the start of its record.
 */

/** EMR record types (the [MS-EMF] `Type` field). */
const EMR = {
  HEADER: 1,
  EOF: 14,
  BITBLT: 76,
  STRETCHBLT: 77,
  SETDIBITSTODEVICE: 80,
  STRETCHDIBITS: 81,
  ALPHABLEND: 114,
  GDICOMMENT: 70,
} as const;

/** Records that carry a bitmap, with where the DIB header (`offBmiSrc` /
 *  `cbBmiSrc`) and the bits (`offBitsSrc` / `cbBitsSrc`) fields sit. */
const BITMAP_RECORDS: Record<number, { bmi: number; bits: number }> = {
  // EMR_BITBLT / EMR_STRETCHBLT / EMR_ALPHABLEND share a prefix: Bounds(16)
  // xDest yDest cxDest cyDest (16) rop(4) xSrc ySrc (8) XformSrc(24)
  // BkColorSrc(4) UsageSrc(4) → offBmiSrc at 84, cbBmiSrc 88, offBitsSrc 92,
  // cbBitsSrc 96.
  [EMR.BITBLT]: { bmi: 84, bits: 92 },
  [EMR.STRETCHBLT]: { bmi: 84, bits: 92 },
  [EMR.ALPHABLEND]: { bmi: 84, bits: 92 },
  // EMR_STRETCHDIBITS / EMR_SETDIBITSTODEVICE: Bounds(16) xDest yDest xSrc
  // ySrc cxSrc cySrc (24) → offBmiSrc at 48, cbBmiSrc 52, offBitsSrc 56,
  // cbBitsSrc 60.
  [EMR.STRETCHDIBITS]: { bmi: 48, bits: 56 },
  [EMR.SETDIBITSTODEVICE]: { bmi: 48, bits: 56 },
};

/** Records that only set state, select objects or comment — a metafile made
 *  of these plus one bitmap record still draws nothing but the bitmap. Every
 *  type outside this set and BITMAP_RECORDS is a drawing primitive (polygons,
 *  paths, text, fills) and disqualifies the file. */
const STATE_RECORDS = new Set<number>([
  EMR.HEADER,
  EMR.EOF,
  9, // SETWINDOWEXTEX
  10, // SETWINDOWORGEX
  11, // SETVIEWPORTEXTEX
  12, // SETVIEWPORTORGEX
  13, // SETBRUSHORGEX
  17, // SETMAPMODE
  18, // SETBKMODE
  19, // SETPOLYFILLMODE
  20, // SETROP2
  21, // SETSTRETCHBLTMODE
  22, // SETTEXTALIGN
  23, // SETCOLORADJUSTMENT
  24, // SETTEXTCOLOR
  25, // SETBKCOLOR
  26, // OFFSETCLIPRGN
  27, // MOVETOEX
  28, // SETMETARGN
  29, // EXCLUDECLIPRECT
  30, // INTERSECTCLIPRECT
  33, // SAVEDC
  34, // RESTOREDC
  35, // SETWORLDTRANSFORM
  36, // MODIFYWORLDTRANSFORM
  37, // SELECTOBJECT
  38, // CREATEPEN
  39, // CREATEBRUSHINDIRECT
  40, // DELETEOBJECT
  67, // SELECTPALETTE
  68, // CREATEPALETTE
  69, // SETPALETTEENTRIES
  EMR.GDICOMMENT,
  75, // SELECTCLIPPATH (state; a path that only clips)
  82, // EXTCREATEFONTINDIRECTW
  95, // EXTCREATEPEN
  98, // SETICMMODE
  99, // CREATECOLORSPACE
  100, // SETCOLORSPACE
  101, // DELETECOLORSPACE
  115, // SETLAYOUT
]);

/**
 * The bitmap inside a bitmap-only EMF as a `data:image/bmp;base64,…` URL, or
 * null when the metafile is not that (vectors, text, several blits, palette
 * indices, or malformed).
 */
export function emfBitmapDataUrl(bytes: Uint8Array): string | null {
  const dib = extractSoleDib(bytes);
  return dib ? `data:image/bmp;base64,${toBase64(dib)}` : null;
}

/** The DIB of a bitmap-only EMF, framed as a complete BMP file. */
export function extractSoleDib(bytes: Uint8Array): Uint8Array | null {
  if (bytes.length < 88) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // ENHMETAHEADER: Type 1, then Size; the ' EMF' signature at 40.
  if (view.getUint32(0, true) !== EMR.HEADER) return null;
  if (view.getUint32(40, true) !== 0x464d4520) return null; // ' EMF'
  let bitmap: { start: number; type: number } | null = null;
  let off = 0;
  while (off + 8 <= bytes.length) {
    const type = view.getUint32(off, true);
    const size = view.getUint32(off + 4, true);
    if (size < 8 || off + size > bytes.length) return null;
    if (BITMAP_RECORDS[type]) {
      if (bitmap) return null; // more than one blit: not a plain picture
      bitmap = { start: off, type };
    } else if (!STATE_RECORDS.has(type)) {
      return null; // a drawing primitive
    }
    if (type === EMR.EOF) break;
    off += size;
  }
  if (!bitmap) return null;
  const f = BITMAP_RECORDS[bitmap.type];
  const rec = bitmap.start;
  const size = view.getUint32(rec + 4, true);
  if (rec + f.bits + 8 > bytes.length) return null;
  const offBmi = view.getUint32(rec + f.bmi, true);
  const cbBmi = view.getUint32(rec + f.bmi + 4, true);
  const offBits = view.getUint32(rec + f.bits, true);
  const cbBits = view.getUint32(rec + f.bits + 4, true);
  // Usage: DIB_RGB_COLORS (0) — palette entries are colours. DIB_PAL_COLORS
  // (1) indexes a logical palette this reader does not carry: bail.
  const usageOff =
    bitmap.type === EMR.STRETCHDIBITS || bitmap.type === EMR.SETDIBITSTODEVICE
      ? 64
      : 80;
  if (view.getUint32(rec + usageOff, true) !== 0) return null;
  if (
    cbBmi < 40 ||
    cbBits === 0 ||
    offBmi + cbBmi > size ||
    offBits + cbBits > size
  )
    return null;
  const bmi = bytes.subarray(rec + offBmi, rec + offBmi + cbBmi);
  const bits = bytes.subarray(rec + offBits, rec + offBits + cbBits);
  // A DIB in a record is BITMAPINFO (header + colour table) followed by the
  // pixel array — a BMP file minus its 14-byte BITMAPFILEHEADER.
  const out = new Uint8Array(14 + cbBmi + cbBits);
  const ov = new DataView(out.buffer);
  out[0] = 0x42; // 'B'
  out[1] = 0x4d; // 'M'
  ov.setUint32(2, out.length, true);
  ov.setUint32(10, 14 + cbBmi, true); // pixel array offset
  out.set(bmi, 14);
  out.set(bits, 14 + cbBmi);
  return out;
}

/** Base64 without a browser-only `btoa` string round-trip (Node has Buffer;
 *  the browser gets a chunked btoa — a 1.7MB DIB is common). */
function toBase64(bytes: Uint8Array): string {
  const B = (
    globalThis as {
      Buffer?: { from(b: Uint8Array): { toString(e: string): string } };
    }
  ).Buffer;
  if (B) return B.from(bytes).toString('base64');
  let s = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK)
    s += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  return btoa(s);
}
