/**
 * Windows metafiles (`.emf`, `.wmf`) in a .docx.
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

// ── WMF ─────────────────────────────────────────────────────────────
// The 16-bit ancestor ([MS-WMF]): an optional 22-byte placeable header
// (magic 0x9AC6CDD7), META_HEADER (18 bytes), then records of
// Size(4, in 16-bit WORDS) + Function(2) + parameters. The bitmap records
// carry the DIB — header, colour table, bits — as one blob to the record's
// end, so unlike EMF there is no offset table: the colour-table size has to be
// worked out from the DIB header to know where the bits start.

const WMF_PLACEABLE = 0x9ac6cdd7;
/** Function → byte offset of the DIB from the record start (Size+Function =
 *  6, then the fixed parameters). */
const WMF_BITMAP_RECORDS: Record<number, number> = {
  0x0f43: 28, // META_STRETCHDIB: rop(4) usage(2) + 8 shorts
  0x0b41: 26, // META_DIBSTRETCHBLT: rop(4) + 8 shorts
  0x0940: 22, // META_DIBBITBLT: rop(4) + 6 shorts
  0x0d33: 24, // META_SETDIBTODEV: usage(2) scans(2) start(2) + 6 shorts
};
/** ColorUsage lives at these offsets (DIB_RGB_COLORS = 0); the two blts have
 *  none (always RGB). */
const WMF_USAGE_OFFSET: Record<number, number> = { 0x0f43: 10, 0x0d33: 6 };
/** State / object / control functions — see STATE_RECORDS for the rule. */
const WMF_STATE_RECORDS = new Set<number>([
  0x0000, // EOF
  0x0035, // REALIZEPALETTE
  0x00f7, // CREATEPALETTE
  0x0102, // SETBKMODE
  0x0103, // SETMAPMODE
  0x0104, // SETROP2
  0x0105, // SETRELABS
  0x0106, // SETPOLYFILLMODE
  0x0107, // SETSTRETCHBLTMODE
  0x0108, // SETTEXTCHAREXTRA
  0x0127, // RESTOREDC
  0x012d, // SELECTOBJECT
  0x012e, // SETTEXTALIGN
  0x0139, // RESIZEPALETTE
  0x0142, // DIBCREATEPATTERNBRUSH
  0x0149, // SETLAYOUT
  0x001e, // SAVEDC
  0x01f0, // DELETEOBJECT
  0x01f9, // CREATEPATTERNBRUSH
  0x0201, // SETBKCOLOR
  0x0209, // SETTEXTCOLOR
  0x020a, // SETTEXTJUSTIFICATION
  0x020b, // SETWINDOWORG
  0x020c, // SETWINDOWEXT
  0x020d, // SETVIEWPORTORG
  0x020e, // SETVIEWPORTEXT
  0x020f, // OFFSETWINDOWORG
  0x0211, // OFFSETVIEWPORTORG
  0x0214, // MOVETO
  0x0220, // OFFSETCLIPRGN
  0x0231, // SETMAPPERFLAGS
  0x0234, // SELECTPALETTE
  0x02fa, // CREATEPENINDIRECT
  0x02fb, // CREATEFONTINDIRECT
  0x02fc, // CREATEBRUSHINDIRECT
  0x0410, // SCALEWINDOWEXT
  0x0412, // SCALEVIEWPORTEXT
  0x0415, // EXCLUDECLIPRECT
  0x0416, // INTERSECTCLIPRECT
  0x0626, // ESCAPE
  0x06ff, // CREATEREGION
  0x012c, // SELECTCLIPREGION
  0x0037, // SETPALENTRIES
  0x0038, // ANIMATEPALETTE
]);

/** The bitmap inside a bitmap-only WMF as a `data:image/bmp;base64,…` URL,
 *  or null (see {@link emfBitmapDataUrl} — same contract). */
export function wmfBitmapDataUrl(bytes: Uint8Array): string | null {
  const dib = extractSoleWmfDib(bytes);
  return dib ? `data:image/bmp;base64,${toBase64(dib)}` : null;
}

/** The DIB of a bitmap-only WMF, framed as a complete BMP file. */
export function extractSoleWmfDib(bytes: Uint8Array): Uint8Array | null {
  if (bytes.length < 18) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let off = view.getUint32(0, true) === WMF_PLACEABLE ? 22 : 0;
  // META_HEADER: Type (1 memory / 2 disk), HeaderSize in words (9).
  const type = view.getUint16(off, true);
  const headerWords = view.getUint16(off + 2, true);
  if ((type !== 1 && type !== 2) || headerWords !== 9) return null;
  off += 18;
  let bitmap: { start: number; size: number; fn: number } | null = null;
  while (off + 6 <= bytes.length) {
    const size = view.getUint32(off, true) * 2;
    const fn = view.getUint16(off + 4, true);
    if (size < 6 || off + size > bytes.length) return null;
    if (WMF_BITMAP_RECORDS[fn] !== undefined) {
      // A META_DIBBITBLT with no DIB is the pattern-fill form: nothing to lift.
      if (size <= WMF_BITMAP_RECORDS[fn]) return null;
      if (bitmap) return null;
      bitmap = { start: off, size, fn };
    } else if (!WMF_STATE_RECORDS.has(fn)) {
      return null;
    }
    if (fn === 0) break;
    off += size;
  }
  if (!bitmap) return null;
  const usageOff = WMF_USAGE_OFFSET[bitmap.fn];
  if (
    usageOff !== undefined &&
    view.getUint16(bitmap.start + usageOff, true) !== 0
  )
    return null;
  const dibStart = bitmap.start + WMF_BITMAP_RECORDS[bitmap.fn];
  const dib = bytes.subarray(dibStart, bitmap.start + bitmap.size);
  const bmiSize = dibInfoSize(dib);
  if (bmiSize === null || bmiSize >= dib.length) return null;
  const out = new Uint8Array(14 + dib.length);
  const ov = new DataView(out.buffer);
  out[0] = 0x42;
  out[1] = 0x4d;
  ov.setUint32(2, out.length, true);
  ov.setUint32(10, 14 + bmiSize, true);
  out.set(dib, 14);
  return out;
}

/** Bytes from a DIB's start to its pixel array: BITMAPINFOHEADER (or a V4/V5
 *  one), the BI_BITFIELDS masks a 40-byte header keeps outside itself, and
 *  the colour table. Null for the 12-byte BITMAPCOREHEADER (RGBTRIPLE
 *  palettes; not seen in Word output) or a truncated header. */
function dibInfoSize(dib: Uint8Array): number | null {
  if (dib.length < 40) return null;
  const v = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const biSize = v.getUint32(0, true);
  if (biSize < 40) return null;
  const bpp = v.getUint16(14, true);
  const compression = v.getUint32(16, true);
  const clrUsed = v.getUint32(32, true);
  let size = biSize;
  if (biSize === 40 && compression === 3) size += 12; // BI_BITFIELDS masks
  if (bpp <= 8) size += (clrUsed || 1 << bpp) * 4;
  else if (clrUsed) size += clrUsed * 4; // optional palette on true colour
  return size;
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
