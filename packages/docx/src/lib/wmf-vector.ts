/**
 * Vector WMF → display list, for the one family of vector metafiles .docx
 * files carry in bulk: MathType / Equation Editor previews.
 *
 * Every `Equation.DSMT4` object in a document embeds a WMF preview
 * (`v:imagedata` → `word/media/*.wmf`) that Word itself paints when MathType
 * isn't installed — so drawing this WMF IS matching Word, not approximating
 * it. Measured over real exam papers, those previews use a tiny closed set of
 * records: glyphs placed one-by-one through EXTTEXTOUT (with per-char
 * advances), fraction bars and radical strokes as MOVETO/LINETO, arrowheads
 * (vector marks over symbols) as one filled POLYGON, and font/pen/colour
 * state — no paths, no bitmaps.
 *
 * This module only PARSES: it walks the records and returns a display list
 * (text runs and lines in z-order, logical coordinates) for a renderer to
 * replay. Files using any drawing primitive outside that set return null and
 * the caller keeps its existing behaviour (bitmap extraction, then
 * placeholder) — same contract as `extractSoleWmfDib`.
 *
 * Record layouts: [MS-WMF] 2.3 (records) and 2.2.1.2 (placeable header).
 * Offsets are from the start of a record's parameter bytes (after the 6-byte
 * Size+Function prefix) unless said otherwise.
 */
import type {
  VectorImageSpec,
  VectorOp,
} from '@shadow-garden/bapbong-contracts';
import { decodeWmfText } from './wmf-charmap.js';
import { mtefLinearFromWmf } from './mtef.js';

/** One glyph run: EXTTEXTOUT plus the graphics state it was issued under. */
export interface WmfTextOp {
  readonly kind: 'text';
  /** Reference point, logical units (window origin already subtracted). */
  readonly x: number;
  readonly y: number;
  /** The string's raw single-byte codes. Fonts like Symbol and MT Extra use
   *  font-specific encodings, so mapping bytes to Unicode needs the facename
   *  — that stays the renderer's job, the parser just keeps the bytes. */
  readonly bytes: Uint8Array;
  /** Per-character advances in logical units (the Dx array), or null when
   *  the record carries none. MathType kerns by hand — a renderer must lay
   *  glyphs by these advances, never by its own measurement. */
  readonly dx: readonly number[] | null;
  /** From the selected font's LogFont. Height keeps its sign: negative is em
   *  size, positive cell height ([MS-WMF] 2.2.1.2). */
  readonly fontHeight: number;
  readonly italic: boolean;
  /** LogFont lfWeight (400 normal, 700 bold). */
  readonly weight: number;
  readonly underline: boolean;
  readonly face: string;
  /** LogFont lfCharSet (0 ANSI, 2 SYMBOL, …). */
  readonly charset: number;
  /** Current text colour as 24-bit 0xRRGGBB. */
  readonly color: number;
  /** Raw SETTEXTALIGN value (TA_LEFT|TA_TOP = 0, TA_BASELINE = 24, …). */
  readonly align: number;
}

/** One stroke: LINETO from the current position under the selected pen. */
export interface WmfLineOp {
  readonly kind: 'line';
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
  /** Pen width in logical units (0 = hairline). */
  readonly width: number;
  /** Pen colour as 24-bit 0xRRGGBB. */
  readonly color: number;
}

/** One POLYGON: outlined by the pen, filled by the brush (arrowheads). */
export interface WmfPolygonOp {
  readonly kind: 'polygon';
  readonly points: readonly { readonly x: number; readonly y: number }[];
  /** Fill colour as 0xRRGGBB, or null for a hollow (BS_NULL) brush. */
  readonly fill: number | null;
  readonly strokeWidth: number;
  readonly strokeColor: number;
}

export type WmfOp = WmfTextOp | WmfLineOp | WmfPolygonOp;

/** A parsed vector WMF: ops in z-order over a logical-unit canvas. */
export interface WmfVectorImage {
  /** SETWINDOWEXT — the logical size ops are placed in ([0,width]×[0,height]
   *  after the origin subtraction the parser already did). */
  readonly width: number;
  readonly height: number;
  /** Logical units per inch from the placeable header, 0 when the file has
   *  none (then the physical size must come from the surrounding markup —
   *  `w:dxaOrig`/`w:dyaOrig` on the `w:object`). */
  readonly unitsPerInch: number;
  readonly ops: readonly WmfOp[];
}

const WMF_PLACEABLE = 0x9ac6cdd7;

const FN = {
  EOF: 0x0000,
  SAVEDC: 0x001e,
  RESTOREDC: 0x0127,
  SETBKMODE: 0x0102,
  SETMAPMODE: 0x0103,
  SETROP2: 0x0104,
  SETRELABS: 0x0105,
  SETPOLYFILLMODE: 0x0106,
  SETSTRETCHBLTMODE: 0x0107,
  SETTEXTCHAREXTRA: 0x0108,
  SETTEXTCOLOR: 0x0209,
  SETBKCOLOR: 0x0201,
  SETTEXTALIGN: 0x012e,
  SETTEXTJUSTIFICATION: 0x020a,
  SETWINDOWORG: 0x020b,
  SETWINDOWEXT: 0x020c,
  SELECTOBJECT: 0x012d,
  DELETEOBJECT: 0x01f0,
  CREATEFONTINDIRECT: 0x02fb,
  CREATEPENINDIRECT: 0x02fa,
  CREATEBRUSHINDIRECT: 0x02fc,
  MOVETO: 0x0214,
  LINETO: 0x0213,
  POLYGON: 0x0324,
  EXTTEXTOUT: 0x0a32,
  TEXTOUT: 0x0521,
  ESCAPE: 0x0626,
} as const;

/** State-only functions a vector preview may carry that change nothing this
 *  parser tracks — tolerated as no-ops. Anything not here, not in FN, is a
 *  drawing primitive (or a blit) outside the supported set: parse fails. */
const NOOP_FNS = new Set<number>([
  FN.SAVEDC,
  FN.RESTOREDC,
  FN.SETMAPMODE,
  FN.SETROP2,
  FN.SETRELABS,
  FN.SETPOLYFILLMODE,
  FN.SETSTRETCHBLTMODE,
  FN.SETTEXTCHAREXTRA,
  FN.SETTEXTJUSTIFICATION,
  FN.ESCAPE,
]);

interface Font {
  readonly kind: 'font';
  readonly height: number;
  readonly italic: boolean;
  readonly weight: number;
  readonly underline: boolean;
  readonly face: string;
  readonly charset: number;
}

interface Pen {
  readonly kind: 'pen';
  readonly width: number;
  readonly color: number;
}

interface Brush {
  readonly kind: 'brush';
  /** Fill colour, or null for BS_NULL (hollow). */
  readonly color: number | null;
}

type WmfObject = Font | Pen | Brush;

/** COLORREF (0x00BBGGRR little-endian) → 0xRRGGBB. */
function colorref(v: number): number {
  return ((v & 0xff) << 16) | (v & 0xff00) | ((v >> 16) & 0xff);
}

const DEFAULT_FONT: Font = {
  kind: 'font',
  height: -12,
  italic: false,
  weight: 400,
  underline: false,
  face: 'System',
  charset: 0,
};
const DEFAULT_PEN: Pen = { kind: 'pen', width: 0, color: 0 };
const DEFAULT_BRUSH: Brush = { kind: 'brush', color: 0xffffff };

/**
 * Parse a vector WMF into a display list, or null when the file is anything
 * beyond the supported record set (then it's not a MathType-style preview and
 * other handling applies).
 */
export function wmfVectorImage(bytes: Uint8Array): WmfVectorImage | null {
  if (bytes.length < 18) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let off = 0;
  let unitsPerInch = 0;
  if (view.getUint32(0, true) === WMF_PLACEABLE) {
    unitsPerInch = view.getUint16(14, true);
    off = 22;
  }
  // META_HEADER: Type (1 memory / 2 disk), HeaderSize in words (9).
  const type = view.getUint16(off, true);
  if ((type !== 1 && type !== 2) || view.getUint16(off + 2, true) !== 9)
    return null;
  off += 18;

  // Graphics state. The window transform in these files is origin+extent
  // only (SETMAPMODE arrives before both), so coordinates normalize by
  // subtracting the origin as they are read.
  let orgX = 0;
  let orgY = 0;
  let extX = 0;
  let extY = 0;
  let posX = 0; // current position (MOVETO/LINETO)
  let posY = 0;
  let textColor = 0;
  let textAlign = 0;
  // The object table: CreateXxx fills the LOWEST free slot, DeleteObject
  // frees one ([MS-WMF] 3.1.4.1) — MathType churns through fonts, so getting
  // this wrong selects the wrong glyphs silently.
  const objects: (WmfObject | null)[] = [];
  let font: Font = DEFAULT_FONT;
  let pen: Pen = DEFAULT_PEN;
  let brush: Brush = DEFAULT_BRUSH;
  const ops: WmfOp[] = [];

  const create = (obj: WmfObject): void => {
    const free = objects.indexOf(null);
    if (free >= 0) objects[free] = obj;
    else objects.push(obj);
  };

  let sawEof = false;
  while (off + 6 <= bytes.length) {
    const size = view.getUint32(off, true) * 2;
    const fn = view.getUint16(off + 4, true);
    if (size < 6 || off + size > bytes.length) return null;
    const p = off + 6; // parameter bytes
    const paramLen = size - 6;

    switch (fn) {
      case FN.EOF:
        sawEof = true;
        break;
      case FN.SETWINDOWORG:
        // Params are Y then X (each int16) — most point-valued WMF records
        // store the pair reversed.
        orgY = view.getInt16(p, true);
        orgX = view.getInt16(p + 2, true);
        break;
      case FN.SETWINDOWEXT:
        extY = view.getInt16(p, true);
        extX = view.getInt16(p + 2, true);
        break;
      case FN.SETTEXTCOLOR:
        textColor = colorref(view.getUint32(p, true));
        break;
      case FN.SETBKCOLOR:
      case FN.SETBKMODE:
        // Text backgrounds: MathType always sets TRANSPARENT, and opaque
        // boxes behind glyphs would be wrong over any page colour — ignored.
        break;
      case FN.SETTEXTALIGN:
        textAlign = view.getUint16(p, true);
        break;
      case FN.CREATEFONTINDIRECT: {
        // LogFont: Height(i16) Width Esc Orient Weight(i16) Italic(u8)
        // Underline StrikeOut CharSet OutPrec ClipPrec Quality PitchFam,
        // then the facename (≤32 bytes, NUL-terminated).
        if (paramLen < 18) return null;
        let end = p + 18;
        const nameEnd = Math.min(p + 18 + 32, off + size);
        while (end < nameEnd && bytes[end] !== 0) end++;
        let face = '';
        for (let i = p + 18; i < end; i++)
          face += String.fromCharCode(bytes[i]);
        create({
          kind: 'font',
          height: view.getInt16(p, true),
          weight: view.getInt16(p + 8, true),
          italic: bytes[p + 10] !== 0,
          underline: bytes[p + 11] !== 0,
          charset: bytes[p + 13],
          face,
        });
        break;
      }
      case FN.CREATEPENINDIRECT:
        // Style(2) Width.x(2) Width.y(2) COLORREF(4); only solid strokes
        // appear in these files, style is ignored.
        if (paramLen < 10) return null;
        create({
          kind: 'pen',
          width: view.getInt16(p + 2, true),
          color: colorref(view.getUint32(p + 6, true)),
        });
        break;
      case FN.CREATEBRUSHINDIRECT:
        // LogBrush: Style(u16) COLORREF(u32) Hatch(u16). BS_NULL (1) fills
        // nothing; every other style is treated as its colour, solid.
        if (paramLen < 8) return null;
        create({
          kind: 'brush',
          color:
            view.getUint16(p, true) === 1
              ? null
              : colorref(view.getUint32(p + 2, true)),
        });
        break;
      case FN.SELECTOBJECT: {
        const obj = objects[view.getUint16(p, true)];
        if (!obj) return null; // selecting a freed slot: corrupt file
        if (obj.kind === 'font') font = obj;
        else if (obj.kind === 'pen') pen = obj;
        else brush = obj;
        break;
      }
      case FN.DELETEOBJECT: {
        const idx = view.getUint16(p, true);
        if (idx >= objects.length || objects[idx] === null) return null;
        objects[idx] = null;
        break;
      }
      case FN.MOVETO:
        posY = view.getInt16(p, true) - orgY;
        posX = view.getInt16(p + 2, true) - orgX;
        break;
      case FN.LINETO: {
        const y = view.getInt16(p, true) - orgY;
        const x = view.getInt16(p + 2, true) - orgX;
        ops.push({
          kind: 'line',
          x1: posX,
          y1: posY,
          x2: x,
          y2: y,
          width: pen.width,
          color: pen.color,
        });
        posX = x;
        posY = y;
        break;
      }
      case FN.POLYGON: {
        // NumberOfPoints(i16), then POINTs — x before y, unlike the
        // reversed pairs of the position/window records.
        if (paramLen < 2) return null;
        const n = view.getInt16(p, true);
        if (n < 2 || 2 + n * 4 > paramLen) return null;
        const points = Array.from({ length: n }, (_, i) => ({
          x: view.getInt16(p + 2 + i * 4, true) - orgX,
          y: view.getInt16(p + 4 + i * 4, true) - orgY,
        }));
        ops.push({
          kind: 'polygon',
          points,
          fill: brush.color,
          strokeWidth: pen.width,
          strokeColor: pen.color,
        });
        break;
      }
      case FN.EXTTEXTOUT: {
        // Y(i16) X(i16) StringLength(i16) fwOpts(u16) [Rectangle(8) when
        // opaque/clipped] String (padded to even) [Dx: StringLength × i16].
        if (paramLen < 8) return null;
        let y = view.getInt16(p, true) - orgY;
        let x = view.getInt16(p + 2, true) - orgX;
        const len = view.getInt16(p + 4, true);
        const fwOpts = view.getUint16(p + 6, true);
        // ETO_OPAQUE (0x0002) | ETO_CLIPPED (0x0004) insert the rectangle.
        const strAt = p + 8 + (fwOpts & 0x0006 ? 8 : 0);
        if (len < 0 || strAt + len > off + size) return null;
        const text = bytes.slice(strAt, strAt + len);
        const dxAt = strAt + len + (len & 1);
        const dx =
          dxAt + len * 2 <= off + size
            ? Array.from({ length: len }, (_, i) =>
                view.getInt16(dxAt + i * 2, true),
              )
            : null;
        // TA_UPDATECP (bit 0): the record's X/Y are IGNORED — text draws at
        // the current position, which then advances by the string's extent
        // (the Dx sum). MathType always draws this way: a MOVETO seeds the
        // position and consecutive runs (base, then Symbol, then MT Extra
        // marks) chain off each other.
        if (textAlign & 0x0001) {
          x = posX;
          y = posY;
          if (dx) posX += dx.reduce((s, d) => s + d, 0);
        }
        ops.push({
          kind: 'text',
          x,
          y,
          bytes: text,
          dx,
          fontHeight: font.height,
          italic: font.italic,
          weight: font.weight,
          underline: font.underline,
          face: font.face,
          charset: font.charset,
          color: textColor,
          align: textAlign,
        });
        break;
      }
      default:
        if (!NOOP_FNS.has(fn)) return null;
    }
    if (sawEof) break;
    off += size;
  }

  if (!sawEof || extX <= 0 || extY <= 0) return null;
  return { width: extX, height: extY, unitsPerInch, ops };
}

/** The registry family a WMF facename paints with. Symbol and MT Extra runs
 *  are already decoded to Unicode, so they render from the same serif face
 *  as the surrounding formula text. */
function familyOf(face: string): string {
  return face === 'Symbol' || face === 'MT Extra' || face === 'System'
    ? 'Times New Roman'
    : face;
}

const cssColor = (c: number): string => `#${c.toString(16).padStart(6, '0')}`;

export interface WmfVectorResult {
  spec: VectorImageSpec;
  /** Linear equation text recovered from the embedded MTEF, or null — what
   *  "Convert to editable equation" inserts. */
  linear: string | null;
  /** Physical size in CSS px from the placeable header, when it has one —
   *  the size fallback for markup that states none. */
  pxWidth: number | null;
  pxHeight: number | null;
}

/**
 * Parse a vector WMF and resolve it to the renderer-facing contract: bytes
 * decoded to Unicode per facename, colours to CSS, LogFont height to an em
 * size, TA flags to a vertical alignment. Null when {@link wmfVectorImage}
 * rejects the file.
 */
export function wmfVectorSpec(bytes: Uint8Array): WmfVectorResult | null {
  const image = wmfVectorImage(bytes);
  if (!image) return null;
  const ops: VectorOp[] = image.ops.map((op) => {
    if (op.kind === 'line')
      return {
        kind: 'line',
        x1: op.x1,
        y1: op.y1,
        x2: op.x2,
        y2: op.y2,
        width: op.width,
        color: cssColor(op.color),
      };
    if (op.kind === 'polygon')
      return {
        kind: 'polygon',
        points: op.points.map((pt) => ({ x: pt.x, y: pt.y })),
        ...(op.fill !== null ? { fill: cssColor(op.fill) } : {}),
        stroke: cssColor(op.strokeColor),
        strokeWidth: op.strokeWidth,
      };
    // TA_BASELINE is both bits (24); TA_BOTTOM is 8 alone; else top. The
    // contract's default is baseline, so only the others are stated.
    const vertical = op.align & 24;
    return {
      kind: 'text',
      x: op.x,
      y: op.y,
      text: decodeWmfText(op.bytes, op.face),
      ...(op.dx ? { dx: [...op.dx] } : {}),
      // Negative LogFont height is the em size; positive is the cell height
      // (em + internal leading), taken as-is — slightly large, and MathType
      // always writes the negative form.
      size: Math.abs(op.fontHeight),
      family: familyOf(op.face),
      ...(op.weight >= 600 ? { bold: true } : {}),
      ...(op.italic ? { italic: true } : {}),
      ...(op.underline ? { underline: true } : {}),
      color: cssColor(op.color),
      ...(vertical !== 24
        ? { vAlign: vertical === 8 ? ('bottom' as const) : ('top' as const) }
        : {}),
    };
  });
  const px = (units: number): number | null =>
    image.unitsPerInch > 0
      ? Math.round((units / image.unitsPerInch) * 96)
      : null;
  return {
    spec: { width: image.width, height: image.height, ops },
    linear: mtefLinearFromWmf(bytes),
    pxWidth: px(image.width),
    pxHeight: px(image.height),
  };
}
