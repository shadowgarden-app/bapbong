import type { CellPadding, TableBorders } from './contracts.js';

/**
 * Table-style theming, resolved: the serializable sheet a document carries in
 * `doc.attrs['tableStyles']`, plus the PURE geometry that decides which
 * conditional branches reach which cell.
 *
 * This module is deliberately OOXML-free — the docx importer builds the sheet
 * (parsing w:tblStylePr / w:tblLook / theme colours into the plain values
 * here), and the layout engine applies it per cell at flatten time. Both
 * sides import THIS copy, so the gating rules cannot drift apart. The rules
 * themselves are Word's, verified cell-by-cell against Word's own PDF on
 * probe F2 (four tblLook combinations over one Light-Grid-shaped style,
 * including the lastRow/lastCol/vBand branches no corpus file exercises).
 */

/** Which conditional formats a table lets through — the six `w:tblLook`
 *  gates. All-false is a legal state; Word's default (element absent) is
 *  firstRow + firstCol + hBand, bitmask 0x04A0. */
export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  hBand: boolean;
  vBand: boolean;
}

/** Where a cell sits in its table — everything the conditional formats ask. */
export interface TableCellPos {
  row: number;
  rowCount: number;
  /** Leading grid column, and how many columns the cell spans. */
  col: number;
  colspan: number;
  colCount: number;
  /** The row carries `w:tblHeader` (a repeated heading row). */
  header: boolean;
}

/** `w:tblStyleRowBandSize` / `w:tblStyleColBandSize` — how many rows/columns
 *  make one band. **Word's default is 0, not the standard's 1**, and a size
 *  of 0 disables that axis's banding entirely (MS-OI29500 §2.1.251). */
export interface TableBandSizes {
  row: number;
  col: number;
}

/** The `w:tblStylePr` region types, in the `@w:type` spelling. */
export type TableCondType =
  | 'firstRow'
  | 'lastRow'
  | 'firstCol'
  | 'lastCol'
  | 'band1Horz'
  | 'band2Horz'
  | 'band1Vert'
  | 'band2Vert'
  | 'nwCell'
  | 'neCell'
  | 'swCell'
  | 'seCell';

/** The run-font delta a style layer contributes (a field set here overrides
 *  the layer below; an absent field inherits). */
export interface TableStyleFont {
  family?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

/** One resolved `w:tblStylePr` branch (or the style's own w:tcPr defaults):
 *  what it contributes to the cells of its region. */
export interface TableStyleCondLayer {
  /** w:rPr, rolled up through basedOn. */
  font?: TableStyleFont;
  /** w:shd fill — hex colour, or null for an explicit "no fill". Absent =
   *  the branch says nothing about shading. */
  background?: string | null;
  /** w:tcBorders with REGION semantics: top/bottom/left/right are the
   *  region's outer edges, insideH/insideV the edges between its cells —
   *  {@link condCellBorders} picks per cell. */
  borders?: TableBorders;
  /** w:tcPr/w:vAlign. */
  vAlign?: 'center' | 'bottom';
  /** w:tcPr/w:tcMar, px per side. */
  padding?: CellPadding;
}

/** One table style, fully resolved (basedOn chain rolled up, theme colours
 *  already hex). Everything a renderer needs; nothing OOXML. */
export interface ResolvedTableStyle {
  /** Whole-table contributions from the style's w:tblPr. */
  table: {
    borders?: TableBorders;
    cellPadding?: CellPadding;
    align?: 'center' | 'right';
  };
  /** The style's own w:rPr — run defaults for every cell. */
  font?: TableStyleFont;
  /** The style's own w:tcPr — cell defaults for every cell (borders here are
   *  ignored, matching Word: only conditional branches contribute borders
   *  below the table level). */
  cell?: TableStyleCondLayer;
  /** Conditional branches by region type. */
  cond: Partial<Record<TableCondType, TableStyleCondLayer>>;
  bands: TableBandSizes;
}

/** Every table style a document's tables reference, by w:styleId. */
export type TableStyleSheet = Record<string, ResolvedTableStyle>;

/**
 * The `w:tblStylePr` types that apply to one cell, base-most FIRST so a plain
 * left-to-right merge reproduces Word's precedence:
 *
 *   *"When specified, Office applies conditional formats in the following
 *    order (therefore subsequent formats override properties on previous
 *    formats): Odd row banding, even row banding · Odd column banding, even
 *    column banding · First column, last column · First row, last row · Top
 *    left, top right, bottom left, bottom right"* — [MS-OI29500] §2.1.250.
 *
 * ISO 29500 §17.7.6.6 orders it differently (columns before rows, first row
 * before first column); we follow Word.
 *
 * Rules that are not in the standard's prose, all verified against Word's own
 * rendering (probe F2 PDF, cell-by-cell on four tblLook combinations):
 *   - A row carrying `w:tblHeader` also takes firstRow formatting (Eric White,
 *     "Assembling Paragraph and Run Properties for Cells").
 *   - Banding counts the BODY only: the first row drops out when firstRow
 *     formatting is on, the last when lastRow is. Same expression for columns.
 *   - Column banding still applies INSIDE the first/last row (probe F2 table
 *     B: r1c2/r1c4 keep the band fill under firstRow's bold), and row banding
 *     inside the first/last column — the row/col gates only exclude their OWN
 *     axis's banding.
 *   - Corner cells need BOTH gates: *"Top left cell – when Header Row and
 *     First Column are used"* ([MS-OI29500] §17.4.55(b)).
 */
export function condTypesFor(
  pos: TableCellPos,
  look: TableLook,
  bands: TableBandSizes,
): TableCondType[] {
  const firstRow = look.firstRow && (pos.row === 0 || pos.header);
  const lastRow = look.lastRow && pos.row === pos.rowCount - 1;
  const firstCol = look.firstCol && pos.col === 0;
  const lastCol = look.lastCol && pos.col + pos.colspan >= pos.colCount;
  const types: TableCondType[] = [];
  if (look.hBand && bands.row > 0 && !firstRow && !lastRow) {
    const body = pos.row - (look.firstRow ? 1 : 0);
    types.push(
      Math.floor(body / bands.row) % 2 === 0 ? 'band1Horz' : 'band2Horz',
    );
  }
  if (look.vBand && bands.col > 0 && !firstCol && !lastCol) {
    const body = pos.col - (look.firstCol ? 1 : 0);
    types.push(
      Math.floor(body / bands.col) % 2 === 0 ? 'band1Vert' : 'band2Vert',
    );
  }
  if (firstCol) types.push('firstCol');
  if (lastCol) types.push('lastCol');
  if (firstRow) types.push('firstRow');
  if (lastRow) types.push('lastRow');
  if (firstRow && firstCol) types.push('nwCell');
  if (firstRow && lastCol) types.push('neCell');
  if (lastRow && firstCol) types.push('swCell');
  if (lastRow && lastCol) types.push('seCell');
  return types;
}

/**
 * The rectangle of grid cells a branch styles. It decides what `insideH` and
 * `insideV` mean inside that branch: they are the edges BETWEEN cells of the
 * region, so a cell on the region's boundary takes top/bottom/left/right there
 * and an interior one takes the inside pair.
 */
export interface TableBranchRegion {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

export function branchRegion(
  type: TableCondType,
  pos: TableCellPos,
  look: TableLook,
  bands: TableBandSizes,
): TableBranchRegion {
  const lastRow = pos.rowCount - 1;
  const lastCol = pos.colCount - 1;
  // Banding counts the body only, so a band's bounds are measured from there.
  const bodyTop = look.firstRow ? 1 : 0;
  const bodyBottom = lastRow - (look.lastRow ? 1 : 0);
  const bodyLeft = look.firstCol ? 1 : 0;
  const bodyRight = lastCol - (look.lastCol ? 1 : 0);
  const wholeRows = { rowStart: 0, rowEnd: lastRow };
  const wholeCols = { colStart: 0, colEnd: lastCol };
  const band = (v: number, from: number, to: number, size: number) => {
    const start = from + Math.floor((v - from) / size) * size;
    return { start, end: Math.min(start + size - 1, to) };
  };
  switch (type) {
    case 'firstRow':
      return { rowStart: 0, rowEnd: 0, ...wholeCols };
    case 'lastRow':
      return { rowStart: lastRow, rowEnd: lastRow, ...wholeCols };
    case 'firstCol':
      return { ...wholeRows, colStart: 0, colEnd: 0 };
    case 'lastCol':
      return { ...wholeRows, colStart: lastCol, colEnd: lastCol };
    case 'band1Horz':
    case 'band2Horz': {
      const b = band(pos.row, bodyTop, bodyBottom, Math.max(1, bands.row));
      return { rowStart: b.start, rowEnd: b.end, ...wholeCols };
    }
    case 'band1Vert':
    case 'band2Vert': {
      const b = band(pos.col, bodyLeft, bodyRight, Math.max(1, bands.col));
      return { ...wholeRows, colStart: b.start, colEnd: b.end };
    }
    // The four corner branches style exactly one cell.
    default:
      return {
        rowStart: pos.row,
        rowEnd: pos.row,
        colStart: pos.col,
        colEnd: pos.col + pos.colspan - 1,
      };
  }
}

/** Everything a table style contributes to ONE cell, after gating and the
 *  precedence merge — the layer a renderer slots UNDER the cell's own
 *  attributes and the document's direct formatting. */
export interface TableCellStyleLayer {
  font?: TableStyleFont;
  background?: string | null;
  borders?: TableBorders | null;
  vAlign?: 'center' | 'bottom';
  padding?: CellPadding | null;
}

/**
 * The cell borders the conditional branches contribute, later branch winning
 * per side. For each branch, the cell's position within the branch's REGION
 * decides which declared side applies: a cell on the region's top edge takes
 * `top`, an interior one takes `insideH` — and an undefined pick keeps what
 * an earlier branch said.
 */
function condCellBorders(
  branches: { type: TableCondType; layer: TableStyleCondLayer }[],
  pos: TableCellPos,
  look: TableLook,
  bands: TableBandSizes,
): TableBorders | null {
  const out: TableBorders = {};
  for (const { type, layer } of branches) {
    const set = layer.borders;
    if (!set) continue;
    const region = branchRegion(type, pos, look, bands);
    const at = {
      top: pos.row === region.rowStart,
      bottom: pos.row === region.rowEnd,
      left: pos.col === region.colStart,
      right: pos.col + pos.colspan - 1 >= region.colEnd,
    };
    const put = (
      side: 'top' | 'bottom' | 'left' | 'right',
      from: TableBorders[keyof TableBorders],
    ) => {
      if (from !== undefined) out[side] = from;
    };
    put('top', at.top ? set.top : set.insideH);
    put('bottom', at.bottom ? set.bottom : set.insideH);
    put('left', at.left ? set.left : set.insideV);
    put('right', at.right ? set.right : set.insideV);
  }
  return Object.keys(out).length > 0 ? out : null;
}

/** Merge one font delta over another (later wins per field). */
export function mergeTableStyleFont(
  base: TableStyleFont | undefined,
  over: TableStyleFont | undefined,
): TableStyleFont | undefined {
  if (!base) return over;
  if (!over) return base;
  return { ...base, ...over };
}

/**
 * Resolve the full style layer for one cell: the style's own cell defaults,
 * then every conditional branch the table's tblLook lets through, in Word's
 * application order. This is THE function both sides share — the layout
 * engine calls it to render, and the importer's shadow check calls it to
 * prove it reproduces the old baked path.
 */
export function cellStyleLayer(
  style: ResolvedTableStyle,
  look: TableLook,
  pos: TableCellPos,
): TableCellStyleLayer {
  const branches = condTypesFor(pos, look, bands0(style)).flatMap((type) => {
    const layer = style.cond[type];
    return layer ? [{ type, layer }] : [];
  });
  const out: TableCellStyleLayer = {};
  let font = mergeTableStyleFont(style.font, undefined);
  const layers = style.cell ? [style.cell] : [];
  for (const b of branches) layers.push(b.layer);
  for (const layer of layers) {
    font = mergeTableStyleFont(font, layer.font);
    if (layer.background !== undefined) out.background = layer.background;
    if (layer.vAlign !== undefined) out.vAlign = layer.vAlign;
    if (layer.padding)
      out.padding = { ...(out.padding ?? {}), ...layer.padding };
  }
  if (font && Object.keys(font).length > 0) out.font = font;
  const borders = condCellBorders(branches, pos, look, bands0(style));
  if (borders) out.borders = borders;
  return out;
}

function bands0(style: ResolvedTableStyle): TableBandSizes {
  return style.bands ?? { row: 0, col: 0 };
}
