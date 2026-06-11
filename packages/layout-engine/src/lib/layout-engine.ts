import { Mark, Node as PMNode } from 'prosemirror-model';
import type {
  Align,
  CellPadding,
  FlowBlock,
  FlowFloat,
  FlowInline,
  FlowParagraph,
  FlowTable,
  FlowTableCell,
  FlowTableRow,
  FontSpec,
  InlineField,
  InlineImage,
  InlineRun,
  LayoutConfig,
  LayoutImageSegment,
  LayoutLine,
  LayoutSegment,
  MeasureMetrics,
  MeasureText,
  ParagraphIndent,
  ResolvedCell,
  ResolvedChrome,
  ResolvedFloat,
  ResolvedLayout,
  ResolvedPage,
  ResolvedTable,
} from '@shadow-garden/bapbong-contracts';

const DEFAULT_FONT: FontSpec = { family: 'Arial', sizePt: 11, bold: false, italic: false };
const PT_TO_PX = 96 / 72;
const LINE_HEIGHT_FACTOR = 1.2;
const BASELINE_FACTOR = 0.8;
const DEFAULT_TAB_WIDTH = 48; // 0.5in, Word's default tab interval
// Default cell padding (px) — Word's w:tblCellMar defaults: 108 twips
// (= 7.2px) left/right, 0 top/bottom. Tables can override via cellPadding.
const CELL_PAD_X = 7.2;
const CELL_PAD_Y = 0;

const sizePx = (font: FontSpec) => font.sizePt * PT_TO_PX;

/** Shared layout inputs, threaded through paragraph/table/cell layout. */
interface Ctx {
  base: FontSpec;
  measure: MeasureText;
  metrics?: MeasureMetrics;
  tabWidth: number;
  /** Width placeholder for page-number fields (digit count of the page
   *  total once known; '1' on the first pass). */
  fieldPlaceholder: string;
}

function findMark(marks: readonly Mark[], name: string): Mark | undefined {
  return marks.find((m) => m.type.name === name);
}

/** Resolve a text node's marks into an InlineRun (font + color + link).
 *  `pos` is the absolute PM position of the run's first character. */
function resolveRun(node: PMNode, base: FontSpec, pos: number): InlineRun {
  const marks = node.marks;
  const font: FontSpec = { ...base };
  if (findMark(marks, 'strong')) font.bold = true;
  if (findMark(marks, 'em')) font.italic = true;
  const size = findMark(marks, 'fontSize');
  if (size) font.sizePt = Number(size.attrs['size']) || base.sizePt;
  const family = findMark(marks, 'fontFamily');
  if (family) font.family = String(family.attrs['family'] ?? base.family);
  const color = findMark(marks, 'textColor');
  const link = findMark(marks, 'link');
  const run: InlineRun = {
    text: node.text ?? '',
    font,
    color: color ? String(color.attrs['color']) : undefined,
    link: link ? String(link.attrs['href']) : undefined,
    pos,
  };
  if (findMark(marks, 'underline')) run.underline = true;
  if (findMark(marks, 'strike')) run.strike = true;
  return run;
}

/** Resolve a page_field node into an InlineField (font/color from marks). */
function resolveField(node: PMNode, base: FontSpec, pos: number): InlineField {
  const marks = node.marks;
  const font: FontSpec = { ...base };
  if (findMark(marks, 'strong')) font.bold = true;
  if (findMark(marks, 'em')) font.italic = true;
  const size = findMark(marks, 'fontSize');
  if (size) font.sizePt = Number(size.attrs['size']) || base.sizePt;
  const family = findMark(marks, 'fontFamily');
  if (family) font.family = String(family.attrs['family'] ?? base.family);
  const color = findMark(marks, 'textColor');
  return {
    field: node.attrs['kind'] === 'pages' ? 'pageCount' : 'pageNumber',
    font,
    color: color ? String(color.attrs['color']) : undefined,
    pos,
  };
}

/** Resolve an image node into an InlineImage. Missing dimensions become 0
 *  (the image then takes no space until real sizing is available). */
function resolveImage(node: PMNode, pos: number): InlineImage {
  const a = node.attrs;
  const link = findMark(node.marks, 'link');
  return {
    src: String(a['src'] ?? ''),
    width: Number(a['width']) || 0,
    height: Number(a['height']) || 0,
    link: link ? String(link.attrs['href']) : undefined,
    pos,
  };
}

/** Flatten one paragraph node into a FlowParagraph (text + inline images).
 *  `nodePos` is the absolute PM position of the paragraph node itself. With
 *  `allowFloats`, anchored images become FlowFloats instead of inline content;
 *  contexts without float support (table cells, chrome) degrade them inline. */
function paragraphToFlow(
  node: PMNode,
  base: FontSpec,
  nodePos: number,
  allowFloats = false,
): FlowParagraph {
  const contentStart = nodePos + 1;
  const runs: FlowInline[] = [];
  const floats: FlowFloat[] = [];
  node.forEach((child, offset) => {
    if (child.isText) runs.push(resolveRun(child, base, contentStart + offset));
    else if (child.type.name === 'image') {
      const float = child.attrs['float'] as Omit<FlowFloat, 'src' | 'width' | 'height'> | null;
      if (float && allowFloats) {
        floats.push({
          ...float,
          src: String(child.attrs['src'] ?? ''),
          width: Number(child.attrs['width']) || 0,
          height: Number(child.attrs['height']) || 0,
        });
      } else {
        runs.push(resolveImage(child, contentStart + offset));
      }
    } else if (child.type.name === 'page_field')
      runs.push(resolveField(child, base, contentStart + offset));
  });
  const list = node.attrs['list'] as { marker?: string } | null;
  const align = node.attrs['align'] as Align | null | undefined;
  const indent = node.attrs['indent'] as ParagraphIndent | null | undefined;
  const flow: FlowParagraph = {
    type: 'paragraph',
    runs,
    marker: list?.marker || undefined,
    align: align ?? undefined,
    indent: indent ?? undefined,
    pos: contentStart,
    end: contentStart + node.content.size,
  };
  if (floats.length > 0) flow.floats = floats;
  return flow;
}

/** Whether the paragraph node anchors any floating image. */
function nodeHasFloats(node: PMNode): boolean {
  let found = false;
  node.forEach((child) => {
    if (child.type.name === 'image' && child.attrs['float']) found = true;
  });
  return found;
}

/** Flatten a block-level node (paragraph or table) into a FlowBlock, or null
 *  for node types we don't model yet. `nodePos` is its absolute PM position. */
function nodeToBlock(
  node: PMNode,
  base: FontSpec,
  nodePos: number,
  allowFloats = false,
): FlowBlock | null {
  if (node.type.name === 'paragraph') return paragraphToFlow(node, base, nodePos, allowFloats);
  if (node.type.name === 'table') return tableToFlow(node, base, nodePos);
  return null;
}

/** Flatten a table node (table → table_row → table_cell → block+). */
function tableToFlow(node: PMNode, base: FontSpec, nodePos: number): FlowTable {
  const rows: FlowTableRow[] = [];
  node.forEach((rowNode, rowOffset) => {
    const rowPos = nodePos + 1 + rowOffset;
    const cells: FlowTableCell[] = [];
    rowNode.forEach((cellNode, cellOffset) => {
      const cellPos = rowPos + 1 + cellOffset;
      const content: FlowBlock[] = [];
      cellNode.forEach((child, childOffset) => {
        const block = nodeToBlock(child, base, cellPos + 1 + childOffset);
        if (block) content.push(block);
      });
      const a = cellNode.attrs;
      cells.push({
        colspan: Number(a['colspan']) || 1,
        rowspan: Number(a['rowspan']) || 1,
        colwidth: (a['colwidth'] as number[] | null) ?? null,
        content,
      });
    });
    const row: FlowTableRow = { cells };
    if (rowNode.attrs['header'] === true) row.header = true;
    rows.push(row);
  });
  const flow: FlowTable = { type: 'table', rows };
  const cellPadding = node.attrs['cellPadding'] as CellPadding | null;
  if (cellPadding) flow.cellPadding = cellPadding;
  return flow;
}

/** Flatten a ProseMirror doc into FlowBlocks (paragraphs + tables). */
export function toFlowBlocks(
  doc: PMNode,
  defaultFont: Partial<FontSpec> = {},
  allowFloats = true,
): FlowBlock[] {
  const base: FontSpec = { ...DEFAULT_FONT, ...defaultFont };
  const blocks: FlowBlock[] = [];
  doc.forEach((node, offset) => {
    const block = nodeToBlock(node, base, offset, allowFloats);
    if (block) blocks.push(block);
  });
  return blocks;
}

interface Token {
  /** Text content (text token), or undefined for an image token. */
  text?: string;
  /** Image payload (image token), or undefined for a text token. */
  image?: InlineImage;
  /** Dynamic page-number field (atomic, like an image). */
  field?: 'pageNumber' | 'pageCount';
  font: FontSpec;
  color?: string;
  link?: string;
  underline?: boolean;
  strike?: boolean;
  width: number;
  isSpace: boolean;
  /** A tab character: its width is resolved to the next tab stop at layout. */
  isTab?: boolean;
  /** Absolute PM position of the token's first character / atom. */
  pos?: number;
  /** Size in PM positions (text length, or 1 for an image atom). */
  size: number;
}

/** Tokenize one inline item: words / spaces / tabs for text, a single atom for
 *  images. Tab widths are placeholders, resolved against tab stops at layout. */
function tokenizeInline(inline: FlowInline, ctx: Ctx): Token[] {
  if ('src' in inline) {
    return [
      {
        image: inline,
        font: ctx.base,
        link: inline.link,
        width: inline.width,
        isSpace: false,
        pos: inline.pos,
        size: 1,
      },
    ];
  }
  if ('field' in inline) {
    return [
      {
        field: inline.field,
        text: ctx.fieldPlaceholder,
        font: inline.font,
        color: inline.color,
        width: ctx.measure(ctx.fieldPlaceholder, inline.font),
        isSpace: false,
        pos: inline.pos,
        size: 1,
      },
    ];
  }
  let offset = 0;
  return inline.text
    .split(/(\t| +)/)
    .filter((part) => part.length > 0)
    .map((part) => {
      const isTab = part === '\t';
      const isSpace = isTab || /^ +$/.test(part);
      const pos = inline.pos != null ? inline.pos + offset : undefined;
      offset += part.length;
      return {
        text: part,
        font: inline.font,
        color: inline.color,
        link: inline.link,
        underline: inline.underline,
        strike: inline.strike,
        width: isTab ? 0 : ctx.measure(part, inline.font),
        isSpace,
        isTab,
        pos,
        size: part.length,
      };
    });
}

/** A laid-out line whose vertical position is not yet assigned (the page/cell
 *  placer fills in `y`). */
interface LineDraft {
  x: number;
  width: number;
  height: number;
  baseline: number;
  segments: LayoutSegment[];
  images: LayoutImageSegment[];
  from?: number;
  to?: number;
}

function draftToLine(d: LineDraft, y: number): LayoutLine {
  const line: LayoutLine = {
    x: d.x,
    y,
    width: d.width,
    height: d.height,
    baseline: d.baseline,
    segments: d.segments,
  };
  if (d.images.length > 0) line.images = d.images;
  if (d.from != null) line.from = d.from;
  if (d.to != null) line.to = d.to;
  return line;
}

/** The content bounds for the line about to be assembled. Queried once per
 *  line, so callers can flow text around floating images (the band narrows
 *  while a float's rectangle is in the way). */
type BandFn = (estHeight: number) => { left: number; right: number };

/** Wrap one paragraph, emitting one LineDraft per line. The band may differ
 *  per line; indents, the list marker and tab stops apply within each band. */
function wrapParagraph(
  block: FlowParagraph,
  ctx: Ctx,
  bandFn: BandFn,
  emit: (d: LineDraft) => void,
): void {
  const { base, measure, metrics, tabWidth } = ctx;
  const indent = block.indent;
  const indentLeft = indent?.left ?? 0;
  const indentRight = indent?.right ?? 0;
  // hanging outdents the first line; firstLine indents it. Mutually exclusive.
  const firstLineDelta = indent?.hanging != null ? -indent.hanging : indent?.firstLine ?? 0;
  const align: Align = block.align ?? 'left';

  const tokens = block.runs.flatMap((inline) => tokenizeInline(inline, ctx));

  // Baseline metrics for the default font seed every line (so empty lines have
  // a sensible height too).
  const baseMetrics = metrics ? metrics(base) : null;
  const nominalH = baseMetrics
    ? baseMetrics.ascent + baseMetrics.descent
    : sizePx(base) * LINE_HEIGHT_FACTOR;

  // Bounds for the line currently being assembled.
  let band = bandFn(nominalH);
  let lineLeft = band.left + indentLeft;
  let lineRight = band.right - indentRight;

  // List marker hangs at the first line's start; text follows after it, and
  // wrapped lines align under that text (hanging indent).
  let marker: LayoutSegment | null = null;
  let markerTextX = 0;
  if (block.marker) {
    marker = {
      x: lineLeft + firstLineDelta,
      text: block.marker,
      font: base,
      width: measure(block.marker, base),
    };
    markerTextX = marker.x + measure(`${block.marker} `, base);
  }
  const firstLineStart = marker ? markerTextX : lineLeft + firstLineDelta;
  let contStart = marker ? Math.max(markerTextX, lineLeft) : lineLeft;

  let lineTokens: Token[] = [];
  let lineWidth = 0; // running width of the current line's tokens
  let firstLine = true;
  let prevTo: number | undefined; // caret slot after the previous line's content
  let maxFontPx = sizePx(base); // tallest text (fallback line-height mode)
  let maxImagePx = 0; // tallest inline image on the line
  let maxAscent = baseMetrics?.ascent ?? 0; // metrics mode
  let maxDescent = baseMetrics?.descent ?? 0;

  const lineStart = () => (firstLine ? firstLineStart : contStart);

  /** Distance from `x` to the next tab stop (every tabWidth from the band left). */
  const tabAdvance = (x: number) => {
    const k = Math.floor((x - lineLeft) / tabWidth) + 1;
    return lineLeft + k * tabWidth - x;
  };

  const flushLine = (isLast: boolean) => {
    const startX = lineStart();
    const avail = lineRight - startX;

    // Trailing whitespace doesn't count toward alignment, nor is it painted.
    let end = lineTokens.length;
    let contentWidth = lineWidth;
    while (end > 0 && lineTokens[end - 1].isSpace) {
      contentWidth -= lineTokens[end - 1].width;
      end--;
    }

    let x = startX;
    let extraPerGap = 0;
    if (align === 'justify' && !isLast) {
      const gaps = lineTokens.slice(0, end).filter((t) => t.isSpace && !t.isTab).length;
      if (gaps > 0) extraPerGap = (avail - contentWidth) / gaps;
    } else if (align === 'center') {
      x += Math.max(0, (avail - contentWidth) / 2);
    } else if (align === 'right') {
      x += Math.max(0, avail - contentWidth);
    }

    const segments: LayoutSegment[] = [];
    const images: LayoutImageSegment[] = [];
    for (let i = 0; i < end; i++) {
      const t = lineTokens[i];
      if (t.image) {
        images.push({
          x,
          src: t.image.src,
          width: t.image.width,
          height: t.image.height,
          link: t.link,
          pos: t.pos,
        });
      } else {
        const seg: LayoutSegment = {
          x,
          text: t.text ?? '',
          font: t.font,
          color: t.color,
          link: t.link,
          underline: t.underline,
          strike: t.strike,
          width: t.width,
          pos: t.pos,
        };
        if (t.field) seg.field = t.field;
        segments.push(seg);
      }
      x += t.width + (t.isSpace && !t.isTab ? extraPerGap : 0);
    }

    // Caret bounds: first painted token's start … last painted token's end.
    // An empty line (empty paragraph) collapses to the paragraph's content pos;
    // continuation lines fall back to the position after the previous content.
    const firstPos = lineTokens.find((t, i) => i < end && t.pos != null)?.pos;
    let lastEnd: number | undefined;
    for (let i = end - 1; i >= 0; i--) {
      const t = lineTokens[i];
      if (t.pos != null) {
        lastEnd = t.pos + t.size;
        break;
      }
    }
    const from = firstPos ?? prevTo ?? block.pos;
    const to = lastEnd ?? from;

    let height: number;
    let baseline: number;
    if (metrics) {
      // Real metrics: line box is ascent + descent; baseline sits at ascent.
      height = maxAscent + maxDescent;
      baseline = maxAscent;
    } else {
      // Fallback: line box grows to the tallest image, else a font-size factor.
      height = Math.max(maxFontPx * LINE_HEIGHT_FACTOR, maxImagePx);
      baseline = height * BASELINE_FACTOR;
    }
    const painted = firstLine && marker ? [marker, ...segments] : segments;
    emit({ x: startX, width: lineRight - startX, height, baseline, segments: painted, images, from, to });
    prevTo = to;

    lineTokens = [];
    lineWidth = 0;
    maxFontPx = sizePx(base);
    maxImagePx = 0;
    maxAscent = baseMetrics?.ascent ?? 0;
    maxDescent = baseMetrics?.descent ?? 0;
    firstLine = false;

    // The next line may sit beside (or past) a float — fetch its band.
    band = bandFn(nominalH);
    lineLeft = band.left + indentLeft;
    lineRight = band.right - indentRight;
    contStart = marker ? Math.max(markerTextX, lineLeft) : lineLeft;
  };

  // Kerning across token boundaries: consecutive same-font text tokens (a word
  // split across runs, e.g. by a mark change) are measured cumulatively, so
  // the advance matches measuring the joined text. Resets on anything that
  // breaks the glyph run: spaces, tabs, images, fields, font changes, wraps.
  let clusterText = '';
  let clusterWidth = 0;
  let clusterFont: FontSpec | null = null;
  const sameFont = (a: FontSpec, b: FontSpec) =>
    a.family === b.family && a.sizePt === b.sizePt && a.bold === b.bold && a.italic === b.italic;
  const resetCluster = () => {
    clusterText = '';
    clusterWidth = 0;
    clusterFont = null;
  };
  const clusterable = (t: Token) =>
    t.text != null && !t.isSpace && !t.isTab && !t.field && !t.image;

  for (const token of tokens) {
    // Skip leading spaces (but keep a leading tab — it indents the line).
    if (token.isSpace && !token.isTab && lineTokens.length === 0) continue;
    if (token.isTab) token.width = tabAdvance(lineStart() + lineWidth);
    if (clusterable(token) && clusterFont && sameFont(clusterFont, token.font)) {
      token.width = Math.max(0, measure(clusterText + token.text, token.font) - clusterWidth);
    }
    const cursor = lineStart() + lineWidth;
    if (!token.isSpace && lineTokens.length > 0 && cursor + token.width > lineRight) {
      flushLine(false);
      resetCluster(); // the wrapped token starts a fresh glyph run
      if (token.isTab) token.width = tabAdvance(lineStart() + lineWidth);
      else if (clusterable(token)) token.width = measure(token.text as string, token.font);
    }
    if (clusterable(token)) {
      if (clusterFont && !sameFont(clusterFont, token.font)) resetCluster();
      clusterText += token.text;
      clusterWidth += token.width;
      clusterFont = token.font;
    } else {
      resetCluster();
    }
    lineTokens.push(token);
    lineWidth += token.width;
    if (token.image) {
      maxImagePx = Math.max(maxImagePx, token.image.height);
      maxAscent = Math.max(maxAscent, token.image.height); // image sits on the baseline
    } else {
      maxFontPx = Math.max(maxFontPx, sizePx(token.font));
      if (metrics) {
        const m = metrics(token.font);
        maxAscent = Math.max(maxAscent, m.ascent);
        maxDescent = Math.max(maxDescent, m.descent);
      }
    }
  }
  flushLine(true); // emit the paragraph's last (or only/empty) line
}

/** Wrap one paragraph into line drafts within [contentLeft, contentRight].
 *  Pure: no pagination, no vertical positioning. */
function layoutParagraph(
  block: FlowParagraph,
  contentLeft: number,
  contentRight: number,
  ctx: Ctx,
): LineDraft[] {
  const drafts: LineDraft[] = [];
  wrapParagraph(block, ctx, () => ({ left: contentLeft, right: contentRight }), (d) =>
    drafts.push(d),
  );
  return drafts;
}

/** Shift a resolved table (and everything inside it) down by `dy` px. */
function offsetTable(table: ResolvedTable, dy: number): void {
  table.y += dy;
  for (const cell of table.cells) {
    cell.y += dy;
    for (const line of cell.lines) line.y += dy;
    if (cell.tables) for (const t of cell.tables) offsetTable(t, dy);
  }
}

/** Lay out a sequence of blocks within a content box, stacking vertically from
 *  y = 0. No pagination — used for table-cell content. */
function layoutFlow(
  blocks: FlowBlock[],
  contentLeft: number,
  contentRight: number,
  ctx: Ctx,
): { lines: LayoutLine[]; tables: ResolvedTable[]; height: number } {
  const lines: LayoutLine[] = [];
  const tables: ResolvedTable[] = [];
  let y = 0;
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const d of layoutParagraph(block, contentLeft, contentRight, ctx)) {
        lines.push(draftToLine(d, y));
        y += d.height;
      }
    } else {
      const table = layoutTable(block, contentLeft, contentRight, ctx);
      offsetTable(table, y);
      tables.push(table);
      y += table.height;
    }
  }
  return { lines, tables, height: y };
}

/** Lay out a table within [contentLeft, contentRight], relative to y = 0.
 *  Columns come from cell `colwidth` (unknowns split the remaining width).
 *  Row heights are the max cell content height; rowspan cells grow the last
 *  spanned row. Whole-table pagination is handled by the caller. */
function layoutTable(
  table: FlowTable,
  contentLeft: number,
  contentRight: number,
  ctx: Ctx,
): ResolvedTable {
  const avail = contentRight - contentLeft;
  const nrows = table.rows.length;
  const ncols = table.rows.reduce(
    (m, r) => Math.max(m, r.cells.reduce((s, c) => s + c.colspan, 0)),
    0,
  );

  // Column widths: take known widths from cells, split the rest equally.
  const colWidths = new Array<number>(ncols).fill(0);
  for (const row of table.rows) {
    let col = 0;
    for (const cell of row.cells) {
      if (cell.colwidth && cell.colwidth.length === cell.colspan) {
        for (let k = 0; k < cell.colspan && col + k < ncols; k++) {
          if (colWidths[col + k] === 0) colWidths[col + k] = cell.colwidth[k];
        }
      }
      col += cell.colspan;
    }
  }
  const known = colWidths.reduce((s, w) => s + w, 0);
  const unknown = colWidths.filter((w) => w === 0).length;
  if (unknown > 0) {
    const share = Math.max(0, (avail - known) / unknown);
    for (let i = 0; i < ncols; i++) if (colWidths[i] === 0) colWidths[i] = share;
  }
  const colX = new Array<number>(ncols + 1).fill(contentLeft);
  for (let i = 0; i < ncols; i++) colX[i + 1] = colX[i] + colWidths[i];
  const tableWidth = colX[ncols] - contentLeft;

  // Lay out each cell's content; remember where it sits in the grid.
  interface CellDraft {
    startRow: number;
    startCol: number;
    colspan: number;
    rowspan: number;
    cellLeft: number;
    cellWidth: number;
    lines: LayoutLine[];
    tables: ResolvedTable[];
    contentHeight: number;
  }
  // Per-table cell margins (w:tblCellMar) override the Word defaults.
  const pad = {
    left: table.cellPadding?.left ?? CELL_PAD_X,
    right: table.cellPadding?.right ?? CELL_PAD_X,
    top: table.cellPadding?.top ?? CELL_PAD_Y,
    bottom: table.cellPadding?.bottom ?? CELL_PAD_Y,
  };

  const cellDrafts: CellDraft[] = [];
  for (let r = 0; r < nrows; r++) {
    let col = 0;
    for (const cell of table.rows[r].cells) {
      let cellWidth = 0;
      for (let k = 0; k < cell.colspan && col + k < ncols; k++) cellWidth += colWidths[col + k];
      const cellLeft = colX[col];
      const flow = layoutFlow(cell.content, cellLeft + pad.left, cellLeft + cellWidth - pad.right, ctx);
      cellDrafts.push({
        startRow: r,
        startCol: col,
        colspan: cell.colspan,
        rowspan: cell.rowspan,
        cellLeft,
        cellWidth,
        lines: flow.lines,
        tables: flow.tables,
        contentHeight: flow.height,
      });
      col += cell.colspan;
    }
  }

  // Row heights: single-row cells set the base; multi-row cells grow the last
  // spanned row if their content needs more than the rows already provide.
  const rowHeight = new Array<number>(nrows).fill(0);
  for (const c of cellDrafts) {
    if (c.rowspan === 1) {
      rowHeight[c.startRow] = Math.max(rowHeight[c.startRow], c.contentHeight + pad.top + pad.bottom);
    }
  }
  for (const c of cellDrafts) {
    if (c.rowspan > 1) {
      const need = c.contentHeight + pad.top + pad.bottom;
      let span = 0;
      for (let r = c.startRow; r < c.startRow + c.rowspan && r < nrows; r++) span += rowHeight[r];
      if (need > span) {
        const last = Math.min(c.startRow + c.rowspan - 1, nrows - 1);
        rowHeight[last] += need - span;
      }
    }
  }
  const rowY = new Array<number>(nrows + 1).fill(0);
  for (let r = 0; r < nrows; r++) rowY[r + 1] = rowY[r] + rowHeight[r];

  // Position cells and shift their content into place.
  const cells: ResolvedCell[] = cellDrafts.map((c) => {
    let height = 0;
    for (let r = c.startRow; r < c.startRow + c.rowspan && r < nrows; r++) height += rowHeight[r];
    const dy = rowY[c.startRow] + pad.top;
    const lines = c.lines.map((ln) => ({ ...ln, y: ln.y + dy }));
    c.tables.forEach((t) => offsetTable(t, dy));
    const cell: ResolvedCell = {
      x: c.cellLeft,
      y: rowY[c.startRow],
      width: c.cellWidth,
      height,
      colspan: c.colspan,
      rowspan: c.rowspan,
      lines,
    };
    if (c.tables.length > 0) cell.tables = c.tables;
    return cell;
  });

  const resolved: ResolvedTable = { x: contentLeft, y: 0, width: tableWidth, height: rowY[nrows], cells };

  // Repeating header band: contiguous header rows from the top, provided no
  // cell spans out of the band (a rowspan into the body would have to split).
  let headerRows = 0;
  while (headerRows < nrows && table.rows[headerRows].header) headerRows++;
  if (headerRows > 0 && headerRows < nrows) {
    const headerBottom = rowY[headerRows];
    const spansOut = cells.some((c) => c.y < headerBottom && c.y + c.height > headerBottom);
    if (!spansOut) resolved.headerBottom = headerBottom;
  }
  return resolved;
}

/** One laid-out top-level block awaiting vertical placement. */
type ParaItem = {
  /** Flattened paragraph (lazy — cache hits skip flattening until needed). */
  getFlow: () => FlowParagraph;
  /** Pre-wrapped constant-band lines; null when the paragraph anchors floats
   *  (those must wrap at placement time, when their y is known). */
  drafts: LineDraft[] | null;
};
type BlockItem = { para: ParaItem } | { table: ResolvedTable };

/** A rectangle text must flow around (a float's box plus its text gaps). */
interface Exclusion {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

/** Narrowest band we'll still flow text into beside a float. */
const MIN_BAND = 24;

/** Shift one cell (box + contents) vertically in place. */
function shiftCell(cell: ResolvedCell, dy: number): ResolvedCell {
  cell.y += dy;
  for (const line of cell.lines) line.y += dy;
  cell.tables?.forEach((t) => offsetTable(t, dy));
  return cell;
}

/** Deep copy of a resolved table (cells, lines, nested tables). */
function cloneTable(t: ResolvedTable): ResolvedTable {
  return { ...t, cells: t.cells.map(cloneCell) };
}

function cloneCell(cell: ResolvedCell): ResolvedCell {
  const copy: ResolvedCell = { ...cell, lines: cell.lines.map((l) => ({ ...l })) };
  if (cell.tables) copy.tables = cell.tables.map(cloneTable);
  return copy;
}

/**
 * Split a resolved table (in its own y = 0 space) at `cut` px.
 *
 * Cells entirely above stay in `top`; cells entirely below shift into `rest`.
 * Cells straddling the cut (the row being broken, or a rowspan crossing a row
 * boundary) are split: lines whose box fits above the cut stay, the remainder
 * re-stacks from the top of a continuation row in `rest` (Word-like — a line
 * straddling the boundary moves down whole). Rows after the split row follow
 * beneath the continuation row.
 */
function splitTableAt(table: ResolvedTable, cut: number): { top: ResolvedTable; rest: ResolvedTable } {
  interface Cont {
    cell: ResolvedCell;
    remLines: LayoutLine[];
    remTables: ResolvedTable[];
    firstY: number;
  }
  const straddlers = new Map<ResolvedCell, Cont>();
  let splitBottom = cut; // bottom of the broken row (cut itself if none breaks)
  let contHeight = 0; // height of the continuation row in `rest`

  for (const cell of table.cells) {
    if (cell.y >= cut || cell.y + cell.height <= cut) continue;
    splitBottom = Math.max(splitBottom, cell.y + cell.height);
    const remLines = cell.lines.filter((l) => l.y + l.height > cut);
    const remTables = (cell.tables ?? []).filter((t) => t.y + t.height > cut);
    const firstY = Math.min(
      remLines.length ? Math.min(...remLines.map((l) => l.y)) : Infinity,
      remTables.length ? Math.min(...remTables.map((t) => t.y)) : Infinity,
    );
    const first = firstY === Infinity ? cut : firstY;
    const extent =
      Math.max(
        remLines.length ? Math.max(...remLines.map((l) => l.y + l.height)) : first,
        remTables.length ? Math.max(...remTables.map((t) => t.y + t.height)) : first,
      ) - first;
    contHeight = Math.max(contHeight, extent);
    straddlers.set(cell, { cell, remLines, remTables, firstY: first });
  }

  const topCells: ResolvedCell[] = [];
  const restCells: ResolvedCell[] = [];
  for (const cell of table.cells) {
    if (cell.y + cell.height <= cut) {
      topCells.push(cell);
    } else if (cell.y >= cut) {
      // a row below the break: follows beneath the continuation row.
      // Copied so the caller can still fall back to placing the original whole.
      restCells.push(shiftCell(cloneCell(cell), contHeight - splitBottom));
    } else {
      const c = straddlers.get(cell) as Cont;
      const topLines = cell.lines.filter((l) => l.y + l.height <= cut);
      const topTables = (cell.tables ?? []).filter((t) => t.y + t.height <= cut);
      const topCell: ResolvedCell = { ...cell, height: cut - cell.y, lines: topLines };
      if (topTables.length > 0) topCell.tables = topTables;
      else delete topCell.tables;
      topCells.push(topCell);

      const delta = -c.firstY;
      const lines = c.remLines.map((l) => ({ ...l, y: l.y + delta }));
      const remTables = c.remTables.map(cloneTable);
      remTables.forEach((t) => offsetTable(t, delta));
      const contCell: ResolvedCell = {
        x: cell.x,
        y: 0,
        width: cell.width,
        height: contHeight,
        colspan: cell.colspan,
        rowspan: cell.rowspan,
        lines,
      };
      if (remTables.length > 0) contCell.tables = remTables;
      restCells.push(contCell);
    }
  }

  return {
    top: { x: table.x, y: 0, width: table.width, height: cut, cells: topCells },
    rest: {
      x: table.x,
      y: 0,
      width: table.width,
      height: contHeight + (table.height - splitBottom),
      cells: restCells,
    },
  };
}

/** Ghost copies of the header-band cells for a continuation fragment. PM
 *  positions are stripped so selection and hit-testing only ever target the
 *  original header; returns null when the band is too complex to repeat. */
function cloneHeaderCells(table: ResolvedTable, headerBottom: number): ResolvedCell[] | null {
  const band = table.cells.filter((c) => c.y + c.height <= headerBottom);
  if (band.some((c) => c.tables && c.tables.length > 0)) return null;
  return band.map((cell) => ({
    ...cell,
    tables: undefined,
    lines: cell.lines.map((l) => {
      const line: LayoutLine = { ...l, segments: l.segments.map((s) => ({ ...s, pos: undefined })) };
      delete line.from;
      delete line.to;
      if (line.images) line.images = line.images.map((im) => ({ ...im, pos: undefined }));
      return line;
    }),
  }));
}

function buildCtx(config: LayoutConfig): Ctx {
  return {
    base: { ...DEFAULT_FONT, ...config.defaultFont },
    measure: config.measureText,
    metrics: config.measureMetrics,
    tabWidth: config.tabWidth ?? DEFAULT_TAB_WIDTH,
    fieldPlaceholder: '1',
  };
}

/** Stack laid-out blocks onto pages (the paginator). `band` overrides the
 *  vertical content bounds (e.g. pushed in by a tall page header/footer). */
function placeBlocks(
  items: Iterable<BlockItem>,
  config: LayoutConfig,
  ctx: Ctx,
  band?: { top: number; bottom: number },
): ResolvedLayout {
  const { page } = config;
  const top = band?.top ?? page.margin.top;
  const bottom = band?.bottom ?? page.height - page.margin.bottom;
  const contentLeft = page.margin.left;
  const contentRight = page.width - page.margin.right;

  const pages: ResolvedPage[] = [];
  let lines: LayoutLine[] = [];
  let tables: ResolvedTable[] = [];
  let pageFloats: ResolvedFloat[] = [];
  let exclusions: Exclusion[] = []; // floats die at their page's end
  let y = top;

  const finalizePage = () => {
    const resolved: ResolvedPage = { index: pages.length, width: page.width, height: page.height, lines };
    if (tables.length > 0) resolved.tables = tables;
    if (pageFloats.length > 0) resolved.floats = pageFloats;
    pages.push(resolved);
    lines = [];
    tables = [];
    pageFloats = [];
    exclusions = [];
    y = top;
  };

  /** Whether the current page already holds content (so a break is meaningful). */
  const pageHasContent = () => lines.length > 0 || tables.length > 0 || pageFloats.length > 0;

  /** Pin a paragraph's floats relative to its start; register text exclusions. */
  const registerFloats = (flow: FlowParagraph, yPara: number) => {
    for (const f of flow.floats ?? []) {
      const baseL = f.hRel === 'page' ? 0 : contentLeft;
      const baseR = f.hRel === 'page' ? page.width : contentRight;
      const fx =
        f.hAlign === 'right'
          ? baseR - f.width
          : f.hAlign === 'center'
            ? (baseL + baseR - f.width) / 2
            : f.hAlign === 'left'
              ? baseL
              : baseL + (f.hOffset ?? 0);
      const fy =
        f.vRel === 'page'
          ? (f.vOffset ?? 0)
          : f.vRel === 'margin'
            ? top + (f.vOffset ?? 0)
            : yPara + (f.vOffset ?? 0);
      pageFloats.push({ x: fx, y: fy, width: f.width, height: f.height, src: f.src });
      if (f.wrap === 'square') {
        exclusions.push({
          left: fx - (f.distL ?? 0),
          right: fx + f.width + (f.distR ?? 0),
          top: fy - (f.distT ?? 0),
          bottom: fy + f.height + (f.distB ?? 0),
        });
      } else if (f.wrap === 'topAndBottom') {
        exclusions.push({
          left: -Infinity,
          right: Infinity,
          top: fy - (f.distT ?? 0),
          bottom: fy + f.height + (f.distB ?? 0),
        });
      } // 'none' paints only
    }
  };

  /** Widest text band at [yy, yy+h) after carving out the exclusions; null
   *  when nothing usable remains (the caller skips below the blocker). */
  const bandAt = (yy: number, h: number): { left: number; right: number } | null => {
    let L = contentLeft;
    let R = contentRight;
    for (const ex of exclusions) {
      if (ex.bottom <= yy || ex.top >= yy + h) continue;
      const leftGap = Math.min(R, ex.left) - L;
      const rightGap = R - Math.max(L, ex.right);
      if (rightGap >= leftGap) L = Math.max(L, ex.right);
      else R = Math.min(R, ex.left);
    }
    return R - L >= MIN_BAND ? { left: L, right: R } : null;
  };

  /** Wrap + place a paragraph line by line, flowing around active floats. */
  const placeParaBanded = (flow: FlowParagraph) => {
    registerFloats(flow, y);
    wrapParagraph(
      flow,
      ctx,
      (estH) => {
        for (;;) {
          if (y + estH > bottom && pageHasContent()) {
            finalizePage(); // fresh page: exclusions are gone
            continue;
          }
          const b = bandAt(y, estH);
          if (b) return b;
          const blockers = exclusions.filter((ex) => ex.top < y + estH && ex.bottom > y);
          if (blockers.length === 0) return { left: contentLeft, right: contentRight };
          y = Math.min(...blockers.map((ex) => ex.bottom)); // skip below the float
        }
      },
      (draft) => {
        if (y + draft.height > bottom && pageHasContent()) finalizePage();
        lines.push(draftToLine(draft, y));
        y += draft.height;
      },
    );
  };

  for (const item of items) {
    if ('para' in item) {
      const drafts = item.para.drafts;
      const draftsHeight = drafts?.reduce((s, d) => s + d.height, 0) ?? 0;
      const floatsAhead = exclusions.some((ex) => ex.bottom > y && ex.top < y + draftsHeight);
      if (drafts && !floatsAhead) {
        for (const d of drafts) {
          if (y + d.height > bottom && pageHasContent()) finalizePage();
          lines.push(draftToLine(d, y));
          y += d.height;
        }
      } else {
        placeParaBanded(item.para.getFlow());
      }
    } else {
      // Tables flow across pages: split at row boundaries when possible, and
      // mid-row when a single row is taller than a whole page. Header rows
      // (w:tblHeader) repeat at the top of every continuation fragment.
      let table = item.table; // laid out relative to y = 0
      for (;;) {
        const avail = bottom - y;
        if (table.height <= avail) {
          offsetTable(table, y);
          tables.push(table);
          y += table.height;
          break;
        }
        // The header band repeats only while it leaves reasonable page room.
        const hb =
          table.headerBottom != null && table.headerBottom < (bottom - top) / 2
            ? table.headerBottom
            : 0;
        // Prefer the lowest row boundary that still fits (never inside the header).
        let cut = 0;
        for (const cell of table.cells) {
          if (cell.y > hb && cell.y <= avail) cut = Math.max(cut, cell.y);
        }
        if (cut === 0) {
          // No row boundary fits the remaining space. Word only moves the row
          // to a fresh page when it would fit one whole; a row taller than a
          // full page starts right here and splits mid-row (no blank gap).
          let firstRowBottom = table.height;
          for (const cell of table.cells) {
            if (cell.y > hb && cell.y < firstRowBottom) firstRowBottom = cell.y;
          }
          const fitsFullPage = firstRowBottom - hb <= bottom - top;
          if (fitsFullPage && pageHasContent()) {
            finalizePage(); // retry with a full fresh page
            continue;
          }
          cut = avail; // split the oversize row in the space we have
        }
        if (cut <= hb) {
          // Header band swallows the remaining space — try a fresh page, or
          // place whole when the geometry is truly degenerate.
          if (pageHasContent()) {
            finalizePage();
            continue;
          }
          offsetTable(table, y);
          tables.push(table);
          y += table.height;
          break;
        }
        const { top: topFrag, rest } = splitTableAt(table, cut);
        if (rest.height >= table.height) {
          // No progress (e.g. one line taller than the page) — place whole.
          // splitTableAt copies whatever it moves, so `table` is intact.
          offsetTable(table, y);
          tables.push(table);
          y += table.height;
          break;
        }
        const topHasContent = topFrag.cells.some(
          (c) => c.lines.length > 0 || (c.tables?.length ?? 0) > 0,
        );
        if (!topHasContent && pageHasContent()) {
          // Not even one line fits the leftover space — don't paint an empty
          // table stub; start on the next page instead.
          finalizePage();
          continue;
        }
        if (hb > 0) {
          const ghosts = cloneHeaderCells(table, hb);
          if (ghosts) {
            for (const cell of rest.cells) shiftCell(cell, hb);
            rest.cells.unshift(...ghosts);
            rest.height += hb;
            rest.headerBottom = hb; // continuations keep repeating it
          }
        }
        offsetTable(topFrag, y);
        tables.push(topFrag);
        finalizePage();
        table = rest;
      }
    }
  }

  if (pageHasContent() || pages.length === 0) finalizePage();
  return { pages };
}

/** Lay out already-flattened blocks into paginated pages. Pure (no DOM);
 *  measurement is injected. */
export function layoutBlocks(blocks: FlowBlock[], config: LayoutConfig): ResolvedLayout {
  const ctx = buildCtx(config);
  const left = config.page.margin.left;
  const right = config.page.width - config.page.margin.right;
  const items: BlockItem[] = blocks.map((block) =>
    block.type === 'paragraph'
      ? {
          para: {
            getFlow: () => block,
            drafts: block.floats?.length ? null : layoutParagraph(block, left, right, ctx),
          },
        }
      : { table: layoutTable(block, left, right, ctx) },
  );
  return placeBlocks(items, config, ctx);
}

// ── Incremental re-layout (M4+) ─────────────────────────────────────

/** Shift every PM position in the drafts by `delta` (geometry is unchanged —
 *  the paragraph's text and wrap are identical, it just moved in the doc). */
function shiftDrafts(drafts: LineDraft[], delta: number): LineDraft[] {
  return drafts.map((d) => ({
    ...d,
    from: d.from != null ? d.from + delta : d.from,
    to: d.to != null ? d.to + delta : d.to,
    segments: d.segments.map((s) => (s.pos != null ? { ...s, pos: s.pos + delta } : s)),
    images: d.images.map((im) => (im.pos != null ? { ...im, pos: im.pos + delta } : im)),
  }));
}

interface ParagraphCacheEntry {
  left: number;
  right: number;
  /** Content-start position the cached drafts were computed at. */
  basePos: number;
  drafts: LineDraft[];
}

/**
 * Paragraph-level layout cache keyed on ProseMirror node identity (PM keeps
 * unchanged nodes identical across transactions). An unchanged paragraph skips
 * measuring/wrapping entirely; one that merely moved in the document gets its
 * positions shifted. Tables are not cached (laid out fresh every time).
 */
export class LayoutCache {
  readonly paragraphs = new WeakMap<PMNode, ParagraphCacheEntry>();
}

export function createLayoutCache(): LayoutCache {
  return new LayoutCache();
}

/** Word's default header/footer distance from the page edge (720 twips). */
const CHROME_DISTANCE = 48;

/** Repeating page furniture passed alongside the body document. */
export interface PageChrome {
  header?: PMNode;
  footer?: PMNode;
}

/** Strip PM positions from chrome content: the band belongs to a separate
 *  document, so its positions must never be caret-addressable. */
function stripPositions(lines: LayoutLine[], tables: ResolvedTable[]): void {
  for (const line of lines) {
    delete line.from;
    delete line.to;
    for (const seg of line.segments) delete seg.pos;
    if (line.images) for (const im of line.images) delete im.pos;
  }
  for (const t of tables) {
    for (const c of t.cells) stripPositions(c.lines, c.tables ?? []);
  }
}

/** Lay out one chrome document (header/footer) and pin it at `topY`. */
function layoutChrome(
  doc: PMNode,
  topY: number,
  left: number,
  right: number,
  ctx: Ctx,
): ResolvedChrome {
  // Chrome can't host floats — anchored images degrade to inline there.
  const flow = layoutFlow(toFlowBlocks(doc, ctx.base, false), left, right, ctx);
  const lines = flow.lines.map((l) => ({ ...l, y: l.y + topY }));
  flow.tables.forEach((t) => offsetTable(t, topY));
  stripPositions(lines, flow.tables);
  return { lines, tables: flow.tables, height: flow.height };
}

/** Lay out a ProseMirror document into paint-ready pages. With a `cache`,
 *  only paragraphs whose node changed since the previous call are re-measured.
 *  `chrome` (page header/footer documents) repeats on every page; the body
 *  band shrinks when a chrome band is taller than the page margin. */
export function layout(
  doc: PMNode,
  config: LayoutConfig,
  cache?: LayoutCache,
  chrome?: PageChrome,
): ResolvedLayout {
  const ctx = buildCtx(config);
  const { page } = config;
  const left = page.margin.left;
  const right = page.width - page.margin.right;

  // Page chrome first — a tall header/footer pushes the body band inward.
  let pageHeader: ResolvedChrome | undefined;
  let pageFooter: ResolvedChrome | undefined;
  let top = page.margin.top;
  let bottom = page.height - page.margin.bottom;
  if (chrome?.header) {
    pageHeader = layoutChrome(chrome.header, CHROME_DISTANCE, left, right, ctx);
    top = Math.max(top, CHROME_DISTANCE + pageHeader.height);
  }
  if (chrome?.footer) {
    const flow = layoutChrome(chrome.footer, 0, left, right, ctx);
    const topY = page.height - CHROME_DISTANCE - flow.height;
    pageFooter = {
      lines: flow.lines.map((l) => ({ ...l, y: l.y + topY })),
      tables: flow.tables,
      height: flow.height,
    };
    pageFooter.tables.forEach((t) => offsetTable(t, topY));
    bottom = Math.min(bottom, topY);
  }

  const items: BlockItem[] = [];
  doc.forEach((node, offset) => {
    if (node.type.name === 'paragraph') {
      const getFlow = () => paragraphToFlow(node, ctx.base, offset, true);
      // Float-anchoring paragraphs always wrap at placement time (their band
      // depends on where they land) — never cached.
      if (nodeHasFloats(node)) {
        items.push({ para: { getFlow, drafts: null } });
        return;
      }
      const contentStart = offset + 1;
      const hit = cache?.paragraphs.get(node);
      if (hit && hit.left === left && hit.right === right) {
        if (hit.basePos !== contentStart) {
          hit.drafts = shiftDrafts(hit.drafts, contentStart - hit.basePos);
          hit.basePos = contentStart;
        }
        items.push({ para: { getFlow, drafts: hit.drafts } });
        return;
      }
      const flow = paragraphToFlow(node, ctx.base, offset, true);
      const drafts = layoutParagraph(flow, left, right, ctx);
      cache?.paragraphs.set(node, { left, right, basePos: contentStart, drafts });
      items.push({ para: { getFlow: () => flow, drafts } });
    } else if (node.type.name === 'table') {
      items.push({ table: layoutTable(tableToFlow(node, ctx.base, offset), left, right, ctx) });
    }
  });
  const resolved = placeBlocks(items, config, ctx, { top, bottom });

  // Chrome with page-number fields: re-lay it out now that the page total is
  // known, so the field slot is as wide as the widest number it will show.
  const hasFields = (c?: ResolvedChrome) =>
    c?.lines.some((l) => l.segments.some((s) => s.field)) ?? false;
  if (hasFields(pageHeader) || hasFields(pageFooter)) {
    const fieldCtx: Ctx = { ...ctx, fieldPlaceholder: String(resolved.pages.length) };
    if (chrome?.header && hasFields(pageHeader)) {
      pageHeader = layoutChrome(chrome.header, CHROME_DISTANCE, left, right, fieldCtx);
    }
    if (chrome?.footer && hasFields(pageFooter)) {
      const flow = layoutChrome(chrome.footer, 0, left, right, fieldCtx);
      const topY = page.height - CHROME_DISTANCE - flow.height;
      pageFooter = {
        lines: flow.lines.map((l) => ({ ...l, y: l.y + topY })),
        tables: flow.tables,
        height: flow.height,
      };
      pageFooter.tables.forEach((t) => offsetTable(t, topY));
    }
  }

  if (pageHeader) resolved.pageHeader = pageHeader;
  if (pageFooter) resolved.pageFooter = pageFooter;
  return resolved;
}
