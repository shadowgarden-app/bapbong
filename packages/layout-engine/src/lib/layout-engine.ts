import { Mark, Node as PMNode } from 'prosemirror-model';
import type {
  Align,
  FlowBlock,
  FlowInline,
  FlowParagraph,
  FlowTable,
  FlowTableCell,
  FlowTableRow,
  FontSpec,
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
  ResolvedLayout,
  ResolvedPage,
  ResolvedTable,
} from '@shadow-garden/bapbong-contracts';

const DEFAULT_FONT: FontSpec = { family: 'Arial', sizePt: 11, bold: false, italic: false };
const PT_TO_PX = 96 / 72;
const LINE_HEIGHT_FACTOR = 1.2;
const BASELINE_FACTOR = 0.8;
const DEFAULT_TAB_WIDTH = 48; // 0.5in, Word's default tab interval
// Cell padding (px). 0 for now — Word's default cell margins (~5.4pt left/right)
// are a later refinement; keeping 0 makes table coordinates predictable.
const CELL_PAD_X = 0;
const CELL_PAD_Y = 0;

const sizePx = (font: FontSpec) => font.sizePt * PT_TO_PX;

/** Shared layout inputs, threaded through paragraph/table/cell layout. */
interface Ctx {
  base: FontSpec;
  measure: MeasureText;
  metrics?: MeasureMetrics;
  tabWidth: number;
}

function findMark(marks: readonly Mark[], name: string): Mark | undefined {
  return marks.find((m) => m.type.name === name);
}

/** Resolve a text node's marks into an InlineRun (font + color + link). */
function resolveRun(node: PMNode, base: FontSpec): InlineRun {
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
  return {
    text: node.text ?? '',
    font,
    color: color ? String(color.attrs['color']) : undefined,
    link: link ? String(link.attrs['href']) : undefined,
  };
}

/** Resolve an image node into an InlineImage. Missing dimensions become 0
 *  (the image then takes no space until real sizing is available). */
function resolveImage(node: PMNode): InlineImage {
  const a = node.attrs;
  const link = findMark(node.marks, 'link');
  return {
    src: String(a['src'] ?? ''),
    width: Number(a['width']) || 0,
    height: Number(a['height']) || 0,
    link: link ? String(link.attrs['href']) : undefined,
  };
}

/** Flatten one paragraph node into a FlowParagraph (text + inline images). */
function paragraphToFlow(node: PMNode, base: FontSpec): FlowParagraph {
  const runs: FlowInline[] = [];
  node.forEach((child) => {
    if (child.isText) runs.push(resolveRun(child, base));
    else if (child.type.name === 'image') runs.push(resolveImage(child));
  });
  const list = node.attrs['list'] as { marker?: string } | null;
  const align = node.attrs['align'] as Align | null | undefined;
  const indent = node.attrs['indent'] as ParagraphIndent | null | undefined;
  return {
    type: 'paragraph',
    runs,
    marker: list?.marker || undefined,
    align: align ?? undefined,
    indent: indent ?? undefined,
  };
}

/** Flatten a block-level node (paragraph or table) into a FlowBlock, or null
 *  for node types we don't model yet. */
function nodeToBlock(node: PMNode, base: FontSpec): FlowBlock | null {
  if (node.type.name === 'paragraph') return paragraphToFlow(node, base);
  if (node.type.name === 'table') return tableToFlow(node, base);
  return null;
}

/** Flatten a table node (table → table_row → table_cell → block+). */
function tableToFlow(node: PMNode, base: FontSpec): FlowTable {
  const rows: FlowTableRow[] = [];
  node.forEach((rowNode) => {
    const cells: FlowTableCell[] = [];
    rowNode.forEach((cellNode) => {
      const content: FlowBlock[] = [];
      cellNode.forEach((child) => {
        const block = nodeToBlock(child, base);
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
    rows.push({ cells });
  });
  return { type: 'table', rows };
}

/** Flatten a ProseMirror doc into FlowBlocks (paragraphs + tables). */
export function toFlowBlocks(doc: PMNode, defaultFont: Partial<FontSpec> = {}): FlowBlock[] {
  const base: FontSpec = { ...DEFAULT_FONT, ...defaultFont };
  const blocks: FlowBlock[] = [];
  doc.forEach((node) => {
    const block = nodeToBlock(node, base);
    if (block) blocks.push(block);
  });
  return blocks;
}

interface Token {
  /** Text content (text token), or undefined for an image token. */
  text?: string;
  /** Image payload (image token), or undefined for a text token. */
  image?: InlineImage;
  font: FontSpec;
  color?: string;
  link?: string;
  width: number;
  isSpace: boolean;
  /** A tab character: its width is resolved to the next tab stop at layout. */
  isTab?: boolean;
}

/** Tokenize one inline item: words / spaces / tabs for text, a single atom for
 *  images. Tab widths are placeholders, resolved against tab stops at layout. */
function tokenizeInline(inline: FlowInline, ctx: Ctx): Token[] {
  if ('src' in inline) {
    return [{ image: inline, font: ctx.base, link: inline.link, width: inline.width, isSpace: false }];
  }
  return inline.text
    .split(/(\t| +)/)
    .filter((part) => part.length > 0)
    .map((part) => {
      const isTab = part === '\t';
      const isSpace = isTab || /^ +$/.test(part);
      return {
        text: part,
        font: inline.font,
        color: inline.color,
        link: inline.link,
        width: isTab ? 0 : ctx.measure(part, inline.font),
        isSpace,
        isTab,
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
  return line;
}

/** Wrap one paragraph into line drafts within [contentLeft, contentRight].
 *  Pure: no pagination, no vertical positioning. */
function layoutParagraph(
  block: FlowParagraph,
  contentLeft: number,
  contentRight: number,
  ctx: Ctx,
): LineDraft[] {
  const { base, measure, metrics, tabWidth } = ctx;
  const indent = block.indent;
  const paraLeft = contentLeft + (indent?.left ?? 0);
  const paraRight = contentRight - (indent?.right ?? 0);
  // hanging outdents the first line; firstLine indents it. Mutually exclusive.
  const firstLineDelta = indent?.hanging != null ? -indent.hanging : indent?.firstLine ?? 0;
  const align: Align = block.align ?? 'left';

  const tokens = block.runs.flatMap((inline) => tokenizeInline(inline, ctx));

  // List marker hangs at the first line's start; text follows after it, and
  // wrapped lines align under that text (hanging indent).
  let marker: LayoutSegment | null = null;
  let markerWidth = 0;
  if (block.marker) {
    marker = { x: paraLeft + firstLineDelta, text: block.marker, font: base };
    markerWidth = measure(`${block.marker} `, base);
  }
  const firstLineStart = marker ? marker.x + markerWidth : paraLeft + firstLineDelta;
  const contStart = marker ? marker.x + markerWidth : paraLeft;

  // Baseline metrics for the default font seed every line (so empty lines have
  // a sensible height too).
  const baseMetrics = metrics ? metrics(base) : null;

  const drafts: LineDraft[] = [];
  let lineTokens: Token[] = [];
  let lineWidth = 0; // running width of the current line's tokens
  let firstLine = true;
  let maxFontPx = sizePx(base); // tallest text (fallback line-height mode)
  let maxImagePx = 0; // tallest inline image on the line
  let maxAscent = baseMetrics?.ascent ?? 0; // metrics mode
  let maxDescent = baseMetrics?.descent ?? 0;

  const lineStart = () => (firstLine ? firstLineStart : contStart);

  /** Distance from `x` to the next tab stop (stops every tabWidth from paraLeft). */
  const tabAdvance = (x: number) => {
    const k = Math.floor((x - paraLeft) / tabWidth) + 1;
    return paraLeft + k * tabWidth - x;
  };

  const flushLine = (isLast: boolean) => {
    const startX = lineStart();
    const avail = paraRight - startX;

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
        images.push({ x, src: t.image.src, width: t.image.width, height: t.image.height, link: t.link });
      } else {
        segments.push({ x, text: t.text ?? '', font: t.font, color: t.color, link: t.link });
      }
      x += t.width + (t.isSpace && !t.isTab ? extraPerGap : 0);
    }

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
    drafts.push({ x: startX, width: paraRight - startX, height, baseline, segments: painted, images });

    lineTokens = [];
    lineWidth = 0;
    maxFontPx = sizePx(base);
    maxImagePx = 0;
    maxAscent = baseMetrics?.ascent ?? 0;
    maxDescent = baseMetrics?.descent ?? 0;
    firstLine = false;
  };

  for (const token of tokens) {
    // Skip leading spaces (but keep a leading tab — it indents the line).
    if (token.isSpace && !token.isTab && lineTokens.length === 0) continue;
    if (token.isTab) token.width = tabAdvance(lineStart() + lineWidth);
    const cursor = lineStart() + lineWidth;
    if (!token.isSpace && lineTokens.length > 0 && cursor + token.width > paraRight) {
      flushLine(false);
      if (token.isTab) token.width = tabAdvance(lineStart() + lineWidth);
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
  const cellDrafts: CellDraft[] = [];
  for (let r = 0; r < nrows; r++) {
    let col = 0;
    for (const cell of table.rows[r].cells) {
      let cellWidth = 0;
      for (let k = 0; k < cell.colspan && col + k < ncols; k++) cellWidth += colWidths[col + k];
      const cellLeft = colX[col];
      const flow = layoutFlow(cell.content, cellLeft + CELL_PAD_X, cellLeft + cellWidth - CELL_PAD_X, ctx);
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
      rowHeight[c.startRow] = Math.max(rowHeight[c.startRow], c.contentHeight + 2 * CELL_PAD_Y);
    }
  }
  for (const c of cellDrafts) {
    if (c.rowspan > 1) {
      const need = c.contentHeight + 2 * CELL_PAD_Y;
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
    const dy = rowY[c.startRow] + CELL_PAD_Y;
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

  return { x: contentLeft, y: 0, width: tableWidth, height: rowY[nrows], cells };
}

/** Lay out already-flattened blocks into paginated pages. Pure (no DOM);
 *  measurement is injected. */
export function layoutBlocks(blocks: FlowBlock[], config: LayoutConfig): ResolvedLayout {
  const base: FontSpec = { ...DEFAULT_FONT, ...config.defaultFont };
  const { page } = config;
  const ctx: Ctx = {
    base,
    measure: config.measureText,
    metrics: config.measureMetrics,
    tabWidth: config.tabWidth ?? DEFAULT_TAB_WIDTH,
  };
  const left = page.margin.left;
  const right = page.width - page.margin.right;
  const top = page.margin.top;
  const bottom = page.height - page.margin.bottom;

  const pages: ResolvedPage[] = [];
  let lines: LayoutLine[] = [];
  let tables: ResolvedTable[] = [];
  let y = top;

  const finalizePage = () => {
    const resolved: ResolvedPage = { index: pages.length, width: page.width, height: page.height, lines };
    if (tables.length > 0) resolved.tables = tables;
    pages.push(resolved);
    lines = [];
    tables = [];
    y = top;
  };

  /** Whether the current page already holds content (so a break is meaningful). */
  const pageHasContent = () => lines.length > 0 || tables.length > 0;

  const placeLine = (d: LineDraft) => {
    if (y + d.height > bottom && pageHasContent()) finalizePage();
    lines.push(draftToLine(d, y));
    y += d.height;
  };

  const placeTable = (table: ResolvedTable) => {
    if (y + table.height > bottom && pageHasContent()) finalizePage();
    offsetTable(table, y); // table was laid out relative to y = 0
    tables.push(table);
    y += table.height;
  };

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const d of layoutParagraph(block, left, right, ctx)) placeLine(d);
    } else {
      placeTable(layoutTable(block, left, right, ctx));
    }
  }

  if (pageHasContent() || pages.length === 0) finalizePage();
  return { pages };
}

/** Lay out a ProseMirror document into paint-ready pages. */
export function layout(doc: PMNode, config: LayoutConfig): ResolvedLayout {
  return layoutBlocks(toFlowBlocks(doc, config.defaultFont), config);
}
