import type {
  CaretRect,
  FontSpec,
  LayoutLine,
  MeasureText,
  PagePoint,
  ResolvedLayout,
  ResolvedPage,
  ResolvedTable,
  SelectionRect,
} from '@shadow-garden/bapbong-contracts';

/** Width of the highlight stub drawn for an empty line inside a selection. */
const EMPTY_LINE_RECT_WIDTH = 6;

/** A caret-addressable inline item on a line (text segment or image). */
interface Item {
  x: number;
  width: number;
  pos: number;
  size: number;
  text?: string;
  font?: FontSpec;
}

/** All caret-addressable lines of a page (top-level + table cells, recursive),
 *  sorted by (y, x). Lines without positions (decoration) are excluded. */
function allLines(page: ResolvedPage): LayoutLine[] {
  const lines = [...page.lines];
  const visit = (t: ResolvedTable) => {
    for (const cell of t.cells) {
      lines.push(...cell.lines);
      cell.tables?.forEach(visit);
    }
  };
  page.tables?.forEach(visit);
  return lines
    .filter((l) => l.from != null)
    .sort((a, b) => a.y - b.y || a.x - b.x);
}

/** Positioned items of a line, sorted by x. Marker segments (no pos) are skipped. */
function lineItems(line: LayoutLine, measure: MeasureText): Item[] {
  const items: Item[] = [];
  for (const seg of line.segments) {
    if (seg.pos == null) continue;
    if (seg.field) {
      // Page-number fields are atoms: one PM position, midpoint rule.
      items.push({
        x: seg.x,
        width: seg.width ?? measure(seg.text, seg.font),
        pos: seg.pos,
        size: 1,
      });
      continue;
    }
    items.push({
      x: seg.x,
      width: measure(seg.text, seg.font),
      pos: seg.pos,
      size: seg.text.length,
      text: seg.text,
      font: seg.font,
    });
  }
  for (const img of line.images ?? []) {
    if (img.pos == null) continue;
    items.push({ x: img.x, width: img.width, pos: img.pos, size: 1 });
  }
  return items.sort((a, b) => a.x - b.x);
}

/** Caret x for `pos` on `line` (clamped to the line's content edges). */
function caretXOnLine(
  line: LayoutLine,
  pos: number,
  measure: MeasureText,
): number {
  const items = lineItems(line, measure);
  if (items.length === 0) return line.x;
  const first = items[0];
  if (pos <= first.pos) return first.x;
  for (const it of items) {
    if (pos > it.pos + it.size) continue;
    if (pos <= it.pos) return it.x; // falls in a swallowed gap before this item
    if (it.text != null && it.font) {
      return it.x + measure(it.text.slice(0, pos - it.pos), it.font);
    }
    return pos === it.pos ? it.x : it.x + it.width; // image atom
  }
  const last = items[items.length - 1];
  return last.x + last.width;
}

/** PM position for a page-local x on `line` (nearest caret boundary). */
function posAtLineX(line: LayoutLine, x: number, measure: MeasureText): number {
  const items = lineItems(line, measure);
  if (items.length === 0) return line.from ?? 0;
  if (x <= items[0].x) return items[0].pos;
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const end = it.x + it.width;
    if (x <= end) {
      if (it.text != null && it.font) {
        // walk char boundaries; place the caret at the nearest one
        let prevW = 0;
        for (let c = 1; c <= it.size; c++) {
          const w = measure(it.text.slice(0, c), it.font);
          if (x < it.x + (prevW + w) / 2) return it.pos + c - 1;
          prevW = w;
        }
        return it.pos + it.size;
      }
      return x < it.x + it.width / 2 ? it.pos : it.pos + it.size; // image atom
    }
    const next = items[i + 1];
    if (next && x < next.x) {
      // inside an inter-item gap (justify space): snap to the nearer edge
      return x < (end + next.x) / 2 ? it.pos + it.size : next.pos;
    }
  }
  const last = items[items.length - 1];
  return last.pos + last.size;
}

/** Pick the line that should display the caret for `pos` on a page, or null. */
function lineForPos(lines: LayoutLine[], pos: number): LayoutLine | null {
  // Strictly inside a line's content (or an empty line's collapsed slot).
  for (const line of lines) {
    const from = line.from as number;
    const to = line.to as number;
    if ((pos >= from && pos < to) || (from === to && pos === from)) return line;
  }
  // Otherwise clamp to the end of the closest line whose content ends at/before
  // pos (covers line-end carets and whitespace swallowed at wrap points).
  let best: LayoutLine | null = null;
  for (const line of lines) {
    const to = line.to as number;
    if (to <= pos && (best == null || to > (best.to as number))) best = line;
  }
  return best;
}

/** Page-local point → PM position (nearest caret slot), or null on an empty page. */
export function hitTest(
  layout: ResolvedLayout,
  point: PagePoint,
  measure: MeasureText,
): number | null {
  const page = layout.pages[point.pageIndex];
  if (!page) return null;
  const lines = allLines(page);
  if (lines.length === 0) return null;

  // Lines whose vertical band contains the point (table cells can sit side by
  // side); otherwise the vertically nearest line.
  const hits = lines.filter((l) => point.y >= l.y && point.y < l.y + l.height);
  let target: LayoutLine;
  if (hits.length > 0) {
    const xDist = (l: LayoutLine) =>
      point.x < l.x
        ? l.x - point.x
        : point.x > l.x + l.width
          ? point.x - (l.x + l.width)
          : 0;
    target = hits.reduce((a, b) => (xDist(b) < xDist(a) ? b : a));
  } else {
    const yDist = (l: LayoutLine) =>
      point.y < l.y ? l.y - point.y : point.y - (l.y + l.height);
    target = lines.reduce((a, b) => (yDist(b) < yDist(a) ? b : a));
  }
  return posAtLineX(target, point.x, measure);
}

/** PM position → caret rectangle in page-local coordinates, or null. */
export function caretRect(
  layout: ResolvedLayout,
  pos: number,
  measure: MeasureText,
): CaretRect | null {
  for (const page of layout.pages) {
    const lines = allLines(page);
    if (lines.length === 0) continue;
    // Skip pages whose content ends before pos (unless it's the last chance).
    const line = lineForPos(lines, pos);
    if (!line) continue;
    const maxTo = Math.max(...lines.map((l) => l.to as number));
    if (pos > maxTo && page.index < layout.pages.length - 1) continue;
    const clamped = Math.min(pos, line.to as number);
    return {
      pageIndex: page.index,
      x: caretXOnLine(line, clamped, measure),
      y: line.y,
      height: line.height,
    };
  }
  return null;
}

/** Selection [from, to) → highlight rectangles (one per line touched). */
export function selectionRects(
  layout: ResolvedLayout,
  from: number,
  to: number,
  measure: MeasureText,
): SelectionRect[] {
  if (to <= from) return [];
  const rects: SelectionRect[] = [];
  for (const page of layout.pages) {
    for (const line of allLines(page)) {
      const lineFrom = line.from as number;
      const lineTo = line.to as number;
      if (lineFrom === lineTo) {
        // Empty line: show a stub when the selection passes over it.
        if (from <= lineFrom && to > lineTo) {
          rects.push({
            pageIndex: page.index,
            x: line.x,
            y: line.y,
            width: EMPTY_LINE_RECT_WIDTH,
            height: line.height,
          });
        }
        continue;
      }
      const s = Math.max(from, lineFrom);
      const e = Math.min(to, lineTo);
      if (e <= s) continue;
      const x1 = caretXOnLine(line, s, measure);
      const x2 = caretXOnLine(line, e, measure);
      rects.push({
        pageIndex: page.index,
        x: x1,
        y: line.y,
        width: Math.max(x2 - x1, 2),
        height: line.height,
      });
    }
  }
  return rects;
}

/** Caret position one visual line above/below `pos`, keeping `goalX`. Returns
 *  null at the document's vertical edges. */
export function verticalCaret(
  layout: ResolvedLayout,
  pos: number,
  dir: -1 | 1,
  goalX: number,
  measure: MeasureText,
): number | null {
  // Locate the current line.
  let pageIdx = -1;
  let current: LayoutLine | null = null;
  for (const page of layout.pages) {
    const lines = allLines(page);
    if (lines.length === 0) continue;
    const line = lineForPos(lines, pos);
    if (line && pos <= Math.max(...lines.map((l) => l.to as number))) {
      pageIdx = page.index;
      current = line;
      break;
    }
  }
  if (!current || pageIdx < 0) return null;

  const pick = (
    cands: LayoutLine[],
    extremeTop: boolean,
  ): LayoutLine | null => {
    if (cands.length === 0) return null;
    const edgeY = extremeTop
      ? Math.min(...cands.map((l) => l.y))
      : Math.max(...cands.map((l) => l.y));
    const band = cands.filter((l) => Math.abs(l.y - edgeY) < 0.5);
    const xDist = (l: LayoutLine) =>
      goalX < l.x
        ? l.x - goalX
        : goalX > l.x + l.width
          ? goalX - (l.x + l.width)
          : 0;
    return band.reduce((a, b) => (xDist(b) < xDist(a) ? b : a));
  };

  // Same page first: nearest band strictly above/below the current line.
  const sameLines = allLines(layout.pages[pageIdx]);
  const cands =
    dir === 1
      ? sameLines.filter((l) => l.y > current.y + 0.5)
      : sameLines.filter((l) => l.y < current.y - 0.5);
  let target = pick(cands, dir === 1);

  // Fall through to the adjacent page.
  if (!target) {
    const nextPage = layout.pages[pageIdx + dir];
    if (!nextPage) return null;
    target = pick(allLines(nextPage), dir === 1);
    if (!target) return null;
  }
  return posAtLineX(target, goalX, measure);
}

/** A document image under a page point: the carrying node's PM position plus
 *  its page-local box — what the resize plugin anchors its frame to. */
export interface ImageHit {
  pos: number;
  pageIndex: number;
  rect: { x: number; y: number; width: number; height: number };
  kind: 'inline' | 'float';
}

/** The topmost document image at a page-local point, or null. Floats win over
 *  inline images (they paint over the text); among floats, the later one wins
 *  (painted last = on top). Chrome/textbox art carries no pos and is skipped. */
export function imageAtPoint(
  layout: ResolvedLayout,
  point: PagePoint,
): ImageHit | null {
  const page = layout.pages[point.pageIndex];
  if (!page) return null;
  // A rotated image paints rotated around its box center; hit-test in its
  // local space by inverse-rotating the point.
  const inside = (
    x: number,
    y: number,
    w: number,
    h: number,
    rotation?: number,
  ) => {
    let px = point.x;
    let py = point.y;
    if (rotation) {
      const rad = (-rotation * Math.PI) / 180;
      const cx = x + w / 2;
      const cy = y + h / 2;
      const dx = px - cx;
      const dy = py - cy;
      px = cx + dx * Math.cos(rad) - dy * Math.sin(rad);
      py = cy + dx * Math.sin(rad) + dy * Math.cos(rad);
    }
    return px >= x && px <= x + w && py >= y && py <= y + h;
  };

  const floats = [...(page.floats ?? [])];
  const visitCells = (t: ResolvedTable) => {
    for (const cell of t.cells) {
      floats.push(...(cell.floats ?? []));
      cell.tables?.forEach(visitCells);
    }
  };
  page.tables?.forEach(visitCells);
  for (let i = floats.length - 1; i >= 0; i--) {
    const f = floats[i];
    if (f.pos == null || !inside(f.x, f.y, f.width, f.height, f.rotation))
      continue;
    return {
      pos: f.pos,
      pageIndex: point.pageIndex,
      rect: { x: f.x, y: f.y, width: f.width, height: f.height },
      kind: 'float',
    };
  }

  for (const line of allLines(page)) {
    for (const img of line.images ?? []) {
      if (img.pos == null) continue;
      // The image's bottom edge sits on the baseline (matches the painter).
      const top = line.y + line.baseline - img.height;
      if (!inside(img.x, top, img.width, img.height, img.rotation)) continue;
      return {
        pos: img.pos,
        pageIndex: point.pageIndex,
        rect: { x: img.x, y: top, width: img.width, height: img.height },
        kind: 'inline',
      };
    }
  }
  return null;
}
