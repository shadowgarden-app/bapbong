import { Mark, Node as PMNode } from 'prosemirror-model';
import type {
  Align,
  FlowBlock,
  FontSpec,
  InlineRun,
  LayoutConfig,
  LayoutLine,
  LayoutSegment,
  ParagraphIndent,
  ResolvedLayout,
  ResolvedPage,
} from '@shadow-garden/bapbong-contracts';

const DEFAULT_FONT: FontSpec = { family: 'Arial', sizePt: 11, bold: false, italic: false };
const PT_TO_PX = 96 / 72;
const LINE_HEIGHT_FACTOR = 1.2;
const BASELINE_FACTOR = 0.8;

const sizePx = (font: FontSpec) => font.sizePt * PT_TO_PX;

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

/** Flatten a ProseMirror doc into FlowBlocks. Paragraphs only for now;
 *  tables/images are handled in a later increment. */
export function toFlowBlocks(doc: PMNode, defaultFont: Partial<FontSpec> = {}): FlowBlock[] {
  const base: FontSpec = { ...DEFAULT_FONT, ...defaultFont };
  const blocks: FlowBlock[] = [];
  doc.forEach((node) => {
    if (node.type.name !== 'paragraph') return;
    const runs: InlineRun[] = [];
    node.forEach((child) => {
      if (child.isText) runs.push(resolveRun(child, base));
    });
    const list = node.attrs['list'] as { marker?: string } | null;
    const align = node.attrs['align'] as Align | null | undefined;
    const indent = node.attrs['indent'] as ParagraphIndent | null | undefined;
    blocks.push({
      type: 'paragraph',
      runs,
      marker: list?.marker || undefined,
      align: align ?? undefined,
      indent: indent ?? undefined,
    });
  });
  return blocks;
}

interface Token {
  text: string;
  font: FontSpec;
  color?: string;
  link?: string;
  width: number;
  isSpace: boolean;
}

function tokenize(run: InlineRun, measure: LayoutConfig['measureText']): Token[] {
  return run.text
    .split(/(\s+)/)
    .filter((part) => part.length > 0)
    .map((text) => ({
      text,
      font: run.font,
      color: run.color,
      link: run.link,
      width: measure(text, run.font),
      isSpace: /^\s+$/.test(text),
    }));
}

/** Lay out already-flattened blocks. Pure (no DOM); measurement is injected. */
export function layoutBlocks(blocks: FlowBlock[], config: LayoutConfig): ResolvedLayout {
  const base: FontSpec = { ...DEFAULT_FONT, ...config.defaultFont };
  const { page, measureText } = config;
  const left = page.margin.left;
  const right = page.width - page.margin.right;
  const top = page.margin.top;
  const bottom = page.height - page.margin.bottom;

  const pages: ResolvedPage[] = [];
  let lines: LayoutLine[] = [];
  let y = top;

  const finalizePage = () => {
    pages.push({ index: pages.length, width: page.width, height: page.height, lines });
    lines = [];
    y = top;
  };

  const pushLine = (segments: LayoutSegment[], x: number, width: number, height: number) => {
    if (y + height > bottom && lines.length > 0) finalizePage();
    lines.push({ x, y, width, height, baseline: height * BASELINE_FACTOR, segments });
    y += height;
  };

  for (const block of blocks) {
    const indent = block.indent;
    const paraLeft = left + (indent?.left ?? 0);
    const paraRight = right - (indent?.right ?? 0);
    // hanging outdents the first line; firstLine indents it. Mutually exclusive.
    const firstLineDelta = indent?.hanging != null ? -indent.hanging : indent?.firstLine ?? 0;
    const align: Align = block.align ?? 'left';

    const tokens = block.runs.flatMap((run) => tokenize(run, measureText));

    // List marker hangs at the first line's start; text follows after it, and
    // wrapped lines align under that text (hanging indent).
    let marker: LayoutSegment | null = null;
    let markerWidth = 0;
    if (block.marker) {
      marker = { x: paraLeft + firstLineDelta, text: block.marker, font: base };
      markerWidth = measureText(`${block.marker} `, base);
    }
    const firstLineStart = marker ? marker.x + markerWidth : paraLeft + firstLineDelta;
    const contStart = marker ? marker.x + markerWidth : paraLeft;

    let lineTokens: Token[] = [];
    let lineWidth = 0; // running width of the current line's tokens
    let firstLine = true;
    let maxSize = sizePx(base);

    const lineStart = () => (firstLine ? firstLineStart : contStart);

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
        const gaps = lineTokens.slice(0, end).filter((t) => t.isSpace).length;
        if (gaps > 0) extraPerGap = (avail - contentWidth) / gaps;
      } else if (align === 'center') {
        x += Math.max(0, (avail - contentWidth) / 2);
      } else if (align === 'right') {
        x += Math.max(0, avail - contentWidth);
      }

      const segments: LayoutSegment[] = [];
      for (let i = 0; i < end; i++) {
        const t = lineTokens[i];
        segments.push({ x, text: t.text, font: t.font, color: t.color, link: t.link });
        x += t.width + (t.isSpace ? extraPerGap : 0);
      }

      const height = maxSize * LINE_HEIGHT_FACTOR;
      const painted = firstLine && marker ? [marker, ...segments] : segments;
      pushLine(painted, startX, paraRight - startX, height);

      lineTokens = [];
      lineWidth = 0;
      maxSize = sizePx(base);
      firstLine = false;
    };

    for (const token of tokens) {
      if (token.isSpace && lineTokens.length === 0) continue; // no leading space
      const cursor = lineStart() + lineWidth;
      if (!token.isSpace && lineTokens.length > 0 && cursor + token.width > paraRight) {
        flushLine(false);
      }
      lineTokens.push(token);
      lineWidth += token.width;
      maxSize = Math.max(maxSize, sizePx(token.font));
    }
    flushLine(true); // emit the paragraph's last (or only/empty) line
  }

  if (lines.length > 0 || pages.length === 0) finalizePage();
  return { pages };
}

/** Lay out a ProseMirror document into paint-ready pages. */
export function layout(doc: PMNode, config: LayoutConfig): ResolvedLayout {
  return layoutBlocks(toFlowBlocks(doc, config.defaultFont), config);
}
