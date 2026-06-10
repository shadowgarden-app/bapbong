import { Mark, Node as PMNode } from 'prosemirror-model';
import type {
  FlowBlock,
  FontSpec,
  InlineRun,
  LayoutConfig,
  LayoutLine,
  LayoutSegment,
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
    blocks.push({ type: 'paragraph', runs, marker: list?.marker || undefined });
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

  const pushLine = (segments: LayoutSegment[], height: number) => {
    if (y + height > bottom && lines.length > 0) finalizePage();
    lines.push({ x: left, y, width: right - left, height, baseline: height * BASELINE_FACTOR, segments });
    y += height;
  };

  for (const block of blocks) {
    const tokens = block.runs.flatMap((run) => tokenize(run, measureText));

    let marker: LayoutSegment | null = null;
    let indent = left;
    if (block.marker) {
      marker = { x: left, text: block.marker, font: base };
      indent = left + measureText(`${block.marker} `, base);
    }

    let segments: LayoutSegment[] = [];
    let x = indent;
    let maxSize = sizePx(base);

    const flushLine = () => {
      const height = maxSize * LINE_HEIGHT_FACTOR;
      pushLine(marker ? [marker, ...segments] : segments, height);
      marker = null; // marker only on the paragraph's first line
      segments = [];
      x = left; // wrapped lines start at the margin (no hanging indent yet)
      maxSize = sizePx(base);
    };

    for (const token of tokens) {
      if (token.isSpace && segments.length === 0) continue; // no leading space
      if (!token.isSpace && segments.length > 0 && x + token.width > right) flushLine();
      segments.push({ x, text: token.text, font: token.font, color: token.color, link: token.link });
      x += token.width;
      maxSize = Math.max(maxSize, sizePx(token.font));
    }
    flushLine(); // emit the paragraph's last (or only/empty) line
  }

  if (lines.length > 0 || pages.length === 0) finalizePage();
  return { pages };
}

/** Lay out a ProseMirror document into paint-ready pages. */
export function layout(doc: PMNode, config: LayoutConfig): ResolvedLayout {
  return layoutBlocks(toFlowBlocks(doc, config.defaultFont), config);
}
