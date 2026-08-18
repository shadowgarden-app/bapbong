import { Mark, Node as PMNode } from 'prosemirror-model';
import { glyphKey, perf, sameGlyphRun } from '@shadow-garden/bapbong-contracts';
import {
  createNumberingCounter,
  type NumberingCounter,
  type NumberingDefs,
} from '@shadow-garden/bapbong-model';
import type {
  Align,
  BorderSide,
  CellDiagonals,
  CellPadding,
  ImageCrop,
  ColumnConfig,
  DocCompat,
  FlowBlock,
  FlowFloat,
  FlowInline,
  FlowParagraph,
  FlowTable,
  FlowTableCell,
  FlowTableRow,
  FontFace,
  FontSpec,
  InlineField,
  InlineImage,
  InlineRun,
  LayoutConfig,
  LayoutImageSegment,
  LayoutLine,
  ParagraphKeeps,
  LayoutSegment,
  MarkerStyle,
  MeasureMetrics,
  MeasureText,
  PageConfig,
  ParagraphBox,
  ResolvedChromeSet,
  ParagraphBorders,
  ParagraphIndent,
  ParagraphSpacing,
  ResolvedCell,
  ResolvedChrome,
  ResolvedFloat,
  ResolvedFootnotes,
  ResolvedLayout,
  ResolvedPage,
  ResolvedTable,
  SectionConfig,
  TableBorders,
  TabStop,
} from '@shadow-garden/bapbong-contracts';

const DEFAULT_FONT: FontSpec = {
  family: 'Arial',
  sizePt: 11,
  bold: false,
  italic: false,
};
const PT_TO_PX = 96 / 72;

/** Default point size per heading level (the run base for a heading paragraph;
 *  explicit fontSize marks — e.g. on imported headings — still override). */
const HEADING_PT: Record<number, number> = {
  1: 24,
  2: 18,
  3: 14,
  4: 12,
  5: 11,
  6: 11,
};
const LINE_HEIGHT_FACTOR = 1.2;
const BASELINE_FACTOR = 0.8;
const DEFAULT_TAB_WIDTH = 48; // 0.5in, Word's default tab interval
const SUPERSUB_SCALE = 0.66; // super/subscript font size relative to the run
// Default cell padding (px) — Word's w:tblCellMar defaults: 108 twips
// (= 7.2px) left/right, 0 top/bottom. Tables can override via cellPadding.
const CELL_PAD_X = 7.2;
const CELL_PAD_Y = 0;
// Footnotes render smaller than the body (Word's footnote style ~ 10/12pt).
const FOOTNOTE_FONT_SCALE = 0.85;
// Vertical gap reserved above the first footnote line — holds the separator
// rule with a little breathing room above and below.
const FOOTNOTE_AREA_GAP = 12;

const sizePx = (font: FontSpec) => font.sizePt * PT_TO_PX;

/** A footnote body laid out flat (relative to y = 0), keyed by display number
 *  for the placer to pull from when a reference lands on a page. */
interface FootnoteBody {
  lines: LayoutLine[];
  height: number;
}

/** Shared layout inputs, threaded through paragraph/table/cell layout. */
interface Ctx {
  base: FontSpec;
  measure: MeasureText;
  metrics?: MeasureMetrics;
  tabWidth: number;
  /** Width placeholder for page-number fields (digit count of the page
   *  total once known; '1' on the first pass). */
  fieldPlaceholder: string;
  /** The document's Word compatibility profile (doc.attrs.compat, else the
   *  config's, else current Word). Rules that differ by Word version ask
   *  this and nothing else. */
  compat: DocCompat;
}

/** Current Word's rules — what an editor-authored document (no settings.xml
 *  behind it) is laid out by. Mirrors what Word 365 writes on a new file. */
const CURRENT_WORD_COMPAT: DocCompat = {
  mode: 15,
  htmlAutoSpacing: true,
  tableIndentToBorder: true,
  normalStyleYieldsToTableStyle: false,
  underlineTrailingSpaces: false,
  expandLineBeforeSoftBreak: true,
};

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
  if (findMark(marks, 'dstrike')) run.dstrike = true;
  if (findMark(marks, 'smallCaps')) run.smallCaps = true;
  const highlight = findMark(marks, 'highlight');
  if (highlight) run.background = String(highlight.attrs['color']);
  const va = findMark(marks, 'vertAlign');
  // The font is reduced in tokenizeInline (one place, both entry paths).
  if (va) run.vertAlign = va.attrs['value'] === 'sub' ? 'sub' : 'super';
  const tracked = findMark(marks, 'letterSpacing');
  if (tracked) {
    // Twips (1/20 pt) → px. Absolute: it is NOT reduced along with the font
    // for superscript or small caps, matching Word.
    const tw = Number(tracked.attrs['twips']);
    if (!Number.isNaN(tw) && tw !== 0)
      font.letterSpacing = (tw / 20) * (96 / 72);
  }
  const kern = findMark(marks, 'kern');
  if (kern) {
    // "The smallest font size which shall have its kerning automatically
    // adjusted; if the font size is smaller than this value, then no font
    // kerning shall be performed." Both sides in half-points. The comparison
    // belongs here, not at import: this is where the run's size is finally
    // settled, after the fontSize mark has had its say over the base.
    //
    // Set BOTH ways round, because the absence of this mark now means "no
    // kerning" (Word's default) rather than "kerning": a run only opts in by
    // carrying a threshold its size clears.
    //
    // A threshold of 0 therefore kerns at every size, which is the literal
    // reading of "the smallest font size". Word 16 does write w:kern w:val="0"
    // — the built-in Normal (Web) style in one of our fixtures carries it
    // under a docDefaults of 2 — and it MAY be meant as "off"; nothing in the
    // spec or in Microsoft's docs says which. Following the letter of the spec
    // affects 3 paragraphs and 4 runs across the corpus, all Vietnamese, where
    // kerning measures 0px either way, so no document can tell the difference.
    const threshold = Number(kern.attrs['halfPoints']);
    if (Number.isFinite(threshold)) font.kerning = threshold <= font.sizePt * 2;
  }
  const scaled = findMark(marks, 'charScale');
  if (scaled) {
    // Percent → factor. 100 is Word's default; the mark still exists at that
    // value (it may override a style), it just has nothing to apply.
    const pct = Number(scaled.attrs['percent']);
    if (!Number.isNaN(pct) && pct > 0 && pct !== 100) font.scaleX = pct / 100;
  }
  const raised = findMark(marks, 'position');
  if (raised) {
    // Half-points → px. Word does NOT grow the line box for a raised run, so
    // neither do we: the shift is purely a paint-time offset.
    const hp = Number(raised.attrs['halfPoints']);
    if (!Number.isNaN(hp) && hp !== 0) run.raise = (hp / 2) * (96 / 72);
  }
  const fn = findMark(marks, 'footnote');
  if (fn) run.footnoteRef = Number(fn.attrs['num']) || undefined;
  const cm = findMark(marks, 'comment');
  if (cm) run.commentIds = cm.attrs['ids'] as number[];
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
  // A sizeless bitmap (inserted without a measured size) still needs a visible
  // box — 96px square, matching the exporter's default extent. Shapes keep 0
  // (a horizontal line legitimately has zero height).
  const fallback = a['shape'] ? 0 : 96;
  return {
    src: String(a['src'] ?? ''),
    width: Number(a['width']) || fallback,
    height: Number(a['height']) || fallback,
    link: link ? String(link.attrs['href']) : undefined,
    ...(a['shape'] ? { shape: a['shape'] as InlineImage['shape'] } : {}),
    ...(a['crop'] ? { crop: a['crop'] as ImageCrop } : {}),
    ...(a['outline'] ? { outline: a['outline'] as BorderSide } : {}),
    ...(Number(a['rotation']) ? { rotation: Number(a['rotation']) } : {}),
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
  marker?: string,
  markerStyle?: MarkerStyle,
): FlowParagraph {
  const contentStart = nodePos + 1;
  // A heading paragraph sizes its runs from the level by default (bigger +
  // bold); explicit fontSize/strong marks on a run still win. Title/Subtitle
  // (named styles without an outline level) get Word-like defaults the same
  // way: Title large, Subtitle a modest italic.
  const headingLevel = node.attrs['heading'] as number | null;
  const styleId = node.attrs['styleId'] as string | null;
  const runBase: FontSpec = headingLevel
    ? { ...base, sizePt: HEADING_PT[headingLevel] ?? base.sizePt, bold: true }
    : styleId === 'Title'
      ? { ...base, sizePt: 28 }
      : styleId === 'Subtitle'
        ? { ...base, sizePt: 14, italic: true }
        : base;
  const runs: FlowInline[] = [];
  const floats: FlowFloat[] = [];
  node.forEach((child, offset) => {
    if (child.isText)
      runs.push(resolveRun(child, runBase, contentStart + offset));
    else if (child.type.name === 'image') {
      const float = child.attrs['float'] as Omit<
        FlowFloat,
        'src' | 'width' | 'height'
      > | null;
      if (float && allowFloats) {
        const floatFallback = child.attrs['shape'] ? 0 : 96;
        const f: FlowFloat = {
          ...float,
          src: String(child.attrs['src'] ?? ''),
          width: Number(child.attrs['width']) || floatFallback,
          height: Number(child.attrs['height']) || floatFallback,
          pos: contentStart + offset,
          ...(child.attrs['shape']
            ? { shape: child.attrs['shape'] as FlowFloat['shape'] }
            : {}),
          ...(child.attrs['crop']
            ? { crop: child.attrs['crop'] as ImageCrop }
            : {}),
          ...(child.attrs['outline']
            ? { outline: child.attrs['outline'] as BorderSide }
            : {}),
          ...(Number(child.attrs['rotation'])
            ? { rotation: Number(child.attrs['rotation']) }
            : {}),
        };
        // Textbox content rides the image node as PM JSON; rebuild it and
        // flatten like any other flow (no nested floats inside the box). It is
        // a whole story, tables included — `nodeToBlock`, not the paragraph
        // path, is what makes a table inside a textbox reach the page.
        const tb = child.attrs['textbox'] as {
          blocks: unknown[];
          inset?: { l: number; t: number; r: number; b: number };
          anchor?: 'ctr' | 'b';
        } | null;
        if (tb && tb.blocks.length > 0) {
          const schema = child.type.schema;
          f.content = tb.blocks
            .map((json, i) => nodeToBlock(schema.nodeFromJSON(json), base, i))
            .filter((b): b is FlowBlock => b !== null);
          if (tb.inset) f.inset = tb.inset;
          if (tb.anchor) f.anchor = tb.anchor;
        }
        floats.push(f);
      } else {
        runs.push(resolveImage(child, contentStart + offset));
      }
    } else if (child.type.name === 'page_field')
      runs.push(resolveField(child, runBase, contentStart + offset));
    else if (child.type.name === 'hard_break')
      runs.push({ break: true, pos: contentStart + offset });
    else if (child.type.name === 'column_break')
      runs.push({ columnBreak: true, pos: contentStart + offset });
  });
  const list = node.attrs['list'] as { marker?: string } | null;
  const align = node.attrs['align'] as Align | null | undefined;
  const indent = node.attrs['indent'] as ParagraphIndent | null | undefined;
  const flow: FlowParagraph = {
    type: 'paragraph',
    runs,
    marker: marker ?? (list?.marker || undefined),
    ...(markerStyle && { markerStyle }),
    align: align ?? undefined,
    indent: indent ?? undefined,
    pos: contentStart,
    end: contentStart + node.content.size,
  };
  if (floats.length > 0) flow.floats = floats;
  // The paragraph mark's font: it sizes the last line together with that
  // line's runs (and the whole line of a run-less paragraph). A stale value
  // left by an edit cannot shrink text — the mark only ever RAISES a line's
  // maxima — so it is carried whether or not the paragraph has runs.
  const markFont = node.attrs['markFont'] as Partial<FontFace> | null;
  if (markFont) flow.markFont = markFont;
  const tabs = node.attrs['tabs'] as TabStop[] | null;
  if (tabs) flow.tabs = tabs;
  const spacing = effectiveSpacing(
    node.attrs['spacing'] as ParagraphSpacing | null,
    node.attrs['contextualSpacing'] as {
      before: boolean;
      after: boolean;
    } | null,
  );
  if (spacing) flow.spacing = spacing;
  if (node.attrs['pageBreakBefore'] === true) flow.pageBreakBefore = true;
  if (node.attrs['keepNext'] === true) flow.keepNext = true;
  if (node.attrs['keepLines'] === true) flow.keepLines = true;
  if (node.attrs['widowControl'] === false) flow.widowControl = false;
  const pBorders = node.attrs['borders'] as ParagraphBorders | null;
  if (pBorders) flow.borders = pBorders;
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

/** The live marker for a list paragraph: counted from the doc's numbering
 *  defs, falling back to a legacy pre-resolved marker on the attr. Carries
 *  the level's label styling (lvlJc / suff / label rPr) alongside the text. */
function markerFor(
  node: PMNode,
  counter: NumberingCounter | undefined,
): { text: string; style?: MarkerStyle } | undefined {
  const list = node.attrs['list'] as {
    numId: string;
    level: number;
    marker?: string;
  } | null;
  if (!list) return undefined;
  const text =
    (counter?.next(list.numId, list.level) || list.marker) ?? undefined;
  if (text === undefined) return undefined;
  const def = counter?.def(list.numId, list.level);
  if (!def || (!def.jc && !def.suff && !def.rPr)) return { text };
  const style: MarkerStyle = {};
  if (def.jc) style.jc = def.jc;
  if (def.suff) style.suff = def.suff;
  if (def.rPr) {
    const { bold, italic, sizePt, family, color } = def.rPr;
    const font: Partial<FontSpec> = {
      ...(bold !== undefined && { bold }),
      ...(italic !== undefined && { italic }),
      ...(sizePt !== undefined && { sizePt }),
      ...(family !== undefined && { family }),
    };
    if (Object.keys(font).length > 0) style.font = font;
    if (color) style.color = color;
  }
  return { text, style };
}

/** Flatten a block-level node (paragraph or table) into a FlowBlock, or null
 *  for node types we don't model yet. `nodePos` is its absolute PM position. */
function nodeToBlock(
  node: PMNode,
  base: FontSpec,
  nodePos: number,
  allowFloats = false,
  counter?: NumberingCounter,
): FlowBlock | null {
  if (node.type.name === 'paragraph') {
    const m = markerFor(node, counter);
    return paragraphToFlow(node, base, nodePos, allowFloats, m?.text, m?.style);
  }
  if (node.type.name === 'table')
    return tableToFlow(node, base, nodePos, counter);
  return null;
}

/** Does this paragraph node carry an inline column break? Checked on the NODE
 *  rather than the flattened flow so a cached paragraph never has to be
 *  re-flattened just to answer it. */
function paragraphHasColumnBreak(node: PMNode): boolean {
  let found = false;
  node.forEach((child) => {
    if (child.type.name === 'column_break') found = true;
  });
  return found;
}

/** Flatten a table node (table → table_row → table_cell → block+). */
function tableToFlow(
  node: PMNode,
  base: FontSpec,
  nodePos: number,
  counter?: NumberingCounter,
): FlowTable {
  const rows: FlowTableRow[] = [];
  node.forEach((rowNode, rowOffset) => {
    const rowPos = nodePos + 1 + rowOffset;
    const cells: FlowTableCell[] = [];
    rowNode.forEach((cellNode, cellOffset) => {
      const cellPos = rowPos + 1 + cellOffset;
      const content: FlowBlock[] = [];
      cellNode.forEach((child, childOffset) => {
        // Cells keep anchored floats: layoutFlow positions them inside the
        // cell box (painted at their offsets; text doesn't wrap around them).
        const block = nodeToBlock(
          child,
          base,
          cellPos + 1 + childOffset,
          true,
          counter,
        );
        if (block) content.push(block);
      });
      const a = cellNode.attrs;
      cells.push({
        colspan: Number(a['colspan']) || 1,
        rowspan: Number(a['rowspan']) || 1,
        colwidth: (a['colwidth'] as number[] | null) ?? null,
        background: (a['background'] as string | null) ?? undefined,
        vAlign: (a['vAlign'] as 'center' | 'bottom' | null) ?? undefined,
        borders: (a['borders'] as TableBorders | null) ?? undefined,
        diagonals: (a['diagonals'] as CellDiagonals | null) ?? undefined,
        padding: (a['padding'] as CellPadding | null) ?? undefined,
        content,
      });
    });
    const row: FlowTableRow = { cells };
    if (rowNode.attrs['header'] === true) row.header = true;
    if (rowNode.attrs['cantSplit'] === true) row.cantSplit = true;
    const height = rowNode.attrs['height'] as {
      value: number;
      exact: boolean;
    } | null;
    if (height) row.height = height;
    rows.push(row);
  });
  const flow: FlowTable = { type: 'table', rows };
  const cellPadding = node.attrs['cellPadding'] as CellPadding | null;
  if (cellPadding) flow.cellPadding = cellPadding;
  const align = node.attrs['align'] as 'center' | 'right' | null;
  if (align) flow.align = align;
  const indent = node.attrs['indent'] as number | null;
  if (indent != null) flow.indent = indent;
  const borders = node.attrs['borders'] as TableBorders | null;
  if (borders) flow.borders = borders;
  return flow;
}

/** Flatten a ProseMirror doc into FlowBlocks (paragraphs + tables). */
export function toFlowBlocks(
  doc: PMNode,
  defaultFont: Partial<FontSpec> = {},
  allowFloats = true,
): FlowBlock[] {
  const base: FontSpec = { ...DEFAULT_FONT, ...defaultFont };
  const counter = createNumberingCounter(
    doc.attrs['numbering'] as NumberingDefs | null,
  );
  const blocks: FlowBlock[] = [];
  doc.forEach((node, offset) => {
    const block = nodeToBlock(node, base, offset, allowFloats, counter);
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
  dstrike?: boolean;
  background?: string;
  vertAlign?: 'super' | 'sub';
  raise?: number;
  footnoteRef?: number;
  commentIds?: number[];
  width: number;
  isSpace: boolean;
  /** A tab character: its width is resolved to the next tab stop at layout. */
  isTab?: boolean;
  /** A forced line break (w:br): flushes the current line. */
  isBreak?: boolean;
  /** A column break (w:br w:type="column"): flushes the line AND ends the
   *  column. Unlike a line break it adds no empty line of its own. */
  isColumnBreak?: boolean;
  /** Original PM position of a tab token (pos is stripped when it becomes a
   *  leader decoration, and must be restorable on a re-resolve after a wrap). */
  origPos?: number;
  /** Absolute PM position of the token's first character / atom. */
  pos?: number;
  /** Size in PM positions (text length, or 1 for an image atom). */
  size: number;
}

/** Small-caps glyph scale (the synthesized lowercase→uppercase size). */
const SMALLCAPS_SCALE = 0.8;

/** Expand a small-caps run into case-spans: lowercase spans become their
 *  uppercase at a reduced size, everything else passes through unchanged.
 *  Per-char mapping keeps the 1:1 text↔PM-position correspondence (a char
 *  whose uppercase form isn't a single char — ß→SS — stays as-is). */
function expandSmallCaps(run: InlineRun): InlineRun[] {
  const isLower = (ch: string): boolean => {
    const up = ch.toUpperCase();
    return up !== ch && up.length === 1;
  };
  const out: InlineRun[] = [];
  const text = run.text;
  let i = 0;
  while (i < text.length) {
    const lower = isLower(text[i]);
    let j = i + 1;
    while (j < text.length && isLower(text[j]) === lower) j++;
    const slice = text.slice(i, j);
    out.push({
      ...run,
      smallCaps: undefined,
      text: lower ? slice.toUpperCase() : slice,
      font: lower
        ? { ...run.font, sizePt: run.font.sizePt * SMALLCAPS_SCALE }
        : run.font,
      pos: run.pos != null ? run.pos + i : undefined,
    });
    i = j;
  }
  return out;
}

/** Tokenize one inline item: words / spaces / tabs for text, a single atom for
 *  images. Tab widths are placeholders, resolved against tab stops at layout. */
function tokenizeInline(inline: FlowInline, ctx: Ctx): Token[] {
  if ('columnBreak' in inline) {
    return [
      {
        isColumnBreak: true,
        font: ctx.base,
        width: 0,
        isSpace: false,
        pos: inline.pos,
        size: 1,
      },
    ];
  }
  if ('break' in inline) {
    return [
      {
        isBreak: true,
        font: ctx.base,
        width: 0,
        isSpace: false,
        pos: inline.pos,
        size: 1,
      },
    ];
  }
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
  // Small caps: split into case-spans first (each re-enters as a plain run).
  if (inline.smallCaps)
    return expandSmallCaps(inline).flatMap((r) => tokenizeInline(r, ctx));
  // Super/subscript render at a reduced size — apply once, here.
  const font = inline.vertAlign
    ? { ...inline.font, sizePt: inline.font.sizePt * SUPERSUB_SCALE }
    : inline.font;
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
        font,
        color: inline.color,
        link: inline.link,
        underline: inline.underline,
        strike: inline.strike,
        dstrike: inline.dstrike,
        background: inline.background,
        vertAlign: inline.vertAlign,
        raise: inline.raise,
        footnoteRef: inline.footnoteRef,
        commentIds: inline.commentIds,
        width: isTab ? 0 : ctx.measure(part, font),
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

/** Materialize a positioned line from a draft. `dx` drops the (column-width)
 *  draft into its column; segments are cloned only when shifting, so cached
 *  drafts are never mutated. */
function draftToLine(d: LineDraft, y: number, dx = 0): LayoutLine {
  const line: LayoutLine = {
    x: d.x + dx,
    y,
    width: d.width,
    height: d.height,
    baseline: d.baseline,
    segments:
      dx === 0 ? d.segments : d.segments.map((s) => ({ ...s, x: s.x + dx })),
  };
  if (d.images.length > 0) {
    line.images =
      dx === 0 ? d.images : d.images.map((im) => ({ ...im, x: im.x + dx }));
  }
  if (d.from != null) line.from = d.from;
  if (d.to != null) line.to = d.to;
  return line;
}

/** The content bounds for the line about to be assembled. Queried once per
 *  line, so callers can flow text around floating images (the band narrows
 *  while a float's rectangle is in the way).
 *
 *  `minWidth` is the width the next UNBREAKABLE item needs. Text never sends
 *  it — text narrows to any band and, alone on a line, breaks at character
 *  level. An inline image can do neither, so a band that passed the MIN_BAND
 *  floor can still be useless to it; the caller should then keep skipping
 *  below the floats (Word's behavior) instead of handing back a band the
 *  image will overflow. Callers with a fixed band (table cells) ignore it. */
type BandFn = (
  estHeight: number,
  /** The paragraph's indents. They apply to the COLUMN — "indentation … on both
   *  the left and the right side of the text margins" (ECMA-376 §17.3.1.12) —
   *  so the band is cut from the indented box, not the other way round. Passing
   *  them in (rather than subtracting from the returned band) is what makes a
   *  float OUTSIDE the box leave it alone, and picks the gap that actually
   *  serves this paragraph when the float is inside it. */
  indents: { left: number; right: number },
  minWidth?: number,
) => {
  left: number;
  right: number;
  /** The text column the band was cut from, indents NOT applied. Tab stops are
   *  measured from here. Returned with the band so the two cannot desync — the
   *  banded flow can switch columns in the middle of a paragraph. */
  column: { left: number; right: number };
};

/** Wrap one paragraph, emitting one LineDraft per line. The band may differ
 *  per line; indents, the list marker and tab stops apply within each band. */
function wrapParagraph(
  block: FlowParagraph,
  ctx: Ctx,
  bandFn: BandFn,
  emit: (d: LineDraft) => void,
  // Resume a paragraph split by a page boundary at this PM position (a line
  // start): earlier tokens are skipped and the first emitted line is a
  // CONTINUATION — no marker, no first-line indent. This is what makes a
  // page that opens mid-way through a float-anchoring paragraph replayable.
  resumeFrom?: number,
  /** Called when a `w:br w:type="column"` is reached: the placer ends the
   *  current column. The next `bandFn` call then answers for the NEW column,
   *  which is why the break has to be handed over here rather than flagged on
   *  a line — by the time a line is emitted its band is already baked in.
   *  Absent where columns cannot exist (table cells), where Word's own
   *  behaviour is to ignore the break. */
  onColumnBreak?: () => void,
): void {
  const { base, measure, metrics, tabWidth } = ctx;
  const indent = block.indent;
  const indentLeft = indent?.left ?? 0;
  const indentRight = indent?.right ?? 0;
  // hanging outdents the first line; firstLine indents it. Mutually exclusive.
  const firstLineDelta =
    indent?.hanging != null ? -indent.hanging : (indent?.firstLine ?? 0);
  const align: Align = block.align ?? 'left';

  let tokens = block.runs.flatMap((inline) => tokenizeInline(inline, ctx));
  if (resumeFrom !== undefined) {
    // Line boundaries are token boundaries, so the cut is exact.
    tokens = tokens.filter((t) => t.pos !== undefined && t.pos >= resumeFrom);
  }

  // A line is as tall as the tallest glyph on it — and the paragraph MARK is
  // a glyph (§17.3.1.29, "a physical character in the document") that sits on
  // the LAST line. So the last line is seeded with the mark's metrics: a mark
  // larger than the text opens that line (Word: "a larger font character,
  // including the paragraph mark, in the last line increases the spacing"),
  // a mark no larger leaves the text's own height alone, and a paragraph
  // with NO runs is exactly one mark-tall line. Every OTHER line has no seed:
  // only its runs decide. Seeding every line from the document default (the
  // old rule) made 8pt text inside 8pt-mark paragraphs 11pt tall, and a rate
  // card's exact-height rows spilled every second line over the row below.
  //
  // The mark's font falls back to the document base when the importer did
  // not resolve one (hand-built flows, older documents).
  const markOnly = block.runs.length === 0;
  const seedFont = block.markFont ? { ...base, ...block.markFont } : base;
  // A mark-only paragraph whose mark IS a section break draws no line box at
  // all — the break is a formatting mark, and the next section starts flush
  // against the last real line (see FlowParagraph.breakMark).
  const baseMetrics =
    markOnly && block.breakMark
      ? { ascent: 0, descent: 0, leading: 0 }
      : metrics
        ? metrics(seedFont)
        : null;
  const nominalH = baseMetrics
    ? baseMetrics.ascent + baseMetrics.descent + (baseMetrics.leading ?? 0)
    : sizePx(seedFont) * LINE_HEIGHT_FACTOR;

  // Bounds for the line currently being assembled.
  //
  // The paragraph's own box comes from the COLUMN — "indentation … on both the
  // left and the right side of the text margins" (ECMA-376 §17.3.1.12). A
  // float's band then CLIPS that box; it must not be the thing the indents are
  // measured from, or a paragraph with a right indent beside a float has the
  // indent subtracted twice. One factsheet in the corpus lands exactly there:
  // column 785px, indents 21 + 402, float band [9, 409] — Word draws a 362px
  // line, subtracting twice gives −23px and one character per line.
  const indents = { left: indentLeft, right: indentRight };
  let band = bandFn(nominalH, indents);
  let lineLeft = band.left;
  let lineRight = band.right;

  // List marker hangs at the first line's start; text follows after it, and
  // wrapped lines align under that text (hanging indent). The label draws
  // with its own font/color (lvl rPr), aligns against the anchor (lvlJc) and
  // separates from the text per w:suff — tab (default) jumps to the hanging
  // text position like Word, space/nothing stay tight.
  const mStyle = block.markerStyle;
  const markerFont: FontSpec = mStyle?.font
    ? { ...base, ...mStyle.font }
    : base;
  let marker: LayoutSegment | null = null;
  let markerTextX = 0;
  const placeMarker = (): void => {
    if (!marker) return;
    const anchor = lineLeft + firstLineDelta;
    const w = marker.width ?? 0;
    marker.x =
      mStyle?.jc === 'right'
        ? anchor - w
        : mStyle?.jc === 'center'
          ? anchor - w / 2
          : anchor;
    const end = marker.x + w;
    markerTextX =
      mStyle?.suff === 'nothing'
        ? end
        : mStyle?.suff === 'space'
          ? end + measure(' ', markerFont)
          : end <= lineLeft
            ? lineLeft
            : end + measure(' ', markerFont);
  };
  if (block.marker) {
    marker = {
      x: 0,
      text: block.marker,
      font: markerFont,
      width: measure(block.marker, markerFont),
      ...(mStyle?.color ? { color: mStyle.color } : {}),
    };
    placeMarker();
  }
  // `let`: re-derived if the first line's band is re-queried for a wide image.
  let firstLineStart = marker ? markerTextX : lineLeft + firstLineDelta;
  let contStart = marker ? Math.max(markerTextX, lineLeft) : lineLeft;

  let lineTokens: Token[] = [];
  let lineWidth = 0; // running width of the current line's tokens
  let firstLine = resumeFrom === undefined; // a resumed line is a continuation
  let prevTo: number | undefined; // caret slot after the previous line's content

  // Zero for a section-break mark, matching baseMetrics above.
  const seedPx = markOnly && block.breakMark ? 0 : sizePx(seedFont);
  // Per-line maxima start EMPTY — the mark joins only the last line, in
  // flushLine (see foldMark).
  let maxFontPx = 0; // tallest text (fallback line-height mode)
  let maxImagePx = 0; // tallest inline image on the line
  let maxAscent = 0; // metrics mode
  let maxDescent = 0;
  // The font's own gap between lines. Tracked as its own maximum rather than
  // folded into the descent so `exact`/`atLeast` can reason about the cell and
  // the gap separately, and so the baseline stays at the ascent.
  let maxLeading = 0;
  // Text-only ascent/descent, excluding image contributions: the w:line
  // 'auto' multiple scales the TEXT box, never an image (Word semantics) —
  // so the spacing code needs the text height by itself.
  let textAscent = 0;
  let textDescent = 0;
  let textLeading = 0;
  /** The paragraph mark takes its place on the line being flushed: it is a
   *  text glyph, so it raises both the line maxima and the text-only ones. */
  const foldMark = (): void => {
    maxFontPx = Math.max(maxFontPx, seedPx);
    if (baseMetrics) {
      maxAscent = Math.max(maxAscent, baseMetrics.ascent);
      maxDescent = Math.max(maxDescent, baseMetrics.descent);
      maxLeading = Math.max(maxLeading, baseMetrics.leading ?? 0);
      textAscent = Math.max(textAscent, baseMetrics.ascent);
      textDescent = Math.max(textDescent, baseMetrics.descent);
      textLeading = Math.max(textLeading, baseMetrics.leading ?? 0);
    }
  };

  const lineStart = () => (firstLine ? firstLineStart : contStart);

  /**
   * Distance from `x` to the next default tab stop.
   *
   * The grid runs from the COLUMN, not from the line's left edge: a tab stop
   * is a position on the ruler, and the ruler starts at the text margin. The
   * paragraph's own indent must not shift it, and neither may a float that
   * narrowed this line. See `resolveTab` for the evidence.
   */
  const tabAdvance = (x: number) => {
    const origin = band.column.left;
    const k = Math.floor((x - origin) / tabWidth) + 1;
    return origin + k * tabWidth - x;
  };

  // ── Custom tab stops (w:tabs) ─────────────────────────────────────
  const tabStops = block.tabs
    ? [...block.tabs].sort((a, b) => a.pos - b.pos)
    : [];
  const LEADER_CHARS = {
    dot: '.',
    hyphen: '-',
    underscore: '_',
    middleDot: '·',
  } as const;
  const MIN_TAB = 2;

  /** The "tab group": tokens after `ti` up to the next tab, whose total width
   *  decides where right/center/decimal-aligned text starts. `decimalPrefix`
   *  is the width before the first '.' or ',' in the group. */
  const tabGroup = (ti: number) => {
    let width = 0;
    let decimalPrefix: number | null = null;
    for (let j = ti + 1; j < tokens.length; j++) {
      const t = tokens[j];
      if (t.isTab) break;
      if (decimalPrefix === null && t.text != null && !t.field) {
        const m = /[.,]/.exec(t.text);
        if (m)
          decimalPrefix = width + measure(t.text.slice(0, m.index), t.font);
      }
      width += t.width;
    }
    return { width, decimalPrefix };
  };

  /**
   * Resolve a tab token at `x`: jump to the next custom stop (aligning the
   * following group for right/center/decimal, synthesizing the leader fill),
   * or fall back to the default grid past the last stop.
   *
   * `w:tab/@w:pos` is measured from the text margin, NOT from the paragraph's
   * indent. A hotel factsheet in the corpus proves it: one paragraph sets
   * `ind left=321tw` with a stop at `pos=1761tw`, and the continuation
   * paragraph below it sets `left=321 + firstLine=1440` — also 1761tw from
   * the margin. Word renders the two flush with each other, which only holds
   * if the stop ignores the 321. Adding the indent (as this did) pushed the
   * second column 21px right of the line beneath it.
   */
  const resolveTab = (token: Token, x: number, ti: number) => {
    // Re-resolves (after a wrap) must start from a pristine tab token.
    token.origPos ??= token.pos;
    token.pos = token.origPos;
    token.text = '\t';

    const origin = band.column.left;
    const stop = tabStops.find((s) => origin + s.pos > x + 0.5);
    if (!stop) {
      token.width = tabAdvance(x);
      return;
    }
    // A stop past the line end clamps to it (Word: TOC stops at the margin).
    const stopX = Math.min(origin + stop.pos, lineRight);
    if (stopX <= x + 0.5) {
      token.width = tabAdvance(x);
      return;
    }
    let w = stopX - x;
    if (stop.val !== 'left') {
      const g = tabGroup(ti);
      const hang =
        stop.val === 'right'
          ? g.width
          : stop.val === 'center'
            ? g.width / 2
            : (g.decimalPrefix ?? g.width); // decimal: align the separator
      w = stopX - x - hang;
      if (w < MIN_TAB) {
        token.width = tabAdvance(x); // group doesn't fit before the stop
        return;
      }
    }
    token.width = w;
    if (stop.leader) {
      const ch = LEADER_CHARS[stop.leader] ?? '.';
      const chW = measure(ch, token.font);
      const n = chW > 0 ? Math.floor(w / chW) : 0;
      if (n > 1) {
        token.text = ch.repeat(n - 1); // a breath of space before the target
        token.pos = undefined; // decoration — not caret-addressable
      }
    }
  };

  const flushLine = (isLast: boolean, softBreak = false) => {
    const startX = lineStart();
    const avail = lineRight - startX;

    // Trailing whitespace doesn't count toward alignment, nor is it painted —
    // Word lets it hang past the line's end. With w:ulTrailSpace the
    // underlined part of it is still painted (an underline runs under the
    // hanging spaces); it hangs all the same, so it never moves the text.
    let end = lineTokens.length;
    let contentWidth = lineWidth;
    while (end > 0 && lineTokens[end - 1].isSpace) {
      contentWidth -= lineTokens[end - 1].width;
      end--;
    }
    let paintEnd = end;
    if (ctx.compat.underlineTrailingSpaces) {
      while (
        paintEnd < lineTokens.length &&
        lineTokens[paintEnd].isSpace &&
        !lineTokens[paintEnd].isTab &&
        lineTokens[paintEnd].underline
      )
        paintEnd++;
    }

    let x = startX;
    let extraPerGap = 0;
    // A line ending in a soft break (w:br) is stretched like any other
    // non-final line — unless w:doNotExpandShiftReturn asks for it to be set
    // ragged, like the last line.
    const ragged =
      isLast || (softBreak && !ctx.compat.expandLineBeforeSoftBreak);
    if (align === 'justify' && !ragged) {
      const gaps = lineTokens
        .slice(0, end)
        .filter((t) => t.isSpace && !t.isTab).length;
      if (gaps > 0) extraPerGap = (avail - contentWidth) / gaps;
    } else if (align === 'center') {
      x += Math.max(0, (avail - contentWidth) / 2);
    } else if (align === 'right') {
      x += Math.max(0, avail - contentWidth);
    }

    const segments: LayoutSegment[] = [];
    const images: LayoutImageSegment[] = [];
    for (let i = 0; i < paintEnd; i++) {
      const t = lineTokens[i];
      if (t.image) {
        images.push({
          x,
          src: t.image.src,
          width: t.image.width,
          height: t.image.height,
          link: t.link,
          ...(t.image.shape ? { shape: t.image.shape } : {}),
          ...(t.image.crop ? { crop: t.image.crop } : {}),
          ...(t.image.outline ? { outline: t.image.outline } : {}),
          ...(t.image.rotation ? { rotation: t.image.rotation } : {}),
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
          background: t.background,
          vertAlign: t.vertAlign,
          raise: t.raise,
          width: t.width,
          pos: t.pos,
        };
        if (t.dstrike) seg.dstrike = true;
        if (t.field) seg.field = t.field;
        if (t.footnoteRef != null) seg.footnoteRef = t.footnoteRef;
        if (t.commentIds) seg.commentIds = t.commentIds;
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

    // The paragraph mark sits on the last line and counts like any glyph.
    if (isLast) foldMark();

    let height: number;
    let baseline: number;
    if (metrics) {
      // Real metrics: the line box is the font's cell PLUS its external
      // leading, which is what Word means by one line — "the cell height plus
      // the external leading is equal to the line spacing" (GDI TEXTMETRIC),
      // and the value LibreOffice's AddExternalLeading compatibility flag adds
      // for the same reason. Without it every line is short by the font's gap:
      // 4.2% of the em in Times New Roman, and the error compounds down a page.
      //
      // The gap sits BELOW the baseline (baseline stays at the ascent) because
      // that is what "space between rows" means. Unverified against Word:
      // measurement pins the line PITCH, not where the baseline sits inside it.
      height = maxAscent + maxDescent + maxLeading;
      baseline = maxAscent;
    } else {
      // Fallback: line box grows to the tallest image, else a font-size factor.
      height = Math.max(maxFontPx * LINE_HEIGHT_FACTOR, maxImagePx);
      baseline = height * BASELINE_FACTOR;
    }
    // w:spacing/w:line: 'auto' multiplies the natural height (extra space split
    // below the baseline), 'exact' forces it, 'atLeast' is a floor.
    //
    // 'auto' scales the TEXT box only — never an image. Word and Google Docs
    // size an image line to the image (it sits on the baseline); multiplying
    // the image height inflated every picture line by the paragraph's line
    // factor (~8% of a 300px screenshot is ~25px), and the accumulated slack
    // paginated documents earlier than Word. The image keeps its natural box;
    // the multiple only wins when the scaled text is taller.
    const sp = block.spacing;
    if (sp?.line) {
      // "One line" for the auto multiple is the whole line spacing, leading
      // included: w:line is "240ths of a single line height", so 360 = 1.5 of
      // THIS. Measured against a document that positions floating shapes off
      // paragraph tops, a 13pt Times line at 1.5 has to come to 22.40pt, and
      // (ascent + descent + leading) × 1.5 gives 22.42 where the cell alone
      // gives 21.59.
      const textH = metrics
        ? textAscent + textDescent + textLeading
        : maxFontPx * LINE_HEIGHT_FACTOR;
      const target =
        sp.lineRule === 'exact'
          ? sp.line
          : sp.lineRule === 'atLeast'
            ? Math.max(height, sp.line)
            : maxImagePx > 0
              ? Math.max(height, textH * sp.line)
              : textH * sp.line;
      baseline +=
        Math.max(0, target - height) *
        (sp.lineRule === 'exact' ? baseline / height : 1);
      height = target;
    }
    const painted = firstLine && marker ? [marker, ...segments] : segments;
    const draft: LineDraft = {
      x: startX,
      width: lineRight - startX,
      height,
      baseline,
      segments: painted,
      images,
      from,
      to,
    };
    emit(draft);
    prevTo = to;

    lineTokens = [];
    lineWidth = 0;
    maxFontPx = 0;
    maxImagePx = 0;
    maxAscent = 0;
    maxDescent = 0;
    maxLeading = 0;
    textAscent = 0;
    textDescent = 0;
    textLeading = 0;
    firstLine = false;

    // The next line may sit beside (or past) a float — its band is fetched
    // LAZILY, at the next token. Fetching it here asked the band callback
    // for a line that need not exist: after the paragraph's LAST line the
    // request walked `y` past floats and could even close the page — a
    // page break decided by a phantom line. (One factsheet lost its section
    // mark to a fresh page exactly that way.)
    bandStale = true;
  };
  let bandStale = false;

  // Kerning across token boundaries: consecutive same-font text tokens (a word
  // split across runs, e.g. by a mark change) are measured cumulatively, so
  // the advance matches measuring the joined text. Resets on anything that
  // breaks the glyph run: spaces, tabs, images, fields, font changes, wraps.
  let clusterText = '';
  let clusterWidth = 0;
  let clusterFont: FontSpec | null = null;
  const resetCluster = () => {
    clusterText = '';
    clusterWidth = 0;
    clusterFont = null;
  };
  const clusterable = (t: Token) =>
    t.text != null && !t.isSpace && !t.isTab && !t.field && !t.image;

  // Spaces at the start of a soft-wrapped continuation line are suppressed
  // (they belong to the previous line's end). But typed leading spaces — the
  // paragraph's first line, or right after a hard break — DO render; Word
  // keeps them, and real documents position text with them.
  let softWrapped = false;

  for (let ti = 0; ti < tokens.length; ti++) {
    if (bandStale) {
      band = bandFn(nominalH, indents);
      lineLeft = band.left;
      lineRight = band.right;
      contStart = marker ? Math.max(markerTextX, lineLeft) : lineLeft;
      bandStale = false;
    }
    const token = tokens[ti];
    // A column break ends the line WITHOUT adding one of its own — unlike a
    // line break, which leaves an empty line behind when it opens a
    // paragraph. What follows resumes in the next column.
    if (token.isColumnBreak) {
      if (token.pos != null) prevTo = token.pos + 1;
      if (lineTokens.length > 0) flushLine(false);
      resetCluster();
      softWrapped = false;
      onColumnBreak?.();
      // Re-ask for the band: it is the new column's now. Both line starts
      // have to follow — a break before the paragraph's first line leaves
      // `firstLine` true, and that reads firstLineStart, not contStart.
      band = bandFn(nominalH, indents);
      lineLeft = band.left;
      lineRight = band.right;
      placeMarker();
      firstLineStart = marker ? markerTextX : lineLeft + firstLineDelta;
      contStart = marker ? Math.max(markerTextX, lineLeft) : lineLeft;
      continue;
    }
    // A forced break (w:br) ends the current line; its PM position is the slot
    // after the line so the caret can sit on it.
    if (token.isBreak) {
      if (token.pos != null) prevTo = token.pos + 1;
      flushLine(false, true);
      resetCluster();
      softWrapped = false;
      continue;
    }
    // Skip leading spaces on wrapped lines (a leading tab always indents).
    if (token.isSpace && !token.isTab && lineTokens.length === 0 && softWrapped)
      continue;
    if (token.isTab) resolveTab(token, lineStart() + lineWidth, ti);
    if (
      clusterable(token) &&
      clusterFont &&
      sameGlyphRun(clusterFont, token.font)
    ) {
      token.width = Math.max(
        0,
        measure(clusterText + token.text, token.font) - clusterWidth,
      );
    }
    const cursor = lineStart() + lineWidth;
    // Spaces hang past the right edge rather than wrapping — but a TAB does
    // not: it wraps to the next line and resolves against the stops there.
    // Measured in the corpus factsheet's PDF: a paragraph ending in a bare
    // w:tab whose next stop (537) lies past the right indent (532.5) gets a
    // second line holding just the tab, 16.8px tall — and the three empty
    // paragraphs after it sit exactly one line pitch lower than an engine
    // that swallows the tab puts them.
    if (
      (!token.isSpace || token.isTab) &&
      lineTokens.length > 0 &&
      cursor + token.width > lineRight
    ) {
      flushLine(false);
      resetCluster(); // the wrapped token starts a fresh glyph run
      softWrapped = true;
      if (token.isTab) resolveTab(token, lineStart() + lineWidth, ti);
      else if (clusterable(token))
        token.width = measure(token.text as string, token.font);
    }
    // An inline image that doesn't fit the band, alone on its line: it cannot
    // break like text, so a float-narrowed band the MIN_BAND floor accepted
    // can still be useless — placing it there overflows straight into the
    // float's rectangle. Ask for a band wide enough; the caller skips below
    // the floats to find one (moving its y), so refresh every band-derived
    // coordinate. If the re-queried band still can't fit it (image wider than
    // the column, or a fixed cell band), keep today's overflow.
    if (
      token.image &&
      lineTokens.length === 0 &&
      lineStart() + token.width > lineRight
    ) {
      // Everything left of the content start (indent, first-line delta, list
      // marker) plus the image plus the right indent — the exact width the
      // band must offer. The left-side geometry is re-derived below relative
      // to the new band, so the relation carries over.
      band = bandFn(
        nominalH,
        indents,
        // The band already has the indents taken out, so the width to ask for
        // is what sits left of the content start plus the image itself.
        lineStart() - band.left + token.width,
      );
      lineLeft = band.left;
      lineRight = band.right;
      placeMarker();
      firstLineStart = marker ? markerTextX : lineLeft + firstLineDelta;
      contStart = marker ? Math.max(markerTextX, lineLeft) : lineLeft;
    }

    // A single word wider than the whole band (narrow table cells): break it
    // at character level, Word-style — fit what we can (at least one char),
    // the remainder re-enters the loop as its own token on the next line.
    if (
      clusterable(token) &&
      lineTokens.length === 0 &&
      token.text!.length > 1 &&
      lineStart() + token.width > lineRight
    ) {
      const avail = lineRight - lineStart();
      let n = 1;
      while (
        n < token.text!.length &&
        measure(token.text!.slice(0, n + 1), token.font) <= avail
      )
        n++;
      const restText = token.text!.slice(n);
      const rest: Token = {
        ...token,
        text: restText,
        width: measure(restText, token.font),
        size: restText.length,
      };
      if (token.pos != null) rest.pos = token.pos + n;
      token.text = token.text!.slice(0, n);
      token.size = n;
      token.width = measure(token.text, token.font);
      tokens.splice(ti + 1, 0, rest);
    }
    if (clusterable(token)) {
      if (clusterFont && !sameGlyphRun(clusterFont, token.font)) resetCluster();
      clusterText += token.text;
      clusterWidth += token.width;
      clusterFont = token.font;
    } else {
      resetCluster();
    }
    lineTokens.push(token);
    lineWidth += token.width;
    if (token.image) {
      // A rotated image paints around the center of its baseline-anchored box,
      // so the line must hold the rotated box's HEIGHT — above and below the
      // baseline. Width is deliberately not grown: horizontal overflow clips
      // (Word keeps the column grid; the picture just pokes out).
      const rotH = rotatedBoxHeight(token.image);
      maxImagePx = Math.max(maxImagePx, rotH);
      maxAscent = Math.max(maxAscent, token.image.height / 2 + rotH / 2);
      maxDescent = Math.max(
        maxDescent,
        Math.max(0, (rotH - token.image.height) / 2),
      );
    } else {
      maxFontPx = Math.max(maxFontPx, sizePx(token.font));
      if (metrics) {
        const m = metrics(token.font);
        maxAscent = Math.max(maxAscent, m.ascent);
        maxDescent = Math.max(maxDescent, m.descent);
        maxLeading = Math.max(maxLeading, m.leading ?? 0);
        textAscent = Math.max(textAscent, m.ascent);
        textDescent = Math.max(textDescent, m.descent);
        textLeading = Math.max(textLeading, m.leading ?? 0);
      }
    }
  }
  flushLine(true); // emit the paragraph's last (or only/empty) line
}

/** Height of the axis-aligned box containing an image after its paint-only
 *  rotation (identity when unrotated). */
function rotatedBoxHeight(img: {
  width: number;
  height: number;
  rotation?: number;
}): number {
  if (!img.rotation) return img.height;
  const rad = (img.rotation * Math.PI) / 180;
  return (
    Math.abs(img.width * Math.sin(rad)) + Math.abs(img.height * Math.cos(rad))
  );
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
  wrapParagraph(
    block,
    ctx,
    (_h, ind) => ({
      left: contentLeft + ind.left,
      right: contentRight - ind.right,
      column: { left: contentLeft, right: contentRight },
    }),
    (d) => drafts.push(d),
  );
  return drafts;
}

/** Shift a resolved table (and everything inside it) down by `dy` px. */
function offsetTable(table: ResolvedTable, dy: number): void {
  table.y += dy;
  for (const cell of table.cells) {
    cell.y += dy;
    for (const line of cell.lines) line.y += dy;
    if (cell.floats) for (const f of cell.floats) f.y += dy;
    if (cell.tables) for (const t of cell.tables) offsetTable(t, dy);
  }
}

/** Shift a laid-out line (box + segments + images) horizontally in place. */
function shiftLineX(line: LayoutLine, dx: number): void {
  line.x += dx;
  for (const seg of line.segments) seg.x += dx;
  if (line.images) for (const im of line.images) im.x += dx;
}

/** Shift a resolved table (and everything inside) horizontally by `dx` px —
 *  drops a column-width table into its column. Safe to mutate: tables are laid
 *  out fresh each pass (never cached). */
function shiftTableX(table: ResolvedTable, dx: number): void {
  if (dx === 0) return;
  table.x += dx;
  for (const cell of table.cells) {
    cell.x += dx;
    for (const line of cell.lines) shiftLineX(line, dx);
    if (cell.floats) for (const f of cell.floats) f.x += dx;
    if (cell.tables) for (const t of cell.tables) shiftTableX(t, dx);
  }
}

/** Word's default textbox interior padding (bodyPr lIns 0.1", tIns 0.05"). */
const TEXTBOX_INSET = { l: 10, t: 5, r: 10, b: 5 };

/** ResolvedFloat for `f` pinned at (x, y). A textbox's paragraphs are flowed
 *  inside the shape's box, in box-local coordinates (the painter translates
 *  by the float's origin) — never caret-addressable, positions stripped. */
function resolveFloat(
  f: FlowFloat,
  x: number,
  y: number,
  ctx: Ctx,
): ResolvedFloat {
  const rf: ResolvedFloat = {
    x,
    y,
    width: f.width,
    height: f.height,
    src: f.src,
  };
  if (f.pos != null) rf.pos = f.pos;
  if (f.rotation) rf.rotation = f.rotation;
  if (f.crop) rf.crop = f.crop;
  if (f.outline) rf.outline = f.outline;
  if (f.shape) rf.shape = f.shape;
  if (f.behind) rf.behind = true;
  if (f.content && f.content.length > 0) {
    const inset = f.inset ?? TEXTBOX_INSET;
    const right = Math.max(inset.l + MIN_BAND, f.width - inset.r);
    const inner = layoutFlow(f.content, inset.l, right, ctx);
    // wps:bodyPr @anchor — the whole text block slides down the leftover
    // height. Word clamps at the top when the text overflows the box, which
    // is what a negative slack would otherwise undo.
    const slack = Math.max(0, f.height - inset.t - inset.b - inner.height);
    const shift = f.anchor === 'ctr' ? slack / 2 : f.anchor === 'b' ? slack : 0;
    const lines = inner.lines.map((l) => ({ ...l, y: l.y + inset.t + shift }));
    // Tables get the same drop as the lines. They are mutated in place rather
    // than copied because layoutFlow builds them fresh on every pass.
    for (const t of inner.tables) offsetTable(t, inset.t + shift);
    stripPositions(lines, inner.tables);
    if (lines.length > 0) rf.lines = lines;
    if (inner.tables.length > 0) rf.tables = inner.tables;
  }
  return rf;
}

/** Lay out a sequence of blocks within a content box, stacking vertically from
 *  y = 0. No pagination — used for table-cell content. Anchored floats are
 *  positioned at their offsets within the box (v1: painted only — the cell's
 *  text does not wrap around them, which matches how Word renders the small
 *  wrapThrough shapes real documents drop into table cells). */
function layoutFlow(
  blocks: FlowBlock[],
  contentLeft: number,
  contentRight: number,
  ctx: Ctx,
): {
  lines: LayoutLine[];
  tables: ResolvedTable[];
  floats: ResolvedFloat[];
  height: number;
} {
  const lines: LayoutLine[] = [];
  const tables: ResolvedTable[] = [];
  const floats: ResolvedFloat[] = [];
  let y = 0;
  // Space-after already added by the previous block, so the next block's
  // space-before collapses against it (see collapsedBefore).
  let pendingAfter = 0;
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      // w:spacing before/after applies inside cells exactly as in the body
      // flow (placeBlocks handles it there) — it was silently swallowed here,
      // so a styled heading lost its breathing room the moment it sat in a
      // table. The float anchor moves with it: the paragraph's top IS below
      // the gap.
      y += collapsedBefore(block.spacing?.before, pendingAfter);
      // ONE keeps object per paragraph, shared by its lines — the reference
      // identity is what lets the row splitter regroup a paragraph's lines
      // and judge split legality per fragment.
      const keeps: ParagraphKeeps = {};
      if (block.keepNext) keeps.keepNext = true;
      if (block.keepLines) keeps.keepLines = true;
      if (block.widowControl === false) keeps.widowControl = false;
      for (const f of block.floats ?? []) {
        // Horizontal: alignment within the box, else offset from its left.
        // (hRel 'page' has no page here — the box is the closest analogue.)
        const fx =
          f.hAlign === 'right'
            ? contentRight - f.width
            : f.hAlign === 'center'
              ? (contentLeft + contentRight - f.width) / 2
              : contentLeft + (f.hOffset ?? 0);
        // Vertical: relative to the anchor paragraph's top (margin/page
        // degrade to the same — a cell has no margin band of its own).
        const fy = y + (f.vOffset ?? 0);
        floats.push({
          ...resolveFloat(f, fx, fy, ctx),
          // Effective offsets for drag-to-move (translation-invariant — the
          // cell-float shift sites need no bookkeeping).
          effHOffset: fx - contentLeft,
          effVOffset: f.vOffset ?? 0,
        });
      }
      for (const d of layoutParagraph(block, contentLeft, contentRight, ctx)) {
        const line = draftToLine(d, y);
        line.keeps = keeps;
        lines.push(line);
        y += d.height;
      }
      y += block.spacing?.after ?? 0;
      pendingAfter = block.spacing?.after ?? 0;
    } else {
      const table = layoutTable(block, contentLeft, contentRight, ctx);
      offsetTable(table, y);
      tables.push(table);
      y += table.height;
      pendingAfter = 0; // a table contributes no space-after to collapse into
    }
  }
  return { lines, tables, floats, height: y };
}

/**
 * Walk every cell assigning its starting column, honoring rowspan occupancy
 * from the rows above: a vertically-merged (rowspan>1) cell reserves its
 * column(s) in the rows it spans, so cells below it shift right instead of
 * overlapping it. Without this, a row beneath a vertical merge is laid out one
 * column too far left and the last column comes out empty.
 */
function eachCell(
  rows: FlowTableRow[],
  ncols: number,
  visit: (rowIndex: number, cell: FlowTableCell, startCol: number) => void,
): void {
  const spanned = new Array<number>(ncols).fill(0); // rows still covered per column
  for (let r = 0; r < rows.length; r++) {
    let col = 0;
    for (const cell of rows[r].cells) {
      while (col < ncols && spanned[col] > 0) col++; // skip columns held by a rowspan above
      visit(r, cell, col);
      for (let k = 0; k < cell.colspan && col + k < ncols; k++)
        spanned[col + k] = cell.rowspan;
      col += cell.colspan;
    }
    for (let i = 0; i < ncols; i++) if (spanned[i] > 0) spanned[i]--; // one row consumed
  }
}

/** Inner-gridline space a rowspan of `span` rows starting at `startRow`
 *  swallows (span-1 lines of insideH width, clamped to the table). */
function bwInsideForSpan(
  startRow: number,
  span: number,
  nrows: number,
  insideH: { width: number } | false | undefined,
): number {
  if (!insideH) return 0;
  const rows = Math.min(span, nrows - startRow);
  return rows > 1 ? (rows - 1) * insideH.width : 0;
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
  // Scale an over-wide grid down to `avail`? TRUE inside narrow spaces —
  // a multi-column section, a cell (nested table), chrome — where spilling
  // would paint over a neighbour's content. FALSE in the single-column body
  // flow: Word renders the STORED tblGrid as-is there and lets the table
  // run into the right margin (autofit only recomputes on edit, not open),
  // so shrinking desynced our columns from every anchored object's x.
  clampToWidth = true,
): ResolvedTable {
  const avail = contentRight - contentLeft;
  const nrows = table.rows.length;
  const ncols = table.rows.reduce(
    (m, r) =>
      Math.max(
        m,
        r.cells.reduce((s, c) => s + c.colspan, 0),
      ),
    0,
  );

  // Column widths: take known widths from cells, split the rest equally.
  const colWidths = new Array<number>(ncols).fill(0);
  eachCell(table.rows, ncols, (_r, cell, startCol) => {
    if (cell.colwidth && cell.colwidth.length === cell.colspan) {
      for (let k = 0; k < cell.colspan && startCol + k < ncols; k++) {
        if (colWidths[startCol + k] === 0)
          colWidths[startCol + k] = cell.colwidth[k];
      }
    }
  });
  const known = colWidths.reduce((s, w) => s + w, 0);
  const unknown = colWidths.filter((w) => w === 0).length;
  if (unknown > 0) {
    const share = Math.max(0, (avail - known) / unknown);
    for (let i = 0; i < ncols; i++)
      if (colWidths[i] === 0) colWidths[i] = share;
  }
  // A tblGrid wider than the available width (e.g. a full-page table dropped
  // into a narrow column) would otherwise overflow into the next column's
  // content. Scale every column down proportionally so it fits — but ONLY
  // where the caller asked for it (see the parameter): the single-column
  // body flow honors the grid and overflows the margin, as Word does.
  const natural = colWidths.reduce((s, w) => s + w, 0);
  if (clampToWidth && natural > avail && natural > 0) {
    const scale = avail / natural;
    for (let i = 0; i < ncols; i++) colWidths[i] *= scale;
  }
  const tableWidth = colWidths.reduce((s, w) => s + w, 0);
  // Per-table cell margins (w:tblCellMar) override the Word defaults; a cell
  // may override again with w:tcMar.
  const pad = {
    left: table.cellPadding?.left ?? CELL_PAD_X,
    right: table.cellPadding?.right ?? CELL_PAD_X,
    top: table.cellPadding?.top ?? CELL_PAD_Y,
    bottom: table.cellPadding?.bottom ?? CELL_PAD_Y,
  };
  // Table alignment (w:jc) shifts the whole grid within the content area.
  const xShift =
    table.align === 'center'
      ? Math.max(0, (avail - tableWidth) / 2)
      : table.align === 'right'
        ? Math.max(0, avail - tableWidth)
        : 0;
  // A left-aligned table's leading border sits at w:tblInd (resolved by the
  // importer to a border offset for the document's compat mode); without one,
  // Word's implicit indent applies: the grid shifts LEFT by the left cell
  // margin, so the first cell's CONTENT (grid + padding) lines up with the
  // body text margin — Word letterheads rely on this. A rate card indents its
  // tables 1648 twips to sit under a centred heading; drawn at the margin
  // they were 110px off Word.
  const indent = table.align ? 0 : (table.indent ?? -pad.left);
  const colX = new Array<number>(ncols + 1).fill(contentLeft + indent + xShift);
  for (let i = 0; i < ncols; i++) colX[i + 1] = colX[i] + colWidths[i];

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
    floats: ResolvedFloat[];
    contentHeight: number;
    background?: string;
    vAlign?: 'center' | 'bottom';
    borders?: TableBorders;
    diagonals?: CellDiagonals;
    pad: { left: number; right: number; top: number; bottom: number };
  }

  const cellDrafts: CellDraft[] = [];
  eachCell(table.rows, ncols, (r, cell, col) => {
    let cellWidth = 0;
    for (let k = 0; k < cell.colspan && col + k < ncols; k++)
      cellWidth += colWidths[col + k];
    const cellLeft = colX[col];
    // w:tcMar overrides the table's margins per cell, per side.
    const cellPad = {
      left: cell.padding?.left ?? pad.left,
      right: cell.padding?.right ?? pad.right,
      top: cell.padding?.top ?? pad.top,
      bottom: cell.padding?.bottom ?? pad.bottom,
    };
    const flow = layoutFlow(
      cell.content,
      cellLeft + cellPad.left,
      cellLeft + cellWidth - cellPad.right,
      ctx,
    );
    cellDrafts.push({
      startRow: r,
      startCol: col,
      colspan: cell.colspan,
      rowspan: cell.rowspan,
      cellLeft,
      cellWidth,
      lines: flow.lines,
      tables: flow.tables,
      floats: flow.floats,
      contentHeight: flow.height,
      background: cell.background,
      vAlign: cell.vAlign,
      borders: cell.borders,
      diagonals: cell.diagonals,
      pad: cellPad,
    });
  });

  // Row heights: single-row cells set the base; multi-row cells grow the last
  // spanned row if their content needs more than the rows already provide.
  const rowHeight = new Array<number>(nrows).fill(0);
  for (const c of cellDrafts) {
    if (c.rowspan === 1) {
      rowHeight[c.startRow] = Math.max(
        rowHeight[c.startRow],
        c.contentHeight + c.pad.top + c.pad.bottom,
      );
    }
  }
  for (const c of cellDrafts) {
    if (c.rowspan > 1) {
      const need = c.contentHeight + c.pad.top + c.pad.bottom;
      let span = 0;
      for (let r = c.startRow; r < c.startRow + c.rowspan && r < nrows; r++)
        span += rowHeight[r];
      // A merged cell also spans the inner gridlines between its rows (the
      // border-height pass below reserves them), so count that space too.
      span += bwInsideForSpan(
        c.startRow,
        c.rowspan,
        nrows,
        table.borders?.insideH,
      );
      if (need > span) {
        const last = Math.min(c.startRow + c.rowspan - 1, nrows - 1);
        rowHeight[last] += need - span;
      }
    }
  }
  // w:trHeight: 'exact' forces the row height, otherwise it's a floor.
  for (let r = 0; r < nrows; r++) {
    const h = table.rows[r]?.height;
    if (h) rowHeight[r] = h.exact ? h.value : Math.max(rowHeight[r], h.value);
  }
  // Horizontal border lines OCCUPY vertical space (Word-verified: a page
  // fits 20 borderless 32px rows but only 16 with 5.25pt borders, and 20
  // flush paragraph lines outside any table — the space belongs to the
  // table's gridlines, not to a page epsilon). Each gridline reserves the
  // table-level border width: the top edge, one line between each pair of
  // rows (insideH), and the bottom edge — rowY[r] is the top of ROW r's
  // content, BELOW the gridline above it; rowY[nrows] adds the bottom edge
  // so the table's height includes it. Per-cell border overrides do not
  // move gridlines (approximation; table-level widths dominate real docs).
  const bw = (side: TableBorders[keyof TableBorders] | undefined): number =>
    side ? side.width : 0;
  const tTop = bw(table.borders?.top);
  const tInside = bw(table.borders?.insideH);
  const tBottom = bw(table.borders?.bottom);
  const rowY = new Array<number>(nrows + 1).fill(0);
  rowY[0] = tTop;
  for (let r = 0; r < nrows; r++)
    rowY[r + 1] = rowY[r] + rowHeight[r] + (r + 1 < nrows ? tInside : tBottom);

  // Position cells and shift their content into place.
  const cells: ResolvedCell[] = cellDrafts.map((c) => {
    let height = 0;
    for (let r = c.startRow; r < c.startRow + c.rowspan && r < nrows; r++)
      height += rowHeight[r];
    // Merged cells swallow the inner gridlines they span.
    height += bwInsideForSpan(
      c.startRow,
      c.rowspan,
      nrows,
      table.borders?.insideH,
    );
    // w:vAlign: distribute the slack above/centered for non-top cells.
    const slack = Math.max(
      0,
      height - c.pad.top - c.pad.bottom - c.contentHeight,
    );
    const vOffset =
      c.vAlign === 'bottom' ? slack : c.vAlign === 'center' ? slack / 2 : 0;
    const dy = rowY[c.startRow] + c.pad.top + vOffset;
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
    if (vOffset > 0) cell.vShift = vOffset;
    if (c.background) cell.background = c.background;
    if (c.borders) cell.borders = c.borders;
    if (c.tables.length > 0) cell.tables = c.tables;
    if (c.floats.length > 0)
      cell.floats = c.floats.map((f) => ({ ...f, y: f.y + dy }));
    return cell;
  });

  const resolved: ResolvedTable = {
    x: colX[0],
    y: 0,
    width: tableWidth,
    height: rowY[nrows],
    cells,
  };
  if (table.borders) resolved.borders = table.borders;

  // Repeating header band: contiguous header rows from the top, provided no
  // cell spans out of the band (a rowspan into the body would have to split).
  let headerRows = 0;
  while (headerRows < nrows && table.rows[headerRows].header) headerRows++;
  if (headerRows > 0 && headerRows < nrows) {
    const headerBottom = rowY[headerRows];
    const spansOut = cells.some(
      (c) => c.y < headerBottom && c.y + c.height > headerBottom,
    );
    if (!spansOut) resolved.headerBottom = headerBottom;
  }
  const cantSplitBands = table.rows
    .map((row, r) =>
      row.cantSplit ? { top: rowY[r], bottom: rowY[r + 1] } : null,
    )
    .filter((b): b is { top: number; bottom: number } => b !== null);
  if (cantSplitBands.length > 0) resolved.cantSplitBands = cantSplitBands;
  // Rows whose FIRST cell opens with a keepNext paragraph: Word does not
  // START such a row in a band's leftover (see ResolvedTable.keepStartBands).
  // Only the first cell counts — nested_table.docx row 2 has keepNext
  // openers in cells 3-4 and Word starts it mid-page regardless, while row 3
  // (keepNext in cell 1) is pushed whole.
  const keepStartBands = table.rows
    .map((row, r) => {
      const first = row.cells[0]?.content[0];
      return first?.type === 'paragraph' && first.keepNext === true
        ? { top: rowY[r], bottom: rowY[r + 1] }
        : null;
    })
    .filter((b): b is { top: number; bottom: number } => b !== null);
  if (keepStartBands.length > 0) resolved.keepStartBands = keepStartBands;
  return resolved;
}

/** Re-base a band list onto a fragment: shift by `delta`, keep only bands
 *  overlapping [limitTop, limitBottom). Shared by every band list on
 *  ResolvedTable (cantSplitBands, keepStartBands) at every re-basing site —
 *  band lists move together or a new one silently goes stale. */
function shiftBands(
  bands: { top: number; bottom: number }[] | undefined,
  delta: number,
  limitTop: number,
  limitBottom: number,
): { top: number; bottom: number }[] | undefined {
  return bands
    ?.map((b) => ({ top: b.top + delta, bottom: b.bottom + delta }))
    .filter((b) => b.bottom > limitTop && b.top < limitBottom);
}

/** Apply shiftBands to BOTH band lists of a fragment in place. */
function rebaseBandsOnto(
  frag: ResolvedTable,
  source: ResolvedTable,
  delta: number,
  limitTop: number,
  limitBottom: number,
): void {
  const cant = shiftBands(source.cantSplitBands, delta, limitTop, limitBottom);
  if (cant?.length) frag.cantSplitBands = cant;
  else delete frag.cantSplitBands;
  const keep = shiftBands(source.keepStartBands, delta, limitTop, limitBottom);
  if (keep?.length) frag.keepStartBands = keep;
  else delete frag.keepStartBands;
}

/** One laid-out top-level block awaiting vertical placement. */
type ParaItem = {
  /** Flattened paragraph (lazy — cache hits skip flattening until needed). */
  getFlow: () => FlowParagraph;
  /** Pre-wrapped constant-band lines; null when the paragraph anchors floats
   *  (those must wrap at placement time, when their y is known). */
  drafts: LineDraft[] | null;
  /** w:spacing before/after gaps (px) and a forced page break. */
  before?: number;
  after?: number;
  pageBreakBefore?: boolean;
  /** The paragraph contains a column break somewhere in its runs. Kept as a
   *  flag so the placer can route it to the placement-time wrap (a mid-
   *  paragraph column switch has no pre-wrapped form) and so a section that
   *  holds one is left unbalanced, without flattening the flow to find out. */
  columnBreak?: boolean;
  /** A next-page section break's mark: keeps its line mid-page and may flow
   *  to the next column, but never OPENS a page — an unfitting one is
   *  clipped at the floor instead (see blockSection.pageBreakMark). */
  pageBreakMark?: boolean;
  /** Pagination keeps (w:keepNext / w:keepLines / w:widowControl off). */
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  /** w:pBdr — a border box painted around the paragraph's lines. */
  borders?: ParagraphBorders;
  /** w:shd — a fill painted behind them, in that same box. */
  shading?: string;
};
type SectionMarker = ColumnConfig & {
  newPage: boolean;
  height?: number;
  /** The section contains an explicit column break, so its columns are NOT
   *  balanced — Word balances what a continuous section break leaves over,
   *  and a manual break is the author saying where the column ends instead.
   *  Filled by assignSectionHeights, which already walks each section. */
  hasColumnBreak?: boolean;
  /** The section is followed by one that starts CONTINUOUS — i.e. it ends in
   *  a continuous section break, which is the only thing that makes Word
   *  balance its columns: "without a section break at the end of your
   *  columned section, Word won't balance the text — it will simply fill the
   *  first column before moving to the next". The document's last section
   *  never has it. Filled by assignSectionHeights, which sees the markers in
   *  order. */
  endsContinuous?: boolean;
  /** Per-section page-geometry override (sanitized); absent → config.page. */
  page?: PageConfig;
  /** Section index — stamps pages (ResolvedPage.chromeIndex) so the painter
   *  picks that section's header/footer set. */
  chromeIndex?: number;
};
type BlockItem = ({ para: ParaItem } | { table: ResolvedTable }) & {
  /** Set on the first block of each section: switch column flow here. `height`
   *  (filled by assignSectionHeights) is the section's total content height,
   *  used to balance the columns on the section's final page. */
  section?: SectionMarker;
  /** The document node this item came from + its PM offset — the identity the
   *  page cache matches on (PM structural sharing keeps unchanged nodes
   *  reference-equal across transactions). */
  node?: PMNode;
  nodeOffset?: number;
};

/**
 * A paragraph's spacing as the placer should see it, with `w:contextualSpacing`
 * applied: a side that borders a paragraph of the SAME style contributes
 * nothing ("ignore spacing above and below when using identical styles").
 * Which sides those are is decided at import, where both neighbours are
 * visible — see resolveContextualSpacing in the docx package.
 */
function effectiveSpacing(
  spacing: ParagraphSpacing | null,
  contextual: { before: boolean; after: boolean } | null,
): ParagraphSpacing | null {
  if (!spacing || !contextual) return spacing;
  if (!contextual.before && !contextual.after) return spacing;
  return {
    ...spacing,
    ...(contextual.before ? { before: 0 } : {}),
    ...(contextual.after ? { after: 0 } : {}),
  };
}

/**
 * The gap owed ABOVE a block, given the space-after the block before it
 * already contributed.
 *
 * OOXML does not add the two: "a consumer shall use the maximum of the
 * inter-line spacing in each paragraph, the spacing after the first paragraph
 * and the spacing before the second paragraph to determine the net spacing
 * between the paragraphs." The previous block's `after` is already on the
 * page, so only whatever `before` asks for BEYOND it is still owed —
 * `after + max(0, before - after)` is exactly `max(after, before)`.
 *
 * The spec's third term, each paragraph's inter-line spacing, is deliberately
 * left out: w:line is already baked into our line boxes, so counting it here
 * would apply it twice. If a document ever spaces wider than Word by exactly
 * one line, this is the omission to revisit.
 */
function collapsedBefore(
  before: number | undefined,
  pendingAfter: number,
): number {
  return Math.max(0, (before ?? 0) - pendingAfter);
}

/** Estimated laid-out height of a block item (drafts + spacing, or table). */
function blockItemHeight(item: BlockItem): number {
  if ('para' in item) {
    const lines = item.para.drafts?.reduce((s, d) => s + d.height, 0) ?? 0;
    return lines + (item.para.before ?? 0) + (item.para.after ?? 0);
  }
  return item.table.height;
}

/** Fill each section marker's `height` with the sum of its blocks' heights, so
 *  the placer can balance columns on the section's final page. */
function assignSectionHeights(items: BlockItem[]): void {
  let current: SectionMarker | null = null;
  // Heights are summed the way they are placed: adjacent spacing collapses,
  // so subtract the overlap rather than counting both sides.
  let pendingAfter = 0;
  for (const item of items) {
    if (item.section) {
      // This section starting continuous is exactly what "the previous one
      // ends in a continuous break" means — the break type lives on the
      // section it BEGINS (ISO 29500 §17.6.18 puts a section's properties at
      // its end, which is what makes the XML read backwards).
      if (current) current.endsContinuous = !item.section.newPage;
      current = item.section;
      current.height = 0;
      current.hasColumnBreak = false;
      pendingAfter = 0;
    }
    if (current && 'para' in item && item.para.columnBreak)
      current.hasColumnBreak = true;
    const overlap =
      'para' in item
        ? Math.min(pendingAfter, item.para.before ?? 0)
        : pendingAfter;
    if (current)
      current.height = (current.height ?? 0) + blockItemHeight(item) - overlap;
    pendingAfter = 'para' in item ? (item.para.after ?? 0) : 0;
  }
}

/** A rectangle text must flow around (a float's box plus its text gaps). */
interface Exclusion {
  left: number;
  right: number;
  top: number;
  bottom: number;
  /** From a through/tight-wrapped float: a paragraph pushed below this
   *  rectangle takes its ANCHOR with it (see FlowFloat.through). */
  through?: boolean;
}

/** Narrowest band we'll still flow text into beside a float. */
const MIN_BAND = 24;

/**
 * How far a NEXT-PAGE section-break mark's line may poke past the page floor
 * and still be absorbed (clipped) rather than spilled to a page of its own.
 *
 * The two sides of this constant are measured, the exact value is not. A
 * generated probe (four sections, marks alternately fitting and not, at
 * bottom margins 720 and 0) paginates in Word as one page per section with
 * marks poking 3–6px; the factsheet's mark pokes 8.9px and Word absorbs it
 * too. Two report covers carry 1.5-spaced marks that Word gives a page of
 * their own, the nearer poking 9.9px. Absorbed at ≤8.9, spilled at ≥9.9 —
 * the value sits in that gap. The gap is one pixel wide in OUR coordinates,
 * so any future change to line metrics near a report cover will trip the
 * pagination baseline; that is the tripwire working, not a mystery.
 */
const PAGE_BREAK_MARK_POKE = 9.4;

/** Shift one cell (box + contents) vertically in place. */
function shiftCell(cell: ResolvedCell, dy: number): ResolvedCell {
  cell.y += dy;
  for (const line of cell.lines) line.y += dy;
  cell.floats?.forEach((f) => (f.y += dy));
  cell.tables?.forEach((t) => offsetTable(t, dy));
  return cell;
}

/** Deep copy of a resolved table (cells, lines, nested tables). */
function cloneTable(t: ResolvedTable): ResolvedTable {
  return { ...t, cells: t.cells.map(cloneCell) };
}

function cloneCell(cell: ResolvedCell): ResolvedCell {
  const copy: ResolvedCell = {
    ...cell,
    lines: cell.lines.map((l) => ({ ...l })),
  };
  if (cell.tables) copy.tables = cell.tables.map(cloneTable);
  if (cell.floats) copy.floats = cell.floats.map((f) => ({ ...f }));
  return copy;
}

/** A copy of the cell with its vertical-alignment slack removed — content
 *  stacked from the cell's top. Word's rule while the cell's row is split:
 *  centering resumes only when the row is whole. Without this, a centered
 *  one-liner sits BELOW the cut, so the first fragment painted an empty box
 *  while its text rode to the next page. */
function unshiftCell(cell: ResolvedCell): ResolvedCell {
  const s = cell.vShift ?? 0;
  if (s <= 0) return cell;
  const copy = cloneCell(cell);
  delete copy.vShift;
  for (const l of copy.lines) l.y -= s;
  copy.tables?.forEach((t) => offsetTable(t, -s));
  copy.floats?.forEach((f) => (f.y -= s));
  return copy;
}

/**
 * The index (0..lines.length) where a page cut through a cell's line stack
 * may legally fall: the largest boundary at or below the geometric cut that
 * violates no keep. Between paragraphs the boundary is blocked by the upper
 * paragraph's w:keepNext; inside one paragraph by w:keepLines, and by widow/
 * orphan control (Word default ON) unless ≥2 of the paragraph's lines stay on
 * EACH side. Lines are grouped into paragraphs by their shared `keeps`
 * reference, and counts are taken fresh from THIS line stack — facts, not
 * precomputed decisions, so a fragment's second split judges the fragment's
 * own remainder (a 3-line leftover of an already-split paragraph must move
 * whole, whatever the original allowed).
 */
function legalPartition(lines: LayoutLine[], cut: number): number {
  let g = 0;
  while (
    g < lines.length &&
    lines[g].y + lines[g].height <= cut + REMAINDER_EPS
  )
    g++;
  for (let p = g; p > 0; p--) {
    if (p >= lines.length) return p; // nothing moves — not a split at all
    const ka = lines[p - 1].keeps;
    const kb = lines[p].keeps;
    if (ka !== kb) {
      // A paragraph boundary (or missing facts on either side).
      if (ka?.keepNext) continue; // a keepNext block may not end a fragment
      return p;
    }
    if (!ka) return p; // no facts on these lines — geometry decides
    if (ka.keepLines) continue;
    if (ka.widowControl !== false) {
      let above = 0;
      for (let i = 0; i < p; i++) if (lines[i].keeps === ka) above++;
      let below = 0;
      for (let i = p; i < lines.length; i++) if (lines[i].keeps === ka) below++;
      if (above < 2 || below < 2) continue;
    }
    return p;
  }
  return 0;
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
function splitTableAt(
  table: ResolvedTable,
  cut: number,
  opts?: {
    /** Recursively split a nested table straddling the cut. Word does this
     *  regardless of the nested table's height (fixture-verified — even one
     *  that would fit whole on a fresh page splits at the page edge), so
     *  the placement loop always passes true; the flag exists so direct
     *  callers (specs) keep pure geometric splitting. Propagates down. */
    splitNested?: boolean;
  },
): { top: ResolvedTable; rest: ResolvedTable } {
  interface Cont {
    /** The cell with centering slack removed — what BOTH fragments read. */
    view: ResolvedCell;
    topLines: LayoutLine[];
    topTables: ResolvedTable[];
    /** Continuation content, already positioned at its 0-based y. */
    contLines: LayoutLine[];
    contTables: ResolvedTable[];
    /** Original y of the first continued item (float re-anchoring). */
    firstY: number;
  }
  const straddlers = new Map<ResolvedCell, Cont>();
  let splitBottom = cut; // bottom of the broken row (cut itself if none breaks)
  let contHeight = 0; // height of the continuation row in `rest`

  for (const cell of table.cells) {
    // Seam tolerance: nested coordinates go through offset round-trips, so
    // an adjacent cell's bottom can exceed the next row's top by an ulp —
    // strict comparisons then invent an empty straddler stub at the seam.
    if (
      cell.y >= cut - REMAINDER_EPS ||
      cell.y + cell.height <= cut + REMAINDER_EPS
    )
      continue;
    splitBottom = Math.max(splitBottom, cell.y + cell.height);
    const view = unshiftCell(cell);
    const part = legalPartition(view.lines, cut);
    const remLines = view.lines.slice(part);
    const topLines = view.lines.slice(0, part);
    const minRemLineY = remLines.length
      ? Math.min(...remLines.map((l) => l.y))
      : Infinity;

    // Classify the cell's nested tables. One continuation item per table or
    // line, in source order; a straddling nested table taller than a full
    // band splits recursively — its top part stays in place, its remainder
    // becomes a continuation item whose height DIFFERS from the slot it
    // vacated, which is what forces the sequential re-stack below.
    interface ContItem {
      origTop: number;
      origBottom: number;
      height: number;
      line?: LayoutLine;
      table?: ResolvedTable;
    }
    const topTables: ResolvedTable[] = [];
    const items: ContItem[] = [];
    let nestedSplit = false;
    for (const t of view.tables ?? []) {
      if (t.y + t.height <= cut + REMAINDER_EPS) {
        topTables.push(t);
        continue;
      }
      if (t.y < cut - REMAINDER_EPS) {
        // A keep-pushed line ABOVE the nested table forces the whole table
        // down with it — splitting would put the line under rows it
        // originally preceded.
        if (opts?.splitNested && minRemLineY >= t.y) {
          const local = cloneTable(t);
          offsetTable(local, -t.y);
          const localCut = cut - t.y;
          let sub = splitTableAt(local, localCut, opts);
          // Same policy as the outer placement loop, one level down: fill
          // mid-row, but fall back to the nested table's own row boundary
          // when the straddling nested row contributes nothing — otherwise
          // its top fragment ends in an empty stub row.
          let b = 0;
          for (const c2 of local.cells)
            if (c2.y > REMAINDER_EPS && c2.y <= localCut) b = Math.max(b, c2.y);
          const rowContributed = sub.top.cells.some(
            (c2) =>
              c2.lines.some((l) => l.y + l.height > b + REMAINDER_EPS) ||
              (c2.tables ?? []).some(
                (nt) => nt.y + nt.height > b + REMAINDER_EPS,
              ),
          );
          if (!rowContributed && b > 0) sub = splitTableAt(local, b, opts);
          const subHasContent = sub.top.cells.some(
            (c2) => c2.lines.length > 0 || (c2.tables?.length ?? 0) > 0,
          );
          if (subHasContent && sub.rest.height < t.height) {
            offsetTable(sub.top, t.y);
            topTables.push(sub.top);
            items.push({
              origTop: cut,
              origBottom: t.y + t.height,
              height: sub.rest.height,
              table: sub.rest,
            });
            nestedSplit = true;
            continue;
          }
        }
      }
      items.push({
        origTop: t.y,
        origBottom: t.y + t.height,
        height: t.height,
        table: cloneTable(t),
      });
    }
    for (const l of remLines)
      items.push({
        origTop: l.y,
        origBottom: l.y + l.height,
        height: l.height,
        line: l,
      });
    items.sort((a, b) => a.origTop - b.origTop);

    const firstY = items.length ? items[0].origTop : cut;
    const contLines: LayoutLine[] = [];
    const contTables: ResolvedTable[] = [];
    let extent = 0;
    if (!nestedSplit) {
      // Pure translation — every item keeps its relative position (the
      // continuation is the original layout shifted up as one block).
      const delta = -firstY;
      let bottom = firstY;
      for (const it of items) {
        if (it.line) contLines.push({ ...it.line, y: it.line.y + delta });
        else {
          const tb = it.table as ResolvedTable;
          offsetTable(tb, delta);
          contTables.push(tb);
        }
        bottom = Math.max(bottom, it.origBottom);
      }
      extent = bottom - firstY;
    } else {
      // A nested table changed height when it split, so everything below it
      // RE-STACKS sequentially: source order, original inter-item gaps, new
      // heights. A single delta would leave the old gap where the vanished
      // rows used to be.
      let yCur = 0;
      let prevBottom: number | null = null;
      for (const it of items) {
        const gap =
          prevBottom === null ? 0 : Math.max(0, it.origTop - prevBottom);
        const newTop = yCur + gap;
        if (it.line) contLines.push({ ...it.line, y: newTop });
        else {
          const tb = it.table as ResolvedTable;
          offsetTable(tb, newTop - tb.y);
          contTables.push(tb);
        }
        yCur = newTop + it.height;
        prevBottom = it.origBottom;
      }
      extent = yCur;
    }
    contHeight = Math.max(contHeight, extent);
    straddlers.set(cell, {
      view,
      topLines,
      topTables,
      contLines,
      contTables,
      firstY,
    });
  }

  const topCells: ResolvedCell[] = [];
  const restCells: ResolvedCell[] = [];
  for (const cell of table.cells) {
    if (cell.y + cell.height <= cut + REMAINDER_EPS) {
      topCells.push(cell);
    } else if (cell.y >= cut - REMAINDER_EPS) {
      // a row below the break: follows beneath the continuation row.
      // Copied so the caller can still fall back to placing the original whole.
      restCells.push(shiftCell(cloneCell(cell), contHeight - splitBottom));
    } else {
      const c = straddlers.get(cell) as Cont;
      // Both fragments read the UN-CENTERED view (see unshiftCell) so the
      // first lines land on the first fragment, as Word splits. The partition
      // is the SNAPPED one — top and rem must complement exactly, or a line
      // pushed down by a keep would also stay up. Nested tables were already
      // classified (and possibly recursively split) in the first pass.
      const topCell: ResolvedCell = {
        ...cell,
        height: cut - cell.y,
        lines: c.topLines,
      };
      delete topCell.vShift; // its content is top-stacked now
      if (c.topTables.length > 0) topCell.tables = c.topTables;
      else delete topCell.tables;
      const topFloats = (c.view.floats ?? []).filter(
        (f) => f.y + f.height <= cut + REMAINDER_EPS,
      );
      if (topFloats.length > 0) topCell.floats = topFloats;
      else delete topCell.floats;
      topCells.push(topCell);

      const lines = c.contLines;
      const remTables = c.contTables;
      const remFloats = (c.view.floats ?? [])
        .filter((f) => f.y + f.height > cut + REMAINDER_EPS)
        .map((f) => ({ ...f, y: f.y - c.firstY }));
      // The continuation inherits everything that describes the CELL —
      // borders, fill — and overrides only what the split changes. Building
      // it field-by-field silently dropped `borders` and `background`: the
      // second fragment of a double-ruled header row fell back to the
      // table's plain grid line, and a shaded cell lost its fill. Only
      // `vShift` is deliberately absent (centering stays suspended while the
      // row is split — see unshiftCell). Word draws a border on BOTH sides of
      // the split, so both fragments must carry the same cell borders.
      const contCell: ResolvedCell = {
        ...cell,
        y: 0,
        height: contHeight,
        lines,
      };
      delete contCell.vShift;
      if (remTables.length > 0) contCell.tables = remTables;
      else delete contCell.tables;
      if (remFloats.length > 0) contCell.floats = remFloats;
      else delete contCell.floats;
      restCells.push(contCell);
    }
  }

  // Both fragments inherit the table's own properties; each then fixes up
  // what the split actually changed. Three kinds of field, spelled out so a
  // NEW one can't be forgotten (the property test in the spec enforces it):
  //   · carried as-is  — borders, and anything describing the table itself
  //   · re-based       — the band lists, measured from the fragment's top
  //     (shiftBands / rebaseBandsOnto)
  //   · dropped        — headerBottom on `rest`: the continuation has no
  //     header band unless ghost rows are actually prepended (the caller
  //     sets it then; cloneHeaderCells can refuse).
  const restDelta = contHeight - splitBottom;
  const top: ResolvedTable = {
    ...table,
    y: 0,
    height: cut,
    cells: topCells,
  };
  rebaseBandsOnto(top, table, 0, 0, cut);

  const rest: ResolvedTable = {
    ...table,
    y: 0,
    height: contHeight + (table.height - splitBottom),
    cells: restCells,
  };
  delete rest.headerBottom;
  rebaseBandsOnto(rest, table, restDelta, 0, rest.height);
  return { top, rest };
}

/** Ghost copies of the header-band cells for a continuation fragment. PM
 *  positions are stripped so selection and hit-testing only ever target the
 *  original header; returns null when the band is too complex to repeat. */
function cloneHeaderCells(
  table: ResolvedTable,
  headerBottom: number,
): ResolvedCell[] | null {
  const band = table.cells.filter((c) => c.y + c.height <= headerBottom);
  if (band.some((c) => c.tables && c.tables.length > 0)) return null;
  return band.map((cell) => ({
    ...cell,
    tables: undefined,
    floats: cell.floats?.map((f) => ({ ...f })),
    lines: cell.lines.map((l) => {
      const line: LayoutLine = {
        ...l,
        segments: l.segments.map((s) => ({ ...s, pos: undefined })),
      };
      delete line.from;
      delete line.to;
      if (line.images)
        line.images = line.images.map((im) => ({ ...im, pos: undefined }));
      return line;
    }),
  }));
}

/**
 * The unplaced remainder of a table mid-pagination, as a cursor into an
 * IMMUTABLE base layout instead of a re-cloned table.
 *
 * The old representation (`table = rest` after each split) deep-cloned every
 * row below the cut once per page — O(rows × pages) for a long table — and the
 * per-page checkpoint cloned the whole remainder again. With this cursor the
 * base is laid out once and never mutated: fragments CLONE the rows they place
 * (each row is cloned exactly once across the table's whole run), checkpoints
 * share `base` by reference and copy only `carry`, and a tail-splice rebase is
 * a posDelta addition instead of a remainder-wide clone.
 */
interface TableRemainder {
  /** The full table in its own y = 0 space. Never mutated after layout —
   *  placed fragments clone out of it, checkpoints reference it. */
  base: ResolvedTable;
  /** Consumed prefix: base content above this y is already placed. Always a
   *  row boundary (the bottom of the last fully-or-partially placed row). */
  fromY: number;
  /** Continuation cells of a row broken at the last cut, re-stacked at y = 0.
   *  Owned by the live remainder (placement consumes them); checkpoints hold
   *  their own clones. Null when the last cut fell on a row boundary. */
  carry: ResolvedCell[] | null;
  /** Stacked height of `carry` (0 when carry is null). */
  carryHeight: number;
  /** Pending PM-position shift for `base` and `carry`, applied lazily when a
   *  checkpointed remainder is materialized for resume. Always 0 on the live
   *  remainder; tail-splice rebasing adds the edit's delta here in O(1). */
  posDelta: number;
}

/** Tolerance for comparing base-row coordinates against the accumulated
 *  consume cursor: `fromY` is recovered from height arithmetic, so it can sit
 *  a few ulps off the row boundary it logically equals. Pure float noise —
 *  far below any real row height. */
const REMAINDER_EPS = 1e-6;

/** A base cell's y in remainder space (y = 0 at the carry's top). The delta
 *  to the consume cursor is SNAPPED when it is float noise: the first
 *  unconsumed row must land exactly at carryHeight, or a cut derived from it
 *  misses the carry's own bottom edge by an ulp and the splitter sees a
 *  straddler where a clean boundary was meant. */
function remainderRelY(r: TableRemainder, cellY: number): number {
  const d = cellY - r.fromY;
  return r.carryHeight + (Math.abs(d) < REMAINDER_EPS ? 0 : d);
}

/** Full height of the unplaced remainder (excluding any repeated header). */
function remainderHeight(r: TableRemainder): number {
  return r.carryHeight + (r.base.height - r.fromY);
}

/** Whether the remainder is mid-table (anything already consumed). A fresh,
 *  untouched table resumes identically from the item's start, which keeps the
 *  page boundary CLEAN for the page cache's tail resync. */
function remainderStarted(r: TableRemainder): boolean {
  return r.fromY > 0 || r.carry !== null;
}

/** Deep-clone one resolved cell, shifting its PM positions by `delta`
 *  (the per-cell body of cloneTableShifted, reused for carry cells). */
function cloneCellPosShifted(cell: ResolvedCell, delta: number): ResolvedCell {
  return {
    ...cell,
    borders: cell.borders ? { ...cell.borders } : cell.borders,
    diagonals: cell.diagonals ? { ...cell.diagonals } : cell.diagonals,
    lines: cell.lines.map((l) => cloneLineShifted(l, delta)),
    tables: cell.tables?.map((nt) => cloneTableShifted(nt, delta)),
    floats: cell.floats?.map((f) => ({ ...f })),
  };
}

/** Whether two checkpointed remainders describe the same resume point,
 *  geometry-only: PM positions may differ by the edit's delta (the splice
 *  shifts them), so only the consume cursor, the carry's shape, and the
 *  base's frame are compared. The caller has already matched the table NODE
 *  by identity, and layout is deterministic, so equal cursors over the same
 *  node imply the same continuation. */
function remainderEq(a: TableRemainder, b: TableRemainder): boolean {
  if (a.fromY !== b.fromY || a.carryHeight !== b.carryHeight) return false;
  if (
    a.base.height !== b.base.height ||
    a.base.width !== b.base.width ||
    a.base.cells.length !== b.base.cells.length
  )
    return false;
  const ac = a.carry ?? [];
  const bc = b.carry ?? [];
  if (ac.length !== bc.length) return false;
  for (let i = 0; i < ac.length; i++) {
    const x = ac[i];
    const y = bc[i];
    if (
      x.x !== y.x ||
      x.y !== y.y ||
      x.width !== y.width ||
      x.height !== y.height ||
      x.lines.length !== y.lines.length
    )
      return false;
  }
  return true;
}

/** Snapshot a live remainder for a checkpoint: `base` is shared (immutable),
 *  only the small carry is cloned. O(carry), not O(remaining rows). */
function snapshotRemainder(r: TableRemainder): TableRemainder {
  return {
    ...r,
    carry: r.carry && r.carry.map((c) => cloneCellPosShifted(c, 0)),
  };
}

/** Materialize a checkpointed remainder for resume: apply the pending
 *  posDelta (base is cloned only when a tail-splice actually shifted it) and
 *  hand the loop its own mutable carry. The checkpoint stays pristine across
 *  retries, exactly like the old clone-on-restore. */
function materializeRemainder(r: TableRemainder): TableRemainder {
  return {
    base: r.posDelta === 0 ? r.base : cloneTableShifted(r.base, r.posDelta),
    fromY: r.fromY,
    carry: r.carry && r.carry.map((c) => cloneCellPosShifted(c, r.posDelta)),
    carryHeight: r.carryHeight,
    posDelta: 0,
  };
}

/**
 * The remainder's cells within `windowH` px of its top, cloned and re-based
 * to a y = 0 view the splitter (or a whole-remainder placement) can consume.
 * Base rows land at `cell.y - fromY + carryHeight`; carry cells sit at their
 * own y (already 0-based) and are handed over AS-IS — they are owned by the
 * live remainder and die with this page either way. Cells at or below the
 * window never enter the view, so they are never cloned: that is the whole
 * point of the representation.
 *
 * Passing Infinity materializes the entire remainder (the final fragment, or
 * a degenerate-geometry overflow placement).
 */
function remainderView(r: TableRemainder, windowH: number): ResolvedTable {
  const shift = r.carryHeight - r.fromY;
  const cells: ResolvedCell[] = [];
  let cloned = 0;
  if (r.carry) cells.push(...r.carry);
  for (const cell of r.base.cells) {
    // Cells starting before the cursor are spent: fully-consumed rows, and
    // straddlers of an earlier cut (rowspan tails included) whose remains
    // live on as carry cells — including the base original again would
    // double their content.
    if (cell.y < r.fromY - REMAINDER_EPS) continue;
    const yr = remainderRelY(r, cell.y);
    if (yr >= windowH - REMAINDER_EPS) continue; // out of this page's reach
    const cellShift = yr - cell.y;
    const copy = cloneCellPosShifted(cell, 0);
    copy.y = yr;
    for (const l of copy.lines) l.y += cellShift;
    copy.floats?.forEach((f) => (f.y += cellShift));
    copy.tables?.forEach((t) => offsetTable(t, cellShift));
    cells.push(copy);
    cloned++;
  }
  perf.bump('table.view.cellsCloned', cloned);
  const view: ResolvedTable = {
    ...r.base,
    y: 0,
    height: remainderHeight(r),
    cells,
  };
  // Repeated headers are a fragment-level concern handled by the placement
  // loop (withGhosts) against the base directly.
  delete view.headerBottom;
  // Band lists re-base onto the view so placed fragments keep carrying them
  // (part of the ResolvedTable contract — the spec asserts a fragment's
  // bands are measured from ITS top); splitTableAt filters them per fragment.
  rebaseBandsOnto(view, r.base, shift, 0, view.height);
  return view;
}

function buildCtx(config: LayoutConfig, docCompat?: DocCompat | null): Ctx {
  return {
    base: { ...DEFAULT_FONT, ...config.defaultFont },
    measure: config.measureText,
    metrics: config.measureMetrics,
    tabWidth: config.tabWidth ?? DEFAULT_TAB_WIDTH,
    fieldPlaceholder: '1',
    compat: { ...CURRENT_WORD_COMPAT, ...config.compat, ...docCompat },
  };
}

/** Everything that crosses a page boundary, captured at each page's start
 *  (the page accumulators are empty there, so this is a handful of scalars
 *  plus the mid-item resume handles). Serves both the fixed-point replay
 *  (Phase 2) and the cross-pass page cache (Phase 3): a page is a pure
 *  function of (its items, this carry), which is what makes reusing it safe. */
interface PageCheckpoint {
  itemIdx: number;
  midTable: TableRemainder | null;
  midDraftIdx: number | null;
  midBandedFrom: number | null;
  unreplayable: boolean;
  firstItem: boolean;
  curPage: PageConfig;
  curChromeIndex: number | undefined;
  pageChromeIndex: number | undefined;
  pageSectionFirst: boolean;
  sectionFirstPending: boolean;
  contentLeft: number;
  contentRight: number;
  contentWidth: number;
  colCount: number;
  colGap: number;
  colWidth: number;
  /** Per-column boxes of the section in flow (identity-stable: applyColumns
   *  builds one array per section, so identity is a sound comparison). */
  colBoxes: ColumnBox[] | null;
  balancing: boolean;
  sectionRemaining: number;
  activePBdr: {
    borders?: ParagraphBorders;
    shading?: string;
    startedEarlier: boolean;
    frag: { x0: number; x1: number; y0: number; y1: number } | null;
  } | null;
}

/** One placed page as recorded for cross-pass reuse: which document nodes fed
 *  it, the carry it started from, and the finished page. `endsMidItem` marks
 *  a page whose last item continues on the next one (a split table or
 *  paragraph), so consecutive entries overlap on that node. */
interface PageRunEntry {
  itemNodes: PMNode[];
  endsMidItem: boolean;
  cp: PageCheckpoint;
  /** PM offset of itemNodes[0] when recorded — the tail-splice shift base. */
  startOffset: number;
  page: ResolvedPage;
}
interface PageRun {
  entries: PageRunEntry[];
}

/** Control-flow signal for the tail splice: placement reached a clean page
 *  boundary whose carry and remaining items match the previous run — every
 *  page from `entryIdx` on can be reused (shifted by the PM-position delta)
 *  instead of re-placed. */
class TailSpliceSignal {
  constructor(readonly entryIdx: number) {}
}

/** A recorded page re-used at a different document position: geometry is
 *  identical by construction (same items, same carry) — only the PM
 *  positions and the page index shift. */
function shiftResolvedPage(
  p: ResolvedPage,
  delta: number,
  index: number,
): ResolvedPage {
  const out: ResolvedPage = {
    ...p,
    index,
    lines: p.lines.map((l) => cloneLineShifted(l, delta)),
  };
  if (p.tables) out.tables = p.tables.map((t) => cloneTableShifted(t, delta));
  if (p.floats)
    out.floats = p.floats.map((f) =>
      f.pos != null ? { ...f, pos: f.pos + delta } : { ...f },
    );
  return out;
}

/** A recorded checkpoint re-based for the spliced tail, so the NEXT pass can
 *  resume from it: item index and PM positions move by the edit's deltas. */
function rebaseCp(
  cp: PageCheckpoint,
  delta: number,
  idxDelta: number,
): PageCheckpoint {
  return {
    ...cp,
    itemIdx: cp.itemIdx + idxDelta,
    // O(1): the shift is recorded, not applied — `base` stays shared and
    // immutable; materializeRemainder applies the accumulated delta on the
    // rare actual resume.
    midTable: cp.midTable && {
      ...cp.midTable,
      posDelta: cp.midTable.posDelta + delta,
    },
    midBandedFrom: cp.midBandedFrom != null ? cp.midBandedFrom + delta : null,
  };
}

/** Control-flow signal for the page fixed-point: a float registered itself
 *  over content already placed above it (Word wraps the WHOLE page around a
 *  float, wherever its anchor sits), so the page must be re-run with that
 *  float's exclusion known from the start. Thrown from deep inside the
 *  placement stack; the page loop catches it, restores the page-start
 *  checkpoint and replays. `drop` retracts seeds whose anchor left the page. */
class PageRetrySignal {
  constructor(
    readonly add: { key: string; rect: Exclusion }[] = [],
    readonly drop: string[] = [],
  ) {}
}

/** Bounded retries per page: margin/page-relative floats converge in one
 *  replay (their rect is page-constant); the cap only guards the rare
 *  paragraph-relative feedback loop. The final attempt places with triggers
 *  disabled, so termination is unconditional. */
const MAX_PAGE_RETRIES = 3;

/** Stack laid-out blocks onto pages (the paginator). `bandFor` computes the
 *  vertical content bounds for a page geometry (e.g. pushed in by a tall page
 *  header/footer) — a function because sections can override the geometry. */
function placeBlocks(
  // An ARRAY (not just an iterable): the keepNext look-ahead indexes forward.
  items: BlockItem[],
  config: LayoutConfig,
  ctx: Ctx,
  bandFor?: (
    p: PageConfig,
    chromeIndex?: number,
  ) => { top: number; bottom: number },
  footnotes?: Map<number, FootnoteBody>,
  // Cross-pass page cache (Phase 3): `store` persists in the host's
  // LayoutCache; `key` fingerprints every input a page depends on besides
  // its items and carry (geometry, chrome bands, tab grid, footnotes).
  pageCache?: { store: Map<string, PageRun>; key: string },
): ResolvedLayout {
  const { page } = config;
  // Which section's chrome the CURRENT page shows: the section in effect where
  // the page starts (a continuous section joining mid-page doesn't change it —
  // Word's rule). `sectionFirst` drives per-section titlePg.
  let curChromeIndex: number | undefined;
  let pageChromeIndex: number | undefined;
  let pageSectionFirst = true;
  let sectionFirstPending = false;
  const bandOf = (p: PageConfig) =>
    bandFor?.(p, curChromeIndex) ?? {
      top: p.margin.top,
      bottom: p.height - p.margin.bottom,
    };
  // Geometry of the page being filled — mutable: a section carrying a `page`
  // override swaps it at its boundary (always a page boundary; a continuous
  // break with differing geometry is promoted to next-page, as Word does).
  let curPage = page;
  let { top, bottom } = bandOf(page);
  let contentLeft = contentLeftOf(page);
  let contentRight = page.width - page.margin.right;

  const pages: ResolvedPage[] = [];
  let lines: LayoutLine[] = [];
  let tables: ResolvedTable[] = [];
  let pageFloats: ResolvedFloat[] = [];
  let exclusions: Exclusion[] = []; // floats die at their page's end
  let y = top;

  // ── Page fixed-point state ──────────────────────────────────────────
  // `idx` is hoisted so page-start checkpoints can record which item was in
  // flight; the mid-item fields record HOW FAR into it the page boundary
  // fell (a table's unplaced remainder / a paragraph's next draft line).
  let idx = 0;
  let midTableRest: TableRemainder | null = null;
  let midParaDraftIdx: number | null = null;
  let inBandedPara = false; // inside a float paragraph's wrap call
  // Mid-wrap resume position of a float-anchoring paragraph: the PM start of
  // the line being placed. UNRESUMABLE (-1) when the flow carries no
  // positions (flat API) — such a page keeps the old can't-replay behavior.
  const UNRESUMABLE = -1;
  let midBandedFrom: number | null = null;
  let resumeTable: TableRemainder | null = null; // handed to the first item
  let resumeDraftIdx: number | null = null; // of a replayed page
  let resumeBandedFrom: number | null = null;
  const seeds = new Map<string, Exclusion>();
  /**
   * First-pass float positions, by the float's PM pos.
   *
   * A replay wraps LINES around the seeded exclusions, but it must not move
   * the ANCHORS: Word's model is incremental — a float's position is fixed
   * when its paragraph is first laid (later floats do not exist yet), and
   * "wrap the whole page around a float" then re-wraps the text alone. Left
   * to recompute, a replayed float whose own exclusion pushed its anchor's
   * paragraph re-registered somewhere new, threw a fresh seed, and the
   * sticky-drop gave up — leaving the factsheet's photo cluster overlapping
   * the lines above it. Pinned, the replay reproduces the first-pass rect
   * and the fixed point closes in one round trip.
   */
  const seedPins = new Map<number, { x: number; y: number; key: string }>();
  // Seeds retracted because they pushed their own anchor off the page —
  // never re-tried, so the oscillating case settles in one round trip.
  const droppedSeeds = new Set<string>();
  const registeredKeys = new Set<string>();
  let retriesLeft = MAX_PAGE_RETRIES;
  let pageStartCp!: PageCheckpoint; // assigned before the loop, re-taken per page

  // ── Page-cache recording (Phase 3) ──────────────────────────────────
  const prior = pageCache?.store.get(pageCache.key);
  // Clean-boundary entries of the previous run, indexed by their first node —
  // the tail-resync lookup.
  const priorFirstNode = new Map<PMNode, number>();
  prior?.entries.forEach((e, i) => {
    const n = e.itemNodes[0];
    const cp = e.cp;
    if (
      n &&
      !priorFirstNode.has(n) &&
      cp.midTable === null &&
      cp.midDraftIdx === null &&
      cp.midBandedFrom === null &&
      cp.activePBdr === null &&
      !cp.unreplayable
    ) {
      priorFirstNode.set(n, i);
    }
  });
  // Mid-TABLE entries of the previous run, indexed by the in-flight table's
  // node — several pages of one long table share that node, so this maps to
  // a candidate list and the resync check picks the entry whose remainder
  // cursor matches. Mid-paragraph boundaries stay non-resyncable.
  const priorMidTableEntries = new Map<PMNode, number[]>();
  prior?.entries.forEach((e, i) => {
    const n = e.itemNodes[0];
    const cp = e.cp;
    if (
      n &&
      cp.midTable !== null &&
      cp.midDraftIdx === null &&
      cp.midBandedFrom === null &&
      cp.activePBdr === null &&
      !cp.unreplayable
    ) {
      const list = priorMidTableEntries.get(n);
      if (list) list.push(i);
      else priorMidTableEntries.set(n, [i]);
    }
  });
  const record: PageRunEntry[] = [];
  let pageNodes: PMNode[] = [];
  let pageFirstOffset: number | null = null;
  let currentPageCp!: PageCheckpoint; // the checkpoint that OPENED this page
  /** The in-flight item continues past this page boundary. */
  const midAny = () =>
    midTableRest !== null || midParaDraftIdx !== null || inBandedPara;
  /** Re-seed the recording state for a page that starts at checkpoint `cp`. */
  const seedRecording = (cp: PageCheckpoint) => {
    currentPageCp = cp;
    const mid =
      cp.midTable !== null ||
      cp.midDraftIdx !== null ||
      cp.midBandedFrom !== null ||
      cp.unreplayable;
    const it = mid ? items[cp.itemIdx] : undefined;
    pageNodes = it?.node ? [it.node] : [];
    pageFirstOffset = it?.node ? (it.nodeOffset ?? null) : null;
  };
  /** The scalar half of a page-start carry comparison — everything except
   *  the mid-table resume handle, whose policy differs per call site. */
  const carryScalarsEq = (a: PageCheckpoint, b: PageCheckpoint): boolean =>
    a.midDraftIdx === null &&
    b.midDraftIdx === null &&
    a.midBandedFrom === null &&
    b.midBandedFrom === null &&
    a.activePBdr === null &&
    b.activePBdr === null &&
    !a.unreplayable &&
    !b.unreplayable &&
    a.firstItem === b.firstItem &&
    a.curChromeIndex === b.curChromeIndex &&
    a.pageChromeIndex === b.pageChromeIndex &&
    a.pageSectionFirst === b.pageSectionFirst &&
    a.sectionFirstPending === b.sectionFirstPending &&
    a.contentLeft === b.contentLeft &&
    a.contentRight === b.contentRight &&
    a.contentWidth === b.contentWidth &&
    a.colCount === b.colCount &&
    a.colGap === b.colGap &&
    a.colWidth === b.colWidth &&
    a.colBoxes === b.colBoxes &&
    a.balancing === b.balancing &&
    a.sectionRemaining === b.sectionRemaining &&
    sameGeom(a.curPage, b.curPage);
  /** Two clean page-start carries describe the same resume point. */
  const carryEq = (a: PageCheckpoint, b: PageCheckpoint): boolean =>
    a.midTable === null && b.midTable === null && carryScalarsEq(a, b);
  /** Two MID-TABLE page starts describe the same resume point: the scalars
   *  match and both remainders sit at the same cursor over the same (node-
   *  identical) table. */
  const midTableCarryEq = (a: PageCheckpoint, b: PageCheckpoint): boolean =>
    a.midTable !== null &&
    b.midTable !== null &&
    remainderEq(a.midTable, b.midTable) &&
    carryScalarsEq(a, b);
  /** The remaining items match the recorded run from `entryIdx` on. */
  const tailMatches = (entryIdx: number): boolean => {
    if (!prior) return false;
    let c = idx;
    let prevEndsMid = false;
    for (let k = entryIdx; k < prior.entries.length; k++) {
      const e = prior.entries[k];
      for (let j = 0; j < e.itemNodes.length; j++) {
        if (j === 0 && prevEndsMid) {
          // Consecutive entries overlap on a split item: its node was
          // already matched as the previous entry's last — just verify.
          if (items[c - 1]?.node !== e.itemNodes[0]) return false;
          continue;
        }
        if (items[c]?.node !== e.itemNodes[j]) return false;
        c++;
      }
      prevEndsMid = e.endsMidItem;
    }
    return c === items.length;
  };
  // w:pBdr tracking: while a bordered paragraph places its lines, a fragment
  // accumulates; a column/page break flushes it as a box (top edge only on
  // the first fragment, bottom only on the last).
  let paraBoxes: ParagraphBox[] = [];
  let activePBdr: {
    borders?: ParagraphBorders;
    shading?: string;
    startedEarlier: boolean;
    frag: { x0: number; x1: number; y0: number; y1: number } | null;
  } | null = null;
  const flushPBdrFrag = (isLast: boolean) => {
    const a = activePBdr;
    if (!a?.frag) return;
    paraBoxes.push({
      x: a.frag.x0,
      y: a.frag.y0,
      width: a.frag.x1 - a.frag.x0,
      height: a.frag.y1 - a.frag.y0,
      ...(a.borders && { borders: a.borders }),
      ...(a.shading && { shading: a.shading }),
      drawTop: !a.startedEarlier,
      drawBottom: isLast,
    });
    a.startedEarlier = true;
    a.frag = null;
  };

  // Multi-column flow. A section's blocks fill column 0 top→bottom, then column
  // 1, etc.; the page finalizes only when the last column overflows. `bandTop`
  // is where columns start on the current page (top, or the handoff y after a
  // continuous section break); `sectionMaxY` tracks the deepest content so the
  // next continuous section can resume below it. count 1 ⇒ ordinary flow.
  let contentWidth = contentRight - contentLeft;
  let colCount = 1;
  let colGap = 0;
  let colWidth = contentWidth;
  // Per-column boxes when the section's columns are not all the same width
  // (w:cols/@w:equalWidth="0"); null keeps the uniform `colWidth + colGap`
  // arithmetic, which every other document uses.
  let colBoxes: ColumnBox[] | null = null;
  let colIndex = 0;
  let bandTop = top;
  let sectionMaxY = top;
  let colDirty = false; // current column holds content (a break would progress)
  /** True while placing a next-page section-break mark. It may still flow to
   *  the next column, but a break that would FINALIZE the page is suppressed
   *  — the mark is clipped at the floor instead (see blockSection). */
  let placingPageBreakMark = false;
  /**
   * Whether this band was opened by a CONTINUOUS section break rather than by
   * a page or a column.
   *
   * Word drops a paragraph's space-before at the top of a page or column, and
   * `colDirty` is how that is spelled here — nothing above to space away from.
   * A section resuming under the previous section's content is neither: it is
   * the middle of a page, and Word keeps the space there. The factsheet's own
   * PDF from Word measures it twice — its section 2 opens with a 155tw
   * space-before that Word honours (10.3px) and section 3 with an 89tw one
   * (5.9px), while both landed flush against the previous section here.
   */
  let bandOpensSection = false;
  // Space-after the previous block already put on the page. The next block's
  // space-before collapses against it (collapsedBefore) — OOXML takes the
  // MAXIMUM of the two, not their sum.
  let pendingAfter = 0;
  const xShift = () =>
    colBoxes ? colBoxes[colIndex].x : colIndex * (colWidth + colGap);
  const colX0 = () => contentLeft + xShift();
  const colX1 = () =>
    colX0() + (colBoxes ? colBoxes[colIndex].width : colWidth);
  const bump = () => {
    if (y > sectionMaxY) sectionMaxY = y;
  };
  const applyColumns = (cols: ColumnConfig) => {
    colCount = Math.max(1, cols.count);
    colGap = colCount > 1 ? cols.gap : 0;
    colWidth = (contentWidth - colGap * (colCount - 1)) / colCount;
    colBoxes = columnBoxes(cols);
  };

  /** Recompute the band for the current page/chrome and restart on it. Only
   *  ever called when the current page is empty. */
  const restartBand = () => {
    ({ top, bottom } = bandOf(curPage));
    colIndex = 0;
    bandTop = top;
    sectionMaxY = top;
    colDirty = false;
    bandOpensSection = false;
    y = top;
  };

  /**
   * Adopt a section's geometry WITHOUT restarting the page.
   *
   * Everything horizontal takes effect from this section's first line, which is
   * what a continuous break changing the left/right margins is for. Everything
   * vertical is deliberately left alone: a page already has a band, and Word
   * cannot apply a new top margin to a page it has begun. It doesn't need to be
   * remembered separately either — `finalizePage` re-reads `curPage` for the
   * next page's band, so the new vertical margins arrive exactly there.
   */
  const adoptGeometry = (p: PageConfig) => {
    curPage = p;
    contentLeft = contentLeftOf(p);
    contentRight = p.width - p.margin.right;
    contentWidth = contentRight - contentLeft;
  };

  /** Swap to a section's page geometry and restart the band on it. Only ever
   *  called at a page boundary (after finalizePage, or before any content). */
  const setGeometry = (p: PageConfig) => {
    adoptGeometry(p);
    restartBand();
  };

  // Footnotes referenced on the current page, in first-appearance order. Their
  // bodies are reserved at the page bottom; `limit()` is the body's lowered
  // floor once that space is committed.
  let pageFnNums: number[] = [];
  const pageFnSet = new Set<number>();
  const noteHeight = (n: number) => footnotes?.get(n)?.height ?? 0;
  const reservedFor = (nums: number[]) =>
    nums.length === 0
      ? 0
      : FOOTNOTE_AREA_GAP + nums.reduce((s, n) => s + noteHeight(n), 0);
  /** Bottom of the body band, lowered by the footnotes committed to this page. */
  const limit = () => bottom - reservedFor(pageFnNums);
  /** Footnote numbers a draft line references that have a known body. */
  const lineFnNums = (segs: LayoutSegment[]): number[] => {
    const out: number[] = [];
    for (const s of segs) {
      if (
        s.footnoteRef != null &&
        footnotes?.has(s.footnoteRef) &&
        !out.includes(s.footnoteRef)
      ) {
        out.push(s.footnoteRef);
      }
    }
    return out;
  };
  /** Commit footnote numbers to the current page (reserve bottom space). */
  const commitFns = (nums: number[]) => {
    for (const n of nums) {
      if (!pageFnSet.has(n)) {
        pageFnNums.push(n);
        pageFnSet.add(n);
      }
    }
  };
  /** Extra bottom space new footnotes would add beyond what's reserved. */
  const addedReserve = (nums: number[]) => {
    const fresh = nums.filter((n) => !pageFnSet.has(n));
    return fresh.length === 0
      ? 0
      : reservedFor([...pageFnNums, ...fresh]) - reservedFor(pageFnNums);
  };
  /** Fresh footnote numbers referenced in a table's cells above `maxY` (cell-y
   *  relative to the table top) — so a placed fragment reserves its notes. */
  const tableFootnoteNums = (t: ResolvedTable, maxY = Infinity): number[] => {
    const out: number[] = [];
    const scan = (ls: LayoutLine[]) => {
      for (const l of ls)
        for (const s of l.segments) {
          const n = s.footnoteRef;
          if (
            n != null &&
            footnotes?.has(n) &&
            !pageFnSet.has(n) &&
            !out.includes(n)
          )
            out.push(n);
        }
    };
    for (const c of t.cells) {
      if (c.y >= maxY) continue;
      scan(c.lines);
      if (c.tables)
        for (const nt of c.tables) for (const nc of nt.cells) scan(nc.lines);
    }
    return out;
  };
  /** Fresh footnote numbers in a table remainder — the repeated-header band
   *  (when it will be prepended), the carry, and the unconsumed base rows —
   *  WITHOUT materializing it: the fits-whole check runs every page. */
  const remainderFootnoteNums = (
    r: TableRemainder,
    ghostH: number,
  ): number[] => {
    const cells: ResolvedCell[] = [];
    if (ghostH > 0)
      for (const c of r.base.cells) if (c.y + c.height <= ghostH) cells.push(c);
    if (r.carry) cells.push(...r.carry);
    for (const c of r.base.cells)
      if (c.y >= r.fromY - REMAINDER_EPS) cells.push(c);
    return tableFootnoteNums({ ...r.base, cells });
  };

  // Column balancing: on a multi-column section's final page, fill every column
  // to an even target height instead of packing column 0 first. `balanceTarget`
  // is that per-column height (non-last columns break there; the last column
  // keeps the full band to absorb rounding). null ⇒ greedy full-height columns.
  let balanceTarget: number | null = null;
  let sectionRemaining = 0; // unplaced height left in the current section
  let balancing = false;
  const colBottom = () =>
    balanceTarget != null && colIndex < colCount - 1
      ? Math.min(limit(), bandTop + balanceTarget)
      : limit();
  /** Recompute the balance target for the current page: balance only when the
   *  section's remaining content fits this page's columns. */
  const rebalance = () => {
    if (!balancing) {
      balanceTarget = null;
      return;
    }
    const pageColH = limit() - bandTop;
    balanceTarget =
      sectionRemaining <= colCount * pageColH
        ? Math.max(1, sectionRemaining / colCount)
        : null;
  };

  /** Lay the page's reserved footnote bodies out at the bottom, above the
   *  footer, with a separator rule. Positions are already stripped. */
  const buildFootnoteArea = (nums: number[]): ResolvedFootnotes => {
    const areaTop = bottom - reservedFor(nums);
    let fy = areaTop + FOOTNOTE_AREA_GAP;
    const out: LayoutLine[] = [];
    for (const n of nums) {
      const body = footnotes?.get(n);
      if (!body) continue;
      for (const l of body.lines) out.push({ ...l, y: l.y + fy });
      fy += body.height;
    }
    return { separatorY: areaTop + FOOTNOTE_AREA_GAP / 2, lines: out };
  };

  const finalizePage = () => {
    // Fixed-point guard: a seeded exclusion whose float never re-registered
    // on the replayed page (the exclusion pushed its own anchor onto the
    // NEXT page) is a phantom — retract it and replay once more. Skipped on
    // the final attempt so termination is unconditional.
    if (seeds.size > 0 && retriesLeft > 0) {
      const phantom = [...seeds.keys()].filter((k) => !registeredKeys.has(k));
      if (phantom.length > 0) throw new PageRetrySignal([], phantom);
    }
    flushPBdrFrag(false); // a bordered paragraph continuing → open-bottom box
    const resolved: ResolvedPage = {
      index: pages.length,
      width: curPage.width,
      height: curPage.height,
      contentLeft,
      contentTop: top,
      lines,
    };
    if (tables.length > 0) resolved.tables = tables;
    if (pageFloats.length > 0) resolved.floats = pageFloats;
    if (pageFnNums.length > 0)
      resolved.footnotes = buildFootnoteArea(pageFnNums);
    if (paraBoxes.length > 0) resolved.paraBoxes = paraBoxes;
    if (pageChromeIndex != null) {
      resolved.chromeIndex = pageChromeIndex;
      resolved.sectionFirst = pageSectionFirst;
    }
    // Page-cache recording: which nodes fed this page, the carry it opened
    // with, and whether its last LISTED item continues on the next page —
    // an in-flight item that placed nothing here isn't an overlap.
    if (pageCache) {
      const inflight = midAny() ? items[idx]?.node : undefined;
      record.push({
        itemNodes: pageNodes,
        endsMidItem: !!inflight && pageNodes[pageNodes.length - 1] === inflight,
        cp: currentPageCp,
        startOffset: pageFirstOffset ?? 0,
        page: resolved,
      });
    }
    // The next page starts under the CURRENT section's chrome — recompute the
    // band so a section with taller/shorter chrome gets its own bounds even
    // when the geometry didn't change.
    pageChromeIndex = curChromeIndex;
    pageSectionFirst = sectionFirstPending;
    sectionFirstPending = false;
    ({ top, bottom } = bandOf(curPage));
    paraBoxes = [];
    pages.push(resolved);
    lines = [];
    tables = [];
    pageFloats = [];
    exclusions = [];
    pageFnNums = [];
    pageFnSet.clear();
    colIndex = 0;
    bandTop = top;
    sectionMaxY = top;
    colDirty = false;
    bandOpensSection = false;
    y = top;
    // A page closed cleanly: its seeds are spent, and the NEXT page gets a
    // fresh retry budget and its own checkpoint (all accumulators are empty
    // here, so the checkpoint is a handful of scalars).
    seeds.clear();
    droppedSeeds.clear();
    seedPins.clear();
    registeredKeys.clear();
    retriesLeft = MAX_PAGE_RETRIES;
    pageStartCp = takeCheckpoint();
    if (pageCache) {
      seedRecording(pageStartCp);
      // Tail resync: this fresh page opens at a boundary whose carry and
      // remaining items match a page of the previous run — every page from
      // there on is reusable (shifted by the position delta) instead of
      // re-placed. Clean boundaries compare by scalars alone; a page opening
      // MID-TABLE also compares the remainder cursor, so an edit above a
      // long table splices its unchanged pages back instead of re-splitting
      // to the table's end. The splice signal unwinds to the page loop.
      if (prior && idx < items.length) {
        const n = items[idx]?.node;
        if (n && !midAny()) {
          const k = priorFirstNode.get(n);
          if (k !== undefined) {
            const e = prior.entries[k];
            if (carryEq(pageStartCp, e.cp) && tailMatches(k)) {
              throw new TailSpliceSignal(k);
            }
          }
        } else if (
          n &&
          midTableRest !== null &&
          midParaDraftIdx === null &&
          !inBandedPara
        ) {
          for (const k of priorMidTableEntries.get(n) ?? []) {
            const e = prior.entries[k];
            if (midTableCarryEq(pageStartCp, e.cp) && tailMatches(k)) {
              throw new TailSpliceSignal(k);
            }
          }
        }
      }
    }
  };

  const takeCheckpoint = (): PageCheckpoint => ({
    itemIdx: idx,
    midTable: midTableRest && snapshotRemainder(midTableRest),
    midDraftIdx: midParaDraftIdx,
    midBandedFrom:
      inBandedPara && midBandedFrom !== UNRESUMABLE ? midBandedFrom : null,
    unreplayable: inBandedPara && midBandedFrom === UNRESUMABLE,
    firstItem,
    curPage,
    curChromeIndex,
    pageChromeIndex,
    pageSectionFirst,
    sectionFirstPending,
    contentLeft,
    contentRight,
    contentWidth,
    colCount,
    colGap,
    colWidth,
    colBoxes,
    balancing,
    sectionRemaining,
    activePBdr: activePBdr && {
      ...activePBdr,
      frag: activePBdr.frag && { ...activePBdr.frag },
    },
  });

  /** Rewind to the page-start checkpoint for a replay: scalars restored,
   *  page accumulators emptied, the seeded exclusions installed, and the
   *  mid-item resume handles armed (fresh clones — the checkpoint stays
   *  valid across multiple retries). */
  const restoreCheckpoint = (cp: PageCheckpoint) => {
    firstItem = cp.firstItem;
    curPage = cp.curPage;
    curChromeIndex = cp.curChromeIndex;
    pageChromeIndex = cp.pageChromeIndex;
    pageSectionFirst = cp.pageSectionFirst;
    sectionFirstPending = cp.sectionFirstPending;
    contentLeft = cp.contentLeft;
    contentRight = cp.contentRight;
    contentWidth = cp.contentWidth;
    colCount = cp.colCount;
    colGap = cp.colGap;
    colWidth = cp.colWidth;
    colBoxes = cp.colBoxes;
    balancing = cp.balancing;
    sectionRemaining = cp.sectionRemaining;
    activePBdr = cp.activePBdr && {
      ...cp.activePBdr,
      frag: cp.activePBdr.frag && { ...cp.activePBdr.frag },
    };
    ({ top, bottom } = bandOf(curPage));
    lines = [];
    tables = [];
    pageFloats = [];
    paraBoxes = [];
    pageFnNums = [];
    pageFnSet.clear();
    registeredKeys.clear();
    exclusions = [...seeds.values()].map((r) => ({ ...r }));
    colIndex = 0;
    bandTop = top;
    sectionMaxY = top;
    colDirty = false;
    bandOpensSection = false;
    pendingAfter = 0; // nothing above to collapse against at a fresh band
    y = top;
    rebalance();
    resumeTable = cp.midTable && materializeRemainder(cp.midTable);
    resumeDraftIdx = cp.midDraftIdx;
    resumeBandedFrom = cp.midBandedFrom;
    midTableRest = null;
    midParaDraftIdx = null;
    midBandedFrom = null;
    inBandedPara = false;
    if (pageCache) seedRecording(cp);
  };

  /** End the current column: move to the next column, or finalize the page when
   *  the last column is full. */
  const breakBand = () => {
    bump();
    flushPBdrFrag(false); // fragment ends at the column/page boundary
    if (colIndex < colCount - 1) {
      colIndex++;
      colDirty = false;
      bandOpensSection = false;
      y = bandTop;
    } else {
      finalizePage();
      rebalance(); // fresh page within a balanced section → new target
    }
  };

  /** Place one finished line, reserving bottom space for any footnotes it
   *  references. Breaks to the next column/page when the line + its (and the
   *  page's already-committed) footnotes no longer fit, or it would cross the
   *  balance target. */
  /** Record that the in-flight item put content on the CURRENT page — the
   *  page cache matches pages by exactly these node lists. Called after any
   *  internal band break has resolved, so it names the right page. */
  const noteNodePlaced = () => {
    const it = items[idx];
    if (!it?.node) return;
    if (pageNodes[pageNodes.length - 1] !== it.node) pageNodes.push(it.node);
    pageFirstOffset ??= it.nodeOffset ?? null;
  };

  /**
   * Place one wrapped line at the cursor.
   *
   * `preShifted` says the draft already carries absolute x: the banded path
   * wraps against `bandAt`, which starts from `colX0()`, whereas the
   * pre-wrapped drafts are measured from the content-area left edge and have
   * to be slid into the current column. Adding the shift to a banded line
   * counts the column offset twice — invisible while the banded path only
   * ever ran in single-column flow (xShift 0), and a 1600px x the moment a
   * multi-column section used it.
   */
  const emitLine = (draft: LineDraft, preShifted = false) => {
    let add = lineFnNums(draft.segments).filter((n) => !pageFnSet.has(n));
    const floor = () =>
      Math.min(colBottom(), bottom - reservedFor([...pageFnNums, ...add]));
    if (
      y + draft.height > floor() &&
      colDirty &&
      !(
        placingPageBreakMark &&
        colIndex >= colCount - 1 &&
        y + draft.height - floor() <= PAGE_BREAK_MARK_POKE
      )
    ) {
      breakBand(); // next column, or next page; footnotes ride the page
      add = lineFnNums(draft.segments).filter((n) => !pageFnSet.has(n));
    }
    noteNodePlaced();
    lines.push(draftToLine(draft, y, preShifted ? 0 : xShift()));
    if (activePBdr) {
      if (!activePBdr.frag) {
        activePBdr.frag = { x0: colX0(), x1: colX1(), y0: y, y1: y };
      }
      activePBdr.frag.y1 = y + draft.height;
    }
    colDirty = true;
    y += draft.height;
    sectionRemaining -= draft.height;
    bump();
    commitFns(add);
  };

  /** Whether the current page already holds content (so a break is meaningful). */
  const pageHasContent = () =>
    lines.length > 0 || tables.length > 0 || pageFloats.length > 0;

  // ── Pagination keeps (w:keepNext / w:keepLines / w:widowControl) ────
  const nominalH = (() => {
    const bm = ctx.metrics ? ctx.metrics(ctx.base) : null;
    return bm ? bm.ascent + bm.descent : sizePx(ctx.base) * LINE_HEIGHT_FACTOR;
  })();
  const draftsTotal = (d: LineDraft[] | null | undefined): number =>
    d?.reduce((s, x) => s + x.height, 0) ?? 0;

  /** The smallest slice of a paragraph that may legally open a band: its
   *  first line, or the first TWO with widow control on (orphan rule) — and
   *  the whole thing when it cannot split at all (keepLines, or ≤3 lines
   *  under widow control, where every split strands a lone line). */
  const minChunkH = (p: ParaItem): number => {
    const d = p.drafts;
    if (!d || d.length === 0) return nominalH;
    if (p.keepLines) return draftsTotal(d);
    if (p.widowControl === false) return d[0].height;
    if (d.length <= 3) return draftsTotal(d);
    return d[0].height + d[1].height;
  };

  /** Height that must follow a keepNext paragraph on its band: consecutive
   *  keepNext blocks in full, then the minimum legal slice of the block that
   *  ends the chain. Tables and un-drafted paragraphs approximate as one
   *  line. */
  const keepAheadH = (idx: number, prevAfter = 0): number => {
    let need = 0;
    for (let j = idx + 1; j < items.length; j++) {
      const it = items[j];
      if (it.section) break; // a section boundary ends the chain
      if ('para' in it) {
        need += collapsedBefore(it.para.before, prevAfter);
        if (it.para.keepNext && it.para.drafts && j + 1 < items.length) {
          need += draftsTotal(it.para.drafts) + (it.para.after ?? 0);
          continue;
        }
        need += minChunkH(it.para);
      } else {
        need += nominalH;
      }
      break;
    }
    return need;
  };

  /** Emit a paragraph's pre-wrapped lines with widow/orphan control: when
   *  the paragraph must split, neither side of the split keeps a lone line
   *  (Word's w:widowControl, on by default). Splits force the band break
   *  here; emitLine's own overflow check stays as the footnote-reserve
   *  fallback. */
  const emitParaDrafts = (
    drafts: LineDraft[],
    widowOn: boolean,
    startI = 0, // a replayed page resumes mid-paragraph here
  ): void => {
    let i = startI;
    while (i < drafts.length) {
      const remaining = drafts.length - i;
      let fit = 0;
      for (
        let yy = y;
        fit < remaining && yy + drafts[i + fit].height <= colBottom();
      ) {
        yy += drafts[i + fit].height;
        fit++;
      }
      if (fit >= remaining) {
        for (; i < drafts.length; i++) {
          // emitLine may still break on footnote reserve. i === 0 means
          // NOTHING of this paragraph placed yet — that boundary is "before
          // the item" (resume re-runs its pre-steps), not mid-item.
          midParaDraftIdx = i === 0 ? null : i;
          emitLine(drafts[i]);
        }
        midParaDraftIdx = null;
        return;
      }
      if (widowOn && remaining >= 2) {
        if (remaining - fit === 1) fit = remaining - 2; // no widow up top
        if (i === 0 && fit === 1) fit = 0; // no orphan down here
        if (fit < 0) fit = 0;
      }
      if (fit === 0 && !colDirty) fit = 1; // an empty band must progress
      for (let k = 0; k < fit; k++, i++) {
        // If placing this line (or the chunk break below) closes the page,
        // a replay resumes at the first UNPLACED draft — emitLine breaks
        // before it pushes, so the current index is exactly that. Index 0
        // stays null: that boundary is "before the item", not mid-item.
        midParaDraftIdx = i === 0 ? null : i;
        emitLine(drafts[i]);
      }
      midParaDraftIdx = i === 0 ? null : i;
      breakBand();
      midParaDraftIdx = null;
    }
  };

  /** Pin a paragraph's floats relative to its start; register text exclusions.
   *  `anchorH` is the anchor line's nominal height — the slice of the page the
   *  anchor occupies, and therefore the slice whose band a column-relative
   *  offset is measured from. */
  const registerFloats = (
    flow: FlowParagraph,
    yPara: number,
    anchorH: number,
  ) => {
    // The column as the ANCHOR LINE sees it, not the bare column box.
    //
    // Word measures a column-relative offset from the text band the anchor
    // got, which is the column narrowed by whatever floats that line already
    // wraps around. Read as the bare column, two of the corpus factsheet's
    // pictures land in the wrong place, and Word's own PDF says by how much:
    // one wants origin 712 where the column starts at 673 — exactly the right
    // edge of the float above it (700) plus that float's 12px distR — and two
    // more want 359 where the column starts at 9, again a float's right edge
    // (347) plus 12. With 359 those two land at 421 and 11, which is where
    // Word has them to the pixel; from the column they land at 71 and −339,
    // the second one entirely off the page.
    //
    // Computed ONCE, before any of this paragraph's own floats join the
    // exclusions — a paragraph anchoring several pictures measures them all
    // from the same band, not each from the one its siblings just narrowed.
    // Indents are excluded deliberately: a paragraph in the factsheet carries
    // an 8px left indent and Word still measures its pictures from 9, the
    // column edge.
    const anchorBand = bandAt(yPara, anchorH, { left: 0, right: 0 });
    for (const f of flow.floats ?? []) {
      // `column` is the text column the anchor paragraph sits in — this runs
      // inside the placer, so colX0/colX1 already name it. Reading it as the
      // content box put six pictures of a two- and three-column factsheet off
      // the page, one of them entirely (x −339..−3).
      const baseL =
        f.hRel === 'page'
          ? 0
          : f.hRel === 'column'
            ? (anchorBand?.left ?? colX0())
            : contentLeft;
      const baseR =
        f.hRel === 'page'
          ? curPage.width
          : f.hRel === 'column'
            ? (anchorBand?.right ?? colX1())
            : contentRight;
      const fx =
        f.hAlign === 'right'
          ? baseR - f.width
          : f.hAlign === 'center'
            ? (baseL + baseR - f.width) / 2
            : f.hAlign === 'left'
              ? baseL
              : baseL + (f.hOffset ?? 0);
      let fy =
        f.vRel === 'page'
          ? (f.vOffset ?? 0)
          : f.vRel === 'margin'
            ? top + (f.vOffset ?? 0)
            : yPara + (f.vOffset ?? 0);
      let fxEff = fx;
      const pin = f.pos !== undefined ? seedPins.get(f.pos) : undefined;
      if (pin) {
        fxEff = pin.x;
        fy = pin.y;
      }
      pageFloats.push({
        ...resolveFloat(f, fxEff, fy, ctx),
        // Effective offsets for drag-to-move: what hOffset/vOffset would put
        // the float at exactly this spot. For an hAlign float this is the
        // alignment resolved to a number, which is what a drag pins it to.
        effHOffset: fxEff - baseL,
        effVOffset: f.vOffset ?? 0,
      });
      colDirty = true;
      const rect: Exclusion | null =
        f.wrap === 'square'
          ? {
              left: fxEff - (f.distL ?? 0),
              right: fxEff + f.width + (f.distR ?? 0),
              top: fy - (f.distT ?? 0),
              bottom: fy + f.height + (f.distB ?? 0),
              ...(f.through && { through: true }),
            }
          : f.wrap === 'topAndBottom'
            ? {
                left: -Infinity,
                right: Infinity,
                top: fy - (f.distT ?? 0),
                bottom: fy + f.height + (f.distB ?? 0),
              }
            : null; // 'none' paints only
      if (rect) {
        const key = `${Math.round(rect.left)}:${Math.round(rect.top)}:${Math.round(rect.right)}:${Math.round(rect.bottom)}`;
        registeredKeys.add(key);
        // Word wraps the WHOLE page around a float, wherever its anchor
        // sits. Content already placed ABOVE this float was laid out
        // without knowing it — replay the page with the exclusion seeded.
        // (Not when the page opened mid-wrap of a float paragraph — that
        // boundary can't be resumed; documented limitation.)
        if (
          retriesLeft > 0 &&
          !seeds.has(key) &&
          !droppedSeeds.has(key) &&
          !pageStartCp.unreplayable &&
          overlapsPlaced(rect)
        ) {
          if (f.pos !== undefined)
            seedPins.set(f.pos, { x: fxEff, y: fy, key });
          throw new PageRetrySignal([{ key, rect }]);
        }
        exclusions.push(rect);
      }
    }
  };

  /** Whether content already placed on this page intersects `rect` — the
   *  page-replay trigger. Lines and table fragments both count. */
  const overlapsPlaced = (rect: Exclusion): boolean =>
    lines.some(
      (l) =>
        l.y < rect.bottom &&
        l.y + l.height > rect.top &&
        l.x < rect.right &&
        l.x + l.width > rect.left,
    ) ||
    tables.some(
      (t) =>
        t.y < rect.bottom &&
        t.y + t.height > rect.top &&
        t.x < rect.right &&
        t.x + t.width > rect.left,
    );

  /** Widest text band at [yy, yy+h) after carving out the exclusions; null
   *  when nothing usable remains (the caller skips below the blocker). */
  const bandAt = (
    yy: number,
    h: number,
    indents: { left: number; right: number },
  ): { left: number; right: number } | null => {
    let L = colX0() + indents.left;
    let R = colX1() - indents.right;
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
  const placeParaBanded = (flow: FlowParagraph, spaceBefore = 0) => {
    // Wrapping is one uninterruptible call: a page boundary inside it has no
    // resumable midpoint, so checkpoints taken while this flag is up are
    // marked unreplayable and float triggers on that page stand down.
    inBandedPara = true;
    try {
      placeParaBandedInner(flow, spaceBefore);
    } finally {
      inBandedPara = false;
      midBandedFrom = null;
    }
  };
  const placeParaBandedInner = (flow: FlowParagraph, spaceBefore: number) => {
    // A paragraph's floats belong on the page its FIRST LINE lands on — but
    // that page-break decision normally happens inside the band callback,
    // AFTER registerFloats has already pushed the floats into the CURRENT
    // page at the current cursor. A paragraph pushed to the next page then
    // left its floats behind, positioned in the previous page's space (the
    // "image bounced to the top of page 1 at page-2's offsets" report). So
    // for float-carrying paragraphs, settle the start first: run the same
    // break/skip loop the band callback would, with a nominal line height —
    // deliberately WITHOUT this paragraph's own floats, whose exclusions
    // don't exist yet (same rule as Word: the anchor's page is decided by
    // the text position, then the drawing follows the anchor).
    // The paragraph's OWN nominal line height, seeded the same way
    // wrapParagraph seeds it (the mark's font, zero for a break mark). It has
    // to be: the through-float walk below advances on this paragraph's LINE
    // GRID, and a step in someone else's pitch puts every landing off — 14pt
    // marks stepped at the 11pt document default even slipped PAST a float's
    // top edge and never saw the blocker at all.
    const seed = flow.markFont ? { ...ctx.base, ...flow.markFont } : ctx.base;
    const bm =
      flow.runs.length === 0 && flow.breakMark
        ? { ascent: 0, descent: 0, leading: 0 }
        : ctx.metrics
          ? ctx.metrics(seed)
          : null;
    const estH = bm
      ? bm.ascent + bm.descent + (bm.leading ?? 0)
      : sizePx(seed) * LINE_HEIGHT_FACTOR;
    // Where the paragraph BELONGS. The loop below walks `y` down past floats
    // to find room for its first LINE, but a float's wrap does not move the
    // paragraph — Word anchors a paragraph-relative drawing at the position
    // the paragraph would have had. The probe pins it to 0.02px, at two
    // different shove distances. A real column/page change is different:
    // that relocates the paragraph itself, so the anchor follows.
    let anchorY = y;
    if (flow.floats && flow.floats.length > 0) {
      for (;;) {
        if (y + estH > colBottom() && colDirty) {
          breakBand();
          anchorY = y; // a new band: the paragraph really did move
          continue;
        }
        // Only asking "is there any room at y" — the paragraph's own indents
        // do not decide where its FLOATS may sit.
        if (bandAt(y, estH, { left: 0, right: 0 })) break;
        // Same column gate as the band callback: a float in another column
        // is not a blocker here.
        const blockers = exclusions.filter(
          (ex) =>
            ex.top < y + estH &&
            ex.bottom > y &&
            ex.left < colX1() &&
            ex.right > colX0(),
        );
        if (blockers.length === 0) break;
        // Two skip modes, measured apart by a generated probe read from
        // Word's own PDF:
        //   square  — the line restarts FLUSH below the float.
        //   through — the line stays on the paragraph's own LINE GRID: whole
        //             line-heights from where it stood. Through-wrap means
        //             the lines still exist across the float's span (each
        //             fully blocked horizontally), so the first free slot is
        //             a grid position, not the float's edge. Three cases pin
        //             it to 0.1px, at three different shove distances.
        // And a through float takes the paragraph WITH it — the anchor
        // follows the line — while a square float displaces the line only.
        if (blockers.some((ex) => ex.through) && estH > 0.5) {
          y += estH;
          anchorY = y;
        } else {
          y = Math.min(...blockers.map((ex) => ex.bottom));
        }
      }
    }
    // A paragraph-relative anchor measures from the TOP OF THE PARAGRAPH, and
    // a paragraph begins where its space-before begins — not where its first
    // line does. Word's PDF of a one-variable probe settles it: an anchor with
    // posOffset 0 on a paragraph carrying 248tw of space-before lands within
    // 0.9px of the paragraph's top and 17px above its first line. `y` here has
    // the space already added (and any float-skip applied to both), so the
    // paragraph's top is that much higher.
    registerFloats(flow, anchorY - spaceBefore, estH);
    bandedEndY = null;
    wrapParagraph(flow, ctx, bandedBandFn, emitBandedLine, undefined, () => {
      if (colDirty) breakBand();
    });
    // A paragraph ends at the bottom of the last line it actually drew.
    //
    // wrapParagraph asks the band callback for one more line after its last
    // token, and bandedBandFn walks the shared `y` down past any float in the
    // way — looking for room for a line that never comes. Left there, that
    // speculative walk became the NEXT paragraph's starting position, so a
    // paragraph anchoring a float shoved the paragraph after it below its own
    // float. Word does not: on a generated probe the following paragraph sat
    // at the previous paragraph's bottom, 104px above where the walk had left
    // us — and 204px when the float was twice as tall.
    //
    // Rewind only within the band the line was drawn in: a genuine column or
    // page change means the paragraph really ended there.
    if (bandedEndY !== null && bandedEndBand === bandTag()) y = bandedEndY;
  };

  const bandedBandFn: BandFn = (estH, indents, minWidth) => {
    for (;;) {
      if (
        y + estH > colBottom() &&
        colDirty &&
        !(placingPageBreakMark && colIndex >= colCount - 1 && y < colBottom())
      ) {
        // (For a section-break mark the loose y-test above only defers the
        // decision — emitLine re-judges with the line's real BASELINE.)
        breakBand(); // next column/page: exclusions are gone
        continue;
      }
      const b = bandAt(y, estH, indents);
      // Only floats that intrude on THIS column block it — one wholly in a
      // neighbouring column shares the y-range but not the band, and letting
      // it into the list made a through float in column 1 grid-step the
      // lines of column 2.
      const blockers = exclusions.filter(
        (ex) =>
          ex.top < y + estH &&
          ex.bottom > y &&
          ex.left < colX1() &&
          ex.right > colX0(),
      );
      // A band that passed the MIN_BAND floor can still be too narrow for
      // an unbreakable item (an inline image). While a float is what
      // narrowed it, keep walking down past the floats — same move as a
      // null band. With no blockers left the band is the full column;
      // return it even if the item is wider (it overflows, as before).
      const column = { left: colX0(), right: colX1() };
      if (
        b &&
        (minWidth === undefined ||
          b.right - b.left >= minWidth ||
          blockers.length === 0)
      )
        return { ...b, column };
      if (blockers.length === 0)
        return {
          left: column.left + indents.left,
          right: column.right - indents.right,
          column,
        };
      // Same two skip modes as the anchor walk above: flush below a square
      // float, whole line-heights below a through one (see the probe notes
      // there). estH here is the line being placed.
      if (blockers.some((ex) => ex.through) && estH > 0.5) y += estH;
      else y = Math.min(...blockers.map((ex) => ex.bottom)); // skip below
    }
  };

  /** Emit one wrapped line of a float-anchoring paragraph, keeping the
   *  mid-wrap resume position current: if placing THIS line closes the page,
   *  a replayed page re-wraps the paragraph from this line's start. Lines
   *  without a PM position (flat-API flows) can't be resumed — the sentinel
   *  marks the checkpoint unreplayable, the pre-existing degraded mode. */
  const emitBandedLine = (draft: LineDraft): void => {
    midBandedFrom = draft.from ?? UNRESUMABLE;
    emitLine(draft, true);
    bandedEndY = y; // emitLine leaves y at this line's bottom
    bandedEndBand = bandTag();
  };
  /** Identity of the band lines are landing in, so the rewind below can tell
   *  "the column we drew on" from "a column the band search wandered into". */
  const bandTag = () => pages.length * 1024 + colIndex;
  let bandedEndY: number | null = null;
  let bandedEndBand = -1;

  /** Re-enter a float paragraph split by the replayed page's start: its
   *  floats stayed on the earlier page (they belong where the first line
   *  landed), so only the remaining lines re-wrap — against this page's
   *  exclusions, seeds included. */
  const resumeParaBanded = (flow: FlowParagraph, fromPos: number) => {
    inBandedPara = true;
    try {
      wrapParagraph(flow, ctx, bandedBandFn, emitBandedLine, fromPos, () => {
        if (colDirty) breakBand();
      });
    } finally {
      inBandedPara = false;
      midBandedFrom = null;
    }
  };

  let firstItem = true;
  pageStartCp = takeCheckpoint();
  if (pageCache) seedRecording(pageStartCp);
  let startIdx = 0;

  // ── Prefix reuse (Phase 3) ──────────────────────────────────────────
  // Pages whose item nodes are reference-identical to the previous run are
  // the same pages — copy them and resume placement at the first difference.
  // The boundary page itself is re-placed (k − 1): its break decisions may
  // have peeked past the boundary (keepNext look-ahead).
  if (prior && prior.entries.length > 1) {
    let cursor = 0;
    let k = 0;
    for (; k < prior.entries.length; k++) {
      const e = prior.entries[k];
      if (e.itemNodes.length === 0) break;
      let c = cursor;
      let ok = true;
      for (let j = 0; j < e.itemNodes.length; j++, c++) {
        if (items[c]?.node !== e.itemNodes[j]) {
          ok = false;
          break;
        }
      }
      if (!ok) break;
      cursor = e.endsMidItem ? c - 1 : c;
    }
    k = Math.max(0, k - 1);
    const cp = prior.entries[k]?.cp;
    if (k > 0 && cp && !cp.unreplayable) {
      for (let i = 0; i < k; i++) {
        pages.push(prior.entries[i].page);
        record.push(prior.entries[i]);
        perf.bump('page.reuse');
      }
      restoreCheckpoint(cp);
      pageStartCp = cp;
      startIdx = cp.itemIdx;
    }
  }

  retry: for (;;) {
    try {
      placeRun(startIdx);
      break retry;
    } catch (e) {
      if (e instanceof TailSpliceSignal) {
        // Splice the previous run's tail in, shifted by the PM delta the
        // edit introduced. Entries are re-based so the NEXT pass can reuse
        // them again without re-placing.
        const entries = (prior as PageRun).entries;
        const first = entries[e.entryIdx];
        const delta = (items[idx].nodeOffset ?? 0) - first.startOffset;
        const idxDelta = idx - first.cp.itemIdx;
        for (let i = e.entryIdx; i < entries.length; i++) {
          const pe = entries[i];
          const newIndex = pages.length;
          const asIs =
            delta === 0 && idxDelta === 0 && pe.page.index === newIndex;
          const pg = asIs
            ? pe.page
            : shiftResolvedPage(pe.page, delta, newIndex);
          pages.push(pg);
          record.push(
            asIs
              ? pe
              : {
                  itemNodes: pe.itemNodes,
                  endsMidItem: pe.endsMidItem,
                  cp: rebaseCp(pe.cp, delta, idxDelta),
                  startOffset: pe.startOffset + delta,
                  page: pg,
                },
          );
          perf.bump('page.splice');
        }
        break retry;
      }
      if (!(e instanceof PageRetrySignal)) throw e;
      retriesLeft--;
      for (const a of e.add) if (!seeds.has(a.key)) seeds.set(a.key, a.rect);
      for (const k of e.drop) {
        seeds.delete(k);
        droppedSeeds.add(k);
        for (const [pp, pn] of seedPins) if (pn.key === k) seedPins.delete(pp);
      }
      restoreCheckpoint(pageStartCp);
      startIdx = pageStartCp.itemIdx;
    }
  }
  if (pageCache) {
    pageCache.store.delete(pageCache.key); // re-insert as most recent
    pageCache.store.set(pageCache.key, { entries: record });
    while (pageCache.store.size > 2) {
      const oldest = pageCache.store.keys().next().value;
      if (oldest === undefined) break;
      pageCache.store.delete(oldest);
    }
  }
  return { pages };

  function placeRun(from: number): void {
    for (idx = from; idx < items.length; idx++) {
      const item = items[idx];
      placingPageBreakMark = false; // per-item; several paths continue early
      // Section boundary: switch column flow (and break) before the block.
      if (item.section) {
        const nextPage = item.section.page ?? page;
        const geomChanges = !sameGeom(nextPage, curPage);
        // Only a different SHEET forces a page; different margins on the same
        // sheet do not (see sameSheet).
        const sheetChanges = !sameSheet(nextPage, curPage);
        curChromeIndex = item.section.chromeIndex;
        if (firstItem) {
          pageChromeIndex = curChromeIndex;
          pageSectionFirst = true;
          if (geomChanges) setGeometry(nextPage);
          else restartBand(); // page is empty — adopt this section's chrome band
          applyColumns(item.section);
        } else if (item.section.newPage || sheetChanges) {
          // The sheet may only change at a page boundary — a continuous break
          // onto a different page size is laid out as next-page (Word's
          // promotion). A margin change alone falls through to the branch
          // below and keeps the page.
          sectionFirstPending = true;
          if (pageHasContent()) {
            finalizePage(); // its tail re-bands for the new section's chrome
          } else {
            // The current (empty) page becomes this section's first page.
            pageChromeIndex = curChromeIndex;
            pageSectionFirst = true;
            sectionFirstPending = false;
            restartBand();
          }
          if (geomChanges) setGeometry(nextPage);
          applyColumns(item.section);
        } else {
          // Continuous break: resume below the finishing section's deepest column,
          // unless that's already at the band floor (then start a fresh page).
          bump();
          if (sectionMaxY >= limit()) {
            sectionFirstPending = true; // the fresh page opens this section
            finalizePage();
            // A page boundary after all, so this section's margins apply whole.
            if (geomChanges) setGeometry(nextPage);
          } else {
            // Same page, possibly new margins: the horizontal ones take effect
            // from here, the vertical ones from the next page.
            if (geomChanges) adoptGeometry(nextPage);
            bandTop = sectionMaxY;
            y = bandTop;
            colIndex = 0;
            colDirty = false;
            bandOpensSection = true; // mid-page: space-before still applies
          }
          applyColumns(item.section);
        }
        // Balance the new section's columns once its content fits a page —
        // but ONLY a section that ends in a continuous break, which is the
        // whole reason people insert one. The document's last section fills
        // column 1 to the bottom and spills into column 2, and so does a
        // section that ends in a page break. A manual column break also turns
        // balancing off: it is the author saying where the column ends.
        balancing =
          colCount > 1 &&
          !item.section.hasColumnBreak &&
          item.section.endsContinuous === true;
        sectionRemaining = item.section.height ?? 0;
        rebalance();
      }
      firstItem = false;

      if ('para' in item) {
        // A replayed page opening mid-way through a FLOAT paragraph's wrap:
        // its floats live on the earlier page; re-wrap only the remaining
        // lines against this page's exclusions (seeds included).
        const resumeBanded = resumeBandedFrom;
        resumeBandedFrom = null;
        if (resumeBanded !== null) {
          resumeParaBanded(item.para.getFlow(), resumeBanded);
          if (activePBdr) {
            flushPBdrFrag(true);
            activePBdr = null;
          }
          if (item.para.after) {
            y += item.para.after;
            sectionRemaining -= item.para.after;
          }
          pendingAfter = item.para.after ?? 0;
          continue;
        }
        // A replayed page opening mid-paragraph: its pre-steps (page break,
        // keeps, space-before, border box creation) already ran before the
        // boundary — only the remaining drafts are placed.
        const resumeAt = resumeDraftIdx;
        resumeDraftIdx = null;
        if (resumeAt !== null && item.para.drafts) {
          emitParaDrafts(
            item.para.drafts,
            item.para.widowControl !== false,
            resumeAt,
          );
          if (activePBdr) {
            flushPBdrFrag(true);
            activePBdr = null;
          }
          if (item.para.after) {
            y += item.para.after;
            sectionRemaining -= item.para.after;
          }
          pendingAfter = item.para.after ?? 0;
          continue;
        }
        if (item.para.pageBreakBefore && pageHasContent()) {
          finalizePage();
          rebalance();
        }
        placingPageBreakMark = item.para.pageBreakMark === true;
        // Pagination keeps: break the band early when this paragraph
        // (keepLines) — or this paragraph plus the head of what must follow
        // it (keepNext) — cannot finish here but WOULD fit a fresh band.
        // A chain taller than a whole band gives up, as Word does.
        // What space-before actually costs here, after collapsing against the
        // previous block's space-after. Used by BOTH the keep math and the
        // placement below, so a keep decision can never be made against a gap
        // the placer will not add.
        const effBefore =
          colDirty || bandOpensSection
            ? collapsedBefore(item.para.before, pendingAfter)
            : 0;
        if (colDirty && item.para.drafts) {
          const selfH = effBefore + draftsTotal(item.para.drafts);
          const bandH = colBottom() - bandTop;
          const avail = colBottom() - y;
          const needKL = item.para.keepLines && selfH > avail && selfH <= bandH;
          const needKN =
            item.para.keepNext &&
            (() => {
              const need =
                selfH +
                (item.para.after ?? 0) +
                keepAheadH(idx, item.para.after ?? 0);
              return need > avail && need <= bandH;
            })();
          if (needKL || needKN) breakBand();
        }
        // Space-before: a gap above the paragraph (collapsed away at a band
        // top, and against the previous block's space-after otherwise).
        if (effBefore) {
          y += effBefore;
          sectionRemaining -= effBefore;
        }
        // One box carries both what surrounds the paragraph and what fills
        // behind it, so either alone is enough to open a fragment.
        if (item.para.borders || item.para.shading) {
          activePBdr = {
            ...(item.para.borders && { borders: item.para.borders }),
            ...(item.para.shading && { shading: item.para.shading }),
            startedEarlier: false,
            frag: null,
          };
        }
        const drafts = item.para.drafts;
        const draftsHeight = drafts?.reduce((s, d) => s + d.height, 0) ?? 0;
        const floatsAhead = exclusions.some(
          (ex) => ex.bottom > y && ex.top < y + draftsHeight,
        );
        if (drafts && !floatsAhead) {
          emitParaDrafts(drafts, item.para.widowControl !== false);
        } else {
          placeParaBanded(item.para.getFlow(), effBefore);
        }
        if (activePBdr) {
          flushPBdrFrag(true); // last fragment closes the box's bottom edge
          activePBdr = null;
        }
        if (item.para.after) {
          y += item.para.after; // space-after gap
          sectionRemaining -= item.para.after;
        }
        pendingAfter = item.para.after ?? 0;
      } else {
        // A table carries no space-after, so nothing collapses into the next
        // paragraph's space-before.
        pendingAfter = 0;
        // Tables flow across columns/pages: split at row boundaries when
        // possible, and mid-row when a single row is taller than a whole band.
        // Header rows (w:tblHeader) repeat at the top of every fragment. Tables
        // are laid out at column width, then shifted into the current column.
        //
        // The unplaced remainder is a CURSOR into the item's immutable layout
        // (see TableRemainder), not a re-cloned table: each page clones only
        // the rows it places (remainderView), so a table spanning K pages
        // costs O(rows) in clones instead of O(rows × K), and the per-page
        // checkpoint snapshots the small carry instead of every remaining row.
        // item.table itself is never mutated — placed fragments are built from
        // clones — so a replay pass re-reads it pristine.
        // A replayed page opening mid-table resumes from the checkpointed
        // remainder instead of the item's start.
        let rem: TableRemainder = resumeTable ?? {
          base: item.table,
          fromY: 0,
          carry: null,
          carryHeight: 0,
          posDelta: 0,
        };
        resumeTable = null;
        // Word never flows a table BESIDE a floating object: text wraps around
        // one (see bandAt), but a table starts below its wrap zone. Without
        // this, a logo anchored at the top of a page painted straight over the
        // first table's opening rows. Skip past whatever the table would run
        // into, as long as the band still has somewhere to go.
        for (let guard = 0; guard < 8; guard++) {
          const bottomOfRun =
            y + Math.min(remainderHeight(rem), colBottom() - y);
          const blockers = exclusions.filter(
            (ex) =>
              ex.bottom > y &&
              ex.top < bottomOfRun &&
              ex.right > colX0() &&
              ex.left < colX1(),
          );
          if (blockers.length === 0) break;
          const next = Math.min(...blockers.map((ex) => ex.bottom));
          if (next <= y || next >= colBottom()) break;
          y = next;
        }
        const placeTable = (t: ResolvedTable) => {
          offsetTable(t, y);
          shiftTableX(t, xShift());
          noteNodePlaced();
          tables.push(t);
          colDirty = true;
          commitFns(tableFootnoteNums(t)); // footnotes referenced in the cells
          y += t.height;
          sectionRemaining -= t.height;
          bump();
        };
        for (;;) {
          const avail = colBottom() - y;
          const H = remainderHeight(rem);
          // The header band repeats only on continuation fragments, and only
          // while it leaves reasonable band room — judged against the page
          // the fragment actually lands on.
          let ghosts: ResolvedCell[] | null = null;
          let ghostH = 0;
          const hbBase = rem.base.headerBottom;
          if (
            remainderStarted(rem) &&
            hbBase != null &&
            hbBase < (limit() - bandTop) / 2
          ) {
            ghosts = cloneHeaderCells(rem.base, hbBase);
            if (ghosts) ghostH = hbBase;
          }
          // A continuation fragment redraws the table's top border and
          // reserves its thickness (Word-verified: a fat-bordered row split
          // across pages fits 19 lines on the SECOND page too, not 20). The
          // ghost header band, cloned from the base's own top, already
          // carries that edge — so the extra reserve applies only when no
          // ghosts are prepended.
          const edgeTop =
            remainderStarted(rem) && ghostH === 0 && rem.base.borders?.top
              ? rem.base.borders.top.width
              : 0;
          const fragTop = ghostH + edgeTop;
          /** Prepend the fragment chrome (continuation edge, ghost header). */
          const withGhosts = (frag: ResolvedTable): ResolvedTable => {
            if (fragTop > 0) {
              for (const cell of frag.cells) shiftCell(cell, fragTop);
              if (ghosts) {
                // Ghost cells keep their base coordinates (their band starts
                // at the table's own top edge) — only the LIVE content moved.
                frag.cells.unshift(...ghosts);
                frag.headerBottom = ghostH;
              }
              frag.height += fragTop;
              // The chrome pushes the content down, so the band lists move
              // with it — they are measured from the fragment's top.
              const bump = (b: { top: number; bottom: number }) => ({
                top: b.top + fragTop,
                bottom: b.bottom + fragTop,
              });
              if (frag.cantSplitBands)
                frag.cantSplitBands = frag.cantSplitBands.map(bump);
              if (frag.keepStartBands)
                frag.keepStartBands = frag.keepStartBands.map(bump);
            }
            return frag;
          };
          const remFns = remainderFootnoteNums(rem, ghostH);
          if (fragTop + H + addedReserve(remFns) <= avail) {
            placeTable(withGhosts(remainderView(rem, Infinity)));
            break;
          }
          const budget = avail - fragTop;
          /** Move the remainder whole to a fresh column/page. */
          const freshBand = () => {
            midTableRest = remainderStarted(rem) ? rem : null;
            breakBand();
            midTableRest = null;
          };
          /** Commit a split: place the top fragment, advance the cursor, and
           *  land the remainder on the next band. Every continuation cell
           *  splitTableAt built has height = contHeight, and (with no
           *  below-cut cells in the view) rest.height = contHeight +
           *  (view.height - splitBottom), which recovers the broken row's
           *  bottom. A cut on a row boundary has no continuation: the
           *  consumed prefix moves to the cut itself. A cut inside the carry
           *  consumes no base content (splitBottom ≤ carry height ⇒ fromY
           *  unchanged). */
          const commitSplit = (
            view: ResolvedTable,
            top: ResolvedTable,
            rest: ResolvedTable,
          ) => {
            const contH = rest.cells.length > 0 ? rest.cells[0].height : 0;
            const splitBottom = view.height - rest.height + contH;
            const advance = splitBottom - rem.carryHeight;
            const next: TableRemainder = {
              base: rem.base,
              fromY: rem.fromY + (advance > REMAINDER_EPS ? advance : 0),
              carry: rest.cells.length > 0 ? rest.cells : null,
              carryHeight: contH,
              posDelta: 0,
            };
            placeTable(withGhosts(top));
            midTableRest = next; // the break lands the remainder on the new page
            breakBand();
            midTableRest = null;
            rem = next;
          };
          if (budget <= 0) {
            // Header band swallows the remaining space — try a fresh band, or
            // place whole when the geometry is truly degenerate.
            if (colDirty) {
              freshBand();
              continue;
            }
            placeTable(withGhosts(remainderView(rem, Infinity)));
            break;
          }
          // Row boundaries in remainder space (y = 0 at the carry's top):
          // `boundary` is the largest at or under the budget, `rowBottom` the
          // first one past it — together they frame the row the page edge
          // lands in. Cells already part of the carry offer no boundary.
          let boundary = 0;
          let rowBottom = H;
          for (const cell of rem.base.cells) {
            if (cell.y < rem.fromY - REMAINDER_EPS) continue; // mid-carry
            const yr = remainderRelY(rem, cell.y);
            if (yr <= REMAINDER_EPS) continue;
            if (yr <= budget) boundary = Math.max(boundary, yr);
            else if (yr < rowBottom) rowBottom = yr;
          }
          const bandShift = rem.carryHeight - rem.fromY;
          const covers = (b: { top: number; bottom: number }) =>
            b.top + bandShift <= boundary + 0.5 &&
            b.bottom + bandShift >= rowBottom - 0.5;
          const straddleCantSplit = (rem.base.cantSplitBands ?? []).some(
            covers,
          );
          // Word's row-start veto: a row whose first cell opens with a
          // keepNext paragraph must not START in a band's leftover — it
          // begins on a fresh band and only then splits normally. The rule
          // is about STARTING: a carry (the row already began) is exempt,
          // and so is a row sitting at the top of a fresh band. A remainder
          // that fits whole never reaches this point — no split touches the
          // row, so the keep is satisfied where it lies.
          const rowStartsHere = boundary > 0 || rem.carryHeight === 0;
          const rowStartVeto =
            rowStartsHere &&
            (boundary > 0 || colDirty) &&
            (rem.base.keepStartBands ?? []).some(covers);
          // 1) Word's default: FILL the leftover — split the row the page
          //    edge lands in, unless it is w:cantSplit. splitTableAt snaps
          //    the cut down to keep-legal line boundaries per cell, so the
          //    attempt is taken only when the straddling row genuinely
          //    contributes content past the row boundary (otherwise the row
          //    boundary gives the same fragment without a continuation).
          //    Fragments carrying fresh footnotes shrink the budget by the
          //    reserve their notes add — monotone, so a bounded walk-down
          //    settles it; footnote-free tables (the norm) skip it entirely.
          // Word splits a nested table whenever the page edge falls inside
          // it — no height threshold (fixture-verified: a 10-row nested
          // table with room to move whole still splits at the cut).
          const splitOpts = { splitNested: true };
          if (!straddleCantSplit && !rowStartVeto) {
            let effBudget = budget;
            let view = remainderView(rem, effBudget);
            let frag = splitTableAt(view, effBudget, splitOpts);
            if (remFns.length > 0) {
              for (let i = 0; i < 3 && effBudget > 0; i++) {
                const need = addedReserve(tableFootnoteNums(frag.top));
                if (ghostH + effBudget + need <= avail) break;
                effBudget = avail - ghostH - need;
                if (effBudget <= 0) break;
                view = remainderView(rem, effBudget);
                frag = splitTableAt(view, effBudget, splitOpts);
              }
            }
            const contributed =
              effBudget > 0 &&
              frag.top.cells.some(
                (c) =>
                  c.lines.some(
                    (l) => l.y + l.height > boundary + REMAINDER_EPS,
                  ) ||
                  (c.tables ?? []).some(
                    (t) => t.y + t.height > boundary + REMAINDER_EPS,
                  ),
              );
            if (contributed && frag.rest.height < H + ghostH) {
              commitSplit(view, frag.top, frag.rest);
              continue;
            }
          }
          // 2) The straddling row moves whole: cut at the row boundary.
          if (boundary > 0) {
            const view = remainderView(rem, boundary);
            const { top: topFrag, rest } = splitTableAt(
              view,
              boundary,
              splitOpts,
            );
            if (rest.height >= H + ghostH) {
              // The cut moved nothing. Retry on a fresh band when this one
              // is already used; only a genuinely band-taller table, with
              // nowhere better to go, is placed as-is (overflowing).
              if (colDirty) {
                freshBand();
                continue;
              }
              placeTable(withGhosts(remainderView(rem, Infinity)));
              break;
            }
            const topHasContent = topFrag.cells.some(
              (c) => c.lines.length > 0 || (c.tables?.length ?? 0) > 0,
            );
            if (!topHasContent && colDirty) {
              freshBand();
              continue;
            }
            commitSplit(view, topFrag, rest);
            continue;
          }
          // 3) No row boundary fits and the straddling row would not (or may
          //    not) split. A w:cantSplit row that would fit a full fresh
          //    band moves there whole; past that, Word splits even a
          //    cantSplit row rather than overflow the page. A row-start
          //    veto moves to the fresh band UNCONDITIONALLY (Word pushes a
          //    keep-start row even when it cannot fit any single band) —
          //    the veto turns itself off there (band no longer dirty).
          if (rowStartVeto && colDirty) {
            freshBand();
            continue;
          }
          if (straddleCantSplit) {
            const fitsFullBand = rowBottom <= limit() - bandTop;
            if (fitsFullBand && colDirty) {
              freshBand();
              continue;
            }
          }
          {
            const view = remainderView(rem, budget);
            const { top: topFrag, rest } = splitTableAt(
              view,
              budget,
              splitOpts,
            );
            const topHasContent = topFrag.cells.some(
              (c) => c.lines.length > 0 || (c.tables?.length ?? 0) > 0,
            );
            if (rest.height >= H + ghostH || !topHasContent) {
              // Not even one line legally fits the leftover space — don't
              // paint an empty table stub; start on the next column/page
              // instead, or place whole when the geometry is degenerate.
              if (colDirty) {
                freshBand();
                continue;
              }
              placeTable(withGhosts(remainderView(rem, Infinity)));
              break;
            }
            commitSplit(view, topFrag, rest);
          }
        }
      }
    }

    if (pageHasContent() || pages.length === 0) finalizePage();
  }
}

/** Lay out already-flattened blocks into paginated pages. Pure (no DOM);
 *  measurement is injected. */
/** A degenerate page config (NaN or absurd values from a hostile import)
 *  must never reach pagination — a non-finite or negative content box turns
 *  the page-fill loop infinite and hangs the app. Sanitize per field. Applied
 *  to the document page AND every per-section override. */
function sanitizePage(p: PageConfig): PageConfig {
  const num = (v: number, fallback: number) =>
    Number.isFinite(v) && v >= 0 ? v : fallback;
  // Small pages are legitimate — only zero/NaN falls back to A4.
  const width = num(p.width, 794) || 794;
  const height = num(p.height, 1123) || 1123;
  let margin = {
    top: num(p.margin.top, 96),
    right: num(p.margin.right, 96),
    bottom: num(p.margin.bottom, 96),
    left: num(p.margin.left, 96),
  };
  // Margins must leave SOME content box on both axes (a zero/negative box
  // never fits a line and the fill loop can't advance).
  if (width - margin.left - margin.right < 10)
    margin = { ...margin, left: 0, right: 0 };
  if (height - margin.top - margin.bottom < 10)
    margin = { ...margin, top: 0, bottom: 0 };
  const same =
    width === p.width &&
    height === p.height &&
    margin.top === p.margin.top &&
    margin.right === p.margin.right &&
    margin.bottom === p.margin.bottom &&
    margin.left === p.margin.left;
  if (same) return p;
  const out: PageConfig = { width, height, margin };
  if (p.headerDistance !== undefined)
    out.headerDistance = num(p.headerDistance, 48);
  if (p.footerDistance !== undefined)
    out.footerDistance = num(p.footerDistance, 48);
  if (p.gutter !== undefined) out.gutter = num(p.gutter, 0);
  return out;
}

function sanitizeConfig(config: LayoutConfig): LayoutConfig {
  const page = sanitizePage(config.page);
  return page === config.page ? config : { ...config, page };
}

/**
 * Whether two section geometries print on the same SHEET.
 *
 * This is the only part of the geometry that can force a page break. A sheet
 * has one size, so a section that changes it cannot share a page with the one
 * before — Word promotes even a continuous break to next-page there. Margins
 * are not the sheet: they describe where text sits on it, and Word lets a
 * continuous section change them without starting a page (the horizontal ones
 * take effect at once — the documented way to change left/right margins
 * mid-page — while the vertical ones simply wait, because "the first place
 * Word can make this margin change is at the top of page 2").
 *
 * Measured, not inferred: guidance on the web says both things, so the corpus
 * factsheet's own PDF from Word decided it. Its section 3 (top margin 280tw)
 * is followed by a CONTINUOUS section 4 (400tw), and every anchored image on
 * that page sits 17px lower in Word's PDF than a sec4-opened page allows —
 * which is exactly 16.5px, the difference between opening the page at sec3
 * (18.67px margin plus its three empty paragraphs, 24.5px) and opening it at
 * sec4 (26.67px). Word put both sections on one page.
 */
function sameSheet(a: PageConfig, b: PageConfig): boolean {
  return a.width === b.width && a.height === b.height;
}

/** Field-equality for page geometry (identity is not enough — overrides come
 *  from doc attrs and are rebuilt on every transaction). */
function sameGeom(a: PageConfig, b: PageConfig): boolean {
  return (
    a.width === b.width &&
    a.height === b.height &&
    a.margin.top === b.margin.top &&
    a.margin.right === b.margin.right &&
    a.margin.bottom === b.margin.bottom &&
    a.margin.left === b.margin.left &&
    a.headerDistance === b.headerDistance &&
    a.footerDistance === b.footerDistance &&
    a.gutter === b.gutter
  );
}

export function layoutBlocks(
  blocks: FlowBlock[],
  config: LayoutConfig,
): ResolvedLayout {
  config = sanitizeConfig(config);
  const ctx = buildCtx(config);
  const left = contentLeftOf(config.page);
  const right = config.page.width - config.page.margin.right;
  // Whole-document column flow: lay every block at the column content width.
  const cols: ColumnConfig = config.columns ?? { count: 1, gap: 0 };
  const docBoxes = columnBoxes(cols);
  const colWidth = docBoxes
    ? docBoxes[0].width
    : columnWidth(right - left, cols);
  const colRight = left + colWidth;
  // See the sectioned builder: columns of differing widths cannot share one
  // pre-wrap, so their paragraphs wrap where they are placed.
  const unevenCols = unevenColumns(cols);
  const items: BlockItem[] = blocks.map((block, i) => {
    const item: BlockItem =
      block.type === 'paragraph'
        ? {
            para: {
              getFlow: () => block,
              drafts:
                block.floats?.length ||
                unevenCols ||
                block.runs.some((r) => 'columnBreak' in r)
                  ? null
                  : layoutParagraph(block, left, colRight, ctx),
              before: block.spacing?.before,
              after: block.spacing?.after,
              pageBreakBefore: block.pageBreakBefore,
              columnBreak: block.runs.some((r) => 'columnBreak' in r),
              keepNext: block.keepNext,
              keepLines: block.keepLines,
              widowControl: block.widowControl,
              borders: block.borders,
              shading: block.shading,
            },
          }
        : { table: layoutTable(block, left, colRight, ctx, cols.count > 1) };
    if (i === 0) item.section = { ...cols, newPage: true };
    return item;
  });
  assignSectionHeights(items);
  return placeBlocks(items, config, ctx);
}

/** Equal-column content width: total width minus the inter-column gaps. */
function columnWidth(totalWidth: number, cols: ColumnConfig): number {
  const count = Math.max(1, cols.count);
  const gap = count > 1 ? cols.gap : 0;
  return (totalWidth - gap * (count - 1)) / count;
}

/** One text column's box, x measured from the content-area left edge. */
type ColumnBox = { x: number; width: number };

/**
 * Per-column boxes for a section that declares its columns individually, or
 * null when there is nothing to declare and the uniform arithmetic applies.
 *
 * Each `w:col` gives its own width and the space that FOLLOWS it, so the
 * columns simply accumulate. The declared numbers are used verbatim, never
 * re-derived: a file may spell out EQUAL widths with a gap that the
 * count-and-divide arithmetic would not reproduce (one in the corpus declares
 * 4470|420|4470 where dividing gives 4320|720|4320), and rebuilding those from
 * `gap` would move both columns while looking like a harmless shortcut.
 */
function columnBoxes(cols: ColumnConfig): ColumnBox[] | null {
  const list = cols.cols;
  if (!list || list.length < 2) return null;
  const boxes: ColumnBox[] = [];
  let x = 0;
  for (const c of list) {
    boxes.push({ x, width: c.width });
    x += c.width + c.space;
  }
  return boxes;
}

/**
 * Whether the section's columns have DIFFERENT widths — a separate question
 * from whether they are declared individually, and the one that decides how
 * paragraphs wrap. Equal columns share one wrap, so a paragraph measured once
 * is valid wherever it lands; unequal ones have no such width, and a paragraph
 * that straddles the boundary has to re-wrap halfway.
 */
function unevenColumns(cols: ColumnConfig): boolean {
  const list = cols.cols;
  return (
    !!list && list.length > 1 && list.some((c) => c.width !== list[0].width)
  );
}

// ── Incremental re-layout (M4+) ─────────────────────────────────────

/** Shift every PM position in the drafts by `delta` **in place** (geometry is
 *  unchanged — the paragraph's text and wrap are identical, it just moved in the
 *  doc). Mutating rather than reallocating is the hot-path win for large docs:
 *  a single keystroke near the top shifts every following block, and cloning
 *  each line + segment there dominated typing latency.
 *
 *  Safety: a draft's segment/image objects can be shared with the *previous*
 *  frame's `LayoutLine`s (single-column `draftToLine` reuses them). That's fine
 *  — layout is synchronous and the previous `resolved` is neither read during
 *  the rebuild nor kept after `layoutDoc` reassigns it; the freshly built layout
 *  re-materializes lines from these same (now-shifted) drafts. Multi-column
 *  `draftToLine` clones segments, snapshotting `pos` after this shift. */
function shiftDraftsInPlace(drafts: LineDraft[], delta: number): void {
  for (const d of drafts) {
    if (d.from != null) d.from += delta;
    if (d.to != null) d.to += delta;
    for (const s of d.segments) if (s.pos != null) s.pos += delta;
    for (const im of d.images) if (im.pos != null) im.pos += delta;
  }
}

interface ParagraphCacheEntry {
  left: number;
  right: number;
  /** Content-start position the cached drafts were computed at. */
  basePos: number;
  /** List marker the drafts were computed with — renumbering invalidates. */
  marker?: string;
  /** The built block item, reused across frames (its `drafts` are shifted in
   *  place when the paragraph moves) so pure cache hits allocate nothing. */
  item: ParaItem;
}

interface TableCacheEntry {
  left: number;
  right: number;
  /** Table node position the cached layout was computed at. */
  basePos: number;
  /** Canonical layout (y = 0, column-0 x, positions at basePos). Cloned per
   *  paint so the placer can mutate (offset/shift/split) without corrupting it. */
  table: ResolvedTable;
}

/** Shift a laid-out line's PM positions by `delta`, deep-cloning it (geometry
 *  is unchanged — the line just moved in the document). */
function cloneLineShifted(l: LayoutLine, delta: number): LayoutLine {
  const out: LayoutLine = {
    ...l,
    segments: l.segments.map((s) =>
      s.pos != null ? { ...s, pos: s.pos + delta } : { ...s },
    ),
  };
  if (l.images)
    out.images = l.images.map((im) =>
      im.pos != null ? { ...im, pos: im.pos + delta } : { ...im },
    );
  if (l.from != null) out.from = l.from + delta;
  if (l.to != null) out.to = l.to + delta;
  return out;
}

/** Deep-clone a resolved table, shifting every PM position by `delta`. Unlike
 *  `cloneTable` (which shares segment arrays), this clones segments too, so the
 *  paginator can offset/shift/split the result without touching the cached one. */
function cloneTableShifted(t: ResolvedTable, delta: number): ResolvedTable {
  return {
    ...t,
    borders: t.borders ? { ...t.borders } : t.borders,
    cells: t.cells.map((c) => ({
      ...c,
      borders: c.borders ? { ...c.borders } : c.borders,
      diagonals: c.diagonals ? { ...c.diagonals } : c.diagonals,
      lines: c.lines.map((l) => cloneLineShifted(l, delta)),
      tables: c.tables?.map((nt) => cloneTableShifted(nt, delta)),
      floats: c.floats?.map((f) => ({ ...f })),
    })),
  };
}

/** Whether a table contains any list paragraph. Such tables advance the live
 *  numbering counter, so they're laid out fresh (never cached) to keep markers
 *  in sync with edits elsewhere in the document. */
function tableHasList(node: PMNode): boolean {
  let found = false;
  node.descendants((n) => {
    if (found) return false;
    if (n.type.name === 'paragraph' && n.attrs['list']) found = true;
    return !found;
  });
  return found;
}

/**
 * Layout cache keyed on ProseMirror node identity (PM keeps unchanged nodes
 * identical across transactions). An unchanged paragraph/table skips
 * measuring/wrapping entirely; one that merely moved gets its positions
 * shifted. Tables that host list paragraphs are not cached (they advance the
 * live numbering counter).
 */
export class LayoutCache {
  readonly paragraphs = new WeakMap<PMNode, ParagraphCacheEntry>();
  readonly tables = new WeakMap<PMNode, TableCacheEntry>();
  /** Placed pages of previous passes, keyed by a layout fingerprint (page
   *  geometry, chrome band heights, tab grid, footnote heights). Unchanged
   *  prefix pages are reused outright; an unchanged tail re-attaches at the
   *  first clean boundary after an edit, shifted by the PM-position delta. */
  readonly pageRuns = new Map<string, PageRun>();
}

export function createLayoutCache(): LayoutCache {
  return new LayoutCache();
}

/** Word's default header/footer distance from the page edge (720 twips). */
const CHROME_DISTANCE = 48;
/** Effective chrome distances / left content edge for a page config. */
const headerDist = (p: PageConfig): number =>
  p.headerDistance ?? CHROME_DISTANCE;
const footerDist = (p: PageConfig): number =>
  p.footerDistance ?? CHROME_DISTANCE;
const contentLeftOf = (p: PageConfig): number =>
  p.margin.left + (p.gutter ?? 0);

/** Repeating page furniture passed alongside the body document. `header`/
 *  `footer` are the default (odd-page) bands; the `*First`/`*Even` variants
 *  apply on page 1 (when `titlePg`) and even pages (when `evenAndOdd`). */
export interface PageChrome {
  header?: PMNode;
  footer?: PMNode;
  headerFirst?: PMNode;
  footerFirst?: PMNode;
  headerEven?: PMNode;
  footerEven?: PMNode;
  /** w:titlePg — page 1 uses the first variant (blank if none declared). */
  titlePg?: boolean;
  /** w:evenAndOddHeaders — even pages use the even variant (blank if none). */
  evenAndOdd?: boolean;
  /** Per-section chrome, aligned with doc.attrs.sections. When present it
   *  wins over the flat fields above: each section's pages take their own
   *  header/footer set (titlePg selects the "first" variant on the SECTION's
   *  first page), laid out against that section's page geometry. */
  sections?: SectionChromeDocs[];
}

/** One section's chrome documents (import side, pre-layout). Inheritance
 *  (Word's "Link to Previous") is resolved by the importer — every section
 *  arrives here with its effective documents. */
export interface SectionChromeDocs {
  header?: PMNode;
  footer?: PMNode;
  headerFirst?: PMNode;
  footerFirst?: PMNode;
  headerEven?: PMNode;
  footerEven?: PMNode;
  titlePg?: boolean;
}

/** Header/footer documents grouped by variant. */
interface ChromeDocs {
  default?: PMNode;
  first?: PMNode;
  even?: PMNode;
}
interface ChromeBands {
  default?: ResolvedChrome;
  first?: ResolvedChrome;
  even?: ResolvedChrome;
}
function layVariants(
  docs: ChromeDocs,
  fn: (doc: PMNode) => ResolvedChrome,
): ChromeBands {
  const out: ChromeBands = {};
  if (docs.default) out.default = fn(docs.default);
  if (docs.first) out.first = fn(docs.first);
  if (docs.even) out.even = fn(docs.even);
  return out;
}
/** Tallest band among present variants (sizes the body band conservatively). */
function maxBandHeight(bands: ChromeBands): number {
  return Math.max(
    0,
    ...[bands.default, bands.first, bands.even]
      .filter(Boolean)
      .map((b) => b!.height),
  );
}
function anyBandHasFields(bands: ChromeBands): boolean {
  return [bands.default, bands.first, bands.even].some(
    (b) => b?.lines.some((l) => l.segments.some((s) => s.field)) ?? false,
  );
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
    for (const c of t.cells) {
      stripPositions(c.lines, c.tables ?? []);
      if (c.floats) for (const f of c.floats) delete f.pos;
    }
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
  // Anchored floats are positioned within the band (painted only — chrome
  // text doesn't wrap around them; layoutFlow places them at their offsets).
  const flow = layoutFlow(toFlowBlocks(doc, ctx.base, true), left, right, ctx);
  const lines = flow.lines.map((l) => ({ ...l, y: l.y + topY }));
  flow.tables.forEach((t) => offsetTable(t, topY));
  stripPositions(lines, flow.tables);
  const band: ResolvedChrome = {
    lines,
    tables: flow.tables,
    height: flow.height,
  };
  if (flow.floats.length > 0)
    band.floats = flow.floats.map((f) => ({
      ...f,
      y: f.y + topY,
      pos: undefined,
    }));
  return band;
}

/** Lay out a footer document pinned `dist` px from the page bottom. */
function layoutFooterChrome(
  doc: PMNode,
  pageHeight: number,
  dist: number,
  left: number,
  right: number,
  ctx: Ctx,
): ResolvedChrome {
  const flow = layoutChrome(doc, 0, left, right, ctx);
  const topY = pageHeight - dist - flow.height;
  const band: ResolvedChrome = {
    lines: flow.lines.map((l) => ({ ...l, y: l.y + topY })),
    tables: flow.tables,
    height: flow.height,
  };
  band.tables.forEach((t) => offsetTable(t, topY));
  if (flow.floats)
    band.floats = flow.floats.map((f) => ({ ...f, y: f.y + topY }));
  return band;
}

/** Lay out one footnote body (reduced font) flat from y = 0, stripping PM
 *  positions (it belongs to a separate note story — never caret-addressable). */
function layoutFootnoteBody(
  doc: PMNode,
  left: number,
  right: number,
  ctx: Ctx,
): FootnoteBody {
  const fnCtx: Ctx = {
    ...ctx,
    base: { ...ctx.base, sizePt: ctx.base.sizePt * FOOTNOTE_FONT_SCALE },
  };
  const flow = layoutFlow(
    toFlowBlocks(doc, fnCtx.base, false),
    left,
    right,
    fnCtx,
  );
  stripPositions(flow.lines, flow.tables);
  return { lines: flow.lines, height: flow.height };
}

// ── Display page numbers (w:pgNumType) ─────────────────────────────

const ROMAN: [number, string][] = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];

function toRoman(n: number): string {
  let out = '';
  for (const [v, s] of ROMAN) {
    while (n >= v) {
      out += s;
      n -= v;
    }
  }
  return out;
}

/** Word's letter numbering repeats the letter past z: 1→a … 26→z, 27→aa,
 *  28→bb (NOT spreadsheet-style aa/ab). */
function toLetters(n: number): string {
  const letter = String.fromCharCode(97 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

/** A display page number in an OOXML ST_NumberFormat. Unrecognized formats
 *  (and values the format can't express, like roman 0) fall back to decimal —
 *  the model keeps the raw fmt string, so nothing is lost by rendering it
 *  conservatively. */
export function formatPageNumber(n: number, fmt?: string): string {
  if (n < 1) return String(n);
  switch (fmt) {
    case 'upperRoman':
      return toRoman(n);
    case 'lowerRoman':
      return toRoman(n).toLowerCase();
    case 'upperLetter':
      return toLetters(n).toUpperCase();
    case 'lowerLetter':
      return toLetters(n);
    default:
      return String(n);
  }
}

/** Per-page display numbers from the sections' w:pgNumType, or undefined when
 *  no section declares one (display number = index + 1, nothing to store).
 *  Each page is governed by the section in effect where it starts
 *  (`chromeIndex`); entering a new section applies the LAST restart declared
 *  by the sections crossed — so a restart on a continuous section wholly
 *  inside a page still takes effect on the next page, matching how Word
 *  settles the same ambiguity. Formats don't inherit: a section without
 *  pgNumType shows decimal (the spec default). */
export function computePageLabels(
  pages: readonly { chromeIndex?: number }[],
  sections: readonly SectionConfig[],
): string[] | undefined {
  if (!sections.some((s) => s.pageNumbers)) return undefined;
  const labels: string[] = [];
  let si = 0;
  let counter = 0;
  for (const page of pages) {
    const next = Math.min(page.chromeIndex ?? si, sections.length - 1);
    let restart: number | undefined;
    // First page scans its own section too (a start on section 0 applies).
    for (let i = labels.length === 0 ? 0 : si + 1; i <= next; i++) {
      const s = sections[i].pageNumbers?.start;
      if (s != null && Number.isInteger(s) && s >= 0) restart = s;
    }
    counter = restart ?? counter + 1;
    si = Math.max(si, next);
    labels.push(formatPageNumber(counter, sections[si].pageNumbers?.fmt));
  }
  return labels;
}

/** Lay out a ProseMirror document into paint-ready pages. With a `cache`,
 *  only paragraphs whose node changed since the previous call are re-measured.
 *  `chrome` (page header/footer documents) repeats on every page; the body
 *  band shrinks when a chrome band is taller than the page margin. `footnotes`
 *  (note bodies keyed by display number) are laid out at the bottom of the page
 *  their reference falls on. */
export function layout(
  doc: PMNode,
  config: LayoutConfig,
  cache?: LayoutCache,
  chrome?: PageChrome,
  footnotes?: Record<number, PMNode>,
): ResolvedLayout {
  config = sanitizeConfig(config);
  const ctx = buildCtx(config, doc.attrs['compat'] as DocCompat | null);
  const { page } = config;
  const left = contentLeftOf(page);
  const right = page.width - page.margin.right;

  // Page chrome first — a tall header/footer pushes the body band inward. The
  // first/even variants are laid out alongside the default; the band shrinks by
  // the TALLEST variant so no page's content overlaps its chrome.
  const headerDocs: ChromeDocs = {
    default: chrome?.header,
    first: chrome?.headerFirst,
    even: chrome?.headerEven,
  };
  const footerDocs: ChromeDocs = {
    default: chrome?.footer,
    first: chrome?.footerFirst,
    even: chrome?.footerEven,
  };
  const layHeaders = (c: Ctx) =>
    layVariants(headerDocs, (d) =>
      layoutChrome(d, headerDist(page), left, right, c),
    );
  const layFooters = (c: Ctx) =>
    layVariants(footerDocs, (d) =>
      layoutFooterChrome(d, page.height, footerDist(page), left, right, c),
    );
  let headers = perf.span('chrome-headers', () => layHeaders(ctx));
  let footers = perf.span('chrome-footers', () => layFooters(ctx));

  // Markers are recounted on every layout, so list edits renumber live. The
  // counter advances for every list paragraph — including cache hits.
  const counter = createNumberingCounter(
    doc.attrs['numbering'] as NumberingDefs | null,
  );

  // Section column flow: each block carries its section's columns and whether
  // it opens a section. Absent → one implicit single-column section.
  const sections = (doc.attrs['sections'] as SectionConfig[] | null) ?? [
    {
      blockCount: doc.childCount,
      columns: { count: 1, gap: 0 },
      newPage: true,
    },
  ];
  const blockSection: {
    columns: ColumnConfig;
    start: boolean;
    newPage: boolean;
    page?: PageConfig;
    chromeIndex: number;
    /** Last block of a section whose break is CONTINUOUS — its paragraph mark
     *  is the break itself and draws nothing (see FlowParagraph.breakMark).
     *  A next-page break keeps its line: the page it ends still has to hold
     *  it, and Word is known to spill a whole extra page over exactly that. */
    breakMark: boolean;
    /** Last block of a section whose break is NEXT-PAGE. Its mark keeps its
     *  line MID-page, and may still flow to the next column — but it never
     *  opens a page: a probe with four such sections, marks alternately
     *  fitting and not, at bottom margins 720 and 0, paginates in Word as
     *  exactly one page per section. An unfitting mark is simply clipped at
     *  the floor of its section's last page. */
    pageBreakMark: boolean;
  }[] = [];
  sections.forEach((sec, si) => {
    // Per-section geometry overrides are attr data — sanitize like config.page
    // (degenerate values would hang the page-fill loop).
    const secPage = sec.page ? sanitizePage(sec.page) : undefined;
    for (let k = 0; k < sec.blockCount; k++) {
      blockSection.push({
        columns: sec.columns,
        start: k === 0,
        newPage: sec.newPage,
        page: secPage,
        chromeIndex: si,
        breakMark:
          k === sec.blockCount - 1 &&
          si < sections.length - 1 &&
          !sections[si + 1].newPage,
        pageBreakMark:
          k === sec.blockCount - 1 &&
          si < sections.length - 1 &&
          sections[si + 1].newPage,
      });
    }
  });
  const lastSec = sections[sections.length - 1];
  const lastSecPage = lastSec.page ? sanitizePage(lastSec.page) : undefined;
  while (blockSection.length < doc.childCount) {
    blockSection.push({
      columns: lastSec.columns,
      start: false,
      newPage: lastSec.newPage,
      page: lastSecPage,
      chromeIndex: sections.length - 1,
      breakMark: false,
      pageBreakMark: false,
    });
  }

  // Per-section chrome: each section's bands laid against its own geometry.
  // setBandH[i] feeds bandFor so a tall section header shrinks exactly that
  // section's pages.
  const secChromeDocs = chrome?.sections;
  let chromeSets: ResolvedChromeSet[] | undefined;
  let setBandH: { header: number; footer: number }[] = [];
  const layChromeSets = (c: Ctx) => {
    if (!secChromeDocs) return;
    chromeSets = [];
    setBandH = [];
    for (let i = 0; i < sections.length; i++) {
      const cs = secChromeDocs[i] ?? {};
      const p = sections[i].page
        ? sanitizePage(sections[i].page as PageConfig)
        : page;
      const l = contentLeftOf(p);
      const r = p.width - p.margin.right;
      const hb = layVariants(
        { default: cs.header, first: cs.headerFirst, even: cs.headerEven },
        (d) => layoutChrome(d, headerDist(p), l, r, c),
      );
      const fb = layVariants(
        { default: cs.footer, first: cs.footerFirst, even: cs.footerEven },
        (d) => layoutFooterChrome(d, p.height, footerDist(p), l, r, c),
      );
      chromeSets.push({
        header: hb.default,
        footer: fb.default,
        headerFirst: hb.first,
        footerFirst: fb.first,
        headerEven: hb.even,
        footerEven: fb.even,
        titlePg: cs.titlePg,
      });
      setBandH.push({ header: maxBandHeight(hb), footer: maxBandHeight(fb) });
    }
  };
  perf.span('chrome-sections', () => layChromeSets(ctx));

  const items: BlockItem[] = [];
  const buildEnd = perf.begin('build-items');
  doc.forEach((node, offset, index) => {
    const bs = blockSection[index] ?? {
      columns: { count: 1, gap: 0 },
      start: false,
      newPage: true,
    };
    // Blocks wrap at their section's column width — computed against the
    // section's own page geometry when it overrides the document's. `bLeft`/
    // `colRight` are the cache key, so a geometry change re-measures exactly
    // the affected sections.
    const secPage = bs.page ?? page;
    const bLeft = contentLeftOf(secPage);
    const bRight = secPage.width - secPage.margin.right;
    // The pre-wrap width comes from the declared boxes when there are any,
    // never from the count-and-divide arithmetic: a landscape section in the
    // corpus declares 4470|420|4470 inside a 12960-twip text area, so dividing
    // gives 408px columns where the file asks for 298px and every line would
    // be wrapped a third too wide for the column it is painted in.
    const secBoxes = columnBoxes(bs.columns);
    const colRight =
      bLeft +
      (secBoxes ? secBoxes[0].width : columnWidth(bRight - bLeft, bs.columns));
    // Columns of DIFFERENT widths have no single width to pre-wrap at: a
    // paragraph wrapped for column 0 is wrong in column 1, and one that
    // straddles the boundary has to re-wrap halfway. Such sections take the
    // placement-time path, which asks the band for the column it actually
    // lands in — the same path float-anchoring paragraphs already use.
    const unevenCols = unevenColumns(bs.columns);
    const tag = (item: BlockItem): BlockItem => {
      if (bs.start) {
        item.section = { ...bs.columns, newPage: bs.newPage };
        if (bs.page) item.section.page = bs.page;
        item.section.chromeIndex = bs.chromeIndex;
      }
      item.node = node;
      item.nodeOffset = offset;
      return item;
    };
    if (node.type.name === 'paragraph') {
      const m = markerFor(node, counter); // advances numbering every pass
      const marker = m?.text;
      const markerStyle = m?.style;
      const contentStart = offset + 1;
      // Fast path: an unchanged paragraph reuses its cached item outright, and a
      // moved one shifts its drafts in place — neither allocates a ParaItem,
      // closure, or attr read. This is the dominant per-keystroke cost on large
      // docs (thousands of blocks re-scanned every edit), so it stays lean.
      // A cache hit also implies the paragraph anchors no floats (float
      // paragraphs are never cached), so the O(children) `nodeHasFloats` scan is
      // skipped here and only paid on a miss.
      const hit = cache?.paragraphs.get(node);
      if (
        hit &&
        hit.left === bLeft &&
        hit.right === colRight &&
        hit.marker === marker
      ) {
        perf.bump(hit.basePos !== contentStart ? 'para.shift' : 'para.hit');
        if (hit.basePos !== contentStart) {
          if (hit.item.drafts)
            shiftDraftsInPlace(hit.item.drafts, contentStart - hit.basePos);
          hit.basePos = contentStart;
        }
        // Fresh 1-field wrapper so `tag` can add the section marker without
        // mutating the cached ParaItem.
        items.push(tag({ para: hit.item }));
        return;
      }
      // Miss (or float-anchoring paragraph, never cached): build fresh.
      const hasFloats = nodeHasFloats(node);
      // Which paragraph ends a section is a property of the SECTION table, not
      // of the node, so it is stamped here rather than by the importer.
      const getFlow = () => {
        const f = paragraphToFlow(
          node,
          ctx.base,
          offset,
          true,
          marker,
          markerStyle,
        );
        if (bs.breakMark) f.breakMark = true;
        return f;
      };
      const sp = effectiveSpacing(
        node.attrs['spacing'] as ParagraphSpacing | null,
        node.attrs['contextualSpacing'] as {
          before: boolean;
          after: boolean;
        } | null,
      );
      const mkItem = (drafts: LineDraft[] | null): ParaItem => ({
        getFlow,
        drafts,
        ...(bs.pageBreakMark &&
          node.childCount === 0 && { pageBreakMark: true }),
        before: sp?.before,
        after: sp?.after,
        pageBreakBefore: node.attrs['pageBreakBefore'] === true,
        columnBreak: paragraphHasColumnBreak(node),
        keepNext: node.attrs['keepNext'] === true,
        keepLines: node.attrs['keepLines'] === true,
        widowControl: node.attrs['widowControl'] !== false,
        borders:
          (node.attrs['borders'] as ParagraphBorders | null) ?? undefined,
        shading: (node.attrs['shading'] as string | null) ?? undefined,
      });
      // Float-anchoring paragraphs always wrap at placement time (their band
      // depends on where they land).
      // A column break splits the paragraph across two columns, which a
      // single pre-wrapped run of lines cannot express.
      // Break marks join the never-cached set: an edit that moves a section
      // boundary changes the flag without changing the node, so a cached
      // draft would keep the old line height.
      if (
        hasFloats ||
        unevenCols ||
        bs.breakMark ||
        bs.pageBreakMark ||
        paragraphHasColumnBreak(node)
      ) {
        items.push(tag({ para: mkItem(null) }));
        return;
      }
      perf.bump('para.miss');
      const flow = getFlow();
      const drafts = layoutParagraph(flow, bLeft, colRight, ctx);
      const item: ParaItem = { ...mkItem(drafts), getFlow: () => flow };
      cache?.paragraphs.set(node, {
        left: bLeft,
        right: colRight,
        basePos: contentStart,
        marker,
        item,
      });
      items.push(tag({ para: item }));
    } else if (node.type.name === 'table') {
      // Cache hit: clone the canonical layout (shifting PM positions if it
      // moved) instead of re-measuring every cell. List-bearing tables are
      // never cached — they advance the live numbering counter.
      const hit = cache?.tables.get(node);
      if (hit && hit.left === bLeft && hit.right === colRight) {
        perf.bump(offset !== hit.basePos ? 'table.shift' : 'table.hit');
        items.push(
          tag({ table: cloneTableShifted(hit.table, offset - hit.basePos) }),
        );
        return;
      }
      perf.bump('table.miss');
      const table = layoutTable(
        tableToFlow(node, ctx.base, offset, counter),
        bLeft,
        colRight,
        ctx,
        bs.columns.count > 1, // clamp only inside a narrow section column
      );
      if (cache && !tableHasList(node)) {
        cache.tables.set(node, {
          left: bLeft,
          right: colRight,
          basePos: offset,
          table,
        });
        items.push(tag({ table: cloneTableShifted(table, 0) }));
      } else {
        items.push(tag({ table }));
      }
    }
  });
  buildEnd();
  // Miss = re-measured this pass; shift = cache hit but drafts re-positioned
  // (every block after the edit point); hit = pure cache reuse, no work.
  perf.counters(
    (c) =>
      `blocks: ${c['para.miss'] ?? 0}+${c['table.miss'] ?? 0} miss · ` +
      `${(c['para.shift'] ?? 0) + (c['table.shift'] ?? 0)} shifted · ` +
      `${(c['para.hit'] ?? 0) + (c['table.hit'] ?? 0)} pure-hit`,
  );
  // Pre-lay footnote bodies so the placer knows their heights up front.
  let fnMap: Map<number, FootnoteBody> | undefined;
  if (footnotes) {
    fnMap = new Map();
    for (const key of Object.keys(footnotes)) {
      const num = Number(key);
      fnMap.set(num, layoutFootnoteBody(footnotes[num], left, right, ctx));
    }
  }
  perf.span('assignSectionHeights', () => assignSectionHeights(items));
  // Band bounds as a function of page geometry (+ the active section's chrome
  // set when per-section chrome is present): the inset applies to whichever
  // page a section switches to.
  const bandFor = (p: PageConfig, chromeIndex?: number) => {
    let t = p.margin.top;
    let b = p.height - p.margin.bottom;
    const sh = chromeIndex != null ? setBandH[chromeIndex] : undefined;
    if (sh) {
      if (sh.header > 0) t = Math.max(t, headerDist(p) + sh.header);
      if (sh.footer > 0) b = Math.min(b, p.height - footerDist(p) - sh.footer);
      return { top: t, bottom: b };
    }
    if (headers.default || headers.first || headers.even)
      t = Math.max(t, headerDist(p) + maxBandHeight(headers));
    if (footers.default || footers.first || footers.even)
      b = Math.min(b, p.height - footerDist(p) - maxBandHeight(footers));
    return { top: t, bottom: b };
  };
  // Page-cache fingerprint (Phase 3): every input a page depends on besides
  // its items and carry. A mismatch (page setup, chrome height, tab grid,
  // footnote heights) invalidates naturally by keying a different run.
  const pageKey =
    cache &&
    JSON.stringify({
      g: [
        page.width,
        page.height,
        page.margin.top,
        page.margin.right,
        page.margin.bottom,
        page.margin.left,
        page.headerDistance ?? -1,
        page.footerDistance ?? -1,
        page.gutter ?? 0,
      ],
      tw: config.tabWidth ?? -1,
      // LayoutConfig.defaultFont is a Partial<FontSpec>, so a host can set a
      // document-wide glyph adjustment. Node identity covers everything that
      // rides a mark; this covers the one input that does not.
      gf: glyphKey({ ...DEFAULT_FONT, ...config.defaultFont }),
      hb: [maxBandHeight(headers), maxBandHeight(footers)],
      sb: setBandH.map((s) => [s.header, s.footer]),
      fn: fnMap
        ? [...fnMap.entries()].map(([n, b]) => [n, Math.round(b.height * 10)])
        : 0,
    });
  const resolved = perf.span('placeBlocks', () =>
    placeBlocks(
      items,
      config,
      ctx,
      bandFor,
      fnMap,
      cache && pageKey ? { store: cache.pageRuns, key: pageKey } : undefined,
    ),
  );

  // Display page numbers (w:pgNumType): a fresh O(pages) pass over the final
  // page list every layout — cached/reused pages are read, never mutated.
  const pageLabels = computePageLabels(resolved.pages, sections);
  if (pageLabels) resolved.pageLabels = pageLabels;

  // Chrome with page-number fields: re-lay every variant now that the page
  // total is known, so each field slot is as wide as the widest number shown.
  const setsHaveFields = (chromeSets ?? []).some(
    (s) =>
      anyBandHasFields({
        default: s.header,
        first: s.headerFirst,
        even: s.headerEven,
      }) ||
      anyBandHasFields({
        default: s.footer,
        first: s.footerFirst,
        even: s.footerEven,
      }),
  );
  if (
    anyBandHasFields(headers) ||
    anyBandHasFields(footers) ||
    setsHaveFields
  ) {
    // The slot must fit the widest value a PAGE/NUMPAGES field can show:
    // the page total, or the longest display label ("xxviii" beats "28").
    const fieldCtx: Ctx = {
      ...ctx,
      fieldPlaceholder: (pageLabels ?? []).reduce(
        (a, b) => (b.length > a.length ? b : a),
        String(resolved.pages.length),
      ),
    };
    headers = layHeaders(fieldCtx);
    footers = layFooters(fieldCtx);
    layChromeSets(fieldCtx); // rebuilds chromeSets with the real page total
  }
  if (chromeSets) resolved.chromeSets = chromeSets;

  if (headers.default) resolved.pageHeader = headers.default;
  if (headers.first) resolved.pageHeaderFirst = headers.first;
  if (headers.even) resolved.pageHeaderEven = headers.even;
  if (footers.default) resolved.pageFooter = footers.default;
  if (footers.first) resolved.pageFooterFirst = footers.first;
  if (footers.even) resolved.pageFooterEven = footers.even;
  if (chrome?.titlePg || chrome?.evenAndOdd) {
    resolved.chromeSelect = {
      titlePg: !!chrome.titlePg,
      evenAndOdd: !!chrome.evenAndOdd,
    };
  }
  return resolved;
}
