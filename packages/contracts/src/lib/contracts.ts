// Shared layout contracts for bapbong. Pure types only — no runtime, no DOM.

/** Page geometry, in CSS pixels. */
export interface PageConfig {
  width: number;
  height: number;
  margin: { top: number; right: number; bottom: number; left: number };
}

/** A resolved font used for both measuring and painting. */
export interface FontSpec {
  family: string;
  /** Size in points (CSS `pt`). */
  sizePt: number;
  bold: boolean;
  italic: boolean;
}

/** Measures the width (CSS px) of `text` rendered with `font`. */
export type MeasureText = (text: string, font: FontSpec) => number;

/** Vertical font metrics (CSS px) for baseline-accurate line boxes. */
export interface FontMetrics {
  /** Distance from the baseline up to the top of the line box. */
  ascent: number;
  /** Distance from the baseline down to the bottom of the line box. */
  descent: number;
}

/** Provides vertical metrics for a font. Injected so the engine stays pure. */
export type MeasureMetrics = (font: FontSpec) => FontMetrics;

// ── Flow input (document flattened, ready for layout) ──────────────

/** A contiguous run of inline text sharing one font/color/link. */
export interface InlineRun {
  text: string;
  font: FontSpec;
  color?: string;
  link?: string;
  underline?: boolean;
  strike?: boolean;
  /** Absolute ProseMirror position of the run's first character. */
  pos?: number;
}

/** An atomic inline image laid out inline with text. Dimensions are CSS px. */
export interface InlineImage {
  src: string;
  width: number;
  height: number;
  link?: string;
  /** Absolute ProseMirror position of the image node (occupies 1 position). */
  pos?: number;
}

/** One piece of a paragraph's inline content. Distinguish with `'src' in x`. */
export type FlowInline = InlineRun | InlineImage;

/** Paragraph horizontal alignment (mirrors w:jc / CSS text-align). */
export type Align = 'left' | 'center' | 'right' | 'justify';

/** Paragraph indentation in CSS px (mirrors w:ind). `firstLine` and `hanging`
 *  are mutually exclusive; if both are present, `hanging` takes precedence. */
export interface ParagraphIndent {
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
}

/** A block flattened and ready for layout (paragraph only, for now). */
export interface FlowParagraph {
  type: 'paragraph';
  /** Ordered inline content: text runs and atomic images. */
  runs: FlowInline[];
  /** List marker text (e.g. "1.", "•") if this paragraph is a list item. */
  marker?: string;
  /** Horizontal alignment; defaults to 'left' when omitted. */
  align?: Align;
  /** Indentation in CSS px; defaults to no indent when omitted. */
  indent?: ParagraphIndent;
  /** Absolute PM position where the paragraph's content starts (nodePos + 1). */
  pos?: number;
  /** Absolute PM position after the paragraph's last character. */
  end?: number;
}

/** A table cell, holding nested flow content (paragraphs / tables). */
export interface FlowTableCell {
  colspan: number;
  rowspan: number;
  /** Px widths of the spanned grid columns, or null to derive equally. */
  colwidth: number[] | null;
  content: FlowBlock[];
}

export interface FlowTableRow {
  cells: FlowTableCell[];
  /** Repeat this row at the top of every page the table spans (w:tblHeader).
   *  Only honored for contiguous header rows at the top of the table. */
  header?: boolean;
}

/** A table flattened and ready for layout. */
export interface FlowTable {
  type: 'table';
  rows: FlowTableRow[];
}

export type FlowBlock = FlowParagraph | FlowTable;

export interface LayoutConfig {
  page: PageConfig;
  measureText: MeasureText;
  /** Optional vertical metrics. When provided, line boxes use real
   *  ascent/descent; otherwise a font-size factor approximates them. */
  measureMetrics?: MeasureMetrics;
  /** Default tab-stop interval in px (Word's default is 48px = 0.5in). */
  tabWidth?: number;
  /** Defaults for runs that don't specify a value. */
  defaultFont?: Partial<FontSpec>;
}

// ── Resolved (paint-ready) output ──────────────────────────────────

/** One painted text segment, positioned in page coordinates (px). */
export interface LayoutSegment {
  x: number;
  text: string;
  font: FontSpec;
  color?: string;
  link?: string;
  underline?: boolean;
  strike?: boolean;
  /** Measured width (px) — lets the painter draw text decorations without
   *  re-measuring at paint time. */
  width?: number;
  /** Absolute PM position of the segment's first character. Segments without
   *  a position (list markers) are decoration — not addressable by a caret. */
  pos?: number;
}

/** A painted inline image; (x) is its left edge, sitting on the line baseline.
 *  width/height are CSS px. */
export interface LayoutImageSegment {
  x: number;
  src: string;
  width: number;
  height: number;
  link?: string;
  /** Absolute PM position of the image node (occupies 1 position). */
  pos?: number;
}

/** A laid-out line; (x, y) is its top-left in page coordinates (px). */
export interface LayoutLine {
  x: number;
  y: number;
  width: number;
  height: number;
  /** Baseline offset from the line's top (px). */
  baseline: number;
  segments: LayoutSegment[];
  /** Inline images on this line, positioned like segments. */
  images?: LayoutImageSegment[];
  /** PM position of the line's first caret slot. */
  from?: number;
  /** PM position after the line's last painted content. */
  to?: number;
}

/** A laid-out table cell. All coordinates are page-absolute (px); `lines` and
 *  `tables` are the cell's content, already positioned inside the cell box. */
export interface ResolvedCell {
  x: number;
  y: number;
  width: number;
  height: number;
  colspan: number;
  rowspan: number;
  lines: LayoutLine[];
  /** Tables nested inside this cell. */
  tables?: ResolvedTable[];
}

/** A laid-out table; (x, y) is its top-left in page coordinates (px). */
export interface ResolvedTable {
  x: number;
  y: number;
  width: number;
  height: number;
  cells: ResolvedCell[];
  /** Bottom edge (px from the table top) of the repeating header band, when
   *  the table's leading rows are marked as header rows. */
  headerBottom?: number;
}

export interface ResolvedPage {
  index: number;
  width: number;
  height: number;
  lines: LayoutLine[];
  /** Tables on this page, positioned in page coordinates. */
  tables?: ResolvedTable[];
}

/** Repeating page furniture (a header or footer band). Coordinates are
 *  page-local; the painter stamps it onto every page. Not caret-addressable —
 *  PM positions are stripped (the band belongs to a separate document). */
export interface ResolvedChrome {
  lines: LayoutLine[];
  tables: ResolvedTable[];
  height: number;
}

/** The paint-ready result the canvas painter consumes (M3). */
export interface ResolvedLayout {
  pages: ResolvedPage[];
  pageHeader?: ResolvedChrome;
  pageFooter?: ResolvedChrome;
}

// ── Interaction (M4): caret, selection, hit-testing ────────────────

/** A point in page-local coordinates (px), tagged with its page. */
export interface PagePoint {
  pageIndex: number;
  x: number;
  y: number;
}

/** Caret placement in page-local coordinates (px). */
export interface CaretRect {
  pageIndex: number;
  x: number;
  y: number;
  height: number;
}

/** One highlighted rectangle of a selection, in page-local coordinates. */
export interface SelectionRect {
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
}
