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

/** A dynamic field whose text is substituted at paint time (page numbers).
 *  Occupies one PM position, like an image atom. */
export interface InlineField {
  field: 'pageNumber' | 'pageCount';
  font: FontSpec;
  color?: string;
  pos?: number;
}

/** One piece of a paragraph's inline content. Distinguish with
 *  `'src' in x` (image) / `'field' in x` (field) / otherwise text run. */
export type FlowInline = InlineRun | InlineImage | InlineField;

/** A floating image (wp:anchor) anchored to a paragraph. Text flows around
 *  its rectangle ('square'), skips below it ('topAndBottom'), or ignores it
 *  ('none' — painted only). All values are CSS px. */
export interface FlowFloat {
  src: string;
  width: number;
  height: number;
  wrap: 'square' | 'topAndBottom' | 'none';
  /** Horizontal: alignment within hRel, or an offset from its left edge. */
  hAlign?: 'left' | 'right' | 'center';
  hOffset?: number;
  hRel?: 'margin' | 'page';
  /** Vertical: offset from the anchor paragraph top / margin / page top. */
  vOffset?: number;
  vRel?: 'paragraph' | 'margin' | 'page';
  /** Text-to-image gaps. */
  distL?: number;
  distR?: number;
  distT?: number;
  distB?: number;
}

/** Paragraph horizontal alignment (mirrors w:jc / CSS text-align). */
export type Align = 'left' | 'center' | 'right' | 'justify';

/** A custom tab stop (w:tabs/w:tab). `pos` is px from the paragraph's content
 *  left edge. `val` says how the following text aligns at the stop; `leader`
 *  fills the jumped-over gap (TOC dots etc.). */
export interface TabStop {
  pos: number;
  val: 'left' | 'right' | 'center' | 'decimal';
  leader?: 'dot' | 'hyphen' | 'underscore' | 'middleDot';
}

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
  /** Floating images anchored to this paragraph. */
  floats?: FlowFloat[];
  /** Custom tab stops; tabs past the last stop use the default grid. */
  tabs?: TabStop[];
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

/** Cell padding overrides (px) from w:tblCellMar; unset sides use defaults. */
export interface CellPadding {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
}

/** A table flattened and ready for layout. */
export interface FlowTable {
  type: 'table';
  rows: FlowTableRow[];
  cellPadding?: CellPadding;
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
  /** Dynamic field: the painter substitutes the page number / page count for
   *  `text` while painting. The segment occupies ONE PM position. */
  field?: 'pageNumber' | 'pageCount';
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

/** A floating image placed on a page (page-local coordinates, px). */
export interface ResolvedFloat {
  x: number;
  y: number;
  width: number;
  height: number;
  src: string;
}

export interface ResolvedPage {
  index: number;
  width: number;
  height: number;
  lines: LayoutLine[];
  /** Tables on this page, positioned in page coordinates. */
  tables?: ResolvedTable[];
  /** Floating images on this page (painted behind the text). */
  floats?: ResolvedFloat[];
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
