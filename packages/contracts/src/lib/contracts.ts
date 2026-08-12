// Shared layout contracts for bapbong. Pure types only — no runtime, no DOM.

/** Page geometry, in CSS pixels. */
export interface PageConfig {
  width: number;
  height: number;
  margin: { top: number; right: number; bottom: number; left: number };
  /** Header/footer band distance from the page edge (w:pgMar @w:header/
   *  @w:footer). Absent → Word's default 720 twips (48px). */
  headerDistance?: number;
  footerDistance?: number;
  /** Binding gutter (w:pgMar @w:gutter) added to the left content edge.
   *  Absent → 0. */
  gutter?: number;
}

/** Multi-column layout (w:cols): `count` equal-width columns separated by
 *  `gap` px. count 1 is the ordinary single-column flow. */
export interface ColumnConfig {
  count: number;
  gap: number;
}

/** One document section's flow properties. Sections are delimited by section
 *  breaks (w:sectPr); each spans `blockCount` top-level blocks. `newPage` is
 *  true for a next-page break (the section starts a fresh page), false for a
 *  continuous break (columns switch mid-page). */
export interface SectionConfig {
  blockCount: number;
  columns: ColumnConfig;
  newPage: boolean;
  /** Page geometry override (w:pgSz/w:pgMar on this section's sectPr) when it
   *  differs from the document default (`doc.attrs.page`) — e.g. a landscape
   *  section inside a portrait document. Absent → the document geometry.
   *  Geometry can only change at a page boundary, so a continuous section
   *  with a differing `page` is laid out as next-page (Word's own promotion). */
  page?: PageConfig;
  /** Page-number restart/format (w:pgNumType) declared on this section's
   *  sectPr. Absent → numbering continues from the previous section in
   *  decimal. `fmt` keeps the raw OOXML ST_NumberFormat value so unmodelled
   *  formats round-trip; the layout falls back to decimal when it doesn't
   *  recognize one. */
  pageNumbers?: PageNumbering;
}

/** w:pgNumType: `start` restarts the page counter at the section's first
 *  page; `fmt` names the display format ("lowerRoman", "decimal", …). */
export interface PageNumbering {
  start?: number;
  fmt?: string;
}

/** `doc.attrs.sectionChromeOverrides` — per-section header/footer story
 *  overrides, the model's first chrome EDIT (the page-number toggle). Keyed
 *  by section index (as a string — attrs are JSON); each story variant holds
 *  the FULL replacement story doc as ProseMirror JSON. An override severs
 *  that story's "Link to Previous" inheritance for that section only; export
 *  writes it as a real header/footer part. */
export type SectionChromeOverrides = Record<
  string,
  {
    headers?: Record<string, unknown>;
    footers?: Record<string, unknown>;
  }
>;

/** A document comment (w:comment) referenced by a w:commentRange in the body.
 *  `id` matches the comment mark's ids; `text` is the flattened comment body. */
export interface CommentData {
  id: number;
  author: string;
  date: string;
  text: string;
}

/** Author of a comment / reply. `id` identifies the user (for own-only edit
 *  checks and @mentions); name/email/avatar are for display. */
export interface IUser {
  id: string;
  name: string;
  email?: string;
  avatar?: string;
}

/** A comment thread node stored on `doc.attrs.comments` (authoring model). The
 *  thread is a tree via `parentId` (null = root, the node a comment mark
 *  anchors to). `body` is commentSchema ProseMirror-doc JSON. `resolved` is
 *  meaningful only on roots. */
export interface CommentNode {
  id: number;
  parentId: number | null;
  user: IUser;
  date: string;
  body: unknown;
  resolved?: boolean;
}

/**
 * Everything the Font dialog reads and writes, shared so the command layer and
 * the widget cannot drift apart on what a field means.
 *
 * Two absent-ish values, deliberately distinct:
 *   `undefined` — leave it alone. The selection was mixed and the user never
 *                 touched the control, so writing anything would flatten
 *                 formatting they never looked at.
 *   `null`      — clear it (Automatic colour, no highlight, no font override).
 *
 * The three spacing values carry no absent case at all: their controls always
 * hold a definite number, and the default (100% / 0pt / 0pt) IS the clear.
 */
export interface CharacterFormatting {
  family?: string | null;
  sizePt?: number | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  doubleStrike?: boolean;
  smallCaps?: boolean;
  vertAlign?: 'super' | 'sub' | null;
  color?: string | null;
  highlight?: string | null;
  /** w:w, percent. 100 = normal. */
  scalePercent: number;
  /** w:spacing, twips. 0 = normal. */
  letterSpacingTwips: number;
  /** w:position, half-points. 0 = on the baseline. */
  positionHalfPoints: number;
}

/** The face itself — everything the CSS font shorthand can say. */
export interface FontFace {
  family: string;
  /** Size in points (CSS `pt`). */
  sizePt: number;
  bold: boolean;
  italic: boolean;
}

/** Per-run glyph adjustments. These change how WIDE text runs, never how TALL:
 *  no adjustment here can move a baseline, which is why {@link MeasureMetrics}
 *  takes only a {@link FontFace} and cannot see them. */
export interface GlyphAdjust {
  /** w:spacing — extra advance added after every character, px. Absolute:
   *  it does not shrink with the font (superscript keeps the same tracking). */
  letterSpacing?: number;
  /** w:w — horizontal glyph scale, 1 = normal. Scales the glyphs and their
   *  advances only; the tracking above rides on top at its absolute value. */
  scaleX?: number;
  /** w:kern — whether pair kerning applies. The OOXML value is a THRESHOLD in
   *  half-points ("kern at this size and above"), resolved against the run's
   *  own size at import, so what reaches here is the yes/no answer. Absent
   *  leaves the platform default (kerned); only an explicit `false` turns it
   *  off. */
  kerning?: boolean;
}

/** A resolved font used for both measuring and painting. */
export interface FontSpec extends FontFace, GlyphAdjust {}

/** Measures the width (CSS px) of `text` rendered with `font`. */
export type MeasureText = (text: string, font: FontSpec) => number;

/** Vertical font metrics (CSS px) for baseline-accurate line boxes. */
export interface FontMetrics {
  /** Distance from the baseline up to the top of the line box. */
  ascent: number;
  /** Distance from the baseline down to the bottom of the line box. */
  descent: number;
}

/** Provides vertical metrics for a face. Injected so the engine stays pure.
 *  Takes {@link FontFace}, not {@link FontSpec}: an implementation physically
 *  cannot read a glyph adjustment, so it cannot let one move a baseline. */
export type MeasureMetrics = (font: FontFace) => FontMetrics;

// ── Flow input (document flattened, ready for layout) ──────────────

/** A contiguous run of inline text sharing one font/color/link. */
export interface InlineRun {
  text: string;
  font: FontSpec;
  color?: string;
  link?: string;
  underline?: boolean;
  strike?: boolean;
  /** Double strikethrough (w:dstrike). */
  dstrike?: boolean;
  /** Small caps (w:smallCaps): lowercase renders as reduced uppercase. */
  smallCaps?: boolean;
  /** Highlight / shading background color, e.g. "#FFFF00". */
  background?: string;
  /** Superscript / subscript (font already reduced; painter shifts baseline). */
  vertAlign?: 'super' | 'sub';
  /** w:position — baseline shift in px, positive UP, at full glyph size.
   *  Independent of vertAlign; a run can carry both. */
  raise?: number;
  /** Footnote reference number (w:footnoteReference): the run's text is the
   *  superscript mark; the body is laid out at the bottom of its page. */
  footnoteRef?: number;
  /** Comment ids covering this run (w:commentRangeStart/End) — painted with a
   *  comment tint; the sidebar locates the range by these ids. */
  commentIds?: number[];
  /** Absolute ProseMirror position of the run's first character. */
  pos?: number;
}

/** A vector shape riding an image box (drawn, not a bitmap — `src` stays '').
 *  Word drawing shapes (wps/VML rect + straight connector) map to this: same
 *  layout semantics as an image (atomic box, inline or anchored), different
 *  paint. Dimensions live on the carrying image; colors are CSS. */
export interface ShapeSpec {
  /** Preset geometry. Names mirror OOXML prstGeom tokens: 'line' is a
   *  corner-to-corner straight connector; 'rightArrow' a block arrow;
   *  'horizontalScroll' paints as a stylized banner with rolled ends. */
  kind:
    | 'rect'
    | 'line'
    | 'ellipse'
    | 'roundRect'
    | 'rightArrow'
    | 'horizontalScroll';
  /** Stroke color; absent = no stroke (a:noFill on the outline). */
  stroke?: string;
  /** Stroke width in px (defaults to 1 when a stroke is drawn). */
  strokeWidth?: number;
  /** Rect fill; absent = transparent (outline only). */
  fill?: string;
  /** Line drawn bottom-left → top-right instead (a:xfrm flipV). */
  flipV?: boolean;
  /** Arrowheads on a 'line' connector (VML startarrow/endarrow, DrawingML
   *  a:headEnd/a:tailEnd). Painted as filled triangles scaled from the
   *  stroke width. */
  arrowStart?: boolean;
  arrowEnd?: boolean;
  /**
   * Dash pattern as alternating dash/gap lengths **in multiples of the stroke
   * width** — the unit both dialects use ("The lengths are relative to the
   * line width: a length of 1 is equal to the line width", VML v:stroke
   * dashstyle; DrawingML a:custDash likewise scales by line width). Absent =
   * solid. Keeping the real numbers rather than a boolean is what lets the
   * painter reproduce the document's own density instead of inventing one.
   */
  dash?: number[];
  /** Stroke end cap (VML v:stroke endcap, DrawingML a:ln@cap). Absent = the
   *  spec default, `flat`. Only meaningful on open paths ('line'). */
  cap?: 'flat' | 'square' | 'round';
  /**
   * Corner rounding of a 'roundRect', as a FRACTION OF THE SHORTER SIDE —
   * 0 square, 0.5 fully rounded. Deliberately not called `cornerRadius`: it
   * is not px. Each dialect states it differently and the importer
   * normalizes: VML `arcsize` is a fraction of HALF the shorter side (so it
   * is halved on the way in), DrawingML `a:gd` adj is adj/100000 of the
   * shorter side. Absent → the painter's DrawingML default (0.16667).
   */
  cornerRatio?: number;
}

/** `a:srcRect` — the part of the bitmap the box shows, as ratios of the
 *  bitmap's own size measured inward from each edge. Positive crops in;
 *  NEGATIVE outsets, reaching past the bitmap so the overhang renders as
 *  nothing. The selected region scales to fill the box, so a crop changes
 *  what is visible, never the box. */
export interface ImageCrop {
  l: number;
  t: number;
  r: number;
  b: number;
}

/** An atomic inline image laid out inline with text. Dimensions are CSS px. */
export interface InlineImage {
  src: string;
  width: number;
  height: number;
  /** Visible sub-rectangle of the bitmap (a:srcRect). */
  crop?: ImageCrop;
  link?: string;
  /** Present when this box is a drawn vector shape instead of a bitmap. */
  shape?: ShapeSpec;
  /** Clockwise degrees around the box center — paint-only: the layout box
   *  stays axis-aligned (Word re-wraps only on commit, not live). */
  rotation?: number;
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

/** A forced line break inside a paragraph (w:br). Occupies one PM position. */
export interface InlineBreak {
  break: true;
  pos?: number;
}

/** One piece of a paragraph's inline content. Distinguish with `'src' in x`
 *  (image) / `'field' in x` (field) / `'break' in x` / otherwise text run. */
export type FlowInline = InlineRun | InlineImage | InlineField | InlineBreak;

/** A floating image (wp:anchor) anchored to a paragraph. Text flows around
 *  its rectangle ('square'), skips below it ('topAndBottom'), or ignores it
 *  ('none' — painted only). All values are CSS px. */
export interface FlowFloat {
  src: string;
  width: number;
  height: number;
  /** Visible sub-rectangle of the bitmap (a:srcRect). */
  crop?: ImageCrop;
  /** Present when this float is a drawn vector shape instead of a bitmap. */
  shape?: ShapeSpec;
  wrap: 'square' | 'topAndBottom' | 'none';
  /** wp:anchor behindDoc="1" — paint UNDER the text (watermarks). Absent =
   *  Word's default: anchored drawings paint over the text. */
  behind?: boolean;
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
  /** Textbox (wps:txbx) paragraphs laid out inside the shape's box. */
  content?: FlowParagraph[];
  /** Textbox interior padding (wps:bodyPr lIns…), px. Absent → Word defaults. */
  inset?: { l: number; t: number; r: number; b: number };
  /** Where the text block sits vertically in the box (wps:bodyPr @anchor).
   *  Absent → the schema default, top. */
  anchor?: 'ctr' | 'b';
  /** Clockwise degrees around the box center — paint-only: wrap exclusions
   *  keep the axis-aligned box. */
  rotation?: number;
  /** Absolute PM position of the carrying image node — lets the editor map a
   *  resolved float back to its node (resize). Absent in chrome (not
   *  caret-addressable). */
  pos?: number;
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

/** Paragraph spacing (mirrors w:spacing). `before`/`after` are px gaps around
 *  the paragraph; `line` is the inter-line measure interpreted by `lineRule`:
 *  'auto' → multiple of the natural line height, 'exact'/'atLeast' → px. */
export interface ParagraphSpacing {
  before?: number;
  after?: number;
  line?: number;
  lineRule?: 'auto' | 'exact' | 'atLeast';
}

/** List label styling from the numbering level (w:lvlJc / w:suff / lvl rPr).
 *  All fields default to Word's own defaults: left-aligned, tab suffix,
 *  label drawn with the paragraph's base font. */
export interface MarkerStyle {
  /** Label alignment against its anchor. */
  jc?: 'center' | 'right';
  /** Separator between label and text. */
  suff?: 'space' | 'nothing';
  /** Label font overrides (family/size/bold/italic from the lvl rPr). */
  font?: Partial<FontSpec>;
  /** Label color ("#RRGGBB"). */
  color?: string;
}

/** A block flattened and ready for layout (paragraph only, for now). */
export interface FlowParagraph {
  type: 'paragraph';
  /** Ordered inline content: text runs and atomic images. */
  runs: FlowInline[];
  /** List marker text (e.g. "1.", "•") if this paragraph is a list item. */
  marker?: string;
  /** Styling for `marker` (alignment, suffix, its own font/color). */
  markerStyle?: MarkerStyle;
  /** Horizontal alignment; defaults to 'left' when omitted. */
  align?: Align;
  /** Indentation in CSS px; defaults to no indent when omitted. */
  indent?: ParagraphIndent;
  /** Line spacing + space before/after; defaults to single, no gaps. */
  spacing?: ParagraphSpacing;
  /** Force this paragraph to start a new page (w:pageBreakBefore / a page
   *  break run at its head). */
  pageBreakBefore?: boolean;
  /** w:keepNext — stay on the same page as the next block's first line. */
  keepNext?: boolean;
  /** w:keepLines — never split this paragraph across pages when it fits a
   *  full band. */
  keepLines?: boolean;
  /** w:widowControl — false disables widow/orphan control (Word's default
   *  is ON; absent means on). */
  widowControl?: boolean;
  /** Absolute PM position where the paragraph's content starts (nodePos + 1). */
  pos?: number;
  /** Absolute PM position after the paragraph's last character. */
  end?: number;
  /** Floating images anchored to this paragraph. */
  floats?: FlowFloat[];
  /** Custom tab stops; tabs past the last stop use the default grid. */
  tabs?: TabStop[];
  /** Paragraph box borders (w:pBdr) — the four outer sides. */
  borders?: ParagraphBorders;
  /** w:shd fill "#RRGGBB" painted behind the paragraph's lines. */
  shading?: string;
}

/** Paragraph box borders (w:pBdr). Only visible sides are present; the
 *  `between` edge is not modelled. */
export interface ParagraphBorders {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
}

/** A table cell, holding nested flow content (paragraphs / tables). */
export interface FlowTableCell {
  colspan: number;
  rowspan: number;
  /** Px widths of the spanned grid columns, or null to derive equally. */
  colwidth: number[] | null;
  /** Cell fill color (w:shd), e.g. "#D9E2F3". */
  background?: string;
  /** Vertical alignment of content within the cell (w:vAlign); top default. */
  vAlign?: 'center' | 'bottom';
  /** Per-cell border overrides (w:tcBorders); each side overrides the table. */
  borders?: TableBorders;
  /** Corner-to-corner rules across the cell (w:tl2br / w:br2tl). */
  diagonals?: CellDiagonals;
  /** Per-cell margin overrides (w:tcMar); each side overrides the table's
   *  cellPadding (and that the Word defaults). */
  padding?: CellPadding;
  content: FlowBlock[];
}

export interface FlowTableRow {
  cells: FlowTableCell[];
  /** Repeat this row at the top of every page the table spans (w:tblHeader).
   *  Only honored for contiguous header rows at the top of the table. */
  header?: boolean;
  /** w:cantSplit — never break this row across pages (Word default allows). */
  cantSplit?: boolean;
  /** Explicit row height (w:trHeight): a floor, or `exact` to force it. */
  height?: { value: number; exact: boolean };
}

/** Which table borders are visible (w:tblBorders). OOXML tables have NO
 *  borders unless declared — absence of this object means draw nothing. */
/** Stroke style of a visible border edge (maps to/from w:val). */
export type BorderStyle = 'solid' | 'dashed' | 'dotted' | 'double';

/** Appearance of one visible border edge. */
export interface BorderSide {
  /** Stroke width in CSS px. */
  width: number;
  style: BorderStyle;
  /** Hex colour, e.g. "#000000". */
  color: string;
  /** OOXML w:space — gap between the border and the content, in px. Optional
   *  (0/absent for the common case); round-tripped, not yet painted. */
  space?: number;
}

/**
 * Per-side border appearance. A side holding a {@link BorderSide} is visible
 * with that width/style/colour; `false` is an explicit "no border"; an absent
 * side inherits (a cell falls back to the table's outer/inside border).
 */
export interface TableBorders {
  top?: BorderSide | false;
  bottom?: BorderSide | false;
  left?: BorderSide | false;
  right?: BorderSide | false;
  insideH?: BorderSide | false;
  insideV?: BorderSide | false;
}

/**
 * Diagonal rules across a table cell (w:tcBorders/w:tl2br and w:br2tl). Not
 * sides — they cross the cell corner to corner, and Vietnamese school tables
 * use them to strike out a cell that has no data. Both may be present, which
 * draws an X.
 */
export interface CellDiagonals {
  /** Top-left to bottom-right. */
  tl2br?: BorderSide;
  /** Bottom-left to top-right. */
  br2tl?: BorderSide;
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
  borders?: TableBorders;
  /** Table alignment within the content area (w:jc); left default. */
  align?: 'center' | 'right';
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
  /** Whole-document column config for the flat (FlowBlock) layout path. The PM
   *  doc path reads per-section columns from the doc instead. */
  columns?: ColumnConfig;
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
  /** Double strikethrough (w:dstrike): the painter draws two thin lines. */
  dstrike?: boolean;
  /** Highlight / shading background painted behind the text. */
  background?: string;
  /** Superscript / subscript: the painter shifts the baseline (font is
   *  already the reduced size). */
  vertAlign?: 'super' | 'sub';
  /** w:position — extra baseline shift in px, positive UP, glyphs unresized. */
  raise?: number;
  /** Measured width (px) — lets the painter draw text decorations without
   *  re-measuring at paint time. */
  width?: number;
  /** Dynamic field: the painter substitutes the page number / page count for
   *  `text` while painting. The segment occupies ONE PM position. */
  field?: 'pageNumber' | 'pageCount';
  /** Footnote reference number: the placer counts which notes land on each
   *  page so it can reserve bottom space and lay their bodies out there. */
  footnoteRef?: number;
  /** Comment ids covering this segment (w:commentRangeStart/End) — the painter
   *  tints it; null/absent means no comment. */
  commentIds?: number[];
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
  /** Present when this box is a drawn vector shape instead of a bitmap. */
  shape?: ShapeSpec;
  /** Visible sub-rectangle of the bitmap (a:srcRect). */
  crop?: ImageCrop;
  /** Clockwise degrees around the box center (paint-only). */
  rotation?: number;
  /** Absolute PM position of the image node (occupies 1 position). */
  pos?: number;
}

/** w:keepNext / w:keepLines / w:widowControl of one flowed paragraph. ONE
 *  object per paragraph, shared by reference across all its lines (see
 *  LayoutLine.keeps) — the identity is what groups a paragraph's lines back
 *  together after clones and fragment splits scatter them. */
export interface ParagraphKeeps {
  /** w:keepNext — stay with the next block's opening. */
  keepNext?: boolean;
  /** w:keepLines — never split this paragraph when it fits a page. */
  keepLines?: boolean;
  /** false disables widow/orphan control; absent = Word's default ON. */
  widowControl?: boolean;
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
  /** Pagination facts of the paragraph this line came from — shared BY
   *  REFERENCE with the paragraph's other lines so a splitter can regroup
   *  them and re-evaluate split legality per fragment (facts, not baked
   *  decisions: a second split sees the fragment's own line counts). Set for
   *  lines flowed inside table cells; layout-only, the painter ignores it. */
  keeps?: ParagraphKeeps;
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
  /** Vertical-alignment slack added above the content (w:vAlign center/
   *  bottom), in px. The row splitter removes it — Word suspends vertical
   *  centering while a row breaks across pages, re-stacking content from the
   *  cell's top. Purely informational for the painter. */
  vShift?: number;
  /** Cell fill color, painted behind the content. */
  background?: string;
  /** Per-cell border overrides; each side overrides the table's edge. */
  borders?: TableBorders;
  /** Corner-to-corner rules across the cell (w:tl2br / w:br2tl). */
  diagonals?: CellDiagonals;
  /** Tables nested inside this cell. */
  tables?: ResolvedTable[];
  /** Anchored images/shapes positioned within this cell (v1: painted at their
   *  anchor offsets; cell text does not wrap around them). */
  floats?: ResolvedFloat[];
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
  /** Vertical bands (px from the table top) of rows marked w:cantSplit — the
   *  paginator moves these whole instead of splitting them mid-content. */
  cantSplitBands?: { top: number; bottom: number }[];
  /** Vertical bands of rows whose FIRST cell opens with a w:keepNext
   *  paragraph. Word does not let such a row START in the leftover of a
   *  band — it begins on a fresh band and only then splits normally.
   *  (Word-verified: a keepNext opener in a LATER cell does not veto —
   *  nested_table.docx row 2 vs row 3.) Layout-only; the painter ignores
   *  it. Band lists re-base together on split fragments. */
  keepStartBands?: { top: number; bottom: number }[];
  /** Visible borders; absent → the table paints borderless (OOXML default). */
  borders?: TableBorders;
}

/** A floating image placed on a page (page-local coordinates, px). */
export interface ResolvedFloat {
  x: number;
  y: number;
  width: number;
  height: number;
  src: string;
  /** Visible sub-rectangle of the bitmap (a:srcRect). */
  crop?: ImageCrop;
  /** Present when this float is a drawn vector shape instead of a bitmap. */
  shape?: ShapeSpec;
  /** Textbox text laid out inside the shape, in BOX-LOCAL coordinates
   *  (origin at the float's top-left) — the painter translates by (x, y).
   *  Never caret-addressable; PM positions are stripped. */
  lines?: LayoutLine[];
  /** Clockwise degrees around the box center (paint-only). */
  rotation?: number;
  /** The offsets that WOULD render the float exactly here — i.e. resolved
   *  against its anchor bases (hRel/vRel, alignment included) at layout time.
   *  Translation-invariant, so the many places that shift a cell's floats
   *  need no bookkeeping. A drag-to-move gesture commits
   *  `hOffset = effHOffset + dx` (dropping hAlign — dragging pins the float
   *  to an explicit position, as Word does). */
  effHOffset?: number;
  effVOffset?: number;
  /** Paint under the text (wp:anchor behindDoc) — see FlowFloat.behind. */
  behind?: boolean;
  /** Absolute PM position of the carrying image node (absent in chrome). */
  pos?: number;
}

export interface ResolvedPage {
  index: number;
  width: number;
  height: number;
  /** Content-area origin (left/top margins in page coords) — lets a plugin
   *  turn a page rect back into margin-relative offsets (e.g. converting an
   *  inline image to a float that keeps its visual position). */
  contentLeft?: number;
  contentTop?: number;
  lines: LayoutLine[];
  /** Tables on this page, positioned in page coordinates. */
  tables?: ResolvedTable[];
  /** Floating images on this page (painted behind the text). */
  floats?: ResolvedFloat[];
  /** Footnote bodies whose references fall on this page, laid out at the
   *  bottom above the footer. Absent when the page has no footnotes. */
  footnotes?: ResolvedFootnotes;
  /** Paragraph boxes (w:pBdr borders and/or w:shd fill) on this page, painted
   *  under the text. */
  paraBoxes?: ParagraphBox[];
  /** Index into ResolvedLayout.chromeSets for this page's header/footer —
   *  the section in effect where the page starts. Absent → the flat
   *  document-level chrome. */
  chromeIndex?: number;
  /** True when this page is the first page of its section — w:titlePg picks
   *  the "first" chrome variant per SECTION, not per document. */
  sectionFirst?: boolean;
}

/** One paragraph's box on a page (page-local px), carrying whatever paints
 *  behind and around its lines: the w:pBdr border sides, the w:shd fill, or
 *  both. A paragraph split across pages emits one box per fragment —
 *  `drawTop` only on the first, `drawBottom` only on the last, while the fill
 *  covers every fragment. */
export interface ParagraphBox {
  x: number;
  y: number;
  width: number;
  height: number;
  /** w:pBdr sides; absent when the paragraph is shaded but unbordered. */
  borders?: ParagraphBorders;
  /** w:shd fill "#RRGGBB", painted under the borders and the text. */
  shading?: string;
  drawTop: boolean;
  drawBottom: boolean;
}

/** Footnote bodies reserved at the bottom of a page. `separatorY` is where the
 *  short rule above the notes is drawn; `lines` are the note bodies, already
 *  positioned in page coordinates (px). */
export interface ResolvedFootnotes {
  separatorY: number;
  lines: LayoutLine[];
}

/** Repeating page furniture (a header or footer band). Coordinates are
 *  page-local; the painter stamps it onto every page. Not caret-addressable —
 *  PM positions are stripped (the band belongs to a separate document). */
export interface ResolvedChrome {
  lines: LayoutLine[];
  tables: ResolvedTable[];
  /** Anchored images/shapes positioned within the band (e.g. the horizontal
   *  rule real headers draw under their contact block). Painted only — the
   *  chrome text does not wrap around them. */
  floats?: ResolvedFloat[];
  height: number;
}

/** The paint-ready result the canvas painter consumes (M3). */
export interface ResolvedLayout {
  pages: ResolvedPage[];
  /** Default (odd-page) header/footer, stamped on every page unless a
   *  first/even variant applies (see chromeSelect). */
  pageHeader?: ResolvedChrome;
  pageFooter?: ResolvedChrome;
  /** Title-page header/footer (w:type="first"), shown on page 1 when
   *  chromeSelect.titlePg is set. */
  pageHeaderFirst?: ResolvedChrome;
  pageFooterFirst?: ResolvedChrome;
  /** Even-page header/footer (w:type="even"), shown on even pages when
   *  chromeSelect.evenAndOdd is set. */
  pageHeaderEven?: ResolvedChrome;
  pageFooterEven?: ResolvedChrome;
  /** Which chrome variant applies per page. `titlePg` → page 1 uses the first
   *  variant (blank if none); `evenAndOdd` → even pages use the even variant.
   *  With chromeSets, titlePg lives per set — only evenAndOdd (a document
   *  setting) is read from here. */
  chromeSelect?: { titlePg: boolean; evenAndOdd: boolean };
  /** Per-section header/footer bands, indexed by ResolvedPage.chromeIndex.
   *  Present only when sections carry distinct chrome; otherwise the flat
   *  pageHeader/pageFooter fields above apply to every page. */
  chromeSets?: ResolvedChromeSet[];
  /** Display page number per page (w:pgNumType applied: per-section restart +
   *  format, e.g. "ii" then "1"). Present only when some section declares
   *  pgNumType; absent → the display number is `index + 1`. Consumers (PAGE
   *  fields, TOC updates) read this instead of the physical index. */
  pageLabels?: string[];
}

/** One section's laid-out chrome bands (all variants), plus its titlePg
 *  toggle. The "first" variant applies on the section's first page. */
export interface ResolvedChromeSet {
  header?: ResolvedChrome;
  footer?: ResolvedChrome;
  headerFirst?: ResolvedChrome;
  footerFirst?: ResolvedChrome;
  headerEven?: ResolvedChrome;
  footerEven?: ResolvedChrome;
  titlePg?: boolean;
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
