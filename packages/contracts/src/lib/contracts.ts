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

// ── Flow input (document flattened, ready for layout) ──────────────

/** A contiguous run of inline text sharing one font/color/link. */
export interface InlineRun {
  text: string;
  font: FontSpec;
  color?: string;
  link?: string;
}

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
  runs: InlineRun[];
  /** List marker text (e.g. "1.", "•") if this paragraph is a list item. */
  marker?: string;
  /** Horizontal alignment; defaults to 'left' when omitted. */
  align?: Align;
  /** Indentation in CSS px; defaults to no indent when omitted. */
  indent?: ParagraphIndent;
}

export type FlowBlock = FlowParagraph;

export interface LayoutConfig {
  page: PageConfig;
  measureText: MeasureText;
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
}

export interface ResolvedPage {
  index: number;
  width: number;
  height: number;
  lines: LayoutLine[];
}

/** The paint-ready result the canvas painter consumes (M3). */
export interface ResolvedLayout {
  pages: ResolvedPage[];
}
