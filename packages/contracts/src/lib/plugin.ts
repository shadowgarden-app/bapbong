import type { EditorState, Transaction } from 'prosemirror-state';
import type { MarkSpec, NodeSpec } from 'prosemirror-model';
import type { CaretRect, PagePoint, ResolvedLayout, SelectionRect } from './contracts.js';

/** A decoration a plugin paints over a document range (comment tint, find
 *  highlight, track-change underline…). Doc positions; the editor resolves them
 *  to page geometry before painting. */
export interface RangeDecoration {
  from: number;
  to: number;
  kind: 'background' | 'underline' | 'strike';
  color: string;
}

/** A {@link RangeDecoration} resolved to page-local rects — what the painter
 *  consumes (it never maps doc positions itself). */
export interface PaintDecoration {
  rects: SelectionRect[];
  kind: 'background' | 'underline' | 'strike';
  color: string;
}

/**
 * Emitted after every editor layout/paint cycle so a host (any framework) and
 * plugins can mirror document state — page count, comment threads (on
 * doc.attrs), selection.
 */
export interface EditorChange {
  /** The current editor state (doc + selection). */
  state: EditorState;
  /** Number of laid-out pages. */
  pageCount: number;
  /** Whether the doc changed this cycle (false = selection-only). Consumers use
   *  it to skip rebuilding doc-derived UI (comment sidebars, JSON panels). */
  docChanged: boolean;
}

/**
 * A pointer event on the canvas, offered to plugins (via
 * {@link EditorPlugin.onPointer}) before the editor's own caret/selection
 * handling. `point` is page-local geometry; `pos` is the document position under
 * it (computed for `down`/`up`/`contextmenu` — null for `move`, which fires on
 * every hover, to keep hover cheap). A plugin returns true to claim the event.
 */
export interface EditorPointerEvent {
  type: 'down' | 'move' | 'up' | 'contextmenu';
  /** Page-local point under the pointer, or null if outside any page. */
  point: PagePoint | null;
  /** Document position under the point (null for `move`). */
  pos: number | null;
  /** Viewport coordinates (for positioning UI like a context menu). */
  clientX: number;
  clientY: number;
  /** Mouse buttons bitmask. */
  buttons: number;
}

/** A transient vertical guide line shown during a drag (page-local geometry) —
 *  e.g. a column-resize preview. The editor positions it on the canvas. */
export interface OverlayGuide {
  pageIndex: number;
  x: number;
  y: number;
  height: number;
}

/** A page-local rect the editor fills as a translucent highlight (e.g. a
 *  selected table-cell block). */
export interface OverlayRect {
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * The controlled surface an {@link EditorPlugin} gets onto the editor core:
 * read/dispatch document state plus the geometry + paint helpers it needs to
 * anchor UI. The editor hands one of these to each plugin's `setup`.
 */
export interface PluginContext {
  /** Current editor state (doc + selection). Live — read it each time. */
  readonly state: EditorState;
  /** Apply a ProseMirror transaction. */
  dispatch(tr: Transaction): void;
  /** Caret geometry at a doc position (page-local), or null. */
  caretRect(pos: number): CaretRect | null;
  /** Map a page-local point to canvas-stack coordinates, or null. */
  pageToCanvas(p: { pageIndex: number; x: number; y: number }): { x: number; y: number } | null;
  /** Move the selection (caret if `to` omitted). */
  setSelection(from: number, to?: number): void;
  /** Scroll the viewport so the caret at `pos` sits `topMargin` px from the top. */
  scrollToPos(pos: number, topMargin?: number): void;
  /** Force a content repaint (e.g. after a plugin's decorations change). */
  requestPaint(): void;
  /** The current paint-ready layout (page/table/cell geometry), or null before
   *  a document loads. Lets pointer plugins hit-test tables/columns. */
  readonly layout: ResolvedLayout | null;
  /** Set the canvas cursor (e.g. `'col-resize'`); null restores the default. */
  setCursor(cursor: string | null): void;
  /** Show a transient vertical guide (e.g. a column-resize preview), or null to
   *  clear it. Page-local geometry; the editor positions it on the canvas. */
  setGuide(guide: OverlayGuide | null): void;
  /** Fill these page-local rects as a translucent highlight (e.g. a selected
   *  table-cell block), or null to clear. The editor renders them on the canvas. */
  setHighlight(rects: OverlayRect[] | null): void;
}

/**
 * A framework-agnostic editor extension. The editor invokes these hooks; the
 * plugin reaches back through {@link PluginContext}. Every hook is optional — a
 * plugin implements only what it needs.
 *
 * This is the stable contract both the editor (which calls the hooks) and
 * plugin packages (which implement them) depend on, so neither needs to import
 * the other. Decoration / schema / docx-part hooks are introduced in later
 * phases; adding them is purely additive.
 */
export interface EditorPlugin {
  /** Stable identifier (also used for diagnostics). */
  readonly name: string;
  /** Schema contributions merged into the editor's document schema, so a plugin
   *  can own its marks/nodes. Composed once when a document loads. */
  schema?: { marks?: Record<string, MarkSpec>; nodes?: Record<string, NodeSpec> };
  /** Called once when the editor is constructed; may return a teardown fn. */
  setup?(ctx: PluginContext): void | (() => void);
  /** Called after every editor layout/paint cycle. */
  onChange?(change: EditorChange): void;
  /** Called when a pointerdown places the caret at doc position `pos`. */
  onCaretPick?(pos: number): void;
  /** Decorations to paint this content frame (collected on every content
   *  repaint; call `ctx.requestPaint()` when they change). */
  decorations?(ctx: PluginContext): RangeDecoration[];
  /** Pointer activity on the canvas, offered before the editor's own
   *  caret/selection handling. Return true to claim the event (the editor then
   *  preventDefaults it, captures the pointer on `down`, and skips its default).
   *  `move` fires on every hover (for cursor feedback / drag). */
  onPointer?(ev: EditorPointerEvent): boolean;
}
