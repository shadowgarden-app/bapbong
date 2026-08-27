import type { EditorState, Transaction } from 'prosemirror-state';
import type { MarkSpec, NodeSpec } from 'prosemirror-model';
import type {
  CaretRect,
  PagePoint,
  ResolvedLayout,
  SelectionRect,
} from './contracts.js';

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
  /** Modifier keys held — e.g. Ctrl/Cmd-click to open a hyperlink. */
  ctrlKey: boolean;
  metaKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

/**
 * Typed handles for plugins that expose a public API. A plugin AUGMENTS this
 * interface from its own module:
 *
 *   declare module '@shadow-garden/bapbong-contracts' {
 *     interface EditorPluginHandles { find: FindPlugin }
 *   }
 *
 * so `editor.plugin('find')` (host) and `ctx.plugin('find')` (a plugin that
 * declared it in `uses`) come back fully typed — and the editor core never
 * hardcodes a plugin name. Plugins without a public API don't register here.
 */
// Empty BY DESIGN: this is the augmentation target. Plugins fill it from their
// own modules, which is what keeps the core free of plugin names.
/* eslint-disable @typescript-eslint/no-empty-object-type, @typescript-eslint/no-empty-interface -- rule renamed across the eslint versions in this workspace; disable both names */
export interface EditorPluginHandles {}
/* eslint-enable @typescript-eslint/no-empty-object-type, @typescript-eslint/no-empty-interface */

/** A keyboard event offered to plugins (via {@link EditorPlugin.onKey})
 *  before the editor's own keymaps see it. */
export interface EditorKeyEvent {
  /** `KeyboardEvent.key` — e.g. "Escape", "Enter", "a". */
  key: string;
  ctrlKey: boolean;
  metaKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

/** A host-built DOM element the editor mounts into the canvas overlay.
 *  Typed opaquely because this package builds WITHOUT the DOM lib — it has to
 *  compile in a Node context — so an HTMLElement satisfies it and the editor
 *  narrows at the boundary. */
export type OverlayPanelElement = object;

/** A vertical line the editor draws on the canvas (page-local geometry).
 *  `drag` is the transient preview of a gesture — a column-resize position;
 *  `caret` is a real insertion point inside an object the hidden text editor
 *  does not own (an equation slot), so it is drawn and blinked like the
 *  document's own caret rather than like a drag guide. */
export interface OverlayGuide {
  pageIndex: number;
  x: number;
  y: number;
  height: number;
  /** Default `drag`. */
  kind?: 'drag' | 'caret';
}

/** One button on the frame's floating action strip. `svg` is inner SVG markup
 *  for a 16×16 viewBox (the editor wraps it in the <svg> element) — plugins
 *  own their icons so the editor stays generic. */
export interface OverlayFrameAction {
  id: string;
  /** Tooltip / accessible name. */
  title: string;
  svg: string;
  /** Render highlighted (the mode currently in effect). */
  active?: boolean;
  /** Render as a thin group divider instead of a button — `title`/`svg` are
   *  unused and clicks never fire (the id still keys the rebuild signature,
   *  so keep it unique). */
  separator?: boolean;
}

/** A selection frame around an object (image resize): border + 8 resize
 *  handles + a rotate knob above the top edge, rotated by `rotation` degrees
 *  around the rect center. Page-local geometry; purely visual — the plugin
 *  hit-tests handles itself from the same numbers. */
export interface OverlayFrame {
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
  /** Clockwise degrees around the rect center (default 0). */
  rotation?: number;
  /** Small readout above the frame (e.g. "320 × 214" during a resize). */
  label?: string;
  /** Floating button strip below the frame (e.g. image wrap modes). Clicks
   *  come back through {@link EditorPlugin.onFrameAction}. Omit during drag
   *  previews — the strip is for the resting selection. */
  actions?: OverlayFrameAction[];
  /**
   * Which resize handles the frame offers. `'all'` (the default) is the
   * eight-handle picture frame; `'corners'` draws only the four corner
   * handles, so the box can only scale PROPORTIONALLY — what an embedded
   * object wants, since stretching one axis distorts the content it is a
   * preview of (Word's own equation objects behave this way).
   */
  handles?: 'all' | 'corners';
  /** Draw the rotate knob and accept rotation gestures. Default true; false
   *  for boxes rotation would be wrong on (Word blocks it while an object
   *  flows inline with text). */
  rotatable?: boolean;
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
  pageToCanvas(p: {
    pageIndex: number;
    x: number;
    y: number;
  }): { x: number; y: number } | null;
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

  /** Float a host-built element beside a page rect, below it when there is
   *  room and above it otherwise — the equation palette. The element belongs
   *  to whoever passed it; the editor only positions and shows it. Pass null
   *  to take it away. */
  setPanel(el: OverlayPanelElement | null, at?: OverlayRect): void;
  /** Show an object-selection frame (border + resize handles + rotate knob),
   *  or null to clear. Page-local geometry; the editor renders it as DOM
   *  overlay in the canvas stack — no canvas repaint. */
  setFrame(frame: OverlayFrame | null): void;
  /** Another plugin's public handle — ONLY names this plugin declared in
   *  `uses` (anything else throws, by design). Resolved live against the
   *  current document's plugin set: instances are document-scoped, so a
   *  cached reference would go stale on the next load — don't keep one. */
  plugin<K extends keyof EditorPluginHandles & string>(
    name: K,
  ): EditorPluginHandles[K];
  /** Show a small action button straddling a page-local point (e.g. a selected
   *  cell block's top-right), invoking `onActivate` on click/tap; null clears it.
   *  A touch-friendly trigger where hover/right-click aren't available. */
  setActionButton(at: PagePoint | null, onActivate?: () => void): void;
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
  /** Names of the plugins this one calls through `ctx.plugin()`. Declared,
   *  not discovered: registration fails fast when a name is missing, setup
   *  order follows the dependency graph, and `ctx.plugin()` refuses anything
   *  undeclared — so the graph in the code is the graph in truth. Omit when
   *  the plugin stands alone (most do). */
  readonly uses?: readonly string[];
  /** Schema contributions merged into the editor's document schema, so a plugin
   *  can own its marks/nodes. Composed once when a document loads. */
  schema?: {
    marks?: Record<string, MarkSpec>;
    nodes?: Record<string, NodeSpec>;
  };
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
  /** Keyboard input, offered before the editor's keymaps (typing, undo,
   *  arrows) — e.g. Escape cancelling an in-flight drag gesture. Return true
   *  to claim (the editor preventDefaults and stops the event). */
  onKey?(ev: EditorKeyEvent): boolean;

  /** Text arriving WITHOUT a usable keydown: an IME committing a composed
   *  syllable, a plain-text paste, a `beforeinput` from an automation tool.
   *  Return true to claim it. A plugin that owns the caret must handle this
   *  as well as `onKey` — otherwise the hidden editor applies the insertion
   *  against its own selection, which for a selected atom means REPLACING the
   *  very object being edited. */
  onTextInput?(text: string): boolean;
  /** A click on one of the frame's action-strip buttons (the plugin put them
   *  there via setFrame's `actions`). Return true to claim. */
  onFrameAction?(id: string): boolean;
}
