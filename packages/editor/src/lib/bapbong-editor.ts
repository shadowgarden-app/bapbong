import {
  DOMParser as PMDOMParser,
  Node as ProseMirrorNode,
  Schema,
} from 'prosemirror-model';
import { imagePasteHandler, insertImageBlobs } from './paste-images';
import { schema as baseSchema } from '@shadow-garden/bapbong-model';
import { RenderCore } from '@shadow-garden/bapbong-view';
import {
  InputBridge,
  IS_MAC,
  moveCaretCommand,
  backspaceOutdent,
  shiftListLevel,
  paragraphEnter,
  type Command,
  // EditorView comes via input-bridge's re-export (not straight from
  // prosemirror-view) so `bridge.view` and this type share ONE identity.
  type EditorView,
  type EditorState,
  type Transaction,
} from '@shadow-garden/bapbong-input-bridge';
import type {
  CaretRect,
  // `Command` here is the headless registry command ({ name, run, isActive… });
  // aliased because input-bridge re-exports ProseMirror's own `Command` type.
  Command as EditorCommand,
  EditorChange,
  EditorPlugin,
  EditorPluginHandles,
  EditorPointerEvent,
  MeasureMetrics,
  MeasureText,
  EditorKeyEvent,
  OverlayFrame,
  OverlayGuide,
  OverlayRect,
  PagePoint,
  PluginContext,
  RangeDecoration,
  ResolvedLayout,
  SectionChromeOverrides,
  SectionConfig,
  SelectionRect,
} from '@shadow-garden/bapbong-contracts';

/** A section boundary surfaced for UI markers. `index` is the break index
 *  passed to `removeSectionBreak`; `rect` is a full-page-width line (page-local)
 *  at the boundary, ready for `pageToCanvas` — null if off the current layout. */
export interface SectionBoundary {
  index: number;
  newPage: boolean;
  columns: number;
  pos: number;
  rect: { pageIndex: number; x: number; y: number; width: number } | null;
}

// The plugin contract's canonical home is `contracts`; re-export it here so a
// plugin author can `import { EditorPlugin, PluginContext, EditorChange } from
// '@shadow-garden/bapbong-editor'`.
export type {
  EditorChange,
  EditorPlugin,
  PluginContext,
} from '@shadow-garden/bapbong-contracts';
// The render core is shared with the read-only viewer; re-export so a host can
// reach it (and `BapbongView`) without a separate import.
export { RenderCore } from '@shadow-garden/bapbong-view';
// Built-in ("internal") plugins ship with the editor (see built-in-plugins.ts)
// and are exposed as typed handles (e.g. editor.find) — no install needed.
import { createBuiltins } from './built-in-plugins';
import {
  Collection,
  KeybindingRegistry,
  perf,
} from '@shadow-garden/bapbong-contracts';
import {
  defaultCommands,
  insertEquation,
  toggleUnicodeHex,
} from '@shadow-garden/bapbong-commands';
export type {
  TableSelectionPlugin,
  CellBlock,
  SelectedCell,
} from './table-selection-plugin';
export type { FindPlugin, FindState } from './find-plugin';
export type { HyperlinkPlugin } from './hyperlink-plugin';
export type { ActiveField, TocPlugin } from './toc-plugin';

const CARET_BLINK_MS = 530;

export interface BapbongEditorOptions {
  /** The scroll viewport the page stack lives in (used for virtualization and
   *  scroll-into-view). Defaults to `stack.closest('.canvas-wrap')`. */
  viewport?: HTMLElement;
  /** Word's symbol AutoCorrect while typing — (c) → ©, --> → →, ... → … (see
   *  contracts AUTOCORRECT_RULES). Default on; a host preference can turn it
   *  off. */
  autoCorrect?: boolean;
  /** Editor plugins. Their lifecycle/event hooks are invoked by the core; they
   *  reach back through the PluginContext handed to `setup`.
   *
   *  Plugin instances are DOCUMENT-SCOPED: every `loadDocx` tears the current
   *  set down and builds a fresh one, so per-document state (a selected
   *  image's position, search matches) cannot survive into a document it
   *  doesn't belong to. Pass a FACTORY to get that isolation (a new closure
   *  per document); a bare instance is also accepted — it is re-setup() each
   *  document and must treat setup/teardown as its reset points. */
  plugins?: Array<EditorPlugin | (() => EditorPlugin)>;
  /** Engine-independent text measurer (e.g. font-file metrics). Defaults to a
   *  canvas-backed one; inject to make wrapping/pagination deterministic across
   *  WebView engines. Pair with {@link measureMetrics}. */
  measureText?: MeasureText;
  /** Vertical-metrics provider paired with {@link measureText}. */
  measureMetrics?: MeasureMetrics;
  /** Vertical gap between pages in layout px (painter default when omitted).
   *  Widen it when the host draws chrome in the gap (section-break markers);
   *  read back via {@link BapbongEditor.getPageGap}. */
  pageGap?: number;
  /** Fallback clipboard reader for hosts whose webview denies the async
   *  Clipboard API (e.g. the desktop shell reads the OS clipboard natively).
   *  Return every flavor present; null when empty / unavailable. */
  readClipboardFallback?: () => Promise<{
    html?: string;
    text?: string;
    image?: Uint8Array;
    imageMime?: string;
  } | null>;
  /** Print channel for hosts whose webview lacks `window.print()` (WKWebView
   *  is a silent no-op). Receives every page as a PNG data URL + its sheet
   *  size in PDF points; resolve true when the host handled the print. */
  printFallback?: (
    pages: { png: string; widthPt: number; heightPt: number }[],
  ) => Promise<boolean>;
}

/**
 * The framework-agnostic core of bapbong: it owns the render → edit loop
 * (import → layout → paint → input → selection → export) and the per-page
 * canvas stack, with no UI chrome of its own. A host wires comment UI / toolbars
 * around it via `onChange`, `caretRect`, `pageToCanvas`, `dispatch`, etc.
 *
 * The render half (layout/paint/scroll/zoom/geometry) lives in a shared
 * {@link RenderCore} (the same one the read-only `BapbongView` uses); this class
 * adds the editing half — a hidden ProseMirror editor as the IME / keyboard
 * sink, pointer-driven caret/selection, plugins, commands and clipboard — and
 * pushes a fresh doc + caret overlay to the core on every transaction.
 */
/**
 * The editor the user last interacted with (focused, or pressed a pointer on).
 *
 * Keydown is listened for on `window` in capture, because a plugin must still
 * see keys mid-gesture — when a claimed pointerdown has left DOM focus on
 * `<body>` rather than in the hidden contenteditable. With ONE editor on the
 * page "focus is on body" unambiguously meant "meant for me". With several
 * live editors it does not, and every one of them would claim the same key.
 *
 * So the body case is resolved by recency: exactly one editor can be the last
 * one touched, and only that one takes it. Editors whose own subtree holds the
 * target are unaffected — that case was never ambiguous.
 *
 * Deliberately not a "split" or "pane" concept: the library still knows nothing
 * about how a host arranges its editors, only that more than one can exist.
 */
let lastInteracted: BapbongEditor | null = null;

function markInteracted(editor: BapbongEditor): void {
  lastInteracted = editor;
}

/** True when a chrome story contains a PAGE/NUMPAGES field atom. */
function storyHasPageField(story: ProseMirrorNode | undefined): boolean {
  let found = false;
  story?.descendants((n) => {
    if (n.type.name === 'page_field') found = true;
    return !found;
  });
  return found;
}

/** A story's JSON minus every page_field atom (PAGE and NUMPAGES both go —
 *  hiding page numbers should not leave a dangling "of 12"). Literal text
 *  around the fields is the user's and stays. */
function stripPageFields(json: unknown): unknown {
  const node = json as { type?: string; content?: unknown[] };
  if (!node || typeof node !== 'object' || !Array.isArray(node.content))
    return json;
  return {
    ...node,
    content: node.content
      .filter((c) => (c as { type?: string } | null)?.type !== 'page_field')
      .map((c) => stripPageFields(c)),
  };
}

export class BapbongEditor {
  private readonly stack: HTMLElement;
  private readonly core: RenderCore;
  private bridge: InputBridge | null = null;

  // Caret blink state (solid on every interaction, toggling while idle).
  private lastCaret: CaretRect | null = null;
  private lastSelection: SelectionRect[] = [];
  private caretVisible = true;
  private blinkTimer: ReturnType<typeof setInterval> | null = null;
  // While dragging, the visual selection leads (painted straight from
  // hit-testing, same frame); the PM model commits once on pointerup.
  private dragAnchor: number | null = null;
  private dragHead: number | null = null;

  private readonly changeListeners = new Set<(c: EditorChange) => void>();
  private readonly caretPickListeners = new Set<(pos: number) => void>();
  // Timestamp of the last keydown on this editor — used to measure the full
  // keydown → painted latency (perf instrumentation; null once consumed).
  private inputStartedAt: number | null = null;

  /** Plugin registry keyed by name — built-ins (internal) + host (external). */
  private plugins: Collection<EditorPlugin>;
  private readonly pluginTeardowns: Array<() => void> = [];
  private readonly pluginCtx: PluginContext;
  /** The host's plugin entries as given — factories re-run per document. */
  private readonly hostPluginEntries: Array<
    EditorPlugin | (() => EditorPlugin)
  >;
  /** Per-plugin gated context (carries that plugin's `uses` allowlist). */
  private pluginCtxs = new Map<EditorPlugin, PluginContext>();
  // Whether any plugin wants pointer events (gates the per-move offer).
  private pointerPlugins = false;
  // Unsubscribe from the core's late-font relayout signal.
  private readonly offFonts: () => void;
  // Transient drag guide (e.g. column-resize preview); lazily created.
  private guideEl: HTMLDivElement | null = null;
  // Pool of translucent highlight divs (e.g. selected table-cell block).
  private readonly highlightEls: HTMLDivElement[] = [];
  // Floating action button (e.g. the cell-block properties trigger).
  private actionEl: HTMLButtonElement | null = null;
  private actionHandler: (() => void) | null = null;
  // Object-selection frame (image resize handles + rotate knob); lazily created.
  private frameEl: HTMLDivElement | null = null;

  /** Headless editor commands keyed by name — the surface a toolbar/menubar
   *  renders and dispatches against (`editor.commands.get('bold')?.run(...)`).
   *  Built from the shared command layer; the same ops run on a backend.
   *  (Plugin-contributed commands are an additive follow-up.) */
  readonly commands: Collection<EditorCommand> = defaultCommands();

  /**
   * Keyboard shortcuts as data — every chord the editor answers to WHILE IT
   * HAS FOCUS, by command name (see contracts `Keybinding`). The bridge's
   * keymap resolves keydowns against it live, so a binding added after mount
   * works at once. The core registers its own here (source `core`); plugins
   * and the host add theirs; a later `add` of the same chord replaces (a host
   * overriding a core binding). App-wide chords that must work with focus
   * elsewhere, or with no document open (⌘S, ⌘F, ⌘W), belong in the HOST's
   * own registry, dispatched by input-bridge's `installWindowKeymap`; the
   * Keyboard-shortcuts dialog lists both.
   */
  readonly keybindings = new KeybindingRegistry(IS_MAC);

  private readonly autoCorrect: boolean;
  private readonly readClipboardFallback?: BapbongEditorOptions['readClipboardFallback'];
  private readonly printFallback?: BapbongEditorOptions['printFallback'];

  constructor(stack: HTMLElement, opts: BapbongEditorOptions = {}) {
    this.stack = stack;
    this.autoCorrect = opts.autoCorrect ?? true;
    this.registerCoreKeys();
    this.readClipboardFallback = opts.readClipboardFallback;
    this.printFallback = opts.printFallback;
    this.core = new RenderCore(stack, {
      viewport: opts.viewport,
      measureText: opts.measureText,
      measureMetrics: opts.measureMetrics,
      pageGap: opts.pageGap,
    });
    // The core resolves plugin decorations to page rects at paint time.
    this.core.setDecorationProvider(() => this.pluginDecorations());
    // After a late-font relayout the core repaints content; recompute the caret
    // overlay (rects moved) and re-anchor the IME against the new layout.
    this.offFonts = this.core.onFontsReloaded(() => {
      if (this.bridge) this.render(this.bridge.state, false);
    });

    stack.addEventListener('pointerdown', this.onPointerDown);
    stack.addEventListener('pointermove', this.onPointerMove);
    stack.addEventListener('pointerup', this.onPointerUp);
    stack.addEventListener('pointercancel', this.onPointerUp);
    stack.addEventListener('dblclick', this.onDblClick);
    stack.addEventListener('contextmenu', this.onContextMenu);
    // Window-level capture: keys must reach plugins even mid-gesture, when a
    // claimed pointerdown left focus outside the hidden editor (see onKeyDown).
    window.addEventListener('keydown', this.onKeyDown, true);

    // Plugins are document-scoped: this first build serves the empty editor,
    // and every loadDocx rebuilds (see buildPlugins).
    this.hostPluginEntries = opts.plugins ?? [];
    this.pluginCtx = this.makePluginContext();
    this.plugins = new Collection<EditorPlugin>([], { idProperty: 'name' });
    this.buildPlugins();
  }

  /** Tear down the current plugin set and build a fresh one.
   *
   *  Runs once at construction and again on every document load. Built-ins
   *  come from their factories, so each document gets NEW closures — the only
   *  structure in which per-document plugin state (a selected image's
   *  position, search matches) cannot leak into a document it doesn't belong
   *  to. No lifecycle flag to remember, nothing to opt into: the state's
   *  container simply stops existing. Host entries passed as factories get
   *  the same isolation; bare instances are re-setup() and documented to
   *  treat setup/teardown as their reset points.
   *
   *  `uses` declarations are enforced here: unknown names fail registration
   *  outright, and setup runs dependencies-first (cycles are an error). */
  private buildPlugins(): void {
    for (const t of this.pluginTeardowns) t();
    this.pluginTeardowns.length = 0;
    this.pluginCtxs = new Map();

    const instances = [
      ...createBuiltins(),
      ...this.hostPluginEntries.map((e) => (typeof e === 'function' ? e() : e)),
    ];
    this.plugins = new Collection<EditorPlugin>(instances, {
      idProperty: 'name',
    });
    this.pointerPlugins = instances.some((p) => p.onPointer);

    for (const p of orderPluginsByUses(instances)) {
      const ctx = this.ctxFor(p);
      this.pluginCtxs.set(p, ctx);
      const teardown = p.setup?.(ctx);
      if (teardown) this.pluginTeardowns.push(teardown);
    }
    // Replay host bindings onto the new instances — without this, a listener
    // registered once at startup would be lost at the first document load.
    for (const bind of this.pluginBindings) bind(this);
  }

  /** The shared context plus this plugin's gated `plugin()` accessor.
   *  Prototype inheritance, NOT a spread: the shared context exposes `state`
   *  and `layout` as live getters, and spreading would invoke them here —
   *  throwing before a document loads, and freezing a snapshot after. */
  private ctxFor(p: EditorPlugin): PluginContext {
    const gated = pluginLookupFor(p, (name) => this.plugins.get(name) ?? null);
    return Object.create(this.pluginCtx, {
      plugin: { value: gated, enumerable: true },
    }) as PluginContext;
  }

  /** The controlled surface handed to each plugin (live state + geometry). */
  private makePluginContext(): PluginContext {
    // Arrow methods capture `this` lexically (no this-alias).
    const ctx = {
      dispatch: (tr: Transaction) => this.dispatch(tr),
      caretRect: (pos: number) => this.core.caretRect(pos),
      pageToCanvas: (p: PagePoint) => this.core.pageToCanvas(p),
      setSelection: (from: number, to?: number) => this.setSelection(from, to),
      scrollToPos: (pos: number, topMargin?: number) =>
        this.core.scrollToPos(pos, topMargin),
      requestPaint: () => this.core.paintContent(this.currentOverlay()),
      setCursor: (cursor: string | null) => {
        this.stack.style.cursor = cursor ?? '';
      },
      setGuide: (guide: OverlayGuide | null) => this.setGuide(guide),
      setHighlight: (rects: OverlayRect[] | null) => this.setHighlight(rects),
      setActionButton: (at: PagePoint | null, onActivate?: () => void) =>
        this.setActionButton(at, onActivate),
      setFrame: (frame: OverlayFrame | null) => this.setFrame(frame),
    };
    // `state` + `layout` are live (read on each access); arrow getters keep them
    // current without throwing at construction (the doc loads later).
    Object.defineProperty(ctx, 'state', {
      enumerable: true,
      get: () => this.state,
    });
    Object.defineProperty(ctx, 'layout', {
      enumerable: true,
      get: () => this.core.layout,
    });
    return ctx as PluginContext;
  }

  // ── Public API ──────────────────────────────────────────────────────

  /** Current editor state (doc + selection). Throws before a document loads. */
  get state(): EditorState {
    if (!this.bridge) throw new Error('BapbongEditor: no document loaded');
    return this.bridge.state;
  }

  /** Number of laid-out pages (0 before the first document). */
  get pageCount(): number {
    return this.core.pageCount;
  }

  /** 1-based index of the page the reader is currently looking at (0 before the
   *  first layout). Cheap — safe to poll on scroll for a "Page X of N" readout. */
  currentPage(): number {
    return this.core.currentPage();
  }

  /** The document schema in use (model's base + plugin schema contributions).
   *  Hosts serialize/parse comment bodies, previews, etc. against this. */
  get schema(): Schema {
    return this.core.schema;
  }

  /** Host subscriptions that must survive document loads, replayed against
   *  each fresh plugin set (see `onPlugin`). */
  private readonly pluginBindings: Array<(e: BapbongEditor) => void> = [];

  /** Bind to a plugin handle in a way that survives document loads.
   *
   *  Plugin instances are document-scoped, so a listener registered once at
   *  startup (`editor.plugin('table-selection').onAction(…)`) would die at the
   *  first `loadDocx` — silently, which is the worst way for a feature to
   *  stop working. This runs `bind` now AND after every rebuild:
   *
   *    editor.onPlugin('table-selection', (p) => p.onAction(openCellProps));
   *
   *  Prefer it over `plugin()` for anything long-lived; `plugin()` stays right
   *  for one-shot calls (`plugin('find').setQuery(q)`). */
  onPlugin<K extends keyof EditorPluginHandles & string>(
    name: K,
    bind: (handle: EditorPluginHandles[K]) => void,
  ): void {
    const run = (e: BapbongEditor) => bind(e.plugin(name));
    this.pluginBindings.push(run);
    if (this.plugins.get(name)) run(this);
  }

  /** A plugin's public handle by name — `editor.plugin('find').setQuery(…)`.
   *
   *  One generic accessor instead of a getter per plugin: the core stays
   *  ignorant of which plugins exist, and a new one becomes reachable just by
   *  augmenting `EditorPluginHandles` from its own module. Types come from
   *  that augmentation, so the name is checked and autocompleted.
   *
   *  Resolves LIVE against the current document's plugin set. Don't cache the
   *  result across a `loadDocx` — instances are document-scoped, and a held
   *  reference would drive a torn-down plugin. */
  plugin<K extends keyof EditorPluginHandles & string>(
    name: K,
  ): EditorPluginHandles[K] {
    const p = this.plugins.get(name);
    if (!p) throw new Error(`BapbongEditor: no plugin named "${name}"`);
    return p as EditorPluginHandles[K];
  }

  /** Import a .docx, lay it out, and paint the first frame. Resolves with the
   *  imported page-chrome keys (the rest of the import rides on the doc model
   *  exposed via `state`). */
  async loadDocx(
    bytes: ArrayBuffer,
    opts?: {
      /** Re-open the document at an in-progress edit state instead of the
       *  pristine disk contents. `bytes` still supplies the page chrome + carry
       *  package (edit-invariant), but the body laid out and mounted is
       *  `restoreState.doc`, and the bridge adopts `restoreState` so its undo
       *  history and selection survive. The state must have been produced by
       *  this editor (same schema) — e.g. captured earlier from `state`. */
      restoreState?: EditorState;
      /** Password for an encrypted document (forwarded to the importer). */
      password?: string;
    },
  ): Promise<{ headerKeys: string[]; footerKeys: string[] }> {
    // Fresh plugin instances for this document — per-document state (a
    // selected image, search matches) belongs to the document it came from,
    // and rebuilding is what makes that true structurally. Must precede
    // composeSchema: the new instances contribute the schema.
    this.buildPlugins();
    // Compose the doc schema from model's base + any plugin schema
    // contributions, and import against it (so plugin-owned marks/nodes parse).
    const composed = composeSchema(baseSchema, this.plugins);
    const restore = opts?.restoreState;
    const { doc, headerKeys, footerKeys } = await this.core.loadDocx(
      bytes,
      // Skip the core's initial paint — `mount` paints with a caret overlay.
      // When restoring, lay out the edited body, not the pristine import.
      {
        schema: composed ?? undefined,
        paint: false,
        layoutTarget: restore?.doc,
        ...(opts?.password !== undefined ? { password: opts.password } : {}),
      },
    );
    this.mount(restore?.doc ?? doc, restore);
    return { headerKeys, footerKeys };
  }

  /** Export the (edited) document back to .docx bytes, carrying the imported
   *  source package so unmodelled parts survive the round-trip. */
  exportDocx(): Promise<Uint8Array> {
    return this.core.exportDocx();
  }

  /** Apply a transaction (the host builds comment/edit transactions against
   *  `state` and dispatches them here). */
  dispatch(tr: Transaction): void {
    this.bridge?.dispatch(tr);
  }

  /** Viewport-space rect of the caret at `pos` (default: the selection head)
   *  — lets hosts anchor floating UI (e.g. a link panel) at the text cursor
   *  or at a stable position like a link's start. Null before the first
   *  layout or when the position is off the built pages. */
  caretViewportRect(
    pos?: number,
  ): { x: number; y: number; height: number } | null {
    const state = this.bridge?.state;
    if (!state) return null;
    const cr = this.core.caretRect(pos ?? state.selection.head);
    if (!cr) return null;
    const top = this.core.pageToCanvas({
      pageIndex: cr.pageIndex,
      x: cr.x,
      y: cr.y,
    });
    const bottom = this.core.pageToCanvas({
      pageIndex: cr.pageIndex,
      x: cr.x,
      y: cr.y + cr.height,
    });
    if (!top || !bottom) return null;
    const host = this.stack.getBoundingClientRect();
    return {
      x: host.left + top.x,
      y: host.top + top.y,
      height: bottom.y - top.y,
    };
  }

  /** Insert an image blob at the selection, measured to its intrinsic size
   *  (display-capped) exactly like a paste. Prefer this over a bare
   *  insertImage command from host insert flows — a node without
   *  width/height lays out as an invisible 0×0 box. Resolves false when no
   *  document is mounted or the blob can't be decoded. */
  insertImageBlob(blob: Blob): Promise<boolean> {
    const view = this.bridge?.view;
    return view ? insertImageBlobs(view, [blob]) : Promise.resolve(false);
  }

  /** Subscribe to layout/paint cycles. Returns an unsubscribe fn. */
  onChange(cb: (c: EditorChange) => void): () => void {
    this.changeListeners.add(cb);
    return () => this.changeListeners.delete(cb);
  }

  /** Subscribe to caret placement (a pointerdown that placed the caret). The
   *  host uses this to e.g. select the comment whose range was clicked. */
  onCaretPick(cb: (pos: number) => void): () => void {
    this.caretPickListeners.add(cb);
    return () => this.caretPickListeners.delete(cb);
  }

  /** Caret geometry at a doc position (page-local), or null. */
  caretRect(pos: number): CaretRect | null {
    return this.core.caretRect(pos);
  }

  /** The current paint-ready layout (page geometry, display page labels), or
   *  null before a document loads. Read-only — hosts use it to label chrome
   *  around the canvas (e.g. the section-break marker's "ii → 1"). */
  get layout(): ResolvedLayout | null {
    return this.core.layout;
  }

  /** Map a page-local point to container (canvas-stack) coordinates, or null. */
  pageToCanvas(p: PagePoint): { x: number; y: number } | null {
    return this.core.pageToCanvas(p);
  }

  /** The document's section boundaries (one per section break), for a host to
   *  draw clickable markers. `pos` is the doc position at the start of the
   *  section after the break — feed it to `caretRect` to place the marker. */
  sectionBoundaries(): SectionBoundary[] {
    if (!this.bridge) return [];
    const doc = this.bridge.state.doc;
    const sections = doc.attrs['sections'] as SectionConfig[] | null;
    if (!sections || sections.length < 2) return [];
    const offsets: number[] = [];
    doc.forEach((_node, offset) => offsets.push(offset));
    const out: SectionBoundary[] = [];
    let blockIdx = 0;
    for (let b = 0; b < sections.length - 1; b++) {
      blockIdx += sections[b].blockCount; // first block of section b+1
      // A NEXT-PAGE break's marker lives in the page gap, so it anchors on
      // the new page's first block. A CONTINUOUS break happens mid-page, and
      // its true position is the BREAK MARK itself — the last paragraph of
      // the ending section, whose (zero-height) line sits exactly on the
      // boundary. Anchoring on the next section's first block put the line
      // at that block's first line, which is BELOW the boundary by its
      // space-before — and above a column that starts flush, the marker cut
      // straight through the neighbouring column's heading.
      const anchorIdx = sections[b + 1].newPage ? blockIdx : blockIdx - 1;
      const pos = (offsets[anchorIdx] ?? 0) + 1;
      const cr = this.core.caretRect(pos);
      const page = cr && this.core.layout?.pages[cr.pageIndex];
      out.push({
        index: b,
        newPage: sections[b + 1].newPage,
        columns: sections[b + 1].columns.count,
        pos,
        rect:
          cr && page
            ? { pageIndex: cr.pageIndex, x: 0, y: cr.y, width: page.width }
            : null,
      });
    }
    return out;
  }

  /** True when section `sectionIndex`'s pages display a page number — some
   *  effective header/footer story (overrides applied, Link-to-Previous
   *  resolved) contains a PAGE field. */
  sectionShowsPageNumbers(sectionIndex: number): boolean {
    const { headers, footers } = this.sectionStories(sectionIndex);
    return [...Object.values(headers), ...Object.values(footers)].some(
      storyHasPageField,
    );
  }

  /**
   * Show/hide page numbers on ONE section, with Word's semantics: the toggle
   * edits the section's header/footer stories themselves, not a display flag.
   * Hiding strips the PAGE fields out of a copy of the effective stories and
   * installs it as the section's own chrome (severing "Link to Previous" for
   * exactly those stories); showing removes that override — or, when the
   * inherited chrome never had a number, inserts a centered PAGE field into
   * the section's footer (Word's Insert ▸ Page Number). Rides
   * `doc.attrs.sectionChromeOverrides`, so it undoes cleanly and exports as
   * real header/footer parts.
   */
  setSectionPageNumbersShown(sectionIndex: number, show: boolean): void {
    if (!this.bridge) return;
    const state = this.bridge.state;
    const schema = state.schema;
    const attr =
      (state.doc.attrs[
        'sectionChromeOverrides'
      ] as SectionChromeOverrides | null) ?? {};
    const key = String(sectionIndex);
    let next: SectionChromeOverrides | null = null;

    if (!show) {
      const { headers, footers } = this.sectionStories(sectionIndex);
      const override: SectionChromeOverrides[string] = {};
      for (const [variant, story] of Object.entries(headers)) {
        if (storyHasPageField(story))
          (override.headers ??= {})[variant] = stripPageFields(story.toJSON());
      }
      for (const [variant, story] of Object.entries(footers)) {
        if (storyHasPageField(story))
          (override.footers ??= {})[variant] = stripPageFields(story.toJSON());
      }
      if (!override.headers && !override.footers) return; // already hidden
      next = { ...attr, [key]: override };
    } else {
      // Inherited = the chrome as imported, IGNORING overrides.
      const inherited =
        this.core.effectiveSectionChrome(null)?.[sectionIndex] ??
        this.core.flatChrome();
      const inheritedShows = [
        ...Object.values(inherited.headers),
        ...Object.values(inherited.footers),
      ].some(storyHasPageField);
      if (inheritedShows) {
        if (!(key in attr)) return; // already shown
        next = { ...attr };
        delete next[key];
        if (Object.keys(next).length === 0) next = null;
      } else {
        const numberPara = schema.node('paragraph', { align: 'center' }, [
          schema.node('page_field', { kind: 'page' }),
        ]);
        const base = inherited.footers['default'];
        const footer = base
          ? base.copy(base.content.addToEnd(numberPara))
          : schema.node('doc', null, [numberPara]);
        next = {
          ...attr,
          [key]: {
            ...attr[key],
            footers: { ...attr[key]?.footers, default: footer.toJSON() },
          },
        };
      }
    }
    this.dispatch(state.tr.setDocAttribute('sectionChromeOverrides', next));
  }

  /** Section `sectionIndex`'s effective stories (overrides > per-section
   *  chrome > the flat document chrome). */
  private sectionStories(sectionIndex: number): {
    headers: Record<string, ProseMirrorNode>;
    footers: Record<string, ProseMirrorNode>;
  } {
    const merged = this.bridge
      ? this.core.effectiveSectionChrome(this.bridge.state.doc)
      : null;
    const s = merged?.[sectionIndex];
    return s
      ? { headers: s.headers, footers: s.footers }
      : this.core.flatChrome();
  }

  /** Move the selection (caret if `to` omitted) and anchor the IME. */
  setSelection(from: number, to?: number): void {
    this.bridge?.setSelection(from, to);
  }

  /** Select the word at a doc position. */
  selectWordAt(pos: number): void {
    this.bridge?.selectWordAt(pos);
  }

  /** Scroll the viewport so the caret at `pos` sits `topMargin` px from the top. */
  scrollToPos(pos: number, topMargin = 80): void {
    this.core.scrollToPos(pos, topMargin);
  }

  /** Set the zoom factor (1 = 100%) and repaint at the new scale. */
  setZoom(zoom: number): void {
    this.core.setZoom(zoom);
  }

  /** The current zoom factor (1 = 100%). */
  getZoom(): number {
    return this.core.getZoom();
  }

  /** The inter-page gap in layout px (multiply by zoom for canvas px). */
  getPageGap(): number {
    return this.core.getPageGap();
  }

  /** Full-resolution PNG snapshot of one page (0-based) — a host's
   *  feedback/report UI uses this for the "what Bapbong renders" side of a
   *  comparison pair. Renders just that page with a throwaway painter;
   *  `width`/`height` are CSS px at zoom 1. Null when the page doesn't exist
   *  or no layout is resolved yet. */
  pageSnapshot(
    index: number,
  ): Promise<{ png: string; width: number; height: number } | null> {
    return this.core.pageSnapshot(index);
  }

  /** Print the whole document — renders every page (not just the visible
   *  ones) and prints one image per sheet. Prefers the host's print channel
   *  (`printFallback`) when provided; falls back to the iframe +
   *  `window.print()` path browsers support. */
  async print(): Promise<void> {
    if (this.printFallback) {
      // CSS px (96/in) → PDF points (72/in): the sheet keeps its physical size.
      const toPt = (px: number) => (px * 72) / 96;
      const pages = (await this.core.pageSnapshots()).map((p) => ({
        png: p.png,
        widthPt: toPt(p.width),
        heightPt: toPt(p.height),
      }));
      if (pages.length) {
        try {
          if (await this.printFallback(pages)) return;
        } catch (err) {
          console.warn(
            '[bapbong] host print channel failed, falling back:',
            err,
          );
        }
      }
    }
    return this.core.print();
  }

  /** Focus the hidden ProseMirror editor (keyboard/IME sink). */
  focus(): void {
    markInteracted(this);
    this.bridge?.focus();
  }

  // ── Clipboard ───────────────────────────────────────────────────────
  // These need the hidden ProseMirror view (selection serialization +
  // clipboard), so they live here, not in the headless command layer. They are
  // browser-gated: copy/cut run via execCommand on the focused view (a user
  // gesture); paste reads the async Clipboard API (permission-gated).

  /** Copy the selection to the clipboard (PM serializes the slice). */
  copy(): boolean {
    if (!this.bridge) return false;
    this.bridge.focus();
    return document.execCommand('copy');
  }

  /** Cut the selection to the clipboard. */
  cut(): boolean {
    if (!this.bridge) return false;
    this.bridge.focus();
    return document.execCommand('cut');
  }

  /** Paste clipboard content — HTML parsed by the schema, image blobs as
   *  embedded images, else plain text.
   *
   *  Menu-driven paste can't ride a native ClipboardEvent, so it climbs a
   *  ladder of acquisition strategies until one yields content:
   *  1. the host's native clipboard reader (`readClipboardFallback`) when
   *     provided — on WKWebView every programmatic clipboard READ (both
   *     `navigator.clipboard.*` and execCommand) pops a "Paste" permission
   *     callout and pends until the user taps it, so the desktop shell
   *     must never touch the DOM clipboard from a menu click;
   *  2. async Clipboard API (`navigator.clipboard.read`) — browsers, where
   *     a user-gesture read is granted silently;
   *  3. `document.execCommand('paste')` — legacy last resort. */
  async paste(): Promise<void> {
    const view = this.bridge?.view;
    if (!view) return;
    // Focus up front: menu clicks steal focus from the hidden view — typing
    // must work right after a paste — and a focused view keeps the call
    // closer to the user gesture for Clipboard API permission checks.
    this.bridge?.focus();
    if (await this.pasteViaHostClipboard(view)) return;
    if (this.readClipboardFallback) return; // host said empty — don't summon webview paste UI
    if (await this.pasteViaClipboardApi(view)) return;
    document.execCommand('paste');
  }

  private async pasteViaClipboardApi(view: EditorView): Promise<boolean> {
    try {
      const items = await navigator.clipboard.read();
      // Mirror imagePasteHandler's preference: HTML with real text wins;
      // image-wrapper-only HTML defers to the blob (remote <img> srcs are
      // rejected by the schema, the blob keeps the pixels).
      const htmlItem = items.find((i) => i.types.includes('text/html'));
      if (htmlItem) {
        const html = await (await htmlItem.getType('text/html')).text();
        // Inert document: no script execution and (unlike innerHTML on a
        // live div) no eager <img src> network fetches while parsing.
        const dom = new DOMParser().parseFromString(html, 'text/html').body;
        if ((dom.textContent ?? '').trim()) {
          const slice = PMDOMParser.fromSchema(view.state.schema).parseSlice(
            dom,
          );
          view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView());
          return true;
        }
      }
      for (const item of items) {
        const type = item.types.find((t) => t.startsWith('image/'));
        if (type && (await insertImageBlobs(view, [await item.getType(type)])))
          return true;
      }
      const text = await navigator.clipboard.readText();
      if (text) {
        view.dispatch(view.state.tr.insertText(text).scrollIntoView());
        return true;
      }
      return false;
    } catch (err) {
      console.warn(
        '[bapbong] Clipboard API paste unavailable, falling back:',
        err,
      );
      return false;
    }
  }

  private async pasteViaHostClipboard(view: EditorView): Promise<boolean> {
    if (!this.readClipboardFallback) return false;
    try {
      const data = await this.readClipboardFallback();
      if (!data) return false;
      // Same preference order as the Clipboard API path.
      if (data.html) {
        const dom = new DOMParser().parseFromString(
          data.html,
          'text/html',
        ).body;
        if ((dom.textContent ?? '').trim()) {
          const slice = PMDOMParser.fromSchema(view.state.schema).parseSlice(
            dom,
          );
          view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView());
          return true;
        }
      }
      if (data.image?.length) {
        const blob = new Blob([data.image as BlobPart], {
          type: data.imageMime ?? 'image/png',
        });
        if (await insertImageBlobs(view, [blob])) return true;
      }
      if (data.text) {
        view.dispatch(view.state.tr.insertText(data.text).scrollIntoView());
        return true;
      }
      return false;
    } catch (err) {
      console.warn('[bapbong] host clipboard fallback failed:', err);
      return false;
    }
  }

  /** Paste clipboard text as plain text (no formatting). */
  async pasteText(): Promise<void> {
    const view = this.bridge?.view;
    if (!view) return;
    this.bridge?.focus(); // see paste()
    // Host reader first, same reason as paste(): a webview readText() would
    // pop the WKWebView paste-permission callout.
    if (this.readClipboardFallback) {
      try {
        const data = await this.readClipboardFallback();
        if (data?.text)
          view.dispatch(view.state.tr.insertText(data.text).scrollIntoView());
      } catch (err) {
        console.warn('[bapbong] host clipboard fallback failed:', err);
      }
      return;
    }
    try {
      const text = await navigator.clipboard.readText();
      if (text) view.dispatch(view.state.tr.insertText(text).scrollIntoView());
    } catch (err) {
      console.warn('[bapbong] Clipboard API readText unavailable:', err);
    }
  }

  destroy(): void {
    // Never leave a destroyed editor as the body-key owner: it would silence
    // the surviving ones (they'd all defer to something that no longer exists).
    if (lastInteracted === this) lastInteracted = null;
    for (const t of this.pluginTeardowns) t();
    this.pluginTeardowns.length = 0;
    this.stopBlink();
    this.offFonts();
    window.removeEventListener('keydown', this.onKeyDown, true);
    this.stack.removeEventListener('pointerdown', this.onPointerDown);
    this.stack.removeEventListener('pointermove', this.onPointerMove);
    this.stack.removeEventListener('pointerup', this.onPointerUp);
    this.stack.removeEventListener('pointercancel', this.onPointerUp);
    this.stack.removeEventListener('dblclick', this.onDblClick);
    this.stack.removeEventListener('contextmenu', this.onContextMenu);
    this.guideEl?.remove();
    this.frameEl?.remove();
    this.frameEl = null;
    this.guideEl = null;
    for (const el of this.highlightEls) el.remove();
    this.highlightEls.length = 0;
    this.actionEl?.remove();
    this.actionEl = null;
    this.core.destroy();
    this.bridge?.destroy();
    this.bridge = null;
  }

  // ── Setup / render loop ─────────────────────────────────────────────

  /** Mount the hidden ProseMirror editor on a doc and run the first paint. When
   *  `state` is given the bridge adopts it verbatim (carrying an in-progress
   *  session's undo history + selection across a rebind); otherwise a fresh
   *  editing state is created from `doc`. */
  private mount(doc: ProseMirrorNode, state?: EditorState): void {
    this.bridge?.destroy();
    this.bridge = new InputBridge({
      doc,
      state,
      // Every chord — Enter/Backspace/Tab/arrows included — comes from the
      // keybinding registry, resolved per keydown (see registerCoreKeys).
      resolveKey: this.resolveEditorKey,
      onUpdate: (state, tr) => this.refresh(state, tr),
      handlePaste: imagePasteHandler,
      autoCorrect: this.autoCorrect,
    });
    // Hidden editor lives in the page-canvas container, so IME anchoring
    // scrolls along (positioned at the painted caret, in container coords).
    this.stack.appendChild(this.bridge.dom);
    // The core already laid out this doc in `loadDocx`; paint the first frame
    // with the caret overlay (no redundant relayout).
    this.render(this.bridge.state, true);
  }

  /** Bridge update hook: relayout (only when the doc changed) → paint → place
   *  IME → notify. */
  private refresh(state: EditorState, tr?: Transaction): void {
    const docChanged = tr?.docChanged ?? true;
    // Time spent between the keydown and this edit's layout starting — i.e. the
    // browser's contenteditable update + ProseMirror's input handling/dispatch.
    if (docChanged && this.inputStartedAt != null) {
      perf.log(
        'key→refresh (PM/input latency)',
        perf.now() - this.inputStartedAt,
      );
    }
    perf.span(docChanged ? 'refresh(edit)' : 'refresh(sel)', () => {
      if (docChanged || !this.core.layout) this.core.layoutDoc(state.doc);
      this.render(state, docChanged);
    });
    // Full keydown → canvas-drawn latency (input handling + layout + paint).
    if (docChanged && this.inputStartedAt != null) {
      perf.log('key→paint (TOTAL)', perf.now() - this.inputStartedAt);
      this.inputStartedAt = null;
    }
  }

  /** Compute the caret/selection overlay from `state`, paint it (full content
   *  when `contentDirty`, else overlay-only), anchor the IME, and notify. */
  private render(state: EditorState, contentDirty: boolean): void {
    const sel = state.selection;
    this.lastCaret = this.core.caretRect(sel.head);
    this.lastSelection = sel.empty
      ? []
      : this.core.selectionRects(sel.from, sel.to);

    // While dragging the caret stays solid and the timer rests; otherwise every
    // interaction restarts the blink phase.
    if (this.dragAnchor != null) {
      this.stopBlink();
      this.caretVisible = true;
    } else {
      this.restartBlink();
    }

    const overlay = this.currentOverlay();
    if (contentDirty) this.core.paintContent(overlay);
    else this.core.paintOverlay(overlay);

    // Anchor the hidden editor (and its IME popup) at the painted caret
    // (pageToCanvas returns container-relative coords, where the editor lives).
    const caret = this.lastCaret;
    if (caret && this.bridge) {
      const pt = this.core.pageToCanvas({
        pageIndex: caret.pageIndex,
        x: caret.x,
        y: caret.y,
      });
      if (pt) this.bridge.place(pt.x, pt.y, caret.height);
    }

    this.emitChange({
      state,
      pageCount: this.core.pageCount,
      docChanged: contentDirty,
    });
  }

  /** The caret/selection overlay in the current blink phase. */
  private currentOverlay(): {
    caret: CaretRect | null;
    selection: SelectionRect[];
  } {
    return {
      caret: this.caretVisible ? this.lastCaret : null,
      selection: this.lastSelection,
    };
  }

  /** Each plugin's doc-range decorations (the core resolves them to rects). */
  private pluginDecorations(): RangeDecoration[] {
    const out: RangeDecoration[] = [];
    for (const p of this.plugins) {
      const decos = p.decorations?.(this.pluginCtx);
      if (decos) out.push(...decos);
    }
    return out;
  }

  private emitChange(c: EditorChange): void {
    for (const cb of this.changeListeners) cb(c);
    for (const p of this.plugins) p.onChange?.(c);
  }

  /** Redraw caret/selection in the current blink phase (overlay-only). */
  private repaintOverlay(): void {
    this.core.paintOverlay(this.currentOverlay());
  }

  // ── Caret blink ─────────────────────────────────────────────────────

  private stopBlink(): void {
    if (this.blinkTimer != null) clearInterval(this.blinkTimer);
    this.blinkTimer = null;
  }

  /** Show the caret solid now, then blink while idle. */
  private restartBlink(): void {
    this.stopBlink();
    this.caretVisible = true;
    this.blinkTimer = setInterval(() => {
      this.caretVisible = !this.caretVisible;
      this.repaintOverlay();
    }, CARET_BLINK_MS);
  }

  // ── Keyboard (vertical caret) ───────────────────────────────────────

  /** ArrowUp/ArrowDown against the canvas layout (the hidden DOM's own line
   *  wrapping is meaningless). With `extend`, Shift+arrow grows the selection. */
  /**
   * The core's own shortcuts: the editing keys that used to be a static map
   * handed to the bridge, plus the Word/OS staples. Registered as COMMANDS
   * (so they list with a title and can be rebound) and as bindings.
   */
  private registerCoreKeys(): void {
    const core = (
      name: string,
      title: string,
      run: Command,
      isEnabled?: (state: EditorState) => boolean,
    ) =>
      this.commands.add({
        name,
        title,
        run: (state, dispatch) => run(state, dispatch),
        ...(isEnabled && { isEnabled }),
      });
    core('paragraph-enter', 'New paragraph / continue list', paragraphEnter);
    core(
      'backspace-outdent',
      'Outdent (Backspace at line start)',
      backspaceOutdent,
    );
    core('list-indent', 'Demote list item', shiftListLevel(1));
    core('list-outdent', 'Promote list item', shiftListLevel(-1));
    core('caret-up', 'Move caret up a line', this.verticalCmd(-1));
    core('caret-down', 'Move caret down a line', this.verticalCmd(1));
    core('select-up', 'Extend selection up a line', this.verticalCmd(-1, true));
    core(
      'select-down',
      'Extend selection down a line',
      this.verticalCmd(1, true),
    );
    const bind = (key: string, command: string, when?: string): void => {
      this.keybindings.add({
        key,
        command,
        source: 'core',
        ...(when && { when }),
      });
    };
    // Order matters only for the dialog's insertion order (it sorts anyway).
    bind('Enter', 'paragraph-enter', 'editing text');
    bind(
      'Backspace',
      'backspace-outdent',
      'at the start of a list or indented paragraph',
    );
    bind('Tab', 'list-indent', 'in a list');
    bind('Shift-Tab', 'list-outdent', 'in a list');
    bind('ArrowUp', 'caret-up', 'editing text');
    bind('ArrowDown', 'caret-down', 'editing text');
    bind('Shift-ArrowUp', 'select-up', 'editing text');
    bind('Shift-ArrowDown', 'select-down', 'editing text');
    bind('Mod-z', 'undo');
    bind('Shift-Mod-z', 'redo');
    bind('Mod-y', 'redo');
    // Word's staples. Nothing bound these before; the base keymap has none.
    bind('Mod-b', 'bold', 'editing text');
    bind('Mod-i', 'italic', 'editing text');
    bind('Mod-u', 'underline', 'editing text');
    // Word's Alt+X: hex before the caret ↔ the character.
    this.commands.add(toggleUnicodeHex());
    bind(
      'Alt-x',
      'toggle-unicode-hex',
      'hex digits or a character before the caret',
    );
    // Word's Alt+=: insert (or convert the selection into) an equation.
    this.commands.add(insertEquation());
    bind('Alt-=', 'insert-equation', 'editing text');
  }

  /** The bridge's live key lookup: a registered chord → the command's run,
   *  as a ProseMirror command. */
  private readonly resolveEditorKey = (name: string): Command | undefined => {
    const b = this.keybindings.get(name);
    if (!b) return undefined;
    const cmd = this.commands.get(b.command);
    if (!cmd) return undefined;
    return (state, dispatch) => cmd.run(state, dispatch);
  };

  private verticalCmd(dir: -1 | 1, extend = false): Command {
    return moveCaretCommand((state) => {
      const head = state.selection.head;
      const cr = this.core.caretRect(head);
      if (!cr) return null;
      return this.core.verticalCaret(head, dir, cr.x);
    }, extend);
  }

  // ── Pointer ─────────────────────────────────────────────────────────

  private onPointerDown = (ev: PointerEvent): void => {
    // Pressing in an editor makes it the one a body-targeted key belongs to —
    // recorded before any claim, since a claimed gesture is exactly the case
    // that strands focus on <body>.
    markInteracted(this);
    // A pointer plugin (e.g. table-column resize) may claim the press; if so,
    // preventDefault + capture the pointer for it and skip caret placement.
    if (this.offerPointer('down', ev)) {
      ev.preventDefault();
      try {
        (ev.target as HTMLElement).setPointerCapture?.(ev.pointerId);
      } catch {
        // capture is a nicety
      }
      return;
    }
    // Only the primary (left) button places the caret / starts a selection drag.
    // A right-click fires pointerdown (button 2) before `contextmenu`; collapsing
    // the selection here would lose it before the context menu opens.
    if (ev.button !== 0) return;
    const pos = this.core.posAtEvent(ev);
    if (pos == null || !this.bridge) return;
    ev.preventDefault(); // keep focus on the hidden editor
    this.dragAnchor = pos;
    this.dragHead = pos;
    // Keep receiving moves when the pointer leaves the canvas mid-drag. Guarded:
    // setPointerCapture throws if no active pointer matches the id (rare in real
    // use; also synthetic events) — it must not abort the caret placement below.
    try {
      (ev.target as HTMLElement).setPointerCapture?.(ev.pointerId);
    } catch {
      // pointer capture is a nicety, not required for selection
    }
    perf.span('pointer.setSelection', () => this.bridge!.setSelection(pos)); // anchors the IME
    perf.span('pointer.focus', () => this.bridge!.focus());
    for (const cb of this.caretPickListeners) cb(pos);
    for (const p of this.plugins) p.onCaretPick?.(pos);
  };

  /** Offer a key to plugins (before the hidden editor's keymaps). Only when
   *  the event targets this editor (the canvas stack / hidden bridge) or has
   *  no specific target (body — where focus lands after a claimed pointer
   *  gesture), so plugins never steal keys from other inputs on the page. */
  private onKeyDown = (ev: KeyboardEvent): void => {
    // `target` is not always a Node: a key dispatched at `window` (or at the
    // document) reports one of those, and handing either to `contains` throws.
    // Those cases name no element, so they fall through to the same
    // recency test as `<body>` below.
    const node = ev.target instanceof Node ? ev.target : null;
    const mine = node !== null && this.stack.contains(node);
    // Focus is on <body>, so the event names no editor. Only the last one the
    // user touched may act on it — otherwise a second live editor on the page
    // would run its plugins over the same keystroke (see lastInteracted).
    if (!mine && lastInteracted !== null && lastInteracted !== this) return;
    if (node !== null && node !== document.body && !mine) return;
    // Stamp the keyboard-event arrival so refresh() can report the full
    // keydown → painted latency. Ignore pure modifier presses (they produce no
    // edit, so their stamp would otherwise inflate the next real keystroke).
    if (
      perf.enabled &&
      ev.key !== 'Shift' &&
      ev.key !== 'Control' &&
      ev.key !== 'Alt' &&
      ev.key !== 'Meta'
    ) {
      this.inputStartedAt = perf.now();
    }
    const offered: EditorKeyEvent = {
      key: ev.key,
      ctrlKey: ev.ctrlKey,
      metaKey: ev.metaKey,
      shiftKey: ev.shiftKey,
      altKey: ev.altKey,
    };
    for (const p of this.plugins) {
      if (p.onKey?.(offered)) {
        ev.preventDefault();
        ev.stopPropagation();
        return;
      }
    }
  };

  private onPointerMove = (ev: PointerEvent): void => {
    // Offer every move (incl. hover) to pointer plugins for cursor/drag; a claim
    // suppresses the editor's own selection drag.
    if (this.offerPointer('move', ev)) return;
    if (this.dragAnchor == null || !(ev.buttons & 1)) return;
    // Freshest coalesced point; the overlay paints SYNCHRONOUSLY in this very
    // frame — no rAF deferral, no PM transaction (the model follows on
    // pointerup, Google Docs-style).
    const coalesced = ev.getCoalescedEvents?.() ?? [];
    const last = coalesced.length > 0 ? coalesced[coalesced.length - 1] : ev;
    const pos = this.core.posAtEvent(last);
    if (pos == null || pos === this.dragHead) return;
    this.dragHead = pos;
    this.paintDragSelection();
  };

  /** Overlay-only repaint of the in-progress drag selection. */
  private paintDragSelection(): void {
    if (this.dragAnchor == null || this.dragHead == null) return;
    const from = Math.min(this.dragAnchor, this.dragHead);
    const to = Math.max(this.dragAnchor, this.dragHead);
    this.lastCaret = this.core.caretRect(this.dragHead);
    this.lastSelection = from === to ? [] : this.core.selectionRects(from, to);
    this.caretVisible = true;
    this.core.paintOverlay({
      caret: this.lastCaret,
      selection: this.lastSelection,
    });
  }

  private onPointerUp = (ev: PointerEvent): void => {
    if (this.offerPointer('up', ev)) {
      this.dragAnchor = null;
      this.dragHead = null;
      return;
    }
    // Commit the dragged selection to the model exactly once.
    if (
      this.dragAnchor != null &&
      this.dragHead != null &&
      this.dragHead !== this.dragAnchor
    ) {
      this.bridge?.setSelection(this.dragAnchor, this.dragHead);
    }
    this.dragAnchor = null;
    this.dragHead = null;
    if (this.bridge) this.restartBlink(); // back to idle blinking
  };

  private onContextMenu = (ev: MouseEvent): void => {
    if (this.offerPointer('contextmenu', ev)) ev.preventDefault();
  };

  private onDblClick = (ev: MouseEvent): void => {
    const pos = this.core.posAtEvent(ev);
    if (pos == null || !this.bridge) return;
    ev.preventDefault();
    this.bridge.selectWordAt(pos);
    this.bridge.focus();
  };

  /** Offer a pointer event to plugins; returns true if one claimed it. `pos` is
   *  resolved for down/up/contextmenu only (move fires on every hover). */
  private offerPointer(
    type: EditorPointerEvent['type'],
    ev: PointerEvent | MouseEvent,
  ): boolean {
    if (!this.pointerPlugins || !this.core.layout) return false;
    const point = this.core.clientToPage(ev.clientX, ev.clientY);
    const pos = type === 'move' || !point ? null : this.core.posAtPoint(point);
    const pev: EditorPointerEvent = {
      type,
      point,
      pos,
      clientX: ev.clientX,
      clientY: ev.clientY,
      buttons: ev.buttons ?? 0,
      ctrlKey: ev.ctrlKey,
      metaKey: ev.metaKey,
      shiftKey: ev.shiftKey,
      altKey: ev.altKey,
    };
    for (const p of this.plugins) {
      if (p.onPointer?.(pev)) return true;
    }
    return false;
  }

  // ── Overlay affordances (guide / highlight / action button) ─────────

  /** Position (or hide, with null) the transient vertical drag guide. Lives in
   *  the canvas stack so it scrolls/zooms with the pages. */
  private setGuide(guide: OverlayGuide | null): void {
    if (!guide) {
      if (this.guideEl) this.guideEl.style.display = 'none';
      return;
    }
    if (!this.guideEl) {
      this.guideEl = document.createElement('div');
      this.guideEl.style.cssText =
        'position:absolute;width:0;border-left:2px dashed #378add;pointer-events:none;z-index:5;';
      this.stack.appendChild(this.guideEl);
    }
    const top = this.core.pageToCanvas({
      pageIndex: guide.pageIndex,
      x: guide.x,
      y: guide.y,
    });
    const bottom = this.core.pageToCanvas({
      pageIndex: guide.pageIndex,
      x: guide.x,
      y: guide.y + guide.height,
    });
    if (!top || !bottom) {
      this.guideEl.style.display = 'none';
      return;
    }
    this.guideEl.style.display = 'block';
    this.guideEl.style.left = `${top.x}px`;
    this.guideEl.style.top = `${top.y}px`;
    this.guideEl.style.height = `${bottom.y - top.y}px`;
  }

  /** Position (or hide, with null) the object-selection frame: border, 8
   *  resize handles, and a rotate knob above the top edge — all one absolutely
   *  positioned container rotated around its center, so a drag updates plain
   *  DOM (no canvas repaint). Handle geometry stays constant-size on screen. */
  private setFrame(frame: OverlayFrame | null): void {
    if (!frame) {
      if (this.frameEl) this.frameEl.style.display = 'none';
      return;
    }
    if (!this.frameEl) {
      const el = document.createElement('div');
      el.style.cssText =
        'position:absolute;pointer-events:none;z-index:6;transform-origin:center;';
      const border = document.createElement('div');
      border.style.cssText =
        'position:absolute;inset:-1px;border:1.5px solid #378add;';
      el.appendChild(border);
      const handle = (left: string, top: string) => {
        const h = document.createElement('div');
        h.style.cssText =
          `position:absolute;width:7px;height:7px;background:#fff;border:1.5px solid #378add;` +
          `left:${left};top:${top};transform:translate(-50%,-50%);`;
        el.appendChild(h);
      };
      for (const lx of ['0%', '50%', '100%'])
        for (const ty of ['0%', '50%', '100%']) {
          if (lx === '50%' && ty === '50%') continue;
          handle(lx, ty);
        }
      const stem = document.createElement('div');
      stem.style.cssText =
        'position:absolute;left:50%;top:-20px;width:1.5px;height:19px;background:#378add;transform:translateX(-50%);';
      el.appendChild(stem);
      const knob = document.createElement('div');
      knob.style.cssText =
        'position:absolute;left:50%;top:-27px;width:13px;height:13px;border-radius:50%;' +
        'background:#fff;border:1.5px solid #378add;transform:translate(-50%,-50%);';
      el.appendChild(knob);
      const label = document.createElement('div');
      label.dataset['role'] = 'label';
      label.style.cssText =
        'position:absolute;right:0;top:-46px;display:none;background:#042c53;color:#b5d4f4;' +
        'font:11px/1.6 ui-monospace,monospace;padding:0 7px;border-radius:4px;white-space:nowrap;';
      el.appendChild(label);
      // Floating action strip below the frame (e.g. image wrap modes). The
      // frame container is pointer-events:none; the strip re-enables them and
      // swallows its pointer events so a button press never falls through to
      // the caret/selection pipeline underneath.
      const strip = document.createElement('div');
      strip.dataset['role'] = 'actions';
      strip.style.cssText =
        'position:absolute;left:50%;top:calc(100% + 10px);transform:translateX(-50%);' +
        'display:none;gap:2px;pointer-events:auto;background:#fff;border:1px solid #d0d4da;' +
        'border-radius:15px;padding:3px 5px;box-shadow:0 2px 8px rgba(0,0,0,0.12);white-space:nowrap;';
      for (const type of ['pointerdown', 'pointerup', 'mousedown', 'click']) {
        strip.addEventListener(type, (e) => {
          e.stopPropagation();
          if (e.type !== 'click') return;
          const btn = (e.target as HTMLElement).closest<HTMLElement>(
            '[data-action]',
          );
          if (!btn) return;
          const id = btn.dataset['action'] as string;
          for (const p of this.plugins) if (p.onFrameAction?.(id)) return;
        });
      }
      el.appendChild(strip);
      this.stack.appendChild(el);
      this.frameEl = el;
    }
    const tl = this.core.pageToCanvas({
      pageIndex: frame.pageIndex,
      x: frame.x,
      y: frame.y,
    });
    const br = this.core.pageToCanvas({
      pageIndex: frame.pageIndex,
      x: frame.x + frame.width,
      y: frame.y + frame.height,
    });
    if (!tl || !br) {
      this.frameEl.style.display = 'none';
      return;
    }
    const el = this.frameEl;
    el.style.display = 'block';
    el.style.left = `${tl.x}px`;
    el.style.top = `${tl.y}px`;
    el.style.width = `${br.x - tl.x}px`;
    el.style.height = `${br.y - tl.y}px`;
    el.style.transform = frame.rotation ? `rotate(${frame.rotation}deg)` : '';
    const label = el.querySelector<HTMLDivElement>('[data-role="label"]');
    if (label) {
      label.style.display = frame.label ? 'block' : 'none';
      if (frame.label) label.textContent = frame.label;
    }
    const strip = el.querySelector<HTMLDivElement>('[data-role="actions"]');
    if (strip) {
      const acts = frame.actions ?? [];
      // Rebuild only when the set (ids + active states) changes — setFrame
      // runs per pointer-move during drags, and those frames carry no actions
      // anyway (previews omit them).
      const sig = acts.map((a) => `${a.id}${a.active ? '!' : ''}`).join(',');
      if (strip.dataset['sig'] !== sig) {
        strip.dataset['sig'] = sig;
        strip.replaceChildren();
        for (const a of acts) {
          if (a.separator) {
            const d = document.createElement('div');
            d.style.cssText =
              'align-self:center;width:1px;height:16px;background:#d0d4da;margin:0 3px;flex:none;';
            strip.appendChild(d);
            continue;
          }
          const b = document.createElement('button');
          b.type = 'button';
          b.dataset['action'] = a.id;
          b.title = a.title;
          b.setAttribute('aria-label', a.title);
          if (a.active) b.setAttribute('aria-pressed', 'true');
          b.style.cssText =
            'display:inline-flex;align-items:center;justify-content:center;width:28px;height:24px;' +
            'border:0;border-radius:12px;cursor:pointer;padding:0;' +
            (a.active
              ? 'background:#e6f1fb;color:#185fa5;'
              : 'background:transparent;color:#5f5e5a;');
          b.innerHTML = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round">${a.svg}</svg>`;
          strip.appendChild(b);
        }
      }
      strip.style.display = acts.length ? 'flex' : 'none';
      // The strip must stay readable under the frame's rotation.
      strip.style.transform = frame.rotation
        ? `translateX(-50%) rotate(${-frame.rotation}deg)`
        : 'translateX(-50%)';
    }
  }

  /** Fill (or clear, with null) translucent highlight rects — e.g. a selected
   *  table-cell block. Reuses a pool of divs in the canvas stack. */
  private setHighlight(rects: OverlayRect[] | null): void {
    const list = rects ?? [];
    while (this.highlightEls.length < list.length) {
      const el = document.createElement('div');
      el.style.cssText =
        'position:absolute;background:rgba(55,138,221,0.20);pointer-events:none;z-index:4;';
      this.stack.appendChild(el);
      this.highlightEls.push(el);
    }
    this.highlightEls.forEach((el, i) => {
      const r = list[i];
      const tl =
        r && this.core.pageToCanvas({ pageIndex: r.pageIndex, x: r.x, y: r.y });
      const br =
        r &&
        this.core.pageToCanvas({
          pageIndex: r.pageIndex,
          x: r.x + r.width,
          y: r.y + r.height,
        });
      if (!tl || !br) {
        el.style.display = 'none';
        return;
      }
      el.style.display = 'block';
      el.style.left = `${tl.x}px`;
      el.style.top = `${tl.y}px`;
      el.style.width = `${br.x - tl.x}px`;
      el.style.height = `${br.y - tl.y}px`;
    });
  }

  /** Show (or hide, with null `at`) a small action button straddling a page
   *  point — a touch-friendly trigger. Stops pointer propagation so clicking it
   *  doesn't reset the selection underneath. */
  private setActionButton(at: PagePoint | null, onActivate?: () => void): void {
    this.actionHandler = onActivate ?? null;
    if (!at) {
      if (this.actionEl) this.actionEl.style.display = 'none';
      return;
    }
    if (!this.actionEl) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.setAttribute('aria-label', 'Cell actions');
      btn.style.cssText =
        'position:absolute;z-index:6;width:24px;height:24px;display:flex;align-items:center;justify-content:center;padding:0;border:0.5px solid #b5d4f4;border-radius:6px;background:#fff;color:#0c447c;box-shadow:0 1px 4px rgba(0,0,0,.18);cursor:pointer;';
      btn.innerHTML =
        '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 7h8M16 7h4M4 17h4M12 17h8"/><circle cx="14" cy="7" r="2"/><circle cx="10" cy="17" r="2"/></svg>';
      btn.addEventListener('pointerdown', (e) => e.stopPropagation());
      btn.addEventListener('mousedown', (e) => e.stopPropagation());
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.actionHandler?.();
      });
      this.stack.appendChild(btn);
      this.actionEl = btn;
    }
    const p = this.core.pageToCanvas(at);
    if (!p) {
      this.actionEl.style.display = 'none';
      return;
    }
    this.actionEl.style.display = 'flex';
    this.actionEl.style.left = `${p.x - 12}px`;
    this.actionEl.style.top = `${p.y - 12}px`;
  }
}

/**
 * Compose the document schema from a base plus each plugin's `schema`
 * contribution (extra marks/nodes appended). Returns `null` when no plugin
 * contributes anything, so callers can keep using the base schema unchanged.
 */
/** `ctx.plugin` for one plugin: it may reach ONLY the names that plugin put in
 *  its `uses`. Undeclared access throws rather than resolving quietly, so the
 *  declared graph is the real graph — a plugin cannot grow a hidden dependency
 *  that setup ordering then fails to honour. Exported for tests. */
export function pluginLookupFor(
  p: EditorPlugin,
  resolve: (name: string) => EditorPlugin | null,
): (name: string) => never {
  return (name: string): never => {
    if (!p.uses?.includes(name))
      throw new Error(
        `plugin "${p.name}" asked for "${name}" without declaring it in \`uses\``,
      );
    const dep = resolve(name);
    if (!dep)
      throw new Error(
        `plugin "${p.name}" uses "${name}", which is not registered`,
      );
    return dep as never;
  };
}

/** Setup order for a plugin set: dependencies (per `uses`) before dependents.
 *  Throws on an unknown name or a cycle — at REGISTRATION, not first use, so
 *  a bad graph cannot ship. Exported for tests. */
export function orderPluginsByUses(plugins: EditorPlugin[]): EditorPlugin[] {
  const byName = new Map(plugins.map((p) => [p.name, p]));
  for (const p of plugins)
    for (const dep of p.uses ?? [])
      if (!byName.has(dep))
        throw new Error(
          `plugin "${p.name}" uses "${dep}", which is not registered`,
        );
  const done = new Set<string>();
  const visiting = new Set<string>();
  const out: EditorPlugin[] = [];
  const visit = (p: EditorPlugin): void => {
    if (done.has(p.name)) return;
    if (visiting.has(p.name))
      throw new Error(
        `plugin dependency cycle through "${p.name}" (check \`uses\`)`,
      );
    visiting.add(p.name);
    for (const dep of p.uses ?? []) visit(byName.get(dep) as EditorPlugin);
    visiting.delete(p.name);
    done.add(p.name);
    out.push(p);
  };
  for (const p of plugins) visit(p);
  return out;
}

export function composeSchema(
  base: Schema,
  plugins: Iterable<EditorPlugin>,
): Schema | null {
  let nodes = base.spec.nodes;
  let marks = base.spec.marks;
  let changed = false;
  for (const p of plugins) {
    if (!p.schema) continue;
    if (p.schema.nodes) {
      nodes = nodes.append(p.schema.nodes);
      changed = true;
    }
    if (p.schema.marks) {
      marks = marks.append(p.schema.marks);
      changed = true;
    }
  }
  return changed ? new Schema({ nodes, marks }) : null;
}
