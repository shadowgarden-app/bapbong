import { Node as ProseMirrorNode } from 'prosemirror-model';
import {
  importDocx,
  exportDocx,
  type DocxImport,
} from '@shadow-garden/bapbong-docx';
import { createLayoutCache, layout } from '@shadow-garden/bapbong-layout-engine';
import {
  createCanvasMeasurer,
  createCanvasMetrics,
  ensureFontsLoaded,
} from '@shadow-garden/bapbong-measuring';
import { CanvasPainter } from '@shadow-garden/bapbong-painter-canvas';
import {
  InputBridge,
  moveCaretCommand,
  splitListItem,
  type Command,
  type EditorState,
  type Transaction,
} from '@shadow-garden/bapbong-input-bridge';
import {
  caretRect,
  hitTest,
  selectionRects,
  verticalCaret,
} from '@shadow-garden/bapbong-selection';
import type {
  CaretRect,
  EditorChange,
  EditorPlugin,
  MeasureMetrics,
  MeasureText,
  PageConfig,
  PluginContext,
  ResolvedLayout,
  SelectionRect,
} from '@shadow-garden/bapbong-contracts';

// The plugin contract's canonical home is `contracts`; re-export it here so a
// plugin author can `import { EditorPlugin, PluginContext, EditorChange } from
// '@shadow-garden/bapbong-editor'`.
export type { EditorChange, EditorPlugin, PluginContext } from '@shadow-garden/bapbong-contracts';

/** A4 at 96 dpi with 1in margins — fallback until a document is imported. */
const A4: PageConfig = {
  width: 794,
  height: 1123,
  margin: { top: 96, right: 96, bottom: 96, left: 96 },
};

const CARET_BLINK_MS = 530;

export interface BapbongEditorOptions {
  /** The scroll viewport the page stack lives in (used for virtualization and
   *  scroll-into-view). Defaults to `stack.closest('.canvas-wrap')`. */
  viewport?: HTMLElement;
  /** Editor plugins. Their lifecycle/event hooks are invoked by the core; they
   *  reach back through the PluginContext handed to `setup`. */
  plugins?: EditorPlugin[];
}

/**
 * The framework-agnostic core of bapbong: it owns the render → edit loop
 * (import → layout → paint → input → selection → export) and the per-page
 * canvas stack, with no UI chrome of its own. A host wires comment UI / toolbars
 * around it via `onChange`, `caretRect`, `pageToCanvas`, `dispatch`, etc.
 *
 * The editor renders into a caller-supplied `stack` element (it appends one
 * `<canvas>` per page, virtualized to the viewport) and mounts a hidden
 * ProseMirror editor there as the IME / keyboard sink.
 */
export class BapbongEditor {
  private readonly stack: HTMLElement;
  private readonly viewport: HTMLElement | null;

  private readonly painter: CanvasPainter;
  private readonly measureText: MeasureText;
  private readonly measureMetrics: MeasureMetrics;
  private bridge: InputBridge | null = null;
  private resolved: ResolvedLayout | null = null;

  // Page chrome from the imported docx, keyed by w:type (default/first/even).
  private chromeHeaders: Record<string, ProseMirrorNode> = {};
  private chromeFooters: Record<string, ProseMirrorNode> = {};
  private chromeTitlePg = false;
  private chromeEvenAndOdd = false;
  // Footnote bodies keyed by display number (laid out at the page bottom).
  private footnotes: Record<number, ProseMirrorNode> | undefined;
  // The imported source package — passed to exportDocx({ carry }) so styles /
  // numbering / headers / media survive a round-trip.
  private importedRaw: DocxImport['raw'] | null = null;
  // Page geometry from the imported docx's sectPr (A4 until imported).
  private page: PageConfig = A4;

  // Incremental re-layout: unchanged paragraphs skip measuring on each keystroke.
  // Replaced wholesale when late-loading fonts invalidate every measurement.
  private layoutCache = createLayoutCache();

  // Caret blink state (solid on every interaction, toggling while idle).
  private lastCaret: CaretRect | null = null;
  private lastSelection: SelectionRect[] = [];
  private caretVisible = true;
  private blinkTimer: ReturnType<typeof setInterval> | null = null;
  // While dragging, the visual selection leads (painted straight from
  // hit-testing, same frame); the PM model commits once on pointerup.
  private dragAnchor: number | null = null;
  private dragHead: number | null = null;
  // Scroll repaint throttle (page virtualization).
  private scrollRaf: number | null = null;

  // Comment ids painted WITHOUT a tint (resolved, or all when hidden). The host
  // owns the comment view policy and pushes the set here via setSuppressedComments.
  private suppressedComments: number[] = [];

  private readonly changeListeners = new Set<(c: EditorChange) => void>();
  private readonly caretPickListeners = new Set<(pos: number) => void>();

  private readonly plugins: EditorPlugin[];
  private readonly pluginTeardowns: Array<() => void> = [];
  private readonly pluginCtx: PluginContext;

  constructor(stack: HTMLElement, opts: BapbongEditorOptions = {}) {
    this.stack = stack;
    this.viewport =
      opts.viewport ?? (stack.closest('.canvas-wrap') as HTMLElement | null);
    this.painter = new CanvasPainter(stack);
    this.measureText = createCanvasMeasurer();
    this.measureMetrics = createCanvasMetrics();

    stack.addEventListener('pointerdown', this.onPointerDown);
    stack.addEventListener('pointermove', this.onPointerMove);
    stack.addEventListener('pointerup', this.onPointerUp);
    stack.addEventListener('pointercancel', this.onPointerUp);
    stack.addEventListener('dblclick', this.onDblClick);
    this.viewport?.addEventListener('scroll', this.onScroll);
    // Fonts that finish loading later invalidate every measurement.
    document.fonts?.addEventListener?.('loadingdone', this.onFontsLoaded);

    // Plugins: build their context and run setup (teardowns collected for destroy).
    this.plugins = opts.plugins ?? [];
    this.pluginCtx = this.makePluginContext();
    for (const p of this.plugins) {
      const teardown = p.setup?.(this.pluginCtx);
      if (teardown) this.pluginTeardowns.push(teardown);
    }
  }

  /** The controlled surface handed to each plugin (live state + geometry). */
  private makePluginContext(): PluginContext {
    // Arrow methods capture `this` lexically (no this-alias).
    const ctx = {
      dispatch: (tr: Transaction) => this.dispatch(tr),
      caretRect: (pos: number) => this.caretRect(pos),
      pageToCanvas: (p: { pageIndex: number; x: number; y: number }) => this.pageToCanvas(p),
      setSelection: (from: number, to?: number) => this.setSelection(from, to),
      scrollToPos: (pos: number, topMargin?: number) => this.scrollToPos(pos, topMargin),
      requestPaint: () => this.repaintContent(),
    };
    // `state` is live (read on each access); an arrow getter keeps it current
    // without throwing at construction (the doc loads later).
    Object.defineProperty(ctx, 'state', { enumerable: true, get: () => this.state });
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
    return this.resolved?.pages.length ?? 0;
  }

  /** Import a .docx, lay it out, and paint the first frame. Resolves with the
   *  imported page-chrome keys (the rest of the import rides on the doc model
   *  exposed via `state`). */
  async loadDocx(
    bytes: ArrayBuffer,
  ): Promise<{ headerKeys: string[]; footerKeys: string[] }> {
    const { doc, headers, footers, footnotes, titlePg, evenAndOdd, page, raw } =
      await importDocx(bytes);
    this.importedRaw = raw; // carried on export so unmodelled parts survive
    this.chromeHeaders = headers;
    this.chromeFooters = footers;
    this.chromeTitlePg = titlePg;
    this.chromeEvenAndOdd = evenAndOdd;
    this.footnotes = footnotes;
    this.page = page;
    // Measure with the real fonts, not their fallbacks.
    await ensureFontsLoaded(
      collectFontFamilies(doc, ...Object.values(headers), ...Object.values(footers)),
    );
    this.mount(doc);
    return { headerKeys: Object.keys(headers), footerKeys: Object.keys(footers) };
  }

  /** Export the (edited) document back to .docx bytes, carrying the imported
   *  source package so unmodelled parts survive the round-trip. */
  async exportDocx(): Promise<Uint8Array> {
    if (!this.bridge) throw new Error('BapbongEditor: no document loaded');
    return exportDocx(
      this.bridge.state.doc,
      this.importedRaw ? { carry: this.importedRaw } : undefined,
    );
  }

  /** Apply a transaction (the host builds comment/edit transactions against
   *  `state` and dispatches them here). */
  dispatch(tr: Transaction): void {
    this.bridge?.dispatch(tr);
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
    return this.resolved ? caretRect(this.resolved, pos, this.measureText) : null;
  }

  /** Map a page-local point to container (canvas-stack) coordinates, or null. */
  pageToCanvas(p: {
    pageIndex: number;
    x: number;
    y: number;
  }): { x: number; y: number } | null {
    return this.painter.pageToCanvas(p);
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
    const cr = this.caretRect(pos);
    const pt = cr && this.painter.pageToCanvas({ pageIndex: cr.pageIndex, x: cr.x, y: cr.y });
    if (pt && this.viewport) this.viewport.scrollTop = Math.max(0, pt.y - topMargin);
  }

  /** Comment ids to paint WITHOUT a tint (resolved, or all when hidden). The
   *  host owns the view policy; this repaints with the new set. */
  setSuppressedComments(ids: number[]): void {
    this.suppressedComments = ids;
    this.repaintContent();
  }

  /** Focus the hidden ProseMirror editor (keyboard/IME sink). */
  focus(): void {
    this.bridge?.focus();
  }

  destroy(): void {
    for (const t of this.pluginTeardowns) t();
    this.pluginTeardowns.length = 0;
    this.stopBlink();
    if (this.scrollRaf != null) cancelAnimationFrame(this.scrollRaf);
    this.stack.removeEventListener('pointerdown', this.onPointerDown);
    this.stack.removeEventListener('pointermove', this.onPointerMove);
    this.stack.removeEventListener('pointerup', this.onPointerUp);
    this.stack.removeEventListener('pointercancel', this.onPointerUp);
    this.stack.removeEventListener('dblclick', this.onDblClick);
    this.viewport?.removeEventListener('scroll', this.onScroll);
    document.fonts?.removeEventListener?.('loadingdone', this.onFontsLoaded);
    this.bridge?.destroy();
    this.bridge = null;
  }

  // ── Setup / render loop ─────────────────────────────────────────────

  /** Mount the hidden ProseMirror editor on a fresh doc and run the first paint. */
  private mount(doc: ProseMirrorNode): void {
    this.bridge?.destroy();
    this.bridge = new InputBridge({
      doc,
      keys: {
        Enter: splitListItem, // continue lists; falls through outside them
        ArrowUp: this.verticalCmd(-1),
        ArrowDown: this.verticalCmd(1),
        'Shift-ArrowUp': this.verticalCmd(-1, true),
        'Shift-ArrowDown': this.verticalCmd(1, true),
      },
      onUpdate: (state, tr) => this.refresh(state, tr),
    });
    // Hidden editor lives in the page-canvas container, so IME anchoring
    // scrolls along (positioned at the painted caret, in container coords).
    this.stack.appendChild(this.bridge.dom);
    this.refresh(this.bridge.state);
  }

  /** Layout (only when the doc changed) → paint → place IME → notify the host. */
  private refresh(state: EditorState, tr?: Transaction): void {
    const docChanged = tr?.docChanged ?? true;

    let contentDirty = false;
    if (docChanged || !this.resolved) {
      this.resolved = layout(
        state.doc,
        { page: this.page, measureText: this.measureText, measureMetrics: this.measureMetrics },
        this.layoutCache,
        {
          header: this.chromeHeaders['default'],
          footer: this.chromeFooters['default'],
          headerFirst: this.chromeHeaders['first'],
          footerFirst: this.chromeFooters['first'],
          headerEven: this.chromeHeaders['even'],
          footerEven: this.chromeFooters['even'],
          titlePg: this.chromeTitlePg,
          evenAndOdd: this.chromeEvenAndOdd,
        },
        this.footnotes,
      );
      contentDirty = true;
    }

    const sel = state.selection;
    this.lastCaret = caretRect(this.resolved, sel.head, this.measureText);
    this.lastSelection = sel.empty
      ? []
      : selectionRects(this.resolved, sel.from, sel.to, this.measureText);

    // While dragging the caret stays solid and the timer rests; otherwise every
    // interaction restarts the blink phase.
    if (this.dragAnchor != null) {
      this.stopBlink();
      this.caretVisible = true;
    } else {
      this.restartBlink();
    }

    if (contentDirty) {
      this.repaintContent();
    } else {
      this.repaintOverlay(); // redraw only the caret/selection page(s)
    }

    // Anchor the hidden editor (and its IME popup) at the painted caret
    // (pageToCanvas returns container-relative coords, where the editor lives).
    const caret = this.lastCaret;
    if (caret && this.bridge) {
      const pt = this.painter.pageToCanvas({ pageIndex: caret.pageIndex, x: caret.x, y: caret.y });
      if (pt) this.bridge.place(pt.x, pt.y, caret.height);
    }

    this.emitChange({ state, pageCount: this.pageCount, docChanged });
  }

  private emitChange(c: EditorChange): void {
    for (const cb of this.changeListeners) cb(c);
    for (const p of this.plugins) p.onChange?.(c);
  }

  /** Redraw caret/selection in the current blink phase (only the affected
   *  page canvases — a blink touches just the caret's page). */
  private repaintOverlay(): void {
    if (!this.resolved) return;
    this.painter.paintOverlay({
      caret: this.caretVisible ? this.lastCaret : null,
      selection: this.lastSelection,
    });
  }

  /** Full content repaint, virtualized to the scroll viewport. */
  private repaintContent(): void {
    if (!this.resolved) return;
    this.painter.paint(this.resolved, {
      caret: this.caretVisible ? this.lastCaret : null,
      selection: this.lastSelection,
      viewport: this.currentViewport(),
      // Ids the host marks tint-suppressed (resolved comments, or all when hidden).
      resolvedComments: this.suppressedComments,
    });
  }

  /** The viewport's window onto the page stack, in container CSS px. */
  private currentViewport(): { top: number; height: number } | undefined {
    const wrap = this.viewport;
    if (!wrap) return undefined;
    const wrapRect = wrap.getBoundingClientRect();
    const stackRect = this.stack.getBoundingClientRect();
    return { top: wrapRect.top - stackRect.top, height: wrap.clientHeight };
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
  private verticalCmd(dir: -1 | 1, extend = false): Command {
    return moveCaretCommand((state) => {
      if (!this.resolved) return null;
      const head = state.selection.head;
      const cr = caretRect(this.resolved, head, this.measureText);
      if (!cr) return null;
      return verticalCaret(this.resolved, head, dir, cr.x, this.measureText);
    }, extend);
  }

  // ── Pointer / scroll ────────────────────────────────────────────────

  private onPointerDown = (ev: PointerEvent): void => {
    const pos = this.posAtEvent(ev);
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
    this.bridge.setSelection(pos); // caret placement (cheap, anchors the IME)
    this.bridge.focus();
    for (const cb of this.caretPickListeners) cb(pos);
    for (const p of this.plugins) p.onCaretPick?.(pos);
  };

  private onPointerMove = (ev: PointerEvent): void => {
    if (this.dragAnchor == null || !(ev.buttons & 1)) return;
    // Freshest coalesced point; the overlay paints SYNCHRONOUSLY in this very
    // frame — no rAF deferral, no PM transaction (the model follows on
    // pointerup, Google Docs-style).
    const coalesced = ev.getCoalescedEvents?.() ?? [];
    const last = coalesced.length > 0 ? coalesced[coalesced.length - 1] : ev;
    const pos = this.posAtEvent(last);
    if (pos == null || pos === this.dragHead) return;
    this.dragHead = pos;
    this.paintDragSelection();
  };

  /** Overlay-only repaint of the in-progress drag selection. */
  private paintDragSelection(): void {
    if (!this.resolved || this.dragAnchor == null || this.dragHead == null) return;
    const from = Math.min(this.dragAnchor, this.dragHead);
    const to = Math.max(this.dragAnchor, this.dragHead);
    this.lastCaret = caretRect(this.resolved, this.dragHead, this.measureText);
    this.lastSelection = from === to ? [] : selectionRects(this.resolved, from, to, this.measureText);
    this.caretVisible = true;
    this.painter.paintOverlay({ caret: this.lastCaret, selection: this.lastSelection });
  }

  private onPointerUp = (): void => {
    // Commit the dragged selection to the model exactly once.
    if (this.dragAnchor != null && this.dragHead != null && this.dragHead !== this.dragAnchor) {
      this.bridge?.setSelection(this.dragAnchor, this.dragHead);
    }
    this.dragAnchor = null;
    this.dragHead = null;
    if (this.bridge) this.restartBlink(); // back to idle blinking
  };

  private onDblClick = (ev: MouseEvent): void => {
    const pos = this.posAtEvent(ev);
    if (pos == null || !this.bridge) return;
    ev.preventDefault();
    this.bridge.selectWordAt(pos);
    this.bridge.focus();
  };

  /** Repaint newly visible pages while scrolling (rAF-throttled). */
  private onScroll = (): void => {
    if (this.scrollRaf != null) return;
    this.scrollRaf = requestAnimationFrame(() => {
      this.scrollRaf = null;
      this.repaintContent();
      // Anchored host UI lives inside the scroll content → it scrolls natively;
      // no per-scroll repositioning needed.
    });
  };

  private posAtEvent(ev: MouseEvent): number | null {
    if (!this.resolved) return null;
    // Page canvases have pointer-events:none, so events land on the container;
    // map client coords to container-relative CSS px (offsetX/Y would be
    // relative to whichever child the pointer happens to be over).
    const rect = this.stack.getBoundingClientRect();
    const pt = this.painter.canvasToPage(ev.clientX - rect.left, ev.clientY - rect.top);
    if (!pt) return null;
    return hitTest(this.resolved, pt, this.measureText);
  }

  private readonly onFontsLoaded = (): void => {
    this.layoutCache = createLayoutCache();
    if (this.bridge) this.refresh(this.bridge.state);
  };
}

/** Every fontFamily mark in the given documents, plus the engine default. */
function collectFontFamilies(...docs: (ProseMirrorNode | undefined)[]): string[] {
  const families = new Set<string>(['Arial']);
  for (const doc of docs) {
    doc?.descendants((node) => {
      for (const mark of node.marks) {
        if (mark.type.name === 'fontFamily' && mark.attrs['family']) {
          families.add(String(mark.attrs['family']));
        }
      }
    });
  }
  return [...families];
}
