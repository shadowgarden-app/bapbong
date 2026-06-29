import { DOMParser as PMDOMParser, Node as ProseMirrorNode, Schema } from 'prosemirror-model';
import { schema as baseSchema } from '@shadow-garden/bapbong-model';
import { RenderCore } from '@shadow-garden/bapbong-view';
import {
  InputBridge,
  moveCaretCommand,
  splitListItem,
  type Command,
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
  EditorPointerEvent,
  OverlayGuide,
  OverlayRect,
  PagePoint,
  PluginContext,
  RangeDecoration,
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
export type { EditorChange, EditorPlugin, PluginContext } from '@shadow-garden/bapbong-contracts';
// The render core is shared with the read-only viewer; re-export so a host can
// reach it (and `BapbongView`) without a separate import.
export { RenderCore } from '@shadow-garden/bapbong-view';
// Built-in ("internal") plugins ship with the editor (see built-in-plugins.ts)
// and are exposed as typed handles (e.g. editor.find) — no install needed.
import { createBuiltins } from './built-in-plugins';
import { Collection } from '@shadow-garden/bapbong-contracts';
import { defaultCommands } from '@shadow-garden/bapbong-commands';
import type { FindPlugin } from './find-plugin';
import type { TableSelectionPlugin } from './table-selection-plugin';
export type { TableSelectionPlugin, CellBlock, SelectedCell } from './table-selection-plugin';
export type { FindPlugin, FindState } from './find-plugin';

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
 * The render half (layout/paint/scroll/zoom/geometry) lives in a shared
 * {@link RenderCore} (the same one the read-only `BapbongView` uses); this class
 * adds the editing half — a hidden ProseMirror editor as the IME / keyboard
 * sink, pointer-driven caret/selection, plugins, commands and clipboard — and
 * pushes a fresh doc + caret overlay to the core on every transaction.
 */
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

  /** Plugin registry keyed by name — built-ins (internal) + host (external). */
  private readonly plugins: Collection<EditorPlugin>;
  private readonly pluginTeardowns: Array<() => void> = [];
  private readonly pluginCtx: PluginContext;
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

  /** Headless editor commands keyed by name — the surface a toolbar/menubar
   *  renders and dispatches against (`editor.commands.get('bold')?.run(...)`).
   *  Built from the shared command layer; the same ops run on a backend.
   *  (Plugin-contributed commands are an additive follow-up.) */
  readonly commands: Collection<EditorCommand> = defaultCommands();

  constructor(stack: HTMLElement, opts: BapbongEditorOptions = {}) {
    this.stack = stack;
    this.core = new RenderCore(stack, { viewport: opts.viewport });
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

    // Plugins: build their context and run setup (teardowns collected for destroy).
    // Internal (built-in) plugins first, then external/host-provided plugins.
    this.plugins = new Collection<EditorPlugin>([...createBuiltins(), ...(opts.plugins ?? [])], {
      idProperty: 'name',
    });
    this.pointerPlugins = [...this.plugins].some((p) => p.onPointer);
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
      caretRect: (pos: number) => this.core.caretRect(pos),
      pageToCanvas: (p: PagePoint) => this.core.pageToCanvas(p),
      setSelection: (from: number, to?: number) => this.setSelection(from, to),
      scrollToPos: (pos: number, topMargin?: number) => this.core.scrollToPos(pos, topMargin),
      requestPaint: () => this.core.paintContent(this.currentOverlay()),
      setCursor: (cursor: string | null) => {
        this.stack.style.cursor = cursor ?? '';
      },
      setGuide: (guide: OverlayGuide | null) => this.setGuide(guide),
      setHighlight: (rects: OverlayRect[] | null) => this.setHighlight(rects),
      setActionButton: (at: PagePoint | null, onActivate?: () => void) => this.setActionButton(at, onActivate),
    };
    // `state` + `layout` are live (read on each access); arrow getters keep them
    // current without throwing at construction (the doc loads later).
    Object.defineProperty(ctx, 'state', { enumerable: true, get: () => this.state });
    Object.defineProperty(ctx, 'layout', { enumerable: true, get: () => this.core.layout });
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

  /** The document schema in use (model's base + plugin schema contributions).
   *  Hosts serialize/parse comment bodies, previews, etc. against this. */
  get schema(): Schema {
    return this.core.schema;
  }

  /** Built-in find-and-replace (always registered; the host renders the bar).
   *  Cast is safe by construction — createBuiltins() always registers it. */
  get find(): FindPlugin {
    return this.plugins.get('find') as FindPlugin;
  }

  /** Built-in table cell-range selection (drag across cells). The host opens
   *  cell properties from its action icon / right-click via `onAction`. */
  get tableSelection(): TableSelectionPlugin {
    return this.plugins.get('table-selection') as TableSelectionPlugin;
  }

  /** Import a .docx, lay it out, and paint the first frame. Resolves with the
   *  imported page-chrome keys (the rest of the import rides on the doc model
   *  exposed via `state`). */
  async loadDocx(bytes: ArrayBuffer): Promise<{ headerKeys: string[]; footerKeys: string[] }> {
    // Compose the doc schema from model's base + any plugin schema
    // contributions, and import against it (so plugin-owned marks/nodes parse).
    const composed = composeSchema(baseSchema, this.plugins);
    const { doc, headerKeys, footerKeys } = await this.core.loadDocx(
      bytes,
      // Skip the core's initial paint — `mount` paints with a caret overlay.
      { schema: composed ?? undefined, paint: false },
    );
    this.mount(doc);
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
      const pos = (offsets[blockIdx] ?? 0) + 1;
      const cr = this.core.caretRect(pos);
      const page = cr && this.core.layout?.pages[cr.pageIndex];
      out.push({
        index: b,
        newPage: sections[b + 1].newPage,
        columns: sections[b + 1].columns.count,
        pos,
        rect: cr && page ? { pageIndex: cr.pageIndex, x: 0, y: cr.y, width: page.width } : null,
      });
    }
    return out;
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

  /** Focus the hidden ProseMirror editor (keyboard/IME sink). */
  focus(): void {
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

  /** Paste clipboard content (HTML → parsed by the schema, else plain text). */
  async paste(): Promise<void> {
    const view = this.bridge?.view;
    if (!view) return;
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes('text/html')) {
          const html = await (await item.getType('text/html')).text();
          const dom = document.createElement('div');
          dom.innerHTML = html;
          const slice = PMDOMParser.fromSchema(view.state.schema).parseSlice(dom);
          view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView());
          return;
        }
      }
      await this.pasteText();
    } catch {
      /* clipboard unavailable (permission / sandbox) */
    }
  }

  /** Paste clipboard text as plain text (no formatting). */
  async pasteText(): Promise<void> {
    const view = this.bridge?.view;
    if (!view) return;
    try {
      const text = await navigator.clipboard.readText();
      if (text) view.dispatch(view.state.tr.insertText(text).scrollIntoView());
    } catch {
      /* clipboard unavailable */
    }
  }

  destroy(): void {
    for (const t of this.pluginTeardowns) t();
    this.pluginTeardowns.length = 0;
    this.stopBlink();
    this.offFonts();
    this.stack.removeEventListener('pointerdown', this.onPointerDown);
    this.stack.removeEventListener('pointermove', this.onPointerMove);
    this.stack.removeEventListener('pointerup', this.onPointerUp);
    this.stack.removeEventListener('pointercancel', this.onPointerUp);
    this.stack.removeEventListener('dblclick', this.onDblClick);
    this.stack.removeEventListener('contextmenu', this.onContextMenu);
    this.guideEl?.remove();
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
    // The core already laid out this doc in `loadDocx`; paint the first frame
    // with the caret overlay (no redundant relayout).
    this.render(this.bridge.state, true);
  }

  /** Bridge update hook: relayout (only when the doc changed) → paint → place
   *  IME → notify. */
  private refresh(state: EditorState, tr?: Transaction): void {
    const docChanged = tr?.docChanged ?? true;
    if (docChanged || !this.core.layout) this.core.layoutDoc(state.doc);
    this.render(state, docChanged);
  }

  /** Compute the caret/selection overlay from `state`, paint it (full content
   *  when `contentDirty`, else overlay-only), anchor the IME, and notify. */
  private render(state: EditorState, contentDirty: boolean): void {
    const sel = state.selection;
    this.lastCaret = this.core.caretRect(sel.head);
    this.lastSelection = sel.empty ? [] : this.core.selectionRects(sel.from, sel.to);

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
      const pt = this.core.pageToCanvas({ pageIndex: caret.pageIndex, x: caret.x, y: caret.y });
      if (pt) this.bridge.place(pt.x, pt.y, caret.height);
    }

    this.emitChange({ state, pageCount: this.core.pageCount, docChanged: contentDirty });
  }

  /** The caret/selection overlay in the current blink phase. */
  private currentOverlay(): { caret: CaretRect | null; selection: SelectionRect[] } {
    return { caret: this.caretVisible ? this.lastCaret : null, selection: this.lastSelection };
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
    this.bridge.setSelection(pos); // caret placement (cheap, anchors the IME)
    this.bridge.focus();
    for (const cb of this.caretPickListeners) cb(pos);
    for (const p of this.plugins) p.onCaretPick?.(pos);
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
    this.core.paintOverlay({ caret: this.lastCaret, selection: this.lastSelection });
  }

  private onPointerUp = (ev: PointerEvent): void => {
    if (this.offerPointer('up', ev)) {
      this.dragAnchor = null;
      this.dragHead = null;
      return;
    }
    // Commit the dragged selection to the model exactly once.
    if (this.dragAnchor != null && this.dragHead != null && this.dragHead !== this.dragAnchor) {
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
  private offerPointer(type: EditorPointerEvent['type'], ev: PointerEvent | MouseEvent): boolean {
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
    const top = this.core.pageToCanvas({ pageIndex: guide.pageIndex, x: guide.x, y: guide.y });
    const bottom = this.core.pageToCanvas({ pageIndex: guide.pageIndex, x: guide.x, y: guide.y + guide.height });
    if (!top || !bottom) {
      this.guideEl.style.display = 'none';
      return;
    }
    this.guideEl.style.display = 'block';
    this.guideEl.style.left = `${top.x}px`;
    this.guideEl.style.top = `${top.y}px`;
    this.guideEl.style.height = `${bottom.y - top.y}px`;
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
      const tl = r && this.core.pageToCanvas({ pageIndex: r.pageIndex, x: r.x, y: r.y });
      const br = r && this.core.pageToCanvas({ pageIndex: r.pageIndex, x: r.x + r.width, y: r.y + r.height });
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
export function composeSchema(base: Schema, plugins: Iterable<EditorPlugin>): Schema | null {
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
