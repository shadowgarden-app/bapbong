import { Component, ElementRef, Injector, OnDestroy, afterNextRender, computed, inject, signal, viewChild } from '@angular/core';
import { NgTemplateOutlet } from '@angular/common';
import { ReplyEditorDirective } from './reply-editor.directive';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import { commentSchema, schema } from '@shadow-garden/bapbong-model';
import { importDocx, exportDocx, type DocxImport } from '@shadow-garden/bapbong-docx';
import { createLayoutCache, layout } from '@shadow-garden/bapbong-layout-engine';
import {
  createCanvasMeasurer,
  createCanvasMetrics,
  ensureFontsLoaded,
} from '@shadow-garden/bapbong-measuring';
import { CanvasPainter } from '@shadow-garden/bapbong-painter-canvas';
import {
  CommentComposer,
  InputBridge,
  type MentionHandlers,
  type MentionUser,
  addCommentTr,
  deleteCommentTr,
  editCommentTr,
  moveCaretCommand,
  replyCommentTr,
  resolveCommentTr,
  splitListItem,
  type Command,
  type EditorState,
  type Transaction,
} from '@shadow-garden/bapbong-input-bridge';
import { caretRect, hitTest, selectionRects, verticalCaret } from '@shadow-garden/bapbong-selection';
import type {
  CaretRect,
  CommentNode,
  IUser,
  MeasureMetrics,
  MeasureText,
  PageConfig,
  ResolvedLayout,
  SelectionRect,
} from '@shadow-garden/bapbong-contracts';

/** A4 at 96 dpi with 1in margins — fallback until a document is imported. */
const A4: PageConfig = {
  width: 794,
  height: 1123,
  margin: { top: 96, right: 96, bottom: 96, left: 96 },
};

const CARET_BLINK_MS = 530;
/** The JSON / DOM-preview panels are inspection aids — sync them lazily. */
const PANEL_SYNC_MS = 250;

/** Minimize-mode bubble diameter (px) and the gap between anchored items —
 *  used to resolve overlapping anchors (see packAnchors). Matches .bubble CSS. */
const BUBBLE_SIZE = 34;
const ANCHOR_GAP = 8;

/** Reply indent step (px) and the depth past which indent stops growing — keeps
 *  deeply-nested threads from overflowing the narrow anchored cards. */
const INDENT_STEP = 14;
const MAX_INDENT_DEPTH = 4;

/** Avatar fills, picked per author so each person reads as a distinct colour. */
const AVATAR_COLORS = ['#378ADD', '#7F77DD', '#1D9E75', '#D85A30', '#BA7517', '#D4537E'];

/** Demo users the comment composer can @mention (beyond the doc's own authors). */
const MENTION_SEED: { id: string; label: string }[] = [
  { id: 'alice', label: 'Alice Nguyễn' },
  { id: 'bob', label: 'Bob Trần' },
  { id: 'charlie', label: 'Charlie Lê' },
];

/** What the single comment composer is currently composing. */
interface ComposerSpec {
  kind: 'add' | 'reply' | 'edit';
  label: string;
  from?: number;
  to?: number;
  parentId?: number;
  id?: number;
}

@Component({
  selector: 'app-root',
  templateUrl: './app.html',
  styleUrl: './app.css',
  imports: [NgTemplateOutlet, ReplyEditorDirective],
})
export class App implements OnDestroy {
  protected readonly fileName = signal<string | null>(null);
  protected readonly error = signal<string | null>(null);
  protected readonly json = signal<string | null>(null);
  protected readonly headerKeys = signal<string[]>([]);
  protected readonly footerKeys = signal<string[]>([]);
  protected readonly loading = signal(false);
  protected readonly pageCount = signal(0);
  // Comment threads (from doc.attrs.comments) + authoring UI state.
  protected readonly comments = signal<CommentNode[]>([]);
  // Self-declared display name; the current user (id = name — no auth yet).
  protected readonly author = signal('Me');
  protected readonly currentUser = computed<IUser>(() => {
    const name = this.author().trim() || 'Me';
    return { id: name, name };
  });
  // Comment view modes (Google-Docs style):
  //  hide     — no tint, nothing shown
  //  minimize — avatar bubbles anchored to each comment's position (margin)
  //  expand   — full cards anchored to each comment's position (margin)
  //  panel    — fixed right-hand list of all comments (authoring view)
  protected readonly commentView = signal<'hide' | 'minimize' | 'expand' | 'panel'>('panel');
  protected readonly showResolved = signal(false);
  /** Roots positioned in the margin (minimize/expand), with on-screen top. */
  protected readonly anchoredComments = signal<{ node: CommentNode; top: number }[]>([]);
  /** In minimize mode, the root whose thread popover is open (or null). */
  protected readonly openBubble = signal<number | null>(null);
  /** In expand mode, the focused card — pinned to its true line on re-pack. */
  protected readonly activeCard = signal<number | null>(null);
  /** Raw (unpacked) expand-card anchors, kept so re-packs start from the true y. */
  private expandRaw: { node: CommentNode; top: number }[] = [];
  /** The anchored layer (bubbles/cards), queried to measure expand-card heights. */
  private readonly anchorLayer = viewChild<ElementRef<HTMLDivElement>>('anchorLayer');
  private readonly injector = inject(Injector);
  private allCommentIds: number[] = [];
  protected readonly hasSelection = signal(false);
  protected readonly composerFor = signal<ComposerSpec | null>(null);
  private readonly composerHost = viewChild<ElementRef<HTMLDivElement>>('composerHost');
  private composer: CommentComposer | null = null;
  private resolvedCommentIds: number[] = [];

  private readonly previewHost = viewChild<ElementRef<HTMLDivElement>>('preview');
  // The painter fills this container with one <canvas> per page (virtualized).
  private readonly stackHost = viewChild<ElementRef<HTMLDivElement>>('canvasStack');
  private readonly serializer = DOMSerializer.fromSchema(schema);

  private painter: CanvasPainter | null = null;
  private measureText: MeasureText | null = null;
  private measureMetrics: MeasureMetrics | null = null;
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
  private dragAnchor: number | null = null;
  // Incremental re-layout: unchanged paragraphs skip measuring on each keystroke.
  // Replaced wholesale when late-loading fonts invalidate every measurement.
  private layoutCache = createLayoutCache();
  private readonly onFontsLoaded = () => {
    this.layoutCache = createLayoutCache();
    if (this.bridge) this.refresh(this.bridge.state);
  };
  // Caret blink state (solid on every interaction, toggling while idle).
  private lastCaret: CaretRect | null = null;
  private lastSelection: SelectionRect[] = [];
  private caretVisible = true;
  private blinkTimer: ReturnType<typeof setInterval> | null = null;
  // While dragging, the visual selection leads (painted straight from
  // hit-testing, same frame); the PM model commits once on pointerup.
  private dragHead: number | null = null;
  // Debounced side panels.
  private panelTimer: ReturnType<typeof setTimeout> | null = null;
  // Scroll repaint throttle (page virtualization).
  private scrollRaf: number | null = null;

  protected async onFile(event: Event): Promise<void> {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;
    await this.load(file.name, await file.arrayBuffer());
  }

  protected async loadSample(name = 'sample.docx'): Promise<void> {
    try {
      const res = await fetch(name);
      if (!res.ok) throw new Error(`${name}: HTTP ${res.status}`);
      await this.load(name, await res.arrayBuffer());
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    }
  }

  /** Export the (edited) document back to a .docx and download it. Carries the
   *  imported source package so unmodelled parts survive the round-trip. */
  protected async downloadDocx(): Promise<void> {
    if (!this.bridge) return;
    try {
      const bytes = await exportDocx(this.bridge.state.doc, this.importedRaw ? { carry: this.importedRaw } : undefined);
      const blob = new Blob([bytes as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = (this.fileName() ?? 'document').replace(/\.docx$/i, '') + '-export.docx';
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    }
  }

  private async load(name: string, bytes: ArrayBuffer): Promise<void> {
    this.loading.set(true);
    this.error.set(null);
    this.json.set(null);
    this.fileName.set(name);

    try {
      const { doc, headers, footers, footnotes, titlePg, evenAndOdd, page, raw } = await importDocx(bytes);
      this.importedRaw = raw; // carried on export so unmodelled parts survive
      this.headerKeys.set(Object.keys(headers));
      this.footerKeys.set(Object.keys(footers));
      this.chromeHeaders = headers;
      this.chromeFooters = footers;
      this.chromeTitlePg = titlePg;
      this.chromeEvenAndOdd = evenAndOdd;
      this.footnotes = footnotes;
      this.closeComposer(); // a stale composer would point at the old doc
      this.page = page;
      // comments now ride doc.attrs.comments — refresh() reads them.
      // Measure with the real fonts, not their fallbacks.
      await ensureFontsLoaded(
        collectFontFamilies(doc, ...Object.values(headers), ...Object.values(footers)),
      );
      this.setupEditor(doc);
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    } finally {
      this.loading.set(false);
    }
  }

  /** M4: mount the hidden ProseMirror editor and run the first paint. */
  private setupEditor(doc: ProseMirrorNode): void {
    const stack = this.stackHost()?.nativeElement;
    if (!stack) return;
    this.measureText ??= createCanvasMeasurer();
    this.measureMetrics ??= createCanvasMetrics();
    this.painter ??= new CanvasPainter(stack);
    // Fonts that finish loading later invalidate every measurement.
    document.fonts?.addEventListener?.('loadingdone', this.onFontsLoaded);
    document.addEventListener('keydown', this.onCommentKey);

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
    stack.appendChild(this.bridge.dom);
    this.refresh(this.bridge.state);
  }

  /** Layout (only when the doc changed) → paint → schedule side panels. */
  private refresh(state: EditorState, tr?: Transaction): void {
    if (!this.painter || !this.measureText || !this.measureMetrics) return;
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
      this.pageCount.set(this.resolved.pages.length);
      this.schedulePanelSync(state);
      contentDirty = true;
      // Comment threads live on doc.attrs.comments — refresh the sidebar +
      // the resolved-id set (resolved comments paint no tint).
      const cs = (state.doc.attrs['comments'] as CommentNode[] | null) ?? [];
      this.comments.set(cs);
      const roots = cs.filter((c) => c.parentId == null);
      this.allCommentIds = roots.map((c) => c.id);
      this.resolvedCommentIds = roots.filter((c) => c.resolved).map((c) => c.id);
      this.recomputeAnchors();
    }

    const sel = state.selection;
    this.hasSelection.set(!sel.empty);
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
  }

  /** Debounced sync of the JSON / DOM-preview inspection panels. */
  private schedulePanelSync(state: EditorState): void {
    if (this.panelTimer != null) clearTimeout(this.panelTimer);
    this.panelTimer = setTimeout(() => {
      this.panelTimer = null;
      this.json.set(JSON.stringify(state.doc.toJSON(), null, 2));
      this.renderPreview(state.doc);
    }, PANEL_SYNC_MS);
  }

  /** Redraw caret/selection in the current blink phase (only the affected
   *  page canvases — a blink touches just the caret's page). */
  private repaintOverlay(): void {
    if (!this.painter || !this.resolved) return;
    this.painter.paintOverlay({
      caret: this.caretVisible ? this.lastCaret : null,
      selection: this.lastSelection,
    });
  }

  /** Full content repaint, virtualized to the scroll viewport. */
  private repaintContent(): void {
    if (!this.painter || !this.resolved) return;
    this.painter.paint(this.resolved, {
      caret: this.caretVisible ? this.lastCaret : null,
      selection: this.lastSelection,
      viewport: this.currentViewport(),
      // "Hide comments" suppresses every tint; otherwise only resolved ones.
      resolvedComments: this.commentView() === 'hide' ? this.allCommentIds : this.resolvedCommentIds,
    });
  }

  /** The canvas-wrap's window onto the page stack, in container CSS px. */
  private currentViewport(): { top: number; height: number } | undefined {
    const stack = this.stackHost()?.nativeElement;
    const wrap = stack?.closest('.canvas-wrap');
    if (!stack || !wrap) return undefined;
    const wrapRect = wrap.getBoundingClientRect();
    const stackRect = stack.getBoundingClientRect();
    return { top: wrapRect.top - stackRect.top, height: wrap.clientHeight };
  }

  /** Repaint newly visible pages while scrolling (rAF-throttled). */
  protected onCanvasScroll(): void {
    if (this.scrollRaf != null) return;
    this.scrollRaf = requestAnimationFrame(() => {
      this.scrollRaf = null;
      this.repaintContent();
      // Anchored comments live inside the scroll content → they scroll natively;
      // no per-scroll repositioning needed.
    });
  }

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

  ngOnDestroy(): void {
    this.stopBlink();
    if (this.panelTimer != null) clearTimeout(this.panelTimer);
    if (this.scrollRaf != null) cancelAnimationFrame(this.scrollRaf);
    document.fonts?.removeEventListener?.('loadingdone', this.onFontsLoaded);
    document.removeEventListener('keydown', this.onCommentKey);
    this.composer?.destroy();
    this.bridge?.destroy();
  }

  /** ArrowUp/ArrowDown against the canvas layout (the hidden DOM's own line
   *  wrapping is meaningless). With `extend`, Shift+arrow grows the selection. */
  private verticalCmd(dir: -1 | 1, extend = false): Command {
    return moveCaretCommand((state) => {
      if (!this.resolved || !this.measureText) return null;
      const head = state.selection.head;
      const cr = caretRect(this.resolved, head, this.measureText);
      if (!cr) return null;
      return verticalCaret(this.resolved, head, dir, cr.x, this.measureText);
    }, extend);
  }

  protected onCanvasPointerDown(ev: PointerEvent): void {
    const pos = this.posAtEvent(ev);
    if (pos == null || !this.bridge) return;
    ev.preventDefault(); // keep focus on the hidden editor
    this.dragAnchor = pos;
    this.dragHead = pos;
    // Keep receiving moves when the pointer leaves the canvas mid-drag.
    (ev.target as HTMLElement).setPointerCapture?.(ev.pointerId);
    this.bridge.setSelection(pos); // caret placement (cheap, anchors the IME)
    this.bridge.focus();
    this.selectCommentAt(pos); // clicking commented text selects its anchored card
  }

  /** The unresolved root comment whose marked range covers `pos`, or null. */
  private commentIdAt(pos: number): number | null {
    for (const [id, r] of this.allCommentRanges()) {
      if (pos < r.from || pos > r.to) continue;
      const c = this.comments().find((x) => x.id === id);
      if (c && c.parentId == null && !c.resolved) return id;
    }
    return null;
  }

  /** Clicking commented text in the doc selects the matching anchored card /
   *  bubble, mirroring a click on the card itself (active-snap / open popover).*/
  private selectCommentAt(pos: number): void {
    const id = this.commentIdAt(pos);
    if (id == null) return;
    if (this.commentView() === 'expand') {
      this.onCardFocus(id);
    } else if (this.commentView() === 'minimize' && this.openBubble() !== id) {
      this.openBubble.set(id);
      this.recomputeAnchors();
    } else if (this.commentView() === 'panel') {
      this.activeCard.set(id); // marks the panel card active
      this.scrollPanelTo(id);
    }
  }

  /** Panel mode: clicking a card highlights its text in the doc (and scrolls
   *  the canvas to it), mirroring the bubble/card click in the other modes. */
  protected onPanelCardClick(rootId: number, ev: Event): void {
    // Don't hijack clicks on the card's controls (resolve/delete, reply input).
    if ((ev.target as HTMLElement).closest('button, input, textarea')) return;
    this.activeCard.set(rootId);
    this.onCommentClick(rootId, true, false); // highlight + scroll canvas, keep card focus
  }

  /** Bring the active panel card into view within the panel list. */
  private scrollPanelTo(id: number): void {
    document
      .querySelector(`.comments.panel .panel-thread[data-id="${id}"]`)
      ?.scrollIntoView({ block: 'nearest' });
  }

  protected onCanvasPointerMove(ev: PointerEvent): void {
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
  }

  /** Overlay-only repaint of the in-progress drag selection. */
  private paintDragSelection(): void {
    if (!this.painter || !this.resolved || !this.measureText) return;
    if (this.dragAnchor == null || this.dragHead == null) return;
    const from = Math.min(this.dragAnchor, this.dragHead);
    const to = Math.max(this.dragAnchor, this.dragHead);
    this.lastCaret = caretRect(this.resolved, this.dragHead, this.measureText);
    this.lastSelection = from === to ? [] : selectionRects(this.resolved, from, to, this.measureText);
    this.caretVisible = true;
    this.painter.paintOverlay({ caret: this.lastCaret, selection: this.lastSelection });
  }

  protected onCanvasPointerUp(): void {
    // Commit the dragged selection to the model exactly once.
    if (this.dragAnchor != null && this.dragHead != null && this.dragHead !== this.dragAnchor) {
      this.bridge?.setSelection(this.dragAnchor, this.dragHead);
    }
    this.dragAnchor = null;
    this.dragHead = null;
    if (this.bridge) this.restartBlink(); // back to idle blinking
  }

  protected onCanvasDblClick(ev: MouseEvent): void {
    const pos = this.posAtEvent(ev);
    if (pos == null || !this.bridge) return;
    ev.preventDefault();
    this.bridge.selectWordAt(pos);
    this.bridge.focus();
  }

  /** Select a comment's range (highlight it); `scroll` also brings it into view.
   *  Bubble clicks pass scroll=false — the bubble is already at the right spot.
   *  `focus=false` highlights without grabbing the IME focus (used when the
   *  click should keep focus on a card / its reply box). */
  protected onCommentClick(id: number, scroll = true, focus = true): void {
    const range = this.commentRange(id);
    if (!range || !this.bridge) return;
    this.bridge.setSelection(range.from, range.to);
    if (focus) this.bridge.focus();
    if (scroll && this.painter && this.resolved && this.measureText) {
      const cr = caretRect(this.resolved, range.from, this.measureText);
      const pt = cr && this.painter.pageToCanvas({ pageIndex: cr.pageIndex, x: cr.x, y: cr.y });
      const wrap = this.stackHost()?.nativeElement.closest('.canvas-wrap') as HTMLElement | null;
      if (pt && wrap) wrap.scrollTop = Math.max(0, pt.y - 80);
    }
  }

  /** PM range (from..to) covered by the comment mark carrying `id`, or null. */
  private commentRange(id: number): { from: number; to: number } | null {
    const doc = this.bridge?.state.doc;
    if (!doc) return null;
    let from = Infinity;
    let to = -Infinity;
    doc.descendants((node, pos) => {
      if (!node.isText) return;
      const m = node.marks.find((mk) => mk.type.name === 'comment');
      if (m && (m.attrs['ids'] as number[]).includes(id)) {
        from = Math.min(from, pos);
        to = Math.max(to, pos + node.nodeSize);
      }
    });
    return from <= to ? { from, to } : null;
  }

  // ── Comment authoring ──────────────────────────────────────────────

  /** First letter of a name, for the minimize-mode avatar bubble. */
  protected initial(name: string): string {
    return (name || '?').trim().charAt(0).toUpperCase() || '?';
  }

  /** A stable avatar colour per author (hash the name into the palette). */
  protected avatarColor(name: string): string {
    let h = 0;
    for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
    return AVATAR_COLORS[h % AVATAR_COLORS.length];
  }

  /** Toggle the minimize-mode thread popover for a bubble (and locate it). */
  protected toggleBubble(rootId: number): void {
    this.closeInlineReply(); // its popover is being opened/closed → editor goes away
    this.openBubble.update((cur) => (cur === rootId ? null : rootId));
    // Highlight the range but DON'T scroll — the bubble is already in view.
    if (this.openBubble() === rootId) {
      this.expandThread(rootId); // the popover shows the full thread
      this.onCommentClick(rootId, false);
    }
    // Re-pack: the open bubble snaps to its true line, the rest flow around it
    // (or, on close, the stack settles back to the resting cascade).
    this.recomputeAnchors();
  }

  /** Expand mode: focusing a card makes it the pinned anchor and re-packs.
   *  Activating expands this thread and collapses the others, so the heights
   *  change in the DOM — re-pack AFTER that render (afterNextRender), not now. */
  protected onCardFocus(id: number): void {
    if (this.activeCard() === id) return;
    this.closeInlineReply(); // a reply editor in the previously-active card goes away
    this.activeCard.set(id);
    this.onCommentClick(id, false, false); // highlight the doc range (keep card focus)
    afterNextRender(() => this.packExpandCards(), { injector: this.injector });
  }

  /** Root threads for the panel list (respecting the show-resolved toggle). */
  protected readonly rootComments = computed<CommentNode[]>(() => {
    const showResolved = this.showResolved();
    return this.comments().filter((c) => c.parentId == null && (showResolved || !c.resolved));
  });

  /** A child comment whose per-node reply editor is shown (root box is always
   *  live, so it doesn't need this). */
  protected readonly replyingTo = signal<number | null>(null);
  /** The reply/composer the @mention popup currently drives (set on focus). */
  private activeEditor: CommentComposer | null = null;
  protected setActiveEditor(c: CommentComposer): void {
    this.activeEditor = c;
  }

  /** Reveal the per-node reply editor under comment `id` (root box is always on). */
  protected startInlineReply(id: number): void {
    this.expandThread(this.rootOf(id)); // reveal the thread so the editor shows
    this.replyingTo.set(id);
    if (this.commentView() === 'expand') afterNextRender(() => this.packExpandCards(), { injector: this.injector });
  }

  /** A reply editor (ReplyEditorDirective) committed its content. */
  protected onReplySubmit(e: { target: number; body: unknown }): void {
    if (this.bridge) {
      this.bridge.dispatch(
        replyCommentTr(this.bridge.state, e.target, {
          user: this.currentUser(),
          date: new Date().toISOString(),
          body: e.body,
        }),
      );
    }
    this.replyingTo.set(null); // close the per-node editor (root box stays)
  }

  /** Esc in a reply editor → close its per-node box (no-op for the root box). */
  protected onReplyCancel(target: number): void {
    if (this.replyingTo() === target) this.replyingTo.set(null);
  }

  protected closeInlineReply(): void {
    this.replyingTo.set(null);
    this.mentionState.set(null);
  }

  /** A root's thread (the root + its reply subtree), flattened depth-first.
   *  `replyTo` is the parent author's name (null for the root) so deep replies
   *  whose indent is capped can still show who they answer. */
  protected threadFor(rootId: number): { node: CommentNode; depth: number; replyTo: string | null }[] {
    const all = this.comments();
    const out: { node: CommentNode; depth: number; replyTo: string | null }[] = [];
    const walk = (id: number, depth: number, replyTo: string | null) => {
      const node = all.find((c) => c.id === id);
      if (!node) return;
      out.push({ node, depth, replyTo });
      for (const c of all) if (c.parentId === id) walk(c.id, depth + 1, node.user.name);
    };
    walk(rootId, 0, null);
    return out;
  }

  /** Reply indentation in px, capped so deep threads don't overflow narrow
   *  anchored cards / popovers (beyond the cap, the "↳ parent" hint carries
   *  the context instead of more indent). */
  protected replyIndent(depth: number): number {
    return Math.min(depth, MAX_INDENT_DEPTH) * INDENT_STEP;
  }

  /** Number of replies under a root (the thread minus the root itself). */
  protected replyCount(rootId: number): number {
    return this.threadFor(rootId).length - 1;
  }

  /** Where the inline reply box goes within `rootId`'s thread: the id of the
   *  LAST node in the reply-target's subtree (so the box sits at the end of that
   *  comment's scope, where the new reply will land), or null if the target
   *  isn't in this thread. */
  protected replyAnchorId(rootId: number): number | null {
    const target = this.replyingTo();
    if (target == null) return null;
    const thread = this.threadFor(rootId);
    const i = thread.findIndex((it) => it.node.id === target);
    if (i < 0) return null;
    let last = i;
    for (let j = i + 1; j < thread.length && thread[j].depth > thread[i].depth; j++) last = j;
    return thread[last].node.id;
  }

  /** Indent (px) for the inline reply box = one level under the reply target. */
  protected replyBoxIndent(rootId: number): number {
    const target = this.replyingTo();
    const it = this.threadFor(rootId).find((i) => i.node.id === target);
    return this.replyIndent((it?.depth ?? 0) + 1);
  }

  // ── Shared post-card thread engine (panel / expand / popover) ───────
  /** Threads whose replies + reply box are revealed (footer 💬 toggles this). */
  protected readonly expandedThreads = signal<ReadonlySet<number>>(new Set());
  protected isThreadExpanded(rootId: number): boolean {
    return this.expandedThreads().has(rootId);
  }
  /** The selected thread — its root input box shows. Minimize uses the open
   *  bubble; panel/expand use the active card. */
  protected isActiveThread(rootId: number): boolean {
    return this.commentView() === 'minimize' ? this.openBubble() === rootId : this.activeCard() === rootId;
  }
  protected toggleThread(rootId: number): void {
    const next = new Set(this.expandedThreads());
    if (!next.delete(rootId)) next.add(rootId);
    this.expandedThreads.set(next);
    if (this.commentView() === 'expand') afterNextRender(() => this.packExpandCards(), { injector: this.injector });
  }
  private expandThread(rootId: number): void {
    if (!this.isThreadExpanded(rootId)) this.expandedThreads.set(new Set(this.expandedThreads()).add(rootId));
  }

  protected commentById(id: number): CommentNode | undefined {
    return this.comments().find((c) => c.id === id);
  }
  /** A thread's replies only (depth-first, excluding the root). */
  protected threadReplies(rootId: number): { node: CommentNode; depth: number; replyTo: string | null }[] {
    return this.threadFor(rootId).slice(1);
  }
  /** Root id of the thread containing `id` (walk parents up). */
  private rootOf(id: number): number {
    let cur = this.commentById(id);
    while (cur && cur.parentId != null) cur = this.commentById(cur.parentId);
    return cur?.id ?? id;
  }

  /** Relative comment timestamp (FB-style): "vừa xong" / "{n} phút" / "{n} giờ"
   *  / "{n} ngày". Falls back to the raw string if it isn't a parseable date. */
  protected formatDate(iso: string): string {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    const sec = Math.max(0, Math.floor((Date.now() - d.getTime()) / 1000));
    if (sec < 60) return 'vừa xong';
    const min = Math.floor(sec / 60);
    if (min < 60) return `${min} phút`;
    const hr = Math.floor(min / 60);
    if (hr < 24) return `${hr} giờ`;
    return `${Math.floor(hr / 24)} ngày`;
  }

  /** Absolute timestamp "DD/MM/YYYY hh:mm AM/PM" for the time tooltip. */
  protected absoluteDate(iso: string): string {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    const pad = (n: number) => String(n).padStart(2, '0');
    const ampm = d.getHours() < 12 ? 'AM' : 'PM';
    return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours() % 12 || 12)}:${pad(d.getMinutes())} ${ampm}`;
  }

  /** Plain-text preview of a comment body (commentSchema doc JSON). */
  protected commentText(body: unknown): string {
    try {
      return commentSchema.nodeFromJSON(body).textContent;
    } catch {
      return '';
    }
  }

  /** Comment body as HTML so @mentions render as chips (via the schema's
   *  toDOM). Angular sanitises [innerHTML]; the `.mention` class survives. */
  protected commentBodyHtml(body: unknown): string {
    try {
      const doc = commentSchema.nodeFromJSON(body);
      const div = document.createElement('div');
      div.appendChild(DOMSerializer.fromSchema(commentSchema).serializeFragment(doc.content));
      return div.innerHTML;
    } catch {
      return '';
    }
  }

  /** Whether the current user may edit/delete a node (UX guard, not security
   *  — there is no authenticated identity; the user id is self-declared). */
  protected canModify(node: CommentNode): boolean {
    return node.user.id === this.currentUser().id;
  }

  protected startAddComment(): void {
    const sel = this.bridge?.state.selection;
    if (!sel || sel.empty) return;
    this.openComposer({ kind: 'add', label: 'Bình luận mới', from: sel.from, to: sel.to });
  }

  private openComposer(spec: ComposerSpec, initialBody?: unknown): void {
    this.closeComposer();
    this.closeInlineReply(); // mutually exclusive with an inline reply editor
    this.composerFor.set(spec);
    const host = this.composerHost()?.nativeElement;
    if (host) {
      const composer = new CommentComposer(commentSchema, host, initialBody, this.mentionHandlers);
      this.composer = composer;
      // Route @mention picks here while this composer is focused.
      composer.view.dom.addEventListener('focusin', () => this.setActiveEditor(composer));
      setTimeout(() => composer.focus(), 0); // after the host un-hides
    }
  }

  // ── @mention popup ─────────────────────────────────────────────────
  /** Users the composer can @mention (existing authors + current user + seed). */
  protected readonly mentionUsers = computed<MentionUser[]>(() => {
    const map = new Map<string, string>(MENTION_SEED.map((u) => [u.id, u.label]));
    for (const c of this.comments()) map.set(c.user.id, c.user.name);
    const me = this.currentUser();
    map.set(me.id, me.name);
    return [...map].map(([id, label]) => ({ id, label }));
  });

  /** Popup state driven by the composer's mention plugin (null = hidden). */
  protected readonly mentionState = signal<{
    query: string;
    items: MentionUser[];
    index: number;
    left: number;
    top: number;
  } | null>(null);

  protected readonly mentionHandlers: MentionHandlers = {
    query: (s) => {
      if (!s) return this.mentionState.set(null);
      const q = s.query.toLowerCase();
      const items = this.mentionUsers()
        .filter((u) => u.label.toLowerCase().includes(q))
        .slice(0, 6);
      // Only show when there are matches; otherwise keys pass through to the editor.
      this.mentionState.set(
        items.length ? { query: s.query, items, index: 0, left: s.coords.left, top: s.coords.bottom } : null,
      );
    },
    key: (k) => {
      const st = this.mentionState();
      if (!st) return false;
      if (k === 'esc') return this.mentionState.set(null), true;
      if (k === 'up') return this.mentionState.set({ ...st, index: (st.index - 1 + st.items.length) % st.items.length }), true;
      if (k === 'down') return this.mentionState.set({ ...st, index: (st.index + 1) % st.items.length }), true;
      if (k === 'enter') return this.pickMention(st.items[st.index]), true;
      return false;
    },
  };

  /** Insert the chosen user as a mention chip into the focused editor + close. */
  protected pickMention(user: MentionUser): void {
    this.activeEditor?.applyMention(user);
    this.mentionState.set(null);
  }

  protected submitComposer(): void {
    const spec = this.composerFor();
    if (!spec || !this.bridge || !this.composer || this.composer.isEmpty()) {
      this.closeComposer();
      return;
    }
    const body = this.composer.getJSON();
    const user = this.currentUser();
    const date = new Date().toISOString();
    const state = this.bridge.state;
    let tr: Transaction | null = null;
    if (spec.kind === 'add' && spec.from != null && spec.to != null) {
      tr = addCommentTr(state, { from: spec.from, to: spec.to }, { user, date, body });
    } else if (spec.kind === 'reply' && spec.parentId != null) {
      tr = replyCommentTr(state, spec.parentId, { user, date, body });
    } else if (spec.kind === 'edit' && spec.id != null) {
      tr = editCommentTr(state, spec.id, body);
    }
    if (tr) this.bridge.dispatch(tr);
    this.closeComposer();
  }

  protected closeComposer(): void {
    this.composer?.destroy();
    this.composer = null;
    this.composerFor.set(null);
    this.mentionState.set(null);
  }

  protected resolveComment(rootId: number, resolved: boolean): void {
    if (this.bridge) this.bridge.dispatch(resolveCommentTr(this.bridge.state, rootId, resolved));
  }

  protected deleteComment(id: number): void {
    if (this.bridge) this.bridge.dispatch(deleteCommentTr(this.bridge.state, id));
  }

  /** Switch the comment view mode; repaint (tint follows) + reposition anchors. */
  protected setCommentView(view: 'hide' | 'minimize' | 'expand' | 'panel'): void {
    if (this.commentView() === view) return;
    this.commentView.set(view);
    if (view !== 'panel') this.closeComposer();
    this.closeInlineReply();
    this.openBubble.set(null);
    this.activeCard.set(null);
    this.repaintContent();
    this.recomputeAnchors();
  }

  /** ⌘⌥⇧ + J/M/E/A → hide / minimize / expand / panel (Google-Docs keys). */
  private readonly onCommentKey = (ev: KeyboardEvent): void => {
    if (!(ev.metaKey && ev.altKey && ev.shiftKey)) return;
    const view = { j: 'hide', m: 'minimize', e: 'expand', a: 'panel' } as const;
    const key = ev.key.toLowerCase();
    if (key in view) {
      ev.preventDefault();
      this.setCommentView(view[key as keyof typeof view]);
    }
  };

  /** Position each (unresolved) root comment in the margin at its anchor's
   *  document y — only for the minimize/expand modes. The anchors are rendered
   *  INSIDE the scroll container (same coordinate space as the pages), so the
   *  browser scrolls them natively; we only recompute on layout changes, never
   *  on scroll. */
  private recomputeAnchors(): void {
    const view = this.commentView();
    if (view !== 'minimize' && view !== 'expand') {
      if (this.anchoredComments().length) this.anchoredComments.set([]);
      return;
    }
    if (!this.painter || !this.resolved || !this.measureText) return;
    const ranges = this.allCommentRanges();
    const out: { node: CommentNode; top: number }[] = [];
    for (const c of this.comments()) {
      if (c.parentId != null || c.resolved) continue; // anchored = unresolved roots
      const range = ranges.get(c.id);
      if (!range) continue;
      const cr = caretRect(this.resolved, range.from, this.measureText);
      const pt = cr && this.painter.pageToCanvas({ pageIndex: cr.pageIndex, x: cr.x, y: cr.y });
      if (pt) out.push({ node: c, top: pt.y }); // container/content y; scrolls natively
    }
    if (view === 'minimize') {
      // Fixed-height bubbles: resolve overlaps now, pinning the open bubble to
      // its true line and pushing the rest away (Google-Docs style).
      const items = out.map((o) => ({ ...o, h: BUBBLE_SIZE }));
      this.packAnchors(items, ANCHOR_GAP, this.openBubble());
      this.setAnchors(items);
    } else {
      // Expand cards have dynamic heights — render at the raw y first, then
      // measure and re-pack once the DOM has them (see packExpandCards). Keep
      // the raw tops so a re-pack always works from the true anchors, never
      // from an already-packed layout (which would drift on repeated fires).
      this.expandRaw = out;
      this.setAnchors(out);
      afterNextRender(() => this.packExpandCards(), { injector: this.injector });
    }
  }

  /** Measure the rendered expand cards and pack them apart (the focused card
   *  stays pinned at its true line). Heights are read live, so a focused card —
   *  whose reply box has revealed — is measured taller and the card below it is
   *  pushed down accordingly (no overlap). */
  private packExpandCards(): void {
    if (this.commentView() !== 'expand') return;
    const layer = this.anchorLayer()?.nativeElement;
    if (!layer) return;
    const heightById = new Map<number, number>();
    for (const el of Array.from(layer.querySelectorAll<HTMLElement>('.anchor')))
      heightById.set(Number(el.dataset['id']), el.offsetHeight);
    const items = this.expandRaw.map((c) => ({
      node: c.node,
      top: c.top, // raw (true) y — packed from scratch every time
      h: heightById.get(c.node.id) ?? 80,
    }));
    this.packAnchors(items, ANCHOR_GAP, this.activeCard());
    this.setAnchors(items);
  }

  /** Publish anchors in a STABLE order (by id), independent of their packed y.
   *  Reordering the @for would move DOM nodes and blur a focused card — which
   *  would hide its reply box mid-measure and corrupt the height. */
  private setAnchors(items: { node: CommentNode; top: number }[]): void {
    this.anchoredComments.set(
      items.map(({ node, top }) => ({ node, top })).sort((a, b) => a.node.id - b.node.id),
    );
  }

  /** Push anchored items apart so their boxes don't overlap, using each item's
   *  own height. With an `activeId`, that item stays pinned at its true line
   *  and the others flow around it (push down below, push up above); otherwise
   *  the whole stack cascades top-down. Mutates `items` (sorted by top). */
  private packAnchors(
    items: { node: CommentNode; top: number; h: number }[],
    gap: number,
    activeId: number | null,
  ): void {
    if (items.length < 2) return;
    items.sort((a, b) => a.top - b.top);
    const ai = activeId == null ? -1 : items.findIndex((o) => o.node.id === activeId);
    const minBelow = (i: number) => items[i - 1].top + items[i - 1].h + gap;
    if (ai < 0) {
      for (let i = 1; i < items.length; i++) items[i].top = Math.max(items[i].top, minBelow(i));
    } else {
      for (let i = ai + 1; i < items.length; i++) items[i].top = Math.max(items[i].top, minBelow(i));
      for (let i = ai - 1; i >= 0; i--)
        items[i].top = Math.min(items[i].top, items[i + 1].top - items[i].h - gap);
    }
    // If the upward push ran past the content top, slide the cluster back down
    // so nothing clips (rare; only when many anchors crowd the very top).
    if (items[0].top < 0) {
      const shift = -items[0].top;
      for (const it of items) it.top += shift;
    }
  }

  /** Every comment's PM range in ONE doc scan (id → from..to). */
  private allCommentRanges(): Map<number, { from: number; to: number }> {
    const m = new Map<number, { from: number; to: number }>();
    const doc = this.bridge?.state.doc;
    if (!doc) return m;
    doc.descendants((node, pos) => {
      if (!node.isText) return;
      const mk = node.marks.find((x) => x.type.name === 'comment');
      if (!mk) return;
      for (const id of mk.attrs['ids'] as number[]) {
        const cur = m.get(id) ?? { from: Infinity, to: -Infinity };
        cur.from = Math.min(cur.from, pos);
        cur.to = Math.max(cur.to, pos + node.nodeSize);
        m.set(id, cur);
      }
    });
    return m;
  }

  private posAtEvent(ev: MouseEvent): number | null {
    if (!this.painter || !this.resolved || !this.measureText) return null;
    // Page canvases have pointer-events:none, so events land on the container;
    // map client coords to container-relative CSS px (offsetX/Y would be
    // relative to whichever child the pointer happens to be over).
    const stack = this.stackHost()?.nativeElement;
    if (!stack) return null;
    const rect = stack.getBoundingClientRect();
    const pt = this.painter.canvasToPage(ev.clientX - rect.left, ev.clientY - rect.top);
    if (!pt) return null;
    return hitTest(this.resolved, pt, this.measureText);
  }

  /** Render the document with the schema's own toDOM rules. */
  private renderPreview(doc: ProseMirrorNode): void {
    const host = this.previewHost()?.nativeElement;
    if (!host) return;
    host.replaceChildren(this.serializer.serializeFragment(doc.content, { document }));
  }
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
