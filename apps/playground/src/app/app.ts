import { Component, ElementRef, OnDestroy, signal, viewChild } from '@angular/core';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from '@shadow-garden/bapbong-docx';
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
import { caretRect, hitTest, selectionRects, verticalCaret } from '@shadow-garden/bapbong-selection';
import type {
  CaretRect,
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

@Component({
  selector: 'app-root',
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App implements OnDestroy {
  protected readonly fileName = signal<string | null>(null);
  protected readonly error = signal<string | null>(null);
  protected readonly json = signal<string | null>(null);
  protected readonly headerKeys = signal<string[]>([]);
  protected readonly footerKeys = signal<string[]>([]);
  protected readonly loading = signal(false);
  protected readonly pageCount = signal(0);

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

  private async load(name: string, bytes: ArrayBuffer): Promise<void> {
    this.loading.set(true);
    this.error.set(null);
    this.json.set(null);
    this.fileName.set(name);

    try {
      const { doc, headers, footers, footnotes, titlePg, evenAndOdd, page } = await importDocx(bytes);
      this.headerKeys.set(Object.keys(headers));
      this.footerKeys.set(Object.keys(footers));
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
