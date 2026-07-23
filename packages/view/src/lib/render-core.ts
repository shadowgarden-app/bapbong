import { Node as ProseMirrorNode, Schema } from 'prosemirror-model';
import { schema as baseSchema } from '@shadow-garden/bapbong-model';
import {
  importDocx,
  exportDocx,
  type DocxImport,
} from '@shadow-garden/bapbong-docx';
import {
  createLayoutCache,
  layout,
} from '@shadow-garden/bapbong-layout-engine';
import {
  createCanvasMeasurer,
  createCanvasMetrics,
  ensureFontsLoaded,
} from '@shadow-garden/bapbong-measuring';
import { CanvasPainter } from '@shadow-garden/bapbong-painter-canvas';
import {
  caretRect,
  hitTest,
  selectionRects,
  verticalCaret,
} from '@shadow-garden/bapbong-selection';
import { A11yMirror } from '@shadow-garden/bapbong-a11y';
import { perf } from '@shadow-garden/bapbong-contracts';
import type {
  CaretRect,
  MeasureMetrics,
  MeasureText,
  PageConfig,
  PagePoint,
  PaintDecoration,
  RangeDecoration,
  ResolvedLayout,
  SelectionRect,
} from '@shadow-garden/bapbong-contracts';

/** A4 at 96 dpi with 1in margins — fallback until a document is imported. */
const A4: PageConfig = {
  width: 794,
  height: 1123,
  margin: { top: 96, right: 96, bottom: 96, left: 96 },
};

export interface RenderCoreOptions {
  /** The scroll viewport the page stack lives in (used for virtualization and
   *  scroll-into-view). Defaults to `stack.closest('.canvas-wrap')`. */
  viewport?: HTMLElement;
  /** Initial zoom factor (1 = 100%). */
  zoom?: number;
  /** Build a visually-hidden ARIA mirror of the document for screen readers
   *  (the canvas is opaque to assistive tech). Default true; pass false to opt
   *  out. `a11yLabel` sets its `aria-label`. */
  a11y?: boolean;
  a11yLabel?: string;
  /** Text-width measurer. Defaults to a canvas-backed one (browser). Inject an
   *  engine-independent measurer (e.g. font-file metrics) to make wrapping and
   *  pagination deterministic across WebView engines / headless. */
  measureText?: MeasureText;
  /** Vertical-metrics provider, paired with {@link measureText}. Defaults to a
   *  canvas-backed one. */
  measureMetrics?: MeasureMetrics;
}

/** Caret + selection to paint on the overlay layer (page-local geometry). */
export interface Overlay {
  caret?: CaretRect | null;
  selection?: SelectionRect[];
}

/**
 * The render half of bapbong, with **no editing**: import → layout → paint →
 * scroll/zoom/virtualize, plus the geometry queries (`caretRect`, `hitTest`,
 * `pageToCanvas`, …) the editor and viewer both need.
 *
 * It holds a **read-only ProseMirror doc** (a `Node`, not an `EditorState`) and
 * knows nothing about the input-bridge / editing. The {@link BapbongView}
 * viewer uses it directly; `BapbongEditor` composes it and feeds it a fresh doc
 * (plus a caret/selection overlay) on every transaction — so layout/paint live
 * in exactly one place.
 */
export class RenderCore {
  readonly stack: HTMLElement;
  private readonly viewport: HTMLElement | null;
  private readonly painter: CanvasPainter;
  private readonly measureText: MeasureText;
  private readonly measureMetrics: MeasureMetrics;

  private resolved: ResolvedLayout | null = null;
  private doc: ProseMirrorNode | null = null;
  private zoomFactor: number;

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
  // The document schema: model's base, or a caller-supplied extension.
  private docSchema: Schema = baseSchema;
  // Page geometry from the imported docx's sectPr (A4 until imported).
  private page: PageConfig = A4;

  // Incremental re-layout: unchanged paragraphs skip measuring on each change.
  // Replaced wholesale when late-loading fonts invalidate every measurement.
  private layoutCache = createLayoutCache();

  // Last painted overlay — reused on scroll/zoom/font-relayout repaints.
  private lastOverlay: Required<Overlay> = { caret: null, selection: [] };
  // Doc-range decorations to paint (comment tint, find highlight…). The editor
  // sets this; the core resolves the ranges to page rects at paint time.
  private decorationProvider: (() => RangeDecoration[]) | null = null;

  // Scroll repaint throttle (page virtualization).
  private scrollRaf: number | null = null;
  // Cached viewport geometry (see currentViewport): the stack's constant offset
  // within the scroll content + the viewport height, so a scroll frame reads
  // only `scrollTop` instead of forcing a reflow with getBoundingClientRect.
  private viewportMetrics: { offset: number; height: number } | null = null;
  // Latest scroll offset, captured cheaply in the scroll event (see onScroll)
  // so paints never read scrollTop off the DOM and force a reflow.
  private lastScrollTop = 0;
  // Late-font re-layout coalescing: families the current doc actually uses
  // (set at load) + a debounce timer so a burst of loadingdone events (fonts
  // land face-by-face) costs ONE re-layout, not one per event.
  private docFamilies = new Set<string>();
  private fontRelayoutTimer: ReturnType<typeof setTimeout> | null = null;
  // Subscribers notified after a late-font relayout (rects moved).
  private readonly fontsListeners = new Set<() => void>();
  // Visually-hidden ARIA mirror of the doc (screen readers; canvas is opaque).
  private readonly a11y: A11yMirror | null;

  constructor(stack: HTMLElement, opts: RenderCoreOptions = {}) {
    this.stack = stack;
    this.viewport =
      opts.viewport ?? (stack.closest('.canvas-wrap') as HTMLElement | null);
    this.painter = new CanvasPainter(stack);
    this.measureText = opts.measureText ?? createCanvasMeasurer();
    this.measureMetrics = opts.measureMetrics ?? createCanvasMetrics();
    this.zoomFactor = opts.zoom ?? 1;
    this.a11y =
      opts.a11y === false
        ? null
        : new A11yMirror(stack, { label: opts.a11yLabel });

    this.viewport?.addEventListener('scroll', this.onScroll);
    // A window resize can move the stack's offset or change the viewport height,
    // so drop the cached geometry. (Deliberately NOT a ResizeObserver on the
    // wrap: on some WebViews it fires spuriously mid-scroll, which would force a
    // getBoundingClientRect reflow on the giant stack every other frame — the
    // very cost currentViewport caches away. A stale offset from a rare non-
    // resize layout shift is harmless: it stays within the paint's page margin.)
    window.addEventListener('resize', this.invalidateViewportMetrics);
    // Fonts that finish loading later invalidate every measurement.
    document.fonts?.addEventListener?.('loadingdone', this.onFontsLoaded);
  }

  // ── State accessors ─────────────────────────────────────────────────

  /** The current read-only document, or null before the first load. */
  get document(): ProseMirrorNode | null {
    return this.doc;
  }

  /** The resolved layout for the current doc, or null. */
  get layout(): ResolvedLayout | null {
    return this.resolved;
  }

  /** Number of laid-out pages (0 before the first document). */
  get pageCount(): number {
    return this.resolved?.pages.length ?? 0;
  }

  /** The document schema in use (model's base, or a caller-supplied extension). */
  get schema(): Schema {
    return this.docSchema;
  }

  /** Resolve doc-range decorations to page rects on each content paint. */
  setDecorationProvider(fn: (() => RangeDecoration[]) | null): void {
    this.decorationProvider = fn;
  }

  // ── Load / export ───────────────────────────────────────────────────

  /**
   * Import a `.docx`, lay it out, and (by default) paint the first frame. Pass
   * `{ schema }` to import against an extended schema (e.g. plugin marks);
   * `{ paint: false }` to skip the initial paint (the editor paints with a caret
   * overlay itself). Resolves with the doc + imported page-chrome keys.
   */
  async loadDocx(
    bytes: ArrayBuffer,
    opts: { schema?: Schema; paint?: boolean } = {},
  ): Promise<{
    doc: ProseMirrorNode;
    headerKeys: string[];
    footerKeys: string[];
  }> {
    const { doc, headers, footers, footnotes, titlePg, evenAndOdd, page, raw } =
      await perf.spanAsync('importDocx', () =>
        importDocx(bytes, opts.schema ? { schema: opts.schema } : undefined),
      );
    this.docSchema = opts.schema ?? baseSchema;
    this.importedRaw = raw; // carried on export so unmodelled parts survive
    this.chromeHeaders = headers;
    this.chromeFooters = footers;
    this.chromeTitlePg = titlePg;
    this.chromeEvenAndOdd = evenAndOdd;
    this.footnotes = footnotes;
    this.page = page;
    // Measure with the real fonts, not their fallbacks.
    const families = collectFontFamilies(
      doc,
      ...Object.values(headers),
      ...Object.values(footers),
    );
    this.docFamilies = new Set(families.map(normalizeFamily));
    await perf.spanAsync('ensureFontsLoaded', () => ensureFontsLoaded(families));
    // Opening a document can change the chrome above the stack (start screen →
    // editor), shifting the stack's offset — re-measure it on the next paint.
    this.invalidateViewportMetrics();
    this.layoutDoc(doc);
    if (opts.paint !== false) this.paintContent();
    return {
      doc,
      headerKeys: Object.keys(headers),
      footerKeys: Object.keys(footers),
    };
  }

  /** Export the current document to .docx bytes, carrying the imported source
   *  package so unmodelled parts survive the round-trip. */
  async exportDocx(): Promise<Uint8Array> {
    if (!this.doc) throw new Error('RenderCore: no document loaded');
    return exportDocx(
      this.doc,
      this.importedRaw ? { carry: this.importedRaw } : undefined,
    );
  }

  // ── Layout / paint ──────────────────────────────────────────────────

  /** Lay out (or re-lay, incrementally) `doc` and store it. Does not paint. */
  layoutDoc(doc: ProseMirrorNode): void {
    this.doc = doc;
    // Keep the accessible mirror in sync (debounced internally).
    perf.span('a11y.update', () => this.a11y?.update(doc));
    this.resolved = perf.span('layout', () =>
      layout(
        doc,
        {
          page: this.page,
          measureText: this.measureText,
          measureMetrics: this.measureMetrics,
        },
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
      ),
    );
  }

  /** Full content repaint, virtualized to the scroll viewport. Pass an overlay
   *  to update the caret/selection painted on top (kept for later repaints). */
  paintContent(overlay?: Overlay): void {
    if (!this.resolved) return;
    if (overlay) this.lastOverlay = normalizeOverlay(overlay);
    const decorations = perf.span('collectDecorations', () =>
      this.collectDecorations(),
    );
    const viewport = perf.span('currentViewport', () => this.currentViewport());
    perf.span('paintContent', () =>
      this.painter.paint(this.resolved!, {
        zoom: this.zoomFactor,
        caret: this.lastOverlay.caret,
        selection: this.lastOverlay.selection,
        viewport,
        decorations,
      }),
    );
  }

  /** Redraw only the caret/selection overlay (just the affected page canvases —
   *  used for caret blink, drag preview, selection-only changes). */
  paintOverlay(overlay: Overlay): void {
    if (!this.resolved) return;
    this.lastOverlay = normalizeOverlay(overlay);
    this.painter.paintOverlay({
      caret: this.lastOverlay.caret,
      selection: this.lastOverlay.selection,
    });
  }

  /** Set the zoom factor (1 = 100%) and repaint content at the new scale. */
  setZoom(zoom: number): void {
    this.zoomFactor = zoom;
    // Zoom rescales the stack, so its offset within the scroll content shifts.
    this.invalidateViewportMetrics();
    this.paintContent();
  }

  /** The current zoom factor. */
  getZoom(): number {
    return this.zoomFactor;
  }

  /**
   * Full-resolution snapshots of every page, in order. The live canvas is
   * virtualized (only visible pages exist), so this renders **every** page
   * with a throwaway painter, untouched by the live view. `width`/`height`
   * are CSS px (layout px at zoom 1); the PNG bitmap itself is
   * devicePixelRatio-scaled for crispness.
   */
  pageSnapshots(): { png: string; width: number; height: number }[] {
    const layout = this.resolved;
    if (!layout) return [];
    const holder = document.createElement('div');
    holder.style.cssText =
      'position:absolute;left:-99999px;top:0;pointer-events:none;';
    document.body.appendChild(holder);
    try {
      // No viewport → the painter mounts a canvas for every page at full size.
      new CanvasPainter(holder).paint(layout, { zoom: 1 });
      const canvases = Array.from(holder.querySelectorAll('canvas')).sort(
        (a, b) =>
          parseFloat(a.style.top || '0') - parseFloat(b.style.top || '0'),
      );
      return canvases.map((c) => ({
        png: c.toDataURL('image/png'),
        width: parseFloat(c.style.width) || c.width,
        height: parseFloat(c.style.height) || c.height,
      }));
    } finally {
      holder.remove();
    }
  }

  /** Print the whole document — one page snapshot per sheet via a hidden
   *  iframe (clean pagination). Hosts whose webview lacks `window.print()`
   *  (WKWebView) print through their own channel instead — see the editor's
   *  `printFallback` option. */
  async print(): Promise<void> {
    const sources = this.pageSnapshots().map((p) => p.png);
    await printImages(sources);
  }

  /** Gather decorations from the provider and resolve each doc range to
   *  page-local rects the painter can fill directly. */
  private collectDecorations(): PaintDecoration[] {
    if (!this.resolved || !this.decorationProvider) return [];
    const out: PaintDecoration[] = [];
    for (const d of this.decorationProvider()) {
      const rects = selectionRects(
        this.resolved,
        d.from,
        d.to,
        this.measureText,
      );
      if (rects.length) out.push({ rects, kind: d.kind, color: d.color });
    }
    return out;
  }

  // ── Geometry queries ────────────────────────────────────────────────

  /** Caret geometry at a doc position (page-local), or null. */
  caretRect(pos: number): CaretRect | null {
    if (!this.resolved) return null;
    return perf.span('caretRect', () =>
      caretRect(this.resolved!, pos, this.measureText),
    );
  }

  /** Selection highlight rects for a doc range (page-local). */
  selectionRects(from: number, to: number): SelectionRect[] {
    if (!this.resolved) return [];
    return perf.span('selectionRects', () =>
      selectionRects(this.resolved!, from, to, this.measureText),
    );
  }

  /** The doc position one line above/below `head` at horizontal `x`, or null. */
  verticalCaret(head: number, dir: -1 | 1, x: number): number | null {
    return this.resolved
      ? verticalCaret(this.resolved, head, dir, x, this.measureText)
      : null;
  }

  /** Map client (viewport) coords to a page-local point, or null. */
  clientToPage(clientX: number, clientY: number): PagePoint | null {
    // Page canvases have pointer-events:none, so events land on the container;
    // map client coords to container-relative CSS px (offsetX/Y would be
    // relative to whichever child the pointer happens to be over).
    const rect = this.stack.getBoundingClientRect();
    return this.painter.canvasToPage(clientX - rect.left, clientY - rect.top);
  }

  /** The doc position under a page-local point, or null. */
  posAtPoint(point: PagePoint): number | null {
    return this.resolved
      ? hitTest(this.resolved, point, this.measureText)
      : null;
  }

  /** The doc position under a pointer/mouse event, or null. */
  posAtEvent(ev: { clientX: number; clientY: number }): number | null {
    const pt = this.clientToPage(ev.clientX, ev.clientY);
    return pt ? this.posAtPoint(pt) : null;
  }

  /** Map a page-local point to container (canvas-stack) coordinates, or null. */
  pageToCanvas(p: PagePoint): { x: number; y: number } | null {
    return this.painter.pageToCanvas(p);
  }

  /** Scroll the viewport so the caret at `pos` sits `topMargin` px from the top. */
  scrollToPos(pos: number, topMargin = 80): void {
    const cr = this.caretRect(pos);
    const pt =
      cr &&
      this.painter.pageToCanvas({ pageIndex: cr.pageIndex, x: cr.x, y: cr.y });
    if (pt && this.viewport) {
      this.viewport.scrollTop = Math.max(0, pt.y - topMargin);
      // Keep the scroll cache in step with a programmatic scroll (the scroll
      // event that would refresh it fires only asynchronously).
      this.lastScrollTop = this.viewport.scrollTop;
    }
  }

  /**
   * The viewport's window onto the page stack, in container CSS px.
   *
   * `top` is the scroll offset of the viewport within the stack; measured
   * directly it is `wrapRect.top − stackRect.top`, but `getBoundingClientRect`
   * on the (page-tall) stack forces a full synchronous reflow — and this runs on
   * every scroll frame, so on a large document it dominated scroll latency.
   *
   * That expression is exactly `wrap.scrollTop + C`, where `C` (the stack's
   * fixed offset within the scroll content) is constant while scrolling. We
   * measure `C` and the viewport height once, cache them, and thereafter read
   * only `wrap.scrollTop` (a cheap scroll-position read) per frame. The cache is
   * invalidated on resize/zoom, where the offset can actually change.
   */
  private currentViewport(): { top: number; height: number } | undefined {
    const wrap = this.viewport;
    if (!wrap) return undefined;
    if (!this.viewportMetrics) {
      // Rare path (first paint / resize / load): read live geometry once and
      // sync the scroll cache, so the constant offset is measured consistently.
      const wrapRect = wrap.getBoundingClientRect();
      const stackRect = this.stack.getBoundingClientRect();
      this.lastScrollTop = wrap.scrollTop;
      this.viewportMetrics = {
        offset: wrapRect.top - stackRect.top - this.lastScrollTop,
        height: wrap.clientHeight,
      };
    }
    return {
      top: this.lastScrollTop + this.viewportMetrics.offset,
      height: this.viewportMetrics.height,
    };
  }

  /** Drop the cached viewport geometry so the next paint re-measures it (after
   *  a resize/zoom/load where the stack's offset or viewport height can shift).
   *  Arrow so it can bind directly as a `resize` listener. */
  private invalidateViewportMetrics = (): void => {
    this.viewportMetrics = null;
  };

  // ── Lifecycle ───────────────────────────────────────────────────────

  /** Subscribe to late-font relayouts (page geometry changed → recompute any
   *  caret/selection rects you hold). Returns an unsubscribe fn. */
  onFontsReloaded(cb: () => void): () => void {
    this.fontsListeners.add(cb);
    return () => this.fontsListeners.delete(cb);
  }

  destroy(): void {
    if (this.scrollRaf != null) cancelAnimationFrame(this.scrollRaf);
    this.scrollRaf = null;
    if (this.fontRelayoutTimer != null) clearTimeout(this.fontRelayoutTimer);
    this.fontRelayoutTimer = null;
    this.viewport?.removeEventListener('scroll', this.onScroll);
    window.removeEventListener('resize', this.invalidateViewportMetrics);
    document.fonts?.removeEventListener?.('loadingdone', this.onFontsLoaded);
    this.fontsListeners.clear();
    this.a11y?.destroy();
  }

  /** Repaint newly visible pages while scrolling (rAF-throttled). */
  private onScroll = (): void => {
    // Read the scroll offset HERE, in the scroll event, and cache it. The event
    // fires from the compositor with the position already known, so the read is
    // cheap; reading it later inside the rAF paint would force a full reflow of
    // the page-tall stack (a canvas mounted on the prior frame left layout
    // dirty). currentViewport consumes this cache instead of touching the DOM.
    if (this.viewport)
      this.lastScrollTop = perf.span('onScroll.read', () => this.viewport!.scrollTop);
    if (this.scrollRaf != null) return;
    this.scrollRaf = requestAnimationFrame(() => {
      this.scrollRaf = null;
      this.paintContent();
      // Anchored host UI lives inside the scroll content → it scrolls natively;
      // no per-scroll repositioning needed.
    });
  };

  private onFontsLoaded = (ev: Event): void => {
    if (!this.doc) return;
    // Only faces the document actually uses invalidate measurements — UI
    // fonts and unused variants must not each trigger a full re-layout.
    // (Families set at loadDocx; a family added by later edits isn't tracked
    // yet — neither is its ensureFontsLoaded, a pre-existing gap.)
    const faces = (ev as { fontfaces?: { family: string }[] }).fontfaces ?? [];
    if (
      faces.length > 0 &&
      !faces.some((f) => this.docFamilies.has(normalizeFamily(f.family)))
    ) {
      return;
    }
    // Coalesce the burst: one re-layout after the last batch settles.
    if (this.fontRelayoutTimer != null) clearTimeout(this.fontRelayoutTimer);
    this.fontRelayoutTimer = setTimeout(() => {
      this.fontRelayoutTimer = null;
      if (!this.doc) return;
      this.layoutCache = createLayoutCache();
      this.layoutDoc(this.doc);
      this.paintContent(); // repaint content with the last overlay (rects may shift)
      // Subscribers (e.g. the editor) recompute caret/selection against the
      // new layout and repaint the overlay precisely.
      for (const cb of this.fontsListeners) cb();
    }, 150);
  };
}

/** Coalesce a partial overlay into a fully-specified one. */
function normalizeOverlay(overlay: Overlay): Required<Overlay> {
  return { caret: overlay.caret ?? null, selection: overlay.selection ?? [] };
}

/** Print page images (one per sheet) via a hidden same-origin iframe. */
function printImages(sources: string[]): Promise<void> {
  return new Promise((resolve) => {
    if (sources.length === 0) {
      resolve();
      return;
    }
    const iframe = document.createElement('iframe');
    iframe.setAttribute('aria-hidden', 'true');
    iframe.style.cssText =
      'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
    document.body.appendChild(iframe);
    const cw = iframe.contentWindow;
    const doc = iframe.contentDocument;
    if (!cw || !doc) {
      iframe.remove();
      resolve();
      return;
    }
    const body = sources.map((s) => `<img src="${s}" alt="" />`).join('');
    doc.open();
    doc.write(
      '<!doctype html><html><head><meta charset="utf-8"><style>' +
        '@page{margin:0}html,body{margin:0;padding:0}' +
        'img{display:block;width:100%;break-after:page;page-break-after:always}' +
        'img:last-child{break-after:auto;page-break-after:auto}' +
        `</style></head><body>${body}</body></html>`,
    );
    doc.close();
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      cw.focus();
      cw.print();
      // Leave the frame up briefly so the print job can spool before teardown.
      setTimeout(() => iframe.remove(), 1000);
      resolve();
    };
    const pending = Array.from(doc.images).filter((im) => !im.complete);
    if (pending.length === 0) {
      finish();
      return;
    }
    let left = pending.length;
    const onOne = () => {
      if (--left === 0) finish();
    };
    for (const im of pending) {
      im.addEventListener('load', onOne);
      im.addEventListener('error', onOne);
    }
    // Safety net if an image never fires (resolve + print what loaded).
    setTimeout(finish, 4000);
  });
}

/** Font-family names as `document.fonts` reports them: unquoted, lowercase. */
function normalizeFamily(family: string): string {
  return family
    .replace(/^["']|["']$/g, '')
    .trim()
    .toLowerCase();
}

/** Every fontFamily mark in the given documents, plus the engine default. */
export function collectFontFamilies(
  ...docs: (ProseMirrorNode | undefined)[]
): string[] {
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
