import type {
  CaretRect,
  FontSpec,
  LayoutLine,
  PagePoint,
  ResolvedLayout,
  ResolvedPage,
  ResolvedTable,
  SelectionRect,
} from '@shadow-garden/bapbong-contracts';

export interface PaintOptions {
  /** Zoom factor (1 = 100%). */
  zoom?: number;
  /** Vertical gap between pages, in layout px. */
  pageGap?: number;
  /** Device pixel ratio override. Defaults to `window.devicePixelRatio`, capped at 2. */
  devicePixelRatio?: number;
  pageBackground?: string;
  pageBorder?: string;
  tableBorder?: string;
  /** Fallback text color for segments without an explicit color. */
  textColor?: string;
  /** Caret to draw (page-local coords), e.g. from bapbong-selection. */
  caret?: CaretRect | null;
  /** Selection highlight rects (page-local coords), drawn under the text. */
  selection?: SelectionRect[];
  caretColor?: string;
  selectionColor?: string;
  /** Visible region of the canvas in CSS px (e.g. the scroll container's
   *  window onto it). Pages outside it paint only their background — call
   *  paint() again on scroll. Omit to paint every page in full. */
  viewport?: { top: number; height: number };
}

type Required_<T> = { [K in keyof T]-?: T[K] };
type ResolvedOptions = Required_<
  Omit<PaintOptions, 'devicePixelRatio' | 'caret' | 'selection' | 'viewport'>
> &
  Pick<PaintOptions, 'caret' | 'selection' | 'viewport'>;

/** Extra band (layout px) painted above/below the viewport so slow scrolls
 *  reveal content instead of blank page. */
const VIEWPORT_MARGIN = 200;

const DEFAULTS: Omit<ResolvedOptions, 'caret' | 'selection'> = {
  zoom: 1,
  pageGap: 24,
  pageBackground: '#ffffff',
  pageBorder: '#c8c8c8',
  tableBorder: '#b0b0b0',
  textColor: '#000000',
  caretColor: '#1a1a1a',
  selectionColor: 'rgba(59, 130, 246, 0.30)',
};

/** CSS font shorthand. Duplicated from bapbong-measuring: the painter may only
 *  depend on contracts (module boundary), and this must stay in sync with how
 *  text was measured. */
const fontCss = (f: FontSpec) =>
  `${f.italic ? 'italic ' : ''}${f.bold ? '700' : '400'} ${f.sizePt}pt ${f.family}`;

const defaultDpr = () =>
  typeof globalThis.devicePixelRatio === 'number' ? Math.min(globalThis.devicePixelRatio, 2) : 1;

/**
 * Paints a ResolvedLayout onto a `<canvas>`. Pages are stacked vertically.
 *
 * The painter consumes pre-computed coordinates only — it never measures or
 * re-flows text. Inline images load asynchronously: the painter caches them
 * and repaints the last layout once an image becomes available.
 */
export class CanvasPainter {
  private readonly ctx: CanvasRenderingContext2D;
  /** Optional second canvas stacked on top: caret + selection live here so
   *  drag/blink redraws never re-rasterize the document text. */
  private readonly overlayCtx: CanvasRenderingContext2D | null;
  private readonly images = new Map<string, HTMLImageElement>();
  private lastLayout: ResolvedLayout | null = null;
  /** Last paint options, minus caret/selection (those live in lastOverlay). */
  private lastOptions: PaintOptions = {};
  private lastOverlay: { caret: CaretRect | null; selection: SelectionRect[] } = {
    caret: null,
    selection: [],
  };
  private lastFrame: { o: ResolvedOptions; dpr: number; width: number; height: number } | null =
    null;

  constructor(
    private readonly canvas: HTMLCanvasElement,
    private readonly overlayCanvas?: HTMLCanvasElement,
  ) {
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('bapbong-painter-canvas: 2D canvas context unavailable');
    this.ctx = ctx;
    this.overlayCtx = overlayCanvas?.getContext('2d') ?? null;
    if (overlayCanvas && !this.overlayCtx) {
      throw new Error('bapbong-painter-canvas: 2D overlay context unavailable');
    }
  }

  paint(layout: ResolvedLayout, options: PaintOptions = {}): void {
    this.lastLayout = layout;
    const { caret, selection, ...rest } = options;
    this.lastOptions = rest;
    if (caret !== undefined || selection !== undefined) {
      this.lastOverlay = { caret: caret ?? null, selection: selection ?? [] };
    }
    const o: ResolvedOptions = { ...DEFAULTS, ...rest };
    const dpr = options.devicePixelRatio ?? defaultDpr();

    const width = layout.pages.reduce((m, p) => Math.max(m, p.width), 0);
    const height =
      layout.pages.reduce((s, p) => s + p.height, 0) +
      Math.max(0, layout.pages.length - 1) * o.pageGap;
    this.lastFrame = { o, dpr, width, height };

    this.sizeCanvas(this.canvas, width, height, o.zoom, dpr);
    if (this.overlayCanvas) this.sizeCanvas(this.overlayCanvas, width, height, o.zoom, dpr);

    const ctx = this.ctx;
    ctx.setTransform(o.zoom * dpr, 0, 0, o.zoom * dpr, 0, 0);
    ctx.clearRect(0, 0, width, height);
    ctx.textBaseline = 'alphabetic';

    // Page virtualization: only pages intersecting the viewport get their
    // content drawn (backgrounds always paint, so scrolling shows pages).
    const vp = options.viewport;
    const vTop = vp ? vp.top / o.zoom - VIEWPORT_MARGIN : -Infinity;
    const vBottom = vp ? (vp.top + vp.height) / o.zoom + VIEWPORT_MARGIN : Infinity;

    let yOffset = 0;
    for (const page of layout.pages) {
      const contentVisible = yOffset + page.height >= vTop && yOffset <= vBottom;
      this.paintPage(page, yOffset, o, this.overlayCtx == null, contentVisible);
      if (contentVisible) {
        // Page chrome (header/footer) repeats on every page.
        for (const chrome of [layout.pageHeader, layout.pageFooter]) {
          if (!chrome) continue;
          for (const line of chrome.lines) this.paintLine(line, yOffset, o);
          for (const table of chrome.tables) this.paintTable(table, yOffset, o);
        }
      }
      yOffset += page.height + o.pageGap;
    }
    if (this.overlayCtx) this.renderOverlay();
  }

  /** Redraw only the caret/selection layer — the cheap path for drag and
   *  blink. Without an overlay canvas this falls back to a full repaint. */
  paintOverlay(overlay: { caret?: CaretRect | null; selection?: SelectionRect[] } = {}): void {
    this.lastOverlay = { caret: overlay.caret ?? null, selection: overlay.selection ?? [] };
    if (!this.overlayCtx) {
      if (this.lastLayout) {
        this.paint(this.lastLayout, {
          ...this.lastOptions,
          caret: this.lastOverlay.caret,
          selection: this.lastOverlay.selection,
        });
      }
      return;
    }
    this.renderOverlay();
  }

  /** Only resize when needed: assigning width/height — even the same value —
   *  clears and reallocates the backing store (expensive on a multi-page canvas). */
  private sizeCanvas(c: HTMLCanvasElement, width: number, height: number, zoom: number, dpr: number): void {
    const deviceW = Math.max(1, Math.round(width * zoom * dpr));
    const deviceH = Math.max(1, Math.round(height * zoom * dpr));
    if (c.width !== deviceW) c.width = deviceW;
    if (c.height !== deviceH) c.height = deviceH;
    c.style.width = `${Math.round(width * zoom)}px`;
    c.style.height = `${Math.round(height * zoom)}px`;
  }

  /** Clear + redraw the overlay canvas from lastOverlay. */
  private renderOverlay(): void {
    const octx = this.overlayCtx;
    const frame = this.lastFrame;
    if (!octx || !frame || !this.lastLayout) return;
    const { o, dpr, width, height } = frame;
    octx.setTransform(o.zoom * dpr, 0, 0, o.zoom * dpr, 0, 0);
    octx.clearRect(0, 0, width, height);
    let yOffset = 0;
    for (const page of this.lastLayout.pages) {
      this.paintPageOverlay(octx, page.index, yOffset, o);
      yOffset += page.height + o.pageGap;
    }
  }

  /** Selection rects + caret for one page onto `ctx` (page-stacking applied). */
  private paintPageOverlay(
    ctx: CanvasRenderingContext2D,
    pageIndex: number,
    yOffset: number,
    o: ResolvedOptions,
  ): void {
    ctx.fillStyle = o.selectionColor;
    for (const r of this.lastOverlay.selection) {
      if (r.pageIndex === pageIndex) ctx.fillRect(r.x, yOffset + r.y, r.width, r.height);
    }
    const caret = this.lastOverlay.caret;
    if (caret && caret.pageIndex === pageIndex) {
      ctx.fillStyle = o.caretColor;
      ctx.fillRect(caret.x, yOffset + caret.y, 1.5, caret.height);
    }
  }

  private paintPage(
    page: ResolvedPage,
    yOffset: number,
    o: ResolvedOptions,
    inlineOverlay: boolean,
    contentVisible: boolean,
  ): void {
    const ctx = this.ctx;
    ctx.fillStyle = o.pageBackground;
    ctx.fillRect(0, yOffset, page.width, page.height);
    ctx.strokeStyle = o.pageBorder;
    ctx.lineWidth = 1;
    ctx.strokeRect(0.5, yOffset + 0.5, page.width - 1, page.height - 1);
    if (!contentVisible) return; // virtualized away — background only

    // Single-canvas mode: selection sits under the text.
    if (inlineOverlay) {
      ctx.fillStyle = o.selectionColor;
      for (const r of this.lastOverlay.selection) {
        if (r.pageIndex === page.index) ctx.fillRect(r.x, yOffset + r.y, r.width, r.height);
      }
    }

    for (const line of page.lines) this.paintLine(line, yOffset, o);
    for (const table of page.tables ?? []) this.paintTable(table, yOffset, o);

    // Single-canvas mode: caret on top.
    if (inlineOverlay) {
      const caret = this.lastOverlay.caret;
      if (caret && caret.pageIndex === page.index) {
        ctx.fillStyle = o.caretColor;
        ctx.fillRect(caret.x, yOffset + caret.y, 1.5, caret.height);
      }
    }
  }

  private paintLine(line: LayoutLine, yOffset: number, o: ResolvedOptions): void {
    const ctx = this.ctx;
    const baselineY = yOffset + line.y + line.baseline;
    for (const seg of line.segments) {
      ctx.font = fontCss(seg.font);
      ctx.fillStyle = seg.color ?? o.textColor;
      ctx.fillText(seg.text, seg.x, baselineY);
      // Text decorations use the width measured at layout time — the painter
      // never measures.
      if ((seg.underline || seg.strike) && seg.width) {
        const em = seg.font.sizePt * (96 / 72);
        const thickness = Math.max(1, em * 0.05);
        if (seg.underline) ctx.fillRect(seg.x, baselineY + Math.max(1, em * 0.1), seg.width, thickness);
        if (seg.strike) ctx.fillRect(seg.x, baselineY - em * 0.27, seg.width, thickness);
      }
    }
    for (const img of line.images ?? []) {
      const el = this.requestImage(img.src);
      if (el?.complete && el.naturalWidth > 0) {
        // The image's bottom edge sits on the baseline (matches the layout).
        ctx.drawImage(el, img.x, baselineY - img.height, img.width, img.height);
      }
    }
  }

  private paintTable(table: ResolvedTable, yOffset: number, o: ResolvedOptions): void {
    const ctx = this.ctx;
    ctx.strokeStyle = o.tableBorder;
    ctx.lineWidth = 1;
    for (const cell of table.cells) {
      ctx.strokeRect(cell.x + 0.5, yOffset + cell.y + 0.5, cell.width, cell.height);
    }
    for (const cell of table.cells) {
      for (const line of cell.lines) this.paintLine(line, yOffset, o);
      for (const nested of cell.tables ?? []) this.paintTable(nested, yOffset, o);
    }
  }

  /** Canvas CSS-px point → page-local point. Points in the gap between pages
   *  clamp to the nearer page edge; null before the first paint. */
  canvasToPage(cssX: number, cssY: number): PagePoint | null {
    if (!this.lastLayout) return null;
    const zoom = this.lastOptions.zoom ?? DEFAULTS.zoom;
    const gap = this.lastOptions.pageGap ?? DEFAULTS.pageGap;
    const x = cssX / zoom;
    const y = cssY / zoom;
    let yOffset = 0;
    const pages = this.lastLayout.pages;
    for (let i = 0; i < pages.length; i++) {
      const page = pages[i];
      const isLast = i === pages.length - 1;
      // Claim half the trailing gap so clicks between pages snap sensibly.
      const claim = page.height + (isLast ? Infinity : gap / 2);
      if (y < yOffset + claim) {
        return { pageIndex: i, x, y: Math.min(Math.max(y - yOffset, 0), page.height) };
      }
      yOffset += page.height + gap;
    }
    return null;
  }

  /** Page-local point → canvas CSS-px point; null before the first paint. */
  pageToCanvas(point: PagePoint): { x: number; y: number } | null {
    if (!this.lastLayout) return null;
    const zoom = this.lastOptions.zoom ?? DEFAULTS.zoom;
    const gap = this.lastOptions.pageGap ?? DEFAULTS.pageGap;
    let yOffset = 0;
    for (let i = 0; i < point.pageIndex; i++) {
      const page = this.lastLayout.pages[i];
      if (!page) return null;
      yOffset += page.height + gap;
    }
    return { x: point.x * zoom, y: (yOffset + point.y) * zoom };
  }

  /** Return the cached image for `src`, kicking off a load (and a repaint on
   *  completion) the first time. Returns undefined where Image is unavailable
   *  (SSR / tests) — the image is simply skipped. */
  private requestImage(src: string): HTMLImageElement | undefined {
    const cached = this.images.get(src);
    if (cached) return cached;
    if (typeof Image === 'undefined') return undefined;
    const el = new Image();
    el.onload = () => {
      if (this.lastLayout) this.paint(this.lastLayout, this.lastOptions);
    };
    el.src = src;
    this.images.set(src, el);
    return el;
  }
}
