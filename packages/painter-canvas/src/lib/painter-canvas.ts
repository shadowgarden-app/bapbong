import type {
  CaretRect,
  FontSpec,
  LayoutLine,
  PagePoint,
  ResolvedChrome,
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
  /** Tint painted behind commented text (w:commentRange). */
  commentColor?: string;
  /** Visible region in container CSS px (the scroll window onto the stack).
   *  Only pages intersecting it (plus a margin) get a canvas + content; the
   *  rest are unmounted to keep memory bounded. Omit to render every page. */
  viewport?: { top: number; height: number };
}

/** Injected so the painter can be unit-tested without a DOM. */
export interface PainterDeps {
  /** Creates a blank canvas element (defaults to document.createElement). */
  createCanvas?: () => HTMLCanvasElement;
}

/** Live numbering for the page being painted (PAGE / NUMPAGES fields). */
interface PageInfo {
  page: number;
  pages: number;
}

type Required_<T> = { [K in keyof T]-?: T[K] };
type ResolvedOptions = Required_<
  Omit<PaintOptions, 'devicePixelRatio' | 'caret' | 'selection' | 'viewport'>
> &
  Pick<PaintOptions, 'caret' | 'selection' | 'viewport'>;

/** Extra band (layout px) mounted above/below the viewport so slow scrolls
 *  reveal content instead of blank page. */
const VIEWPORT_MARGIN = 200;
/** Idle page canvases kept for reuse rather than discarded on scroll. */
const POOL_LIMIT = 8;

const DEFAULTS: Omit<ResolvedOptions, 'caret' | 'selection'> = {
  zoom: 1,
  pageGap: 24,
  pageBackground: '#ffffff',
  pageBorder: '#c8c8c8',
  tableBorder: '#b0b0b0',
  textColor: '#000000',
  caretColor: '#1a1a1a',
  selectionColor: 'rgba(59, 130, 246, 0.30)',
  commentColor: 'rgba(255, 193, 7, 0.28)',
};

/** CSS font shorthand. Duplicated from bapbong-measuring: the painter may only
 *  depend on contracts (module boundary), and this must stay in sync with how
 *  text was measured. */
const fontCss = (f: FontSpec) =>
  `${f.italic ? 'italic ' : ''}${f.bold ? '700' : '400'} ${f.sizePt}pt ${f.family}`;

const defaultDpr = () =>
  typeof globalThis.devicePixelRatio === 'number' ? Math.min(globalThis.devicePixelRatio, 2) : 1;

/** One mounted page: its own `<canvas>` element + 2D context. */
interface PageSlot {
  canvas: HTMLCanvasElement;
  ctx: CanvasRenderingContext2D;
}

/**
 * Paints a ResolvedLayout into a scroll container, one `<canvas>` per page.
 *
 * A single canvas can't represent a long document: browsers cap a canvas at
 * 65535 px per side, so a tall doc (≈57+ A4 pages) silently renders blank. One
 * canvas per page keeps every backing store small, and only the pages near the
 * viewport are allocated — the rest are unmounted, so memory stays bounded no
 * matter how long the document is.
 *
 * The painter consumes pre-computed coordinates only — it never measures or
 * re-flows text. Inline images load asynchronously: the painter caches them and
 * repaints once an image becomes available.
 */
export class CanvasPainter {
  private readonly createCanvas: () => HTMLCanvasElement;
  private readonly images = new Map<string, HTMLImageElement>();
  /** pageIndex → its mounted canvas; idle canvases wait in `pool`. */
  private readonly mounted = new Map<number, PageSlot>();
  private readonly pool: PageSlot[] = [];
  /** The 2D context of the page currently being drawn (paint* read this). */
  private ctx!: CanvasRenderingContext2D;

  private lastLayout: ResolvedLayout | null = null;
  /** Last paint options, minus caret/selection (those live in lastOverlay). */
  private lastOptions: PaintOptions = {};
  private lastOverlay: { caret: CaretRect | null; selection: SelectionRect[] } = {
    caret: null,
    selection: [],
  };
  private lastFrame: { o: ResolvedOptions; dpr: number } | null = null;
  /** Top edge (layout px) of each page in the stacked document. */
  private pageY: number[] = [];
  /** Pages currently showing caret/selection (so an overlay change can clear them). */
  private overlayPages = new Set<number>();

  constructor(
    private readonly container: HTMLElement,
    deps: PainterDeps = {},
  ) {
    this.createCanvas =
      deps.createCanvas ?? (() => document.createElement('canvas'));
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
    this.lastFrame = { o, dpr };

    // Stacked geometry: page top edges + the container's scroll range.
    const width = layout.pages.reduce((m, p) => Math.max(m, p.width), 0);
    this.pageY = [];
    let acc = 0;
    for (const page of layout.pages) {
      this.pageY.push(acc);
      acc += page.height + o.pageGap;
    }
    const totalHeight = Math.max(0, acc - o.pageGap);
    this.container.style.position ||= 'relative';
    this.container.style.width = `${Math.round(width * o.zoom)}px`;
    this.container.style.height = `${Math.round(totalHeight * o.zoom)}px`;

    // Which pages to keep mounted: those intersecting the viewport (+ margin).
    const vp = options.viewport;
    const vTop = vp ? vp.top / o.zoom - VIEWPORT_MARGIN : -Infinity;
    const vBottom = vp ? (vp.top + vp.height) / o.zoom + VIEWPORT_MARGIN : Infinity;
    const desired = new Set<number>();
    layout.pages.forEach((page, i) => {
      const top = this.pageY[i];
      if (top + page.height >= vTop && top <= vBottom) desired.add(i);
    });

    for (const idx of [...this.mounted.keys()]) {
      if (!desired.has(idx)) this.unmountPage(idx);
    }
    this.overlayPages = new Set();
    for (const i of desired) this.drawPage(i, o, dpr);
    if (this.lastOverlay.caret) this.overlayPages.add(this.lastOverlay.caret.pageIndex);
    for (const r of this.lastOverlay.selection) this.overlayPages.add(r.pageIndex);
  }

  /** Redraw just the pages whose caret/selection changed — the cheap path for
   *  blink and drag (one page for a blink, a handful for a drag). */
  paintOverlay(overlay: { caret?: CaretRect | null; selection?: SelectionRect[] } = {}): void {
    this.lastOverlay = { caret: overlay.caret ?? null, selection: overlay.selection ?? [] };
    const frame = this.lastFrame;
    if (!frame || !this.lastLayout) return;
    const next = new Set<number>();
    if (this.lastOverlay.caret) next.add(this.lastOverlay.caret.pageIndex);
    for (const r of this.lastOverlay.selection) next.add(r.pageIndex);
    const affected = new Set([...this.overlayPages, ...next]);
    this.overlayPages = next;
    for (const i of affected) {
      if (this.mounted.has(i)) this.drawPage(i, frame.o, frame.dpr);
    }
  }

  /** Mount (or reuse) page `i`'s canvas, size + position it, and draw it. */
  private drawPage(i: number, o: ResolvedOptions, dpr: number): void {
    const page = this.lastLayout?.pages[i];
    if (!page) return;
    const slot = this.mountPage(i);
    this.sizeCanvas(slot.canvas, page.width, page.height, o.zoom, dpr);
    slot.canvas.style.top = `${Math.round(this.pageY[i] * o.zoom)}px`;

    this.ctx = slot.ctx;
    this.ctx.setTransform(o.zoom * dpr, 0, 0, o.zoom * dpr, 0, 0);
    this.ctx.clearRect(0, 0, page.width, page.height);
    this.ctx.textBaseline = 'alphabetic';

    const pageInfo = { page: page.index + 1, pages: this.lastLayout?.pages.length ?? 1 };
    // Each canvas is page-local, so everything draws at yOffset 0.
    this.paintPage(page, 0, o, pageInfo);
    for (const chrome of [this.chromeFor(i, 'header'), this.chromeFor(i, 'footer')]) {
      if (!chrome) continue;
      for (const line of chrome.lines) this.paintLine(line, 0, o, pageInfo);
      for (const table of chrome.tables) this.paintTable(table, 0, o, pageInfo);
    }
  }

  /** Pick the header/footer band for page `i`: the first variant on page 1 when
   *  titlePg is set, the even variant on even pages when evenAndOdd is set, else
   *  the default. A selected-but-absent variant means a blank band (no fallback
   *  to the default — that's Word's behavior for title/even pages). */
  private chromeFor(i: number, kind: 'header' | 'footer'): ResolvedChrome | undefined {
    const L = this.lastLayout;
    if (!L) return undefined;
    const s = L.chromeSelect;
    const pick = (def?: ResolvedChrome, first?: ResolvedChrome, even?: ResolvedChrome) => {
      if (s?.titlePg && i === 0) return first;
      if (s?.evenAndOdd && (i + 1) % 2 === 0) return even;
      return def;
    };
    return kind === 'header'
      ? pick(L.pageHeader, L.pageHeaderFirst, L.pageHeaderEven)
      : pick(L.pageFooter, L.pageFooterFirst, L.pageFooterEven);
  }

  /** Get page `i`'s slot, reusing a pooled canvas or creating one. */
  private mountPage(i: number): PageSlot {
    const existing = this.mounted.get(i);
    if (existing) return existing;
    let slot = this.pool.pop();
    if (!slot) {
      const canvas = this.createCanvas();
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error('bapbong-painter-canvas: 2D canvas context unavailable');
      canvas.style.position = 'absolute';
      canvas.style.left = '0';
      canvas.style.display = 'block';
      canvas.style.pointerEvents = 'none'; // pointer events go to the container
      slot = { canvas, ctx };
    }
    this.container.appendChild(slot.canvas);
    this.mounted.set(i, slot);
    return slot;
  }

  /** Detach page `i`, freeing its backing store (kept in the pool for reuse). */
  private unmountPage(i: number): void {
    const slot = this.mounted.get(i);
    if (!slot) return;
    this.mounted.delete(i);
    slot.canvas.width = 0;
    slot.canvas.height = 0; // release the (large) backing store
    slot.canvas.parentNode?.removeChild(slot.canvas);
    if (this.pool.length < POOL_LIMIT) this.pool.push(slot);
  }

  /** Only resize when needed: assigning width/height — even the same value —
   *  clears and reallocates the backing store. */
  private sizeCanvas(c: HTMLCanvasElement, width: number, height: number, zoom: number, dpr: number): void {
    const deviceW = Math.max(1, Math.round(width * zoom * dpr));
    const deviceH = Math.max(1, Math.round(height * zoom * dpr));
    if (c.width !== deviceW) c.width = deviceW;
    if (c.height !== deviceH) c.height = deviceH;
    c.style.width = `${Math.round(width * zoom)}px`;
    c.style.height = `${Math.round(height * zoom)}px`;
  }

  private paintPage(page: ResolvedPage, yOffset: number, o: ResolvedOptions, pageInfo?: PageInfo): void {
    const ctx = this.ctx;
    ctx.fillStyle = o.pageBackground;
    ctx.fillRect(0, yOffset, page.width, page.height);
    ctx.strokeStyle = o.pageBorder;
    ctx.lineWidth = 1;
    ctx.strokeRect(0.5, yOffset + 0.5, page.width - 1, page.height - 1);

    // Floating images sit behind the text (text already flows around them).
    for (const f of page.floats ?? []) {
      const el = this.requestImage(f.src);
      if (el?.complete && el.naturalWidth > 0) {
        ctx.drawImage(el, f.x, yOffset + f.y, f.width, f.height);
      }
    }

    // Selection sits under the text.
    ctx.fillStyle = o.selectionColor;
    for (const r of this.lastOverlay.selection) {
      if (r.pageIndex === page.index) ctx.fillRect(r.x, yOffset + r.y, r.width, r.height);
    }

    for (const line of page.lines) this.paintLine(line, yOffset, o, pageInfo);
    for (const table of page.tables ?? []) this.paintTable(table, yOffset, o, pageInfo);

    // Footnotes at the page bottom: a short separator rule above the bodies.
    if (page.footnotes) {
      const fn = page.footnotes;
      const startX = fn.lines[0]?.x ?? 0;
      const sepY = Math.round(yOffset + fn.separatorY) + 0.5;
      ctx.strokeStyle = o.pageBorder;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(startX, sepY);
      ctx.lineTo(startX + Math.min(192, page.width * 0.3), sepY);
      ctx.stroke();
      for (const line of fn.lines) this.paintLine(line, yOffset, o, pageInfo);
    }

    // Caret on top.
    const caret = this.lastOverlay.caret;
    if (caret && caret.pageIndex === page.index) {
      ctx.fillStyle = o.caretColor;
      ctx.fillRect(caret.x, yOffset + caret.y, 1.5, caret.height);
    }
  }

  private paintLine(
    line: LayoutLine,
    yOffset: number,
    o: ResolvedOptions,
    pageInfo?: PageInfo,
  ): void {
    const ctx = this.ctx;
    const baselineY = yOffset + line.y + line.baseline;
    // Highlight / shading + comment tint behind the text first, so glyphs sit on top.
    for (const seg of line.segments) {
      if (seg.background && seg.width) {
        ctx.fillStyle = seg.background;
        ctx.fillRect(seg.x, yOffset + line.y, seg.width, line.height);
      }
      if (seg.commentIds?.length && seg.width) {
        ctx.fillStyle = o.commentColor;
        ctx.fillRect(seg.x, yOffset + line.y, seg.width, line.height);
      }
    }
    for (const seg of line.segments) {
      ctx.font = fontCss(seg.font);
      ctx.fillStyle = seg.color ?? o.textColor;
      // Page-number fields render the live value for the page being painted.
      const text =
        seg.field && pageInfo
          ? String(seg.field === 'pageNumber' ? pageInfo.page : pageInfo.pages)
          : seg.text;
      // Super/subscript shift the (already-reduced) glyphs off the baseline.
      const em = seg.font.sizePt * (96 / 72);
      const segY =
        seg.vertAlign === 'super' ? baselineY - em * 0.5 : seg.vertAlign === 'sub' ? baselineY + em * 0.2 : baselineY;
      ctx.fillText(text, seg.x, segY);
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

  private paintTable(
    table: ResolvedTable,
    yOffset: number,
    o: ResolvedOptions,
    pageInfo?: PageInfo,
  ): void {
    const ctx = this.ctx;
    // Cell fills (w:shd) first — behind borders and content.
    for (const cell of table.cells) {
      if (cell.background) {
        ctx.fillStyle = cell.background;
        ctx.fillRect(cell.x, yOffset + cell.y, cell.width, cell.height);
      }
    }
    // OOXML tables are borderless unless w:tblBorders (or a table style) says
    // otherwise. Outer edges use top/bottom/left/right; shared edges insideH/V.
    // A cell's own w:tcBorders override its four edges.
    const b = table.borders;
    if (b || table.cells.some((c) => c.borders)) {
      ctx.strokeStyle = o.tableBorder;
      ctx.lineWidth = 1;
      ctx.beginPath();
      const edge = (x1: number, y1: number, x2: number, y2: number) => {
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
      };
      const eps = 0.5;
      for (const cell of table.cells) {
        const cb = cell.borders;
        const x0 = cell.x + 0.5;
        const x1 = cell.x + cell.width + 0.5;
        const y0 = yOffset + cell.y + 0.5;
        const y1 = yOffset + cell.y + cell.height + 0.5;
        const topOuter = Math.abs(cell.y - table.y) < eps;
        const bottomOuter = Math.abs(cell.y + cell.height - (table.y + table.height)) < eps;
        const leftOuter = Math.abs(cell.x - table.x) < eps;
        const rightOuter = Math.abs(cell.x + cell.width - (table.x + table.width)) < eps;
        // Cell border wins; else the table's outer/inside edge.
        if (cb?.top ?? (topOuter ? b?.top : b?.insideH)) edge(x0, y0, x1, y0);
        if (cb?.bottom ?? (bottomOuter ? b?.bottom : b?.insideH)) edge(x0, y1, x1, y1);
        if (cb?.left ?? (leftOuter ? b?.left : b?.insideV)) edge(x0, y0, x0, y1);
        if (cb?.right ?? (rightOuter ? b?.right : b?.insideV)) edge(x1, y0, x1, y1);
      }
      ctx.stroke();
    }
    for (const cell of table.cells) {
      for (const line of cell.lines) this.paintLine(line, yOffset, o, pageInfo);
      for (const nested of cell.tables ?? []) this.paintTable(nested, yOffset, o, pageInfo);
    }
  }

  /** Container CSS-px point → page-local point. Points in the gap between pages
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

  /** Page-local point → container CSS-px point; null before the first paint. */
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
      if (this.lastLayout) {
        this.paint(this.lastLayout, {
          ...this.lastOptions,
          caret: this.lastOverlay.caret,
          selection: this.lastOverlay.selection,
        });
      }
    };
    el.src = src;
    this.images.set(src, el);
    return el;
  }
}
