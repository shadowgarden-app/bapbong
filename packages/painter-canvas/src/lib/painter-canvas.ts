import type {
  BorderSide,
  CaretRect,
  FontSpec,
  LayoutLine,
  PagePoint,
  PaintDecoration,
  ResolvedChrome,
  ResolvedFloat,
  ResolvedLayout,
  ResolvedPage,
  ResolvedTable,
  SelectionRect,
  ShapeSpec,
} from '@shadow-garden/bapbong-contracts';
import { perf } from '@shadow-garden/bapbong-contracts';

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
  /** Plugin-contributed decorations, pre-resolved to page-local rects. Background
   *  kinds paint behind the text; underline/strike paint over it. */
  decorations?: PaintDecoration[];
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

/** Word's hyperlink blue — used for linked runs with no explicit color. */
const LINK_COLOR = '#0563c1';
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
  decorations: [],
};

/** CSS font shorthand. Duplicated from bapbong-measuring: the painter may only
 *  depend on contracts (module boundary), and this must stay in sync with how
 *  text was measured. */
const fontCss = (f: FontSpec) =>
  `${f.italic ? 'italic ' : ''}${f.bold ? '700' : '400'} ${f.sizePt}pt ${f.family}`;

const defaultDpr = () =>
  typeof globalThis.devicePixelRatio === 'number'
    ? Math.min(globalThis.devicePixelRatio, 2)
    : 1;

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
  private lastOverlay: { caret: CaretRect | null; selection: SelectionRect[] } =
    {
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
    const desired = perf.span('paint.setup', () => {
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
      const vBottom = vp
        ? (vp.top + vp.height) / o.zoom + VIEWPORT_MARGIN
        : Infinity;
      const set = new Set<number>();
      layout.pages.forEach((page, i) => {
        const top = this.pageY[i];
        if (top + page.height >= vTop && top <= vBottom) set.add(i);
      });

      for (const idx of [...this.mounted.keys()]) {
        if (!set.has(idx)) this.unmountPage(idx);
      }
      return set;
    });
    this.overlayPages = new Set();
    perf.span(`drawPages(${desired.size}/${layout.pages.length})`, () => {
      for (const i of desired) this.drawPage(i, o, dpr);
    });
    if (this.lastOverlay.caret)
      this.overlayPages.add(this.lastOverlay.caret.pageIndex);
    for (const r of this.lastOverlay.selection)
      this.overlayPages.add(r.pageIndex);
  }

  /** Redraw just the pages whose caret/selection changed — the cheap path for
   *  blink and drag (one page for a blink, a handful for a drag). */
  paintOverlay(
    overlay: { caret?: CaretRect | null; selection?: SelectionRect[] } = {},
  ): void {
    this.lastOverlay = {
      caret: overlay.caret ?? null,
      selection: overlay.selection ?? [],
    };
    const frame = this.lastFrame;
    if (!frame || !this.lastLayout) return;
    const next = new Set<number>();
    if (this.lastOverlay.caret) next.add(this.lastOverlay.caret.pageIndex);
    for (const r of this.lastOverlay.selection) next.add(r.pageIndex);
    const affected = new Set([...this.overlayPages, ...next]);
    this.overlayPages = next;
    perf.span(`paintOverlay(${affected.size}pg full-redraw)`, () => {
      for (const i of affected) {
        if (this.mounted.has(i)) this.drawPage(i, frame.o, frame.dpr);
      }
    });
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

    const pageInfo = {
      page: page.index + 1,
      pages: this.lastLayout?.pages.length ?? 1,
    };
    // Each canvas is page-local, so everything draws at yOffset 0.
    this.paintPage(page, 0, o, pageInfo);
    for (const chrome of [
      this.chromeFor(i, 'header'),
      this.chromeFor(i, 'footer'),
    ]) {
      if (!chrome) continue;
      // Watermarks live in headers as behindDoc anchors — under the chrome
      // text; the rest stays on top as before.
      for (const f of chrome.floats ?? [])
        if (f.behind) this.paintFloat(f, 0, o, pageInfo);
      for (const line of chrome.lines) this.paintLine(line, 0, o, pageInfo);
      for (const table of chrome.tables) this.paintTable(table, 0, o, pageInfo);
      for (const f of chrome.floats ?? [])
        if (!f.behind) this.paintFloat(f, 0, o, pageInfo);
    }
  }

  /** Pick the header/footer band for page `i`: the first variant on page 1 when
   *  titlePg is set, the even variant on even pages when evenAndOdd is set, else
   *  the default. A selected-but-absent variant means a blank band (no fallback
   *  to the default — that's Word's behavior for title/even pages). */
  private chromeFor(
    i: number,
    kind: 'header' | 'footer',
  ): ResolvedChrome | undefined {
    const L = this.lastLayout;
    if (!L) return undefined;
    const s = L.chromeSelect;
    // Per-section chrome: the page carries its section's set index; titlePg
    // applies on the SECTION's first page (evenAndOdd stays document-wide).
    const page = L.pages[i];
    const set =
      L.chromeSets && page?.chromeIndex != null
        ? L.chromeSets[page.chromeIndex]
        : undefined;
    if (set) {
      if (set.titlePg && page.sectionFirst) {
        return kind === 'header' ? set.headerFirst : set.footerFirst;
      }
      if (s?.evenAndOdd && (i + 1) % 2 === 0) {
        return kind === 'header' ? set.headerEven : set.footerEven;
      }
      return kind === 'header' ? set.header : set.footer;
    }
    const pick = (
      def?: ResolvedChrome,
      first?: ResolvedChrome,
      even?: ResolvedChrome,
    ) => {
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
      if (!ctx)
        throw new Error(
          'bapbong-painter-canvas: 2D canvas context unavailable',
        );
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
  private sizeCanvas(
    c: HTMLCanvasElement,
    width: number,
    height: number,
    zoom: number,
    dpr: number,
  ): void {
    const deviceW = Math.max(1, Math.round(width * zoom * dpr));
    const deviceH = Math.max(1, Math.round(height * zoom * dpr));
    if (c.width !== deviceW) c.width = deviceW;
    if (c.height !== deviceH) c.height = deviceH;
    c.style.width = `${Math.round(width * zoom)}px`;
    c.style.height = `${Math.round(height * zoom)}px`;
  }

  private paintPage(
    page: ResolvedPage,
    yOffset: number,
    o: ResolvedOptions,
    pageInfo?: PageInfo,
  ): void {
    const ctx = this.ctx;
    ctx.fillStyle = o.pageBackground;
    ctx.fillRect(0, yOffset, page.width, page.height);
    ctx.strokeStyle = o.pageBorder;
    ctx.lineWidth = 1;
    ctx.strokeRect(0.5, yOffset + 0.5, page.width - 1, page.height - 1);

    // Two-pass float painting, Word's z-order: only behindDoc drawings sit
    // under the text; everything else paints OVER it (see below). Painting
    // all floats first was wrong two ways — it hid Word's default-front
    // drawings under overlapping text, and it disagreed with the cell/chrome
    // paths, which already painted floats on top.
    for (const f of page.floats ?? [])
      if (f.behind) this.paintFloat(f, yOffset, o, pageInfo);

    // Selection sits under the text.
    ctx.fillStyle = o.selectionColor;
    for (const r of this.lastOverlay.selection) {
      if (r.pageIndex === page.index)
        ctx.fillRect(r.x, yOffset + r.y, r.width, r.height);
    }

    // Plugin background decorations (comment tint, find highlight…) behind text.
    for (const d of o.decorations) {
      if (d.kind !== 'background') continue;
      ctx.fillStyle = d.color;
      for (const r of d.rects) {
        if (r.pageIndex === page.index)
          ctx.fillRect(r.x, yOffset + r.y, r.width, r.height);
      }
    }

    // Paragraph borders (w:pBdr) under the text, over backgrounds.
    for (const b of page.paraBorders ?? []) {
      const y0 = yOffset + b.y;
      const y1 = yOffset + b.y + b.height;
      if (b.drawTop && b.borders.top)
        this.strokeBorder(b.borders.top, b.x, y0, b.x + b.width, y0);
      if (b.drawBottom && b.borders.bottom)
        this.strokeBorder(b.borders.bottom, b.x, y1, b.x + b.width, y1);
      if (b.borders.left) this.strokeBorder(b.borders.left, b.x, y0, b.x, y1);
      if (b.borders.right)
        this.strokeBorder(
          b.borders.right,
          b.x + b.width,
          y0,
          b.x + b.width,
          y1,
        );
    }

    for (const line of page.lines) this.paintLine(line, yOffset, o, pageInfo);
    for (const table of page.tables ?? [])
      this.paintTable(table, yOffset, o, pageInfo);

    // Anchored drawings without behindDoc: over the text (Word's default).
    for (const f of page.floats ?? [])
      if (!f.behind) this.paintFloat(f, yOffset, o, pageInfo);

    // Plugin underline/strike decorations over the text.
    for (const d of o.decorations) {
      if (d.kind === 'background') continue;
      ctx.fillStyle = d.color;
      for (const r of d.rects) {
        if (r.pageIndex !== page.index) continue;
        const thickness = Math.max(1, r.height * 0.06);
        const y =
          d.kind === 'underline'
            ? yOffset + r.y + r.height - thickness
            : yOffset + r.y + r.height / 2;
        ctx.fillRect(r.x, y, r.width, thickness);
      }
    }

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

  /** Run `draw` rotated `deg` clockwise around the center of the given box —
   *  the paint-only rotation of images/shapes (the layout box stays put). */
  private withRotation(
    deg: number | undefined,
    x: number,
    y: number,
    w: number,
    h: number,
    draw: () => void,
  ): void {
    if (!deg) {
      draw();
      return;
    }
    const ctx = this.ctx;
    ctx.save();
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate((deg * Math.PI) / 180);
    ctx.translate(-(x + w / 2), -(y + h / 2));
    draw();
    ctx.restore();
  }

  /** One anchored float: vector shape or bitmap, then any textbox text laid
   *  out inside it (box-local lines — translate to the float's origin, and
   *  clip to the box the way Word hides textbox overflow). */
  private paintFloat(
    f: ResolvedFloat,
    yOffset: number,
    o: ResolvedOptions,
    pageInfo?: PageInfo,
  ): void {
    this.withRotation(f.rotation, f.x, yOffset + f.y, f.width, f.height, () => {
      if (f.shape) {
        this.drawShape(f.shape, f.x, yOffset + f.y, f.width, f.height);
      } else {
        const el = this.requestImage(f.src);
        if (el?.complete && el.naturalWidth > 0) {
          this.ctx.drawImage(el, f.x, yOffset + f.y, f.width, f.height);
        }
      }
      if (f.lines && f.lines.length > 0) {
        const ctx = this.ctx;
        ctx.save();
        ctx.beginPath();
        ctx.rect(f.x, yOffset + f.y, f.width, f.height);
        ctx.clip();
        ctx.translate(f.x, yOffset + f.y);
        for (const line of f.lines) this.paintLine(line, 0, o, pageInfo);
        ctx.restore();
      }
    });
  }

  private paintLine(
    line: LayoutLine,
    yOffset: number,
    o: ResolvedOptions,
    pageInfo?: PageInfo,
  ): void {
    const ctx = this.ctx;
    const baselineY = yOffset + line.y + line.baseline;
    // Segment highlight / shading behind the text first, so glyphs sit on top.
    // (Comment tint is no longer special-cased here — it arrives as a generic
    // plugin decoration painted in paintPage.)
    for (const seg of line.segments) {
      if (seg.background && seg.width) {
        ctx.fillStyle = seg.background;
        ctx.fillRect(seg.x, yOffset + line.y, seg.width, line.height);
      }
    }
    for (const seg of line.segments) {
      ctx.font = fontCss(seg.font);
      // Hyperlinks without an explicit color get Word's hyperlink look
      // (blue + underline) — otherwise a fresh link paints like plain text
      // and inserting one reads as "nothing happened".
      ctx.fillStyle = seg.color ?? (seg.link ? LINK_COLOR : o.textColor);
      // Page-number fields render the live value for the page being painted.
      const text =
        seg.field && pageInfo
          ? String(seg.field === 'pageNumber' ? pageInfo.page : pageInfo.pages)
          : seg.text;
      // Super/subscript shift the (already-reduced) glyphs off the baseline.
      const em = seg.font.sizePt * (96 / 72);
      const segY =
        seg.vertAlign === 'super'
          ? baselineY - em * 0.5
          : seg.vertAlign === 'sub'
            ? baselineY + em * 0.2
            : baselineY;
      ctx.fillText(text, seg.x, segY);
      // Text decorations use the width measured at layout time — the painter
      // never measures.
      const underline = seg.underline || (!!seg.link && !seg.color);
      if ((underline || seg.strike) && seg.width) {
        const em = seg.font.sizePt * (96 / 72);
        const thickness = Math.max(1, em * 0.05);
        if (underline)
          ctx.fillRect(
            seg.x,
            baselineY + Math.max(1, em * 0.1),
            seg.width,
            thickness,
          );
        if (seg.strike)
          ctx.fillRect(seg.x, baselineY - em * 0.27, seg.width, thickness);
      }
    }
    for (const img of line.images ?? []) {
      this.withRotation(
        img.rotation,
        img.x,
        baselineY - img.height,
        img.width,
        img.height,
        () => {
          if (img.shape) {
            // Same box the bitmap would occupy: bottom edge on the baseline.
            this.drawShape(
              img.shape,
              img.x,
              baselineY - img.height,
              img.width,
              img.height,
            );
            return;
          }
          const el = this.requestImage(img.src);
          if (el?.complete && el.naturalWidth > 0) {
            // The image's bottom edge sits on the baseline (matches the layout).
            ctx.drawImage(
              el,
              img.x,
              baselineY - img.height,
              img.width,
              img.height,
            );
          }
        },
      );
    }
  }

  /** Vector shape in an image box, per ShapeSpec.kind. The path is built with
   *  primitive calls (no Path2D) and filled then stroked; strokes stay inside
   *  the box so thick outlines don't bleed into text. */
  private drawShape(
    s: ShapeSpec,
    x: number,
    y: number,
    w: number,
    h: number,
  ): void {
    const ctx = this.ctx;
    const lw = s.strokeWidth || 1;
    const fillStroke = () => {
      if (s.fill) {
        ctx.fillStyle = s.fill;
        ctx.fill();
      }
      if (s.stroke) {
        ctx.strokeStyle = s.stroke;
        ctx.lineWidth = lw;
        ctx.stroke();
      }
    };
    switch (s.kind) {
      case 'rect': {
        if (s.fill) {
          ctx.fillStyle = s.fill;
          ctx.fillRect(x, y, w, h);
        }
        if (s.stroke) {
          ctx.strokeStyle = s.stroke;
          ctx.lineWidth = lw;
          ctx.strokeRect(
            x + lw / 2,
            y + lw / 2,
            Math.max(0, w - lw),
            Math.max(0, h - lw),
          );
        }
        return;
      }
      case 'line': {
        if (!s.stroke) return;
        ctx.strokeStyle = s.stroke;
        ctx.lineWidth = lw;
        ctx.beginPath();
        if (s.flipV) {
          ctx.moveTo(x, y + h);
          ctx.lineTo(x + w, y);
        } else {
          ctx.moveTo(x, y);
          ctx.lineTo(x + w, y + h);
        }
        ctx.stroke();
        return;
      }
      case 'ellipse': {
        ctx.beginPath();
        ctx.ellipse(
          x + w / 2,
          y + h / 2,
          Math.max(0, (w - lw) / 2),
          Math.max(0, (h - lw) / 2),
          0,
          0,
          Math.PI * 2,
        );
        fillStroke();
        return;
      }
      case 'roundRect': {
        // OOXML default corner adj 16667/100000 of the shorter side.
        const r = Math.min(0.16667 * Math.min(w, h), w / 2, h / 2);
        const [x0, y0, x1, y1] = [
          x + lw / 2,
          y + lw / 2,
          x + w - lw / 2,
          y + h - lw / 2,
        ];
        ctx.beginPath();
        ctx.moveTo(x0 + r, y0);
        ctx.lineTo(x1 - r, y0);
        ctx.quadraticCurveTo(x1, y0, x1, y0 + r);
        ctx.lineTo(x1, y1 - r);
        ctx.quadraticCurveTo(x1, y1, x1 - r, y1);
        ctx.lineTo(x0 + r, y1);
        ctx.quadraticCurveTo(x0, y1, x0, y1 - r);
        ctx.lineTo(x0, y0 + r);
        ctx.quadraticCurveTo(x0, y0, x0 + r, y0);
        ctx.closePath();
        fillStroke();
        return;
      }
      case 'rightArrow': {
        // Block arrow, OOXML defaults: shaft height h/2, head length half the
        // shorter side.
        const head = Math.min(0.5 * Math.min(w, h), w);
        const hx = x + w - head;
        ctx.beginPath();
        ctx.moveTo(x, y + h / 4);
        ctx.lineTo(hx, y + h / 4);
        ctx.lineTo(hx, y);
        ctx.lineTo(x + w, y + h / 2);
        ctx.lineTo(hx, y + h);
        ctx.lineTo(hx, y + (3 * h) / 4);
        ctx.lineTo(x, y + (3 * h) / 4);
        ctx.closePath();
        fillStroke();
        return;
      }
      case 'horizontalScroll': {
        // Stylized banner: paper band between two vertical rolled ends
        // (full-height ellipses), roll radius per the preset's default adj.
        const r = Math.min(0.125 * Math.min(w, h), w / 4);
        ctx.beginPath();
        ctx.rect(x + r, y + lw / 2, w - 2 * r, h - lw);
        fillStroke();
        for (const cx of [x + r, x + w - r]) {
          ctx.beginPath();
          ctx.ellipse(
            cx,
            y + h / 2,
            r,
            Math.max(0, (h - lw) / 2),
            0,
            0,
            Math.PI * 2,
          );
          fillStroke();
        }
        return;
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
    // Cell fills (w:shd) first — behind everything.
    for (const cell of table.cells) {
      if (cell.background) {
        ctx.fillStyle = cell.background;
        ctx.fillRect(cell.x, yOffset + cell.y, cell.width, cell.height);
      }
    }
    // behindDoc drawings in cells: under the cell text, over the fills.
    for (const cell of table.cells)
      for (const f of cell.floats ?? [])
        if (f.behind) this.paintFloat(f, yOffset, o, pageInfo);
    // Cell content next: run shading/highlight boxes can reach the cell edge,
    // so borders must stroke AFTER them (Word paints grid lines on top).
    for (const cell of table.cells) {
      // A rotated image's box can poke past the cell horizontally (the row
      // only grew vertically) — clip the cell like overflow:hidden so it
      // doesn't paint over neighbouring columns.
      const clip = cell.lines.some((l) => l.images?.some((im) => im.rotation));
      if (clip) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(cell.x, yOffset + cell.y, cell.width, cell.height);
        ctx.clip();
      }
      for (const line of cell.lines) this.paintLine(line, yOffset, o, pageInfo);
      for (const nested of cell.tables ?? [])
        this.paintTable(nested, yOffset, o, pageInfo);
      if (clip) ctx.restore();
    }
    // OOXML tables are borderless unless w:tblBorders (or a table style) says
    // otherwise. Outer edges use top/bottom/left/right; shared edges insideH/V.
    // A cell's own w:tcBorders override its four edges.
    const b = table.borders;
    if (b || table.cells.some((c) => c.borders)) {
      const eps = 0.5;
      for (const cell of table.cells) {
        const cb = cell.borders;
        const x0 = cell.x + 0.5;
        const x1 = cell.x + cell.width + 0.5;
        const y0 = yOffset + cell.y + 0.5;
        const y1 = yOffset + cell.y + cell.height + 0.5;
        const topOuter = Math.abs(cell.y - table.y) < eps;
        const bottomOuter =
          Math.abs(cell.y + cell.height - (table.y + table.height)) < eps;
        const leftOuter = Math.abs(cell.x - table.x) < eps;
        const rightOuter =
          Math.abs(cell.x + cell.width - (table.x + table.width)) < eps;
        // Cell border wins; else the table's outer/inside edge. A BorderSide is
        // visible, `false` is explicit-none, absent inherits.
        const top = cb?.top ?? (topOuter ? b?.top : b?.insideH);
        const bottom = cb?.bottom ?? (bottomOuter ? b?.bottom : b?.insideH);
        const left = cb?.left ?? (leftOuter ? b?.left : b?.insideV);
        const right = cb?.right ?? (rightOuter ? b?.right : b?.insideV);
        if (top) this.strokeBorder(top, x0, y0, x1, y0);
        if (bottom) this.strokeBorder(bottom, x0, y1, x1, y1);
        if (left) this.strokeBorder(left, x0, y0, x0, y1);
        if (right) this.strokeBorder(right, x1, y0, x1, y1);
      }
      ctx.setLineDash([]);
    }
    // Anchored images/shapes positioned in the cells, on top (Word paints
    // non-behindDoc drawings over the table grid).
    for (const cell of table.cells) {
      for (const f of cell.floats ?? [])
        if (!f.behind) this.paintFloat(f, yOffset, o, pageInfo);
    }
  }

  /** Stroke one border edge with its width / style / colour. The edge is always
   *  horizontal or vertical (table grid lines). */
  private strokeBorder(
    side: BorderSide,
    x1: number,
    y1: number,
    x2: number,
    y2: number,
  ): void {
    const ctx = this.ctx;
    ctx.strokeStyle = side.color;
    if (side.style === 'double') {
      // Two thin parallel lines straddling the edge.
      ctx.setLineDash([]);
      ctx.lineWidth = Math.max(0.75, side.width / 3);
      const gap = Math.max(2, side.width) / 2;
      const horiz = y1 === y2;
      ctx.beginPath();
      ctx.moveTo(x1 + (horiz ? 0 : -gap), y1 + (horiz ? -gap : 0));
      ctx.lineTo(x2 + (horiz ? 0 : -gap), y2 + (horiz ? -gap : 0));
      ctx.moveTo(x1 + (horiz ? 0 : gap), y1 + (horiz ? gap : 0));
      ctx.lineTo(x2 + (horiz ? 0 : gap), y2 + (horiz ? gap : 0));
      ctx.stroke();
      return;
    }
    ctx.lineWidth = side.width;
    ctx.setLineDash(
      side.style === 'dashed'
        ? [side.width * 3, side.width * 2]
        : side.style === 'dotted'
          ? [side.width, side.width * 1.6]
          : [],
    );
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
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
        return {
          pageIndex: i,
          x,
          y: Math.min(Math.max(y - yOffset, 0), page.height),
        };
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
    const t0 = perf.now();
    el.onload = () => {
      perf.log(
        `image.load(${(src.length / 1024).toFixed(0)}KB src)`,
        perf.now() - t0,
      );
      if (this.lastLayout) {
        perf.span('image.onload-repaint', () =>
          this.paint(this.lastLayout!, {
            ...this.lastOptions,
            caret: this.lastOverlay.caret,
            selection: this.lastOverlay.selection,
          }),
        );
      }
    };
    el.src = src;
    this.images.set(src, el);
    return el;
  }
}
