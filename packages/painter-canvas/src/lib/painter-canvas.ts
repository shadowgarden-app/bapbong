import { applyGlyphSpec } from '@shadow-garden/bapbong-contracts';
import type {
  BorderSide,
  CaretRect,
  ImageCrop,
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

/** Live numbering for the page being painted (PAGE / NUMPAGES fields).
 *  `page` is the DISPLAY number — w:pgNumType restart/format applied
 *  (layout.pageLabels), so it may be "ii" rather than a digit string. */
interface PageInfo {
  page: string;
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

/**
 * A shape's dash pattern (multiples of the stroke width) → canvas px.
 *
 * The cap adjustment is an INFERENCE, not something either spec spells out:
 * canvas — like PDF — grows every dash by half the stroke width at each end
 * once the cap is round or square, so a "1 1" dotted connector would come out
 * a solid line. Word plainly shows those as separated dots, so the cap must
 * be drawn INSIDE the dash length. Shortening each dash by one stroke width
 * (and lengthening each gap to match) reproduces that: "1 1" round becomes a
 * zero-length dash — a dot — spaced two widths apart. If a document ever
 * renders wrong because of this, it is this assumption to revisit first.
 */
function dashArray(
  pattern: readonly number[],
  lw: number,
  cap: 'flat' | 'square' | 'round' | undefined,
): number[] {
  const round = cap === 'round' || cap === 'square';
  return pattern.map((n, i) => {
    const len = n * lw;
    // Even slots are dashes, odd are gaps.
    if (!round) return len;
    return i % 2 === 0 ? Math.max(0, len - lw) : len + lw;
  });
}

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
  /** Sources the engine could not decode (a vector metafile, a broken data
   *  URL, an unreachable link). Painted as a placeholder box — the layout
   *  reserved the space, and an empty gap where the document has a picture
   *  reads as "nothing here" rather than "something we cannot show". */
  private readonly undecodable = new Set<string>();
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
      page: this.lastLayout?.pageLabels?.[page.index] ?? String(page.index + 1),
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

    // Paragraph boxes (w:shd fill, then w:pBdr sides) under the text, over
    // backgrounds. The fill covers every fragment of a split paragraph; the
    // horizontal rules only close the box at its real top and bottom.
    for (const b of page.paraBoxes ?? []) {
      const y0 = yOffset + b.y;
      const y1 = yOffset + b.y + b.height;
      if (b.shading) {
        ctx.fillStyle = b.shading;
        ctx.fillRect(b.x, y0, b.width, b.height);
      }
      const bd = b.borders;
      if (!bd) continue;
      if (b.drawTop && bd.top)
        this.strokeBorder(bd.top, b.x, y0, b.x + b.width, y0);
      if (b.drawBottom && bd.bottom)
        this.strokeBorder(bd.bottom, b.x, y1, b.x + b.width, y1);
      if (bd.left) this.strokeBorder(bd.left, b.x, y0, b.x, y1);
      if (bd.right)
        this.strokeBorder(bd.right, b.x + b.width, y0, b.x + b.width, y1);
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
          this.drawBitmap(el, f.crop, f.x, yOffset + f.y, f.width, f.height);
        } else if (this.undecodable.has(f.src)) {
          this.drawPlaceholder(f.x, yOffset + f.y, f.width, f.height);
        }
        // Word's picture border sits ON the box edge, over the bitmap.
        if (f.outline)
          this.strokeBox(f.outline, f.x, yOffset + f.y, f.width, f.height);
      }
      const lines = f.lines ?? [];
      const tables = f.tables ?? [];
      if (lines.length > 0 || tables.length > 0) {
        const ctx = this.ctx;
        ctx.save();
        ctx.beginPath();
        ctx.rect(f.x, yOffset + f.y, f.width, f.height);
        ctx.clip();
        ctx.translate(f.x, yOffset + f.y);
        // Same order as a page: text, then tables over it.
        for (const line of lines) this.paintLine(line, 0, o, pageInfo);
        for (const table of tables) this.paintTable(table, 0, o, pageInfo);
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
      const scaleX = applyGlyphSpec(ctx, seg.font);
      // Hyperlinks without an explicit color get Word's hyperlink look
      // (blue + underline) — otherwise a fresh link paints like plain text
      // and inserting one reads as "nothing happened".
      ctx.fillStyle = seg.color ?? (seg.link ? LINK_COLOR : o.textColor);
      // Page-number fields render the live value for the page being painted.
      const text =
        seg.field && pageInfo
          ? seg.field === 'pageNumber'
            ? pageInfo.page
            : String(pageInfo.pages)
          : seg.text;
      // Super/subscript shift the (already-reduced) glyphs off the baseline;
      // w:position adds its own shift on top, at full size.
      const em = seg.font.sizePt * (96 / 72);
      const segY =
        (seg.vertAlign === 'super'
          ? baselineY - em * 0.5
          : seg.vertAlign === 'sub'
            ? baselineY + em * 0.2
            : baselineY) - (seg.raise ?? 0);
      if (scaleX === 1) {
        ctx.fillText(text, seg.x, segY);
      } else {
        // Horizontal glyph scaling (w:w). Only the GLYPHS are squeezed, so
        // the transform wraps nothing but the fillText — seg.width already
        // has the scale baked in by the measurer, and the decorations below
        // are drawn from it in unscaled space. Putting them inside here
        // would apply the scale a second time.
        ctx.save();
        ctx.scale(scaleX, 1);
        ctx.fillText(text, seg.x / scaleX, segY);
        ctx.restore();
      }
      // Text decorations use the width measured at layout time — the painter
      // never measures.
      const underline = seg.underline || (!!seg.link && !seg.color);
      if ((underline || seg.strike || seg.dstrike) && seg.width) {
        const em = seg.font.sizePt * (96 / 72);
        const thickness = Math.max(1, em * 0.05);
        // Decorations ride with a raised run — a rule left at the original
        // baseline would detach from the glyphs it belongs to. (Super/sub
        // deliberately keep drawing theirs at the base line, as before.)
        const decoY = baselineY - (seg.raise ?? 0);
        if (underline)
          ctx.fillRect(
            seg.x,
            decoY + Math.max(1, em * 0.1),
            seg.width,
            thickness,
          );
        if (seg.strike)
          ctx.fillRect(seg.x, decoY - em * 0.27, seg.width, thickness);
        if (seg.dstrike) {
          // Double strikethrough: two thin lines straddling the single-strike
          // position.
          ctx.fillRect(seg.x, decoY - em * 0.34, seg.width, thickness);
          ctx.fillRect(seg.x, decoY - em * 0.2, seg.width, thickness);
        }
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
            this.drawBitmap(
              el,
              img.crop,
              img.x,
              baselineY - img.height,
              img.width,
              img.height,
            );
            if (img.outline)
              this.strokeBox(
                img.outline,
                img.x,
                baselineY - img.height,
                img.width,
                img.height,
              );
          } else if (this.undecodable.has(img.src)) {
            this.drawPlaceholder(
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
        const x1 = x;
        const y1 = s.flipV ? y + h : y;
        const x2 = x + w;
        const y2 = s.flipV ? y : y + h;
        ctx.strokeStyle = s.stroke;
        ctx.lineWidth = lw;
        if (s.dash?.length) ctx.setLineDash(dashArray(s.dash, lw, s.cap));
        // Only touch lineCap when it actually changes: the context is shared
        // and long-lived, so an unconditional write (and reset) would be one
        // more state change on every line the document draws.
        const capped = s.cap === 'round' || s.cap === 'square';
        if (capped) ctx.lineCap = s.cap === 'square' ? 'square' : 'round';
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
        ctx.setLineDash([]);
        if (capped) ctx.lineCap = 'butt';
        // Arrowheads: filled triangles scaled from the stroke width (Word's
        // "block" arrow at medium size ≈ 3×). Drawn along the line direction.
        if (s.arrowStart || s.arrowEnd) {
          const ang = Math.atan2(y2 - y1, x2 - x1);
          const len = Math.max(6, lw * 4);
          const half = Math.max(2.5, lw * 1.6);
          const head = (px: number, py: number, a: number) => {
            ctx.fillStyle = s.stroke as string;
            ctx.beginPath();
            ctx.moveTo(px, py);
            ctx.lineTo(
              px - len * Math.cos(a) + half * Math.sin(a),
              py - len * Math.sin(a) - half * Math.cos(a),
            );
            ctx.lineTo(
              px - len * Math.cos(a) - half * Math.sin(a),
              py - len * Math.sin(a) + half * Math.cos(a),
            );
            ctx.closePath();
            ctx.fill();
          };
          if (s.arrowEnd) head(x2, y2, ang);
          if (s.arrowStart) head(x1, y1, ang + Math.PI);
        }
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
        // The importer normalizes each dialect's own statement of roundness
        // to a fraction of the shorter side; absent, fall back to the
        // DrawingML default adj (16667/100000 of the shorter side).
        const ratio = s.cornerRatio ?? 0.16667;
        const r = Math.min(ratio * Math.min(w, h), w / 2, h / 2);
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
        // Horizontal gridlines OCCUPY vertical space (the layout reserves
        // their width above/below the cell boxes), so outer-edge detection
        // allows for the reserved band, and the strokes sit OUTSIDE the box,
        // centered in the band — adjacent cells' strokes then coincide. For
        // 1px borders this is the same half-pixel crispness as before.
        const topBandW = (b?.top ? b.top.width : 0) + eps;
        const bottomBandW = (b?.bottom ? b.bottom.width : 0) + eps;
        const topOuter = cell.y - table.y < topBandW;
        const bottomOuter =
          table.y + table.height - (cell.y + cell.height) < bottomBandW;
        const leftOuter = Math.abs(cell.x - table.x) < eps;
        const rightOuter =
          Math.abs(cell.x + cell.width - (table.x + table.width)) < eps;
        // Cell border wins; else the table's outer/inside edge. A BorderSide is
        // visible, `false` is explicit-none, absent inherits.
        const top = cb?.top ?? (topOuter ? b?.top : b?.insideH);
        const bottom = cb?.bottom ?? (bottomOuter ? b?.bottom : b?.insideH);
        const left = cb?.left ?? (leftOuter ? b?.left : b?.insideV);
        const right = cb?.right ?? (rightOuter ? b?.right : b?.insideV);
        const y0 = yOffset + cell.y - (top ? top.width / 2 : -0.5);
        const y1 =
          yOffset + cell.y + cell.height + (bottom ? bottom.width / 2 : 0.5);
        if (top) this.strokeBorder(top, x0, y0, x1, y0);
        if (bottom) this.strokeBorder(bottom, x0, y1, x1, y1);
        if (left) this.strokeBorder(left, x0, y0, x0, y1);
        if (right) this.strokeBorder(right, x1, y0, x1, y1);
      }
      ctx.setLineDash([]);
    }
    // w:tl2br / w:br2tl cross the cell corner to corner instead of running
    // along an edge, so they belong to the cell alone — painted whether or
    // not the table declares any borders at all. Both present draws an X.
    for (const cell of table.cells) {
      const d = cell.diagonals;
      if (!d) continue;
      const x0 = cell.x + 0.5;
      const x1 = cell.x + cell.width + 0.5;
      const y0 = yOffset + cell.y + 0.5;
      const y1 = yOffset + cell.y + cell.height + 0.5;
      if (d.tl2br) this.strokeBorder(d.tl2br, x0, y0, x1, y1);
      if (d.br2tl) this.strokeBorder(d.br2tl, x0, y1, x1, y0);
      ctx.setLineDash([]);
    }
    // Anchored images/shapes positioned in the cells, on top (Word paints
    // non-behindDoc drawings over the table grid).
    for (const cell of table.cells) {
      for (const f of cell.floats ?? [])
        if (!f.behind) this.paintFloat(f, yOffset, o, pageInfo);
    }
  }

  /** Four edges of a box in one border style — a picture's own outline. */
  /** The stand-in for a picture the engine cannot decode: a light box with a
   *  hairline border and a small "picture" glyph (frame + sun + hills) in the
   *  middle, sized to the box — the same idea as Word's "cannot be displayed"
   *  frame, without its text (whose font would be a guess). Kept flat and
   *  neutral so it never competes with the page. */
  private drawPlaceholder(x: number, y: number, w: number, h: number): void {
    if (w <= 0 || h <= 0) return;
    const ctx = this.ctx;
    ctx.save();
    ctx.fillStyle = '#f2f2f2';
    ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = '#b8b8b8';
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, Math.max(0, w - 1), Math.max(0, h - 1));
    const s = Math.min(w, h) * 0.5;
    if (s >= 8) {
      const gx = x + (w - s) / 2;
      const gy = y + (h - s) / 2;
      ctx.strokeStyle = '#9a9a9a';
      ctx.lineWidth = Math.max(1, s / 24);
      ctx.strokeRect(gx, gy, s, s * 0.8);
      ctx.fillStyle = '#9a9a9a';
      ctx.beginPath(); // sun
      ctx.arc(gx + s * 0.3, gy + s * 0.25, s * 0.09, 0, Math.PI * 2);
      ctx.fill();
      ctx.beginPath(); // hills
      ctx.moveTo(gx + s * 0.08, gy + s * 0.72);
      ctx.lineTo(gx + s * 0.38, gy + s * 0.4);
      ctx.lineTo(gx + s * 0.56, gy + s * 0.58);
      ctx.lineTo(gx + s * 0.68, gy + s * 0.46);
      ctx.lineTo(gx + s * 0.92, gy + s * 0.72);
      ctx.closePath();
      ctx.fill();
    }
    ctx.restore();
  }

  private strokeBox(
    side: BorderSide,
    x: number,
    y: number,
    w: number,
    h: number,
  ): void {
    this.strokeBorder(side, x, y, x + w, y);
    this.strokeBorder(side, x, y + h, x + w, y + h);
    this.strokeBorder(side, x, y, x, y + h);
    this.strokeBorder(side, x + w, y, x + w, y + h);
    this.ctx.setLineDash([]);
  }

  /** Draw a bitmap into a box, honouring an `a:srcRect` crop.
   *
   *  Uncropped goes through the 5-argument form, which is what every image
   *  used before this existed. Cropped picks the sub-rectangle out of the
   *  BITMAP's own pixels and lets it scale to fill the same box — the box
   *  never changes size, only what is inside it.
   *
   *  Negative offsets (Word's outset) make the source rectangle reach past
   *  the bitmap. That needs no special case: the canvas spec clips such a
   *  rectangle to the image and shrinks the destination by the same ratio, so
   *  the overhang lands as empty space rather than stretched pixels. */
  private drawBitmap(
    el: CanvasImageSource & { naturalWidth: number; naturalHeight: number },
    crop: ImageCrop | undefined,
    x: number,
    y: number,
    w: number,
    h: number,
  ): void {
    if (!crop) {
      this.ctx.drawImage(el, x, y, w, h);
      return;
    }
    const iw = el.naturalWidth;
    const ih = el.naturalHeight;
    const sx = crop.l * iw;
    const sy = crop.t * ih;
    const sw = iw - (crop.l + crop.r) * iw;
    const sh = ih - (crop.t + crop.b) * ih;
    // A crop that leaves nothing (l + r >= 1) would ask for a zero or
    // negative source width, which throws. Word shows nothing in that case
    // and so do we.
    if (sw <= 0 || sh <= 0) return;
    this.ctx.drawImage(el, sx, sy, sw, sh, x, y, w, h);
  }

  /** Stroke one border edge with its width / style / colour. Edges are
   *  horizontal or vertical (table grid lines), with one exception: the cell
   *  diagonals w:tl2br / w:br2tl. Dashes and solid strokes follow the line
   *  whatever its angle; the `double` style offsets its two rules
   *  perpendicular to a horizontal or vertical edge, so on a diagonal it
   *  falls back to the vertical offset. Word draws double diagonals so
   *  rarely that a proper normal-vector offset would be untested code. */
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
    el.onerror = () => {
      // Undecodable: remember, and repaint so the placeholder replaces the
      // blank the first paint left where the picture belongs.
      this.undecodable.add(src);
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
