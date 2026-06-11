import type {
  FontSpec,
  LayoutLine,
  ResolvedLayout,
  ResolvedPage,
  ResolvedTable,
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
}

type Required_<T> = { [K in keyof T]-?: T[K] };
type ResolvedOptions = Required_<Omit<PaintOptions, 'devicePixelRatio'>>;

const DEFAULTS: ResolvedOptions = {
  zoom: 1,
  pageGap: 24,
  pageBackground: '#ffffff',
  pageBorder: '#c8c8c8',
  tableBorder: '#b0b0b0',
  textColor: '#000000',
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
  private readonly images = new Map<string, HTMLImageElement>();
  private lastLayout: ResolvedLayout | null = null;
  private lastOptions: PaintOptions = {};

  constructor(private readonly canvas: HTMLCanvasElement) {
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('bapbong-painter-canvas: 2D canvas context unavailable');
    this.ctx = ctx;
  }

  paint(layout: ResolvedLayout, options: PaintOptions = {}): void {
    this.lastLayout = layout;
    this.lastOptions = options;
    const o: ResolvedOptions = { ...DEFAULTS, ...options };
    const dpr = options.devicePixelRatio ?? defaultDpr();

    const width = layout.pages.reduce((m, p) => Math.max(m, p.width), 0);
    const height =
      layout.pages.reduce((s, p) => s + p.height, 0) +
      Math.max(0, layout.pages.length - 1) * o.pageGap;

    this.canvas.width = Math.max(1, Math.round(width * o.zoom * dpr));
    this.canvas.height = Math.max(1, Math.round(height * o.zoom * dpr));
    this.canvas.style.width = `${Math.round(width * o.zoom)}px`;
    this.canvas.style.height = `${Math.round(height * o.zoom)}px`;

    const ctx = this.ctx;
    ctx.setTransform(o.zoom * dpr, 0, 0, o.zoom * dpr, 0, 0);
    ctx.clearRect(0, 0, width, height);
    ctx.textBaseline = 'alphabetic';

    let yOffset = 0;
    for (const page of layout.pages) {
      this.paintPage(page, yOffset, o);
      yOffset += page.height + o.pageGap;
    }
  }

  private paintPage(page: ResolvedPage, yOffset: number, o: ResolvedOptions): void {
    const ctx = this.ctx;
    ctx.fillStyle = o.pageBackground;
    ctx.fillRect(0, yOffset, page.width, page.height);
    ctx.strokeStyle = o.pageBorder;
    ctx.lineWidth = 1;
    ctx.strokeRect(0.5, yOffset + 0.5, page.width - 1, page.height - 1);

    for (const line of page.lines) this.paintLine(line, yOffset, o);
    for (const table of page.tables ?? []) this.paintTable(table, yOffset, o);
  }

  private paintLine(line: LayoutLine, yOffset: number, o: ResolvedOptions): void {
    const ctx = this.ctx;
    const baselineY = yOffset + line.y + line.baseline;
    for (const seg of line.segments) {
      ctx.font = fontCss(seg.font);
      ctx.fillStyle = seg.color ?? o.textColor;
      ctx.fillText(seg.text, seg.x, baselineY);
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
