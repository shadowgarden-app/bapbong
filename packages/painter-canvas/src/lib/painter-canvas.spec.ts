import type {
  FontSpec,
  ResolvedLayout,
} from '@shadow-garden/bapbong-contracts';
import { CanvasPainter } from './painter-canvas.js';

interface Call {
  method: string;
  args: unknown[];
  font: string;
  fillStyle: string;
  strokeStyle: string;
}

/** Minimal 2D-context stand-in that records draw calls with current state. */
class RecordingCtx {
  calls: Call[] = [];
  font = '';
  fillStyle = '';
  strokeStyle = '';
  lineWidth = 1;
  textBaseline = '';

  setTransform(...args: unknown[]) {
    this.record('setTransform', args);
  }
  clearRect(...args: unknown[]) {
    this.record('clearRect', args);
  }
  fillRect(...args: unknown[]) {
    this.record('fillRect', args);
  }
  strokeRect(...args: unknown[]) {
    this.record('strokeRect', args);
  }
  fillText(...args: unknown[]) {
    this.record('fillText', args);
  }
  drawImage(...args: unknown[]) {
    this.record('drawImage', args);
  }
  beginPath(...args: unknown[]) {
    this.record('beginPath', args);
  }
  moveTo(...args: unknown[]) {
    this.record('moveTo', args);
  }
  lineTo(...args: unknown[]) {
    this.record('lineTo', args);
  }
  stroke(...args: unknown[]) {
    this.record('stroke', args);
  }
  setLineDash(...args: unknown[]) {
    this.record('setLineDash', args);
  }
  fill(...args: unknown[]) {
    this.record('fill', args);
  }
  rect(...args: unknown[]) {
    this.record('rect', args);
  }
  ellipse(...args: unknown[]) {
    this.record('ellipse', args);
  }
  quadraticCurveTo(...args: unknown[]) {
    this.record('quadraticCurveTo', args);
  }
  closePath(...args: unknown[]) {
    this.record('closePath', args);
  }
  save(...args: unknown[]) {
    this.record('save', args);
  }
  restore(...args: unknown[]) {
    this.record('restore', args);
  }
  translate(...args: unknown[]) {
    this.record('translate', args);
  }
  rotate(...args: unknown[]) {
    this.record('rotate', args);
  }
  clip(...args: unknown[]) {
    this.record('clip', args);
  }

  of(method: string): Call[] {
    return this.calls.filter((c) => c.method === method);
  }

  private record(method: string, args: unknown[]) {
    this.calls.push({
      method,
      args,
      font: this.font,
      fillStyle: this.fillStyle,
      strokeStyle: this.strokeStyle,
    });
  }
}

/** A fake <canvas>: own recording ctx, tracked width/height/style/parent. */
interface FakeCanvas {
  width: number;
  height: number;
  style: Record<string, string>;
  parentNode: FakeContainer | null;
  _ctx: RecordingCtx;
  getContext(): RecordingCtx;
}
function makeCanvas(): FakeCanvas {
  const ctx = new RecordingCtx();
  return {
    width: 0,
    height: 0,
    style: {},
    parentNode: null,
    _ctx: ctx,
    getContext: () => ctx,
  };
}

/** A fake container element the painter fills with page canvases. */
interface FakeContainer {
  style: Record<string, string>;
  children: FakeCanvas[];
  appendChild(c: FakeCanvas): FakeCanvas;
  removeChild(c: FakeCanvas): FakeCanvas;
}
function makeContainer(): FakeContainer {
  const children: FakeCanvas[] = [];
  return {
    style: {},
    children,
    appendChild(c) {
      c.parentNode = this;
      children.push(c);
      return c;
    },
    removeChild(c) {
      const i = children.indexOf(c);
      if (i >= 0) children.splice(i, 1);
      c.parentNode = null;
      return c;
    },
  };
}

/** Painter whose pages each get their own recording ctx (created on demand). */
function setup() {
  const container = makeContainer();
  const painter = new CanvasPainter(container as unknown as HTMLElement, {
    createCanvas: () => makeCanvas() as unknown as HTMLCanvasElement,
  });
  return { painter, container };
}
/** The recording ctx of the i-th currently-mounted page canvas. */
const ctxAt = (container: FakeContainer, i: number) =>
  container.children[i]._ctx;

const font = (over: Partial<FontSpec> = {}): FontSpec => ({
  family: 'Arial',
  sizePt: 11,
  bold: false,
  italic: false,
  ...over,
});

const page = (lines: ResolvedLayout['pages'][0]['lines'], index = 0) => ({
  index,
  width: 200,
  height: 300,
  lines,
});

const helloLine = {
  x: 20,
  y: 20,
  width: 160,
  height: 16,
  baseline: 12,
  segments: [
    { x: 20, text: 'Hello', font: font({ bold: true }), color: '#ff0000' },
  ],
};

describe('CanvasPainter', () => {
  it('sizes each page canvas for zoom × dpr and the container for scroll range', () => {
    const { painter, container } = setup();
    painter.paint(
      { pages: [page([helloLine]), page([], 1)] },
      { zoom: 2, devicePixelRatio: 2, pageGap: 10 },
    );
    const c0 = container.children[0];
    expect(c0.width).toBe(200 * 2 * 2);
    expect(c0.height).toBe(300 * 2 * 2);
    expect(c0.style['width']).toBe('400px');
    expect(ctxAt(container, 0).of('setTransform')[0].args).toEqual([
      4, 0, 0, 4, 0, 0,
    ]);
    // The container carries the full stacked size so the wrap can scroll.
    expect(container.style['width']).toBe('400px');
    expect(container.style['height']).toBe(`${(300 + 10 + 300) * 2}px`);
  });

  it('paints the page background and text at the baseline', () => {
    const { painter, container } = setup();
    painter.paint({ pages: [page([helloLine])] }, { devicePixelRatio: 1 });
    const ctx = ctxAt(container, 0);
    const bg = ctx.of('fillRect')[0];
    expect(bg.args).toEqual([0, 0, 200, 300]);
    expect(bg.fillStyle).toBe('#ffffff');
    const text = ctx.of('fillText')[0];
    expect(text.args).toEqual(['Hello', 20, 32]); // line.y 20 + baseline 12
    expect(text.font).toBe('700 11pt Arial');
    expect(text.fillStyle).toBe('#ff0000');
  });

  it('positions the second page canvas by page height + gap (content is page-local)', () => {
    const { painter, container } = setup();
    painter.paint(
      { pages: [page([]), page([helloLine], 1)] },
      { devicePixelRatio: 1, pageGap: 10 },
    );
    const page1 = container.children[1];
    expect(page1.style['top']).toBe('310px'); // 300 + 10
    const ctx1 = ctxAt(container, 1);
    expect(ctx1.of('fillRect')[0].args).toEqual([0, 0, 200, 300]); // bg, page-local
    expect(ctx1.of('fillText')[0].args).toEqual(['Hello', 20, 32]); // local baseline
  });

  it('fills segment highlight behind the text and cell shading behind content', () => {
    const { painter, container } = setup();
    const hl = {
      ...helloLine,
      segments: [
        { x: 20, text: 'Hi', font: font(), background: '#FFFF00', width: 30 },
      ],
    };
    const p = {
      ...page([hl]),
      tables: [
        {
          x: 20,
          y: 60,
          width: 100,
          height: 16,
          cells: [
            {
              x: 20,
              y: 60,
              width: 100,
              height: 16,
              colspan: 1,
              rowspan: 1,
              background: '#D9E2F3',
              lines: [],
            },
          ],
        },
      ],
    };
    painter.paint({ pages: [p] }, { devicePixelRatio: 1 });
    const ctx = ctxAt(container, 0);
    const fills = ctx.of('fillRect');
    const hlFill = fills.find((c) => c.fillStyle === '#FFFF00');
    expect(hlFill?.args).toEqual([20, 20, 30, 16]); // segment bg over the line box
    const cellFill = fills.find((c) => c.fillStyle === '#D9E2F3');
    expect(cellFill?.args).toEqual([20, 60, 100, 16]);
    // highlight is painted before the glyph
    expect(ctx.calls.indexOf(hlFill as never)).toBeLessThan(
      ctx.calls.findIndex((c) => c.method === 'fillText'),
    );
  });

  it('fills a background decoration behind the glyphs', () => {
    // Comment tint / find-highlight etc. arrive as generic plugin decorations,
    // pre-resolved to page-local rects (the painter no longer knows "comments").
    const { painter, container } = setup();
    painter.paint(
      { pages: [page([helloLine])] },
      {
        devicePixelRatio: 1,
        decorations: [
          {
            rects: [{ pageIndex: 0, x: 20, y: 20, width: 30, height: 16 }],
            kind: 'background',
            color: 'rgba(255, 193, 7, 0.28)',
          },
        ],
      },
    );
    const ctx = ctxAt(container, 0);
    const tint = ctx
      .of('fillRect')
      .find((c) => c.fillStyle.startsWith('rgba(255'));
    expect(tint?.args).toEqual([20, 20, 30, 16]);
    // painted before the glyph
    expect(ctx.calls.indexOf(tint as never)).toBeLessThan(
      ctx.calls.findIndex((c) => c.method === 'fillText'),
    );
  });

  it('draws underline and strike from layout-measured widths', () => {
    const { painter, container } = setup();
    const decorated = {
      ...helloLine,
      segments: [
        {
          x: 20,
          text: 'Hi',
          font: font(),
          underline: true,
          strike: true,
          width: 30,
        },
      ],
    };
    painter.paint({ pages: [page([decorated])] }, { devicePixelRatio: 1 });
    const rects = ctxAt(container, 0).of('fillRect').slice(1); // [0] is the page background
    // 11pt → em ≈ 14.67px: underline ≈ baseline+1.47, strike ≈ baseline−3.96.
    expect(rects).toHaveLength(2);
    const [under, strike] = rects;
    expect(under.args[0]).toBe(20);
    expect(under.args[2]).toBe(30); // layout width, not re-measured
    expect(under.args[1] as number).toBeGreaterThan(32);
    expect(strike.args[1] as number).toBeLessThan(32);
  });

  it('draws declared table borders and paints cell content', () => {
    const { painter, container } = setup();
    const s = { width: 1, style: 'solid' as const, color: '#000' };
    const p = {
      ...page([]),
      tables: [
        {
          x: 20,
          y: 20,
          width: 100,
          height: 16,
          borders: {
            top: s,
            bottom: s,
            left: s,
            right: s,
            insideH: s,
            insideV: s,
          },
          cells: [
            {
              x: 20,
              y: 20,
              width: 100,
              height: 16,
              colspan: 1,
              rowspan: 1,
              lines: [helloLine],
            },
          ],
        },
      ],
    };
    painter.paint({ pages: [p] }, { devicePixelRatio: 1 });
    const ctx = ctxAt(container, 0);
    // single cell, all edges outer → 4 edges drawn in one path
    expect(ctx.of('lineTo')).toHaveLength(4);
    expect(ctx.of('stroke').length).toBeGreaterThan(0);
    expect(ctx.of('fillText')[0].args).toEqual(['Hello', 20, 32]);
  });

  it('draws per-cell border overrides even with no table borders', () => {
    const { painter, container } = setup();
    const p = {
      ...page([]),
      tables: [
        {
          x: 20,
          y: 20,
          width: 100,
          height: 16,
          cells: [
            {
              x: 20,
              y: 20,
              width: 100,
              height: 16,
              colspan: 1,
              rowspan: 1,
              borders: {
                bottom: { width: 1, style: 'solid' as const, color: '#000' },
              },
              lines: [],
            },
          ],
        },
      ],
    };
    painter.paint({ pages: [p] }, { devicePixelRatio: 1 });
    // only the bottom edge → one line segment despite no table borders
    expect(ctxAt(container, 0).of('lineTo')).toHaveLength(1);
  });

  it('paints tables WITHOUT borders when none are declared (OOXML default)', () => {
    const { painter, container } = setup();
    const p = {
      ...page([]),
      tables: [
        {
          x: 20,
          y: 20,
          width: 100,
          height: 16,
          cells: [
            {
              x: 20,
              y: 20,
              width: 100,
              height: 16,
              colspan: 1,
              rowspan: 1,
              lines: [helloLine],
            },
          ],
        },
      ],
    };
    painter.paint({ pages: [p] }, { devicePixelRatio: 1 });
    const ctx = ctxAt(container, 0);
    expect(ctx.of('lineTo')).toHaveLength(0); // no borders
    expect(ctx.of('strokeRect')).toHaveLength(1); // only the page border
    expect(ctx.of('fillText')[0].args).toEqual(['Hello', 20, 32]); // content intact
  });

  it('draws selection under the text and the caret on top (page-local)', () => {
    const { painter, container } = setup();
    painter.paint(
      { pages: [page([]), page([helloLine], 1)] },
      {
        devicePixelRatio: 1,
        pageGap: 10,
        selection: [{ pageIndex: 1, x: 20, y: 20, width: 30, height: 16 }],
        caret: { pageIndex: 1, x: 50, y: 20, height: 16 },
      },
    );
    expect(container.children[1].style['top']).toBe('310px'); // page-offset lives here
    const ctx1 = ctxAt(container, 1);
    const fills = ctx1.of('fillRect');
    const sel = fills.find((c) => c.fillStyle.startsWith('rgba'));
    expect(sel?.args).toEqual([20, 20, 30, 16]); // page-local
    const caret = fills[fills.length - 1];
    expect(caret.args).toEqual([50, 20, 1.5, 16]);
    // selection painted before the line's text, caret after.
    const textIdx = ctx1.calls.findIndex((c) => c.method === 'fillText');
    expect(ctx1.calls.indexOf(sel as never)).toBeLessThan(textIdx);
    expect(ctx1.calls.indexOf(caret)).toBeGreaterThan(textIdx);
  });

  it('paintOverlay redraws only the pages whose caret/selection changed', () => {
    const { painter, container } = setup();
    painter.paint(
      {
        pages: [page([helloLine]), page([helloLine], 1), page([helloLine], 2)],
      },
      {
        devicePixelRatio: 1,
        caret: { pageIndex: 1, x: 50, y: 20, height: 16 },
      },
    );
    const before = container.children.map((c) => c._ctx.calls.length);
    painter.paintOverlay({
      caret: { pageIndex: 1, x: 60, y: 20, height: 16 },
      selection: [],
    });
    const after = container.children.map((c) => c._ctx.calls.length);
    expect(after[0]).toBe(before[0]); // page 0 untouched
    expect(after[2]).toBe(before[2]); // page 2 untouched
    expect(after[1]).toBeGreaterThan(before[1]); // only the caret's page redrew
    expect(ctxAt(container, 1).of('fillRect').at(-1)?.args).toEqual([
      60, 20, 1.5, 16,
    ]);
  });

  it('skips the backing-store resize when dimensions are unchanged', () => {
    const container = makeContainer();
    let widthSets = 0;
    const painter = new CanvasPainter(container as unknown as HTMLElement, {
      createCanvas: () => {
        const c = makeCanvas();
        let w = 0;
        Object.defineProperty(c, 'width', {
          get: () => w,
          set: (v: number) => {
            widthSets++;
            w = v;
          },
        });
        return c as unknown as HTMLCanvasElement;
      },
    });
    const layout: ResolvedLayout = { pages: [page([helloLine])] };
    painter.paint(layout, { devicePixelRatio: 1 });
    expect(widthSets).toBe(1);
    painter.paint(layout, { devicePixelRatio: 1 }); // same size → no reassignment
    expect(widthSets).toBe(1);
  });

  it('paints floats in Word z-order: default over the text, behindDoc under', async () => {
    class FakeImage {
      onload: (() => void) | null = null;
      complete = false;
      naturalWidth = 0;
      set src(_v: string) {
        queueMicrotask(() => {
          this.complete = true;
          this.naturalWidth = 80;
          this.onload?.();
        });
      }
    }
    vi.stubGlobal('Image', FakeImage);
    try {
      // Default (no behind flag): the anchored drawing paints OVER the text.
      const { painter, container } = setup();
      painter.paint(
        {
          pages: [
            {
              ...page([helloLine]),
              floats: [{ x: 140, y: 20, width: 80, height: 40, src: 'f' }],
            },
          ],
        },
        { devicePixelRatio: 1 },
      );
      await new Promise((r) => setTimeout(r, 0));
      const ctx = ctxAt(container, 0);
      const draw = ctx.of('drawImage')[0];
      expect(draw.args.slice(1)).toEqual([140, 20, 80, 40]);
      const lastDraw = ctx.calls.lastIndexOf(
        ctx.of('drawImage').at(-1) as never,
      );
      const lastText = ctx.calls.lastIndexOf(
        ctx.of('fillText').at(-1) as never,
      );
      expect(lastDraw).toBeGreaterThan(lastText);

      // behindDoc: under the text (watermarks).
      const b = setup();
      b.painter.paint(
        {
          pages: [
            {
              ...page([helloLine]),
              floats: [
                {
                  x: 140,
                  y: 20,
                  width: 80,
                  height: 40,
                  src: 'f',
                  behind: true,
                },
              ],
            },
          ],
        },
        { devicePixelRatio: 1 },
      );
      await new Promise((r) => setTimeout(r, 0));
      const bctx = ctxAt(b.container, 0);
      const bDraw = bctx.calls.lastIndexOf(
        bctx.of('drawImage').at(-1) as never,
      );
      const bText = bctx.calls.lastIndexOf(bctx.of('fillText').at(-1) as never);
      expect(bDraw).toBeLessThan(bText);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('substitutes live page numbers into chrome field segments', () => {
    const { painter, container } = setup();
    const fieldLine = {
      ...helloLine,
      segments: [
        {
          x: 20,
          text: '1',
          font: font(),
          field: 'pageNumber' as const,
          width: 10,
        },
        {
          x: 35,
          text: '1',
          font: font(),
          field: 'pageCount' as const,
          width: 10,
        },
      ],
    };
    painter.paint(
      {
        pages: [page([]), page([], 1)],
        pageFooter: { lines: [fieldLine], tables: [], height: 16 },
      },
      { devicePixelRatio: 1, pageGap: 10 },
    );
    // Each page's footer renders its own page number against the total.
    expect(
      ctxAt(container, 0)
        .of('fillText')
        .map((c) => c.args[0]),
    ).toEqual(['1', '2']);
    expect(
      ctxAt(container, 1)
        .of('fillText')
        .map((c) => c.args[0]),
    ).toEqual(['2', '2']);
  });

  it('stamps page chrome (header/footer) onto every page', () => {
    const { painter, container } = setup();
    const headerLine = {
      ...helloLine,
      y: 5,
      segments: [{ x: 20, text: 'hdr', font: font() }],
    };
    painter.paint(
      {
        pages: [page([]), page([], 1)],
        pageHeader: { lines: [headerLine], tables: [], height: 16 },
      },
      { devicePixelRatio: 1, pageGap: 10 },
    );
    // Header drawn on each page canvas at its page-local y (offset is in style.top).
    expect(ctxAt(container, 0).of('fillText')[0].args).toEqual(['hdr', 20, 17]); // 5 + baseline 12
    expect(ctxAt(container, 1).of('fillText')[0].args).toEqual(['hdr', 20, 17]);
    expect(container.children[1].style['top']).toBe('310px');
  });

  it('selects first/even header variants per page', () => {
    const { painter, container } = setup();
    const chrome = (text: string) => ({
      lines: [{ ...helloLine, segments: [{ x: 20, text, font: font() }] }],
      tables: [],
      height: 16,
    });
    painter.paint(
      {
        pages: [page([]), page([], 1), page([], 2), page([], 3)],
        pageHeader: chrome('def'),
        pageHeaderFirst: chrome('first'),
        pageHeaderEven: chrome('even'),
        chromeSelect: { titlePg: true, evenAndOdd: true },
      },
      { devicePixelRatio: 1 },
    );
    const txt = (i: number) =>
      ctxAt(container, i)
        .of('fillText')
        .map((c) => c.args[0]);
    expect(txt(0)).toEqual(['first']); // page 1 → title page
    expect(txt(1)).toEqual(['even']); // page 2 → even
    expect(txt(2)).toEqual(['def']); // page 3 → odd/default
    expect(txt(3)).toEqual(['even']); // page 4 → even
  });

  it('shows a blank band when the selected variant is absent', () => {
    const { painter, container } = setup();
    const chrome = (text: string) => ({
      lines: [{ ...helloLine, segments: [{ x: 20, text, font: font() }] }],
      tables: [],
      height: 16,
    });
    painter.paint(
      {
        pages: [page([]), page([], 1)],
        pageHeader: chrome('def'),
        chromeSelect: { titlePg: true, evenAndOdd: false },
      },
      { devicePixelRatio: 1 },
    );
    expect(ctxAt(container, 0).of('fillText')).toHaveLength(0); // page 1: titlePg, no first → blank
    expect(
      ctxAt(container, 1)
        .of('fillText')
        .map((c) => c.args[0]),
    ).toEqual(['def']);
  });

  it('virtualizes pages outside the viewport (unmounted, no canvas)', () => {
    const { painter, container } = setup();
    const layout: ResolvedLayout = { pages: [page([]), page([helloLine], 1)] };
    // Page 1 spans y 310..610 (gap 10). Viewport 0..50 + 200 margin → hidden.
    painter.paint(layout, {
      devicePixelRatio: 1,
      pageGap: 10,
      viewport: { top: 0, height: 50 },
    });
    expect(container.children).toHaveLength(1); // only page 0 mounted
    expect(ctxAt(container, 0).of('fillText')).toHaveLength(0);

    // Scrolled down: page 1 enters the viewport → mounted, its text paints.
    painter.paint(layout, {
      devicePixelRatio: 1,
      pageGap: 10,
      viewport: { top: 320, height: 50 },
    });
    const allText = container.children.flatMap((c) => c._ctx.of('fillText'));
    expect(allText).toHaveLength(1);
    expect(allText[0].args).toEqual(['Hello', 20, 32]);
  });

  it('maps container coords to page-local coords and back (zoom + gap aware)', () => {
    const { painter } = setup();
    painter.paint(
      { pages: [page([]), page([], 1)] },
      { devicePixelRatio: 1, zoom: 2, pageGap: 10 },
    );
    // CSS y = (300 + 10 + 5) * 2 = 630 → page 1, y = 5.
    expect(painter.canvasToPage(40, 630)).toEqual({
      pageIndex: 1,
      x: 20,
      y: 5,
    });
    // Gap point clamps to the nearer edge (here: bottom of page 0).
    expect(painter.canvasToPage(0, (300 + 2) * 2)).toEqual({
      pageIndex: 0,
      x: 0,
      y: 300,
    });
    expect(painter.pageToCanvas({ pageIndex: 1, x: 20, y: 5 })).toEqual({
      x: 40,
      y: 630,
    });
  });

  it('skips images where Image is unavailable, draws after async load', async () => {
    const layout: ResolvedLayout = {
      pages: [
        page([
          {
            ...helloLine,
            segments: [],
            images: [{ x: 30, src: 'data:img', width: 40, height: 10 }],
          },
        ]),
      ],
    };

    // Node (no Image): paint must not throw and must not draw.
    {
      const { painter, container } = setup();
      painter.paint(layout, { devicePixelRatio: 1 });
      expect(ctxAt(container, 0).of('drawImage')).toHaveLength(0);
    }

    // Stubbed Image: load completes on a microtask → painter repaints with it.
    class FakeImage {
      onload: (() => void) | null = null;
      complete = false;
      naturalWidth = 0;
      set src(_v: string) {
        queueMicrotask(() => {
          this.complete = true;
          this.naturalWidth = 40;
          this.onload?.();
        });
      }
    }
    vi.stubGlobal('Image', FakeImage);
    try {
      const { painter, container } = setup();
      painter.paint(layout, { devicePixelRatio: 1 });
      expect(ctxAt(container, 0).of('drawImage')).toHaveLength(0); // not loaded yet
      await new Promise((r) => setTimeout(r, 0));
      const draw = ctxAt(container, 0).of('drawImage')[0];
      // x, baselineY (20 + 12) − height 10 = 22, w, h
      expect(draw.args.slice(1)).toEqual([30, 22, 40, 10]);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('draws preset geometry floats (ellipse, roundRect, horizontalScroll)', () => {
    const { painter, container } = setup();
    const floats = [
      {
        x: 10,
        y: 20,
        width: 40,
        height: 20,
        src: '',
        shape: {
          kind: 'ellipse' as const,
          fill: '#FFEE00',
          stroke: '#000000',
          strokeWidth: 2,
        },
      },
      {
        x: 0,
        y: 60,
        width: 40,
        height: 40,
        src: '',
        shape: { kind: 'roundRect' as const, stroke: '#112233' },
      },
      {
        x: 0,
        y: 120,
        width: 80,
        height: 40,
        src: '',
        shape: {
          kind: 'horizontalScroll' as const,
          fill: '#FFFFFF',
          stroke: '#000000',
        },
      },
    ];
    painter.paint(
      { pages: [{ ...page([]), floats }] },
      { devicePixelRatio: 1 },
    );
    const ctx = ctxAt(container, 0);

    // Ellipse: centered in its box, radii shrunk by the stroke width.
    const ell = ctx.of('ellipse');
    expect(ell[0].args).toEqual([30, 30, 19, 9, 0, 0, Math.PI * 2]);
    expect(ctx.of('fill').length).toBeGreaterThan(0);

    // Round rect traces four corner curves.
    expect(ctx.of('quadraticCurveTo')).toHaveLength(4);

    // Scroll: paper band + two rolled ends (full-height ellipses).
    // r = 0.125 × min(80,40) = 5 → band from x+5, width 70; rolls at 5 and 75.
    expect(
      ctx.of('rect').some((c) => c.args[0] === 5 && c.args[2] === 70),
    ).toBe(true);
    expect(ell.filter((c) => c.args[0] === 5 || c.args[0] === 75)).toHaveLength(
      2,
    );
  });

  it('clips a cell holding a rotated image to the cell box (overflow hidden)', () => {
    const { painter, container } = setup();
    const line = {
      x: 20,
      y: 60,
      width: 100,
      height: 100,
      baseline: 60,
      segments: [],
      images: [
        {
          x: 20,
          src: '',
          width: 100,
          height: 20,
          rotation: 90,
          shape: { kind: 'rect' as const, stroke: '#000' },
        },
      ],
    };
    const p = {
      ...page([]),
      tables: [
        {
          x: 20,
          y: 60,
          width: 100,
          height: 100,
          cells: [
            {
              x: 20,
              y: 60,
              width: 100,
              height: 100,
              colspan: 1,
              rowspan: 1,
              lines: [line],
            },
          ],
        },
      ],
    };
    painter.paint({ pages: [p] }, { devicePixelRatio: 1 });
    const ctx = ctxAt(container, 0);
    // save → rect(cell box) → clip … restore around the cell's content.
    expect(ctx.of('clip').length).toBe(1);
    expect(
      ctx
        .of('rect')
        .some(
          (c) =>
            c.args[0] === 20 &&
            c.args[1] === 60 &&
            c.args[2] === 100 &&
            c.args[3] === 100,
        ),
    ).toBe(true);
    expect(ctx.of('rotate').length).toBe(1); // the image still paints rotated
  });

  it('rotates a float shape around its box center (paint-only)', () => {
    const { painter, container } = setup();
    const floats = [
      {
        x: 20,
        y: 40,
        width: 60,
        height: 20,
        src: '',
        rotation: 90,
        shape: { kind: 'rect' as const, stroke: '#000000' },
      },
    ];
    painter.paint(
      { pages: [{ ...page([]), floats }] },
      { devicePixelRatio: 1 },
    );
    const ctx = ctxAt(container, 0);
    // translate(center) → rotate(π/2) → translate(−center) around (50, 50).
    const tr = ctx.of('translate');
    expect(tr.some((c) => c.args[0] === 50 && c.args[1] === 50)).toBe(true);
    expect(ctx.of('rotate')[0].args[0]).toBeCloseTo(Math.PI / 2);
    expect(ctx.of('strokeRect').length).toBeGreaterThan(0);
  });
});
