import type { FontSpec, ResolvedLayout } from '@shadow-garden/bapbong-contracts';
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

  setTransform(...args: unknown[]) { this.record('setTransform', args); }
  clearRect(...args: unknown[]) { this.record('clearRect', args); }
  fillRect(...args: unknown[]) { this.record('fillRect', args); }
  strokeRect(...args: unknown[]) { this.record('strokeRect', args); }
  fillText(...args: unknown[]) { this.record('fillText', args); }
  drawImage(...args: unknown[]) { this.record('drawImage', args); }
  beginPath(...args: unknown[]) { this.record('beginPath', args); }
  moveTo(...args: unknown[]) { this.record('moveTo', args); }
  lineTo(...args: unknown[]) { this.record('lineTo', args); }
  stroke(...args: unknown[]) { this.record('stroke', args); }

  of(method: string): Call[] {
    return this.calls.filter((c) => c.method === method);
  }

  private record(method: string, args: unknown[]) {
    this.calls.push({ method, args, font: this.font, fillStyle: this.fillStyle, strokeStyle: this.strokeStyle });
  }
}

function makeCanvas(ctx: RecordingCtx) {
  return {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
}

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
  segments: [{ x: 20, text: 'Hello', font: font({ bold: true }), color: '#ff0000' }],
};

describe('CanvasPainter', () => {
  it('sizes the canvas for zoom × devicePixelRatio and styles in CSS px', () => {
    const ctx = new RecordingCtx();
    const canvas = makeCanvas(ctx);
    new CanvasPainter(canvas).paint(
      { pages: [page([helloLine]), page([], 1)] },
      { zoom: 2, devicePixelRatio: 2, pageGap: 10 },
    );
    expect(canvas.width).toBe(200 * 2 * 2);
    expect(canvas.height).toBe((300 + 10 + 300) * 2 * 2);
    expect((canvas.style as unknown as Record<string, string>)['width']).toBe('400px');
    expect(ctx.of('setTransform')[0].args).toEqual([4, 0, 0, 4, 0, 0]);
  });

  it('paints the page background and text at the baseline', () => {
    const ctx = new RecordingCtx();
    new CanvasPainter(makeCanvas(ctx)).paint(
      { pages: [page([helloLine])] },
      { devicePixelRatio: 1 },
    );
    const bg = ctx.of('fillRect')[0];
    expect(bg.args).toEqual([0, 0, 200, 300]);
    expect(bg.fillStyle).toBe('#ffffff');
    const text = ctx.of('fillText')[0];
    expect(text.args).toEqual(['Hello', 20, 32]); // line.y 20 + baseline 12
    expect(text.font).toBe('700 11pt Arial');
    expect(text.fillStyle).toBe('#ff0000');
  });

  it('offsets the second page by page height + gap', () => {
    const ctx = new RecordingCtx();
    new CanvasPainter(makeCanvas(ctx)).paint(
      { pages: [page([]), page([helloLine], 1)] },
      { devicePixelRatio: 1, pageGap: 10 },
    );
    expect(ctx.of('fillRect')[1].args).toEqual([0, 310, 200, 300]);
    expect(ctx.of('fillText')[0].args).toEqual(['Hello', 20, 310 + 32]);
  });

  it('fills segment highlight behind the text and cell shading behind content', () => {
    const ctx = new RecordingCtx();
    const hl = { ...helloLine, segments: [{ x: 20, text: 'Hi', font: font(), background: '#FFFF00', width: 30 }] };
    const p = {
      ...page([hl]),
      tables: [
        {
          x: 20, y: 60, width: 100, height: 16,
          cells: [{ x: 20, y: 60, width: 100, height: 16, colspan: 1, rowspan: 1, background: '#D9E2F3', lines: [] }],
        },
      ],
    };
    new CanvasPainter(makeCanvas(ctx)).paint({ pages: [p] }, { devicePixelRatio: 1 });
    const fills = ctx.of('fillRect');
    const hlFill = fills.find((c) => c.fillStyle === '#FFFF00');
    expect(hlFill?.args).toEqual([20, 20, 30, 16]); // segment bg over the line box
    const cellFill = fills.find((c) => c.fillStyle === '#D9E2F3');
    expect(cellFill?.args).toEqual([20, 60, 100, 16]);
    // highlight is painted before the glyph
    expect(ctx.calls.indexOf(hlFill as never)).toBeLessThan(ctx.calls.findIndex((c) => c.method === 'fillText'));
  });

  it('draws underline and strike from layout-measured widths', () => {
    const ctx = new RecordingCtx();
    const decorated = {
      ...helloLine,
      segments: [{ x: 20, text: 'Hi', font: font(), underline: true, strike: true, width: 30 }],
    };
    new CanvasPainter(makeCanvas(ctx)).paint({ pages: [page([decorated])] }, { devicePixelRatio: 1 });
    const rects = ctx.of('fillRect').slice(1); // [0] is the page background
    // 11pt → em ≈ 14.67px: underline ≈ baseline+1.47, strike ≈ baseline−3.96.
    expect(rects).toHaveLength(2);
    const [under, strike] = rects;
    expect(under.args[0]).toBe(20);
    expect(under.args[2]).toBe(30); // layout width, not re-measured
    expect(under.args[1] as number).toBeGreaterThan(32);
    expect(strike.args[1] as number).toBeLessThan(32);
  });

  it('draws declared table borders and paints cell content', () => {
    const ctx = new RecordingCtx();
    const p = {
      ...page([]),
      tables: [
        {
          x: 20,
          y: 20,
          width: 100,
          height: 16,
          borders: { top: true, bottom: true, left: true, right: true, insideH: true, insideV: true },
          cells: [
            { x: 20, y: 20, width: 100, height: 16, colspan: 1, rowspan: 1, lines: [helloLine] },
          ],
        },
      ],
    };
    new CanvasPainter(makeCanvas(ctx)).paint({ pages: [p] }, { devicePixelRatio: 1 });
    // single cell, all edges outer → 4 edges drawn in one path
    expect(ctx.of('lineTo')).toHaveLength(4);
    expect(ctx.of('stroke').length).toBeGreaterThan(0);
    expect(ctx.of('fillText')[0].args).toEqual(['Hello', 20, 32]);
  });

  it('paints tables WITHOUT borders when none are declared (OOXML default)', () => {
    const ctx = new RecordingCtx();
    const p = {
      ...page([]),
      tables: [
        {
          x: 20,
          y: 20,
          width: 100,
          height: 16,
          cells: [
            { x: 20, y: 20, width: 100, height: 16, colspan: 1, rowspan: 1, lines: [helloLine] },
          ],
        },
      ],
    };
    new CanvasPainter(makeCanvas(ctx)).paint({ pages: [p] }, { devicePixelRatio: 1 });
    expect(ctx.of('lineTo')).toHaveLength(0); // no borders
    expect(ctx.of('strokeRect')).toHaveLength(1); // only the page border
    expect(ctx.of('fillText')[0].args).toEqual(['Hello', 20, 32]); // content intact
  });

  it('draws selection under the text and the caret on top, page-offset', () => {
    const ctx = new RecordingCtx();
    new CanvasPainter(makeCanvas(ctx)).paint(
      { pages: [page([]), page([helloLine], 1)] },
      {
        devicePixelRatio: 1,
        pageGap: 10,
        selection: [{ pageIndex: 1, x: 20, y: 20, width: 30, height: 16 }],
        caret: { pageIndex: 1, x: 50, y: 20, height: 16 },
      },
    );
    // page 1 starts at y = 310; selection rect lands at 330, caret at 330.
    const fills = ctx.of('fillRect');
    const sel = fills.find((c) => c.fillStyle.startsWith('rgba'));
    expect(sel?.args).toEqual([20, 330, 30, 16]);
    const caret = fills[fills.length - 1];
    expect(caret.args).toEqual([50, 330, 1.5, 16]);
    // selection painted before the line's text, caret after.
    const textIdx = ctx.calls.findIndex((c) => c.method === 'fillText');
    expect(ctx.calls.indexOf(sel as never)).toBeLessThan(textIdx);
    expect(ctx.calls.indexOf(caret)).toBeGreaterThan(textIdx);
  });

  it('routes caret/selection to the overlay canvas when one is provided', () => {
    const content = new RecordingCtx();
    const overlay = new RecordingCtx();
    const painter = new CanvasPainter(makeCanvas(content), makeCanvas(overlay));
    painter.paint(
      { pages: [page([helloLine])] },
      { devicePixelRatio: 1, caret: { pageIndex: 0, x: 50, y: 20, height: 16 } },
    );
    // content canvas: text only — no caret rect, no selection fill.
    expect(content.of('fillText')).toHaveLength(1);
    expect(content.of('fillRect')).toHaveLength(1); // page background only
    // overlay canvas: cleared + caret drawn, no text.
    expect(overlay.of('fillText')).toHaveLength(0);
    expect(overlay.of('fillRect')[0].args).toEqual([50, 20, 1.5, 16]);

    // paintOverlay redraws ONLY the overlay (content untouched).
    const contentCalls = content.calls.length;
    painter.paintOverlay({
      caret: null,
      selection: [{ pageIndex: 0, x: 20, y: 20, width: 30, height: 16 }],
    });
    expect(content.calls.length).toBe(contentCalls);
    const sel = overlay.of('fillRect').at(-1);
    expect(sel?.args).toEqual([20, 20, 30, 16]);
    expect(sel?.fillStyle.startsWith('rgba')).toBe(true);
  });

  it('skips the backing-store resize when dimensions are unchanged', () => {
    const ctx = new RecordingCtx();
    let widthSets = 0;
    const canvas = makeCanvas(ctx) as unknown as { width: number };
    let w = 0;
    Object.defineProperty(canvas, 'width', {
      get: () => w,
      set: (v: number) => {
        widthSets++;
        w = v;
      },
    });
    const painter = new CanvasPainter(canvas as unknown as HTMLCanvasElement);
    const layout: ResolvedLayout = { pages: [page([helloLine])] };
    painter.paint(layout, { devicePixelRatio: 1 });
    expect(widthSets).toBe(1);
    painter.paint(layout, { devicePixelRatio: 1 }); // same size → no reassignment
    expect(widthSets).toBe(1);
  });

  it('paints page floats behind the text once loaded', async () => {
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
      const ctx = new RecordingCtx();
      new CanvasPainter(makeCanvas(ctx)).paint(
        { pages: [{ ...page([helloLine]), floats: [{ x: 140, y: 20, width: 80, height: 40, src: 'f' }] }] },
        { devicePixelRatio: 1 },
      );
      await new Promise((r) => setTimeout(r, 0));
      const draw = ctx.of('drawImage')[0];
      expect(draw.args.slice(1)).toEqual([140, 20, 80, 40]);
      // behind the text: drawImage precedes the repaint's fillText
      const lastDraw = ctx.calls.lastIndexOf(ctx.of('drawImage')[0]);
      const lastText = ctx.calls.lastIndexOf(ctx.of('fillText').at(-1) as never);
      expect(lastDraw).toBeLessThan(lastText);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('substitutes live page numbers into chrome field segments', () => {
    const ctx = new RecordingCtx();
    const fieldLine = {
      ...helloLine,
      segments: [
        { x: 20, text: '1', font: font(), field: 'pageNumber' as const, width: 10 },
        { x: 35, text: '1', font: font(), field: 'pageCount' as const, width: 10 },
      ],
    };
    new CanvasPainter(makeCanvas(ctx)).paint(
      { pages: [page([]), page([], 1)], pageFooter: { lines: [fieldLine], tables: [], height: 16 } },
      { devicePixelRatio: 1, pageGap: 10 },
    );
    const texts = ctx.of('fillText').map((c) => c.args[0]);
    expect(texts).toEqual(['1', '2', '2', '2']); // page1: 1/2 — page2: 2/2
  });

  it('stamps page chrome (header/footer) onto every page', () => {
    const ctx = new RecordingCtx();
    const headerLine = { ...helloLine, y: 5, segments: [{ x: 20, text: 'hdr', font: font() }] };
    new CanvasPainter(makeCanvas(ctx)).paint(
      {
        pages: [page([]), page([], 1)],
        pageHeader: { lines: [headerLine], tables: [], height: 16 },
      },
      { devicePixelRatio: 1, pageGap: 10 },
    );
    const hdrCalls = ctx.of('fillText').filter((c) => c.args[0] === 'hdr');
    expect(hdrCalls).toHaveLength(2); // once per page
    expect(hdrCalls.map((c) => c.args[2])).toEqual([5 + 12, 310 + 5 + 12]); // y + baseline, page-offset
  });

  it('virtualizes pages outside the viewport (background only)', () => {
    const ctx = new RecordingCtx();
    const painter = new CanvasPainter(makeCanvas(ctx));
    const layout: ResolvedLayout = { pages: [page([]), page([helloLine], 1)] };
    // Page 1 spans y 310..610 (gap 10). Viewport 0..50 + 200 margin → hidden.
    painter.paint(layout, { devicePixelRatio: 1, pageGap: 10, viewport: { top: 0, height: 50 } });
    expect(ctx.of('fillText')).toHaveLength(0);
    expect(ctx.of('fillRect')).toHaveLength(2); // both page backgrounds

    // Scrolled down: viewport reaches page 1 → its text paints.
    painter.paint(layout, { devicePixelRatio: 1, pageGap: 10, viewport: { top: 320, height: 50 } });
    expect(ctx.of('fillText')).toHaveLength(1);
  });

  it('maps canvas coords to page-local coords and back (zoom + gap aware)', () => {
    const ctx = new RecordingCtx();
    const painter = new CanvasPainter(makeCanvas(ctx));
    painter.paint({ pages: [page([]), page([], 1)] }, { devicePixelRatio: 1, zoom: 2, pageGap: 10 });
    // CSS y = (300 + 10 + 5) * 2 = 630 → page 1, y = 5.
    expect(painter.canvasToPage(40, 630)).toEqual({ pageIndex: 1, x: 20, y: 5 });
    // Gap point clamps to the nearer edge (here: bottom of page 0).
    expect(painter.canvasToPage(0, (300 + 2) * 2)).toEqual({ pageIndex: 0, x: 0, y: 300 });
    expect(painter.pageToCanvas({ pageIndex: 1, x: 20, y: 5 })).toEqual({ x: 40, y: 630 });
  });

  it('skips images where Image is unavailable, draws after async load', async () => {
    const ctx = new RecordingCtx();
    const imageLine = {
      ...helloLine,
      segments: [],
      images: [{ x: 30, src: 'data:img', width: 40, height: 10 }],
    };
    const layout: ResolvedLayout = { pages: [page([imageLine])] };

    // Node (no Image): paint must not throw and must not draw.
    new CanvasPainter(makeCanvas(ctx)).paint(layout, { devicePixelRatio: 1 });
    expect(ctx.of('drawImage')).toHaveLength(0);

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
      const ctx2 = new RecordingCtx();
      new CanvasPainter(makeCanvas(ctx2)).paint(layout, { devicePixelRatio: 1 });
      expect(ctx2.of('drawImage')).toHaveLength(0); // not loaded yet
      await new Promise((r) => setTimeout(r, 0));
      const draw = ctx2.of('drawImage')[0];
      // x, baselineY (20 + 12) − height 10 = 22, w, h
      expect(draw.args.slice(1)).toEqual([30, 22, 40, 10]);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
