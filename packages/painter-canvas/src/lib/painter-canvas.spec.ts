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

  it('strokes table cell borders and paints cell content', () => {
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
    const border = ctx.of('strokeRect')[1]; // [0] is the page border
    expect(border.args).toEqual([20.5, 20.5, 100, 16]);
    expect(border.strokeStyle).toBe('#b0b0b0');
    expect(ctx.of('fillText')[0].args).toEqual(['Hello', 20, 32]);
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
