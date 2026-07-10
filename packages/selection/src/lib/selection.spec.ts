import type {
  FontSpec,
  LayoutLine,
  MeasureText,
  ResolvedLayout,
  ResolvedPage,
} from '@shadow-garden/bapbong-contracts';
import { caretRect, hitTest, imageAtPoint, selectionRects, verticalCaret } from './selection.js';

const F: FontSpec = { family: 'Arial', sizePt: 10, bold: false, italic: false };
/** 10px per character. */
const measure: MeasureText = (text) => text.length * 10;

const seg = (x: number, text: string, pos: number) => ({ x, text, font: F, pos });

const line = (
  y: number,
  segments: ReturnType<typeof seg>[],
  from: number,
  to: number,
  over: Partial<LayoutLine> = {},
): LayoutLine => ({
  x: 20,
  y,
  width: 200,
  height: 16,
  baseline: 12,
  segments,
  from,
  to,
  ...over,
});

const page = (lines: LayoutLine[], index = 0, over: Partial<ResolvedPage> = {}): ResolvedPage => ({
  index,
  width: 240,
  height: 100,
  lines,
  ...over,
});

// "hello world" wrapped as "hello" (pos 1..6) / "world" (pos 7..12);
// the space at pos 6 is swallowed by the wrap.
const twoLines: ResolvedLayout = {
  pages: [
    page([
      line(20, [seg(20, 'hello', 1)], 1, 6),
      line(36, [seg(20, 'world', 7)], 7, 12),
    ]),
  ],
};

describe('hitTest', () => {
  it('maps x to the nearest character boundary', () => {
    // chars span 20-30, 30-40, …; the midpoint rule decides the boundary.
    expect(hitTest(twoLines, { pageIndex: 0, x: 20, y: 25 }, measure)).toBe(1);
    expect(hitTest(twoLines, { pageIndex: 0, x: 34, y: 25 }, measure)).toBe(2); // 34 < mid(35) → before 2nd char
    expect(hitTest(twoLines, { pageIndex: 0, x: 36, y: 25 }, measure)).toBe(3); // 36 > mid(35) → after it
    expect(hitTest(twoLines, { pageIndex: 0, x: 300, y: 25 }, measure)).toBe(6); // right of line → end
  });

  it('clamps to the vertically nearest line', () => {
    expect(hitTest(twoLines, { pageIndex: 0, x: 20, y: 90 }, measure)).toBe(7); // below all → line 2
    expect(hitTest(twoLines, { pageIndex: 0, x: 20, y: 5 }, measure)).toBe(1); // above all → line 1
  });

  it('treats an image as an atom with a midpoint rule', () => {
    const l = line(20, [seg(20, 'ab', 1)], 1, 4, {
      images: [{ x: 40, src: 'i', width: 30, height: 10, pos: 3 }],
    });
    const layout: ResolvedLayout = { pages: [page([l])] };
    expect(hitTest(layout, { pageIndex: 0, x: 50, y: 25 }, measure)).toBe(3); // left half
    expect(hitTest(layout, { pageIndex: 0, x: 62, y: 25 }, measure)).toBe(4); // right half
  });

  it('picks the side-by-side cell line nearest in x', () => {
    const left = line(20, [seg(20, 'aa', 1)], 1, 3, { x: 20, width: 80 });
    const right = line(20, [seg(120, 'bb', 10)], 10, 12, { x: 120, width: 80 });
    const layout: ResolvedLayout = {
      pages: [
        page([], 0, {
          tables: [
            {
              x: 20,
              y: 20,
              width: 180,
              height: 16,
              cells: [
                { x: 20, y: 20, width: 80, height: 16, colspan: 1, rowspan: 1, lines: [left] },
                { x: 120, y: 20, width: 80, height: 16, colspan: 1, rowspan: 1, lines: [right] },
              ],
            },
          ],
        }),
      ],
    };
    expect(hitTest(layout, { pageIndex: 0, x: 22, y: 25 }, measure)).toBe(1);
    expect(hitTest(layout, { pageIndex: 0, x: 122, y: 25 }, measure)).toBe(10);
  });
});

describe('caretRect', () => {
  it('positions the caret inside a segment', () => {
    expect(caretRect(twoLines, 3, measure)).toEqual({ pageIndex: 0, x: 40, y: 20, height: 16 });
  });

  it('clamps a swallowed wrap-space position to the line end', () => {
    // pos 6 is the space dropped at the wrap → caret at end of "hello".
    expect(caretRect(twoLines, 6, measure)).toEqual({ pageIndex: 0, x: 70, y: 20, height: 16 });
  });

  it('places the caret on an empty line', () => {
    const layout: ResolvedLayout = { pages: [page([line(20, [], 5, 5)])] };
    expect(caretRect(layout, 5, measure)).toEqual({ pageIndex: 0, x: 20, y: 20, height: 16 });
  });
});

describe('selectionRects', () => {
  it('builds one rect per touched line', () => {
    const rects = selectionRects(twoLines, 3, 9, measure);
    expect(rects).toEqual([
      { pageIndex: 0, x: 40, y: 20, width: 30, height: 16 }, // "llo"
      { pageIndex: 0, x: 20, y: 36, width: 20, height: 16 }, // "wo"
    ]);
  });

  it('spans pages and marks empty lines with a stub', () => {
    const layout: ResolvedLayout = {
      pages: [
        page([line(20, [seg(20, 'aa', 1)], 1, 3)], 0),
        page([line(20, [], 5, 5), line(36, [seg(20, 'bb', 7)], 7, 9)], 1),
      ],
    };
    const rects = selectionRects(layout, 1, 9, measure);
    expect(rects.map((r) => r.pageIndex)).toEqual([0, 1, 1]);
    expect(rects[1].width).toBe(6); // empty-line stub
  });
});

describe('verticalCaret', () => {
  it('moves down to the nearest x on the next line', () => {
    expect(verticalCaret(twoLines, 3, 1, 40, measure)).toBe(9); // x=40 on "world"
    expect(verticalCaret(twoLines, 9, -1, 40, measure)).toBe(3);
  });

  it('returns null at the document edges', () => {
    expect(verticalCaret(twoLines, 3, -1, 40, measure)).toBeNull();
    expect(verticalCaret(twoLines, 9, 1, 40, measure)).toBeNull();
  });

  it('crosses onto the adjacent page', () => {
    const layout: ResolvedLayout = {
      pages: [
        page([line(20, [seg(20, 'aa', 1)], 1, 3)], 0),
        page([line(20, [seg(20, 'bb', 7)], 7, 9)], 1),
      ],
    };
    expect(verticalCaret(layout, 2, 1, 30, measure)).toBe(8);
    expect(verticalCaret(layout, 8, -1, 30, measure)).toBe(2);
  });
});

describe('imageAtPoint', () => {
  const layout: ResolvedLayout = {
    pages: [
      page(
        [
          // Inline image on the line: x 60..100, bottom on the baseline
          // (line y 20 + baseline 12 = 32), height 10 → top 22.
          line(20, [seg(20, 'himg', 1)], 1, 7, {
            images: [{ x: 60, src: 'a', width: 40, height: 10, pos: 5 }],
          }),
        ],
        0,
        {
          floats: [
            { x: 120, y: 40, width: 50, height: 30, src: 'f1', pos: 9 },
            // Overlapping later float wins (painted on top). This one has no
            // pos (chrome-style) so it must be skipped even though it's on top.
            { x: 130, y: 45, width: 50, height: 30, src: 'deco' },
          ],
          tables: [
            {
              x: 20, y: 80, width: 100, height: 20,
              cells: [{
                x: 20, y: 80, width: 100, height: 20, colspan: 1, rowspan: 1,
                lines: [],
                floats: [{ x: 30, y: 85, width: 20, height: 10, src: 'cf', pos: 15 }],
              }],
            },
          ],
        },
      ),
    ],
  };
  it('hits an inline image via its baseline-anchored box', () => {
    const hit = imageAtPoint(layout, { pageIndex: 0, x: 70, y: 25 });
    expect(hit).toMatchObject({ pos: 5, kind: 'inline', rect: { x: 60, y: 22, width: 40, height: 10 } });
    // Just above the image box (line top band) → no hit.
    expect(imageAtPoint(layout, { pageIndex: 0, x: 70, y: 21 })).toBeNull();
  });

  it('hits floats (topmost with a pos) and cell floats', () => {
    const f = imageAtPoint(layout, { pageIndex: 0, x: 125, y: 50 });
    expect(f).toMatchObject({ pos: 9, kind: 'float' });
    // Point only inside the pos-less decoration float → skipped → null.
    expect(imageAtPoint(layout, { pageIndex: 0, x: 175, y: 60 })).toBeNull();
    const cf = imageAtPoint(layout, { pageIndex: 0, x: 35, y: 90 });
    expect(cf).toMatchObject({ pos: 15, kind: 'float', rect: { x: 30, y: 85 } });
  });

  it('prefers the float over an underlying inline image', () => {
    const both: ResolvedLayout = {
      pages: [page(
        [line(20, [seg(20, 'x', 1)], 1, 3, { images: [{ x: 20, src: 'a', width: 40, height: 10, pos: 2 }] })],
        0,
        { floats: [{ x: 20, y: 20, width: 40, height: 12, src: 'f', pos: 30 }] },
      )],
    };
    expect(imageAtPoint(both, { pageIndex: 0, x: 30, y: 26 })?.pos).toBe(30);
  });
});
