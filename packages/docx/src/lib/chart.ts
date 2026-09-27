/**
 * What a chart paints while its drawing is carried rather than modelled.
 *
 * A DrawingML chart (`c:chart`) holds no picture: Word renders it from the
 * chart part's series each time, and a file saved by Word carries no
 * fallback image for it. Until charts are rendered, the box shows where the
 * chart is and how much room it takes — a flat neutral frame with a small
 * bar-chart glyph, the same look as the painter's "cannot display" picture
 * stand-in, so the page reads the same and the text around the chart
 * flows exactly as in Word.
 *
 * Pure: a size in, a display list in px of the box out.
 */
import type {
  VectorImageSpec,
  VectorOp,
} from '@shadow-garden/bapbong-contracts';

const FILL = '#f2f2f2';
const BORDER = '#b8b8b8';
const GLYPH = '#9a9a9a';

const rect = (
  x: number,
  y: number,
  w: number,
  h: number,
  fill: string,
  stroke?: string,
): VectorOp => ({
  kind: 'polygon',
  points: [
    { x, y },
    { x: x + w, y },
    { x: x + w, y: y + h },
    { x, y: y + h },
  ],
  fill,
  ...(stroke && { stroke, strokeWidth: 1 }),
});

export function chartPlaceholder(
  width: number,
  height: number,
): VectorImageSpec {
  const ops: VectorOp[] = [
    rect(
      0.5,
      0.5,
      Math.max(0, width - 1),
      Math.max(0, height - 1),
      FILL,
      BORDER,
    ),
  ];
  // The glyph: three bars of rising height on a baseline, centred, half the
  // box's shorter side — too small a box keeps just the frame.
  const s = Math.min(width, height) * 0.5;
  if (s >= 8) {
    const gx = (width - s) / 2;
    const gy = (height - s) / 2;
    const bar = s * 0.2;
    const gap = (s - 3 * bar) / 4;
    [0.45, 0.75, 0.6].forEach((h, i) =>
      ops.push(
        rect(gx + gap + i * (bar + gap), gy + s * (1 - h), bar, s * h, GLYPH),
      ),
    );
    ops.push({
      kind: 'line',
      x1: gx,
      y1: gy + s,
      x2: gx + s,
      y2: gy + s,
      width: Math.max(1, s / 24),
      color: GLYPH,
    });
  }
  return { width, height, ops };
}
