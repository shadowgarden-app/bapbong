# @shadow-garden/bapbong-painter-canvas

Paints a `ResolvedLayout` (from `@shadow-garden/bapbong-layout-engine`) onto
HTML `<canvas>`. bapbong's core differentiator: the painter consumes
pre-computed coordinates only — **it never measures text at paint time**.

- **Scope:** `scope:painter`
- **Depends on:** `@shadow-garden/bapbong-contracts`

## What it does

- **One `<canvas>` per page** (virtualized): only on-screen pages get a backing
  store, sidestepping the browser's ~65535px max canvas dimension on long docs.
- Draws text runs, images, tables (borders/shading), floats, footnotes, and
  header/footer chrome; tints commented ranges (skipping resolved threads).
- `pageToCanvas({ pageIndex, x, y })` maps page coords → container pixels for the
  caret/selection overlay and comment anchors.

```ts
import { CanvasPainter } from '@shadow-garden/bapbong-painter-canvas';

const painter = new CanvasPainter(container, { createCanvas });
painter.paint(resolvedLayout, { zoom: 1, resolvedComments });
painter.paintOverlay({ caret, selection }); // caret/selection, no relayout
```

Also: `PaintOptions`, `PainterDeps`.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-painter-canvas
pnpm nx test  @shadow-garden/bapbong-painter-canvas
```
