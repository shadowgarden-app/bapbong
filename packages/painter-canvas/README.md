# @shadow-garden/bapbong-painter-canvas

Paints a `ResolvedLayout` (from `@shadow-garden/bapbong-layout-engine`) onto an
HTML `<canvas>`. This is bapbong's core differentiator: the painter consumes
pre-computed coordinates only — **it never measures text at paint time**.

```ts
import { CanvasPainter } from '@shadow-garden/bapbong-painter-canvas';

const painter = new CanvasPainter(canvas);
painter.paint(resolvedLayout, { zoom: 1 });
```
