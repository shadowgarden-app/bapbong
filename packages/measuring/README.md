# @shadow-garden/bapbong-measuring

Text measurement + font-metrics, behind a small injectable interface. The layout
engine and selection math never touch the canvas directly — they call a
`MeasureText` / `MeasureMetrics` provided by this package, so they stay testable
in a headless (no-DOM) environment.

- **Scope:** `scope:measuring`
- **Depends on:** `@shadow-garden/bapbong-contracts`

## What it provides

- **`createCanvasMeasurer()` / `createCanvasMetrics()`** — real measurement via a
  shared offscreen `CanvasRenderingContext2D` (`ctx.measureText`,
  ascent/descent), with a per-font cache.
- **`createApproxMeasurer()` / `createApproxMetrics()`** — DOM-free approximations
  (average advance widths) for unit tests and SSR.
- **`fontToCss(spec)`** — `FontSpec` → CSS `font` shorthand.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-measuring
pnpm nx test  @shadow-garden/bapbong-measuring
```
