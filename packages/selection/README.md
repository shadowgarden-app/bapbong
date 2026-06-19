# @shadow-garden/bapbong-selection

Pure caret/selection math over a `ResolvedLayout`. No DOM, no ProseMirror
dependency — positions come from the layout (`LayoutSegment.pos`,
`LayoutLine.from/to`) and text widths from an injected `MeasureText`.

- **Scope:** `scope:selection`
- **Depends on:** `@shadow-garden/bapbong-contracts`

## Public API

- `hitTest(layout, point, measure)` — page-local `(x, y)` → ProseMirror position
- `caretRect(layout, pos, measure)` — PM position → caret rectangle (with page index)
- `selectionRects(layout, from, to, measure)` — selection range → highlight rects
- `verticalCaret(layout, pos, dir, goalX, measure)` — Up/Down caret motion

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-selection
pnpm nx test  @shadow-garden/bapbong-selection
```
