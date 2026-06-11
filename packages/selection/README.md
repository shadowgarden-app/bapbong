# @shadow-garden/bapbong-selection

Pure caret/selection math over a `ResolvedLayout`:

- `hitTest(layout, point, measure)` — page-local (x, y) → ProseMirror position
- `caretRect(layout, pos, measure)` — PM position → caret rectangle
- `selectionRects(layout, from, to, measure)` — selection → highlight rects
- `verticalCaret(layout, pos, dir, goalX, measure)` — Up/Down caret motion

No DOM, no ProseMirror dependency: positions are read from the layout
(`LayoutSegment.pos`, `LayoutLine.from/to`), text widths come from the
injected `MeasureText`.
