# @shadow-garden/bapbong-layout-engine

The heart of the renderer: turns a ProseMirror document into a **paint-ready
`ResolvedLayout`** — every glyph run, line, table cell, float and footnote with
absolute page coordinates. The painter then just draws; it never measures.

- **Scope:** `scope:engine`
- **Depends on:** `@shadow-garden/bapbong-contracts`, `@shadow-garden/bapbong-model` (measurement is injected via `MeasureText`)

## What it does

- `toFlowBlocks(doc)` — schema document → flow model (`FlowParagraph` / `FlowTable` / …).
- Line breaking, **pagination**, paragraph spacing/indent, tabs, hanging indents.
- **Multi-column** sections (with column-height balancing), **floats** (text wrap),
  page **headers/footers** (first/even variants), and **footnotes** reserved at the
  page bottom (including refs inside table cells).
- Threads `commentIds` / `footnoteRef` onto segments for the painter.

## Public API

```ts
import { layout, createLayoutCache } from '@shadow-garden/bapbong-layout-engine';

const cache = createLayoutCache();            // incremental table re-layout
const resolved = layout(doc, config, measureText, { cache });
```

Also: `layoutBlocks`, `LayoutCache`, `PageChrome`.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-layout-engine
pnpm nx test  @shadow-garden/bapbong-layout-engine
```
