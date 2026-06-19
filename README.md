# bapbong

> A **canvas-rendered DOCX editor** for the browser — like [[ref]](https://github.com/[ref]-dev/[ref]),
> but the presentation/render layer paints to **HTML Canvas 2D** instead of the DOM.

**Status: active.** Import → layout → canvas paint → editing all work end-to-end
in the playground: DOCX import (text/marks, lists, tables, images, multi-column,
headers/footers, footnotes, comments), paginated canvas rendering, caret &
selection, IME typing, and a full comment system (4 view modes, threading,
resolve, @mention, threaded-comment round-trip on import). See the roadmap in
[`PLAN.md`](PLAN.md).

## Why canvas?

DOM/`contenteditable` editors fight the browser for Word-grade fidelity and true
pagination. Rendering to canvas gives pixel-accurate layout and a real page model —
the path Google Docs and OnlyOffice took. The hard parts (line-break/pagination
math, a hidden ProseMirror editor as the IME/undo input sink, caret/selection
overlay) are independent of the paint target; bapbong swaps the **painter** to canvas.

## Pipeline

```
.docx ─importDocx→ ProseMirror doc ─layout()→ ResolvedLayout ─CanvasPainter→ <canvas>
 (docx)             (model)          (layout-engine)            (painter-canvas)
                                      ↑ measuring                hidden ProseMirror (input-bridge)
                                      + contracts (types)        + caret/selection (selection)
```

## Packages

All packages publish under the `@shadow-garden` scope with the `bapbong-` prefix.
Module boundaries are enforced by Nx scope tags (see `eslint.config.mjs`).

| Package | Scope | Purpose |
|---|---|---|
| `@shadow-garden/bapbong-contracts` | `pure` | Shared types (`FlowParagraph`, `ResolvedLayout`, …) |
| `@shadow-garden/bapbong-model` | `model` | ProseMirror schema + list numbering |
| `@shadow-garden/bapbong-docx` | `io` | DOCX (OOXML) import → model |
| `@shadow-garden/bapbong-measuring` | `measuring` | Text measurement + font-metrics cache |
| `@shadow-garden/bapbong-layout-engine` | `engine` | Line-break + pagination → `ResolvedLayout` |
| `@shadow-garden/bapbong-painter-canvas` | `painter` | Canvas 2D renderer (the core differentiator) |
| `@shadow-garden/bapbong-selection` | `selection` | Caret/selection math + hit-testing |
| `@shadow-garden/bapbong-input-bridge` | `input` | Hidden ProseMirror (IME/undo) + comment authoring |
| `@shadow-garden/bapbong-editor` | `app` | Umbrella entry point / public API (in progress) |

[`apps/playground`](apps/playground) is the reference app that wires everything
together and is where features are dogfooded.

## Tech stack

- **Nx** monorepo (task caching, `nx affected`, module-boundary enforcement) on **pnpm**
- **TypeScript**, libraries built via `@nx/esbuild` (ESM), tests via **Vitest**
- Editing on **ProseMirror**; DOCX unzip via **jszip**
- Versioning/publish via **`nx release`** (conventional commits)

## Development

```sh
pnpm install

# Run the playground (the reference app)
pnpm exec nx serve playground

# Build / test / lint (one project or all)
pnpm exec nx build @shadow-garden/bapbong-layout-engine
pnpm exec nx run-many -t lint test build

# Only what changed vs main
pnpm exec nx affected -t lint test build

# Visualize the project graph
pnpm exec nx graph
```

## Releasing

```sh
pnpm exec nx release            # version + changelog + publish (conventional commits)
pnpm exec nx release --dry-run  # preview without releasing
```

## License

MIT
