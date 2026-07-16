# bapbong

[![npm](https://img.shields.io/npm/v/@shadow-garden/bapbong-editor?label=npm%20%C2%B7%20%40shadow-garden&color=e8722a)](https://www.npmjs.com/org/shadow-garden)
[![CI](https://github.com/shadowgarden-app/bapbong/actions/workflows/ci.yml/badge.svg)](https://github.com/shadowgarden-app/bapbong/actions/workflows/ci.yml)
[![license](https://img.shields.io/badge/license-MIT-3da639)](LICENSE)
[![node](https://img.shields.io/badge/types-TypeScript-3178c6)](tsconfig.base.json)

> A **canvas-rendered DOCX editor** engine for the browser — the presentation
> layer paints to **HTML Canvas 2D** instead of the DOM, for pixel-accurate,
> truly paginated documents.

**Status: active.** Import → layout → canvas paint → editing all work end-to-end
in the playground: DOCX import (text/marks, lists, tables, images, multi-column,
headers/footers, footnotes, comment marks), paginated canvas rendering, caret &
selection, IME typing, find & replace, and DOCX export round-trip.

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

All packages publish under the [`@shadow-garden`](https://www.npmjs.com/org/shadow-garden)
scope with the `bapbong-` prefix. Module boundaries are enforced by Nx scope
tags (see `eslint.config.mjs`). Package names link to the source; badges link to npm.

| Package                                             | npm                                                                                                                                                          | Purpose                                                                |
| --------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------- |
| [`bapbong-contracts`](packages/contracts)           | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-contracts?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-contracts)           | Shared types & plugin contracts (`FlowParagraph`, `ResolvedLayout`, …) |
| [`bapbong-model`](packages/model)                   | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-model?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-model)                   | ProseMirror schema + list numbering                                    |
| [`bapbong-docx`](packages/docx)                     | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-docx?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-docx)                     | DOCX (OOXML) import / export                                           |
| [`bapbong-measuring`](packages/measuring)           | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-measuring?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-measuring)           | Text measurement + font-metrics registry                               |
| [`bapbong-layout-engine`](packages/layout-engine)   | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-layout-engine?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-layout-engine)   | Line-breaking + pagination → `ResolvedLayout`                          |
| [`bapbong-painter-canvas`](packages/painter-canvas) | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-painter-canvas?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-painter-canvas) | Canvas 2D renderer (the core differentiator)                           |
| [`bapbong-selection`](packages/selection)           | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-selection?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-selection)           | Caret/selection math + hit-testing over `ResolvedLayout`               |
| [`bapbong-input-bridge`](packages/input-bridge)     | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-input-bridge?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-input-bridge)     | Hidden ProseMirror editor as the input/IME sink                        |
| [`bapbong-editor`](packages/editor)                 | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-editor?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-editor)                 | The umbrella editor: render/edit loop + public API                     |
| [`bapbong-view`](packages/view)                     | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-view?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-view)                     | Read-only render core (zoom + page virtualization, no editing)         |
| [`bapbong-ui`](packages/ui)                         | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-ui?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-ui)                         | Framework-agnostic toolbar/menubar/dialogs bound to `editor.commands`  |
| [`bapbong-commands`](packages/commands)             | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-commands?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-commands)             | Headless commands + queries + registry (UI and backends share ops)     |
| [`bapbong-headless`](packages/headless)             | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-headless?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-headless)             | Server-side façade: import → edit → export, no DOM                     |
| [`bapbong-mcp`](packages/mcp)                       | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-mcp?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-mcp)                       | MCP server — AI agents edit documents over a `DocumentSession` port    |
| [`bapbong-a11y`](packages/a11y)                     | [![npm](https://img.shields.io/npm/v/%40shadow-garden%2Fbapbong-a11y?label=)](https://www.npmjs.com/package/@shadow-garden/bapbong-a11y)                     | Screen-reader mirror of the canvas document (ARIA DOM)                 |

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
