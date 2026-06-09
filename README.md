# bapbong

> A **canvas-rendered DOCX editor** for the browser — like [[ref]](https://github.com/[ref]-dev/[ref]),
> but the presentation/render layer paints to **HTML Canvas 2D** instead of the DOM.

🚧 **Status: early / work-in-progress.** Currently at M0 (workspace scaffolding). The
canvas renderer, layout engine, and editing layer are not built yet — see the roadmap.

## Why canvas?

DOM/`contenteditable` editors fight the browser for Word-grade fidelity and true
pagination. Rendering to canvas gives pixel-accurate layout and real page model — the
path Google Docs and OnlyOffice took. The hard parts (line-break/pagination math, a
hidden ProseMirror editor as the IME/undo input sink, caret/selection overlay) live
upstream and are independent of the paint target; bapbong swaps the **painter** to canvas.

## Packages

All packages publish under the `@shadow-garden` scope with the `bapbong-` prefix.

| Package | Status | Purpose |
|---|---|---|
| `@shadow-garden/bapbong-editor` | ✅ scaffolded | Umbrella entry point / public API |
| `@shadow-garden/bapbong-contracts` | planned | Shared types (`FlowBlock`, `ResolvedLayout`, …) |
| `@shadow-garden/bapbong-docx` | planned | DOCX (OOXML) import + export |
| `@shadow-garden/bapbong-word-layout` | planned | Pure paragraph/list/tab layout math |
| `@shadow-garden/bapbong-measuring` | planned | Text measurement + font-metrics cache |
| `@shadow-garden/bapbong-layout-engine` | planned | Line-break + pagination → `ResolvedLayout` |
| `@shadow-garden/bapbong-painter-canvas` | planned | Canvas 2D renderer (the core differentiator) |
| `@shadow-garden/bapbong-input-bridge` | planned | Hidden ProseMirror + IME forwarding |
| `@shadow-garden/bapbong-selection` | planned | Caret/selection overlay + hit-testing |
| `@shadow-garden/bapbong-angular` | planned | Angular component wrapper |

## Tech stack

- **Nx** monorepo (task caching, `nx affected`, module-boundary enforcement) on **pnpm**
- **TypeScript**, build via `@nx/rollup` (ESM), tests via **Vitest**
- Versioning/publish via **`nx release`** (conventional commits)

## Development

```sh
pnpm install

# Build / test / lint / typecheck (one project or all)
pnpm exec nx build @shadow-garden/bapbong-editor
pnpm exec nx run-many -t lint test build typecheck

# Only what changed vs main
pnpm exec nx affected -t lint test build typecheck

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
