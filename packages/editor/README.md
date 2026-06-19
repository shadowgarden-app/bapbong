# @shadow-garden/bapbong-editor

Umbrella entry point for bapbong — the package a host app installs to get the
canvas-rendered DOCX editor. It wires the pipeline together (import → model →
layout → paint → input/selection) behind one public API.

- **Scope:** `scope:app` (may depend on every other package)
- **Depends on:** the bapbong packages (`docx`, `model`, `layout-engine`, `measuring`, `painter-canvas`, `selection`, `input-bridge`, `contracts`)

> **Status:** thin scaffold. The full assembled API is still being extracted from
> the reference implementation in [`apps/playground`](../../apps/playground),
> which currently drives the whole pipeline directly.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-editor
pnpm nx test  @shadow-garden/bapbong-editor
```
