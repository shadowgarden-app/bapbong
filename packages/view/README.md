# @shadow-garden/bapbong-view

The **render core + read-only viewer** for bapbong. It loads a `.docx`, lays it
out, and paints it to a virtualized `<canvas>` page-stack — with zoom, text
selection + copy, and an accessibility mirror — but **no editing** and **no
input-bridge**, so the preview bundle never pulls in the ProseMirror editing
surface.

This is the **preview tier** of the 3-tier architecture: headless
(`@shadow-garden/bapbong-headless`) / preview (this) / full editor
(`@shadow-garden/bapbong-editor`). The editor composes the same `RenderCore`
exported here, so layout/paint live in exactly one place.

## Install

```sh
pnpm add @shadow-garden/bapbong-view
```

## `BapbongView` — read-only viewer

Give it a scrollable host element (`overflow: auto` + a bounded height); it
renders the document inside.

```ts
import { BapbongView } from '@shadow-garden/bapbong-view';

const view = new BapbongView(hostEl);     // hostEl: overflow:auto, fixed height
await view.loadDocx(bytes);               // ArrayBuffer of a .docx
view.setZoom(1.25);
view.onChange((c) => console.log(c.pageCount));
// view.print();        // render every page → one image per sheet
// view.destroy();
```

Constructor options (`BapbongViewOptions`):

| Option | Default | Meaning |
| --- | --- | --- |
| `viewport` | the host | Scroll container for virtualization / scroll-into-view |
| `zoom` | `1` | Initial zoom factor |
| `selectable` | `true` | Drag to select text, double-click word, Ctrl/⌘+A, Ctrl/⌘+C copy |
| `a11y` | `true` | Build a visually-hidden ARIA mirror for screen readers |
| `a11yLabel` | `"Document content"` | `aria-label` for the mirror |

Methods: `loadDocx` · `scrollToPos` · `setZoom` / `getZoom` · `print` ·
`exportDocx` · `onChange` · `destroy`; getters `pageCount` · `schema`.

## `RenderCore` — the shared render engine

Lower-level than `BapbongView` (no host wrapper); the editor composes it. It
owns load → layout → paint → scroll/zoom/virtualize plus the geometry queries
(`caretRect`, `selectionRects`, `posAtEvent`, `pageToCanvas`, …) and holds a
**read-only ProseMirror doc** (a `Node`, not an `EditorState`). You normally use
`BapbongView`; reach for `RenderCore` only to build a custom surface.

## Notes

- **Selection is painted on the canvas** (using the same layout geometry the
  painter uses, so it aligns at any zoom) and copies via the document model.
  Because it isn't a native browser selection, the browser's right-click "Copy"
  and find-in-page don't apply — `Ctrl/⌘+C` does.
- **Print is rasterized** (each page is a high-res PNG, one per sheet) — pixel
  faithful to the screen, but the printout's text isn't selectable. Vector/PDF
  output is a planned separate path (`@shadow-garden/bapbong-pdf`).
- Deps: `contracts`, `model`, `docx`, `measuring`, `layout-engine`,
  `painter-canvas`, `selection`, `a11y` — **not** `input-bridge`.
