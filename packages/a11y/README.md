# @shadow-garden/bapbong-a11y

Accessibility for the canvas-rendered editor. A `<canvas>` is **opaque to screen
readers** — there is no text in the accessibility tree. `A11yMirror` keeps a
parallel, visually-hidden DOM subtree that mirrors the document so assistive
tech can read and navigate it.

The mirror is built from the document's **own schema `toDOM`** (via
ProseMirror's `DOMSerializer`), so it is faithfully semantic: paragraphs,
lists, tables, and real `<a href>` links — the same structure the model would
export.

```ts
import { A11yMirror } from '@shadow-garden/bapbong-a11y';

const mirror = new A11yMirror(hostEl, { label: 'My document' });
mirror.update(doc);   // call on each document change — debounced internally
// …
mirror.destroy();
```

Wired automatically by `RenderCore` (so both `@shadow-garden/bapbong-editor` and
`@shadow-garden/bapbong-view` are accessible out of the box); you rarely
construct it by hand.

## Notes

- **Visually hidden, AT-visible.** Uses the standard `sr-only` clip recipe — the
  mirror never shows on screen but stays in the accessibility tree.
- **Debounced.** `update()` coalesces rapid edits (200 ms by default) so typing
  stays snappy; the mirror is for reading/navigation, not live keystroke echo.
- **Light.** Inline-image `src` (base64 data-URLs) is stripped — assistive tech
  reads the `alt`, and duplicating data-URLs would bloat the DOM.
