# @shadow-garden/bapbong-editor

The framework-agnostic core of bapbong — the package a host app installs to get
a canvas-rendered DOCX editor. It owns the whole render → edit loop (import →
model → layout → paint → input/selection → export) and the per-page `<canvas>`
stack behind one vanilla `BapbongEditor` class, with **no UI chrome of its own**.
A host (the Angular [`apps/playground`](../../apps/playground), a future React
adapter, etc.) wires comment UI / toolbars around it.

- **Scope:** `scope:app` (may depend on every other package)
- **Depends on:** the bapbong packages (`docx`, `layout-engine`, `measuring`, `painter-canvas`, `selection`, `input-bridge`, `contracts`)

## Usage

```ts
import { BapbongEditor } from '@shadow-garden/bapbong-editor';

// `stack` is the element the editor fills with one <canvas> per page; the
// optional `viewport` is the scroll container (defaults to .canvas-wrap).
const editor = new BapbongEditor(stackEl, { viewport: wrapEl });

editor.onChange(({ state, pageCount, docChanged }) => {
  // Mirror doc-derived UI: page count, comment threads (state.doc.attrs.comments),
  // selection. `docChanged` is false on selection-only cycles.
});
editor.onCaretPick((pos) => {
  // A pointerdown placed the caret at doc position `pos` (e.g. select the
  // comment whose range was clicked).
});

await editor.loadDocx(bytes);          // ArrayBuffer → render first frame
const out = await editor.exportDocx(); // Uint8Array, carries the source package
editor.focus();
editor.destroy();
```

### Public API

| Member | Purpose |
| --- | --- |
| `loadDocx(bytes)` | Import a `.docx`, lay it out, paint; resolves with the page-chrome keys. |
| `exportDocx()` | Serialize the edited doc back to `.docx` bytes (round-trips unmodelled parts). |
| `state` / `dispatch(tr)` | Read the `EditorState`; apply a ProseMirror transaction. |
| `onChange(cb)` / `onCaretPick(cb)` | Subscribe to render cycles / caret placement (each returns an unsubscribe). |
| `caretRect(pos)` / `pageToCanvas(pt)` | Geometry helpers for anchoring host UI (comment bubbles/cards). |
| `setSelection(from, to?)` / `selectWordAt(pos)` / `scrollToPos(pos)` | Move/scroll the selection. |
| `setSuppressedComments(ids)` | Comment ids to paint with no tint (resolved, or all when hidden). |
| `focus()` / `destroy()` | Focus the IME sink; tear down listeners + the hidden editor. |

Comment marks and threads ride on the editor's doc model (`state.doc.attrs`), so
the host stays in sync purely through `dispatch`/`state` plus the geometry
helpers — the editor itself is comment-UI agnostic.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-editor
pnpm nx test  @shadow-garden/bapbong-editor
```
