# @shadow-garden/bapbong-input-bridge

The canvas can't receive text input, so this package wraps a **hidden
ProseMirror `EditorView`** (a real, visually-invisible contenteditable) as the
input sink. The browser delivers keyboard and **IME composition** (Vietnamese,
CJK, …) to it; ProseMirror owns the document, undo history and clipboard. Every
transaction fires `onUpdate`, where the host re-runs layout and repaints.

- **Scope:** `scope:input`
- **Depends on:** `@shadow-garden/bapbong-model`, `prosemirror-{state,view,model,commands,keymap,history,transform}`

## What it provides

- **`InputBridge`** — the hidden editor: `state`, `dispatch(tr)`, `setSelection`,
  `selectWordAt`, `place(x, y, h)` (anchors the IME popup at the caret), `focus`.
- **Layout-aware commands:** `moveCaretCommand`, `splitListItem`, `wordRangeAt`.
- **Comment authoring** (undoable transactions on `doc.attrs.comments`):
  `addCommentTr`, `replyCommentTr`, `resolveCommentTr`, `editCommentTr`,
  `deleteCommentTr`.
- **`CommentComposer`** — a small standalone PM editor for comment bodies, with an
  optional **@mention** suggestion plugin (`MentionHandlers` / `MentionUser` /
  `MentionCoords`), a `placeholder`, and `onEnter` / `onEscape` hooks.

```ts
const bridge = new InputBridge({ doc, keys: { ArrowUp, ArrowDown }, onUpdate: repaint });
container.appendChild(bridge.dom);
bridge.setSelection(pos); bridge.focus(); bridge.place(x, y, h);
```

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-input-bridge
pnpm nx test  @shadow-garden/bapbong-input-bridge
```
