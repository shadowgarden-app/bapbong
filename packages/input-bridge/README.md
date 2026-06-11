# @shadow-garden/bapbong-input-bridge

The canvas can't receive text input — this package wraps a **hidden
ProseMirror `EditorView`** (a real contenteditable, visually invisible) that
acts as the input sink. The browser delivers keyboard and **IME composition**
(Vietnamese, CJK, …) to it; ProseMirror owns the document model, undo history
and clipboard. Every transaction triggers `onUpdate`, where the host app
re-runs the layout engine and repaints the canvas.

```ts
const bridge = new InputBridge({
  doc,
  keys: { ArrowUp: myLayoutAwareUp, ArrowDown: myLayoutAwareDown },
  onUpdate: (state) => repaint(state),
});
container.appendChild(bridge.dom);
bridge.setSelection(pos);
bridge.focus();
bridge.place(caretCssX, caretCssY, caretCssHeight); // anchors the IME popup
```
