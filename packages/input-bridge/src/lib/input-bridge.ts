import { baseKeymap } from 'prosemirror-commands';
import { history, redo, undo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState, TextSelection, type Command, type Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';

// Re-exported so hosts type against ONE prosemirror-state identity (mixing
// module resolutions across packages makes TS treat duplicates as unrelated).
export type { Command, EditorState, Transaction } from 'prosemirror-state';

/** A command that moves the caret to the position computed by `compute`
 *  (e.g. layout-aware ArrowUp/ArrowDown from bapbong-selection). Returns
 *  false — leaving the key to the next handler — when `compute` yields null. */
export function moveCaretCommand(compute: (state: EditorState) => number | null): Command {
  return (state, dispatch) => {
    const pos = compute(state);
    if (pos == null) return false;
    dispatch?.(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    return true;
  };
}

export interface InputBridgeOptions {
  /** The initial document (its schema drives the editor). */
  doc: PMNode;
  /** Extra bindings, checked before the base keymap — e.g. ArrowUp/ArrowDown
   *  wired to layout-aware caret motion from bapbong-selection. */
  keys?: Record<string, Command>;
  /** Called after every dispatched transaction (typing, IME composition
   *  steps, undo, selection changes). Re-layout + repaint here. */
  onUpdate: (state: EditorState, tr: Transaction) => void;
}

/** Editing state with history + base keymap; exported for headless tests. */
export function createEditingState(doc: PMNode, keys: Record<string, Command> = {}): EditorState {
  return EditorState.create({
    doc,
    plugins: [
      history(),
      keymap({ 'Mod-z': undo, 'Shift-Mod-z': redo, 'Mod-y': redo }),
      keymap(keys),
      keymap(baseKeymap),
    ],
  });
}

/**
 * Hidden ProseMirror editor acting as the canvas's input sink.
 *
 * The browser routes keyboard and IME composition into a real (but invisible)
 * contenteditable; ProseMirror keeps the model, history and clipboard. The
 * host positions `dom` at the canvas caret via `place()` so IME candidate
 * popups appear next to the visible (painted) caret.
 */
export class InputBridge {
  /** Host element of the hidden editor. Append it near the canvas. */
  readonly dom: HTMLElement;
  readonly view: EditorView;

  constructor(options: InputBridgeOptions) {
    this.dom = document.createElement('div');
    this.dom.className = 'bapbong-input-bridge';
    Object.assign(this.dom.style, {
      position: 'absolute',
      top: '0',
      left: '0',
      width: '1px',
      height: '1em',
      overflow: 'hidden',
      opacity: '0',
      // Keep it focusable but out of the way of canvas pointer events.
      pointerEvents: 'none',
      zIndex: '-1',
    } satisfies Partial<CSSStyleDeclaration>);

    this.view = new EditorView(this.dom, {
      state: createEditingState(options.doc, options.keys),
      dispatchTransaction: (tr) => {
        const state = this.view.state.apply(tr);
        this.view.updateState(state);
        options.onUpdate(state, tr);
      },
    });
  }

  get state(): EditorState {
    return this.view.state;
  }

  focus(): void {
    this.view.focus();
  }

  /** Set a text selection (collapsed caret or anchor→head range). Positions
   *  are clamped to the nearest valid text slots. */
  setSelection(anchor: number, head: number = anchor): void {
    const { doc } = this.view.state;
    const clamp = (p: number) => Math.max(0, Math.min(p, doc.content.size));
    const sel = TextSelection.between(doc.resolve(clamp(anchor)), doc.resolve(clamp(head)));
    this.view.dispatch(this.view.state.tr.setSelection(sel));
  }

  /** Move the hidden editor to the painted caret (CSS px, relative to the
   *  positioned ancestor) so IME popups anchor in the right place. */
  place(x: number, y: number, height: number): void {
    this.dom.style.transform = `translate(${x}px, ${y}px)`;
    this.dom.style.height = `${height}px`;
  }

  destroy(): void {
    this.view.destroy();
    this.dom.remove();
  }
}
