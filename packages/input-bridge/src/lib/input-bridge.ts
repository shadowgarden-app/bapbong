import { baseKeymap } from 'prosemirror-commands';
import { history, redo, undo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import type { Node as PMNode, Schema } from 'prosemirror-model';
import { EditorState, Plugin, PluginKey, TextSelection, type Command, type Transaction } from 'prosemirror-state';
import { canSplit } from 'prosemirror-transform';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';

// Re-exported so hosts type against ONE prosemirror-state identity (mixing
// module resolutions across packages makes TS treat duplicates as unrelated).
export type { Command, EditorState, Transaction } from 'prosemirror-state';

/** A command that moves the caret to the position computed by `compute`
 *  (e.g. layout-aware ArrowUp/ArrowDown from bapbong-selection). With `extend`,
 *  the anchor stays put and only the head moves (Shift+arrow selection).
 *  Returns false — leaving the key to the next handler — when `compute`
 *  yields null. */
export function moveCaretCommand(
  compute: (state: EditorState) => number | null,
  extend = false,
): Command {
  return (state, dispatch) => {
    const pos = compute(state);
    if (pos == null) return false;
    const sel = extend
      ? TextSelection.create(state.doc, state.selection.anchor, pos)
      : TextSelection.create(state.doc, pos);
    dispatch?.(state.tr.setSelection(sel));
    return true;
  };
}

/** Enter inside a list item: split the paragraph KEEPING its list attrs (the
 *  layout engine recounts markers, so the new item numbers itself and
 *  everything below renumbers). Enter on an EMPTY item exits the list, like
 *  Word. Returns false outside lists so the base keymap takes over. */
export const splitListItem: Command = (state, dispatch) => {
  const { $from, $to } = state.selection;
  const parent = $from.parent;
  if (!parent.isTextblock || !parent.attrs['list']) return false;
  if ($from.parent !== $to.parent) return false; // cross-block selection → default handling

  if (parent.content.size === 0) {
    // Empty item: leave the list instead of adding another empty marker.
    dispatch?.(
      state.tr.setNodeMarkup($from.before(), null, { ...parent.attrs, list: null }),
    );
    return true;
  }

  const tr = state.tr.deleteSelection();
  const pos = tr.selection.from;
  if (!canSplit(tr.doc, pos)) return false;
  tr.split(pos, 1, [{ type: parent.type, attrs: parent.attrs }]);
  dispatch?.(tr);
  return true;
};

const WORD_CHAR = /[\p{L}\p{N}_]/u;

/** The word range around `pos` (for double-click selection), or null when the
 *  position isn't inside a textblock or doesn't touch a word character.
 *  Unicode-aware: Vietnamese diacritics count as word characters. */
export function wordRangeAt(doc: PMNode, pos: number): { from: number; to: number } | null {
  if (pos < 0 || pos > doc.content.size) return null;
  const $pos = doc.resolve(pos);
  const parent = $pos.parent;
  if (!parent.isTextblock) return null;
  // Leaf nodes (images) become U+FFFC, which is not a word character.
  const text = parent.textBetween(0, parent.content.size, undefined, '￼');
  const offset = pos - $pos.start();
  let from = offset;
  let to = offset;
  while (from > 0 && WORD_CHAR.test(text[from - 1])) from--;
  while (to < text.length && WORD_CHAR.test(text[to])) to++;
  if (from === to) return null;
  return { from: $pos.start() + from, to: $pos.start() + to };
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
  /** Clip layer appended near the canvas: it fills the positioned ancestor
   *  and `overflow: hidden`s the 800px-wide host riding the caret inside —
   *  otherwise that tail pokes past the right page edge and gifts the
   *  scroll viewport a horizontal scrollbar. */
  readonly dom: HTMLElement;
  /** The hidden editor's host — the element `place()` actually moves. */
  private readonly host: HTMLElement;
  readonly view: EditorView;

  constructor(options: InputBridgeOptions) {
    this.dom = document.createElement('div');
    this.dom.className = 'bapbong-input-bridge';
    Object.assign(this.dom.style, {
      position: 'absolute',
      inset: '0',
      overflow: 'hidden',
      pointerEvents: 'none',
      zIndex: '-1',
    } satisfies Partial<CSSStyleDeclaration>);

    this.host = document.createElement('div');
    Object.assign(this.host.style, {
      position: 'absolute',
      top: '0',
      left: '0',
      // MUST be wide: a 1px host makes the hidden document wrap one character
      // per line (tens of thousands of line boxes), and every DOM-selection
      // placement then re-runs that layout — ~500ms PER CLICK on multi-page
      // documents. At a page-ish width the same layout is ~10ms and stays
      // warm. The host is invisible either way (opacity 0, clipped height).
      width: '800px',
      height: '1em',
      overflow: 'hidden',
      opacity: '0',
      // Keep it focusable but out of the way of canvas pointer events.
      pointerEvents: 'none',
    } satisfies Partial<CSSStyleDeclaration>);
    this.dom.appendChild(this.host);

    this.view = new EditorView(this.host, {
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

  /** Dispatch a transaction (e.g. a comment authoring command) through the
   *  editor, so it routes through onUpdate + history like any edit. */
  dispatch(tr: Transaction): void {
    this.view.dispatch(tr);
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

  /** Select the word around `pos` (double-click); falls back to a collapsed
   *  caret when there is no word at that position. */
  selectWordAt(pos: number): void {
    const range = wordRangeAt(this.view.state.doc, pos);
    if (range) this.setSelection(range.from, range.to);
    else this.setSelection(pos);
  }

  /** Move the hidden editor to the painted caret (CSS px, relative to the
   *  positioned ancestor) so IME popups anchor in the right place. */
  place(x: number, y: number, height: number): void {
    this.host.style.transform = `translate(${x}px, ${y}px)`;
    this.host.style.height = `${height}px`;
  }

  destroy(): void {
    this.view.destroy();
    this.dom.remove();
  }
}

/** A user the composer can @mention. */
export interface MentionUser {
  id: string;
  label: string;
}

/** Where the @query popup should sit (client coords of the caret). */
export interface MentionCoords {
  left: number;
  top: number;
  bottom: number;
}

/** Host hooks for the @mention popup. The composer owns detection + insertion;
 *  the host owns rendering the popup and choosing the user. */
export interface MentionHandlers {
  /** Popup should show/move (state) or hide (null). */
  query(state: { query: string; coords: MentionCoords } | null): void;
  /** A nav key fired while the popup is open; return true if the host consumed
   *  it (then the composer swallows it instead of editing text). */
  key(key: 'up' | 'down' | 'enter' | 'esc'): boolean;
}

interface MentionState {
  active: boolean;
  from: number; // doc position of the triggering '@'
  query: string;
}

const mentionKey = new PluginKey<MentionState>('mention');
const INACTIVE: MentionState = { active: false, from: 0, query: '' };
// '@' at a word boundary, then Unicode letters/digits/_ (so Vietnamese works).
const MENTION_RE = /(?:^|\s)@([\p{L}\p{N}_]*)$/u;

/** Shows `text` as placeholder while the editor holds a single empty block. */
function placeholderPlugin(text: string): Plugin {
  return new Plugin({
    props: {
      decorations(state) {
        const doc = state.doc;
        const first = doc.firstChild;
        if (doc.childCount === 1 && first?.isTextblock && first.content.size === 0) {
          return DecorationSet.create(doc, [
            Decoration.node(0, first.nodeSize, { class: 'is-empty', 'data-placeholder': text }),
          ]);
        }
        return null;
      },
    },
  });
}

/** Detects a trailing "@query" before the caret and drives the host popup. */
function mentionPlugin(handlers: MentionHandlers): Plugin<MentionState> {
  return new Plugin<MentionState>({
    key: mentionKey,
    state: {
      init: () => INACTIVE,
      apply(tr) {
        const sel = tr.selection;
        if (!sel.empty) return INACTIVE;
        const $from = sel.$from;
        const textBefore = $from.parent.textBetween(0, $from.parentOffset, '\n', '￼');
        const m = MENTION_RE.exec(textBefore);
        if (!m) return INACTIVE;
        const query = m[1];
        return { active: true, from: $from.pos - query.length - 1, query };
      },
    },
    view: (editorView) => {
      const emit = (view: EditorView) => {
        const st = mentionKey.getState(view.state);
        if (!st?.active) return handlers.query(null);
        const c = view.coordsAtPos(view.state.selection.from);
        handlers.query({ query: st.query, coords: { left: c.left, top: c.top, bottom: c.bottom } });
      };
      emit(editorView);
      return { update: emit, destroy: () => handlers.query(null) };
    },
    props: {
      handleKeyDown(view, event) {
        if (!mentionKey.getState(view.state)?.active) return false;
        const map: Record<string, 'up' | 'down' | 'enter' | 'esc'> = {
          ArrowUp: 'up',
          ArrowDown: 'down',
          Enter: 'enter',
          Tab: 'enter',
          Escape: 'esc',
        };
        const k = map[event.key];
        if (k && handlers.key(k)) {
          event.preventDefault();
          return true;
        }
        return false;
      },
    },
  });
}

/**
 * A small standalone ProseMirror editor for composing a comment body. The host
 * passes the body schema (bapbong-model's commentSchema) so input-bridge stays
 * schema-agnostic; `getJSON()` returns the composed doc to store on the thread.
 * Pass `mention` handlers to enable @-mentions (the host renders the popup).
 * `opts.onEnter` makes Enter submit (Shift-Enter newlines) — used by the inline
 * reply box; `opts.onEscape` cancels. The mention plugin is ordered first so it
 * gets Enter/Esc while its popup is open, before these keymaps.
 */
export class CommentComposer {
  readonly view: EditorView;

  constructor(
    private readonly schema: Schema,
    mount: HTMLElement,
    initialDoc?: unknown,
    mention?: MentionHandlers,
    opts?: { onEnter?: () => void; onEscape?: () => void; placeholder?: string },
  ) {
    const doc = initialDoc ? schema.nodeFromJSON(initialDoc) : undefined;
    const plugins = [history()];
    if (opts?.placeholder) plugins.push(placeholderPlugin(opts.placeholder));
    if (mention && schema.nodes['mention']) plugins.push(mentionPlugin(mention));
    const keys: Record<string, Command> = {};
    if (opts?.onEnter) {
      keys['Enter'] = () => (opts.onEnter?.(), true);
      keys['Shift-Enter'] = baseKeymap['Enter']; // newline within the reply
    }
    if (opts?.onEscape) keys['Escape'] = () => (opts.onEscape?.(), true);
    if (Object.keys(keys).length) plugins.push(keymap(keys));
    plugins.push(keymap({ 'Mod-z': undo, 'Shift-Mod-z': redo }), keymap(baseKeymap));
    this.view = new EditorView(mount, {
      state: EditorState.create({ ...(doc ? { doc } : { schema }), plugins }),
    });
  }

  /** Replace the active "@query" with a mention node + trailing space. */
  applyMention(user: MentionUser): void {
    const st = mentionKey.getState(this.view.state);
    if (!st?.active) return;
    const to = this.view.state.selection.from;
    const node = this.schema.nodes['mention'].create({ id: user.id, label: user.label });
    const tr = this.view.state.tr.replaceWith(st.from, to, [node, this.schema.text(' ')]);
    tr.setSelection(TextSelection.create(tr.doc, st.from + 2)); // after mention + space
    this.view.dispatch(tr);
    this.view.focus();
  }

  /** Whether the composer has no visible text. */
  isEmpty(): boolean {
    return this.view.state.doc.textContent.trim().length === 0;
  }

  /** The composed body as commentSchema doc JSON. */
  getJSON(): unknown {
    return this.view.state.doc.toJSON();
  }

  /** Reset to an empty document. */
  clear(): void {
    const empty = EditorState.create({ schema: this.schema, plugins: this.view.state.plugins });
    this.view.updateState(empty);
  }

  focus(): void {
    this.view.focus();
  }

  destroy(): void {
    this.view.destroy();
  }
}
