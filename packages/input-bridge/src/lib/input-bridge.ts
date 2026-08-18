import { baseKeymap, chainCommands } from 'prosemirror-commands';
import { history, redo, undo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import {
  DOMSerializer,
  type Node as PMNode,
  type Schema,
} from 'prosemirror-model';
import {
  EditorState,
  Plugin,
  PluginKey,
  TextSelection,
  type Command,
  type Transaction,
} from 'prosemirror-state';
import { canSplit } from 'prosemirror-transform';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';

// Re-exported so hosts type against ONE prosemirror-state identity (mixing
// module resolutions across packages makes TS treat duplicates as unrelated).
export type { Command, EditorState, Transaction } from 'prosemirror-state';
// Same rationale for the view type: `InputBridge.view` crosses the package
// boundary into the editor, so consumers must import EditorView from HERE (not
// straight from prosemirror-view) to share one identity — otherwise a stale
// `tsc --build` composite state can treat the same file as two unrelated types.
export type { EditorView } from 'prosemirror-view';

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
      state.tr.setNodeMarkup($from.before(), null, {
        ...parent.attrs,
        list: null,
      }),
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

/** Enter at the END of a paragraph: the new paragraph keeps the paragraph
 *  formatting, the way Word's does — alignment, indents, tabs, spacing,
 *  pagination keeps, borders/shading and the paragraph MARK's own font
 *  (`markFont`, what sizes an empty line and what typed text picks up). Not
 *  carried: heading/styleId (Word moves on to the style's "next style",
 *  Normal), bookmarks and field (anchors and generated content stay with the
 *  paragraph they were on), list (see splitListItem, which runs first) and
 *  carry (unmodelled XML). Enter MID-paragraph is left to the base keymap,
 *  which already splits in place with the attrs intact; only the at-end
 *  split fell back to a bare default paragraph. */
export const splitParagraphKeepFormat: Command = (state, dispatch) => {
  const { $from, $to } = state.selection;
  const parent = $from.parent;
  if (!(state.selection instanceof TextSelection)) return false;
  if (!parent.isTextblock || $from.parent !== $to.parent) return false;
  if ($to.parentOffset !== parent.content.size) return false; // not at end
  const attrs = {
    ...parent.attrs,
    heading: null,
    styleId: null,
    bookmarks: null,
    field: null,
    list: null,
    carry: null,
  };
  const tr = state.tr.deleteSelection();
  const pos = tr.selection.from;
  const types = [{ type: parent.type, attrs }];
  if (!canSplit(tr.doc, pos, 1, types)) return false;
  tr.split(pos, 1, types);
  dispatch?.(tr.scrollIntoView());
  return true;
};

/** The Enter key: continue a list, else keep the paragraph formatting on an
 *  at-end split; anything else falls through to the base keymap. */
export const paragraphEnter: Command = chainCommands(
  splitListItem,
  splitParagraphKeepFormat,
);

/** Word: text typed into an EMPTY paragraph takes the paragraph mark's font
 *  — the ¶ is the only thing on that line, and it is what holds the
 *  formatting there. ProseMirror has no such glyph: an empty textblock offers
 *  no marks to inherit, so typing would fall back to the document default.
 *  This plugin seeds the STORED marks from `markFont` whenever a caret comes
 *  to rest in an empty paragraph without any (a stored-mark set left by a
 *  font command — bold toggled off, say — is respected: `[]` is not null).
 *  Only the four properties the mark models are seeded (family, size, bold,
 *  italic), as resolved marks — the same shape imported runs carry. */
export const markFontSeed = new Plugin({
  appendTransaction(_trs, _old, state) {
    if (state.storedMarks) return null;
    const { empty, $from } = state.selection;
    const parent = $from.parent;
    if (!empty || !parent.isTextblock || parent.content.size !== 0) return null;
    const mf = parent.attrs['markFont'] as {
      family?: string;
      sizePt?: number;
      bold?: boolean;
      italic?: boolean;
    } | null;
    if (!mf) return null;
    const { marks } = state.schema;
    const seeded = [];
    if (mf.family && marks['fontFamily'])
      seeded.push(marks['fontFamily'].create({ family: mf.family }));
    if (mf.sizePt != null && marks['fontSize'])
      seeded.push(marks['fontSize'].create({ size: mf.sizePt }));
    if (mf.bold && marks['strong']) seeded.push(marks['strong'].create());
    if (mf.italic && marks['em']) seeded.push(marks['em'].create());
    return seeded.length ? state.tr.setStoredMarks(seeded) : null;
  },
});

/** Backspace at the very start of a paragraph outdents in Word-like steps
 *  instead of immediately joining backward:
 *    1. a list item drops its marker (the `list` attr clears, indent kept);
 *    2. an indented paragraph clears its indent (caret returns to the margin);
 *    3. a plain, unindented paragraph defers to the base keymap (joins backward).
 *  Only fires on a collapsed caret at the block start; returns false elsewhere
 *  so ranged deletes and mid-line Backspace run normally. */
export const backspaceOutdent: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  if (!empty) return false; // ranged Backspace deletes the selection
  const parent = $from.parent;
  if (!parent.isTextblock || $from.parentOffset !== 0) return false;
  const attrs = parent.attrs;
  const outdent =
    attrs['list'] != null
      ? { list: null } // step 1: drop the marker, keep the indent
      : attrs['indent'] != null
        ? { indent: null } // step 2: clear the indent, caret to the margin
        : null; // step 3: nothing left to outdent
  if (!outdent) return false;
  dispatch?.(
    state.tr.setNodeMarkup($from.before(), null, { ...attrs, ...outdent }),
  );
  return true;
};

/** Indent step per nesting level, matching the hanging indent Word gives
 *  editor-authored lists (0.25" = 24px). */
const LEVEL_INDENT = 24;
/** Word's numbering depth limit (levels 0–8) when the definition has no
 *  explicit level table to cap against. */
const MAX_LEVEL = 8;

/** Tab / Shift-Tab in a list: change the nesting level of every list item the
 *  selection touches (Word demote/promote). The level is capped by the levels
 *  the numbering definition actually defines (fallback 0–8), and the left
 *  indent shifts 24px per level so nesting stays visible even for defs without
 *  per-level indents. Returns false when the selection touches no list item,
 *  so Tab keeps its default behavior outside lists. */
export function shiftListLevel(delta: 1 | -1): Command {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    const targets: { pos: number; node: PMNode }[] = [];
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isTextblock) return true; // descend into tables etc.
      if (node.attrs['list']) targets.push({ pos, node });
      return false;
    });
    if (targets.length === 0) return false;

    const defs = state.doc.attrs['numbering'] as {
      [numId: string]: { levels?: Record<string, unknown> };
    } | null;
    let tr = state.tr;
    let changed = false;
    for (const { pos, node } of targets) {
      const list = node.attrs['list'] as { numId: string; level?: number };
      const levels = defs?.[list.numId]?.levels;
      const defined = levels ? Object.keys(levels).map(Number) : [];
      const max = defined.length > 0 ? Math.max(...defined) : MAX_LEVEL;
      const level = Math.max(0, Math.min(max, (list.level ?? 0) + delta));
      if (level === (list.level ?? 0)) continue;
      const indent = (node.attrs['indent'] as { left?: number } | null) ?? null;
      const left = Math.max(0, (indent?.left ?? 0) + delta * LEVEL_INDENT);
      const nextIndent = { ...indent, left };
      if (left === 0) delete (nextIndent as { left?: number }).left;
      tr = tr.setNodeMarkup(pos, null, {
        ...node.attrs,
        list: { ...list, level },
        indent: Object.keys(nextIndent).length > 0 ? nextIndent : null,
      });
      changed = true;
    }
    if (!changed) return false;
    dispatch?.(tr.scrollIntoView());
    return true;
  };
}

const WORD_CHAR = /[\p{L}\p{N}_]/u;

/** The word range around `pos` (for double-click selection), or null when the
 *  position isn't inside a textblock or doesn't touch a word character.
 *  Unicode-aware: Vietnamese diacritics count as word characters. */
export function wordRangeAt(
  doc: PMNode,
  pos: number,
): { from: number; to: number } | null {
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
  /** A pre-built editor state to adopt verbatim instead of creating a fresh
   *  one from `doc` — carries an in-progress editing session (undo/redo history,
   *  selection) across a rebind. When given, its `doc` must already be the one
   *  passed as `doc` (same schema). `keys` still drives the keymaps it was built
   *  with, so the state must have been produced by `createEditingState`. */
  state?: EditorState;
  /** Extra bindings, checked before the base keymap — e.g. ArrowUp/ArrowDown
   *  wired to layout-aware caret motion from bapbong-selection. */
  keys?: Record<string, Command>;
  /** Called after every dispatched transaction (typing, IME composition
   *  steps, undo, selection changes). Re-layout + repaint here. */
  onUpdate: (state: EditorState, tr: Transaction) => void;
  /** editorProps.handlePaste — return true to claim the paste (e.g. image
   *  blobs); false lets ProseMirror's default clipboard parsing run. */
  handlePaste?: (view: EditorView, event: ClipboardEvent) => boolean;
}

/** Editing state with history + base keymap; exported for headless tests. */
export function createEditingState(
  doc: PMNode,
  keys: Record<string, Command> = {},
): EditorState {
  return EditorState.create({
    doc,
    plugins: [
      history(),
      keymap({ 'Mod-z': undo, 'Shift-Mod-z': redo, 'Mod-y': redo }),
      keymap(keys),
      keymap(baseKeymap),
      markFontSeed,
    ],
  });
}

// ── DOM windowing (large-document WebKit fix) ─────────────────────────
//
// The bridge holds the WHOLE document in its ProseMirror state, but it must
// not hold the whole document as DOM: on WebKit, the mere presence of a
// several-thousand-block hidden tree stalls the rendering pipeline for
// seconds after every scroll (frames stop while the main thread stays idle)
// — ablation-verified: with the tree out of the document the stalls vanish,
// and no amount of `contain` / `content-visibility` / off-layer positioning
// helps while it is present.
//
// So only a WINDOW of blocks around the selection is rendered for real; every
// other top-level block renders as an empty stub element. The model, history,
// clipboard and IME are untouched (they live on the full PM state — positions
// never remapped); the DOM selection always sits inside the window, and the
// window follows the selection: a `decorations` prop marks in-window blocks,
// and the block node views recreate themselves whenever their membership
// changes (update() → false). Stub blocks are opaque leaves (no contentDOM),
// so ProseMirror never renders their children at all.

/** Blocks kept fully rendered before/after the selection ends. Large enough
 *  that Enter/Backspace joins and multi-line IME always touch real DOM. */
const WINDOW_BUFFER = 4;

const IN_WINDOW = 'bapbongInWindow';

/** Node decorations marking the top-level blocks that must render for real:
 *  a ±buffer around the selection's endpoints (two segments when a long
 *  selection spans more — the blocks between stay stubs; DOM selection only
 *  needs real anchor/head nodes, and copy serializes from the model). */
function windowDecorations(state: EditorState): DecorationSet {
  const doc = state.doc;
  if (doc.childCount === 0) return DecorationSet.empty;
  const last = doc.childCount - 1;
  const clamp = (i: number) => Math.max(0, Math.min(last, i));
  const fromIdx = doc.resolve(state.selection.from).index(0);
  const toIdx = doc.resolve(state.selection.to).index(0);
  const K = WINDOW_BUFFER;
  const segs: [number, number][] =
    toIdx - fromIdx <= 2 * K
      ? [[clamp(fromIdx - K), clamp(toIdx + K)]]
      : [
          [clamp(fromIdx - K), clamp(fromIdx + K)],
          [clamp(toIdx - K), clamp(toIdx + K)],
        ];
  const decos: Decoration[] = [];
  let idx = 0;
  doc.forEach((node, offset) => {
    if (segs.some(([a, b]) => idx >= a && idx <= b)) {
      decos.push(
        Decoration.node(
          offset,
          offset + node.nodeSize,
          {},
          { [IN_WINDOW]: true },
        ),
      );
    }
    idx++;
  });
  return DecorationSet.create(doc, decos);
}

/** Whether the outer decorations mark this block as inside the window. */
function isInWindow(decorations: readonly Decoration[]): boolean {
  return decorations.some(
    (d) => (d.spec as Record<string, unknown>)[IN_WINDOW],
  );
}

/** Node view for top-level blocks: real (default-equivalent, with contentDOM)
 *  inside the window, an empty stub outside it. Nested occurrences (e.g.
 *  paragraphs inside table cells) always render for real — their table is
 *  already window-gated. Returning false from update() recreates the view when
 *  window membership flips. */
function windowedBlockView(
  node: PMNode,
  view: EditorView,
  getPos: () => number | undefined,
  decorations: readonly Decoration[],
): {
  dom: HTMLElement;
  contentDOM?: HTMLElement;
  update?: (n: PMNode, decos: readonly Decoration[]) => boolean;
  ignoreMutation?: () => boolean;
} {
  const pos = getPos();
  const topLevel =
    typeof pos === 'number' && view.state.doc.resolve(pos).depth === 0;
  if (!topLevel || isInWindow(decorations)) {
    // Default-equivalent rendering (schema toDOM) so PM manages the children.
    let dom: HTMLElement;
    let contentDOM: HTMLElement | undefined;
    const spec = node.type.spec.toDOM?.(node);
    if (spec != null) {
      const rendered = DOMSerializer.renderSpec(document, spec);
      dom = rendered.dom as HTMLElement;
      contentDOM =
        (rendered.contentDOM as HTMLElement | undefined) ?? undefined;
    } else {
      dom = document.createElement('div');
      contentDOM = dom;
    }
    return {
      dom,
      contentDOM,
      update: (n, decos) => {
        if (n.type !== node.type) return false;
        // Fell out of the window → recreate as a stub.
        if (topLevel && !isInWindow(decos)) return false;
        return true; // PM syncs children into contentDOM
      },
    };
  }
  // Stub: an opaque, empty placeholder — its children are never rendered and
  // model changes inside it need no DOM work (the model is the source of
  // truth; the canvas paints from the layout engine, not from this DOM).
  const dom = document.createElement(node.isTextblock ? 'p' : 'div');
  return {
    dom,
    update: (n, decos) => {
      if (n.type !== node.type) return false;
      if (isInWindow(decos)) return false; // entered the window → materialize
      return true;
    },
    ignoreMutation: () => true,
  };
}

/**
 * Hidden ProseMirror editor acting as the canvas's input sink.
 *
 * The browser routes keyboard and IME composition into a real (but invisible)
 * contenteditable; ProseMirror keeps the model, history and clipboard. The
 * host positions `dom` at the canvas caret via `place()` so IME candidate
 * popups appear next to the visible (painted) caret. Only the blocks around
 * the selection exist as real DOM — see the DOM-windowing section above.
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
      // In the page-canvas stack (absolute), so the host — and the IME popup
      // anchored to it — scrolls along with the painted caret. Safe because the
      // windowed DOM above keeps this subtree small even for huge documents.
      position: 'absolute',
      inset: '0',
      overflow: 'hidden',
      pointerEvents: 'none',
      zIndex: '-1',
      // Full containment: external layout changes never reach this subtree.
      contain: 'strict',
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
      // Suppress the browser's NATIVE text caret. The visible caret is painted
      // on the canvas; the hidden editor only sinks input. Left visible, the
      // native caret blinks in this whole-document contenteditable, and on
      // WebKit each blink re-lays-out the (huge) hidden DOM — a ~900ms
      // main-thread stall every ~500ms that froze large docs.
      caretColor: 'transparent',
      // Fixed-size containment root: the host is explicitly 800px × 1em with the
      // document clipped inside, so its layout must not participate in (or be
      // invalidated by) the surrounding canvas stack's.
      contain: 'strict',
    } satisfies Partial<CSSStyleDeclaration>);
    this.dom.appendChild(this.host);

    // The keymap only sees Enter as a clean keydown. When an IME composition
    // is active (Vietnamese Telex composes on EVERY letter), the committing
    // Enter arrives with isComposing=true, the keymap skips it, and the
    // browser applies its NATIVE insertParagraph to the hidden contenteditable
    // — WebKit's DOM edit re-parses into a paragraph OUTSIDE isolating nodes
    // (caret jumps below the table it was typing in). Claim the beforeinput
    // and run the same Enter chain the keymap would have.
    const enterChain = chainCommands(
      options.keys?.['Enter'] ?? (() => false),
      baseKeymap['Enter'],
    );
    this.view = new EditorView(this.host, {
      state: options.state ?? createEditingState(options.doc, options.keys),
      handlePaste: options.handlePaste,
      // DOM windowing: only blocks near the selection render for real.
      decorations: windowDecorations,
      nodeViews: {
        // Top-level blocks window-gate themselves (stub away from the caret).
        paragraph: windowedBlockView,
        table: windowedBlockView,
        // The hidden editor is invisible and only sinks input; it never needs
        // the actual bitmaps (the canvas paints from the model). Render images
        // as empty, correctly-sized placeholders so the in-window DOM does not
        // hold — and WebKit does not decode — the doc's media.
        image(node) {
          const dom = document.createElement('img');
          const w = Number(node.attrs['width']) || 0;
          const h = Number(node.attrs['height']) || 0;
          if (w) dom.style.width = `${w}px`;
          if (h) dom.style.height = `${h}px`;
          return { dom };
        },
      },
      handleDOMEvents: {
        beforeinput: (view, event) => {
          if (event.inputType !== 'insertParagraph') return false;
          event.preventDefault();
          enterChain(view.state, view.dispatch);
          return true;
        },
      },
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
    const sel = TextSelection.between(
      doc.resolve(clamp(anchor)),
      doc.resolve(clamp(head)),
    );
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
        if (
          doc.childCount === 1 &&
          first?.isTextblock &&
          first.content.size === 0
        ) {
          return DecorationSet.create(doc, [
            Decoration.node(0, first.nodeSize, {
              class: 'is-empty',
              'data-placeholder': text,
            }),
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
        const textBefore = $from.parent.textBetween(
          0,
          $from.parentOffset,
          '\n',
          '￼',
        );
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
        handlers.query({
          query: st.query,
          coords: { left: c.left, top: c.top, bottom: c.bottom },
        });
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
    opts?: {
      onEnter?: () => void;
      onEscape?: () => void;
      placeholder?: string;
    },
  ) {
    const doc = initialDoc ? schema.nodeFromJSON(initialDoc) : undefined;
    const plugins = [history()];
    if (opts?.placeholder) plugins.push(placeholderPlugin(opts.placeholder));
    if (mention && schema.nodes['mention'])
      plugins.push(mentionPlugin(mention));
    const keys: Record<string, Command> = {};
    if (opts?.onEnter) {
      keys['Enter'] = () => (opts.onEnter?.(), true);
      keys['Shift-Enter'] = baseKeymap['Enter']; // newline within the reply
    }
    if (opts?.onEscape) keys['Escape'] = () => (opts.onEscape?.(), true);
    if (Object.keys(keys).length) plugins.push(keymap(keys));
    plugins.push(
      keymap({ 'Mod-z': undo, 'Shift-Mod-z': redo }),
      keymap(baseKeymap),
    );
    this.view = new EditorView(mount, {
      state: EditorState.create({ ...(doc ? { doc } : { schema }), plugins }),
    });
  }

  /** Replace the active "@query" with a mention node + trailing space. */
  applyMention(user: MentionUser): void {
    const st = mentionKey.getState(this.view.state);
    if (!st?.active) return;
    const to = this.view.state.selection.from;
    const node = this.schema.nodes['mention'].create({
      id: user.id,
      label: user.label,
    });
    const tr = this.view.state.tr.replaceWith(st.from, to, [
      node,
      this.schema.text(' '),
    ]);
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
    const empty = EditorState.create({
      schema: this.schema,
      plugins: this.view.state.plugins,
    });
    this.view.updateState(empty);
  }

  focus(): void {
    this.view.focus();
  }

  destroy(): void {
    this.view.destroy();
  }
}
