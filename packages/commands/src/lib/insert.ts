import { NodeSelection, type EditorState } from 'prosemirror-state';
import type { Mark as Mark0, Node as PMNode0 } from 'prosemirror-model';
import type {
  Command,
  ResolvedTableStyle,
  TableLook,
  TableStyleSheet,
} from '@shadow-garden/bapbong-contracts';
import { MATH_ALPHABETS, mathLetters } from '@shadow-garden/bapbong-contracts';
import { isMarkActive } from './marks.js';
import { WORD_DEFAULT_TABLE_LOOK } from './table-style.js';

/** Toggle "start on a new page" (`w:pageBreakBefore`) on the paragraph(s) the
 *  selection touches; the first such paragraph decides the new on/off state. */
export function pageBreakCommand(): Command {
  const paragraphs = (state: EditorState) => {
    const { from, to } = state.selection;
    const out: Array<{ pos: number; on: boolean }> = [];
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name === 'paragraph')
        out.push({ pos, on: !!node.attrs['pageBreakBefore'] });
    });
    return out;
  };
  return {
    name: 'page-break',
    run(state, dispatch) {
      const ps = paragraphs(state);
      if (ps.length === 0) return false;
      if (dispatch) {
        const next = !ps[0].on;
        const tr = state.tr;
        for (const p of ps) tr.setNodeAttribute(p.pos, 'pageBreakBefore', next);
        dispatch(tr);
      }
      return true;
    },
    isActive: (state) => paragraphs(state).some((p) => p.on),
  };
}

/**
 * Insert `text` at the selection (replacing it), in the formatting of the
 * caret — the stored marks, else the marks of the position — exactly what a
 * typed character gets, and what Word's Insert › Symbol does with a
 * "(normal text)" character: it takes the surrounding font. Empty text is a
 * no-op that still reports handled (nothing to undo).
 */
export function insertText(text: string): Command {
  return {
    name: 'insert-text',
    run(state, dispatch) {
      if (!text) return true;
      if (dispatch) dispatch(state.tr.insertText(text).scrollIntoView());
      return true;
    },
  };
}

/** The table STYLE a new table is born with — Word's "Table Grid" — as the
 *  host resolves it (the docx package's catalog). */
export interface InsertTableStyle {
  styleId: string;
  /** Gates; absent = Word's default (0x04A0). */
  look?: TableLook;
  /** The resolved definition, injected into doc.attrs.tableStyles when the
   *  sheet lacks the id (the same move applyTableStyle makes). */
  style?: ResolvedTableStyle;
}

/**
 * Insert a `rows`×`cols` table of empty cells, replacing the selection.
 *
 * A new table must not be invisible — OOXML tables are borderless unless
 * declared — and Word gives it the "Table Grid" STYLE, not a grid of direct
 * borders. The difference shows the moment the style changes: a direct
 * border outranks any style, so a table born with a direct 1px grid kept
 * its black vertical lines under Medium Shading (whose look has none), and
 * its black frame under every style. With `style` the table carries the
 * style pair instead, and the gallery replaces it cleanly. Without one (or
 * on a schema without table styling), the direct grid remains the fallback.
 */
export function insertTable(
  rows = 2,
  cols = 2,
  style?: InsertTableStyle,
): Command {
  return {
    name: 'insert-table',
    run(state, dispatch) {
      const { table, table_row, table_cell, paragraph } = state.schema.nodes;
      if (!table || !table_row || !table_cell || !paragraph) return false;
      const makeRow = () =>
        table_row.create(
          null,
          Array.from({ length: cols }, () =>
            table_cell.create(null, paragraph.create()),
          ),
        );
      const styled = !!style && !!table.spec.attrs?.['styleId'];
      let attrs: Record<string, unknown>;
      if (styled) {
        attrs = {
          styleId: style.styleId,
          look: { ...(style.look ?? WORD_DEFAULT_TABLE_LOOK) },
        };
      } else {
        const side = { width: 1, style: 'solid', color: '#000000' };
        attrs = {
          borders: {
            top: side,
            bottom: side,
            left: side,
            right: side,
            insideH: side,
            insideV: side,
          },
        };
      }
      const node = table.create(attrs, Array.from({ length: rows }, makeRow));
      if (dispatch) {
        const tr = state.tr.replaceSelectionWith(node);
        const sheet = (state.doc.attrs['tableStyles'] ?? {}) as TableStyleSheet;
        if (
          styled &&
          style.style &&
          !sheet[style.styleId] &&
          state.schema.nodes['doc'].spec.attrs?.['tableStyles']
        )
          tr.setDocAttribute('tableStyles', {
            ...sheet,
            [style.styleId]: style.style,
          });
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
  };
}

/** Insert an inline image at the selection. */
export function insertImage(src: string, alt = ''): Command {
  return {
    name: 'insert-image',
    run(state, dispatch) {
      const image = state.schema.nodes['image'];
      if (!image || !src) return false;
      if (dispatch)
        dispatch(
          state.tr
            .replaceSelectionWith(image.create({ src, alt }), false)
            .scrollIntoView(),
        );
      return true;
    },
  };
}

/** Scheme-less input ("google.com") gets https:// so the link actually
 *  opens; in-document anchors ("#…") and real schemes pass through. */
function normalizeHref(href: string): string {
  const t = href.trim();
  if (t.startsWith('#') || /^[a-z][a-z0-9+.-]*:/i.test(t)) return t;
  return `https://${t}`;
}

/** The contiguous inline range around a collapsed caret whose children carry
 *  the link mark — what "remove link" unlinks without a selection. */
function linkRangeAt(state: EditorState): { from: number; to: number } | null {
  const type = state.schema.marks['link'];
  if (!type) return null;
  const $pos = state.selection.$from;
  const parent = $pos.parent;
  const linked = (i: number) =>
    i >= 0 && i < parent.childCount && !!type.isInSet(parent.child(i).marks);
  let start = $pos.index();
  // The child BEFORE the caret only counts when the caret sits exactly on the
  // boundary — a caret strictly inside an unlinked run next to a link is NOT
  // "in" that link.
  const atBoundary = $pos.textOffset === 0;
  if (!linked(start) && !(atBoundary && linked(start - 1))) return null;
  if (!linked(start)) start--;
  while (linked(start - 1)) start--;
  let end = start + 1;
  while (linked(end)) end++;
  let from = $pos.start();
  let to = from;
  for (let i = 0; i < end; i++) {
    const size = parent.child(i).nodeSize;
    if (i < start) from += size;
    to += size;
  }
  return { from, to };
}

/** The link under a collapsed caret: href + range + its visible text — what
 *  a link panel shows in view mode and prefills for editing. */
export function linkInfoAt(
  state: EditorState,
): { href: string; from: number; to: number; text: string } | null {
  const type = state.schema.marks['link'];
  if (!type) return null;
  const range = linkRangeAt(state);
  if (!range) return null;
  let href = '';
  state.doc.nodesBetween(range.from, range.to, (n) => {
    if (href) return;
    const mark = type.isInSet(n.marks);
    if (mark) href = String(mark.attrs['href']);
  });
  return { href, ...range, text: state.doc.textBetween(range.from, range.to) };
}

/** Apply (or clear, with `null`) a hyperlink. Over a selection the mark wraps
 *  the range; at a collapsed caret, inserting types `label` (or the address
 *  itself, Word-style) as linked text, a caret INSIDE a link edits that run
 *  in place, and clearing unlinks it — no silent no-ops. */
export function setLink(href: string | null, label?: string): Command {
  return {
    name: 'link',
    run(state, dispatch) {
      const type = state.schema.marks['link'];
      if (!type) return false;
      if (!state.selection.empty) {
        if (dispatch) {
          const { from, to } = state.selection;
          const tr = state.tr.removeMark(from, to, type);
          if (href)
            tr.addMark(from, to, type.create({ href: normalizeHref(href) }));
          dispatch(tr);
        }
        return true;
      }
      const range = linkRangeAt(state);
      if (href) {
        const mark = type.create({ href: normalizeHref(href) });
        const text = (label ?? '').trim();
        if (range) {
          // Edit the linked run under the caret: new href, and new text when
          // the label changed. Non-link marks on the run (bold, size…) carry
          // over to replacement text.
          if (dispatch) {
            const current = state.doc.textBetween(range.from, range.to);
            if (text && text !== current) {
              const $from = state.doc.resolve(range.from);
              const keep = ($from.nodeAfter?.marks ?? []).filter(
                (m) => m.type !== type,
              );
              dispatch(
                state.tr.replaceWith(
                  range.from,
                  range.to,
                  state.schema.text(text, [...keep, mark]),
                ),
              );
            } else {
              dispatch(
                state.tr
                  .removeMark(range.from, range.to, type)
                  .addMark(range.from, range.to, mark),
              );
            }
          }
          return true;
        }
        const shown = text || href.trim();
        if (!shown) return false;
        if (dispatch) {
          dispatch(
            state.tr
              .replaceSelectionWith(state.schema.text(shown, [mark]), false)
              .scrollIntoView(),
          );
        }
        return true;
      }
      if (!range) return false;
      if (dispatch) dispatch(state.tr.removeMark(range.from, range.to, type));
      return true;
    },
    isActive: (state) => isMarkActive(state, 'link'),
    isEnabled: (state) => !!state.schema.marks['link'],
  };
}

/** The empty-slot glyph a fresh equation starts with — inserted SELECTED, so
 *  the first keystroke replaces it (Word's dotted placeholder box). */
export const EQUATION_PLACEHOLDER = '□';

/**
 * Insert ▸ Equation (Word's Alt+=), linear v1.
 *
 * With a selection: the range becomes an equation — letters restyle through
 * the math-italic alphabet (a → 𝑎, the OMML default letterform) and the
 * whole range takes the math mark, exactly what the importer produces for an
 * OMML equation, so the exporter turns it back into m:oMath.
 *
 * With a caret: a placeholder slot is inserted math-marked and selected;
 * typing replaces it and — the mark being inclusive — keeps extending the
 * equation until the caret leaves it.
 */
export function insertEquation(): Command {
  return {
    name: 'insert-equation',
    run(state, dispatch) {
      const math = state.schema.marks['math'];
      if (!math) return false;
      const { from, to, empty } = state.selection;
      if (!state.selection.$from.parent.isTextblock) return false;
      if (!dispatch) return true;
      const tr = state.tr;
      if (empty) {
        tr.insertText(EQUATION_PLACEHOLDER, from);
        tr.addMark(from, from + EQUATION_PLACEHOLDER.length, math.create());
        const Sel = state.selection.constructor as unknown as {
          create(
            doc: PMNode0,
            from: number,
            to: number,
          ): typeof state.selection;
        };
        tr.setSelection(
          Sel.create(tr.doc, from, from + EQUATION_PLACEHOLDER.length),
        );
        dispatch(tr.scrollIntoView());
        return true;
      }
      // Restyle end → start so earlier positions stay valid; letter → math
      // italic changes UTF-16 lengths (𝑎 is a surrogate pair).
      const italic = MATH_ALPHABETS['italic'];
      const pieces: {
        a: number;
        b: number;
        text: string;
        marks: readonly Mark0[];
      }[] = [];
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText) return;
        const a = Math.max(from, pos);
        const b = Math.min(to, pos + node.nodeSize);
        const slice = (node.text ?? '').slice(a - pos, b - pos);
        pieces.push({
          a,
          b,
          text: mathLetters(slice, italic),
          marks: node.marks,
        });
      });
      let end = to;
      for (const p of [...pieces].reverse()) {
        if (
          p.text.length !== p.b - p.a ||
          p.text !== state.doc.textBetween(p.a, p.b)
        ) {
          tr.replaceWith(
            p.a,
            p.b,
            state.schema.text(p.text, p.marks as Mark0[]),
          );
          end += p.text.length - (p.b - p.a);
        }
      }
      tr.addMark(from, end, math.create());
      dispatch(tr.scrollIntoView());
      return true;
    },
    isActive: (state) => isMarkActive(state, 'math'),
  };
}

/**
 * Insert a 2D equation carrying `ast` at the selection — what the built-in
 * gallery does. The node is selected afterwards so the equation plugin can
 * open its slot editor on it (Word drops the caret into a new equation too).
 */
export function insertEquationNode(ast: unknown): Command {
  return {
    name: 'insert-equation-node',
    run(state, dispatch) {
      const type = state.schema.nodes['equation'];
      if (!type) return false;
      if (!state.selection.$from.parent.isTextblock) return false;
      if (!dispatch) return true;
      const from = state.selection.from;
      const node = type.create({ ast, sizePt: 12 });
      const tr = state.tr.replaceSelectionWith(node, false);
      // A NodeSelection over the fresh equation: the plugin picks it up and
      // opens its slot editor, so typing continues inside the equation.
      tr.setSelection(NodeSelection.create(tr.doc, from));
      dispatch(tr.scrollIntoView());
      return true;
    },
  };
}
