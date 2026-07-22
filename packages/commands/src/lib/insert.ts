import type { EditorState } from 'prosemirror-state';
import type { Command } from '@shadow-garden/bapbong-contracts';
import { isMarkActive } from './marks.js';

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

/** Insert a `rows`×`cols` table of empty cells, replacing the selection.
 *  New tables get Word's default look: a full 1px solid grid — OOXML tables
 *  are borderless unless declared, and an invisible fresh table reads as
 *  "nothing happened". */
export function insertTable(rows = 2, cols = 2): Command {
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
      const side = { width: 1, style: 'solid', color: '#000000' };
      const borders = {
        top: side,
        bottom: side,
        left: side,
        right: side,
        insideH: side,
        insideV: side,
      };
      const node = table.create(
        { borders },
        Array.from({ length: rows }, makeRow),
      );
      if (dispatch)
        dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
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
