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

/** Apply (or clear, with `null`) a hyperlink over the selected range. */
export function setLink(href: string | null): Command {
  return {
    name: 'link',
    run(state, dispatch) {
      const type = state.schema.marks['link'];
      if (!type || state.selection.empty) return false;
      if (dispatch) {
        const { from, to } = state.selection;
        const tr = state.tr.removeMark(from, to, type);
        if (href) tr.addMark(from, to, type.create({ href }));
        dispatch(tr);
      }
      return true;
    },
    isActive: (state) => isMarkActive(state, 'link'),
    isEnabled: (state) =>
      !!state.schema.marks['link'] && !state.selection.empty,
  };
}
