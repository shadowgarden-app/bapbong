import type {
  EditorPlugin,
  PluginContext,
  RangeDecoration,
} from '@shadow-garden/bapbong-contracts';
import {
  anchorName,
  fieldAt,
  findBookmark,
  type FieldInfo,
} from '@shadow-garden/bapbong-model';

/** Word's field shading grey, at the translucency the painter blends over
 *  page white without dulling the text. */
const FIELD_SHADE = 'rgba(0, 0, 0, 0.08)';

/** A field the caret currently sits in, with the paragraph range it covers. */
export interface ActiveField {
  field: FieldInfo;
  from: number;
  to: number;
}

export interface TocPlugin extends EditorPlugin {
  /** The generated field under `pos` (the caret by default), or null. Hosts
   *  use it to decide whether to offer field actions instead of ordinary
   *  text/link editing. */
  fieldAt(pos?: number): ActiveField | null;
  /** Recompute the page number of every entry in the TOC field covering
   *  `pos` (the caret by default) from the CURRENT layout — Word's "Update
   *  Field ▸ page numbers only". Returns how many entries changed. */
  updatePageNumbers(pos?: number): number;
}

declare module '@shadow-garden/bapbong-contracts' {
  interface EditorPluginHandles {
    toc: TocPlugin;
  }
}

/** The page number closing a TOC entry — the DIGITS only, positioned. Word's
 *  entries end "text → right tab (dot leader) → number", and the importer
 *  keeps the tab in the same text node ("\t12"), so the digit span is what
 *  may be replaced: rewriting the whole node would eat the tab and with it
 *  the leader dots. Null when the paragraph doesn't end in a number. */
interface NumberSpan {
  from: number;
  to: number;
  text: string;
  /** The original run's marks — the replacement keeps the entry's formatting
   *  (and its hyperlink) instead of dropping to bare text. */
  marks: readonly import('prosemirror-model').Mark[];
}

function trailingNumber(
  para: import('prosemirror-model').Node,
  paraPos: number,
): NumberSpan | null {
  let hit: NumberSpan | null = null;
  para.forEach((child, offset) => {
    if (!child.isText || !child.text) return;
    const text = child.text;
    const m = /^(\s*)(\d+)\s*$/.exec(text);
    if (m) {
      const start = paraPos + 1 + offset + m[1].length;
      hit = {
        from: start,
        to: start + m[2].length,
        text: m[2],
        marks: child.marks,
      };
    } else if (text.trim()) {
      hit = null; // real content after a number → that number wasn't the tail
    }
  });
  return hit;
}

/** The bookmark an entry paragraph links to (its first internal anchor). */
function entryAnchor(para: import('prosemirror-model').Node): string | null {
  let name: string | null = null;
  para.forEach((child) => {
    if (name || !child.isText) return;
    const link = child.marks.find((m) => m.type.name === 'link');
    if (link) name = anchorName(link.attrs['href'] as string);
  });
  return name;
}

/**
 * Internal plugin: **generated fields behave like Word's, not like ordinary
 * text**. A table of contents imported from a .docx is not a list of
 * hyperlinks a reader should edit link-by-link — it is field output that Word
 * shades grey when the caret enters it and regenerates on demand.
 *
 * This plugin supplies both halves: the shading (a decoration over the whole
 * field span, shown only while the caret is inside it) and the update command
 * that recounts each entry's page number from the real layout. The importer
 * models the span (paragraph `field` attr) and the anchors (`bookmarks`).
 */
export function tocPlugin(): TocPlugin {
  let ctx: PluginContext | null = null;

  const at = (pos?: number): ActiveField | null => {
    if (!ctx) return null;
    const p = pos ?? ctx.state.selection.from;
    return fieldAt(ctx.state.doc, p);
  };

  return {
    name: 'toc',
    setup(c) {
      ctx = c;
    },
    fieldAt: at,
    updatePageNumbers(pos) {
      if (!ctx) return 0;
      const active = at(pos);
      if (!active || active.field.kind !== 'toc') return 0;
      const { doc } = ctx.state;
      const edits: NumberSpan[] = [];
      doc.forEach((node, offset) => {
        if (offset < active.from || offset >= active.to) return;
        if (node.type.name !== 'paragraph') return;
        const name = entryAnchor(node);
        const tail = trailingNumber(node, offset);
        if (!name || !tail) return;
        const target = findBookmark(doc, name);
        if (target === null) return; // stale entry: leave its number alone
        const page = ctx?.caretRect(target)?.pageIndex;
        if (page == null) return;
        const text = String(page + 1);
        if (text !== tail.text) edits.push({ ...tail, text });
      });
      if (edits.length === 0) return 0;
      // Apply back-to-front so earlier edits don't shift later positions.
      // One pass: a changed number can itself repaginate the document, and
      // Word settles that the same way — by updating the field again.
      const tr = ctx.state.tr;
      for (const e of edits.reverse()) {
        tr.replaceWith(
          e.from,
          e.to,
          ctx.state.schema.text(e.text, [...e.marks]),
        );
      }
      ctx.dispatch(tr);
      return edits.length;
    },
    decorations(): RangeDecoration[] {
      const active = at();
      if (!active) return [];
      return [
        {
          from: active.from,
          to: active.to,
          kind: 'background',
          color: FIELD_SHADE,
        },
      ];
    },
  };
}
