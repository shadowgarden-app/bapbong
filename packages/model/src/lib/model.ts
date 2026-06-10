import { Schema } from 'prosemirror-model';

/**
 * bapbong's ProseMirror document schema.
 *
 * Deliberately minimal for M1 (DOCX import vertical slice): block paragraphs
 * of inline text with the four common character toggles. Lists, tables,
 * headings, images, etc. are added in later milestones — when they land they
 * extend THIS schema so the importer, layout engine, and (canvas) painter all
 * agree on one document model.
 */
export const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },

    paragraph: {
      group: 'block',
      content: 'inline*',
      // `list` is null for normal paragraphs, or ListInfo for list items.
      // DOCX stores lists flat (each w:p carries w:numPr), so we keep them
      // flat and let the marker be computed by the numbering counter.
      attrs: { list: { default: null } },
      parseDOM: [{ tag: 'p' }],
      toDOM: () => ['p', 0],
    },

    text: { group: 'inline' },

    // Inline image. `src` is typically a data URL (the importer inlines the
    // embedded media); width/height are CSS pixels, null if unspecified.
    image: {
      inline: true,
      group: 'inline',
      draggable: true,
      attrs: {
        src: {},
        alt: { default: '' },
        width: { default: null },
        height: { default: null },
      },
      toDOM(node) {
        const a = node.attrs;
        const attrs: Record<string, string> = { src: a['src'] as string, alt: a['alt'] as string };
        if (a['width'] != null) attrs['width'] = String(a['width']);
        if (a['height'] != null) attrs['height'] = String(a['height']);
        return ['img', attrs];
      },
    },

    // Tables: kept structural (table → row → cell → block+). Horizontal spans
    // map to colspan; vertical merges aren't collapsed yet (rowspan default 1).
    table: {
      group: 'block',
      content: 'table_row+',
      isolating: true,
      parseDOM: [{ tag: 'table' }],
      toDOM: () => ['table', ['tbody', 0]],
    },
    table_row: {
      content: 'table_cell+',
      parseDOM: [{ tag: 'tr' }],
      toDOM: () => ['tr', 0],
    },
    table_cell: {
      content: 'block+',
      isolating: true,
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        colwidth: { default: null }, // px widths of the spanned columns, or null
      },
      parseDOM: [{ tag: 'td' }, { tag: 'th' }],
      toDOM(node) {
        const attrs: Record<string, string> = {};
        if (node.attrs['colspan'] !== 1) attrs['colspan'] = String(node.attrs['colspan']);
        if (node.attrs['rowspan'] !== 1) attrs['rowspan'] = String(node.attrs['rowspan']);
        return ['td', attrs, 0];
      },
    },
  },

  marks: {
    // w:b
    strong: {
      parseDOM: [{ tag: 'strong' }, { tag: 'b' }],
      toDOM: () => ['strong', 0],
    },
    // w:i
    em: {
      parseDOM: [{ tag: 'em' }, { tag: 'i' }],
      toDOM: () => ['em', 0],
    },
    // w:u
    underline: {
      parseDOM: [{ tag: 'u' }],
      toDOM: () => ['u', 0],
    },
    // w:strike
    strike: {
      parseDOM: [{ tag: 's' }, { tag: 'strike' }],
      toDOM: () => ['s', 0],
    },

    // w:color — hex "#RRGGBB"
    textColor: {
      attrs: { color: {} },
      parseDOM: [{ style: 'color', getAttrs: (value) => ({ color: value as string }) }],
      toDOM(mark) {
        return ['span', { style: `color: ${mark.attrs['color'] as string}` }, 0];
      },
    },
    // w:sz — size in points
    fontSize: {
      attrs: { size: {} },
      parseDOM: [
        {
          style: 'font-size',
          getAttrs: (value) => {
            const pt = parseFloat(value as string);
            return Number.isNaN(pt) ? false : { size: pt };
          },
        },
      ],
      toDOM(mark) {
        return ['span', { style: `font-size: ${mark.attrs['size'] as number}pt` }, 0];
      },
    },
    // w:rFonts — font family
    fontFamily: {
      attrs: { family: {} },
      parseDOM: [{ style: 'font-family', getAttrs: (value) => ({ family: value as string }) }],
      toDOM(mark) {
        return ['span', { style: `font-family: ${mark.attrs['family'] as string}` }, 0];
      },
    },
    // w:hyperlink — external URL or "#anchor"
    link: {
      attrs: { href: {} },
      inclusive: false,
      toDOM(mark) {
        return ['a', { href: mark.attrs['href'] as string, rel: 'noopener', target: '_blank' }, 0];
      },
    },
  },
});

/** Concrete schema type, handy for typing Node/Mark across packages. */
export type BapbongSchema = typeof schema;

/** Value of a list paragraph's `list` attribute. `marker` is the resolved
 *  number/bullet string (e.g. "1.", "2.a", "•"); the numbering engine owns
 *  recomputation when the document is edited. */
export interface ListInfo {
  numId: string;
  level: number;
  marker: string;
}
