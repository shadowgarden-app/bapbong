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
      //
      // `align` mirrors w:jc ('left'|'center'|'right'|'justify'), null = default.
      // `indent` mirrors w:ind, all measured in CSS px (Indent | null):
      //   { left, right, firstLine, hanging }. firstLine and hanging are
      //   mutually exclusive in OOXML; hanging wins if both appear.
      attrs: {
        list: { default: null },
        align: { default: null },
        indent: { default: null },
      },
      // No getAttrs: nothing in the pipeline parses paragraphs from the DOM
      // yet (the importer builds nodes directly). align/indent still round-trip
      // out through toDOM. Revisit when HTML paste lands.
      parseDOM: [{ tag: 'p' }],
      toDOM(node) {
        const style = paragraphStyle(node.attrs as ParagraphAttrs);
        return ['p', style ? { style } : {}, 0];
      },
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
        // wp:anchor (floating image): { wrap: 'square'|'topAndBottom'|'none',
        // hAlign?, hOffset?, hRel?, vOffset?, vRel?, distL?, distR?, distT?,
        // distB? } in px, or null for inline images. Importer-set.
        float: { default: null },
      },
      toDOM(node) {
        const a = node.attrs;
        const attrs: Record<string, string> = { src: a['src'] as string, alt: a['alt'] as string };
        if (a['width'] != null) attrs['width'] = String(a['width']);
        if (a['height'] != null) attrs['height'] = String(a['height']);
        return ['img', attrs];
      },
    },

    // PAGE / NUMPAGES fields: atoms whose text is computed per page at paint
    // time. `kind` is 'page' (current page number) or 'pages' (total count).
    page_field: {
      inline: true,
      group: 'inline',
      atom: true,
      attrs: { kind: {} },
      parseDOM: [{ tag: 'span[data-page-field]' }],
      toDOM: (node) => [
        'span',
        { 'data-page-field': String(node.attrs['kind']) },
        node.attrs['kind'] === 'pages' ? '##' : '#',
      ],
    },

    // Tables: kept structural (table → row → cell → block+). Horizontal spans
    // map to colspan; vertical merges aren't collapsed yet (rowspan default 1).
    table: {
      group: 'block',
      content: 'table_row+',
      isolating: true,
      attrs: {
        // w:tblCellMar overrides (px: {left,right,top,bottom}), or null for
        // Word defaults. Importer-set (same rationale as paragraph attrs).
        cellPadding: { default: null },
      },
      parseDOM: [{ tag: 'table' }],
      toDOM: () => ['table', ['tbody', 0]],
    },
    table_row: {
      content: 'table_cell+',
      attrs: {
        header: { default: false }, // w:trPr/w:tblHeader — repeat on every page
      },
      // No getAttrs (same rationale as paragraph): the importer sets attrs
      // directly; revisit when HTML paste lands.
      parseDOM: [{ tag: 'tr' }],
      toDOM: (node) => (node.attrs['header'] ? ['tr', { 'data-header': 'true' }, 0] : ['tr', 0]),
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

/** Paragraph horizontal alignment (mirrors w:jc). */
export type Align = 'left' | 'center' | 'right' | 'justify';

/** Paragraph indentation in CSS px (mirrors w:ind). `firstLine` and `hanging`
 *  are mutually exclusive; if both are present, `hanging` takes precedence. */
export interface Indent {
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
}

/** Shape of the paragraph node's attrs (for typed toDOM/serialization). */
export interface ParagraphAttrs {
  list: ListInfo | null;
  align: Align | null;
  indent: Indent | null;
}

/** Build an inline CSS `style` string for a paragraph's align/indent, or ''
 *  when nothing applies. Used by toDOM (and the DOM preview in playground). */
function paragraphStyle(attrs: ParagraphAttrs): string {
  const parts: string[] = [];
  if (attrs.align) parts.push(`text-align: ${attrs.align}`);
  const ind = attrs.indent;
  if (ind) {
    if (ind.left) parts.push(`margin-left: ${ind.left}px`);
    if (ind.right) parts.push(`margin-right: ${ind.right}px`);
    // hanging wins over firstLine; negative text-indent renders the hang.
    if (ind.hanging) parts.push(`text-indent: ${-ind.hanging}px`);
    else if (ind.firstLine) parts.push(`text-indent: ${ind.firstLine}px`);
  }
  return parts.join('; ');
}

/** Value of a list paragraph's `list` attribute. `marker` is the resolved
 *  number/bullet string (e.g. "1.", "2.a", "•"); the numbering engine owns
 *  recomputation when the document is edited. */
export interface ListInfo {
  numId: string;
  level: number;
  marker: string;
}
