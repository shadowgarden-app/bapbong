import { Schema } from 'prosemirror-model';

/** Structural view of a DOM element — this package has no DOM lib, so parseDOM
 *  getAttrs callbacks narrow structurally (same idiom as the footnote mark). */
interface DomEl {
  getAttribute(name: string): string | null;
}

/** getAttrs for pasted <p>/<h1>–<h6>: heading level from the tag, alignment
 *  from inline style. Everything else keeps its schema default. */
function pastedParagraphAttrs(el: unknown, heading: number | null) {
  const style = (el as DomEl).getAttribute('style') ?? '';
  const m = /(?:^|;)\s*text-align\s*:\s*(center|right|justify)/i.exec(style);
  return { heading, align: m ? m[1].toLowerCase() : null };
}

/** Pasted <img>: only embedded bitmaps survive. Remote URLs are rejected —
 *  the canvas painter can't reliably fetch them (CORS) and DOCX export embeds
 *  media bytes; the paste layer converts blobs to data URLs before insert. */
function pastedImageAttrs(el: unknown) {
  const e = el as DomEl;
  const src = e.getAttribute('src') ?? '';
  if (!/^data:image\//i.test(src)) return false as const;
  const dim = (v: string | null) => {
    const n = parseFloat(v ?? '');
    return Number.isFinite(n) && n > 0 ? Math.round(n) : null;
  };
  return {
    src,
    alt: e.getAttribute('alt') ?? '',
    width: dim(e.getAttribute('width')),
    height: dim(e.getAttribute('height')),
  };
}

/** Pasted <a href>: allow web/mail/anchor protocols only (blocks javascript:
 *  and friends). A rejected rule drops the mark but keeps the text. */
function pastedLinkAttrs(el: unknown) {
  const href = ((el as DomEl).getAttribute('href') ?? '').trim();
  return /^(https?:|mailto:|#)/i.test(href) ? { href } : (false as const);
}

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
    doc: {
      content: 'block+',
      attrs: {
        // NumberingDefs (numbering.ts) — list markers are recomputed from
        // these at layout time, so edits renumber live. Importer-set.
        numbering: { default: null },
        // SectionConfig[] (contracts) — per-section column flow, delimited by
        // w:sectPr breaks. null/absent → one implicit single-column section.
        sections: { default: null },
        // CommentNode[] (contracts) — comment threads keyed to `comment` marks.
        // Edited via setDocAttribute so add/reply/resolve/delete undo cleanly.
        comments: { default: null },
      },
    },

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
        // Heading level 1–6 (maps to a Word "Heading N" style on export, and to
        // an <h1>–<h6> in toDOM so the a11y mirror is semantic), or null for a
        // body paragraph.
        heading: { default: null },
        // w:tabs — [{ pos, val: 'left'|'right'|'center'|'decimal', leader? }]
        // in px from the paragraph's content left edge, or null. Importer-set.
        tabs: { default: null },
        // w:spacing — { before?, after?, line?, lineRule? }, or null.
        spacing: { default: null },
        // w:pageBreakBefore — start this paragraph on a new page.
        pageBreakBefore: { default: false },
      },
      // HTML paste path: recover heading level from h1–h6 and alignment from
      // inline style. Other attrs (list/indent/tabs/spacing) stay importer-only
      // — pasted HTML rarely carries them faithfully.
      parseDOM: [
        { tag: 'p', getAttrs: (el) => pastedParagraphAttrs(el, null) },
        ...[1, 2, 3, 4, 5, 6].map((level) => ({
          tag: `h${level}`,
          getAttrs: (el: unknown) => pastedParagraphAttrs(el, level),
        })),
      ],
      toDOM(node) {
        const attrs = node.attrs as ParagraphAttrs;
        const style = paragraphStyle(attrs);
        const tag = attrs.heading ? `h${attrs.heading}` : 'p';
        return [tag, style ? { style } : {}, 0];
      },
    },

    text: { group: 'inline' },

    // A soft line break inside a paragraph (w:br) — forces a new line without
    // ending the paragraph. Occupies one PM position, like an image atom.
    hard_break: {
      inline: true,
      group: 'inline',
      selectable: false,
      parseDOM: [{ tag: 'br' }],
      toDOM: () => ['br'],
    },

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
        // Drawn vector shape (wps rect / straight connector) riding this box:
        // { kind: 'rect'|'line', stroke?, strokeWidth?, fill?, flipV? } — src
        // is then '' and the box paints as vector. Null for real bitmaps.
        shape: { default: null },
        // Textbox (wps:txbx) content riding a shape: { paragraphs: <paragraph
        // node JSON>[], inset?: {l,t,r,b} px }. The layout engine flows the
        // paragraphs inside the shape's box (paint-only, not editable v1).
        textbox: { default: null },
        // Clockwise rotation in degrees around the box center (a:xfrm@rot).
        // Paint-only: the layout box stays axis-aligned.
        rotation: { default: 0 },
      },
      parseDOM: [{ tag: 'img[src]', getAttrs: pastedImageAttrs }],
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
        // w:tblBorders visibility { top, bottom, left, right, insideH,
        // insideV }, or null — OOXML tables are borderless unless declared.
        borders: { default: null },
        // w:tblPr/w:jc — 'center' | 'right' table alignment, or null (left).
        align: { default: null },
      },
      parseDOM: [{ tag: 'table' }],
      toDOM: (node) => ['table', node.attrs['borders'] ? { 'data-borders': '1' } : {}, ['tbody', 0]],
    },
    table_row: {
      content: 'table_cell+',
      attrs: {
        header: { default: false }, // w:trPr/w:tblHeader — repeat on every page
        // w:trPr/w:cantSplit — the row must not break across pages. Absent
        // (Word's default) means the paginator may split the row mid-content.
        cantSplit: { default: false },
        // w:trHeight — { value: px, exact: boolean } or null (auto).
        height: { default: null },
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
        background: { default: null }, // w:shd w:fill — cell fill "#RRGGBB"
        vAlign: { default: null }, // w:vAlign — 'center' | 'bottom' (top default)
        borders: { default: null }, // w:tcBorders per-side visibility override
      },
      parseDOM: [{ tag: 'td' }, { tag: 'th' }],
      toDOM(node) {
        const attrs: Record<string, string> = {};
        if (node.attrs['colspan'] !== 1) attrs['colspan'] = String(node.attrs['colspan']);
        if (node.attrs['rowspan'] !== 1) attrs['rowspan'] = String(node.attrs['rowspan']);
        if (node.attrs['background']) attrs['style'] = `background-color: ${node.attrs['background']}`;
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
    // w:vertAlign — superscript / subscript
    vertAlign: {
      attrs: { value: {} }, // 'super' | 'sub'
      parseDOM: [
        { tag: 'sup', getAttrs: () => ({ value: 'super' }) },
        { tag: 'sub', getAttrs: () => ({ value: 'sub' }) },
      ],
      toDOM: (mark) => [mark.attrs['value'] === 'sub' ? 'sub' : 'sup', 0],
    },
    // w:highlight / w:shd w:fill — run background color ("#RRGGBB")
    highlight: {
      attrs: { color: {} },
      parseDOM: [
        { style: 'background-color', getAttrs: (value) => ({ color: value as string }) },
      ],
      toDOM(mark) {
        return ['span', { style: `background-color: ${mark.attrs['color'] as string}` }, 0];
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
      parseDOM: [{ tag: 'a[href]', getAttrs: pastedLinkAttrs }],
      toDOM(mark) {
        return ['a', { href: mark.attrs['href'] as string, rel: 'noopener', target: '_blank' }, 0];
      },
    },
    // w:footnoteReference — the carrier text is the superscript number; `num`
    // lets the layout engine match the reference to its page-bottom body.
    footnote: {
      attrs: { num: {} },
      inclusive: false,
      parseDOM: [
        {
          tag: 'sup[data-footnote]',
          // `el` is an HTMLElement at runtime; this package has no DOM lib, so
          // narrow structurally rather than naming the type.
          getAttrs: (el) => ({
            num: Number((el as { getAttribute(n: string): string | null }).getAttribute('data-footnote')) || 0,
          }),
        },
      ],
      toDOM(mark) {
        return ['sup', { 'data-footnote': String(mark.attrs['num'] as number) }, 0];
      },
    },
    // The `comment` mark (w:commentRangeStart/End) is contributed by the comment
    // plugin (@shadow-garden/bapbong-comments) via the editor's schema
    // composition — it is NOT part of the base document schema, so a build
    // without comments stays free of comment marks. The `comments` doc attr
    // above is inert storage the plugin populates when present.
  },
});

/** Concrete schema type, handy for typing Node/Mark across packages. */
export type BapbongSchema = typeof schema;

/**
 * Minimal schema for composing a comment body (the comment sidebar's little
 * editors). Rich enough for paragraphs of text plus an inline `mention` atom
 * (@user). Comment bodies are stored as this schema's JSON on the comment
 * thread, kept separate from the document schema.
 */
export const commentSchema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*', parseDOM: [{ tag: 'p' }], toDOM: () => ['p', 0] },
    text: { group: 'inline' },
    // @mention: an inline atom carrying the mentioned user's id + display name.
    // `leafText` lets textContent include "@Name" (search / plain-text preview).
    mention: {
      group: 'inline',
      inline: true,
      atom: true,
      selectable: false,
      attrs: { id: {}, label: {} },
      leafText: (node) => `@${node.attrs['label']}`,
      toDOM: (node) => [
        'span',
        { class: 'mention', 'data-id': String(node.attrs['id']) },
        `@${node.attrs['label']}`,
      ],
      parseDOM: [
        {
          tag: 'span.mention',
          getAttrs: (el: unknown) => {
            const e = el as { getAttribute(n: string): string | null; textContent: string | null };
            return { id: e.getAttribute('data-id'), label: (e.textContent ?? '').replace(/^@/, '') };
          },
        },
      ],
    },
  },
  marks: {},
});

export type CommentSchema = typeof commentSchema;

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
export interface Spacing {
  before?: number;
  after?: number;
  line?: number;
  lineRule?: 'auto' | 'exact' | 'atLeast';
}

export interface ParagraphAttrs {
  list: ListInfo | null;
  align: Align | null;
  indent: Indent | null;
  spacing?: Spacing | null;
  heading?: number | null;
}

/** Build an inline CSS `style` string for a paragraph's align/indent/spacing,
 *  or '' when nothing applies. Used by toDOM (and the DOM preview). */
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
  const sp = attrs.spacing;
  if (sp) {
    if (sp.before) parts.push(`margin-top: ${sp.before}px`);
    if (sp.after) parts.push(`margin-bottom: ${sp.after}px`);
    if (sp.line && sp.lineRule === 'auto') parts.push(`line-height: ${sp.line}`);
    else if (sp.line) parts.push(`line-height: ${sp.line}px`);
  }
  return parts.join('; ');
}

/** Value of a list paragraph's `list` attribute. The marker string ("1.",
 *  "2.a", "•") is NOT stored — it's recomputed at layout time from the doc's
 *  numbering defs, so edits renumber live. `marker` remains only for legacy
 *  callers that pass pre-resolved markers straight to the layout engine. */
export interface ListInfo {
  numId: string;
  level: number;
  marker?: string;
}
