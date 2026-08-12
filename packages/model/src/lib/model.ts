import { Schema } from 'prosemirror-model';

/** Structural view of a DOM element — this package has no DOM lib, so parseDOM
 *  getAttrs callbacks narrow structurally (same idiom as the footnote mark). */
interface DomEl {
  getAttribute(name: string): string | null;
}

/** Complex attrs ride the DOM as data-* JSON so an internal copy/paste
 *  (ProseMirror's clipboard is a toDOM → parseDOM round-trip) keeps them.
 *  External HTML simply lacks the attribute → schema default. */
function dataJson(el: unknown, name: string): unknown {
  const raw = (el as DomEl).getAttribute(name);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

/** getAttrs for pasted <p>/<h1>–<h6>: heading level from the tag, alignment
 *  from inline style. Everything else keeps its schema default. */
function pastedParagraphAttrs(el: unknown, heading: number | null) {
  const style = (el as DomEl).getAttribute('style') ?? '';
  const m = /(?:^|;)\s*text-align\s*:\s*(center|right|justify)/i.exec(style);
  return {
    heading,
    align: m ? m[1].toLowerCase() : null,
    borders: dataJson(el, 'data-borders'),
    carry: dataJson(el, 'data-carry'),
  };
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
    crop: dataJson(el, 'data-crop'),
    outline: dataJson(el, 'data-outline'),
  };
}

/** Pasted <a href>: allow web/mail/anchor protocols only (blocks javascript:
 *  and friends). A rejected rule drops the mark but keeps the text. */
function pastedLinkAttrs(el: unknown) {
  const e = el as DomEl;
  const href = (e.getAttribute('href') ?? '').trim();
  return /^(https?:|mailto:|#)/i.test(href)
    ? { href, targetFrame: e.getAttribute('data-target-frame') }
    : (false as const);
}

/** getAttrs for <td>/<th>: spans from standard attributes, the rich cell
 *  attrs from their data-* JSON carriers, the fill from inline style. */
function pastedCellAttrs(el: unknown) {
  const e = el as DomEl;
  const span = (name: string) => {
    const n = parseInt(e.getAttribute(name) ?? '1', 10);
    return Number.isFinite(n) && n > 0 ? n : 1;
  };
  const style = e.getAttribute('style') ?? '';
  const bg = /(?:^|;)\s*background-color\s*:\s*([^;]+)/i.exec(style);
  return {
    colspan: span('colspan'),
    rowspan: span('rowspan'),
    colwidth: dataJson(el, 'data-colwidth'),
    background: bg ? bg[1].trim() : null,
    vAlign: e.getAttribute('data-valign'),
    borders: dataJson(el, 'data-borders'),
    diagonals: dataJson(el, 'data-diagonals'),
    padding: dataJson(el, 'data-padding'),
    carry: dataJson(el, 'data-carry'),
  };
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
        // PageConfig (contracts) — page size + margins in CSS px, from the
        // body w:sectPr. Edited via setDocAttribute (page-setup commands) so
        // orientation/paper-size changes undo cleanly. null → A4 default.
        page: { default: null },
        // CommentNode[] (contracts) — comment threads keyed to `comment` marks.
        // Edited via setDocAttribute so add/reply/resolve/delete undo cleanly.
        comments: { default: null },
        // Per-section header/footer story OVERRIDES, keyed by section index →
        // story ('headers'|'footers') → variant ('default'|'first'|'even') →
        // the full story doc as JSON. The first chrome EDIT the model supports
        // (the section marker's page-number toggle): an override replaces the
        // section's inherited story (Word's "unlink from previous"), rides the
        // doc so it undoes cleanly, and exports as a real header/footer part.
        sectionChromeOverrides: { default: null },
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
        // Named Word paragraph style with no outline level: 'Title' |
        // 'Subtitle', or null. Mutually exclusive with `heading` — the
        // setParagraphStyle command is the only writer and keeps the
        // invariant (styleId set ⇒ heading null).
        styleId: { default: null },
        // w:tabs — [{ pos, val: 'left'|'right'|'center'|'decimal', leader? }]
        // in px from the paragraph's content left edge, or null. Importer-set.
        tabs: { default: null },
        // w:spacing — { before?, after?, line?, lineRule? }, or null.
        spacing: { default: null },
        // w:bookmarkStart names anchored in this paragraph (["_Toc89595219"]),
        // or null. Link hrefs of the form "#name" resolve against these —
        // paragraph-level is the right altitude: Word's TOC bookmarks wrap a
        // heading's text, and jumping to the heading is what a reader wants.
        bookmarks: { default: null },
        // The generated field this paragraph belongs to ({ kind: 'toc',
        // instr } for a TOC entry), or null for ordinary content. Word paints
        // such content with field shading and regenerates it on update.
        field: { default: null },
        // w:pageBreakBefore — start this paragraph on a new page.
        pageBreakBefore: { default: false },
        // w:contextualSpacing — "ignore spacing above and below when using
        // identical styles". Present ⇒ the flag is on; the two booleans say
        // which side actually borders a paragraph of the SAME style, which is
        // resolved at import (only there is both sides of a boundary visible).
        contextualSpacing: { default: null },
        // w:keepNext — stay on the same page as the next block's first line.
        keepNext: { default: false },
        // w:keepLines — never split this paragraph across pages.
        keepLines: { default: false },
        // w:widowControl — Word's default is ON; false only when the document
        // explicitly disables widow/orphan control for this paragraph.
        widowControl: { default: true },
        // w:pBdr — { top?, bottom?, left?, right? } of BorderSide, or null.
        // Importer-set; painted as a box around the paragraph's lines.
        borders: { default: null },
        // w:shd at the paragraph layer, resolved to a fill "#RRGGBB" (or
        // null). Painted behind the paragraph's lines, in the same box the
        // borders use.
        shading: { default: null },
        // Carry-through fidelity (docx round-trip): OOXML paragraph
        // properties the model does NOT represent, preserved verbatim so a
        // customer's save never drops them. { pPr?: string, markRPr?: string }
        // — raw XML fragments (pPr extras / the paragraph-mark w:rPr), or
        // null. Importer-set; the exporter splices them back into w:pPr.
        carry: { default: null },
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
        const dom: Record<string, string> = style ? { style } : {};
        if (attrs.styleId) dom['data-style'] = attrs.styleId;
        if (node.attrs['borders'])
          dom['data-borders'] = JSON.stringify(node.attrs['borders']);
        if (node.attrs['shading'])
          dom['data-shading'] = String(node.attrs['shading']);
        // Round-trip fidelity survives in-app copy/paste (PM's clipboard is
        // DOM-serialized); external HTML paste simply has no such attribute.
        if (node.attrs['carry'])
          dom['data-carry'] = JSON.stringify(node.attrs['carry']);
        return [tag, dom, 0];
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
        // distB?, behind? } in px (behind = behindDoc: paint under the text),
        // or null for inline images. Importer-set.
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
        // a:srcRect — { l, t, r, b } ratios of the BITMAP, inward from each
        // edge, or null for the whole image. Negative values outset. The box
        // keeps its size; the selected region scales to fill it.
        crop: { default: null },
        // a:ln on pic:spPr — Word's picture border, as a BorderSide, or null.
        outline: { default: null },
      },
      parseDOM: [{ tag: 'img[src]', getAttrs: pastedImageAttrs }],
      toDOM(node) {
        const a = node.attrs;
        const attrs: Record<string, string> = {
          src: a['src'] as string,
          alt: a['alt'] as string,
        };
        if (a['width'] != null) attrs['width'] = String(a['width']);
        if (a['height'] != null) attrs['height'] = String(a['height']);
        if (a['crop']) attrs['data-crop'] = JSON.stringify(a['crop']);
        if (a['outline']) attrs['data-outline'] = JSON.stringify(a['outline']);
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
        // Carry-through fidelity: unmodelled w:tblPr children (tblStyle,
        // tblLayout, tblLook, tblInd, floating tblpPr, …) as one raw XML
        // string ({ tblPr: string }), or null. Importer-set; the exporter
        // splices it back so a save never drops them.
        carry: { default: null },
      },
      // Complex attrs round-trip as data-* JSON — ProseMirror's clipboard is
      // a toDOM → parseDOM pass, so without this an internal copy/paste
      // dropped borders / cell padding / alignment.
      parseDOM: [
        {
          tag: 'table',
          getAttrs: (el) => ({
            borders: dataJson(el, 'data-borders'),
            cellPadding: dataJson(el, 'data-cell-padding'),
            align: (el as DomEl).getAttribute('data-align'),
            carry: dataJson(el, 'data-carry'),
          }),
        },
      ],
      toDOM: (node) => {
        const a = node.attrs;
        const dom: Record<string, string> = {};
        if (a['borders']) dom['data-borders'] = JSON.stringify(a['borders']);
        if (a['cellPadding'])
          dom['data-cell-padding'] = JSON.stringify(a['cellPadding']);
        if (a['align']) dom['data-align'] = String(a['align']);
        if (a['carry']) dom['data-carry'] = JSON.stringify(a['carry']);
        return ['table', dom, ['tbody', 0]];
      },
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
        // Carry-through fidelity: unmodelled w:trPr children ({ trPr: string
        // } — gridBefore/wBefore, cnfStyle, …), or null. Importer-set.
        carry: { default: null },
      },
      parseDOM: [
        {
          tag: 'tr',
          getAttrs: (el) => ({
            header: (el as DomEl).getAttribute('data-header') === 'true',
            cantSplit: (el as DomEl).getAttribute('data-cant-split') === 'true',
            height: dataJson(el, 'data-height'),
            carry: dataJson(el, 'data-carry'),
          }),
        },
      ],
      toDOM: (node) => {
        const dom: Record<string, string> = {};
        if (node.attrs['header']) dom['data-header'] = 'true';
        if (node.attrs['cantSplit']) dom['data-cant-split'] = 'true';
        if (node.attrs['height'])
          dom['data-height'] = JSON.stringify(node.attrs['height']);
        if (node.attrs['carry'])
          dom['data-carry'] = JSON.stringify(node.attrs['carry']);
        return ['tr', dom, 0];
      },
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
        // w:tcBorders/w:tl2br + w:br2tl — rules across the cell's corners
        diagonals: { default: null },
        padding: { default: null }, // w:tcMar per-side margin override (px)
        // Carry-through fidelity: unmodelled w:tcPr children ({ tcPr: string
        // } — textDirection, noWrap, tcFitText, …), or null. Importer-set.
        carry: { default: null },
      },
      parseDOM: [
        { tag: 'td', getAttrs: pastedCellAttrs },
        { tag: 'th', getAttrs: pastedCellAttrs },
      ],
      toDOM(node) {
        const attrs: Record<string, string> = {};
        if (node.attrs['colspan'] !== 1)
          attrs['colspan'] = String(node.attrs['colspan']);
        if (node.attrs['rowspan'] !== 1)
          attrs['rowspan'] = String(node.attrs['rowspan']);
        if (node.attrs['background'])
          attrs['style'] = `background-color: ${node.attrs['background']}`;
        if (node.attrs['colwidth'])
          attrs['data-colwidth'] = JSON.stringify(node.attrs['colwidth']);
        if (node.attrs['vAlign'])
          attrs['data-valign'] = String(node.attrs['vAlign']);
        if (node.attrs['borders'])
          attrs['data-borders'] = JSON.stringify(node.attrs['borders']);
        if (node.attrs['diagonals'])
          attrs['data-diagonals'] = JSON.stringify(node.attrs['diagonals']);
        if (node.attrs['padding'])
          attrs['data-padding'] = JSON.stringify(node.attrs['padding']);
        if (node.attrs['carry'])
          attrs['data-carry'] = JSON.stringify(node.attrs['carry']);
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
    // w:dstrike — double strikethrough
    dstrike: {
      parseDOM: [
        {
          style: 'text-decoration-style=double',
          getAttrs: (value) => (value === 'double' ? {} : false),
        },
      ],
      toDOM: () => [
        'span',
        {
          style: 'text-decoration: line-through; text-decoration-style: double',
        },
        0,
      ],
    },
    // w:smallCaps
    smallCaps: {
      parseDOM: [
        {
          style: 'font-variant-caps',
          getAttrs: (value) => (value === 'small-caps' ? {} : false),
        },
      ],
      toDOM: () => ['span', { style: 'font-variant-caps: small-caps' }, 0],
    },

    // w:color — hex "#RRGGBB"
    textColor: {
      attrs: { color: {} },
      parseDOM: [
        { style: 'color', getAttrs: (value) => ({ color: value as string }) },
      ],
      toDOM(mark) {
        return [
          'span',
          { style: `color: ${mark.attrs['color'] as string}` },
          0,
        ];
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
        return [
          'span',
          { style: `font-size: ${mark.attrs['size'] as number}pt` },
          0,
        ];
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
    // w:spacing (rPr) — tracking in twips, positive = expanded. Absolute:
    // unlike the font size it does not shrink for superscript or small caps.
    letterSpacing: {
      attrs: { twips: {} },
      parseDOM: [
        {
          style: 'letter-spacing',
          getAttrs: (value) => {
            const m = /^(-?[\d.]+)pt$/.exec(String(value));
            return m ? { twips: Math.round(Number(m[1]) * 20) } : false;
          },
        },
      ],
      toDOM: (mark) => [
        'span',
        { style: `letter-spacing: ${(mark.attrs['twips'] as number) / 20}pt` },
        0,
      ],
    },
    // w:kern — the smallest font size (HALF-POINTS) that gets pair kerning.
    // A threshold, not a switch: whether it bites depends on the run's own
    // size, so the declared number is what the model keeps and the layout
    // compares. No DOM representation — CSS font-kerning cannot express a
    // size threshold, so an external paste simply has no opinion.
    kern: {
      attrs: { halfPoints: {} },
      parseDOM: [
        {
          tag: 'span[data-kern]',
          getAttrs: (el) => {
            const raw = (el as DomEl).getAttribute('data-kern');
            const n = Number(raw);
            return Number.isFinite(n) ? { halfPoints: n } : false;
          },
        },
      ],
      toDOM: (mark) => [
        'span',
        { 'data-kern': String(mark.attrs['halfPoints']) },
        0,
      ],
    },
    // w:w — horizontal glyph scale as a PERCENT (100 = normal). Squeezes the
    // glyphs and their advances; independent of letterSpacing, which rides on
    // top at its absolute value.
    charScale: {
      attrs: { percent: {} },
      parseDOM: [
        {
          style: 'transform',
          getAttrs: (value) => {
            const m = /^scaleX\(([\d.]+)\)$/.exec(String(value).trim());
            return m ? { percent: Math.round(Number(m[1]) * 100) } : false;
          },
        },
      ],
      toDOM: (mark) => [
        'span',
        {
          style:
            `display: inline-block; transform: scaleX(` +
            `${(mark.attrs['percent'] as number) / 100})`,
        },
        0,
      ],
    },
    // w:position — baseline shift in half-points, positive up. Unlike
    // super/subscript this does NOT resize the glyphs; Word treats the two
    // as independent and documents combine them.
    position: {
      attrs: { halfPoints: {} },
      parseDOM: [
        {
          style: 'vertical-align',
          getAttrs: (value) => {
            const m = /^(-?[\d.]+)pt$/.exec(String(value));
            return m ? { halfPoints: Number(m[1]) * 2 } : false;
          },
        },
      ],
      toDOM: (mark) => [
        'span',
        {
          style: `vertical-align: ${
            (mark.attrs['halfPoints'] as number) / 2
          }pt`,
        },
        0,
      ],
    },
    // w:highlight / w:shd w:fill — run background color ("#RRGGBB")
    highlight: {
      attrs: { color: {} },
      parseDOM: [
        {
          style: 'background-color',
          getAttrs: (value) => ({ color: value as string }),
        },
      ],
      toDOM(mark) {
        return [
          'span',
          { style: `background-color: ${mark.attrs['color'] as string}` },
          0,
        ];
      },
    },
    // w:rFonts — font family
    fontFamily: {
      attrs: { family: {} },
      parseDOM: [
        {
          style: 'font-family',
          getAttrs: (value) => ({ family: value as string }),
        },
      ],
      toDOM(mark) {
        return [
          'span',
          { style: `font-family: ${mark.attrs['family'] as string}` },
          0,
        ];
      },
    },
    // w:hyperlink — external URL or "#anchor"
    link: {
      // w:hyperlink @w:tgtFrame — the frame the link was meant to open in
      // ("_blank", "_top", or a frame name), from the HTML-frames era. There
      // is no frame here to obey it and nothing acts on it; it is modelled so
      // that saving the file gives the attribute back instead of dropping it.
      attrs: { href: {}, targetFrame: { default: null } },
      inclusive: false,
      parseDOM: [{ tag: 'a[href]', getAttrs: pastedLinkAttrs }],
      toDOM(mark) {
        return [
          'a',
          {
            href: mark.attrs['href'] as string,
            rel: 'noopener',
            target: '_blank',
            ...(mark.attrs['targetFrame']
              ? { 'data-target-frame': String(mark.attrs['targetFrame']) }
              : {}),
          },
          0,
        ];
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
            num:
              Number(
                (el as { getAttribute(n: string): string | null }).getAttribute(
                  'data-footnote',
                ),
              ) || 0,
          }),
        },
      ],
      toDOM(mark) {
        return [
          'sup',
          { 'data-footnote': String(mark.attrs['num'] as number) },
          0,
        ];
      },
    },
    // Carry-through fidelity (docx round-trip): run properties the model does
    // NOT represent (w:rtl, w:kern, w:szCs, …), preserved verbatim as one raw
    // XML fragment so saving a customer's file never drops them. Invisible —
    // no rendering; the docx exporter splices `xml` back into the run's rPr.
    // Typing inside/at the edge of a carried run extends the mark (inclusive
    // default), which is the faithful behavior for properties like w:rtl.
    carryRPr: {
      attrs: { xml: {} },
      toDOM(mark) {
        // Ride the clipboard (in-app copy/paste keeps fidelity); renders as
        // an unstyled span otherwise.
        return ['span', { 'data-carry-rpr': String(mark.attrs['xml']) }, 0];
      },
      parseDOM: [
        {
          tag: 'span[data-carry-rpr]',
          getAttrs: (el) => ({
            xml:
              (el as { getAttribute(n: string): string | null }).getAttribute(
                'data-carry-rpr',
              ) ?? '',
          }),
        },
      ],
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
    paragraph: {
      group: 'block',
      content: 'inline*',
      parseDOM: [{ tag: 'p' }],
      toDOM: () => ['p', 0],
    },
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
            const e = el as {
              getAttribute(n: string): string | null;
              textContent: string | null;
            };
            return {
              id: e.getAttribute('data-id'),
              label: (e.textContent ?? '').replace(/^@/, ''),
            };
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
  styleId?: 'Title' | 'Subtitle' | null;
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
    if (sp.line && sp.lineRule === 'auto')
      parts.push(`line-height: ${sp.line}`);
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
