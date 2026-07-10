import JSZip from 'jszip';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { commentSchema } from '@shadow-garden/bapbong-model';
import type { BorderStyle, CommentNode, SectionConfig, ShapeSpec, TableBorders } from '@shadow-garden/bapbong-contracts';

/**
 * DOCX export (round-trip). Phases:
 *  E1 ✅ paragraphs + runs + common character marks + pPr + hard breaks.
 *  E2 ✅ lists (numPr), tables, inline images (+ media + rels), hyperlinks (+ rels).
 *  E3 ⬜ comments (ranges + comments.xml + commentsExtended.xml).
 *  E4 ⬜ carry the original imported parts (styles/numbering/headers/media).
 */

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** Deterministic 8-hex w14:paraId for a comment id (links replies in w15). */
const paraId = (id: number) => (0x10000000 + id).toString(16).toUpperCase();

/** px → twips (1px @96dpi = 15 twips); inverse of the importer's twipsToPx. */
const pxToTwips = (px: number) => Math.round(px * 15);
/** px → EMU (9525 EMU/px); inverse of emuToPx. */
const pxToEmu = (px: number) => Math.round(px * 9525);

/** Accumulates document.xml.rels + media parts emitted during serialization. */
interface ExportCtx {
  rels: string[];
  media: { path: string; base64: string }[];
  exts: Set<string>; // image extensions → content-type defaults
  nextId: number; // shared rId / media / drawing counter
  // Comment ranges (E3): the doc's known comment ids, the last inline-node index
  // each id covers (so we can close the range), the currently-open ids, and a
  // running inline-node index aligned with the precompute.
  knownComments: Set<number>;
  lastRun: Map<number, number>;
  openComments: Set<number>;
  runIdx: number;
}

const commentIdsOf = (node: PMNode): number[] =>
  (node.marks.find((m) => m.type.name === 'comment')?.attrs['ids'] as number[] | undefined) ?? [];

/** Is this an inline leaf the comment-range index counts (text / image / break)? */
const isInlineLeaf = (node: PMNode): boolean => node.isText || node.type.name === 'image' || node.type.name === 'hard_break';

function esc(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string);
}

const MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/jpg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
};

// ── runs ────────────────────────────────────────────────────────────

/** A run's w:rPr from its marks (character marks only; link is a wrapper). */
function runProps(marks: readonly Mark[]): string {
  const byName = new Map(marks.map((m) => [m.type.name, m]));
  const out: string[] = [];
  const fam = byName.get('fontFamily')?.attrs['family'] as string | undefined;
  if (fam) out.push(`<w:rFonts w:ascii="${esc(fam)}" w:hAnsi="${esc(fam)}"/>`);
  if (byName.has('strong')) out.push('<w:b/>');
  if (byName.has('em')) out.push('<w:i/>');
  if (byName.has('underline')) out.push('<w:u w:val="single"/>');
  if (byName.has('strike')) out.push('<w:strike/>');
  const color = byName.get('textColor')?.attrs['color'] as string | undefined;
  if (color) out.push(`<w:color w:val="${color.replace(/^#/, '')}"/>`);
  const size = byName.get('fontSize')?.attrs['size'] as number | undefined;
  if (size != null) out.push(`<w:sz w:val="${Math.round(size * 2)}"/>`);
  const hl = byName.get('highlight')?.attrs['color'] as string | undefined;
  if (hl) out.push(`<w:shd w:val="clear" w:color="auto" w:fill="${hl.replace(/^#/, '')}"/>`);
  const va = byName.get('vertAlign')?.attrs['value'] as string | undefined;
  if (va) out.push(`<w:vertAlign w:val="${va === 'sub' ? 'subscript' : 'superscript'}"/>`);
  return out.length ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

/** Float attrs → the wp:anchor wrapper (position + wrap) around a graphic. */
function anchorXml(float: Record<string, unknown>, cx: number, cy: number, n: number, graphic: string): string {
  const dist = (k: string) => pxToEmu((float[k] as number) ?? 0);
  const hRel = float['hRel'] === 'page' ? 'page' : 'column';
  const vRel = float['vRel'] === 'page' ? 'page' : float['vRel'] === 'margin' ? 'margin' : 'paragraph';
  const posH = float['hAlign']
    ? `<wp:align>${float['hAlign']}</wp:align>`
    : `<wp:posOffset>${pxToEmu((float['hOffset'] as number) ?? 0)}</wp:posOffset>`;
  const posV = `<wp:posOffset>${pxToEmu((float['vOffset'] as number) ?? 0)}</wp:posOffset>`;
  const wrap =
    float['wrap'] === 'topAndBottom'
      ? '<wp:wrapTopAndBottom/>'
      : float['wrap'] === 'none'
        ? '<wp:wrapNone/>'
        : '<wp:wrapSquare wrapText="bothSides"/>';
  return (
    `<wp:anchor distT="${dist('distT')}" distB="${dist('distB')}" distL="${dist('distL')}" distR="${dist('distR')}" ` +
    `simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${hRel}">${posH}</wp:positionH>` +
    `<wp:positionV relativeFrom="${vRel}">${posV}</wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    wrap +
    `<wp:docPr id="${n}" name="Shape ${n}"/><wp:cNvGraphicFramePr/>` +
    graphic +
    `</wp:anchor>`
  );
}

/** ` rot="..."` for a:xfrm from the node's rotation attr (1/60000 deg), or ''. */
function rotAttr(node: PMNode): string {
  const deg = Number(node.attrs['rotation']) || 0;
  return deg ? ` rot="${Math.round(deg * 60000)}"` : '';
}

/** A drawn shape (rect/line) → wps drawing; inline or anchored per `float`. */
function shapeXml(node: PMNode, ctx: ExportCtx): string {
  const s = node.attrs['shape'] as ShapeSpec;
  const rot = rotAttr(node);
  const n = ctx.nextId++;
  const cx = pxToEmu((node.attrs['width'] as number) ?? 0);
  const cy = pxToEmu((node.attrs['height'] as number) ?? 0);
  const fill = s.fill ? `<a:solidFill><a:srgbClr val="${s.fill.replace(/^#/, '')}"/></a:solidFill>` : '<a:noFill/>';
  const ln = s.stroke
    ? `<a:ln w="${pxToEmu(s.strokeWidth ?? 1)}"><a:solidFill><a:srgbClr val="${s.stroke.replace(/^#/, '')}"/></a:solidFill></a:ln>`
    : '<a:ln><a:noFill/></a:ln>';
  // Textbox paragraphs ride the node as PM JSON; re-emit them as txbxContent
  // so the text survives the round-trip.
  const tb = node.attrs['textbox'] as
    | { paragraphs: unknown[]; inset?: { l: number; t: number; r: number; b: number } }
    | null;
  const txbx = tb
    ? `<wps:txbx><w:txbxContent>${tb.paragraphs
        .map((json) => paragraphXml(node.type.schema.nodeFromJSON(json), ctx))
        .join('')}</w:txbxContent></wps:txbx>`
    : '';
  const bodyPr = tb?.inset
    ? `<wps:bodyPr lIns="${pxToEmu(tb.inset.l)}" tIns="${pxToEmu(tb.inset.t)}" rIns="${pxToEmu(tb.inset.r)}" bIns="${pxToEmu(tb.inset.b)}"/>`
    : '<wps:bodyPr/>';
  const graphic =
    `<a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp><wps:cNvSpPr${tb ? ' txBox="1"' : ''}/>` +
    `<wps:spPr><a:xfrm${rot}${s.flipV ? ' flipV="1"' : ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${s.kind}"><a:avLst/></a:prstGeom>${fill}${ln}</wps:spPr>` +
    `${txbx}${bodyPr}</wps:wsp></a:graphicData></a:graphic>`;
  const float = node.attrs['float'] as Record<string, unknown> | null;
  const body = float
    ? anchorXml(float, cx, cy, n, graphic)
    : `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:docPr id="${n}" name="Shape ${n}"/>${graphic}</wp:inline>`;
  return `<w:r><w:drawing>${body}</w:drawing></w:r>`;
}

function imageXml(node: PMNode, ctx: ExportCtx): string {
  if (node.attrs['shape']) return shapeXml(node, ctx);
  const src = String(node.attrs['src'] ?? '');
  const m = /^data:([^;]+);base64,(.+)$/.exec(src);
  if (!m) return '';
  const ext = MIME_EXT[m[1].toLowerCase()] ?? 'png';
  ctx.exts.add(ext);
  const n = ctx.nextId++;
  const rid = `rId${n}`;
  ctx.media.push({ path: `word/media/image${n}.${ext}`, base64: m[2] });
  ctx.rels.push(`<Relationship Id="${rid}" Type="${R_NS}/image" Target="media/image${n}.${ext}"/>`);
  const cx = pxToEmu((node.attrs['width'] as number) ?? 96);
  const cy = pxToEmu((node.attrs['height'] as number) ?? 96);
  return (
    `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${n}" name="Picture ${n}"/>` +
    `<a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${n}" name="image${n}.${ext}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm${rotAttr(node)}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  );
}

/** One inline node → its run XML (excluding the link wrapper). */
function inlineXml(node: PMNode, ctx: ExportCtx): string {
  if (node.type.name === 'hard_break') return '<w:r><w:br/></w:r>';
  if (node.type.name === 'image') return imageXml(node, ctx);
  if (node.isText) {
    const fn = node.marks.find((m) => m.type.name === 'footnote');
    if (fn) return `<w:r>${runProps(node.marks)}<w:footnoteReference w:id="${fn.attrs['num']}"/></w:r>`;
    return `<w:r>${runProps(node.marks)}<w:t xml:space="preserve">${esc(node.text ?? '')}</w:t></w:r>`;
  }
  return '';
}

const linkHref = (node: PMNode): string | null =>
  (node.marks.find((m) => m.type.name === 'link')?.attrs['href'] as string | undefined) ?? null;

/** Inline content of a paragraph/cell, grouping consecutive link runs into one
 *  w:hyperlink (external → rel + r:id; "#anchor" → w:anchor). */
/** A single inline node wrapped in its hyperlink, if it carries a link mark. */
function inlineUnit(node: PMNode, ctx: ExportCtx): string {
  const inner = inlineXml(node, ctx);
  const href = linkHref(node);
  if (!href || !inner) return inner;
  if (href.startsWith('#')) return `<w:hyperlink w:anchor="${esc(href.slice(1))}">${inner}</w:hyperlink>`;
  const n = ctx.nextId++;
  ctx.rels.push(`<Relationship Id="rId${n}" Type="${R_NS}/hyperlink" Target="${esc(href)}" TargetMode="External"/>`);
  return `<w:hyperlink r:id="rId${n}">${inner}</w:hyperlink>`;
}

/** Inline content with comment-range markers (w:commentRangeStart/End +
 *  w:commentReference) emitted as runs transition across comment ids. */
function inlineContent(node: PMNode, ctx: ExportCtx): string {
  let out = '';
  node.forEach((child) => {
    if (!isInlineLeaf(child)) return;
    const ids = commentIdsOf(child).filter((id) => ctx.knownComments.has(id));
    for (const id of ids) {
      if (!ctx.openComments.has(id)) {
        out += `<w:commentRangeStart w:id="${id}"/>`;
        ctx.openComments.add(id);
      }
    }
    out += inlineUnit(child, ctx);
    const here = ctx.runIdx++;
    for (const id of [...ctx.openComments]) {
      if (ctx.lastRun.get(id) === here) {
        out += `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
        ctx.openComments.delete(id);
      }
    }
  });
  return out;
}

// ── paragraphs ──────────────────────────────────────────────────────

/** A section break's w:sectPr (column flow + break type). Page geometry is
 *  inherited from the body sectPr, so it isn't repeated here. */
function sectionSectPr(s: SectionConfig): string {
  const type = `<w:type w:val="${s.newPage ? 'nextPage' : 'continuous'}"/>`;
  const cols =
    s.columns.count > 1
      ? `<w:cols w:num="${s.columns.count}" w:space="${pxToTwips(s.columns.gap)}"/>`
      : `<w:cols w:space="${pxToTwips(s.columns.gap)}"/>`;
  return `<w:sectPr>${type}${cols}</w:sectPr>`;
}

/** A paragraph's w:pPr children (no wrapper). */
function paraProps(node: PMNode): string {
  const a = node.attrs;
  const out: string[] = [];
  // w:pStyle is the first pPr child. (A round-tripped doc carries the matching
  // "HeadingN" style def; authored-from-scratch headings need styles.xml regen
  // to render the style in Word — a follow-up, like list numbering.xml.)
  const heading = a['heading'] as number | null;
  if (heading) out.push(`<w:pStyle w:val="Heading${heading}"/>`);
  if (a['pageBreakBefore']) out.push('<w:pageBreakBefore/>');
  const list = a['list'] as { numId: string; level: number } | null;
  if (list) out.push(`<w:numPr><w:ilvl w:val="${list.level}"/><w:numId w:val="${esc(list.numId)}"/></w:numPr>`);
  const sp = a['spacing'] as { before?: number; after?: number; line?: number; lineRule?: string } | null;
  if (sp) {
    const at: string[] = [];
    if (sp.before != null) at.push(`w:before="${pxToTwips(sp.before)}"`);
    if (sp.after != null) at.push(`w:after="${pxToTwips(sp.after)}"`);
    if (sp.line != null) {
      const auto = sp.lineRule === 'auto' || sp.lineRule == null;
      at.push(`w:line="${auto ? Math.round(sp.line * 240) : pxToTwips(sp.line)}"`);
      at.push(`w:lineRule="${sp.lineRule ?? 'auto'}"`);
    }
    if (at.length) out.push(`<w:spacing ${at.join(' ')}/>`);
  }
  const ind = a['indent'] as { left?: number; right?: number; firstLine?: number; hanging?: number } | null;
  if (ind) {
    const at: string[] = [];
    if (ind.left != null) at.push(`w:left="${pxToTwips(ind.left)}"`);
    if (ind.right != null) at.push(`w:right="${pxToTwips(ind.right)}"`);
    if (ind.hanging != null) at.push(`w:hanging="${pxToTwips(ind.hanging)}"`);
    else if (ind.firstLine != null) at.push(`w:firstLine="${pxToTwips(ind.firstLine)}"`);
    if (at.length) out.push(`<w:ind ${at.join(' ')}/>`);
  }
  const align = a['align'] as string | null;
  if (align) out.push(`<w:jc w:val="${align === 'justify' ? 'both' : align}"/>`);
  return out.join('');
}

/** `sectPr` (a section break) appends inside this paragraph's pPr, last. */
function paragraphXml(node: PMNode, ctx: ExportCtx, sectPr = ''): string {
  const props = paraProps(node) + sectPr;
  const pPr = props ? `<w:pPr>${props}</w:pPr>` : '';
  return `<w:p>${pPr}${inlineContent(node, ctx)}</w:p>`;
}

// ── tables ──────────────────────────────────────────────────────────

const TABLE_SIDES = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const;
const CELL_SIDES = ['top', 'bottom', 'left', 'right'] as const;

const BORDER_STYLE_OUT: Record<BorderStyle, string> = {
  solid: 'single',
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
};

function bordersXml(tag: string, borders: TableBorders, sides: readonly string[]): string {
  const inner = sides
    .filter((s) => s in borders)
    .map((s) => {
      const side = borders[s as keyof TableBorders];
      if (!side) return `<w:${s} w:val="nil"/>`;
      const sz = Math.max(2, Math.round(side.width * 6)); // px → eighths of a point
      const color = side.color === '#b0b0b0' ? 'auto' : side.color.replace(/^#/, '');
      return `<w:${s} w:val="${BORDER_STYLE_OUT[side.style] ?? 'single'}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
    })
    .join('');
  return inner ? `<${tag}>${inner}</${tag}>` : '';
}

function cellXml(cell: PMNode, ctx: ExportCtx): string {
  const a = cell.attrs;
  const pr: string[] = [];
  const colwidth = a['colwidth'] as number[] | null;
  if (colwidth?.length) pr.push(`<w:tcW w:w="${pxToTwips(colwidth.reduce((x, y) => x + y, 0))}" w:type="dxa"/>`);
  if ((a['colspan'] as number) > 1) pr.push(`<w:gridSpan w:val="${a['colspan']}"/>`);
  const borders = a['borders'] as TableBorders | null;
  if (borders) pr.push(bordersXml('w:tcBorders', borders, CELL_SIDES));
  const bg = a['background'] as string | null;
  if (bg) pr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${bg.replace(/^#/, '')}"/>`);
  const vAlign = a['vAlign'] as string | null;
  if (vAlign) pr.push(`<w:vAlign w:val="${vAlign}"/>`);
  let content = '';
  cell.forEach((b) => (content += blockXml(b, ctx)));
  if (!content) content = '<w:p/>'; // a cell must contain at least one block
  return `<w:tc><w:tcPr>${pr.join('')}</w:tcPr>${content}</w:tc>`;
}

function rowXml(row: PMNode, ctx: ExportCtx): string {
  const pr: string[] = [];
  if (row.attrs['header']) pr.push('<w:tblHeader/>');
  if (row.attrs['cantSplit']) pr.push('<w:cantSplit/>');
  const h = row.attrs['height'] as { value: number; exact: boolean } | null;
  if (h) pr.push(`<w:trHeight w:val="${pxToTwips(h.value)}" w:hRule="${h.exact ? 'exact' : 'atLeast'}"/>`);
  const trPr = pr.length ? `<w:trPr>${pr.join('')}</w:trPr>` : '';
  let cells = '';
  row.forEach((c) => (cells += cellXml(c, ctx)));
  return `<w:tr>${trPr}${cells}</w:tr>`;
}

function tableXml(node: PMNode, ctx: ExportCtx): string {
  const a = node.attrs;
  const pr: string[] = [];
  if (a['align']) pr.push(`<w:jc w:val="${a['align']}"/>`);
  const borders = a['borders'] as TableBorders | null;
  if (borders) pr.push(bordersXml('w:tblBorders', borders, TABLE_SIDES));
  const pad = a['cellPadding'] as { left?: number; right?: number; top?: number; bottom?: number } | null;
  if (pad) {
    const m = (['top', 'left', 'bottom', 'right'] as const)
      .filter((s) => pad[s] != null)
      .map((s) => `<w:${s} w:w="${pxToTwips(pad[s] as number)}" w:type="dxa"/>`)
      .join('');
    if (m) pr.push(`<w:tblCellMar>${m}</w:tblCellMar>`);
  }
  // Grid columns from the first row's cell widths (flattened across spans).
  const firstRow = node.firstChild;
  const grid: number[] = [];
  firstRow?.forEach((c) => (c.attrs['colwidth'] as number[] | null)?.forEach((w) => grid.push(w)));
  const gridXml = grid.length ? `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${pxToTwips(w)}"/>`).join('')}</w:tblGrid>` : '';
  let rows = '';
  node.forEach((r) => (rows += rowXml(r, ctx)));
  return `<w:tbl><w:tblPr>${pr.join('')}</w:tblPr>${gridXml}${rows}</w:tbl>`;
}

/** A top-level block → its OOXML. A `sectPr` marks a section break ending at
 *  this block: it goes inside a paragraph's pPr, or in a trailing empty
 *  paragraph after a table. */
function blockXml(node: PMNode, ctx: ExportCtx, sectPr = ''): string {
  if (node.type.name === 'paragraph') return paragraphXml(node, ctx, sectPr);
  let out = node.type.name === 'table' ? tableXml(node, ctx) : '';
  if (sectPr) out += `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;
  return out;
}

/** Block index → section-break sectPr, for every section but the last (whose
 *  properties live in the body sectPr). */
function sectionBoundaries(doc: PMNode): Map<number, string> {
  const sections = doc.attrs['sections'] as SectionConfig[] | null;
  const out = new Map<number, string>();
  if (!sections || sections.length < 2) return out;
  let acc = 0;
  for (let i = 0; i < sections.length - 1; i++) {
    acc += sections[i].blockCount;
    out.set(acc - 1, sectionSectPr(sections[i])); // last block of section i
  }
  return out;
}

// ── packaging ───────────────────────────────────────────────────────

// ── comments (E3) ───────────────────────────────────────────────────

/** A comment body (commentSchema doc JSON) → w:p runs; the first paragraph
 *  carries `firstParaId`. Mentions serialize as plain "@label" text. */
function commentBodyXml(body: unknown, firstParaId: string): string {
  let doc: PMNode | null = null;
  try {
    doc = commentSchema.nodeFromJSON(body);
  } catch {
    doc = null;
  }
  if (!doc) return `<w:p w14:paraId="${firstParaId}"/>`;
  const ps: string[] = [];
  doc.forEach((p, _o, i) => {
    let runs = '';
    p.forEach((inline) => {
      if (inline.isText) runs += `<w:r><w:t xml:space="preserve">${esc(inline.text ?? '')}</w:t></w:r>`;
      else if (inline.type.name === 'mention') runs += `<w:r><w:t xml:space="preserve">@${esc(String(inline.attrs['label'] ?? ''))}</w:t></w:r>`;
    });
    ps.push(`<w:p${i === 0 ? ` w14:paraId="${firstParaId}"` : ''}>${runs}</w:p>`);
  });
  return ps.join('') || `<w:p w14:paraId="${firstParaId}"/>`;
}

const initialsOf = (name: string): string =>
  name.trim().split(/\s+/).map((w) => w[0] ?? '').join('').slice(0, 3).toUpperCase();

function commentsXml(comments: CommentNode[]): string {
  const body = comments
    .map(
      (c) =>
        `<w:comment w:id="${c.id}" w:author="${esc(c.user.name)}" w:date="${esc(c.date)}" w:initials="${esc(initialsOf(c.user.name))}">` +
        `${commentBodyXml(c.body, paraId(c.id))}</w:comment>`,
    )
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">${body}</w:comments>`;
}

function commentsExtendedXml(comments: CommentNode[]): string {
  const body = comments
    .map((c) => {
      const parent = c.parentId != null ? ` w15:paraIdParent="${paraId(c.parentId)}"` : '';
      return `<w15:commentEx w15:paraId="${paraId(c.id)}"${parent} w15:done="${c.resolved ? 1 : 0}"/>`;
    })
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w15:commentsEx xmlns:w15="${W15_NS}">${body}</w15:commentsEx>`;
}

// ── packaging ───────────────────────────────────────────────────────

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`;

function contentTypes(exts: Set<string>, hasComments: boolean): string {
  const parts = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
  ];
  for (const ext of exts) parts.push(`<Default Extension="${ext}" ContentType="image/${ext === 'jpg' ? 'jpeg' : ext}"/>`);
  parts.push('<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>');
  if (hasComments) {
    parts.push('<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>');
    parts.push('<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>');
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CT_NS}">${parts.join('')}</Types>`;
}

/** Ensure image content-type defaults + comment overrides exist in an original
 *  [Content_Types].xml (E4 merge). */
function mergeContentTypes(xml: string, exts: Set<string>, hasComments: boolean): string {
  let out = xml;
  const add = (frag: string, key: string) => {
    if (!out.includes(`"${key}"`)) out = out.replace('</Types>', `${frag}</Types>`);
  };
  for (const ext of exts) add(`<Default Extension="${ext}" ContentType="image/${ext === 'jpg' ? 'jpeg' : ext}"/>`, ext);
  if (hasComments) {
    add('<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>', '/word/comments.xml');
    add('<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>', '/word/commentsExtended.xml');
  }
  return out;
}

/** The body-level w:sectPr from the original document.xml (page geometry +
 *  header/footer references). It's the last sectPr in the file (after every
 *  paragraph + any section-break sectPr), so re-attaching it keeps page setup
 *  and headers/footers — whose parts + rels are carried — wired up. */
function extractBodySectPr(xml: string): string {
  const all = xml.match(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/g);
  return all ? all[all.length - 1] : '';
}

/** Original document rels minus any comment(sExtended) rels (regenerated),
 *  plus the freshly-emitted rels (E4 merge). */
function mergeRels(xml: string | undefined, newRels: string[]): string {
  const base = (xml ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PR_NS}"></Relationships>`).replace(
    /<Relationship\b[^>]*Target="comments(?:Extended)?\.xml"[^>]*\/>/g,
    '',
  );
  return base.replace('</Relationships>', `${newRels.join('')}</Relationships>`);
}

/**
 * Serialise a bapbong document back to `.docx` bytes.
 *
 * With `opts.carry` (the source package from `importDocx().raw`) the original
 * parts bapbong doesn't model yet — styles, numbering, headers/footers,
 * footnotes, settings, theme, media — are preserved; only document.xml and the
 * comment parts are regenerated. Without it, a minimal package is built from
 * scratch (E1–E3 content only).
 */
export async function exportDocx(doc: PMNode, opts?: { carry?: JSZip }): Promise<Uint8Array> {
  const comments = (doc.attrs['comments'] as CommentNode[] | null) ?? [];
  // Precompute the last inline-leaf index each comment id covers, in document
  // order, so inlineContent can close the range at the right run.
  const knownComments = new Set(comments.map((c) => c.id));
  const lastRun = new Map<number, number>();
  let idx = 0;
  doc.descendants((n) => {
    if (!isInlineLeaf(n)) return;
    for (const id of commentIdsOf(n)) if (knownComments.has(id)) lastRun.set(id, idx);
    idx++;
  });

  const ctx: ExportCtx = {
    rels: [],
    media: [],
    exts: new Set(),
    nextId: 100,
    knownComments,
    lastRun,
    openComments: new Set(),
    runIdx: 0,
  };
  const boundaries = sectionBoundaries(doc);
  let body = '';
  doc.forEach((block, _offset, i) => (body += blockXml(block, ctx, boundaries.get(i))));

  const hasComments = comments.length > 0;
  if (hasComments) {
    ctx.rels.push(`<Relationship Id="rIdComments" Type="${R_NS}/comments" Target="comments.xml"/>`);
    ctx.rels.push(`<Relationship Id="rIdCommentsExt" Type="${R_NS}/commentsExtended" Target="commentsExtended.xml"/>`);
  }

  const zip = new JSZip();
  let sectPr = ''; // re-attached from the original (carry) for page setup + headers
  if (opts?.carry) {
    // E4: start from the original package so unmodelled parts survive.
    const carry = opts.carry;
    for (const [path, f] of Object.entries(carry.files)) {
      if (!f.dir) zip.file(path, await f.async('uint8array'));
    }
    const ct = await carry.file('[Content_Types].xml')?.async('string');
    zip.file('[Content_Types].xml', ct ? mergeContentTypes(ct, ctx.exts, hasComments) : contentTypes(ctx.exts, hasComments));
    const rels = await carry.file('word/_rels/document.xml.rels')?.async('string');
    zip.file('word/_rels/document.xml.rels', mergeRels(rels, ctx.rels));
    const origDoc = await carry.file('word/document.xml')?.async('string');
    if (origDoc) sectPr = extractBodySectPr(origDoc);
  } else {
    zip.file('[Content_Types].xml', contentTypes(ctx.exts, hasComments));
    zip.file('_rels/.rels', ROOT_RELS);
    if (ctx.rels.length) {
      zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PR_NS}">${ctx.rels.join('')}</Relationships>`);
    }
  }

  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" xmlns:wps="${WPS_NS}">` +
    `<w:body>${body}${sectPr}</w:body></w:document>`;
  zip.file('word/document.xml', documentXml);
  if (hasComments) {
    zip.file('word/comments.xml', commentsXml(comments));
    zip.file('word/commentsExtended.xml', commentsExtendedXml(comments));
  }
  for (const { path, base64 } of ctx.media) zip.file(path, base64, { base64: true });
  return zip.generateAsync({ type: 'uint8array' });
}
