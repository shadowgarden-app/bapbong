import JSZip from 'jszip';
import type { Mark, Node as PMNode } from 'prosemirror-model';

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
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

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
}

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

function imageXml(node: PMNode, ctx: ExportCtx): string {
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
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  );
}

/** One inline node → its run XML (excluding the link wrapper). */
function inlineXml(node: PMNode, ctx: ExportCtx): string {
  if (node.type.name === 'hard_break') return '<w:r><w:br/></w:r>';
  if (node.type.name === 'image') return imageXml(node, ctx);
  if (node.isText) {
    return `<w:r>${runProps(node.marks)}<w:t xml:space="preserve">${esc(node.text ?? '')}</w:t></w:r>`;
  }
  return '';
}

const linkHref = (node: PMNode): string | null =>
  (node.marks.find((m) => m.type.name === 'link')?.attrs['href'] as string | undefined) ?? null;

/** Inline content of a paragraph/cell, grouping consecutive link runs into one
 *  w:hyperlink (external → rel + r:id; "#anchor" → w:anchor). */
function inlineContent(node: PMNode, ctx: ExportCtx): string {
  const kids: PMNode[] = [];
  node.forEach((c) => kids.push(c));
  let out = '';
  for (let i = 0; i < kids.length; ) {
    const href = linkHref(kids[i]);
    if (href) {
      let j = i;
      let inner = '';
      while (j < kids.length && linkHref(kids[j]) === href) inner += inlineXml(kids[j++], ctx);
      if (href.startsWith('#')) {
        out += `<w:hyperlink w:anchor="${esc(href.slice(1))}">${inner}</w:hyperlink>`;
      } else {
        const n = ctx.nextId++;
        ctx.rels.push(`<Relationship Id="rId${n}" Type="${R_NS}/hyperlink" Target="${esc(href)}" TargetMode="External"/>`);
        out += `<w:hyperlink r:id="rId${n}">${inner}</w:hyperlink>`;
      }
      i = j;
    } else {
      out += inlineXml(kids[i++], ctx);
    }
  }
  return out;
}

// ── paragraphs ──────────────────────────────────────────────────────

function paraProps(node: PMNode): string {
  const a = node.attrs;
  const out: string[] = [];
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
  return out.length ? `<w:pPr>${out.join('')}</w:pPr>` : '';
}

function paragraphXml(node: PMNode, ctx: ExportCtx): string {
  return `<w:p>${paraProps(node)}${inlineContent(node, ctx)}</w:p>`;
}

// ── tables ──────────────────────────────────────────────────────────

const TABLE_SIDES = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const;
const CELL_SIDES = ['top', 'bottom', 'left', 'right'] as const;

function bordersXml(tag: string, borders: Record<string, boolean>, sides: readonly string[]): string {
  const inner = sides
    .filter((s) => s in borders)
    .map((s) => (borders[s] ? `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>` : `<w:${s} w:val="nil"/>`))
    .join('');
  return inner ? `<${tag}>${inner}</${tag}>` : '';
}

function cellXml(cell: PMNode, ctx: ExportCtx): string {
  const a = cell.attrs;
  const pr: string[] = [];
  const colwidth = a['colwidth'] as number[] | null;
  if (colwidth?.length) pr.push(`<w:tcW w:w="${pxToTwips(colwidth.reduce((x, y) => x + y, 0))}" w:type="dxa"/>`);
  if ((a['colspan'] as number) > 1) pr.push(`<w:gridSpan w:val="${a['colspan']}"/>`);
  const borders = a['borders'] as Record<string, boolean> | null;
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
  const borders = a['borders'] as Record<string, boolean> | null;
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

/** A top-level block → its OOXML. */
function blockXml(node: PMNode, ctx: ExportCtx): string {
  if (node.type.name === 'paragraph') return paragraphXml(node, ctx);
  if (node.type.name === 'table') return tableXml(node, ctx);
  return '';
}

// ── packaging ───────────────────────────────────────────────────────

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`;

function contentTypes(exts: Set<string>): string {
  const defaults = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>'];
  for (const ext of exts) defaults.push(`<Default Extension="${ext}" ContentType="image/${ext === 'jpg' ? 'jpeg' : ext}"/>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CT_NS}">${defaults.join('')}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
}

/**
 * Serialise a bapbong document back to `.docx` bytes (E1 + E2). Comments and the
 * original imported parts are added in later phases.
 */
export async function exportDocx(doc: PMNode): Promise<Uint8Array> {
  const ctx: ExportCtx = { rels: [], media: [], exts: new Set(), nextId: 100 };
  let body = '';
  doc.forEach((block) => (body += blockXml(block, ctx)));

  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">` +
    `<w:body>${body}</w:body></w:document>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes(ctx.exts));
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('word/document.xml', documentXml);
  if (ctx.rels.length) {
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PR_NS}">${ctx.rels.join('')}</Relationships>`,
    );
  }
  for (const { path, base64 } of ctx.media) zip.file(path, base64, { base64: true });
  return zip.generateAsync({ type: 'uint8array' });
}
