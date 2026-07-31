import JSZip from 'jszip';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { perf } from '@shadow-garden/bapbong-contracts';
import { commentSchema } from '@shadow-garden/bapbong-model';
import type {
  BorderStyle,
  CommentNode,
  PageConfig,
  SectionConfig,
  ShapeSpec,
  TableBorders,
} from '@shadow-garden/bapbong-contracts';
import { parsePageGeometry } from './docx.js';
import { child, parseXml } from './ooxml.js';

/**
 * DOCX export (round-trip). Phases:
 *  E1 ✅ paragraphs + runs + common character marks + pPr + hard breaks.
 *  E2 ✅ lists (numPr), tables, inline images (+ media + rels), hyperlinks (+ rels).
 *  E3 ⬜ comments (ranges + comments.xml + commentsExtended.xml).
 *  E4 ⬜ carry the original imported parts (styles/numbering/headers/media).
 */

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const WPS_NS =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
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
  /** Doc numId → output w:numId. Editor-authored ids (`bb-*`) are remapped to
   *  fresh integers backed by generated numbering.xml defs; carried ids pass
   *  through. */
  numIdMap: Map<string, string>;
  // Comment ranges (E3): the doc's known comment ids, the last inline-node index
  // each id covers (so we can close the range), the currently-open ids, and a
  // running inline-node index aligned with the precompute.
  knownComments: Set<number>;
  lastRun: Map<number, number>;
  openComments: Set<number>;
  runIdx: number;
}

const commentIdsOf = (node: PMNode): number[] =>
  (node.marks.find((m) => m.type.name === 'comment')?.attrs['ids'] as
    | number[]
    | undefined) ?? [];

/** Is this an inline leaf the comment-range index counts (text / image / break)? */
const isInlineLeaf = (node: PMNode): boolean =>
  node.isText ||
  node.type.name === 'image' ||
  node.type.name === 'hard_break' ||
  // Without this, inlineContent skips page_field entirely and PAGE/NUMPAGES
  // vanish from the body on export. Both the comment-range precompute and the
  // writer use this predicate, so they stay in step.
  node.type.name === 'page_field';

function esc(s: string): string {
  return s.replace(
    /[&<>"]/g,
    (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string,
  );
}

// Every MIME the importer accepts (docx.ts mimeOf) must map here — an
// unmapped type would fall back to a .png extension around foreign bytes,
// and Word shows a missing-image box for the mislabelled part.
const MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/jpg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
  'image/tiff': 'tiff',
  'image/avif': 'avif',
};

/** [Content_Types].xml Default per media extension (inverse of MIME_EXT). */
const EXT_MIME: Record<string, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  tiff: 'image/tiff',
  avif: 'image/avif',
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
  if (hl)
    out.push(
      `<w:shd w:val="clear" w:color="auto" w:fill="${hl.replace(/^#/, '')}"/>`,
    );
  const va = byName.get('vertAlign')?.attrs['value'] as string | undefined;
  if (va)
    out.push(
      `<w:vertAlign w:val="${va === 'sub' ? 'subscript' : 'superscript'}"/>`,
    );
  return out.length ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

/** Float attrs → the wp:anchor wrapper (position + wrap) around a graphic. */
function anchorXml(
  float: Record<string, unknown>,
  cx: number,
  cy: number,
  n: number,
  graphic: string,
  docPr?: string,
): string {
  const dist = (k: string) => pxToEmu((float[k] as number) ?? 0);
  const hRel = float['hRel'] === 'page' ? 'page' : 'column';
  const vRel =
    float['vRel'] === 'page'
      ? 'page'
      : float['vRel'] === 'margin'
        ? 'margin'
        : 'paragraph';
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
    `simplePos="0" relativeHeight="251658240" behindDoc="${float['behind'] ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${hRel}">${posH}</wp:positionH>` +
    `<wp:positionV relativeFrom="${vRel}">${posV}</wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    wrap +
    (docPr ?? `<wp:docPr id="${n}" name="Shape ${n}"/>`) +
    `<wp:cNvGraphicFramePr/>` +
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
  const fill = s.fill
    ? `<a:solidFill><a:srgbClr val="${s.fill.replace(/^#/, '')}"/></a:solidFill>`
    : '<a:noFill/>';
  const ln = s.stroke
    ? `<a:ln w="${pxToEmu(s.strokeWidth ?? 1)}"><a:solidFill><a:srgbClr val="${s.stroke.replace(/^#/, '')}"/></a:solidFill></a:ln>`
    : '<a:ln><a:noFill/></a:ln>';
  // Textbox paragraphs ride the node as PM JSON; re-emit them as txbxContent
  // so the text survives the round-trip.
  const tb = node.attrs['textbox'] as {
    paragraphs: unknown[];
    inset?: { l: number; t: number; r: number; b: number };
  } | null;
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
  let n: number;
  let name: string;
  let blip: string;
  if (m) {
    const ext = MIME_EXT[m[1].toLowerCase()] ?? 'png';
    ctx.exts.add(ext);
    n = ctx.nextId++;
    name = `image${n}.${ext}`;
    ctx.media.push({ path: `word/media/${name}`, base64: m[2] });
    ctx.rels.push(
      `<Relationship Id="rId${n}" Type="${R_NS}/image" Target="media/${name}"/>`,
    );
    blip = `<a:blip r:embed="rId${n}"/>`;
  } else if (/^https?:\/\//i.test(src)) {
    // "Insert image from URL": no bytes on hand — write an externally-linked
    // picture rather than silently dropping the node.
    n = ctx.nextId++;
    name = `image${n}`;
    ctx.rels.push(
      `<Relationship Id="rId${n}" Type="${R_NS}/image" Target="${esc(src)}" TargetMode="External"/>`,
    );
    blip = `<a:blip r:link="rId${n}"/>`;
  } else {
    return ''; // blob:/object URL — its bytes are unreachable at export time
  }
  const cx = pxToEmu((node.attrs['width'] as number) ?? 96);
  const cy = pxToEmu((node.attrs['height'] as number) ?? 96);
  // descr is where Word keeps alt text. Dropping it is silent — the image
  // still renders, so nothing looks wrong; only screen readers lose it.
  const alt = (node.attrs['alt'] as string | null) ?? '';
  const descr = alt ? ` descr="${esc(alt)}"` : '';
  const graphic =
    `<a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${n}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm${rotAttr(node)}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>`;
  const docPr = `<wp:docPr id="${n}" name="Picture ${n}"${descr}/>`;
  // A floating BITMAP goes out as wp:anchor like the drawn shapes always did —
  // the unconditional wp:inline here flattened every floating picture to
  // inline on save (shapes were fine; they take the shapeXml path).
  const float = node.attrs['float'] as Record<string, unknown> | null;
  const body = float
    ? anchorXml(float, cx, cy, n, graphic, docPr)
    : `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>${docPr}${graphic}</wp:inline>`;
  return `<w:r><w:drawing>${body}</w:drawing></w:r>`;
}

/** One inline node → its run XML (excluding the link wrapper). */
function inlineXml(node: PMNode, ctx: ExportCtx): string {
  if (node.type.name === 'hard_break') return '<w:r><w:br/></w:r>';
  if (node.type.name === 'image') return imageXml(node, ctx);
  // w:fldSimple, not the three-part w:fldChar dance: it is the form the
  // importer already reads back, and Word recomputes the result on open — the
  // run we emit is only a placeholder that carries the formatting.
  if (node.type.name === 'page_field') {
    const instr = node.attrs['kind'] === 'pages' ? ' NUMPAGES ' : ' PAGE ';
    return `<w:fldSimple w:instr="${instr}"><w:r>${runProps(node.marks)}<w:t>1</w:t></w:r></w:fldSimple>`;
  }
  if (node.isText) {
    const fn = node.marks.find((m) => m.type.name === 'footnote');
    if (fn)
      return `<w:r>${runProps(node.marks)}<w:footnoteReference w:id="${fn.attrs['num']}"/></w:r>`;
    return `<w:r>${runProps(node.marks)}<w:t xml:space="preserve">${esc(node.text ?? '')}</w:t></w:r>`;
  }
  return '';
}

const linkHref = (node: PMNode): string | null =>
  (node.marks.find((m) => m.type.name === 'link')?.attrs['href'] as
    | string
    | undefined) ?? null;

/** Inline content of a paragraph/cell, grouping consecutive link runs into one
 *  w:hyperlink (external → rel + r:id; "#anchor" → w:anchor). */
/** A single inline node wrapped in its hyperlink, if it carries a link mark. */
function inlineUnit(node: PMNode, ctx: ExportCtx): string {
  const inner = inlineXml(node, ctx);
  const href = linkHref(node);
  if (!href || !inner) return inner;
  if (href.startsWith('#'))
    return `<w:hyperlink w:anchor="${esc(href.slice(1))}">${inner}</w:hyperlink>`;
  const n = ctx.nextId++;
  ctx.rels.push(
    `<Relationship Id="rId${n}" Type="${R_NS}/hyperlink" Target="${esc(href)}" TargetMode="External"/>`,
  );
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
function paraProps(node: PMNode, ctx: ExportCtx): string {
  const a = node.attrs;
  const out: string[] = [];
  // w:pStyle is the first pPr child. ensureStyleDefs() guarantees the
  // referenced style exists in styles.xml (generated from scratch, or merged
  // into a carried package that lacks it).
  const heading = a['heading'] as number | null;
  const styleId = a['styleId'] as string | null;
  if (heading) out.push(`<w:pStyle w:val="Heading${heading}"/>`);
  else if (styleId === 'Title' || styleId === 'Subtitle')
    out.push(`<w:pStyle w:val="${styleId}"/>`);
  if (a['pageBreakBefore']) out.push('<w:pageBreakBefore/>');
  const list = a['list'] as { numId: string; level: number } | null;
  if (list) {
    const outId = ctx.numIdMap.get(list.numId) ?? list.numId;
    out.push(
      `<w:numPr><w:ilvl w:val="${list.level}"/><w:numId w:val="${esc(outId)}"/></w:numPr>`,
    );
  }
  // w:pBdr sits between numPr and spacing in the pPr sequence.
  const pBdr = a['borders'] as TableBorders | null;
  if (pBdr) {
    const box = bordersXml('w:pBdr', pBdr, CELL_SIDES);
    if (box) out.push(box);
  }
  // w:tabs sits between pBdr and spacing in the pPr sequence. The importer
  // resolves these through the style cascade into px; they go back as twips.
  const tabs = a['tabs'] as
    | { pos: number; val: string; leader?: string }[]
    | null;
  if (tabs?.length) {
    const stops = tabs
      .map(
        (t) =>
          `<w:tab w:val="${t.val}" w:pos="${pxToTwips(t.pos)}"${
            t.leader ? ` w:leader="${t.leader}"` : ''
          }/>`,
      )
      .join('');
    out.push(`<w:tabs>${stops}</w:tabs>`);
  }
  const sp = a['spacing'] as {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: string;
  } | null;
  if (sp) {
    const at: string[] = [];
    if (sp.before != null) at.push(`w:before="${pxToTwips(sp.before)}"`);
    if (sp.after != null) at.push(`w:after="${pxToTwips(sp.after)}"`);
    if (sp.line != null) {
      const auto = sp.lineRule === 'auto' || sp.lineRule == null;
      at.push(
        `w:line="${auto ? Math.round(sp.line * 240) : pxToTwips(sp.line)}"`,
      );
      at.push(`w:lineRule="${sp.lineRule ?? 'auto'}"`);
    }
    if (at.length) out.push(`<w:spacing ${at.join(' ')}/>`);
  }
  const ind = a['indent'] as {
    left?: number;
    right?: number;
    firstLine?: number;
    hanging?: number;
  } | null;
  if (ind) {
    const at: string[] = [];
    if (ind.left != null) at.push(`w:left="${pxToTwips(ind.left)}"`);
    if (ind.right != null) at.push(`w:right="${pxToTwips(ind.right)}"`);
    if (ind.hanging != null) at.push(`w:hanging="${pxToTwips(ind.hanging)}"`);
    else if (ind.firstLine != null)
      at.push(`w:firstLine="${pxToTwips(ind.firstLine)}"`);
    if (at.length) out.push(`<w:ind ${at.join(' ')}/>`);
  }
  const align = a['align'] as string | null;
  if (align)
    out.push(`<w:jc w:val="${align === 'justify' ? 'both' : align}"/>`);
  return out.join('');
}

/** `sectPr` (a section break) appends inside this paragraph's pPr, last. */
function paragraphXml(node: PMNode, ctx: ExportCtx, sectPr = ''): string {
  const props = paraProps(node, ctx) + sectPr;
  const pPr = props ? `<w:pPr>${props}</w:pPr>` : '';
  return `<w:p>${pPr}${inlineContent(node, ctx)}</w:p>`;
}

// ── tables ──────────────────────────────────────────────────────────

const TABLE_SIDES = [
  'top',
  'bottom',
  'left',
  'right',
  'insideH',
  'insideV',
] as const;
const CELL_SIDES = ['top', 'bottom', 'left', 'right'] as const;

const BORDER_STYLE_OUT: Record<BorderStyle, string> = {
  solid: 'single',
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
};

function bordersXml(
  tag: string,
  borders: TableBorders,
  sides: readonly string[],
): string {
  const inner = sides
    .filter((s) => s in borders)
    .map((s) => {
      const side = borders[s as keyof TableBorders];
      if (!side) return `<w:${s} w:val="nil"/>`;
      const sz = Math.max(2, Math.round(side.width * 6)); // px → eighths of a point
      const color =
        side.color === '#b0b0b0' ? 'auto' : side.color.replace(/^#/, '');
      return `<w:${s} w:val="${BORDER_STYLE_OUT[side.style] ?? 'single'}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
    })
    .join('');
  return inner ? `<${tag}>${inner}</${tag}>` : '';
}

function cellXml(cell: PMNode, ctx: ExportCtx): string {
  const a = cell.attrs;
  const pr: string[] = [];
  const colwidth = a['colwidth'] as number[] | null;
  if (colwidth?.length)
    pr.push(
      `<w:tcW w:w="${pxToTwips(colwidth.reduce((x, y) => x + y, 0))}" w:type="dxa"/>`,
    );
  if ((a['colspan'] as number) > 1)
    pr.push(`<w:gridSpan w:val="${a['colspan']}"/>`);
  // A vertical merge starts here; the rows it covers get their placeholder
  // cells from rowXml. Without this the merge is lost on export — the model
  // carries it as `rowspan`, but OOXML expresses it as a `w:vMerge` on the
  // first cell plus a real `w:tc` in every covered row.
  if ((a['rowspan'] as number) > 1) pr.push('<w:vMerge w:val="restart"/>');
  const borders = a['borders'] as TableBorders | null;
  if (borders) pr.push(bordersXml('w:tcBorders', borders, CELL_SIDES));
  const bg = a['background'] as string | null;
  if (bg)
    pr.push(
      `<w:shd w:val="clear" w:color="auto" w:fill="${bg.replace(/^#/, '')}"/>`,
    );
  const vAlign = a['vAlign'] as string | null;
  if (vAlign) pr.push(`<w:vAlign w:val="${vAlign}"/>`);
  const padding = a['padding'] as {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
  } | null;
  if (padding) {
    const sides = (['top', 'left', 'bottom', 'right'] as const)
      .filter((s) => padding[s] != null)
      .map(
        (s) =>
          `<w:${s} w:type="dxa" w:w="${pxToTwips(padding[s] as number)}"/>`,
      )
      .join('');
    if (sides) pr.push(`<w:tcMar>${sides}</w:tcMar>`);
  }
  let content = '';
  cell.forEach((b) => (content += blockXml(b, ctx)));
  if (!content) content = '<w:p/>'; // a cell must contain at least one block
  return `<w:tc><w:tcPr>${pr.join('')}</w:tcPr>${content}</w:tc>`;
}

/** A vertical merge still open below the row that started it. */
interface PendingMerge {
  /** Rows it still has to cover, not counting the one it started in. */
  rowsLeft: number;
  colspan: number;
  colwidth: number[] | null;
}

/** The placeholder a covered row needs. OOXML has no "this row is shorter"
 *  concept: a merged-away slot is still a real `w:tc`, just one carrying an
 *  empty `w:vMerge`. Omitting it leaves the row with fewer cells than
 *  `w:tblGrid` declares, which is what made Word draw the table ragged. */
function mergeContinuationXml(p: PendingMerge): string {
  const pr: string[] = [];
  if (p.colwidth?.length)
    pr.push(
      `<w:tcW w:w="${pxToTwips(p.colwidth.reduce((x, y) => x + y, 0))}" w:type="dxa"/>`,
    );
  if (p.colspan > 1) pr.push(`<w:gridSpan w:val="${p.colspan}"/>`);
  pr.push('<w:vMerge/>');
  return `<w:tc><w:tcPr>${pr.join('')}</w:tcPr><w:p/></w:tc>`;
}

/** `pending` is carried across rows by the caller — the model absorbs merged
 *  cells into a `rowspan` on the cell above (see the importer's logical grid),
 *  so the rows below have no node for that slot and we have to put one back. */
function rowXml(
  row: PMNode,
  ctx: ExportCtx,
  pending: Map<number, PendingMerge>,
): string {
  const pr: string[] = [];
  if (row.attrs['header']) pr.push('<w:tblHeader/>');
  if (row.attrs['cantSplit']) pr.push('<w:cantSplit/>');
  const h = row.attrs['height'] as { value: number; exact: boolean } | null;
  if (h)
    pr.push(
      `<w:trHeight w:val="${pxToTwips(h.value)}" w:hRule="${h.exact ? 'exact' : 'atLeast'}"/>`,
    );
  const trPr = pr.length ? `<w:trPr>${pr.join('')}</w:trPr>` : '';

  const nodes: PMNode[] = [];
  row.forEach((c) => nodes.push(c));
  let cells = '';
  let col = 0;
  let i = 0;
  // Walk grid columns, not nodes: a covered slot has no node, so the two
  // indices drift apart exactly where the merge is.
  while (i < nodes.length || pending.size) {
    const open = pending.get(col);
    if (open) {
      cells += mergeContinuationXml(open);
      const at = col;
      col += open.colspan;
      if (--open.rowsLeft <= 0) pending.delete(at);
      continue;
    }
    if (i < nodes.length) {
      const c = nodes[i++];
      cells += cellXml(c, ctx);
      const colspan = (c.attrs['colspan'] as number) || 1;
      const rowspan = (c.attrs['rowspan'] as number) || 1;
      if (rowspan > 1)
        pending.set(col, {
          rowsLeft: rowspan - 1,
          colspan,
          colwidth: c.attrs['colwidth'] as number[] | null,
        });
      col += colspan;
      continue;
    }
    // Out of nodes but merges are still open further right — skip the gap
    // rather than invent cells for it.
    const next = [...pending.keys()]
      .filter((k) => k > col)
      .sort((x, y) => x - y)[0];
    if (next === undefined) break;
    col = next;
  }
  return `<w:tr>${trPr}${cells}</w:tr>`;
}

function tableXml(node: PMNode, ctx: ExportCtx): string {
  const a = node.attrs;
  const pr: string[] = [];
  if (a['align']) pr.push(`<w:jc w:val="${a['align']}"/>`);
  const borders = a['borders'] as TableBorders | null;
  if (borders) pr.push(bordersXml('w:tblBorders', borders, TABLE_SIDES));
  const pad = a['cellPadding'] as {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
  } | null;
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
  firstRow?.forEach((c) =>
    (c.attrs['colwidth'] as number[] | null)?.forEach((w) => grid.push(w)),
  );
  const gridXml = grid.length
    ? `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${pxToTwips(w)}"/>`).join('')}</w:tblGrid>`
    : '';
  let rows = '';
  // One map per table, threaded through the rows: a merge opened in row N has
  // to be closed out by rows N+1… — state that cannot live inside rowXml.
  const pending = new Map<number, PendingMerge>();
  node.forEach((r) => (rows += rowXml(r, ctx, pending)));
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
      if (inline.isText)
        runs += `<w:r><w:t xml:space="preserve">${esc(inline.text ?? '')}</w:t></w:r>`;
      else if (inline.type.name === 'mention')
        runs += `<w:r><w:t xml:space="preserve">@${esc(String(inline.attrs['label'] ?? ''))}</w:t></w:r>`;
    });
    ps.push(
      `<w:p${i === 0 ? ` w14:paraId="${firstParaId}"` : ''}>${runs}</w:p>`,
    );
  });
  return ps.join('') || `<w:p w14:paraId="${firstParaId}"/>`;
}

const initialsOf = (name: string): string =>
  name
    .trim()
    .split(/\s+/)
    .map((w) => w[0] ?? '')
    .join('')
    .slice(0, 3)
    .toUpperCase();

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
      const parent =
        c.parentId != null ? ` w15:paraIdParent="${paraId(c.parentId)}"` : '';
      return `<w15:commentEx w15:paraId="${paraId(c.id)}"${parent} w15:done="${c.resolved ? 1 : 0}"/>`;
    })
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w15:commentsEx xmlns:w15="${W15_NS}">${body}</w15:commentsEx>`;
}

// ── packaging ───────────────────────────────────────────────────────

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`;

// ── styles.xml ──────────────────────────────────────────────────────
// Definitions for every style id a w:pStyle can reference. Sizes are
// half-points and mirror the layout engine's defaults (HEADING_PT + the
// Title/Subtitle run bases), so Word renders what the canvas showed.

const STYLE_DEFS: Record<string, string> = (() => {
  const heading = (level: number, halfPt: number) =>
    `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
    `<w:rPr><w:b/><w:sz w:val="${halfPt}"/><w:szCs w:val="${halfPt}"/></w:rPr></w:style>`;
  return {
    Heading1: heading(1, 48),
    Heading2: heading(2, 36),
    Heading3: heading(3, 28),
    Heading4: heading(4, 24),
    Heading5: heading(5, 22),
    Heading6: heading(6, 22),
    Title:
      `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
      `<w:pPr><w:spacing w:after="80"/></w:pPr><w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>`,
    Subtitle:
      `<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
      `<w:rPr><w:i/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>`,
  };
})();

/** Every style id the document's paragraphs reference via w:pStyle. */
function usedStyleIds(doc: PMNode): Set<string> {
  const used = new Set<string>();
  doc.descendants((n) => {
    if (n.type.name !== 'paragraph') return;
    const heading = n.attrs['heading'] as number | null;
    const styleId = n.attrs['styleId'] as string | null;
    if (heading) used.add(`Heading${heading}`);
    else if (styleId && STYLE_DEFS[styleId]) used.add(styleId);
  });
  return used;
}

/** A from-scratch word/styles.xml: docDefaults + Normal + the used defs. */
function stylesXml(used: Set<string>): string {
  const defs = [...used].map((id) => STYLE_DEFS[id]).join('');
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${W_NS}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    defs +
    `</w:styles>`
  );
}

/** Append any used-but-missing style defs to a carried styles.xml, so
 *  headings/Title/Subtitle authored in bapbong render styled in Word even
 *  when the source document never defined them. */
function mergeStyles(xml: string, used: Set<string>): string {
  const missing = [...used].filter((id) => !xml.includes(`w:styleId="${id}"`));
  if (!missing.length) return xml;
  return xml.replace(
    '</w:styles>',
    `${missing.map((id) => STYLE_DEFS[id]).join('')}</w:styles>`,
  );
}

function contentTypes(
  exts: Set<string>,
  hasComments: boolean,
  hasNumbering = false,
): string {
  const parts = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
  ];
  for (const ext of exts)
    parts.push(
      `<Default Extension="${ext}" ContentType="${EXT_MIME[ext] ?? `image/${ext}`}"/>`,
    );
  parts.push(
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  );
  parts.push(
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
  );
  if (hasNumbering) parts.push(NUMBERING_OVERRIDE);
  if (hasComments) {
    parts.push(
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
    );
    parts.push(
      '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>',
    );
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CT_NS}">${parts.join('')}</Types>`;
}

/** Ensure image content-type defaults + comment overrides exist in an original
 *  [Content_Types].xml (E4 merge). */
function mergeContentTypes(
  xml: string,
  exts: Set<string>,
  hasComments: boolean,
  hasNewNumbering = false,
): string {
  let out = xml;
  const add = (frag: string, key: string) => {
    if (!out.includes(`"${key}"`))
      out = out.replace('</Types>', `${frag}</Types>`);
  };
  if (hasNewNumbering) add(NUMBERING_OVERRIDE, '/word/numbering.xml');
  for (const ext of exts)
    add(
      `<Default Extension="${ext}" ContentType="${EXT_MIME[ext] ?? `image/${ext}`}"/>`,
      ext,
    );
  if (hasComments) {
    add(
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
      '/word/comments.xml',
    );
    add(
      '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>',
      '/word/commentsExtended.xml',
    );
  }
  return out;
}

/** The body-level w:sectPr from the original document.xml (page geometry +
 *  header/footer references). It's the last sectPr in the file (after every
 *  paragraph + any section-break sectPr), so re-attaching it keeps page setup
 *  and headers/footers — whose parts + rels are carried — wired up. */
function extractBodySectPr(xml: string): string {
  const all = xml.match(
    /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/g,
  );
  return all ? all[all.length - 1] : '';
}

// ── Page geometry (w:pgSz / w:pgMar) ────────────────────────────────

/** A4 @96dpi with 1in margins — what layout shows when a doc carries no page
 *  attr, so export must emit the same (Word's own default is Letter). */
const A4_PAGE: PageConfig = {
  width: 794,
  height: 1123,
  margin: { top: 96, right: 96, bottom: 96, left: 96 },
};

const PGSZ_RX = /<w:pgSz\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/w:pgSz>)/;
const PGMAR_RX = /<w:pgMar\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/w:pgMar>)/;

/** Canonical twip dimensions for common paper sizes, keyed by portrait px
 *  size. Emitting the canonical value (not px×15, which lands a few twips off
 *  after import rounding) keeps Word's paper-size dropdown naming the size
 *  ("A4") instead of showing "Custom". */
const PAPER_TWIPS: Record<string, [number, number]> = {
  '794x1123': [11906, 16838], // A4
  '816x1056': [12240, 15840], // Letter
  '816x1344': [12240, 20160], // Legal
  '1123x1587': [16838, 23811], // A3
  '559x794': [8391, 11906], // A5
};

/** w:pgSz for the modelled page (px→twips); landscape swaps the emitted
 *  dimensions and rides w:orient, the shape Word itself writes. */
function pgSzXml(page: PageConfig): string {
  const landscape = page.width > page.height;
  const [pw, ph] = landscape
    ? [page.height, page.width]
    : [page.width, page.height];
  const [tw, th] = PAPER_TWIPS[`${pw}x${ph}`] ?? [pxToTwips(pw), pxToTwips(ph)];
  const [w, h] = landscape ? [th, tw] : [tw, th];
  return `<w:pgSz w:w="${w}" w:h="${h}"${landscape ? ' w:orient="landscape"' : ''}/>`;
}

/** w:pgMar for the modelled margins. `keep` preserves the original pgMar's
 *  header/footer/gutter distances (not modelled); Word defaults otherwise. */
function pgMarXml(
  m: PageConfig['margin'],
  keep?: { header?: string; footer?: string; gutter?: string },
): string {
  return (
    `<w:pgMar w:top="${pxToTwips(m.top)}" w:right="${pxToTwips(m.right)}"` +
    ` w:bottom="${pxToTwips(m.bottom)}" w:left="${pxToTwips(m.left)}"` +
    ` w:header="${keep?.header ?? '720'}" w:footer="${keep?.footer ?? '720'}"` +
    ` w:gutter="${keep?.gutter ?? '0'}"/>`
  );
}

/** True when the carried sectPr's geometry equals the modelled one — i.e. the
 *  user never touched page setup. Compared in px through the importer's own
 *  parser so px↔twips rounding can't produce a false "edited". */
function sectPrMatchesPage(sectPr: string, page: PageConfig): boolean {
  const parsed = parsePageGeometry(child(parseXml(sectPr), 'w:sectPr'));
  return (
    parsed.width === page.width &&
    parsed.height === page.height &&
    parsed.margin.top === page.margin.top &&
    parsed.margin.right === page.margin.right &&
    parsed.margin.bottom === page.margin.bottom &&
    parsed.margin.left === page.margin.left
  );
}

/** Replace the carried body sectPr's w:pgSz/w:pgMar with the modelled page
 *  geometry, in place (child order — headerReference, type, pgSz, pgMar,
 *  cols… — is schema-significant). Only called on a real page-setup edit; an
 *  untouched doc keeps its original bytes (px↔twips rounding would otherwise
 *  drift values on every save). */
function splicePageGeometry(sectPr: string, page: PageConfig): string {
  // A childless self-closing sectPr needs a slot to insert into.
  const self = /^<w:sectPr\b([^>]*)\/>\s*$/.exec(sectPr.trim());
  let out = self ? `<w:sectPr${self[1]}></w:sectPr>` : sectPr;
  const oldMar = PGMAR_RX.exec(out)?.[0] ?? '';
  const keepAttr = (name: string) =>
    new RegExp(`\\bw:${name}="([^"]*)"`).exec(oldMar)?.[1];
  const sz = pgSzXml(page);
  const mar = pgMarXml(page.margin, {
    header: keepAttr('header'),
    footer: keepAttr('footer'),
    gutter: keepAttr('gutter'),
  });
  out = PGSZ_RX.test(out)
    ? out.replace(PGSZ_RX, sz)
    : PGMAR_RX.test(out)
      ? out.replace(PGMAR_RX, (m) => sz + m)
      : out.replace('</w:sectPr>', `${sz}</w:sectPr>`);
  out = PGMAR_RX.test(out)
    ? out.replace(PGMAR_RX, mar)
    : out.replace(sz, sz + mar);
  return out;
}

/** Original document rels minus any comment(sExtended) rels (regenerated),
 *  plus the freshly-emitted rels (E4 merge). */
function mergeRels(xml: string | undefined, newRels: string[]): string {
  const base = (
    xml ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PR_NS}"></Relationships>`
  ).replace(
    /<Relationship\b[^>]*Target="comments(?:Extended)?\.xml"[^>]*\/>/g,
    '',
  );
  return base.replace(
    '</Relationships>',
    `${newRels.join('')}</Relationships>`,
  );
}

// ── Numbering export ─────────────────────────────────────────────────
// Editor-authored lists carry string numIds (`bb-bullet`, `bb-ordered-paren`)
// whose definitions live only on the doc node's `numbering` attr — invalid in
// OOXML (w:numId must be an integer) and undefined in Word. At export they are
// remapped to fresh integer ids and their defs regenerated into
// word/numbering.xml (merged into a carried part, or a new part from scratch),
// so the file opens in Word with the same markers bapbong paints.

type NumberingDefEntry = {
  key: string;
  levels: Record<
    number,
    { numFmt: string; lvlText: string; start?: number } | undefined
  >;
};

interface NumberingPlan {
  /** Doc numId → output w:numId (identity for ids the carried part covers). */
  map: Map<string, string>;
  /** Generated `<w:abstractNum>` / `<w:num>` fragments ('' when none). */
  abstractXml: string;
  numXml: string;
}

/** One abstractNum from a doc-attr definition. `w:suff="space"` matches the
 *  marker-space-text layout bapbong paints (Word's default suffix is a tab). */
function abstractNumXml(absId: number, def: NumberingDefEntry): string {
  const lvls = Object.keys(def.levels)
    .map(Number)
    .sort((a, b) => a - b)
    .map((ilvl) => {
      const l = def.levels[ilvl];
      if (!l) return '';
      return (
        `<w:lvl w:ilvl="${ilvl}"><w:start w:val="${l.start ?? 1}"/>` +
        `<w:numFmt w:val="${esc(l.numFmt)}"/><w:suff w:val="space"/>` +
        `<w:lvlText w:val="${esc(l.lvlText)}"/><w:lvlJc w:val="left"/></w:lvl>`
      );
    })
    .join('');
  return `<w:abstractNum w:abstractNumId="${absId}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls}</w:abstractNum>`;
}

/** Decide which numIds the output file must define, and mint their ids. */
function planNumbering(doc: PMNode, carried: string | null): NumberingPlan {
  const used = new Set<string>();
  doc.descendants((n) => {
    const list = n.attrs['list'] as { numId?: string } | null | undefined;
    if (list?.numId) used.add(list.numId);
  });
  const map = new Map<string, string>();
  if (used.size === 0) return { map, abstractXml: '', numXml: '' };

  const defs =
    (doc.attrs['numbering'] as Record<string, NumberingDefEntry> | null) ?? {};
  const inCarried = (id: string) =>
    carried != null && new RegExp(`<w:num w:numId="${id}"[ />]`).test(carried);

  const generate: string[] = [];
  for (const id of used) {
    if (inCarried(id) || !defs[id])
      map.set(id, id); // covered, or no def (degraded passthrough)
    else generate.push(id);
  }
  if (generate.length === 0) return { map, abstractXml: '', numXml: '' };

  // Mint ids above everything the carried part uses (numId and abstractNumId
  // share one counter for simplicity — the namespaces are independent, so
  // this only costs unused integers) and above kept integer ids.
  const taken = new Set<number>([0]);
  if (carried)
    for (const m of carried.matchAll(/w:(?:numId|abstractNumId)="(\d+)"/g))
      taken.add(Number(m[1]));
  for (const id of generate) if (/^\d+$/.test(id)) taken.add(Number(id));
  let next = Math.max(...taken) + 1;

  // Ids sharing an abstract definition (`key`) share one abstractNum, so
  // their counters keep advancing together, mirroring w:abstractNumId.
  const absIdByKey = new Map<string, number>();
  const abstracts: string[] = [];
  const nums: string[] = [];
  for (const id of generate) {
    const def = defs[id];
    const key = def.key || id;
    let absId = absIdByKey.get(key);
    if (absId == null) {
      absId = next++;
      absIdByKey.set(key, absId);
      abstracts.push(abstractNumXml(absId, def));
    }
    const outId = /^\d+$/.test(id) ? id : String(next++);
    map.set(id, outId);
    nums.push(
      `<w:num w:numId="${outId}"><w:abstractNumId w:val="${absId}"/></w:num>`,
    );
  }
  return { map, abstractXml: abstracts.join(''), numXml: nums.join('') };
}

/** The generated fragments merged into a carried numbering.xml (schema order:
 *  every abstractNum precedes the first w:num), or a fresh part. */
function numberingPartXml(plan: NumberingPlan, carried: string | null): string {
  if (carried && carried.includes('</w:numbering>')) {
    let out = carried;
    const firstNum = out.search(/<w:num[ >]/);
    if (firstNum >= 0)
      out = out.slice(0, firstNum) + plan.abstractXml + out.slice(firstNum);
    else
      out = out.replace('</w:numbering>', `${plan.abstractXml}</w:numbering>`);
    return out.replace('</w:numbering>', `${plan.numXml}</w:numbering>`);
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:numbering xmlns:w="${W_NS}">${plan.abstractXml}${plan.numXml}</w:numbering>`
  );
}

const NUMBERING_OVERRIDE =
  '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>';

/**
 * Serialise a bapbong document back to `.docx` bytes.
 *
 * With `opts.carry` (the source package from `importDocx().raw`) the original
 * parts bapbong doesn't model yet — styles, numbering, headers/footers,
 * footnotes, settings, theme, media — are preserved; only document.xml and the
 * comment parts are regenerated. Without it, a minimal package is built from
 * scratch (E1–E3 content only).
 */
export async function exportDocx(
  doc: PMNode,
  opts?: { carry?: JSZip },
): Promise<Uint8Array> {
  const comments = (doc.attrs['comments'] as CommentNode[] | null) ?? [];
  // Precompute the last inline-leaf index each comment id covers, in document
  // order, so inlineContent can close the range at the right run.
  const knownComments = new Set(comments.map((c) => c.id));
  const lastRun = new Map<number, number>();
  let idx = 0;
  doc.descendants((n) => {
    if (!isInlineLeaf(n)) return;
    for (const id of commentIdsOf(n))
      if (knownComments.has(id)) lastRun.set(id, idx);
    idx++;
  });

  // Numbering must be planned before the body serialises (numPr remapping).
  const carriedNumbering =
    (await opts?.carry?.file('word/numbering.xml')?.async('string')) ?? null;
  const numbering = planNumbering(doc, carriedNumbering);

  const ctx: ExportCtx = {
    rels: [],
    media: [],
    exts: new Set(),
    nextId: 100,
    numIdMap: numbering.map,
    knownComments,
    lastRun,
    openComments: new Set(),
    runIdx: 0,
  };
  const boundaries = sectionBoundaries(doc);
  let body = '';
  perf.span('export.body', () =>
    doc.forEach(
      (block, _offset, i) => (body += blockXml(block, ctx, boundaries.get(i))),
    ),
  );

  const hasComments = comments.length > 0;
  if (hasComments) {
    ctx.rels.push(
      `<Relationship Id="rIdComments" Type="${R_NS}/comments" Target="comments.xml"/>`,
    );
    ctx.rels.push(
      `<Relationship Id="rIdCommentsExt" Type="${R_NS}/commentsExtended" Target="commentsExtended.xml"/>`,
    );
  }

  // Regenerated numbering: merged into the carried part, or a brand-new part
  // (which then needs its relationship + content-type override).
  const numberingPart = numbering.numXml
    ? numberingPartXml(numbering, carriedNumbering)
    : null;
  const newNumberingPart = numberingPart != null && carriedNumbering == null;
  if (newNumberingPart)
    ctx.rels.push(
      `<Relationship Id="rIdNumbering" Type="${R_NS}/numbering" Target="numbering.xml"/>`,
    );

  const zip = new JSZip();
  const styleIds = usedStyleIds(doc);
  let sectPr = ''; // re-attached from the original (carry) for page setup + headers
  if (opts?.carry) {
    // E4: start from the original package so unmodelled parts survive.
    const carry = opts.carry;
    for (const [path, f] of Object.entries(carry.files)) {
      if (!f.dir) zip.file(path, await f.async('uint8array'));
    }
    const ct = await carry.file('[Content_Types].xml')?.async('string');
    zip.file(
      '[Content_Types].xml',
      ct
        ? mergeContentTypes(ct, ctx.exts, hasComments, newNumberingPart)
        : contentTypes(ctx.exts, hasComments, numberingPart != null),
    );
    // Styles referenced by pStyle but never defined by the source (headings /
    // Title / Subtitle authored in bapbong) get their defs appended.
    const carriedStyles = await carry.file('word/styles.xml')?.async('string');
    if (carriedStyles)
      zip.file('word/styles.xml', mergeStyles(carriedStyles, styleIds));
    const rels = await carry
      .file('word/_rels/document.xml.rels')
      ?.async('string');
    zip.file('word/_rels/document.xml.rels', mergeRels(rels, ctx.rels));
    const origDoc = await carry.file('word/document.xml')?.async('string');
    if (origDoc) sectPr = extractBodySectPr(origDoc);
  } else {
    ctx.rels.push(
      `<Relationship Id="rIdStyles" Type="${R_NS}/styles" Target="styles.xml"/>`,
    );
    zip.file('word/styles.xml', stylesXml(styleIds));
    zip.file(
      '[Content_Types].xml',
      contentTypes(ctx.exts, hasComments, numberingPart != null),
    );
    zip.file('_rels/.rels', ROOT_RELS);
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PR_NS}">${ctx.rels.join('')}</Relationships>`,
    );
  }

  if (numberingPart) zip.file('word/numbering.xml', numberingPart);

  // Page setup: the modelled geometry (doc.attrs.page) wins over the carried
  // sectPr — but only a real edit rewrites it (byte fidelity otherwise). No
  // sectPr at all (fresh doc, or a carry without one) → emit the modelled
  // geometry outright; omitting it would hand Word ITS default (Letter), not
  // the A4 bapbong displayed.
  const pageAttr = doc.attrs['page'] as PageConfig | null;
  if (sectPr && pageAttr && !sectPrMatchesPage(sectPr, pageAttr)) {
    sectPr = splicePageGeometry(sectPr, pageAttr);
  } else if (!sectPr) {
    const p = pageAttr ?? A4_PAGE;
    sectPr = `<w:sectPr>${pgSzXml(p)}${pgMarXml(p.margin)}</w:sectPr>`;
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
  perf.span('export.media', () => {
    for (const { path, base64 } of ctx.media)
      zip.file(path, base64, { base64: true });
  });
  perf.bump('export.mediaCount', ctx.media.length);
  return perf.spanAsync('export.generate', () =>
    zip.generateAsync({ type: 'uint8array' }),
  );
}
