import JSZip from 'jszip';
import { Node as PMNode, Mark } from 'prosemirror-model';
import { schema, type Align, type Indent, type ListInfo } from '@shadow-garden/bapbong-model';
import {
  attrOf,
  child,
  children,
  findDescendant,
  mergeRunProps,
  OoxmlNode,
  parseRunProps,
  parseXml,
  RunProps,
} from './ooxml';
import { buildStyleRegistry, StyleRegistry } from './styles';
import { buildNumbering, NumberingResolver } from './numbering';
import { buildRels, Relationship } from './rels';
import { buildThemeResolver, ThemeResolver } from './theme';

export type DocxInput = ArrayBuffer | Uint8Array | Blob;

/**
 * Result of importing a .docx: the ProseMirror document plus the raw
 * `word/document.xml` we parsed it from. The raw string is kept so a later
 * export step can round-trip parts of the document we don't model yet.
 */
export interface DocxImport {
  doc: PMNode;
  rawDocumentXml: string;
  /** Header stories keyed by w:type ("default" | "first" | "even"). */
  headers: Record<string, PMNode>;
  /** Footer stories keyed by w:type. */
  footers: Record<string, PMNode>;
}

interface Ctx {
  styles: StyleRegistry;
  numbering: NumberingResolver;
  rels: Map<string, Relationship>;
  media: Map<string, string>; // zip path → data URL
  resolveTheme: ThemeResolver;
}

/** 1440 twips = 1 inch = 96 px. */
const twipsToPx = (twips: number) => Math.round(twips / 15);
/** 914400 EMU = 1 inch = 96 px → 9525 EMU/px. */
function emuToPx(emu: string | undefined): number | null {
  const n = Number(emu ?? '0');
  return Number.isNaN(n) || n === 0 ? null : Math.round(n / 9525);
}

function mimeOf(path: string): string {
  switch (path.slice(path.lastIndexOf('.') + 1).toLowerCase()) {
    case 'png':
      return 'image/png';
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'bmp':
      return 'image/bmp';
    case 'svg':
      return 'image/svg+xml';
    case 'webp':
      return 'image/webp';
    case 'tif':
    case 'tiff':
      return 'image/tiff';
    default:
      return 'application/octet-stream';
  }
}

/** Turn resolved run properties into ProseMirror marks. */
function propsToMarks(p: RunProps): Mark[] {
  const marks: Mark[] = [];
  if (p.bold) marks.push(schema.marks.strong.create());
  if (p.italic) marks.push(schema.marks.em.create());
  if (p.underline) marks.push(schema.marks.underline.create());
  if (p.strike) marks.push(schema.marks.strike.create());
  if (p.color) marks.push(schema.marks.textColor.create({ color: p.color }));
  if (p.sizePt !== undefined) marks.push(schema.marks.fontSize.create({ size: p.sizePt }));
  if (p.fontFamily) marks.push(schema.marks.fontFamily.create({ family: p.fontFamily }));
  return marks;
}

/** A run's text content in order: w:t segments plus w:tab elements (→ \t). */
function runText(run: OoxmlNode): string {
  let out = '';
  for (const node of run.children) {
    if (node.name === 'w:t') out += node.text;
    else if (node.name === 'w:tab') out += '\t';
  }
  return out;
}

/** EMU → px where 0 is meaningful (offsets/gaps), unlike emuToPx. */
function emuToPxZero(emu: string | undefined): number | undefined {
  if (emu === undefined) return undefined;
  const n = Number(emu);
  return Number.isNaN(n) ? undefined : Math.round(n / 9525);
}

/** Floating-image placement from a wp:anchor, or null for wp:inline. */
function parseAnchorFloat(drawing: OoxmlNode): Record<string, unknown> | null {
  const anchor = child(drawing, 'wp:anchor');
  if (!anchor) return null;

  const wrap = child(anchor, 'wp:wrapTopAndBottom')
    ? 'topAndBottom'
    : child(anchor, 'wp:wrapSquare') || child(anchor, 'wp:wrapTight') || child(anchor, 'wp:wrapThrough')
      ? 'square'
      : 'none'; // wrapNone / absent: paints without affecting text

  const float: Record<string, unknown> = { wrap };

  const posH = child(anchor, 'wp:positionH');
  if (posH) {
    const align = child(posH, 'wp:align')?.text.trim();
    if (align === 'left' || align === 'right' || align === 'center') float['hAlign'] = align;
    const off = emuToPxZero(child(posH, 'wp:posOffset')?.text);
    if (off !== undefined && float['hAlign'] === undefined) float['hOffset'] = off;
    const rel = attrOf(posH, 'relativeFrom');
    float['hRel'] = rel === 'page' ? 'page' : 'margin'; // column/margin/… ≈ margin
  }
  const posV = child(anchor, 'wp:positionV');
  if (posV) {
    const off = emuToPxZero(child(posV, 'wp:posOffset')?.text);
    if (off !== undefined) float['vOffset'] = off;
    const rel = attrOf(posV, 'relativeFrom');
    float['vRel'] = rel === 'page' ? 'page' : rel === 'margin' ? 'margin' : 'paragraph';
  }
  // Text-to-image gaps (EMU attrs on the anchor itself).
  for (const side of ['distL', 'distR', 'distT', 'distB'] as const) {
    const v = emuToPxZero(attrOf(anchor, side));
    if (v !== undefined) float[side] = v;
  }
  return float;
}

/** Extract an image (inline or floating) from a run's w:drawing, if any. */
function parseImage(run: OoxmlNode, ctx: Ctx): PMNode | null {
  const drawing = child(run, 'w:drawing');
  if (!drawing) return null;
  const blip = findDescendant(drawing, 'a:blip');
  const embed = attrOf(blip, 'r:embed') ?? attrOf(blip, 'r:link');
  const rel = embed ? ctx.rels.get(embed) : undefined;
  if (!rel) return null;
  const target = rel.target.replace(/^\/+/, '');
  const src = ctx.media.get(`word/${target}`) ?? ctx.media.get(target);
  if (!src) return null;

  const extent = findDescendant(drawing, 'wp:extent');
  const docPr = findDescendant(drawing, 'wp:docPr');
  const float = parseAnchorFloat(drawing);
  return schema.nodes.image.create({
    src,
    width: emuToPx(attrOf(extent, 'cx')),
    height: emuToPx(attrOf(extent, 'cy')),
    alt: attrOf(docPr, 'descr') ?? attrOf(docPr, 'title') ?? '',
    float,
  });
}

/** Effective marks for a run (docDefaults+paraStyle → run style → inline rPr). */
function runMarks(run: OoxmlNode | undefined, paraBase: RunProps, ctx: Ctx, href: string | null) {
  const rPr = child(run, 'w:rPr');
  const rStyleId = attrOf(child(rPr, 'w:rStyle'), 'w:val');
  const effective = [
    paraBase,
    ctx.styles.resolveStyle(rStyleId),
    parseRunProps(rPr, ctx.resolveTheme),
  ].reduce(mergeRunProps, {} as RunProps);
  const marks = propsToMarks(effective);
  if (href) marks.push(schema.marks.link.create({ href }));
  return marks;
}

/** Map one w:r into inline nodes (image or marked text), optionally hyperlinked. */
function runToInline(run: OoxmlNode, paraBase: RunProps, ctx: Ctx, href: string | null): PMNode[] {
  const marks = runMarks(run, paraBase, ctx, href);

  const image = parseImage(run, ctx);
  if (image) return [href ? image.mark([schema.marks.link.create({ href })]) : image];

  const text = runText(run);
  if (text.length === 0) return [];
  return [schema.text(text, marks)];
}

/** PAGE / NUMPAGES from a field instruction, or null for any other field. */
function fieldKind(instr: string): 'page' | 'pages' | null {
  if (/\bNUMPAGES\b/.test(instr)) return 'pages';
  if (/\bPAGE\b/.test(instr)) return 'page';
  return null;
}

/** A page_field node formatted like `formatRun` (the field's result run). */
function pageFieldNode(
  kind: 'page' | 'pages',
  formatRun: OoxmlNode | undefined,
  paraBase: RunProps,
  ctx: Ctx,
): PMNode {
  return schema.nodes.page_field.create({ kind }).mark(runMarks(formatRun, paraBase, ctx, null));
}

/** State for a complex field (w:fldChar begin … instrText … separate … end). */
interface FieldState {
  instr: string;
  resultRuns: OoxmlNode[];
  phase: 'instr' | 'result';
}

function parseParagraph(p: OoxmlNode, ctx: Ctx): PMNode {
  const pPr = child(p, 'w:pPr');
  const pStyleId = attrOf(child(pPr, 'w:pStyle'), 'w:val');
  // Base for every run: docDefaults → paragraph style's run properties.
  const paraBase = mergeRunProps(ctx.styles.docDefaults, ctx.styles.resolveStyle(pStyleId));
  // Paragraph-property cascade, base-most first; later layers win:
  // docDefaults pPrDefault → style chain (w:basedOn ancestors → style) → inline.
  const pPrChain: (OoxmlNode | undefined)[] = [
    ctx.styles.docDefaultsPPr,
    ...ctx.styles.resolveStylePPr(pStyleId),
    pPr,
  ];
  const list = parseList(lastWith(pPrChain, 'w:numPr'), ctx.numbering);
  const align = resolveAlign(pPrChain);
  const indent = resolveIndent(pPrChain);
  const tabs = resolveTabs(pPrChain);

  const inline: PMNode[] = [];
  let field: FieldState | null = null;
  for (const node of p.children) {
    if (node.name === 'w:r') {
      const fldType = attrOf(child(node, 'w:fldChar'), 'w:fldCharType');
      if (fldType === 'begin') {
        field = { instr: '', resultRuns: [], phase: 'instr' };
        continue;
      }
      if (field) {
        if (fldType === 'separate') {
          field.phase = 'result';
        } else if (fldType === 'end') {
          const kind = fieldKind(field.instr);
          if (kind) {
            inline.push(pageFieldNode(kind, field.resultRuns[0], paraBase, ctx));
          } else {
            // Unknown instruction: keep the cached result text.
            for (const r of field.resultRuns) inline.push(...runToInline(r, paraBase, ctx, null));
          }
          field = null;
        } else if (field.phase === 'instr') {
          field.instr += children(node, 'w:instrText')
            .map((t) => t.text)
            .join('');
        } else {
          field.resultRuns.push(node);
        }
        continue;
      }
      inline.push(...runToInline(node, paraBase, ctx, null));
    } else if (node.name === 'w:fldSimple') {
      const kind = fieldKind(attrOf(node, 'w:instr') ?? '');
      const resultRuns = children(node, 'w:r');
      if (kind) {
        inline.push(pageFieldNode(kind, resultRuns[0], paraBase, ctx));
      } else {
        for (const r of resultRuns) inline.push(...runToInline(r, paraBase, ctx, null));
      }
    } else if (node.name === 'w:hyperlink') {
      const rel = attrOf(node, 'r:id') ? ctx.rels.get(attrOf(node, 'r:id') as string) : undefined;
      const anchor = attrOf(node, 'w:anchor');
      const href = rel?.target ?? (anchor ? `#${anchor}` : null);
      for (const run of children(node, 'w:r')) {
        inline.push(...runToInline(run, paraBase, ctx, href));
      }
    }
  }
  const attrs: {
    list?: ListInfo;
    align?: Align;
    indent?: Indent;
    tabs?: { pos: number; val: string; leader?: string }[];
  } = {};
  if (list) attrs.list = list;
  if (align) attrs.align = align;
  if (indent) attrs.indent = indent;
  if (tabs) attrs.tabs = tabs;
  return schema.nodes.paragraph.create(attrs, inline);
}

/** Custom tab stops from the cascade: the most-derived w:tabs list wins
 *  (per-layer merging with w:val="clear" is a later refinement). 'clear' and
 *  unsupported 'bar' stops are dropped; 'num' behaves like 'left'. */
function resolveTabs(
  chain: (OoxmlNode | undefined)[],
): { pos: number; val: string; leader?: string }[] | null {
  const layer = lastWith(chain, 'w:tabs');
  if (!layer) return null;
  const stops: { pos: number; val: string; leader?: string }[] = [];
  for (const tab of children(child(layer, 'w:tabs'), 'w:tab')) {
    const val = attrOf(tab, 'w:val') ?? 'left';
    if (val === 'clear' || val === 'bar') continue;
    const pos = attrOf(tab, 'w:pos');
    if (pos === undefined) continue;
    const stop: { pos: number; val: string; leader?: string } = {
      pos: twipsToPx(Number(pos)),
      val: val === 'right' || val === 'center' || val === 'decimal' ? val : 'left',
    };
    const leader = attrOf(tab, 'w:leader');
    if (leader && leader !== 'none') {
      stop.leader =
        leader === 'hyphen' ? 'hyphen' : leader === 'underscore' ? 'underscore' : leader === 'middleDot' ? 'middleDot' : 'dot';
    }
    stops.push(stop);
  }
  return stops.length > 0 ? stops.sort((a, b) => a.pos - b.pos) : null;
}

/** The last (most-derived) pPr layer that carries `childName`, if any. */
function lastWith(chain: (OoxmlNode | undefined)[], childName: string): OoxmlNode | undefined {
  for (let i = chain.length - 1; i >= 0; i--) {
    if (child(chain[i], childName)) return chain[i];
  }
  return undefined;
}

/** Resolve alignment through the cascade: the last layer with a w:jc wins. */
function resolveAlign(chain: (OoxmlNode | undefined)[]): Align | null {
  for (let i = chain.length - 1; i >= 0; i--) {
    const align = parseAlign(chain[i]);
    if (align) return align;
  }
  return null;
}

/** Resolve indentation through the cascade. w:ind merges per attribute (a
 *  style's left can combine with an inline firstLine); when both hanging and
 *  firstLine survive the merge, hanging wins. */
function resolveIndent(chain: (OoxmlNode | undefined)[]): Indent | null {
  let left: number | undefined;
  let right: number | undefined;
  let firstLine: number | undefined;
  let hanging: number | undefined;
  for (const pPr of chain) {
    const ind = child(pPr, 'w:ind');
    if (!ind) continue;
    const px = (attr: string): number | undefined => {
      const v = attrOf(ind, attr);
      return v === undefined ? undefined : twipsToPx(Number(v));
    };
    left = px('w:left') ?? px('w:start') ?? left;
    right = px('w:right') ?? px('w:end') ?? right;
    const fl = px('w:firstLine');
    const hg = px('w:hanging');
    // A layer that sets either first-line property replaces the pair.
    if (fl !== undefined || hg !== undefined) {
      firstLine = fl;
      hanging = hg;
    }
  }
  const out: Indent = {};
  if (left !== undefined) out.left = left;
  if (right !== undefined) out.right = right;
  if (hanging !== undefined) out.hanging = hanging;
  else if (firstLine !== undefined) out.firstLine = firstLine;
  return Object.keys(out).length > 0 ? out : null;
}

/** Map w:jc to an alignment, or null when absent/unrecognized. */
function parseAlign(pPr: OoxmlNode | undefined): Align | null {
  switch (attrOf(child(pPr, 'w:jc'), 'w:val')) {
    case 'center':
      return 'center';
    case 'right':
    case 'end':
      return 'right';
    case 'both':
    case 'distribute':
      return 'justify';
    case 'left':
    case 'start':
      return 'left';
    default:
      return null;
  }
}

/** Read a paragraph's list membership (w:numPr) and advance the counter. */
function parseList(pPr: OoxmlNode | undefined, numbering: NumberingResolver): ListInfo | null {
  const numPr = child(pPr, 'w:numPr');
  const numId = attrOf(child(numPr, 'w:numId'), 'w:val');
  if (numId === undefined || numId === '0') return null; // 0 cancels numbering
  const ilvl = Number(attrOf(child(numPr, 'w:ilvl'), 'w:val') ?? '0');
  const level = Number.isNaN(ilvl) ? 0 : ilvl;
  return { numId, level, marker: numbering.next(numId, level) };
}

interface LogicalCell {
  startCol: number; // grid column this cell starts at
  colspan: number;
  vMerge: 'restart' | 'continue' | null;
  colwidth: number[] | null;
  content: PMNode[];
}

function emptyCell(): PMNode {
  return schema.nodes.table_cell.create(null, [schema.nodes.paragraph.create()]);
}

/** w:tblPr/w:tblCellMar overrides (px), or null for Word defaults.
 *  type="nil" forces 0; only "dxa" widths are interpreted. */
function parseCellMargins(
  tbl: OoxmlNode,
): { left?: number; right?: number; top?: number; bottom?: number } | null {
  const mar = child(child(tbl, 'w:tblPr'), 'w:tblCellMar');
  if (!mar) return null;
  const side = (name: string): number | undefined => {
    const el = child(mar, name);
    if (!el) return undefined;
    if (attrOf(el, 'w:type') === 'nil') return 0;
    const w = attrOf(el, 'w:w');
    return w === undefined ? undefined : twipsToPx(Number(w));
  };
  const out: { left?: number; right?: number; top?: number; bottom?: number } = {};
  const left = side('w:left') ?? side('w:start');
  const right = side('w:right') ?? side('w:end');
  const top = side('w:top');
  const bottom = side('w:bottom');
  if (left !== undefined) out.left = left;
  if (right !== undefined) out.right = right;
  if (top !== undefined) out.top = top;
  if (bottom !== undefined) out.bottom = bottom;
  return Object.keys(out).length > 0 ? out : null;
}

function parseTable(tbl: OoxmlNode, ctx: Ctx): PMNode {
  const grid = children(child(tbl, 'w:tblGrid'), 'w:gridCol').map((c) =>
    Number(attrOf(c, 'w:w') ?? '0'),
  );

  // Phase 1: logical grid — every w:tc (incl. vMerge-continue placeholders),
  // tracking each cell's starting grid column.
  const logicalRows: LogicalCell[][] = children(tbl, 'w:tr').map((tr) => {
    const cells: LogicalCell[] = [];
    let col = 0;
    for (const tc of children(tr, 'w:tc')) {
      const tcPr = child(tc, 'w:tcPr');
      const colspan = Number(attrOf(child(tcPr, 'w:gridSpan'), 'w:val') ?? '1') || 1;
      const vMergeEl = child(tcPr, 'w:vMerge');
      const vMerge = !vMergeEl
        ? null
        : attrOf(vMergeEl, 'w:val') === 'restart'
          ? 'restart'
          : 'continue'; // omitted w:val defaults to continue
      const widths = grid.length ? grid.slice(col, col + colspan).map(twipsToPx) : [];
      const content = parseBlocks(tc, ctx);
      if (content.length === 0) content.push(schema.nodes.paragraph.create());
      cells.push({ startCol: col, colspan, vMerge, colwidth: widths.length ? widths : null, content });
      col += colspan;
    }
    return cells;
  });

  // w:trPr/w:tblHeader — OOXML on/off: present means true unless val says no.
  const headerFlags = children(tbl, 'w:tr').map((tr) => {
    const el = child(child(tr, 'w:trPr'), 'w:tblHeader');
    if (!el) return false;
    const val = attrOf(el, 'w:val');
    return val !== 'false' && val !== '0';
  });

  const colIndex = logicalRows.map((cells) => new Map(cells.map((c) => [c.startCol, c])));

  // Phase 2: drop continue cells; non-continue cells absorb the continues
  // directly below them in the same column as rowspan.
  const rows: PMNode[] = logicalRows.map((cells, r) => {
    const emitted: PMNode[] = [];
    for (const cell of cells) {
      if (cell.vMerge === 'continue') continue; // absorbed by the cell above
      let rowspan = 1;
      for (let r2 = r + 1; r2 < logicalRows.length; r2++) {
        const below = colIndex[r2].get(cell.startCol);
        if (below && below.vMerge === 'continue') rowspan++;
        else break;
      }
      emitted.push(
        schema.nodes.table_cell.create(
          { colspan: cell.colspan, rowspan, colwidth: cell.colwidth },
          cell.content,
        ),
      );
    }
    return schema.nodes.table_row.create(
      headerFlags[r] ? { header: true } : null,
      emitted.length > 0 ? emitted : [emptyCell()],
    );
  });

  const cellPadding = parseCellMargins(tbl);
  return schema.nodes.table.create(
    cellPadding ? { cellPadding } : null,
    rows.length > 0 ? rows : [schema.nodes.table_row.create(null, [emptyCell()])],
  );
}

/** Walk an element's children in document order, mapping w:p / w:tbl to blocks. */
function parseBlocks(parent: OoxmlNode, ctx: Ctx): PMNode[] {
  const blocks: PMNode[] = [];
  for (const node of parent.children) {
    if (node.name === 'w:p') blocks.push(parseParagraph(node, ctx));
    else if (node.name === 'w:tbl') blocks.push(parseTable(node, ctx));
  }
  return blocks;
}

async function readPart(zip: JSZip, path: string): Promise<string | undefined> {
  const entry = zip.file(path);
  return entry ? entry.async('string') : undefined;
}

/** Rels file that accompanies a part: "word/header1.xml" → "word/_rels/header1.xml.rels". */
async function readPartRels(zip: JSZip, partPath: string): Promise<OoxmlNode | undefined> {
  const slash = partPath.lastIndexOf('/');
  const relsPath = `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
  const xml = await readPart(zip, relsPath);
  return xml ? parseXml(xml) : undefined;
}

function storyDoc(blocks: PMNode[]): PMNode {
  // doc content is `block+` — guarantee at least one paragraph.
  return schema.nodes.doc.create(null, blocks.length > 0 ? blocks : [schema.nodes.paragraph.create()]);
}

async function extractMedia(zip: JSZip): Promise<Map<string, string>> {
  const media = new Map<string, string>();
  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith('word/media/')) continue;
    const entry = zip.file(path);
    if (!entry || entry.dir) continue;
    media.set(path, `data:${mimeOf(path)};base64,${await entry.async('base64')}`);
  }
  return media;
}

/**
 * Parse the bytes of a .docx file into a bapbong ProseMirror document.
 *
 * Scope so far: paragraphs, tables (col/row spans), text runs with the
 * run-property cascade resolved to bold/italic/underline/strike/color/size/font
 * marks, flat list paragraphs with multilevel numbering, hyperlinks (link mark),
 * and inline images (data-URL). Headers/footers, theme colors, and export are
 * later milestones; unmodeled XML is preserved on `rawDocumentXml`.
 */
export async function importDocx(input: DocxInput): Promise<DocxImport> {
  const zip = await JSZip.loadAsync(input);

  const rawDocumentXml = await readPart(zip, 'word/document.xml');
  if (rawDocumentXml === undefined) {
    throw new Error('bapbong-docx: word/document.xml not found in archive');
  }

  const stylesXml = await readPart(zip, 'word/styles.xml');
  const numberingXml = await readPart(zip, 'word/numbering.xml');
  const themeXml = await readPart(zip, 'word/theme/theme1.xml');

  // Stateless/shared pieces; numbering counters are per-story (built fresh below).
  const resolveTheme = buildThemeResolver(themeXml ? parseXml(themeXml) : undefined);
  const styles = buildStyleRegistry(stylesXml ? parseXml(stylesXml) : undefined, resolveTheme);
  const numberingRoot = numberingXml ? parseXml(numberingXml) : undefined;
  const media = await extractMedia(zip);
  const makeCtx = (rels: Map<string, Relationship>): Ctx => ({
    styles,
    numbering: buildNumbering(numberingRoot),
    rels,
    media,
    resolveTheme,
  });

  const docRels = await readPart(zip, 'word/_rels/document.xml.rels');
  const ctx = makeCtx(buildRels(docRels ? parseXml(docRels) : undefined));

  const body = child(child(parseXml(rawDocumentXml), 'w:document'), 'w:body');
  const doc = storyDoc(body ? parseBlocks(body, ctx) : []);

  // Headers/footers referenced by the section properties.
  const headers: Record<string, PMNode> = {};
  const footers: Record<string, PMNode> = {};
  const sectPr = body ? child(body, 'w:sectPr') : undefined;
  if (sectPr) {
    const collect = async (refName: string, store: Record<string, PMNode>, root: string) => {
      for (const ref of children(sectPr, refName)) {
        const type = attrOf(ref, 'w:type') ?? 'default';
        const rId = attrOf(ref, 'r:id');
        const target = rId ? ctx.rels.get(rId)?.target : undefined;
        if (!target || store[type]) continue;
        const partPath = `word/${target.replace(/^\/+/, '')}`;
        const xml = await readPart(zip, partPath);
        if (!xml) continue;
        const partCtx = makeCtx(buildRels(await readPartRels(zip, partPath)));
        const el = child(parseXml(xml), root);
        store[type] = storyDoc(el ? parseBlocks(el, partCtx) : []);
      }
    };
    await collect('w:headerReference', headers, 'w:hdr');
    await collect('w:footerReference', footers, 'w:ftr');
  }

  return { doc, rawDocumentXml, headers, footers };
}
