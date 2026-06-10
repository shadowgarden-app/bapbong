import JSZip from 'jszip';
import { Node as PMNode, Mark } from 'prosemirror-model';
import { schema, type ListInfo } from '@shadow-garden/bapbong-model';
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

/** Concatenate a run's `w:t` text segments. */
function runText(run: OoxmlNode): string {
  return children(run, 'w:t')
    .map((t) => t.text)
    .join('');
}

/** Extract an inline image from a run's w:drawing, if any. */
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
  return schema.nodes.image.create({
    src,
    width: emuToPx(attrOf(extent, 'cx')),
    height: emuToPx(attrOf(extent, 'cy')),
    alt: attrOf(docPr, 'descr') ?? attrOf(docPr, 'title') ?? '',
  });
}

/** Map one w:r into inline nodes (image or marked text), optionally hyperlinked. */
function runToInline(run: OoxmlNode, paraBase: RunProps, ctx: Ctx, href: string | null): PMNode[] {
  const rPr = child(run, 'w:rPr');
  const rStyleId = attrOf(child(rPr, 'w:rStyle'), 'w:val');
  // Cascade: docDefaults+paraStyle → run style → inline rPr (later wins).
  const effective = [
    paraBase,
    ctx.styles.resolveStyle(rStyleId),
    parseRunProps(rPr, ctx.resolveTheme),
  ].reduce(mergeRunProps, {} as RunProps);
  const marks = propsToMarks(effective);
  if (href) marks.push(schema.marks.link.create({ href }));

  const image = parseImage(run, ctx);
  if (image) return [href ? image.mark([schema.marks.link.create({ href })]) : image];

  const text = runText(run);
  if (text.length === 0) return [];
  return [schema.text(text, marks)];
}

function parseParagraph(p: OoxmlNode, ctx: Ctx): PMNode {
  const pPr = child(p, 'w:pPr');
  const pStyleId = attrOf(child(pPr, 'w:pStyle'), 'w:val');
  // Base for every run: docDefaults → paragraph style's run properties.
  const paraBase = mergeRunProps(ctx.styles.docDefaults, ctx.styles.resolveStyle(pStyleId));
  const list = parseList(pPr, ctx.numbering);

  const inline: PMNode[] = [];
  for (const node of p.children) {
    if (node.name === 'w:r') {
      inline.push(...runToInline(node, paraBase, ctx, null));
    } else if (node.name === 'w:hyperlink') {
      const rel = attrOf(node, 'r:id') ? ctx.rels.get(attrOf(node, 'r:id') as string) : undefined;
      const anchor = attrOf(node, 'w:anchor');
      const href = rel?.target ?? (anchor ? `#${anchor}` : null);
      for (const run of children(node, 'w:r')) {
        inline.push(...runToInline(run, paraBase, ctx, href));
      }
    }
  }
  return schema.nodes.paragraph.create(list ? { list } : null, inline);
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
    return schema.nodes.table_row.create(null, emitted.length > 0 ? emitted : [emptyCell()]);
  });

  return schema.nodes.table.create(
    null,
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
