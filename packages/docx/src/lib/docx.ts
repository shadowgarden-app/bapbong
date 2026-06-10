import JSZip from 'jszip';
import { Node as PMNode, Mark } from 'prosemirror-model';
import { schema, type ListInfo } from '@shadow-garden/bapbong-model';
import {
  attrOf,
  child,
  children,
  mergeRunProps,
  OoxmlNode,
  parseRunProps,
  parseXml,
  RunProps,
} from './ooxml';
import { buildStyleRegistry, StyleRegistry } from './styles';
import { buildNumbering, NumberingResolver } from './numbering';

export type DocxInput = ArrayBuffer | Uint8Array | Blob;

/**
 * Result of importing a .docx: the ProseMirror document plus the raw
 * `word/document.xml` we parsed it from. The raw string is kept so a later
 * export step can round-trip parts of the document we don't model yet.
 */
export interface DocxImport {
  doc: PMNode;
  rawDocumentXml: string;
}

interface Ctx {
  styles: StyleRegistry;
  numbering: NumberingResolver;
}

/** 1440 twips = 1 inch = 96 px. */
const twipsToPx = (twips: number) => Math.round(twips / 15);

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

/** Read a paragraph's list membership (w:numPr) and advance the counter. */
function parseList(pPr: OoxmlNode | undefined, numbering: NumberingResolver): ListInfo | null {
  const numPr = child(pPr, 'w:numPr');
  const numId = attrOf(child(numPr, 'w:numId'), 'w:val');
  if (numId === undefined || numId === '0') return null; // 0 cancels numbering
  const ilvl = Number(attrOf(child(numPr, 'w:ilvl'), 'w:val') ?? '0');
  const level = Number.isNaN(ilvl) ? 0 : ilvl;
  return { numId, level, marker: numbering.next(numId, level) };
}

function parseParagraph(p: OoxmlNode, ctx: Ctx): PMNode {
  const pPr = child(p, 'w:pPr');
  const pStyleId = attrOf(child(pPr, 'w:pStyle'), 'w:val');
  // Base for every run: docDefaults → paragraph style's run properties.
  const paraBase = mergeRunProps(ctx.styles.docDefaults, ctx.styles.resolveStyle(pStyleId));
  const list = parseList(pPr, ctx.numbering);

  const inline: PMNode[] = [];
  for (const run of children(p, 'w:r')) {
    const text = runText(run);
    if (text.length === 0) continue;

    const rPr = child(run, 'w:rPr');
    const rStyleId = attrOf(child(rPr, 'w:rStyle'), 'w:val');
    // Cascade: docDefaults+paraStyle → run style → inline rPr (later wins).
    const effective = [paraBase, ctx.styles.resolveStyle(rStyleId), parseRunProps(rPr)].reduce(
      mergeRunProps,
      {} as RunProps,
    );

    inline.push(schema.text(text, propsToMarks(effective)));
  }
  return schema.nodes.paragraph.create(list ? { list } : null, inline);
}

function parseTable(tbl: OoxmlNode, ctx: Ctx): PMNode {
  const grid = children(child(tbl, 'w:tblGrid'), 'w:gridCol').map((c) =>
    Number(attrOf(c, 'w:w') ?? '0'),
  );

  const rows: PMNode[] = [];
  for (const tr of children(tbl, 'w:tr')) {
    const cells: PMNode[] = [];
    let col = 0;
    for (const tc of children(tr, 'w:tc')) {
      const tcPr = child(tc, 'w:tcPr');
      const colspan = Number(attrOf(child(tcPr, 'w:gridSpan'), 'w:val') ?? '1') || 1;
      const widths = grid.length ? grid.slice(col, col + colspan).map(twipsToPx) : [];
      col += colspan;

      const content = parseBlocks(tc, ctx);
      if (content.length === 0) content.push(schema.nodes.paragraph.create());
      cells.push(
        schema.nodes.table_cell.create(
          { colspan, rowspan: 1, colwidth: widths.length ? widths : null },
          content,
        ),
      );
    }
    if (cells.length > 0) rows.push(schema.nodes.table_row.create(null, cells));
  }

  if (rows.length === 0) {
    rows.push(
      schema.nodes.table_row.create(null, [
        schema.nodes.table_cell.create(null, [schema.nodes.paragraph.create()]),
      ]),
    );
  }
  return schema.nodes.table.create(null, rows);
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

/**
 * Parse the bytes of a .docx file into a bapbong ProseMirror document.
 *
 * Scope so far: paragraphs, tables, and text runs with the run-property cascade
 * (docDefaults → paragraph/run styles → inline) resolved to bold / italic /
 * underline / strike / color / size / font marks, plus flat list paragraphs
 * with multilevel numbering markers. Vertical cell merges, headers/footers, and
 * export are later milestones; unmodeled XML is preserved on `rawDocumentXml`.
 */
export async function importDocx(input: DocxInput): Promise<DocxImport> {
  const zip = await JSZip.loadAsync(input);

  const rawDocumentXml = await readPart(zip, 'word/document.xml');
  if (rawDocumentXml === undefined) {
    throw new Error('bapbong-docx: word/document.xml not found in archive');
  }

  const stylesXml = await readPart(zip, 'word/styles.xml');
  const numberingXml = await readPart(zip, 'word/numbering.xml');
  const ctx: Ctx = {
    styles: buildStyleRegistry(stylesXml ? parseXml(stylesXml) : undefined),
    numbering: buildNumbering(numberingXml ? parseXml(numberingXml) : undefined),
  };

  const body = child(child(parseXml(rawDocumentXml), 'w:document'), 'w:body');
  const blocks = body ? parseBlocks(body, ctx) : [];

  // doc content is `block+` — guarantee at least one paragraph.
  if (blocks.length === 0) blocks.push(schema.nodes.paragraph.create());

  return { doc: schema.nodes.doc.create(null, blocks), rawDocumentXml };
}
