import JSZip from 'jszip';
import { Node as PMNode, Mark } from 'prosemirror-model';
import { schema, type ListInfo } from '@shadow-garden/bapbong-model';
import { asRecord, attr, mergeRunProps, parseRunProps, parser, RunProps, toArray } from './ooxml';
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

/** Concatenate a run's `w:t` text (each may be a string or `{ '#text' }`). */
function runText(run: Record<string, unknown>): string {
  return toArray(run['w:t'])
    .map((t) => {
      const rec = asRecord(t);
      return rec ? String(rec['#text'] ?? '') : String(t ?? '');
    })
    .join('');
}

/** Read a paragraph's list membership (w:numPr) and advance the counter. */
function parseList(
  pPr: Record<string, unknown> | undefined,
  numbering: NumberingResolver,
): ListInfo | null {
  const numPr = asRecord(pPr?.['w:numPr']);
  const numId = attr(numPr?.['w:numId'], '@_w:val');
  if (numId === undefined || numId === '0') return null; // 0 cancels numbering
  const ilvl = Number(attr(numPr?.['w:ilvl'], '@_w:val') ?? '0');
  const level = Number.isNaN(ilvl) ? 0 : ilvl;
  return { numId, level, marker: numbering.next(numId, level) };
}

function parseParagraph(
  p: Record<string, unknown>,
  styles: StyleRegistry,
  numbering: NumberingResolver,
): PMNode {
  const pPr = asRecord(p['w:pPr']);
  const pStyleId = attr(pPr?.['w:pStyle'], '@_w:val');
  // Base for every run: docDefaults → paragraph style's run properties.
  const paraBase = mergeRunProps(styles.docDefaults, styles.resolveStyle(pStyleId));
  const list = parseList(pPr, numbering);

  const inline: PMNode[] = [];
  for (const runUnknown of toArray(p['w:r'])) {
    const run = asRecord(runUnknown);
    if (!run) continue;
    const text = runText(run);
    if (text.length === 0) continue;

    const rPr = asRecord(run['w:rPr']);
    const rStyleId = attr(rPr?.['w:rStyle'], '@_w:val');
    // Cascade: docDefaults+paraStyle → run style → inline rPr (later wins).
    const effective = [paraBase, styles.resolveStyle(rStyleId), parseRunProps(rPr)].reduce(
      mergeRunProps,
      {} as RunProps,
    );

    inline.push(schema.text(text, propsToMarks(effective)));
  }
  return schema.nodes.paragraph.create(list ? { list } : null, inline);
}

async function readPart(zip: JSZip, path: string): Promise<string | undefined> {
  const entry = zip.file(path);
  return entry ? entry.async('string') : undefined;
}

function parseOptionalPart(xml: string | undefined): Record<string, unknown> | undefined {
  return xml ? (parser.parse(xml) as Record<string, unknown>) : undefined;
}

/**
 * Parse the bytes of a .docx file into a bapbong ProseMirror document.
 *
 * Scope so far: paragraphs and text runs with the run-property cascade
 * (docDefaults → paragraph/run styles → inline) resolved to bold / italic /
 * underline / strike / color / size / font marks, plus flat list paragraphs
 * with multilevel numbering markers. Tables, headers, and export are later
 * milestones; unmodeled XML is preserved on `rawDocumentXml`.
 */
export async function importDocx(input: DocxInput): Promise<DocxImport> {
  const zip = await JSZip.loadAsync(input);

  const rawDocumentXml = await readPart(zip, 'word/document.xml');
  if (rawDocumentXml === undefined) {
    throw new Error('bapbong-docx: word/document.xml not found in archive');
  }

  const styles = buildStyleRegistry(parseOptionalPart(await readPart(zip, 'word/styles.xml')));
  const numbering = buildNumbering(parseOptionalPart(await readPart(zip, 'word/numbering.xml')));

  const tree = parser.parse(rawDocumentXml) as Record<string, unknown>;
  const body = asRecord(asRecord(tree['w:document'])?.['w:body']);
  const paragraphs = toArray(body?.['w:p']).map((p) =>
    parseParagraph(asRecord(p) ?? {}, styles, numbering),
  );

  // doc content is `block+` — guarantee at least one paragraph.
  if (paragraphs.length === 0) paragraphs.push(schema.nodes.paragraph.create());

  return { doc: schema.nodes.doc.create(null, paragraphs), rawDocumentXml };
}
