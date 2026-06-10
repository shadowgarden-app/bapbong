import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { Node as PMNode, Mark } from 'prosemirror-model';
import { schema } from '@shadow-garden/bapbong-model';

export type DocxInput = ArrayBuffer | Uint8Array | Blob;

/**
 * Result of importing a .docx: the ProseMirror document plus the raw
 * `word/document.xml` we parsed it from. The raw string is kept so a later
 * export step can round-trip parts of the document we don't model yet
 * (M1 only maps paragraphs + the four character toggles).
 */
export interface DocxImport {
  doc: PMNode;
  rawDocumentXml: string;
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  // Keep significant whitespace in <w:t> (e.g. xml:space="preserve" "Hello ").
  trimValues: false,
  // Repeated OOXML elements must always be arrays so the walker is uniform.
  isArray: (name) => name === 'w:p' || name === 'w:r' || name === 'w:t',
});

function toArray<T>(v: T | T[] | undefined | null): T[] {
  if (v == null) return [];
  return Array.isArray(v) ? v : [v];
}

function asRecord(v: unknown): Record<string, unknown> | undefined {
  return v && typeof v === 'object' ? (v as Record<string, unknown>) : undefined;
}

/** A property element's `w:val` attribute, if present. */
function val(el: unknown): string | undefined {
  const rec = asRecord(el);
  const v = rec?.['@_w:val'];
  return v == null ? undefined : String(v);
}

const OFF = new Set(['false', '0', 'off']);

/** OOXML on/off toggle (`<w:b/>`, `<w:i/>`, `<w:strike/>`): present ⇒ on,
 *  unless an explicit falsy `w:val` turns it off. */
function toggleOn(el: unknown): boolean {
  if (el === undefined) return false;
  const v = val(el);
  return v === undefined || !OFF.has(v.toLowerCase());
}

/** Underline (`<w:u w:val="single|none|...">`): on unless `none`/falsy. */
function underlineOn(el: unknown): boolean {
  if (el === undefined) return false;
  const v = val(el);
  if (v === undefined) return true;
  const lower = v.toLowerCase();
  return lower !== 'none' && !OFF.has(lower);
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

function runMarks(run: Record<string, unknown>): Mark[] {
  const rPr = asRecord(run['w:rPr']);
  if (!rPr) return [];
  const marks: Mark[] = [];
  if (toggleOn(rPr['w:b'])) marks.push(schema.marks.strong.create());
  if (toggleOn(rPr['w:i'])) marks.push(schema.marks.em.create());
  if (underlineOn(rPr['w:u'])) marks.push(schema.marks.underline.create());
  if (toggleOn(rPr['w:strike'])) marks.push(schema.marks.strike.create());
  return marks;
}

function parseParagraph(p: Record<string, unknown>): PMNode {
  const inline: PMNode[] = [];
  for (const runUnknown of toArray(p['w:r'])) {
    const run = asRecord(runUnknown);
    if (!run) continue;
    const text = runText(run);
    if (text.length === 0) continue;
    inline.push(schema.text(text, runMarks(run)));
  }
  return schema.nodes.paragraph.create(null, inline);
}

async function readDocumentXml(input: DocxInput): Promise<string> {
  const zip = await JSZip.loadAsync(input);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('bapbong-docx: word/document.xml not found in archive');
  return entry.async('string');
}

/**
 * Parse the bytes of a .docx file into a bapbong ProseMirror document.
 *
 * M1 scope: paragraphs and text runs with bold/italic/underline/strike.
 * Anything else in the body is ignored for now (the raw XML is preserved on
 * the result for future round-tripping).
 */
export async function importDocx(input: DocxInput): Promise<DocxImport> {
  const rawDocumentXml = await readDocumentXml(input);
  const tree = parser.parse(rawDocumentXml) as Record<string, unknown>;

  const body = asRecord(asRecord(tree['w:document'])?.['w:body']);
  const paragraphs = toArray(body?.['w:p']).map((p) => parseParagraph(asRecord(p) ?? {}));

  // doc content is `block+` — guarantee at least one paragraph.
  if (paragraphs.length === 0) paragraphs.push(schema.nodes.paragraph.create());

  return { doc: schema.nodes.doc.create(null, paragraphs), rawDocumentXml };
}
