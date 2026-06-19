import JSZip from 'jszip';
import type { Mark, Node as PMNode } from 'prosemirror-model';

/**
 * DOCX export (round-trip), phase E1: serialise the bapbong ProseMirror document
 * back to a minimal but valid `.docx` — paragraphs, runs and the common
 * character marks. Lists/tables/images (E2) and comments (E3) follow; carrying
 * the original imported parts (styles/numbering/media) is E4.
 */

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OR_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** px → twips (1px @96dpi = 15 twips); inverse of the importer's twipsToPx. */
const pxToTwips = (px: number) => Math.round(px * 15);

function esc(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string);
}

/** A run's w:rPr from its marks (only the character marks; link/footnote/comment
 *  are handled at the inline/range level, added in later phases). */
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
  if (size != null) out.push(`<w:sz w:val="${Math.round(size * 2)}"/>`); // half-points
  const hl = byName.get('highlight')?.attrs['color'] as string | undefined;
  if (hl) out.push(`<w:shd w:val="clear" w:color="auto" w:fill="${hl.replace(/^#/, '')}"/>`);
  const va = byName.get('vertAlign')?.attrs['value'] as string | undefined;
  if (va) out.push(`<w:vertAlign w:val="${va === 'sub' ? 'subscript' : 'superscript'}"/>`);
  return out.length ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

/** One inline node → run XML (text run, or a w:br for a hard break). */
function inlineXml(node: PMNode): string {
  if (node.type.name === 'hard_break') return '<w:r><w:br/></w:r>';
  if (node.isText) {
    const rpr = runProps(node.marks);
    return `<w:r>${rpr}<w:t xml:space="preserve">${esc(node.text ?? '')}</w:t></w:r>`;
  }
  return ''; // image / page_field / fields — later phases
}

/** A paragraph's w:pPr from its attrs (OOXML child order: pageBreakBefore,
 *  spacing, ind, jc). */
function paraProps(node: PMNode): string {
  const a = node.attrs;
  const out: string[] = [];
  if (a['pageBreakBefore']) out.push('<w:pageBreakBefore/>');
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

function paragraphXml(node: PMNode): string {
  let runs = '';
  node.forEach((child) => (runs += inlineXml(child)));
  return `<w:p>${paraProps(node)}${runs}</w:p>`;
}

/** A top-level block → its OOXML (E1: paragraphs only; tables come in E2). */
function blockXml(node: PMNode): string {
  if (node.type.name === 'paragraph') return paragraphXml(node);
  return '';
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${OR_NS}/officeDocument" Target="word/document.xml"/></Relationships>`;

/**
 * Serialise a bapbong document back to `.docx` bytes. Phase E1 covers
 * paragraphs + runs + the common character marks; other content is skipped for
 * now (round-trips as empty / plain).
 */
export async function exportDocx(doc: PMNode): Promise<Uint8Array> {
  let body = '';
  doc.forEach((block) => (body += blockXml(block)));
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'uint8array' });
}
