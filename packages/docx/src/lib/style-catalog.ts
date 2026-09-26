/**
 * The styles a document defines, and what one of them looks like on a
 * paragraph — for a caller that wants to apply "Quote" or a company style
 * without a style sheet in the model (paragraph styles are baked into
 * attrs and marks at import).
 *
 * The look is not re-derived here: a one-paragraph package is built around
 * the document's own styles.xml (and the parts a style can reach: theme,
 * numbering, settings) and run through the real importer, so a style comes
 * out exactly as it would have from the file.
 */
import JSZip from 'jszip';
import { importDocx } from './docx.js';

/** One style of `word/styles.xml`. */
export interface DocumentStyle {
  id: string;
  /** The name Word shows ("Heading 1", "Quote", "My Style"). */
  name: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering';
  basedOn?: string;
  /** Defined by the document's author, not one of Word's built-ins. */
  custom: boolean;
  /** The type's default (Normal, for paragraphs). */
  isDefault: boolean;
  /** Hidden from Word's style gallery (semiHidden / hidden). */
  hidden: boolean;
}

const decode = (s: string) =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');

const on = (tag: string, name: string) => {
  const v = new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
  return v === '1' || v === 'true' || v === 'on';
};

/** Every style `stylesXml` defines, in file order. */
export function documentStyles(stylesXml: string): DocumentStyle[] {
  const out: DocumentStyle[] = [];
  for (const m of stylesXml.matchAll(
    /<w:style\b([^>]*?)(?:\/>|>([\s\S]*?)<\/w:style>)/g,
  )) {
    const tag = m[1];
    const body = m[2] ?? '';
    const id = /\bw:styleId="([^"]*)"/.exec(tag)?.[1];
    if (!id) continue;
    const type = (/\bw:type="([^"]*)"/.exec(tag)?.[1] ??
      'paragraph') as DocumentStyle['type'];
    const name = /<w:name\b[^>]*\bw:val="([^"]*)"/.exec(body)?.[1];
    const basedOn = /<w:basedOn\b[^>]*\bw:val="([^"]*)"/.exec(body)?.[1];
    out.push({
      id: decode(id),
      name: decode(name ?? id),
      type,
      ...(basedOn ? { basedOn: decode(basedOn) } : {}),
      custom: on(tag, 'w:customStyle'),
      isDefault: on(tag, 'w:default'),
      hidden: /<w:(?:semiHidden|hidden)\b(?![^>]*w:val="(?:0|false)")/.test(
        body,
      ),
    });
  }
  return out;
}

/** What a paragraph in style `styleId` gets from the importer: its
 *  paragraph attrs and the marks on its text (as JSON, schema-free). */
export interface StyleFormatting {
  attrs: Record<string, unknown>;
  marks: { type: string; attrs?: Record<string, unknown> }[];
}

/** Parts a style can reach while it resolves. */
const STYLE_PARTS = [
  'word/styles.xml',
  'word/theme/theme1.xml',
  'word/numbering.xml',
  'word/settings.xml',
  'word/fontTable.xml',
];

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const escAttr = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/**
 * The look of paragraph style `styleId` from the package `source` (a
 * document's carried JSZip), or null when the package has no such
 * paragraph style.
 */
export async function styleFormatting(
  source: JSZip,
  styleId: string,
): Promise<StyleFormatting | null> {
  const stylesXml = await source.file('word/styles.xml')?.async('string');
  if (!stylesXml) return null;
  const style = documentStyles(stylesXml).find(
    (s) => s.id === styleId && s.type === 'paragraph',
  );
  if (!style) return null;
  const probe = new JSZip();
  for (const path of STYLE_PARTS) {
    const part = source.file(path);
    if (part) probe.file(path, await part.async('uint8array'));
  }
  probe.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${W_NS}"><w:body>` +
      `<w:p><w:pPr><w:pStyle w:val="${escAttr(styleId)}"/></w:pPr>` +
      `<w:r><w:t>x</w:t></w:r></w:p>` +
      `<w:sectPr/></w:body></w:document>`,
  );
  const bytes = await probe.generateAsync({ type: 'arraybuffer' });
  const { doc } = await importDocx(bytes);
  const p = doc.firstChild;
  if (!p) return null;
  const text = p.firstChild;
  return {
    attrs: { ...p.attrs },
    marks: (text?.marks ?? []).map(
      (m) => m.toJSON() as StyleFormatting['marks'][number],
    ),
  };
}
