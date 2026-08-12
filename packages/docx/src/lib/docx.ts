import JSZip from 'jszip';
import { Node as PMNode, Mark, Schema } from 'prosemirror-model';
import {
  DocxImportError,
  IMPORT_ERROR_MESSAGES,
  errorForSniff,
  sniffDocx,
} from './sniff.js';
import { decryptOfficeFile, WrongPasswordError } from './crypto-docx.js';

/** Decrypt a password-protected document, mapping the crypto layer's failures
 *  onto the classified import errors a shell already handles. */
async function decryptDocx(
  bytes: Uint8Array,
  password: string,
): Promise<Uint8Array> {
  try {
    return await decryptOfficeFile(bytes, password);
  } catch (err) {
    if (err instanceof WrongPasswordError) {
      throw new DocxImportError(
        'wrong-password',
        IMPORT_ERROR_MESSAGES['wrong-password'],
      );
    }
    // An unsupported scheme (Word 2007 "Standard", a cipher we don't do) is
    // a dead end for the unlock prompt — say so rather than looping on it.
    throw new DocxImportError(
      'unsupported-encryption',
      IMPORT_ERROR_MESSAGES['unsupported-encryption'],
      err instanceof Error ? err.message : String(err),
    );
  }
}
import {
  commentSchema,
  schema,
  type Align,
  type FieldInfo,
  type Indent,
  type ListInfo,
  type NumberingDefs,
  type Spacing,
} from '@shadow-garden/bapbong-model';
import {
  attrOf,
  child,
  children,
  findDescendant,
  mergeRunProps,
  normalizeHex,
  OoxmlNode,
  parseRunProps,
  parseXml,
  RunProps,
  serializeOoxml,
  shdFill,
} from './ooxml.js';
import type {
  BorderSide,
  BorderStyle,
  ShapeSpec,
  TableBorders,
} from '@shadow-garden/bapbong-contracts';
import { audit } from './audit.js';
import { buildStyleRegistry, StyleRegistry } from './styles.js';
import { buildNumbering, NumberingResolver } from './numbering.js';
import { buildRels, Relationship } from './rels.js';
import {
  buildThemeFillResolver,
  buildThemeFontResolver,
  buildThemeResolver,
  drawingColor,
  ThemeFillResolver,
  ThemeFontResolver,
  ThemeResolver,
} from './theme.js';

export type DocxInput = ArrayBuffer | Uint8Array | Blob;

/**
 * Result of importing a .docx: the ProseMirror document plus the raw
 * `word/document.xml` we parsed it from. The raw string is kept so a later
 * export step can round-trip parts of the document we don't model yet.
 */
/** Page geometry in CSS px (structurally a bapbong-contracts PageConfig). */
export interface PageConfig {
  width: number;
  height: number;
  margin: { top: number; right: number; bottom: number; left: number };
  /** Header/footer band distance from the page edge (w:pgMar @w:header/
   *  @w:footer). Absent → Word's default 720 twips (48px). */
  headerDistance?: number;
  footerDistance?: number;
  /** Binding gutter (w:pgMar @w:gutter) added to the left content edge. */
  gutter?: number;
}

/** A document comment (structurally a bapbong-contracts CommentData). */
export interface CommentData {
  id: number;
  author: string;
  date: string;
  text: string;
}

/** A comment author (structurally a bapbong-contracts IUser). */
interface IUser {
  id: string;
  name: string;
  email?: string;
  avatar?: string;
}

/** A comment thread root for doc.attrs.comments (structurally a
 *  bapbong-contracts CommentNode); `body` is commentSchema doc JSON. */
interface CommentNode {
  id: number;
  parentId: number | null;
  user: IUser;
  date: string;
  body: unknown;
  resolved: boolean;
}

/** One section's effective chrome stories (inheritance already applied). */
export interface SectionChrome {
  headers: Record<string, PMNode>;
  footers: Record<string, PMNode>;
  titlePg: boolean;
}

export interface DocxImport {
  doc: PMNode;
  rawDocumentXml: string;
  /** Header stories keyed by w:type ("default" | "first" | "even"). */
  headers: Record<string, PMNode>;
  /** Footer stories keyed by w:type. */
  footers: Record<string, PMNode>;
  /** Footnote body stories keyed by display number (w:footnoteReference). Laid
   *  out at the bottom of the page their reference falls on. Endnotes are NOT
   *  here — they're appended to `doc`. */
  footnotes: Record<number, PMNode>;
  /** w:titlePg — page 1 uses the "first" header/footer (headers/footers['first']). */
  titlePg: boolean;
  /** w:evenAndOddHeaders — even pages use the "even" header/footer. */
  evenAndOdd: boolean;
  /** Per-section chrome, aligned with doc.attrs.sections, with Word's "Link
   *  to Previous" resolved: a section without its own w:headerReference of a
   *  type inherits the previous section's story. Present only when the doc
   *  has ≥2 sections AND at least one declares its own chrome or titlePg —
   *  the flat `headers`/`footers` (the LAST section's) cover the rest. */
  sectionChrome?: SectionChrome[];
  /** Comments referenced by the body (w:commentRange), in appearance order. */
  comments: CommentData[];
  /** Page size + margins from w:sectPr (A4 @96dpi when unspecified). */
  page: PageConfig;
  /** Default tab interval in px (settings w:defaultTabStop); absent → the
   *  layout engine's 0.5" default. */
  tabWidth?: number;
  /** The loaded source package — pass to `exportDocx(doc, { carry })` so the
   *  parts bapbong doesn't model yet (styles, numbering, headers/footers, …)
   *  survive the round-trip instead of being dropped. */
  raw: JSZip;
}

/** Footnote/endnote bodies + a counter that numbers references in document
 *  order. Footnote bodies are laid out at the bottom of the page their
 *  reference falls on (by the layout engine, via `DocxImport.footnotes`);
 *  endnote bodies are appended at the document end. */
interface NotesRegistry {
  bodies: { footnote: Map<string, OoxmlNode>; endnote: Map<string, OoxmlNode> };
  refs: { kind: 'footnote' | 'endnote'; id: string; num: number }[];
  counter: { footnote: number; endnote: number };
  /** Assign (and remember) the display number for a reference. */
  ref(kind: 'footnote' | 'endnote', id: string): number;
}

/** Comment bodies (w:comment) + the live set covering the text being parsed.
 *  `active` toggles on w:commentRangeStart/End; `used` records referenced
 *  comments in first-appearance order. `paraToId` + `ext` carry the threaded-
 *  comment data from word/commentsExtended.xml (w15): replies link by the
 *  parent paragraph's w14:paraId, and `done` is the resolved flag. */
interface CommentsRegistry {
  defs: Map<
    number,
    { author: string; date: string; body: OoxmlNode; paraIds: string[] }
  >;
  paraToId: Map<string, number>;
  ext: Map<string, { parentParaId: string | null; done: boolean }>;
  active: Set<number>;
  used: number[];
}

interface Ctx {
  styles: StyleRegistry;
  numbering: NumberingResolver;
  rels: Map<string, Relationship>;
  media: Map<string, string>; // zip path → data URL
  resolveTheme: ThemeResolver;
  resolveFont: ThemeFontResolver;
  /** Shape fill from an `a:fillRef` (theme format scheme + placeholder). */
  resolveThemeFill: ThemeFillResolver;
  notes: NotesRegistry;
  comments: CommentsRegistry;
  /** Schema the doc nodes/marks are created with (model's by default; the editor
   *  may inject a composed schema so plugin-contributed marks are imported). */
  schema: Schema;
  /** Page content-box width in px (page minus side margins) — what
   *  percentage-based table widths (w:tblW/w:tcW type="pct") resolve against. */
  contentWidth: number;
  /** VML shapetype registry (`v:shapetype` id → o:spt), filled as picts are
   *  parsed in document order — Word always defines a type before its first
   *  `type="#id"` reference. Lazily created by parseVmlShape. */
  vmlShapeTypes?: Map<string, number>;
  /** Generated fields open across paragraph boundaries. A TOC field begins in
   *  its first entry's paragraph and closes many paragraphs later, so the span
   *  can't live in parseParagraph's local state: while this stack is non-empty
   *  every paragraph parsed belongs to the innermost field and is stamped with
   *  that SHARED object — identity is what marks the span (see the model's
   *  `fieldAt`). */
  openFields: FieldInfo[];
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
function propsToMarks(p: RunProps, ctx: Ctx): Mark[] {
  const marks: Mark[] = [];
  if (p.bold) marks.push(ctx.schema.marks['strong'].create());
  if (p.italic) marks.push(ctx.schema.marks['em'].create());
  if (p.underline) marks.push(ctx.schema.marks['underline'].create());
  if (p.strike) marks.push(ctx.schema.marks['strike'].create());
  if (p.dstrike) marks.push(ctx.schema.marks['dstrike'].create());
  if (p.smallCaps) marks.push(ctx.schema.marks['smallCaps'].create());
  if (p.color)
    marks.push(ctx.schema.marks['textColor'].create({ color: p.color }));
  if (p.sizePt !== undefined)
    marks.push(ctx.schema.marks['fontSize'].create({ size: p.sizePt }));
  if (p.fontFamily)
    marks.push(ctx.schema.marks['fontFamily'].create({ family: p.fontFamily }));
  if (p.highlight)
    marks.push(ctx.schema.marks['highlight'].create({ color: p.highlight }));
  if (p.vertAlign)
    marks.push(ctx.schema.marks['vertAlign'].create({ value: p.vertAlign }));
  // 0 is a real value (an explicit "back to the baseline" override), so test
  // for presence, not truthiness.
  if (p.position !== undefined && ctx.schema.marks['position'])
    marks.push(ctx.schema.marks['position'].create({ halfPoints: p.position }));
  if (p.letterSpacing !== undefined && ctx.schema.marks['letterSpacing'])
    marks.push(
      ctx.schema.marks['letterSpacing'].create({ twips: p.letterSpacing }),
    );
  if (p.charScale !== undefined && ctx.schema.marks['charScale'])
    marks.push(ctx.schema.marks['charScale'].create({ percent: p.charScale }));
  return marks;
}

/** Wingdings code point → the Unicode character Word draws for it. Keyed by
 *  the LOW byte, because the same glyph reaches us two ways: as `w:sym
 *  w:char="F06F"` (PUA-offset) and as an ordinary run of text whose rFonts
 *  says Wingdings — a form's ticked checkbox is often just the letter "x"
 *  in that font. Both go through symbolChar. */
const WINGDINGS: Record<number, string> = {
  0x4a: '☺',
  0x4b: '😐',
  0x4c: '☹',
  0x4d: '💣',
  0x4e: '☠',
  0x51: '✈',
  0x6c: '●',
  0x6e: '■',
  // The checkbox family: Word draws these at text size, so the Unicode
  // BALLOT BOX forms match it far better than the tiny ▫/□ geometric ones.
  0x6f: '☐',
  0x70: '☐',
  0x71: '☐',
  0x72: '☐',
  0x73: '☐',
  0x75: '◆',
  0x78: '☒', // "x" — a ticked box, the usual mark in Vietnamese HR forms
  0xa7: '▪',
  0xa8: '☐',
  0xb7: '•',
  0xe0: '→',
  0xfb: '✗',
  0xfc: '✔',
  0xfd: '☒',
  0xfe: '☑',
};

/** Symbol font: only the few glyphs documents actually lean on (its letters
 *  are Greek, which we leave to the font). */
const SYMBOL_FONT: Record<number, string> = {
  0xb7: '•',
  0xd7: '×',
  0xb0: '°',
  0xa0: '€',
};

/** Fonts whose bytes are pictures, not letters. */
function symbolTable(font: string | undefined): Record<number, string> | null {
  if (!font) return null;
  const f = font.toLowerCase();
  if (f.startsWith('wingdings')) return WINGDINGS;
  if (f === 'symbol') return SYMBOL_FONT;
  return null;
}

/** One character of a symbol font → the Unicode Word shows. Mapped chars come
 *  back font-independent (no `font`), so they render anywhere; unmapped ones
 *  keep their code point AND the font name, and the caller tags them with a
 *  fontFamily mark so the glyph still appears where the font is installed
 *  (and survives a save either way). */
function symbolChar(
  code: string | undefined,
  font: string | undefined,
): { text: string; font?: string } {
  if (!code) return { text: '' };
  const n = parseInt(code, 16);
  if (Number.isNaN(n)) return { text: '' };
  // w:sym codes sit in the PUA (F0xx); a plain run's character is the byte.
  const mapped = symbolTable(font)?.[n & 0xff];
  if (mapped) return { text: mapped };
  return { text: String.fromCodePoint(n), ...(font && { font }) };
}

/** A run of ordinary text set in a symbol font, translated to Unicode. Null
 *  when the font isn't one (the overwhelmingly common case — one cheap map
 *  lookup) or when nothing in the text maps, so the run passes through
 *  untouched and keeps its font. */
function symbolFontText(text: string, font: string | undefined): string | null {
  const table = symbolTable(font);
  if (!table) return null;
  let changed = false;
  let out = '';
  for (const ch of text) {
    const mapped = table[ch.codePointAt(0) as number];
    if (mapped) {
      out += mapped;
      changed = true;
    } else {
      out += ch;
    }
  }
  return changed ? out : null;
}

/** Whether a run carries an explicit page break (w:br w:type="page"). */
function hasPageBreak(run: OoxmlNode): boolean {
  return run.children.some(
    (n) => n.name === 'w:br' && attrOf(n, 'w:type') === 'page',
  );
}

/** Flatten tracked changes for the "accept all changes" view: w:ins unwraps
 *  to its child runs (inserted text was otherwise lost — the run loop only
 *  walked top-level w:r), w:del is dropped (deleted text). Other nodes pass
 *  through. Recurses so ins/del can nest or wrap hyperlinks. */
function effectiveChildren(nodes: OoxmlNode[]): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  for (const node of nodes) {
    if (node.name === 'w:del' || node.name === 'w:moveFrom') {
      // Deliberately dropped (accept-all-changes view) — subtree included.
      audit.markSubtree(node);
      continue;
    }
    if (node.name === 'w:ins' || node.name === 'w:moveTo') {
      audit.mark(node);
      out.push(...effectiveChildren(node.children));
    } else if (node.name === 'w:sdt') {
      out.push(...effectiveChildren(unwrapSdt([node])));
    } else {
      out.push(node);
    }
  }
  return out;
}

/** Content controls (w:sdt) unwrap to their w:sdtContent children — the
 *  control chrome is dropped, the content (paragraphs/tables at block level,
 *  runs inline — a w14:checkbox's ☒/☐ glyph run included) survives. Recurses
 *  so nested controls (cover pages hold several) fully unwrap. */
function unwrapSdt(nodes: OoxmlNode[]): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  for (const node of nodes) {
    if (node.name === 'w:sdt') {
      audit.mark(node);
      const content = child(node, 'w:sdtContent');
      if (content) out.push(...unwrapSdt(content.children));
    } else {
      out.push(node);
    }
  }
  return out;
}

// Run children the loop below consumes (page-type w:br is handled at the
// paragraph level via hasPageBreak, but it IS handled — audit-marked here).
const RUN_CHILD_TAGS = new Set([
  'w:t',
  'w:tab',
  'w:sym',
  'w:footnoteReference',
  'w:endnoteReference',
  'w:br',
]);

/** Inline nodes for a run, splitting text at soft w:br into hard_break nodes
 *  (page breaks are handled at the paragraph level, not here). */
function runInlineNodes(run: OoxmlNode, marks: Mark[], ctx: Ctx): PMNode[] {
  const group = parseGroup(run, ctx);
  if (group) return group;
  // Shapetypes are registered for every pict up front — see
  // registerVmlShapeTypes for why this cannot live in the shape parser.
  registerVmlShapeTypes(run, ctx);
  const image =
    parseImage(run, ctx) ??
    parseShape(run, ctx) ??
    parseVmlImage(run, ctx) ??
    parseVmlShape(run, ctx);
  if (image) return [image];

  const out: PMNode[] = [];
  // A run set in a symbol font spells its pictures with ordinary letters —
  // a ticked checkbox is the letter "x" in Wingdings. Translate those to
  // Unicode and drop the font mark, so the glyph survives on machines
  // without the font (and stays legible in the a11y mirror and on export).
  const runFont = marks.find((m) => m.type.name === 'fontFamily')?.attrs[
    'family'
  ] as string | undefined;
  // Only when something ACTUALLY translates: a symbol-font run we can't read
  // must keep its font, or its glyph is lost for everyone who has it.
  const asSymbols =
    symbolTable(runFont) &&
    run.children.some(
      (c) => c.name === 'w:t' && symbolFontText(c.text, runFont) !== null,
    )
      ? runFont
      : undefined;
  const textMarks = asSymbols
    ? marks.filter((m) => m.type.name !== 'fontFamily')
    : marks;
  let buf = '';
  const flush = () => {
    if (buf.length > 0) out.push(ctx.schema.text(buf, textMarks));
    buf = '';
  };
  for (const node of run.children) {
    if (RUN_CHILD_TAGS.has(node.name)) audit.mark(node);
    if (node.name === 'w:t')
      buf += (asSymbols && symbolFontText(node.text, asSymbols)) || node.text;
    else if (node.name === 'w:tab') buf += '\t';
    else if (node.name === 'w:sym') {
      const sym = symbolChar(attrOf(node, 'w:char'), attrOf(node, 'w:font'));
      if (!sym.text) continue;
      if (sym.font) {
        // Unmapped symbol: its glyph lives in the symbol font's PUA range —
        // switch just this character to that font.
        flush();
        out.push(
          ctx.schema.text(sym.text, [
            ...marks.filter((m) => m.type.name !== 'fontFamily'),
            ctx.schema.marks['fontFamily'].create({ family: sym.font }),
          ]),
        );
      } else {
        buf += sym.text;
      }
    } else if (
      node.name === 'w:footnoteReference' ||
      node.name === 'w:endnoteReference'
    ) {
      const kind = node.name === 'w:footnoteReference' ? 'footnote' : 'endnote';
      const id = attrOf(node, 'w:id');
      if (id && ctx.notes.bodies[kind].has(id)) {
        flush();
        const num = ctx.notes.ref(kind, id);
        const refMarks = [
          ...marks,
          ctx.schema.marks['vertAlign'].create({ value: 'super' }),
        ];
        // Footnotes carry a `footnote` mark so the layout engine can match the
        // reference to its page-bottom body; endnotes stay plain superscripts
        // (their bodies are appended at the document end).
        if (kind === 'footnote')
          refMarks.push(ctx.schema.marks['footnote'].create({ num }));
        out.push(ctx.schema.text(String(num), refMarks));
      }
    } else if (node.name === 'w:br' && attrOf(node, 'w:type') !== 'page') {
      flush();
      out.push(ctx.schema.nodes['hard_break'].create());
    }
  }
  flush();
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
    : child(anchor, 'wp:wrapSquare') ||
        child(anchor, 'wp:wrapTight') ||
        child(anchor, 'wp:wrapThrough')
      ? 'square'
      : 'none'; // wrapNone / absent: paints without affecting text

  const float: Record<string, unknown> = { wrap };
  // z-order: behindDoc="1" puts the drawing UNDER the text (watermarks, page
  // backgrounds). Word's default is in front. Dropping this flag on import is
  // what silently pulled behind-text images in front of the text on save.
  if (attrOf(anchor, 'behindDoc') === '1') float['behind'] = true;

  const posH = child(anchor, 'wp:positionH');
  if (posH) {
    const align = child(posH, 'wp:align')?.text.trim();
    if (align === 'left' || align === 'right' || align === 'center')
      float['hAlign'] = align;
    const off = emuToPxZero(child(posH, 'wp:posOffset')?.text);
    if (off !== undefined && float['hAlign'] === undefined)
      float['hOffset'] = off;
    const rel = attrOf(posH, 'relativeFrom');
    float['hRel'] = rel === 'page' ? 'page' : 'margin'; // column/margin/… ≈ margin
  }
  const posV = child(anchor, 'wp:positionV');
  if (posV) {
    const off = emuToPxZero(child(posV, 'wp:posOffset')?.text);
    if (off !== undefined) float['vOffset'] = off;
    const rel = attrOf(posV, 'relativeFrom');
    float['vRel'] =
      rel === 'page' ? 'page' : rel === 'margin' ? 'margin' : 'paragraph';
  }
  // Text-to-image gaps (EMU attrs on the anchor itself).
  for (const side of ['distL', 'distR', 'distT', 'distB'] as const) {
    const v = emuToPxZero(attrOf(anchor, side));
    if (v !== undefined) float[side] = v;
  }
  return float;
}

/** The run's w:drawing, looking through mc:AlternateContent — Word wraps
 *  shapes and some images in Choice/Fallback pairs (`Requires="wps"` etc.);
 *  the Choice branch carries the richer DrawingML, so it wins over Fallback. */
function runDrawing(run: OoxmlNode): OoxmlNode | undefined {
  const direct = child(run, 'w:drawing');
  if (direct) return direct;
  const alt = child(run, 'mc:AlternateContent');
  if (!alt) return undefined;
  for (const branch of ['mc:Choice', 'mc:Fallback']) {
    for (const b of children(alt, branch)) {
      const d = child(b, 'w:drawing');
      if (d) return d;
    }
  }
  return undefined;
}

/** A wpg group (wp:anchor holding wpg:wgp) flattened to one floating image
 *  per member picture: child coordinates live in the group's child space
 *  (a:chOff/chExt) and scale to its on-page extent (a:ext), each member
 *  becoming its own float offset from the group anchor. V1 handles pic:pic
 *  members (bitmaps); member wps shapes and nested groups stay unmodelled.
 *  Inline (non-anchored) groups fall through to the single-image path. */
function parseGroup(run: OoxmlNode, ctx: Ctx): PMNode[] | null {
  const drawing = runDrawing(run);
  if (!drawing) return null;
  const wgp = findDescendant(drawing, 'wpg:wgp');
  if (!wgp) return null;
  const baseFloat = parseAnchorFloat(drawing);
  if (!baseFloat) return null;
  const num = (n: OoxmlNode | undefined, a: string) =>
    Number(attrOf(n, a) ?? '0');
  const xfrm = child(child(wgp, 'wpg:grpSpPr'), 'a:xfrm');
  const ext = child(xfrm, 'a:ext');
  const chOff = child(xfrm, 'a:chOff');
  const chExt = child(xfrm, 'a:chExt');
  const chW = num(chExt, 'cx') || num(ext, 'cx') || 1;
  const chH = num(chExt, 'cy') || num(ext, 'cy') || 1;
  const sx = (num(ext, 'cx') || chW) / chW;
  const sy = (num(ext, 'cy') || chH) / chH;

  const out: PMNode[] = [];
  for (const pic of children(wgp, 'pic:pic')) {
    const blip = findDescendant(pic, 'a:blip');
    const embed = attrOf(blip, 'r:embed') ?? attrOf(blip, 'r:link');
    const rel = embed ? ctx.rels.get(embed) : undefined;
    if (!rel) continue;
    const target = rel.target.replace(/^\/+/, '');
    const src =
      ctx.media.get(`word/${target}`) ??
      ctx.media.get(target) ??
      (/^https?:\/\//i.test(rel.target) ? rel.target : undefined);
    if (!src) continue;
    const picXfrm = child(child(pic, 'pic:spPr'), 'a:xfrm');
    const off = child(picXfrm, 'a:off');
    const cext = child(picXfrm, 'a:ext');
    const emuPx = (emu: number) => Math.round(emu / 9525);
    out.push(
      ctx.schema.nodes['image'].create({
        src,
        width: emuPx(num(cext, 'cx') * sx),
        height: emuPx(num(cext, 'cy') * sy),
        alt: attrOf(findDescendant(pic, 'pic:cNvPr'), 'descr') ?? '',
        float: {
          ...baseFloat,
          hOffset:
            ((baseFloat['hOffset'] as number) ?? 0) +
            emuPx((num(off, 'x') - num(chOff, 'x')) * sx),
          vOffset:
            ((baseFloat['vOffset'] as number) ?? 0) +
            emuPx((num(off, 'y') - num(chOff, 'y')) * sy),
        },
      }),
    );
  }
  return out.length > 0 ? out : null;
}

/** Rotation (clockwise degrees) from a subtree's a:xfrm@rot (1/60000 deg). */
function xfrmRotation(root: OoxmlNode | undefined): number {
  const rot = Number(attrOf(findDescendant(root, 'a:xfrm'), 'rot'));
  return rot ? Math.round((rot / 60000) * 100) / 100 : 0;
}

/** Extract an image (inline or floating) from a run's w:drawing, if any. */
function parseImage(run: OoxmlNode, ctx: Ctx): PMNode | null {
  const drawing = runDrawing(run);
  if (!drawing) return null;
  const blip = findDescendant(drawing, 'a:blip');
  const embed = attrOf(blip, 'r:embed') ?? attrOf(blip, 'r:link');
  const rel = embed ? ctx.rels.get(embed) : undefined;
  if (!rel) return null;
  const target = rel.target.replace(/^\/+/, '');
  // Externally-linked picture (TargetMode="External"): no media part — the
  // rel target IS the image URL.
  const src =
    ctx.media.get(`word/${target}`) ??
    ctx.media.get(target) ??
    (/^https?:\/\//i.test(rel.target) ? rel.target : undefined);
  if (!src) return null;

  const extent = findDescendant(drawing, 'wp:extent');
  const docPr = findDescendant(drawing, 'wp:docPr');
  const float = parseAnchorFloat(drawing);
  return ctx.schema.nodes['image'].create({
    src,
    width: emuToPx(attrOf(extent, 'cx')),
    height: emuToPx(attrOf(extent, 'cy')),
    alt: attrOf(docPr, 'descr') ?? attrOf(docPr, 'title') ?? '',
    float,
    rotation: xfrmRotation(drawing),
  });
}

/** Legacy VML image (w:object / w:pict holding v:shape + v:imagedata) — how
 *  older Word versions and OLE embeds carry pictures. The bitmap rides
 *  v:imagedata's relationship; the display size lives in the v:shape style
 *  ("width:108.3pt;height:61.35pt"), with w:object's dxaOrig/dyaOrig (twips)
 *  as the fallback. */
/**
 * Register a pict's shapetypes (id → o:spt) doc-wide. Word defines a type
 * once, before its first reference, and later picts point at it by `@type`.
 *
 * Runs for EVERY pict before the image/shape paths are tried, not inside the
 * shape parser: a pict that holds a picture returns from `parseVmlImage`
 * first, so a type declared alongside it would never be registered — and a
 * later shape referencing it would be dropped for want of a kind.
 */
function registerVmlShapeTypes(run: OoxmlNode, ctx: Ctx): void {
  const pict = child(run, 'w:pict');
  if (!pict) return;
  const types = (ctx.vmlShapeTypes ??= new Map());
  for (const st of children(pict, 'v:shapetype')) {
    const id = attrOf(st, 'id');
    const spt = Number(attrOf(st, 'o:spt'));
    if (id && Number.isFinite(spt)) types.set(id, spt);
    // A shapetype is a definition, not content — nothing inside it is ours
    // to render, so the whole subtree counts as consumed by design.
    audit.markSubtree(st);
  }
}

function parseVmlImage(run: OoxmlNode, ctx: Ctx): PMNode | null {
  const holder = child(run, 'w:object') ?? child(run, 'w:pict');
  if (!holder) return null;
  const imagedata = findDescendant(holder, 'v:imagedata');
  const rid = attrOf(imagedata, 'r:id');
  const rel = rid ? ctx.rels.get(rid) : undefined;
  if (!rel) return null;
  const target = rel.target.replace(/^\/+/, '');
  const src = ctx.media.get(`word/${target}`) ?? ctx.media.get(target);
  if (!src) return null;

  const vshape = findDescendant(holder, 'v:shape');
  const style = attrOf(vshape, 'style') ?? '';
  // Alt text, in the order the formats define it: VML's own `alt` is THE
  // alternative text ("displayed instead of a graphic", read out by screen
  // readers); o:title is an Office label, useful only as a fallback.
  //
  // Both are read BEFORE choosing. `a || b` would short-circuit, and the
  // audit counts an attribute as covered only when the code asks for it —
  // so the fallback would show up as an unread gap on every file that has
  // the first one.
  const altAttr = attrOf(vshape, 'alt');
  const oTitle = attrOf(imagedata, 'o:title');
  const vmlAltText = altAttr || oTitle || '';
  // The frame preset this picture rides (`type="#_x0000_t75"`, o:spt 75).
  // Read so the shape is fully accounted for — the registry has already
  // recorded the definition it points at.
  attrOf(vshape, 'type');
  const ptToPx = (m: RegExpExecArray | null) =>
    m ? Math.round((parseFloat(m[1]) * 96) / 72) : null;
  const width =
    ptToPx(/(?:^|;)width:([\d.]+)pt/.exec(style)) ??
    (Number(attrOf(holder, 'w:dxaOrig'))
      ? twipsToPx(Number(attrOf(holder, 'w:dxaOrig')))
      : null);
  const height =
    ptToPx(/(?:^|;)height:([\d.]+)pt/.exec(style)) ??
    (Number(attrOf(holder, 'w:dyaOrig'))
      ? twipsToPx(Number(attrOf(holder, 'w:dyaOrig')))
      : null);

  return ctx.schema.nodes['image'].create({
    src,
    width,
    height,
    alt: vmlAltText,
    float: null,
  });
}

/** Color of a node's <a:solidFill>, through the full DrawingML colour union
 *  and its transform stack (see theme.ts `drawingColor`). */
function solidFillColor(
  node: OoxmlNode | undefined,
  ctx: Ctx,
): string | undefined {
  return drawingColor(child(node, 'a:solidFill'), ctx.resolveTheme);
}

/** A drawn wps shape (rect / straight connector) in a run's drawing — the
 *  checkbox squares and horizontal rules real documents draw with Shapes.
 *  Rides the image node (same box semantics) with a `shape` payload; other
 *  prstGeom kinds stay unmodelled (dropped) for now. */
function parseShape(run: OoxmlNode, ctx: Ctx): PMNode | null {
  const drawing = runDrawing(run);
  if (!drawing) return null;
  const wsp = findDescendant(drawing, 'wps:wsp');
  const spPr = child(wsp, 'wps:spPr');
  const prst = attrOf(child(spPr, 'a:prstGeom'), 'prst');
  const textbox = parseTextbox(wsp, ctx);
  // Geometry we paint natively (ShapeSpec kinds mirror the prst tokens);
  // anything else with a textbox degrades to a rect frame — the text matters
  // more than the fancy outline — and without one stays unmodelled.
  const KIND: Record<string, ShapeSpec['kind']> = {
    rect: 'rect',
    line: 'line',
    straightConnector1: 'line',
    ellipse: 'ellipse',
    roundRect: 'roundRect',
    rightArrow: 'rightArrow',
    horizontalScroll: 'horizontalScroll',
  };
  const kind = (prst && KIND[prst]) || (textbox ? 'rect' : null);
  if (!kind) return null;

  const shape: Record<string, unknown> = { kind };
  // Adjust value → corner ratio, and ONLY for roundRect: `adj` is per-preset
  // (a horizontalScroll's adj is the size of its curl), so reading it blindly
  // would feed one preset's parameter into another's geometry.
  if (prst === 'roundRect') {
    const gd = child(child(child(spPr, 'a:prstGeom'), 'a:avLst'), 'a:gd');
    const fmla = attrOf(gd, 'fmla'); // "val 12500"
    const n = fmla?.startsWith('val ') ? Number(fmla.slice(4)) : NaN;
    if (Number.isFinite(n))
      shape['cornerRatio'] = Math.min(Math.max(n / 100000, 0), 0.5);
  }
  const style = child(wsp, 'wps:style');
  const ln = child(spPr, 'a:ln');
  if (!child(ln, 'a:noFill')) {
    const w = attrOf(ln, 'w'); // outline width in EMU
    shape['strokeWidth'] = w ? Math.max(1, Math.round(Number(w) / 9525)) : 1;
    // Dash: a:custDash states the pattern exactly (d/sp are ST_Percentage
    // "relative to the line width", 100000 = 100% = one stroke width — the
    // same unit the model uses), so it converts without loss. A named
    // a:prstDash only says WHICH preset, and the spec doesn't publish their
    // lengths, so it keeps the generic pattern rather than an invented one.
    const custDash = child(ln, 'a:custDash');
    const stops = custDash ? children(custDash, 'a:ds') : [];
    if (stops.length > 0) {
      const pattern: number[] = [];
      for (const ds of stops) {
        pattern.push(pctToRatio(attrOf(ds, 'd')), pctToRatio(attrOf(ds, 'sp')));
      }
      shape['dash'] = pattern;
    } else {
      const prstDash = attrOf(child(ln, 'a:prstDash'), 'val');
      if (prstDash && prstDash !== 'solid') shape['dash'] = DML_NAMED_DASH;
    }
    // cap: honored only when stated. ECMA says an omitted cap means `square`,
    // but every Word-authored theme in the wild writes cap="flat" on its line
    // styles, and shapes inherit that — so assuming square for a bare a:ln
    // would lengthen the ends of every existing line by half a stroke.
    // `flat` is the model's absent case, so it is not stored — that keeps one
    // canonical shape for a value with one meaning, and makes the round-trip
    // symmetric (no cap → cap="flat" on write → no cap on read).
    const cap = attrOf(ln, 'cap');
    if (cap === 'rnd') shape['cap'] = 'round';
    else if (cap === 'sq') shape['cap'] = 'square';
    // Arrowheads (ST_LineEndType: none|triangle|stealth|diamond|oval|arrow).
    // The model records only their presence, so any head that isn't `none`
    // counts — the painter draws one triangle either way.
    const headEnd = attrOf(child(ln, 'a:headEnd'), 'type');
    const tailEnd = attrOf(child(ln, 'a:tailEnd'), 'type');
    if (headEnd && headEnd !== 'none') shape['arrowStart'] = true;
    if (tailEnd && tailEnd !== 'none') shape['arrowEnd'] = true;
    // Direct outline color, else the style's line reference (how Word themes
    // shape outlines), else black.
    shape['stroke'] =
      solidFillColor(ln, ctx) ??
      drawingColor(findDescendant(style, 'a:lnRef'), ctx.resolveTheme) ??
      '#000000';
  }
  // Fill: an explicit a:noFill wins, then a direct a:solidFill, then the
  // shape style's a:fillRef resolved through the theme's format scheme —
  // which is how Word fills every shape inserted from the shape gallery.
  const fill = child(spPr, 'a:noFill')
    ? undefined
    : (solidFillColor(spPr, ctx) ??
      ctx.resolveThemeFill(child(style, 'a:fillRef')));
  if (fill) shape['fill'] = fill;
  if (attrOf(child(spPr, 'a:xfrm'), 'flipV') === '1') shape['flipV'] = true;

  const extent = findDescendant(drawing, 'wp:extent');
  const docPr = findDescendant(drawing, 'wp:docPr');
  const shapeDescr = attrOf(docPr, 'descr');
  const shapeTitle = attrOf(docPr, 'title');
  return ctx.schema.nodes['image'].create({
    src: '',
    width: emuToPxZero(attrOf(extent, 'cx')) ?? 0,
    height: emuToPxZero(attrOf(extent, 'cy')) ?? 0,
    // descr is "Alternative Text for Object"; `name` is what Word auto-names
    // the shape ("Rectangle 5"), which is no more a description than an id.
    // Both attrs are read up front (see parseVmlImage) so neither reads as a
    // gap when the other one wins.
    alt: shapeDescr || shapeTitle || '',
    float: parseAnchorFloat(drawing),
    shape,
    textbox,
    rotation: xfrmRotation(spPr),
  });
}

/** Textbox content of a wps shape (wps:txbx/w:txbxContent), as paragraph node
 *  JSON the layout engine flows inside the shape's box, plus the interior
 *  padding from wps:bodyPr (EMU attrs; absent → Word's 0.1"/0.05" defaults). */
function parseTextbox(
  wsp: OoxmlNode | undefined,
  ctx: Ctx,
): {
  paragraphs: unknown[];
  inset?: { l: number; t: number; r: number; b: number };
} | null {
  const paragraphs = txbxParagraphs(
    child(child(wsp, 'wps:txbx'), 'w:txbxContent'),
    ctx,
  );
  if (!paragraphs) return null;
  const bodyPr = child(wsp, 'wps:bodyPr');
  const ins = (name: string): number | undefined =>
    emuToPxZero(attrOf(bodyPr, name));
  const l = ins('lIns'),
    t = ins('tIns'),
    r = ins('rIns'),
    b = ins('bIns');
  const inset =
    l !== undefined || t !== undefined || r !== undefined || b !== undefined
      ? { l: l ?? 10, t: t ?? 5, r: r ?? 10, b: b ?? 5 }
      : undefined;
  return inset ? { paragraphs, inset } : { paragraphs };
}

/** Paragraph JSON of a `w:txbxContent` — shared by the modern (wps:txbx) and
 *  legacy VML (v:textbox) textbox paths. */
function txbxParagraphs(
  content: OoxmlNode | undefined,
  ctx: Ctx,
): unknown[] | null {
  if (!content) return null;
  const paragraphs = children(content, 'w:p').map((p) =>
    parseParagraph(p, ctx).toJSON(),
  );
  return paragraphs.length > 0 ? paragraphs : null;
}

// ── legacy VML shapes ───────────────────────────────────────────────
// Old Word (and .doc conversions) draws flowcharts as w:pict + v:* — rounded
// boxes with v:textbox content and straight connectors — with geometry in a
// CSS-ish `style` string. Everything maps onto the SAME image-node payload
// (shape / float / textbox) the modern wps path builds, so layout, painting,
// dragging and export need nothing new past line arrowheads.

/** A CSS-style VML length ("135pt", ".05in", "12px"; bare number = px) in
 *  px, or undefined. */
function cssLenToPx(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^(-?[\d.]+)(pt|in|cm|mm|px)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  if (!Number.isFinite(n)) return undefined;
  const u = m[2] ?? 'px';
  const f =
    u === 'pt'
      ? 96 / 72
      : u === 'in'
        ? 96
        : u === 'cm'
          ? 96 / 2.54
          : u === 'mm'
            ? 96 / 25.4
            : 1;
  return n * f;
}

/** The geometry bits of a VML `style` attribute. */
function parseVmlStyle(style: string): {
  absolute: boolean;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  zIndex?: number;
  flipV?: boolean;
} {
  const out: ReturnType<typeof parseVmlStyle> = { absolute: false };
  for (const part of style.split(';')) {
    const i = part.indexOf(':');
    if (i < 0) continue;
    const k = part.slice(0, i).trim();
    const v = part.slice(i + 1).trim();
    if (k === 'position') out.absolute = v === 'absolute';
    else if (k === 'margin-left') out.left = cssLenToPx(v);
    else if (k === 'margin-top') out.top = cssLenToPx(v);
    else if (k === 'width') out.width = cssLenToPx(v);
    else if (k === 'height') out.height = cssLenToPx(v);
    else if (k === 'z-index') out.zIndex = Number(v) || 0;
    else if (k === 'flip') out.flipV = v.includes('y');
  }
  return out;
}

const VML_COLORS: Record<string, string> = {
  black: '#000000',
  white: '#ffffff',
  red: '#ff0000',
  green: '#008000',
  blue: '#0000ff',
  yellow: '#ffff00',
  gray: '#808080',
  silver: '#c0c0c0',
};

/** A VML colour attribute ("white", "#FF0000", "white [3212]"), or undefined. */
function vmlColor(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const c = v.trim().split(' ')[0];
  if (c.startsWith('#')) return normalizeHex(c.slice(1)) ?? undefined;
  return VML_COLORS[c.toLowerCase()];
}

/**
 * `v:roundrect @arcsize` → the model's corner ratio.
 *
 * The spec states the value as a fraction **of half the shorter side** (0%
 * square, 100% circular, default 20%), while the model — and DrawingML's
 * `a:gd` adj — measure against the WHOLE shorter side. Hence the ÷2: a Word
 * 2003 rounded rectangle at the usual `10923f` is rounded 0.083 of its short
 * side, half as much as the DrawingML default the painter falls back to.
 *
 * Three encodings appear in the wild: `10923f` (fixed point, /65536), `25%`,
 * and a bare fraction.
 */
function vmlArcSize(v: string | undefined): number {
  const raw = v?.trim();
  let a = 0.2; // spec default
  if (raw) {
    const num = Number.parseFloat(raw);
    if (Number.isFinite(num)) {
      if (raw.endsWith('f')) a = num / 65536;
      else if (raw.endsWith('%')) a = num / 100;
      else a = num;
    }
  }
  return Math.min(Math.max(a, 0), 1) / 2;
}

/** Fallback pattern for the NAMED dashstyles (dash, longdashdot, …). The spec
 *  lists the names but not their lengths, and no document in our corpus uses
 *  one — rather than invent a table per name, every named style keeps the
 *  generic dash this code has always drawn. Numeric styles are exact. */
const VML_NAMED_DASH = [4, 3];

/** Same reasoning for DrawingML's a:prstDash names (dash, sysDot, lgDashDot…):
 *  the preset is named, its lengths are not published, so it renders generic.
 *  a:custDash, which DOES state lengths, is read exactly. */
const DML_NAMED_DASH = VML_NAMED_DASH;

/** ST_Percentage → ratio. The type is thousandths of a percent (100000 =
 *  100%), so a dash stop's `d="100000"` is one stroke width. */
function pctToRatio(v: string | undefined): number {
  const n = Number(v);
  return Number.isFinite(n) ? Math.max(0, n / 100000) : 0;
}

/** `v:stroke @dashstyle` → dash pattern in multiples of the stroke width.
 *  Numeric form ("1 1", "3 1 1 1") is the document's own measurement and is
 *  taken verbatim; a named form falls back (see {@link VML_NAMED_DASH}). */
function vmlDashPattern(v: string): number[] {
  const nums = v
    .trim()
    .split(/[\s,]+/)
    .map(Number);
  return nums.length > 0 && nums.every((n) => Number.isFinite(n) && n >= 0)
    ? nums
    : VML_NAMED_DASH;
}

/** v:textbox content + inset ("l,t,r,b" CSS lengths; defaults match the wps
 *  bodyPr defaults — both are Word's 0.1"/0.05"). */
function parseVmlTextbox(
  el: OoxmlNode,
  ctx: Ctx,
): {
  paragraphs: unknown[];
  inset?: { l: number; t: number; r: number; b: number };
} | null {
  const tb = child(el, 'v:textbox');
  const paragraphs = txbxParagraphs(child(tb, 'w:txbxContent'), ctx);
  if (!paragraphs) return null;
  const insetAttr = attrOf(tb, 'inset');
  if (insetAttr) {
    const p = insetAttr.split(',').map((s) => {
      const px = cssLenToPx(s.trim());
      return px === undefined ? undefined : Math.round(px);
    });
    return {
      paragraphs,
      inset: { l: p[0] ?? 10, t: p[1] ?? 5, r: p[2] ?? 10, b: p[3] ?? 5 },
    };
  }
  return { paragraphs };
}

/** A drawn legacy VML shape in a run's w:pict: v:roundrect / v:rect / v:oval
 *  boxes (with their v:textbox content) and straight connectors
 *  (o:connectortype, or a `type="#id"` reference to a v:shapetype with
 *  o:spt 32). Pictures took the parseVmlImage exit before this; unmapped
 *  VML stays unmodelled and visible in the XML audit. */
function parseVmlShape(run: OoxmlNode, ctx: Ctx): PMNode | null {
  const pict = child(run, 'w:pict');
  if (!pict) return null;
  // Filled by registerVmlShapeTypes for every pict, so a type declared in
  // one that turned out to hold a picture is still resolvable here.
  const types = ctx.vmlShapeTypes ?? new Map<string, number>();

  let el = child(pict, 'v:roundrect');
  let kind: ShapeSpec['kind'] | null = el ? 'roundRect' : null;
  if (!el) {
    el = child(pict, 'v:rect');
    if (el) kind = 'rect';
  }
  if (!el) {
    el = child(pict, 'v:oval');
    if (el) kind = 'ellipse';
  }
  if (!el) {
    el = child(pict, 'v:shape');
    if (el) {
      const ref = attrOf(el, 'type');
      const spt = ref?.startsWith('#') ? types.get(ref.slice(1)) : undefined;
      kind = attrOf(el, 'o:connectortype')
        ? 'line'
        : spt === 32
          ? 'line'
          : child(el, 'v:textbox') || spt === 202
            ? 'rect'
            : null;
    }
  }
  if (!el || !kind) return null;

  const st = parseVmlStyle(attrOf(el, 'style') ?? '');
  const textbox = kind === 'line' ? null : parseVmlTextbox(el, ctx);

  const shape: Record<string, unknown> = { kind };
  // VML defaults: stroked in black at 1px, boxes filled white — both until
  // the attribute says otherwise ("f").
  if (attrOf(el, 'stroked') !== 'f') {
    shape['stroke'] = vmlColor(attrOf(el, 'strokecolor')) ?? '#000000';
    const w = cssLenToPx(attrOf(el, 'strokeweight'));
    shape['strokeWidth'] = w ? Math.max(1, Math.round(w)) : 1;
  }
  if (kind !== 'line' && attrOf(el, 'filled') !== 'f')
    shape['fill'] = vmlColor(attrOf(el, 'fillcolor')) ?? '#ffffff';
  if (st.flipV) shape['flipV'] = true;
  if (kind === 'roundRect')
    shape['cornerRatio'] = vmlArcSize(attrOf(el, 'arcsize'));
  const strokeEl = child(el, 'v:stroke');
  const startArrow = attrOf(strokeEl, 'startarrow');
  const endArrow = attrOf(strokeEl, 'endarrow');
  if (startArrow && startArrow !== 'none') shape['arrowStart'] = true;
  if (endArrow && endArrow !== 'none') shape['arrowEnd'] = true;
  // Dashed connectors ("1 1" dotted, "dash", …) — anything but solid.
  const dashstyle = attrOf(strokeEl, 'dashstyle');
  if (dashstyle && dashstyle !== 'solid')
    shape['dash'] = vmlDashPattern(dashstyle);
  // endcap: flat (the spec default) needs no model field — the painter's
  // canvas default is already butt.
  const endcap = attrOf(strokeEl, 'endcap');
  if (endcap === 'round' || endcap === 'square') shape['cap'] = endcap;

  // v:textbox style layout-flow:vertical (+ mso-layout-flow-alt) — text runs
  // along the box's vertical axis. The painter rotates a float's ENTIRE
  // payload (shape + textbox lines), so model the box laid out SIDEWAYS
  // (w/h swapped around the same center) and rotated back into place;
  // bottom-to-top reads upward = counter-clockwise.
  let rotation = 0;
  const tbStyle = attrOf(child(el, 'v:textbox'), 'style');
  if (tbStyle && /layout-flow:vertical/.test(tbStyle)) {
    rotation = /mso-layout-flow-alt:bottom-to-top/.test(tbStyle) ? -90 : 90;
  }

  // position:absolute floats at margin-left/top from the anchor paragraph
  // (VML's default relative frame is the text column / the paragraph).
  const w = st.width ?? 0;
  const h = st.height ?? 0;
  // A rotated (vertical-text) box swaps its laid-out dimensions; the offsets
  // shift so the rotated box lands on the ORIGINAL rectangle (same center).
  const boxW = rotation ? h : w;
  const boxH = rotation ? w : h;
  const dx = rotation ? (w - h) / 2 : 0;
  const dy = rotation ? (h - w) / 2 : 0;

  let float: Record<string, unknown> | null = null;
  if (st.absolute) {
    float = { wrap: 'none', hRel: 'margin', vRel: 'paragraph' };
    if (st.left !== undefined) float['hOffset'] = Math.round(st.left + dx);
    if (st.top !== undefined) float['vOffset'] = Math.round(st.top + dy);
    if ((st.zIndex ?? 0) < 0) float['behind'] = true;
  }

  // NB: no blanket markSubtree here — properties the model does NOT honor
  // (v:shadow, o:extrusion, …) must stay visible in the XML audit;
  // everything consumed above was marked by reading it.
  audit.mark(pict);
  return ctx.schema.nodes['image'].create({
    src: '',
    width: Math.round(boxW),
    height: Math.round(boxH),
    // The shape's OWN alt text, or nothing. `id` used to stand in here, which
    // meant a screen reader announced "_x0000_s1026" — an identifier is not a
    // description, and an empty alt is the honest way to say a shape carries
    // no description of its own.
    alt: attrOf(el, 'alt') || '',
    float,
    shape,
    textbox,
    ...(rotation ? { rotation } : {}),
  });
}

// ── carry-through fidelity ──────────────────────────────────────────
// Body XML is REGENERATED on export, so any property the model doesn't
// represent would be lost the moment a customer saves. Unmodelled rPr/pPr
// children are therefore preserved verbatim (raw XML on a mark / paragraph
// attr) and spliced back by the exporter. See model.ts `carryRPr` / `carry`.

/** rPr children whose VALUE already lives in the model (re-emitted from
 *  marks on export) — carrying them too would duplicate or contradict. */
const CONSUMED_RPR = new Set([
  'w:b',
  'w:i',
  'w:u',
  'w:strike',
  'w:dstrike',
  'w:smallCaps',
  'w:color',
  'w:sz',
  'w:rFonts',
  'w:vertAlign',
  'w:position',
  // Character tracking. The pPr set below has its own 'w:spacing' entry — the
  // two elements share a name and nothing else.
  'w:spacing',
  'w:w',
  'w:highlight',
  'w:shd',
  // w:rStyle: resolved into the cascade. NOT carried on purpose — re-emitting
  // it beside flattened direct props would resurrect style formatting the
  // user removed (style bold + user unbolds → bold comes back). Known loss.
  'w:rStyle',
]);

/** Inline pPr children the model represents (or handles elsewhere). */
const CONSUMED_PPR = new Set([
  'w:pStyle',
  'w:numPr',
  'w:jc',
  'w:ind',
  'w:spacing',
  'w:tabs',
  'w:pBdr',
  'w:pageBreakBefore',
  'w:keepNext',
  'w:keepLines',
  'w:widowControl',
  'w:outlineLvl',
  'w:sectPr', // section breaks — parsed by parseBodyBlocks
  'w:rPr', // the paragraph mark's run props — carried separately (markRPr)
]);

/** tblPr children the model represents and re-emits. Everything else —
 *  including w:tblStyle — is carried: unlike w:rStyle (whose re-emit could
 *  resurrect character formatting the user removed via the always-visible
 *  toggle buttons), table styles are rarely edited away in-app, and dropping
 *  the link visibly strips banding/theme colors the moment a customer saves. */
const CONSUMED_TBLPR = new Set(['w:tblBorders', 'w:tblCellMar', 'w:jc']);

/** trPr children the model represents (header / height / cantSplit). */
const CONSUMED_TRPR = new Set(['w:tblHeader', 'w:trHeight', 'w:cantSplit']);

/** tcPr children the model represents. */
const CONSUMED_TCPR = new Set([
  'w:tcW',
  'w:gridSpan',
  'w:vMerge',
  'w:shd',
  'w:vAlign',
  'w:tcBorders',
  'w:tcMar',
]);

/** Property records of tracked changes: after the user edits, replaying them
 *  would be a lie. Never carried (the accept-all view dropped their runs). */
const CARRY_NEVER = new Set([
  'w:rPrChange',
  'w:pPrChange',
  'w:sectPrChange',
  'w:tblPrChange',
  'w:trPrChange',
  'w:tcPrChange',
  'w:tblGridChange',
]);

// Carried fragments are embedded into OUR generated document.xml, whose root
// declares a fixed namespace set — only elements/attrs that stay well-formed
// there may pass. (w14:/wp14: etc. would be undeclared → invalid XML.)
const CARRY_FILTER = {
  element: (name: string) => name.startsWith('w:') && !CARRY_NEVER.has(name),
  attr: (name: string) =>
    name.startsWith('w:') || name === 'xml:space' || !name.includes(':'),
};

/** The unmodelled children of a property bag as one XML string (original
 *  order), or null. Carried nodes count as consumed for the audit — they are
 *  preserved, not lost; remaining RENDER gaps are tracked in the plan. */
function collectCarry(
  bag: OoxmlNode | undefined,
  consumed: Set<string>,
): string | null {
  if (!bag) return null;
  let out = '';
  for (const c of bag.children) {
    if (consumed.has(c.name) || !CARRY_FILTER.element(c.name)) continue;
    const xml = serializeOoxml(c, CARRY_FILTER);
    if (xml) {
      out += xml;
      audit.markSubtree(c);
    }
  }
  return out || null;
}

/** Effective marks for a run (docDefaults+paraStyle → run style → inline rPr). */
function runMarks(
  run: OoxmlNode | undefined,
  paraBase: RunProps,
  ctx: Ctx,
  href: string | null,
) {
  const rPr = child(run, 'w:rPr');
  // No w:rStyle still means the default character style ("Default Paragraph
  // Font"), which is empty in a stock document but need not be.
  const rStyleId =
    attrOf(child(rPr, 'w:rStyle'), 'w:val') ??
    ctx.styles.defaultStyleIdFor('character');
  const effective = [
    paraBase,
    ctx.styles.resolveStyle(rStyleId),
    parseRunProps(rPr, ctx.resolveTheme, ctx.resolveFont),
  ].reduce(mergeRunProps, {} as RunProps);
  const marks = propsToMarks(effective, ctx);
  // Unmodelled INLINE rPr children (w:rtl, w:kern, w:szCs, …) ride a carry
  // mark so a save keeps them. Style-layer properties stay in styles.xml.
  const carryXml = collectCarry(rPr, CONSUMED_RPR);
  if (carryXml && ctx.schema.marks['carryRPr']) {
    marks.push(ctx.schema.marks['carryRPr'].create({ xml: carryXml }));
  }
  if (href) marks.push(ctx.schema.marks['link'].create({ href }));
  // Comments active over this run (w:commentRangeStart/End) become a comment
  // mark — only when the schema in use actually has it (the comment plugin
  // contributed it; otherwise comment ranges are silently skipped).
  if (ctx.comments.active.size > 0 && ctx.schema.marks['comment']) {
    const ids = [...ctx.comments.active].sort((a, b) => a - b);
    marks.push(ctx.schema.marks['comment'].create({ ids }));
  }
  return marks;
}

/** Map one w:r into inline nodes (image, hard breaks or marked text). */
function runToInline(
  run: OoxmlNode,
  paraBase: RunProps,
  ctx: Ctx,
  href: string | null,
): PMNode[] {
  const marks = runMarks(run, paraBase, ctx, href);
  const nodes = runInlineNodes(run, marks, ctx);
  // A linked image keeps its link mark; image nodes carry no marks otherwise.
  return href
    ? nodes.map((n) => (n.type.name === 'image' ? n.mark(marks) : n))
    : nodes;
}

/** PAGE / NUMPAGES from a field instruction, or null for any other field. */
function fieldKind(instr: string): 'page' | 'pages' | null {
  if (/\bNUMPAGES\b/.test(instr)) return 'pages';
  if (/\bPAGE\b/.test(instr)) return 'page';
  return null;
}

/** Kind for a field whose result SPANS paragraphs. Only the table of contents
 *  is modelled by name (its entries are regenerated on update); anything else
 *  is still generated content and shades the same, just anonymously. */
function fieldSpanKind(instr: string): string {
  return /^\s*TOC\b/i.test(instr) ? 'toc' : 'field';
}

/** A page_field node formatted like `formatRun` (the field's result run). */
function pageFieldNode(
  kind: 'page' | 'pages',
  formatRun: OoxmlNode | undefined,
  paraBase: RunProps,
  ctx: Ctx,
): PMNode {
  return ctx.schema.nodes['page_field']
    .create({ kind })
    .mark(runMarks(formatRun, paraBase, ctx, null));
}

/** State for a complex field (w:fldChar begin … instrText … separate … end).
 *  Fields NEST (a TOC field wraps a PAGEREF per entry) — `parent` is the
 *  enclosing open field, and result content renders eagerly into `result`
 *  so an inner field's output lands inside the outer field's result. */
interface FieldState {
  instr: string;
  /** Rendered result content (cached field result, kept as-is). */
  result: PMNode[];
  /** First result run's XML — formatting source for PAGE/NUMPAGES nodes. */
  firstResultRun?: OoxmlNode;
  phase: 'instr' | 'result';
  parent: FieldState | null;
}

/** Heading level (1–6) for a paragraph, from its style id ("Heading1"…) or an
 *  explicit `w:outlineLvl` (0-based) in the pPr cascade; undefined for body. */
function headingLevel(
  pStyleId: string | undefined,
  pPrChain: (OoxmlNode | undefined)[],
): number | undefined {
  if (pStyleId) {
    const m = /^heading\s*([1-9])$/i.exec(pStyleId);
    if (m) {
      // The style name settled it, so the chain's w:outlineLvl is never
      // consulted — but Word bakes the matching level into every built-in
      // HeadingN style, so it agrees and was handled, not missed.
      if (audit.collecting)
        markOverridden(pPrChain, pPrChain.length, 'w:outlineLvl');
      return Math.min(6, Number(m[1]));
    }
  }
  const ol = lastWith(pPrChain, 'w:outlineLvl');
  if (ol) {
    const v = Number(attrOf(ol, 'w:val'));
    if (!Number.isNaN(v) && v >= 0 && v <= 8) return Math.min(6, v + 1);
  }
  return undefined;
}

const SUB_DIGITS = '₀₁₂₃₄₅₆₇₈₉';
const SUP_DIGITS = '⁰¹²³⁴⁵⁶⁷⁸⁹';

/** Digits mapped through a Unicode sub/superscript alphabet, or null when the
 *  text isn't digits-only (falls back to `_(…)` / `^(…)` linear format). */
function scriptDigits(text: string, alphabet: string): string | null {
  if (!/^[0-9]+$/.test(text)) return null;
  return [...text].map((d) => alphabet[Number(d)]).join('');
}

/** OMML (`m:oMath`) flattened to readable plain text — v1 keeps the equation's
 *  CONTENT, not its typesetting: `t` sub `1` → "t₁", `x` sup `2` → "x²",
 *  fractions → "num/den", radicals → "√(…)", delimiters → "(…)". Unknown
 *  constructs concatenate their children's text so nothing is dropped. */
function flattenOmml(node: OoxmlNode): string {
  const flat = (n: OoxmlNode | undefined): string => (n ? flattenOmml(n) : '');
  switch (node.name) {
    case 'm:t':
      return node.text;
    case 'm:f': {
      // Multi-term sides get parens so "t₁+t₂+t₃ over 3" doesn't flatten to
      // the ambiguous "t₁+t₂+t₃/3".
      const side = (s: string): string => (/[+\-±×÷/ ]/.test(s) ? `(${s})` : s);
      return `${side(flat(child(node, 'm:num')))}/${side(flat(child(node, 'm:den')))}`;
    }
    case 'm:sSub': {
      const sub = flat(child(node, 'm:sub'));
      return (
        flat(child(node, 'm:e')) +
        (scriptDigits(sub, SUB_DIGITS) ?? `_(${sub})`)
      );
    }
    case 'm:sSup': {
      const sup = flat(child(node, 'm:sup'));
      return (
        flat(child(node, 'm:e')) +
        (scriptDigits(sup, SUP_DIGITS) ?? `^(${sup})`)
      );
    }
    case 'm:sSubSup': {
      const sub = flat(child(node, 'm:sub'));
      const sup = flat(child(node, 'm:sup'));
      return (
        flat(child(node, 'm:e')) +
        (scriptDigits(sub, SUB_DIGITS) ?? `_(${sub})`) +
        (scriptDigits(sup, SUP_DIGITS) ?? `^(${sup})`)
      );
    }
    case 'm:rad': {
      const deg = flat(child(node, 'm:deg'));
      return `${deg}√(${flat(child(node, 'm:e'))})`;
    }
    case 'm:d': {
      // Delimiters: explicit m:begChr/m:endChr/m:sepChr in m:dPr, parens/comma
      // by default (empty w:val means "none").
      const pr = child(node, 'm:dPr');
      const chr = (name: string, dflt: string): string =>
        attrOf(child(pr, name), 'm:val') ?? dflt;
      const args = children(node, 'm:e').map(flat);
      return (
        chr('m:begChr', '(') +
        args.join(chr('m:sepChr', ',')) +
        chr('m:endChr', ')')
      );
    }
    default:
      // Property containers hold formatting (and ctrlPr), never content.
      if (node.name.startsWith('m:') && node.name.endsWith('Pr')) return '';
      return node.children.map(flat).join('');
  }
}

// Paragraph children the loop in parseParagraph consumes (w:pPr is marked by
// the child() accessor; OMML branches are subtree-marked where flattened).
const PARA_CHILD_TAGS = new Set([
  'w:r',
  'w:fldSimple',
  'w:hyperlink',
  'w:commentRangeStart',
  'w:commentRangeEnd',
]);

function parseParagraph(p: OoxmlNode, ctx: Ctx): PMNode {
  const pPr = child(p, 'w:pPr');
  const pStyleId = attrOf(child(pPr, 'w:pStyle'), 'w:val');
  // Unstyled content still gets Word's DEFAULT paragraph style ("Normal"),
  // which is where a document keeps its line spacing and space-after. A
  // paragraph that names its own style does not fall back to Normal — the
  // style's basedOn chain decides what it inherits, as in Word.
  const styleId = pStyleId ?? ctx.styles.defaultStyleIdFor('paragraph');
  // Base for every run: docDefaults → paragraph style's run properties.
  const paraBase = mergeRunProps(
    ctx.styles.docDefaults,
    ctx.styles.resolveStyle(styleId),
  );
  // Paragraph-property cascade, base-most first; later layers win:
  // docDefaults pPrDefault → style chain (w:basedOn ancestors → style)
  // → numbering lvl pPr (the per-level list indent) → inline.
  const pPrChain: (OoxmlNode | undefined)[] = [
    ctx.styles.docDefaultsPPr,
    ...ctx.styles.resolveStylePPr(styleId),
    pPr,
  ];
  const parsedList = parseList(lastWith(pPrChain, 'w:numPr'));
  let list: ListInfo | null = null;
  if (parsedList) {
    const { explicitLevel, ...info } = parsedList;
    list = info;
    // Numbered heading styles: a lvl carrying w:pStyle claims paragraphs of
    // that style for ITS level. Only a written w:ilvl (direct intent) wins
    // over the link — the style-chain numPr usually has none.
    if (!explicitLevel && pStyleId !== undefined) {
      const linked = ctx.numbering.levelForStyle(info.numId, pStyleId);
      if (linked !== undefined) list.level = linked;
    }
    const lvlPPr = ctx.numbering.levelPPr(list.numId, list.level);
    if (lvlPPr) pPrChain.splice(pPrChain.length - 1, 0, lvlPPr);
  }
  const align = resolveAlign(pPrChain);
  const indent = resolveIndent(pPrChain);
  const spacing = resolveSpacing(pPrChain, list !== null);
  const tabs = resolveTabs(pPrChain);
  const heading = headingLevel(pStyleId, pPrChain);

  const inline: PMNode[] = [];
  let bookmarks: string[] | null = null;
  // The generated field this paragraph sits in: one already open when it
  // starts, else one it opens itself (resolved after the run loop).
  const fieldAtStart = ctx.openFields[0] ?? null;
  let field: FieldState | null = null;
  // Toggle-valued: w:pageBreakBefore w:val="false" (an override cancelling an
  // inherited break) must count as OFF — presence alone read it backwards.
  const pbLayer = lastWith(pPrChain, 'w:pageBreakBefore');
  let pageBreak = pbLayer
    ? isToggleOn(child(pbLayer, 'w:pageBreakBefore'))
    : false;
  // Pagination keeps, same toggle semantics. widowControl defaults ON in
  // Word — only an explicit off is worth an attr.
  const toggleLayer = (name: string, dflt: boolean): boolean => {
    const layer = lastWith(pPrChain, name);
    return layer ? isToggleOn(child(layer, name)) : dflt;
  };
  const keepNext = toggleLayer('w:keepNext', false);
  const keepLines = toggleLayer('w:keepLines', false);
  const widowControl = toggleLayer('w:widowControl', true);
  // Where rendered inline content lands: the open field's result if we're
  // past its separate mark, the paragraph otherwise.
  const sink = (): PMNode[] =>
    field && field.phase === 'result' ? field.result : inline;
  const emitRun = (run: OoxmlNode, href: string | null): void => {
    if (field && field.phase === 'result' && !field.firstResultRun)
      field.firstResultRun = run;
    sink().push(...runToInline(run, paraBase, ctx, href));
  };

  // One run through the field state machine. `href` marks runs living inside
  // a w:hyperlink wrapper — TOC entries put whole PAGEREF fields there, so
  // the machine must run for them too (they used to bypass it entirely).
  const handleRun = (node: OoxmlNode, href: string | null): void => {
    if (hasPageBreak(node)) pageBreak = true;
    const fldChars = children(node, 'w:fldChar');
    if (fldChars.length === 0) {
      if (field && field.phase === 'instr') {
        field.instr += children(node, 'w:instrText')
          .map((t) => t.text)
          .join('');
        // Anything else in an instruction-side run is field plumbing we
        // drop as Word does — consumed by design, not a coverage gap.
        audit.markSubtree(node);
        return;
      }
      emitRun(node, href);
      return;
    }
    // The run carries fldChar(s). Some producers (Google Docs) pack
    // begin + instrText + separate + end into a SINGLE run — treating the
    // run as one fldChar (the old shape) left the field open forever and
    // swallowed the rest of the paragraph. Walk the run's children in
    // order through the same state machine instead.
    const rPr = child(node, 'w:rPr');
    // A marker-only run's rPr formats the invisible field mark — dropped by
    // design (any actual content is re-read through the synth run below).
    if (rPr) audit.markSubtree(rPr);
    let plain: OoxmlNode[] = [];
    const flushPlain = () => {
      if (plain.length === 0) return;
      const synth: OoxmlNode = {
        name: 'w:r',
        attrs: node.attrs,
        children: rPr ? [rPr, ...plain] : plain,
        text: '',
      };
      // phase 'instr': stray content between begin and separate is
      // instruction-side noise — dropped, as Word does.
      if (!field || field.phase === 'result') emitRun(synth, href);
      plain = [];
    };
    for (const c of node.children) {
      if (c.name === 'w:fldChar' || c.name === 'w:instrText') audit.mark(c);
      if (c.name === 'w:fldChar') {
        const t = attrOf(c, 'w:fldCharType');
        if (t === 'begin') {
          flushPlain();
          field = { instr: '', result: [], phase: 'instr', parent: field };
        } else if (t === 'separate') {
          if (field) field.phase = 'result';
        } else if (t === 'end') {
          flushPlain(); // result text inside this run, before the end mark
          if (field) {
            const done = field;
            field = done.parent;
            const kind = fieldKind(done.instr);
            // PAGE/NUMPAGES: the cached result is recomputed by our layout —
            // only the first result run's formatting is read; everything
            // else keeps the cached rendering.
            sink().push(
              ...(kind
                ? [pageFieldNode(kind, done.firstResultRun, paraBase, ctx)]
                : done.result),
            );
          } else if (ctx.openFields.length > 0) {
            // Closes a field that OPENED in an earlier paragraph (a TOC ends
            // in its last entry) — the span stops with this paragraph.
            ctx.openFields.pop();
          }
        }
      } else if (c.name === 'w:instrText') {
        if (field && field.phase === 'instr') field.instr += c.text;
      } else if (c.name !== 'w:rPr') {
        plain.push(c);
      }
    }
    flushPlain();
  };

  for (const node of effectiveChildren(p.children)) {
    if (PARA_CHILD_TAGS.has(node.name)) audit.mark(node);
    if (node.name === 'w:r') {
      handleRun(node, null);
    } else if (node.name === 'w:fldSimple') {
      const kind = fieldKind(attrOf(node, 'w:instr') ?? '');
      const resultRuns = children(node, 'w:r');
      if (kind) {
        // Same as the fldChar path: cached PAGE/NUMPAGES text is recomputed.
        for (const r of resultRuns) audit.markSubtree(r);
        inline.push(pageFieldNode(kind, resultRuns[0], paraBase, ctx));
      } else {
        for (const r of resultRuns)
          inline.push(...runToInline(r, paraBase, ctx, null));
      }
    } else if (node.name === 'w:hyperlink') {
      const rel = attrOf(node, 'r:id')
        ? ctx.rels.get(attrOf(node, 'r:id') as string)
        : undefined;
      const anchor = attrOf(node, 'w:anchor');
      const href = rel?.target ?? (anchor ? `#${anchor}` : null);
      for (const run of children(node, 'w:r')) {
        handleRun(run, href);
      }
    } else if (node.name === 'm:oMath' || node.name === 'm:oMathPara') {
      // OMML equations, flattened to a plain-text run (v1: content over
      // typesetting — see flattenOmml). Formatted like the first math run.
      // Deliberate wholesale flattening — the whole subtree counts consumed.
      audit.markSubtree(node);
      const text = flattenOmml(node);
      if (text.length > 0) {
        const first = findDescendant(node, 'm:r');
        inline.push(
          ctx.schema.text(text, runMarks(first, paraBase, ctx, null)),
        );
      }
    } else if (node.name === 'w:commentRangeStart') {
      const id = Number(attrOf(node, 'w:id'));
      if (!Number.isNaN(id) && ctx.comments.defs.has(id)) {
        ctx.comments.active.add(id);
        if (!ctx.comments.used.includes(id)) ctx.comments.used.push(id);
      }
    } else if (node.name === 'w:commentRangeEnd') {
      const id = Number(attrOf(node, 'w:id'));
      if (!Number.isNaN(id)) ctx.comments.active.delete(id);
    } else if (node.name === 'w:bookmarkStart') {
      // Named anchor: link hrefs "#name" resolve to this paragraph. Word's
      // own cursor bookmark is noise.
      audit.mark(node);
      const name = attrOf(node, 'w:name');
      if (name && name !== '_GoBack') (bookmarks ??= []).push(name);
    }
  }
  // A field still open here spans paragraphs (the TOC field wraps ALL its
  // entry paragraphs; each parseParagraph sees only its slice) — keep the
  // cached result content, and register the span so the FOLLOWING paragraphs
  // know they are inside it too. Registered outermost-first (the loop unwinds
  // innermost-first) to match the stack's own order.
  const spans: FieldInfo[] = [];
  while (field) {
    const done: FieldState = field;
    field = done.parent;
    sink().push(...done.result);
    spans.unshift({
      kind: fieldSpanKind(done.instr),
      instr: done.instr.trim(),
    });
  }
  ctx.openFields.push(...spans);
  // w:pBdr from the most-derived cascade layer: keep visible sides only
  // ('between' isn't modelled; w:val="none" sides drop out).
  const pBdrLayer = lastWith(pPrChain, 'w:pBdr');
  const pBdrSides = pBdrLayer
    ? parseBordersEl(child(pBdrLayer, 'w:pBdr'), CELL_SIDES)
    : null;
  const paraBorders: Record<string, BorderSide> = {};
  if (pBdrSides) {
    for (const side of CELL_SIDES) {
      const s = pBdrSides[side];
      if (s) paraBorders[side] = s;
    }
  }

  const attrs: {
    list?: ListInfo;
    align?: Align;
    indent?: Indent;
    spacing?: Spacing;
    tabs?: { pos: number; val: string; leader?: string }[];
    bookmarks?: string[];
    field?: FieldInfo;
    pageBreakBefore?: boolean;
    keepNext?: boolean;
    keepLines?: boolean;
    widowControl?: boolean;
    heading?: number;
    styleId?: string;
    borders?: Record<string, BorderSide>;
    carry?: { pPr?: string; markRPr?: string };
  } = {};
  // Carry-through: unmodelled INLINE pPr children + the paragraph mark's
  // w:rPr, preserved verbatim for export (see collectCarry).
  const carryPPr = collectCarry(pPr, CONSUMED_PPR);
  const carryMarkRPr = collectCarry(child(pPr, 'w:rPr'), new Set());
  if (carryPPr || carryMarkRPr) {
    attrs.carry = {
      ...(carryPPr && { pPr: carryPPr }),
      ...(carryMarkRPr && { markRPr: carryMarkRPr }),
    };
  }
  if (list) attrs.list = list;
  if (align) attrs.align = align;
  if (heading) attrs.heading = heading;
  // Title/Subtitle: named styles without an outline level. Only when the
  // paragraph isn't already a heading (invariant: styleId ⇒ heading null).
  else if (pStyleId && /^(title|subtitle)$/i.test(pStyleId)) {
    attrs.styleId = pStyleId.toLowerCase() === 'title' ? 'Title' : 'Subtitle';
  }
  if (indent) attrs.indent = indent;
  if (spacing) attrs.spacing = spacing;
  if (tabs) attrs.tabs = tabs;
  if (bookmarks) attrs.bookmarks = bookmarks;
  // A paragraph that OPENS a spanning field belongs to it as much as the ones
  // that follow (the TOC's first entry is part of the TOC).
  const fieldForPara = fieldAtStart ?? ctx.openFields[0] ?? null;
  if (fieldForPara) attrs.field = fieldForPara;
  if (pageBreak) attrs.pageBreakBefore = true;
  if (keepNext) attrs.keepNext = true;
  if (keepLines) attrs.keepLines = true;
  if (!widowControl) attrs.widowControl = false;
  if (Object.keys(paraBorders).length > 0) attrs.borders = paraBorders;
  return ctx.schema.nodes['paragraph'].create(attrs, inline);
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
      val:
        val === 'right' || val === 'center' || val === 'decimal' ? val : 'left',
    };
    const leader = attrOf(tab, 'w:leader');
    if (leader && leader !== 'none') {
      stop.leader =
        leader === 'hyphen'
          ? 'hyphen'
          : leader === 'underscore'
            ? 'underscore'
            : leader === 'middleDot'
              ? 'middleDot'
              : 'dot';
    }
    stops.push(stop);
  }
  return stops.length > 0 ? stops.sort((a, b) => a.pos - b.pos) : null;
}

/** Mark every `childName` in the layers BELOW `winner` as handled.
 *
 *  A cascade resolves by walking back from the most-derived layer and taking
 *  the first hit, so the layers under it are never touched — and to the audit
 *  an untouched node is indistinguishable from an unsupported one. But being
 *  overridden is the correct outcome, not a gap: eight Heading3 paragraphs
 *  carrying an inline `w:tabs` made the style's own `w:tabs` look unread,
 *  when the renderer had honoured the cascade exactly right.
 *
 *  Pass `chain.length` for `winner` when something OUTSIDE the chain settled
 *  the property (a style name deciding a heading level). */
function markOverridden(
  chain: (OoxmlNode | undefined)[],
  winner: number,
  childName: string,
): void {
  for (let i = winner - 1; i >= 0; i--) {
    const el = child(chain[i], childName);
    if (el) audit.markSubtree(el);
  }
}

/** The last (most-derived) pPr layer that carries `childName`, if any. */
function lastWith(
  chain: (OoxmlNode | undefined)[],
  childName: string,
): OoxmlNode | undefined {
  for (let i = chain.length - 1; i >= 0; i--) {
    if (child(chain[i], childName)) {
      if (audit.collecting) markOverridden(chain, i, childName);
      return chain[i];
    }
  }
  return undefined;
}

/** Resolve alignment through the cascade: the last layer with a w:jc wins.
 *  A layer whose w:val we don't map (thaiDistribute…) falls through to the
 *  one below, so the winner is not simply the last layer holding a w:jc. */
function resolveAlign(chain: (OoxmlNode | undefined)[]): Align | null {
  for (let i = chain.length - 1; i >= 0; i--) {
    const align = parseAlign(chain[i]);
    if (!align) continue;
    if (audit.collecting) markOverridden(chain, i, 'w:jc');
    return align;
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

/** Resolve w:spacing through the cascade: the last layer with each attribute
 *  wins. before/after are twips→px; line is 240ths of a line for lineRule
 *  'auto' (→ multiplier), else twips→px for 'exact'/'atLeast'. */
function resolveSpacing(
  chain: (OoxmlNode | undefined)[],
  isList = false,
): Spacing | null {
  const out: Spacing = {};
  for (const pPr of chain) {
    const sp = child(pPr, 'w:spacing');
    if (!sp) continue;
    const before = attrOf(sp, 'w:before');
    const after = attrOf(sp, 'w:after');
    const line = attrOf(sp, 'w:line');
    const rule = attrOf(sp, 'w:lineRule');
    // HTML-era auto spacing: the literal value is Word's own cached auto
    // amount, so it stands — except list items, which drop auto spacing.
    const onFlag = (name: string): boolean => {
      const v = attrOf(sp, name);
      return v !== undefined && v !== '0' && v.toLowerCase() !== 'false';
    };
    const beforeAuto = onFlag('w:beforeAutospacing');
    const afterAuto = onFlag('w:afterAutospacing');
    if (before !== undefined || beforeAuto)
      out.before = beforeAuto && isList ? 0 : twipsToPx(Number(before ?? '0'));
    if (after !== undefined || afterAuto)
      out.after = afterAuto && isList ? 0 : twipsToPx(Number(after ?? '0'));
    if (line !== undefined) {
      const lineRule = rule === 'exact' || rule === 'atLeast' ? rule : 'auto';
      out.lineRule = lineRule;
      out.line =
        lineRule === 'auto' ? Number(line) / 240 : twipsToPx(Number(line));
    }
  }
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

/** Read a paragraph's list membership (w:numPr). The marker string is NOT
 *  resolved here — the layout engine recounts markers from the doc's
 *  numbering defs every pass, so edits renumber live. */
function parseList(
  pPr: OoxmlNode | undefined,
): (ListInfo & { explicitLevel: boolean }) | null {
  const numPr = child(pPr, 'w:numPr');
  const numId = attrOf(child(numPr, 'w:numId'), 'w:val');
  if (numId === undefined || numId === '0') return null; // 0 cancels numbering
  const ilvlAttr = attrOf(child(numPr, 'w:ilvl'), 'w:val');
  const ilvl = Number(ilvlAttr ?? '0');
  return {
    numId,
    level: Number.isNaN(ilvl) ? 0 : ilvl,
    // A written w:ilvl is direct intent; its absence leaves room for the
    // lvl w:pStyle link to pick the level (numbered heading styles).
    explicitLevel: ilvlAttr !== undefined,
  };
}

interface LogicalCell {
  startCol: number; // grid column this cell starts at
  colspan: number;
  vMerge: 'restart' | 'continue' | null;
  colwidth: number[] | null;
  background: string | null;
  vAlign: string | null;
  borders: TableBorders | null;
  padding: {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
  } | null;
  /** Unmodelled tcPr children, carried verbatim (see collectCarry). */
  carry: string | null;
  content: PMNode[];
}

function emptyCell(ctx: Ctx): PMNode {
  return ctx.schema.nodes['table_cell'].create(null, [
    ctx.schema.nodes['paragraph'].create(),
  ]);
}

/** Four-side margins (px) from a w:tblCellMar / w:tcMar element, or null.
 *  type="nil" forces 0; only "dxa" widths are interpreted. */
function parseMarginsEl(
  mar: OoxmlNode | undefined,
): { left?: number; right?: number; top?: number; bottom?: number } | null {
  if (!mar) return null;
  const side = (name: string): number | undefined => {
    const el = child(mar, name);
    if (!el) return undefined;
    if (attrOf(el, 'w:type') === 'nil') return 0;
    const w = attrOf(el, 'w:w');
    return w === undefined ? undefined : twipsToPx(Number(w));
  };
  const out: { left?: number; right?: number; top?: number; bottom?: number } =
    {};
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

/** w:tblPr/w:tblCellMar overrides (px), or null for Word defaults. */
function parseCellMargins(
  tbl: OoxmlNode,
  ctx: Ctx,
): { left?: number; right?: number; top?: number; bottom?: number } | null {
  const tblPr = child(tbl, 'w:tblPr');
  const inline = parseMarginsEl(child(tblPr, 'w:tblCellMar'));
  if (inline) return inline;
  // No inline margins: the table STYLE's w:tblCellMar applies (Word default
  // styles carry 108-twip side margins there — dropping them cramped text
  // against the cell borders).
  const styleId = attrOf(child(tblPr, 'w:tblStyle'), 'w:val');
  return parseMarginsEl(ctx.styles.resolveTableCellMar(styleId));
}

/** OOXML w:val border styles → our {@link BorderStyle} (unknowns → solid). */
const BORDER_STYLE_IN: Record<string, BorderStyle> = {
  single: 'solid',
  thick: 'solid',
  dashed: 'dashed',
  dashSmallGap: 'dashed',
  dotted: 'dotted',
  dotDash: 'dashed',
  dotDotDash: 'dashed',
  double: 'double',
};

/** A side element (w:top/bottom/…) → its appearance, or `false` when hidden.
 *  w:sz is eighths of a point; w:color "auto"/absent keeps the default grey. */
function parseBorderSide(el: OoxmlNode): BorderSide | false {
  const val = attrOf(el, 'w:val');
  if (val === 'none' || val === 'nil') {
    // Hidden side: its sz/color/space are meaningless, but "ask" them so the
    // coverage audit doesn't flag decoration attrs on borders we DID handle.
    attrOf(el, 'w:sz');
    attrOf(el, 'w:color');
    attrOf(el, 'w:space');
    return false;
  }
  const sz = Number(attrOf(el, 'w:sz') ?? '4');
  const width = Math.max(0.75, (sz / 8) * (96 / 72));
  const style = BORDER_STYLE_IN[val ?? 'single'] ?? 'solid';
  const colorAttr = attrOf(el, 'w:color');
  const color =
    colorAttr && colorAttr !== 'auto'
      ? (normalizeHex(colorAttr) ?? '#b0b0b0')
      : '#b0b0b0';
  const side: BorderSide = { width, style, color };
  // w:space (points) — border-to-content gap; kept for round-trip fidelity.
  const sp = Number(attrOf(el, 'w:space') ?? '0');
  if (Number.isFinite(sp) && sp > 0) side.space = Math.round(sp * (96 / 72));
  return side;
}

/** Per-side border appearance from a w:tblBorders / w:tcBorders node. An absent
 *  side is omitted (inherits); a present side is a {@link BorderSide} or false. */
function parseBordersEl(
  bordersEl: OoxmlNode | undefined,
  sides: readonly string[],
): TableBorders | null {
  if (!bordersEl) return null;
  const out: TableBorders = {};
  for (const side of sides) {
    const el = child(bordersEl, `w:${side}`);
    if (!el) continue;
    out[side as keyof TableBorders] = parseBorderSide(el);
  }
  return Object.keys(out).length > 0 ? out : null;
}

const TABLE_SIDES = [
  'top',
  'bottom',
  'left',
  'right',
  'insideH',
  'insideV',
] as const;
const CELL_SIDES = ['top', 'bottom', 'left', 'right'] as const;

/** Border visibility from a w:tblBorders node (direct tblPr or table style).
 *  OOXML tables are borderless unless declared; val none/nil hides a side. */
function parseTableBorders(tbl: OoxmlNode, ctx: Ctx): TableBorders | null {
  const tblPr = child(tbl, 'w:tblPr');
  const styleId = attrOf(child(tblPr, 'w:tblStyle'), 'w:val');
  const bordersEl =
    child(tblPr, 'w:tblBorders') ?? ctx.styles.resolveTableBorders(styleId);
  const out = parseBordersEl(bordersEl, TABLE_SIDES);
  // Only treat the table as bordered if at least one side is actually visible.
  return out && Object.values(out).some(Boolean) ? out : null;
}

/** A pct-typed width (w:tblW / w:tcW w:type="pct") as a percentage, or null.
 *  The value is either literal ("30%") or OOXML's 50ths-of-a-percent. */
function pctWidth(el: OoxmlNode | undefined): number | null {
  if (!el || attrOf(el, 'w:type') !== 'pct') return null;
  const raw = String(attrOf(el, 'w:w') ?? '').trim();
  const n = raw.endsWith('%') ? Number(raw.slice(0, -1)) : Number(raw) / 50;
  return Number.isFinite(n) && n > 0 ? n : null;
}

/** Effective per-grid-column pixel widths. Percentage-sized tables resolve
 *  against the page content width — their w:tblGrid is often a placeholder
 *  (e.g. 100 twips per column ≈ 7px, which would stack one character per
 *  line) — while ordinary tables keep the twips grid. */
function tableColumnWidths(tbl: OoxmlNode, grid: number[], ctx: Ctx): number[] {
  const tablePct = pctWidth(child(child(tbl, 'w:tblPr'), 'w:tblW'));
  const firstRow = children(tbl, 'w:tr')[0];
  const cells = firstRow
    ? children(firstRow, 'w:tc').map((tc) => {
        const tcPr = child(tc, 'w:tcPr');
        const tcW = child(tcPr, 'w:tcW');
        return {
          pct: pctWidth(tcW),
          // dxa width — the fallback grid when w:tblGrid is missing or
          // degenerate (some producers write no grid at all).
          dxa:
            attrOf(tcW, 'w:type') !== 'pct'
              ? Number(attrOf(tcW, 'w:w') ?? '0') || 0
              : 0,
          span: Number(attrOf(child(tcPr, 'w:gridSpan'), 'w:val') ?? '1') || 1,
        };
      })
    : [];
  if (!grid.some((w) => w > 0) && cells.some((c) => c.dxa > 0)) {
    grid = cells.flatMap((c) =>
      new Array<number>(c.span).fill(Math.round(c.dxa / c.span)),
    );
  }
  const spanSum = cells.reduce((s, c) => s + c.span, 0);
  const cols = grid.length || spanSum;
  if (tablePct === null && !cells.some((c) => c.pct !== null)) {
    return grid.map(twipsToPx);
  }
  const total = (ctx.contentWidth * (tablePct ?? 100)) / 100;
  const px = new Array<number>(cols).fill(Math.round(total / (cols || 1)));
  if (cells.length && cells.every((c) => c.pct !== null) && spanSum === cols) {
    let i = 0;
    for (const c of cells) {
      const w = Math.round((total * (c.pct as number)) / 100 / c.span);
      for (let k = 0; k < c.span && i < cols; k++) px[i++] = w;
    }
  } else if (grid.length) {
    // Pct table without full cell pcts: keep the grid's PROPORTIONS, scaled
    // to the pct width.
    const gsum = grid.reduce((a, b) => a + b, 0);
    if (gsum > 0)
      for (let i = 0; i < cols; i++)
        px[i] = Math.round((grid[i] / gsum) * total);
  }
  return px;
}

function parseTable(tbl: OoxmlNode, ctx: Ctx): PMNode {
  const grid = children(child(tbl, 'w:tblGrid'), 'w:gridCol').map((c) =>
    Number(attrOf(c, 'w:w') ?? '0'),
  );
  const colPx = tableColumnWidths(tbl, grid, ctx);

  // Phase 1: logical grid — every w:tc (incl. vMerge-continue placeholders),
  // tracking each cell's starting grid column.
  const logicalRows: LogicalCell[][] = children(tbl, 'w:tr').map((tr) => {
    const cells: LogicalCell[] = [];
    let col = 0;
    for (const tc of children(tr, 'w:tc')) {
      const tcPr = child(tc, 'w:tcPr');
      const colspan =
        Number(attrOf(child(tcPr, 'w:gridSpan'), 'w:val') ?? '1') || 1;
      const vMergeEl = child(tcPr, 'w:vMerge');
      const vMerge = !vMergeEl
        ? null
        : attrOf(vMergeEl, 'w:val') === 'restart'
          ? 'restart'
          : 'continue'; // omitted w:val defaults to continue
      const widths = colPx.length ? colPx.slice(col, col + colspan) : [];
      const background =
        shdFill(child(tcPr, 'w:shd'), ctx.resolveTheme) ?? null;
      const vAlignVal = attrOf(child(tcPr, 'w:vAlign'), 'w:val');
      const vAlign =
        vAlignVal === 'center' || vAlignVal === 'bottom' ? vAlignVal : null;
      const borders = parseBordersEl(child(tcPr, 'w:tcBorders'), CELL_SIDES);
      const padding = parseMarginsEl(child(tcPr, 'w:tcMar'));
      const carry = collectCarry(tcPr, CONSUMED_TCPR);
      const content = parseBlocks(tc, ctx);
      if (content.length === 0)
        content.push(ctx.schema.nodes['paragraph'].create());
      cells.push({
        startCol: col,
        colspan,
        vMerge,
        colwidth: widths.length ? widths : null,
        background,
        vAlign,
        borders,
        padding,
        carry,
        content,
      });
      col += colspan;
    }
    return cells;
  });

  // Per-row trPr: w:tblHeader (on/off) + w:trHeight (px floor / exact).
  const rowProps = children(tbl, 'w:tr').map((tr) => {
    const trPr = child(tr, 'w:trPr');
    // isToggleOn: also treats w:val="off" as off (the ad-hoc checks missed it).
    const header = isToggleOn(child(trPr, 'w:tblHeader'));
    const trH = child(trPr, 'w:trHeight');
    const hv = attrOf(trH, 'w:val');
    const height =
      hv !== undefined
        ? {
            value: twipsToPx(Number(hv)),
            exact: attrOf(trH, 'w:hRule') === 'exact',
          }
        : null;
    const cantSplit = isToggleOn(child(trPr, 'w:cantSplit'));
    const carry = collectCarry(trPr, CONSUMED_TRPR);
    return { header, height, cantSplit, carry };
  });

  const colIndex = logicalRows.map(
    (cells) => new Map(cells.map((c) => [c.startCol, c])),
  );

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
        ctx.schema.nodes['table_cell'].create(
          {
            colspan: cell.colspan,
            rowspan,
            colwidth: cell.colwidth,
            background: cell.background,
            vAlign: cell.vAlign,
            borders: cell.borders,
            padding: cell.padding,
            carry: cell.carry ? { tcPr: cell.carry } : null,
          },
          cell.content,
        ),
      );
    }
    const rp = rowProps[r];
    const rowAttrs: Record<string, unknown> = {};
    if (rp.header) rowAttrs['header'] = true;
    if (rp.height) rowAttrs['height'] = rp.height;
    if (rp.cantSplit) rowAttrs['cantSplit'] = true;
    if (rp.carry) rowAttrs['carry'] = { trPr: rp.carry };
    return ctx.schema.nodes['table_row'].create(
      Object.keys(rowAttrs).length > 0 ? rowAttrs : null,
      emitted.length > 0 ? emitted : [emptyCell(ctx)],
    );
  });

  const cellPadding = parseCellMargins(tbl, ctx);
  const borders = parseTableBorders(tbl, ctx);
  const jc = attrOf(child(child(tbl, 'w:tblPr'), 'w:jc'), 'w:val');
  const attrs: Record<string, unknown> = {};
  if (cellPadding) attrs['cellPadding'] = cellPadding;
  if (borders) attrs['borders'] = borders;
  if (jc === 'center' || jc === 'right' || jc === 'end')
    attrs['align'] = jc === 'end' ? 'right' : jc;
  // Carry-through: tblStyle/tblW/tblLayout/tblInd/tblLook/… survive the save.
  const tblCarry = collectCarry(child(tbl, 'w:tblPr'), CONSUMED_TBLPR);
  if (tblCarry) attrs['carry'] = { tblPr: tblCarry };
  return ctx.schema.nodes['table'].create(
    Object.keys(attrs).length > 0 ? attrs : null,
    rows.length > 0
      ? rows
      : [ctx.schema.nodes['table_row'].create(null, [emptyCell(ctx)])],
  );
}

/** Walk an element's children in document order, mapping w:p / w:tbl to blocks. */
function parseBlocks(parent: OoxmlNode, ctx: Ctx): PMNode[] {
  const blocks: PMNode[] = [];
  for (const node of unwrapSdt(parent.children)) {
    if (node.name === 'w:p' || node.name === 'w:tbl') audit.mark(node);
    if (node.name === 'w:p') blocks.push(parseParagraph(node, ctx));
    else if (node.name === 'w:tbl') blocks.push(parseTable(node, ctx));
  }
  return blocks;
}

interface SectionConfig {
  blockCount: number;
  columns: { count: number; gap: number };
  newPage: boolean;
  /** Geometry override when this section's w:pgSz/w:pgMar differ from the
   *  document default (the body sectPr). */
  page?: PageConfig;
  /** w:pgNumType — page-number restart (`start`) and/or display format
   *  (`fmt`, raw ST_NumberFormat value). Absent when the sectPr carries none:
   *  numbering then continues from the previous section in decimal. */
  pageNumbers?: { start?: number; fmt?: string };
}

/** A section's w:pgNumType, or undefined when it declares neither a restart
 *  value nor a format (an empty element is a no-op). */
function parsePageNumbers(
  sectPr: OoxmlNode | undefined,
): SectionConfig['pageNumbers'] {
  const pgNum = sectPr && child(sectPr, 'w:pgNumType');
  if (!pgNum) return undefined;
  const startRaw = attrOf(pgNum, 'w:start');
  const fmt = attrOf(pgNum, 'w:fmt');
  const start = startRaw == null ? NaN : Number(startRaw);
  const out: SectionConfig['pageNumbers'] = {};
  if (Number.isInteger(start) && start >= 0) out.start = start;
  if (fmt) out.fmt = fmt;
  return out.start != null || out.fmt ? out : undefined;
}

/** Column flow from a section's w:cols (equal-width only). count defaults to 1;
 *  the gap is w:space (twips→px), Word's default 720 twips = 0.5in. */
function parseColumns(sectPr: OoxmlNode | undefined): {
  count: number;
  gap: number;
} {
  const cols = sectPr && child(sectPr, 'w:cols');
  const num = Number(attrOf(cols, 'w:num') ?? '1');
  const explicit = cols ? children(cols, 'w:col').length : 0;
  const count = Math.max(Number.isNaN(num) ? 1 : num, explicit, 1);
  const spaceTw = Number(attrOf(cols, 'w:space') ?? '720');
  return { count, gap: twipsToPx(Number.isNaN(spaceTw) ? 720 : spaceTw) };
}

/** A continuous section break switches columns mid-page; every other type
 *  (next/odd/even page, or unspecified) starts the section on a new page. */
function sectionStartsNewPage(sectPr: OoxmlNode | undefined): boolean {
  return attrOf(child(sectPr, 'w:type'), 'w:val') !== 'continuous';
}

/** Parse the body into blocks while recording section boundaries. A w:p whose
 *  w:pPr carries a w:sectPr ends a section (with that sectPr's columns); the
 *  trailing w:body/w:sectPr ends the final section. */
function parseBodyBlocks(
  body: OoxmlNode,
  ctx: Ctx,
): { blocks: PMNode[]; sections: SectionConfig[] } {
  const blocks: PMNode[] = [];
  const sections: SectionConfig[] = [];
  let start = 0;
  for (const node of unwrapSdt(body.children)) {
    if (node.name === 'w:p' || node.name === 'w:tbl') audit.mark(node);
    if (node.name === 'w:p') {
      blocks.push(parseParagraph(node, ctx));
      const sectPr = child(child(node, 'w:pPr'), 'w:sectPr');
      if (sectPr) {
        const section: SectionConfig = {
          blockCount: blocks.length - start,
          columns: parseColumns(sectPr),
          newPage: sectionStartsNewPage(sectPr),
          // Every OOXML sectPr is self-contained — parse its geometry too.
          // importDocx drops `page` again on sections matching the document
          // default, so single-geometry docs stay on the fast path.
          page: parsePageGeometry(sectPr),
        };
        const pageNumbers = parsePageNumbers(sectPr);
        if (pageNumbers) section.pageNumbers = pageNumbers;
        sections.push(section);
        start = blocks.length;
      }
    } else if (node.name === 'w:tbl') {
      blocks.push(parseTable(node, ctx));
    }
  }
  // The trailing body sectPr closes the last (or only) section.
  const bodySectPr = child(body, 'w:sectPr');
  if (blocks.length > start || sections.length === 0) {
    const section: SectionConfig = {
      blockCount: blocks.length - start,
      columns: parseColumns(bodySectPr),
      newPage: sectionStartsNewPage(bodySectPr),
    };
    const pageNumbers = parsePageNumbers(bodySectPr);
    if (pageNumbers) section.pageNumbers = pageNumbers;
    sections.push(section);
  }
  return { blocks, sections };
}

async function readPart(zip: JSZip, path: string): Promise<string | undefined> {
  const entry = zip.file(path);
  return entry ? entry.async('string') : undefined;
}

/** Rels file that accompanies a part: "word/header1.xml" → "word/_rels/header1.xml.rels". */
async function readPartRels(
  zip: JSZip,
  partPath: string,
): Promise<OoxmlNode | undefined> {
  const slash = partPath.lastIndexOf('/');
  const relsPath = `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
  const xml = await readPart(zip, relsPath);
  if (!xml) return undefined;
  const root = parseXml(xml);
  audit.registerPart(relsPath, root);
  return root;
}

/** Load footnotes.xml / endnotes.xml into a numbering registry. Separator
 *  notes (negative ids / w:type separator) are skipped. */
async function buildNotesRegistry(zip: JSZip): Promise<NotesRegistry> {
  const bodies = {
    footnote: new Map<string, OoxmlNode>(),
    endnote: new Map<string, OoxmlNode>(),
  };
  const load = async (
    path: string,
    root: string,
    tag: string,
    into: Map<string, OoxmlNode>,
  ) => {
    const xml = await readPart(zip, path);
    if (!xml) return;
    const parsed = parseXml(xml);
    audit.registerPart(path, parsed);
    for (const note of children(child(parsed, root), tag)) {
      const id = attrOf(note, 'w:id');
      const type = attrOf(note, 'w:type');
      if (id === undefined || Number(id) < 1 || (type && type !== 'normal')) {
        // Separator/continuation chrome — skipped by design, subtree and all.
        audit.markSubtree(note);
        continue;
      }
      into.set(id, note);
    }
  };
  await load(
    'word/footnotes.xml',
    'w:footnotes',
    'w:footnote',
    bodies.footnote,
  );
  await load('word/endnotes.xml', 'w:endnotes', 'w:endnote', bodies.endnote);
  const reg: NotesRegistry = {
    bodies,
    refs: [],
    counter: { footnote: 0, endnote: 0 },
    ref(kind, id) {
      const num = ++reg.counter[kind];
      reg.refs.push({ kind, id, num });
      return num;
    },
  };
  return reg;
}

/** Load comments.xml (+ commentsExtended.xml) into a registry: bodies/author/
 *  date keyed by id, plus the paraId→id map and the w15 thread/resolved data. */
async function buildCommentsRegistry(zip: JSZip): Promise<CommentsRegistry> {
  const defs = new Map<
    number,
    { author: string; date: string; body: OoxmlNode; paraIds: string[] }
  >();
  const paraToId = new Map<string, number>();
  const xml = await readPart(zip, 'word/comments.xml');
  if (xml) {
    const parsed = parseXml(xml);
    audit.registerPart('word/comments.xml', parsed);
    for (const c of children(child(parsed, 'w:comments'), 'w:comment')) {
      // Comment bodies are deliberately flattened to plain text (collectText)
      // — their formatting subtree counts as consumed.
      audit.markSubtree(c);
      const id = Number(attrOf(c, 'w:id'));
      if (Number.isNaN(id)) continue;
      const paraIds = children(c, 'w:p')
        .map((p) => attrOf(p, 'w14:paraId'))
        .filter((v): v is string => !!v);
      for (const pid of paraIds) paraToId.set(pid, id);
      defs.set(id, {
        author: attrOf(c, 'w:author') ?? '',
        date: attrOf(c, 'w:date') ?? '',
        body: c,
        paraIds,
      });
    }
  }
  // Threaded comments (Word 2013+): commentsExtended.xml links replies by the
  // parent paragraph's paraId and carries the resolved flag (w15:done).
  const ext = new Map<string, { parentParaId: string | null; done: boolean }>();
  const extXml = await readPart(zip, 'word/commentsExtended.xml');
  if (extXml) {
    const parsed = parseXml(extXml);
    audit.registerPart('word/commentsExtended.xml', parsed);
    for (const ex of children(
      child(parsed, 'w15:commentsEx'),
      'w15:commentEx',
    )) {
      const paraId = attrOf(ex, 'w15:paraId');
      if (!paraId) continue;
      const done = attrOf(ex, 'w15:done');
      ext.set(paraId, {
        parentParaId: attrOf(ex, 'w15:paraIdParent') ?? null,
        done: done === '1' || done === 'true',
      });
    }
  }
  return { defs, paraToId, ext, active: new Set(), used: [] };
}

/** All w:t text under a node, concatenated (no side effects). */
function collectText(node: OoxmlNode): string {
  if (node.name === 'w:t') return node.text ?? '';
  return node.children.map(collectText).join('');
}

/** Referenced comments as flat data, in first-appearance order. */
function buildCommentsList(ctx: Ctx): CommentData[] {
  const out: CommentData[] = [];
  for (const id of ctx.comments.used) {
    const def = ctx.comments.defs.get(id);
    if (def)
      out.push({
        id,
        author: def.author,
        date: def.date,
        text: collectText(def.body).trim(),
      });
  }
  return out;
}

/** A comment body (w:comment) as commentSchema doc JSON — paragraphs of text. */
function commentBodyJSON(comment: OoxmlNode): unknown {
  const paras = children(comment, 'w:p').map((p) => {
    const text = collectText(p);
    return commentSchema.node(
      'paragraph',
      null,
      text ? [commentSchema.text(text)] : [],
    );
  });
  return commentSchema
    .node('doc', null, paras.length ? paras : [commentSchema.node('paragraph')])
    .toJSON();
}

/** Referenced comments as authoring thread nodes for doc.attrs.comments. OOXML
 *  comments carry only an author name, so the user id is the name itself.
 *  Thread parent + resolved come from commentsExtended.xml (w15) when present;
 *  without it (plain OOXML) every comment is a flat, unresolved root. */
function buildCommentNodes(ctx: Ctx): CommentNode[] {
  const { defs, paraToId, ext, used } = ctx.comments;
  const usedSet = new Set(used);
  // Resolve each comment's parent + resolved flag from the w15 thread data.
  const parentOf = new Map<number, number | null>();
  const resolvedOf = new Map<number, boolean>();
  for (const [id, def] of defs) {
    const exEntry = def.paraIds.map((p) => ext.get(p)).find(Boolean);
    const pid =
      exEntry?.parentParaId != null
        ? (paraToId.get(exEntry.parentParaId) ?? null)
        : null;
    parentOf.set(id, pid === id ? null : pid); // guard self-parent
    resolvedOf.set(id, exEntry?.done ?? false);
  }
  // Replies aren't referenced in the body (only the root's range is), so include
  // any comment whose thread root is referenced — walk up the parent chain.
  const referenced = (id: number): boolean => {
    for (
      let cur: number | null = id, n = 0;
      cur != null && n < 100;
      cur = parentOf.get(cur) ?? null, n++
    )
      if (usedSet.has(cur)) return true;
    return false;
  };
  // Roots in body-appearance order, then the remaining replies in id order.
  const ids = [
    ...used.filter((id) => defs.has(id)),
    ...[...defs.keys()].filter((id) => !usedSet.has(id)),
  ];
  const out: CommentNode[] = [];
  for (const id of ids) {
    const def = defs.get(id);
    if (!def || !referenced(id)) continue;
    const user: IUser = {
      id: def.author || 'unknown',
      name: def.author || 'Unknown',
    };
    out.push({
      id,
      parentId: parentOf.get(id) ?? null,
      user,
      date: def.date,
      body: commentBodyJSON(def.body),
      resolved: resolvedOf.get(id) ?? false,
    });
  }
  return out;
}

/** Parse one note body into blocks, prefixing the first paragraph with its
 *  display number (so "1. note text" reads naturally). Shared by the appended
 *  endnote section and the page-bottom footnote map. */
function noteBlocks(note: OoxmlNode, num: number, ctx: Ctx): PMNode[] {
  const blocks = parseBlocks(note, ctx);
  const marker = ctx.schema.text(`${num}. `, [
    ctx.schema.marks['vertAlign'].create({ value: 'super' }),
  ]);
  const first = blocks[0];
  if (first && first.type.name === 'paragraph') {
    const kids: PMNode[] = [marker];
    first.forEach((k) => kids.push(k));
    blocks[0] = ctx.schema.nodes['paragraph'].create(first.attrs, kids);
  } else {
    blocks.unshift(ctx.schema.nodes['paragraph'].create(null, [marker]));
  }
  return blocks;
}

/** Footnote bodies keyed by display number, each a standalone story document.
 *  The layout engine lays these out and reserves space for them at the bottom
 *  of whichever page their reference falls on. */
function buildFootnotesMap(ctx: Ctx): Record<number, PMNode> {
  const out: Record<number, PMNode> = {};
  for (const { kind, id, num } of ctx.notes.refs) {
    if (kind !== 'footnote') continue;
    const note = ctx.notes.bodies.footnote.get(id);
    if (note) out[num] = storyDoc(ctx, noteBlocks(note, num, ctx), null);
  }
  return out;
}

/** Endnotes render as an appended section at the document end (a heading + one
 *  paragraph per note). Footnotes are NOT appended here — they go to the bottom
 *  of their page (see buildFootnotesMap). */
function buildNotesSection(ctx: Ctx): PMNode[] {
  const endnotes = ctx.notes.refs.filter((r) => r.kind === 'endnote');
  if (endnotes.length === 0) return [];
  const out: PMNode[] = [
    ctx.schema.nodes['paragraph'].create({ spacing: { before: 12 } }, [
      ctx.schema.text('Ghi chú cuối', [ctx.schema.marks['strong'].create()]),
    ]),
  ];
  for (const { id, num } of endnotes) {
    const note = ctx.notes.bodies.endnote.get(id);
    if (note) out.push(...noteBlocks(note, num, ctx));
  }
  return out;
}

function storyDoc(
  ctx: Ctx,
  blocks: PMNode[],
  numbering: NumberingDefs | null,
  sections: SectionConfig[] | null = null,
  comments: CommentNode[] | null = null,
  page: PageConfig | null = null,
): PMNode {
  // doc content is `block+` — guarantee at least one paragraph. The numbering
  // defs (live markers), section column flow, page geometry, and comment
  // threads ride the doc as attrs.
  const attrs: Record<string, unknown> = {};
  if (numbering) attrs['numbering'] = numbering;
  if (sections) attrs['sections'] = sections;
  if (comments && comments.length > 0) attrs['comments'] = comments;
  if (page) attrs['page'] = page;
  return ctx.schema.nodes['doc'].create(
    Object.keys(attrs).length > 0 ? attrs : null,
    blocks.length > 0 ? blocks : [ctx.schema.nodes['paragraph'].create()],
  );
}

async function extractMedia(zip: JSZip): Promise<Map<string, string>> {
  const media = new Map<string, string>();
  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith('word/media/')) continue;
    const entry = zip.file(path);
    if (!entry || entry.dir) continue;
    media.set(
      path,
      `data:${mimeOf(path)};base64,${await entry.async('base64')}`,
    );
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
export async function importDocx(
  input: DocxInput,
  opts?: { schema?: Schema; password?: string },
): Promise<DocxImport> {
  // XML-audit session (no-op unless the flag is on): every part parsed inside
  // registers its root; endImport sweeps for untouched tags/attrs. Paired in
  // try/finally so a failed import can't wedge the audit's depth counter —
  // overlapping imports merge into one report (see audit.ts).
  const size =
    input instanceof Blob ? input.size : (input.byteLength ?? undefined);
  audit.beginImport(`document.xml (${size ?? '?'} bytes)`);
  try {
    return await importDocxImpl(input, opts);
  } finally {
    audit.endImport();
  }
}

async function importDocxImpl(
  input: DocxInput,
  opts?: { schema?: Schema; password?: string },
): Promise<DocxImport> {
  // Classify before parsing: a renamed PDF/.doc/encrypted file fails with a
  // cause a shell can act on, instead of JSZip's central-directory riddle.
  let bytes =
    input instanceof Blob ? new Uint8Array(await input.arrayBuffer()) : input;
  let sniff = sniffDocx(bytes);
  if (sniff === 'encrypted') {
    // Password-protected: without a password this is still a classified
    // failure the shell turns into its unlock prompt; with one, decrypting
    // yields the ordinary .docx zip and the rest of the import is unchanged.
    if (opts?.password === undefined) throw errorForSniff('encrypted');
    bytes = await decryptDocx(
      bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes),
      opts.password,
    );
    sniff = sniffDocx(bytes);
  }
  if (sniff !== 'zip') throw errorForSniff(sniff);
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch (err) {
    // A PK header whose archive won't open: truncated download/copy.
    const detail = err instanceof Error ? err.message : String(err);
    throw new DocxImportError(
      'corrupt-zip',
      IMPORT_ERROR_MESSAGES['corrupt-zip'],
      detail,
    );
  }

  const rawDocumentXml = await readPart(zip, 'word/document.xml');
  if (rawDocumentXml === undefined) {
    // A real zip without a Word body — most often a sibling Office format
    // renamed; say which one when its marker part is there.
    const kind = zip.file('xl/workbook.xml')
      ? ('xlsx' as const)
      : zip.file('ppt/presentation.xml')
        ? ('pptx' as const)
        : ('no-document' as const);
    throw new DocxImportError(kind, IMPORT_ERROR_MESSAGES[kind]);
  }

  const stylesXml = await readPart(zip, 'word/styles.xml');
  const numberingXml = await readPart(zip, 'word/numbering.xml');
  const themeXml = await readPart(zip, 'word/theme/theme1.xml');

  const parsePart = (name: string, xml: string): OoxmlNode => {
    const root = parseXml(xml);
    audit.registerPart(name, root);
    return root;
  };

  // Stateless/shared pieces; numbering counters are per-story (built fresh below).
  const themeRoot = themeXml
    ? parsePart('word/theme/theme1.xml', themeXml)
    : undefined;
  const resolveTheme = buildThemeResolver(themeRoot);
  const resolveFont = buildThemeFontResolver(themeRoot);
  const resolveThemeFill = buildThemeFillResolver(themeRoot, resolveTheme);
  const styles = buildStyleRegistry(
    stylesXml ? parsePart('word/styles.xml', stylesXml) : undefined,
    resolveTheme,
    resolveFont,
  );
  const numberingRoot = numberingXml
    ? parsePart('word/numbering.xml', numberingXml)
    : undefined;
  const media = await extractMedia(zip);
  const notes = await buildNotesRegistry(zip);
  const comments = await buildCommentsRegistry(zip);
  // Page geometry up front: pct-based table widths need the content width
  // while the body is being parsed.
  const body = child(
    child(parsePart('word/document.xml', rawDocumentXml), 'w:document'),
    'w:body',
  );
  const sectPr = body ? child(body, 'w:sectPr') : undefined;
  const pageGeom = parsePageGeometry(sectPr);
  const contentWidth =
    pageGeom.width - pageGeom.margin.left - pageGeom.margin.right;

  // Stateless — markers are recounted by the layout engine, so one resolver
  // serves every story (and its audit usage-tracking sees them all).
  const numbering = buildNumbering(numberingRoot, resolveTheme, resolveFont);
  const makeCtx = (rels: Map<string, Relationship>): Ctx => ({
    styles,
    numbering,
    rels,
    media,
    resolveTheme,
    resolveFont,
    resolveThemeFill,
    notes,
    comments,
    schema: opts?.schema ?? schema,
    contentWidth,
    // Per-story: a field opened in the body can't leak into a header.
    openFields: [],
  });

  const docRels = await readPart(zip, 'word/_rels/document.xml.rels');
  const ctx = makeCtx(
    buildRels(
      docRels ? parsePart('word/_rels/document.xml.rels', docRels) : undefined,
    ),
  );

  const parsed = body
    ? parseBodyBlocks(body, ctx)
    : { blocks: [], sections: [] };
  // References are now assigned (in document order); footnotes go to the page
  // bottom, endnotes stay appended (and fall into the last section).
  const footnotes = buildFootnotesMap(ctx);
  const endnoteBlocks = buildNotesSection(ctx);
  const { sections } = parsed;
  if (endnoteBlocks.length > 0 && sections.length > 0) {
    sections[sections.length - 1].blockCount += endnoteBlocks.length;
  }
  // Sections whose geometry matches the document default carry no override —
  // the common all-one-geometry doc keeps every `page` absent.
  const samePage = (a: PageConfig, b: PageConfig) =>
    a.width === b.width &&
    a.height === b.height &&
    a.margin.top === b.margin.top &&
    a.margin.right === b.margin.right &&
    a.margin.bottom === b.margin.bottom &&
    a.margin.left === b.margin.left;
  for (const s of sections) {
    if (s.page && samePage(s.page, pageGeom)) delete s.page;
  }
  // Only ride the sections attr when it changes layout (>1 section, columns,
  // a per-section geometry override, or page numbering — even a one-section
  // doc can restart/reformat its page numbers).
  const multiSection =
    sections.length > 1 ||
    sections.some((s) => s.columns.count > 1 || s.page || s.pageNumbers);
  // Comment threads only ride the doc when the schema carries the comment mark
  // (the comment plugin is present); otherwise comment values are filtered out.
  const hasComments = !!ctx.schema.marks['comment'];
  const doc = storyDoc(
    ctx,
    [...parsed.blocks, ...endnoteBlocks],
    ctx.numbering.defs,
    multiSection ? sections : null,
    hasComments ? buildCommentNodes(ctx) : null,
    pageGeom, // page geometry rides the doc so page-setup edits are undoable
  );

  // Headers/footers, per section. Every sectPr (paragraph-level breaks +
  // the body's) is visited in document order; a section that doesn't declare
  // a w:headerReference/w:footerReference of a type inherits the previous
  // section's story — Word's "Link to Previous". Parts are parsed once and
  // shared (memoized by path) across the sections referencing them.
  const partMemo = new Map<string, PMNode | undefined>();
  const loadChromePart = async (
    rId: string | undefined,
    root: string,
  ): Promise<PMNode | undefined> => {
    const target = rId ? ctx.rels.get(rId)?.target : undefined;
    if (!target) return undefined;
    const partPath = `word/${target.replace(/^\/+/, '')}`;
    if (partMemo.has(partPath)) return partMemo.get(partPath);
    const xml = await readPart(zip, partPath);
    let story: PMNode | undefined;
    if (xml) {
      const partCtx = makeCtx(buildRels(await readPartRels(zip, partPath)));
      const el = child(parsePart(partPath, xml), root);
      story = storyDoc(
        partCtx,
        el ? parseBlocks(el, partCtx) : [],
        ctx.numbering.defs,
      );
    }
    partMemo.set(partPath, story);
    return story;
  };
  const chromeOf = async (sp: OoxmlNode | undefined) => {
    const own: SectionChrome = { headers: {}, footers: {}, titlePg: false };
    let declared = false;
    if (sp) {
      const collect = async (
        refName: string,
        store: Record<string, PMNode>,
        root: string,
      ) => {
        for (const ref of children(sp, refName)) {
          const type = attrOf(ref, 'w:type') ?? 'default';
          if (store[type]) continue;
          const story = await loadChromePart(attrOf(ref, 'r:id'), root);
          if (story) {
            store[type] = story;
            declared = true;
          }
        }
      };
      await collect('w:headerReference', own.headers, 'w:hdr');
      await collect('w:footerReference', own.footers, 'w:ftr');
      own.titlePg = isToggleOn(child(sp, 'w:titlePg'));
      if (own.titlePg) declared = true;
    }
    return { own, declared };
  };
  // sectPrs in document order, aligned with `sections` (the body sectPr is
  // the last section's — unless the doc ended AT a paragraph break and the
  // body sectPr closed no section of its own).
  const sectPrList: (OoxmlNode | undefined)[] = [];
  if (body) {
    for (const node of unwrapSdt(body.children)) {
      if (node.name === 'w:p') {
        const sp = child(child(node, 'w:pPr'), 'w:sectPr');
        if (sp) sectPrList.push(sp);
      }
    }
  }
  sectPrList.push(sectPr);
  const sectionChrome: SectionChrome[] = [];
  let anyDeclaredBeforeLast = false;
  for (let i = 0; i < sections.length; i++) {
    const { own, declared } = await chromeOf(sectPrList[i]);
    const prev = sectionChrome[i - 1];
    sectionChrome.push({
      headers: { ...(prev?.headers ?? {}), ...own.headers },
      footers: { ...(prev?.footers ?? {}), ...own.footers },
      titlePg: own.titlePg,
    });
    if (declared && i < sections.length - 1) anyDeclaredBeforeLast = true;
  }
  const lastChrome = sectionChrome[sectionChrome.length - 1] ?? {
    headers: {},
    footers: {},
    titlePg: false,
  };
  // The flat fields stay the LAST section's resolved chrome (what the old
  // body-sectPr-only path produced, plus inheritance).
  const headers = lastChrome.headers;
  const footers = lastChrome.footers;

  // w:titlePg (section) → page 1 uses the "first" chrome; w:evenAndOddHeaders
  // (document settings) → even pages use the "even" chrome.
  const titlePg = lastChrome.titlePg;
  const settingsXml = await readPart(zip, 'word/settings.xml');
  const settings = settingsXml
    ? child(parsePart('word/settings.xml', settingsXml), 'w:settings')
    : undefined;
  const evenAndOdd = settings
    ? isToggleOn(child(settings, 'w:evenAndOddHeaders'))
    : false;
  // Default tab interval (w:defaultTabStop, twips) — drives the layout's
  // implicit tab grid when a paragraph has no explicit stops.
  const defaultTabStop = settings
    ? Number(attrOf(child(settings, 'w:defaultTabStop'), 'w:val'))
    : NaN;
  const tabWidth =
    Number.isFinite(defaultTabStop) && defaultTabStop > 0
      ? twipsToPx(defaultTabStop)
      : undefined;

  // Audit: with every story parsed (body, notes, headers/footers), styles
  // and numbering levels nothing referenced are swept out of the report —
  // see StyleRegistry / NumberingResolver.
  styles.auditMarkUnusedStyles();
  numbering.auditMarkUnusedLevels();

  const out: DocxImport = {
    doc,
    rawDocumentXml,
    headers,
    footers,
    footnotes,
    titlePg,
    evenAndOdd,
    comments: hasComments ? buildCommentsList(ctx) : [],
    page: pageGeom,
    ...(tabWidth !== undefined && { tabWidth }),
    raw: zip,
  };
  // Only when some non-last section declares its own chrome/titlePg does the
  // per-section set differ from the flat fields.
  if (sections.length > 1 && anyDeclaredBeforeLast) {
    out.sectionChrome = sectionChrome;
  }
  return out;
}

/** An OOXML on/off toggle element (w:titlePg, w:evenAndOddHeaders, …): present
 *  means on, unless it carries w:val="false"/"0"/"off". */
function isToggleOn(el: OoxmlNode | undefined): boolean {
  if (!el) return false;
  const val = attrOf(el, 'w:val');
  return val === undefined || !['false', '0', 'off'].includes(val);
}

/** Page size + margins from w:sectPr (twips→px). Defaults to A4 @96dpi with
 *  1in margins; landscape swaps w/h. Header/footer distances aren't returned —
 *  the layout engine uses its own chrome distance. Exported for the exporter,
 *  which re-parses the carried sectPr to detect page-setup edits. */
export function parsePageGeometry(sectPr: OoxmlNode | undefined): PageConfig {
  const A4: PageConfig = {
    width: 794,
    height: 1123,
    margin: { top: 96, right: 96, bottom: 96, left: 96 },
  };
  if (!sectPr) return A4;
  const pgSz = child(sectPr, 'w:pgSz');
  const pgMar = child(sectPr, 'w:pgMar');
  // OOXML measurements are plain twips numbers, but non-Word producers ship
  // unit suffixes in the wild (w:top="20pt") — Number() would go NaN and NaN
  // margins send pagination into an infinite loop. Parse unit-aware and fall
  // back on anything unparseable.
  const px = (el: OoxmlNode | undefined, attr: string, fallback: number) => {
    const v = el && attrOf(el, attr);
    if (v === undefined || v === null) return fallback;
    const m = /^(-?\d+(?:\.\d+)?)(pt|in|cm|mm)?$/.exec(String(v).trim());
    if (!m) return fallback;
    const n = Number(m[1]);
    const PX_PER: Record<string, number> = {
      pt: 96 / 72,
      in: 96,
      cm: 96 / 2.54,
      mm: 96 / 25.4,
    };
    const out = m[2] ? Math.round(n * PX_PER[m[2]]) : twipsToPx(n);
    return Number.isFinite(out) ? out : fallback;
  };
  let width = px(pgSz, 'w:w', A4.width);
  let height = px(pgSz, 'w:h', A4.height);
  if (attrOf(pgSz, 'w:orient') === 'landscape' && height > width) {
    [width, height] = [height, width];
  }
  const geom: PageConfig = {
    width,
    height,
    margin: {
      top: px(pgMar, 'w:top', 96),
      right: px(pgMar, 'w:right', 96),
      bottom: px(pgMar, 'w:bottom', 96),
      left: px(pgMar, 'w:left', 96),
    },
  };
  // Chrome distances + binding gutter: set only when declared (the layout
  // falls back to Word's defaults), so untouched docs stay byte-stable.
  if (pgMar && attrOf(pgMar, 'w:header') !== undefined)
    geom.headerDistance = px(pgMar, 'w:header', 48);
  if (pgMar && attrOf(pgMar, 'w:footer') !== undefined)
    geom.footerDistance = px(pgMar, 'w:footer', 48);
  const gutter = pgMar ? px(pgMar, 'w:gutter', 0) : 0;
  if (gutter > 0) geom.gutter = gutter;
  return geom;
}
