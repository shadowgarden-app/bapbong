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
  MATH_ALPHABETS,
  mathLetters,
  type MathAlphabet,
} from '@shadow-garden/bapbong-contracts';
import {
  attrOf,
  child,
  children,
  findDescendant,
  isToggleOn,
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
  CellDiagonals,
  GradientFill,
  BorderStyle,
  ColumnConfig,
  DocCompat,
  ShapeSpec,
  TableBorders,
} from '@shadow-garden/bapbong-contracts';
import { audit } from './audit.js';
import { buildStyleRegistry, CondLayer, StyleRegistry } from './styles.js';
import { parseCompat } from './compat.js';
import { emfBitmapDataUrl, wmfBitmapDataUrl } from './emf.js';
import { wmfVectorSpec, WmfVectorResult } from './wmf-vector.js';
import { buildNumbering, NumberingResolver } from './numbering.js';
import { buildRels, Relationship } from './rels.js';
import {
  parseGradient,
  buildThemeFillResolver,
  buildThemeFontResolver,
  buildThemeLineResolver,
  buildThemeResolver,
  drawingColor,
  ThemeFillResolver,
  ThemeFontResolver,
  ThemeLineResolver,
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
  /** `num` keys the body (every note gets one); `display` is the VISIBLE
   *  auto number — absent for custom-marked notes, which show their own
   *  glyph and, per Word (measured), do not consume a number. */
  refs: {
    kind: 'footnote' | 'endnote';
    id: string;
    num: number;
    display?: number;
  }[];
  counter: { footnote: number; endnote: number };
  display: { footnote: number; endnote: number };
  /** Assign (and remember) the key + display number for a reference.
   *  `custom` (w:customMarkFollows) skips the display counter. */
  ref(
    kind: 'footnote' | 'endnote',
    id: string,
    custom?: boolean,
  ): { num: number; display?: number };
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

/** Structured-document-tag wrapper attached to its boundary blocks: the
 *  serialized INNER XML of w:sdtPr / w:sdtEndPr rides the first block, the
 *  end flag rides the last, and export re-wraps the range. */
export interface SdtBoundary {
  start?: { pr: string | null; endPr: string | null };
  end?: boolean;
}

interface Ctx {
  styles: StyleRegistry;
  numbering: NumberingResolver;
  rels: Map<string, Relationship>;
  media: Map<string, string>; // zip path → data URL
  /** Vector WMFs (MathType equation previews) resolved to display lists,
   *  keyed like `media`. The media entry keeps the original bytes. */
  vectorMedia: Map<string, WmfVectorResult>;
  resolveTheme: ThemeResolver;
  resolveFont: ThemeFontResolver;
  /** Shape fill from an `a:fillRef` (theme format scheme + placeholder). */
  resolveThemeFill: ThemeFillResolver;
  /** The theme line style an `a:lnRef` names (width, dash, cap). */
  resolveThemeLine: ThemeLineResolver;
  /**
   * Table styles of the tables currently being parsed, innermost LAST.
   *
   * A table style's w:pPr/w:rPr are defaults for the content inside that
   * table, so a paragraph needs to know which table encloses it. The stack
   * holds layers already resolved — one roll-up per TABLE, not per paragraph:
   * the registry memoises nothing and khbd has 275 tables holding 4704
   * paragraphs between them.
   *
   * Only the innermost entry applies. A nested table is a different table
   * with a style of its own, and Word does not cascade the outer one into it.
   * Stories that merely sit inside a cell — a textbox's w:txbxContent — are
   * not "content of the table" either, so the parser CLEARS this while
   * reading one (see txbxParagraphs).
   */
  tableStyles: { pPr: OoxmlNode[]; rPr: RunProps }[];
  notes: NotesRegistry;
  /** SDT wrappers keyed by their boundary w:p / w:tbl nodes — filled by
   *  unwrapContainers, read by parseParagraph / parseTable into `carry`. */
  sdtBoundaries: WeakMap<OoxmlNode, SdtBoundary>;
  comments: CommentsRegistry;
  /** Schema the doc nodes/marks are created with (model's by default; the editor
   *  may inject a composed schema so plugin-contributed marks are imported). */
  schema: Schema;
  /** Page content-box width in px (page minus side margins) — what
   *  percentage-based table widths (w:tblW/w:tcW type="pct") resolve against. */
  contentWidth: number;
  /** The document's Word compatibility profile (settings.xml `w:compat`,
   *  resolved once by parseCompat). Every mode-dependent rule reads from here. */
  compat: DocCompat;
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
    case 'emf':
      return 'image/emf';
    case 'wmf':
      return 'image/wmf';
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
  if (p.kern !== undefined && ctx.schema.marks['kern'])
    marks.push(ctx.schema.marks['kern'].create({ halfPoints: p.kern }));
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

/** Wingdings 2 is a DIFFERENT font with its own layout — not a superset of
 *  Wingdings, though it shares the name's stem: 0x52 is the ticked box here
 *  and a pointing hand there. Word's HR forms tick their boxes with it
 *  (`w:sym w:font="Wingdings 2" w:char="F052"`). Codes per Microsoft's chart
 *  as tabulated at alanwood.net/demos/wingdings-2.html; the Unicode 7 astral
 *  equivalents (🗴, 🗵, ⯾) are traded for BMP look-alikes that every font
 *  set has (✗, ☒, ⊗). */
const WINGDINGS_2: Record<number, string> = {
  0x4f: '✗', // 🗴 ballot script X
  0x50: '✓',
  0x51: '☒', // 🗵 ballot box with script X
  0x52: '☑',
  0x53: '☒',
  0x54: '☒',
  0x55: '⊗', // ⯾ circled X
  0x56: '⊗',
  0x95: '•',
  0x96: '●',
  0x97: '●',
  0x98: '●',
  0x99: '○',
  0x9a: '○',
  0x9b: '○',
  0x9c: '○',
  0x9d: '◉',
  0x9e: '⦿',
  0x9f: '◾',
  0xa0: '■',
  0xa1: '◼',
  0xa2: '■',
  // The empty boxes: BALLOT BOX, like WINGDINGS' 0x6F–0x73, so an unticked
  // one sits beside a ticked one at the same size.
  0xa3: '☐',
  0xa4: '☐',
  0xa5: '☐',
};

/** Monotype Sorts — Monotype's clone of ITC Zapf Dingbats, same encoding.
 *  Only the box family, which forms lean on: 0x6F–0x72 are Zapf's four
 *  squares. 0x7F is unassigned in Zapf Dingbats yet Word draws an EMPTY
 *  BALLOT BOX for it (an application form pairs "Monotype Sorts F07F Not
 *  yet" with "Wingdings 2 F052 Yes", and Word shows ☐ beside ☑). */
const MONOTYPE_SORTS: Record<number, string> = {
  0x6f: '❏',
  0x70: '❐',
  0x71: '❑',
  0x72: '❒',
  0x7f: '☐',
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
  const f = font.toLowerCase().trim();
  // Exact names: "Wingdings 2" and "Wingdings 3" are different fonts, and a
  // prefix match once sent Wingdings 2's ticked box through Wingdings'
  // table (0x52 there is a pointing hand — the box came out as tofu).
  if (f === 'wingdings') return WINGDINGS;
  if (f === 'wingdings 2') return WINGDINGS_2;
  if (f === 'monotype sorts') return MONOTYPE_SORTS;
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

/** The wrappers' own property bags — the markup's name/uri metadata, never
 *  content. Dropped whole so the audit counts them handled, not missed. */
const WRAPPER_PROPS = new Set(['w:smartTagPr', 'w:customXmlPr']);

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
    if (WRAPPER_PROPS.has(node.name)) {
      audit.markSubtree(node);
      continue;
    }
    if (node.name === 'w:ins' || node.name === 'w:moveTo') {
      audit.mark(node);
      out.push(...effectiveChildren(node.children));
    } else if (node.name === 'w:smartTag') {
      // CT_SmartTagRun: a transparent wrapper around inline content (runs,
      // hyperlinks, fields, math, other smart tags). Word writes these when
      // it recognises a place or a date; the tag itself renders nothing, but
      // whatever it wraps is ordinary text — bc_rieng lost the "Nam" of
      // "Việt Nam" to one before this unwrapped.
      audit.mark(node);
      out.push(...effectiveChildren(node.children));
    } else if (node.name === 'w:sdt' || node.name === 'w:customXml') {
      out.push(...effectiveChildren(unwrapContainers([node])));
    } else {
      out.push(node);
    }
  }
  return out;
}

/** Chrome wrappers that carry content and nothing of their own, at BLOCK or
 *  inline level: w:sdt (content controls — content sits in w:sdtContent) and
 *  w:customXml (custom-XML markup — content is the children themselves). The
 *  chrome is dropped, the content (paragraphs/tables at block level, runs
 *  inline — a w14:checkbox's ☒/☐ glyph run included) survives. Recurses so
 *  nested wrappers (cover pages hold several) fully unwrap. */
function unwrapContainers(
  nodes: OoxmlNode[],
  sdt?: WeakMap<OoxmlNode, SdtBoundary>,
): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  for (const node of nodes) {
    if (WRAPPER_PROPS.has(node.name)) {
      audit.markSubtree(node);
    } else if (node.name === 'w:sdt') {
      audit.mark(node);
      const content = child(node, 'w:sdtContent');
      if (content) {
        const inner = unwrapContainers(content.children, sdt);
        // The wrapper's properties ride its boundary blocks so a save can
        // re-wrap the range (page-number building blocks, TOC chrome). One
        // flat level only — the OUTERMOST sdt wins and wipes any markers an
        // inner one left, so export always sees balanced, non-crossing
        // ranges. Serialized through CARRY_FILTER (w:-prefixed data only:
        // a w14:checkbox does not survive — a flat re-emit would need the
        // extended namespaces declared on our root).
        if (sdt && inner.length > 0) {
          const ser = (el: OoxmlNode | undefined): string | null => {
            if (!el) return null;
            audit.markSubtree(el);
            return (
              el.children
                .map((c) => serializeOoxml(c, CARRY_FILTER))
                .join('') || null
            );
          };
          const pr = ser(child(node, 'w:sdtPr'));
          const endPr = ser(child(node, 'w:sdtEndPr'));
          if (pr !== null || endPr !== null) {
            for (const n of inner) sdt.delete(n);
            const first = inner[0];
            const last = inner[inner.length - 1];
            sdt.set(first, { start: { pr, endPr } });
            sdt.set(last, { ...(sdt.get(last) ?? {}), end: true });
          }
        }
        out.push(...inner);
      }
    } else if (node.name === 'w:customXml') {
      audit.mark(node);
      out.push(...unwrapContainers(node.children));
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
  let textMarks = asSymbols
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
      // Read OUTSIDE the body-exists guard: a dangling reference must not
      // leave the attribute looking unread to the audit.
      const customAttr = attrOf(node, 'w:customMarkFollows');
      const custom = customAttr === '1' || customAttr === 'true';
      if (id && ctx.notes.bodies[kind].has(id)) {
        flush();
        // w:customMarkFollows: the glyph is the REST OF THIS RUN's text (the
        // † the author typed), no auto number is drawn and the note does not
        // consume one — measured: marks render 1, †, 2, not 1, 2†, 3.
        const { num, display } = ctx.notes.ref(kind, id, custom);
        if (custom) {
          // The custom glyph needs the footnote mark so the layout engine
          // still links it to its page-bottom body.
          if (kind === 'footnote')
            textMarks = [
              ...textMarks,
              ctx.schema.marks['footnote'].create({ num, id }),
            ];
        } else {
          const refMarks = [
            ...marks,
            ctx.schema.marks['vertAlign'].create({ value: 'super' }),
          ];
          // Footnotes carry a `footnote` mark so the layout engine can match
          // the reference to its page-bottom body; endnotes stay plain
          // superscripts (their bodies are appended at the document end).
          if (kind === 'footnote')
            refMarks.push(ctx.schema.marks['footnote'].create({ num, id }));
          out.push(ctx.schema.text(String(display), refMarks));
        }
      }
    } else if (node.name === 'w:br') {
      // A page break is a paragraph-level property (hasPageBreak reads it);
      // a COLUMN break is inline, because ST_BrType `column` restarts "the
      // next character" — everything before it, floats included, belongs to
      // the column being left behind. Everything else is a line break.
      const type = attrOf(node, 'w:type');
      if (type !== 'page') {
        flush();
        out.push(
          ctx.schema.nodes[
            type === 'column' ? 'column_break' : 'hard_break'
          ].create(),
        );
      }
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
  // Tight/through carve the same rectangle here (the polygon is not
  // modelled) but anchor differently — see FlowFloat.through.
  if (child(anchor, 'wp:wrapTight') || child(anchor, 'wp:wrapThrough'))
    float['through'] = true;
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
    // ST_RelFromH: margin | page | column | character | leftMargin |
    // rightMargin | insideMargin | outsideMargin. `column` positions the
    // object "with respect to the column it resides in" — identical to the
    // margin box in a single-column section, which is how it went unnoticed,
    // and 67 of the 69 anchors in the corpus use it. The four side-margin
    // bases and `character` have no model counterpart (and no document here
    // uses them); they keep the margin reading rather than a made-up one.
    float['hRel'] =
      rel === 'page' ? 'page' : rel === 'column' ? 'column' : 'margin';
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

/** Namespace prefixes an `mc:Choice` may require and still be renderable
 *  here: the 2010 shape/group extensions, plus the drawing extras that only
 *  add sizing hints. Anything else — `aink` (pen strokes), `am3d`, future
 *  namespaces — means the Choice holds markup we would silently drop, so the
 *  Fallback is the better branch. Being conservative is safe by design: a
 *  Fallback is, by definition, what an older consumer can render. */
const MC_UNDERSTOOD = new Set(['wps', 'wpg', 'wp14', 'w14', 'pic', 'v', 'o']);

/** The run's w:drawing, resolving `mc:AlternateContent` the way ISO/IEC
 *  29500-3 says to: walk the `mc:Choice` branches in order, take the first
 *  whose `@Requires` namespaces are ALL understood, else fall back to
 *  `mc:Fallback`.
 *
 *  This used to take whichever branch simply had a w:drawing, which is not
 *  the same thing. khbd's inked pages are the counter-example: five
 *  `Requires="aink"` Choices hand us a w:drawing whose graphicData holds
 *  `w14:contentPart` (a pen stroke we cannot draw), so the drawing was found,
 *  then abandoned — while the Fallback right beside it holds the same stroke
 *  rasterised as an ordinary picture.
 *
 *  Only w:drawing goes through here. A `w:pict` inside a Fallback is the
 *  legacy twin of a Choice we already took, so the VML paths keep reading the
 *  run's own child; a document whose ONLY representation is Choice(unknown)
 *  + Fallback(w:pict) would still render nothing. No file in the corpus has
 *  that shape. */
function runDrawing(run: OoxmlNode): OoxmlNode | undefined {
  const direct = child(run, 'w:drawing');
  if (direct) return direct;
  const alt = child(run, 'mc:AlternateContent');
  if (!alt) return undefined;
  const skipped: OoxmlNode[] = [];
  for (const c of children(alt, 'mc:Choice')) {
    const requires = (attrOf(c, 'Requires') ?? '').split(/\s+/).filter(Boolean);
    if (!requires.every((ns) => MC_UNDERSTOOD.has(ns))) {
      skipped.push(c);
      continue;
    }
    const hit = child(c, 'w:drawing');
    if (hit) return hit;
  }
  for (const b of children(alt, 'mc:Fallback')) {
    const hit = child(b, 'w:drawing');
    // Passing over a Choice is a decision, not a miss — but only once we
    // have somewhere else to go, so the subtrees are written off here.
    if (hit) {
      for (const s of skipped) audit.markSubtree(s);
      return hit;
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
  const picDescr = attrOf(docPr, 'descr');
  const picTitle = attrOf(docPr, 'title');
  const float = parseAnchorFloat(drawing);
  const crop = parseSrcRect(findDescendant(drawing, 'a:srcRect'));
  // Word's "picture border": an a:ln on the picture's own spPr. Shapes have
  // always had their outline read; pictures never did, so a framed photo lost
  // its frame. Reuses the border reader the tables use, so width/style/colour
  // are interpreted one way across the converter. Not modelled: @cmpd
  // (single / double / thickThin …), which is a second line rather than a
  // property of this one.
  const picSpPr = findDescendant(drawing, 'pic:spPr');
  const picLn = child(picSpPr, 'a:ln');
  // A fill on the picture's own spPr paints BEHIND the bitmap — visible
  // through transparent pixels. Word draws it (measured); dropping it turns
  // a matted logo transparent.
  const background = solidFillColor(picSpPr, ctx) ?? null;
  // Both attributes are asked for whether or not the outline paints. Word
  // writes `<a:ln w="9525"><a:noFill/>…</a:ln>` on plenty of pictures — an
  // explicit "no border" that still states a width — and reading the element
  // without reading them would report that boilerplate as a gap. Same move
  // parseBorderSide makes for a hidden side.
  const picW = picLn ? attrOf(picLn, 'w') : undefined;
  if (picLn) attrOf(picLn, 'cmpd');
  const outline =
    picLn && !child(picLn, 'a:noFill')
      ? pictureOutline(picLn, picW, ctx)
      : null;
  // w:position on the picture's run — same baseline shift the legacy object
  // path reads (see parseVmlImage); our own export writes it back here.
  const posHp = Number(
    attrOf(child(child(run, 'w:rPr'), 'w:position'), 'w:val'),
  );
  return ctx.schema.nodes['image'].create({
    src,
    width: emuToPx(attrOf(extent, 'cx')),
    height: emuToPx(attrOf(extent, 'cy')),
    raise: Number.isFinite(posHp) ? (posHp / 2) * (96 / 72) : 0,
    // Both docPr attrs are read EAGERLY: `??` used to short-circuit past
    // @title whenever @descr existed (even empty), leaving @title unread —
    // and dropped on save. alt is @descr; @title rides its own attr.
    alt: picDescr ?? '',
    ...(picTitle != null && { title: picTitle }),
    ...(background && { background }),
    float,
    ...(crop && { crop }),
    ...(outline && { outline }),
    rotation: xfrmRotation(drawing),
  });
}

/** A picture's outline as a {@link BorderSide}. DrawingML states the width in
 *  EMU and the colour as a fill, where a table border states eighths of a
 *  point and a hex — same shape, different units, so the conversion happens
 *  here and the painter treats both alike. */
function pictureOutline(
  ln: OoxmlNode,
  w: string | undefined,
  ctx: Ctx,
): BorderSide {
  const width = w ? Math.max(0.75, Number(w) / 9525) : 1;
  const prstDash = attrOf(child(ln, 'a:prstDash'), 'val');
  return {
    width,
    style: prstDash && prstDash !== 'solid' ? 'dashed' : 'solid',
    color: solidFillColor(ln, ctx) ?? '#000000',
  };
}

/**
 * `a:srcRect` — which part of the bitmap the picture shows.
 *
 * CT_RelativeRect: four ST_Percentage offsets measured from the matching edge
 * of the bounding box, each defaulting to 0. Positive is an INSET (crop in),
 * negative an OUTSET (the source rectangle reaches past the bitmap, and the
 * overhang shows as nothing). The selected rectangle is then scaled to fill
 * the picture's box, so a crop changes what you see, never how big the box is.
 *
 * Returns null for the all-zero case — which is most of them: Word stamps an
 * empty `<a:srcRect/>` on nearly every picture it writes (18 of khbd's 21),
 * and the audit already classifies that shape as inert. Reading the four
 * attributes here and then declining to store anything keeps that verdict
 * intact instead of trading 18 inert elements for 72 unread attributes.
 */
function parseSrcRect(
  el: OoxmlNode | undefined,
): { l: number; t: number; r: number; b: number } | null {
  if (!el) return null;
  const side = (name: string) => signedPct(attrOf(el, name));
  const l = side('l'),
    t = side('t'),
    r = side('r'),
    b = side('b');
  return l || t || r || b ? { l, t, r, b } : null;
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
  const mediaKey = ctx.media.has(`word/${target}`) ? `word/${target}` : target;
  const src = ctx.media.get(mediaKey);
  if (!src) return null;
  // Equation previews (vector WMF) resolved at extract time: the node keeps
  // the original bytes in `src` and carries the display list beside them.
  const vector = ctx.vectorMedia.get(mediaKey);

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
  // MathType aligns the equation's baseline with the text line by LOWERING
  // the object run (w:position, half-points, negative = down) — without it
  // every formula floats above the line. Same convention as InlineRun.raise:
  // px, positive UP, paint-only (Word does not grow the line box).
  const posHp = Number(
    attrOf(child(child(run, 'w:rPr'), 'w:position'), 'w:val'),
  );
  const raise = Number.isFinite(posHp) ? (posHp / 2) * (96 / 72) : 0;
  const width =
    ptToPx(/(?:^|;)width:([\d.]+)pt/.exec(style)) ??
    (Number(attrOf(holder, 'w:dxaOrig'))
      ? twipsToPx(Number(attrOf(holder, 'w:dxaOrig')))
      : (vector?.pxWidth ?? null));
  const height =
    ptToPx(/(?:^|;)height:([\d.]+)pt/.exec(style)) ??
    (Number(attrOf(holder, 'w:dyaOrig'))
      ? twipsToPx(Number(attrOf(holder, 'w:dyaOrig')))
      : (vector?.pxHeight ?? null));

  return ctx.schema.nodes['image'].create({
    src,
    width,
    height,
    alt: vmlAltText,
    float: null,
    vector: vector?.spec ?? null,
    raise,
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
  // A bodyPr on a shape with NO text configures nothing — Word still stamps
  // the full attribute set on every connector it writes (15 per drawing in
  // one corpus file). Consumed as a decision, not silently ignored: the
  // moment the shape gains a txbx, parseTextbox reads it for real.
  if (!textbox) {
    const idleBodyPr = child(wsp, 'wps:bodyPr');
    if (idleBodyPr) audit.markSubtree(idleBodyPr);
  }
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
    shape['strokeWidth'] = 1;
    // A shape's outline starts from the theme line style its a:lnRef names
    // and then takes whatever its own a:ln states on top — so a shape from
    // the gallery, which carries only the ref, still gets the theme's width.
    // Both passes go through the same reader, so the two sources cannot
    // drift apart in how they are interpreted.
    applyLineProps(ctx.resolveThemeLine(child(style, 'a:lnRef')), shape);
    applyLineProps(ln, shape);
    // Direct outline color, else the style's line reference (how Word themes
    // shape outlines), else black. BOTH are read before either is chosen: a
    // `??` chain would leave the loser untouched, and to the coverage audit
    // an untouched node is indistinguishable from an unsupported one. This is
    // the third time that has bitten (v:imagedata @o:title, wp:docPr
    // descr/title), so it is worth spelling out rather than shortening.
    const directStroke = solidFillColor(ln, ctx);
    const themedStroke = drawingColor(
      findDescendant(style, 'a:lnRef'),
      ctx.resolveTheme,
    );
    shape['stroke'] = directStroke ?? themedStroke ?? '#000000';
  }
  // Fill: an explicit a:noFill wins, then the shape's own a:solidFill or
  // a:gradFill, then the shape style's a:fillRef resolved through the
  // theme's format scheme — which is how Word fills every shape inserted
  // from the shape gallery. A gradient (direct or themed) carries resolved
  // stops; anything else stays a flat colour.
  if (!child(spPr, 'a:noFill')) {
    let fill = solidFillColor(spPr, ctx);
    let gradient: GradientFill | undefined;
    const directGrad = child(spPr, 'a:gradFill');
    if (fill === undefined && directGrad)
      gradient = parseGradient(directGrad, ctx.resolveTheme);
    if (fill === undefined && gradient === undefined) {
      const themed = ctx.resolveThemeFill(child(style, 'a:fillRef'));
      if (typeof themed === 'string') fill = themed;
      else if (themed) gradient = themed;
    }
    if (fill) shape['fill'] = fill;
    if (gradient) shape['gradient'] = gradient;
  }
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
    alt: shapeDescr ?? '',
    ...(shapeTitle != null && { title: shapeTitle }),
    float: parseAnchorFloat(drawing),
    shape,
    textbox,
    rotation: xfrmRotation(spPr),
  });
}

/**
 * Everything an `a:ln` says about how a line looks, folded onto a ShapeSpec.
 *
 * Called twice per shape — once for the theme line style the shape's
 * `a:lnRef` names, then once for the shape's own outline — so a property
 * states itself or inherits, with the direct one winning. Only what the node
 * actually states is written, which is what makes the second call an override
 * rather than a reset. The colour is not here: it comes from the ref or the
 * outline's own fill, resolved at the call site.
 */
function applyLineProps(
  ln: OoxmlNode | undefined,
  shape: Record<string, unknown>,
): void {
  if (!ln) return;
  const w = attrOf(ln, 'w'); // outline width in EMU
  if (w) shape['strokeWidth'] = Math.max(1, Math.round(Number(w) / 9525));
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
}

/** Textbox content of a wps shape (wps:txbx/w:txbxContent), as paragraph node
 *  JSON the layout engine flows inside the shape's box, plus the interior
 *  padding from wps:bodyPr (EMU attrs; absent → Word's 0.1"/0.05" defaults). */
function parseTextbox(
  wsp: OoxmlNode | undefined,
  ctx: Ctx,
): {
  blocks: unknown[];
  inset?: { l: number; t: number; r: number; b: number };
  anchor?: 'ctr' | 'b';
  autofit?: boolean;
  autoWidth?: boolean;
  compatLnSpc?: boolean;
} | null {
  const blocks = txbxBlocks(
    child(child(wsp, 'wps:txbx'), 'w:txbxContent'),
    ctx,
  );
  if (!blocks) return null;
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
  // ST_TextAnchoringType: where the text block sits in the box vertically.
  // 't' is the default and needs nothing. 'just' and 'dist' stretch the gaps
  // BETWEEN paragraphs to fill the box rather than moving the block — a
  // different mechanism, not implemented, so they fall back to top. The attr
  // is read either way, which means the audit cannot tell those two apart
  // from the handled ones (the known value-level blind spot in audit.ts).
  const anchorAttr = attrOf(bodyPr, 'anchor');
  const anchor =
    anchorAttr === 'ctr' ? 'ctr' : anchorAttr === 'b' ? 'b' : undefined;
  // a:spAutoFit: "resize shape to fit text". Word ignores the stored extent
  // and regrows the height at render time (probe B5), so the flag rides the
  // model and the layout engine grows the box — the stored extent stays on
  // the node untouched, which is also exactly what Word writes back.
  const autofit = !!child(bodyPr, 'a:spAutoFit');
  // @wrap="none": the text never wraps — Word grows the box WIDE to the
  // longest line instead (probe B9: declared 108pt, drawn 315.4pt = text +
  // both insets, height untouched).
  const autoWidth = attrOf(bodyPr, 'wrap') === 'none';
  // @compatLnSpc ("simplified line spacing"): measured to change NOTHING —
  // probes B6/B7 render 200% line spacing at an identical 32.9pt step with
  // the flag on and off. Carried for the round-trip, deliberately not fed
  // into layout.
  const compatAttr = attrOf(bodyPr, 'compatLnSpc');
  const compatLnSpc = compatAttr === '1' || compatAttr === 'true';
  return {
    blocks,
    ...(inset && { inset }),
    ...(anchor && { anchor }),
    ...(autofit && { autofit }),
    ...(autoWidth && { autoWidth }),
    ...(compatLnSpc && { compatLnSpc }),
  };
}

/**
 * Block JSON of a `w:txbxContent` — shared by the modern (wps:txbx) and legacy
 * VML (v:textbox) textbox paths.
 *
 * A textbox holds a whole story, not a run of paragraphs: CT_TxbxContent is
 * `EG_BlockLevelElts`, the same group `w:body` uses, so `w:tbl`, `w:sdt` and
 * `w:customXml` are as legal in here as `w:p` is. Reading only `w:p` dropped
 * two entire tables from a hotel factsheet whose textboxes are invisible
 * frames (`filled="f" stroked="f"`) holding nothing else — the boxes rendered
 * empty and the tables vanished.
 *
 * So this hands the content to `parseBlocks`, the reader `w:body` and table
 * cells already use, rather than keeping a second, poorer content model here.
 * That brings the container unwrapping, block bookmarks and the contextual /
 * automatic spacing resolution the textbox path never had.
 */
function txbxBlocks(
  content: OoxmlNode | undefined,
  ctx: Ctx,
): unknown[] | null {
  if (!content) return null;
  // A textbox anchored in a table cell is still its own story: the text in it
  // is not "content of the table", so the enclosing table's style must not
  // reach it. The stack is emptied for the duration and put back after —
  // parseTextbox runs mid-run, deep inside parseTable.
  const enclosing = ctx.tableStyles.splice(0, ctx.tableStyles.length);
  try {
    const blocks = parseBlocks(content, ctx).map((b) => b.toJSON());
    return blocks.length > 0 ? blocks : null;
  } finally {
    ctx.tableStyles.push(...enclosing);
  }
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
  /** mso-position-horizontal-relative: what margin-left is measured from
   *  (margin | page | text | char). */
  hRelRaw?: string;
  /** mso-position-vertical-relative (margin | page | text | line). */
  vRelRaw?: string;
  /** mso-wrap-distance-left/right/top/bottom — the gaps wrapped text keeps
   *  from the shape's box (px). Absent entries take Word's defaults. */
  wrapDist?: { l?: number; r?: number; t?: number; b?: number };
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
    else if (k === 'mso-position-horizontal-relative') out.hRelRaw = v;
    else if (k === 'mso-position-vertical-relative') out.vRelRaw = v;
    else if (k.startsWith('mso-wrap-distance-')) {
      const px = cssLenToPx(v);
      if (px !== undefined) {
        const side = k.slice('mso-wrap-distance-'.length);
        out.wrapDist ??= {};
        if (side === 'left') out.wrapDist.l = px;
        else if (side === 'right') out.wrapDist.r = px;
        else if (side === 'top') out.wrapDist.t = px;
        else if (side === 'bottom') out.wrapDist.b = px;
      }
    }
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
 *  100%), so a dash stop's `d="100000"` is one stroke width. Clamped at zero:
 *  a negative dash length has no meaning. Where a negative IS meaningful —
 *  a:srcRect, where it outsets — use {@link signedPct}. */
function pctToRatio(v: string | undefined): number {
  const n = Number(v);
  return Number.isFinite(n) ? Math.max(0, n / 100000) : 0;
}

/** ST_Percentage → ratio, sign kept. */
function signedPct(v: string | undefined): number {
  const n = Number(v);
  return Number.isFinite(n) ? n / 100000 : 0;
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
  blocks: unknown[];
  inset?: { l: number; t: number; r: number; b: number };
} | null {
  const tb = child(el, 'v:textbox');
  const blocks = txbxBlocks(child(tb, 'w:txbxContent'), ctx);
  if (!blocks) return null;
  const insetAttr = attrOf(tb, 'inset');
  if (insetAttr) {
    const p = insetAttr.split(',').map((s) => {
      const px = cssLenToPx(s.trim());
      return px === undefined ? undefined : Math.round(px);
    });
    return {
      blocks,
      inset: { l: p[0] ?? 10, t: p[1] ?? 5, r: p[2] ?? 10, b: p[3] ?? 5 },
    };
  }
  return { blocks };
}

/**
 * What a VML float's `margin-left` is measured from.
 *
 * Two sources say it and they agree wherever both appear: `w10:wrap/@anchorx`
 * (ECMA-376 Part 4 — "If this attribute is omitted, then its value shall be
 * assumed to be `page`") and the `mso-position-horizontal-relative` style
 * property (margin | page | text | char). An EXPLICIT style property beats the
 * element's default, because an omitted attribute must not overrule something
 * the document states outright — a file here has `<w10:wrap type="through"/>`
 * with no anchorx next to a style that says `text`, and reading the default
 * there would move 15 shapes.
 *
 * `text` and `char` mean the text column and the anchor character; the model
 * has no such anchor base, so they keep the historical `margin` reading rather
 * than being mapped to something they are not.
 */
function vmlHRel(
  wrap: OoxmlNode | undefined,
  styleValue: string | undefined,
): 'margin' | 'page' {
  const explicit = attrOf(wrap, 'anchorx') ?? styleValue;
  if (explicit === 'page') return 'page';
  if (explicit === 'margin') return 'margin';
  // Element present but silent, and the style says nothing either: the spec's
  // default. No document in the corpus reaches this branch.
  if (!explicit && wrap) return 'page';
  return 'margin';
}

/**
 * How text flows around a VML float, from `w10:wrap/@type`.
 *
 * An ABSENT `w10:wrap` element means no wrapping at all: "If this element is
 * omitted, then no text wrapping shall be performed (i.e. the object shall be
 * presented in line with text)" (ISO/IEC 29500-4 §19.3.2.6). An absent @type
 * on a PRESENT element has no documented default in either 29500 or
 * MS-OI29500; LibreOffice — which had to match Word — leaves such a shape at
 * `WrapTextMode_THROUGH`, i.e. text flows through it, and only the four named
 * values change that. Both cases land on 'none' here.
 *
 * `tight` and `through` wrap around a polygon (`wrapcoords`) that the model
 * has no room for, so they fold into 'square' — the same approximation the
 * DrawingML path already makes for wp:wrapTight / wp:wrapThrough, and the
 * same one LibreOffice makes outside its contour flag.
 */
function vmlWrap(
  wrap: OoxmlNode | undefined,
): 'square' | 'topAndBottom' | 'none' {
  const type = attrOf(wrap, 'type');
  if (type === 'square' || type === 'tight' || type === 'through')
    return 'square';
  if (type === 'topAndBottom') return 'topAndBottom';
  return 'none';
}

/** Word's wrap gap when the shape does not state one: 114300 EMU = 9pt to the
 *  left and right, nothing above or below. The VML spec says zero, "but Word
 *  implements a non-zero value" — the figure is [MS-ODRAW] §2.3.4.9, by way of
 *  LibreOffice, and every shape in the corpus that spells it out writes
 *  exactly this. */
const VML_WRAP_DIST = { l: 12, r: 12, t: 0, b: 0 };

/** The vertical twin of {@link vmlHRel}, from `w10:wrap/@anchory` or
 *  `mso-position-vertical-relative` (margin | page | text | line). No source
 *  found states a default for `anchory`, so an absent value keeps the
 *  paragraph-relative reading instead of inventing one. `line` is the line the
 *  anchor sits on — unmodelled, and the anchor paragraph is the closest thing
 *  we have. */
function vmlVRel(
  wrap: OoxmlNode | undefined,
  styleValue: string | undefined,
): 'paragraph' | 'margin' | 'page' {
  const explicit = attrOf(wrap, 'anchory') ?? styleValue;
  return explicit === 'page'
    ? 'page'
    : explicit === 'margin'
      ? 'margin'
      : 'paragraph';
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
    const wrapEl = child(el, 'w10:wrap');
    const wrap = vmlWrap(wrapEl);
    float = {
      wrap,
      hRel: vmlHRel(wrapEl, st.hRelRaw),
      vRel: vmlVRel(wrapEl, st.vRelRaw),
    };
    // Only a wrapping shape keeps text at a distance; on a float the text
    // runs through, the gaps would describe nothing.
    if (wrap !== 'none') {
      const d = st.wrapDist ?? {};
      float['distL'] = Math.round(d.l ?? VML_WRAP_DIST.l);
      float['distR'] = Math.round(d.r ?? VML_WRAP_DIST.r);
      float['distT'] = Math.round(d.t ?? VML_WRAP_DIST.t);
      float['distB'] = Math.round(d.b ?? VML_WRAP_DIST.b);
    }
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

/** The paragraph MARK's rPr children whose value lives in `markFont` (family,
 *  size, bold, italic — re-emitted from it on export, so a font command that
 *  re-sizes the mark is what the file says too). Everything else on the mark
 *  is carried verbatim. */
const MARK_CONSUMED_RPR = new Set(['w:rFonts', 'w:sz', 'w:b', 'w:i']);

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
  'w:kern',
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
  'w:shd',
  'w:pageBreakBefore',
  'w:keepNext',
  'w:keepLines',
  'w:widowControl',
  'w:outlineLvl',
  'w:contextualSpacing',
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
  tgtFrame: string | null = null,
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
  if (href)
    marks.push(
      ctx.schema.marks['link'].create({
        href,
        ...(tgtFrame ? { targetFrame: tgtFrame } : {}),
      }),
    );
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
  tgtFrame: string | null = null,
): PMNode[] {
  const marks = runMarks(run, paraBase, ctx, href, tgtFrame);
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

/** Heading level (1–6) for a paragraph, from `w:outlineLvl` in the pPr cascade
 *  or, failing that, from a style id spelled "Heading1"…; undefined for body.
 *
 *  `w:outlineLvl` comes first because it is the property Word actually ranks
 *  paragraphs by (Navigation pane, TOC, collapse), and `lastWith` has already
 *  resolved the cascade so direct formatting beats the style. Values 0–8 are
 *  heading levels 1–9; **9 means body text** and makes the paragraph a
 *  non-heading even when its style is named like one. bc_rieng shows why that
 *  matters: `TOCHeading` is `basedOn="Heading1"` with `outlineLvl=9` — Word's
 *  way of building a heading-looking title that stays out of its own TOC.
 *
 *  The style-name regex is a fallback for documents that declare no outline
 *  level at all. It only reads the paragraph's own style id, so it is English-
 *  only; matching on the style's `w:name` is a separate job. */
function headingLevel(
  pStyleId: string | undefined,
  pPrChain: (OoxmlNode | undefined)[],
): number | undefined {
  const olLayer = lastWith(pPrChain, 'w:outlineLvl');
  const ol = child(olLayer, 'w:outlineLvl');
  if (ol) {
    const v = Number(attrOf(ol, 'w:val'));
    if (Number.isInteger(v) && v >= 0 && v <= 8) return Math.min(6, v + 1);
    if (v === 9) return undefined;
  }
  if (pStyleId) {
    const m = /^heading\s*([1-9])$/i.exec(pStyleId);
    if (m) return Math.min(6, Number(m[1]));
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

/** The math alphabet an `m:r` renders its letters in, or null for upright
 *  plain text (m:nor, or m:sty val="p"). Defaults to italic — that IS the
 *  OMML default ("use italics for characters" absent any override). */
function mathVariantOf(run: OoxmlNode): MathAlphabet | null {
  const rPr = child(run, 'm:rPr');
  if (child(rPr, 'm:nor')) return null;
  const scr = attrOf(child(rPr, 'm:scr'), 'm:val');
  const sty = attrOf(child(rPr, 'm:sty'), 'm:val') ?? 'i';
  if (scr && scr !== 'roman')
    return MATH_ALPHABETS[scr] ?? MATH_ALPHABETS['italic'];
  if (sty === 'p') return null;
  if (sty === 'b') return MATH_ALPHABETS['bold'];
  if (sty === 'bi') return MATH_ALPHABETS['bold-italic'];
  return MATH_ALPHABETS['italic'];
}

/** OMML (`m:oMath`) flattened to readable plain text — v1 keeps the equation's
 *  CONTENT, not its typesetting: `t` sub `1` → "t₁", `x` sup `2` → "x²",
 *  fractions → "num/den", radicals → "√(…)", delimiters → "(…)". Letters keep
 *  their math letterform via Unicode math alphabets (𝑥, 𝒫, ℝ). Unknown
 *  constructs concatenate their children's text so nothing is dropped. */
function flattenOmml(node: OoxmlNode): string {
  const flat = (n: OoxmlNode | undefined): string => (n ? flattenOmml(n) : '');
  switch (node.name) {
    case 'm:t':
      return node.text;
    case 'm:r': {
      const variant = mathVariantOf(node);
      const text = children(node, 'm:t')
        .map((t) => t.text)
        .join('');
      return variant ? mathLetters(text, variant) : text;
    }
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
  // The table this paragraph sits in, if any — innermost only (see Ctx).
  const tableLayer = ctx.tableStyles[ctx.tableStyles.length - 1];
  // Base for every run: docDefaults → TABLE style → paragraph style. Same
  // order as the paragraph cascade below, and the same reason: a table
  // style's w:rPr is a default for the runs inside that table.
  //
  // The toggle properties (b, i, smallCaps, strike and the eight RunProps does
  // not model) come through here as ordinary values. ISO 29500 says a style
  // that sets one TOGGLES it — "setting this property shall toggle the current
  // state of the property as specified up to this point in the hierarchy" —
  // but Word does not implement that: "Word resets the value of the toggle
  // property to the value specified by the paragraph style if a value is
  // present; for example, a value of 1 resets the property's state to True
  // instead of toggling it" (MS-OI29500 §2.1.258). A flat overwrite IS Word's
  // behaviour, so a table style asking for bold gets bold, and a paragraph
  // style that also asks for bold stays bold rather than cancelling it.
  //
  // The one place Word keeps a toggle rule is §17.7.3(a): when the document
  // DEFAULTS set the property true and it reappears down the style hierarchy,
  // Word counts levels ("true if and only if its effective value is false for
  // an even number of levels"). Not implemented, deliberately: no docDefaults
  // in any of the 18 documents in this repo turns a toggle on, so there is
  // nothing to test it against and a guess would be worse than the omission.
  //
  // Word 2007's reading (DocCompat.normalStyleYieldsToTableStyle): the DEFAULT
  // paragraph style's 11pt/12pt size does not beat the table style's size,
  // and its left justification does not beat the table style's — a stock
  // Normal must not undo a table style's look. Only those two properties, only
  // the default style, only when the table style sets them.
  const legacyTableRule =
    !!tableLayer &&
    ctx.compat.normalStyleYieldsToTableStyle &&
    styleId === ctx.styles.defaultStyleIdFor('paragraph');
  let styleRPr = ctx.styles.resolveStyle(styleId);
  if (
    legacyTableRule &&
    tableLayer.rPr.sizePt !== undefined &&
    (styleRPr.sizePt === 11 || styleRPr.sizePt === 12)
  ) {
    const { sizePt: _dropped, ...rest } = styleRPr;
    styleRPr = rest;
  }
  const paraBase = [
    ctx.styles.docDefaults,
    ...(tableLayer ? [tableLayer.rPr] : []),
    styleRPr,
  ].reduce(mergeRunProps, {} as RunProps);
  // Paragraph-property cascade, base-most first; later layers win:
  // docDefaults pPrDefault → TABLE style (for content inside a table)
  // → style chain (w:basedOn ancestors → style) → numbering lvl pPr (the
  // per-level list indent) → inline.
  //
  // The table layer's slot is the spec's: "the global default paragraph
  // properties · the table style paragraph properties · the paragraph
  // properties applied directly to a paragraph". A paragraph style therefore
  // overrides a table style, which is why a table style's formatting appears
  // to do nothing in documents whose paragraphs carry a style that sets the
  // same property.
  //
  // The numbering layer sitting AFTER the style chain rather than before it is
  // a pre-existing, deliberate deviation (Word's list indents win); untouched
  // here.
  let stylePPr = ctx.styles.resolveStylePPr(styleId);
  if (legacyTableRule && lastWith(tableLayer.pPr, 'w:jc')) {
    // Same rule for justification: the default style's LEFT gives way (any
    // other alignment it names still wins, as does an inline w:jc below).
    const styleJc = lastWith(stylePPr, 'w:jc');
    const val = attrOf(child(styleJc, 'w:jc'), 'w:val');
    if (val === 'left' || val === 'start')
      stylePPr = stylePPr.map((n) =>
        n === styleJc
          ? { ...n, children: n.children.filter((c) => c.name !== 'w:jc') }
          : n,
      );
  }
  const pPrChain: (OoxmlNode | undefined)[] = [
    ctx.styles.docDefaultsPPr,
    ...(tableLayer?.pPr ?? []),
    ...stylePPr,
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
  const spacing = resolveSpacing(pPrChain);
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
  // "Ignore spacing above and below when using identical styles" — resolved
  // against neighbours later (parseBodyBlocks / the cell loop), since only a
  // sibling walk can tell whether the paragraph beside this one shares its
  // style. Stored here as "the flag is on, nothing collapsed yet".
  const contextualSpacing = toggleLayer('w:contextualSpacing', false);
  const widowControl = toggleLayer('w:widowControl', true);
  // Where rendered inline content lands: the open field's result if we're
  // past its separate mark, the paragraph otherwise.
  const sink = (): PMNode[] =>
    field && field.phase === 'result' ? field.result : inline;
  const emitRun = (
    run: OoxmlNode,
    href: string | null,
    tgtFrame: string | null = null,
  ): void => {
    if (field && field.phase === 'result' && !field.firstResultRun)
      field.firstResultRun = run;
    sink().push(...runToInline(run, paraBase, ctx, href, tgtFrame));
  };

  // One run through the field state machine. `href` marks runs living inside
  // a w:hyperlink wrapper — TOC entries put whole PAGEREF fields there, so
  // the machine must run for them too (they used to bypass it entirely).
  const handleRun = (
    node: OoxmlNode,
    href: string | null,
    tgtFrame: string | null = null,
  ): void => {
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
      emitRun(node, href, tgtFrame);
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
      // Which frame the link was meant to open in. Kept so a save gives it
      // back; nothing here navigates, so nothing obeys it.
      const tgtFrame = attrOf(node, 'w:tgtFrame') ?? null;
      // Through the same wrappers as the paragraph loop: a link's runs can sit
      // inside w:ins or a smart tag. Non-run content a link may legally hold
      // (a nested field, math) still falls through — a separate gap.
      for (const c of effectiveChildren(node.children)) {
        if (c.name !== 'w:r') continue;
        audit.mark(c); // children() marked these before the loop changed
        handleRun(c, href, tgtFrame);
      }
    } else if (node.name === 'm:oMath' || node.name === 'm:oMathPara') {
      // OMML equations, flattened to a plain-text run (v1: content over
      // typesetting — see flattenOmml). Formatted like the first math run.
      // Deliberate wholesale flattening — the whole subtree counts consumed.
      audit.markSubtree(node);
      const text = flattenOmml(node);
      if (text.length > 0) {
        const first = findDescendant(node, 'm:r');
        const marks = runMarks(first, paraBase, ctx, null);
        // The math mark keeps the run addressable as an equation — and the
        // exporter rebuilds m:oMath from it, so the region survives a save.
        const math = ctx.schema.marks['math'];
        inline.push(
          ctx.schema.text(text, math ? [...marks, math.create()] : marks),
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
    ? parseBordersEl(child(pBdrLayer, 'w:pBdr'), CELL_SIDES, ctx.resolveTheme)
    : null;
  const paraBorders: Record<string, BorderSide> = {};
  if (pBdrSides) {
    for (const side of CELL_SIDES) {
      const s = pBdrSides[side];
      if (s) paraBorders[side] = s;
    }
  }
  // w:shd at the PARAGRAPH layer: "the shading applied to the contents of the
  // paragraph … the background color behind the paragraph, then the pattern
  // color using the mask supplied by the pattern over that background"
  // (ECMA-376). Runs and cells have had this since the start via shdFill; the
  // paragraph layer was simply never read. Only the fill is honoured — the
  // pattern masks (pct*, diagStripe, …) are not painted, so a shd carrying
  // ONLY a pattern colour resolves to nothing and stays out of the model.
  const shdLayer = lastWith(pPrChain, 'w:shd');
  const shading = shdLayer
    ? (shdFill(child(shdLayer, 'w:shd'), ctx.resolveTheme) ?? null)
    : null;

  const attrs: {
    list?: ListInfo;
    align?: Align;
    indent?: Indent;
    spacing?: Spacing;
    tabs?: { pos: number; val: string; leader?: string }[];
    bookmarks?: string[];
    field?: FieldInfo;
    pageBreakBefore?: boolean;
    contextualSpacing?: { before: boolean; after: boolean };
    keepNext?: boolean;
    keepLines?: boolean;
    widowControl?: boolean;
    heading?: number;
    styleId?: string;
    borders?: Record<string, BorderSide>;
    shading?: string;
    markFont?: {
      family?: string;
      sizePt?: number;
      bold?: boolean;
      italic?: boolean;
    };
    carry?: {
      pPr?: string;
      markRPr?: string;
      sdtStart?: SdtBoundary['start'];
      sdtEnd?: boolean;
    };
  } = {};
  // The paragraph mark's own font. The mark is "a physical character in the
  // document" (§17.3.1.29) and it sits on the paragraph's LAST line — Word
  // sizes that line to the tallest glyph on it, mark included, so a mark
  // larger than the text opens the last line and a mark no larger leaves the
  // text's own height alone. Emitted for EVERY paragraph: with no runs the
  // mark is the whole line (one corpus factsheet ends a section with two 3pt
  // marks and a 10pt one; sized from the document default they came to 17px
  // each, most of a spurious page), and with runs it is what the layout must
  // NOT seed the line from the document default instead — a rate card sets
  // 8pt text AND an 8pt mark inside exact-height rows, and a default-font
  // seed spilled every second line over the row below.
  //
  // Resolved through the run cascade, not read raw: the size can come from
  // docDefaults or the paragraph style with the mark's own rPr setting only
  // the font (`paraBase` is that cascade, minus the run layer).
  {
    const markRPr = child(pPr, 'w:rPr');
    const eff = [
      paraBase,
      ctx.styles.resolveStyle(
        attrOf(child(markRPr, 'w:rStyle'), 'w:val') ??
          ctx.styles.defaultStyleIdFor('character'),
      ),
      parseRunProps(markRPr, ctx.resolveTheme, ctx.resolveFont),
    ].reduce(mergeRunProps, {} as RunProps);
    const markFont = {
      ...(eff.fontFamily !== undefined && { family: eff.fontFamily }),
      ...(eff.sizePt !== undefined && { sizePt: eff.sizePt }),
      ...(eff.bold !== undefined && { bold: eff.bold }),
      ...(eff.italic !== undefined && { italic: eff.italic }),
    };
    if (Object.keys(markFont).length > 0) attrs.markFont = markFont;
  }
  // Carry-through: unmodelled INLINE pPr children + the paragraph mark's
  // w:rPr, preserved verbatim for export (see collectCarry).
  const carryPPr = collectCarry(pPr, CONSUMED_PPR);
  const carryMarkRPr = collectCarry(child(pPr, 'w:rPr'), MARK_CONSUMED_RPR);
  const sdtB = ctx.sdtBoundaries.get(p);
  if (carryPPr || carryMarkRPr || sdtB) {
    attrs.carry = {
      ...(carryPPr && { pPr: carryPPr }),
      ...(carryMarkRPr && { markRPr: carryMarkRPr }),
      ...(sdtB?.start && { sdtStart: sdtB.start }),
      ...(sdtB?.end && { sdtEnd: true }),
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
  // Both sides start false: whether a side actually collapses depends on the
  // neighbouring paragraph, which only the block walk can see.
  if (contextualSpacing)
    attrs.contextualSpacing = { before: false, after: false };
  if (keepNext) attrs.keepNext = true;
  if (keepLines) attrs.keepLines = true;
  if (!widowControl) attrs.widowControl = false;
  if (Object.keys(paraBorders).length > 0) attrs.borders = paraBorders;
  if (shading) attrs.shading = shading;
  return ctx.schema.nodes['paragraph'].create(attrs, inline);
}

/**
 * Custom tab stops from the cascade — ACCUMULATED, not overridden.
 *
 * Unlike the other pPr properties, tab stops from each layer add to the ones
 * below: that is the only reading under which `w:val="clear"` means anything.
 * The spec defines clear as a stop that "shall be removed"
 * (ECMA-376, w:tab/@val), and python-docx's analysis says out loud why it
 * exists — it "allows a tab stop inherited from a style, for example, to be
 * ignored". You cannot remove from a set the derived layer has already
 * replaced. Taking only the most-derived w:tabs, as this did, silently kept
 * inherited stops the document had explicitly cleared, and skipped clear
 * entries without even reading their @w:pos.
 *
 * Positions key the set in TWIPS, the unit the file writes, so a clear finds
 * exactly the stop it names — matching in px would let rounding collide two
 * neighbouring stops into one.
 *
 * 'bar' is a vertical rule rather than a stop and is not painted, so it never
 * enters the set; 'num' behaves like 'left'.
 */
function resolveTabs(
  chain: (OoxmlNode | undefined)[],
): { pos: number; val: string; leader?: string }[] | null {
  const stops = new Map<
    number,
    { pos: number; val: string; leader?: string }
  >();
  for (const layer of chain) {
    const tabs = child(layer, 'w:tabs');
    if (!tabs) continue;
    for (const tab of children(tabs, 'w:tab')) {
      const val = attrOf(tab, 'w:val') ?? 'left';
      const pos = attrOf(tab, 'w:pos');
      if (pos === undefined) continue;
      const tw = Number(pos);
      if (!Number.isFinite(tw)) continue;
      if (val === 'clear') {
        stops.delete(tw);
        continue;
      }
      if (val === 'bar') continue;
      const stop: { pos: number; val: string; leader?: string } = {
        pos: twipsToPx(tw),
        val:
          val === 'right' || val === 'center' || val === 'decimal'
            ? val
            : 'left',
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
      stops.set(tw, stop);
    }
  }
  if (stops.size === 0) return null;
  return [...stops.values()].sort((a, b) => a.pos - b.pos);
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

/**
 * Resolve w:spacing through the cascade: the last layer with each attribute
 * wins. before/after are twips→px; line is 240ths of a line for lineRule
 * 'auto' (→ multiplier), else twips→px for 'exact'/'atLeast'.
 *
 * `w:beforeAutospacing` / `w:afterAutospacing` merge per attribute like the
 * rest, and when one is on the literal value on THAT SIDE is dropped: "if this
 * attribute is specified, then any value in the before or beforeLines
 * attributes is ignored" (ECMA-376 §17.3.1.33). The number Word leaves in the
 * file next to the flag is its own cached guess, not an instruction. What the
 * gap actually becomes depends on the paragraph's neighbours, so that is
 * settled later by resolveAutoSpacing.
 */
function resolveSpacing(chain: (OoxmlNode | undefined)[]): Spacing | null {
  const out: Spacing = {};
  let beforeAuto: boolean | undefined;
  let afterAuto: boolean | undefined;
  for (const pPr of chain) {
    const sp = child(pPr, 'w:spacing');
    if (!sp) continue;
    const before = attrOf(sp, 'w:before');
    const after = attrOf(sp, 'w:after');
    const line = attrOf(sp, 'w:line');
    const rule = attrOf(sp, 'w:lineRule');
    const flag = (name: string): boolean | undefined => {
      const v = attrOf(sp, name);
      return v === undefined ? undefined : isOn(v) || v === 'on';
    };
    beforeAuto = flag('w:beforeAutospacing') ?? beforeAuto;
    afterAuto = flag('w:afterAutospacing') ?? afterAuto;
    if (before !== undefined) out.before = twipsToPx(Number(before));
    if (after !== undefined) out.after = twipsToPx(Number(after));
    if (line !== undefined) {
      const lineRule = rule === 'exact' || rule === 'atLeast' ? rule : 'auto';
      out.lineRule = lineRule;
      out.line =
        lineRule === 'auto' ? Number(line) / 240 : twipsToPx(Number(line));
    }
  }
  if (beforeAuto) {
    out.beforeAuto = true;
    delete out.before;
  }
  if (afterAuto) {
    out.afterAuto = true;
    delete out.after;
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
  if (numId === undefined || numId === '0') {
    // "A value of 0 shall never be used to point to a numbering definition
    // instance, and shall instead only be used to designate the removal of
    // numbering properties" (ECMA-376). The level that rides along with it is
    // therefore meaningless — but it is asked for, because leaving it unread
    // makes the coverage audit report a level we deliberately ignored as one
    // we failed to handle. Eight paragraphs in large_sample are this shape.
    attrOf(child(numPr, 'w:ilvl'), 'w:val');
    return null;
  }
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
  diagonals: CellDiagonals | null;
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

/**
 * w:tblPr/w:tblCellMar overrides (px), or null for Word defaults.
 *
 * The style's margins come first and the table's own element lands on top
 * SIDE BY SIDE, which is how cell margins combine everywhere else in this
 * importer — `resolveCellProps` merges `w:tcMar` per side one layer down, and
 * a table declaring only its left and right insets has no business dropping
 * the style's top and bottom. The spec talks about the whole element
 * ("if this element is omitted, then it shall inherit … from the associated
 * table style") but says nothing about the partial case, and its own fallback
 * sentence is per margin: "each margin shall use its default margin size".
 *
 * `resolveTableCellMar(undefined)` resolves the w:default table style, which
 * is where Word keeps the 108-twip side margins every unstyled table gets —
 * so consulting it unconditionally is also what makes those margins reach a
 * table that declares some of its own.
 */
function parseCellMargins(
  tbl: OoxmlNode,
  ctx: Ctx,
): { left?: number; right?: number; top?: number; bottom?: number } | null {
  const tblPr = child(tbl, 'w:tblPr');
  const styleId = attrOf(child(tblPr, 'w:tblStyle'), 'w:val');
  const fromStyle = parseMarginsEl(ctx.styles.resolveTableCellMar(styleId));
  const inline = parseMarginsEl(child(tblPr, 'w:tblCellMar'));
  if (!fromStyle && !inline) return null;
  return { ...(fromStyle ?? {}), ...(inline ?? {}) };
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
function parseBorderSide(
  el: OoxmlNode,
  resolveTheme?: ThemeResolver,
): BorderSide | false {
  const val = attrOf(el, 'w:val');
  if (val === 'none' || val === 'nil') {
    // Hidden side: its sz/color/space are meaningless, but "ask" them so the
    // coverage audit doesn't flag decoration attrs on borders we DID handle.
    attrOf(el, 'w:sz');
    attrOf(el, 'w:color');
    attrOf(el, 'w:themeColor');
    attrOf(el, 'w:space');
    return false;
  }
  const sz = Number(attrOf(el, 'w:sz') ?? '4');
  const width = Math.max(0.75, (sz / 8) * (96 / 72));
  const style = BORDER_STYLE_IN[val ?? 'single'] ?? 'solid';
  // "If the border specifies the use of a theme color via the themeColor
  // attribute, this value is superseded by the theme color value" — w:color
  // is the cached rendering of the theme slot, kept for consumers with no
  // theme part. Every attribute is read into a variable BEFORE choosing, so
  // the audit sees them all: a short-circuit here would hide whichever one
  // lost. (Read the theme even when the two agree, as they do in
  // large_sample, where accent1 IS 5B9BD5.)
  const colorAttr = attrOf(el, 'w:color');
  const themeColor = attrOf(el, 'w:themeColor');
  const themeTint = attrOf(el, 'w:themeTint');
  const themeShade = attrOf(el, 'w:themeShade');
  const themed =
    themeColor && resolveTheme
      ? resolveTheme(themeColor, themeTint, themeShade)
      : undefined;
  const color =
    themed ??
    (colorAttr && colorAttr !== 'auto'
      ? (normalizeHex(colorAttr) ?? '#b0b0b0')
      : '#b0b0b0');
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
  resolveTheme?: ThemeResolver,
): TableBorders | null {
  if (!bordersEl) return null;
  const out: TableBorders = {};
  for (const side of sides) {
    const el = child(bordersEl, `w:${side}`);
    if (!el) continue;
    out[side as keyof TableBorders] = parseBorderSide(el, resolveTheme);
  }
  return Object.keys(out).length > 0 ? out : null;
}

/** The two corner-to-corner rules a w:tcBorders may carry. Unlike the four
 *  sides these have no "explicit none" state worth modelling — a hidden
 *  diagonal is simply absent — so parseBorderSide's `false` drops out. */
function parseDiagonals(
  tcBorders: OoxmlNode | undefined,
  resolveTheme?: ThemeResolver,
): CellDiagonals | null {
  if (!tcBorders) return null;
  const out: CellDiagonals = {};
  for (const [key, tag] of [
    ['tl2br', 'w:tl2br'],
    ['br2tl', 'w:br2tl'],
  ] as const) {
    const el = child(tcBorders, tag);
    if (!el) continue;
    const side = parseBorderSide(el, resolveTheme);
    if (side) out[key] = side;
  }
  return Object.keys(out).length > 0 ? out : null;
}

/** A row's w:tblPrEx border set folded into ONE cell of that row. The cell's
 *  own w:tcBorders still win — the exception replaces the TABLE's edges, not
 *  the cell's. Which of the exception's six sides reaches a given edge is the
 *  same question the painter answers for the table: an edge on the table's
 *  outline takes top/bottom/left/right, an interior one takes insideH/V. */
function applyRowException(
  own: TableBorders | null,
  ex: TableBorders,
  at: {
    firstRow: boolean;
    lastRow: boolean;
    firstCol: boolean;
    lastCol: boolean;
  },
): TableBorders | null {
  const out: TableBorders = { ...(own ?? {}) };
  const put = (
    side: 'top' | 'bottom' | 'left' | 'right',
    from: BorderSide | false | undefined,
  ) => {
    if (out[side] === undefined && from !== undefined) out[side] = from;
  };
  put('top', at.firstRow ? ex.top : ex.insideH);
  put('bottom', at.lastRow ? ex.bottom : ex.insideH);
  put('left', at.firstCol ? ex.left : ex.insideV);
  put('right', at.lastCol ? ex.right : ex.insideV);
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
  const out = parseBordersEl(bordersEl, TABLE_SIDES, ctx.resolveTheme);
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

/** Which conditional formats a table lets through — `w:tblLook`. */
interface TblLook {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  hBand: boolean;
  vBand: boolean;
}

/** ECMA-376 1st edition had only the `w:val` bitmask; the six booleans came
 *  later. Both are still written (in agreement) by every producer in our
 *  corpus, but three files carry the bitmask alone. */
const TBL_LOOK_BITS: [keyof TblLook, number, boolean][] = [
  ['firstRow', 0x0020, true],
  ['lastRow', 0x0040, true],
  ['firstCol', 0x0080, true],
  ['lastCol', 0x0100, true],
  // noHBand/noVBand are NEGATIVE: the bit set means "do not band".
  ['hBand', 0x0200, false],
  ['vBand', 0x0400, false],
];

/** Attribute name per flag, for the modern (2nd edition) spelling. Note the
 *  spelling shift: `w:tblLook` says firstColumn, `w:tblStylePr` says firstCol. */
const TBL_LOOK_ATTRS: [keyof TblLook, string, boolean][] = [
  ['firstRow', 'w:firstRow', true],
  ['lastRow', 'w:lastRow', true],
  ['firstCol', 'w:firstColumn', true],
  ['lastCol', 'w:lastColumn', true],
  ['hBand', 'w:noHBand', false],
  ['vBand', 'w:noVBand', false],
];

/**
 * The table's conditional-formatting gates.
 *
 * Two Word behaviours the standard does not describe, both from
 * [MS-OI29500] §17.4.55:
 *
 *   - *"In Word, when the tblLook element is omitted, the bitmask of table
 *     style options on the current table is assumed to be **0x04A0**"* — first
 *     row + first column + no vertical banding, NOT the standard's 0x0000.
 *   - *"Word reads the val attribute … **if, and only if, none of the
 *     attributes** specified in this subsection are present"* — so the modern
 *     booleans win outright, and the legacy bitmask is the fallback.
 */
function tblLookFlags(tblPr: OoxmlNode | undefined): TblLook {
  const el = child(tblPr, 'w:tblLook');
  const flags: TblLook = {
    firstRow: false,
    lastRow: false,
    firstCol: false,
    lastCol: false,
    hBand: false,
    vBand: false,
  };
  const fromBits = (bits: number) => {
    for (const [key, mask, positive] of TBL_LOOK_BITS)
      flags[key] = (bits & mask) !== 0 ? positive : !positive;
  };
  if (!el) {
    fromBits(0x04a0);
    return flags;
  }
  // Read every attribute unconditionally — the audit must see the whole
  // element, not just the branch we ended up using.
  const attrs = TBL_LOOK_ATTRS.map(
    ([, name]) => [name, attrOf(el, name)] as const,
  );
  const val = attrOf(el, 'w:val');
  if (attrs.some(([, v]) => v !== undefined)) {
    for (const [key, name, positive] of TBL_LOOK_ATTRS) {
      const on = isOn(attrs.find(([n]) => n === name)?.[1]);
      flags[key] = on ? positive : !positive;
    }
    return flags;
  }
  // A present element with neither form says nothing; the schema defaults are
  // all "off", which for the two no*Band flags means banding IS allowed.
  fromBits(val === undefined ? 0x0000 : Number.parseInt(val, 16) || 0);
  return flags;
}

/** ST_OnOff, for attributes rather than elements ("1"/"true"/"on" = on). */
function isOn(v: string | undefined): boolean {
  return v === '1' || v === 'true' || v === 'on';
}

/** Where a cell sits in its table — everything the conditional formats ask. */
interface CellPos {
  row: number;
  rowCount: number;
  /** Leading grid column, and how many columns the cell spans. */
  col: number;
  colspan: number;
  colCount: number;
  /** The row carries `w:tblHeader` (a repeated heading row). */
  header: boolean;
}

/**
 * The `w:tblStylePr` types that apply to one cell, base-most FIRST so a plain
 * left-to-right merge reproduces Word's precedence:
 *
 *   *"When specified, Office applies conditional formats in the following
 *    order (therefore subsequent formats override properties on previous
 *    formats): Odd row banding, even row banding · Odd column banding, even
 *    column banding · First column, last column · First row, last row · Top
 *    left, top right, bottom left, bottom right"* — [MS-OI29500] §2.1.250.
 *
 * ISO 29500 §17.7.6.6 orders it differently (columns before rows, first row
 * before first column); we follow Word.
 *
 * Two rules that are not in the standard's prose:
 *   - A row carrying `w:tblHeader` also takes firstRow formatting (Eric White,
 *     "Assembling Paragraph and Run Properties for Cells").
 *   - Banding counts the BODY only: the first row drops out when firstRow
 *     formatting is on, the last when lastRow is. Verified against Word's own
 *     `w:cnfStyle` output on a 39-row Light Grid table — row 0 is firstRow
 *     alone, row 1 is band1Horz, and the last row keeps its band because that
 *     table's tblLook has lastRow off. Same expression for columns, which no
 *     file in the corpus exercises (every one of them sets noVBand).
 *   - Corner cells need BOTH gates: *"Top left cell – when Header Row and
 *     First Column are used"* ([MS-OI29500] §17.4.55(b)).
 */
function condTypesFor(
  pos: CellPos,
  look: TblLook,
  bands: { row: number; col: number },
): string[] {
  const firstRow = look.firstRow && (pos.row === 0 || pos.header);
  const lastRow = look.lastRow && pos.row === pos.rowCount - 1;
  const firstCol = look.firstCol && pos.col === 0;
  const lastCol = look.lastCol && pos.col + pos.colspan >= pos.colCount;
  const types: string[] = [];
  if (look.hBand && bands.row > 0 && !firstRow && !lastRow) {
    const body = pos.row - (look.firstRow ? 1 : 0);
    types.push(
      Math.floor(body / bands.row) % 2 === 0 ? 'band1Horz' : 'band2Horz',
    );
  }
  if (look.vBand && bands.col > 0 && !firstCol && !lastCol) {
    const body = pos.col - (look.firstCol ? 1 : 0);
    types.push(
      Math.floor(body / bands.col) % 2 === 0 ? 'band1Vert' : 'band2Vert',
    );
  }
  if (firstCol) types.push('firstCol');
  if (lastCol) types.push('lastCol');
  if (firstRow) types.push('firstRow');
  if (lastRow) types.push('lastRow');
  if (firstRow && firstCol) types.push('nwCell');
  if (firstRow && lastCol) types.push('neCell');
  if (lastRow && firstCol) types.push('swCell');
  if (lastRow && lastCol) types.push('seCell');
  return types;
}

/** The paragraph/run defaults a table (or one cell of it) hands its content. */
type TableLayer = Ctx['tableStyles'][number];

/** A conditional branch that reaches a cell, with the @w:type it came from —
 *  the type names the REGION its insideH/insideV edges belong to. */
interface CellBranch {
  type: string;
  layer: CondLayer;
}

/** The conditional branches that reach one cell, in application order. */
function cellBranches(tblCond: TableCond, pos: CellPos): CellBranch[] {
  if (tblCond.cond.size === 0) return [];
  return condTypesFor(pos, tblCond.look, tblCond.bands).flatMap((type) => {
    const layer = tblCond.cond.get(type);
    return layer ? [{ type, layer }] : [];
  });
}

/**
 * The rectangle of grid cells a branch styles. It decides what `insideH` and
 * `insideV` mean inside that branch: they are the edges BETWEEN cells of the
 * region, so a cell on the region's boundary takes top/bottom/left/right there
 * and an interior one takes the inside pair — the same question
 * `applyRowException` answers for a row's tblPrEx, one region smaller.
 */
interface BranchRegion {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

function branchRegion(
  type: string,
  pos: CellPos,
  look: TblLook,
  bands: { row: number; col: number },
): BranchRegion {
  const lastRow = pos.rowCount - 1;
  const lastCol = pos.colCount - 1;
  // Banding counts the body only, so a band's bounds are measured from there.
  const bodyTop = look.firstRow ? 1 : 0;
  const bodyBottom = lastRow - (look.lastRow ? 1 : 0);
  const bodyLeft = look.firstCol ? 1 : 0;
  const bodyRight = lastCol - (look.lastCol ? 1 : 0);
  const wholeRows = { rowStart: 0, rowEnd: lastRow };
  const wholeCols = { colStart: 0, colEnd: lastCol };
  const band = (v: number, from: number, to: number, size: number) => {
    const start = from + Math.floor((v - from) / size) * size;
    return { start, end: Math.min(start + size - 1, to) };
  };
  switch (type) {
    case 'firstRow':
      return { rowStart: 0, rowEnd: 0, ...wholeCols };
    case 'lastRow':
      return { rowStart: lastRow, rowEnd: lastRow, ...wholeCols };
    case 'firstCol':
      return { ...wholeRows, colStart: 0, colEnd: 0 };
    case 'lastCol':
      return { ...wholeRows, colStart: lastCol, colEnd: lastCol };
    case 'band1Horz':
    case 'band2Horz': {
      const b = band(pos.row, bodyTop, bodyBottom, Math.max(1, bands.row));
      return { rowStart: b.start, rowEnd: b.end, ...wholeCols };
    }
    case 'band1Vert':
    case 'band2Vert': {
      const b = band(pos.col, bodyLeft, bodyRight, Math.max(1, bands.col));
      return { ...wholeRows, colStart: b.start, colEnd: b.end };
    }
    // The four corner branches style exactly one cell.
    default:
      return {
        rowStart: pos.row,
        rowEnd: pos.row,
        colStart: pos.col,
        colEnd: pos.col + pos.colspan - 1,
      };
  }
}

/**
 * The cell borders the conditional branches contribute, later branch winning.
 * Sits BELOW the row's tblPrEx and the cell's own w:tcBorders: a table style
 * cannot overrule what the document says about this particular cell.
 */
function branchBorders(
  branches: CellBranch[],
  pos: CellPos,
  look: TblLook,
  bands: { row: number; col: number },
  ctx: Ctx,
): TableBorders | null {
  const out: TableBorders = {};
  for (const { type, layer } of branches) {
    const region = branchRegion(type, pos, look, bands);
    const at = {
      top: pos.row === region.rowStart,
      bottom: pos.row === region.rowEnd,
      left: pos.col === region.colStart,
      right: pos.col + pos.colspan - 1 >= region.colEnd,
    };
    for (const tcPr of layer.tcPr) {
      const set = parseBordersEl(
        child(tcPr, 'w:tcBorders'),
        TABLE_SIDES,
        ctx.resolveTheme,
      );
      if (!set) continue;
      const put = (
        side: 'top' | 'bottom' | 'left' | 'right',
        from: BorderSide | false | undefined,
      ) => {
        if (from !== undefined) out[side] = from;
      };
      put('top', at.top ? set.top : set.insideH);
      put('bottom', at.bottom ? set.bottom : set.insideH);
      put('left', at.left ? set.left : set.insideV);
      put('right', at.right ? set.right : set.insideV);
    }
  }
  return Object.keys(out).length > 0 ? out : null;
}

/**
 * The table layer as one cell sees it: the table style's own w:pPr/w:rPr with
 * the branches that reach this cell merged on top. A cell that takes no branch
 * gets the table's layer back BY IDENTITY, so `withTableLayer` can skip the
 * push entirely.
 */
function layerWithBranches(
  base: TableLayer | undefined,
  branches: CellBranch[],
): TableLayer | undefined {
  if (!base || branches.length === 0) return base;
  return {
    pPr: [...base.pPr, ...branches.flatMap((b) => b.layer.pPr)],
    rPr: branches.reduce((acc, b) => mergeRunProps(acc, b.layer.rPr), base.rPr),
  };
}

/** Cell padding, in px per side. */
type CellPadding = {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
};

/**
 * Resolve one cell's `w:tcPr`-borne properties down the layers Word stacks:
 * the table style's own w:tcPr, the conditional branches, then the cell's own.
 *
 * Later layers win per property, and an explicit no-shading (`w:fill="auto"`,
 * which shdFill reads as "no colour") CLEARS an inherited fill rather than
 * being skipped. w:tcMar merges per SIDE, the way cell margins do everywhere
 * else, so a branch that only widens the left inset keeps the rest.
 */
function resolveCellProps(
  layers: OoxmlNode[],
  ctx: Ctx,
): {
  background: string | null;
  vAlign: 'center' | 'bottom' | null;
  padding: CellPadding | null;
} {
  let background: string | null = null;
  let vAlign: 'center' | 'bottom' | null = null;
  let padding: CellPadding | null = null;
  for (const layer of layers) {
    const shd = child(layer, 'w:shd');
    if (shd) background = shdFill(shd, ctx.resolveTheme) ?? null;
    const vAlignEl = child(layer, 'w:vAlign');
    if (vAlignEl) {
      const v = attrOf(vAlignEl, 'w:val');
      // ST_VerticalJc also has "both" (vertically justified), which renders as
      // top for the single-paragraph cells it ever appears on.
      vAlign = v === 'center' || v === 'bottom' ? v : null;
    }
    const mar = parseMarginsEl(child(layer, 'w:tcMar'));
    if (mar) padding = { ...(padding ?? {}), ...mar };
  }
  return { background, vAlign, padding };
}

/** Run `fn` with `layer` as the innermost table layer. */
function withTableLayer<T>(
  ctx: Ctx,
  layer: TableLayer | undefined,
  fn: () => T,
): T {
  if (!layer || layer === ctx.tableStyles[ctx.tableStyles.length - 1])
    return fn();
  ctx.tableStyles.push(layer);
  try {
    return fn();
  } finally {
    ctx.tableStyles.pop();
  }
}

function parseTable(tbl: OoxmlNode, ctx: Ctx): PMNode {
  // The table style's paragraph/run defaults are rolled up ONCE here and left
  // on the stack for every paragraph in every cell to read. try/finally
  // because the pop has to happen even if a malformed table throws — a leaked
  // entry would style the paragraphs that follow the table.
  const styleId = attrOf(child(child(tbl, 'w:tblPr'), 'w:tblStyle'), 'w:val');
  ctx.tableStyles.push({
    pPr: ctx.styles.resolveTableStylePPr(styleId),
    rPr: ctx.styles.resolveTableStyleRPr(styleId),
  });
  try {
    return parseTableRows(tbl, ctx, {
      styleId,
      cond: ctx.styles.resolveTableStyleCond(styleId),
      look: tblLookFlags(child(tbl, 'w:tblPr')),
      bands: ctx.styles.resolveTableBandSizes(styleId),
      styleTcPr: ctx.styles.resolveTableStyleTcPr(styleId),
    });
  } finally {
    ctx.tableStyles.pop();
  }
}

/** A table's conditional formatting, resolved once for the whole table: the
 *  branch definitions, the gates that select them, and the band sizes. */
interface TableCond {
  /** w:tblPr/w:tblStyle of this table, for the resolvers that need it again. */
  styleId: string | undefined;
  cond: Map<string, CondLayer>;
  look: TblLook;
  bands: { row: number; col: number };
  /** The table style's OWN w:tcPr chain (basedOn ancestors → style), the
   *  base-most cell layer — below the conditional branches and the cell's own
   *  w:tcPr. */
  styleTcPr: OoxmlNode[];
}

function parseTableRows(tbl: OoxmlNode, ctx: Ctx, tblCond: TableCond): PMNode {
  const grid = children(child(tbl, 'w:tblGrid'), 'w:gridCol').map((c) =>
    Number(attrOf(c, 'w:w') ?? '0'),
  );
  const colPx = tableColumnWidths(tbl, grid, ctx);

  // Phase 1: logical grid — every w:tc (incl. vMerge-continue placeholders),
  // tracking each cell's starting grid column.
  const rowEls = children(tbl, 'w:tr');
  const spanTotals = rowEls.map((tr) =>
    children(tr, 'w:tc').reduce(
      (n, tc) =>
        n +
        (Number(
          attrOf(child(child(tc, 'w:tcPr'), 'w:gridSpan'), 'w:val') ?? '1',
        ) || 1),
      0,
    ),
  );
  // "Last column" for conditional formatting is the TABLE's right edge, not
  // the row's — a row whose cells are merged into one still has its last cell
  // sitting at the table edge.
  const colCount = Math.max(grid.length, ...spanTotals, 0);
  const tableLayer = ctx.tableStyles[ctx.tableStyles.length - 1];
  const logicalRows: LogicalCell[][] = rowEls.map((tr, rowIdx) => {
    const cells: LogicalCell[] = [];
    let col = 0;
    // w:tblPrEx — "table properties which shall be applied to the contents of
    // this row IN PLACE OF the table properties" (ECMA-376). Word writes it
    // when two tables are merged and the second one's look has to survive.
    // The painter only knows table edges and per-cell overrides, so the row's
    // exception is resolved HERE into the cells it covers, using the same
    // outer-vs-inside mapping the painter applies to the table's own set.
    const rowEx = parseBordersEl(
      child(child(tr, 'w:tblPrEx'), 'w:tblBorders'),
      TABLE_SIDES,
      ctx.resolveTheme,
    );
    const rowCols = spanTotals[rowIdx];
    const rowHeader = isToggleOn(child(child(tr, 'w:trPr'), 'w:tblHeader'));
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
      const pos: CellPos = {
        row: rowIdx,
        rowCount: rowEls.length,
        col,
        colspan,
        colCount,
        header: rowHeader,
      };
      const branches = cellBranches(tblCond, pos);
      const { background, vAlign, padding } = resolveCellProps(
        [
          ...tblCond.styleTcPr,
          ...branches.flatMap((b) => b.layer.tcPr),
          ...(tcPr ? [tcPr] : []),
        ],
        ctx,
      );
      const ownBorders = parseBordersEl(
        child(tcPr, 'w:tcBorders'),
        CELL_SIDES,
        ctx.resolveTheme,
      );
      const withEx = rowEx
        ? applyRowException(ownBorders, rowEx, {
            firstRow: rowIdx === 0,
            lastRow: rowIdx === rowEls.length - 1,
            firstCol: col === 0,
            lastCol: col + colspan >= rowCols,
          })
        : ownBorders;
      // The table style's conditional borders are the lowest layer: they fill
      // the sides nothing else claimed.
      const fromStyle = branchBorders(
        branches,
        pos,
        tblCond.look,
        tblCond.bands,
        ctx,
      );
      const borders = fromStyle ? { ...fromStyle, ...(withEx ?? {}) } : withEx;
      const diagonals = parseDiagonals(
        child(tcPr, 'w:tcBorders'),
        ctx.resolveTheme,
      );
      const carry = collectCarry(tcPr, CONSUMED_TCPR);
      // Conditional formatting is a per-CELL layer: which branches reach this
      // cell depends on where it sits. Cells that take none reuse the table's
      // own layer object — identity, no re-merge — so a document like khbd
      // (275 tables, 4704 paragraphs, zero tblStylePr) pays nothing for this.
      const content = withTableLayer(
        ctx,
        layerWithBranches(tableLayer, branches),
        () => parseBlocks(tc, ctx),
      );
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
        diagonals,
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
    // w:tblPrEx rides along verbatim MINUS its w:tblBorders, which is already
    // resolved into this row's cells (see applyRowException) and would be
    // written twice if it came back here as well. What survives is what the
    // model has nowhere else to put: w:jc, the per-row table alignment, and
    // w:tblLook — the TABLE-level tblLook already drove the conditional
    // formats at import (condTypesFor), so a row-level copy only needs to
    // round-trip.
    const exCarry = collectCarry(
      child(tr, 'w:tblPrEx'),
      new Set(['w:tblBorders']),
    );
    return { header, height, cantSplit, carry, exCarry };
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
            diagonals: cell.diagonals,
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
    if (rp.carry || rp.exCarry)
      rowAttrs['carry'] = {
        ...(rp.carry && { trPr: rp.carry }),
        ...(rp.exCarry && { tblPrEx: rp.exCarry }),
      };
    return ctx.schema.nodes['table_row'].create(
      Object.keys(rowAttrs).length > 0 ? rowAttrs : null,
      emitted.length > 0 ? emitted : [emptyCell(ctx)],
    );
  });

  const cellPadding = parseCellMargins(tbl, ctx);
  const borders = parseTableBorders(tbl, ctx);
  // Alignment: the table's own w:jc, else the table style's — the same
  // fallback parseTableBorders and parseCellMargins already use.
  //
  // Not implemented, and measured rather than assumed: PER-ROW alignment
  // (w:trPr/w:jc, w:tblPrEx/w:jc), which ECMA-376 defines as "the alignment of
  // a single row in the parent table with respect to the text margins".
  // Honouring it means shifting one row's cells independently of the grid,
  // through the cell drafts, the borders, hit-testing and the painter. Across
  // all 18 documents in the repo exactly ONE table would look different: 22
  // tables carry row-level jc, 21 of them declare the same value the table
  // itself does, and the odd one out is a 13-row table in a lesson plan where
  // 5 rows say center and the table says nothing.
  const jc =
    attrOf(child(child(tbl, 'w:tblPr'), 'w:jc'), 'w:val') ??
    ctx.styles.resolveTableJc(tblCond.styleId);
  const attrs: Record<string, unknown> = {};
  if (cellPadding) attrs['cellPadding'] = cellPadding;
  if (borders) attrs['borders'] = borders;
  if (jc === 'center' || jc === 'right' || jc === 'end')
    attrs['align'] = jc === 'end' ? 'right' : jc;
  // w:tblInd — the table's own, else its style's — resolved here to where
  // the leading BORDER lands relative to the text margin, so the layout has
  // one number to add and no compat mode to know: to-border in Word 2013+
  // documents, to-the-first-cell's-text (border = indent − left cell margin)
  // in older ones (see DocCompat.tableIndentToBorder). Only a left-aligned table
  // is indented — w:jc center/right position the table on their own. The
  // element itself still rides carry, verbatim, for the save.
  // Both layers of the cascade are read — the style's tblInd is what the
  // table inherits, and the table's own overrides it.
  const indentOf = (el: OoxmlNode | undefined): number | undefined => {
    if (!el) return undefined;
    const type = attrOf(el, 'w:type') ?? 'dxa';
    if (type === 'nil') return 0;
    if (type !== 'dxa') return undefined; // pct/auto: not modelled
    return twipsToPx(Number(attrOf(el, 'w:w') ?? 0));
  };
  const styleInd = indentOf(ctx.styles.resolveTableInd(tblCond.styleId));
  const ownInd = indentOf(child(child(tbl, 'w:tblPr'), 'w:tblInd'));
  if (!attrs['align']) {
    // No element at all is an indent of 0 — which the two rules still place
    // differently, so it is resolved rather than left to a layout default.
    const value = ownInd ?? styleInd ?? 0;
    // The older rule measures to the text: the border sits one left cell
    // margin further out (Word's 0.08" default when the table names none).
    const leftPad = cellPadding?.left ?? twipsToPx(108);
    attrs['indent'] = ctx.compat.tableIndentToBorder ? value : value - leftPad;
  }
  // Carry-through: tblStyle/tblW/tblLayout/tblInd/tblLook/… survive the save.
  const tblCarry = collectCarry(child(tbl, 'w:tblPr'), CONSUMED_TBLPR);
  const tblSdt = ctx.sdtBoundaries.get(tbl);
  if (tblCarry || tblSdt)
    attrs['carry'] = {
      ...(tblCarry && { tblPr: tblCarry }),
      ...(tblSdt?.start && { sdtStart: tblSdt.start }),
      ...(tblSdt?.end && { sdtEnd: true }),
    };
  return ctx.schema.nodes['table'].create(
    Object.keys(attrs).length > 0 ? attrs : null,
    rows.length > 0
      ? rows
      : [ctx.schema.nodes['table_row'].create(null, [emptyCell(ctx)])],
  );
}

/** Walk an element's children in document order, mapping w:p / w:tbl to blocks. */
/**
 * A `w:bookmarkStart` sitting BETWEEN blocks rather than inside a paragraph.
 *
 * The element belongs to EG_RangeMarkupElements, which EG_BlockLevelElts
 * includes, so this is ordinary valid markup — Word writes it for TOC anchors
 * ahead of a content control. parseParagraph only ever saw the ones inside a
 * w:p, so a link to one of these resolved to nothing.
 *
 * Returns true when it consumed the node, so the caller can skip it.
 */
function takeBlockBookmark(node: OoxmlNode, pending: string[]): boolean {
  if (node.name !== 'w:bookmarkStart') return false;
  audit.mark(node);
  const name = attrOf(node, 'w:name');
  // colFirst/colLast narrow a bookmark to a span of table columns. We anchor
  // to a block, so they are asked for and dropped rather than left to read as
  // an unsupported gap.
  attrOf(node, 'w:colFirst');
  attrOf(node, 'w:colLast');
  // Word's own cursor bookmark is noise, same as in parseParagraph.
  if (name && name !== '_GoBack') pending.push(name);
  return true;
}

/** Fold the bookmarks collected before a block onto it. They go to the NEXT
 *  paragraph — that is where a reader following the link wants to land — and
 *  ahead of the paragraph's own names, matching document order. Tables carry
 *  no bookmarks attr, so a bookmark before one waits for the paragraph after
 *  it; `pending` is emptied in place. */
function withBookmarks(block: PMNode, pending: string[]): PMNode {
  if (pending.length === 0 || block.type.name !== 'paragraph') return block;
  const own = (block.attrs['bookmarks'] as string[] | null) ?? [];
  const merged = [...pending, ...own];
  pending.length = 0;
  return block.type.create(
    { ...block.attrs, bookmarks: merged },
    block.content,
    block.marks,
  );
}

/** Bookmarks left over at the end of a container (nothing follows them) go to
 *  the last paragraph instead — the nearest thing a link can reach. */
function attachTrailingBookmarks(blocks: PMNode[], pending: string[]): void {
  if (pending.length === 0) return;
  for (let i = blocks.length - 1; i >= 0; i--) {
    if (blocks[i].type.name === 'paragraph') {
      const own = (blocks[i].attrs['bookmarks'] as string[] | null) ?? [];
      blocks[i] = blocks[i].type.create(
        { ...blocks[i].attrs, bookmarks: [...own, ...pending] },
        blocks[i].content,
        blocks[i].marks,
      );
      break;
    }
  }
  pending.length = 0;
}

function parseBlocks(parent: OoxmlNode, ctx: Ctx): PMNode[] {
  const blocks: PMNode[] = [];
  const styleKeys: (string | null)[] = [];
  const pending: string[] = [];
  for (const node of unwrapContainers(parent.children, ctx.sdtBoundaries)) {
    if (takeBlockBookmark(node, pending)) continue;
    if (node.name === 'w:p' || node.name === 'w:tbl') audit.mark(node);
    if (node.name === 'w:p') {
      blocks.push(withBookmarks(parseParagraph(node, ctx), pending));
      styleKeys.push(paraStyleKey(node, ctx));
    } else if (node.name === 'w:tbl') {
      blocks.push(parseTable(node, ctx));
      styleKeys.push(null);
    }
  }
  attachTrailingBookmarks(blocks, pending);
  // Contextual spacing is resolved per container: a cell's first paragraph
  // has no predecessor to collapse against, exactly as Word treats it.
  return resolveAutoSpacing(resolveContextualSpacing(blocks, styleKeys), ctx);
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

/**
 * Column flow from a section's `w:cols`, by Word's rules rather than the
 * standard's — they differ here, and the corpus cannot tell them apart.
 * count defaults to 1; the gap is `w:space` (twips→px, Word's default 720 =
 * 0.5in); unequal sections carry each column's own width instead.
 */
function parseColumns(sectPr: OoxmlNode | undefined): ColumnConfig {
  const cols = sectPr && child(sectPr, 'w:cols');
  // MS-OI29500 §17.6.4(a): Word restricts @w:num to 1..45.
  const num = Number(attrOf(cols, 'w:num') ?? '1');
  const count = Math.min(45, Math.max(1, Number.isNaN(num) ? 1 : num));
  const spaceTw = Number(attrOf(cols, 'w:space') ?? '720');
  const gap = twipsToPx(Number.isNaN(spaceTw) ? 720 : spaceTw);

  // Equal width is Word's default for an ABSENT attribute — the standard
  // declares none, "Word uses a default value of true" (MS-OI29500
  // §17.6.4(b)) — and when it is on, "the col elements are ignored".
  const eq = attrOf(cols, 'w:equalWidth');
  if (eq === undefined || !['false', '0', 'off'].includes(eq))
    return { count, gap };

  // Unequal: each w:col carries its own width and the space AFTER it, and
  // the cols/@w:space above is disregarded.
  //
  // The count still comes from @w:num, NOT from the number of children. The
  // standard says the opposite — num "is ignored in favor of the number of
  // child col elements" — but Word "requires that the value of the num
  // attribute matches the number of child col elements. If the num attribute
  // is not specified, then Word assumes a value of 1" (MS-OI29500 §17.6.4(c)).
  // Every unequal section in the corpus satisfies that requirement, so no
  // fixture can tell the two readings apart; following the standard here
  // would be wrong in a way the tests could never catch.
  const colEls = cols ? children(cols, 'w:col') : [];
  if (colEls.length !== count) return { count, gap };
  const list = colEls.map((c) => {
    const w = Number(attrOf(c, 'w:w') ?? '0');
    const s = Number(attrOf(c, 'w:space') ?? '0');
    return {
      width: twipsToPx(Number.isNaN(w) ? 0 : w),
      space: twipsToPx(Number.isNaN(s) ? 0 : s),
    };
  });
  // The widths ride through verbatim even when they happen to be uniform —
  // re-deriving them from a gap would silently rescale a section whose
  // declared widths do not add up to the text area. Whether the columns
  // differ is the layout's question, not this one's.
  return { count, gap, cols: list };
}

/** A continuous section break switches columns mid-page; every other type
 *  (next/odd/even page, or unspecified) starts the section on a new page. */
function sectionStartsNewPage(sectPr: OoxmlNode | undefined): boolean {
  return attrOf(child(sectPr, 'w:type'), 'w:val') !== 'continuous';
}

/** Parse the body into blocks while recording section boundaries. A w:p whose
 *  w:pPr carries a w:sectPr ends a section (with that sectPr's columns); the
 *  trailing w:body/w:sectPr ends the final section. */
/**
 * Word's HTML auto spacing, in px per side.
 *
 * The amount is 14pt, and three independent implementations say so:
 *
 *   - *"When you set paragraph Space Before and Space After to Auto, Microsoft
 *     Word adds 14 points spacing between paragraphs automatically"* —
 *     Aspose.Words, ParagraphFormat.SpaceAfterAuto (an engine built to
 *     reproduce Word's layout).
 *   - Word's own help: *"a standard amount of spacing… usually about 14 points
 *     for a 12-point font size, similar to browsers"*.
 *   - LibreOffice's .docx importer, writerfilter/DomainMapper.cxx, which picks
 *     the number outright:
 *
 *       // See SwWW8ImplReader::GetParagraphAutoSpace() on why these are 100 and 280
 *       default_spacing = 100;
 *       if (!GetSettingsTable()->GetDoNotUseHTMLParagraphAutoSpacing()) {
 *           if (GetView() == ST_View_web) default_spacing = 49;
 *           else                          default_spacing = 280;   // ← 14pt
 *       }
 *
 * The dissenting source is one answer on MS Q&A describing Word 2013+ as 0em
 * before and 1em after (scaled to the font size) — one voice against three,
 * and our corpus cannot referee it: every paragraph that ends up with auto
 * spacing on is alone in a table cell, where the suppression rules below cut
 * the gap to 0 under either reading. Still one constant, if it ever has to
 * change.
 *
 * LibreOffice's 49-twip "web view" branch has no counterpart here — we have no
 * web layout mode, and its own comment calls 49 a leftover to be removed.
 */
const AUTO_SPACING_PX = twipsToPx(280);

/** `w:doNotUseHTMLParagraphAutoSpacing`: "5 points of spacing before and 10
 *  points after", fixed, instead of the HTML emulation. The asymmetry is the
 *  standard's; LibreOffice uses 100 twips on BOTH sides for this case, which
 *  is 5pt/5pt. We follow the standard, since that is the normative text and
 *  nothing here can test either: no document in the repo sets the flag, so
 *  this path has unit tests and nothing else. */
const AUTO_SPACING_FIXED = { before: twipsToPx(100), after: twipsToPx(200) };

/**
 * Turn `beforeAuto`/`afterAuto` into actual gaps, which needs the paragraph's
 * neighbours — the same reason resolveContextualSpacing runs here.
 *
 * Auto spacing exists only at a boundary between two paragraphs. Word's own
 * documentation lists the cases as "no spacing before the first paragraph in a
 * document", "…in a table cell", "no spacing after the last paragraph in a
 * table cell", and none between list items of the same list; Aspose states the
 * same set the other way round ("auto spacing is only applied between two
 * neighboring paragraphs… spacing is added only after the last item in the
 * list"). One rule covers all of them: a side gets the spacing only if a
 * paragraph of this same container sits on that side, and that neighbour is not
 * an item of the same list.
 *
 * Two things this deliberately does not model: a paragraph next to a TABLE
 * rather than a paragraph gets 0 (Aspose says Word does add spacing after a
 * table, but our model has nowhere to hang a table's own spacing), and the
 * nested-list nuance — "in a nested bulleted or numbered list spacing is not
 * added" — is only honoured where the neighbour shares the numbering.
 */
function resolveAutoSpacing(blocks: PMNode[], ctx: Ctx): PMNode[] {
  const auto = (n: PMNode) => {
    const sp = n.attrs['spacing'] as Spacing | null;
    return sp && (sp.beforeAuto || sp.afterAuto) ? sp : null;
  };
  if (!blocks.some((b) => b.type.name === 'paragraph' && auto(b)))
    return blocks;
  const listOf = (n: PMNode | undefined) =>
    n?.type.name === 'paragraph'
      ? ((n.attrs['list'] as { numId: string } | null)?.numId ?? null)
      : undefined;
  return blocks.map((node, i) => {
    const sp = node.type.name === 'paragraph' ? auto(node) : null;
    if (!sp) return node;
    const mine = listOf(node);
    // undefined = no paragraph on that side at all; null = not a list item.
    const neighbour = (at: number) => {
      const side = listOf(blocks[at]);
      if (side === undefined) return 0;
      if (mine !== null && mine !== undefined && side === mine) return 0;
      return null; // spacing applies
    };
    const amount = (side: 'before' | 'after', at: number) =>
      ctx.compat.htmlAutoSpacing
        ? (neighbour(at) ?? AUTO_SPACING_PX)
        : AUTO_SPACING_FIXED[side];
    return node.type.create(
      {
        ...node.attrs,
        spacing: {
          ...sp,
          ...(sp.beforeAuto ? { before: amount('before', i - 1) } : {}),
          ...(sp.afterAuto ? { after: amount('after', i + 1) } : {}),
        },
      },
      node.content,
      node.marks,
    );
  });
}

/**
 * Resolve `w:contextualSpacing` against neighbours: *"any space specified
 * before or after this paragraph … should not be applied when the preceding
 * and following paragraphs are of the same paragraph style"*. Each paragraph's
 * own flag governs its own two sides; the neighbour only has to SHARE its
 * style, flag or not.
 *
 * This runs over assembled blocks because it is the one place both sides of a
 * boundary are visible. Doing it in layout instead would mean carrying every
 * paragraph's style name in the model and folding the neighbour into the
 * paragraph cache key — a stale-layout bug waiting to happen.
 *
 * Known deviation, deliberate: the spec subtracts the flagged side from the
 * collapsed gap, whereas this drops that side outright. The two agree whenever
 * both neighbours carry the flag — which is always, when it comes from the
 * shared style. They differ only if the flag is set INLINE on exactly one of
 * two same-styled paragraphs (spec's worked example: after 10pt vs before
 * 12pt → 2pt, where this yields 12pt). No document in our corpus declares it
 * inline. Closing that gap means deferring space-after inside the placer,
 * which would have to join the page-cache checkpoint state.
 */
function resolveContextualSpacing(
  blocks: PMNode[],
  styleKeys: (string | null)[],
): PMNode[] {
  return blocks.map((node, i) => {
    const flag = node.attrs['contextualSpacing'] as {
      before: boolean;
      after: boolean;
    } | null;
    if (!flag) return node;
    const mine = styleKeys[i];
    const before = mine !== null && i > 0 && styleKeys[i - 1] === mine;
    const after =
      mine !== null && i + 1 < blocks.length && styleKeys[i + 1] === mine;
    if (!before && !after) return node;
    return node.type.create(
      { ...node.attrs, contextualSpacing: { before, after } },
      node.content,
      node.marks,
    );
  });
}

/** The style a paragraph resolves to, for "same paragraph style" comparisons.
 *  Unstyled content still lands on the default style, so two bare paragraphs
 *  count as identical. Non-paragraph blocks get null and never match. */
function paraStyleKey(node: OoxmlNode, ctx: Ctx): string | null {
  if (node.name !== 'w:p') return null;
  return (
    attrOf(child(child(node, 'w:pPr'), 'w:pStyle'), 'w:val') ??
    ctx.styles.defaultStyleIdFor('paragraph') ??
    'Normal'
  );
}

function parseBodyBlocks(
  body: OoxmlNode,
  ctx: Ctx,
): { blocks: PMNode[]; sections: SectionConfig[] } {
  const blocks: PMNode[] = [];
  const styleKeys: (string | null)[] = [];
  const sections: SectionConfig[] = [];
  let start = 0;
  const pendingBookmarks: string[] = [];
  for (const node of unwrapContainers(body.children, ctx.sdtBoundaries)) {
    if (takeBlockBookmark(node, pendingBookmarks)) continue;
    if (node.name === 'w:p' || node.name === 'w:tbl') audit.mark(node);
    if (node.name === 'w:p') {
      blocks.push(withBookmarks(parseParagraph(node, ctx), pendingBookmarks));
      styleKeys.push(paraStyleKey(node, ctx));
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
      styleKeys.push(null);
    }
  }
  attachTrailingBookmarks(blocks, pendingBookmarks);
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
  return {
    blocks: resolveAutoSpacing(
      resolveContextualSpacing(blocks, styleKeys),
      ctx,
    ),
    sections,
  };
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
      // Separator/continuation chrome is identified by its TYPE, not its id:
      // Word does stamp -1/0 on the separators, but it also always types
      // them — while Google Docs exports number REAL footnotes from 0 with
      // no separators at all, and an id-based guard silently ate the first
      // footnote of every such file.
      if (id === undefined || (type && type !== 'normal')) {
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
    display: { footnote: 0, endnote: 0 },
    ref(kind, id, custom) {
      const num = ++reg.counter[kind];
      const display = custom ? undefined : ++reg.display[kind];
      reg.refs.push({ kind, id, num, ...(display && { display }) });
      return { num, display };
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
 *  display number. The marker is the bare superscript number — measured:
 *  Word's own body marker is just the w:footnoteRef glyph; the space that
 *  usually follows is AUTHORED text inside the note. A custom-marked note
 *  (display == null) gets no marker at all: its body already starts with
 *  the author's glyph. Shared by the appended endnote section and the
 *  page-bottom footnote map. */
function noteBlocks(
  note: OoxmlNode,
  display: number | null,
  ctx: Ctx,
): PMNode[] {
  const blocks = parseBlocks(note, ctx);
  if (display == null) return blocks;
  const marker = ctx.schema.text(`${display}`, [
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
  for (const { kind, id, num, display } of ctx.notes.refs) {
    if (kind !== 'footnote') continue;
    const note = ctx.notes.bodies.footnote.get(id);
    if (note)
      out[num] = storyDoc(ctx, noteBlocks(note, display ?? null, ctx), null);
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
  for (const { id, display } of endnotes) {
    const note = ctx.notes.bodies.endnote.get(id);
    if (note) out.push(...noteBlocks(note, display ?? null, ctx));
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
  // The compat profile rides every story doc (body, headers, notes) so a
  // layout rule that depends on it can ask the doc, wherever it is laid out.
  if (ctx.schema.nodes['doc'].spec.attrs?.['compat'])
    attrs['compat'] = ctx.compat;
  return ctx.schema.nodes['doc'].create(
    Object.keys(attrs).length > 0 ? attrs : null,
    blocks.length > 0 ? blocks : [ctx.schema.nodes['paragraph'].create()],
  );
}

async function extractMedia(zip: JSZip): Promise<{
  media: Map<string, string>;
  vectorMedia: Map<string, WmfVectorResult>;
}> {
  const media = new Map<string, string>();
  const vectorMedia = new Map<string, WmfVectorResult>();
  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith('word/media/')) continue;
    const entry = zip.file(path);
    if (!entry || entry.dir) continue;
    // A metafile no browser can decode — but the kind Word writes most, a
    // single bitmap in EMF clothing, is re-framed as the .bmp it really is
    // (see emf.ts). Vector WMFs of the MathType-preview shape resolve to a
    // display list the painter replays (wmf-vector.ts) — the media entry
    // still keeps the original bytes, for the export. Anything else keeps
    // its bytes and paints as the placeholder.
    const lower = path.toLowerCase();
    if (lower.endsWith('.emf') || lower.endsWith('.wmf')) {
      const bytes = await entry.async('uint8array');
      const bmp = lower.endsWith('.emf')
        ? emfBitmapDataUrl(bytes)
        : wmfBitmapDataUrl(bytes);
      if (bmp) {
        media.set(path, bmp);
        continue;
      }
      if (lower.endsWith('.wmf')) {
        const vector = wmfVectorSpec(bytes);
        if (vector) vectorMedia.set(path, vector);
      }
    }
    media.set(
      path,
      `data:${mimeOf(path)};base64,${await entry.async('base64')}`,
    );
  }
  return { media, vectorMedia };
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
  const settingsPart = await readPart(zip, 'word/settings.xml');

  const parsePart = (name: string, xml: string): OoxmlNode => {
    const root = parseXml(xml);
    audit.registerPart(name, root);
    return root;
  };

  // Settings are read before everything else: the theme resolver needs
  // w:clrSchemeMapping (bg1/tx1 → which theme slot) and the compat profile is
  // consulted while the body is parsed (auto spacing, table indents).
  const settingsEl = settingsPart
    ? child(parsePart('word/settings.xml', settingsPart), 'w:settings')
    : undefined;
  const compat = parseCompat(settingsEl);

  // Stateless/shared pieces; numbering counters are per-story (built fresh below).
  const themeRoot = themeXml
    ? parsePart('word/theme/theme1.xml', themeXml)
    : undefined;
  const resolveTheme = buildThemeResolver(
    themeRoot,
    child(settingsEl, 'w:clrSchemeMapping'),
  );
  const resolveFont = buildThemeFontResolver(themeRoot);
  const resolveThemeFill = buildThemeFillResolver(themeRoot, resolveTheme);
  const resolveThemeLine = buildThemeLineResolver(themeRoot);
  const styles = buildStyleRegistry(
    stylesXml ? parsePart('word/styles.xml', stylesXml) : undefined,
    resolveTheme,
    resolveFont,
  );
  const numberingRoot = numberingXml
    ? parsePart('word/numbering.xml', numberingXml)
    : undefined;
  const { media, vectorMedia } = await extractMedia(zip);
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
    compat,
    styles,
    numbering,
    rels,
    media,
    vectorMedia,
    resolveTheme,
    resolveFont,
    resolveThemeFill,
    resolveThemeLine,
    tableStyles: [],
    notes,
    sdtBoundaries: new WeakMap(),
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
    for (const node of unwrapContainers(body.children)) {
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
  const settings = settingsEl;
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
