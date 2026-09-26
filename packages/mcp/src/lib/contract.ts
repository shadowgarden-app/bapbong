/**
 * The host-agnostic contract between MCP tools and a document host.
 *
 * Tools (packages/mcp/src/lib/server.ts) are written once against
 * {@link DocumentSession}; hosts implement it for their document store:
 * the desktop app binds the live editor (WebView), a server binds a headless
 * ProseMirror EditorState ({@link ../headless-session}). Anything here must
 * therefore stay meaningful for BOTH — no DOM, no desktop-isms.
 *
 * Anchoring semantics (the part language-agnostic hosts must reproduce):
 * - Text anchors match within a single block (paragraph/cell paragraph);
 *   matches never span blocks or inline atoms (images/fields).
 * - A mutation anchor must match exactly once, or the caller must pass
 *   `occurrence` (1-based, in document order). Ambiguity is an error that
 *   lists the occurrence count — the model retries with `occurrence`.
 * - `docVersion` is an opaque string that changes whenever the document
 *   changes. Mutations MAY pass `expectedVersion`; a mismatch raises
 *   {@link VersionConflictError} and the caller re-reads before retrying.
 */

import type { Content, TabStop, TableEdit } from './blocks.js';
export type { Block, Cell, Content, TableEdit, TabStop } from './blocks.js';

/**
 * What kind of picture an image box holds. The agent needs this before it
 * offers to change one: only a `bitmap` is a picture in the everyday sense.
 *
 * - `shape`   a drawn shape (rect, line, …) the editor models;
 * - `drawing` art we paint but do not model — a group of freeform shapes
 *             from the file, kept verbatim so a save never loses it.
 *             Replacing one turns editable art into a flat picture;
 * - `equation` a MathType/OLE object whose picture is only its preview.
 */
export type DocImageKind = 'bitmap' | 'shape' | 'drawing' | 'equation';

/** An inline image (bitmap picture or drawn shape) inside a block —
 *  addressable as (block index, image index) by updateImage. */
export interface DocImage {
  /** 0-based among the block's images, in block order. */
  index: number;
  alt: string;
  /** CSS px. */
  width: number;
  height: number;
  /** Clockwise degrees around the image center (0 when unrotated). */
  rotation: number;
  kind: DocImageKind;
  /** Anchored to the page and wrapped by the text, rather than sitting
   *  inline in it. Set only when it floats. */
  float?: boolean;
}

/** Where the bytes of a new picture come from. The session never fetches
 *  anything itself: the host reads the file, downloads the attachment or
 *  rasterizes the markup, and declares which of these it can do
 *  ({@link SessionCapabilities.images}). */
export type ImageSourceKind = 'path' | 'attachment' | 'svg';

export type ImageSource =
  | { kind: 'path'; path: string }
  | { kind: 'attachment'; attachmentId: string }
  | { kind: 'svg'; svg: string };

/** A host's answer to an {@link ImageSource}: the picture's bytes, and the
 *  media type it was announced as (advisory — the bytes decide). */
export interface ImageBytes {
  bytes: Uint8Array;
  mediaType?: string;
  /**
   * Device pixels per CSS pixel in those bytes; default 1.
   *
   * A host that rasterizes (SVG → PNG) draws at 2× so the picture stays
   * sharp in print, and says so here — otherwise a diagram the agent drew
   * 400 wide would come into the document at 800 and be shrunk to whatever
   * the page allows. The extra pixels ride along; only the box is divided.
   */
  scale?: number;
}

/** How a new picture is sized and described. */
export interface ImagePlacement {
  /** Display width in CSS px. Absent: the picture's own size, shrunk to the
   *  text width when it is wider; on a replacement, the box it takes over. */
  width?: number;
  /** Alt text (Word's "Description"). A replacement keeps the old one when
   *  this is absent. */
  alt?: string;
}

/** The box a picture ended up occupying, in CSS px. */
export interface ImageBox {
  width: number;
  height: number;
}

/** One addressable block in reading order (table-cell paragraphs included). */
export interface DocBlock {
  /** Stable only within one docVersion — re-read after any change. */
  index: number;
  /** 'paragraph' | 'heading1'..'heading6' | future kinds. */
  type: string;
  text: string;
  /** The block's inline images, when it has any. */
  images?: DocImage[];
  /** Set when the block sits in a table cell: the table's 0-based index in
   *  the document (what edit_table addresses), the row and the cell. */
  table?: { index: number; row: number; cell: number };
  /** Set when the block is a list item: which kind, and its nesting level
   *  (1 = top). The number or bullet itself is drawn, not part of `text`. */
  list?: { kind: 'bullet' | 'number'; level: number };
  /** The section (1-based) the block is in — set only when the document has
   *  more than one, which is when page_setup's `section` means something. */
  section?: number;
  /** The hyperlinks in the block: the linked text and where it goes. */
  links?: { text: string; href: string }[];
}

/** A header or footer as the pages show it. Read-only here: no command
 *  edits one yet. */
export interface DocChrome {
  part: 'header' | 'footer';
  /** "default" (every page), "first" (a section's first page) or "even". */
  variant: string;
  /** The 1-based sections that show it — a story linked to the previous
   *  section's is listed once, with both. */
  sections: number[];
  /** Its text, one line per paragraph; a page number shows as the field's
   *  placeholder. */
  text: string;
}

export interface DocSnapshot {
  docVersion: string;
  blocks: DocBlock[];
  /** Headers and footers with text in them, when the host can read them. */
  chrome?: DocChrome[];
  /** Host metadata (file name, dirty state…) — informational only. */
  meta: { name?: string; dirty?: boolean };
}

export interface FindMatch {
  blockIndex: number;
  /** 1-based occurrence of the query across the whole document. */
  occurrence: number;
  /** The match with surrounding text — enough to disambiguate. */
  context: string;
}

/** Where insertContent puts new blocks. */
export type InsertAnchor =
  | { position: 'before' | 'after'; text: string; occurrence?: number }
  | { position: 'document_end' };

/** Character marks are tri-state: true = apply, false = remove, absent = keep. */
export interface Formatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font size in points for the target text. */
  fontSize?: number;
  /** Font family for the target text; null removes it (the style's font). */
  fontFamily?: string | null;
  /** Text colour "#RRGGBB"; null removes it (automatic). */
  color?: string | null;
  /** Highlight colour "#RRGGBB"; null removes it. */
  highlight?: string | null;
  /** Raised or lowered text; null puts it back on the baseline. */
  verticalAlign?: 'superscript' | 'subscript' | null;
  /** Strip the target's character formatting first (links, comments and
   *  footnotes stay), then apply whatever else this Formatting says. */
  clear?: boolean;
  /** Make the target text a hyperlink to this address (http(s), mailto,
   *  tel, "#bookmark"); null unlinks it. */
  link?: string | null;
  /** Space above / below the containing paragraph, in points. */
  spaceBefore?: number;
  spaceAfter?: number;
  /** Line spacing of the containing paragraph: a multiple of single
   *  (1, 1.15, 1.5, 2), or an exact / minimum line height in points. */
  lineSpacing?: LineSpacing;
  /** Indents of the containing paragraph (0 removes one). `firstLine` and
   *  `hanging` exclude each other: setting one clears the other. */
  indent?: IndentChange;
  /** Paragraph alignment of the block(s) containing the target text. */
  align?: 'left' | 'center' | 'right' | 'justify';
  /** Word "Heading N" for the containing paragraph; 0 or null = body text. */
  heading?: number | null;
  /** Named paragraph style; null = body text. Clears `heading`. */
  style?: 'Title' | 'Subtitle' | null;
  /** The containing paragraph's tab stops (replaces them all). */
  tabs?: TabStop[];
  /** Make the containing paragraph a list item of this kind (it joins a list
   *  of the same kind right above it, else starts one), or null to make it
   *  body text again. */
  list?: 'bullet' | 'number' | null;
  /** Nesting level of the containing list item, 1 (top) to 3. */
  listLevel?: number;
}

/** A length as an agent writes it: a number is centimetres, or a string
 *  such as "2cm", "1in", "25mm". */
export type Length = number | string;

/** Word's paper sizes, by the names Word shows. */
export type PaperName = 'A3' | 'A4' | 'A5' | 'Letter' | 'Legal' | 'Executive';

/** A page-setup request as an agent makes it (see page_setup). Section
 *  numbers are 1-based, like the ones get_document reports. */
export interface PageSetup {
  /** The section to change; absent = every section. */
  section?: number;
  orientation?: 'portrait' | 'landscape';
  paper?: PaperName;
  /** A Word preset, or lengths per side (a side left out keeps its value). */
  margins?:
    | 'normal'
    | 'narrow'
    | 'moderate'
    | 'wide'
    | { top?: Length; right?: Length; bottom?: Length; left?: Length };
  /** Text columns, 1-3. */
  columns?: number;
  /** Start a new section after this block (an index from get_document; a
   *  paragraph outside tables). `newPage` false makes it continuous. */
  sectionBreakAfter?: { blockIndex: number; newPage?: boolean };
  /** Remove section break n (1-based: the one ending section n). */
  removeSectionBreak?: number;
}

/** What a host's page-setup port takes: {@link PageSetup} resolved to the
 *  model — 0-based indexes, px, the model's paper keys. */
export interface PageSetupChange {
  section?: number;
  orientation?: 'portrait' | 'landscape';
  paper?: 'a3' | 'a4' | 'a5' | 'letter' | 'legal' | 'executive';
  margins?:
    | 'normal'
    | 'narrow'
    | 'moderate'
    | 'wide'
    | { top?: number; right?: number; bottom?: number; left?: number };
  columns?: number;
  sectionBreak?: { after: number; newPage: boolean };
  removeSectionBreak?: number;
}

/** Line spacing: a multiple of single, or a height in points. */
export type LineSpacing = number | { exact: number } | { atLeast: number };

/** Paragraph indents as lengths (cm as a number, or "1cm", "0.5in"). */
export interface IndentChange {
  left?: Length;
  right?: Length;
  firstLine?: Length;
  hanging?: Length;
}

/** What applyFormatting addresses: exact text (matched once, or with
 *  occurrence) or a whole block by index from the latest snapshot. */
export type FormatTarget = string | { blockIndex: number };

/** Partial image update — absent fields keep their current value. */
export interface ImageChanges {
  /** CSS px (rounded; must be positive). */
  width?: number;
  /** CSS px (rounded; must be positive). */
  height?: number;
  /** Clockwise degrees around the center — normalized to [0, 360). */
  rotation?: number;
}

export interface MutationResult {
  docVersion: string;
  /** The affected range (PM positions) — informational; hosts use it to
   *  surface the edit (the desktop selects + scrolls to it). Positions are
   *  only stable within the returned docVersion. */
  range?: { from: number; to: number };
}

export interface MutationOptions {
  expectedVersion?: string;
  /** 1-based pick when the anchor text matches more than once. */
  occurrence?: number;
}

/** What a host supports; tools that need a missing capability aren't offered. */
export interface SessionCapabilities {
  /** A live user selection exists (desktop editor) — enables get_selection. */
  selection: boolean;
  /** Which image sources the host can fetch. Empty (or absent) means it
   *  cannot take new pictures at all, and the image tools aren't offered. */
  images?: readonly ImageSourceKind[];
  /** The session can change page geometry and sections — enables
   *  page_setup. */
  pageSetup?: boolean;
}

/** The port every document host implements. */
export interface DocumentSession {
  readonly capabilities: SessionCapabilities;
  snapshot(): Promise<DocSnapshot>;
  find(query: string): Promise<FindMatch[]>;
  replaceText(
    oldText: string,
    newText: string,
    opts?: MutationOptions,
  ): Promise<MutationResult>;
  /** `content`: plain text; each line becomes one paragraph. */
  insertContent(
    content: Content,
    anchor: InsertAnchor,
    opts?: MutationOptions,
  ): Promise<MutationResult>;
  applyFormatting(
    target: FormatTarget,
    format: Formatting,
    opts?: MutationOptions,
  ): Promise<MutationResult>;
  /** Change an existing table — rows, merges, widths, borders — addressed by
   *  its 0-based index from the latest snapshot. One transaction. */
  editTable(
    tableIndex: number,
    edit: TableEdit,
    opts?: MutationOptions,
  ): Promise<MutationResult & { rows: number; cols: number }>;
  /** Resize/rotate one image, addressed as (blockIndex, imageIndex) from the
   *  latest snapshot. One transaction — a single undo step in a live editor. */
  updateImage(
    blockIndex: number,
    imageIndex: number,
    changes: ImageChanges,
    opts?: MutationOptions,
  ): Promise<MutationResult>;
  /** Put a new picture in its own paragraph at `anchor`. Only when
   *  `capabilities.images` lists the source's kind. */
  insertImage(
    source: ImageSource,
    anchor: InsertAnchor,
    placement?: ImagePlacement,
    opts?: MutationOptions,
  ): Promise<MutationResult & ImageBox>;
  /** Swap the picture at (blockIndex, imageIndex) for a new one, keeping
   *  where it sits: its anchor, its wrap, and its width unless `placement`
   *  says otherwise (the height follows the new picture's proportions). */
  replaceImage(
    blockIndex: number,
    imageIndex: number,
    source: ImageSource,
    placement?: ImagePlacement,
    opts?: MutationOptions,
  ): Promise<MutationResult & ImageBox>;
  /** Remove the picture at (blockIndex, imageIndex). Its paragraph stays. */
  deleteImage(
    blockIndex: number,
    imageIndex: number,
    opts?: MutationOptions,
  ): Promise<MutationResult>;
  /** Remove `count` consecutive blocks from `blockIndex` (snapshot order),
   *  whole — paragraph, text, pictures and all. A paragraph in a table cell
   *  goes only while its cell keeps another; the document keeps one empty
   *  paragraph if nothing else is left. One transaction. */
  deleteBlocks(
    blockIndex: number,
    count: number,
    opts?: MutationOptions,
  ): Promise<MutationResult & { deleted: number }>;
  /** Page geometry, columns and section breaks — for one section or all
   *  of them. Only when capabilities.pageSetup. One transaction. */
  pageSetup(
    setup: PageSetup,
    opts?: MutationOptions,
  ): Promise<MutationResult & { sections: number }>;
  /** Only when capabilities.selection — the user's current selection. */
  getSelection?(): Promise<{ text: string; blockIndex: number } | null>;
  /** Persist to the host's backing store (file, DB…). */
  save(): Promise<void>;
  close(): Promise<void>;
}

/**
 * Resolves the session a tool call operates on.
 *
 * `documentId` is how a host that knows several documents picks one — a server
 * keyed by tenancy, say. There is deliberately no way to *enumerate* them here:
 * the tools this package registers cover editing ONE document, the one the host
 * hands back when the id is omitted. A host that wants agents working across a
 * set of documents supplies its own way to discover them, by registering a tool
 * of its own on the returned server.
 *
 * For the common case — one document, no ids — use
 * {@link singleDocumentProvider}.
 */
export interface SessionProvider {
  get(documentId?: string): Promise<DocumentSession | null>;
}

/** The provider for a host with exactly one document: it answers when the id is
 *  omitted, and reports "no such document" for anything else. */
export function singleDocumentProvider(
  session: DocumentSession | (() => DocumentSession | null),
): SessionProvider {
  const resolve = typeof session === 'function' ? session : () => session;
  return {
    get: async (documentId?: string) =>
      documentId === undefined ? resolve() : null,
  };
}

/** The document changed since `expectedVersion` — re-read, then retry. */
export class VersionConflictError extends Error {
  constructor(current: string, expected: string) {
    super(
      `Document changed (version is now ${current}, you expected ${expected}). ` +
        `Call get_document again, re-locate your anchor, then retry.`,
    );
    this.name = 'VersionConflictError';
  }
}

/** Anchor text not found, or ambiguous without `occurrence`. */
export class AnchorError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AnchorError';
  }
}

/** Structured content that cannot be built (ragged table rows, a widths
 *  list that does not fit…) — the message says what to change. */
export class ContentError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ContentError';
  }
}

/** The host has no document to operate on (desktop: nothing open yet). */
export class NoDocumentError extends Error {
  constructor(
    message = 'No document is open. Ask the user to open a document first.',
  ) {
    super(message);
    this.name = 'NoDocumentError';
  }
}

/** The session is readable but not writable — the host granted read-only
 *  access to this document (see {@link ReadOnlySession}). */
export class ReadOnlyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ReadOnlyError';
  }
}
