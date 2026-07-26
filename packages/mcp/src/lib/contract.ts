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
  /** Drawn vector shape (rect/ellipse/…) vs a bitmap picture. */
  kind: 'bitmap' | 'shape';
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
}

export interface DocSnapshot {
  docVersion: string;
  blocks: DocBlock[];
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
  /** Paragraph alignment of the block(s) containing the target text. */
  align?: 'left' | 'center' | 'right' | 'justify';
}

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
}

/** The port every document host implements. */
export interface DocumentSession {
  readonly capabilities: SessionCapabilities;
  snapshot(): Promise<DocSnapshot>;
  find(query: string): Promise<FindMatch[]>;
  replaceText(oldText: string, newText: string, opts?: MutationOptions): Promise<MutationResult>;
  /** `content`: plain text; each line becomes one paragraph. */
  insertContent(content: string, anchor: InsertAnchor, opts?: MutationOptions): Promise<MutationResult>;
  applyFormatting(target: string, format: Formatting, opts?: MutationOptions): Promise<MutationResult>;
  /** Resize/rotate one image, addressed as (blockIndex, imageIndex) from the
   *  latest snapshot. One transaction — a single undo step in a live editor. */
  updateImage(blockIndex: number, imageIndex: number, changes: ImageChanges, opts?: MutationOptions): Promise<MutationResult>;
  /** Only when capabilities.selection — the user's current selection. */
  getSelection?(): Promise<{ text: string; blockIndex: number } | null>;
  /** Persist to the host's backing store (file, DB…). */
  save(): Promise<void>;
  close(): Promise<void>;
}

/** One document a host can serve, as listed by {@link SessionProvider.list}. */
export interface DocumentRef {
  /** The value to pass back as `documentId`. */
  id: string;
  /** Human-readable label (file name, title…). */
  name: string;
  /** Whether this document is the one currently open/focused in the host. */
  open?: boolean;
}

/** Resolves the session a tool call operates on. Desktop: `documentId`
 *  omitted = the open document. Server: id is required (tenancy). */
export interface SessionProvider {
  get(documentId?: string): Promise<DocumentSession | null>;
  /** The documents this host can serve, so a client can discover the ids it
   *  may pass as `documentId`. Optional: a single-document host omits it and
   *  the list_documents tool is simply not offered. */
  list?(): Promise<DocumentRef[]>;
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

/** The host has no document to operate on (desktop: nothing open yet). */
export class NoDocumentError extends Error {
  constructor(message = 'No document is open. Ask the user to open a document first.') {
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
