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

/** One addressable block in reading order (table-cell paragraphs included). */
export interface DocBlock {
  /** Stable only within one docVersion — re-read after any change. */
  index: number;
  /** 'paragraph' | 'heading1'..'heading6' | future kinds. */
  type: string;
  text: string;
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

export interface MutationResult {
  docVersion: string;
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
  /** Only when capabilities.selection — the user's current selection. */
  getSelection?(): Promise<{ text: string; blockIndex: number } | null>;
  /** Persist to the host's backing store (file, DB…). */
  save(): Promise<void>;
  close(): Promise<void>;
}

/** Resolves the session a tool call operates on. Desktop: `documentId`
 *  omitted = the open document. Server: id is required (tenancy). */
export interface SessionProvider {
  get(documentId?: string): Promise<DocumentSession | null>;
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
