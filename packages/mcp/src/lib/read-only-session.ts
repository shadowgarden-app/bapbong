/**
 * A read-only view over any {@link DocumentSession}.
 *
 * Reads pass straight through; every mutation (and `save`) throws
 * {@link ReadOnlyError} with a message the model can act on, so an agent that
 * tries to edit a document the host only lets it read gets a clear refusal
 * instead of a silent no-op. Hosts use this to widen *reading* to documents
 * they are not willing to have written — e.g. the desktop app exposing every
 * document in the open workspace folder while only the document the user
 * actually has open is writable.
 */
import type {
  DocSnapshot,
  DocumentSession,
  FindMatch,
  MutationResult,
  SessionCapabilities,
} from './contract.js';
import { ReadOnlyError } from './contract.js';

export interface ReadOnlySessionOptions {
  /** Appended to the refusal so the agent learns *why* and what to suggest. */
  reason?: string;
}

const DEFAULT_REASON =
  'it is open for reading only. Ask the user to open this document in bapbong if they want it edited.';

export class ReadOnlySession implements DocumentSession {
  constructor(
    private readonly inner: DocumentSession,
    private readonly opts: ReadOnlySessionOptions = {},
  ) {}

  get capabilities(): SessionCapabilities {
    return this.inner.capabilities;
  }

  private refuse(what: string): never {
    throw new ReadOnlyError(
      `Cannot ${what} — ${this.opts.reason ?? DEFAULT_REASON}`,
    );
  }

  // ── Reads: straight through ─────────────────────────────────────────
  snapshot(): Promise<DocSnapshot> {
    return this.inner.snapshot();
  }
  find(query: string): Promise<FindMatch[]> {
    return this.inner.find(query);
  }
  getSelection(): Promise<{ text: string; blockIndex: number } | null> {
    return this.inner.getSelection?.() ?? Promise.resolve(null);
  }
  close(): Promise<void> {
    return this.inner.close();
  }

  // ── Mutations: refused ──────────────────────────────────────────────
  replaceText(): Promise<MutationResult> {
    this.refuse('replace text in this document');
  }
  insertContent(): Promise<MutationResult> {
    this.refuse('insert content into this document');
  }
  applyFormatting(): Promise<MutationResult> {
    this.refuse('format this document');
  }
  updateImage(): Promise<MutationResult> {
    this.refuse('change images in this document');
  }
  save(): Promise<void> {
    this.refuse('save this document');
  }
}
