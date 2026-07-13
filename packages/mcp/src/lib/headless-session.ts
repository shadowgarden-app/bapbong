/**
 * DocumentSession over a headless ProseMirror EditorState — no DOM, no UI.
 *
 * A thin host around {@link PmDocSession} (which owns ALL the anchoring /
 * locking / mutation semantics): this file just holds the state, counts
 * versions, and imports/exports the .docx bytes. It is both the unit-test
 * host for the tool semantics and the seed of a future server-side document
 * service; the desktop app hosts the same PmDocSession over its live editor.
 */
import { importDocx, exportDocx, type DocxImport } from '@shadow-garden/bapbong-headless';
import { EditorState, type Transaction } from 'prosemirror-state';
import type {
  DocSnapshot,
  DocumentSession,
  FindMatch,
  Formatting,
  ImageChanges,
  InsertAnchor,
  MutationOptions,
  MutationResult,
  SessionCapabilities,
} from './contract.js';
import { PmDocSession, type PmSessionHost } from './pm-session.js';

export interface HeadlessSessionOptions {
  /** Where save() writes the exported bytes (file, DB, test sink…). */
  onSave?: (bytes: Uint8Array) => void | Promise<void>;
  name?: string;
}

export class HeadlessSession implements DocumentSession {
  readonly capabilities: SessionCapabilities = { selection: false };

  private state: EditorState;
  private version = 1;
  private dirty = false;
  private readonly inner: PmDocSession;

  private constructor(
    state: EditorState,
    private readonly raw: DocxImport['raw'],
    private readonly opts: HeadlessSessionOptions,
  ) {
    this.state = state;
    const host: PmSessionHost = {
      getState: () => this.state,
      apply: (tr: Transaction) => {
        this.state = this.state.apply(tr);
        this.version++;
        this.dirty = true;
      },
      getVersion: () => this.docVersion,
      meta: () => ({ name: this.opts.name, dirty: this.dirty }),
      save: async () => {
        const bytes = await exportDocx(this.state.doc, this.raw ? { carry: this.raw } : undefined);
        await this.opts.onSave?.(bytes);
        this.dirty = false;
      },
      // no selection() — headless documents have no user selection
    };
    this.inner = new PmDocSession(host);
  }

  static async open(bytes: ArrayBuffer | Uint8Array, opts: HeadlessSessionOptions = {}): Promise<HeadlessSession> {
    const { doc, raw } = await importDocx(bytes instanceof Uint8Array ? (bytes.slice().buffer as ArrayBuffer) : bytes);
    return new HeadlessSession(EditorState.create({ doc }), raw, opts);
  }

  get docVersion(): string {
    return `v${this.version}`;
  }

  snapshot(): Promise<DocSnapshot> {
    return this.inner.snapshot();
  }
  find(query: string): Promise<FindMatch[]> {
    return this.inner.find(query);
  }
  replaceText(oldText: string, newText: string, opts?: MutationOptions): Promise<MutationResult> {
    return this.inner.replaceText(oldText, newText, opts);
  }
  insertContent(content: string, anchor: InsertAnchor, opts?: MutationOptions): Promise<MutationResult> {
    return this.inner.insertContent(content, anchor, opts);
  }
  applyFormatting(target: string, format: Formatting, opts?: MutationOptions): Promise<MutationResult> {
    return this.inner.applyFormatting(target, format, opts);
  }
  updateImage(blockIndex: number, imageIndex: number, changes: ImageChanges, opts?: MutationOptions): Promise<MutationResult> {
    return this.inner.updateImage(blockIndex, imageIndex, changes, opts);
  }
  save(): Promise<void> {
    return this.inner.save();
  }
  close(): Promise<void> {
    return this.inner.close();
  }
}
