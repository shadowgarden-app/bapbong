/**
 * DocumentSession over a headless ProseMirror EditorState — no DOM, no UI.
 *
 * A thin host around {@link PmDocSession} (which owns ALL the anchoring /
 * locking / mutation semantics): this file just holds the state, counts
 * versions, and imports/exports the .docx bytes. It is both the unit-test
 * host for the tool semantics and the seed of a future server-side document
 * service; the desktop app hosts the same PmDocSession over its live editor.
 */
import {
  catalogTableStyles,
  importDocx,
  exportDocx,
  pageSetupTransaction,
  type DocxImport,
} from '@shadow-garden/bapbong-headless';
import { EditorState, type Transaction } from 'prosemirror-state';
import type {
  Content,
  DocSnapshot,
  DocumentSession,
  FindMatch,
  FormatTarget,
  Formatting,
  TableEdit,
  ImageBox,
  ImageBytes,
  ImageChanges,
  ImagePlacement,
  ImageSource,
  ImageSourceKind,
  InsertAnchor,
  MutationOptions,
  MutationResult,
  PageSetup,
  SessionCapabilities,
} from './contract.js';
import { PmDocSession, type PmSessionHost } from './pm-session.js';

export interface HeadlessSessionOptions {
  /** Where save() writes the exported bytes (file, DB, test sink…). */
  onSave?: (bytes: Uint8Array) => void | Promise<void>;
  name?: string;
  /** Document identity woven into `docVersion` (whatever the host keys
   *  documents by — a path, a row id). A host serving several documents MUST
   *  set it: without it every document counts `v1, v2, …` independently, so an
   *  `expectedVersion` read from one document can silently satisfy the
   *  optimistic lock of another. */
  id?: string;
  /** Fetch a picture's bytes for insert_image / replace_image, and which
   *  sources that covers. A host without this takes no new pictures. */
  readImage?: (source: ImageSource) => Promise<ImageBytes>;
  imageSources?: readonly ImageSourceKind[];
}

export class HeadlessSession implements DocumentSession {
  get capabilities(): SessionCapabilities {
    return this.inner.capabilities;
  }

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
        const bytes = await exportDocx(
          this.state.doc,
          this.raw ? { carry: this.raw } : undefined,
        );
        await this.opts.onSave?.(bytes);
        this.dirty = false;
      },
      // no selection() — headless documents have no user selection
      ...(this.opts.readImage
        ? {
            readImage: this.opts.readImage,
            imageSources: this.opts.imageSources ?? [],
          }
        : {}),
      // The editor's own Layout commands, composed.
      pageSetup: (state, change) => pageSetupTransaction(state, change),
      tableStyle: () => {
        const grid = catalogTableStyles().find((t) => t.id === 'TableGrid');
        return grid ? { styleId: grid.id, style: grid.style } : undefined;
      },
    };
    this.inner = new PmDocSession(host);
  }

  static async open(
    bytes: ArrayBuffer | Uint8Array,
    opts: HeadlessSessionOptions = {},
  ): Promise<HeadlessSession> {
    const { doc, raw } = await importDocx(
      bytes instanceof Uint8Array
        ? (bytes.slice().buffer as ArrayBuffer)
        : bytes,
    );
    return new HeadlessSession(EditorState.create({ doc }), raw, opts);
  }

  /** Optimistic-lock token. Prefixed with the document id when the host gave
   *  one, so a version read from one document never satisfies another's lock. */
  get docVersion(): string {
    return this.opts.id
      ? `${this.opts.id}:v${this.version}`
      : `v${this.version}`;
  }

  snapshot(): Promise<DocSnapshot> {
    return this.inner.snapshot();
  }
  find(query: string): Promise<FindMatch[]> {
    return this.inner.find(query);
  }
  replaceText(
    oldText: string,
    newText: string,
    opts?: MutationOptions,
  ): Promise<MutationResult> {
    return this.inner.replaceText(oldText, newText, opts);
  }
  insertContent(
    content: Content,
    anchor: InsertAnchor,
    opts?: MutationOptions,
  ): Promise<MutationResult> {
    return this.inner.insertContent(content, anchor, opts);
  }
  editTable(
    tableIndex: number,
    edit: TableEdit,
    opts?: MutationOptions,
  ): Promise<MutationResult & { rows: number; cols: number }> {
    return this.inner.editTable(tableIndex, edit, opts);
  }
  applyFormatting(
    target: FormatTarget,
    format: Formatting,
    opts?: MutationOptions,
  ): Promise<MutationResult> {
    return this.inner.applyFormatting(target, format, opts);
  }
  updateImage(
    blockIndex: number,
    imageIndex: number,
    changes: ImageChanges,
    opts?: MutationOptions,
  ): Promise<MutationResult> {
    return this.inner.updateImage(blockIndex, imageIndex, changes, opts);
  }
  insertImage(
    source: ImageSource,
    anchor: InsertAnchor,
    placement?: ImagePlacement,
    opts?: MutationOptions,
  ): Promise<MutationResult & ImageBox> {
    return this.inner.insertImage(source, anchor, placement, opts);
  }
  replaceImage(
    blockIndex: number,
    imageIndex: number,
    source: ImageSource,
    placement?: ImagePlacement,
    opts?: MutationOptions,
  ): Promise<MutationResult & ImageBox> {
    return this.inner.replaceImage(
      blockIndex,
      imageIndex,
      source,
      placement,
      opts,
    );
  }
  deleteImage(
    blockIndex: number,
    imageIndex: number,
    opts?: MutationOptions,
  ): Promise<MutationResult> {
    return this.inner.deleteImage(blockIndex, imageIndex, opts);
  }
  deleteBlocks(
    blockIndex: number,
    count: number,
    opts?: MutationOptions,
  ): Promise<MutationResult & { deleted: number }> {
    return this.inner.deleteBlocks(blockIndex, count, opts);
  }
  pageSetup(
    setup: PageSetup,
    opts?: MutationOptions,
  ): Promise<MutationResult & { sections: number }> {
    return this.inner.pageSetup(setup, opts);
  }
  save(): Promise<void> {
    return this.inner.save();
  }
  close(): Promise<void> {
    return this.inner.close();
  }
}
