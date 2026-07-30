/**
 * The wire protocol between a DocumentSession host process and the process
 * where the document actually lives, when they are not the same — on desktop,
 * the Bun process hosts the MCP server while the document lives in the
 * WebView editor.
 *
 * Bun side: {@link RemoteSession} — a DocumentSession whose every method
 * serializes into one {@link SessionOpRequest} (pushed over the app's SSE
 * command channel) and awaits the matching {@link SessionOpResponse}.
 * WebView side: {@link executeOp} — runs the request against the local
 * session and produces the response, round-tripping contract errors by name
 * so optimistic-locking / anchoring semantics survive the process hop.
 */
import {
  AnchorError,
  NoDocumentError,
  VersionConflictError,
  type DocumentSession,
  type Formatting,
  type ImageChanges,
  type InsertAnchor,
  type MutationOptions,
  type SessionCapabilities,
} from './contract.js';

export type SessionOpName =
  | 'snapshot'
  | 'find'
  | 'replaceText'
  | 'insertContent'
  | 'applyFormatting'
  | 'updateImage'
  | 'getSelection'
  | 'save'
  // Host-level, handled by the caller before executeOp (see its note):
  /** Ask the user to allow an AI client in. */
  | 'consent'
  /** Host → UI: something host-side changed; `args[0]` is a host-defined
   *  topic the UI re-reads (it carries no document state itself). */
  | 'notify';

export interface SessionOpRequest {
  id: string;
  op: SessionOpName;
  args: unknown[];
  /** Which document the op is for, in the host's own id space; absent means
   *  "the one the user is working in".
   *
   *  A host with a single live document never needs it. One that can hold
   *  several open at once — side-by-side panes, several windows — does: the
   *  target would otherwise be implied by whatever happens to be focused when
   *  the request lands, which is a race, not an address. The wire stays
   *  ignorant of what a pane or a window is; it just carries the id. */
  doc?: string;
}

export type SessionOpResponse =
  | { id: string; ok: true; value: unknown }
  | { id: string; ok: false; error: { name: string; message: string } };

/** Run one wire op against a local session ('consent' and 'notify' are
 *  host-level and are handled by the caller before this). Never throws —
 *  errors are encoded. */
export async function executeOp(
  session: DocumentSession | null,
  request: SessionOpRequest,
): Promise<SessionOpResponse> {
  try {
    if (!session) throw new NoDocumentError();
    const value = await run(session, request);
    return { id: request.id, ok: true, value };
  } catch (err) {
    const e = err instanceof Error ? err : new Error(String(err));
    return {
      id: request.id,
      ok: false,
      error: { name: e.name, message: e.message },
    };
  }
}

function run(
  session: DocumentSession,
  { op, args }: SessionOpRequest,
): Promise<unknown> {
  switch (op) {
    case 'snapshot':
      return session.snapshot();
    case 'find':
      return session.find(args[0] as string);
    case 'replaceText':
      return session.replaceText(
        args[0] as string,
        args[1] as string,
        args[2] as MutationOptions | undefined,
      );
    case 'insertContent':
      return session.insertContent(
        args[0] as string,
        args[1] as InsertAnchor,
        args[2] as MutationOptions | undefined,
      );
    case 'applyFormatting':
      return session.applyFormatting(
        args[0] as string,
        args[1] as Formatting,
        args[2] as MutationOptions | undefined,
      );
    case 'updateImage':
      return session.updateImage(
        args[0] as number,
        args[1] as number,
        args[2] as ImageChanges,
        args[3] as MutationOptions | undefined,
      );
    case 'getSelection':
      return session.getSelection?.() ?? Promise.resolve(null);
    case 'save':
      return session.save();
    default:
      return Promise.reject(new Error(`unknown session op: ${op as string}`));
  }
}

/** Rebuild a contract error from its wire form so upstream catch-by-class
 *  (createMcpServer's teaching errors) keeps working across the hop. */
export function reviveError(error: { name: string; message: string }): Error {
  switch (error.name) {
    case 'AnchorError':
      return new AnchorError(error.message);
    case 'NoDocumentError':
      return new NoDocumentError(error.message);
    case 'VersionConflictError': {
      const revived = new VersionConflictError('?', '?');
      revived.message = error.message;
      return revived;
    }
    default: {
      const generic = new Error(error.message);
      generic.name = error.name;
      return generic;
    }
  }
}

/** DocumentSession proxy: forwards every call over `send` (the transport is
 *  the host's business — desktop uses its SSE command channel + POST-back). */
export class RemoteSession implements DocumentSession {
  constructor(
    private readonly send: (
      op: SessionOpName,
      args: unknown[],
      doc?: string,
    ) => Promise<unknown>,
    readonly capabilities: SessionCapabilities,
    /** Pins this session to one document (see SessionOpRequest.doc). Omit for
     *  "whatever the user is working in". */
    private readonly doc?: string,
  ) {}

  private async call<T>(op: SessionOpName, args: unknown[] = []): Promise<T> {
    return (await this.send(op, args, this.doc)) as T;
  }

  snapshot() {
    return this.call<Awaited<ReturnType<DocumentSession['snapshot']>>>(
      'snapshot',
    );
  }
  find(query: string) {
    return this.call<Awaited<ReturnType<DocumentSession['find']>>>('find', [
      query,
    ]);
  }
  replaceText(oldText: string, newText: string, opts?: MutationOptions) {
    return this.call<Awaited<ReturnType<DocumentSession['replaceText']>>>(
      'replaceText',
      [oldText, newText, opts],
    );
  }
  insertContent(content: string, anchor: InsertAnchor, opts?: MutationOptions) {
    return this.call<Awaited<ReturnType<DocumentSession['insertContent']>>>(
      'insertContent',
      [content, anchor, opts],
    );
  }
  applyFormatting(target: string, format: Formatting, opts?: MutationOptions) {
    return this.call<Awaited<ReturnType<DocumentSession['applyFormatting']>>>(
      'applyFormatting',
      [target, format, opts],
    );
  }
  updateImage(
    blockIndex: number,
    imageIndex: number,
    changes: ImageChanges,
    opts?: MutationOptions,
  ) {
    return this.call<Awaited<ReturnType<DocumentSession['updateImage']>>>(
      'updateImage',
      [blockIndex, imageIndex, changes, opts],
    );
  }
  async getSelection() {
    return this.call<{ text: string; blockIndex: number } | null>(
      'getSelection',
    );
  }
  async save() {
    await this.call<void>('save');
  }
  async close() {
    /* the document belongs to the remote host — nothing to release here */
  }
}
