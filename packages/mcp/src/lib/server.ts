/**
 * The MCP server over a {@link SessionProvider} — the host-agnostic half of
 * M15. Hosts supply sessions (desktop: the live editor; server: headless
 * EditorStates) and a transport (loopback Bun.serve, TLS gateway, stdio…);
 * everything the agent sees — tool names, schemas, anchoring/versioning
 * semantics, error texts that teach the retry — is defined once, here.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { AnchorError, NoDocumentError, VersionConflictError, type DocumentSession, type SessionProvider } from './contract.js';

export interface CreateMcpServerOptions {
  name?: string;
  version?: string;
  /** Offer get_selection (hosts with a live user selection — the desktop). */
  selection?: boolean;
}

const documentId = z
  .string()
  .optional()
  .describe('Target document id. Omit for the currently open document (desktop).');

const expectedVersion = z
  .string()
  .optional()
  .describe('docVersion you last read. If the document changed since, the call fails and you must re-read.');

const occurrence = z
  .number()
  .int()
  .min(1)
  .optional()
  .describe('1-based pick when the anchor text matches more than once (document order).');

/** JSON text content for a tool result. */
function json(value: unknown): { content: { type: 'text'; text: string }[] } {
  return { content: [{ type: 'text', text: JSON.stringify(value, null, 1) }] };
}

function errorText(message: string): { content: { type: 'text'; text: string }[]; isError: true } {
  return { content: [{ type: 'text', text: message }], isError: true };
}

export function createMcpServer(provider: SessionProvider, opts: CreateMcpServerOptions = {}): McpServer {
  const server = new McpServer({
    name: opts.name ?? 'bapbong',
    version: opts.version ?? '0.0.1',
  });

  /** Resolve the session and translate contract errors into teaching text. */
  const withSession = async (
    id: string | undefined,
    fn: (session: DocumentSession) => Promise<ReturnType<typeof json>>,
  ): Promise<ReturnType<typeof json> | ReturnType<typeof errorText>> => {
    const session = await provider.get(id);
    if (!session) {
      return errorText(
        id === undefined
          ? 'No document is open. Ask the user to open a document first.'
          : `No document with id ${JSON.stringify(id)}.`,
      );
    }
    try {
      return await fn(session);
    } catch (err) {
      if (err instanceof AnchorError || err instanceof VersionConflictError || err instanceof NoDocumentError) {
        return errorText(err.message);
      }
      throw err;
    }
  };

  server.registerTool(
    'get_document',
    {
      title: 'Read the document',
      description:
        'Read the whole document as numbered blocks (paragraphs, headings — table-cell paragraphs included, in reading order). ' +
        'Returns docVersion: pass it as expectedVersion to mutation tools so concurrent edits are detected. ' +
        'Block indexes are only stable within one docVersion.',
      inputSchema: { documentId },
    },
    async ({ documentId: id }) => withSession(id, async (s) => json(await s.snapshot())),
  );

  server.registerTool(
    'find_text',
    {
      title: 'Find text',
      description:
        'Find every occurrence of a text in the document. Matches are within one paragraph (they never span paragraphs or inline objects). ' +
        'Returns each match with its block index, 1-based occurrence number, and surrounding context.',
      inputSchema: { documentId, query: z.string().min(1).describe('Exact text to find (case-sensitive).') },
    },
    async ({ documentId: id, query }) => withSession(id, async (s) => json({ matches: await s.find(query) })),
  );

  server.registerTool(
    'replace_text',
    {
      title: 'Replace text',
      description:
        'Replace one occurrence of exact text. old_text must match exactly once in the document — if it matches more, ' +
        'either pass occurrence or use a longer, unique anchor (include surrounding words). Formatting of the replaced range is kept.',
      inputSchema: {
        documentId,
        old_text: z.string().min(1).describe('Exact existing text (must match uniquely, or pass occurrence).'),
        new_text: z.string().describe('Replacement text (empty string deletes).'),
        occurrence,
        expectedVersion,
      },
    },
    async ({ documentId: id, old_text, new_text, occurrence: occ, expectedVersion: ver }) =>
      withSession(id, async (s) => json(await s.replaceText(old_text, new_text, { occurrence: occ, expectedVersion: ver }))),
  );

  server.registerTool(
    'insert_content',
    {
      title: 'Insert content',
      description:
        'Insert new paragraphs. Each line of `content` becomes one paragraph. Anchor by exact text with position before/after ' +
        '(the paragraph containing the anchor), or position document_end to append.',
      inputSchema: {
        documentId,
        content: z.string().min(1).describe('Plain text; every line becomes a paragraph.'),
        position: z.enum(['before', 'after', 'document_end']),
        anchor_text: z.string().optional().describe('Required for before/after: exact text inside the anchor paragraph.'),
        occurrence,
        expectedVersion,
      },
    },
    async ({ documentId: id, content, position, anchor_text, occurrence: occ, expectedVersion: ver }) =>
      withSession(id, async (s) => {
        if (position === 'document_end') {
          return json(await s.insertContent(content, { position: 'document_end' }, { expectedVersion: ver }));
        }
        if (!anchor_text) return errorText('anchor_text is required when position is before/after.');
        return json(
          await s.insertContent(content, { position, text: anchor_text, occurrence: occ }, { expectedVersion: ver }),
        );
      }),
  );

  server.registerTool(
    'apply_formatting',
    {
      title: 'Apply formatting',
      description:
        'Format one occurrence of exact text: character marks (bold/italic/underline/strike — true applies, false removes) ' +
        'and/or paragraph alignment of the containing paragraph. Same anchoring rules as replace_text.',
      inputSchema: {
        documentId,
        target_text: z.string().min(1).describe('Exact text to format (must match uniquely, or pass occurrence).'),
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        underline: z.boolean().optional(),
        strike: z.boolean().optional(),
        align: z.enum(['left', 'center', 'right', 'justify']).optional(),
        occurrence,
        expectedVersion,
      },
    },
    async ({ documentId: id, target_text, occurrence: occ, expectedVersion: ver, ...format }) =>
      withSession(id, async (s) => json(await s.applyFormatting(target_text, format, { occurrence: occ, expectedVersion: ver }))),
  );

  server.registerTool(
    'update_image',
    {
      title: 'Resize / rotate an image',
      description:
        'Resize and/or rotate one image. Read get_document first: blocks list their images, and ' +
        '(block_index, image_index) addresses one. Omitted fields keep their current value; width/height are CSS px, ' +
        'rotation is clockwise degrees around the image center (normalized to 0-360; 0 = upright). ' +
        'Block indexes change with every edit — pass expectedVersion from your last read.',
      inputSchema: {
        documentId,
        block_index: z.number().int().min(0).describe('Index of the block holding the image (from get_document).'),
        image_index: z.number().int().min(0).optional().describe('0-based image within the block (default 0, the first).'),
        width: z.number().positive().optional().describe('New width in CSS px.'),
        height: z.number().positive().optional().describe('New height in CSS px.'),
        rotation: z.number().optional().describe('Clockwise degrees around the center.'),
        expectedVersion,
      },
    },
    async ({ documentId: id, block_index, image_index, width, height, rotation, expectedVersion: ver }) =>
      withSession(id, async (s) => {
        if (width === undefined && height === undefined && rotation === undefined) {
          return errorText('Pass at least one of width, height, rotation.');
        }
        return json(
          await s.updateImage(block_index, image_index ?? 0, { width, height, rotation }, { expectedVersion: ver }),
        );
      }),
  );

  server.registerTool(
    'save_document',
    {
      title: 'Save the document',
      description: 'Persist the document to its backing store (file on desktop). Usually optional — the desktop autosaves.',
      inputSchema: { documentId },
    },
    async ({ documentId: id }) =>
      withSession(id, async (s) => {
        await s.save();
        return json({ saved: true });
      }),
  );

  if (opts.selection) {
    server.registerTool(
      'get_selection',
      {
        title: "Read the user's selection",
        description: "The text the user currently has selected in the editor, or null when nothing is selected.",
        inputSchema: { documentId },
      },
      async ({ documentId: id }) =>
        withSession(id, async (s) => json((await s.getSelection?.()) ?? { selection: null })),
    );
  }

  server.registerResource(
    'document',
    'bapbong://document',
    {
      title: 'Open document',
      description: 'The currently open document as plain text (one line per block).',
      mimeType: 'text/plain',
    },
    async (uri) => {
      const session = await provider.get(undefined);
      if (!session) return { contents: [{ uri: uri.href, text: '(no document open)' }] };
      const snap = await session.snapshot();
      return { contents: [{ uri: uri.href, text: snap.blocks.map((b) => b.text).join('\n') }] };
    },
  );

  return server;
}
