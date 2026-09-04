/**
 * The MCP adapter over a {@link SessionProvider} — the host-agnostic half of
 * M15. Hosts supply sessions (desktop: the live editor; server: headless
 * EditorStates) and a transport (loopback Bun.serve, TLS gateway, stdio…).
 *
 * Everything the agent sees — tool names, schemas, anchoring/versioning
 * semantics, error texts that teach the retry — lives in the command records
 * (./document-commands); this file only registers them on an McpServer.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { SessionProvider } from './contract.js';
import { registerCommands } from './catalog.js';
import { documentCommands } from './document-commands.js';

export interface CreateMcpServerOptions {
  name?: string;
  version?: string;
  /** Offer get_selection (hosts with a live user selection — the desktop). */
  selection?: boolean;
}

export function createMcpServer(
  provider: SessionProvider,
  opts: CreateMcpServerOptions = {},
): McpServer {
  const server = new McpServer({
    name: opts.name ?? 'bapbong',
    version: opts.version ?? '0.0.1',
  });

  registerCommands(server, documentCommands, provider, {
    selection: opts.selection,
  });

  server.registerResource(
    'document',
    'bapbong://document',
    {
      title: 'Open document',
      description:
        'The currently open document as plain text (one line per block).',
      mimeType: 'text/plain',
    },
    async (uri) => {
      const session = await provider.get(undefined);
      if (!session)
        return { contents: [{ uri: uri.href, text: '(no document open)' }] };
      const snap = await session.snapshot();
      return {
        contents: [
          { uri: uri.href, text: snap.blocks.map((b) => b.text).join('\n') },
        ],
      };
    },
  );

  return server;
}
