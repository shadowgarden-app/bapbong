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
import { registerCommands, registerDocumentResource } from './catalog.js';
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

  registerDocumentResource(server, provider);

  return server;
}
