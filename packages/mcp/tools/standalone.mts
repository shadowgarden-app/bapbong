/**
 * Standalone MCP document server — dev tool + seed of the future server host.
 *
 *   bun packages/mcp/tools/standalone.mts <file.docx> [port]
 *
 * Serves the document over MCP Streamable HTTP at http://127.0.0.1:<port>/mcp
 * (default 3845). Edits stay in memory; save_document writes a SIBLING file
 * `<file>.ai.docx` — the original is never touched (dev-tool safety).
 *
 * Connect Claude Desktop through the generic stdio→HTTP proxy while the real
 * shim doesn't exist yet (M15-P3):
 *
 *   { "mcpServers": { "bapbong": {
 *       "command": "npx",
 *       "args": ["-y", "mcp-remote", "http://127.0.0.1:3845/mcp"] } } }
 *
 * Stateless transport: every request gets a fresh server+transport pair; the
 * document session itself lives in this process for the server's lifetime.
 */
import { readFileSync } from 'node:fs';
import { WebStandardStreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js';
import { createMcpServer, HeadlessSession } from '../src/index.ts';

const [file, portArg] = process.argv.slice(2);
if (!file) {
  console.error('usage: bun packages/mcp/tools/standalone.mts <file.docx> [port]');
  process.exit(1);
}
const port = Number(portArg) || 3845;
const outPath = file.replace(/\.docx$/i, '') + '.ai.docx';

const session = await HeadlessSession.open(readFileSync(file), {
  name: file.split('/').pop(),
  onSave: async (bytes) => {
    await Bun.write(outPath, bytes);
    console.log(`[standalone] saved ${bytes.byteLength}B -> ${outPath}`);
  },
});
const provider = { get: async () => session };

Bun.serve({
  hostname: '127.0.0.1',
  port,
  idleTimeout: 120,
  async fetch(req) {
    const url = new URL(req.url);
    if (url.pathname !== '/mcp') return new Response('bapbong mcp standalone — POST /mcp', { status: 404 });
    const server = createMcpServer(provider, { name: 'bapbong-standalone' });
    const transport = new WebStandardStreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    await server.connect(transport);
    const res = await transport.handleRequest(req);
    console.log(`[standalone] ${req.method} /mcp -> ${res.status}`);
    return res;
  },
});

console.log(`[standalone] serving ${file}`);
console.log(`[standalone] MCP endpoint: http://127.0.0.1:${port}/mcp`);
console.log(`[standalone] save_document writes to: ${outPath}`);
