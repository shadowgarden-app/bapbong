/**
 * bapbong-mcp shim — connects stdio MCP clients (Claude Desktop) to the
 * RUNNING bapbong desktop app.
 *
 * Claude Desktop only spawns stdio servers, but the document lives in the
 * already-running app; this shim finds the app through its discovery file
 * (port + token, written at app launch) and pipes MCP messages 1:1 between
 * stdio and the app's Streamable HTTP endpoint. No tool logic lives here —
 * it is a disposable adapter (Claude Code connects to the HTTP URL directly).
 *
 * claude_desktop_config.json:
 *   { "mcpServers": { "bapbong": {
 *       "command": "bun",
 *       "args": ["<repo>/packages/mcp/tools/shim.mts"] } } }
 */
import { readFileSync } from 'node:fs';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';

const DISCOVERY = `${process.env.HOME}/Library/Application Support/bapbong/mcp.json`;

function die(message: string): never {
  console.error(`bapbong-mcp: ${message}`);
  process.exit(1);
}

let discovery: { port: number; token: string; endpoint: string };
try {
  discovery = JSON.parse(readFileSync(DISCOVERY, 'utf8'));
} catch {
  die('bapbong is not running (no discovery file). Open the bapbong app, then reconnect.');
}

const url = new URL(`http://127.0.0.1:${discovery.port}${discovery.endpoint}`);
const http = new StreamableHTTPClientTransport(url, {
  requestInit: { headers: { Authorization: `Bearer ${discovery.token}` } },
});
const stdio = new StdioServerTransport();

// A dumb, faithful pipe: every JSON-RPC message crosses unchanged, so the
// app's server is the single source of truth for tools and semantics.
stdio.onmessage = (message) => {
  void http.send(message).catch((err) => {
    console.error(`bapbong-mcp: lost the app (${err instanceof Error ? err.message : err})`);
    process.exit(1);
  });
};
http.onmessage = (message) => {
  void stdio.send(message);
};
stdio.onclose = () => process.exit(0);
http.onclose = () => {
  console.error('bapbong-mcp: the app closed the connection.');
  process.exit(1);
};

await http.start().catch(() => die('bapbong is not reachable — is the app still running?'));
await stdio.start();
console.error(`bapbong-mcp: connected to ${url.href}`);
