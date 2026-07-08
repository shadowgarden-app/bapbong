/**
 * `@shadow-garden/bapbong-mcp` — MCP server for AI agents (M15).
 *
 * Host-agnostic by design: tools are written once against the
 * {@link DocumentSession} port; the desktop app binds its live editor, a
 * server binds {@link HeadlessSession} (headless ProseMirror EditorStates).
 * See PLAN.md §M15 for the architecture and the frozen contract semantics.
 */
export * from './lib/contract.js';
export { PmDocSession, type PmSessionHost } from './lib/pm-session.js';
export { HeadlessSession, type HeadlessSessionOptions } from './lib/headless-session.js';
export { createMcpServer, type CreateMcpServerOptions } from './lib/server.js';
export {
  executeOp,
  reviveError,
  RemoteSession,
  type SessionOpName,
  type SessionOpRequest,
  type SessionOpResponse,
} from './lib/wire.js';
