/**
 * `@shadow-garden/bapbong-mcp` — MCP server for AI agents (M15).
 *
 * Host-agnostic by design: tools are written once against the
 * {@link DocumentSession} port; the desktop app binds its live editor, a
 * server binds {@link HeadlessSession} (headless ProseMirror EditorStates).
 */
export * from './lib/contract.js';
export { PmDocSession, type PmSessionHost } from './lib/pm-session.js';
export {
  contentToNodes,
  lengthToPx,
  referencedTableStyles,
  type Block,
  type BuildOptions,
  type Cell,
  type Content,
  type Inline,
  type ParagraphBlock,
  type TableBlock,
  type TableStyleSource,
  type TabStop,
} from './lib/blocks.js';
export {
  CONTENT_GRAMMAR,
  blockSchema,
  contentSchema,
} from './lib/blocks-schema.js';
export {
  HeadlessSession,
  type HeadlessSessionOptions,
} from './lib/headless-session.js';
export { createMcpServer, type CreateMcpServerOptions } from './lib/server.js';
export {
  defineCommand,
  errorText,
  isOffered,
  json,
  registerCommands,
  registerDocumentResource,
  withSession,
  type AgentCommand,
  type CallContext,
  type CommandArgs,
  type DocumentEffect,
  type CommandRequirement,
  type CommandResult,
  type HostCapabilities,
} from './lib/catalog.js';
export * from './lib/document-commands.js';
export {
  executeOp,
  reviveError,
  RemoteSession,
  type SessionOpName,
  type SessionOpRequest,
  type SessionOpResponse,
} from './lib/wire.js';
