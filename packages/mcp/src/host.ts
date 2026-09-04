/**
 * Host-process surface: everything in ./session plus the MCP server factory
 * (which brings the @modelcontextprotocol/sdk + zod runtime deps). The
 * desktop Bun process imports THIS; it deliberately excludes HeadlessSession
 * so a host that proxies to a live editor doesn't drag the docx pipeline in.
 */
export * from './session.js';
export { createMcpServer, type CreateMcpServerOptions } from './lib/server.js';
export {
  defineCommand,
  errorText,
  isOffered,
  json,
  registerCommands,
  withSession,
  type AgentCommand,
  type CommandArgs,
  type CommandEffect,
  type CommandRequirement,
  type CommandResult,
  type HostCapabilities,
} from './lib/catalog.js';
export * from './lib/document-commands.js';
