/**
 * Browser-safe surface: the contract, the shared PM session semantics, and
 * the wire protocol — zero runtime dependencies (prosemirror imports are
 * type-only). The desktop WebView imports THIS to execute session ops against
 * its live editor; it must never pull the MCP SDK into the web bundle.
 */
export * from './lib/contract.js';
export { PmDocSession, type PmSessionHost } from './lib/pm-session.js';
export {
  ReadOnlySession,
  type ReadOnlySessionOptions,
} from './lib/read-only-session.js';
export {
  executeOp,
  reviveError,
  RemoteSession,
  type SessionOpName,
  type SessionOpRequest,
  type SessionOpResponse,
} from './lib/wire.js';
