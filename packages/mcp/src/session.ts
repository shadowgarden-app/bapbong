/**
 * Browser-safe surface: the contract, the shared PM session semantics, and
 * the wire protocol — zero runtime dependencies (prosemirror imports are
 * type-only). The desktop WebView imports THIS to execute session ops against
 * its live editor; it must never pull the MCP SDK into the web bundle.
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
  type TableEdit,
  type TableStyleSource,
  type TabStop,
} from './lib/blocks.js';
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
