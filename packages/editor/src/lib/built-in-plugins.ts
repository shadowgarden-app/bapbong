import type { EditorPlugin } from '@shadow-garden/bapbong-contracts';
import { findPlugin } from './find-plugin';
import { tableResizePlugin } from './table-resize-plugin';

/**
 * Built-in ("internal") plugins — shipped with the editor, no install needed
 * (unlike external `@shadow-garden/bapbong-*` plugins the host passes via
 * `{ plugins }`). The editor registers these (keyed by `name`) ahead of the
 * host's plugins and exposes the ones with a richer API as typed handles
 * (e.g. `editor.find`).
 *
 * Add a new internal plugin by importing its factory and returning it here.
 * This is a FACTORY, not a shared array: each plugin holds per-editor state, so
 * every editor needs its own fresh instances.
 */
export function createBuiltins(): EditorPlugin[] {
  return [findPlugin(), tableResizePlugin()];
}
