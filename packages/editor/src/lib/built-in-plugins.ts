import { findPlugin, type FindPlugin } from './find-plugin';

/**
 * Built-in ("internal") plugins — shipped with the editor, no install needed
 * (unlike external `@shadow-garden/bapbong-*` plugins the host passes in via
 * `{ plugins }`). The editor instantiates these per construction, registers
 * them ahead of the host's plugins, and exposes each as a typed handle
 * (e.g. `editor.find`).
 *
 * Add a new internal plugin by importing its factory, adding a field to
 * {@link Builtins} + {@link createBuiltins}, and exposing a getter on
 * `BapbongEditor`. This is a FACTORY, not a shared array: each plugin holds
 * per-editor state, so every editor needs its own fresh instances.
 */
export interface Builtins {
  find: FindPlugin;
}

export function createBuiltins(): Builtins {
  return {
    find: findPlugin(),
  };
}
