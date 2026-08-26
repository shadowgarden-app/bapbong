import type { EditorPlugin } from '@shadow-garden/bapbong-contracts';
import { equationPlugin } from './equation-plugin';
import { findPlugin } from './find-plugin';
import { imageResizePlugin } from './image-resize-plugin';
import { tableResizePlugin } from './table-resize-plugin';
import { tableSelectionPlugin } from './table-selection-plugin';
import { hyperlinkPlugin } from './hyperlink-plugin';
import { tocPlugin } from './toc-plugin';

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
  // Order matters for the pointer hook: hyperlink claims modifier-clicks
  // first, then the resizers claim their geometry (table column borders,
  // then image boxes) before table selection.
  return [
    hyperlinkPlugin(),
    tocPlugin(),
    equationPlugin(),
    findPlugin(),
    tableResizePlugin(),
    imageResizePlugin(),
    tableSelectionPlugin(),
  ];
}
