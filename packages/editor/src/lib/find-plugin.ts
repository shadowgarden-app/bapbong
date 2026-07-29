import type {
  EditorChange,
  EditorPlugin,
  PluginContext,
  RangeDecoration,
} from '@shadow-garden/bapbong-contracts';

/** All matches get a soft highlight; the active one a stronger fill. */
const MATCH = 'rgba(255, 214, 0, 0.40)';
const ACTIVE = 'rgba(255, 138, 0, 0.65)';

export interface FindState {
  query: string;
  /** Total matches. */
  count: number;
  /** 1-based index of the active match (0 when there are none). */
  active: number;
}

export interface FindPlugin extends EditorPlugin {
  setQuery(q: string): void;
  next(): void;
  prev(): void;
  clear(): void;
  /** Replace the active match with `text` (one undoable transaction). */
  replaceCurrent(text: string): void;
  /** Replace every match with `text` (one undoable transaction). */
  replaceAll(text: string): void;
  /** Subscribe to query/count/active changes (for the host's find bar). */
  onState(cb: (s: FindState) => void): () => void;
}

/** Register the handle's type so `editor.plugin('find')` is typed without the
 *  core knowing this plugin exists. */
declare module '@shadow-garden/bapbong-contracts' {
  interface EditorPluginHandles {
    find: FindPlugin;
  }
}

/**
 * Find-and-replace as a built-in ("internal") editor plugin: the editor
 * instantiates it and reaches it as `editor.plugin('find')`. Highlighting rides the
 * decoration pipeline (read side); replace dispatches transactions (write side)
 * — exercising both halves of the plugin contract. Matches are recomputed only
 * when the query or the doc changes (cached for cheap per-scroll repaints).
 *
 * Matching is case-insensitive, within a single text node (cross-node matches
 * are out of scope for v1). The host owns the find-bar UI.
 */
export function findPlugin(): FindPlugin {
  let ctx: PluginContext | null = null;
  let query = '';
  let matches: { from: number; to: number }[] = [];
  let active = 0; // index into matches
  const listeners = new Set<(s: FindState) => void>();

  const notify = (): void => {
    const s: FindState = {
      query,
      count: matches.length,
      active: matches.length ? active + 1 : 0,
    };
    for (const cb of listeners) cb(s);
  };

  const recompute = (): void => {
    matches = [];
    if (ctx && query) {
      const needle = query.toLowerCase();
      ctx.state.doc.descendants((node, pos) => {
        if (!node.isText || !node.text) return;
        const hay = node.text.toLowerCase();
        for (
          let i = hay.indexOf(needle);
          i !== -1;
          i = hay.indexOf(needle, i + query.length)
        ) {
          matches.push({ from: pos + i, to: pos + i + query.length });
        }
      });
    }
    if (active >= matches.length)
      active = matches.length ? matches.length - 1 : 0;
  };

  const revealActive = (): void => {
    const m = matches[active];
    if (m && ctx) {
      ctx.setSelection(m.from, m.to);
      ctx.scrollToPos(m.from);
    }
  };

  return {
    name: 'find',
    setup(c: PluginContext) {
      ctx = c;
    },
    onChange(change: EditorChange) {
      if (!change.docChanged) return;
      recompute(); // positions shifted (and replaces land here too)
      ctx?.requestPaint();
      notify();
    },
    decorations(): RangeDecoration[] {
      return matches.map((m, i) => ({
        from: m.from,
        to: m.to,
        kind: 'background',
        color: i === active ? ACTIVE : MATCH,
      }));
    },
    setQuery(q: string) {
      query = q;
      active = 0;
      recompute();
      revealActive();
      ctx?.requestPaint();
      notify();
    },
    next() {
      if (!matches.length) return;
      active = (active + 1) % matches.length;
      revealActive();
      ctx?.requestPaint();
      notify();
    },
    prev() {
      if (!matches.length) return;
      active = (active - 1 + matches.length) % matches.length;
      revealActive();
      ctx?.requestPaint();
      notify();
    },
    clear() {
      query = '';
      matches = [];
      active = 0;
      ctx?.requestPaint();
      notify();
    },
    replaceCurrent(text: string) {
      const m = matches[active];
      if (!m || !ctx) return;
      // onChange (docChanged) recomputes matches + notifies.
      ctx.dispatch(ctx.state.tr.insertText(text, m.from, m.to));
    },
    replaceAll(text: string) {
      if (!matches.length || !ctx) return;
      const tr = ctx.state.tr;
      // High→low so each replace doesn't shift the positions still to come.
      for (let i = matches.length - 1; i >= 0; i--) {
        tr.insertText(text, matches[i].from, matches[i].to);
      }
      ctx.dispatch(tr);
    },
    onState(cb: (s: FindState) => void): () => void {
      listeners.add(cb);
      return () => listeners.delete(cb);
    },
  };
}
