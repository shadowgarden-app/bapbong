import type {
  CommentNode,
  EditorChange,
  EditorPlugin,
  PluginContext,
  RangeDecoration,
} from '@shadow-garden/bapbong-contracts';

/** Highlight painted behind text covered by an unresolved comment. */
const COMMENT_TINT = 'rgba(255, 193, 7, 0.28)';

export interface CommentTintPlugin extends EditorPlugin {
  /** Toggle the tint (e.g. off in the "hide" view). Triggers a repaint. */
  setTintEnabled(enabled: boolean): void;
}

/**
 * A framework-agnostic editor plugin that paints the comment tint as generic
 * {@link RangeDecoration}s — reading the `comment` mark + `doc.attrs.comments`
 * to decide which ranges are unresolved. Replaces the painter's old hardcoded
 * comment-tint path (and the editor's `setSuppressedComments`).
 *
 * Ranges are recomputed only when the doc changes (cached for cheap per-scroll
 * repaints). This will move into `@shadow-garden/bapbong-comments` alongside the
 * comment mark + transactions in a later step.
 */
export function commentsPlugin(): CommentTintPlugin {
  let ctx: PluginContext | null = null;
  let enabled = true;
  let cache: RangeDecoration[] = [];

  const recompute = (): void => {
    if (!ctx) {
      cache = [];
      return;
    }
    const doc = ctx.state.doc;
    const comments = (doc.attrs['comments'] as CommentNode[] | null) ?? [];
    const resolvedRoots = new Set(
      comments.filter((c) => c.parentId == null && c.resolved).map((c) => c.id),
    );
    const out: RangeDecoration[] = [];
    doc.descendants((node, pos) => {
      if (!node.isText) return;
      const mark = node.marks.find((m) => m.type.name === 'comment');
      if (!mark) return;
      const ids = mark.attrs['ids'] as number[];
      // Tint unless every covering comment is a resolved root.
      if (ids.some((id) => !resolvedRoots.has(id))) {
        out.push({ from: pos, to: pos + node.nodeSize, kind: 'background', color: COMMENT_TINT });
      }
    });
    cache = out;
  };

  return {
    name: 'comments',
    setup(c: PluginContext) {
      ctx = c; // doc loads later — first recompute happens on the first change
    },
    onChange(change: EditorChange) {
      if (!change.docChanged) return;
      recompute();
      ctx?.requestPaint(); // reflect new / resolved / deleted comments in the tint
    },
    decorations(): RangeDecoration[] {
      return enabled ? cache : [];
    },
    setTintEnabled(value: boolean) {
      if (enabled === value) return;
      enabled = value;
      ctx?.requestPaint();
    },
  };
}
