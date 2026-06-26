import type { EditorPlugin, PluginContext } from '@shadow-garden/bapbong-contracts';

/** The `link` mark href covering doc position `pos`, or null. */
function linkHrefAt(state: PluginContext['state'], pos: number): string | null {
  const linkType = state.schema.marks['link'];
  if (!linkType) return null;
  const $pos = state.doc.resolve(pos);
  // `link` is non-inclusive, so a boundary position belongs to the side that
  // carries the mark — check the node after the point, then the one before.
  const after = $pos.nodeAfter;
  const before = $pos.nodeBefore;
  const mark =
    (after && linkType.isInSet(after.marks)) || (before && linkType.isInSet(before.marks)) || null;
  return mark ? (mark.attrs['href'] as string) : null;
}

/** Open an external hyperlink in a new tab. In-document anchors (`#…`) are left
 *  to a future "scroll to bookmark" follow-up. */
function openHref(href: string): void {
  if (!href || href.startsWith('#')) return;
  window.open(href, '_blank', 'noopener,noreferrer');
}

/**
 * Internal plugin: **Ctrl/Cmd-click a hyperlink to open it** (the Word / Google
 * Docs convention — a plain click still places the caret so links stay
 * editable). Claims the press when a modified primary click lands on a `link`
 * mark, so the editor skips caret placement and the URL opens in a new tab.
 *
 * Pairs with the export hyperlink round-trip (M6) and the accessible mirror
 * (M10), whose `<a href>` elements expose the same links to assistive tech.
 */
export function hyperlinkPlugin(): EditorPlugin {
  let ctx: PluginContext | null = null;
  return {
    name: 'hyperlink',
    setup(c) {
      ctx = c;
    },
    onPointer(ev): boolean {
      if (ev.type !== 'down' || !ctx || ev.pos == null) return false;
      if (!(ev.metaKey || ev.ctrlKey) || ev.buttons !== 1) return false; // modifier + primary only
      const href = linkHrefAt(ctx.state, ev.pos);
      if (!href) return false;
      openHref(href);
      return true; // claim → suppress caret placement
    },
  };
}
