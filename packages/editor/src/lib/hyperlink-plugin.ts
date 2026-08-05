import type {
  EditorPlugin,
  PluginContext,
} from '@shadow-garden/bapbong-contracts';
import {
  anchorName,
  bookmarkLabel,
  findBookmark,
} from '@shadow-garden/bapbong-model';

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
    (after && linkType.isInSet(after.marks)) ||
    (before && linkType.isInSet(before.marks)) ||
    null;
  return mark ? (mark.attrs['href'] as string) : null;
}

/** Open an external hyperlink in a new tab. Scheme-less hrefs (a stored
 *  "www.google.com") get https:// — window.open would treat them as relative
 *  paths and silently 404 inside the app. */
function openExternal(href: string): void {
  const url = /^[a-z][a-z0-9+.-]*:/i.test(href) ? href : `https://${href}`;
  window.open(url, '_blank', 'noopener,noreferrer');
}

/** Public API — the host's link UI resolves and follows links through this
 *  rather than reimplementing bookmark lookup. */
export interface HyperlinkPlugin extends EditorPlugin {
  /** Follow a link: external URLs open in a new tab, `#name` scrolls to the
   *  bookmarked paragraph and puts the caret there. False when an internal
   *  target no longer exists (a stale anchor). */
  follow(href: string): boolean;
  /** What a reader should see for this href: the target paragraph's text for
   *  an internal anchor (Word never shows the machine-generated `_Toc…` id),
   *  the URL itself otherwise. Null when an internal target is missing. */
  describe(href: string): string | null;
  /** Whether the href points inside this document. */
  isInternal(href: string): boolean;
}

declare module '@shadow-garden/bapbong-contracts' {
  interface EditorPluginHandles {
    hyperlink: HyperlinkPlugin;
  }
}

/**
 * Internal plugin: **Ctrl/Cmd-click a hyperlink to follow it** (the Word /
 * Google Docs convention — a plain click still places the caret so links stay
 * editable). Claims the press when a modified primary click lands on a `link`
 * mark, so the editor skips caret placement.
 *
 * External URLs open in a new tab; in-document anchors (`#_Toc89595219`, what
 * every TOC entry points at) scroll to the bookmarked paragraph and place the
 * caret there, which is what Ctrl-clicking a TOC entry does in Word.
 *
 * Pairs with the export hyperlink round-trip (M6) and the accessible mirror
 * (M10), whose `<a href>` elements expose the same links to assistive tech.
 */
export function hyperlinkPlugin(): HyperlinkPlugin {
  let ctx: PluginContext | null = null;

  const follow = (href: string): boolean => {
    if (!href) return false;
    const name = anchorName(href);
    if (!name) {
      openExternal(href);
      return true;
    }
    if (!ctx) return false;
    const pos = findBookmark(ctx.state.doc, name);
    if (pos === null) return false; // stale anchor — nothing to jump to
    ctx.setSelection(pos);
    ctx.scrollToPos(pos);
    return true;
  };

  return {
    name: 'hyperlink',
    setup(c) {
      ctx = c;
    },
    follow,
    describe(href) {
      const name = anchorName(href);
      if (!name) return href;
      return ctx ? bookmarkLabel(ctx.state.doc, name) : null;
    },
    isInternal: (href) => anchorName(href) !== null,
    onPointer(ev): boolean {
      if (ev.type !== 'down' || !ctx || ev.pos == null) return false;
      if (!(ev.metaKey || ev.ctrlKey) || ev.buttons !== 1) return false; // modifier + primary only
      const href = linkHrefAt(ctx.state, ev.pos);
      if (!href) return false;
      // A dead internal anchor is NOT claimed: let the click place the caret
      // instead of swallowing it with nothing to show for it.
      return follow(href);
    },
  };
}
