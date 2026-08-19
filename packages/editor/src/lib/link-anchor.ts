import type { BapbongEditor } from './bapbong-editor';

/**
 * What the link panel should say about a link the caret sits in when that
 * link points INSIDE the document (`#_Toc…` — every table-of-contents entry,
 * a cross-reference): Word never shows the raw bookmark id, it shows where
 * the link GOES and offers the jump; and when the link is field output (a TOC
 * entry) it does not let you edit or unlink it one by one — updating the
 * field is the way to change those.
 *
 * Decided HERE, once, from the built-in `hyperlink` and `toc` plugins every
 * editor carries — not in each host. The desktop shell had copied the
 * playground's link-panel wiring before this rule existed and never gained
 * it, so the same document's TOC was a jump popover in one app and a raw
 * "#_Toc376184353" with an unlink button in the other. Returns null for an
 * external link (or none), which the panel renders as a URL.
 *
 * The shape is `LinkPanelOptions['internal']` (bapbong-ui), structurally —
 * this package does not import the UI.
 */
export function internalLinkFor(
  editor: BapbongEditor,
  info: { href: string; from: number } | null | undefined,
): { label: string | null; generated: boolean; onGo(): void } | null {
  if (!info?.href) return null;
  const link = editor.plugin('hyperlink');
  if (!link.isInternal(info.href)) return null;
  const href = info.href;
  return {
    label: link.describe(href),
    generated: editor.plugin('toc').fieldAt(info.from) !== null,
    onGo: () => {
      link.follow(href);
      editor.focus();
    },
  };
}
