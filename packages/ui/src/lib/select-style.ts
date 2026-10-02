/**
 * The library's `<select>`s drawn as ours, not the platform's.
 *
 * A select keeps its native look — macOS draws a grey popup button with
 * up/down arrows in WKWebView — whatever background the theme gives it,
 * unless `appearance: none` says otherwise. The toolbar's Styles / Font /
 * Size boxes were that grey button on the desktop's glass, in every
 * theme. With the native look off, the arrow is ours: a chevron any theme
 * can replace through `--bb-ui-select-chevron`; the default grey reads on
 * light and dark alike.
 *
 * Append it AFTER a rule's `background` shorthand (which would reset the
 * image), and theme hover / press with `background-color`, never the
 * shorthand, or the chevron disappears under the pointer.
 */
export const SELECT_CHEVRON =
  'appearance:none;-webkit-appearance:none;' +
  "background-image:var(--bb-ui-select-chevron,url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 10 10'%3E%3Cpath d='M2 3.5l3 3 3-3' fill='none' stroke='%238a8780' stroke-width='1.4' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E\"));" +
  'background-repeat:no-repeat;background-position:right 7px center;background-size:10px 10px;padding-right:22px';
