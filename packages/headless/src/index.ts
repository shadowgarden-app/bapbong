/**
 * `@shadow-garden/bapbong-headless` — the backend façade.
 *
 * One install that re-exports the whole **isomorphic tier** (no DOM, no canvas):
 *   - `bapbong-contracts` — shared types (PageConfig, Align, Command, …)
 *   - `bapbong-model`      — the ProseMirror `schema` + numbering
 *   - `bapbong-docx`       — `importDocx` / `exportDocx`
 *   - `bapbong-commands`   — headless editor ops (toggleMark, setAlign, …)
 *
 * The whole surface runs on Node/server. A backend imports a `.docx`, builds a
 * ProseMirror `EditorState` from the doc, mutates it through the same `commands`
 * the editor UI uses, then exports a `.docx` back — with zero browser deps.
 *
 *   import { importDocx, exportDocx, setAlign } from '@shadow-garden/bapbong-headless';
 *   import { EditorState } from 'prosemirror-state';
 *
 * Layout/pagination is included too: font-file metrics made it DOM-free, so a
 * backend can paginate (e.g. server-side PDF) with the SAME pure `layout` engine
 * and measurers the editor uses — identical measurer + doc ⇒ identical page
 * breaks (client↔server parity). Only the DOM-free measuring surface is
 * re-exported; the canvas measurers stay out (they need a browser).
 *
 * Do NOT add `bapbong-editor` or `bapbong-view` here — they drag in canvas/DOM
 * and would break Node usage. The `scope:headless` lint boundary enforces this.
 */
export * from '@shadow-garden/bapbong-contracts';
export * from '@shadow-garden/bapbong-model';
export * from '@shadow-garden/bapbong-docx';
export * from '@shadow-garden/bapbong-commands';
export {
  layout,
  createLayoutCache,
} from '@shadow-garden/bapbong-layout-engine';
export {
  FontRegistry,
  createFontRegistryMeasurer,
  createFontRegistryMetrics,
  createApproxMeasurer,
  createApproxMetrics,
  fontToCss,
  type FontVariant,
} from '@shadow-garden/bapbong-measuring';
// How a FontSpec becomes glyphs on a context. Prefer this over `fontToCss`:
// the shorthand cannot carry tracking, so a context built from it alone draws
// text narrower than the engine measured it.
export {
  applyGlyphSpec,
  glyphCount,
  glyphKey,
  sameGlyphRun,
  type GlyphContext,
} from '@shadow-garden/bapbong-contracts';

// A handful of names are declared (structurally-identically) in more than one
// layer. Re-export the canonical contract versions explicitly so they stay
// accessible from the façade — an explicit re-export wins over the otherwise
// ambiguous `export *`, which would silently drop these names.
export type {
  Align,
  PageConfig,
  CommentData,
} from '@shadow-garden/bapbong-contracts';
