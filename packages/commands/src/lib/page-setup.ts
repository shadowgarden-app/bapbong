import type { EditorState } from 'prosemirror-state';
import type { Command, PageConfig } from '@shadow-garden/bapbong-contracts';

/**
 * Page-setup commands (File → Page setup): paper size + orientation. The Word
 * model: these are document/section properties, not content — they live on
 * `doc.attrs.page` and are edited via setDocAttribute so they undo cleanly.
 * Export maps known sizes back to canonical twips (Word then names the size
 * "A4"/"Letter" instead of "Custom").
 */

/** Common paper sizes in CSS px @96dpi, portrait. Derived from their canonical
 *  twip dimensions (÷15, rounded) so import → command → export round-trips. */
export const PAPER_SIZES = {
  a4: { width: 794, height: 1123 },
  letter: { width: 816, height: 1056 },
  legal: { width: 816, height: 1344 },
  a3: { width: 1123, height: 1587 },
  a5: { width: 559, height: 794 },
} as const;
export type PaperSize = keyof typeof PAPER_SIZES;

/** A4 @96dpi with 1in margins — what layout assumes when a doc predates the
 *  page attr (same fallback as the view and the exporter). */
const A4_PAGE: PageConfig = {
  width: 794,
  height: 1123,
  margin: { top: 96, right: 96, bottom: 96, left: 96 },
};

/** The doc's page geometry (a copy, safe to mutate). */
function currentPage(state: EditorState): PageConfig {
  const raw = state.doc.attrs['page'] as PageConfig | null;
  const p = raw ?? A4_PAGE;
  return { width: p.width, height: p.height, margin: { ...p.margin } };
}

const isLandscape = (p: PageConfig) => p.width > p.height;

/**
 * Set the page orientation. Landscape swaps width/height; margins are kept as
 * named (top stays top) — with the symmetric margins virtually every document
 * uses, this matches Word's behavior exactly.
 */
export function setOrientation(o: 'portrait' | 'landscape'): Command {
  return {
    name: `orientation-${o}`,
    run(state, dispatch) {
      const page = currentPage(state);
      const want = o === 'landscape';
      if (isLandscape(page) === want) return true; // already there — no undo step
      if (dispatch) {
        [page.width, page.height] = [page.height, page.width];
        dispatch(state.tr.setDocAttribute('page', page).scrollIntoView());
      }
      return true;
    },
    isActive: (state) => isLandscape(currentPage(state)) === (o === 'landscape'),
  };
}

/**
 * Set the paper size, keeping the current orientation and margins (choosing
 * "A4" while landscape yields A4 landscape — the Word model).
 */
export function setPaperSize(size: PaperSize): Command {
  const dims = PAPER_SIZES[size];
  return {
    name: `paper-${size}`,
    run(state, dispatch) {
      const page = currentPage(state);
      const landscape = isLandscape(page);
      const [w, h] = landscape
        ? [dims.height, dims.width]
        : [dims.width, dims.height];
      if (page.width === w && page.height === h) return true; // no-op
      if (dispatch) {
        page.width = w;
        page.height = h;
        dispatch(state.tr.setDocAttribute('page', page).scrollIntoView());
      }
      return true;
    },
    isActive: (state) => {
      const page = currentPage(state);
      const [pw, ph] = isLandscape(page)
        ? [page.height, page.width]
        : [page.width, page.height];
      return pw === dims.width && ph === dims.height;
    },
  };
}
