import { TextSelection } from 'prosemirror-state';
import type { EditorState } from 'prosemirror-state';
import type {
  Command,
  PageConfig,
  SectionConfig,
} from '@shadow-garden/bapbong-contracts';
import { currentSections, headBlockIndex, sectionAt } from './sections.js';

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
 * Insert a landscape page after the block at the caret: a fresh empty
 * paragraph, fenced off as its own section whose geometry is the document
 * page rotated. The manual equivalent — two next-page breaks plus a
 * per-section orientation — is exactly what this compounds, and what Word
 * does when you scope Orientation to a selection. The following content
 * resumes on the document geometry.
 */
export function insertLandscapeSection(): Command {
  return {
    name: 'insert-landscape-section',
    run(state, dispatch) {
      if (state.doc.childCount === 0) return false;
      const docPage = currentPage(state);
      if (isLandscape(docPage)) return false; // already landscape everywhere
      if (dispatch) {
        const sections = currentSections(state);
        const bi = headBlockIndex(state);
        const { i, start } = sectionAt(sections, bi);
        const S = sections[i];
        // Split S at the caret block; the landscape page sits between the
        // halves. An empty second half is dropped (caret at section end).
        const firstCount = bi + 1 - start;
        const secondCount = S.blockCount - firstCount;
        const landscape: PageConfig = {
          width: docPage.height,
          height: docPage.width,
          margin: { ...docPage.margin },
        };
        const next: SectionConfig[] = [
          ...sections.slice(0, i),
          { blockCount: firstCount, columns: { ...S.columns }, newPage: S.newPage },
          {
            blockCount: 1,
            columns: { count: 1, gap: 0 },
            newPage: true,
            page: landscape,
          },
          ...(secondCount > 0
            ? [
                {
                  blockCount: secondCount,
                  columns: { ...S.columns },
                  newPage: true,
                },
              ]
            : []),
          ...sections.slice(i + 1),
        ];
        // Insert the empty paragraph after the caret block, then rewrite the
        // section map in the same transaction so undo reverts both at once.
        const insertAt = state.selection.$head.after(1);
        const para = state.schema.nodes['paragraph'].create();
        const tr = state.tr
          .insert(insertAt, para)
          .setDocAttribute('sections', next);
        // Land the caret on the new page, ready to type.
        tr.setSelection(TextSelection.create(tr.doc, insertAt + 1));
        dispatch(tr.scrollIntoView());
      }
      return true;
    },
    isEnabled: (state) =>
      state.doc.childCount > 0 && !isLandscape(currentPage(state)),
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
