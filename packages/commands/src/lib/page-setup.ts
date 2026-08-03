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

/** Common paper sizes in CSS px @96dpi, portrait. `width`/`height` are the
 *  model values (derived from the canonical twips ÷15, so import → command →
 *  export round-trips); `cm` is the exact nominal size for display — the px
 *  rounding would otherwise show A4 as 21.006 cm instead of 21. */
export const PAPER_SIZES = {
  letter: { width: 816, height: 1056, cm: [21.59, 27.94], label: 'Letter' },
  legal: { width: 816, height: 1344, cm: [21.59, 35.56], label: 'Legal' },
  executive: {
    width: 696,
    height: 1008,
    cm: [18.415, 26.67],
    label: 'Executive',
  },
  a4: { width: 794, height: 1123, cm: [21, 29.7], label: 'A4' },
  a5: { width: 559, height: 794, cm: [14.8, 21], label: 'A5' },
  a3: { width: 1123, height: 1587, cm: [29.7, 42], label: 'A3' },
} as const;
export type PaperSize = keyof typeof PAPER_SIZES;

/** Word's margin presets, in CSS px @96dpi (1 cm = 96/2.54 px, all exact). */
export const MARGIN_PRESETS = {
  normal: { top: 96, right: 96, bottom: 96, left: 96 }, // 2.54 cm
  narrow: { top: 48, right: 48, bottom: 48, left: 48 }, // 1.27 cm
  moderate: { top: 96, right: 72, bottom: 96, left: 72 }, // 2.54 / 1.905
  wide: { top: 96, right: 192, bottom: 96, left: 192 }, // 2.54 / 5.08
  office2003: { top: 96, right: 120, bottom: 96, left: 120 }, // 2.54 / 3.175
} as const;
export type MarginPreset = keyof typeof MARGIN_PRESETS;

/** A4 @96dpi with 1in margins — what layout assumes when a doc predates the
 *  page attr (same fallback as the view and the exporter). */
const A4_PAGE: PageConfig = {
  width: 794,
  height: 1123,
  margin: { top: 96, right: 96, bottom: 96, left: 96 },
};

/** The doc's page geometry (a copy, safe to mutate). Exported so a host can
 *  seed its page-setup UI with the values currently in effect. */
export function currentPageConfig(state: EditorState): PageConfig {
  const raw = state.doc.attrs['page'] as PageConfig | null;
  const p = raw ?? A4_PAGE;
  return { width: p.width, height: p.height, margin: { ...p.margin } };
}
const currentPage = currentPageConfig;

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
    isActive: (state) =>
      isLandscape(currentPage(state)) === (o === 'landscape'),
  };
}

/** Smallest page/margin the layout engine can lay a line into — anything less
 *  yields a degenerate content box (which `sanitizePage` would then override
 *  wholesale). Clamped here so a typed value can't silently become something
 *  else downstream. */
const MIN_PAGE = 96; // 1in — well under any real paper
const MIN_CONTENT = 48; // margins must leave at least this much across

/** Apply a partial page-geometry change (merged over the current one). */
function applyPage(
  name: string,
  patch: (page: PageConfig) => PageConfig | null,
  isActive?: (page: PageConfig) => boolean,
): Command {
  return {
    name,
    run(state, dispatch) {
      const next = patch(currentPage(state));
      if (!next) return true; // already in effect — no undo step
      if (dispatch)
        dispatch(state.tr.setDocAttribute('page', next).scrollIntoView());
      return true;
    },
    ...(isActive ? { isActive: (s) => isActive(currentPage(s)) } : {}),
  };
}

const sameMargin = (a: PageConfig['margin'], b: PageConfig['margin']) =>
  a.top === b.top &&
  a.right === b.right &&
  a.bottom === b.bottom &&
  a.left === b.left;

/**
 * Apply a margin preset (Layout ▸ Margins). Page size is untouched.
 */
export function setMargins(preset: MarginPreset): Command {
  const m = MARGIN_PRESETS[preset];
  return applyPage(
    `margins-${preset}`,
    (page) =>
      sameMargin(page.margin, m) ? null : { ...page, margin: { ...m } },
    (page) => sameMargin(page.margin, m),
  );
}

/**
 * Set explicit margins in px (Layout ▸ Margins ▸ Custom margins). Values are
 * clamped so the content box stays layout-able on both axes.
 */
export function setPageMargins(margin: PageConfig['margin']): Command {
  return applyPage('page-margins-custom', (page) => {
    const fit = (a: number, b: number, extent: number): [number, number] => {
      const lo = Math.max(0, a);
      const hi = Math.max(0, b);
      if (extent - lo - hi >= MIN_CONTENT) return [lo, hi];
      // Shrink both sides proportionally rather than rejecting the input.
      const room = Math.max(0, extent - MIN_CONTENT);
      const total = lo + hi || 1;
      return [(lo / total) * room, (hi / total) * room];
    };
    const [left, right] = fit(margin.left, margin.right, page.width);
    const [top, bottom] = fit(margin.top, margin.bottom, page.height);
    const next = {
      top: Math.round(top),
      right: Math.round(right),
      bottom: Math.round(bottom),
      left: Math.round(left),
    };
    return sameMargin(page.margin, next) ? null : { ...page, margin: next };
  });
}

/**
 * Set an explicit page size in px (Layout ▸ Page size ▸ Custom page size),
 * keeping the current margins. Unlike {@link setPaperSize} the dimensions are
 * taken as given — the caller states orientation by which one is larger.
 */
export function setPageDimensions(width: number, height: number): Command {
  return applyPage('page-size-custom', (page) => {
    const w = Math.max(MIN_PAGE, Math.round(width));
    const h = Math.max(MIN_PAGE, Math.round(height));
    if (page.width === w && page.height === h) return null;
    return { ...page, width: w, height: h };
  });
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
          {
            blockCount: firstCount,
            columns: { ...S.columns },
            newPage: S.newPage,
          },
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
