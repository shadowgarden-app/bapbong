import type { EditorState } from 'prosemirror-state';
import type { Command, SectionConfig } from '@shadow-garden/bapbong-contracts';

/** Default column gap (px) when switching a section to multi-column — roughly
 *  Word's default `w:space` (425 twips ≈ 28px). */
const DEFAULT_GAP = 28;

/** The doc's sections, defaulting to one implicit single-column section over
 *  every top-level block. Returns a deep-ish copy (safe to mutate). */
function currentSections(state: EditorState): SectionConfig[] {
  const raw = state.doc.attrs['sections'] as SectionConfig[] | null;
  if (raw && raw.length) return raw.map((s) => ({ ...s, columns: { ...s.columns } }));
  return [{ blockCount: state.doc.childCount, columns: { count: 1, gap: 0 }, newPage: false }];
}

/** Index of the top-level block containing the selection head. */
function headBlockIndex(state: EditorState): number {
  return state.selection.$head.index(0);
}

/** The section index covering top-level block `bi`, plus its first block. */
function sectionAt(sections: SectionConfig[], bi: number): { i: number; start: number } {
  let start = 0;
  for (let i = 0; i < sections.length; i++) {
    if (bi < start + sections[i].blockCount) return { i, start };
    start += sections[i].blockCount;
  }
  const last = sections.length - 1;
  return { i: last, start: start - sections[last].blockCount };
}

/**
 * Insert a section break after the block at the caret: the section the caret
 * sits in is split in two at that boundary. **Columns are unchanged** — the new
 * section inherits the current column layout — so this only sets where the
 * following content begins (`newPage` → a fresh page, else continuous). Change
 * the column count separately with {@link setColumns} (the Word model: columns
 * are a per-section property, set via Format → Columns).
 *
 * Sections are stored positionally (`doc.attrs.sections`, by `blockCount`), so
 * like imported sections they can drift if blocks are later added/removed
 * elsewhere — a known model limitation, not introduced here.
 */
export function insertSectionBreak(opts: { newPage: boolean }): Command {
  return {
    name: opts.newPage ? 'section-break-next-page' : 'section-break-continuous',
    run(state, dispatch) {
      if (state.doc.childCount < 2) return false;
      const sections = currentSections(state);
      const bi = headBlockIndex(state);
      const { i, start } = sectionAt(sections, bi);
      const S = sections[i];
      const firstCount = bi + 1 - start; // [start, bi] stays with the section
      const secondCount = S.blockCount - firstCount; // [bi+1, sectionEnd) splits off
      if (secondCount <= 0) return false; // caret already at the section's last block
      if (dispatch) {
        const next: SectionConfig[] = [
          ...sections.slice(0, i),
          { blockCount: firstCount, columns: { ...S.columns }, newPage: S.newPage },
          { blockCount: secondCount, columns: { ...S.columns }, newPage: opts.newPage },
          ...sections.slice(i + 1),
        ];
        dispatch(state.tr.setDocAttribute('sections', next).scrollIntoView());
      }
      return true;
    },
    isEnabled: (state) => state.doc.childCount > 1,
  };
}

/**
 * Remove the section break at `boundaryIndex` (0-based; there are
 * `sections.length - 1` breaks, break `b` sits between section `b` and `b+1`).
 * The two sections merge into one. Word semantics: the upper content adopts the
 * **following** section's column layout (`columns` taken from section `b+1`),
 * while the merged section keeps section `b`'s page-start.
 */
export function removeSectionBreak(boundaryIndex: number): Command {
  return {
    name: 'remove-section-break',
    run(state, dispatch) {
      const raw = state.doc.attrs['sections'] as SectionConfig[] | null;
      const b = boundaryIndex;
      if (!raw || b < 0 || b >= raw.length - 1) return false;
      if (dispatch) {
        const upper = raw[b];
        const lower = raw[b + 1];
        const merged: SectionConfig = {
          blockCount: upper.blockCount + lower.blockCount,
          columns: { ...lower.columns }, // following section's layout wins (Word)
          newPage: upper.newPage, // merged range still begins where the upper did
        };
        const next = [...raw.slice(0, b), merged, ...raw.slice(b + 2)];
        dispatch(state.tr.setDocAttribute('sections', next).scrollIntoView());
      }
      return true;
    },
  };
}

/**
 * Set the column count of the section the caret sits in (Format → Columns). A
 * count > 1 needs a section to flow into, so a single-section doc becomes
 * multi-column as a whole; scope columns to a region by inserting section
 * breaks around it first ({@link insertSectionBreak}).
 */
export function setColumns(count: number): Command {
  return {
    name: `columns-${count}`,
    run(state, dispatch) {
      if (state.doc.childCount === 0) return false;
      const sections = currentSections(state);
      const { i } = sectionAt(sections, headBlockIndex(state));
      if (dispatch) {
        const gap = count > 1 ? sections[i].columns.gap || DEFAULT_GAP : sections[i].columns.gap;
        const next = sections.map((s, j) => (j === i ? { ...s, columns: { count, gap } } : s));
        dispatch(state.tr.setDocAttribute('sections', next).scrollIntoView());
      }
      return true;
    },
    isActive: (state) => {
      const sections = currentSections(state);
      return sections[sectionAt(sections, headBlockIndex(state)).i].columns.count === count;
    },
  };
}
