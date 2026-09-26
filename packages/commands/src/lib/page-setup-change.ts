import type { EditorState, Transaction } from 'prosemirror-state';
import type {
  Command,
  PageConfig,
  SectionConfig,
} from '@shadow-garden/bapbong-contracts';
import {
  MARGIN_PRESETS,
  currentPageConfig,
  definedSides,
  setMargins,
  setOrientation,
  setPageMargins,
  setPaperSize,
  setSectionOrientation,
  setSectionPageMargins,
  setSectionPaperSize,
  type MarginPreset,
  type PaperSize,
} from './page-setup.js';
import {
  insertSectionBreakAfter,
  removeSectionBreak,
  setSectionColumns,
} from './sections.js';

/**
 * A page-setup change addressed by index rather than by the caret — what a
 * caller that works on blocks (an agent) asks for. The Layout menu's
 * commands, composed: see {@link pageSetupTransaction}.
 */
export interface PageSetupChange {
  /** 0-based section. Absent: every section of the document. */
  section?: number;
  orientation?: 'portrait' | 'landscape';
  paper?: PaperSize;
  /** A Word preset, or px per side (a side left out keeps its value). */
  margins?: MarginPreset | Partial<PageConfig['margin']>;
  columns?: number;
  /** Start a new section after top-level block `after` (0-based). */
  sectionBreak?: { after: number; newPage: boolean };
  /** Remove section break `n` (0-based: between section n and n + 1). */
  removeSectionBreak?: number;
}

/**
 * One transaction for `change`, or null when all of it is already in
 * effect. The steps of each command are replayed onto one transaction, so
 * the whole change is a single undo step. Throws an Error whose message says
 * what cannot be done (no such section, nothing after that block to split
 * off) — nothing is applied then.
 *
 * A break is a change of its own: section numbers shift with it, so a call
 * that moves a break cannot also set up a section.
 */
export function pageSetupTransaction(
  state: EditorState,
  change: PageSetupChange,
): Transaction | null {
  const geometry =
    change.orientation !== undefined ||
    change.paper !== undefined ||
    change.margins !== undefined ||
    change.columns !== undefined;
  const breaks =
    change.sectionBreak !== undefined ||
    change.removeSectionBreak !== undefined;
  if (breaks && (geometry || change.section !== undefined)) {
    throw new Error(
      'Insert or remove a section break on its own — the sections are numbered anew after it — then set the new section up in a second call.',
    );
  }
  if (change.sectionBreak && change.removeSectionBreak !== undefined) {
    throw new Error('Insert a section break or remove one, not both at once.');
  }

  const sections =
    (state.doc.attrs['sections'] as SectionConfig[] | null) ?? null;
  const count = sections?.length ?? 1;
  if (
    change.section !== undefined &&
    (change.section < 0 || change.section >= count)
  ) {
    throw new Error(
      `No section ${change.section + 1} — the document has ${count} section(s).`,
    );
  }

  const commands: [Command, string][] = [];
  if (change.sectionBreak) {
    const { after, newPage } = change.sectionBreak;
    commands.push([
      insertSectionBreakAfter(after, { newPage }),
      `Cannot start a section after block ${after}: it is the last block of the document, or the last of its section already.`,
    ]);
  }
  if (change.removeSectionBreak !== undefined) {
    const b = change.removeSectionBreak;
    commands.push([
      removeSectionBreak(b),
      `No section break ${b + 1} — the document has ${count - 1}.`,
    ]);
  }

  // Geometry: per section when the document has sections (the last one's
  // geometry is the document's page, the rest carry overrides), else the
  // document's page itself.
  const targets =
    !sections || sections.length < 2
      ? null // one section: its geometry IS the document's page
      : change.section !== undefined
        ? [change.section]
        : sections.map((_s, i) => i);
  const noSection = `That section cannot be changed.`;
  if (change.orientation) {
    const o = change.orientation;
    if (targets)
      for (const i of targets)
        commands.push([setSectionOrientation(i, o), noSection]);
    else commands.push([setOrientation(o), noSection]);
  }
  if (change.paper) {
    const p = change.paper;
    if (targets)
      for (const i of targets)
        commands.push([setSectionPaperSize(i, p), noSection]);
    else commands.push([setPaperSize(p), noSection]);
  }
  if (change.margins !== undefined) {
    const m = change.margins;
    if (typeof m === 'string') {
      if (targets)
        for (const i of targets)
          commands.push([
            setSectionPageMargins(i, MARGIN_PRESETS[m]),
            noSection,
          ]);
      else commands.push([setMargins(m), noSection]);
    } else if (targets) {
      for (const i of targets)
        commands.push([setSectionPageMargins(i, m), noSection]);
    } else {
      commands.push([
        setPageMargins({
          ...currentPageConfig(state).margin,
          ...definedSides(m),
        }),
        noSection,
      ]);
    }
  }
  if (change.columns !== undefined) {
    const n = change.columns;
    for (const i of targets ?? [0])
      commands.push([setSectionColumns(i, n), noSection]);
  }

  // Run them one after another, each on the state the last one left, and
  // collect their steps onto one transaction.
  const tr = state.tr;
  let current = state;
  for (const [command, refusal] of commands) {
    let produced: Transaction | null = null;
    const ok = command.run(current, (t) => {
      produced = t;
    });
    if (!ok) throw new Error(refusal);
    const t = produced as Transaction | null;
    if (!t) continue; // already in effect
    for (const step of t.steps) tr.step(step);
    current = current.apply(t);
  }
  return tr.docChanged || tr.steps.length > 0 ? tr : null;
}
