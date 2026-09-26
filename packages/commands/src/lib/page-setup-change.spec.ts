import { describe, expect, it } from 'vitest';
import { Schema, type Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import type {
  PageConfig,
  SectionConfig,
} from '@shadow-garden/bapbong-contracts';
import { pageSetupTransaction } from './page-setup-change.js';

const schema = new Schema({
  nodes: {
    doc: {
      content: 'block+',
      attrs: { page: { default: null }, sections: { default: null } },
    },
    paragraph: { group: 'block', content: 'text*' },
    text: {},
  },
});

const A4: PageConfig = {
  width: 794,
  height: 1123,
  margin: { top: 96, right: 96, bottom: 96, left: 96 },
};
const section = (blockCount: number, page?: PageConfig): SectionConfig => ({
  blockCount,
  columns: { count: 1, gap: 0 },
  newPage: true,
  ...(page ? { page } : {}),
});

function stateOf(
  n: number,
  attrs: { page?: PageConfig; sections?: SectionConfig[] } = {},
): EditorState {
  const blocks: PMNode[] = Array.from({ length: n }, (_x, i) =>
    schema.node('paragraph', null, [schema.text(`p${i}`)]),
  );
  return EditorState.create({
    doc: schema.node('doc', { page: A4, ...attrs }, blocks),
  });
}

const run = (
  state: EditorState,
  change: Parameters<typeof pageSetupTransaction>[1],
) => {
  const tr = pageSetupTransaction(state, change);
  return tr ? state.apply(tr) : state;
};
const pageOf = (s: EditorState) => s.doc.attrs['page'] as PageConfig;
const sectionsOf = (s: EditorState) =>
  s.doc.attrs['sections'] as SectionConfig[];

describe('pageSetupTransaction', () => {
  it('a one-section document: orientation, paper, margins and columns in one transaction', () => {
    const state = stateOf(3);
    const tr = pageSetupTransaction(state, {
      orientation: 'landscape',
      paper: 'letter',
      margins: { left: 48, right: 48 },
      columns: 2,
    });
    expect(tr).not.toBeNull();
    const next = state.apply(tr!);
    expect(pageOf(next)).toEqual({
      width: 1056,
      height: 816,
      margin: { top: 96, right: 48, bottom: 96, left: 48 },
    });
    expect(sectionsOf(next)[0].columns.count).toBe(2);
    // Naming the one section is the same as naming none.
    expect(
      pageOf(run(state, { section: 0, orientation: 'landscape' })),
    ).toMatchObject({ width: 1123 });
    // Already in effect: nothing to do.
    expect(pageSetupTransaction(next, { orientation: 'landscape' })).toBeNull();
  });

  it('one section of several: only it changes; every section when none is named', () => {
    const state = stateOf(4, { sections: [section(2), section(2)] });
    const first = run(state, { section: 0, orientation: 'landscape' });
    expect(sectionsOf(first)[0].page).toMatchObject({
      width: 1123,
      height: 794,
    });
    expect(pageOf(first)).toMatchObject({ width: 794 }); // the last section's
    const all = run(state, { margins: 'narrow' });
    expect(sectionsOf(all)[0].page?.margin).toEqual({
      top: 48,
      right: 48,
      bottom: 48,
      left: 48,
    });
    expect(pageOf(all).margin).toEqual({
      top: 48,
      right: 48,
      bottom: 48,
      left: 48,
    });
  });

  it('inserts and removes a section break after a block', () => {
    const state = stateOf(4);
    const split = run(state, { sectionBreak: { after: 1, newPage: true } });
    expect(sectionsOf(split).map((s) => s.blockCount)).toEqual([2, 2]);
    const joined = run(split, { removeSectionBreak: 0 });
    expect(sectionsOf(joined).map((s) => s.blockCount)).toEqual([4]);
  });

  it('says what it cannot do, and does nothing then', () => {
    const state = stateOf(2, { sections: [section(1), section(1)] });
    expect(() =>
      pageSetupTransaction(state, { section: 5, columns: 2 }),
    ).toThrow(/No section 6/);
    expect(() =>
      pageSetupTransaction(state, {
        sectionBreak: { after: 1, newPage: true },
      }),
    ).toThrow(/Cannot start a section after block 1/);
    expect(() =>
      pageSetupTransaction(state, { removeSectionBreak: 3 }),
    ).toThrow(/No section break 4/);
    expect(() =>
      pageSetupTransaction(state, {
        sectionBreak: { after: 0, newPage: true },
        columns: 2,
      }),
    ).toThrow(/on its own/);
  });
});
