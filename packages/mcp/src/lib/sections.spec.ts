import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';
import { keepSections } from './pm-session.js';

const section = (blockCount: number, newPage: boolean) => ({
  blockCount,
  columns: { count: 1, gap: 0 },
  newPage,
});

/** A B | C D — two sections of two paragraphs, the second on a new page. */
function twoSections(): PMNode {
  return schema.node(
    'doc',
    { sections: [section(2, false), section(2, true)] },
    ['A', 'B', 'C', 'D'].map((t) =>
      schema.node('paragraph', null, [schema.text(t)]),
    ),
  );
}

const counts = (doc: PMNode) =>
  (doc.attrs['sections'] as { blockCount: number }[]).map((s) => s.blockCount);

const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;

describe('an agent edit keeps section breaks where they were', () => {
  it('the document the tests start from has its two sections', async () => {
    const s = await HeadlessSession.open(await exportDocx(twoSections()));
    expect(counts(docOf(s))).toEqual([2, 2]);
  });

  it('content lands in the section of its anchor: after the last block above the break, before the first below it', async () => {
    const s = await HeadlessSession.open(await exportDocx(twoSections()));
    await s.insertContent('x\ny', { position: 'after', text: 'B' });
    expect(counts(docOf(s))).toEqual([4, 2]);
    await s.insertContent('z', { position: 'before', text: 'C' });
    expect(counts(docOf(s))).toEqual([4, 3]);
    await s.insertContent('end', { position: 'document_end' });
    expect(counts(docOf(s))).toEqual([4, 4]);
    // The break is still between B and whatever now opens section 2.
    const texts: string[] = [];
    docOf(s).forEach((n) => texts.push(n.textContent));
    expect(texts.slice(0, 4)).toEqual(['A', 'B', 'x', 'y']);
    expect(texts[4]).toBe('z');
  });

  it('an edit inside a paragraph adds no top-level block and changes nothing', () => {
    const doc = twoSections();
    const state = EditorState.create({ doc });
    const tr = state.tr.insertText('!', 2);
    expect(keepSections(doc, tr, 1)).toBe(tr);
  });

  it('removing a block moves the counts; emptying a section is refused', () => {
    const doc = twoSections();
    const state = EditorState.create({ doc });
    const a = doc.child(0);
    const tr = keepSections(doc, state.tr.delete(0, a.nodeSize), 1);
    expect(counts(tr.doc)).toEqual([1, 2]);
    const both = doc.child(0).nodeSize + doc.child(1).nodeSize;
    expect(() => keepSections(doc, state.tr.delete(0, both), 1)).toThrow(
      ContentError,
    );
  });
});
