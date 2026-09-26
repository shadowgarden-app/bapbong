import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { AnchorError, ContentError, ReadOnlyError } from './contract.js';
import { HeadlessSession } from './headless-session.js';
import { ReadOnlySession } from './read-only-session.js';

const p = (t: string) =>
  schema.node('paragraph', null, t ? [schema.text(t)] : []);

async function open(
  blocks: PMNode[],
  attrs: Record<string, unknown> | null = null,
): Promise<HeadlessSession> {
  return HeadlessSession.open(
    await exportDocx(schema.node('doc', attrs, blocks)),
  );
}

const texts = async (s: HeadlessSession) =>
  (await s.snapshot()).blocks.map((b) => b.text);

describe('delete_block', () => {
  it('takes whole paragraphs away — one, or several in a row', async () => {
    const s = await open([p('A'), p('B'), p('C'), p('D')]);
    const r = await s.deleteBlocks(1, 1);
    expect(r.deleted).toBe(1);
    expect(await texts(s)).toEqual(['A', 'C', 'D']);
    await s.deleteBlocks(0, 2);
    expect(await texts(s)).toEqual(['D']);
  });

  it('leaves one empty paragraph when everything goes', async () => {
    const s = await open([p('A'), p('B')]);
    await s.deleteBlocks(0, 2);
    expect(await texts(s)).toEqual(['']);
  });

  it('a paragraph in a table cell goes only while the cell keeps another', async () => {
    const s = await open([p('Intro')]);
    await s.insertContent(
      [
        {
          table: [
            ['a', 'b'],
            ['c', 'd'],
          ],
        },
      ],
      { position: 'document_end' },
    );
    // Block 1 is cell (0,0) "a", its only paragraph.
    await expect(s.deleteBlocks(1, 1)).rejects.toThrow(
      /last paragraph of its table cell/,
    );
    // Give the cell a second paragraph, then one of the two can go.
    await s.insertContent('a2', { position: 'after', text: 'a' });
    expect(await texts(s)).toEqual(['Intro', 'a', 'a2', 'b', 'c', 'd']);
    await s.deleteBlocks(1, 1);
    expect(await texts(s)).toEqual(['Intro', 'a2', 'b', 'c', 'd']);
    // A range over two cells would empty the first: refused, nothing changed.
    await expect(s.deleteBlocks(1, 2)).rejects.toThrow(ContentError);
    expect(await texts(s)).toEqual(['Intro', 'a2', 'b', 'c', 'd']);
  });

  it('refuses a range past the end, and a count below one', async () => {
    const s = await open([p('A'), p('B')]);
    await expect(s.deleteBlocks(1, 2)).rejects.toThrow(AnchorError);
    await expect(s.deleteBlocks(0, 0)).rejects.toThrow(ContentError);
  });

  it('keeps the section breaks: counts follow, and a section is never emptied', async () => {
    const section = (blockCount: number) => ({
      blockCount,
      columns: { count: 1, gap: 0 },
      newPage: true,
    });
    const s = await open([p('A'), p('B'), p('C'), p('D')], {
      sections: [section(2), section(2)],
    });
    await s.deleteBlocks(0, 1);
    const doc = (s as unknown as { state: { doc: PMNode } }).state.doc;
    expect(
      (doc.attrs['sections'] as { blockCount: number }[]).map(
        (x) => x.blockCount,
      ),
    ).toEqual([1, 2]);
    await expect(s.deleteBlocks(0, 1)).rejects.toThrow(/empty section 1/);
  });

  it('is refused on a read-only session', async () => {
    const s = new ReadOnlySession(await open([p('A')]), { reason: 'no' });
    // Like every refusal of this wrapper, it throws before any promise.
    expect(() => s.deleteBlocks()).toThrow(ReadOnlyError);
  });
});
