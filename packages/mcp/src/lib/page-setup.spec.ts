import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import {
  exportDocx,
  importDocx,
  schema,
} from '@shadow-garden/bapbong-headless';
import { ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';

type Page = {
  width: number;
  height: number;
  margin: Record<string, number>;
};
type Section = { blockCount: number; page?: Page; columns: { count: number } };

const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;
const pageOf = (s: HeadlessSession) => docOf(s).attrs['page'] as Page;
const sectionsOf = (s: HeadlessSession) =>
  (docOf(s).attrs['sections'] as Section[] | null) ?? [];

async function open(lines: string[]): Promise<HeadlessSession> {
  return HeadlessSession.open(
    await exportDocx(
      schema.node(
        'doc',
        null,
        lines.map((l) => schema.node('paragraph', null, [schema.text(l)])),
      ),
    ),
  );
}

describe('page_setup', () => {
  it('the whole document: orientation, paper, margins in cm, columns', async () => {
    const s = await open(['A', 'B']);
    expect(s.capabilities.pageSetup).toBe(true);
    const r = await s.pageSetup({
      orientation: 'landscape',
      paper: 'A5',
      margins: { left: 2, right: '2cm', top: '1in' },
      columns: 2,
    });
    expect(r.sections).toBe(1);
    expect(pageOf(s)).toMatchObject({
      width: 794,
      height: 559,
      margin: { left: 76, right: 76, top: 96 },
    });
    expect(sectionsOf(s)[0].columns.count).toBe(2);
  });

  it('a landscape page in a portrait document: two breaks, then that section — and back', async () => {
    const s = await open(['Intro', 'Wide table here', 'Outro']);
    await s.pageSetup({ sectionBreakAfter: { blockIndex: 0 } });
    const second = await s.pageSetup({ sectionBreakAfter: { blockIndex: 1 } });
    expect(second.sections).toBe(3);
    expect((await s.snapshot()).blocks.map((b) => b.section)).toEqual([
      1, 2, 3,
    ]);
    await s.pageSetup({ section: 2, orientation: 'landscape' });
    const secs = sectionsOf(s);
    expect(secs[1].page).toMatchObject({ width: 1123, height: 794 });
    expect(secs[0].page).toBeUndefined();
    expect(pageOf(s)).toMatchObject({ width: 794, height: 1123 });

    // It survives the file.
    let bytes: Uint8Array | undefined;
    (
      s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }
    ).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    const { doc } = await importDocx(bytes!.slice().buffer as ArrayBuffer);
    const back = doc.attrs['sections'] as Section[];
    expect(back.map((x) => x.blockCount)).toEqual([1, 1, 1]);
    expect(back[1].page).toMatchObject({ width: 1123, height: 794 });

    // Joining sections 2 and 3 leaves two.
    const joined = await s.pageSetup({ removeSectionBreak: 2 });
    expect(joined.sections).toBe(2);
  });

  it('refuses what it cannot do, in words', async () => {
    const s = await open(['A', 'B']);
    await s.insertContent([{ table: [['x']] }], { position: 'document_end' });
    await expect(
      s.pageSetup({ sectionBreakAfter: { blockIndex: 2 } }),
    ).rejects.toThrow(/inside a table/);
    await expect(s.pageSetup({ margins: { left: '10%' } })).rejects.toThrow(
      /not a share of the page/,
    );
    await s.pageSetup({ sectionBreakAfter: { blockIndex: 0 } });
    await expect(s.pageSetup({ section: 5, columns: 2 })).rejects.toThrow(
      /No section 5/,
    );
    await expect(
      s.pageSetup({
        sectionBreakAfter: { blockIndex: 1 },
        orientation: 'landscape',
      }),
    ).rejects.toThrow(ContentError);
  });
});
