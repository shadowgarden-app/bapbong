import { describe, expect, it } from 'vitest';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { AnchorError, ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';

async function open(): Promise<HeadlessSession> {
  return HeadlessSession.open(
    await exportDocx(
      schema.node('doc', null, [
        schema.node('paragraph', null, [
          schema.text('First claim. Second claim.'),
        ]),
      ]),
    ),
  );
}

const refsOf = async (s: HeadlessSession) =>
  (await s.snapshot()).blocks[0].text;

describe('insert_footnote', () => {
  it('numbers follow the document: a note put in front moves the others up', async () => {
    const s = await open();
    expect(
      (await s.insertFootnote({ text: 'Second claim.' }, 'Source B.')).number,
    ).toBe('1');
    expect(
      (await s.insertFootnote({ text: 'First claim.' }, 'Source A.')).number,
    ).toBe('1');
    expect(await refsOf(s)).toBe('First claim.1 Second claim.2');
    expect((await s.snapshot()).footnotes).toEqual([
      { number: '1', text: 'Source A.' },
      { number: '2', text: 'Source B.' },
    ]);
  });

  it('survives the file: Word gets real footnotes, bapbong reads them back', async () => {
    const s = await open();
    await s.insertFootnote({ text: 'First claim.' }, [
      { paragraph: ['See ', { text: 'Annex 2', italic: true }, '.'] },
    ]);
    let bytes: Uint8Array | undefined;
    (
      s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }
    ).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    const back = await HeadlessSession.open(bytes!);
    expect((await back.snapshot()).footnotes).toEqual([
      { number: '1', text: 'See Annex 2.' },
    ]);
    // And a second one on the reopened file sits beside it.
    await back.insertFootnote({ text: 'Second claim.' }, 'Another.');
    expect((await back.snapshot()).footnotes?.map((f) => f.number)).toEqual([
      '1',
      '2',
    ]);
  });

  it('refuses what a note cannot hold, and an anchor that is not there', async () => {
    const s = await open();
    await expect(
      s.insertFootnote({ text: 'First claim.' }, [{ table: [['x']] }]),
    ).rejects.toThrow(ContentError);
    await expect(
      s.insertFootnote({ text: 'First claim.' }, [
        { paragraph: 'x', list: 'bullet' },
      ]),
    ).rejects.toThrow(/no tables or lists/);
    await expect(
      s.insertFootnote({ text: 'First claim.' }, [
        { paragraph: [{ text: 'site', link: 'https://example.com' }] },
      ]),
    ).rejects.toThrow(/link/);
    await expect(s.insertFootnote({ text: 'Nowhere' }, 'x')).rejects.toThrow(
      AnchorError,
    );
  });
});
