import { describe, expect, it } from 'vitest';
import {
  exportDocx,
  importDocx,
  schema,
} from '@shadow-garden/bapbong-headless';
import { contentToNodes, linkTarget } from './blocks.js';
import { ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';

async function open(text: string): Promise<HeadlessSession> {
  return HeadlessSession.open(
    await exportDocx(
      schema.node('doc', null, [
        schema.node('paragraph', null, [schema.text(text)]),
      ]),
    ),
  );
}

describe('hyperlinks', () => {
  it('a link target is a web address, mailto, tel or a bookmark — never a script', () => {
    expect(linkTarget('example.com/docs')).toBe('https://example.com/docs');
    expect(linkTarget('mailto:a@b.c')).toBe('mailto:a@b.c');
    expect(linkTarget('#Terms')).toBe('#Terms');
    expect(() => linkTarget('javascript:alert(1)')).toThrow(ContentError);
    expect(() => linkTarget('file:///etc/passwd')).toThrow(/not "file:"/);
  });

  it('an inline carries its link; get_document lists the block links', async () => {
    const [p] = contentToNodes(
      [
        {
          paragraph: [
            'See ',
            { text: 'the guide', link: 'example.com/guide', bold: true },
            '.',
          ],
        },
      ],
      schema,
      { contentWidth: 602 },
    );
    const linked = p.child(1);
    expect(
      linked.marks.find((m) => m.type.name === 'link')?.attrs['href'],
    ).toBe('https://example.com/guide');
    const s = await open('Intro');
    await s.insertContent(
      [{ paragraph: [{ text: 'Mail us', link: 'mailto:hi@x.io' }] }],
      {
        position: 'document_end',
      },
    );
    expect((await s.snapshot()).blocks[1].links).toEqual([
      { text: 'Mail us', href: 'mailto:hi@x.io' },
    ]);
  });

  it('apply_formatting links text and unlinks it; a script target is refused', async () => {
    const s = await open('Read the terms first.');
    await s.applyFormatting('the terms', { link: '#Terms' });
    expect((await s.snapshot()).blocks[0].links).toEqual([
      { text: 'the terms', href: '#Terms' },
    ]);
    await expect(
      s.applyFormatting('first', { link: 'javascript:void(0)' }),
    ).rejects.toThrow(ContentError);
    await s.applyFormatting('the terms', { link: null });
    expect((await s.snapshot()).blocks[0].links).toBeUndefined();
  });

  it('round-trips through .docx', async () => {
    const s = await open('Visit our site.');
    await s.applyFormatting('our site', { link: 'https://example.com' });
    let bytes: Uint8Array | undefined;
    (
      s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }
    ).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    const back = await HeadlessSession.open(bytes!);
    expect((await back.snapshot()).blocks[0].links).toEqual([
      { text: 'our site', href: 'https://example.com' },
    ]);
    // And a plain import sees the same mark.
    const { doc } = await importDocx(bytes!.slice().buffer as ArrayBuffer);
    expect(doc.textContent).toBe('Visit our site.');
  });
});
