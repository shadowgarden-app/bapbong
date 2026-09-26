import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { AnchorError, ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';

const story = (...lines: (string | PMNode)[]): PMNode =>
  schema.node(
    'doc',
    null,
    lines.map((l) =>
      typeof l === 'string'
        ? schema.node('paragraph', null, l ? [schema.text(l)] : [])
        : l,
    ),
  );
const bold = (t: string) =>
  schema.node('paragraph', null, [
    schema.text(t, [schema.marks['strong'].create()]),
  ]);

/** A document whose section(s) carry a header, as a file would. */
async function withHeader(sections = 1): Promise<HeadlessSession> {
  const section = (blockCount: number) => ({
    blockCount,
    columns: { count: 1, gap: 0 },
    newPage: true,
  });
  const doc = schema.node(
    'doc',
    {
      ...(sections > 1
        ? { sections: Array.from({ length: sections }, () => section(1)) }
        : {}),
      sectionChromeOverrides: Object.fromEntries(
        Array.from({ length: sections }, (_x, i) => [
          String(i),
          {
            headers: {
              default: story(bold('Acme Ltd'), 'Annual report').toJSON(),
            },
            footers: { default: story('Confidential').toJSON() },
          },
        ]),
      ),
    },
    Array.from({ length: sections }, (_x, i) =>
      schema.node('paragraph', null, [schema.text(`Body ${i + 1}`)]),
    ),
  );
  return HeadlessSession.open(await exportDocx(doc));
}

async function reopen(s: HeadlessSession): Promise<HeadlessSession> {
  let bytes: Uint8Array | undefined;
  (s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }).opts.onSave =
    (b) => {
      bytes = b;
    };
  await s.save();
  return HeadlessSession.open(bytes!);
}

const chromeOf = async (s: HeadlessSession) =>
  ((await s.snapshot()).chrome ?? []).map((c) => [
    c.part,
    c.sections.join(','),
    c.text,
  ]);

describe('edit_header_footer', () => {
  it('replaces text inside a header, formatting kept, and it survives the file', async () => {
    const s = await withHeader();
    expect(s.capabilities.headerFooter).toBe(true);
    const r = await s.editChrome({
      part: 'header',
      replace: { oldText: 'Acme', newText: 'Beta' },
    });
    expect(r.sections).toEqual([1]);
    const back = await reopen(s);
    expect(await chromeOf(back)).toEqual([
      ['header', '1', 'Beta Ltd\nAnnual report'],
      ['footer', '1', 'Confidential'],
    ]);
  });

  it('rewrites a footer with the page number and count, which come back as fields', async () => {
    const s = await withHeader();
    await s.editChrome({
      part: 'footer',
      content: [
        {
          paragraph: ['Page ', { field: 'page' }, ' of ', { field: 'pages' }],
          align: 'center',
        },
      ],
    });
    expect((await chromeOf(s))[1]).toEqual([
      'footer',
      '1',
      'Page {page} of {pages}',
    ]);
    const back = await reopen(s);
    expect((await chromeOf(back))[1]).toEqual([
      'footer',
      '1',
      'Page {page} of {pages}',
    ]);
  });

  it('one section, or every section; across all, a section without the text is left alone', async () => {
    const s = await withHeader(2);
    await s.editChrome({
      part: 'footer',
      section: 2,
      content: 'Appendix',
    });
    expect(await chromeOf(s)).toEqual([
      ['header', '1,2', 'Acme Ltd\nAnnual report'],
      ['footer', '1', 'Confidential'],
      ['footer', '2', 'Appendix'],
    ]);
    const all = await s.editChrome({
      part: 'footer',
      replace: { oldText: 'Confidential', newText: 'Internal' },
    });
    expect(all.sections).toEqual([1]);
    await expect(
      s.editChrome({
        part: 'footer',
        replace: { oldText: 'Nowhere', newText: 'x' },
      }),
    ).rejects.toThrow(AnchorError);
  });

  it('refuses what it cannot do, in words', async () => {
    const s = await withHeader();
    await expect(
      s.editChrome({ part: 'header', variant: 'first', content: 'Cover' }),
    ).rejects.toThrow(/no first-page header/);
    await expect(
      s.editChrome({
        part: 'header',
        content: [{ paragraph: 'a', list: 'bullet' }],
      }),
    ).rejects.toThrow(/cannot hold a list/);
    await expect(
      s.editChrome({
        part: 'header',
        content: 'x',
        replace: { oldText: 'Acme', newText: 'B' },
      }),
    ).rejects.toThrow(ContentError);
    await expect(
      s.editChrome({ part: 'header', section: 3, content: 'x' }),
    ).rejects.toThrow(/No section 3/);
  });
});
