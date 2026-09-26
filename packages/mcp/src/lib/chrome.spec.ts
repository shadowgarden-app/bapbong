import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { HeadlessSession } from './headless-session.js';
import { PmDocSession } from './pm-session.js';

const story = (...lines: string[]): PMNode =>
  schema.node(
    'doc',
    null,
    lines.map((l) => schema.node('paragraph', null, l ? [schema.text(l)] : [])),
  );

describe('get_document reports headers and footers', () => {
  it('reads what the file carries', async () => {
    const doc = schema.node(
      'doc',
      {
        sectionChromeOverrides: {
          '0': {
            headers: {
              default: story('Acme Ltd', 'Quarterly report').toJSON(),
            },
            footers: { default: story('Confidential').toJSON() },
          },
        },
      },
      [schema.node('paragraph', null, [schema.text('Body')])],
    );
    const s = await HeadlessSession.open(await exportDocx(doc));
    expect((await s.snapshot()).chrome).toEqual([
      {
        part: 'header',
        variant: 'default',
        sections: [1],
        text: 'Acme Ltd\nQuarterly report',
      },
      {
        part: 'footer',
        variant: 'default',
        sections: [1],
        text: 'Confidential',
      },
    ]);
  });

  it('lists a story shared by sections once, and leaves out empty ones', async () => {
    const shared = story('Acme Ltd');
    const state = EditorState.create({
      doc: schema.node('doc', null, [schema.node('paragraph')]),
    });
    const s = new PmDocSession({
      getState: () => state,
      apply: () => undefined,
      getVersion: () => 'v',
      meta: () => ({}),
      save: async () => undefined,
      chrome: () => [
        { headers: { default: shared }, footers: { default: story('') } },
        { headers: { default: shared, first: story('Cover') }, footers: {} },
      ],
    });
    expect((await s.snapshot()).chrome).toEqual([
      {
        part: 'header',
        variant: 'default',
        sections: [1, 2],
        text: 'Acme Ltd',
      },
      { part: 'header', variant: 'first', sections: [2], text: 'Cover' },
    ]);
  });

  it('a document without headers says nothing about them', async () => {
    const s = await HeadlessSession.open(
      await exportDocx(schema.node('doc', null, [schema.node('paragraph')])),
    );
    expect((await s.snapshot()).chrome).toBeUndefined();
  });
});
