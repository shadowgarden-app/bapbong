import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from './docx.js';
import { exportDocx } from './export.js';

const story = (text: string): unknown =>
  schema
    .node('doc', null, [schema.node('paragraph', null, [schema.text(text)])])
    .toJSON();

const firstText = (n: PMNode | undefined) => n?.textContent ?? null;

describe('chrome overrides on a file bapbong saved before', () => {
  it('new parts do not reuse the part names or rel ids the carried package has', async () => {
    // First save: a header and a footer, written as headerB1 / footerB2.
    const first = await exportDocx(
      schema.node(
        'doc',
        {
          sectionChromeOverrides: {
            '0': {
              headers: { default: story('Acme') },
              footers: { default: story('Page') },
            },
          },
        },
        [schema.node('paragraph', null, [schema.text('Body')])],
      ),
    );
    const opened = await importDocx(first.slice().buffer as ArrayBuffer);
    // Second edit: only the footer changes.
    const edited = opened.doc.type.create(
      {
        ...opened.doc.attrs,
        sectionChromeOverrides: {
          '0': { footers: { default: story('New footer') } },
        },
      },
      opened.doc.content,
    );
    const second = await exportDocx(edited, { carry: opened.raw });
    const back = await importDocx(second.slice().buffer as ArrayBuffer);
    expect(firstText(back.headers['default'])).toBe('Acme');
    expect(firstText(back.footers['default'])).toBe('New footer');
  });
});
