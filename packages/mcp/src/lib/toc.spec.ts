import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import {
  exportDocx,
  importDocx,
  schema,
} from '@shadow-garden/bapbong-headless';
import { ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';
import { PmDocSession } from './pm-session.js';

const h = (level: number, text: string) =>
  schema.node('paragraph', { heading: level }, [schema.text(text)]);
const p = (text: string) => schema.node('paragraph', null, [schema.text(text)]);

async function open(): Promise<HeadlessSession> {
  return HeadlessSession.open(
    await exportDocx(
      schema.node('doc', null, [
        p('Cover'),
        h(1, 'Introduction'),
        p('text'),
        h(2, 'Scope'),
        h(3, 'Fine print'),
        h(1, 'Results'),
      ]),
    ),
  );
}

const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;

describe('insert_toc', () => {
  it('lists the headings to the level asked, each linked to a bookmark on its heading', async () => {
    const s = await open();
    const r = await s.insertToc(
      { levels: 2, title: 'Contents' },
      { position: 'after', text: 'Cover' },
    );
    expect(r).toMatchObject({ entries: 3, pageNumbers: 'pending' });
    const doc = docOf(s);
    expect(doc.child(1).textContent).toBe('Contents');
    const entries = [2, 3, 4].map((i) => doc.child(i));
    expect(entries.map((e) => e.textContent)).toEqual([
      'Introduction\t1',
      'Scope\t1',
      'Results\t1',
    ]);
    expect(entries.map((e) => e.attrs['styleId'])).toEqual([
      'TOC1',
      'TOC2',
      'TOC1',
    ]);
    // One field for all of them — that is what makes them one TOC.
    expect(entries[1].attrs['field']).toBe(entries[0].attrs['field']);
    expect(entries[0].attrs['field']).toMatchObject({
      kind: 'toc',
      instr: 'TOC \\o "1-2" \\h \\z \\u',
      dirty: true,
    });
    const href = entries[1].firstChild!.marks.find(
      (m) => m.type.name === 'link',
    )?.attrs['href'] as string;
    const scope = doc.child(7);
    expect(scope.textContent).toBe('Scope');
    expect(scope.attrs['bookmarks']).toEqual([href.slice(1)]);
    // Level 3 is left out.
    expect(entries.some((e) => e.textContent.startsWith('Fine print'))).toBe(
      false,
    );
  });

  it('is a real TOC in the file: one field span, links to the headings, Word asked to update it', async () => {
    const s = await open();
    await s.insertToc({}, { position: 'after', text: 'Cover' });
    let bytes: Uint8Array | undefined;
    (
      s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }
    ).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    const { doc } = await importDocx(bytes!.slice().buffer as ArrayBuffer);
    const f = doc.child(1).attrs['field'];
    expect(f).toMatchObject({ kind: 'toc' });
    expect([2, 3, 4].map((i) => doc.child(i).attrs['field'])).toEqual([
      f,
      f,
      f,
    ]);
    expect(doc.child(5).attrs['field']).toBeNull();
  });

  it('asks the host to fill the page numbers in when it can', async () => {
    let state = EditorState.create({
      doc: schema.node('doc', null, [p('Cover'), h(1, 'Only')]),
    });
    let asked: number | null = null;
    const s = new PmDocSession({
      getState: () => state,
      apply: (tr: Transaction) => {
        state = state.apply(tr);
      },
      getVersion: () => 'v',
      meta: () => ({}),
      save: async () => undefined,
      refreshToc: async (pos) => {
        asked = pos;
        return true;
      },
    });
    const r = await s.insertToc({}, { position: 'document_end' });
    expect(r.pageNumbers).toBe('updated');
    expect(state.doc.resolve(asked!).parent.attrs['field']).toMatchObject({
      kind: 'toc',
    });
  });

  it('refuses a document without headings, and a spot inside a table', async () => {
    const s = await HeadlessSession.open(
      await exportDocx(schema.node('doc', null, [p('Just text')])),
    );
    await expect(s.insertToc({}, { position: 'document_end' })).rejects.toThrow(
      /no headings/,
    );
    const t = await open();
    await t.insertContent([{ table: [['cell']] }], {
      position: 'document_end',
    });
    await expect(
      t.insertToc({}, { position: 'after', text: 'cell' }),
    ).rejects.toThrow(ContentError);
  });
});
