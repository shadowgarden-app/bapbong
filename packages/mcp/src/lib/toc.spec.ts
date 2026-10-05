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
        return 0;
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

describe('update_toc', () => {
  const pmSession = (
    doc: PMNode,
    refresh?: (pos: number, state: EditorState) => number | null,
  ) => {
    let state = EditorState.create({ doc });
    const asked: number[] = [];
    const s = new PmDocSession({
      getState: () => state,
      apply: (tr: Transaction) => {
        state = state.apply(tr);
      },
      getVersion: () => 'v',
      meta: () => ({}),
      save: async () => undefined,
      ...(refresh
        ? {
            refreshToc: async (pos: number) => {
              asked.push(pos);
              return refresh(pos, state);
            },
          }
        : {}),
    });
    /** Rename the heading `from` (not its TOC entry) to `to`. */
    const rename = (from: string, to: string) => {
      state.doc.forEach((n, offset) => {
        if (n.attrs['heading'] && n.textContent === from) {
          state = state.apply(
            state.tr.insertText(to, offset + 1, offset + 1 + from.length),
          );
        }
      });
    };
    return { s, asked, doc: () => state.doc, rename };
  };
  const entryTexts = (doc: PMNode) => {
    const out: string[] = [];
    doc.forEach((n) => {
      if (n.attrs['field']) out.push(n.textContent);
    });
    return out;
  };

  it('says there is nothing to update when the document has no TOC', async () => {
    const s = await open();
    await expect(s.updateToc({})).rejects.toThrow(/insert_toc/);
  });

  it('refreshes page numbers through the host, and stays "pending" without one', async () => {
    const s = await open();
    await s.insertToc({}, { position: 'after', text: 'Cover' });
    expect(await s.updateToc({})).toMatchObject({
      tables: 1,
      entries: 4,
      changed: 0,
      rebuilt: false,
      pageNumbers: 'pending',
    });

    const { s: pm, asked } = pmSession(
      schema.node('doc', null, [p('Cover'), h(1, 'A'), h(1, 'B')]),
      () => 2,
    );
    await pm.insertToc({}, { position: 'after', text: 'Cover' });
    asked.length = 0;
    const r = await pm.updateToc({});
    expect(r).toMatchObject({ tables: 1, changed: 2, pageNumbers: 'updated' });
    expect(asked).toHaveLength(1);
  });

  it('rebuild lists the headings as they are now, keeping the levels', async () => {
    const { s, doc, rename } = pmSession(
      schema.node('doc', null, [
        p('Cover'),
        h(1, 'Introduction'),
        h(2, 'Scope'),
        h(3, 'Fine print'),
        h(1, 'Results'),
      ]),
    );
    await s.insertToc({ levels: 2 }, { position: 'after', text: 'Cover' });
    expect(entryTexts(doc())).toEqual([
      'Introduction\t1',
      'Scope\t1',
      'Results\t1',
    ]);
    // Rename two headings, add a new one at the end.
    rename('Results', 'Findings');
    rename('Scope', 'Aims');
    await s.insertContent([{ paragraph: 'Limits', heading: 2 }], {
      position: 'document_end',
    });
    const before = doc().childCount;
    // Page numbers only: entries untouched.
    await s.updateToc({});
    expect(entryTexts(doc())).toEqual([
      'Introduction\t1',
      'Scope\t1',
      'Results\t1',
    ]);
    const r = await s.updateToc({ rebuild: true });
    expect(r).toMatchObject({ tables: 1, entries: 4, rebuilt: true });
    expect(entryTexts(doc())).toEqual([
      'Introduction\t1',
      'Aims\t1',
      'Findings\t1',
      'Limits\t1',
    ]);
    expect(doc().childCount).toBe(before + 1);
    // Still one field, still levels 1-2, every entry linked to a bookmark
    // that a heading carries.
    const fields = new Set<unknown>();
    doc().forEach((n) => {
      if (n.attrs['field']) fields.add(n.attrs['field']);
    });
    expect(fields.size).toBe(1);
    expect([...fields][0]).toMatchObject({
      instr: 'TOC \\o "1-2" \\h \\z \\u',
    });
    const names = new Set<string>();
    doc().forEach((n) => {
      for (const b of (n.attrs['bookmarks'] as string[] | null) ?? [])
        names.add(b);
    });
    doc().forEach((n) => {
      if (!n.attrs['field']) return;
      const href = n.firstChild!.marks.find((m) => m.type.name === 'link')
        ?.attrs['href'] as string;
      expect(names.has(href.slice(1))).toBe(true);
    });
  });

  it('rebuilds every table, each with its own levels', async () => {
    const { s, doc } = pmSession(
      schema.node('doc', null, [
        p('Front'),
        p('Middle'),
        h(1, 'One'),
        h(2, 'One.a'),
        h(1, 'Two'),
      ]),
    );
    await s.insertToc({ levels: 1 }, { position: 'after', text: 'Front' });
    await s.insertToc({ levels: 2 }, { position: 'after', text: 'Middle' });
    await s.insertContent([{ paragraph: 'Three', heading: 1 }], {
      position: 'document_end',
    });
    const r = await s.updateToc({ rebuild: true });
    expect(r).toMatchObject({ tables: 2, entries: 3 + 4 });
    expect(entryTexts(doc())).toEqual([
      'One\t1',
      'Two\t1',
      'Three\t1',
      'One\t1',
      'One.a\t1',
      'Two\t1',
      'Three\t1',
    ]);
  });

  it('is refused on a read-only document', async () => {
    const s = await open();
    await s.insertToc({}, { position: 'after', text: 'Cover' });
    const { ReadOnlySession } = await import('./read-only-session.js');
    const ro = new ReadOnlySession(s);
    await expect(
      Promise.resolve().then(() => ro.updateToc()),
    ).rejects.toThrow();
  });
});
