import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import {
  createNumberingCounter,
  exportDocx,
  importDocx,
  listPresets,
  schema,
} from '@shadow-garden/bapbong-headless';
import { buildContent, LIST_DEFS } from './blocks.js';
import { ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';

const WIDTH = 602;

/** The session's CURRENT document (its state is replaced on every apply). */
const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;

/** The markers the layout would draw, one per paragraph ('' = no list). */
function labels(doc: PMNode): string[] {
  const counter = createNumberingCounter(doc.attrs['numbering']);
  const out: string[] = [];
  doc.descendants((n) => {
    if (n.type.name !== 'paragraph') return true;
    const list = n.attrs['list'] as { numId: string; level?: number } | null;
    out.push(list ? counter.next(list.numId, list.level ?? 0) : '');
    return false;
  });
  return out;
}

async function open(lines: string[]): Promise<HeadlessSession> {
  const doc = schema.node(
    'doc',
    null,
    lines.map((l) => schema.node('paragraph', null, [schema.text(l)])),
  );
  return HeadlessSession.open(await exportDocx(doc));
}

describe('lists in blocks', () => {
  it("a new list is born with the editor's own default markers", () => {
    const [bullet] = listPresets('bullet');
    const [ordered] = listPresets('ordered');
    expect(LIST_DEFS.bullet.numId).toBe(bullet.numId);
    expect(LIST_DEFS.number.numId).toBe(ordered.numId);
    expect(Object.values(LIST_DEFS.bullet.levels)).toEqual(bullet.levels);
    expect(Object.values(LIST_DEFS.number.levels)).toEqual(ordered.levels);
  });

  it('consecutive items are one list; anything between two numbered lists starts the second at 1', () => {
    const { nodes, numbering } = buildContent(
      [
        { paragraph: 'one', list: 'number' },
        { paragraph: 'two', list: 'number' },
        { paragraph: 'two a', list: 'number', level: 2 },
        { paragraph: 'Next', heading: 2 },
        { paragraph: 'again one', list: 'number' },
        { paragraph: 'dot', list: 'bullet' },
      ],
      schema,
      { contentWidth: WIDTH },
    );
    const doc = schema.node('doc', { numbering }, nodes);
    expect(labels(doc)).toEqual(['1.', '2.', 'a.', '', '1.', '•']);
    // The nested item is indented the way Tab indents it in the editor.
    expect(nodes[2].attrs['indent']).toEqual({ left: 24 });
    expect(nodes[0].attrs['indent']).toBeNull();
    expect(Object.keys(numbering ?? {}).sort()).toEqual([
      'bb-bullet',
      'bb-ordered',
      'bb-ordered-n2',
    ]);
  });

  it('keeps the ids the document already uses, and says when nothing was added', () => {
    const existing = { 'bb-ordered': { key: 'bb-ordered', levels: {} } };
    const { numbering } = buildContent(
      [{ paragraph: 'x', list: 'number' }],
      schema,
      { contentWidth: WIDTH, numbering: existing },
    );
    expect(Object.keys(numbering ?? {})).toEqual([
      'bb-ordered',
      'bb-ordered-n2',
    ]);
    expect(
      buildContent(['plain'], schema, { contentWidth: WIDTH }).numbering,
    ).toBeNull();
  });

  it('refuses a level without a list, and a level deeper than three', () => {
    expect(() =>
      buildContent([{ paragraph: 'x', level: 2 }], schema, {
        contentWidth: WIDTH,
      }),
    ).toThrow(/level is for list items/);
    expect(() =>
      buildContent([{ paragraph: 'x', list: 'bullet', level: 4 }], schema, {
        contentWidth: WIDTH,
      }),
    ).toThrow(ContentError);
  });
});

describe('lists through a headless session', () => {
  it('inserting right after a list item continues that list', async () => {
    const s = await open(['Steps', 'After']);
    await s.insertContent(
      [
        { paragraph: 'first', list: 'number' },
        { paragraph: 'second', list: 'number' },
      ],
      { position: 'after', text: 'Steps' },
    );
    await s.insertContent([{ paragraph: 'third', list: 'number' }], {
      position: 'after',
      text: 'second',
    });
    expect(labels(docOf(s))).toEqual(['', '1.', '2.', '3.', '']);
    const snap = await s.snapshot();
    expect(snap.blocks[1].list).toEqual({ kind: 'number', level: 1 });
    expect(snap.blocks[0].list).toBeUndefined();
  });

  it('apply_formatting turns paragraphs into one list top to bottom, nests one, and ends it', async () => {
    const s = await open(['Intro', 'a', 'b', 'c']);
    for (const i of [1, 2, 3])
      await s.applyFormatting({ blockIndex: i }, { list: 'number' });
    expect(labels(docOf(s))).toEqual(['', '1.', '2.', '3.']);

    await s.applyFormatting({ blockIndex: 2 }, { listLevel: 2 });
    expect(labels(docOf(s))).toEqual(['', '1.', 'a.', '2.']);
    expect(docOf(s).child(2).attrs['indent']).toEqual({ left: 24 });
    expect((await s.snapshot()).blocks[2].list).toEqual({
      kind: 'number',
      level: 2,
    });

    // Switching kind leaves the list it was in; body text drops the indent.
    await s.applyFormatting('c', { list: 'bullet' });
    expect(labels(docOf(s))).toEqual(['', '1.', 'a.', '•']);
    await s.applyFormatting({ blockIndex: 2 }, { list: null });
    expect(docOf(s).child(2).attrs['list']).toBeNull();
    expect(docOf(s).child(2).attrs['indent']).toBeNull();
  });

  it('refuses a level for a paragraph that is not a list item, or deeper than its list goes', async () => {
    const s = await open(['plain', 'item']);
    await expect(
      s.applyFormatting({ blockIndex: 0 }, { listLevel: 2 }),
    ).rejects.toThrow(/not a list item/);
    await s.applyFormatting({ blockIndex: 1 }, { list: 'bullet' });
    await expect(
      s.applyFormatting({ blockIndex: 1 }, { listLevel: 4 }),
    ).rejects.toThrow(/1 to 3 levels/);
  });

  it('round-trips through .docx: the lists come back with the same markers', async () => {
    const s = await open(['Intro']);
    await s.insertContent(
      [
        { paragraph: 'one', list: 'number' },
        { paragraph: 'two', list: 'number' },
        { paragraph: 'Break', heading: 2 },
        { paragraph: 'again', list: 'number' },
        { paragraph: 'dot', list: 'bullet', level: 2 },
      ],
      { position: 'document_end' },
    );
    let bytes: Uint8Array | undefined;
    (
      s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }
    ).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    const { doc } = await importDocx(bytes!.slice().buffer as ArrayBuffer);
    expect(labels(doc)).toEqual(['', '1.', '2.', '', '1.', '◦']);
  });
});
