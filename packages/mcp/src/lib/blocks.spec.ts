import { describe, expect, it } from 'vitest';
import {
  exportDocx,
  importDocx,
  schema,
} from '@shadow-garden/bapbong-headless';
import { contentToNodes, lengthToPx } from './blocks.js';
import { ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';

const WIDTH = 602;
const build = (content: Parameters<typeof contentToNodes>[0]) =>
  contentToNodes(content, schema, { contentWidth: WIDTH });

describe('lengthToPx', () => {
  it('reads cm numbers, %, cm, in, px', () => {
    expect(lengthToPx(2.54, WIDTH)).toBeCloseTo(96);
    expect(lengthToPx('50%', WIDTH)).toBe(301);
    expect(lengthToPx('1in', WIDTH)).toBe(96);
    expect(lengthToPx('40px', WIDTH)).toBe(40);
    expect(() => lengthToPx('wide', WIDTH)).toThrow(ContentError);
  });
});

describe('contentToNodes', () => {
  it('a plain string is one paragraph per line — unchanged behaviour', () => {
    const nodes = build('a\n\nb');
    expect(nodes.map((n) => n.type.name)).toEqual(['paragraph', 'paragraph', 'paragraph']);
    expect(nodes[0].textContent).toBe('a');
    expect(nodes[1].childCount).toBe(0);
  });

  it('paragraph blocks carry heading / style / align / tabs / marks', () => {
    const [h, t, p] = build([
      { paragraph: 'Title', heading: 1, align: 'center' },
      { paragraph: 'Sub', style: 'Subtitle' },
      {
        paragraph: ['Name: ', { tab: true }],
        tabs: [{ at: '100%', align: 'right', leader: 'dot' }],
        bold: true,
      },
    ]);
    expect(h.attrs['heading']).toBe(1);
    expect(h.attrs['align']).toBe('center');
    expect(t.attrs['styleId']).toBe('Subtitle');
    expect(t.attrs['heading']).toBeNull();
    expect(p.attrs['tabs']).toEqual([{ pos: WIDTH, val: 'right', leader: 'dot' }]);
    expect(p.textContent).toBe('Name: \t');
    expect(p.firstChild!.marks.map((m) => m.type.name)).toEqual(['strong']);
  });

  it('inline marks add to the paragraph marks', () => {
    const [p] = build([{ paragraph: ['a', { text: 'b', italic: true }], bold: true }]);
    expect(p.child(1).marks.map((m) => m.type.name).sort()).toEqual(['em', 'strong']);
  });

  it('refuses a newline inside a paragraph', () => {
    expect(() => build([{ paragraph: 'a\nb' }])).toThrow(/No "\\n"/);
    expect(() => build([{ paragraph: [{ text: 'a\nb' }] }])).toThrow(ContentError);
  });

  it('a table block: equal columns by default, header row bold + repeated, Table Grid when the host has it', () => {
    const [t] = contentToNodes(
      [{ table: [['A', 'B'], ['1', '2']], header: true }],
      schema,
      { contentWidth: WIDTH, tableStyle: { styleId: 'TableGrid' } },
    );
    expect(t.type.name).toBe('table');
    expect(t.attrs['styleId']).toBe('TableGrid');
    expect(t.childCount).toBe(2);
    const header = t.child(0);
    expect(header.attrs['header']).toBe(true);
    const cell = header.child(0);
    expect(cell.attrs['colwidth']).toEqual([301]);
    expect(cell.firstChild!.firstChild!.marks.map((m) => m.type.name)).toEqual(['strong']);
    expect(t.child(1).child(0).firstChild!.firstChild!.marks).toEqual([]);
  });

  it('without a host style a grid is direct borders; "outer" nils the inside; "none" has none', () => {
    const [grid, outer, none] = build([
      { table: [['a']] },
      { table: [['a']], borders: 'outer' },
      { table: [['a']], borders: 'none' },
    ]);
    expect(grid.attrs['styleId']).toBeNull();
    expect((grid.attrs['borders'] as { insideV: unknown }).insideV).toBeTruthy();
    expect((outer.attrs['borders'] as { insideV: unknown; top: unknown }).insideV).toBeNull();
    expect((outer.attrs['borders'] as { insideV: unknown; top: unknown }).top).toBeTruthy();
    expect(none.attrs['borders']).toBeNull();
    expect(none.attrs['styleId']).toBeNull();
  });

  it('widths, colspan, shading, cell alignment', () => {
    const [t] = build([
      {
        table: [
          ['#', 'Item', 'Total'],
          [{ text: 'Sum', colspan: 2, align: 'right' }, { text: '9', shading: '#eeeeee' }],
        ],
        widths: ['10%', '60%', '30%'],
        align: 'center',
      },
    ]);
    expect(t.attrs['align']).toBe('center');
    const r0 = t.child(0);
    expect(r0.child(0).attrs['colwidth']).toEqual([60]);
    expect(r0.child(1).attrs['colwidth']).toEqual([361]);
    const merged = t.child(1).child(0);
    expect(merged.attrs['colspan']).toBe(2);
    expect(merged.attrs['colwidth']).toEqual([60, 361]);
    expect(merged.firstChild!.attrs['align']).toBe('right');
    expect(t.child(1).child(1).attrs['background']).toBe('#eeeeee');
  });

  it('reports ragged rows and a widths list that does not fit, with the block index', () => {
    expect(() => build([{ table: [['a', 'b'], ['c']] }])).toThrow(/Block 0: Row 1 covers 1 column/);
    expect(() => build(['x', { table: [['a', 'b']], widths: [1] }])).toThrow(/Block 1: widths lists 1/);
    expect(() => build([{ table: [] }])).toThrow(ContentError);
  });
});

describe('insert_content with blocks, through a headless session', () => {
  async function open(): Promise<HeadlessSession> {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('Intro')]),
      schema.node('paragraph', null, [schema.text('Outro')]),
    ]);
    return HeadlessSession.open(await exportDocx(doc));
  }

  it('inserts a heading and a table after an anchor and reads the cells back as blocks', async () => {
    const s = await open();
    await s.insertContent(
      [
        { paragraph: 'Items', heading: 2 },
        { table: [['Item', 'Qty'], ['Pen', '2']], header: true, widths: [10, 4] },
      ],
      { position: 'after', text: 'Intro' },
    );
    const snap = await s.snapshot();
    expect(snap.blocks.map((b) => [b.type, b.text])).toEqual([
      ['paragraph', 'Intro'],
      ['heading2', 'Items'],
      ['paragraph', 'Item'],
      ['paragraph', 'Qty'],
      ['paragraph', 'Pen'],
      ['paragraph', '2'],
      ['paragraph', 'Outro'],
    ]);
    // The document's sheet now defines Table Grid for the layout.
    const state = (s as unknown as { state: { doc: { attrs: Record<string, unknown> } } }).state;
    expect((state.doc.attrs['tableStyles'] as Record<string, unknown>)['TableGrid']).toBeTruthy();
    // Cell text is editable like any paragraph.
    await s.replaceText('Pen', 'Pencil');
    expect((await s.snapshot()).blocks[4].text).toBe('Pencil');
  });

  it('round-trips through .docx: table, merged cell, tab leader, heading survive export → import', async () => {
    const s = await open();
    await s.insertContent(
      [
        { paragraph: 'Form', heading: 1 },
        {
          paragraph: ['Name: ', { tab: true }],
          tabs: [{ at: '100%', align: 'right', leader: 'dot' }],
        },
        {
          table: [
            ['a', 'b', 'c'],
            [{ text: 'sum', colspan: 2 }, 'z'],
          ],
          borders: 'none',
        },
      ],
      { position: 'document_end' },
    );
    let bytes: Uint8Array | undefined;
    (s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    expect(bytes).toBeDefined();
    const { doc } = await importDocx(bytes!.slice().buffer as ArrayBuffer);
    const kinds: string[] = [];
    doc.descendants((n) => {
      if (n.type.name === 'table') kinds.push('table');
      if (n.type.name === 'paragraph' && n.attrs['heading']) kinds.push(`h${n.attrs['heading']}`);
      if (n.type.name === 'paragraph' && n.attrs['tabs']) kinds.push('tabs');
      if (n.type.name === 'table_cell' && n.attrs['colspan'] === 2) kinds.push('span2');
      return true;
    });
    expect(kinds).toEqual(['h1', 'tabs', 'table', 'span2']);
    const tabbed = (() => {
      let found: Record<string, unknown> | null = null;
      doc.descendants((n) => {
        if (n.type.name === 'paragraph' && n.attrs['tabs']) found = n.attrs;
        return !found;
      });
      return found as unknown as { tabs: { val: string; leader?: string }[] };
    })();
    expect(tabbed.tabs[0]).toMatchObject({ val: 'right', leader: 'dot' });
  });

  it('a content error comes back as a tool error, not a crash', async () => {
    const s = await open();
    await expect(
      s.insertContent([{ table: [['a', 'b'], ['c']] }], { position: 'document_end' }),
    ).rejects.toThrow(ContentError);
  });
});
