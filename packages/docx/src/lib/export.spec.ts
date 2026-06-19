import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from './docx';
import { exportDocx } from './export';

/** Build a schema doc from paragraphs of `{ text, marks }` runs. */
function makeDoc(paras: { text: string; marks?: string[]; attrs?: Record<string, unknown> }[][], pAttrs: Record<string, unknown>[] = []) {
  const ps = paras.map((runs, i) =>
    schema.node(
      'paragraph',
      pAttrs[i] ?? null,
      runs.map((r) => schema.text(r.text, (r.marks ?? []).map((m) => markFor(m, r.attrs)))),
    ),
  );
  return schema.node('doc', null, ps);
}

function markFor(name: string, attrs?: Record<string, unknown>) {
  return schema.marks[name].create(attrs);
}

describe('exportDocx (round-trip)', () => {
  it('round-trips paragraphs and common marks', async () => {
    const doc = makeDoc([
      [{ text: 'Hello ' }, { text: 'bold', marks: ['strong'] }, { text: ' and ' }, { text: 'italic', marks: ['em'] }],
      [{ text: 'plain second paragraph' }],
    ]);

    const bytes = await exportDocx(doc);
    const { doc: back } = await importDocx(bytes);

    expect(back.childCount).toBe(2);
    expect(back.child(0).textContent).toBe('Hello bold and italic');
    expect(back.child(1).textContent).toBe('plain second paragraph');
    const boldRun = [...range(back.child(0))].find((n) => n.text === 'bold');
    expect(boldRun?.marks.map((m) => m.type.name)).toContain('strong');
    const italicRun = [...range(back.child(0))].find((n) => n.text === 'italic');
    expect(italicRun?.marks.map((m) => m.type.name)).toContain('em');
  });

  it('round-trips color, size, vertAlign, paragraph alignment + page break', async () => {
    const doc = makeDoc(
      [
        [
          { text: 'red', marks: ['textColor'], attrs: { color: '#C0392B' } },
          { text: 'big', marks: ['fontSize'], attrs: { size: 18 } },
          { text: 'x', marks: ['vertAlign'], attrs: { value: 'super' } },
        ],
      ],
      [{ align: 'center', pageBreakBefore: true }],
    );

    const { doc: back } = await importDocx(await exportDocx(doc));
    const p0 = back.child(0);
    expect(p0.attrs['align']).toBe('center');
    expect(p0.attrs['pageBreakBefore']).toBe(true);
    const runs = [...range(p0)];
    expect(runs.find((n) => n.text === 'red')?.marks.find((m) => m.type.name === 'textColor')?.attrs['color']).toBe('#C0392B');
    expect(runs.find((n) => n.text === 'big')?.marks.find((m) => m.type.name === 'fontSize')?.attrs['size']).toBe(18);
    expect(runs.find((n) => n.text === 'x')?.marks.find((m) => m.type.name === 'vertAlign')?.attrs['value']).toBe('super');
  });

  it('round-trips a hard break within a paragraph', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('a'), schema.node('hard_break'), schema.text('b')]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const p = back.child(0);
    expect(p.childCount).toBe(3);
    expect(p.child(1).type.name).toBe('hard_break');
  });
});

/** Iterate a block's inline children. */
function* range(block: import('prosemirror-model').Node) {
  for (let i = 0; i < block.childCount; i++) yield block.child(i);
}

// 1×1 transparent PNG.
const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

describe('exportDocx (E2: lists / tables / images / hyperlinks)', () => {
  it('round-trips a list paragraph (numId + level)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { list: { numId: '3', level: 1 } }, [schema.text('an item')]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    expect(back.child(0).attrs['list']).toEqual({ numId: '3', level: 1 });
    expect(back.child(0).textContent).toBe('an item');
  });

  it('round-trips a table with borders + colspan', async () => {
    const cell = (text: string, attrs?: Record<string, unknown>) =>
      schema.node('table_cell', attrs ?? null, [schema.node('paragraph', null, [schema.text(text)])]);
    const doc = schema.node('doc', null, [
      schema.node('table', { borders: { top: true, bottom: true, left: true, right: true, insideH: true, insideV: true } }, [
        schema.node('table_row', null, [cell('A1', { colwidth: [100] }), cell('B1', { colwidth: [120] })]),
        schema.node('table_row', null, [cell('wide', { colspan: 2, colwidth: [100, 120] })]),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const tbl = back.child(0);
    expect(tbl.type.name).toBe('table');
    expect(tbl.attrs['borders']).toBeTruthy();
    expect(tbl.childCount).toBe(2); // two rows
    expect(tbl.child(0).childCount).toBe(2); // two cells
    expect(tbl.child(0).child(0).textContent).toBe('A1');
    expect(tbl.child(1).child(0).attrs['colspan']).toBe(2);
    expect(tbl.child(1).child(0).textContent).toBe('wide');
  });

  it('round-trips an inline image (src + dimensions)', async () => {
    const src = `data:image/png;base64,${PNG_1PX}`;
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [schema.node('image', { src, width: 50, height: 40 })]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const img = [...range(back.child(0))].find((n) => n.type.name === 'image');
    expect(img).toBeTruthy();
    expect(img?.attrs['width']).toBe(50);
    expect(img?.attrs['height']).toBe(40);
    expect(String(img?.attrs['src'])).toBe(src);
  });

  it('round-trips a hyperlink (link mark + href)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('see '),
        schema.text('the site', [schema.marks['link'].create({ href: 'https://prosemirror.net/' })]),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const linked = [...range(back.child(0))].find((n) => n.text === 'the site');
    expect(linked?.marks.find((m) => m.type.name === 'link')?.attrs['href']).toBe('https://prosemirror.net/');
  });
});
