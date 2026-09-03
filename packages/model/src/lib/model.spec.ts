import { commentSchema, schema } from './model';

describe('schema', () => {
  it('defines doc/paragraph/text nodes', () => {
    expect(schema.nodes.doc).toBeDefined();
    expect(schema.nodes.paragraph).toBeDefined();
    expect(schema.nodes.text).toBeDefined();
  });

  it('defines character + formatting marks', () => {
    expect(Object.keys(schema.marks)).toEqual(
      expect.arrayContaining([
        'strong',
        'em',
        'underline',
        'strike',
        'textColor',
        'fontSize',
        'fontFamily',
      ]),
    );
  });

  it('formatting marks carry attributes', () => {
    expect(
      schema.marks.textColor.create({ color: '#FF0000' }).attrs.color,
    ).toBe('#FF0000');
    expect(schema.marks.fontSize.create({ size: 12 }).attrs.size).toBe(12);
    expect(
      schema.marks.fontFamily.create({ family: 'Calibri' }).attrs.family,
    ).toBe('Calibri');
  });

  describe('paste parseDOM rules (structural getAttrs, no DOM needed)', () => {
    // Minimal stand-in for an HTMLElement in the node-env tests.
    const el = (attrs: Record<string, string>) => ({
      getAttribute: (n: string) => attrs[n] ?? null,
    });
    type GetAttrs = (el: unknown) => Record<string, unknown> | false;
    const rules = (spec: {
      parseDOM?: readonly { tag?: string; getAttrs?: unknown }[];
    }) =>
      new Map(
        (spec.parseDOM ?? []).map((r) => [r.tag, r.getAttrs as GetAttrs]),
      );

    it('parses h1–h6 into the heading attr and text-align into align', () => {
      const p = rules(schema.nodes.paragraph.spec as never);
      const base = { borders: null, markFont: null, carry: null };
      expect(p.get('p')!(el({}))).toEqual({
        heading: null,
        align: null,
        ...base,
      });
      expect(p.get('h2')!(el({}))).toEqual({
        heading: 2,
        align: null,
        ...base,
      });
      expect(
        p.get('h6')!(el({ style: 'color:red;text-align: center' })),
      ).toEqual({
        heading: 6,
        align: 'center',
        ...base,
      });
      // 'left' is the default — normalized to null
      expect(p.get('p')!(el({ style: 'text-align:left' }))).toEqual({
        heading: null,
        align: null,
        ...base,
      });
    });

    it('accepts only http(s)/mailto/#anchor link hrefs', () => {
      const a = rules(schema.marks.link.spec as never).get('a[href]')!;
      // Pasted HTML carries no w:tgtFrame equivalent unless it came from us.
      expect(a(el({ href: 'https://x.vn/a' }))).toEqual({
        href: 'https://x.vn/a',
        targetFrame: null,
      });
      expect(a(el({ href: 'mailto:x@y.z' }))).toEqual({
        href: 'mailto:x@y.z',
        targetFrame: null,
      });
      expect(a(el({ href: '#s1' }))).toEqual({
        href: '#s1',
        targetFrame: null,
      });
      expect(a(el({ href: '#s1', 'data-target-frame': '_blank' }))).toEqual({
        href: '#s1',
        targetFrame: '_blank',
      });
      // eslint-disable-next-line no-script-url
      expect(a(el({ href: 'javascript:alert(1)' }))).toBe(false);
      expect(a(el({ href: 'data:text/html,x' }))).toBe(false);
      expect(a(el({ href: '' }))).toBe(false);
    });

    it('accepts only data:image/* img srcs, with numeric dimensions', () => {
      const img = rules(schema.nodes.image.spec as never).get('img[src]')!;
      expect(
        img(
          el({
            src: 'data:image/png;base64,AAAA',
            width: '120',
            height: '80.6',
          }),
        ),
      ).toEqual({
        src: 'data:image/png;base64,AAAA',
        alt: '',
        title: null,
        width: 120,
        height: 81,
        // No data-crop / data-outline / data-effect-extent on pasted HTML —
        // an external <img> has no opinion on which part of the bitmap
        // shows, on a border, or on the room Word keeps around one.
        crop: null,
        outline: null,
        effectExtent: null,
      });
      expect(
        img(el({ src: 'data:image/jpeg;base64,BB', alt: 'chart' })),
      ).toEqual({
        src: 'data:image/jpeg;base64,BB',
        alt: 'chart',
        title: null,
        width: null,
        height: null,
        crop: null,
        outline: null,
        effectExtent: null,
      });
      expect(img(el({ src: 'https://cdn.x/pic.png' }))).toBe(false); // remote: CORS/export unsafe
      expect(img(el({ src: 'data:text/html,x' }))).toBe(false);
    });
  });

  it('paragraph carries an optional list attribute (default null)', () => {
    expect(schema.nodes.paragraph.create().attrs.list).toBeNull();
    const p = schema.nodes.paragraph.create({
      list: { numId: '1', level: 0, marker: '1.' },
    });
    expect((p.attrs.list as { marker: string }).marker).toBe('1.');
  });

  it('supports the link mark and image node', () => {
    expect(schema.marks.link.create({ href: 'https://x' }).attrs.href).toBe(
      'https://x',
    );
    const img = schema.nodes.image.create({
      src: 'data:,',
      width: 10,
      height: 20,
      alt: 'a',
    });
    expect(img.type.name).toBe('image');
    expect(img.attrs.width).toBe(10);
    expect(img.isInline).toBe(true);
  });

  it('can build a table (row -> cell -> block) with colspan', () => {
    const cell = schema.nodes.table_cell.create({ colspan: 2 }, [
      schema.nodes.paragraph.create(null, [schema.text('x')]),
    ]);
    const table = schema.nodes.table.create(null, [
      schema.nodes.table_row.create(null, [cell]),
    ]);
    expect(table.type.name).toBe('table');
    expect(table.firstChild?.firstChild?.attrs.colspan).toBe(2);
    expect(table.textContent).toBe('x');
  });

  it('can build a paragraph with a marked text node', () => {
    const p = schema.nodes.paragraph.create(null, [
      schema.text('hi', [schema.marks.strong.create()]),
    ]);
    const doc = schema.nodes.doc.create(null, [p]);
    expect(doc.textContent).toBe('hi');
    expect(doc.child(0).child(0).marks[0].type.name).toBe('strong');
  });
});

describe('commentSchema', () => {
  it('mention is an inline atom whose textContent reads "@label"', () => {
    const mention = commentSchema.nodes['mention'].create({
      id: 'alice',
      label: 'Alice Nguyễn',
    });
    const p = commentSchema.nodes['paragraph'].create(null, [
      commentSchema.text('hi '),
      mention,
      commentSchema.text(' bye'),
    ]);
    const doc = commentSchema.nodes['doc'].create(null, [p]);
    expect(mention.isAtom).toBe(true);
    expect(mention.isInline).toBe(true);
    expect(doc.textContent).toBe('hi @Alice Nguyễn bye'); // leafText
  });

  it('round-trips rich table/cell attrs through toDOM/parseDOM (clipboard)', () => {
    // ProseMirror's clipboard is a toDOM → parseDOM pass: complex attrs ride
    // data-* JSON carriers or an internal copy/paste silently drops them.
    const side = { width: 1, style: 'solid', color: '#000000' };
    const asEl = (attrs: Record<string, string>) => ({
      getAttribute: (n: string) => attrs[n] ?? null,
    });
    const getAttrs = (spec: unknown, i = 0) =>
      (spec as { parseDOM: { getAttrs: (el: unknown) => unknown }[] }).parseDOM[
        i
      ].getAttrs;

    const cell = schema.nodes.table_cell.create(
      {
        colspan: 2,
        colwidth: [80, 120],
        background: '#D9E2F3',
        vAlign: 'center',
        borders: { top: side },
        padding: { left: 5, top: 4 },
      },
      [schema.nodes.paragraph.create()],
    );
    const row = schema.nodes.table_row.create(
      { header: true, height: { value: 50, exact: false } },
      [cell],
    );
    const look = {
      firstRow: true,
      lastRow: false,
      firstCol: true,
      lastCol: false,
      hBand: true,
      vBand: false,
    };
    const table = schema.nodes.table.create(
      {
        borders: { top: side, insideH: side },
        cellPadding: { left: 4 },
        align: 'center',
        styleId: 'LightGrid-Accent3',
        look,
      },
      [row],
    );

    const domAttrs = (node: typeof table) =>
      (
        node.type.spec.toDOM as (
          n: typeof table,
        ) => [string, Record<string, string>, ...unknown[]]
      )(node)[1];

    const tableBack = getAttrs(schema.nodes.table.spec)(
      asEl(domAttrs(table)),
    ) as Record<string, unknown>;
    expect(tableBack['borders']).toEqual({ top: side, insideH: side });
    expect(tableBack['cellPadding']).toEqual({ left: 4 });
    expect(tableBack['align']).toBe('center');
    // The live-theming pair: without these an internal copy/paste turns a
    // styled table into an unstyled one.
    expect(tableBack['styleId']).toBe('LightGrid-Accent3');
    expect(tableBack['look']).toEqual(look);

    const rowBack = getAttrs(schema.nodes.table_row.spec)(
      asEl(domAttrs(row)),
    ) as Record<string, unknown>;
    expect(rowBack['header']).toBe(true);
    expect(rowBack['height']).toEqual({ value: 50, exact: false });

    const cellBack = getAttrs(schema.nodes.table_cell.spec)(
      asEl(domAttrs(cell)),
    ) as Record<string, unknown>;
    expect(cellBack['colspan']).toBe(2);
    expect(cellBack['colwidth']).toEqual([80, 120]);
    expect(cellBack['background']).toBe('#D9E2F3');
    expect(cellBack['vAlign']).toBe('center');
    expect(cellBack['borders']).toEqual({ top: side });
    expect(cellBack['padding']).toEqual({ left: 5, top: 4 });

    const para = schema.nodes.paragraph.create({ borders: { top: side } });
    const paraBack = getAttrs(schema.nodes.paragraph.spec)(
      asEl(domAttrs(para)),
    ) as Record<string, unknown>;
    expect(paraBack['borders']).toEqual({ top: side });
  });

  it('mention round-trips through JSON keeping its attrs', () => {
    const doc = commentSchema.nodes['doc'].create(null, [
      commentSchema.nodes['paragraph'].create(null, [
        commentSchema.nodes['mention'].create({ id: 'bob', label: 'Bob' }),
      ]),
    ]);
    const back = commentSchema.nodeFromJSON(doc.toJSON());
    const node = back.child(0).child(0);
    expect(node.type.name).toBe('mention');
    expect(node.attrs['id']).toBe('bob');
    expect(node.attrs['label']).toBe('Bob');
  });
});
