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
    expect(schema.marks.textColor.create({ color: '#FF0000' }).attrs.color).toBe('#FF0000');
    expect(schema.marks.fontSize.create({ size: 12 }).attrs.size).toBe(12);
    expect(schema.marks.fontFamily.create({ family: 'Calibri' }).attrs.family).toBe('Calibri');
  });

  describe('paste parseDOM rules (structural getAttrs, no DOM needed)', () => {
    // Minimal stand-in for an HTMLElement in the node-env tests.
    const el = (attrs: Record<string, string>) => ({
      getAttribute: (n: string) => attrs[n] ?? null,
    });
    type GetAttrs = (el: unknown) => Record<string, unknown> | false;
    const rules = (spec: { parseDOM?: readonly { tag?: string; getAttrs?: unknown }[] }) =>
      new Map((spec.parseDOM ?? []).map((r) => [r.tag, r.getAttrs as GetAttrs]));

    it('parses h1–h6 into the heading attr and text-align into align', () => {
      const p = rules(schema.nodes.paragraph.spec as never);
      expect(p.get('p')!(el({}))).toEqual({ heading: null, align: null });
      expect(p.get('h2')!(el({}))).toEqual({ heading: 2, align: null });
      expect(p.get('h6')!(el({ style: 'color:red;text-align: center' }))).toEqual({
        heading: 6,
        align: 'center',
      });
      // 'left' is the default — normalized to null
      expect(p.get('p')!(el({ style: 'text-align:left' }))).toEqual({ heading: null, align: null });
    });

    it('accepts only http(s)/mailto/#anchor link hrefs', () => {
      const a = rules(schema.marks.link.spec as never).get('a[href]')!;
      expect(a(el({ href: 'https://x.vn/a' }))).toEqual({ href: 'https://x.vn/a' });
      expect(a(el({ href: 'mailto:x@y.z' }))).toEqual({ href: 'mailto:x@y.z' });
      expect(a(el({ href: '#s1' }))).toEqual({ href: '#s1' });
      // eslint-disable-next-line no-script-url
      expect(a(el({ href: 'javascript:alert(1)' }))).toBe(false);
      expect(a(el({ href: 'data:text/html,x' }))).toBe(false);
      expect(a(el({ href: '' }))).toBe(false);
    });

    it('accepts only data:image/* img srcs, with numeric dimensions', () => {
      const img = rules(schema.nodes.image.spec as never).get('img[src]')!;
      expect(img(el({ src: 'data:image/png;base64,AAAA', width: '120', height: '80.6' }))).toEqual(
        { src: 'data:image/png;base64,AAAA', alt: '', width: 120, height: 81 },
      );
      expect(img(el({ src: 'data:image/jpeg;base64,BB', alt: 'chart' }))).toEqual({
        src: 'data:image/jpeg;base64,BB',
        alt: 'chart',
        width: null,
        height: null,
      });
      expect(img(el({ src: 'https://cdn.x/pic.png' }))).toBe(false); // remote: CORS/export unsafe
      expect(img(el({ src: 'data:text/html,x' }))).toBe(false);
    });
  });

  it('paragraph carries an optional list attribute (default null)', () => {
    expect(schema.nodes.paragraph.create().attrs.list).toBeNull();
    const p = schema.nodes.paragraph.create({ list: { numId: '1', level: 0, marker: '1.' } });
    expect((p.attrs.list as { marker: string }).marker).toBe('1.');
  });

  it('supports the link mark and image node', () => {
    expect(schema.marks.link.create({ href: 'https://x' }).attrs.href).toBe('https://x');
    const img = schema.nodes.image.create({ src: 'data:,', width: 10, height: 20, alt: 'a' });
    expect(img.type.name).toBe('image');
    expect(img.attrs.width).toBe(10);
    expect(img.isInline).toBe(true);
  });

  it('can build a table (row -> cell -> block) with colspan', () => {
    const cell = schema.nodes.table_cell.create({ colspan: 2 }, [
      schema.nodes.paragraph.create(null, [schema.text('x')]),
    ]);
    const table = schema.nodes.table.create(null, [schema.nodes.table_row.create(null, [cell])]);
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
    const mention = commentSchema.nodes['mention'].create({ id: 'alice', label: 'Alice Nguyễn' });
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
