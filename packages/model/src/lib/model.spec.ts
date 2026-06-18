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
