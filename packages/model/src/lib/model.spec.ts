import { schema } from './model';

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

  it('can build a paragraph with a marked text node', () => {
    const p = schema.nodes.paragraph.create(null, [
      schema.text('hi', [schema.marks.strong.create()]),
    ]);
    const doc = schema.nodes.doc.create(null, [p]);
    expect(doc.textContent).toBe('hi');
    expect(doc.child(0).child(0).marks[0].type.name).toBe('strong');
  });
});
