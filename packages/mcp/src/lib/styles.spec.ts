import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { ContentError } from './contract.js';
import { HeadlessSession } from './headless-session.js';

const MY_QUOTE =
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="MyQuote">' +
  '<w:name w:val="My Quote"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:ind w:left="720"/><w:jc w:val="center"/></w:pPr>' +
  '<w:rPr><w:i/><w:color w:val="1F4E79"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Hidden1"><w:name w:val="Hidden one"/><w:semiHidden/></w:style>';

type Raw = {
  file(path: string): { async(t: 'string'): Promise<string> } | null;
  file(path: string, data: string): unknown;
};

/** A session over a document whose styles.xml defines "My Quote". */
async function open(): Promise<HeadlessSession> {
  const link = schema.marks['link'].create({ href: 'https://example.com' });
  const s = await HeadlessSession.open(
    await exportDocx(
      schema.node('doc', null, [
        schema.node('paragraph', null, [
          schema.text('Wise words, '),
          schema.text('source', [link]),
        ]),
        schema.node('paragraph', null, [schema.text('Plain')]),
      ]),
    ),
  );
  const raw = (s as unknown as { raw: Raw }).raw;
  const xml = await raw.file('word/styles.xml')!.async('string');
  raw.file(
    'word/styles.xml',
    xml.replace('</w:styles>', `${MY_QUOTE}</w:styles>`),
  );
  return s;
}

const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;
const marksOf = (n: PMNode) => n.marks.map((m) => m.type.name).sort();

describe("the document's own paragraph styles", () => {
  it('list_styles: the paragraph styles with names, custom flag and use; hidden ones only when used', async () => {
    const s = await open();
    expect(s.capabilities.styles).toBe(true);
    const styles = await s.listStyles();
    const quote = styles.find((st) => st.id === 'MyQuote');
    expect(quote).toMatchObject({ name: 'My Quote', custom: true, used: 0 });
    expect(styles.some((st) => st.id === 'Hidden1')).toBe(false);
    expect(styles.some((st) => st.id === 'Normal')).toBe(true);
  });

  it('apply_formatting by name gives the paragraph the style: its look, its name, the link kept', async () => {
    const s = await open();
    await s.applyFormatting({ blockIndex: 0 }, { style: 'my quote' });
    const p = docOf(s).child(0);
    expect(p.attrs).toMatchObject({
      styleId: 'MyQuote',
      align: 'center',
      indent: { left: 48 },
      heading: null,
    });
    // The importer bakes the size the style chain resolves, so it comes too.
    expect(marksOf(p.child(0))).toEqual(['em', 'fontSize', 'textColor']);
    expect(marksOf(p.child(1))).toEqual([
      'em',
      'fontSize',
      'link',
      'textColor',
    ]);
    expect((await s.listStyles()).find((st) => st.id === 'MyQuote')?.used).toBe(
      1,
    );
    // A style with an outline level makes a heading.
    await s.applyFormatting({ blockIndex: 1 }, { style: 'heading 1' });
    expect(docOf(s).child(1).attrs['heading']).toBe(1);
  });

  it('a block in a document style gets its look; formatting on the run adds to it', async () => {
    const s = await open();
    await s.insertContent(
      [
        {
          paragraph: ['Quoted, ', { text: 'firmly', bold: true }],
          style: 'MyQuote',
        },
      ],
      { position: 'document_end' },
    );
    const p = docOf(s).lastChild!;
    expect(p.attrs['styleId']).toBe('MyQuote');
    expect(p.attrs['align']).toBe('center');
    expect(marksOf(p.child(1))).toEqual([
      'em',
      'fontSize',
      'strong',
      'textColor',
    ]);
  });

  it('an unknown style is named in the refusal with the ones there are', async () => {
    const s = await open();
    await expect(
      s.applyFormatting({ blockIndex: 0 }, { style: 'Fancy' }),
    ).rejects.toThrow(/No paragraph style "Fancy".*"My Quote"/);
    await expect(
      s.insertContent([{ paragraph: 'x', style: 'Fancy' }], {
        position: 'document_end',
      }),
    ).rejects.toThrow(ContentError);
  });

  it('the style name survives the file', async () => {
    const s = await open();
    await s.applyFormatting({ blockIndex: 0 }, { style: 'MyQuote' });
    let bytes: Uint8Array | undefined;
    (
      s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }
    ).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    const back = await HeadlessSession.open(bytes!);
    expect(docOf(back).child(0).attrs['styleId']).toBe('MyQuote');
  });
});
