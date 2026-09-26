import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { schema } from '@shadow-garden/bapbong-model';
import { exportDocx } from './export.js';
import { documentStyles, styleFormatting } from './style-catalog.js';

const MY_QUOTE =
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="MyQuote">' +
  '<w:name w:val="My Quote &amp; Note"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:before="240"/><w:ind w:left="720"/><w:jc w:val="center"/></w:pPr>' +
  '<w:rPr><w:i/><w:color w:val="1F4E79"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Secret"><w:name w:val="Secret"/><w:semiHidden/></w:style>' +
  '<w:style w:type="character" w:styleId="Strong2"><w:name w:val="Strong 2"/><w:rPr><w:b/></w:rPr></w:style>';

async function pack(): Promise<JSZip> {
  const base = await exportDocx(
    schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('x')]),
    ]),
  );
  const zip = await JSZip.loadAsync(base);
  const styles = await zip.file('word/styles.xml')!.async('string');
  zip.file(
    'word/styles.xml',
    styles.replace('</w:styles>', `${MY_QUOTE}</w:styles>`),
  );
  return zip;
}

describe('the document style catalog', () => {
  it('lists the styles with their names, kinds and flags', async () => {
    const xml = await (await pack()).file('word/styles.xml')!.async('string');
    const styles = documentStyles(xml);
    const byId = Object.fromEntries(styles.map((s) => [s.id, s]));
    expect(byId['MyQuote']).toMatchObject({
      name: 'My Quote & Note',
      type: 'paragraph',
      basedOn: 'Normal',
      custom: true,
      hidden: false,
    });
    expect(byId['Normal']).toMatchObject({
      type: 'paragraph',
      isDefault: true,
    });
    expect(byId['Secret'].hidden).toBe(true);
    expect(byId['Strong2'].type).toBe('character');
  });

  it("gives a paragraph style's look exactly as the importer resolves it", async () => {
    const f = await styleFormatting(await pack(), 'MyQuote');
    expect(f?.attrs).toMatchObject({
      styleId: 'MyQuote',
      align: 'center',
      indent: { left: 48 },
      spacing: expect.objectContaining({ before: 16 }),
    });
    expect(f?.marks.map((m) => m.type)).toEqual(
      expect.arrayContaining(['em', 'textColor']),
    );
    expect(await styleFormatting(await pack(), 'Strong2')).toBeNull(); // not a paragraph style
    expect(await styleFormatting(await pack(), 'Nope')).toBeNull();
  });
});
