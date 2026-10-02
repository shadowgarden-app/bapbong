/**
 * A document saved after a web paste held `<w:color w:val="rgb(0, 0, 0)"/>`
 * on every run: the importer passed it on as "#RGB(0, 0, 0)", the canvas
 * could not draw it, and the whole text took a stray light blue.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from './docx';
import { exportDocx } from './export';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function docxWith(rPr: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>Mưa lớn</w:t></w:r></w:p></w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

const colorOf = (doc: import('prosemirror-model').Node) =>
  doc
    .child(0)
    .child(0)
    .marks.find((m) => m.type.name === 'textColor')?.attrs['color'];

describe('a CSS colour in w:color', () => {
  it('is read as the colour it names', async () => {
    const { doc } = await importDocx(
      await docxWith('<w:color w:val="rgb(0, 0, 0)"/>'),
    );
    expect(colorOf(doc)).toBe('#000000');
  });

  it('that names no colour is no colour, not a broken one', async () => {
    const { doc } = await importDocx(
      await docxWith('<w:color w:val="rgb(oops)"/>'),
    );
    expect(colorOf(doc)).toBeUndefined();
  });

  it('is written back as hex, whatever the model held', async () => {
    const run = schema.text('Mưa lớn', [
      schema.marks['textColor'].create({ color: 'rgb(31, 78, 121)' }),
    ]);
    const out = await exportDocx(
      schema.node('doc', null, [schema.node('paragraph', null, [run])]),
    );
    const xml = await (await JSZip.loadAsync(out))
      .file('word/document.xml')!
      .async('string');
    expect(xml).toContain('<w:color w:val="1F4E79"/>');
    expect(xml).not.toContain('rgb(');
  });
});
