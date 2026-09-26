import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from './docx.js';
import { exportDocx } from './export.js';

const MY_QUOTE =
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="MyQuote">' +
  '<w:name w:val="My Quote"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="1F4E79"/></w:rPr></w:style>';

/** A package whose styles.xml defines MyQuote and whose first paragraph
 *  uses it, the second plain Normal. */
async function withCustomStyle(): Promise<Uint8Array> {
  const base = await exportDocx(
    schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('Quoted')]),
      schema.node('paragraph', null, [schema.text('Plain')]),
    ]),
  );
  const zip = await JSZip.loadAsync(base);
  const styles = await zip.file('word/styles.xml')!.async('string');
  zip.file(
    'word/styles.xml',
    styles.replace('</w:styles>', `${MY_QUOTE}</w:styles>`),
  );
  const doc = await zip.file('word/document.xml')!.async('string');
  zip.file(
    'word/document.xml',
    doc.replace('<w:p>', '<w:p><w:pPr><w:pStyle w:val="MyQuote"/></w:pPr>'),
  );
  return zip.generateAsync({ type: 'uint8array' });
}

describe("a paragraph keeps its document's own style name", () => {
  it('import keeps the id; the default style needs none', async () => {
    const { doc } = await importDocx(
      (await withCustomStyle()).slice().buffer as ArrayBuffer,
    );
    expect(doc.child(0).attrs['styleId']).toBe('MyQuote');
    expect(doc.child(1).attrs['styleId']).toBeNull();
    // Its look is baked, as before: italic and the colour.
    const marks = doc.child(0).firstChild!.marks.map((m) => m.type.name);
    expect(marks).toEqual(expect.arrayContaining(['em', 'textColor']));
  });

  it('export writes it back where the package defines it, and not where it does not', async () => {
    const bytes = await withCustomStyle();
    const opened = await importDocx(bytes.slice().buffer as ArrayBuffer);
    const again = await exportDocx(opened.doc, { carry: opened.raw });
    const xml = await (await JSZip.loadAsync(again))
      .file('word/document.xml')!
      .async('string');
    expect(xml).toContain('<w:pStyle w:val="MyQuote"/>');
    // A new document has no MyQuote to point at.
    const fresh = await exportDocx(opened.doc);
    const freshXml = await (await JSZip.loadAsync(fresh))
      .file('word/document.xml')!
      .async('string');
    expect(freshXml).not.toContain('MyQuote');
  });
});
