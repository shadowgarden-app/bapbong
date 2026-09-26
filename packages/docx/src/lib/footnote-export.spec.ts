import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from './docx.js';
import { exportDocx } from './export.js';

const ref = (num: number, shown: string, id: string | null = null) =>
  schema.text(shown, [
    schema.marks['vertAlign'].create({ value: 'super' }),
    schema.marks['footnote'].create({ num, id }),
  ]);
const body = (text: string) =>
  schema
    .node('doc', null, [schema.node('paragraph', null, [schema.text(text)])])
    .toJSON();

const texts = (notes: Record<number, PMNode>) =>
  Object.values(notes).map((n) => n.textContent);

describe('footnotes written in bapbong are saved', () => {
  it('into a new footnotes part, with the separators, the relationship and the content type', async () => {
    const doc = schema.node(
      'doc',
      { footnoteBodies: { '1': body('See the annex.') } },
      [schema.node('paragraph', null, [schema.text('Claim'), ref(1, '1')])],
    );
    const bytes = await exportDocx(doc);
    const zip = await JSZip.loadAsync(bytes);
    const notes = await zip.file('word/footnotes.xml')!.async('string');
    expect(notes).toContain('w:type="separator" w:id="-1"');
    expect(notes).toMatch(
      /<w:footnote w:id="1">[\s\S]*<w:footnoteRef\/>[\s\S]*See the annex\./,
    );
    expect(await zip.file('word/document.xml')!.async('string')).toContain(
      '<w:footnoteReference w:id="1"/>',
    );
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
      'PartName="/word/footnotes.xml"',
    );
    expect(
      await zip.file('word/_rels/document.xml.rels')!.async('string'),
    ).toContain('Target="footnotes.xml"');
    const back = await importDocx(bytes.slice().buffer as ArrayBuffer);
    expect(texts(back.footnotes)).toEqual(['1 See the annex.']);
  });

  it('next to the footnotes a file already has, under an id above theirs', async () => {
    const first = await exportDocx(
      schema.node('doc', { footnoteBodies: { '1': body('Old note.') } }, [
        schema.node('paragraph', null, [schema.text('A'), ref(1, '1')]),
      ]),
    );
    const opened = await importDocx(first.slice().buffer as ArrayBuffer);
    const para = opened.doc.child(0);
    const edited = opened.doc.type.create(
      { ...opened.doc.attrs, footnoteBodies: { '7': body('New note.') } },
      [
        para.type.create(para.attrs, [
          ...contentOf(para),
          schema.text(' B'),
          ref(7, '2'),
        ]),
      ],
    );
    const second = await exportDocx(edited, { carry: opened.raw });
    const notes = await (await JSZip.loadAsync(second))
      .file('word/footnotes.xml')!
      .async('string');
    expect(notes).toContain('<w:footnote w:id="2">');
    const back = await importDocx(second.slice().buffer as ArrayBuffer);
    expect(texts(back.footnotes)).toEqual(['1 Old note.', '2 New note.']);
  });
});

function contentOf(n: PMNode): PMNode[] {
  const out: PMNode[] = [];
  n.forEach((c) => out.push(c));
  return out;
}
