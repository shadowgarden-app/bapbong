import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from './docx.js';
import { exportDocx } from './export.js';

const INSTR = 'TOC \\o "1-3" \\h \\z \\u';

function tocDoc(dirty: boolean) {
  const field = {
    kind: 'toc',
    instr: INSTR,
    ...(dirty ? { dirty: true } : {}),
  };
  const entry = (text: string, anchor: string, level: number) =>
    schema.node(
      'paragraph',
      {
        field,
        styleId: `TOC${level}`,
        tabs: [{ pos: 600, val: 'right', leader: 'dot' }],
      },
      [
        schema.text(`${text}\t1`, [
          schema.marks['link'].create({ href: `#${anchor}` }),
        ]),
      ],
    );
  return schema.node('doc', null, [
    entry('Intro', '_Toc1', 1),
    entry('Detail', '_Toc2', 2),
    schema.node('paragraph', { heading: 1, bookmarks: ['_Toc1'] }, [
      schema.text('Intro'),
    ]),
    schema.node('paragraph', { heading: 2, bookmarks: ['_Toc2'] }, [
      schema.text('Detail'),
    ]),
  ]);
}

const documentXml = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');

describe('a table of contents is saved as a field', () => {
  it('begin + instruction + separate before the first entry, end after the last, TOC styles defined', async () => {
    const bytes = await exportDocx(tocDoc(false));
    const xml = await documentXml(bytes);
    const paras = xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [];
    expect(paras[0]).toMatch(
      /fldCharType="begin"\/>[\s\S]*TOC \\o &quot;1-3&quot;[\s\S]*fldCharType="separate"/,
    );
    expect(paras[0]).not.toContain('fldCharType="end"');
    expect(paras[1]).toContain('fldCharType="end"');
    expect(paras[1]).toContain('<w:pStyle w:val="TOC2"/>');
    expect(xml).not.toContain('w:dirty');
    const styles = await (await JSZip.loadAsync(bytes))
      .file('word/styles.xml')!
      .async('string');
    expect(styles).toContain('w:styleId="TOC2"');
  });

  it('comes back as one TOC span with its links, and a stale one asks Word to update', async () => {
    const bytes = await exportDocx(tocDoc(false));
    const { doc } = await importDocx(bytes.slice().buffer as ArrayBuffer);
    const f0 = doc.child(0).attrs['field'];
    expect(f0).toMatchObject({ kind: 'toc', instr: INSTR });
    expect(doc.child(1).attrs['field']).toBe(f0);
    expect(doc.child(2).attrs['field']).toBeNull();
    const dirty = await documentXml(await exportDocx(tocDoc(true)));
    expect(dirty).toContain(
      '<w:fldChar w:fldCharType="begin" w:dirty="true"/>',
    );
  });
});

describe('a one-entry table of contents', () => {
  it('survives a round trip: begin and end in one paragraph is still the TOC', async () => {
    const field = { kind: 'toc', instr: INSTR };
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { field }, [
        schema.text('Only\t1', [
          schema.marks['link'].create({ href: '#_Toc1' }),
        ]),
      ]),
      schema.node('paragraph', { heading: 1, bookmarks: ['_Toc1'] }, [
        schema.text('Only'),
      ]),
    ]);
    const bytes = await exportDocx(doc);
    const xml = await documentXml(bytes);
    const first = (xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [])[0];
    expect(first).toContain('fldCharType="begin"');
    expect(first).toContain('fldCharType="end"');
    const back = (await importDocx(bytes.slice().buffer as ArrayBuffer)).doc;
    expect(back.child(0).attrs['field']).toMatchObject({
      kind: 'toc',
      instr: INSTR,
    });
    expect(back.child(0).textContent).toBe('Only\t1');
    expect(back.child(1).attrs['field']).toBeNull();
  });
});
