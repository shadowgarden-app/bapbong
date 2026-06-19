import JSZip from 'jszip';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from './docx';
import { exportDocx } from './export';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** Build a schema doc from paragraphs of `{ text, marks }` runs. */
function makeDoc(paras: { text: string; marks?: string[]; attrs?: Record<string, unknown> }[][], pAttrs: Record<string, unknown>[] = []) {
  const ps = paras.map((runs, i) =>
    schema.node(
      'paragraph',
      pAttrs[i] ?? null,
      runs.map((r) => schema.text(r.text, (r.marks ?? []).map((m) => markFor(m, r.attrs)))),
    ),
  );
  return schema.node('doc', null, ps);
}

function markFor(name: string, attrs?: Record<string, unknown>) {
  return schema.marks[name].create(attrs);
}

describe('exportDocx (round-trip)', () => {
  it('round-trips paragraphs and common marks', async () => {
    const doc = makeDoc([
      [{ text: 'Hello ' }, { text: 'bold', marks: ['strong'] }, { text: ' and ' }, { text: 'italic', marks: ['em'] }],
      [{ text: 'plain second paragraph' }],
    ]);

    const bytes = await exportDocx(doc);
    const { doc: back } = await importDocx(bytes);

    expect(back.childCount).toBe(2);
    expect(back.child(0).textContent).toBe('Hello bold and italic');
    expect(back.child(1).textContent).toBe('plain second paragraph');
    const boldRun = [...range(back.child(0))].find((n) => n.text === 'bold');
    expect(boldRun?.marks.map((m) => m.type.name)).toContain('strong');
    const italicRun = [...range(back.child(0))].find((n) => n.text === 'italic');
    expect(italicRun?.marks.map((m) => m.type.name)).toContain('em');
  });

  it('round-trips color, size, vertAlign, paragraph alignment + page break', async () => {
    const doc = makeDoc(
      [
        [
          { text: 'red', marks: ['textColor'], attrs: { color: '#C0392B' } },
          { text: 'big', marks: ['fontSize'], attrs: { size: 18 } },
          { text: 'x', marks: ['vertAlign'], attrs: { value: 'super' } },
        ],
      ],
      [{ align: 'center', pageBreakBefore: true }],
    );

    const { doc: back } = await importDocx(await exportDocx(doc));
    const p0 = back.child(0);
    expect(p0.attrs['align']).toBe('center');
    expect(p0.attrs['pageBreakBefore']).toBe(true);
    const runs = [...range(p0)];
    expect(runs.find((n) => n.text === 'red')?.marks.find((m) => m.type.name === 'textColor')?.attrs['color']).toBe('#C0392B');
    expect(runs.find((n) => n.text === 'big')?.marks.find((m) => m.type.name === 'fontSize')?.attrs['size']).toBe(18);
    expect(runs.find((n) => n.text === 'x')?.marks.find((m) => m.type.name === 'vertAlign')?.attrs['value']).toBe('super');
  });

  it('round-trips a hard break within a paragraph', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('a'), schema.node('hard_break'), schema.text('b')]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const p = back.child(0);
    expect(p.childCount).toBe(3);
    expect(p.child(1).type.name).toBe('hard_break');
  });
});

/** Iterate a block's inline children. */
function* range(block: import('prosemirror-model').Node) {
  for (let i = 0; i < block.childCount; i++) yield block.child(i);
}

// 1×1 transparent PNG.
const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

describe('exportDocx (E2: lists / tables / images / hyperlinks)', () => {
  it('round-trips a list paragraph (numId + level)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { list: { numId: '3', level: 1 } }, [schema.text('an item')]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    expect(back.child(0).attrs['list']).toEqual({ numId: '3', level: 1 });
    expect(back.child(0).textContent).toBe('an item');
  });

  it('round-trips a table with borders + colspan', async () => {
    const cell = (text: string, attrs?: Record<string, unknown>) =>
      schema.node('table_cell', attrs ?? null, [schema.node('paragraph', null, [schema.text(text)])]);
    const doc = schema.node('doc', null, [
      schema.node('table', { borders: { top: true, bottom: true, left: true, right: true, insideH: true, insideV: true } }, [
        schema.node('table_row', null, [cell('A1', { colwidth: [100] }), cell('B1', { colwidth: [120] })]),
        schema.node('table_row', null, [cell('wide', { colspan: 2, colwidth: [100, 120] })]),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const tbl = back.child(0);
    expect(tbl.type.name).toBe('table');
    expect(tbl.attrs['borders']).toBeTruthy();
    expect(tbl.childCount).toBe(2); // two rows
    expect(tbl.child(0).childCount).toBe(2); // two cells
    expect(tbl.child(0).child(0).textContent).toBe('A1');
    expect(tbl.child(1).child(0).attrs['colspan']).toBe(2);
    expect(tbl.child(1).child(0).textContent).toBe('wide');
  });

  it('round-trips an inline image (src + dimensions)', async () => {
    const src = `data:image/png;base64,${PNG_1PX}`;
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [schema.node('image', { src, width: 50, height: 40 })]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const img = [...range(back.child(0))].find((n) => n.type.name === 'image');
    expect(img).toBeTruthy();
    expect(img?.attrs['width']).toBe(50);
    expect(img?.attrs['height']).toBe(40);
    expect(String(img?.attrs['src'])).toBe(src);
  });

  it('round-trips a hyperlink (link mark + href)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('see '),
        schema.text('the site', [schema.marks['link'].create({ href: 'https://prosemirror.net/' })]),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const linked = [...range(back.child(0))].find((n) => n.text === 'the site');
    expect(linked?.marks.find((m) => m.type.name === 'link')?.attrs['href']).toBe('https://prosemirror.net/');
  });
});

describe('exportDocx (E3: comments round-trip)', () => {
  const body = (text: string) => ({ type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text }] }] });
  // Plain text from a commentSchema body JSON (no DOM/model import needed).
  const bodyText = (b: unknown): string =>
    ((b as { content?: { content?: { text?: string }[] }[] } | undefined)?.content ?? [])
      .map((p) => (p.content ?? []).map((r) => r.text ?? '').join(''))
      .join('\n');

  it('round-trips comment ranges + thread + resolved + a mention in the body', async () => {
    const mark = (id: number) => schema.marks['comment'].create({ ids: [id] });
    const comments = [
      { id: 1, parentId: null, user: { id: 'a', name: 'Alice Nguyễn' }, date: '2026-06-17T09:00:00Z', body: body('root note'), resolved: false },
      { id: 2, parentId: 1, user: { id: 'b', name: 'Bob' }, date: '2026-06-17T09:05:00Z', body: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'cảm ơn ' }, { type: 'mention', attrs: { id: 'x', label: 'Xuân' } }] }] }, resolved: false },
      { id: 3, parentId: null, user: { id: 'c', name: 'Carol Lee' }, date: '2026-06-17T09:10:00Z', body: body('done thread'), resolved: true },
    ];
    const doc = schema.node('doc', { comments }, [
      schema.node('paragraph', null, [schema.text('before '), schema.text('commented', [mark(1)]), schema.text(' after')]),
      schema.node('paragraph', null, [schema.text('a ', [mark(3)]), schema.text('span', [mark(3)])]),
    ]);

    const { doc: back } = await importDocx(await exportDocx(doc));

    // thread + resolved survive
    const nodes = back.attrs['comments'] as { id: number; parentId: number | null; resolved: boolean }[];
    const byId = new Map(nodes.map((n) => [n.id, n]));
    expect(nodes).toHaveLength(3);
    expect(byId.get(1)).toMatchObject({ parentId: null, resolved: false });
    expect(byId.get(2)).toMatchObject({ parentId: 1, resolved: false });
    expect(byId.get(3)).toMatchObject({ parentId: null, resolved: true });

    // the range marks land on the right runs (and only those)
    const p0 = [...range(back.child(0))];
    expect(p0.find((n) => n.text === 'commented')?.marks.find((m) => m.type.name === 'comment')?.attrs['ids']).toEqual([1]);
    expect(p0.find((n) => n.text === 'before ')?.marks.some((m) => m.type.name === 'comment')).toBeFalsy();
    expect(back.child(1).textContent).toBe('a span'); // both runs carry comment 3

    // mention in the reply body serialized as "@Xuân" text
    const reply = (back.attrs['comments'] as { id: number; body: unknown }[]).find((c) => c.id === 2);
    expect(bodyText(reply?.body)).toContain('@Xuân');
  });
});

describe('exportDocx (E4: carry original parts)', () => {
  // Minimal source .docx with a numbered list paragraph + numbering.xml.
  async function sourceBytes(): Promise<Uint8Array> {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
    );
    zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file(
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item one</w:t></w:r></w:p></w:body></w:document>`,
    );
    zip.file(
      'word/numbering.xml',
      `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    );
    return zip.generateAsync({ type: 'uint8array' });
  }

  it('preserves numbering.xml with carry (and drops it without)', async () => {
    const { doc, raw } = await importDocx(await sourceBytes());
    expect(doc.attrs['numbering']).toBeTruthy(); // imported defs
    expect(doc.child(0).attrs['list']).toEqual({ numId: '1', level: 0 });

    // With carry: numbering.xml survives the round-trip.
    const withCarry = await importDocx(await exportDocx(doc, { carry: raw }));
    expect(withCarry.doc.attrs['numbering']).toBeTruthy();
    expect(withCarry.doc.child(0).attrs['list']).toEqual({ numId: '1', level: 0 });

    // Without carry: the list paragraph still round-trips, but the defs are gone.
    const fromScratch = await importDocx(await exportDocx(doc));
    expect(fromScratch.doc.child(0).attrs['list']).toEqual({ numId: '1', level: 0 });
    expect(fromScratch.doc.attrs['numbering']).toBeNull();
  });
});

describe('exportDocx (E4 fidelity: sectPr + header refs)', () => {
  async function sourceWithHeader(): Promise<Uint8Array> {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
    );
    zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="${PR_NS}"><Relationship Id="rIdH" Type="${R_NS}/header" Target="header1.xml"/></Relationships>`);
    zip.file('word/header1.xml', `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>MY HEADER</w:t></w:r></w:p></w:hdr>`);
    zip.file(
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body><w:p><w:r><w:t>body text</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
    );
    return zip.generateAsync({ type: 'uint8array' });
  }

  it('re-attaches sectPr so headers + page geometry survive (carry)', async () => {
    const { doc, headers, page, raw } = await importDocx(await sourceWithHeader());
    expect(headers['default']?.textContent).toBe('MY HEADER');
    expect(page.width).toBe(Math.round(11906 / 15));

    const back = await importDocx(await exportDocx(doc, { carry: raw }));
    expect(back.headers['default']?.textContent).toBe('MY HEADER'); // ref resolved
    expect(back.page.width).toBe(Math.round(11906 / 15)); // pgSz preserved
    expect(back.doc.child(0).textContent).toBe('body text');

    // Without carry there's no sectPr → no header.
    const noCarry = await importDocx(await exportDocx(doc));
    expect(Object.keys(noCarry.headers)).toHaveLength(0);
  });
});

describe('exportDocx (sections + footnotes)', () => {
  // The exported word/document.xml, for inspecting serialization directly.
  async function docXml(doc: import('prosemirror-model').Node, opts?: Parameters<typeof exportDocx>[1]): Promise<string> {
    const zip = await JSZip.loadAsync(await exportDocx(doc, opts));
    return (await zip.file('word/document.xml')?.async('string')) ?? '';
  }

  it('serializes a section break (w:cols + w:type) into the first section', async () => {
    const doc = schema.node(
      'doc',
      {
        sections: [
          { blockCount: 1, columns: { count: 2, gap: 24 }, newPage: true },
          { blockCount: 1, columns: { count: 1, gap: 36 }, newPage: false },
        ],
      },
      [
        schema.node('paragraph', null, [schema.text('section one')]),
        schema.node('paragraph', null, [schema.text('section two')]),
      ],
    );
    const xml = await docXml(doc);
    // Section 0's break sits in the first paragraph's pPr.
    expect(xml).toContain('<w:sectPr><w:type w:val="nextPage"/><w:cols w:num="2" w:space="360"/></w:sectPr>');
    // Only the non-last section is serialized inline (last → body sectPr).
    expect(xml.match(/<w:sectPr>/g)?.length).toBe(1);
  });

  it('serializes a footnote reference for a footnote-marked run', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('text'), schema.text('1', [schema.marks['footnote'].create({ num: 2 })])]),
    ]);
    const xml = await docXml(doc);
    expect(xml).toContain('<w:footnoteReference w:id="2"/>');
    expect(xml).not.toContain('>1</w:t>'); // the carrier number isn't emitted as text
  });
});
