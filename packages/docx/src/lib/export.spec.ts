import JSZip from 'jszip';
import { Schema } from 'prosemirror-model';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from './docx';
import { exportDocx } from './export';

// The comment mark now lives in the comment plugin, not the base schema. Tests
// that round-trip comments compose a schema carrying it (a minimal local spec,
// to avoid a docx→plugin dependency) and import against it.
const withComments = new Schema({
  nodes: schema.spec.nodes,
  marks: schema.spec.marks.append({ comment: { attrs: { ids: {} } } }),
});

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** Build a schema doc from paragraphs of `{ text, marks }` runs. */
function makeDoc(
  paras: {
    text: string;
    marks?: string[];
    attrs?: Record<string, unknown>;
  }[][],
  pAttrs: Record<string, unknown>[] = [],
) {
  const ps = paras.map((runs, i) =>
    schema.node(
      'paragraph',
      pAttrs[i] ?? null,
      runs.map((r) =>
        schema.text(
          r.text,
          (r.marks ?? []).map((m) => markFor(m, r.attrs)),
        ),
      ),
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
      [
        { text: 'Hello ' },
        { text: 'bold', marks: ['strong'] },
        { text: ' and ' },
        { text: 'italic', marks: ['em'] },
      ],
      [{ text: 'plain second paragraph' }],
    ]);

    const bytes = await exportDocx(doc);
    const { doc: back } = await importDocx(bytes);

    expect(back.childCount).toBe(2);
    expect(back.child(0).textContent).toBe('Hello bold and italic');
    expect(back.child(1).textContent).toBe('plain second paragraph');
    const boldRun = [...range(back.child(0))].find((n) => n.text === 'bold');
    expect(boldRun?.marks.map((m) => m.type.name)).toContain('strong');
    const italicRun = [...range(back.child(0))].find(
      (n) => n.text === 'italic',
    );
    expect(italicRun?.marks.map((m) => m.type.name)).toContain('em');
  });

  it('round-trips smallCaps and dstrike marks', async () => {
    const doc = makeDoc([
      [
        { text: 'Heading', marks: ['smallCaps'] },
        { text: 'removed', marks: ['dstrike'] },
      ],
    ]);

    const bytes = await exportDocx(doc);
    const { doc: back } = await importDocx(bytes);

    const sc = [...range(back.child(0))].find((n) => n.text === 'Heading');
    expect(sc?.marks.map((m) => m.type.name)).toContain('smallCaps');
    const ds = [...range(back.child(0))].find((n) => n.text === 'removed');
    expect(ds?.marks.map((m) => m.type.name)).toContain('dstrike');
  });

  it('carry-through: unmodelled rPr/pPr props survive import → export', async () => {
    // A source docx with properties the model does NOT represent: run-level
    // w:rtl/w:kern/w:szCs, paragraph-level w:contextualSpacing/w:keepNext and
    // a paragraph-mark w:rPr. Saving must not drop any of them.
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    );
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0"?><Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    );
    zip.file(
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
        <w:p>
          <w:pPr>
            <w:keepNext/><w:contextualSpacing/><w:jc w:val="center"/>
            <w:rPr><w:b/><w:sz w:val="16"/></w:rPr>
            <w:pPrChange w:id="1" w:author="x"><w:pPr/></w:pPrChange>
          </w:pPr>
          <w:r>
            <w:rPr><w:b/><w:rtl/><w:kern w:val="28"/><w:szCs w:val="30"/><w14:glow w14:rad="1"/></w:rPr>
            <w:t>ab&lt;c</w:t>
          </w:r>
        </w:p>
      </w:body></w:document>`,
    );
    const source = await zip.generateAsync({ type: 'uint8array' });

    const { doc, raw } = await importDocx(source);
    const run = doc.child(0).child(0);
    const carryMark = run.marks.find((m) => m.type.name === 'carryRPr');
    expect(carryMark?.attrs['xml']).toBe(
      '<w:rtl/><w:kern w:val="28"/><w:szCs w:val="30"/>',
    );
    // keepNext is MODELLED now (an attr, re-emitted by the exporter) — only
    // contextualSpacing still rides the carry.
    expect(doc.child(0).attrs['keepNext']).toBe(true);
    expect(doc.child(0).attrs['carry']).toEqual({
      pPr: '<w:contextualSpacing/>',
      markRPr: '<w:b/><w:sz w:val="16"/>',
    });

    const outBytes = await exportDocx(doc, { carry: raw });
    const outZip = await JSZip.loadAsync(outBytes);
    const xml = (await outZip.file('word/document.xml')?.async('string')) ?? '';
    // Run: modelled bold once, carried extras present, foreign ns dropped.
    expect(xml).toContain('<w:rtl/><w:kern w:val="28"/><w:szCs w:val="30"/>');
    expect(xml.match(/<w:b\/>/g)?.length).toBe(2); // run rPr + mark rPr — no dupes
    expect(xml).not.toContain('w14:glow');
    // Paragraph: modelled keepNext re-emitted (schema-first in pPr), the
    // carried extra + the paragraph-mark rPr; revision record dropped.
    expect(xml).toContain('<w:keepNext/>');
    expect(xml).toContain('<w:contextualSpacing/>');
    expect(xml).toContain('<w:rPr><w:b/><w:sz w:val="16"/></w:rPr>');
    expect(xml).not.toContain('w:pPrChange');
    expect(xml).toContain('ab&lt;c'); // escaping intact through the trip

    // Second trip: everything carried again, byte-identical fragments.
    const { doc: doc2 } = await importDocx(outBytes);
    const run2 = doc2.child(0).child(0);
    expect(
      run2.marks.find((m) => m.type.name === 'carryRPr')?.attrs['xml'],
    ).toBe('<w:rtl/><w:kern w:val="28"/><w:szCs w:val="30"/>');
    expect(doc2.child(0).attrs['carry']).toEqual(doc.child(0).attrs['carry']);
  });

  it('carry-through: unmodelled table props survive import → export', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    );
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0"?><Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    );
    zip.file(
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>
        <w:tbl>
          <w:tblPr>
            <w:tblStyle w:val="TableGrid"/>
            <w:tblW w:w="5000" w:type="dxa"/>
            <w:jc w:val="center"/>
            <w:tblInd w:w="720" w:type="dxa"/>
            <w:tblBorders><w:top w:val="single" w:sz="8" w:space="4" w:color="FF0000"/></w:tblBorders>
            <w:tblLayout w:type="fixed"/>
            <w:tblLook w:val="04A0" w:firstRow="1"/>
            <w:tblPrChange w:id="9" w:author="x"><w:tblPr/></w:tblPrChange>
          </w:tblPr>
          <w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
          <w:tr>
            <w:trPr><w:gridBefore w:val="1"/><w:wBefore w:w="500" w:type="dxa"/><w:tblHeader/></w:trPr>
            <w:tc>
              <w:tcPr><w:tcW w:w="2500" w:type="dxa"/><w:textDirection w:val="btLr"/><w:noWrap/></w:tcPr>
              <w:p><w:r><w:t>A</w:t></w:r></w:p>
            </w:tc>
            <w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
      </w:body></w:document>`,
    );
    const source = await zip.generateAsync({ type: 'uint8array' });

    const { doc, raw } = await importDocx(source);
    const tbl = doc.child(0);
    expect((tbl.attrs['carry'] as { tblPr: string }).tblPr).toBe(
      '<w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="dxa"/><w:tblInd w:w="720" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1"/>',
    );
    const row = tbl.child(0);
    expect((row.attrs['carry'] as { trPr: string }).trPr).toBe(
      '<w:gridBefore w:val="1"/><w:wBefore w:w="500" w:type="dxa"/>',
    );
    expect(row.attrs['header']).toBe(true); // modelled prop still modelled
    expect((row.child(0).attrs['carry'] as { tcPr: string }).tcPr).toBe(
      '<w:textDirection w:val="btLr"/><w:noWrap/>',
    );
    expect(row.child(1).attrs['carry']).toBeNull();

    const out = await exportDocx(doc, { carry: raw });
    const xml =
      (await (await JSZip.loadAsync(out))
        .file('word/document.xml')
        ?.async('string')) ?? '';
    // tblPr: carried extras first (tblStyle at the head), modelled after, no dupes.
    expect(xml).toContain('<w:tblPr><w:tblStyle w:val="TableGrid"/>');
    expect(xml).toContain('<w:tblLayout w:type="fixed"/>');
    expect(xml).toContain('<w:tblLook w:val="04A0" w:firstRow="1"/>');
    expect(xml.match(/<w:jc w:val="center"\/>/g)?.length).toBe(1);
    expect(xml).not.toContain('w:tblPrChange'); // revision record dropped
    // Border w:space round-trips (4pt → px → 4pt).
    expect(xml).toMatch(
      /<w:top w:val="single" w:sz="\d+" w:space="4" w:color="FF0000"\/>/,
    );
    // trPr/tcPr extras.
    expect(xml).toContain(
      '<w:gridBefore w:val="1"/><w:wBefore w:w="500" w:type="dxa"/>',
    );
    expect(xml).toContain('<w:tblHeader/>');
    expect(xml).toContain('<w:textDirection w:val="btLr"/><w:noWrap/>');
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
    expect(
      runs
        .find((n) => n.text === 'red')
        ?.marks.find((m) => m.type.name === 'textColor')?.attrs['color'],
    ).toBe('#C0392B');
    expect(
      runs
        .find((n) => n.text === 'big')
        ?.marks.find((m) => m.type.name === 'fontSize')?.attrs['size'],
    ).toBe(18);
    expect(
      runs
        .find((n) => n.text === 'x')
        ?.marks.find((m) => m.type.name === 'vertAlign')?.attrs['value'],
    ).toBe('super');
  });

  it('round-trips w:position (baseline shift)', async () => {
    const doc = makeDoc([
      [
        { text: 'down', marks: ['position'], attrs: { halfPoints: -2 } },
        { text: 'up', marks: ['position'], attrs: { halfPoints: 6 } },
        { text: 'flat' },
      ],
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const runs = [...range(back.child(0))];
    const raise = (t: string) =>
      runs
        .find((n) => n.text === t)
        ?.marks.find((m) => m.type.name === 'position')?.attrs['halfPoints'];
    expect(raise('down')).toBe(-2);
    expect(raise('up')).toBe(6);
    expect(raise('flat')).toBeUndefined();
  });

  it('round-trips rPr w:spacing (tracking)', async () => {
    const doc = makeDoc([
      [
        { text: 'wide', marks: ['letterSpacing'], attrs: { twips: 26 } },
        { text: 'tight', marks: ['letterSpacing'], attrs: { twips: -8 } },
        { text: 'plain' },
      ],
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const runs = [...range(back.child(0))];
    const track = (t: string) =>
      runs
        .find((n) => n.text === t)
        ?.marks.find((m) => m.type.name === 'letterSpacing')?.attrs['twips'];
    expect(track('wide')).toBe(26);
    expect(track('tight')).toBe(-8);
    expect(track('plain')).toBeUndefined();
  });

  it('round-trips a hard break within a paragraph', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('a'),
        schema.node('hard_break'),
        schema.text('b'),
      ]),
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
      schema.node('paragraph', { list: { numId: '3', level: 1 } }, [
        schema.text('an item'),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    expect(back.child(0).attrs['list']).toEqual({ numId: '3', level: 1 });
    expect(back.child(0).textContent).toBe('an item');
  });

  it('round-trips a heading level (w:pStyle HeadingN)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { heading: 2 }, [schema.text('My Heading')]),
      schema.node('paragraph', null, [schema.text('body text')]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    expect(back.child(0).attrs['heading']).toBe(2);
    expect(back.child(0).textContent).toBe('My Heading');
    expect(back.child(1).attrs['heading']).toBeNull();
  });

  it('round-trips a table with borders + colspan', async () => {
    const cell = (text: string, attrs?: Record<string, unknown>) =>
      schema.node('table_cell', attrs ?? null, [
        schema.node('paragraph', null, [schema.text(text)]),
      ]);
    const s = { width: 1.5, style: 'solid' as const, color: '#000000' };
    const doc = schema.node('doc', null, [
      schema.node(
        'table',
        {
          borders: {
            top: s,
            bottom: s,
            left: s,
            right: s,
            insideH: s,
            insideV: s,
          },
        },
        [
          schema.node('table_row', null, [
            cell('A1', { colwidth: [100] }),
            cell('B1', { colwidth: [120] }),
          ]),
          schema.node('table_row', null, [
            cell('wide', { colspan: 2, colwidth: [100, 120] }),
          ]),
        ],
      ),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const tbl = back.child(0);
    expect(tbl.type.name).toBe('table');
    // width/style/colour survive the round-trip (1.5px → w:sz 9 → 1.5px).
    expect(tbl.attrs['borders'].top).toEqual({
      width: 1.5,
      style: 'solid',
      color: '#000000',
    });
    expect(tbl.childCount).toBe(2); // two rows
    expect(tbl.child(0).childCount).toBe(2); // two cells
    expect(tbl.child(0).child(0).textContent).toBe('A1');
    expect(tbl.child(1).child(0).attrs['colspan']).toBe(2);
    expect(tbl.child(1).child(0).textContent).toBe('wide');
  });

  it('round-trips behindDoc (float.behind) both ways', async () => {
    const img = (behind: boolean) =>
      schema.node('paragraph', null, [
        schema.node('image', {
          src: `data:image/png;base64,${PNG_1PX}`,
          width: 10,
          height: 10,
          float: {
            wrap: 'none',
            hOffset: 50,
            vOffset: 60,
            ...(behind ? { behind: true } : {}),
          },
        }),
      ]);
    const doc = schema.node('doc', null, [img(true), img(false)]);
    const bytes = await exportDocx(doc);
    const xml = await (await JSZip.loadAsync(bytes))
      .file('word/document.xml')!
      .async('string');
    // Hardcoding behindDoc="0" was silently pulling behind-text images in
    // front of the text on every save.
    expect((xml.match(/behindDoc="1"/g) ?? []).length).toBe(1);
    expect((xml.match(/behindDoc="0"/g) ?? []).length).toBe(1);

    const { doc: back } = await importDocx(bytes);
    const floats: unknown[] = [];
    back.descendants((n) => {
      if (n.type.name === 'image') floats.push(n.attrs['float']);
    });
    expect(floats[0]).toMatchObject({ wrap: 'none', behind: true });
    expect((floats[1] as Record<string, unknown>)['behind']).toBeUndefined();
  });

  it('round-trips custom tab stops (w:tabs)', async () => {
    const tabs = [
      { pos: 200, val: 'right', leader: 'dot' },
      { pos: 400, val: 'decimal' },
    ];
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { tabs }, [schema.text('x')]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    expect(back.child(0).attrs['tabs']).toEqual(tabs);
  });

  it('round-trips a PAGE field in the body (w:fldSimple)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.node('page_field', { kind: 'page' }),
      ]),
      schema.node('paragraph', null, [
        schema.node('page_field', { kind: 'pages' }),
      ]),
    ]);
    const bytes = await exportDocx(doc);
    const xml = await (await JSZip.loadAsync(bytes))
      .file('word/document.xml')!
      .async('string');
    expect(xml).toContain('w:instr=" PAGE "');
    expect(xml).toContain('w:instr=" NUMPAGES "');

    const { doc: back } = await importDocx(bytes);
    const kinds: string[] = [];
    back.descendants((n) => {
      if (n.type.name === 'page_field') kinds.push(n.attrs['kind'] as string);
    });
    expect(kinds).toEqual(['page', 'pages']);
  });

  it('round-trips image alt text (wp:docPr@descr)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.node('image', {
          src: `data:image/png;base64,${PNG_1PX}`,
          width: 10,
          height: 10,
          alt: 'a duck & a "goose"',
        }),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    let alt: unknown = null;
    back.descendants((n) => {
      if (n.type.name === 'image') alt = n.attrs['alt'];
    });
    expect(alt).toBe('a duck & a "goose"');
  });

  it('round-trips a vertical merge (rowspan → w:vMerge)', async () => {
    const cell = (text: string, attrs?: Record<string, unknown>) =>
      schema.node('table_cell', attrs ?? null, [
        schema.node('paragraph', null, [schema.text(text)]),
      ]);
    // Column 1 spans all three rows, so rows 2 and 3 hold ONE node each — the
    // covered slot is absorbed into the rowspan above and has no node at all.
    const doc = schema.node('doc', null, [
      schema.node('table', null, [
        schema.node('table_row', null, [
          cell('tall', { rowspan: 3, colwidth: [80] }),
          cell('B1', { colwidth: [120] }),
        ]),
        schema.node('table_row', null, [cell('B2', { colwidth: [120] })]),
        schema.node('table_row', null, [cell('B3', { colwidth: [120] })]),
      ]),
    ]);

    const bytes = await exportDocx(doc);
    const xml = await (await JSZip.loadAsync(bytes))
      .file('word/document.xml')!
      .async('string');
    // OOXML keeps a real w:tc in every covered row; dropping them is what left
    // rows shorter than w:tblGrid and made Word render the table ragged.
    expect((xml.match(/<w:tc>/g) ?? []).length).toBe(6);
    expect((xml.match(/<w:vMerge w:val="restart"\/>/g) ?? []).length).toBe(1);
    expect((xml.match(/<w:vMerge\/>/g) ?? []).length).toBe(2);

    const { doc: back } = await importDocx(bytes);
    const tbl = back.child(0);
    expect(tbl.childCount).toBe(3);
    expect(tbl.child(0).child(0).attrs['rowspan']).toBe(3);
    expect(tbl.child(0).child(0).textContent).toBe('tall');
    // The covered rows come back with one cell each, as they went out.
    expect(tbl.child(1).childCount).toBe(1);
    expect(tbl.child(1).child(0).textContent).toBe('B2');
    expect(tbl.child(2).child(0).textContent).toBe('B3');
  });

  it('round-trips an inline image (src + dimensions)', async () => {
    const src = `data:image/png;base64,${PNG_1PX}`;
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.node('image', { src, width: 50, height: 40 }),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const img = [...range(back.child(0))].find((n) => n.type.name === 'image');
    expect(img).toBeTruthy();
    expect(img?.attrs['width']).toBe(50);
    expect(img?.attrs['height']).toBe(40);
    expect(String(img?.attrs['src'])).toBe(src);
  });

  it('embeds a webp data-URL image with the right extension + content type', async () => {
    // Regression: unmapped MIME types used to be written as image<n>.png —
    // foreign bytes behind a .png label, which Word renders as a missing box.
    const src = `data:image/webp;base64,${PNG_1PX}`;
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.node('image', { src, width: 50, height: 40 }),
      ]),
    ]);
    const bytes = await exportDocx(doc);
    const zip = await JSZip.loadAsync(bytes.slice().buffer);
    expect(zip.file(/word\/media\/image\d+\.webp/)).toHaveLength(1);
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('Extension="webp" ContentType="image/webp"');
    const { doc: back } = await importDocx(bytes);
    const img = [...range(back.child(0))].find((n) => n.type.name === 'image');
    expect(String(img?.attrs['src'])).toBe(src);
  });

  it('round-trips an http-src image as an externally-linked picture', async () => {
    // Regression: non-data-URL srcs used to be silently dropped on export.
    const src = 'https://example.com/pic.png';
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.node('image', { src, width: 50, height: 40 }),
      ]),
    ]);
    const bytes = await exportDocx(doc);
    const zip = await JSZip.loadAsync(bytes.slice().buffer);
    const rels = await zip
      .file('word/_rels/document.xml.rels')!
      .async('string');
    expect(rels).toContain(`Target="${src}" TargetMode="External"`);
    const docXml = await zip.file('word/document.xml')!.async('string');
    expect(docXml).toContain('r:link=');
    const { doc: back } = await importDocx(bytes);
    const img = [...range(back.child(0))].find((n) => n.type.name === 'image');
    expect(String(img?.attrs['src'])).toBe(src);
    expect(img?.attrs['width']).toBe(50);
  });

  it('round-trips paragraph borders (w:pBdr) and cell margins (w:tcMar)', async () => {
    const side = { width: 1.5, style: 'solid' as const, color: '#CCCCCC' };
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { borders: { top: side } }, [
        schema.text('ruled'),
      ]),
      schema.node('table', null, [
        schema.node('table_row', null, [
          schema.node('table_cell', { padding: { top: 4, left: 5 } }, [
            schema.node('paragraph', null, [schema.text('c')]),
          ]),
        ]),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    expect(back.child(0).attrs['borders']).toEqual({ top: side });
    expect(back.child(1).child(0).child(0).attrs['padding']).toEqual({
      top: 4,
      left: 5,
    });
  });

  it('round-trips w:cantSplit on table rows', async () => {
    const doc = schema.node('doc', null, [
      schema.node('table', null, [
        schema.node('table_row', { cantSplit: true }, [
          schema.node('table_cell', null, [
            schema.node('paragraph', null, [schema.text('keep me whole')]),
          ]),
        ]),
        schema.node('table_row', null, [
          schema.node('table_cell', null, [
            schema.node('paragraph', null, [schema.text('free to split')]),
          ]),
        ]),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const table = back.child(0);
    expect(table.child(0).attrs['cantSplit']).toBe(true);
    expect(table.child(1).attrs['cantSplit']).toBe(false);
  });

  it('round-trips drawn shapes (anchored rect + inline line)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.node('image', {
          src: '',
          width: 18,
          height: 16,
          float: {
            wrap: 'square',
            hOffset: 319,
            hRel: 'margin',
            vOffset: -7,
            vRel: 'paragraph',
          },
          shape: { kind: 'rect', stroke: '#4472C4', strokeWidth: 2 },
        }),
        schema.node('image', {
          src: '',
          width: 100,
          height: 0,
          shape: {
            kind: 'line',
            stroke: '#C45911',
            strokeWidth: 1,
            flipV: true,
          },
        }),
        schema.node('image', {
          src: '',
          width: 60,
          height: 40,
          shape: {
            kind: 'ellipse',
            stroke: '#000000',
            strokeWidth: 1,
            fill: '#FFEE00',
          },
          rotation: 15.5,
        }),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const shapes = [...range(back.child(0))].filter(
      (n) => n.type.name === 'image',
    );
    expect(shapes).toHaveLength(3);
    expect(shapes[0].attrs['shape']).toEqual({
      kind: 'rect',
      stroke: '#4472C4',
      strokeWidth: 2,
    });
    expect(shapes[0].attrs['width']).toBe(18);
    expect(shapes[0].attrs['float']).toMatchObject({
      wrap: 'square',
      hOffset: 319,
      vOffset: -7,
      vRel: 'paragraph',
    });
    expect(shapes[1].attrs['shape']).toEqual({
      kind: 'line',
      stroke: '#C45911',
      strokeWidth: 1,
      flipV: true,
    });
    expect(shapes[1].attrs['float']).toBeNull();
    // Preset geometry beyond rect/line keeps its kind through the round-trip.
    expect(shapes[2].attrs['shape']).toEqual({
      kind: 'ellipse',
      stroke: '#000000',
      strokeWidth: 1,
      fill: '#FFEE00',
    });
    // Rotation round-trips via a:xfrm@rot (1/60000 deg).
    expect(shapes[2].attrs['rotation']).toBeCloseTo(15.5);
    expect(shapes[0].attrs['rotation']).toBe(0);
  });

  it('round-trips textbox text on a drawn shape', async () => {
    const boxParas = [
      schema.node('paragraph', null, [
        schema.text('Phiếu học tập: ', [schema.marks['strong'].create()]),
        schema.text('Học sinh trả lời.'),
      ]),
      schema.node('paragraph', null, [schema.text('Câu 1')]),
    ].map((p) => p.toJSON());
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.node('image', {
          src: '',
          width: 643,
          height: 274,
          float: {
            wrap: 'none',
            hOffset: 7,
            hRel: 'margin',
            vOffset: 14,
            vRel: 'paragraph',
          },
          shape: {
            kind: 'rect',
            stroke: '#000000',
            strokeWidth: 1,
            fill: '#FFFFFF',
          },
          textbox: {
            paragraphs: boxParas,
            inset: { l: 19, t: 10, r: 10, b: 5 },
          },
        }),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const node = [...range(back.child(0))].find((n) => n.type.name === 'image');
    const tb = node?.attrs['textbox'] as {
      paragraphs: unknown[];
      inset?: unknown;
    };
    expect(tb).toBeTruthy();
    expect(tb.paragraphs).toHaveLength(2);
    const p0 = schema.nodeFromJSON(tb.paragraphs[0] as never);
    expect(p0.textContent).toBe('Phiếu học tập: Học sinh trả lời.');
    expect(p0.child(0).marks.map((m) => m.type.name)).toContain('strong');
    expect(tb.inset).toEqual({ l: 19, t: 10, r: 10, b: 5 });
  });

  it('round-trips a hyperlink (link mark + href)', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('see '),
        schema.text('the site', [
          schema.marks['link'].create({ href: 'https://prosemirror.net/' }),
        ]),
      ]),
    ]);
    const { doc: back } = await importDocx(await exportDocx(doc));
    const linked = [...range(back.child(0))].find((n) => n.text === 'the site');
    expect(
      linked?.marks.find((m) => m.type.name === 'link')?.attrs['href'],
    ).toBe('https://prosemirror.net/');
  });
});

describe('exportDocx (E3: comments round-trip)', () => {
  const body = (text: string) => ({
    type: 'doc',
    content: [{ type: 'paragraph', content: [{ type: 'text', text }] }],
  });
  // Plain text from a commentSchema body JSON (no DOM/model import needed).
  const bodyText = (b: unknown): string =>
    (
      (b as { content?: { content?: { text?: string }[] }[] } | undefined)
        ?.content ?? []
    )
      .map((p) => (p.content ?? []).map((r) => r.text ?? '').join(''))
      .join('\n');

  it('round-trips comment ranges + thread + resolved + a mention in the body', async () => {
    const mark = (id: number) =>
      withComments.marks['comment'].create({ ids: [id] });
    const comments = [
      {
        id: 1,
        parentId: null,
        user: { id: 'a', name: 'Alice Nguyễn' },
        date: '2026-06-17T09:00:00Z',
        body: body('root note'),
        resolved: false,
      },
      {
        id: 2,
        parentId: 1,
        user: { id: 'b', name: 'Bob' },
        date: '2026-06-17T09:05:00Z',
        body: {
          type: 'doc',
          content: [
            {
              type: 'paragraph',
              content: [
                { type: 'text', text: 'cảm ơn ' },
                { type: 'mention', attrs: { id: 'x', label: 'Xuân' } },
              ],
            },
          ],
        },
        resolved: false,
      },
      {
        id: 3,
        parentId: null,
        user: { id: 'c', name: 'Carol Lee' },
        date: '2026-06-17T09:10:00Z',
        body: body('done thread'),
        resolved: true,
      },
    ];
    const doc = withComments.node('doc', { comments }, [
      withComments.node('paragraph', null, [
        withComments.text('before '),
        withComments.text('commented', [mark(1)]),
        withComments.text(' after'),
      ]),
      withComments.node('paragraph', null, [
        withComments.text('a ', [mark(3)]),
        withComments.text('span', [mark(3)]),
      ]),
    ]);

    const { doc: back } = await importDocx(await exportDocx(doc), {
      schema: withComments,
    });

    // thread + resolved survive
    const nodes = back.attrs['comments'] as {
      id: number;
      parentId: number | null;
      resolved: boolean;
    }[];
    const byId = new Map(nodes.map((n) => [n.id, n]));
    expect(nodes).toHaveLength(3);
    expect(byId.get(1)).toMatchObject({ parentId: null, resolved: false });
    expect(byId.get(2)).toMatchObject({ parentId: 1, resolved: false });
    expect(byId.get(3)).toMatchObject({ parentId: null, resolved: true });

    // the range marks land on the right runs (and only those)
    const p0 = [...range(back.child(0))];
    expect(
      p0
        .find((n) => n.text === 'commented')
        ?.marks.find((m) => m.type.name === 'comment')?.attrs['ids'],
    ).toEqual([1]);
    expect(
      p0
        .find((n) => n.text === 'before ')
        ?.marks.some((m) => m.type.name === 'comment'),
    ).toBeFalsy();
    expect(back.child(1).textContent).toBe('a span'); // both runs carry comment 3

    // mention in the reply body serialized as "@Xuân" text
    const reply = (
      back.attrs['comments'] as { id: number; body: unknown }[]
    ).find((c) => c.id === 2);
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
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    );
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
    expect(withCarry.doc.child(0).attrs['list']).toEqual({
      numId: '1',
      level: 0,
    });

    // Without carry: numbering.xml is regenerated from the doc-attr defs, so
    // the definition survives too (the numId may be re-minted).
    const fromScratch = await importDocx(await exportDocx(doc));
    const list = fromScratch.doc.child(0).attrs['list'] as { numId: string };
    const defs = fromScratch.doc.attrs['numbering'] as Record<
      string,
      { levels: Record<number, { lvlText: string }> }
    >;
    expect(defs[list.numId]?.levels[0]?.lvlText).toBe('%1.');
  });

  it('regenerates numbering.xml for editor-authored presets (integer numIds Word accepts)', async () => {
    const numbering = {
      'bb-ordered-paren': {
        key: 'bb-ordered-paren',
        levels: {
          0: { numFmt: 'decimal', lvlText: '%1)', start: 1 },
          1: { numFmt: 'lowerLetter', lvlText: '%2)', start: 1 },
          2: { numFmt: 'lowerRoman', lvlText: '%3)', start: 1 },
        },
      },
      'bb-bullet': {
        key: 'bb-bullet',
        levels: { 0: { numFmt: 'bullet', lvlText: '•', start: 1 } },
      },
    };
    const doc = schema.node('doc', { numbering }, [
      schema.node(
        'paragraph',
        { list: { numId: 'bb-ordered-paren', level: 0 } },
        [schema.text('first')],
      ),
      schema.node(
        'paragraph',
        { list: { numId: 'bb-ordered-paren', level: 1 } },
        [schema.text('nested')],
      ),
      schema.node('paragraph', { list: { numId: 'bb-bullet', level: 0 } }, [
        schema.text('bullet'),
      ]),
    ]);

    const bytes = await exportDocx(doc);
    const zip = await JSZip.loadAsync(bytes);
    const docXml = await zip.file('word/document.xml')!.async('string');
    // No bb-* id may leak into the document — Word requires integers.
    expect(docXml).not.toContain('bb-');
    const numIds = [...docXml.matchAll(/<w:numId w:val="([^"]+)"/g)].map(
      (m) => m[1],
    );
    expect(numIds).toHaveLength(3);
    for (const id of numIds) expect(id).toMatch(/^\d+$/);

    const numberingXml = await zip.file('word/numbering.xml')!.async('string');
    expect(numberingXml).toContain('w:numFmt w:val="lowerLetter"');
    expect(numberingXml).toContain('w:lvlText w:val="•"');
    // abstractNum defs must all precede the first w:num (schema order).
    expect(numberingXml.lastIndexOf('<w:abstractNum ')).toBeLessThan(
      numberingXml.search(/<w:num[ >]/),
    );
    // Part is wired up: relationship + content-type override.
    const rels = await zip
      .file('word/_rels/document.xml.rels')!
      .async('string');
    expect(rels).toContain('Target="numbering.xml"');
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('/word/numbering.xml');

    // Full round-trip: markers and levels identical when reopened.
    const { doc: back } = await importDocx(bytes);
    const backDefs = back.attrs['numbering'] as Record<
      string,
      { levels: Record<number, { numFmt: string; lvlText: string }> }
    >;
    const l0 = back.child(0).attrs['list'] as { numId: string; level: number };
    const l1 = back.child(1).attrs['list'] as { numId: string; level: number };
    expect(l0.level).toBe(0);
    expect(l1.level).toBe(1);
    expect(l1.numId).toBe(l0.numId); // same list, same def
    expect(backDefs[l0.numId].levels[0]).toMatchObject({
      numFmt: 'decimal',
      lvlText: '%1)',
    });
    expect(backDefs[l0.numId].levels[1]).toMatchObject({
      numFmt: 'lowerLetter',
      lvlText: '%2)',
    });
    const bullet = back.child(2).attrs['list'] as { numId: string };
    expect(backDefs[bullet.numId].levels[0].lvlText).toBe('•');
  });

  it('merges editor-preset defs into a carried numbering.xml without touching original ids', async () => {
    const { doc, raw } = await importDocx(await sourceBytes());
    // User adds a bulleted paragraph in the editor: defs gain bb-bullet.
    const defs = {
      ...(doc.attrs['numbering'] as Record<string, unknown>),
      'bb-bullet': {
        key: 'bb-bullet',
        levels: { 0: { numFmt: 'bullet', lvlText: '•', start: 1 } },
      },
    };
    const doc2 = doc.type.create({ ...doc.attrs, numbering: defs }, [
      doc.child(0),
      schema.node('paragraph', { list: { numId: 'bb-bullet', level: 0 } }, [
        schema.text('new bullet'),
      ]),
    ]);

    const bytes = await exportDocx(doc2, { carry: raw });
    const zip = await JSZip.loadAsync(bytes);
    const numberingXml = await zip.file('word/numbering.xml')!.async('string');
    // Original def untouched, new def appended in schema order.
    expect(numberingXml).toContain('<w:num w:numId="1">');
    expect(numberingXml).toContain('w:lvlText w:val="•"');
    expect(numberingXml.lastIndexOf('<w:abstractNum ')).toBeLessThan(
      numberingXml.search(/<w:num[ >]/),
    );

    const { doc: back } = await importDocx(bytes);
    expect(back.child(0).attrs['list']).toEqual({ numId: '1', level: 0 }); // untouched
    const bullet = back.child(1).attrs['list'] as { numId: string };
    expect(bullet.numId).toMatch(/^\d+$/);
    expect(bullet.numId).not.toBe('1');
    const backDefs = back.attrs['numbering'] as Record<
      string,
      { levels: Record<number, { lvlText: string }> }
    >;
    expect(backDefs[bullet.numId].levels[0].lvlText).toBe('•');
  });
});

// Module-scoped: shared by the E4-fidelity and page-setup describes.
async function sourceWithHeader(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="${PR_NS}"><Relationship Id="rIdH" Type="${R_NS}/header" Target="header1.xml"/></Relationships>`,
  );
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>MY HEADER</w:t></w:r></w:p></w:hdr>`,
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body><w:p><w:r><w:t>body text</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

describe('exportDocx (E4 fidelity: sectPr + header refs)', () => {
  it('re-attaches sectPr so headers + page geometry survive (carry)', async () => {
    const { doc, headers, page, raw } = await importDocx(
      await sourceWithHeader(),
    );
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

describe('exportDocx (page setup: w:pgSz / w:pgMar)', () => {
  async function xmlOf(
    doc: import('prosemirror-model').Node,
    opts?: Parameters<typeof exportDocx>[1],
  ): Promise<string> {
    const zip = await JSZip.loadAsync(await exportDocx(doc, opts));
    return (await zip.file('word/document.xml')?.async('string')) ?? '';
  }
  const para = (t: string) => schema.node('paragraph', null, [schema.text(t)]);

  it('emits an A4 body sectPr for a fresh doc without a page attr', async () => {
    // Without one, Word applies ITS default (usually Letter) — not the A4
    // bapbong laid the doc out against.
    const xml = await xmlOf(schema.node('doc', null, [para('x')]));
    expect(xml).toContain('<w:pgSz w:w="11906" w:h="16838"/>'); // canonical A4
    expect(xml).toContain(
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"',
    );
  });

  it('emits doc.attrs.page — landscape Letter swaps dims and rides w:orient', async () => {
    const page = {
      width: 1056, // Letter landscape (px)
      height: 816,
      margin: { top: 48, right: 72, bottom: 48, left: 72 },
    };
    const xml = await xmlOf(schema.node('doc', { page }, [para('x')]));
    expect(xml).toContain(
      '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>',
    );
    expect(xml).toContain(
      '<w:pgMar w:top="720" w:right="1080" w:bottom="720" w:left="1080"',
    );
  });

  it('keeps the carried sectPr byte-identical when page setup is untouched', async () => {
    const { doc, raw } = await importDocx(await sourceWithHeader());
    const xml = await xmlOf(doc, { carry: raw });
    // The exact original sectPr — px↔twips rounding must not drift it.
    expect(xml).toContain(
      '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
    );
  });

  it('splices an edited page setup into the carried sectPr, preserving header refs', async () => {
    const { doc, raw } = await importDocx(await sourceWithHeader());
    const edited = doc.type.create(
      {
        ...doc.attrs,
        page: {
          width: 1123, // rotated to A4 landscape
          height: 794,
          margin: { top: 96, right: 96, bottom: 96, left: 96 },
        },
      },
      doc.content,
    );
    const xml = await xmlOf(edited, { carry: raw });
    expect(xml).toContain('<w:headerReference w:type="default" r:id="rIdH"/>'); // survives
    expect(xml).toContain(
      '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>',
    );
    // Replaced in place: the old portrait pgSz is gone.
    expect(xml).not.toContain('<w:pgSz w:w="11906" w:h="16838"/>');
    // Round-trip: importing the export yields the edited geometry.
    const back = await importDocx(await exportDocx(edited, { carry: raw }));
    expect(back.page.width).toBe(1123);
    expect(back.doc.attrs.page).toMatchObject({ width: 1123, height: 794 });
  });

  it('survives the export worker hop (toJSON → nodeFromJSON)', async () => {
    // The desktop shell autosaves off the main thread: app.ts posts
    // doc.toJSON() to export.worker.ts, which rebuilds it with
    // schema.nodeFromJSON before calling exportDocx. Page setup lives in a doc
    // ATTR, so a serialization that dropped attrs would show the new geometry
    // on screen and write the old one to disk — the failure a user only
    // notices after reopening the file.
    const page = {
      width: 1123, // A4 landscape
      height: 794,
      margin: { top: 48, right: 72, bottom: 48, left: 72 },
    };
    const doc = schema.node('doc', { page }, [para('after page setup')]);
    const rebuilt = schema.nodeFromJSON(
      JSON.parse(JSON.stringify(doc.toJSON())),
    );
    expect(rebuilt.attrs['page']).toEqual(page);

    const back = await importDocx(await exportDocx(rebuilt));
    // The exporter writes Word-default chrome distances (720 twips), which
    // the importer now reads back as explicit 48px distances.
    expect(back.page).toEqual({
      ...page,
      headerDistance: 48,
      footerDistance: 48,
    });
    expect(back.doc.child(0).textContent).toBe('after page setup');
  });

  it('round-trips a per-section geometry override (landscape section)', async () => {
    const landscape = {
      width: 1123, // A4 landscape
      height: 794,
      margin: { top: 96, right: 96, bottom: 96, left: 96 },
    };
    const doc = schema.node(
      'doc',
      {
        sections: [
          {
            blockCount: 1,
            columns: { count: 1, gap: 0 },
            newPage: true,
            page: landscape,
          },
          { blockCount: 1, columns: { count: 1, gap: 0 }, newPage: true },
        ],
      },
      [para('wide table page'), para('back to portrait')],
    );
    const xml = await xmlOf(doc);
    // The break's sectPr carries the override; the body sectPr stays portrait.
    expect(xml).toContain(
      '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>',
    );
    expect(xml).toContain('<w:pgSz w:w="11906" w:h="16838"/>');

    const back = await importDocx(await exportDocx(doc));
    const sections = back.doc.attrs.sections as {
      page?: { width: number; height: number };
    }[];
    expect(sections).toHaveLength(2);
    expect(sections[0].page).toMatchObject({ width: 1123, height: 794 });
    expect(sections[1].page).toBeUndefined();
  });
});

describe('exportDocx (per-section chrome round-trip)', () => {
  async function twoSectionSource(): Promise<Uint8Array> {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
    );
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0"?><Relationships xmlns="${PR_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    );
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="${PR_NS}"><Relationship Id="rIdH1" Type="${R_NS}/header" Target="header1.xml"/><Relationship Id="rIdH2" Type="${R_NS}/header" Target="header2.xml"/></Relationships>`,
    );
    const hdr = (t: string) =>
      `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:hdr>`;
    zip.file('word/header1.xml', hdr('SECTION ONE'));
    zip.file('word/header2.xml', hdr('SECTION TWO'));
    zip.file(
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:titlePg/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH2"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    );
    return zip.generateAsync({ type: 'uint8array' });
  }

  it('re-attaches intermediate header refs + titlePg on carry export', async () => {
    const { doc, raw, sectionChrome } = await importDocx(
      await twoSectionSource(),
    );
    expect(sectionChrome?.[0].headers['default']?.textContent).toBe(
      'SECTION ONE',
    );
    const back = await importDocx(await exportDocx(doc, { carry: raw }));
    expect(back.sectionChrome?.[0].headers['default']?.textContent).toBe(
      'SECTION ONE',
    );
    expect(back.sectionChrome?.[0].titlePg).toBe(true);
    expect(back.sectionChrome?.[1].headers['default']?.textContent).toBe(
      'SECTION TWO',
    );
  });
});

describe('exportDocx (sections + footnotes)', () => {
  // The exported word/document.xml, for inspecting serialization directly.
  async function docXml(
    doc: import('prosemirror-model').Node,
    opts?: Parameters<typeof exportDocx>[1],
  ): Promise<string> {
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
    // Section 0's break sits in the first paragraph's pPr. Every sectPr is
    // self-contained, so it carries the document geometry (A4 default) too.
    expect(xml).toContain(
      '<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"' +
        ' w:header="720" w:footer="720" w:gutter="0"/>' +
        '<w:cols w:num="2" w:space="360"/></w:sectPr>',
    );
    // The non-last section is serialized inline; the last section's slot is
    // the body sectPr (now always emitted, carrying the page geometry).
    expect(xml.match(/<w:sectPr>/g)?.length).toBe(2);
  });

  it('serializes a footnote reference for a footnote-marked run', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('text'),
        schema.text('1', [schema.marks['footnote'].create({ num: 2 })]),
      ]),
    ]);
    const xml = await docXml(doc);
    expect(xml).toContain('<w:footnoteReference w:id="2"/>');
    expect(xml).not.toContain('>1</w:t>'); // the carrier number isn't emitted as text
  });
});

describe('named paragraph styles (Title/Subtitle/HeadingN) + styles.xml', () => {
  it('round-trips Title/Subtitle through pStyle and generates their defs', async () => {
    const doc = makeDoc(
      [
        [{ text: 'Báo cáo' }],
        [{ text: 'quý 3' }],
        [{ text: 'Mục A' }],
        [{ text: 'nội dung' }],
      ],
      [{ styleId: 'Title' }, { styleId: 'Subtitle' }, { heading: 2 }, {}],
    );
    const bytes = await exportDocx(doc);

    const zip = await JSZip.loadAsync(bytes);
    const documentXml = await zip.file('word/document.xml')!.async('string');
    expect(documentXml).toContain('<w:pStyle w:val="Title"/>');
    expect(documentXml).toContain('<w:pStyle w:val="Subtitle"/>');
    expect(documentXml).toContain('<w:pStyle w:val="Heading2"/>');

    const stylesXml = await zip.file('word/styles.xml')!.async('string');
    for (const id of ['Title', 'Subtitle', 'Heading2']) {
      expect(stylesXml).toContain(`w:styleId="${id}"`);
    }
    expect(stylesXml).not.toContain('w:styleId="Heading5"'); // only used defs

    const rels = await zip
      .file('word/_rels/document.xml.rels')!
      .async('string');
    expect(rels).toContain('Target="styles.xml"');
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('/word/styles.xml');

    const { doc: back } = await importDocx(bytes);
    expect(back.child(0).attrs['styleId']).toBe('Title');
    expect(back.child(0).attrs['heading']).toBeNull();
    expect(back.child(1).attrs['styleId']).toBe('Subtitle');
    expect(back.child(2).attrs['heading']).toBe(2);
    expect(back.child(2).attrs['styleId']).toBeNull();
    expect(back.child(3).attrs['styleId']).toBeNull();
  });

  it('merges missing style defs into a carried styles.xml', async () => {
    // A carried package whose styles.xml only defines Normal.
    const src = makeDoc([[{ text: 'x' }]], [{}]);
    const zip = await JSZip.loadAsync(await exportDocx(src));
    zip.file(
      'word/styles.xml',
      `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
    );

    const doc = makeDoc([[{ text: 'tiêu đề' }]], [{ styleId: 'Title' }]);
    const out = await JSZip.loadAsync(await exportDocx(doc, { carry: zip }));
    const styles = await out.file('word/styles.xml')!.async('string');
    expect(styles).toContain('w:styleId="Normal"');
    expect(styles).toContain('w:styleId="Title"');
  });
});
