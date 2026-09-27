/**
 * DrawingML charts (D-2609-LSTQ: a thesis with five charts, every one of them
 * gone after a save). A chart's drawing holds only `<c:chart r:id>`; the
 * chart lives in its own part, which relates to an embedded workbook and a
 * theme override. The fixture is that shape: an inline chart reaching all
 * three, and an anchored one with a part of its own.
 */
import JSZip from 'jszip';
import { audit } from './audit';
import { chartPlaceholder } from './chart';
import { importDocx } from './docx';
import { exportDocx } from './export';
import type { CarriedPart, CarriedRel } from './package-parts';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CHART_REL = `${R_NS}/chart`;
const CHART_CT =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const XLSX_CT =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** What Word writes for a chart frame: the chart declares c: on itself. */
const chartGraphic = (rId: string) =>
  `<a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" r:id="${rId}"/></a:graphicData></a:graphic>`;

const INLINE_CHART =
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="5486400" cy="3200400"/><wp:effectExtent l="0" t="0" r="19050" b="19050"/>` +
  `<wp:docPr id="1" name="Chart 1" descr="Growth"/><wp:cNvGraphicFramePr/>` +
  `${chartGraphic('rId5')}</wp:inline></w:drawing></w:r>`;

const ANCHORED_CHART =
  `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>95250</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="4572000" cy="2743200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>` +
  `<wp:docPr id="2" name="Chart 2"/><wp:cNvGraphicFramePr/>` +
  `${chartGraphic('rId6')}</wp:anchor></w:drawing></w:r>`;

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}"><w:body>` +
  `<w:p>${INLINE_CHART}</w:p><w:p>${ANCHORED_CHART}<w:r><w:t>Below the chart.</w:t></w:r></w:p>` +
  `</w:body></w:document>`;

const CHART1_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C_NS}" xmlns:r="${R_NS}"><c:chart><c:plotArea/></c:chart><c:externalData r:id="rId2"/></c:chartSpace>`;
const CHART2_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C_NS}"><c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>`;
const THEME_OVERRIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:themeOverride xmlns:a="${A_NS}"/>`;
/** Not a real workbook — bytes that are not text, which is what matters. */
const WORKBOOK = new Uint8Array([
  0x50, 0x4b, 0x03, 0x04, 0x00, 0xff, 0x80, 0x7f,
]);

const rels = (entries: [string, string, string][]) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR_NS}">` +
  entries
    .map(
      ([id, type, target]) =>
        `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`,
    )
    .join('') +
  `</Relationships>`;

/** The source package. `chart1` swaps the first chart's XML — the same ids
 *  and paths holding a different chart, as another document would. */
async function makePackage(chart1 = CHART1_XML): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="xlsx" ContentType="${XLSX_CT}"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/charts/chart1.xml" ContentType="${CHART_CT}"/>` +
      `<Override PartName="/word/charts/chart2.xml" ContentType="${CHART_CT}"/>` +
      `<Override PartName="/word/theme/themeOverride1.xml" ContentType="application/vnd.openxmlformats-officedocument.themeOverride+xml"/>` +
      `</Types>`,
  );
  zip.file(
    '_rels/.rels',
    rels([['rId1', `${R_NS}/officeDocument`, 'word/document.xml']]),
  );
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file(
    'word/_rels/document.xml.rels',
    rels([
      ['rId5', CHART_REL, 'charts/chart1.xml'],
      ['rId6', CHART_REL, 'charts/chart2.xml'],
    ]),
  );
  zip.file('word/charts/chart1.xml', chart1);
  zip.file(
    'word/charts/_rels/chart1.xml.rels',
    rels([
      ['rId1', `${R_NS}/themeOverride`, '../theme/themeOverride1.xml'],
      [
        'rId2',
        `${R_NS}/package`,
        '../embeddings/Microsoft_Excel_Worksheet1.xlsx',
      ],
    ]),
  );
  zip.file('word/charts/chart2.xml', CHART2_XML);
  zip.file('word/theme/themeOverride1.xml', THEME_OVERRIDE_XML);
  zip.file('word/embeddings/Microsoft_Excel_Worksheet1.xlsx', WORKBOOK);
  return zip.generateAsync({ type: 'uint8array' });
}

interface Raw {
  xml: string;
  rels: CarriedRel[];
  parts: CarriedPart[];
}

const chartsOf = (doc: import('prosemirror-model').Node) => {
  const out: import('prosemirror-model').Node[] = [];
  doc.descendants((n) => {
    if (n.type.name === 'image' && n.attrs['rawDrawing']) out.push(n);
  });
  return out;
};

/** document.xml's chart references, in order, resolved to part paths. */
async function chartTargets(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const docXml = await zip.file('word/document.xml')!.async('string');
  const relsXml = await zip
    .file('word/_rels/document.xml.rels')!
    .async('string');
  const ids = [...docXml.matchAll(/<c:chart\b[^>]*\br:id="([^"]+)"/g)].map(
    (m) => m[1],
  );
  const targets = ids.map((id) => {
    const m = new RegExp(`<Relationship Id="${id}" [^>]*Target="([^"]+)"`).exec(
      relsXml,
    );
    return m ? `word/${m[1]}` : null;
  });
  return { zip, ids, targets };
}

describe('a DrawingML chart', () => {
  it('imports as a box the size of its frame', async () => {
    const { doc } = await importDocx(await makePackage());
    const [inline, anchored] = chartsOf(doc);
    expect(inline.attrs['width']).toBe(576);
    expect(inline.attrs['height']).toBe(336);
    expect(inline.attrs['alt']).toBe('Growth');
    expect(inline.attrs['float']).toBeNull();
    expect(inline.attrs['vector']).toEqual(chartPlaceholder(576, 336));
    expect(anchored.attrs['width']).toBe(480);
    expect(anchored.attrs['float']).toMatchObject({
      wrap: 'topAndBottom',
      hOffset: 10,
      vOffset: 20,
    });
  });

  it('carries its drawing and every part the chart reaches', async () => {
    const { doc } = await importDocx(await makePackage());
    const raw = chartsOf(doc)[0].attrs['rawDrawing'] as Raw;
    expect(raw.xml).toContain('<c:chart r:id="rId5"/>');
    // The chart's own namespace declaration moves to the carried root.
    expect(raw.xml).toMatch(new RegExp(`^<w:drawing[^>]* xmlns:c="${C_NS}"`));
    expect(raw.rels).toEqual([
      { id: 'rId5', type: CHART_REL, target: 'word/charts/chart1.xml' },
    ]);
    expect(raw.parts.map((p) => p.path)).toEqual([
      'word/charts/chart1.xml',
      'word/theme/themeOverride1.xml',
      'word/embeddings/Microsoft_Excel_Worksheet1.xlsx',
    ]);
    const [chart, , workbook] = raw.parts;
    expect(chart.contentType).toBe(CHART_CT);
    expect(chart.data).toBe(CHART1_XML);
    expect(chart.rels?.map((r) => r.target)).toEqual([
      'word/theme/themeOverride1.xml',
      'word/embeddings/Microsoft_Excel_Worksheet1.xlsx',
    ]);
    // The workbook is bytes (default content type by extension).
    expect(workbook.contentType).toBe(XLSX_CT);
    expect(workbook.base64).toBe(true);
  });

  it('is read in full — nothing of it reported as a gap', async () => {
    audit.setEnabled(true);
    try {
      await importDocx(await makePackage());
      const unknown = (audit.lastReport?.unknown ?? []).map((e) => e.key);
      expect(unknown.filter((k) => /^(wp|a|c):/.test(k))).toEqual([]);
    } finally {
      audit.setEnabled(false);
    }
  });

  it('keeps the source parts on a save that still has them', async () => {
    const { doc, raw } = await importDocx(await makePackage());
    const out = await exportDocx(doc, { carry: raw });
    const { zip, ids, targets } = await chartTargets(out);
    expect(ids).toEqual(['rId5', 'rId6']);
    expect(targets).toEqual([
      'word/charts/chart1.xml',
      'word/charts/chart2.xml',
    ]);
    // No copies written beside them.
    expect(
      Object.keys(zip.files)
        .filter((p) => p.startsWith('word/charts/') && !zip.files[p].dir)
        .sort(),
    ).toEqual([
      'word/charts/_rels/chart1.xml.rels',
      'word/charts/chart1.xml',
      'word/charts/chart2.xml',
    ]);
  });

  it('writes the parts itself when there is no source package', async () => {
    const { doc } = await importDocx(await makePackage());
    const out = await exportDocx(doc);
    const { zip, targets } = await chartTargets(out);
    expect(targets).toEqual([
      'word/charts/chart1.xml',
      'word/charts/chart2.xml',
    ]);
    expect(await zip.file('word/charts/chart1.xml')!.async('string')).toBe(
      CHART1_XML,
    );
    const chartRels = await zip
      .file('word/charts/_rels/chart1.xml.rels')!
      .async('string');
    expect(chartRels).toContain('Target="../theme/themeOverride1.xml"');
    expect(chartRels).toContain(
      'Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"',
    );
    expect(
      await zip
        .file('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!
        .async('uint8array'),
    ).toEqual(WORKBOOK);
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).toContain(
      `<Override PartName="/word/charts/chart1.xml" ContentType="${CHART_CT}"/>`,
    );
    expect(types).toContain(
      `<Override PartName="/word/embeddings/Microsoft_Excel_Worksheet1.xlsx" ContentType="${XLSX_CT}"/>`,
    );
    // And the saved file reads back to the same two charts.
    const again = chartsOf((await importDocx(out)).doc);
    expect(again).toHaveLength(2);
    expect((again[0].attrs['rawDrawing'] as Raw).parts).toHaveLength(3);
  });

  it('gives a second copy of a chart parts of its own', async () => {
    const { doc, raw } = await importDocx(await makePackage());
    const para = doc.child(0);
    const twice = para.copy(para.content.append(para.content));
    const out = await exportDocx(doc.copy(doc.content.replaceChild(0, twice)), {
      carry: raw,
    });
    const { zip, ids, targets } = await chartTargets(out);
    expect(ids).toHaveLength(3);
    expect(ids[0]).toBe('rId5');
    expect(ids[1]).not.toBe('rId5');
    // A fresh name: the source's chart1/chart2 are both still in the package.
    expect(targets[0]).toBe('word/charts/chart1.xml');
    expect(targets[1]).toBe('word/charts/chart3.xml');
    expect(await zip.file('word/charts/chart3.xml')!.async('string')).toBe(
      CHART1_XML,
    );
    expect(
      await zip.file('word/charts/_rels/chart3.xml.rels')!.async('string'),
    ).toContain('Target="../embeddings/Microsoft_Excel_Worksheet2.xlsx"');
  });

  it('does not borrow the parts another document keeps under the same ids', async () => {
    const { doc } = await importDocx(await makePackage());
    const other = await importDocx(
      await makePackage(
        CHART1_XML.replace(
          '<c:plotArea/>',
          '<c:plotArea><c:pieChart/></c:plotArea>',
        ),
      ),
    );
    const out = await exportDocx(doc, { carry: other.raw });
    const { zip, ids, targets } = await chartTargets(out);
    // Chart 1 differs from what the other package holds: written fresh.
    expect(ids[0]).not.toBe('rId5');
    expect(await zip.file(targets[0]!)!.async('string')).toBe(CHART1_XML);
    // Chart 2 is byte-identical there, so its id and part are kept.
    expect(ids[1]).toBe('rId6');
  });
});
