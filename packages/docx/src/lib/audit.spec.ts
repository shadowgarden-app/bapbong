import JSZip from 'jszip';
import { schema } from '@shadow-garden/bapbong-model';
import { audit } from './audit';
import { importDocx } from './docx';
import { exportDocx } from './export';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Minimal in-memory .docx from a document.xml (+ optional styles.xml). */
async function makeDocx(
  documentXml: string,
  stylesXml?: string,
  themeXml?: string,
  settingsXml?: string,
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file('word/document.xml', documentXml);
  if (stylesXml) zip.file('word/styles.xml', stylesXml);
  if (themeXml) zip.file('word/theme/theme1.xml', themeXml);
  if (settingsXml) zip.file('word/settings.xml', settingsXml);
  return zip.generateAsync({ type: 'uint8array' });
}

// A body exercising the audit's cases: an unread pPr child (w:keepNext),
// unread attrs on read elements (w:pgMar header/footer/gutter), an entirely
// unknown element (w:fictional), and ignored-by-design noise (w:proofErr,
// rsid attrs).
const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}">
  <w:body>
    <w:p w:rsidR="00AA00AA">
      <w:pPr><w:keepNext/><w:jc w:val="both"/></w:pPr>
      <w:proofErr w:type="spellStart"/>
      <w:r><w:t xml:space="preserve">Hello</w:t></w:r>
      <w:fictional w:val="1"/>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const keys = (entries: { key: string }[]) => entries.map((e) => e.key);

const NS =
  `xmlns:w="${W_NS}" ` +
  `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;

/** A run holding a textbox shape. Attributes are only ever attr-checked on
 *  elements the converter VISITED, so wps:bodyPr has to be reached the real
 *  way — through drawing → graphicData → wsp → txbxContent. */
const textboxRun = (bodyPrAttrs: string) =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>` +
  `<a:graphic><a:graphicData><wps:wsp>` +
  `<wps:spPr><a:prstGeom prst="rect"/></wps:spPr>` +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
  `<wps:bodyPr ${bodyPrAttrs}/>` +
  `</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

describe('xml audit (import)', () => {
  afterEach(() => audit.setEnabled(false));

  it('stays silent when the flag is off', async () => {
    audit.setEnabled(false);
    const before = audit.lastReport;
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await makeDocx(DOCUMENT_XML));
    expect(audit.lastReport).toBe(before);
    expect(
      log.mock.calls.filter((c) => String(c[0]).includes('[xml-audit]')),
    ).toHaveLength(0);
    log.mockRestore();
  });

  it('reports untouched tags/attrs as UNKNOWN and noise as ignored', async () => {
    audit.setEnabled(true);
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await makeDocx(DOCUMENT_XML));
    log.mockRestore();

    const report = audit.lastReport;
    expect(report?.mode).toBe('import');
    const unknown = keys(report?.unknown ?? []);
    const ignored = keys(report?.ignored ?? []);

    // Real gaps surface as UNKNOWN. (w:keepNext used to be one, but inline
    // pPr extras are now CARRIED — preserved verbatim on export — so the
    // audit rightly counts them consumed.)
    expect(unknown).not.toContain('w:keepNext');
    expect(unknown).toContain('w:fictional');
    // pgMar chrome distances + gutter are read into PageConfig now.
    expect(unknown).not.toContain('w:pgMar @w:header');
    expect(unknown).not.toContain('w:pgMar @w:footer');
    expect(unknown).not.toContain('w:pgMar @w:gutter');

    // Deliberately-skipped noise is classified, not flagged.
    expect(ignored).toContain('w:proofErr');
    expect(ignored).toContain('w:p @w:rsidR');
    expect(unknown).not.toContain('w:proofErr');

    // What the importer reads must NOT be reported.
    expect(unknown).not.toContain('w:jc');
    expect(unknown).not.toContain('w:t');
    expect(unknown).not.toContain('w:pgSz');
    expect(unknown).not.toContain('w:sectPr');
    // Asked-for attrs count as covered.
    expect(unknown).not.toContain('w:jc @w:val');
    expect(unknown).not.toContain('w:pgMar @w:top');
  });

  it('judges w:compat entry by entry: read, not adopted, or a real gap', async () => {
    // settings.xml used to be ignored wholesale as "editor UI". Its w:compat
    // block is layout: two entries are READ into DocCompat, the ones Word
    // writes on every file but this program has not adopted are classified
    // by name (each a candidate field), and anything new stays UNKNOWN.
    audit.setEnabled(true);
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    const cs = (name: string, val: string) =>
      `<w:compatSetting w:name="${name}" w:uri="http://schemas.microsoft.com/office/word" w:val="${val}"/>`;
    await importDocx(
      await makeDocx(
        DOCUMENT_XML,
        undefined,
        undefined,
        `<?xml version="1.0"?><w:settings xmlns:w="${W_NS}"><w:compat>` +
          '<w:ulTrailSpace/><w:doNotUseHTMLParagraphAutoSpacing/><w:fictionalCompat/>' +
          cs('compatibilityMode', '15') +
          cs('overrideTableStyleFontSizeAndJustification', '1') +
          cs('enableOpenTypeFeatures', '1') +
          cs('brandNewSetting', '1') +
          '</w:compat></w:settings>',
      ),
    );
    log.mockRestore();
    const report = audit.lastReport;
    const unknown = keys(report?.unknown ?? []);
    const ignored = keys(report?.ignored ?? []);
    // Read.
    expect(unknown.filter((k) => k.includes('compatibilityMode'))).toEqual([]);
    expect(
      unknown.filter((k) => k.includes('overrideTableStyleFontSize')),
    ).toEqual([]);
    expect(unknown).not.toContain('w:doNotUseHTMLParagraphAutoSpacing');
    // Not adopted, by name — keyed per setting so the list stays readable.
    expect(ignored).toContain('w:compatSetting[enableOpenTypeFeatures] @w:val');
    expect(ignored).toContain('w:ulTrailSpace');
    // The namespace URI is plumbing on every entry.
    expect(unknown.some((k) => k.endsWith('@w:uri'))).toBe(false);
    // New to us: a real gap, and it says which one.
    expect(unknown).toContain('w:compatSetting[brandNewSetting] @w:val');
    expect(unknown).toContain('w:fictionalCompat');
  });

  it('demotes a value that asks for nothing — and only that value', async () => {
    audit.setEnabled(true);
    const anchored = (wrap: string) =>
      `<?xml version="1.0"?><w:document ${NS}><w:body><w:p><w:r><w:drawing>` +
      `<wp:anchor><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="914400" cy="914400"/>${wrap}` +
      `<a:graphic><a:graphicData><wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr></wps:wsp></a:graphicData></a:graphic>` +
      `</wp:anchor></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`;
    const box = (x2: number, y2: number, extra = '') =>
      `<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="${y2}"/>` +
      `<wp:lineTo x="${x2}" y="${y2}"/>${extra}<wp:lineTo x="${x2}" y="0"/>` +
      `<wp:lineTo x="0" y="0"/></wp:wrapPolygon>`;
    const run = async (xml: string) => {
      const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
      await importDocx(await makeDocx(xml));
      log.mockRestore();
      return {
        unknown: keys(audit.lastReport?.unknown ?? []),
        inert: keys(audit.lastReport?.inert ?? []),
      };
    };

    // bothSides is the default and is what our square wrap does; a polygon
    // that IS the box (Word rounds its far edge a few thousandths short of
    // 21600) is the rectangle we already exclude.
    const plain = await run(
      anchored(
        `<wp:wrapThrough wrapText="bothSides">${box(21435, 21236)}</wp:wrapThrough>`,
      ),
    );
    expect(plain.inert).toContain('wp:wrapThrough @wrapText');
    expect(plain.inert).toContain('wp:wrapPolygon');
    expect(plain.unknown).not.toContain('wp:wrapPolygon');

    // One side only is a real gap.
    const side = await run(
      anchored(
        `<wp:wrapThrough wrapText="left">${box(21435, 21236)}</wp:wrapThrough>`,
      ),
    );
    expect(side.unknown).toContain('wp:wrapThrough @wrapText');

    // A rectangle INSET from the box would let text closer than we allow…
    const inset = await run(
      anchored(
        `<wp:wrapThrough wrapText="bothSides">${box(15000, 21236)}</wp:wrapThrough>`,
      ),
    );
    expect(inset.unknown).toContain('wp:wrapPolygon');

    // …and a polygon with a fifth corner carves a shape we cannot follow.
    const carved = await run(
      anchored(
        `<wp:wrapThrough wrapText="bothSides">${box(21435, 21236, '<wp:lineTo x="10000" y="9000"/>')}</wp:wrapThrough>`,
      ),
    );
    expect(carved.unknown).toContain('wp:wrapPolygon');
  });

  it('demotes autoSpace only when it is switched OFF', async () => {
    audit.setEnabled(true);
    // In a paragraph's own pPr these would be CARRIED (preserved verbatim on
    // export) and so count as consumed; a style is where they show up
    // unhandled, which is where the corpus has them.
    const styles = (attrs: string) =>
      `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">` +
      `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/>` +
      `<w:pPr><w:autoSpaceDE ${attrs}/><w:autoSpaceDN ${attrs}/></w:pPr>` +
      `</w:style></w:styles>`;
    const body =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body><w:p>` +
      `<w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>x</w:t></w:r>` +
      `</w:p><w:sectPr/></w:body></w:document>`;
    const run = async (attrs: string) => {
      const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
      await importDocx(await makeDocx(body, styles(attrs)));
      log.mockRestore();
      return {
        unknown: keys(audit.lastReport?.unknown ?? []),
        inert: keys(audit.lastReport?.inert ?? []),
      };
    };
    // Off asks for exactly what we do — we never insert East-Asian spacing.
    const off = await run('w:val="0"');
    expect(off.inert).toContain('w:autoSpaceDE');
    expect(off.unknown).not.toContain('w:autoSpaceDE');
    // On (including the absent-means-on default) is a feature we lack.
    expect((await run('')).unknown).toContain('w:autoSpaceDE');
    expect((await run('w:val="1"')).unknown).toContain('w:autoSpaceDN');
  });

  it('separates a declaration nothing points at from a real gap', async () => {
    audit.setEnabled(true);
    const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    // A theme whose format scheme offers a gradient nobody selects.
    const theme =
      `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}" name="t"><a:themeElements>` +
      `<a:fmtScheme name="Office"><a:fillStyleLst>` +
      `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill>` +
      `</a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
    const body = (ref: string) =>
      `<?xml version="1.0"?><w:document ${NS}><w:body><w:p><w:r><w:drawing>` +
      `<wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData><wps:wsp>` +
      `<wps:spPr><a:prstGeom prst="rect"/></wps:spPr>${ref}` +
      `</wps:wsp></a:graphicData></a:graphic></wp:inline>` +
      `</w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`;
    const run = async (ref: string) => {
      const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
      await importDocx(await makeDocx(body(ref), undefined, theme));
      log.mockRestore();
      const r = audit.lastReport;
      return {
        unknown: keys(r?.unknown ?? []),
        unref: keys(r?.unreferenced ?? []),
      };
    };

    // No shape selects from the matrix: the gradient renders nothing here.
    const dead = await run('');
    expect(dead.unref).toContain('a:gradFill');
    expect(dead.unknown).not.toContain('a:gradFill');

    // One a:lnRef anywhere and the whole matrix is back in play — a shape is
    // picking entries out of it, and the ones we do not read are real gaps.
    const live = await run(
      `<wps:style><a:lnRef idx="2"><a:srgbClr val="000000"/></a:lnRef></wps:style>`,
    );
    expect(live.unknown).toContain('a:gradFill');
    expect(live.unref).not.toContain('a:gradFill');
  });

  it('counts a smart tag handled once its runs are unwrapped', async () => {
    audit.setEnabled(true);
    const body =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body><w:p>` +
      `<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">` +
      `<w:smartTagPr><w:attr w:name="ProductID" w:val="x"/></w:smartTagPr>` +
      `<w:r><w:t>Nam</w:t></w:r></w:smartTag>` +
      `</w:p><w:sectPr/></w:body></w:document>`;
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    const { doc } = await importDocx(await makeDocx(body));
    log.mockRestore();

    expect(doc.child(0).textContent).toBe('Nam');
    const unknown = keys(audit.lastReport?.unknown ?? []);
    const ignored = keys(audit.lastReport?.ignored ?? []);
    // The wrapper is consumed, and its property bag with it.
    expect(unknown).not.toContain('w:smartTag');
    expect(unknown).not.toContain('w:smartTagPr');
    // Which recogniser claimed the text is metadata, not content.
    expect(ignored).toContain('w:smartTag @w:element');
    expect(ignored).toContain('w:smartTag @w:uri');
  });

  it('demotes unread no-op values to inert, keeps real values UNKNOWN', async () => {
    audit.setEnabled(true);
    // A textbox carrying the schema-default set Word stamps on every shape,
    // plus the no-op elements it writes unprompted. The loose elements in the
    // second run sit where the audit can see them unvisited; what is under
    // test is the bucketing rule, not where Word would really put them.
    const body =
      `<?xml version="1.0"?><w:document ${NS}><w:body><w:p>` +
      textboxRun('rot="0" wrap="square" compatLnSpc="1"') +
      `<w:r>` +
      `<a:effectLst/><a:srcRect/>` +
      `<a:ln><a:noFill/></a:ln>` +
      `<a:prstTxWarp prst="textNoShape"/>` +
      `<w:tblInd w:w="0" w:type="dxa"/>` +
      `<w:docGrid w:linePitch="360"/>` +
      `</w:r></w:p><w:sectPr/></w:body></w:document>`;
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await makeDocx(body));
    log.mockRestore();

    const unknown = keys(audit.lastReport?.unknown ?? []);
    const inert = keys(audit.lastReport?.inert ?? []);

    // Attributes sitting at their ECMA-376 default, and elements the spec
    // defines as no-ops in exactly this state.
    expect(inert).toContain('wps:bodyPr @rot');
    expect(inert).toContain('wps:bodyPr @wrap');
    expect(inert).toEqual(expect.arrayContaining(['a:effectLst', 'a:srcRect']));
    expect(inert).toContain('a:ln');
    expect(inert).toContain('a:prstTxWarp');
    expect(inert).toContain('w:tblInd');
    expect(inert).toContain('w:docGrid');
    for (const k of inert) expect(unknown).not.toContain(k);

    // @compatLnSpc's default is "0" — a written "1" really does change line
    // spacing in the shape, so it must NOT be demoted along with its
    // neighbours just because it rode in on the same element.
    expect(unknown).toContain('wps:bodyPr @compatLnSpc');
  });

  it('demotes settings/OLE chrome and the no-op shapes of shared elements', async () => {
    audit.setEnabled(true);
    const body =
      `<?xml version="1.0"?><w:document ${NS}><w:body><w:p><w:r>` +
      // Two shapes of the same key. Word writes the first on every document;
      // the second configures note numbering and must stay a gap.
      `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>` +
      `<w:endnotePr><w:numFmt w:val="lowerRoman"/></w:endnotePr>` +
      // Same for a content control's chrome: ids and gallery metadata say
      // nothing, an rPr might.
      `<w:sdtPr><w:id w:val="1"/><w:docPartObj/></w:sdtPr>` +
      `<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>` +
      `</w:r></w:p><w:sectPr/></w:body></w:document>`;
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await makeDocx(body));
    log.mockRestore();

    const unknown = keys(audit.lastReport?.unknown ?? []);
    const inert = keys(audit.lastReport?.inert ?? []);
    expect(inert).toContain('w:footnotePr');
    expect(inert).toContain('w:sdtPr');
    // The other shape of the very same key is still a gap.
    expect(unknown).toContain('w:endnotePr');
    expect(unknown).toContain('w:sdtEndPr');
  });

  it('demotes by value, not by name: the same key can be both', async () => {
    audit.setEnabled(true);
    const body =
      `<?xml version="1.0"?><w:document ${NS}><w:body><w:p>` +
      textboxRun('vert="horz"') +
      textboxRun('vert="vert270"') +
      `</w:p><w:sectPr/></w:body></w:document>`;
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await makeDocx(body));
    log.mockRestore();

    // An ignore-list keyed by NAME would silence both. Keying by VALUE keeps
    // the one that actually moves text on the page visible.
    expect(keys(audit.lastReport?.inert ?? [])).toContain('wps:bodyPr @vert');
    expect(keys(audit.lastReport?.unknown ?? [])).toContain('wps:bodyPr @vert');
  });

  it('does not demote values that merely match our own fallback', async () => {
    audit.setEnabled(true);
    // TableNormal's 108-twip cell margins equal layout's CELL_PAD_X (7.2px),
    // so this renders identically today — by coincidence, not by spec. It is
    // a real gap and has to stay countable.
    const styles =
      `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">` +
      `<w:style w:type="table" w:default="1" w:styleId="TableNormal">` +
      `<w:tblPr><w:tblCellMar>` +
      `<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>` +
      `<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>` +
      `</w:tblCellMar></w:tblPr></w:style></w:styles>`;
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await makeDocx(DOCUMENT_XML, styles));
    log.mockRestore();

    const unknown = keys(audit.lastReport?.unknown ?? []);
    expect(unknown).toContain('w:left');
    expect(keys(audit.lastReport?.inert ?? [])).not.toContain('w:left');
  });

  it('a cascade layer that was overridden is handled, not missing', async () => {
    audit.setEnabled(true);
    // Heading3 carries w:tabs and w:outlineLvl; the paragraph overrides the
    // tabs inline and its style NAME settles the heading level. Both style
    // properties were resolved correctly, so neither is a gap — before this,
    // the cascade's early return left them looking unread.
    const stylesXml =
      `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">` +
      `<w:style w:type="paragraph" w:styleId="Heading3"><w:pPr>` +
      `<w:tabs><w:tab w:val="left" w:pos="-48"/></w:tabs>` +
      `<w:jc w:val="left"/><w:outlineLvl w:val="2"/>` +
      `</w:pPr></w:style></w:styles>`;
    const documentXml =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body><w:p><w:pPr>` +
      `<w:pStyle w:val="Heading3"/>` +
      `<w:tabs><w:tab w:val="center" w:pos="2160"/></w:tabs>` +
      `<w:jc w:val="center"/>` +
      `</w:pPr><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await makeDocx(documentXml, stylesXml));
    log.mockRestore();

    const unknown = keys(audit.lastReport?.unknown ?? []);
    expect(unknown).not.toContain('w:tabs');
    expect(unknown).not.toContain('w:jc');
    expect(unknown).not.toContain('w:outlineLvl');
  });

  it('separator footnotes and comment bodies count as consumed subtrees', async () => {
    audit.setEnabled(true);
    const zip = new JSZip();
    const bytes = await makeDocx(
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>x</w:t><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>`,
    );
    await zip.loadAsync(bytes);
    zip.file(
      'word/footnotes.xml',
      `<?xml version="1.0"?><w:footnotes xmlns:w="${W_NS}">` +
        `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote>` +
        `</w:footnotes>`,
    );
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    log.mockRestore();

    const unknown = keys(audit.lastReport?.unknown ?? []);
    // The separator note's content is skipped by design — no w:separator noise.
    expect(unknown).not.toContain('w:separator');
    expect(unknown).not.toContain('w:p');
  });
});

describe('xml audit (export + round-trip baseline)', () => {
  afterEach(() => audit.setEnabled(false));

  it('logs nothing for a doc made only of handled types', async () => {
    audit.setEnabled(true);
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('plain '),
        schema.text('bold', [schema.marks['strong'].create()]),
      ]),
    ]);
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await exportDocx(doc);
    log.mockRestore();
    expect(audit.lastReport?.mode).toBe('export');
    expect(audit.lastReport?.unknown).toEqual([]);
  });

  it('round-trip: what our own exporter writes that the importer never reads', async () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { heading: 2 }, [schema.text('Heading')]),
      schema.node('paragraph', null, [schema.text('Body')]),
    ]);
    const bytes = await exportDocx(doc);

    audit.setEnabled(true);
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(bytes);
    log.mockRestore();

    // BASELINE of known round-trip losses (see the bb3.docx survey). It is
    // EMPTY: everything our exporter writes, our importer reads back. Any
    // entry appearing here is a regression — either the exporter started
    // writing something new, or the importer stopped reading something.
    // (w:style @w:type left when the registry began reading every style's
    // kind; w:outlineLvl when the cascade started marking the layers a more
    // derived one overrides; the pgMar chrome went earlier, into PageConfig.)
    const unknown = keys(audit.lastReport?.unknown ?? []).sort();
    expect(unknown).toEqual([]);
  });
});
