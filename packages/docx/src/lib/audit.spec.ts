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

  it('demotes unread no-op values to inert, keeps real values UNKNOWN', async () => {
    audit.setEnabled(true);
    // A textbox carrying the schema-default set Word stamps on every shape,
    // plus the no-op elements it writes unprompted. The loose elements in the
    // second run sit where the audit can see them unvisited; what is under
    // test is the bucketing rule, not where Word would really put them.
    const body =
      `<?xml version="1.0"?><w:document ${NS}><w:body><w:p>` +
      textboxRun('anchor="t" rot="0" wrap="square" compatLnSpc="1"') +
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
    expect(inert).toContain('wps:bodyPr @anchor');
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

  it('demotes by value, not by name: the same key can be both', async () => {
    audit.setEnabled(true);
    const body =
      `<?xml version="1.0"?><w:document ${NS}><w:body><w:p>` +
      textboxRun('anchor="t"') +
      textboxRun('anchor="ctr"') +
      `</w:p><w:sectPr/></w:body></w:document>`;
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    await importDocx(await makeDocx(body));
    log.mockRestore();

    // An ignore-list keyed by NAME would silence both. Keying by VALUE keeps
    // the one that actually moves text on the page visible.
    expect(keys(audit.lastReport?.inert ?? [])).toContain('wps:bodyPr @anchor');
    expect(keys(audit.lastReport?.unknown ?? [])).toContain(
      'wps:bodyPr @anchor',
    );
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

    // BASELINE of known round-trip losses (see the bb3.docx survey). Shrink
    // this list as gaps get fixed; a new entry appearing here means the
    // exporter started writing something the importer doesn't read back.
    const unknown = keys(audit.lastReport?.unknown ?? []).sort();
    expect(unknown).toEqual([
      // Read in general, but headingLevel() short-circuits on the "Heading2"
      // style id before asking for outlineLvl — unread in THIS document.
      'w:outlineLvl',
      // NOTE the pgMar chrome distances/gutter no longer appear (read into
      // PageConfig), and w:style @w:default is read for the unused-style
      // sweep. The FUNCTIONAL gap (default styles are not applied to
      // unstyled content) still exists — the audit just can't see an attr
      // that is read for classification; tracked in the support plan.
      'w:style @w:type',
    ]);
  });
});
