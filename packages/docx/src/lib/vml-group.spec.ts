/**
 * VML groups (v:group) — legacy diagrams, framed captions and whole bar
 * charts drawn in VML (D-2609-DHQ8: 18 groups, every one dropped). The
 * fixture is that file's shape in miniature: a page-anchored group in its
 * own coordinate space holding a freeform drawn with relative moves and
 * run-together commands, a text frame (shapetype 202), a nested group, and
 * a picture.
 */
import JSZip from 'jszip';
import type {
  VectorImageOp,
  VectorPolygonOp,
} from '@shadow-garden/bapbong-contracts';
import { importDocx } from './docx';
import { exportDocx } from './export';
import { parseVmlPath, vmlGroupDrawing } from './vml-group';
import { child, parseXml } from './ooxml';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const V_NS = 'urn:schemas-microsoft-com:vml';
const O_NS = 'urn:schemas-microsoft-com:office:office';
const W10_NS = 'urn:schemas-microsoft-com:office:word';

// A 1×1 PNG.
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/** 150pt × 75pt (200 × 100 px) over coordinates 1000…1400 × 2000…2200 —
 *  half a px per unit. */
const GROUP = `<w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"/><v:group id="g1" style="position:absolute;margin-left:30pt;margin-top:6pt;width:150pt;height:75pt;z-index:-1;mso-position-horizontal-relative:page" coordorigin="1000,2000" coordsize="400,200"><v:shape id="s1" style="position:absolute;left:1000;top:2000;width:100;height:100" coordorigin="1000,2000" coordsize="100,100" path="m1000,2000r100,l1100,2100xe" fillcolor="#9a3365" stroked="f"/><v:shape id="s2" type="#_x0000_t202" style="position:absolute;left:1200;top:2000;width:200;height:40" filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>Label</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape><v:group id="g2" style="position:absolute;left:1200;top:2100;width:200;height:100" coordorigin="0,0" coordsize="20,10"><v:line id="l1" from="0,5" to="20,5" strokecolor="red"/></v:group><v:shape id="p1" style="position:absolute;left:1000;top:2100;width:40;height:40"><v:imagedata r:id="rId9" o:title=""/></v:shape><w10:wrap anchorx="page"/></v:group></w:pict>`;

async function makeDocx(body: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R_NS}/image" Target="media/image4.png"/></Relationships>`,
  );
  zip.file('word/media/image4.png', PNG, { base64: true });
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:v="${V_NS}" xmlns:o="${O_NS}" xmlns:w10="${W10_NS}"><w:body>${body}</w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

const groupDoc = () => makeDocx(`<w:p><w:r>${GROUP}</w:r></w:p>`);

describe('parseVmlPath', () => {
  it('splits run-together commands — `xe` closes, then ends', () => {
    expect(parseVmlPath('m0,0l10,0,10,10xe')).toEqual([
      {
        pts: [
          { x: 0, y: 0 },
          { x: 10, y: 0 },
          { x: 10, y: 10 },
        ],
        closed: true,
      },
    ]);
  });

  it('reads an empty argument as 0', () => {
    // The unit box Word writes for shapetype 202.
    const [box] = parseVmlPath('m,l,21600r21600,l21600,xe')!;
    expect(box.pts).toEqual([
      { x: 0, y: 0 },
      { x: 0, y: 21600 },
      { x: 21600, y: 21600 },
      { x: 21600, y: 0 },
    ]);
    expect(box.closed).toBe(true);
  });

  it('moves relative to the pen with r, and opens a subpath per m', () => {
    const subs = parseVmlPath('m1142,1973r,22m1142,939r,1034e')!;
    expect(subs).toEqual([
      {
        pts: [
          { x: 1142, y: 1973 },
          { x: 1142, y: 1995 },
        ],
        closed: false,
      },
      {
        pts: [
          { x: 1142, y: 939 },
          { x: 1142, y: 1973 },
        ],
        closed: false,
      },
    ]);
  });

  it('gives up on formula references', () => {
    expect(parseVmlPath('m@4@5l@4@11xe')).toBeNull();
  });
});

describe('vmlGroupDrawing', () => {
  const draw = () => {
    const root = parseXml(
      `<w:r xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:v="${V_NS}" xmlns:o="${O_NS}" xmlns:w10="${W10_NS}">${GROUP}</w:r>`,
    );
    const group = child(child(child(root, 'w:r'), 'w:pict'), 'v:group')!;
    return vmlGroupDrawing(group, 200, 100, {
      textbox: () => ({ blocks: [{ type: 'paragraph' }] }),
      color: (v) => (v === 'red' ? '#ff0000' : v?.toUpperCase()),
      length: () => 1,
      shapeType: (id) => (id === '_x0000_t202' ? 202 : undefined),
      image: () => 'data:image/png;base64,' + PNG,
      read: () => undefined,
    });
  };

  it('maps member coordinates through the group space to px', () => {
    const { vector } = draw();
    expect(vector).toMatchObject({ width: 200, height: 100 });
    const fill = vector.ops[0] as VectorPolygonOp;
    expect(fill).toMatchObject({ kind: 'polygon', fill: '#9A3365' });
    expect(fill.stroke).toBeUndefined();
    expect(fill.points).toEqual([
      { x: 0, y: 0 },
      { x: 50, y: 0 },
      { x: 50, y: 50 },
    ]);
  });

  it('places a text frame and draws no box for it', () => {
    const { vector, frames } = draw();
    expect(frames).toEqual([
      {
        x: 100,
        y: 0,
        width: 100,
        height: 20,
        textbox: { blocks: [{ type: 'paragraph' }] },
      },
    ]);
    // filled="f" stroked="f": the frame is text only.
    expect(vector.ops.filter((o) => o.kind === 'polygon')).toHaveLength(1);
  });

  it('nests a group in its own coordinate space', () => {
    const line = draw().vector.ops.find((o) => o.kind === 'line');
    expect(line).toMatchObject({
      x1: 100,
      y1: 75,
      x2: 200,
      y2: 75,
      color: '#ff0000',
    });
  });

  it('places a picture member', () => {
    const img = draw().vector.ops.find(
      (o) => o.kind === 'image',
    ) as VectorImageOp;
    expect(img).toMatchObject({ x: 0, y: 50, width: 20, height: 20 });
    expect(img.src).toMatch(/^data:image\/png;base64,/);
  });
});

describe('a VML group in a document', () => {
  it('imports as one float with its drawing and text frames', async () => {
    const { doc } = await importDocx(await groupDoc());
    const img = doc.child(0).child(0);
    expect(img.type.name).toBe('image');
    expect(img.attrs['width']).toBe(200);
    expect(img.attrs['height']).toBe(100);
    expect(img.attrs['float']).toMatchObject({
      hRel: 'page',
      hOffset: 40,
      vOffset: 8,
      behind: true,
    });
    const frames = img.attrs['frames'] as { blocks: unknown[] }[];
    expect(frames).toHaveLength(1);
    expect(JSON.stringify(frames[0].blocks)).toContain('Label');
  });

  it('writes the group back verbatim, its picture with it', async () => {
    const { doc } = await importDocx(await groupDoc());
    const zip = await JSZip.loadAsync(await exportDocx(doc));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('<v:group id="g1"');
    expect(xml).toContain('<v:shapetype id="_x0000_t202"');
    // The picture's relationship is re-established in the new package.
    const rid = /<v:imagedata r:id="([^"]+)"/.exec(xml)![1];
    const rels = await zip
      .file('word/_rels/document.xml.rels')!
      .async('string');
    const target = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(
      rels,
    )![1];
    expect(await zip.file(`word/${target}`)!.async('base64')).toBe(PNG);
    // And it reads back as the same drawing.
    const again = await importDocx(await exportDocx(doc));
    expect(again.doc.child(0).child(0).attrs['frames']).toHaveLength(1);
  });

  it('is left alone when a picture it references cannot travel', async () => {
    const zip = await JSZip.loadAsync(await groupDoc());
    zip.remove('word/media/image4.png');
    const { doc } = await importDocx(
      await zip.generateAsync({ type: 'uint8array' }),
    );
    expect(doc.child(0).childCount).toBe(0);
  });

  it('draws and carries a lone freeform shape (an arrow)', async () => {
    const arrow = `<w:pict><v:shape id="a1" style="position:absolute;margin-left:30pt;margin-top:6pt;width:30pt;height:7.5pt;z-index:1;mso-position-horizontal-relative:page" coordorigin="100,100" coordsize="40,10" o:spt="100" adj="0,,0" path="m100,100r40,5l100,110xe" fillcolor="black" stroked="f"><v:path arrowok="t"/></v:shape></w:pict>`;
    const { doc } = await importDocx(
      await makeDocx(`<w:p><w:r>${arrow}</w:r></w:p>`),
    );
    const img = doc.child(0).child(0);
    expect(img.attrs['width']).toBe(40);
    expect(img.attrs['height']).toBe(10);
    expect(img.attrs['vector'].ops).toEqual([
      {
        kind: 'polygon',
        points: [
          { x: 0, y: 0 },
          { x: 40, y: 5 },
          { x: 0, y: 10 },
        ],
        fill: '#000000',
      },
    ]);
    const xml = await (await JSZip.loadAsync(await exportDocx(doc)))
      .file('word/document.xml')!
      .async('string');
    expect(xml).toContain('path="m100,100r40,5l100,110xe"');
  });
});
