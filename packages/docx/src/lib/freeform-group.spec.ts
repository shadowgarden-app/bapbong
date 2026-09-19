/**
 * wpg groups of freeform shapes — the diagrams Word's converter makes of
 * legacy drawing canvases (D-2609-PREK: a project network of circles and
 * arrows, 6 groups, 69 custGeom shapes, 9 nested groups). The fixture below
 * is that file's shape in miniature: an anchored, behind-text group inside
 * mc:AlternateContent (Choice wpg + a VML Fallback twin), holding a nested
 * group, a filled triangle (an arrowhead) and an open polyline — one of them
 * with a path that states only its width, as 6 of the file's paths do.
 */
import JSZip from 'jszip';
import type {
  VectorPolygonOp,
  VectorPolylineOp,
} from '@shadow-garden/bapbong-contracts';
import { importDocx } from './docx';
import { exportDocx } from './export';
import { freeformGroupOps } from './freeform-group';
import { child, findDescendant, parseXml, type OoxmlNode } from './ooxml';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPG_NS =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const WPS_NS =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WP14_NS =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const V_NS = 'urn:schemas-microsoft-com:vml';

/** One freeform member: its box in the parent's space, a path, a look. */
export function wsp(opts: {
  off: [number, number];
  ext: [number, number];
  path: string;
  pathW: number;
  pathH?: number;
  fill?: string;
  lineW?: number;
}): string {
  const fill = opts.fill
    ? `<a:solidFill><a:srgbClr val="${opts.fill}"/></a:solidFill>`
    : '<a:noFill/>';
  const ln =
    opts.lineW === undefined
      ? '<a:ln><a:noFill/></a:ln>'
      : `<a:ln w="${opts.lineW}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:round/><a:headEnd/><a:tailEnd/></a:ln>`;
  const h = opts.pathH === undefined ? '' : ` h="${opts.pathH}"`;
  return (
    `<wps:wsp><wps:cNvPr id="7" name="Freeform"/><wps:cNvSpPr/>` +
    `<wps:spPr bwMode="auto"><a:xfrm><a:off x="${opts.off[0]}" y="${opts.off[1]}"/><a:ext cx="${opts.ext[0]}" cy="${opts.ext[1]}"/></a:xfrm>` +
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
    `<a:pathLst><a:path w="${opts.pathW}"${h}>${opts.path}</a:path></a:pathLst></a:custGeom>` +
    `${fill}${ln}</wps:spPr><wps:bodyPr/></wps:wsp>`
  );
}

export const pt = (x: number, y: number) => `<a:pt x="${x}" y="${y}"/>`;

/**
 * The group, 400×200 px on the page (3810000×1905000 EMU), child space
 * 800×400 units at offset (1000, 2000): one child unit = half a px.
 */
export function groupDocXml(members: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:wpg="${WPG_NS}" xmlns:wps="${WPS_NS}" xmlns:mc="${MC_NS}" xmlns:wp14="${WP14_NS}" xmlns:v="${V_NS}" mc:Ignorable="wp14"><w:body>
<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="1A2B3C4D"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>952500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>-190500</wp:posOffset></wp:positionV><wp:extent cx="3810000" cy="1905000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="42" name="Group 42"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${WPG_NS}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/><a:chOff x="1000" y="2000"/><a:chExt cx="800" cy="400"/></a:xfrm></wpg:grpSpPr>${members}</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:group style="position:absolute;margin-left:75pt;margin-top:-15pt;width:300pt;height:150pt" coordorigin="1000,2000" coordsize="800,400"/></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t>A, 3</w:t></w:r></w:p>
</w:body></w:document>`;
}

export async function makeDocx(documentXml: string): Promise<Uint8Array> {
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
  return zip.generateAsync({ type: 'uint8array' });
}

/** A filled arrowhead and an open polyline, both straight in child space. */
const MEMBERS =
  wsp({
    off: [1100, 2100],
    ext: [40, 20],
    pathW: 40,
    pathH: 20,
    path: `<a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(40, 10)}</a:lnTo><a:lnTo>${pt(0, 20)}</a:lnTo><a:close/>`,
    fill: '000000',
    lineW: 3048,
  }) +
  wsp({
    off: [1000, 2000],
    ext: [200, 100],
    pathW: 200,
    path: `<a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(200, 100)}</a:lnTo>`,
    lineW: 12192,
  });

async function importFixture(members = MEMBERS) {
  return importDocx(await makeDocx(groupDocXml(members)));
}

async function documentXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file('word/document.xml')!.async('string');
}

describe('a wpg group of freeform shapes', () => {
  it('imports as one behind-text float the size of the group', async () => {
    const { doc } = await importFixture();
    const para = doc.child(0);
    const img = para.child(0);
    expect(img.type.name).toBe('image');
    expect(img.attrs['width']).toBe(400);
    expect(img.attrs['height']).toBe(200);
    expect(img.attrs['float']).toMatchObject({
      wrap: 'none',
      behind: true,
      hOffset: 100,
      vOffset: -20,
      hRel: 'page',
      vRel: 'paragraph',
    });
    // The label beside it is ordinary text, as in the reported file.
    expect(para.textContent).toBe('A, 3');
  });

  it('paints its members from their own geometry', async () => {
    const { doc } = await importFixture();
    const v = doc.child(0).child(0).attrs['vector'] as {
      width: number;
      height: number;
      ops: { kind: string }[];
    };
    expect(v.width).toBe(400);
    expect(v.height).toBe(200);
    expect(v.ops.map((o) => o.kind)).toEqual(['polygon', 'polyline']);
  });

  it('carries its XML — Choice and Fallback — through a save', async () => {
    const { doc } = await importFixture();
    const out = await documentXmlOf(await exportDocx(doc));
    expect(out).toContain('<mc:AlternateContent');
    expect(out).toContain('<wpg:wgp>');
    expect(out).toContain('<v:group');
    expect(out.match(/<wps:wsp>/g)).toHaveLength(2);
    // The fragment declares the prefixes the generated root does not.
    expect(out).toMatch(/<mc:AlternateContent[^>]*xmlns:wpg="[^"]+"/);
    expect(out).toMatch(/<mc:AlternateContent[^>]*xmlns:v="[^"]+"/);
    expect(out).toMatch(/<mc:AlternateContent[^>]*mc:Ignorable="wp14"/);
    // And the round-trip reads back to the same float.
    const again = await importDocx(await exportDocx(doc));
    expect(again.doc.child(0).child(0).attrs['float']).toMatchObject({
      hOffset: 100,
      vOffset: -20,
      behind: true,
    });
  });

  it('gets a fresh drawing id, not the source one', async () => {
    const { doc } = await importFixture();
    const out = await documentXmlOf(await exportDocx(doc));
    const ids = [...out.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)].map(
      (m) => m[1],
    );
    expect(ids).toHaveLength(1);
    expect(ids[0]).not.toBe('42');
  });

  it('moves its anchor when the reader moved the float', async () => {
    const { doc } = await importFixture();
    const img = doc.child(0).child(0);
    const moved = doc.type.schema.nodes['image'].create({
      ...img.attrs,
      float: { ...(img.attrs['float'] as object), hOffset: 150 },
    });
    const para = doc.child(0).copy(doc.child(0).content.replaceChild(0, moved));
    const out = await documentXmlOf(
      await exportDocx(doc.copy(doc.content.replaceChild(0, para))),
    );
    expect(out).toMatch(
      /<wp:positionH relativeFrom="page"><wp:posOffset>1428750<\/wp:posOffset>/,
    );
    // The vertical offset was not touched.
    expect(out).toMatch(
      /<wp:positionV relativeFrom="paragraph"><wp:posOffset>-190500<\/wp:posOffset>/,
    );
  });

  it('leaves a group holding pictures to the per-picture path', async () => {
    const withPic = MEMBERS.replace(
      '<wps:bodyPr/>',
      '<wps:bodyPr/><a:blip r:embed="rId9"/>',
    );
    const { doc } = await importFixture(withPic);
    expect(doc.child(0).child(0).attrs['rawDrawing'] ?? null).toBeNull();
  });
});

describe('freeformGroupOps (the group → display list)', () => {
  /** The fixture's wgp, parsed, with `members` inside. */
  const wgpOf = (members: string) =>
    findDescendant(parseXml(groupDocXml(members)), 'wpg:wgp')!;
  const black = {
    solidFill: (el: OoxmlNode | undefined) =>
      child(el, 'a:solidFill') ? '#000000' : undefined,
  };
  const round = (p: { x: number; y: number }) => ({
    x: Math.round(p.x * 100) / 100,
    y: Math.round(p.y * 100) / 100,
  });

  it('carries a closed path through shape box and child space to px', () => {
    const { ops, drawn } = freeformGroupOps(wgpOf(MEMBERS), black);
    expect(drawn).toBe(2);
    const tri = ops[0] as VectorPolygonOp;
    expect(tri.kind).toBe('polygon');
    // child unit = 0.5 px; the triangle's box sits at (1100, 2100).
    expect(tri.points.map(round)).toEqual([
      { x: 50, y: 50 },
      { x: 70, y: 55 },
      { x: 50, y: 60 },
    ]);
    expect(tri.fill).toBe('#000000');
    expect(tri.strokeWidth).toBeCloseTo(0.32, 2); // 3048 EMU, unscaled
    expect(tri.join).toBe('round');
  });

  it('reads a path with no h in the shape’s own units on that axis', () => {
    const { ops } = freeformGroupOps(wgpOf(MEMBERS), black);
    const line = ops[1] as VectorPolylineOp;
    expect(line.kind).toBe('polyline');
    expect(line.points.map(round)).toEqual([
      { x: 0, y: 0 },
      { x: 100, y: 50 },
    ]);
    expect(line.strokeWidth).toBeCloseTo(1.28, 2);
  });

  it('composes a nested group’s own child space', () => {
    const nested =
      `<wpg:grpSp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="1400" y="2200"/><a:ext cx="200" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="400" cy="200"/></a:xfrm></wpg:grpSpPr>` +
      wsp({
        off: [0, 0],
        ext: [400, 200],
        pathW: 400,
        pathH: 200,
        path: `<a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(400, 200)}</a:lnTo>`,
        lineW: 9525,
      }) +
      `</wpg:grpSp>`;
    const { ops } = freeformGroupOps(wgpOf(nested), black);
    const line = ops[0] as VectorPolylineOp;
    // (0,0)–(400,200) in the nested space → (1400,2200)–(1600,2300) in the
    // top child space → px at half a px per unit.
    expect(line.points.map(round)).toEqual([
      { x: 200, y: 100 },
      { x: 300, y: 150 },
    ]);
    // The nested group halves its content; the line keeps its 1 px.
    expect(line.strokeWidth).toBeCloseTo(1, 5);
  });

  it('mirrors a flipped shape inside its box', () => {
    const flipped = wsp({
      off: [1000, 2000],
      ext: [200, 100],
      pathW: 200,
      pathH: 100,
      path: `<a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(200, 100)}</a:lnTo>`,
      lineW: 9525,
    }).replace('<a:xfrm>', '<a:xfrm flipH="1">');
    const { ops } = freeformGroupOps(wgpOf(flipped), black);
    expect((ops[0] as VectorPolylineOp).points.map(round)).toEqual([
      { x: 100, y: 0 },
      { x: 0, y: 50 },
    ]);
  });

  it('leaves out a shape it cannot draw, whole', () => {
    const arc = wsp({
      off: [1000, 2000],
      ext: [200, 100],
      pathW: 200,
      pathH: 100,
      path: `<a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(100, 0)}</a:lnTo><a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>`,
      lineW: 9525,
    });
    const { ops, drawn, skipped } = freeformGroupOps(
      wgpOf(arc + MEMBERS),
      black,
    );
    expect(skipped).toBe(1);
    expect(drawn).toBe(2);
    expect(ops).toHaveLength(2);
  });

  it('flattens a cubic curve into points along it', () => {
    const curve = wsp({
      off: [1000, 2000],
      ext: [200, 100],
      pathW: 200,
      pathH: 100,
      path: `<a:moveTo>${pt(0, 0)}</a:moveTo><a:cubicBezTo>${pt(0, 100)}${pt(200, 100)}${pt(200, 0)}</a:cubicBezTo>`,
      lineW: 9525,
    });
    const { ops } = freeformGroupOps(wgpOf(curve), black);
    const pts = (ops[0] as VectorPolylineOp).points;
    expect(pts.length).toBeGreaterThan(8);
    expect(round(pts[pts.length - 1])).toEqual({ x: 100, y: 0 });
    // The curve bulges down to 3/4 of the control height at its middle.
    const mid = pts[Math.floor(pts.length / 2)];
    expect(mid.y).toBeCloseTo(37.5, 0);
  });
});
