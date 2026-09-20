import { describe, expect, it } from 'vitest';
import { EditorState, type Transaction } from 'prosemirror-state';
import { schema } from '@shadow-garden/bapbong-headless';
import { ContentError, type ImageSource } from './contract.js';
import { base64, dataUrl, sniffImage } from './image-bytes.js';
import { PmDocSession, type PmSessionHost } from './pm-session.js';

// ── fixtures ─────────────────────────────────────────────────────────

/** A PNG header with the given size. Only the first 24 bytes are ever read
 *  (signature + IHDR), so the pixels are left out. */
function png(width: number, height: number): Uint8Array {
  const b = new Uint8Array(24);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  b.set([0, 0, 0, 13], 8);
  b.set([0x49, 0x48, 0x44, 0x52], 12); // "IHDR"
  new DataView(b.buffer).setUint32(16, width);
  new DataView(b.buffer).setUint32(20, height);
  return b;
}

/** A JPEG whose size sits behind a comment segment — the sniffer must skip
 *  that by its length rather than stopping at the first 0xFF it meets. */
function jpeg(width: number, height: number): Uint8Array {
  const comment = [0xff, 0xfe, 0x00, 0x06, 0x68, 0x65, 0x6c, 0x6f];
  const sof = [0xff, 0xc0, 0x00, 0x11, 0x08];
  const b = [0xff, 0xd8, ...comment, ...sof];
  b.push(
    (height >> 8) & 0xff,
    height & 0xff,
    (width >> 8) & 0xff,
    width & 0xff,
  );
  b.push(...new Array(8).fill(0));
  return new Uint8Array(b);
}

function gif(width: number, height: number): Uint8Array {
  const b = new Uint8Array(13);
  b.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]); // GIF89a
  new DataView(b.buffer).setUint16(6, width, true);
  new DataView(b.buffer).setUint16(8, height, true);
  return b;
}

function bmp(width: number, height: number): Uint8Array {
  const b = new Uint8Array(30);
  b.set([0x42, 0x4d]);
  const view = new DataView(b.buffer);
  view.setInt32(18, width, true);
  view.setInt32(22, -height, true); // bottom-up: a negative height
  return b;
}

/** A session over a document built here, with a host that answers with
 *  whatever picture the test wants. */
function session(
  doc: ReturnType<typeof schema.node>,
  images?: {
    bytes: (source: ImageSource) => Uint8Array;
    sources?: ('path' | 'attachment' | 'svg')[];
  },
) {
  let state = EditorState.create({ doc });
  let version = 1;
  const host: PmSessionHost = {
    getState: () => state,
    apply: (tr: Transaction) => {
      state = state.apply(tr);
      version++;
    },
    getVersion: () => `v${version}`,
    meta: () => ({ name: 'test.docx' }),
    save: async () => undefined,
    ...(images
      ? {
          readImage: async (source: ImageSource) => ({
            bytes: images.bytes(source),
          }),
          imageSources: images.sources ?? ['path', 'attachment', 'svg'],
        }
      : {}),
  };
  return { s: new PmDocSession(host), doc: () => state.doc };
}

const para = (text: string) =>
  schema.node('paragraph', null, [schema.text(text)]);

// ── the bytes ────────────────────────────────────────────────────────

describe('sniffImage', () => {
  it('reads the size out of each format Word embeds', () => {
    expect(sniffImage(png(640, 480))).toEqual({
      mediaType: 'image/png',
      width: 640,
      height: 480,
    });
    expect(sniffImage(jpeg(300, 200))).toEqual({
      mediaType: 'image/jpeg',
      width: 300,
      height: 200,
    });
    expect(sniffImage(gif(12, 34))).toEqual({
      mediaType: 'image/gif',
      width: 12,
      height: 34,
    });
    // A bottom-up BMP declares a negative height; the picture is not upside
    // down, that is just where its first row lives.
    expect(sniffImage(bmp(50, 25))).toEqual({
      mediaType: 'image/bmp',
      width: 50,
      height: 25,
    });
  });

  it('refuses what Word would not take, and says what it does take', () => {
    expect(() => sniffImage(new Uint8Array())).toThrow(ContentError);
    expect(() => sniffImage(new Uint8Array([1, 2, 3, 4]))).toThrow(
      /not an image in a format Word reads.*PNG, JPEG, GIF, BMP/s,
    );
    const webp = new Uint8Array(12);
    webp.set([0x52, 0x49, 0x46, 0x46], 0); // RIFF
    webp.set([0x57, 0x45, 0x42, 0x50], 8); // WEBP
    expect(() => sniffImage(webp)).toThrow(/WebP/);
    // The declared type is only there to make the refusal specific.
    expect(() => sniffImage(new Uint8Array([1]), 'image/heic')).toThrow(
      /announced image\/heic/,
    );
  });

  it('encodes bytes the way a data URL needs, padding included', () => {
    expect(base64(new Uint8Array([0]))).toBe('AA==');
    expect(base64(new Uint8Array([0, 0]))).toBe('AAA=');
    expect(base64(new Uint8Array([77, 97, 110]))).toBe('TWFu');
    expect(dataUrl(new Uint8Array([77, 97, 110]), 'image/png')).toBe(
      'data:image/png;base64,TWFu',
    );
  });
});

// ── the tools ────────────────────────────────────────────────────────

describe('insertImage', () => {
  it('puts the picture in its own paragraph at the anchor, at its own size', async () => {
    const { s, doc } = session(
      schema.node('doc', null, [para('Before'), para('After')]),
      { bytes: () => png(200, 100) },
    );
    const out = await s.insertImage(
      { kind: 'path', path: '/tmp/diagram.png' },
      { position: 'before', text: 'After' },
      { alt: 'Network diagram' },
    );
    expect(out).toMatchObject({ width: 200, height: 100 });

    const inserted = doc().child(1);
    expect(inserted.type.name).toBe('paragraph');
    expect(inserted.child(0).attrs['alt']).toBe('Network diagram');
    expect(inserted.child(0).attrs['src']).toMatch(/^data:image\/png;base64,/);
    expect(doc().child(2).textContent).toBe('After');
  });

  it('shrinks a picture wider than the text, and keeps its proportions', async () => {
    const { s, doc } = session(schema.node('doc', null, [para('Body')]), {
      bytes: () => png(4000, 1000),
    });
    // A4 with 1in margins = 602px of text.
    await s.insertImage(
      { kind: 'svg', svg: '<svg/>' },
      { position: 'document_end' },
    );
    expect(doc().lastChild!.child(0).attrs).toMatchObject({
      width: 602,
      height: 151,
    });

    const asked = await s.insertImage(
      { kind: 'svg', svg: '<svg/>' },
      { position: 'document_end' },
      { width: 300 },
    );
    expect(asked).toMatchObject({ width: 300, height: 75 });
  });

  it('treats a rasterizer\u2019s extra pixels as sharpness, not size', async () => {
    // The window draws SVG at 2\u00d7 for print; the box stays the size the
    // agent drew, and the picture simply has twice the pixels.
    let state = EditorState.create({
      doc: schema.node('doc', null, [para('Body')]),
    });
    const host: PmSessionHost = {
      getState: () => state,
      apply: (tr) => {
        state = state.apply(tr);
      },
      getVersion: () => 'v1',
      meta: () => ({}),
      save: async () => undefined,
      readImage: async () => ({ bytes: png(800, 400), scale: 2 }),
      imageSources: ['svg'],
    };
    const out = await new PmDocSession(host).insertImage(
      { kind: 'svg', svg: '<svg/>' },
      { position: 'document_end' },
    );
    expect(out).toMatchObject({ width: 400, height: 200 });
  });

  it('is refused, by name, when the host cannot fetch that kind', async () => {
    const { s } = session(schema.node('doc', null, [para('Body')]), {
      bytes: () => png(10, 10),
      sources: ['attachment'],
    });
    expect(s.capabilities.images).toEqual(['attachment']);
    await expect(
      s.insertImage(
        { kind: 'path', path: '/tmp/x.png' },
        { position: 'document_end' },
      ),
    ).rejects.toThrow(/cannot take a picture from path.*accepts: attachment/s);

    const without = session(schema.node('doc', null, [para('Body')]));
    expect(without.s.capabilities.images).toEqual([]);
    await expect(
      without.s.insertImage(
        { kind: 'svg', svg: '<svg/>' },
        { position: 'document_end' },
      ),
    ).rejects.toThrow(/no new pictures at all/);
  });
});

describe('replaceImage', () => {
  /** A floating drawing carried verbatim from the file — what a diagram
   *  imported as a group of freeform shapes looks like. */
  const drawing = () =>
    schema.node('doc', null, [
      para('Intro'),
      schema.node('paragraph', null, [
        schema.node('image', {
          src: '',
          alt: 'So do mang',
          width: 400,
          height: 200,
          float: { wrap: 'none', hOffset: 120, vOffset: 40 },
          rawDrawing: {
            xml: '<mc:AlternateContent/>',
            float: { wrap: 'none' },
          },
        }),
      ]),
    ]);

  it('keeps the box and the anchor, and lets go of the old drawing', async () => {
    const { s, doc } = session(drawing(), { bytes: () => png(800, 200) });
    const out = await s.replaceImage(1, 0, { kind: 'svg', svg: '<svg/>' });

    // The width it took over, the height from the NEW picture's proportions.
    expect(out).toMatchObject({ width: 400, height: 100 });
    const img = doc().child(1).child(0);
    expect(img.attrs['float']).toEqual({
      wrap: 'none',
      hOffset: 120,
      vOffset: 40,
    });
    expect(img.attrs['alt']).toBe('So do mang');
    // Without this the exporter would write the original drawing back out and
    // the replacement would never reach the file.
    expect(img.attrs['rawDrawing']).toBeNull();
    expect(img.attrs['src']).toMatch(/^data:image\/png;base64,/);
  });

  it('takes a width and an alt when the caller wants them', async () => {
    const { s, doc } = session(drawing(), { bytes: () => png(800, 200) });
    await s.replaceImage(
      1,
      0,
      { kind: 'path', path: '/a.png' },
      {
        width: 200,
        alt: 'Redrawn',
      },
    );
    expect(doc().child(1).child(0).attrs).toMatchObject({
      width: 200,
      height: 50,
      alt: 'Redrawn',
    });
  });

  it('says which picture it could not find', async () => {
    const { s } = session(drawing(), { bytes: () => png(10, 10) });
    await expect(
      s.replaceImage(0, 0, { kind: 'path', path: '/a.png' }),
    ).rejects.toThrow(/Block 0 has no images/);
    await expect(
      s.replaceImage(1, 3, { kind: 'path', path: '/a.png' }),
    ).rejects.toThrow(/imageIndex 3 is out of range — block 1 has 1 image/);
  });
});

describe('deleteImage', () => {
  it('removes the picture and leaves its paragraph standing', async () => {
    const { s, doc } = session(
      schema.node('doc', null, [
        schema.node('paragraph', null, [
          schema.text('Figure: '),
          schema.node('image', {
            src: 'data:image/png;base64,AA==',
            width: 10,
            height: 10,
          }),
          schema.node('image', {
            src: 'data:image/png;base64,AB==',
            width: 20,
            height: 20,
          }),
        ]),
      ]),
    );
    await s.deleteImage(0, 0);
    const block = doc().child(0);
    expect(block.textContent).toBe('Figure: ');
    expect(block.childCount).toBe(2);
    expect(block.child(1).attrs['width']).toBe(20);
  });
});

describe('get_document on pictures', () => {
  it('names what each box holds, and whether it floats', async () => {
    const { s } = session(
      schema.node('doc', null, [
        schema.node('paragraph', null, [
          schema.node('image', {
            src: 'data:image/png;base64,AA==',
            width: 10,
            height: 10,
          }),
          schema.node('image', {
            src: '',
            width: 10,
            height: 10,
            shape: { kind: 'rect' },
          }),
          schema.node('image', {
            src: '',
            width: 10,
            height: 10,
            rawDrawing: { xml: '<x/>' },
            float: { wrap: 'none' },
          }),
          schema.node('image', {
            src: '',
            width: 10,
            height: 10,
            oleProgId: 'Equation.DSMT4',
          }),
        ]),
      ]),
    );
    const { blocks } = await s.snapshot();
    expect(blocks[0].images?.map((i) => i.kind)).toEqual([
      'bitmap',
      'shape',
      'drawing',
      'equation',
    ]);
    expect(blocks[0].images?.map((i) => i.float ?? false)).toEqual([
      false,
      false,
      true,
      false,
    ]);
  });
});
