/**
 * What an image's bytes say about themselves: format and intrinsic size.
 *
 * The desktop runs agent edits headlessly in Bun — no DOM, no
 * `createImageBitmap` — and the same code has to work in the WebView and in
 * tests, so the dimensions are read out of the file's own header instead.
 * Four formats, which is what Word embeds without converting anything.
 *
 * Nothing here fetches: hosts hand over bytes (a file they were allowed to
 * read, an attachment they downloaded, an SVG they rasterized) and this
 * turns them into the `data:` URL the image node's `src` has always been —
 * the same representation the DOCX importer and the paste path produce, so
 * layout, painting and export need no new case.
 */
import { ContentError } from './contract.js';

export interface ImageInfo {
  mediaType: string;
  /** Intrinsic pixels. */
  width: number;
  height: number;
}

const PNG_SIG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

const be32 = (b: Uint8Array, at: number) =>
  ((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]) >>> 0;
const be16 = (b: Uint8Array, at: number) => (b[at] << 8) | b[at + 1];
const le16 = (b: Uint8Array, at: number) => b[at] | (b[at + 1] << 8);
const le32 = (b: Uint8Array, at: number) =>
  b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24) | 0;

const starts = (b: Uint8Array, sig: number[]) =>
  b.length >= sig.length && sig.every((v, i) => b[i] === v);

const ascii = (b: Uint8Array, at: number, len: number) =>
  String.fromCharCode(...b.subarray(at, at + len));

/** JPEG: walk the marker segments to the frame header that carries the size.
 *  Everything else (EXIF, comments, quantization tables) is skipped by its
 *  own length, so a photo with a thumbnail does not report the thumbnail. */
function jpegSize(b: Uint8Array): { width: number; height: number } | null {
  let i = 2;
  while (i + 9 < b.length) {
    if (b[i] !== 0xff) {
      i++; // fill byte or padding — resync on the next marker
      continue;
    }
    const marker = b[i + 1];
    if (
      marker === 0xd8 ||
      marker === 0x01 ||
      (marker >= 0xd0 && marker <= 0xd7)
    ) {
      i += 2; // standalone markers carry no length
      continue;
    }
    const len = be16(b, i + 2);
    if (len < 2) return null;
    // SOF0-3, 5-7, 9-11, 13-15: a start-of-frame. C4/C8/CC are tables, not frames.
    const sof =
      marker >= 0xc0 &&
      marker <= 0xcf &&
      marker !== 0xc4 &&
      marker !== 0xc8 &&
      marker !== 0xcc;
    if (sof) return { height: be16(b, i + 5), width: be16(b, i + 7) };
    if (marker === 0xda) return null; // scan data begins; no frame header found
    i += 2 + len;
  }
  return null;
}

/** Format + intrinsic size from the bytes themselves. `declared` (a media
 *  type the host was told, e.g. an attachment's) is only used to make the
 *  refusal specific — the bytes decide. */
export function sniffImage(bytes: Uint8Array, declared?: string): ImageInfo {
  // Annotated on the variable, not just the arrow: TypeScript only lets a
  // `never`-returning call end a code path when the binding says so.
  const refuse: (why: string) => never = (why) => {
    throw new ContentError(
      `${why} Supported: PNG, JPEG, GIF, BMP` +
        (declared ? ` (this one announced ${declared}).` : '.'),
    );
  };

  if (starts(bytes, PNG_SIG)) {
    // IHDR is always the first chunk: 8 signature + 4 length + 4 type.
    if (bytes.length < 24 || ascii(bytes, 12, 4) !== 'IHDR')
      refuse('This PNG has no image header.');
    return {
      mediaType: 'image/png',
      width: be32(bytes, 16),
      height: be32(bytes, 20),
    };
  }

  if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    const size = jpegSize(bytes);
    if (!size) refuse('This JPEG has no frame header.');
    return { mediaType: 'image/jpeg', ...size! };
  }

  if (bytes.length >= 10 && ascii(bytes, 0, 4) === 'GIF8')
    return {
      mediaType: 'image/gif',
      width: le16(bytes, 6),
      height: le16(bytes, 8),
    };

  if (bytes.length >= 26 && bytes[0] === 0x42 && bytes[1] === 0x4d)
    return {
      mediaType: 'image/bmp',
      width: Math.abs(le32(bytes, 18)),
      // Bottom-up bitmaps declare a negative height.
      height: Math.abs(le32(bytes, 22)),
    };

  if (
    bytes.length >= 12 &&
    ascii(bytes, 0, 4) === 'RIFF' &&
    ascii(bytes, 8, 4) === 'WEBP'
  )
    refuse('WebP is not an image format Word embeds.');

  refuse(
    bytes.length === 0
      ? 'That image is empty.'
      : 'Those bytes are not an image in a format Word reads.',
  );
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/** base64 without `btoa`/`Buffer` — this tier runs in a WebView, in Bun and
 *  in plain Node, and one encoder means one behaviour. */
export function base64(bytes: Uint8Array): string {
  let out = '';
  let i = 0;
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
    out +=
      B64[(n >> 18) & 63] +
      B64[(n >> 12) & 63] +
      B64[(n >> 6) & 63] +
      B64[n & 63];
  }
  const left = bytes.length - i;
  if (left === 1) {
    const n = bytes[i] << 16;
    out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63] + '==';
  } else if (left === 2) {
    const n = (bytes[i] << 16) | (bytes[i + 1] << 8);
    out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63] + B64[(n >> 6) & 63] + '=';
  }
  return out;
}

/** The `data:` URL an image node's `src` holds. */
export function dataUrl(bytes: Uint8Array, mediaType: string): string {
  return `data:${mediaType};base64,${base64(bytes)}`;
}
