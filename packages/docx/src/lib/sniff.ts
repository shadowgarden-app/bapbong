/**
 * Cheap content sniffing for "is this really a .docx?".
 *
 * A file with a .docx name routinely isn't one: a PDF or web page renamed, a
 * Word 97–2003 .doc, a password-protected document (encrypted OOXML is an OLE
 * container, not a zip), a truncated download. JSZip's "can't find end of
 * central directory" tells the user nothing — these few magic-byte checks
 * name the actual problem, and are cheap enough for a shell to run BEFORE
 * committing a tab slot to the bytes.
 *
 * `sniffDocx` reads at most a few hundred bytes of header (plus one bounded
 * scan of an OLE file for its stream names) and never parses the container.
 * 'zip' means "plausibly a docx — importing is the only way to know more";
 * every other verdict is a certainty that importDocx would fail.
 */

export type DocxSniff =
  | 'zip' // PK header — the only candidate worth importing
  | 'empty'
  | 'pdf'
  | 'legacy-doc' // OLE container holding a WordDocument stream (Word 97–2003)
  | 'encrypted' // OLE container holding EncryptionInfo (password-protected OOXML)
  | 'ole' // OLE container that is neither (some other Office-era file)
  | 'rtf'
  | 'html'
  | 'unknown';

/** Machine-readable cause carried by {@link DocxImportError}. Supersets
 *  {@link DocxSniff} with the failures only visible after opening the zip. */
export type DocxImportErrorKind =
  | Exclude<DocxSniff, 'zip'>
  | 'corrupt-zip' // PK header but the archive won't open (truncated/partial)
  | 'xlsx' // a real zip, but it's an Excel workbook
  | 'pptx' // … or a PowerPoint deck
  | 'no-document'; // a zip without word/document.xml (damaged docx)

/** importDocx failure with a classified cause. `message` stays human-readable
 *  (and is what a shell that doesn't switch on `kind` will show); `detail`
 *  carries the underlying library error, when one exists, for diagnostics. */
export class DocxImportError extends Error {
  constructor(
    readonly kind: DocxImportErrorKind,
    message: string,
    readonly detail?: string,
  ) {
    super(message);
    this.name = 'DocxImportError';
  }
}

/** Bytes of `ascii` encoded as UTF-16LE — how OLE directory entries store
 *  stream names ("EncryptionInfo", "WordDocument"). */
function utf16leBytes(ascii: string): Uint8Array {
  const out = new Uint8Array(ascii.length * 2);
  for (let i = 0; i < ascii.length; i++) out[i * 2] = ascii.charCodeAt(i);
  return out;
}

function indexOfBytes(haystack: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = 0; i + needle.length <= haystack.length; i++) {
    for (let j = 0; j < needle.length; j++) {
      if (haystack[i + j] !== needle[j]) continue outer;
    }
    return i;
  }
  return -1;
}

const ENCRYPTION_INFO = utf16leBytes('EncryptionInfo');
const WORD_DOCUMENT = utf16leBytes('WordDocument');

/** How much of an OLE file to scan for stream names. Directory entries live
 *  in the FAT-addressed sectors; for the tiny files Word writes they sit well
 *  within the first 64 KB — a full scan of a mis-named 2 GB video is not. */
const OLE_SCAN_LIMIT = 64 * 1024;

const starts = (b: Uint8Array, sig: number[]): boolean =>
  sig.every((v, i) => b[i] === v);

export function sniffDocx(input: ArrayBuffer | Uint8Array): DocxSniff {
  const b = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (b.length === 0) return 'empty';
  if (starts(b, [0x50, 0x4b])) return 'zip'; // "PK"
  if (starts(b, [0x25, 0x50, 0x44, 0x46, 0x2d])) return 'pdf'; // "%PDF-"
  if (starts(b, [0x7b, 0x5c, 0x72, 0x74, 0x66])) return 'rtf'; // "{\rtf"
  if (starts(b, [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])) {
    // OLE compound file. Encrypted OOXML and legacy .doc are BOTH this
    // container — the stream names tell them apart without an OLE parser.
    const head = b.subarray(0, OLE_SCAN_LIMIT);
    if (indexOfBytes(head, ENCRYPTION_INFO) >= 0) return 'encrypted';
    if (indexOfBytes(head, WORD_DOCUMENT) >= 0) return 'legacy-doc';
    return 'ole';
  }
  // Web pages saved with a .docx name (a common "export" from webmail): match
  // an HTML opener within leading whitespace/BOM.
  const leadText = new TextDecoder('utf-8', { fatal: false })
    .decode(b.subarray(0, 512))
    .replace(/^﻿/, '')
    .trimStart()
    .slice(0, 15)
    .toLowerCase();
  if (leadText.startsWith('<!doctype') || leadText.startsWith('<html'))
    return 'html';
  return 'unknown';
}

/** The default English message for each failure — what shells show verbatim
 *  unless they localize by `kind`. One sentence of WHAT, one of what to DO. */
export const IMPORT_ERROR_MESSAGES: Record<DocxImportErrorKind, string> = {
  empty: 'The file is empty — nothing was ever written to it.',
  pdf: 'This file is a PDF renamed to .docx. Renaming it back to .pdf should open it.',
  'legacy-doc':
    'This is a Word 97–2003 (.doc) file. Open it in Word and save it as .docx to edit it here.',
  encrypted:
    'This document is password-protected. Remove the password in Word, then reopen it.',
  ole: 'This is an older Office file, not a .docx document.',
  rtf: 'This is an RTF document renamed to .docx. Renaming it back to .rtf should open it.',
  html: 'This file contains a web page (HTML), not a Word document.',
  unknown: 'The contents are not a Word document.',
  'corrupt-zip':
    'The file is damaged or incomplete — it may be a partial download or copy. Try getting the file again.',
  xlsx: 'This is an Excel workbook renamed to .docx. Renaming it back to .xlsx should open it.',
  pptx: 'This is a PowerPoint presentation renamed to .docx. Renaming it back to .pptx should open it.',
  'no-document':
    'The archive opened but has no document content — the file is likely damaged.',
};

/** A classified import failure for a sniff verdict that rules importing out. */
export function errorForSniff(kind: Exclude<DocxSniff, 'zip'>): DocxImportError {
  return new DocxImportError(kind, IMPORT_ERROR_MESSAGES[kind]);
}
