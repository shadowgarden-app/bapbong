/**
 * Test-only builders for password-protected documents. Word is the only thing
 * that can create a real one, so nothing can be checked in: these write a CFB
 * container with a genuinely Agile-encrypted package straight from the
 * MS-OFFCRYPTO spec, independently of the reader under test — a mistake in the
 * block keys, the KDF order or the segment IVs shows up as a round-trip
 * failure rather than passing on both sides.
 *
 * Named `.spec-helper.ts` so the library build excludes it (see
 * tsconfig.lib.json) while spec files can import it without dragging another
 * suite's tests into their run.
 */
import { aesCbcEncryptNoPad, bytesToB64, deriveKey } from './crypto-docx.js';

// ── Test fixture builders ───────────────────────────────────────────
// There is no password-protected .docx checked in (creating one needs Word),
// so the tests build one: a CFB container holding a real Agile-encrypted
// package, written from the same MS-OFFCRYPTO spec the reader implements —
// independently enough that a mistake in the block keys, the KDF order or the
// segment IVs shows up as a round-trip failure.

const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const SECTOR = 512;
const MINI = 64;
const MINI_CUTOFF = 4096;

/** Build a compound file holding the named streams. Streams under the mini
 *  cutoff land in the mini stream (as Word writes EncryptionInfo), larger ones
 *  in ordinary FAT sectors — so both read paths get exercised. */
export function buildCfb(
  streams: { name: string; data: Uint8Array }[],
): Uint8Array {
  const small = streams.filter((s) => s.data.length < MINI_CUTOFF);
  const large = streams.filter((s) => s.data.length >= MINI_CUTOFF);

  // Mini stream: small streams concatenated, each starting on a mini sector.
  const miniChunks: { name: string; start: number; size: number }[] = [];
  let miniLen = 0;
  for (const s of small) {
    miniChunks.push({
      name: s.name,
      start: miniLen / MINI,
      size: s.data.length,
    });
    miniLen += Math.ceil(s.data.length / MINI) * MINI;
  }
  const miniStream = new Uint8Array(miniLen);
  {
    let w = 0;
    for (const s of small) {
      miniStream.set(s.data, w);
      w += Math.ceil(s.data.length / MINI) * MINI;
    }
  }

  const miniStreamSectors = Math.ceil(miniStream.length / SECTOR) || 0;
  const largeSectors = large.map((s) => Math.ceil(s.data.length / SECTOR));
  // Layout: 0 = FAT, 1 = directory, 2 = mini-FAT, 3… = mini stream, then each
  // large stream in turn.
  const MINI_STREAM_START = 3;
  const totalSectors =
    3 + miniStreamSectors + largeSectors.reduce((a, b) => a + b, 0);

  const fat = new Uint32Array(SECTOR / 4).fill(FREESECT);
  fat[0] = FATSECT;
  fat[1] = ENDOFCHAIN;
  fat[2] = ENDOFCHAIN;
  const chainRange = (start: number, count: number) => {
    for (let i = 0; i < count; i++) {
      fat[start + i] = i === count - 1 ? ENDOFCHAIN : start + i + 1;
    }
  };
  chainRange(MINI_STREAM_START, miniStreamSectors);
  const largeStarts: number[] = [];
  let next = MINI_STREAM_START + miniStreamSectors;
  for (const n of largeSectors) {
    largeStarts.push(next);
    chainRange(next, n);
    next += n;
  }

  const miniFat = new Uint32Array(SECTOR / 4).fill(FREESECT);
  for (const c of miniChunks) {
    const n = Math.ceil(c.size / MINI);
    for (let i = 0; i < n; i++) {
      miniFat[c.start + i] = i === n - 1 ? ENDOFCHAIN : c.start + i + 1;
    }
  }

  const out = new Uint8Array((1 + totalSectors) * SECTOR);
  const view = new DataView(out.buffer);
  const sectorOffset = (n: number) => SECTOR + n * SECTOR;

  // Header
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
  view.setUint16(24, 0x003e, true); // minor version
  view.setUint16(26, 0x0003, true); // major version (512-byte sectors)
  view.setUint16(28, 0xfffe, true); // little-endian marker
  view.setUint16(30, 9, true); // sector shift → 512
  view.setUint16(32, 6, true); // mini sector shift → 64
  view.setUint32(44, 1, true); // FAT sector count
  view.setUint32(48, 1, true); // first directory sector
  view.setUint32(56, MINI_CUTOFF, true);
  view.setUint32(60, 2, true); // first mini-FAT sector
  view.setUint32(64, 1, true); // mini-FAT sector count
  view.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector
  view.setUint32(72, 0, true); // DIFAT sector count
  for (let i = 0; i < 109; i++) view.setUint32(76 + i * 4, FREESECT, true);
  view.setUint32(76, 0, true); // DIFAT[0] → FAT sector 0

  for (let i = 0; i < fat.length; i++)
    view.setUint32(sectorOffset(0) + i * 4, fat[i], true);
  for (let i = 0; i < miniFat.length; i++)
    view.setUint32(sectorOffset(2) + i * 4, miniFat[i], true);
  out.set(miniStream, sectorOffset(MINI_STREAM_START));
  large.forEach((s, i) => out.set(s.data, sectorOffset(largeStarts[i])));

  // Directory: root + one entry per stream.
  const dirBase = sectorOffset(1);
  const writeEntry = (
    slot: number,
    name: string,
    type: number,
    start: number,
    size: number,
  ) => {
    const off = dirBase + slot * 128;
    for (let i = 0; i < name.length; i++)
      view.setUint16(off + i * 2, name.charCodeAt(i), true);
    view.setUint16(off + name.length * 2, 0, true);
    view.setUint16(off + 64, (name.length + 1) * 2, true);
    view.setUint8(off + 66, type);
    view.setUint8(off + 67, 1); // black
    view.setUint32(off + 68, FREESECT, true); // left sibling
    view.setUint32(off + 72, FREESECT, true); // right sibling
    view.setUint32(off + 76, FREESECT, true); // child
    view.setUint32(off + 116, start, true);
    view.setUint32(off + 120, size, true);
  };
  writeEntry(0, 'Root Entry', 5, MINI_STREAM_START, miniStream.length);
  const placed: { name: string; slot: number }[] = [];
  let slot = 1;
  for (const c of miniChunks) {
    writeEntry(slot, c.name, 2, c.start, c.size);
    placed.push({ name: c.name, slot: slot++ });
  }
  large.forEach((s, i) => {
    writeEntry(slot, s.name, 2, largeStarts[i], s.data.length);
    placed.push({ name: s.name, slot: slot++ });
  });
  // Directory TREE links (MS-CFB §2.6.4: siblings ordered by name length,
  // then uppercased name). A reader walks from the root's child pointer;
  // without these the container looks empty to everything but a linear scan.
  placed.sort((a, b) =>
    a.name.length !== b.name.length
      ? a.name.length - b.name.length
      : a.name.toUpperCase() < b.name.toUpperCase()
        ? -1
        : 1,
  );
  if (placed.length) {
    view.setUint32(dirBase + 0 * 128 + 76, placed[0].slot, true); // root child
    for (let i = 0; i + 1 < placed.length; i++)
      view.setUint32(
        dirBase + placed[i].slot * 128 + 72,
        placed[i + 1].slot,
        true,
      );
  }
  return out;
}

const rand = (n: number) => crypto.getRandomValues(new Uint8Array(n));

/** Encrypt `plain` (a .docx zip's bytes) into a password-protected container.
 *  `spinCount` is kept tiny in tests — the KDF cost is the point in
 *  production, not here. */
export async function buildEncryptedDocx(
  plain: Uint8Array,
  password: string,
  spinCount = 50,
): Promise<Uint8Array> {
  const keySalt = rand(16);
  const pwSalt = rand(16);
  const secretKey = rand(32);
  const verifierInput = rand(16);

  const spec = {
    saltValue: pwSalt,
    blockSize: 16,
    keyBits: 256,
    hashSize: 64,
    hashAlgorithm: 'SHA512',
    spinCount,
    encryptedVerifierHashInput: new Uint8Array(),
    encryptedVerifierHashValue: new Uint8Array(),
    encryptedKeyValue: new Uint8Array(),
  };
  const BK_IN = new Uint8Array([
    0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79,
  ]);
  const BK_VAL = new Uint8Array([
    0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e,
  ]);
  const BK_SECRET = new Uint8Array([
    0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6,
  ]);

  const encVerifierInput = await aesCbcEncryptNoPad(
    await deriveKey(password, spec, BK_IN),
    pwSalt,
    verifierInput,
  );
  const verifierHash = new Uint8Array(
    await crypto.subtle.digest('SHA-512', verifierInput),
  );
  const encVerifierHash = await aesCbcEncryptNoPad(
    await deriveKey(password, spec, BK_VAL),
    pwSalt,
    verifierHash,
  );
  const encSecretKey = await aesCbcEncryptNoPad(
    await deriveKey(password, spec, BK_SECRET),
    pwSalt,
    secretKey,
  );

  // Package: 4096-byte segments, each with its own salt-derived IV.
  const padded = new Uint8Array(Math.ceil(plain.length / 16) * 16);
  padded.set(plain);
  const body = new Uint8Array(padded.length);
  for (let i = 0, w = 0; w < padded.length; i++) {
    const chunk = padded.subarray(i * 4096, (i + 1) * 4096);
    if (!chunk.length) break;
    const ivFull = new Uint8Array(
      await crypto.subtle.digest(
        'SHA-512',
        (() => {
          const b = new Uint8Array(keySalt.length + 4);
          b.set(keySalt);
          new DataView(b.buffer).setUint32(keySalt.length, i, true);
          return b;
        })(),
      ),
    );
    body.set(
      await aesCbcEncryptNoPad(secretKey, ivFull.subarray(0, 16), chunk),
      w,
    );
    w += chunk.length;
  }
  const encryptedPackage = new Uint8Array(8 + body.length);
  new DataView(encryptedPackage.buffer).setUint32(0, plain.length, true);
  encryptedPackage.set(body, 8);

  // Real files carry dataIntegrity; readers expect it (see buildDataIntegrity).
  const hmacKey = rand(64);
  const ivFor = async (bk: number[]) => {
    const b = new Uint8Array(keySalt.length + bk.length);
    b.set(keySalt);
    b.set(bk, keySalt.length);
    return new Uint8Array(await crypto.subtle.digest('SHA-512', b)).subarray(
      0,
      16,
    );
  };
  const encHmacKey = await aesCbcEncryptNoPad(
    secretKey,
    await ivFor([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
    hmacKey,
  );
  const macKey = await crypto.subtle.importKey(
    'raw',
    hmacKey,
    { name: 'HMAC', hash: 'SHA-512' },
    false,
    ['sign'],
  );
  const mac = new Uint8Array(
    await crypto.subtle.sign('HMAC', macKey, encryptedPackage),
  );
  const encHmacValue = await aesCbcEncryptNoPad(
    secretKey,
    await ivFor([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
    mac,
  );

  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ` +
    `xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" ` +
    `cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" ` +
    `saltValue="${bytesToB64(keySalt)}"/>` +
    `<dataIntegrity encryptedHmacKey="${bytesToB64(encHmacKey)}" ` +
    `encryptedHmacValue="${bytesToB64(encHmacValue)}"/>` +
    `<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16" keyBits="256" ` +
    `hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" ` +
    `saltValue="${bytesToB64(pwSalt)}" ` +
    `encryptedVerifierHashInput="${bytesToB64(encVerifierInput)}" ` +
    `encryptedVerifierHashValue="${bytesToB64(encVerifierHash)}" ` +
    `encryptedKeyValue="${bytesToB64(encSecretKey)}"/>` +
    `</keyEncryptor></keyEncryptors></encryption>`;
  const xmlBytes = new TextEncoder().encode(xml);
  const info = new Uint8Array(8 + xmlBytes.length);
  new DataView(info.buffer).setUint16(0, 4, true); // major
  new DataView(info.buffer).setUint16(2, 4, true); // minor
  new DataView(info.buffer).setUint32(4, 0x40, true); // flags
  info.set(xmlBytes, 8);

  return buildCfb([
    { name: 'EncryptionInfo', data: info },
    { name: 'EncryptedPackage', data: encryptedPackage },
  ]);
}
