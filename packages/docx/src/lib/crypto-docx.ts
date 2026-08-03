/**
 * Decrypting password-protected Office documents (MS-OFFCRYPTO "Agile
 * Encryption", what Word 2010+ writes).
 *
 * The file is an OLE container (see cfb.ts) holding two streams:
 *   EncryptionInfo    — version header + XML describing salts, cipher and KDF
 *   EncryptedPackage  — uint64 plaintext size, then AES-CBC ciphertext of the
 *                       ordinary .docx zip, in independently-IV'd segments
 *
 * The password is turned into a key by a deliberately slow KDF (`spinCount`
 * hash rounds — 100 000 in Word's default), which is why callers should run
 * this off the main thread. Password correctness is checked LOCALLY against a
 * stored verifier hash, so a wrong password is known before any bulk work.
 *
 * Everything runs on WebCrypto (SubtleCrypto), available in browsers, workers
 * and Node ≥ 16 — no dependency, no key material outside this module.
 */

/** Block keys from MS-OFFCRYPTO §2.3.4.12 — fixed constants that separate the
 *  KDF outputs used for different purposes. */
const BLOCK_KEY_VERIFIER_INPUT = new Uint8Array([
  0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79,
]);
const BLOCK_KEY_VERIFIER_VALUE = new Uint8Array([
  0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e,
]);
const BLOCK_KEY_SECRET = new Uint8Array([
  0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6,
]);

/** Bulk data is encrypted in independently-IV'd chunks of this size. */
const SEGMENT_SIZE = 4096;

export class DocxCryptoError extends Error {}
/** The password did not match the stored verifier. */
export class WrongPasswordError extends DocxCryptoError {
  constructor() {
    super('Incorrect password.');
  }
}

interface KeyDataSpec {
  saltValue: Uint8Array;
  blockSize: number;
  keyBits: number;
  hashSize: number;
  hashAlgorithm: string;
}

interface PasswordKeySpec extends KeyDataSpec {
  spinCount: number;
  encryptedVerifierHashInput: Uint8Array;
  encryptedVerifierHashValue: Uint8Array;
  encryptedKeyValue: Uint8Array;
}

export interface AgileEncryptionInfo {
  keyData: KeyDataSpec;
  password: PasswordKeySpec;
}

function b64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function bytesToB64(bytes: Uint8Array): string {
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

/** "SHA512" / "SHA-512" → the WebCrypto name. */
function webcryptoHash(name: string): string {
  const n = name.toUpperCase().replace('-', '');
  if (n === 'SHA512') return 'SHA-512';
  if (n === 'SHA384') return 'SHA-384';
  if (n === 'SHA256') return 'SHA-256';
  if (n === 'SHA1') return 'SHA-1';
  throw new DocxCryptoError(`unsupported hash algorithm: ${name}`);
}

const attr = (xml: string, tag: string, name: string): string | undefined => {
  // The element may carry a namespace prefix (p:encryptedKey); match on the
  // local name and read the attribute out of that one element's tag text.
  const el = new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>`).exec(xml)?.[0];
  if (!el) return undefined;
  return new RegExp(`\\b${name}="([^"]*)"`).exec(el)?.[1];
};

/** Parse the EncryptionInfo stream. Throws for anything that isn't the Agile
 *  scheme (Word 2007's "Standard" RC4/AES is a different, older layout). */
export function parseEncryptionInfo(stream: Uint8Array): AgileEncryptionInfo {
  if (stream.length < 8) throw new DocxCryptoError('EncryptionInfo too short');
  const view = new DataView(
    stream.buffer,
    stream.byteOffset,
    stream.byteLength,
  );
  const major = view.getUint16(0, true);
  const minor = view.getUint16(2, true);
  if (major !== 4 || minor !== 4) {
    throw new DocxCryptoError(
      `unsupported encryption version ${major}.${minor} (only Agile 4.4 is supported)`,
    );
  }
  const xml = new TextDecoder().decode(stream.subarray(8));
  const num = (tag: string, name: string, fallback?: number): number => {
    const raw = attr(xml, tag, name);
    if (raw === undefined) {
      if (fallback !== undefined) return fallback;
      throw new DocxCryptoError(`EncryptionInfo: missing ${tag}/@${name}`);
    }
    return Number(raw);
  };
  const bin = (tag: string, name: string): Uint8Array => {
    const raw = attr(xml, tag, name);
    if (raw === undefined)
      throw new DocxCryptoError(`EncryptionInfo: missing ${tag}/@${name}`);
    return b64ToBytes(raw);
  };
  const str = (tag: string, name: string): string => {
    const raw = attr(xml, tag, name);
    if (raw === undefined)
      throw new DocxCryptoError(`EncryptionInfo: missing ${tag}/@${name}`);
    return raw;
  };

  const cipher = str('keyData', 'cipherAlgorithm');
  if (!/^AES$/i.test(cipher))
    throw new DocxCryptoError(`unsupported cipher: ${cipher}`);
  const chaining = str('keyData', 'cipherChaining');
  if (!/ChainingModeCBC/i.test(chaining))
    throw new DocxCryptoError(`unsupported chaining mode: ${chaining}`);

  return {
    keyData: {
      saltValue: bin('keyData', 'saltValue'),
      blockSize: num('keyData', 'blockSize', 16),
      keyBits: num('keyData', 'keyBits'),
      hashSize: num('keyData', 'hashSize'),
      hashAlgorithm: str('keyData', 'hashAlgorithm'),
    },
    password: {
      saltValue: bin('encryptedKey', 'saltValue'),
      blockSize: num('encryptedKey', 'blockSize', 16),
      keyBits: num('encryptedKey', 'keyBits'),
      hashSize: num('encryptedKey', 'hashSize'),
      hashAlgorithm: str('encryptedKey', 'hashAlgorithm'),
      spinCount: num('encryptedKey', 'spinCount'),
      encryptedVerifierHashInput: bin('encryptedKey', 'encryptedVerifierHashInput'),
      encryptedVerifierHashValue: bin('encryptedKey', 'encryptedVerifierHashValue'),
      encryptedKeyValue: bin('encryptedKey', 'encryptedKeyValue'),
    },
  };
}

const subtle = () => {
  const c = globalThis.crypto;
  if (!c?.subtle)
    throw new DocxCryptoError('WebCrypto is unavailable in this environment');
  return c.subtle;
};

const digest = async (alg: string, data: Uint8Array): Promise<Uint8Array> =>
  new Uint8Array(await subtle().digest(webcryptoHash(alg), data));

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let w = 0;
  for (const p of parts) {
    out.set(p, w);
    w += p.length;
  }
  return out;
}

const le32 = (n: number): Uint8Array => {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n, true);
  return b;
};

/** Truncate to `size`, or pad with 0x36 — MS-OFFCRYPTO §2.3.4.11. */
function fit(bytes: Uint8Array, size: number): Uint8Array {
  if (bytes.length === size) return bytes;
  if (bytes.length > size) return bytes.subarray(0, size);
  const out = new Uint8Array(size).fill(0x36);
  out.set(bytes);
  return out;
}

/**
 * The password KDF (§2.3.4.11): H₀ = H(salt ‖ UTF16LE(password)), then
 * `spinCount` rounds of H(iterator ‖ Hₙ), then one final H(Hₙ ‖ blockKey).
 * The spin is the whole point — it's what makes guessing expensive — so this
 * is the slow call (~1–3 s at Word's 100 000 rounds) and belongs in a worker.
 */
export async function deriveKey(
  password: string,
  spec: PasswordKeySpec,
  blockKey: Uint8Array,
): Promise<Uint8Array> {
  const pw = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) {
    new DataView(pw.buffer).setUint16(i * 2, password.charCodeAt(i), true);
  }
  let h = await digest(spec.hashAlgorithm, concat(spec.saltValue, pw));
  for (let i = 0; i < spec.spinCount; i++) {
    h = await digest(spec.hashAlgorithm, concat(le32(i), h));
  }
  const final = await digest(spec.hashAlgorithm, concat(h, blockKey));
  return fit(final, spec.keyBits / 8);
}

/**
 * AES-CBC without PKCS#7 — what Office writes, and what WebCrypto refuses to
 * produce (it always pads). The fix is to append one ciphertext block that
 * decrypts to a full block of valid padding, so the built-in unpadding strips
 * exactly that block and hands back the raw plaintext. That extra block is
 * ECB(0x10…0x10 ⊕ lastCipherBlock), and CBC-encrypt with IV = lastCipherBlock
 * yields precisely that as its first block.
 */
async function aesCbcDecryptNoPad(
  key: Uint8Array,
  iv: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  if (data.length === 0) return new Uint8Array(0);
  if (data.length % 16 !== 0)
    throw new DocxCryptoError('ciphertext is not a whole number of blocks');
  const k = await subtle().importKey(
    'raw',
    key,
    { name: 'AES-CBC' },
    false,
    ['decrypt', 'encrypt'],
  );
  const lastBlock = data.subarray(data.length - 16);
  const padPlain = new Uint8Array(16).fill(16);
  const padEnc = new Uint8Array(
    await subtle().encrypt(
      { name: 'AES-CBC', iv: lastBlock },
      k,
      padPlain,
    ),
  ).subarray(0, 16);
  const full = concat(data, padEnc);
  const plain = await subtle().decrypt(
    { name: 'AES-CBC', iv: iv },
    k,
    full,
  );
  return new Uint8Array(plain);
}

/** Encrypt without PKCS#7 (test fixtures / a future re-encrypt path): CBC
 *  encrypt then drop the padding block WebCrypto appends. */
export async function aesCbcEncryptNoPad(
  key: Uint8Array,
  iv: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  if (data.length % 16 !== 0)
    throw new DocxCryptoError('plaintext is not a whole number of blocks');
  const k = await subtle().importKey(
    'raw',
    key,
    { name: 'AES-CBC' },
    false,
    ['encrypt'],
  );
  const out = new Uint8Array(
    await subtle().encrypt(
      { name: 'AES-CBC', iv: iv },
      k,
      data,
    ),
  );
  return out.subarray(0, data.length);
}

/**
 * Check a password against the stored verifier — a purely local comparison
 * (§2.3.4.13), so a wrong password costs one KDF and nothing else. Returns the
 * derived intermediate key when it matches, so callers don't spin twice.
 */
export async function verifyPassword(
  password: string,
  info: AgileEncryptionInfo,
): Promise<Uint8Array> {
  const spec = info.password;
  const kIn = await deriveKey(password, spec, BLOCK_KEY_VERIFIER_INPUT);
  const verifierInput = await aesCbcDecryptNoPad(
    kIn,
    spec.saltValue,
    spec.encryptedVerifierHashInput,
  );
  const kVal = await deriveKey(password, spec, BLOCK_KEY_VERIFIER_VALUE);
  const storedHash = await aesCbcDecryptNoPad(
    kVal,
    spec.saltValue,
    spec.encryptedVerifierHashValue,
  );
  const actual = await digest(spec.hashAlgorithm, verifierInput);
  for (let i = 0; i < spec.hashSize; i++) {
    if (actual[i] !== storedHash[i]) throw new WrongPasswordError();
  }
  // Correct password → unwrap the key the package is actually encrypted with.
  const kSecret = await deriveKey(password, spec, BLOCK_KEY_SECRET);
  const secret = await aesCbcDecryptNoPad(
    kSecret,
    spec.saltValue,
    spec.encryptedKeyValue,
  );
  return secret.subarray(0, spec.keyBits / 8);
}

/**
 * Decrypt the EncryptedPackage stream with the unwrapped key. The payload is
 * split into 4096-byte segments, each with its own IV derived from the
 * keyData salt and the segment index (§2.3.4.15), so segments are independent.
 */
export async function decryptPackage(
  encryptedPackage: Uint8Array,
  secretKey: Uint8Array,
  keyData: KeyDataSpec,
): Promise<Uint8Array> {
  if (encryptedPackage.length < 8)
    throw new DocxCryptoError('EncryptedPackage too short');
  const view = new DataView(
    encryptedPackage.buffer,
    encryptedPackage.byteOffset,
    encryptedPackage.byteLength,
  );
  // uint64 LE plaintext length; the high dword is always 0 for real documents.
  const size = view.getUint32(0, true) + view.getUint32(4, true) * 2 ** 32;
  const body = encryptedPackage.subarray(8);
  const out = new Uint8Array(Math.min(size, body.length));
  for (let i = 0, w = 0; w < out.length; i++) {
    const start = i * SEGMENT_SIZE;
    const chunk = body.subarray(start, start + SEGMENT_SIZE);
    if (chunk.length === 0) break;
    const iv = fit(
      await digest(keyData.hashAlgorithm, concat(keyData.saltValue, le32(i))),
      keyData.blockSize,
    );
    const plain = await aesCbcDecryptNoPad(secretKey, iv, chunk);
    const n = Math.min(plain.length, out.length - w);
    out.set(plain.subarray(0, n), w);
    w += n;
  }
  return out;
}
