import { CfbReader } from './cfb';
import {
  decryptOfficeFile,
  decryptPackage,
  parseEncryptionInfo,
  reencryptOfficeFile,
  unlockOfficeFile,
  verifyPassword,
  WrongPasswordError,
} from './crypto-docx';
import { buildCfb, buildEncryptedDocx } from './crypto-docx.spec-helper';


// ── Tests ───────────────────────────────────────────────────────────

describe('CfbReader', () => {
  it('reads streams from both the mini-FAT and the FAT', async () => {
    const small = new TextEncoder().encode('a small stream, under the cutoff');
    const big = crypto.getRandomValues(new Uint8Array(9000)); // > 4096 → FAT
    const cfb = buildCfb([
      { name: 'EncryptionInfo', data: small },
      { name: 'EncryptedPackage', data: big },
    ]);
    const r = new CfbReader(cfb);
    expect(r.streamNames().sort()).toEqual([
      'EncryptedPackage',
      'EncryptionInfo',
    ]);
    expect(r.readStream('EncryptionInfo')).toEqual(small);
    expect(r.readStream('EncryptedPackage')).toEqual(big);
    expect(r.readStream('Nope')).toBeNull();
  });

  it('rejects bytes that are not a compound file', () => {
    expect(() => new CfbReader(new TextEncoder().encode('PK not ole'))).toThrow(
      /compound file/,
    );
  });
});

describe('Agile decryption', () => {
  const plain = () => {
    // Something longer than one 4096-byte segment, so segment IV derivation
    // is exercised rather than assumed.
    const b = crypto.getRandomValues(new Uint8Array(10_000));
    b.set([0x50, 0x4b, 0x03, 0x04]); // looks like the zip it stands in for
    return b;
  };

  it('round-trips a password-protected package', async () => {
    const source = plain();
    const file = await buildEncryptedDocx(source, 'hùng lâm 123');
    const r = new CfbReader(file);
    const info = parseEncryptionInfo(r.readStream('EncryptionInfo')!);
    expect(info.password.spinCount).toBe(50);
    expect(info.keyData.keyBits).toBe(256);

    const key = await verifyPassword('hùng lâm 123', info);
    const out = await decryptPackage(
      r.readStream('EncryptedPackage')!,
      key,
      info.keyData,
    );
    expect(out.length).toBe(source.length);
    expect(out).toEqual(source);
  });

  it('rejects a wrong password locally, before any bulk work', async () => {
    const file = await buildEncryptedDocx(plain(), 'correct horse');
    const r = new CfbReader(file);
    const info = parseEncryptionInfo(r.readStream('EncryptionInfo')!);
    await expect(verifyPassword('battery staple', info)).rejects.toBeInstanceOf(
      WrongPasswordError,
    );
  });

  it('decryptOfficeFile unwraps straight to the inner bytes', async () => {
    const source = plain();
    const file = await buildEncryptedDocx(source, 'mật khẩu');
    expect(await decryptOfficeFile(file, 'mật khẩu')).toEqual(source);
    await expect(decryptOfficeFile(file, 'sai')).rejects.toBeInstanceOf(
      WrongPasswordError,
    );
  });

  it('re-encrypts edited content so the SAME password still opens it', async () => {
    const original = plain();
    const file = await buildEncryptedDocx(original, 'giữ nguyên mật khẩu');
    const { plain: got, material } = await unlockOfficeFile(
      file,
      'giữ nguyên mật khẩu',
    );
    expect(got).toEqual(original);

    // Edit, save, reopen — with the password the user already knows.
    const edited = plain();
    const saved = await reencryptOfficeFile(edited, material);
    const reopened = await decryptOfficeFile(saved, 'giữ nguyên mật khẩu');
    expect(reopened).toEqual(edited);
    // …and only that password.
    await expect(decryptOfficeFile(saved, 'khác')).rejects.toBeInstanceOf(
      WrongPasswordError,
    );
  });

  it('gives every save a fresh package key and salt', async () => {
    // Same content saved twice must not produce the same ciphertext: the
    // keyData salt derives the segment IVs, so reusing it would leak which
    // parts of a document changed between two versions.
    const file = await buildEncryptedDocx(plain(), 'pw');
    const { material } = await unlockOfficeFile(file, 'pw');
    const content = plain();
    const a = await reencryptOfficeFile(content, material);
    const b = await reencryptOfficeFile(content, material);
    const pkg = (f: Uint8Array) =>
      new CfbReader(f).readStream('EncryptedPackage')!;
    expect(pkg(a)).not.toEqual(pkg(b));
    // Both still open, and to the same content.
    expect(await decryptOfficeFile(a, 'pw')).toEqual(content);
    expect(await decryptOfficeFile(b, 'pw')).toEqual(content);
  });

  it('re-encrypting keeps the original key-derivation parameters', async () => {
    const file = await buildEncryptedDocx(plain(), 'pw', 77);
    const { material } = await unlockOfficeFile(file, 'pw');
    const saved = await reencryptOfficeFile(plain(), material);
    const info = parseEncryptionInfo(
      new CfbReader(saved).readStream('EncryptionInfo')!,
    );
    expect(info.password.spinCount).toBe(77); // not silently weakened
    expect(info.keyData.keyBits).toBe(256);
  });

  it('refuses non-Agile encryption instead of guessing', () => {
    const legacy = new Uint8Array(16);
    new DataView(legacy.buffer).setUint16(0, 3, true); // Standard, not Agile
    new DataView(legacy.buffer).setUint16(2, 2, true);
    expect(() => parseEncryptionInfo(legacy)).toThrow(/only Agile/);
  });
});
