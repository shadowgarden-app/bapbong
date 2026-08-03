/**
 * Minimal reader for the Compound File Binary format (MS-CFB) — the OLE
 * container Office wraps an ENCRYPTED .docx in. We only ever need two streams
 * out of it (`EncryptionInfo` and `EncryptedPackage`), so this reads the
 * header, walks the FAT/mini-FAT chains and returns stream bytes by name; it
 * does not write, and ignores storages beyond finding entries inside them.
 *
 * Sizes and offsets follow MS-CFB §2.1–2.6.
 */

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Sector chain terminators (MS-CFB §2.2). */
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;

/** A directory entry we care about: a stream and where its bytes live. */
interface DirEntry {
  name: string;
  /** 0 unknown, 1 storage, 2 stream, 5 root. */
  type: number;
  startSector: number;
  size: number;
}

export class CfbError extends Error {}

/** Guard: a corrupt/hostile file must not send a chain walk infinite, and a
 *  declared stream size must not make us allocate gigabytes. Both are far
 *  above anything a real encrypted .docx needs. */
const MAX_CHAIN = 1 << 22; // sectors
const MAX_STREAM = 512 * 1024 * 1024;

export class CfbReader {
  private readonly view: DataView;
  private readonly bytes: Uint8Array;
  private readonly sectorSize: number;
  private readonly miniSectorSize: number;
  private readonly miniCutoff: number;
  private readonly fat: Uint32Array;
  private readonly miniFat: Uint32Array;
  private readonly entries: DirEntry[] = [];
  /** The mini stream's own bytes (the root entry's stream), assembled once. */
  private miniStream: Uint8Array | null = null;

  constructor(input: ArrayBuffer | Uint8Array) {
    this.bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    if (this.bytes.length < 512) throw new CfbError('not a compound file');
    for (let i = 0; i < SIGNATURE.length; i++) {
      if (this.bytes[i] !== SIGNATURE[i])
        throw new CfbError('not a compound file');
    }
    this.view = new DataView(
      this.bytes.buffer,
      this.bytes.byteOffset,
      this.bytes.byteLength,
    );
    this.sectorSize = 1 << this.view.getUint16(30, true);
    this.miniSectorSize = 1 << this.view.getUint16(32, true);
    this.miniCutoff = this.view.getUint32(56, true);
    if (this.sectorSize < 128 || this.sectorSize > 1 << 20)
      throw new CfbError('unsupported sector size');

    this.fat = this.readFat();
    this.miniFat = this.readChainAsUint32(
      this.view.getUint32(60, true), // first mini-FAT sector
    );
    this.readDirectory(this.view.getUint32(48, true));
  }

  /** Byte offset of sector `n` (sector 0 starts right after the 512B header). */
  private sectorOffset(n: number): number {
    return 512 + n * this.sectorSize;
  }

  /** The FAT itself, assembled from the DIFAT (109 entries in the header,
   *  then any DIFAT sectors). */
  private readFat(): Uint32Array {
    const fatSectors: number[] = [];
    for (let i = 0; i < 109; i++) {
      const s = this.view.getUint32(76 + i * 4, true);
      if (s === FREESECT || s === ENDOFCHAIN) break;
      fatSectors.push(s);
    }
    // Extra DIFAT sectors: each holds (sectorSize/4 - 1) FAT sector numbers
    // plus a pointer to the next DIFAT sector in its last slot.
    let difat = this.view.getUint32(68, true);
    const perSector = this.sectorSize / 4 - 1;
    let guard = 0;
    while (difat !== ENDOFCHAIN && difat !== FREESECT && guard++ < MAX_CHAIN) {
      const off = this.sectorOffset(difat);
      if (off + this.sectorSize > this.bytes.length) break;
      for (let i = 0; i < perSector; i++) {
        const s = this.view.getUint32(off + i * 4, true);
        if (s === FREESECT || s === ENDOFCHAIN) continue;
        fatSectors.push(s);
      }
      difat = this.view.getUint32(off + perSector * 4, true);
    }

    const out = new Uint32Array((fatSectors.length * this.sectorSize) / 4);
    let w = 0;
    for (const s of fatSectors) {
      const off = this.sectorOffset(s);
      if (off + this.sectorSize > this.bytes.length) break;
      for (let i = 0; i < this.sectorSize / 4; i++) {
        out[w++] = this.view.getUint32(off + i * 4, true);
      }
    }
    return out;
  }

  /** Sector numbers of a chain starting at `start`, following the FAT. */
  private chain(start: number): number[] {
    const out: number[] = [];
    let s = start;
    let guard = 0;
    while (s !== ENDOFCHAIN && s !== FREESECT && guard++ < MAX_CHAIN) {
      out.push(s);
      if (s >= this.fat.length) break;
      s = this.fat[s];
    }
    return out;
  }

  /** A FAT chain's bytes read as little-endian uint32s (used for the mini-FAT). */
  private readChainAsUint32(start: number): Uint32Array {
    const sectors = this.chain(start);
    const out = new Uint32Array((sectors.length * this.sectorSize) / 4);
    let w = 0;
    for (const s of sectors) {
      const off = this.sectorOffset(s);
      if (off + this.sectorSize > this.bytes.length) break;
      for (let i = 0; i < this.sectorSize / 4; i++) {
        out[w++] = this.view.getUint32(off + i * 4, true);
      }
    }
    return out;
  }

  /** Every directory entry, walked as a flat array (we look entries up by
   *  name, so the red-black tree structure doesn't matter here). */
  private readDirectory(firstDirSector: number): void {
    for (const sector of this.chain(firstDirSector)) {
      const base = this.sectorOffset(sector);
      if (base + this.sectorSize > this.bytes.length) break;
      for (let off = base; off + 128 <= base + this.sectorSize; off += 128) {
        const nameLen = this.view.getUint16(off + 64, true);
        const type = this.view.getUint8(off + 66);
        if (type === 0) continue; // unallocated
        // nameLen counts BYTES including the UTF-16 null terminator.
        let name = '';
        for (let i = 0; i + 1 < Math.min(nameLen, 64); i += 2) {
          const c = this.view.getUint16(off + i, true);
          if (c === 0) break;
          name += String.fromCharCode(c);
        }
        this.entries.push({
          name,
          type,
          startSector: this.view.getUint32(off + 116, true),
          // Streams here are far under 4 GB; the high dword is ignored on
          // purpose (a bogus one would only inflate an allocation).
          size: this.view.getUint32(off + 120, true),
        });
      }
    }
  }

  /** Assemble bytes from regular (FAT) sectors. */
  private readFromFat(start: number, size: number): Uint8Array {
    const out = new Uint8Array(size);
    let w = 0;
    for (const s of this.chain(start)) {
      if (w >= size) break;
      const off = this.sectorOffset(s);
      const n = Math.min(this.sectorSize, size - w, this.bytes.length - off);
      if (n <= 0) break;
      out.set(this.bytes.subarray(off, off + n), w);
      w += n;
    }
    return out;
  }

  /** Assemble bytes from the mini stream (mini-FAT sectors). */
  private readFromMini(start: number, size: number): Uint8Array {
    if (!this.miniStream) {
      const root = this.entries.find((e) => e.type === 5);
      this.miniStream = root
        ? this.readFromFat(root.startSector, root.size)
        : new Uint8Array(0);
    }
    const out = new Uint8Array(size);
    let w = 0;
    let s = start;
    let guard = 0;
    while (s !== ENDOFCHAIN && s !== FREESECT && w < size && guard++ < MAX_CHAIN) {
      const off = s * this.miniSectorSize;
      const n = Math.min(
        this.miniSectorSize,
        size - w,
        Math.max(0, this.miniStream.length - off),
      );
      if (n <= 0) break;
      out.set(this.miniStream.subarray(off, off + n), w);
      w += n;
      if (s >= this.miniFat.length) break;
      s = this.miniFat[s];
    }
    return out;
  }

  /** Names of the streams present (diagnostics / tests). */
  streamNames(): string[] {
    return this.entries.filter((e) => e.type === 2).map((e) => e.name);
  }

  /** A stream's bytes by name, or null when it isn't in the container. */
  readStream(name: string): Uint8Array | null {
    const e = this.entries.find((x) => x.name === name && x.type === 2);
    if (!e) return null;
    if (e.size > MAX_STREAM) throw new CfbError(`${name}: stream too large`);
    return e.size < this.miniCutoff
      ? this.readFromMini(e.startSector, e.size)
      : this.readFromFat(e.startSector, e.size);
  }
}
