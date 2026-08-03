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

// ── Writing ──────────────────────────────────────────────────────────

const FATSECT = 0xfffffffd;
const SECTOR = 512;
const MINI_SECTOR = 64;
/** Streams smaller than this live in the mini stream (MS-CFB §2.2). */
const MINI_CUTOFF = 4096;

/**
 * Write a compound file holding `streams` — the container an encrypted
 * document needs (EncryptionInfo + EncryptedPackage). Deliberately minimal:
 * one FAT sector, one mini-FAT sector and a flat directory, which covers
 * documents up to ~64 MB. Larger inputs throw rather than emit a file that
 * would silently truncate.
 */
export function writeCfb(
  streams: { name: string; data: Uint8Array }[],
): Uint8Array {
  const small = streams.filter((s) => s.data.length < MINI_CUTOFF);
  const large = streams.filter((s) => s.data.length >= MINI_CUTOFF);
  if (streams.length > 3) throw new CfbError('too many streams for this writer');

  // Mini stream: the small streams, each starting on a mini-sector boundary.
  const miniChunks: { name: string; start: number; size: number }[] = [];
  let miniLen = 0;
  for (const s of small) {
    miniChunks.push({
      name: s.name,
      start: miniLen / MINI_SECTOR,
      size: s.data.length,
    });
    miniLen += Math.ceil(s.data.length / MINI_SECTOR) * MINI_SECTOR;
  }
  const miniStream = new Uint8Array(miniLen);
  {
    let w = 0;
    for (const s of small) {
      miniStream.set(s.data, w);
      w += Math.ceil(s.data.length / MINI_SECTOR) * MINI_SECTOR;
    }
  }

  const miniStreamSectors = Math.ceil(miniStream.length / SECTOR);
  const largeSectors = large.map((s) => Math.ceil(s.data.length / SECTOR));
  // Sector 0 = FAT, 1 = directory, 2 = mini-FAT, 3… = mini stream, then the
  // large streams in order.
  const MINI_START = 3;
  const totalSectors =
    3 + miniStreamSectors + largeSectors.reduce((a, b) => a + b, 0);
  const fatCapacity = SECTOR / 4;
  if (totalSectors > fatCapacity)
    throw new CfbError('document too large for this writer');
  if (miniStream.length / MINI_SECTOR > fatCapacity)
    throw new CfbError('too many mini sectors for this writer');

  const fat = new Uint32Array(fatCapacity).fill(FREESECT);
  fat[0] = FATSECT;
  fat[1] = ENDOFCHAIN;
  fat[2] = ENDOFCHAIN;
  const chainRange = (start: number, count: number) => {
    for (let i = 0; i < count; i++)
      fat[start + i] = i === count - 1 ? ENDOFCHAIN : start + i + 1;
  };
  chainRange(MINI_START, miniStreamSectors);
  const largeStarts: number[] = [];
  let next = MINI_START + miniStreamSectors;
  for (const n of largeSectors) {
    largeStarts.push(next);
    chainRange(next, n);
    next += n;
  }

  const miniFat = new Uint32Array(fatCapacity).fill(FREESECT);
  for (const c of miniChunks) {
    const n = Math.ceil(c.size / MINI_SECTOR);
    for (let i = 0; i < n; i++)
      miniFat[c.start + i] = i === n - 1 ? ENDOFCHAIN : c.start + i + 1;
  }

  const out = new Uint8Array((1 + totalSectors) * SECTOR);
  const view = new DataView(out.buffer);
  const sectorAt = (n: number) => SECTOR + n * SECTOR;

  out.set(SIGNATURE, 0);
  view.setUint16(24, 0x003e, true); // minor version
  view.setUint16(26, 0x0003, true); // major version (512-byte sectors)
  view.setUint16(28, 0xfffe, true); // byte-order marker
  view.setUint16(30, 9, true); // sector shift → 512
  view.setUint16(32, 6, true); // mini-sector shift → 64
  view.setUint32(44, 1, true); // FAT sector count
  view.setUint32(48, 1, true); // first directory sector
  view.setUint32(56, MINI_CUTOFF, true);
  view.setUint32(60, 2, true); // first mini-FAT sector
  view.setUint32(64, 1, true); // mini-FAT sector count
  view.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector
  view.setUint32(72, 0, true); // DIFAT sector count
  for (let i = 0; i < 109; i++) view.setUint32(76 + i * 4, FREESECT, true);
  view.setUint32(76, 0, true); // DIFAT[0] → the FAT sector

  for (let i = 0; i < fat.length; i++)
    view.setUint32(sectorAt(0) + i * 4, fat[i], true);
  for (let i = 0; i < miniFat.length; i++)
    view.setUint32(sectorAt(2) + i * 4, miniFat[i], true);
  out.set(miniStream, sectorAt(MINI_START));
  large.forEach((s, i) => out.set(s.data, sectorAt(largeStarts[i])));

  const dirBase = sectorAt(1);
  const entry = (
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
    view.setUint8(off + 67, 1); // colour: black
    view.setUint32(off + 68, FREESECT, true); // left sibling
    view.setUint32(off + 72, FREESECT, true); // right sibling
    view.setUint32(off + 76, FREESECT, true); // child
    view.setUint32(off + 116, start, true);
    view.setUint32(off + 120, size, true);
  };
  entry(0, 'Root Entry', 5, MINI_START, miniStream.length);
  const placed: { name: string; slot: number }[] = [];
  let slot = 1;
  for (const c of miniChunks) {
    entry(slot, c.name, 2, c.start, c.size);
    placed.push({ name: c.name, slot: slot++ });
  }
  large.forEach((st, i) => {
    entry(slot, st.name, 2, largeStarts[i], st.data.length);
    placed.push({ name: st.name, slot: slot++ });
  });

  // Link the entries into the directory TREE. Readers walk it from the root's
  // child pointer — leaving the pointers empty (as a linear scan of the
  // sectors does not need) yields a container whose streams are invisible to
  // every conforming reader, including Word's.
  //
  // MS-CFB §2.6.4 orders siblings by name LENGTH first, then by the uppercased
  // name. A right-leaning chain is a valid (if unbalanced) binary tree for the
  // two or three entries written here.
  placed.sort((a, b) =>
    a.name.length !== b.name.length
      ? a.name.length - b.name.length
      : a.name.toUpperCase() < b.name.toUpperCase()
        ? -1
        : 1,
  );
  const RIGHT_SIBLING = 72;
  const CHILD = 76;
  const field = (s2: number, off: number, v: number) =>
    view.setUint32(dirBase + s2 * 128 + off, v, true);
  if (placed.length > 0) {
    field(0, CHILD, placed[0].slot); // root → first entry
    for (let i = 0; i + 1 < placed.length; i++) {
      field(placed[i].slot, RIGHT_SIBLING, placed[i + 1].slot);
    }
  }
  return out;
}
