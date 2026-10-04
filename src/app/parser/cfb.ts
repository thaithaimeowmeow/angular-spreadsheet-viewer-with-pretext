/** OLE2 / Compound File Binary reader (the container of legacy .xls files). */

const END_OF_CHAIN = 0xfffffffe;
const FREE = 0xffffffff;

export interface CfbEntry { name: string; type: number; start: number; size: number }

export class CompoundFile {
  readonly entries: CfbEntry[] = [];
  private readonly sectorSize: number;
  private readonly miniSize: number;
  private readonly miniCutoff: number;
  private readonly fat: Uint32Array;
  private readonly miniFat: Uint32Array;
  private readonly miniStream: Uint8Array;

  constructor(private readonly buf: Uint8Array) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    if (dv.getUint32(0, true) !== 0xe011cfd0 || dv.getUint32(4, true) !== 0xe11ab1a1) {
      throw new Error('Not an OLE2 compound file.');
    }
    this.sectorSize = 1 << dv.getUint16(0x1e, true);
    this.miniSize = 1 << dv.getUint16(0x20, true);
    const fatCount = dv.getUint32(0x2c, true);
    const dirStart = dv.getUint32(0x30, true);
    this.miniCutoff = dv.getUint32(0x38, true) || 4096;
    const miniFatStart = dv.getUint32(0x3c, true);
    let difatSector = dv.getUint32(0x44, true);

    // DIFAT: 109 entries in the header, then a chain of DIFAT sectors.
    const fatSectors: number[] = [];
    for (let i = 0; i < 109 && fatSectors.length < fatCount; i++) fatSectors.push(dv.getUint32(0x4c + i * 4, true));
    const perSector = this.sectorSize / 4;
    for (let guard = 0; difatSector < END_OF_CHAIN && fatSectors.length < fatCount && guard < 1e6; guard++) {
      const off = this.offset(difatSector);
      for (let i = 0; i < perSector - 1 && fatSectors.length < fatCount; i++) fatSectors.push(dv.getUint32(off + i * 4, true));
      difatSector = dv.getUint32(off + (perSector - 1) * 4, true);
    }

    this.fat = new Uint32Array(fatSectors.length * perSector);
    fatSectors.forEach((s, k) => {
      const off = this.offset(s);
      for (let i = 0; i < perSector && off + i * 4 + 4 <= buf.length; i++) this.fat[k * perSector + i] = dv.getUint32(off + i * 4, true);
    });

    const dir = this.readChain(dirStart);
    const ddv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
    for (let p = 0; p + 128 <= dir.length; p += 128) {
      const nameLen = ddv.getUint16(p + 0x40, true);
      let name = '';
      for (let i = 0; i + 2 < nameLen && i < 64; i += 2) name += String.fromCharCode(ddv.getUint16(p + i, true));
      this.entries.push({
        name,
        type: dir[p + 0x42],
        start: ddv.getUint32(p + 0x74, true),
        size: this.sectorSize === 512 ? ddv.getUint32(p + 0x78, true) : Number(ddv.getBigUint64(p + 0x78, true)),
      });
    }

    const root = this.entries[0];
    this.miniStream = root ? this.readChain(root.start, root.size) : new Uint8Array(0);
    const mf = miniFatStart < END_OF_CHAIN ? this.readChain(miniFatStart) : new Uint8Array(0);
    this.miniFat = new Uint32Array(mf.buffer.slice(mf.byteOffset, mf.byteOffset + (mf.byteLength & ~3)));
  }

  find(name: string): CfbEntry | undefined {
    const n = name.toLowerCase();
    return this.entries.find((e) => e.type === 2 && e.name.toLowerCase() === n);
  }

  read(entry: CfbEntry): Uint8Array {
    if (entry.size < this.miniCutoff) return this.readMini(entry.start, entry.size);
    return this.readChain(entry.start, entry.size);
  }

  private offset(sector: number) {
    return (sector + 1) * this.sectorSize;
  }

  private readChain(start: number, size?: number): Uint8Array {
    const sectors: number[] = [];
    for (let s = start; s < END_OF_CHAIN && s !== FREE && sectors.length <= this.fat.length; s = this.fat[s]) {
      sectors.push(s);
      if (s >= this.fat.length) break;
    }
    const total = size ?? sectors.length * this.sectorSize;
    const out = new Uint8Array(Math.min(total, sectors.length * this.sectorSize));
    sectors.forEach((s, i) => {
      const off = this.offset(s);
      const len = Math.min(this.sectorSize, out.length - i * this.sectorSize);
      if (len > 0) out.set(this.buf.subarray(off, off + len), i * this.sectorSize);
    });
    return out;
  }

  private readMini(start: number, size: number): Uint8Array {
    const out = new Uint8Array(size);
    let pos = 0;
    for (let s = start, guard = 0; s < END_OF_CHAIN && pos < size && guard <= this.miniFat.length; s = this.miniFat[s], guard++) {
      const off = s * this.miniSize;
      const len = Math.min(this.miniSize, size - pos);
      out.set(this.miniStream.subarray(off, off + len), pos);
      pos += len;
    }
    return out;
  }
}
