/**
 * ZIP reader over a File: reads the central directory, then streams single entries through
 * the native DecompressionStream. The archive is never loaded into memory as a whole.
 */

export interface ZipEntry {
  name: string;
  method: number;
  compSize: number;
  offset: number;
}

async function view(file: File, start: number, end: number): Promise<DataView> {
  return new DataView(await file.slice(start, end).arrayBuffer());
}

export class ZipArchive {
  private readonly lower = new Map<string, ZipEntry>();

  private constructor(private readonly file: File, readonly entries: Map<string, ZipEntry>) {
    for (const [name, e] of entries) this.lower.set(name.toLowerCase(), e);
  }

  static async open(file: File): Promise<ZipArchive> {
    const tailLen = Math.min(file.size, 65557);
    const tail = await view(file, file.size - tailLen, file.size);
    let eocd = -1;
    for (let i = tailLen - 22; i >= 0; i--) {
      if (tail.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('The file is damaged or incomplete (ZIP end record not found).');

    let count = tail.getUint16(eocd + 10, true);
    let cdSize = tail.getUint32(eocd + 12, true);
    let cdOffset = tail.getUint32(eocd + 16, true);

    if (cdOffset === 0xffffffff || count === 0xffff || cdSize === 0xffffffff) {
      const loc = eocd - 20;
      if (loc >= 0 && tail.getUint32(loc, true) === 0x07064b50) {
        const z64 = Number(tail.getBigUint64(loc + 8, true));
        const z = await view(file, z64, z64 + 56);
        count = Number(z.getBigUint64(32, true));
        cdSize = Number(z.getBigUint64(40, true));
        cdOffset = Number(z.getBigUint64(48, true));
      }
    }

    const cd = await view(file, cdOffset, cdOffset + cdSize);
    const utf8 = new TextDecoder();
    const latin = new TextDecoder('windows-1252');
    const entries = new Map<string, ZipEntry>();
    let p = 0;
    for (let i = 0; i < count && p + 46 <= cd.byteLength; i++) {
      if (cd.getUint32(p, true) !== 0x02014b50) break;
      const flags = cd.getUint16(p + 8, true);
      const method = cd.getUint16(p + 10, true);
      let compSize = cd.getUint32(p + 20, true);
      let size = cd.getUint32(p + 24, true);
      const nameLen = cd.getUint16(p + 28, true);
      const extraLen = cd.getUint16(p + 30, true);
      const commentLen = cd.getUint16(p + 32, true);
      let offset = cd.getUint32(p + 42, true);
      const nameBytes = new Uint8Array(cd.buffer, p + 46, nameLen);
      const name = ((flags & 0x800) ? utf8 : latin).decode(nameBytes).replace(/\\/g, '/').replace(/^\/+/, '');

      // ZIP64 extended info: only the fields saturated in the main record are present.
      let x = p + 46 + nameLen;
      const xEnd = x + extraLen;
      while (x + 4 <= xEnd) {
        const id = cd.getUint16(x, true);
        const len = cd.getUint16(x + 2, true);
        if (id === 0x0001) {
          let q = x + 4;
          if (size === 0xffffffff) { size = Number(cd.getBigUint64(q, true)); q += 8; }
          if (compSize === 0xffffffff) { compSize = Number(cd.getBigUint64(q, true)); q += 8; }
          if (offset === 0xffffffff) offset = Number(cd.getBigUint64(q, true));
        }
        x += 4 + len;
      }
      if (flags & 0x1) throw new Error('This file is password-protected. Remove the password in Excel and try again.');
      entries.set(name, { name, method, compSize, offset });
      p += 46 + nameLen + extraLen + commentLen;
    }
    return new ZipArchive(file, entries);
  }

  /** Finds a part by path: exact, then case-insensitive, then URL-decoded. */
  find(path: string): ZipEntry | undefined {
    path = path.replace(/^\/+/, '');
    let decoded = path;
    try { decoded = decodeURIComponent(path); } catch { /* keep as is */ }
    return this.entries.get(path) ?? this.lower.get(path.toLowerCase())
      ?? this.entries.get(decoded) ?? this.lower.get(decoded.toLowerCase());
  }

  async bytes(entry: ZipEntry, onBytes?: (n: number) => void): Promise<ReadableStream<Uint8Array>> {
    const local = await view(this.file, entry.offset, entry.offset + 30);
    if (local.getUint32(0, true) !== 0x04034b50) throw new Error(`The file is damaged (bad ZIP entry ${entry.name}).`);
    const start = entry.offset + 30 + local.getUint16(26, true) + local.getUint16(28, true);

    let stream = this.file.slice(start, start + entry.compSize).stream() as ReadableStream<Uint8Array>;
    if (onBytes) {
      stream = stream.pipeThrough(new TransformStream<Uint8Array, Uint8Array>({
        transform(chunk, ctl) { onBytes(chunk.byteLength); ctl.enqueue(chunk); },
      }));
    }
    if (entry.method === 8) {
      return stream.pipeThrough(new DecompressionStream('deflate-raw') as unknown as TransformStream<Uint8Array, Uint8Array>);
    }
    if (entry.method === 0) return stream;
    throw new Error(`Unsupported ZIP compression method ${entry.method} in ${entry.name}.`);
  }

  async text(entry: ZipEntry, onBytes?: (n: number) => void): Promise<ReadableStream<string>> {
    return (await this.bytes(entry, onBytes))
      .pipeThrough(new TextDecoderStream() as unknown as TransformStream<Uint8Array, string>);
  }

  async readText(path: string): Promise<string | null> {
    const entry = this.find(path);
    if (!entry) return null;
    return new Response(await this.bytes(entry)).text();
  }
}
