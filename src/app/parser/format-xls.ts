/** Legacy .xls: BIFF8 (Excel 97–2003) and BIFF5 (Excel 5.0/95) inside an OLE2 container. */
import { CompoundFile } from './cfb';
import type { Book, Job, RowEmitter } from './job';
import { Kind, builtinKind, formatNumber, kindOfFormat } from './numfmt';
import { normalizeNewlines } from './xml';

const R = {
  BOF: 0x0809, EOF: 0x000a, CONTINUE: 0x003c, CODEPAGE: 0x0042, DATEMODE: 0x0022, FILEPASS: 0x002f,
  BOUNDSHEET: 0x0085, SST: 0x00fc, FORMAT: 0x041e, FORMAT_OLD: 0x001e, XF: 0x00e0,
  LABELSST: 0x00fd, LABEL: 0x0204, RSTRING: 0x00d6, NUMBER: 0x0203, RK: 0x027e, MULRK: 0x00bd,
  BOOLERR: 0x0205, FORMULA: 0x0006, STRING: 0x0207,
} as const;

const ERRORS: Record<number, string> = {
  0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A',
};

interface Rec { type: number; data: Uint8Array; conts: Uint8Array[] }

function codepageName(cp: number): string {
  if (cp === 1200) return 'utf-16le';
  if (cp === 367 || cp === 20127) return 'us-ascii';
  if (cp === 10000 || cp === 32768) return 'macintosh';
  if (cp === 932) return 'shift_jis';
  if (cp === 936) return 'gbk';
  if (cp === 949) return 'euc-kr';
  if (cp === 950) return 'big5';
  if (cp === 874 || (cp >= 1250 && cp <= 1258)) return `windows-${cp}`;
  return 'windows-1252';
}

/** Splits a BIFF stream into records, attaching CONTINUE records to their parent. */
function records(stream: Uint8Array, start: number): { recs: Rec[]; end: number } {
  const recs: Rec[] = [];
  let depth = 0;
  let p = start;
  while (p + 4 <= stream.length) {
    const type = stream[p] | (stream[p + 1] << 8);
    const len = stream[p + 2] | (stream[p + 3] << 8);
    const data = stream.subarray(p + 4, p + 4 + len);
    p += 4 + len;
    if (type === R.CONTINUE && recs.length) {
      recs[recs.length - 1].conts.push(data);
      continue;
    }
    recs.push({ type, data, conts: [] });
    if (type === R.BOF || type === 0x0009 || type === 0x0209 || type === 0x0409) depth++;
    else if (type === R.EOF && --depth <= 0) break;
  }
  return { recs, end: p };
}

/** Reads across a record and its CONTINUE records; strings re-read the flags byte at a split. */
class Reader {
  private si = 0;
  pos: number;
  private readonly segs: Uint8Array[];

  constructor(rec: Rec, start = 0, private readonly decoder?: TextDecoder) {
    this.segs = [rec.data, ...rec.conts];
    this.pos = start;
  }

  private seg() {
    while (this.si < this.segs.length - 1 && this.pos >= this.segs[this.si].length) {
      this.si++;
      this.pos = 0;
    }
    return this.segs[this.si];
  }

  get eof() {
    const s = this.seg();
    return this.si === this.segs.length - 1 && this.pos >= s.length;
  }

  u8() {
    return this.seg()[this.pos++] ?? 0;
  }

  u16() {
    return this.u8() | (this.u8() << 8);
  }

  u32() {
    return (this.u16() | (this.u16() << 16)) >>> 0;
  }

  skip(n: number) {
    while (n > 0 && !this.eof) {
      const s = this.seg();
      const take = Math.min(n, s.length - this.pos);
      this.pos += take;
      n -= take;
    }
  }

  /** BIFF8 character data: `high` = UTF-16LE, otherwise compressed Latin-1. */
  chars(cch: number, high: boolean): string {
    let out = '';
    while (cch > 0) {
      let s = this.segs[this.si];
      if (this.pos >= s.length) {
        if (this.si >= this.segs.length - 1) break;
        this.si++;
        this.pos = 0;
        s = this.segs[this.si];
        high = (s[this.pos++] & 1) === 1; // continuation restates the encoding
      }
      const avail = high ? (s.length - this.pos) >> 1 : s.length - this.pos;
      const take = Math.min(cch, avail);
      if (take <= 0) { this.pos = s.length; continue; }
      const codes = new Array<number>(take);
      for (let i = 0; i < take; i++) {
        codes[i] = high ? s[this.pos + 2 * i] | (s[this.pos + 2 * i + 1] << 8) : s[this.pos + i];
      }
      for (let i = 0; i < take; i += 8192) out += String.fromCharCode.apply(null, codes.slice(i, i + 8192));
      this.pos += high ? take * 2 : take;
      cch -= take;
    }
    return out;
  }

  /** BIFF5 8-bit string in the workbook code page. */
  bytes8(n: number): string {
    const s = this.seg();
    const out = this.decoder!.decode(s.subarray(this.pos, this.pos + n));
    this.pos += n;
    return out;
  }

  /** XLUnicodeRichExtendedString (SST, LABEL, STRING, FORMAT in BIFF8). */
  unicode(cchBytes: 1 | 2): string {
    const cch = cchBytes === 1 ? this.u8() : this.u16();
    const flags = this.u8();
    const runs = flags & 0x08 ? this.u16() : 0;
    const ext = flags & 0x04 ? this.u32() : 0;
    const text = this.chars(cch, (flags & 1) === 1);
    this.skip(runs * 4 + ext);
    return text;
  }
}

function f64(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset + offset, 8).getFloat64(0, true);
}

function rk(v: number): number {
  let n: number;
  if (v & 0x02) {
    n = v >> 2; // signed 30-bit integer
  } else {
    const dv = new DataView(new ArrayBuffer(8));
    dv.setUint32(4, v & 0xfffffffc, true);
    n = dv.getFloat64(0, true);
  }
  return v & 0x01 ? n / 100 : n;
}

export async function openXls(buffer: ArrayBuffer, job: Job): Promise<Book> {
  const cfb = new CompoundFile(new Uint8Array(buffer));
  if (cfb.find('EncryptedPackage') || cfb.find('EncryptionInfo')) {
    throw new Error('This file is password-protected. Remove the password in Excel and try again.');
  }
  const entry = cfb.find('Workbook') ?? cfb.find('Book');
  if (!entry) throw new Error('This OLE2 file is not an Excel workbook (no Workbook stream). It may be a Word or other Office file.');
  const stream = cfb.read(entry);
  job.check();

  const { recs } = records(stream, 0);
  const bof = recs[0];
  if (!bof || bof.type !== R.BOF) {
    throw new Error('Excel 2.x–4.x worksheets are not supported. Re-save the file as .xlsx in Excel.');
  }
  const biff8 = (bof.data[0] | (bof.data[1] << 8)) === 0x0600;
  let decoder = new TextDecoder('windows-1252');
  let date1904 = false;
  const sheets: { name: string; pos: number }[] = [];
  const sst: string[] = [];
  const formats = new Map<number, Kind>();
  const xfFormat: number[] = [];

  for (const rec of recs) {
    const d = rec.data;
    switch (rec.type) {
      case R.FILEPASS:
        throw new Error('This file is password-protected. Remove the password in Excel and try again.');
      case R.CODEPAGE:
        try { decoder = new TextDecoder(codepageName(d[0] | (d[1] << 8))); } catch { /* keep default */ }
        break;
      case R.DATEMODE:
        date1904 = (d[0] | (d[1] << 8)) === 1;
        break;
      case R.BOUNDSHEET: {
        const pos = (d[0] | (d[1] << 8) | (d[2] << 16) | (d[3] << 24)) >>> 0;
        const kind = d[5];
        const r = new Reader(rec, 6, decoder);
        const name = biff8 ? r.unicode(1) : r.bytes8(r.u8());
        if (kind === 0 || kind === 1) sheets.push({ name, pos });
        break;
      }
      case R.SST: {
        const r = new Reader(rec, 8);
        const count = (d[4] | (d[5] << 8) | (d[6] << 16) | (d[7] << 24)) >>> 0;
        for (let i = 0; i < count && !r.eof; i++) sst.push(normalizeNewlines(r.unicode(2)));
        break;
      }
      case R.FORMAT:
      case R.FORMAT_OLD: {
        const r = new Reader(rec, 0, decoder);
        const id = r.u16();
        const code = biff8 ? r.unicode(2) : r.bytes8(r.u8());
        formats.set(id, kindOfFormat(code));
        break;
      }
      case R.XF:
        xfFormat.push(d[2] | (d[3] << 8));
        break;
    }
  }
  if (!sheets.length) throw new Error('The workbook has no worksheets.');

  const kindOf = (xf: number): Kind => {
    const id = xfFormat[xf];
    return id === undefined ? Kind.None : formats.get(id) ?? builtinKind(id);
  };
  const num = (n: number, xf: number) => formatNumber(n, kindOf(xf), date1904);

  return {
    format: biff8 ? 'Excel 97–2003 Workbook' : 'Excel 5.0/95 Workbook',
    sheets: sheets.map((s) => s.name),
    async load(index, job, out: RowEmitter) {
      const { recs } = records(stream, sheets[index].pos);
      const grid: string[][] = [];
      const set = (r: number, c: number, v: string) => {
        if (v === '') return;
        (grid[r] ??= [])[c] = v;
      };
      let pendingString: { r: number; c: number } | null = null;
      let depth = 0;

      for (let i = 0; i < recs.length; i++) {
        if ((i & 4095) === 0) await job.tick();
        const rec = recs[i];
        const d = rec.data;
        if (rec.type === R.BOF) { depth++; continue; }
        if (rec.type === R.EOF) { depth--; continue; }
        if (depth !== 1 || d.length < 6) {
          if (rec.type === R.STRING && pendingString && depth === 1) {
            const r = new Reader(rec, 0, decoder);
            set(pendingString.r, pendingString.c, normalizeNewlines(biff8 ? r.unicode(2) : r.bytes8(r.u16())));
            pendingString = null;
          }
          continue;
        }
        const row = d[0] | (d[1] << 8);
        const col = d[2] | (d[3] << 8);
        const xf = d[4] | (d[5] << 8);
        switch (rec.type) {
          case R.LABELSST:
            set(row, col, sst[(d[6] | (d[7] << 8) | (d[8] << 16) | (d[9] << 24)) >>> 0] ?? '');
            break;
          case R.LABEL:
          case R.RSTRING: {
            const r = new Reader(rec, 6, decoder);
            set(row, col, normalizeNewlines(biff8 ? r.unicode(2) : r.bytes8(r.u16())));
            break;
          }
          case R.NUMBER:
            set(row, col, num(f64(d, 6), xf));
            break;
          case R.RK:
            set(row, col, num(rk((d[6] | (d[7] << 8) | (d[8] << 16) | (d[9] << 24)) >>> 0), xf));
            break;
          case R.MULRK: {
            const last = d[d.length - 2] | (d[d.length - 1] << 8);
            for (let c = col, p = 4; c <= last && p + 6 <= d.length - 2; c++, p += 6) {
              const cxf = d[p] | (d[p + 1] << 8);
              set(row, c, num(rk((d[p + 2] | (d[p + 3] << 8) | (d[p + 4] << 16) | (d[p + 5] << 24)) >>> 0), cxf));
            }
            break;
          }
          case R.BOOLERR:
            set(row, col, d[7] ? ERRORS[d[6]] ?? '#ERR' : d[6] ? 'TRUE' : 'FALSE');
            break;
          case R.FORMULA:
            if (d[12] === 0xff && d[13] === 0xff) {
              const t = d[6];
              if (t === 0) pendingString = { r: row, c: col };
              else if (t === 1) set(row, col, d[8] ? 'TRUE' : 'FALSE');
              else if (t === 2) set(row, col, ERRORS[d[8]] ?? '#ERR');
            } else {
              set(row, col, num(f64(d, 6), xf));
            }
            break;
          case R.STRING:
            if (pendingString) {
              const r = new Reader(rec, 0, decoder);
              set(pendingString.r, pendingString.c, normalizeNewlines(biff8 ? r.unicode(2) : r.bytes8(r.u16())));
              pendingString = null;
            }
            break;
        }
      }
      for (let r = 0; r < grid.length; r++) if (grid[r]) out.row(r, grid[r]);
    },
  };
}

