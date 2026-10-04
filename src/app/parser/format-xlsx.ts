/** .xlsx / .xlsm / .xltx / .xltm (Office Open XML, transitional and strict). */
import type { Book, Job, RowEmitter } from './job';
import { Kind, builtinKind, formatNumber, kindOfFormat } from './numfmt';
import type { Row } from './protocol';
import { P, attr, element, own, scanElements, unescapeXml } from './xml';
import type { ZipArchive } from './zip';

interface Rel { id: string; type: string; target: string }

function resolvePath(base: string, target: string): string {
  if (/^[a-z]+:/i.test(target)) return target; // external link
  if (target.startsWith('/')) return target.slice(1);
  const out: string[] = [];
  for (const part of (base + target).split('/')) {
    if (part === '..') out.pop();
    else if (part !== '.' && part !== '') out.push(part);
  }
  return out.join('/');
}

function parseRels(xml: string | null, base: string): Rel[] {
  if (!xml) return [];
  return [...xml.matchAll(element('Relationship'))].map((m) => ({
    id: attr(m[1], 'Id') ?? '',
    type: attr(m[1], 'Type') ?? '',
    target: resolvePath(base, attr(m[1], 'Target') ?? ''),
  }));
}

function relsPathFor(part: string): string {
  const i = part.lastIndexOf('/');
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}

function parseStyles(xml: string | null): Kind[] {
  if (!xml) return [];
  const custom = new Map<number, Kind>();
  for (const m of xml.matchAll(element('numFmt'))) {
    custom.set(Number(attr(m[1], 'numFmtId')), kindOfFormat(attr(m[1], 'formatCode') ?? ''));
  }
  const xfs = element('cellXfs', '').exec(xml)?.[2];
  if (!xfs) return [];
  const kinds: Kind[] = [];
  // Opening tags only: an <xf> may contain children such as <alignment/>.
  for (const m of xfs.matchAll(new RegExp(`<${P}xf(?=[\\s/>])([^>]*)>`, 'g'))) {
    const id = Number(attr(m[1], 'numFmtId') ?? 0);
    kinds.push(custom.get(id) ?? builtinKind(id));
  }
  return kinds;
}

const T_TAG = element('t');
const PHONETIC = element('rPh');

/** Concatenates every <t> run (rich text), ignoring phonetic hints. */
function collectText(inner: string): string {
  if (inner.indexOf('rPh') !== -1) inner = inner.replace(PHONETIC, '');
  let out = '';
  T_TAG.lastIndex = 0;
  for (let m = T_TAG.exec(inner); m; m = T_TAG.exec(inner)) out += m[2] ?? '';
  return own(unescapeXml(out));
}

function colIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

const CELL = element('c');
const CELL_REF = /(?:^|\s)r\s*=\s*["']\$?([A-Za-z]+)\$?\d*["']/;
const CELL_TYPE = /(?:^|\s)t\s*=\s*["'](\w+)["']/;
const CELL_STYLE = /(?:^|\s)s\s*=\s*["'](\d+)["']/;
const VALUE = element('v', '');
const INLINE = element('is', '');
const ROW_REF = /(?:^|\s)r\s*=\s*["'](\d+)["']/;

export async function openXlsx(zip: ZipArchive, job: Job): Promise<Book> {
  const rootRels = parseRels(await zip.readText('_rels/.rels'), '');
  const wbPath = rootRels.find((r) => /\/officeDocument$/.test(r.type))?.target ?? 'xl/workbook.xml';
  if (/\.bin$/i.test(wbPath)) {
    throw new Error('Excel Binary Workbook (.xlsb) is not supported. Save it as .xlsx in Excel and try again.');
  }
  const wbXml = await zip.readText(wbPath);
  if (!wbXml) throw new Error('This ZIP file does not contain an Excel workbook.');
  const wbBase = wbPath.slice(0, wbPath.lastIndexOf('/') + 1);
  const wbRels = parseRels(await zip.readText(relsPathFor(wbPath)), wbBase);
  const relById = new Map(wbRels.map((r) => [r.id, r]));
  const date1904 = new RegExp(`<${P}workbookPr\\b[^>]*\\bdate1904\\s*=\\s*["'](1|true|on)["']`).test(wbXml);

  const sheets: { name: string; path: string }[] = [];
  for (const m of wbXml.matchAll(element('sheet'))) {
    const rid = /(?:^|\s)[\w.-]+:id\s*=\s*["']([^"']+)["']/.exec(m[1])?.[1];
    const rel = rid ? relById.get(rid) : undefined;
    if (!rel || /\/chartsheet$/.test(rel.type) || !zip.find(rel.target)) continue;
    sheets.push({ name: attr(m[1], 'name') ?? `Sheet${sheets.length + 1}`, path: rel.target });
  }
  if (!sheets.length) throw new Error('The workbook has no worksheets.');

  const stylesRel = wbRels.find((r) => /\/styles$/.test(r.type));
  const styleKinds = parseStyles(stylesRel ? await zip.readText(stylesRel.target) : null);

  const sharedStrings: string[] = [];
  const ssRel = wbRels.find((r) => /\/sharedStrings$/.test(r.type));
  const ssEntry = ssRel ? zip.find(ssRel.target) : undefined;
  if (ssEntry) {
    const stream = await zip.text(ssEntry, job.byteProgress('strings', ssEntry.compSize));
    await scanElements(stream, 'si', job, (_a, inner) => sharedStrings.push(inner ? collectText(inner) : ''));
  }

  function cellValue(attrs: string, body: string): string {
    const type = CELL_TYPE.exec(attrs)?.[1];
    if (type === 'inlineStr') {
      const is = INLINE.exec(body)?.[2];
      return is ? collectText(is) : '';
    }
    const v = VALUE.exec(body)?.[2];
    if (v === undefined) return '';
    switch (type) {
      case 's': return sharedStrings[+v] ?? '';
      case 'b': return v.trim() === '1' || v.trim() === 'true' ? 'TRUE' : 'FALSE';
      case 'str':
      case 'e':
      case 'd': return own(unescapeXml(v));
    }
    const n = Number(v);
    if (v.trim() === '' || !Number.isFinite(n)) return own(unescapeXml(v));
    const style = CELL_STYLE.exec(attrs);
    return formatNumber(n, style ? styleKinds[+style[1]] ?? Kind.None : Kind.None, date1904);
  }

  function parseRow(inner: string): Row {
    const row: string[] = [];
    let col = 0;
    CELL.lastIndex = 0;
    for (let m = CELL.exec(inner); m; m = CELL.exec(inner)) {
      const ref = CELL_REF.exec(m[1]);
      if (ref) col = colIndex(ref[1].toUpperCase());
      if (m[2]) {
        const value = cellValue(m[1], m[2]);
        if (value !== '') row[col] = value;
      }
      col++;
    }
    return row.length ? row : null;
  }

  return {
    format: 'Excel Workbook',
    sheets: sheets.map((s) => s.name),
    async load(index, job, out: RowEmitter) {
      const entry = zip.find(sheets[index].path)!;
      const stream = await zip.text(entry, job.byteProgress('rows', entry.compSize));
      let next = 0;
      await scanElements(stream, 'row', job, (attrs, inner) => {
        const ref = ROW_REF.exec(attrs);
        const r = ref ? +ref[1] - 1 : next;
        next = r + 1;
        out.row(r, inner ? parseRow(inner) : null);
      });
    },
  };
}
