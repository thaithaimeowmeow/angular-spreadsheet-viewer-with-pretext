/** .ods (OpenDocument Spreadsheet). Cell text is taken from <text:p>, i.e. already formatted. */
import type { Book, Job, RowEmitter } from './job';
import { formatGeneral } from './numfmt';
import { P, attr, own, scanStream, unescapeXml } from './xml';
import type { ZipArchive } from './zip';

const MAX_REPEAT_COLS = 1024;
const MAX_REPEAT_ROWS = 10_000;

// Table open / close tags and complete rows, in document order.
const TOKENS = new RegExp(
  `<${P}table(?=[\\s/>])([^>]*?)(/?)>|</${P}table\\s*>|<${P}table-row(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</${P}table-row\\s*>)`,
  'g',
);
const CELL = new RegExp(
  `<${P}(table-cell|covered-table-cell)(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</${P}\\1\\s*>)`,
  'g',
);
const PARAGRAPH = new RegExp(`<${P}(?:p|h)(?=[\\s/>])[^>]*?(?:/>|>([\\s\\S]*?)</${P}(?:p|h)\\s*>)`, 'g');
const ANNOTATION = new RegExp(`<${P}annotation(?=[\\s/>])[\\s\\S]*?</${P}annotation\\s*>`, 'g');
const INLINE = new RegExp(`<${P}(s|tab|line-break)(?=[\\s/>])([^>]*)>|<[^>]+>`, 'g');

function paragraphText(inner: string): string {
  return inner.replace(INLINE, (_, tag: string | undefined, attrs: string) => {
    if (tag === 's') return ' '.repeat(Math.max(1, Number(attr(attrs, 'c', true) ?? 1)));
    if (tag === 'tab') return '\t';
    if (tag === 'line-break') return '\n';
    return '';
  });
}

function cellText(attrs: string, inner: string | undefined): string {
  if (inner) {
    if (inner.indexOf('annotation') !== -1) inner = inner.replace(ANNOTATION, '');
    const parts: string[] = [];
    PARAGRAPH.lastIndex = 0;
    for (let m = PARAGRAPH.exec(inner); m; m = PARAGRAPH.exec(inner)) parts.push(m[1] ? paragraphText(m[1]) : '');
    if (parts.length) return unescapeXml(parts.join('\n'), false);
  }
  const type = attr(attrs, 'value-type', true);
  if (!type) return '';
  if (type === 'boolean') return attr(attrs, 'boolean-value', true) === 'true' ? 'TRUE' : 'FALSE';
  if (type === 'date') return (attr(attrs, 'date-value', true) ?? '').replace('T', ' ');
  if (type === 'time') return attr(attrs, 'time-value', true) ?? '';
  if (type === 'string') return attr(attrs, 'string-value', true) ?? '';
  const v = Number(attr(attrs, 'value', true));
  return Number.isFinite(v) ? formatGeneral(v) : '';
}

function parseRow(inner: string): string[] | null {
  const row: string[] = [];
  let col = 0;
  CELL.lastIndex = 0;
  for (let m = CELL.exec(inner); m; m = CELL.exec(inner)) {
    const repeat = Math.max(1, Number(attr(m[2], 'number-columns-repeated', true) ?? 1) || 1);
    const text = cellText(m[2], m[3]);
    if (text === '') {
      col += repeat;
      continue;
    }
    const stored = own(text); // one copy, shared by repeated columns
    for (let k = 0; k < Math.min(repeat, MAX_REPEAT_COLS); k++) row[col + k] = stored;
    col += repeat;
  }
  return row.length ? row : null;
}

/** Walks content.xml; calls back with (tableIndex, rowAttrs, rowInner). Return true to stop. */
async function walk(
  zip: ZipArchive,
  job: Job,
  onTable: (index: number, attrs: string) => boolean | void,
  onRow: (table: number, attrs: string, inner: string | undefined) => boolean | void,
  onBytes?: (n: number) => void,
) {
  const entry = zip.find('content.xml');
  if (!entry) throw new Error('The OpenDocument file has no content.xml.');
  let tableCount = 0;
  let current = -1;
  await scanStream(await zip.text(entry, onBytes), ['table-row', 'table'], job, (chunk) => {
    TOKENS.lastIndex = 0;
    for (let m = TOKENS.exec(chunk); m; m = TOKENS.exec(chunk)) {
      if (m[0].startsWith('</')) {
        current = -1;
      } else if (m[1] !== undefined) {
        current = tableCount++;
        if (onTable(current, m[1])) return true;
        if (m[2] === '/') current = -1;
      } else if (current >= 0 && onRow(current, m[3], m[4])) {
        return true;
      }
    }
    return false;
  });
}

export async function openOds(zip: ZipArchive, job: Job): Promise<Book> {
  const names: string[] = [];
  await walk(zip, job, (i, attrs) => { names[i] = attr(attrs, 'name', true) ?? `Sheet${i + 1}`; }, () => {});
  if (!names.length) throw new Error('The spreadsheet has no sheets.');

  return {
    format: 'OpenDocument Spreadsheet',
    sheets: names,
    async load(index, job, out: RowEmitter) {
      const entry = zip.find('content.xml')!;
      let r = 0;
      await walk(zip, job, (i) => i > index, (table, attrs, inner) => {
        if (table !== index) return false;
        const repeat = Math.max(1, Number(attr(attrs, 'number-rows-repeated', true) ?? 1) || 1);
        const row = inner ? parseRow(inner) : null;
        if (row) for (let k = 0; k < Math.min(repeat, MAX_REPEAT_ROWS); k++) out.row(r + k, row);
        r += repeat; // trailing empty (often 1M repeated) rows are never materialized
        return false;
      }, job.byteProgress('rows', entry.compSize));
    },
  };
}
