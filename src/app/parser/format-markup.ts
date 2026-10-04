/**
 * Markup "spreadsheets" that are often saved with an .xls extension:
 *  - HTML tables (very common from web/report exports)
 *  - SpreadsheetML 2003 (Excel "XML Spreadsheet")
 * Both are parsed in memory; such files are rarely large.
 */
import type { Book, RowEmitter } from './job';
import { formatGeneral } from './numfmt';
import { P, attr, element, own, unescapeXml } from './xml';

type Grid = string[][];

const HTML_ENTITIES: Record<string, string> = {
  nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", copy: '©', reg: '®', deg: '°',
  ndash: '–', mdash: '—', hellip: '…', laquo: '«', raquo: '»', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”',
  middot: '·', times: '×', divide: '÷', euro: '€', bull: '•', trade: '™',
};

function decodeHtml(s: string): string {
  return s.replace(/&(#\d+|#[xX][0-9a-fA-F]+|[a-zA-Z]+);?/g, (m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : +e.slice(1);
      return cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : '';
    }
    return HTML_ENTITIES[e] ?? HTML_ENTITIES[e.toLowerCase()] ?? m;
  });
}

function htmlCellText(html: string): string {
  return decodeHtml(
    html
      .replace(/<!--[\s\S]*?-->/g, '')
      .replace(/<(style|script)\b[\s\S]*?<\/\1\s*>/gi, '')
      .replace(/\s+/g, ' ')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|li)\s*>/gi, '\n')
      .replace(/<[^>]*>/g, ''),
  )
    .replace(/ /g, ' ')
    .replace(/ *\n */g, '\n')
    .trim();
}

/** Places cells honoring colspan/rowspan. */
function placer() {
  const grid: Grid = [];
  const taken = new Map<number, Set<number>>();
  return {
    grid,
    put(r: number, c: number, text: string, rowspan: number, colspan: number): number {
      while (taken.get(r)?.has(c)) c++;
      if (text !== '') (grid[r] ??= [])[c] = text;
      for (let dr = 0; dr < Math.min(rowspan, 1000); dr++) {
        for (let dc = 0; dc < Math.min(colspan, 1000); dc++) {
          if (dr || dc) {
            let set = taken.get(r + dr);
            if (!set) taken.set(r + dr, (set = new Set()));
            set.add(c + dc);
          }
        }
      }
      return c + colspan;
    },
    skipTaken(r: number, c: number) {
      while (taken.get(r)?.has(c)) c++;
      return c;
    },
  };
}

function emit(grid: Grid, out: RowEmitter) {
  for (let r = 0; r < grid.length; r++) if (grid[r]) out.row(r, grid[r]);
}

export function openHtml(text: string, name: string): Book {
  const body = text.replace(/<!--[\s\S]*?-->/g, '').replace(/<(style|script)\b[\s\S]*?<\/\1\s*>/gi, '');
  // Split on <table> openings; handles tables that are never closed.
  const tables = body.split(/<table\b[^>]*>/i).slice(1).map((t) => t.split(/<\/table\s*>/i)[0]);
  const grids: Grid[] = [];
  for (const table of tables) {
    const g = placer();
    let r = 0;
    for (const tr of table.split(/<tr\b[^>]*>/i).slice(1)) {
      const rowHtml = tr.split(/<\/tr\s*>/i)[0];
      let c = 0;
      const cells = rowHtml.split(/(?=<t[dh]\b)/i).filter((x) => /^<t[dh]\b/i.test(x));
      for (const cell of cells) {
        const open = /^<t[dh]\b([^>]*)>/i.exec(cell)!;
        const inner = cell.slice(open[0].length).replace(/<\/t[dh]\s*>[\s\S]*$/i, '');
        const rowspan = Math.max(1, parseInt(/rowspan\s*=\s*["']?(\d+)/i.exec(open[1])?.[1] ?? '1', 10) || 1);
        const colspan = Math.max(1, parseInt(/colspan\s*=\s*["']?(\d+)/i.exec(open[1])?.[1] ?? '1', 10) || 1);
        c = g.put(r, c, own(htmlCellText(inner)), rowspan, colspan);
      }
      r++;
    }
    if (g.grid.length) grids.push(g.grid);
  }
  if (!grids.length) throw new Error('No table found in this HTML file.');
  return {
    format: 'HTML table',
    sheets: grids.length === 1 ? [name] : grids.map((_, i) => `Table ${i + 1}`),
    async load(index, _job, out) {
      emit(grids[index], out);
    },
  };
}

export function isSpreadsheetMl(head: string): boolean {
  return /urn:schemas-microsoft-com:office:spreadsheet/.test(head) && /<(?:[\w.-]+:)?Workbook\b/.test(head);
}

export function openSpreadsheetMl(text: string): Book {
  const sheets: { name: string; xml: string }[] = [];
  for (const m of text.matchAll(element('Worksheet'))) {
    sheets.push({ name: attr(m[1], 'Name', true) ?? `Sheet${sheets.length + 1}`, xml: m[2] ?? '' });
  }
  if (!sheets.length) throw new Error('The XML spreadsheet has no worksheets.');
  const rowRe = element('Row');
  const cellRe = element('Cell');
  const dataRe = new RegExp(`<${P}Data(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</${P}Data\\s*>)`);

  return {
    format: 'XML Spreadsheet 2003',
    sheets: sheets.map((s) => s.name),
    async load(index, _job, out) {
      const g = placer();
      let r = 0;
      for (const row of sheets[index].xml.matchAll(rowRe)) {
        const ri = Number(attr(row[1], 'Index', true));
        if (ri > 0) r = ri - 1;
        let c = 0;
        for (const cell of (row[2] ?? '').matchAll(cellRe)) {
          const ci = Number(attr(cell[1], 'Index', true));
          c = ci > 0 ? ci - 1 : g.skipTaken(r, c);
          const across = Number(attr(cell[1], 'MergeAcross', true) ?? 0) + 1;
          const down = Number(attr(cell[1], 'MergeDown', true) ?? 0) + 1;
          const data = dataRe.exec(cell[2] ?? '');
          let value = data ? unescapeXml((data[2] ?? '').replace(/<[^>]*>/g, ''), false) : '';
          const type = data ? attr(data[1], 'Type', true) : undefined;
          if (type === 'Number' && value.trim() !== '' && Number.isFinite(Number(value))) value = formatGeneral(Number(value));
          else if (type === 'Boolean') value = value.trim() === '1' ? 'TRUE' : 'FALSE';
          else if (type === 'DateTime') value = value.replace(/T00:00:00(\.0+)?$/, '').replace('T', ' ').replace(/\.0+$/, '');
          c = g.put(r, c, own(value), down, across);
        }
        r++;
      }
      emit(g.grid, out);
    },
  };
}
