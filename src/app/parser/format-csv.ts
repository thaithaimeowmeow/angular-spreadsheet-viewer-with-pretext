/** Delimited text: .csv / .tsv / .txt. Streams the file; delimiter and encoding are detected. */
import type { Book, Job, RowEmitter } from './job';
import { normalizeNewlines, own } from './xml';

const QUOTE = 34;
const CR = 13;
const LF = 10;
const CANDIDATES = [',', ';', '\t', '|'];

/** Picks the delimiter whose per-line count is most consistent over the first lines. */
export function detectDelimiter(sample: string, preferTab: boolean): string {
  const lines: string[] = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < sample.length && lines.length < 50; i++) {
    const ch = sample[i];
    if (ch === '"') inQuotes = !inQuotes;
    if (!inQuotes && (ch === '\n' || ch === '\r')) {
      if (cur) lines.push(cur);
      cur = '';
    } else if (!inQuotes) {
      cur += ch;
    }
  }
  if (cur && lines.length < 50) lines.push(cur);
  if (!lines.length) return preferTab ? '\t' : ',';

  let best = preferTab ? '\t' : ',';
  let bestScore = 0;
  for (const d of CANDIDATES) {
    const counts = lines.map((l) => l.split(d).length - 1);
    const freq = new Map<number, number>();
    for (const c of counts) if (c > 0) freq.set(c, (freq.get(c) ?? 0) + 1);
    let score = 0;
    for (const [count, times] of freq) score = Math.max(score, times * 1000 + count);
    if (score > bestScore || (score === bestScore && preferTab && d === '\t')) {
      best = d;
      bestScore = score;
    }
  }
  return best;
}

export function openCsv(file: File, encoding: string, sample: string, name: string): Book {
  let delimiter = detectDelimiter(sample, /\.tsv$/i.test(file.name));
  // Excel's "sep=;" first-line directive.
  const sep = /^(?:﻿)?sep=(.)\r?\n/i.exec(sample);
  if (sep) delimiter = sep[1];

  return {
    format: delimiter === '\t' ? 'Tab-separated text' : `Delimited text (${delimiter === ',' ? 'comma' : delimiter === ';' ? 'semicolon' : delimiter})`,
    sheets: [name],
    async load(_index, job: Job, out: RowEmitter) {
      const delim = delimiter.charCodeAt(0);
      const onBytes = job.byteProgress('rows', file.size);
      const reader = file.stream()
        .pipeThrough(new TransformStream<Uint8Array, Uint8Array>({
          transform(chunk, ctl) { onBytes(chunk.byteLength); ctl.enqueue(chunk); },
        }))
        .pipeThrough(new TextDecoderStream(encoding) as unknown as TransformStream<Uint8Array, string>)
        .getReader();

      // States: 0 field start, 1 unquoted, 2 quoted, 3 quote seen inside quoted field.
      let state = 0;
      let field = '';
      let row: string[] = [];
      let r = 0;
      let skipLF = false;
      let skipDirective = !!sep;

      const endField = () => {
        row.push(own(normalizeNewlines(field)));
        field = '';
      };
      const endRow = () => {
        if (skipDirective) {
          skipDirective = false;
        } else {
          const empty = row.length === 1 && row[0] === '';
          let last = row.length;
          while (last > 0 && row[last - 1] === '') last--;
          if (last < row.length) row.length = last;
          out.row(r++, empty || !row.length ? null : row);
        }
        row = [];
      };

      try {
        for (;;) {
          const { done, value: s } = await reader.read();
          if (done) break;
          await job.tick();
          let segStart = 0;
          const n = s.length;
          for (let i = 0; i < n; i++) {
            const c = s.charCodeAt(i);
            if (skipLF) {
              skipLF = false;
              if (c === LF && state !== 2) { segStart = i + 1; continue; }
            }
            switch (state) {
              case 0:
                if (c === QUOTE) { state = 2; segStart = i + 1; }
                else if (c === delim) endField();
                else if (c === CR || c === LF) { endField(); endRow(); skipLF = c === CR; }
                else { state = 1; segStart = i; }
                break;
              case 1:
                if (c === delim) { field += s.slice(segStart, i); endField(); state = 0; }
                else if (c === CR || c === LF) { field += s.slice(segStart, i); endField(); endRow(); state = 0; skipLF = c === CR; }
                break;
              case 2:
                if (c === QUOTE) { field += s.slice(segStart, i); state = 3; }
                break;
              case 3:
                if (c === QUOTE) { field += '"'; state = 2; segStart = i + 1; }
                else if (c === delim) { endField(); state = 0; }
                else if (c === CR || c === LF) { endField(); endRow(); state = 0; skipLF = c === CR; }
                else { state = 1; segStart = i; } // text after a closing quote: keep it, like Excel
                break;
            }
          }
          if (state === 1 || state === 2) field += s.slice(segStart, n);
        }
        if (state !== 0 || field !== '' || row.length) {
          endField();
          endRow();
        }
      } catch (err) {
        reader.cancel().catch(() => {});
        throw err;
      }
    },
  };
}
