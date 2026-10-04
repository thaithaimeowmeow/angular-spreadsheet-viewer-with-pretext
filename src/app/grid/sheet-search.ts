/**
 * Incremental sheet search. One compiled RegExp per query, tested against every cell in time
 * slices; matches are kept as sorted numeric keys (row * MAX_COLS + col) for O(log n) next/prev.
 *
 * Accent-insensitive for Vietnamese by default ("nguyen" finds "Nguyễn"); a letter typed WITH
 * an accent must match that accent. Case-insensitive.
 */
import type { Row } from '../parser/protocol';

export const MAX_COLS = 16384;
const MAX_MATCHES = 1_000_000;

const VARIANTS: Record<string, string> = {
  a: 'aàáạảãâầấậẩẫăằắặẳẵ',
  e: 'eèéẹẻẽêềếệểễ',
  i: 'iìíịỉĩ',
  o: 'oòóọỏõôồốộổỗơờớợởỡ',
  u: 'uùúụủũưừứựửữ',
  y: 'yỳýỵỷỹ',
  d: 'dđ',
};

// Only syntax characters may be escaped in 'u' mode (e.g. "\-" would throw).
const escapeRe = (ch: string) => ch.replace(/[\\^$.*+?()[\]{}|/]/g, '\\$&');

/** Builds the matcher source for `query`, or null for an empty query. */
export function matcherSource(query: string): string | null {
  query = query.normalize('NFC');
  if (!query.trim()) return null;
  let src = '';
  for (const ch of query.toLowerCase()) {
    const variants = VARIANTS[ch];
    // Plain letters match every accented form; a combining-mark tail tolerates NFD text.
    src += variants ? `[${variants}][\\u0300-\\u036f]*` : escapeRe(ch);
  }
  return src;
}

export interface TextPart { t: string; hit: boolean }

/** Splits `text` into highlighted / plain parts for rendering. */
export function highlightParts(text: string, global: RegExp): TextPart[] | null {
  global.lastIndex = 0;
  let m = global.exec(text);
  if (!m) return null;
  const parts: TextPart[] = [];
  let last = 0;
  while (m) {
    if (m.index > last) parts.push({ t: text.slice(last, m.index), hit: false });
    parts.push({ t: m[0], hit: true });
    last = m.index + m[0].length;
    if (m[0].length === 0) global.lastIndex++;
    m = global.exec(text);
  }
  if (last < text.length) parts.push({ t: text.slice(last), hit: false });
  return parts;
}

export class SheetSearch {
  /** Test regex (no 'g': no lastIndex state) and highlight regex ('g'). */
  test: RegExp | null = null;
  global: RegExp | null = null;
  private keys = new Float64Array(1024);
  count = 0;
  /** Next row to scan. */
  cursor = 0;
  truncated = false;

  setQuery(query: string): boolean {
    const src = matcherSource(query);
    this.test = src ? new RegExp(src, 'iu') : null;
    this.global = src ? new RegExp(src, 'giu') : null;
    this.count = 0;
    this.cursor = 0;
    this.truncated = false;
    return !!src;
  }

  get active() {
    return this.test !== null;
  }

  /** Scans rows from the cursor for up to `budgetMs`. Returns true when caught up with `rows`. */
  scan(rows: Row[], budgetMs: number): boolean {
    const re = this.test;
    if (!re) return true;
    const t0 = performance.now();
    const n = rows.length;
    while (this.cursor < n) {
      const stop = Math.min(n, this.cursor + 256);
      for (let r = this.cursor; r < stop; r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          const text = row[c];
          if (text && re.test(text)) this.push(r * MAX_COLS + c);
        }
      }
      this.cursor = stop;
      if (this.truncated) {
        this.cursor = n;
        break;
      }
      if (performance.now() - t0 > budgetMs) break;
    }
    return this.cursor >= n;
  }

  key(i: number) {
    return this.keys[i];
  }

  /** Copy of match keys [from, to) (worker → UI transfer). */
  keysSlice(from: number, to: number): Float64Array {
    return this.keys.slice(from, to);
  }

  /** Appends match keys found elsewhere (the UI receives them from the worker). */
  append(keys: Float64Array) {
    for (let i = 0; i < keys.length; i++) this.push(keys[i]);
  }

  /** Index of the first match with key > k (dir 1) or the last with key < k (dir -1), wrapping. */
  neighbor(k: number, dir: 1 | -1): number {
    if (!this.count) return -1;
    let lo = 0;
    let hi = this.count;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.keys[mid] <= k) lo = mid + 1;
      else hi = mid;
    }
    // lo = first index with key > k
    if (dir === 1) return lo < this.count ? lo : 0;
    const before = this.keys[lo - 1] === k ? lo - 2 : lo - 1;
    return before >= 0 ? before : this.count - 1;
  }

  indexOf(k: number): number {
    let lo = 0;
    let hi = this.count - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const v = this.keys[mid];
      if (v === k) return mid;
      if (v < k) lo = mid + 1;
      else hi = mid - 1;
    }
    return -1;
  }

  private push(k: number) {
    if (this.count >= MAX_MATCHES) {
      this.truncated = true;
      return;
    }
    if (this.count === this.keys.length) {
      const next = new Float64Array(this.keys.length * 2);
      next.set(this.keys);
      this.keys = next;
    }
    this.keys[this.count++] = k;
  }
}
