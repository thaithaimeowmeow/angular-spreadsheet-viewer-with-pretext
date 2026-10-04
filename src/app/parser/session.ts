/**
 * One sheet session in the worker: owns the parsed rows of the visible sheet and answers the UI's
 * queries (row windows, height estimates, search, autofit, copy). Only the current sheet is kept,
 * so switching sheets frees the previous one.
 */
import { DEFAULT_COL, MAX_AUTO_COL, estimateColWidth, estimateRowHeight } from '../grid/metrics';
import { SheetSearch } from '../grid/sheet-search';
import { type RowEmitter, isActiveJob, macrotask, post } from './job';
import type { Row } from './protocol';

/** Short values repeat a lot (status, city, gender...): keep one shared copy of each. */
const INTERN_MAX_LEN = 24;
const INTERN_MAX_ENTRIES = 200_000;
const AUTO_SIZE_ROWS = 400;
const ESTIMATE_CHUNK = 4096;
const SLICE_MS = 10;
const MAX_COPY_CELLS = 3_000_000;

export class SheetSession implements RowEmitter {
  /** Holey: rows never stored (empty) read as undefined. */
  private readonly rows: Row[] = [];
  private next = 0;
  private colCount = 0;
  private complete = false;
  private readonly startedAt = performance.now();
  private lastGrow = 0;
  private columnsSent = false;
  private readonly interned = new Map<string, string>();

  private widths: number[] = [];
  private version = 0;
  private layoutCursor = 0;
  private layoutRunning = false;
  private layoutDoneVersion = -1;

  private readonly search = new SheetSearch();
  private searchId = 0;
  private searchSent = 0;
  private searchDoneSent = false;
  private searchRunning = false;

  constructor(readonly job: number, private wrap: boolean) {}

  private get alive() {
    return isActiveJob(this.job);
  }

  // ------------------------------------------------------------ rows from the parser

  row(r: number, row: Row) {
    if (row) {
      for (let c = 0; c < row.length; c++) {
        const t = row[c];
        if (!t || t.length > INTERN_MAX_LEN) continue;
        const shared = this.interned.get(t);
        if (shared !== undefined) row[c] = shared;
        else if (this.interned.size < INTERN_MAX_ENTRIES) this.interned.set(t, t);
      }
      if (row.length > this.colCount) this.colCount = row.length;
    }
    if (r < this.next) {
      // Out-of-order row: replace it and tell the UI to drop any cached copy.
      this.rows[r] = row;
      this.layoutCursor = Math.min(this.layoutCursor, r);
      post({ type: 'invalidate', job: this.job, start: r, end: r + 1 });
      return;
    }
    // Empty rows only exist as gaps before a later non-empty row, so trailing empty rows
    // (Excel often stores thousands of blank, formatted rows) never appear.
    if (!row) return;
    this.rows[r] = row;
    this.next = r + 1;
    this.maybeGrow();
  }

  private maybeGrow() {
    const now = performance.now();
    if (!this.columnsSent) {
      if (this.next < 100 && now - this.startedAt < 30) return;
      this.sendColumns();
    } else if (now - this.lastGrow < 100) {
      return;
    }
    this.lastGrow = now;
    post({ type: 'grow', job: this.job, rowCount: this.next, colCount: this.colCount });
    this.kickLayout();
    this.kickSearch();
  }

  /** Auto-sizes columns from the first rows (estimated text widths). */
  private sendColumns() {
    this.columnsSent = true;
    const widths = new Array<number>(this.colCount).fill(0);
    const sample = Math.min(this.next, AUTO_SIZE_ROWS);
    for (let r = 0; r < sample; r++) {
      const row = this.rows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const t = row[c];
        if (t && widths[c] < MAX_AUTO_COL) widths[c] = Math.max(widths[c], estimateColWidth(t));
      }
    }
    this.widths = widths.map((w) => w || DEFAULT_COL);
    post({ type: 'columns', job: this.job, widths: this.widths });
  }

  finish() {
    this.complete = true;
    if (!this.columnsSent) this.sendColumns();
    post({ type: 'grow', job: this.job, rowCount: this.next, colCount: this.colCount });
    post({ type: 'done', job: this.job, rowCount: this.next, colCount: this.colCount });
    this.kickLayout();
    this.kickSearch();
  }

  // ------------------------------------------------------------ height estimates

  setLayout(version: number, widths: number[], wrap: boolean) {
    if (version <= this.version) return;
    this.version = version;
    this.widths = widths;
    this.wrap = wrap;
    this.layoutCursor = 0;
    this.kickLayout();
  }

  private kickLayout() {
    if (!this.layoutRunning) void this.layoutLoop();
  }

  private async layoutLoop() {
    this.layoutRunning = true;
    try {
      while (this.alive) {
        const version = this.version;
        // Wrap off: the UI uses the base height everywhere, nothing to estimate.
        if (!this.wrap || this.layoutCursor >= this.next) {
          if ((!this.wrap || this.complete) && this.layoutDoneVersion !== version) {
            this.layoutDoneVersion = version;
            post({ type: 'layoutDone', job: this.job, version });
          }
          return; // more rows, finish() or a layout change will kick the loop again
        }
        const start = this.layoutCursor;
        const end = Math.min(this.next, start + ESTIMATE_CHUNK);
        const out = new Uint16Array(end - start);
        const t0 = performance.now();
        let r = start;
        while (r < end) {
          out[r - start] = estimateRowHeight(this.rows[r] ?? null, this.widths);
          r++;
          if ((r & 63) === 0 && performance.now() - t0 > SLICE_MS) break;
        }
        if (version === this.version && start === this.layoutCursor) {
          this.layoutCursor = r;
          const heights = r === end ? out : out.slice(0, r - start);
          post({ type: 'heights', job: this.job, version, start, heights }, [heights.buffer]);
        }
        await macrotask();
      }
    } finally {
      this.layoutRunning = false;
    }
  }

  // ------------------------------------------------------------ queries

  window(reqId: number, start: number, end: number) {
    const stop = Math.min(end, this.next);
    const rows: Row[] = new Array(Math.max(0, stop - start));
    for (let i = 0; i < rows.length; i++) rows[i] = this.rows[start + i] ?? null;
    post({ type: 'window', job: this.job, reqId, start, rows });
  }

  autofit(reqId: number, col: number, from: number) {
    let width = 48;
    const scan = (r: number) => {
      const t = this.rows[r]?.[col];
      if (t) width = Math.max(width, estimateColWidth(t, MAX_AUTO_COL * 2));
    };
    const head = Math.min(this.next, 1000);
    for (let r = 0; r < head; r++) scan(r);
    for (let r = Math.max(from, head); r < Math.min(this.next, from + 200); r++) scan(r);
    post({ type: 'autofit', job: this.job, reqId, width });
  }

  copy(reqId: number, r0: number, r1: number, c0: number, c1: number) {
    r1 = Math.min(r1, this.next - 1);
    const cells = Math.max(0, r1 - r0 + 1) * Math.max(0, c1 - c0 + 1);
    if (cells > MAX_COPY_CELLS) {
      post({
        type: 'copy', job: this.job, reqId, text: null, cells,
        error: `Selection too large to copy: ${cells.toLocaleString()} cells (limit ${MAX_COPY_CELLS.toLocaleString()}).`,
      });
      return;
    }
    const lines: string[] = [];
    const out: string[] = [];
    for (let r = r0; r <= r1; r++) {
      const row = this.rows[r];
      out.length = 0;
      for (let c = c0; c <= c1; c++) {
        const t = row?.[c] ?? '';
        out.push(/[\t\n\r"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
      }
      lines.push(out.join('\t'));
    }
    post({ type: 'copy', job: this.job, reqId, text: lines.join('\r\n'), cells });
  }

  // ------------------------------------------------------------ search

  setSearch(searchId: number, query: string) {
    this.searchId = searchId;
    this.searchSent = 0;
    this.searchDoneSent = false;
    if (!this.search.setQuery(query)) {
      post({ type: 'searchHits', job: this.job, searchId, keys: new Float64Array(0), done: true, truncated: false });
      this.searchDoneSent = true;
      return;
    }
    this.kickSearch();
  }

  private kickSearch() {
    if (!this.searchRunning && this.search.active && !this.searchDoneSent) void this.searchLoop();
  }

  private async searchLoop() {
    this.searchRunning = true;
    try {
      while (this.alive && this.search.active) {
        const id = this.searchId;
        const caughtUp = this.search.scan(this.rows, SLICE_MS);
        if (id !== this.searchId) continue;
        const done = caughtUp && this.complete;
        const count = this.search.count;
        if (count > this.searchSent || (done && !this.searchDoneSent)) {
          const keys = this.search.keysSlice(this.searchSent, count);
          this.searchSent = count;
          this.searchDoneSent = done;
          post({ type: 'searchHits', job: this.job, searchId: id, keys, done, truncated: this.search.truncated }, [keys.buffer]);
        }
        if (caughtUp) return; // more rows (or finish) will kick the loop again
        await macrotask();
      }
    } finally {
      this.searchRunning = false;
    }
  }
}
