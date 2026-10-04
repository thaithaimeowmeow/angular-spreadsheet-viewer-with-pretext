import { signal } from '@angular/core';
import { BLOCK_ROWS, type FromWorker, type Row, type ToWorker } from '../parser/protocol';

/** Row blocks kept in the UI (64 × 256 = 16k rows); everything else stays in the worker. */
const MAX_BLOCKS = 64;

export interface CopyResult { text: string | null; cells: number; error?: string }

/**
 * The UI's view of one sheet session. Rows live in the worker; this keeps an LRU cache of row
 * blocks fetched for what is on screen, and relays layout / search / autofit / copy requests.
 */
export class SheetSource {
  rowCount = 0;
  colCount = 0;
  complete = false;
  /** Auto-sized widths from the worker (null until the first rows are parsed). */
  widths: number[] | null = null;
  /** Layout version of the last widths/wrap sent; worker heights for other versions are dropped. */
  layoutVersion = 0;

  /** Bumped when counts, cached blocks or widths change. */
  readonly dataVersion = signal(0);
  /** Share of rows with a height estimate for the current layout (0..1). */
  readonly layoutProgress = signal(0);

  onHeights: ((start: number, heights: Uint16Array) => void) | null = null;
  onInvalidate: ((start: number, end: number) => void) | null = null;
  onSearchHits: ((keys: Float64Array, done: boolean, truncated: boolean) => void) | null = null;

  private readonly blocks = new Map<number, Row[]>(); // insertion order = LRU order
  private readonly pending = new Set<number>();
  private readonly waiting = new Map<number, (value: never) => void>();
  private reqId = 0;
  private searchId = 0;
  private covered = 0;

  constructor(
    readonly job: number,
    readonly initialWrap: boolean,
    private readonly send: (msg: ToWorker) => void,
  ) {}

  /** Cached row (null = empty), or undefined while its block is being fetched. */
  getRow(r: number): Row | undefined {
    if (r < 0 || r >= this.rowCount) return undefined;
    const b = Math.floor(r / BLOCK_ROWS);
    const rows = this.blocks.get(b);
    const i = r - b * BLOCK_ROWS;
    if (rows && i < rows.length) return rows[i];
    this.request(b);
    return undefined;
  }

  /** Fetches blocks covering rows [from, to) ahead of time (e.g. one screen above and below). */
  prefetch(from: number, to: number) {
    const last = Math.min(to, this.rowCount) - 1;
    for (let b = Math.max(0, Math.floor(from / BLOCK_ROWS)); b <= Math.floor(last / BLOCK_ROWS); b++) {
      const rows = this.blocks.get(b);
      if (rows) {
        this.blocks.delete(b); // touch: most recently used
        this.blocks.set(b, rows);
        if (rows.length < BLOCK_ROWS && b * BLOCK_ROWS + rows.length < this.rowCount) this.request(b);
      } else {
        this.request(b);
      }
    }
  }

  private request(b: number) {
    if (this.pending.has(b)) return;
    this.pending.add(b);
    const start = b * BLOCK_ROWS;
    this.send({ type: 'window', job: this.job, reqId: ++this.reqId, start, end: start + BLOCK_ROWS });
  }

  setLayout(widths: number[], wrap: boolean) {
    this.layoutVersion++;
    this.covered = 0;
    this.layoutProgress.set(wrap ? 0 : 1);
    this.send({ type: 'layout', job: this.job, version: this.layoutVersion, widths: widths.slice(), wrap });
  }

  autofit(col: number, from: number): Promise<number> {
    const reqId = ++this.reqId;
    this.send({ type: 'autofit', job: this.job, reqId, col, from });
    return new Promise((resolve) => this.waiting.set(reqId, resolve as (value: never) => void));
  }

  copy(r0: number, r1: number, c0: number, c1: number): Promise<CopyResult> {
    const reqId = ++this.reqId;
    this.send({ type: 'copy', job: this.job, reqId, r0, r1, c0, c1 });
    return new Promise((resolve) => this.waiting.set(reqId, resolve as (value: never) => void));
  }

  search(query: string) {
    this.send({ type: 'search', job: this.job, searchId: ++this.searchId, query });
  }

  /** Routes a worker message of this session. */
  handle(msg: FromWorker) {
    switch (msg.type) {
      case 'columns':
        this.widths = msg.widths;
        this.bump();
        break;
      case 'grow':
        this.rowCount = msg.rowCount;
        this.colCount = msg.colCount;
        this.bump();
        break;
      case 'done':
        this.rowCount = msg.rowCount;
        this.colCount = msg.colCount;
        this.complete = true;
        this.bump();
        break;
      case 'window': {
        const b = Math.floor(msg.start / BLOCK_ROWS);
        this.pending.delete(b);
        this.blocks.delete(b);
        this.blocks.set(b, msg.rows);
        while (this.blocks.size > MAX_BLOCKS) this.blocks.delete(this.blocks.keys().next().value!);
        this.bump();
        break;
      }
      case 'invalidate':
        for (let b = Math.floor(msg.start / BLOCK_ROWS); b <= Math.floor((msg.end - 1) / BLOCK_ROWS); b++) this.blocks.delete(b);
        this.onInvalidate?.(msg.start, msg.end);
        this.bump();
        break;
      case 'heights':
        if (msg.version !== this.layoutVersion) return;
        if (msg.start <= this.covered) this.covered = Math.max(this.covered, msg.start + msg.heights.length);
        this.layoutProgress.set(this.rowCount ? Math.min(1, this.covered / this.rowCount) : 1);
        this.onHeights?.(msg.start, msg.heights);
        break;
      case 'layoutDone':
        if (msg.version === this.layoutVersion) this.layoutProgress.set(1);
        break;
      case 'autofit':
        this.resolve(msg.reqId, msg.width);
        break;
      case 'copy':
        this.resolve(msg.reqId, { text: msg.text, cells: msg.cells, error: msg.error });
        break;
      case 'searchHits':
        if (msg.searchId === this.searchId) this.onSearchHits?.(msg.keys, msg.done, msg.truncated);
        break;
    }
  }

  dispose() {
    this.blocks.clear();
    this.pending.clear();
    this.waiting.clear();
    this.onHeights = this.onInvalidate = this.onSearchHits = null;
  }

  private resolve(reqId: number, value: unknown) {
    const done = this.waiting.get(reqId);
    this.waiting.delete(reqId);
    done?.(value as never);
  }

  private bump() {
    this.dataVersion.update((v) => v + 1);
  }
}
