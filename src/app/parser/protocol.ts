/**
 * Contract between the UI thread and the sheet worker.
 *
 * The worker parses the file and KEEPS the current sheet's rows. The UI holds only row heights
 * plus a small cache of row blocks it fetches with `window` requests for what is on screen.
 * Off-screen heights are worker ESTIMATES (no pretext); the UI measures visible rows exactly
 * with pretext and those values win.
 */

/** A row is a (possibly holey) array of display strings, or null when empty. */
export type Row = string[] | null;

export type LoadPhase = 'reading' | 'strings' | 'rows';

/** Rows per block for `window` requests and the UI row cache. */
export const BLOCK_ROWS = 256;

export type ToWorker =
  // ---- control: cancel any in-flight open/sheet work and start a new job
  | { type: 'open'; job: number; file: File }
  /** Load sheet `index`; `job` names this sheet session for all later messages. */
  | { type: 'sheet'; job: number; index: number; wrap: boolean }
  // ---- queries: never cancel anything; ignored unless `job` is the current sheet session
  /** Rows [start, end). */
  | { type: 'window'; job: number; reqId: number; start: number; end: number }
  /** New column widths / wrap mode; the worker re-estimates every height for `version`. */
  | { type: 'layout'; job: number; version: number; widths: number[]; wrap: boolean }
  /** Width that fits column `col` (scans the first rows and rows from `from`). */
  | { type: 'autofit'; job: number; reqId: number; col: number; from: number }
  /** Start a search (empty query stops it). Hits stream back as `searchHits`. */
  | { type: 'search'; job: number; searchId: number; query: string }
  /** Tab-separated text of the range, for the clipboard. */
  | { type: 'copy'; job: number; reqId: number; r0: number; r1: number; c0: number; c1: number };

export type FromWorker =
  | { type: 'meta'; job: number; sheets: string[]; format: string }
  | { type: 'progress'; job: number; phase: LoadPhase; loaded: number; total: number }
  /** Auto-sized column widths, sent once per session before the first `grow`. Layout version 0. */
  | { type: 'columns'; job: number; widths: number[] }
  /** More rows are available. Counts only, no data. */
  | { type: 'grow'; job: number; rowCount: number; colCount: number }
  /** Estimated heights for rows [start, start + heights.length) at layout `version`. Transferred. */
  | { type: 'heights'; job: number; version: number; start: number; heights: Uint16Array }
  /** Every row has an estimate for `version` (sheet fully parsed). */
  | { type: 'layoutDone'; job: number; version: number }
  /** Rows [start, end) changed after they may have been fetched (out-of-order data). */
  | { type: 'invalidate'; job: number; start: number; end: number }
  | { type: 'window'; job: number; reqId: number; start: number; rows: Row[] }
  | { type: 'autofit'; job: number; reqId: number; width: number }
  /** New match keys (row * MAX_COLS + col, ascending) appended to the search `searchId`. Transferred. */
  | { type: 'searchHits'; job: number; searchId: number; keys: Float64Array; done: boolean; truncated: boolean }
  | { type: 'copy'; job: number; reqId: number; text: string | null; cells: number; error?: string }
  | { type: 'done'; job: number; rowCount: number; colCount: number }
  | { type: 'error'; job: number; message: string };
