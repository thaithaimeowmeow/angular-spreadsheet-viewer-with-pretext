/**
 * Exact cell measurement with @chenglou/pretext (UI thread, visible rows only).
 *
 * pretext splits measurement in two:
 *   prepare()  – segments the text and measures it with canvas (the costly part, cached here)
 *   layout()   – pure arithmetic line-breaking for a given width (no DOM, no reflow)
 *
 * Rows are capped at MAX_LINES, so a cell that provably needs ≥ MAX_LINES lines is answered from
 * a cheap per-character lower bound without calling pretext at all.
 */
import { layout, measureNaturalWidth, prepareWithSegments, type PreparedTextWithSegments } from '@chenglou/pretext';
import type { Row } from '../parser/protocol';
import {
  BASE_ROW, BORDER, CELL_FONT, DEFAULT_COL, FONT_SIZE, LINE_HEIGHT, MAX_AUTO_COL, MAX_LINES, PAD_X, PAD_Y,
  availWidth, heightForLines, isNumeric, minLines,
} from './metrics';

export * from './metrics';

// Two generations rotated by total characters: short texts get a large cache, long texts can't
// blow up memory. Dropping a whole generation is O(1) (evicting one-by-one from a Map's head
// degrades badly in V8 because of tombstones).
const GENERATION_CHARS = 2_000_000;
const GENERATION_ENTRIES = 60_000;
let hot = new Map<string, PreparedTextWithSegments>();
let cold = new Map<string, PreparedTextWithSegments>();
let hotChars = 0;

function prepared(text: string): PreparedTextWithSegments {
  let p = hot.get(text);
  if (p) return p;
  p = cold.get(text) ?? prepareWithSegments(text, CELL_FONT, { whiteSpace: 'pre-wrap' });
  if (hotChars >= GENERATION_CHARS || hot.size >= GENERATION_ENTRIES) {
    cold = hot;
    hot = new Map();
    hotChars = 0;
  }
  hot.set(text, p);
  hotChars += text.length;
  return p;
}

/** Drops cached prepared text (call when a file is closed). */
export function clearTextCache() {
  hot = new Map();
  cold = new Map();
  hotChars = 0;
}

/** Wrapped line count of `text` inside a cell `width` px wide (capped at MAX_LINES). */
export function lineCount(text: string, width: number): number {
  const avail = availWidth(width);
  if (avail <= 0) return 1;
  // Fast path: no glyph in our font is wider than 1em, so short single-line text always fits.
  if (text.length * FONT_SIZE <= avail && text.indexOf('\n') === -1 && text.indexOf('\t') === -1) return 1;
  if (minLines(text, avail) >= MAX_LINES) return MAX_LINES;
  return Math.max(1, layout(prepared(text), avail, LINE_HEIGHT).lineCount);
}

export function rowHeight(row: Row, colWidths: ArrayLike<number>): number {
  if (!row) return BASE_ROW;
  let lines = 1;
  for (let c = 0; c < row.length; c++) {
    const text = row[c];
    if (!text || isNumeric(text)) continue;
    const n = lineCount(text, colWidths[c] ?? DEFAULT_COL);
    if (n > lines) {
      lines = n;
      if (lines >= MAX_LINES) break;
    }
  }
  return heightForLines(lines);
}

/** Upper limit for "fit row to all text", so one huge cell can't make a 100k px row. */
const MAX_FIT_LINES = 500;

/** Height that shows every line of every cell in the row, ignoring the MAX_LINES cap. */
export function fullRowHeight(row: Row, colWidths: ArrayLike<number>): number {
  if (!row) return BASE_ROW;
  let lines = 1;
  for (let c = 0; c < row.length; c++) {
    const text = row[c];
    if (!text || isNumeric(text)) continue;
    const avail = availWidth(colWidths[c] ?? DEFAULT_COL);
    if (avail <= 0) continue;
    if (text.length * FONT_SIZE <= avail && text.indexOf('\n') === -1 && text.indexOf('\t') === -1) continue;
    lines = Math.max(lines, layout(prepared(text), avail, LINE_HEIGHT).lineCount);
  }
  return Math.min(lines, MAX_FIT_LINES) * LINE_HEIGHT + PAD_Y * 2 + BORDER;
}

/** Width needed to show `text` without wrapping (widest hard line). */
export function naturalWidth(text: string): number {
  if (text.length * 3 > MAX_AUTO_COL) return MAX_AUTO_COL; // can't fit anyway, skip measuring
  return Math.ceil(measureNaturalWidth(prepared(text))) + PAD_X * 2 + BORDER;
}
