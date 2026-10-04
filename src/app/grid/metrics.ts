/**
 * Cell metrics shared by the UI and the sheet worker. No pretext and no DOM here, so the worker
 * bundle stays small: the worker only ESTIMATES heights; the UI measures visible rows exactly
 * with pretext (text-metrics.ts).
 */

// Must match .cell in sheet-grid.css exactly, or measured heights will drift from the rendered DOM.
export const FONT_SIZE = 13;
export const CELL_FONT = `${FONT_SIZE}px Arial, Helvetica, sans-serif`;
export const LINE_HEIGHT = 18;
export const PAD_X = 6;
export const PAD_Y = 3;
export const BORDER = 1;
export const MAX_LINES = 8;
export const BASE_ROW = LINE_HEIGHT + PAD_Y * 2 + BORDER;
export const DEFAULT_COL = 100;
export const MIN_COL = 32;
export const MAX_AUTO_COL = 320;

const NUMERIC = /^[-+]?[\d,]*\.?\d+(?:[eE][-+]?\d+)?%?$/;

export function isNumeric(text: string): boolean {
  return text.length < 32 && NUMERIC.test(text);
}

export const availWidth = (width: number) => width - PAD_X * 2 - BORDER;

export function heightForLines(lines: number): number {
  return Math.min(Math.max(1, lines), MAX_LINES) * LINE_HEIGHT + PAD_Y * 2 + BORDER;
}

// ------------------------------------------------------------------ per-character widths

const TABLE_SIZE = 0x2100;
const TABLE_RANGES = [[0x20, 0x24f], [0x300, 0x4ff], [0x1e00, 0x1eff], [0x2000, 0x20cf]];

interface CharTables {
  /** Isolated advance per character, NaN when not measured. */
  upper: Float32Array;
  /** Conservative advance for lower bounds: 0 for spaces (they hang at wraps), marks, unknowns. */
  lower: Float32Array;
  /** Widest measured glyph; below this column width a single glyph may overflow a line. */
  maxGlyph: number;
}

let tables: CharTables | null = null;

function contributesNothing(cp: number): boolean {
  return cp === 0x20 || cp === 0xad
    || (cp >= 0x300 && cp <= 0x36f) // combining marks: zero advance once shaped
    || (cp >= 0x2000 && cp <= 0x200f) || (cp >= 0x2028 && cp <= 0x202f) || (cp >= 0x205f && cp <= 0x206f);
}

function charTables(): CharTables {
  if (tables) return tables;
  const ctx = (typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(1, 1).getContext('2d')
    : typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d')
    : null) as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
  if (ctx) ctx.font = CELL_FONT;
  // Without any canvas (old engines, Node tests): a rough average advance. Only estimates use it.
  const measure = (ch: string) => (ctx ? ctx.measureText(ch).width : FONT_SIZE * 0.6);
  const upper = new Float32Array(TABLE_SIZE).fill(NaN);
  const lower = new Float32Array(TABLE_SIZE);
  let maxGlyph = 0;
  for (const [from, to] of TABLE_RANGES) {
    for (let cp = from; cp <= to; cp++) {
      const w = measure(String.fromCharCode(cp));
      upper[cp] = w;
      if (!contributesNothing(cp)) {
        // 15% margin: kerning makes shaped text slightly narrower than the sum of its glyphs.
        lower[cp] = w * 0.85;
        if (w > maxGlyph) maxGlyph = w;
      }
    }
  }
  return (tables = { upper, lower, maxGlyph });
}

/** Approximate single-line width: sum of isolated glyph advances; unknown characters = 1em. */
export function estimateWidth(text: string): number {
  const { upper } = charTables();
  let w = 0;
  for (let i = 0; i < text.length; i++) {
    const cp = text.charCodeAt(i);
    const cw = cp < TABLE_SIZE ? upper[cp] : NaN;
    w += cw === cw ? cw : FONT_SIZE; // cw === cw is false only for NaN
  }
  return w;
}

/**
 * Lower bound of the wrapped line count. Every line holds at most `avail` px of non-hanging
 * content, so lines ≥ (sum of lower glyph widths) / avail; hard breaks add lines.
 * Only valid when any single glyph fits a line (avail ≥ widest glyph).
 */
export function minLines(text: string, avail: number): number {
  const { lower, maxGlyph } = charTables();
  let w = 0;
  let breaks = 0;
  for (let i = 0; i < text.length; i++) {
    const cp = text.charCodeAt(i);
    if (cp === 10) breaks++;
    else if (cp < TABLE_SIZE) w += lower[cp];
  }
  if (avail < maxGlyph) return breaks + 1;
  return Math.max(breaks + 1, Math.ceil(w / (avail + 1)));
}

/**
 * Estimated wrapped line count (no pretext). Per hard line: width / available width, plus 8%
 * for space lost at word breaks. [Estimate] Used only for rows that are not on screen.
 */
export function estimateLines(text: string, width: number): number {
  const avail = availWidth(width);
  if (avail <= 0 || isNumeric(text)) return 1;
  if (text.length * FONT_SIZE <= avail && text.indexOf('\n') === -1) return 1;
  let lines = 0;
  let start = 0;
  for (;;) {
    const nl = text.indexOf('\n', start);
    const part = nl === -1 ? text.slice(start) : text.slice(start, nl);
    lines += part ? Math.max(1, Math.ceil((estimateWidth(part) * 1.08) / avail)) : 1;
    if (lines >= MAX_LINES || nl === -1) break;
    start = nl + 1;
  }
  return Math.min(lines, MAX_LINES);
}

export function estimateRowHeight(row: (string | undefined)[] | null, widths: ArrayLike<number>): number {
  if (!row) return BASE_ROW;
  let lines = 1;
  for (let c = 0; c < row.length; c++) {
    const text = row[c];
    if (!text) continue;
    const n = estimateLines(text, widths[c] ?? DEFAULT_COL);
    if (n > lines) {
      lines = n;
      if (lines >= MAX_LINES) break;
    }
  }
  return heightForLines(lines);
}

/** Column width that shows `text` on one line (estimate), clamped to [48, max]. */
export function estimateColWidth(text: string, max = MAX_AUTO_COL): number {
  if (text.length * 3 > max) return max;
  let widest = 0;
  for (const line of text.split('\n')) widest = Math.max(widest, estimateWidth(line));
  return Math.min(max, Math.max(48, Math.ceil(widest) + PAD_X * 2 + BORDER + 2));
}
