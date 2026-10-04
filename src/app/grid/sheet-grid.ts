import {
  ChangeDetectionStrategy, Component, DestroyRef, ElementRef, afterNextRender, computed, effect,
  inject, input, output, signal, untracked, viewChild,
} from '@angular/core';
import { BLOCK_ROWS } from '../parser/protocol';
import { RowHeights } from './row-heights';
import { MAX_COLS, SheetSearch, type TextPart, highlightParts } from './sheet-search';
import type { SheetSource } from './sheet-source';
import { BASE_ROW, DEFAULT_COL, LINE_HEIGHT, MIN_COL, fullRowHeight, isNumeric, rowHeight } from './text-metrics';

export interface CellSelection {
  ref: string;
  row: number;
  col: number;
  value: string;
}

interface ViewCell {
  x: number; y: number; w: number; h: number; text: string; num: boolean;
  parts: TextPart[] | null; current: boolean;
}
interface ViewHead { key: number; pos: number; size: number; label: string }

const HEADER_H = 24;
const MIN_ROW = 12;
const OVERSCAN_PX = 120;
/** While dragging a column edge, send new widths to the worker at most this often. */
const LAYOUT_THROTTLE_MS = 120;
/** Browsers cap element height (~17.8M px in Firefox). Beyond this we scale the scrollbar. */
const MAX_SCROLL_PX = 10_000_000;

export function colName(c: number): string {
  let s = '';
  for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
}

/**
 * Virtualized sheet view. Rows live in the worker; this component holds row heights (Fenwick
 * tree) and reads row text through SheetSource's block cache. Visible rows are measured exactly
 * with pretext; all other heights are worker estimates.
 */
@Component({
  selector: 'app-sheet-grid',
  templateUrl: './sheet-grid.html',
  styleUrl: './sheet-grid.css',
  host: { '[class.nowrap]': '!wrap()' },
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class SheetGrid {
  readonly source = input.required<SheetSource>();
  readonly wrap = input(true);
  readonly selectionChange = output<CellSelection | null>();
  /** Short status messages (copy result, limits) for the host to show. */
  readonly notice = output<string>();

  /** Last time (ms) to measure the visible rows exactly; shown in the UI. */
  readonly lastLayoutMs = signal(0);
  readonly searchCount = signal(0);
  readonly searchIndex = signal(-1);
  readonly searchDone = signal(true);
  readonly searchTruncated = signal(false);

  private readonly scrollerRef = viewChild<ElementRef<HTMLDivElement>>('scroller');
  private readonly spacerRef = viewChild<ElementRef<HTMLDivElement>>('spacer');

  private readonly heights = new RowHeights(BASE_ROW);
  private colWidths: number[] = [];
  private colOffsets = new Float64Array(1);
  private widthsReady = false;
  private wrapOn = true;
  private layoutTimer = 0;

  // Scroll state. `virtualTop` is authoritative; the native scrollTop may be scaled.
  private virtualTop = 0;
  private scrollLeft = 0;
  private lastScrollTop = 0;
  private viewW = 0;
  private viewH = 0;
  protected gutterW = 48;

  /** Selection: `anchor` is the active cell, `focus` the opposite corner of the range. */
  private anchor: { r: number; c: number } | null = null;
  private focus: { r: number; c: number } | null = null;
  private dragMode: 'cells' | 'rows' | 'cols' | null = null;
  /** The active cell's row was not cached when selected: re-emit its value when it arrives. */
  private selectionStale = false;
  private resizing: { c: number; startX: number; startW: number } | null = null;
  private resizingRow: { r: number; startY: number; startH: number } | null = null;
  /** Heights the user set by dragging or double-clicking a row edge; they win over measurement. */
  private rowOverrides = new Map<number, number>();
  /** Cell whose full text is shown in the viewer panel. */
  private viewerCell: { r: number; c: number } | null = null;

  /** Match keys (from the worker) and the highlight regex. */
  private readonly search = new SheetSearch();
  private searchQuery = '';
  private currentMatch = -1;
  /** Enter pressed before any match arrived: jump as soon as the first one does. */
  private pendingFind: 0 | 1 | -1 = 0;

  private rafId = 0;
  private syncRafId = 0;

  protected readonly frame = signal(0);
  protected readonly view = computed(() => {
    this.frame();
    return this.buildView();
  });

  constructor() {
    effect((onCleanup) => {
      const source = this.source();
      untracked(() => this.attach(source));
      onCleanup(() => source.dispose());
    });
    effect(() => {
      this.source().dataVersion();
      untracked(() => this.syncData());
    });
    effect(() => {
      const wrap = this.wrap();
      untracked(() => this.applyWrap(wrap));
    });

    const destroyRef = inject(DestroyRef);
    afterNextRender(() => {
      const el = this.scrollerRef()!.nativeElement;
      const ro = new ResizeObserver(() => {
        this.viewW = el.clientWidth;
        this.viewH = el.clientHeight;
        this.render();
      });
      ro.observe(el);
      const wheel = (e: WheelEvent) => this.onWheel(e);
      el.addEventListener('wheel', wheel, { passive: false });
      destroyRef.onDestroy(() => {
        ro.disconnect();
        el.removeEventListener('wheel', wheel);
        cancelAnimationFrame(this.rafId);
        cancelAnimationFrame(this.syncRafId);
        clearTimeout(this.layoutTimer);
      });
    });
  }

  // ------------------------------------------------------------ data

  /** New sheet session: reset view state and subscribe to the worker's pushes. */
  private attach(source: SheetSource) {
    // A width update queued for the previous sheet must not reach this one.
    clearTimeout(this.layoutTimer);
    this.layoutTimer = 0;
    this.resizing = this.resizingRow = null;
    this.heights.reset();
    this.rowOverrides.clear();
    this.viewerCell = null;
    this.colWidths = [];
    this.widthsReady = false;
    this.virtualTop = 0;
    this.scrollLeft = 0;
    this.anchor = this.focus = null;
    this.dragMode = null;
    this.selectionChange.emit(null);
    const el = this.scrollerRef()?.nativeElement;
    if (el) el.scrollTop = el.scrollLeft = 0;
    this.lastScrollTop = 0;

    source.onHeights = (start, values) => {
      if (!this.wrapOn) return;
      this.heights.resize(source.rowCount);
      this.keepAnchor(() => this.heights.applyEstimates(start, values));
      // Rows on screen are measured exactly, so only the scrollbar changes.
      this.syncScrollSoon();
    };
    source.onInvalidate = (start, end) => {
      for (let r = start; r < end && r < this.heights.count; r++) this.heights.unmark(r);
      this.render();
    };
    source.onSearchHits = (keys, done, truncated) => this.onSearchHits(keys, done, truncated);

    this.rebuildColumns();
    this.setSearch(this.searchQuery);
    this.render();
  }

  private syncData() {
    const source = this.source();
    if (!this.widthsReady && source.widths) {
      this.colWidths = source.widths.slice();
      this.widthsReady = true;
      // The session started with the wrap mode of the moment; send ours if it changed since.
      if (this.wrapOn !== source.initialWrap) source.setLayout(this.colWidths, this.wrapOn);
    }
    this.rebuildColumns();
    const n = source.rowCount;
    // New rows start at the base height, which is also what wrap-off uses.
    this.heights.resize(n);
    this.gutterW = Math.max(40, String(n).length * 8 + 18);
    if (this.selectionStale && this.anchor && source.getRow(this.anchor.r) !== undefined) {
      this.emitSelection();
    }
    this.render();
  }

  private row(r: number) {
    return this.source().getRow(r);
  }

  private rebuildColumns() {
    const count = Math.max(this.source().colCount, this.widthsReady ? this.colWidths.length : 0);
    while (this.colWidths.length < count) this.colWidths.push(DEFAULT_COL);
    const offsets = new Float64Array(count + 1);
    for (let c = 0; c < count; c++) offsets[c + 1] = offsets[c] + this.colWidths[c];
    this.colOffsets = offsets;
  }

  private applyWrap(wrap: boolean) {
    const changed = this.wrapOn !== wrap;
    this.wrapOn = wrap;
    this.keepAnchor(() => {
      if (wrap) this.heights.invalidate();
      else this.fillBase();
    });
    if (changed && this.widthsReady) this.source().setLayout(this.colWidths, wrap);
    this.render();
  }

  /** Sends the current widths to the worker (throttled during column drags). */
  private pushLayout(immediate = false) {
    if (!this.widthsReady) return;
    clearTimeout(this.layoutTimer);
    const send = () => {
      this.layoutTimer = 0;
      if (this.widthsReady) this.source().setLayout(this.colWidths, this.wrapOn);
    };
    if (immediate) send();
    else this.layoutTimer = setTimeout(send, LAYOUT_THROTTLE_MS);
  }

  // ------------------------------------------------------------ exact layout (visible rows)

  /** Wrap off: every row at the base height, except rows the user resized. */
  private fillBase() {
    this.heights.fill(BASE_ROW);
    for (const [r, h] of this.rowOverrides) if (r < this.heights.count) this.heights.set(r, h);
  }

  /** Measures cached visible rows exactly with pretext; uncached rows wait for their block. */
  private measureVisible() {
    if (!this.wrapOn) return;
    const t0 = performance.now();
    const h = this.heights;
    const end = this.virtualTop + this.viewH - HEADER_H + OVERSCAN_PX;
    let r = h.rowAt(this.virtualTop);
    let y = h.offsetOf(r);
    let measured = 0;
    for (; r < h.count && y < end; r++) {
      if (!h.isMeasured(r)) {
        const override = this.rowOverrides.get(r);
        const data = override === undefined ? this.row(r) : null;
        if (override !== undefined || data !== undefined) {
          h.set(r, override ?? rowHeight(data!, this.colWidths));
          h.markMeasured(r);
          measured++;
        }
      }
      y += h.height(r);
    }
    if (measured) this.lastLayoutMs.set(+(performance.now() - t0).toFixed(2));
  }

  /** Runs a height mutation while keeping the top visible row pinned on screen. */
  private keepAnchor(mutate: () => void) {
    const anchor = this.heights.rowAt(this.virtualTop);
    const before = this.heights.offsetOf(anchor);
    const intra = Math.min(this.virtualTop - before, this.heights.count ? this.heights.height(anchor) : 0);
    mutate();
    const after = this.heights.offsetOf(anchor);
    if (after !== before) this.virtualTop = after + intra;
  }

  // ------------------------------------------------------------ scrolling

  private contentHeight() {
    return HEADER_H + this.heights.total();
  }

  private scrollRatio() {
    const total = this.contentHeight();
    if (total <= MAX_SCROLL_PX) return 1;
    return (total - this.viewH) / (MAX_SCROLL_PX - this.viewH);
  }

  /** Updates spacer size and scroll position on the next frame, without rebuilding the cells. */
  private syncScrollSoon() {
    if (this.rafId || this.syncRafId) return;
    this.syncRafId = requestAnimationFrame(() => {
      this.syncRafId = 0;
      this.syncScroll();
    });
  }

  private render() {
    if (this.rafId) return;
    this.rafId = requestAnimationFrame(() => {
      this.rafId = 0;
      this.measureVisible();
      this.syncScroll();
      this.frame.update((v) => v + 1);
    });
  }

  /** Writes spacer size and native scroll position from the virtual state. */
  private syncScroll() {
    const el = this.scrollerRef()?.nativeElement;
    const spacer = this.spacerRef()?.nativeElement;
    if (!el || !spacer) return;
    const total = this.contentHeight();
    const maxTop = Math.max(0, total - this.viewH);
    this.virtualTop = Math.min(Math.max(0, this.virtualTop), maxTop);
    spacer.style.height = `${Math.min(total, MAX_SCROLL_PX)}px`;
    spacer.style.width = `${this.gutterW + this.colOffsets[this.colOffsets.length - 1]}px`;
    const target = this.virtualTop / this.scrollRatio();
    if (Math.abs(el.scrollTop - target) >= 1) {
      el.scrollTop = target;
      this.lastScrollTop = el.scrollTop;
    }
    if (el.scrollLeft !== this.scrollLeft) {
      el.scrollLeft = this.scrollLeft;
      this.scrollLeft = el.scrollLeft;
    }
  }

  protected onScroll() {
    const el = this.scrollerRef()!.nativeElement;
    const st = el.scrollTop;
    // An echo of our own scrollTop write (e.g. anchoring after estimates): nothing on screen
    // moved, so skip the redraw.
    if (Math.abs(st - this.lastScrollTop) < 1 && el.scrollLeft === this.scrollLeft) return;
    // Ignore echoes of our own scrollTop writes so scaled scrolling stays precise.
    if (Math.abs(st - this.lastScrollTop) >= 1) this.virtualTop = st * this.scrollRatio();
    this.lastScrollTop = st;
    this.scrollLeft = el.scrollLeft;
    this.render();
  }

  /** In scaled mode the native wheel step would be multiplied; scroll virtually instead. */
  private onWheel(e: WheelEvent) {
    if (this.scrollRatio() === 1 || e.ctrlKey) return;
    e.preventDefault();
    const unit = e.deltaMode === 1 ? LINE_HEIGHT : e.deltaMode === 2 ? this.viewH : 1;
    const dy = e.shiftKey && !e.deltaX ? 0 : e.deltaY * unit;
    const dx = e.shiftKey && !e.deltaX ? e.deltaY * unit : e.deltaX * unit;
    this.virtualTop += dy;
    this.scrollLeft = Math.max(0, this.scrollLeft + dx);
    this.render();
  }

  // ------------------------------------------------------------ view model

  private colAt(x: number): number {
    const o = this.colOffsets;
    let lo = 0;
    let hi = o.length - 2;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (o[mid] <= x) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  private buildView() {
    const h = this.heights;
    const bodyW = this.viewW - this.gutterW;
    const bodyH = this.viewH - HEADER_H;

    const colHeads: ViewHead[] = [];
    const lastCol = this.colOffsets.length - 1;
    for (let c = this.colAt(this.scrollLeft); c < lastCol; c++) {
      const x = this.colOffsets[c] - this.scrollLeft;
      if (x >= bodyW) break;
      colHeads.push({ key: c, pos: x, size: this.colWidths[c], label: colName(c) });
    }

    const rowHeads: ViewHead[] = [];
    if (h.count) {
      let r = h.rowAt(this.virtualTop);
      let y = h.offsetOf(r) - this.virtualTop;
      for (; r < h.count && y < bodyH; r++) {
        const rh = h.height(r);
        rowHeads.push({ key: r, pos: y, size: rh, label: String(r + 1) });
        y += rh;
      }
      // Keep a block above and below cached so short scrolls never show empty rows.
      if (rowHeads.length) this.source().prefetch(rowHeads[0].key - BLOCK_ROWS, rowHeads[rowHeads.length - 1].key + BLOCK_ROWS);
    }

    // Cells are tracked by position ($index) in the template, so scrolling reuses the same DOM
    // elements and only updates their text and position.
    const cells: ViewCell[] = [];
    const hl = this.search.global;
    const currentKey = this.currentMatch >= 0 && this.currentMatch < this.search.count ? this.search.key(this.currentMatch) : -1;
    for (const rh of rowHeads) {
      const row = this.row(rh.key);
      for (const ch of colHeads) {
        const text = row?.[ch.key] ?? '';
        cells.push({
          x: ch.pos, y: rh.pos, w: ch.size, h: rh.size,
          text, num: text !== '' && isNumeric(text),
          parts: hl && text ? highlightParts(text, hl) : null,
          current: rh.key * MAX_COLS + ch.key === currentKey,
        });
      }
    }

    // Selection range, clipped to the visible area so a whole-column selection stays cheap.
    let selection: { x: number; y: number; w: number; h: number } | null = null;
    let active: { x: number; y: number; w: number; h: number } | null = null;
    const rg = this.range();
    if (rg) {
      const top = Math.max(h.offsetOf(rg.r0) - this.virtualTop, -4);
      const bottom = Math.min(h.offsetOf(rg.r1 + 1) - this.virtualTop, bodyH + 4);
      const left = Math.max(this.colOffsets[rg.c0] - this.scrollLeft, -4);
      const right = Math.min(this.colOffsets[rg.c1 + 1] - this.scrollLeft, bodyW + 4);
      if (bottom > top && right > left && (rg.r0 !== rg.r1 || rg.c0 !== rg.c1)) {
        selection = { x: left, y: top, w: right - left, h: bottom - top };
      }
      const a = this.anchor!;
      active = {
        x: this.colOffsets[a.c] - this.scrollLeft,
        y: h.offsetOf(a.r) - this.virtualTop,
        w: this.colWidths[a.c],
        h: h.height(a.r),
      };
    }

    let viewer: { x: number; y: number; w: number; maxH: number; ref: string; text: string; parts: TextPart[] | null } | null = null;
    const vc = this.viewerCell;
    if (vc && vc.r < h.count && vc.c < this.colWidths.length) {
      const w = Math.min(Math.max(this.colWidths[vc.c], 360), this.viewW - 16);
      const cellX = this.gutterW + this.colOffsets[vc.c] - this.scrollLeft;
      const cellY = HEADER_H + h.offsetOf(vc.r) - this.virtualTop;
      const x = Math.min(Math.max(8, cellX), this.viewW - w - 8);
      const y = Math.min(Math.max(HEADER_H + 4, cellY), Math.max(HEADER_H + 4, this.viewH * 0.45));
      const text = this.cellText(vc.r, vc.c);
      viewer = {
        x, y, w, maxH: this.viewH - y - 8, ref: `${colName(vc.c)}${vc.r + 1}`, text,
        parts: hl ? highlightParts(text, hl) : null,
      };
    }

    return {
      cells, colHeads, rowHeads, selection, active, viewer,
      selR0: rg?.r0 ?? -1, selR1: rg?.r1 ?? -2, selC0: rg?.c0 ?? -1, selC1: rg?.c1 ?? -2,
      viewW: this.viewW, viewH: this.viewH, gutterW: this.gutterW,
    };
  }

  // ------------------------------------------------------------ selection

  private range(): { r0: number; r1: number; c0: number; c1: number } | null {
    const a = this.anchor;
    const f = this.focus;
    if (!a || !f || !this.heights.count || !this.colWidths.length) return null;
    const maxR = this.heights.count - 1;
    const maxC = this.colWidths.length - 1;
    return {
      r0: Math.min(a.r, f.r, maxR), r1: Math.min(Math.max(a.r, f.r), maxR),
      c0: Math.min(a.c, f.c, maxC), c1: Math.min(Math.max(a.c, f.c), maxC),
    };
  }

  private lastRow() {
    return Math.max(0, this.heights.count - 1);
  }

  private lastCol() {
    return Math.max(0, this.colWidths.length - 1);
  }

  private setSelection(anchor: { r: number; c: number }, focus = anchor) {
    const v = this.viewerCell;
    if (v && (v.r !== anchor.r || v.c !== anchor.c)) this.viewerCell = null;
    this.anchor = anchor;
    this.focus = focus;
    this.emitSelection();
    this.render();
  }

  private emitSelection() {
    const a = this.anchor;
    const rg = this.range();
    if (!a || !rg) return;
    const single = rg.r0 === rg.r1 && rg.c0 === rg.c1;
    this.selectionStale = this.row(a.r) === undefined;
    this.selectionChange.emit({
      ref: single ? `${colName(a.c)}${a.r + 1}` : `${colName(rg.c0)}${rg.r0 + 1}:${colName(rg.c1)}${rg.r1 + 1}`,
      row: a.r,
      col: a.c,
      value: this.cellText(a.r, a.c),
    });
  }

  selectAll() {
    if (!this.heights.count || !this.colWidths.length) return;
    this.setSelection({ r: 0, c: 0 }, { r: this.lastRow(), c: this.lastCol() });
  }

  /** Copies the selection as tab-separated text (pastes into Excel as cells). */
  copySelection() {
    const rg = this.range();
    if (!rg) return;
    // The worker builds the text; ClipboardItem accepts it as a promise so the copy still counts
    // as part of the key press. Fallback: await it, then writeText.
    const result = this.source().copy(rg.r0, rg.r1, rg.c0, rg.c1);
    const report = (res: { cells: number; error?: string }) =>
      this.notice.emit(res.error ?? (res.cells === 1 ? 'Copied 1 cell' : `Copied ${res.cells.toLocaleString()} cells`));
    const failed = () => this.notice.emit('Copy failed: the browser denied clipboard access.');
    const blob = result.then((res) => {
      if (res.text === null) throw new Error(res.error);
      return new Blob([res.text], { type: 'text/plain' });
    });
    blob.catch(() => {}); // handled below; avoids an unhandled-rejection warning
    try {
      navigator.clipboard.write([new ClipboardItem({ 'text/plain': blob })])
        .then(() => result.then(report))
        .catch(() => result.then((res) => {
          if (res.text === null) return report(res);
          return navigator.clipboard.writeText(res.text).then(() => report(res), failed);
        }));
    } catch {
      result.then((res) => {
        if (res.text === null) return report(res);
        return navigator.clipboard.writeText(res.text).then(() => report(res), failed);
      });
    }
  }

  // ------------------------------------------------------------ search

  /** Starts a new search over the whole sheet in the worker (empty query clears it). */
  setSearch(query: string) {
    this.searchQuery = query;
    const on = this.search.setQuery(query);
    this.currentMatch = -1;
    this.pendingFind = 0;
    this.searchIndex.set(-1);
    this.searchCount.set(0);
    this.searchTruncated.set(false);
    this.searchDone.set(!on);
    this.source().search(on ? query : '');
    this.render();
  }

  private onSearchHits(keys: Float64Array, done: boolean, truncated: boolean) {
    if (!this.search.active) return;
    this.search.append(keys);
    this.searchCount.set(this.search.count);
    this.searchTruncated.set(truncated);
    if (done) this.searchDone.set(true);
    if (this.pendingFind && this.search.count) {
      const dir = this.pendingFind;
      this.pendingFind = 0;
      this.findNext(dir);
    } else if (done) {
      this.pendingFind = 0;
    }
  }

  /** Selects the next (1) or previous (-1) match after the active cell, wrapping around. */
  findNext(dir: 1 | -1) {
    const n = this.search.count;
    if (!n) {
      if (this.search.active && !this.searchDone()) this.pendingFind = dir;
      return;
    }
    const a = this.anchor;
    let i: number;
    if (this.currentMatch >= 0 && this.currentMatch < n && a && this.search.key(this.currentMatch) === a.r * MAX_COLS + a.c) {
      i = (this.currentMatch + dir + n) % n;
    } else {
      i = this.search.neighbor(a ? a.r * MAX_COLS + a.c : -1, dir);
    }
    this.currentMatch = i;
    this.searchIndex.set(i);
    const k = this.search.key(i);
    const r = Math.floor(k / MAX_COLS);
    const c = k % MAX_COLS;
    this.setSelection({ r, c });
    this.scrollIntoView(r, c);
  }

  // ------------------------------------------------------------ pointer

  private point(e: MouseEvent) {
    const rect = this.scrollerRef()!.nativeElement.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }

  /** Row / column under a point, clamped to the data (for drags that leave the grid). */
  private clampedCell(x: number, y: number) {
    const offset = Math.min(Math.max(0, this.virtualTop + y - HEADER_H), Math.max(0, this.heights.total() - 1));
    const colX = Math.min(Math.max(0, this.scrollLeft + x - this.gutterW), Math.max(0, this.colOffsets[this.colOffsets.length - 1] - 1));
    return { r: this.heights.rowAt(offset), c: this.colAt(colX) };
  }

  /** What a pointer event hits: a cell, a row number, a column letter, the corner, or nothing. */
  private hit(e: MouseEvent): { zone: 'cell' | 'row' | 'col' | 'corner'; r: number; c: number } | null {
    const { x, y } = this.point(e);
    if (x >= this.viewW || y >= this.viewH || !this.heights.count || !this.colWidths.length) return null;
    const inRows = this.virtualTop + y - HEADER_H < this.heights.total();
    const inCols = this.scrollLeft + x - this.gutterW < this.colOffsets[this.colOffsets.length - 1];
    const { r, c } = this.clampedCell(x, y);
    if (x < this.gutterW && y < HEADER_H) return { zone: 'corner', r, c };
    if (x < this.gutterW) return inRows ? { zone: 'row', r, c } : null;
    if (y < HEADER_H) return inCols ? { zone: 'col', r, c } : null;
    return inRows && inCols ? { zone: 'cell', r, c } : null;
  }

  protected onPointerDown(e: PointerEvent) {
    if (e.button !== 0) return;
    this.closeViewer();
    const target = this.hit(e);
    if (!target) return;
    const el = this.scrollerRef()!.nativeElement;
    el.focus({ preventScroll: true });
    const a = e.shiftKey && this.anchor ? this.anchor : null;
    switch (target.zone) {
      case 'corner':
        this.selectAll();
        return;
      case 'row':
        this.dragMode = 'rows';
        this.setSelection({ r: a ? a.r : target.r, c: 0 }, { r: target.r, c: this.lastCol() });
        break;
      case 'col':
        this.dragMode = 'cols';
        this.setSelection({ r: 0, c: a ? a.c : target.c }, { r: this.lastRow(), c: target.c });
        break;
      case 'cell':
        this.dragMode = 'cells';
        this.setSelection(a ?? { r: target.r, c: target.c }, { r: target.r, c: target.c });
        break;
    }
    el.setPointerCapture(e.pointerId);
  }

  protected onPointerMove(e: PointerEvent) {
    if (!this.dragMode || !this.anchor) return;
    const { x, y } = this.point(e);
    // Auto-scroll while dragging past an edge.
    const step = 24;
    if (y > this.viewH - 8) this.virtualTop += step;
    else if (y < HEADER_H && this.dragMode !== 'cols') this.virtualTop = Math.max(0, this.virtualTop - step);
    if (x > this.viewW - 8) this.scrollLeft += step;
    else if (x < this.gutterW && this.dragMode !== 'rows') this.scrollLeft = Math.max(0, this.scrollLeft - step);
    const p = this.clampedCell(x, y);
    const focus = this.dragMode === 'rows' ? { r: p.r, c: this.lastCol() }
      : this.dragMode === 'cols' ? { r: this.lastRow(), c: p.c }
      : p;
    if (this.focus && focus.r === this.focus.r && focus.c === this.focus.c) {
      this.render();
      return;
    }
    this.setSelection(this.anchor, focus);
  }

  protected onPointerUp(e: PointerEvent) {
    this.dragMode = null;
    const el = this.scrollerRef()?.nativeElement;
    if (el?.hasPointerCapture(e.pointerId)) el.releasePointerCapture(e.pointerId);
  }

  /** Double-click a cell: show its full text in a scrollable panel. */
  protected onDoubleClick(e: MouseEvent) {
    const target = this.hit(e);
    if (target?.zone === 'cell') this.openViewer(target.r, target.c);
  }

  private openViewer(r: number, c: number) {
    if (!this.cellText(r, c)) return;
    this.viewerCell = { r, c };
    this.render();
  }

  protected closeViewer() {
    if (!this.viewerCell) return;
    this.viewerCell = null;
    this.render();
  }

  // ------------------------------------------------------------ keyboard

  protected onKeyDown(e: KeyboardEvent) {
    if (e.key === 'Escape') {
      this.closeViewer();
      return;
    }
    if (!this.colWidths.length || !this.heights.count) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && (e.key === 'a' || e.key === 'A')) {
      e.preventDefault();
      this.selectAll();
      return;
    }
    if (mod && (e.key === 'c' || e.key === 'C')) {
      e.preventDefault();
      this.copySelection();
      return;
    }
    const a = this.anchor ?? { r: 0, c: 0 };
    if (e.key === 'F2') {
      e.preventDefault();
      this.openViewer(a.r, a.c);
      return;
    }
    // Shift moves the far corner of the range; otherwise the active cell moves and the range collapses.
    const from = e.shiftKey && this.focus ? this.focus : a;
    const pageRows = Math.max(1, Math.floor((this.viewH - HEADER_H) / BASE_ROW) - 1);
    const moves: Record<string, [number, number]> = {
      ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1],
      PageUp: [-pageRows, 0], PageDown: [pageRows, 0], Enter: [1, 0],
    };
    let r = from.r;
    let c = from.c;
    if (e.key in moves) {
      const [dr, dc] = moves[e.key];
      r += mod && dr ? Math.sign(dr) * this.heights.count : dr;
      c += mod && dc ? Math.sign(dc) * this.colWidths.length : dc;
    } else if (e.key === 'Home') {
      c = 0;
      if (mod) r = 0;
    } else if (e.key === 'End') {
      c = this.lastCol();
      if (mod) r = this.lastRow();
    } else {
      return;
    }
    e.preventDefault();
    r = Math.min(Math.max(0, r), this.lastRow());
    c = Math.min(Math.max(0, c), this.lastCol());
    if (e.shiftKey && e.key !== 'Enter') this.setSelection(a, { r, c });
    else this.setSelection({ r, c });
    this.scrollIntoView(r, c);
  }

  private cellText(r: number, c: number) {
    return this.row(r)?.[c] ?? '';
  }

  private scrollIntoView(r: number, c: number) {
    const top = this.heights.offsetOf(r);
    const bottom = top + this.heights.height(r);
    const bodyH = this.viewH - HEADER_H;
    if (top < this.virtualTop) this.virtualTop = top;
    else if (bottom > this.virtualTop + bodyH) this.virtualTop = bottom - bodyH;

    const left = this.colOffsets[c];
    const right = this.colOffsets[c + 1];
    const bodyW = this.viewW - this.gutterW;
    if (left < this.scrollLeft) this.scrollLeft = left;
    else if (right > this.scrollLeft + bodyW) this.scrollLeft = right - bodyW;
    this.render();
  }

  // ------------------------------------------------------------ resizing

  // Column resize: every pointermove re-wraps the visible rows through pretext's
  // arithmetic-only layout() — no DOM measurement, so it stays smooth on huge sheets.
  protected onResizeStart(e: PointerEvent, c: number) {
    e.preventDefault();
    e.stopPropagation();
    (e.target as HTMLElement).setPointerCapture(e.pointerId);
    this.resizing = { c, startX: e.clientX, startW: this.colWidths[c] };
  }

  protected onResizeMove(e: PointerEvent) {
    if (!this.resizing) return;
    const { c, startX, startW } = this.resizing;
    this.setColWidth(c, Math.max(MIN_COL, Math.round(startW + e.clientX - startX)));
  }

  protected onResizeEnd() {
    if (!this.resizing) return;
    this.resizing = null;
    this.pushLayout(true);
  }

  // Row resize: drag the bottom edge of a row number; double-click it to fit all the row's text.
  protected onRowResizeStart(e: PointerEvent, r: number) {
    e.preventDefault();
    e.stopPropagation();
    (e.target as HTMLElement).setPointerCapture(e.pointerId);
    this.resizingRow = { r, startY: e.clientY, startH: this.heights.height(r) };
  }

  protected onRowResizeMove(e: PointerEvent) {
    if (!this.resizingRow) return;
    const { r, startY, startH } = this.resizingRow;
    this.setRowHeight(r, Math.max(MIN_ROW, Math.round(startH + e.clientY - startY)));
  }

  protected onRowResizeEnd() {
    this.resizingRow = null;
  }

  protected onRowAutoFit(r: number) {
    const data = this.row(r);
    if (data !== undefined) this.setRowHeight(r, fullRowHeight(data, this.colWidths));
  }

  private setRowHeight(r: number, h: number) {
    this.rowOverrides.set(r, h);
    this.keepAnchor(() => {
      this.heights.set(r, h);
      this.heights.markMeasured(r);
    });
    this.render();
  }

  protected onAutoFit(c: number) {
    this.source().autofit(c, this.heights.rowAt(this.virtualTop)).then((w) => {
      this.setColWidth(c, w);
      this.pushLayout(true);
    });
  }

  private setColWidth(c: number, w: number) {
    if (this.colWidths[c] === w) return;
    this.colWidths[c] = w;
    this.rebuildColumns();
    if (this.wrapOn) {
      // Visible rows are re-measured exactly on the next frame; the worker re-estimates the rest.
      this.keepAnchor(() => this.heights.invalidate());
      this.pushLayout();
    }
    this.render();
  }
}
