import {
  ChangeDetectionStrategy, Component, DestroyRef, ElementRef, computed, effect, inject, signal, untracked, viewChild,
} from '@angular/core';
import { CellSelection, SheetGrid } from './grid/sheet-grid';
import { SheetSource } from './grid/sheet-source';
import type { FromWorker, LoadPhase, ToWorker } from './parser/protocol';

type Status = 'idle' | LoadPhase | 'done' | 'error';

/** Above this many rows wrapping starts off: no height work at all until the user enables it. */
const AUTO_NOWRAP_ROWS = 500_000;

@Component({
  selector: 'app-root',
  imports: [SheetGrid],
  templateUrl: './app.html',
  styleUrl: './app.css',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { '(document:keydown)': 'onDocumentKey($event)' },
})
export class App {
  private readonly worker = new Worker(new URL('./parser/parser.worker', import.meta.url), { type: 'module' });
  private job = 0;
  private startedAt = 0;

  protected readonly fileName = signal('');
  protected readonly fileSize = signal(0);
  protected readonly format = signal('');
  protected readonly accept = '.xlsx,.xlsm,.xltx,.xltm,.xls,.xlt,.ods,.csv,.tsv,.txt,.xml,.htm,.html';
  protected readonly sheets = signal<string[]>([]);
  protected readonly activeSheet = signal(0);
  protected readonly store = signal<SheetSource | null>(null);
  /** Per sheet: wrap was switched off automatically once; never again after the user decides. */
  private autoWrapDone = false;
  protected readonly status = signal<Status>('idle');
  protected readonly progress = signal(0);
  protected readonly error = signal('');
  protected readonly elapsed = signal(0);
  protected readonly rowCount = signal(0);
  protected readonly colCount = signal(0);
  protected readonly wrap = signal(true);
  protected readonly selection = signal<CellSelection | null>(null);
  protected readonly dragging = signal(false);

  private readonly grid = viewChild(SheetGrid);
  protected readonly layoutProgress = computed(() => this.store()?.layoutProgress() ?? 1);
  protected readonly lastLayoutMs = computed(() => this.grid()?.lastLayoutMs() ?? 0);

  // Search
  private readonly searchInput = viewChild<ElementRef<HTMLInputElement>>('searchInput');
  protected readonly searchOpen = signal(false);
  protected readonly searchText = signal('');
  /** Query the grid is actually searching (applied after a short typing pause). */
  private appliedQuery = '';
  private searchTimer = 0;
  protected readonly searchLabel = computed(() => {
    const g = this.grid();
    if (!g || !this.searchText().trim()) return '';
    const count = g.searchCount();
    const done = g.searchDone();
    const more = g.searchTruncated() ? '+' : '';
    const busy = done ? '' : '…';
    if (!count) return done ? 'No results' : 'Searching…';
    const i = g.searchIndex();
    return i >= 0 ? `${(i + 1).toLocaleString()} / ${count.toLocaleString()}${more}${busy}` : `${count.toLocaleString()}${more} found${busy}`;
  });

  protected readonly notice = signal('');
  private noticeTimer = 0;

  protected readonly loading = computed(() => {
    const s = this.status();
    return s === 'reading' || s === 'strings' || s === 'rows';
  });
  protected readonly statusLabel = computed(() => {
    const pct = Math.round(this.progress() * 100);
    switch (this.status()) {
      case 'reading': return 'Reading archive…';
      case 'strings': return `Shared strings ${pct}%`;
      case 'rows': return `Parsing rows ${pct}%`;
      case 'done': return `Parsed in ${(this.elapsed() / 1000).toFixed(2)}s`;
      default: return '';
    }
  });

  constructor() {
    this.worker.onmessage = (e: MessageEvent<FromWorker>) => this.onWorker(e.data);
    this.worker.onerror = (e) => this.fail(e.message || 'Worker crashed');
    inject(DestroyRef).onDestroy(() => {
      this.worker.terminate();
      clearTimeout(this.searchTimer);
      clearTimeout(this.noticeTimer);
    });
    // A new grid (new file) starts with the current search applied.
    effect(() => {
      const g = this.grid();
      if (g) untracked(() => g.setSearch(this.appliedQuery));
    });
  }

  protected openFile(file: File | undefined) {
    if (!file) return;
    // No extension check: the worker detects the real format from the file's bytes.
    this.fileName.set(file.name);
    this.format.set('');
    this.fileSize.set(file.size);
    this.sheets.set([]);
    this.activeSheet.set(0);
    this.store()?.dispose();
    this.store.set(null);
    this.selection.set(null);
    this.error.set('');
    this.startedAt = performance.now();
    this.send({ type: 'open', file, job: ++this.job });
  }

  protected selectSheet(index: number) {
    if (index === this.activeSheet() && this.store()) return;
    this.activeSheet.set(index);
    this.rowCount.set(0);
    this.colCount.set(0);
    this.selection.set(null);
    this.autoWrapDone = false;
    this.startedAt = performance.now();
    const job = ++this.job;
    const wrap = this.wrap();
    this.store()?.dispose();
    this.store.set(new SheetSource(job, wrap, (msg) => this.worker.postMessage(msg)));
    this.send({ type: 'sheet', index, job, wrap });
  }

  protected onPick(e: Event) {
    const input = e.target as HTMLInputElement;
    this.openFile(input.files?.[0]);
    input.value = '';
  }

  protected onDrop(e: DragEvent) {
    e.preventDefault();
    this.dragging.set(false);
    this.openFile(e.dataTransfer?.files[0]);
  }

  protected onDragOver(e: DragEvent) {
    e.preventDefault();
    this.dragging.set(true);
  }

  protected formatBytes(n: number) {
    const units = ['B', 'KB', 'MB', 'GB'];
    let i = 0;
    while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
    return `${n.toFixed(i ? 1 : 0)} ${units[i]}`;
  }

  // ------------------------------------------------------------ search

  protected onDocumentKey(e: KeyboardEvent) {
    if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F') && this.store()) {
      e.preventDefault(); // the browser's find can't see virtualized rows
      this.openSearch();
    }
  }

  protected toggleSearch() {
    if (this.searchOpen()) this.closeSearch();
    else this.openSearch();
  }

  private openSearch() {
    this.searchOpen.set(true);
    setTimeout(() => {
      const input = this.searchInput()?.nativeElement;
      input?.focus();
      input?.select();
    });
  }

  protected closeSearch() {
    this.searchOpen.set(false);
    this.searchText.set('');
    this.applySearch('');
  }

  protected onSearchInput(value: string) {
    this.searchText.set(value);
    clearTimeout(this.searchTimer);
    this.searchTimer = setTimeout(() => this.applySearch(value), 120);
  }

  protected onSearchKey(e: KeyboardEvent) {
    if (e.key === 'Enter') {
      e.preventDefault();
      this.find(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      this.closeSearch();
    }
  }

  protected find(dir: 1 | -1) {
    // Typing pause not over yet: search the current text right away.
    if (this.searchText() !== this.appliedQuery) this.applySearch(this.searchText());
    this.grid()?.findNext(dir);
  }

  private applySearch(query: string) {
    clearTimeout(this.searchTimer);
    this.appliedQuery = query;
    this.grid()?.setSearch(query);
  }

  protected showNotice(message: string) {
    this.notice.set(message);
    clearTimeout(this.noticeTimer);
    this.noticeTimer = setTimeout(() => this.notice.set(''), 2500);
  }

  private send(msg: ToWorker) {
    this.status.set('reading');
    this.progress.set(0);
    this.worker.postMessage(msg);
  }

  private onWorker(msg: FromWorker) {
    if (msg.job !== this.job) return; // stale message from a cancelled job
    this.store()?.handle(msg);
    switch (msg.type) {
      case 'meta':
        this.sheets.set(msg.sheets);
        this.format.set(msg.format);
        this.selectSheet(0);
        break;
      case 'progress':
        this.status.set(msg.phase);
        this.progress.set(msg.total ? Math.min(1, msg.loaded / msg.total) : 0);
        break;
      case 'grow':
        this.rowCount.set(msg.rowCount);
        this.colCount.set(msg.colCount);
        if (!this.autoWrapDone && msg.rowCount > AUTO_NOWRAP_ROWS && this.wrap()) {
          this.autoWrapDone = true;
          this.wrap.set(false);
          this.showNotice(`Wrap text turned off for this large sheet (over ${AUTO_NOWRAP_ROWS.toLocaleString()} rows). Turn it on any time.`);
        }
        break;
      case 'done':
        this.rowCount.set(msg.rowCount);
        this.colCount.set(msg.colCount);
        this.elapsed.set(performance.now() - this.startedAt);
        this.status.set('done');
        break;
      case 'error':
        this.fail(msg.message);
        break;
    }
  }

  private fail(message: string) {
    this.error.set(message);
    this.status.set('error');
  }
}
