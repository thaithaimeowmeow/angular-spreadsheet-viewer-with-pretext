import type { FromWorker, LoadPhase, Row } from './protocol';

const postRaw = postMessage as unknown as (msg: FromWorker, transfer: Transferable[]) => void;
export const post = (msg: FromWorker, transfer: Transferable[] = []) => postRaw(msg, transfer);

export class Cancelled extends Error {}

let activeJob = 0;
export const setActiveJob = (id: number) => (activeJob = id);
export const isActiveJob = (id: number) => id === activeJob;

// A macrotask via MessageChannel: unlike setTimeout it is not clamped to 4 ms when nested.
const channel = new MessageChannel();
const waiters: (() => void)[] = [];
channel.port1.onmessage = () => waiters.shift()?.();
export const macrotask = () => new Promise<void>((resolve) => {
  waiters.push(resolve);
  channel.port2.postMessage(0);
});

/** One unit of work (open or load sheet). A newer control message makes older jobs throw Cancelled. */
export class Job {
  private lastYield = performance.now();

  constructor(readonly id: number) {}

  check() {
    if (this.id !== activeJob) throw new Cancelled();
  }

  /**
   * Stream reads resolve as microtasks; yield a macrotask every ~12 ms so queries (window,
   * search, ...) and newer control messages get served while parsing.
   */
  async tick() {
    if (performance.now() - this.lastYield > 12) {
      await macrotask();
      this.lastYield = performance.now();
    }
    this.check();
  }

  progress(phase: LoadPhase, loaded: number, total: number) {
    post({ type: 'progress', job: this.id, phase, loaded, total });
  }

  /** Returns a byte counter that posts throttled progress. */
  byteProgress(phase: LoadPhase, total: number) {
    let loaded = 0;
    let lastPost = 0;
    this.progress(phase, 0, total);
    return (n: number) => {
      loaded += n;
      const now = performance.now();
      if (now - lastPost > 80) {
        lastPost = now;
        this.progress(phase, loaded, total);
      }
    };
  }
}

/** Receives parsed rows in (mostly) increasing order; implemented by the worker's sheet session. */
export interface RowEmitter {
  row(r: number, row: Row): void;
}

export interface Book {
  format: string;
  sheets: string[];
  load(index: number, job: Job, out: RowEmitter): Promise<void>;
}
