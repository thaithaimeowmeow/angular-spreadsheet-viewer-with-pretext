/**
 * Variable row heights backed by a Fenwick (binary indexed) tree:
 * O(log n) height update, row → offset, and offset → row — fine for millions of rows.
 */
export class RowHeights {
  private heights = new Float32Array(0);
  private tree = new Float64Array(1);
  private stamp = new Uint32Array(0);
  private generation = 1;
  count = 0;

  constructor(private readonly base: number) {}

  get capacity() {
    return this.heights.length;
  }

  /** Grows to `n` rows; new rows start at the base height (unmeasured). */
  resize(n: number) {
    if (n <= this.count) return;
    if (n > this.capacity) {
      const cap = Math.max(n, this.capacity * 2, 1024);
      const heights = new Float32Array(cap);
      heights.set(this.heights);
      const stamp = new Uint32Array(cap);
      stamp.set(this.stamp);
      heights.fill(this.base, this.count, n);
      this.heights = heights;
      this.stamp = stamp;
      this.count = n;
      this.rebuild();
      return;
    }
    for (let i = this.count; i < n; i++) this.set(i, this.base);
    this.count = n;
  }

  reset() {
    this.heights = new Float32Array(0);
    this.tree = new Float64Array(1);
    this.stamp = new Uint32Array(0);
    this.count = 0;
    this.generation++;
  }

  height(i: number) {
    return this.heights[i];
  }

  set(i: number, h: number) {
    const d = h - this.heights[i];
    if (d === 0) return;
    this.heights[i] = h;
    const cap = this.capacity;
    for (let j = i + 1; j <= cap; j += j & -j) this.tree[j] += d;
  }

  isMeasured(i: number) {
    return this.stamp[i] === this.generation;
  }

  markMeasured(i: number) {
    this.stamp[i] = this.generation;
  }

  unmark(i: number) {
    this.stamp[i] = 0;
  }

  /**
   * Applies worker height estimates to rows that were not measured exactly in this generation.
   * Large chunks are written in bulk and the tree rebuilt once (O(n)) instead of per-row updates.
   */
  applyEstimates(start: number, values: Uint16Array) {
    const end = Math.min(this.count, start + values.length);
    if (end <= start) return;
    if (end - start > this.count / 16) {
      for (let i = start; i < end; i++) {
        if (this.stamp[i] !== this.generation) this.heights[i] = values[i - start];
      }
      this.rebuild();
      return;
    }
    for (let i = start; i < end; i++) {
      if (this.stamp[i] !== this.generation) this.set(i, values[i - start]);
    }
  }

  /** Marks every row stale while keeping current heights as estimates (no visual jump). */
  invalidate() {
    this.generation++;
  }

  /** Sets all rows to one height and marks them measured (used when wrapping is off). */
  fill(h: number) {
    this.heights.fill(h, 0, this.count);
    this.rebuild();
    this.generation++;
    this.stamp.fill(this.generation, 0, this.count);
  }

  /** Sum of heights of rows [0, i). */
  offsetOf(i: number): number {
    let s = 0;
    for (let j = Math.min(i, this.count); j > 0; j -= j & -j) s += this.tree[j];
    return s;
  }

  total(): number {
    return this.offsetOf(this.count);
  }

  /** Index of the row containing `offset`, clamped to [0, count - 1]. */
  rowAt(offset: number): number {
    if (this.count === 0 || offset <= 0) return 0;
    const cap = this.capacity;
    let pos = 0;
    let rem = offset;
    for (let step = 1 << Math.floor(Math.log2(cap)); step > 0; step >>= 1) {
      const next = pos + step;
      if (next <= cap && this.tree[next] <= rem) {
        pos = next;
        rem -= this.tree[next];
      }
    }
    return Math.min(pos, this.count - 1);
  }

  private rebuild() {
    const cap = this.capacity;
    const tree = new Float64Array(cap + 1);
    for (let i = 0; i < cap; i++) tree[i + 1] = this.heights[i];
    for (let i = 1; i <= cap; i++) {
      const j = i + (i & -i);
      if (j <= cap) tree[j] += tree[i];
    }
    this.tree = tree;
  }
}
