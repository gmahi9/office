export class SizeIndex {
  readonly count: number;
  private readonly sizes: Float32Array;
  private readonly prefix: Float64Array;

  constructor(
    count: number,
    fallback: number,
    hidden: Iterable<number>,
    overrides?: Record<number, number>,
    scale = 1,
  ) {
    this.count = count;
    const hiddenSet = new Set(hidden);
    this.sizes = new Float32Array(count);
    this.prefix = new Float64Array(count + 1);
    let acc = 0;
    for (let i = 0; i < count; i += 1) {
      const custom = overrides?.[i];
      const size = hiddenSet.has(i) ? 0 : (custom ?? fallback) * scale;
      this.sizes[i] = size;
      acc += size;
      this.prefix[i + 1] = acc;
    }
  }

  size(index: number): number {
    return this.sizes[index] ?? 0;
  }

  offset(index: number): number {
    if (index <= 0) return 0;
    if (index >= this.count) return this.prefix[this.count] ?? 0;
    return this.prefix[index] ?? 0;
  }

  total(): number {
    return this.prefix[this.count] ?? 0;
  }

  span(start: number, end: number): number {
    return this.offset(end + 1) - this.offset(start);
  }

  /** Largest index whose leading edge is at or before `px`, skipping hidden tracks. */
  indexAt(px: number): number {
    if (this.count === 0) return 0;
    if (px <= 0) return this.nextVisible(0);
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if ((this.prefix[mid] ?? 0) <= px) lo = mid;
      else hi = mid - 1;
    }
    return this.nextVisible(lo);
  }

  nextVisible(index: number): number {
    let n = Math.max(0, Math.min(this.count - 1, index));
    while (n < this.count - 1 && this.sizes[n] === 0) n += 1;
    return n;
  }

  prevVisible(index: number): number {
    let n = Math.max(0, Math.min(this.count - 1, index));
    while (n > 0 && this.sizes[n] === 0) n -= 1;
    return n;
  }

  isHidden(index: number): boolean {
    return this.sizes[index] === 0;
  }
}

export function visibleSpan(index: SizeIndex, origin: number, scroll: number, view: number, buffer = 2): { start: number; end: number } {
  const base = index.offset(origin);
  const start = Math.max(origin, index.indexAt(base + scroll) - buffer);
  const endIndex = index.indexAt(base + scroll + view);
  const end = Math.min(index.count - 1, endIndex + buffer);
  return { start, end: Math.max(start, end) };
}
