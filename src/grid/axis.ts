/**
 * Row/column geometry with sparse custom sizes: offset(i) and indexAt(px) in O(log k) where k is the
 * number of rows/columns with a non-default size (hidden ones have size 0).
 */
export class Axis {
  private keys: number[] = [];
  private deltaPrefix: number[] = [0]; // deltaPrefix[j] = sum of (size - def) for keys[0..j-1]
  private sizes = new Map<number, number>();

  constructor(public def: number, custom: Map<number, number>, hidden: Set<number>, public count: number) {
    for (const [i, s] of custom) this.sizes.set(i, s);
    for (const i of hidden) this.sizes.set(i, 0);
    this.keys = [...this.sizes.keys()].filter((i) => this.sizes.get(i) !== def).sort((a, b) => a - b);
    let acc = 0;
    for (const k of this.keys) { acc += (this.sizes.get(k)! - def); this.deltaPrefix.push(acc); }
  }

  size(i: number): number { const s = this.sizes.get(i); return s == null ? this.def : s; }

  /** Number of custom keys strictly less than i. */
  private rank(i: number): number {
    let lo = 0, hi = this.keys.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (this.keys[m] < i) lo = m + 1; else hi = m; }
    return lo;
  }

  offset(i: number): number { return i * this.def + this.deltaPrefix[this.rank(i)]; }

  total(): number { return this.offset(this.count); }

  /** Index containing pixel offset px (clamped to [0, count-1]). */
  indexAt(px: number): number {
    if (px <= 0) return 0;
    let lo = 0, hi = this.count - 1;
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1;
      if (this.offset(m) <= px) lo = m; else hi = m - 1;
    }
    // skip zero-size (hidden) entries
    while (lo < this.count - 1 && this.size(lo) === 0) lo++;
    return lo;
  }

  isHidden(i: number): boolean { return this.size(i) === 0; }

  /** Next visible index in a direction, or -1. */
  nextVisible(i: number, dir: 1 | -1): number {
    let j = i + dir;
    while (j >= 0 && j < this.count && this.size(j) === 0) j += dir;
    return j >= 0 && j < this.count ? j : -1;
  }
}
