import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';

describe('engine performance', () => {
  it('recalculates 50k chained formulas and incremental edits quickly', () => {
    const b = new Workbook();
    const e = new Engine(b);
    const s = b.sheet;
    const N = 50000;
    b.transact('fill', () => {
      for (let r = 0; r < N; r++) {
        s.set(r, 0, { v: r });
        s.set(r, 1, { f: `A${r + 1}*2` });
        s.set(r, 2, { f: r ? `C${r}+B${r + 1}` : 'B1' });
      }
      s.set(0, 3, { f: `SUM(B1:B${N})` });
    });
    let t = performance.now();
    e.recalc();
    const full = performance.now() - t;
    expect(s.value(0, 3)).toBe(N * (N - 1));
    t = performance.now();
    b.transact('edit', () => s.set(N - 1, 0, { v: 0 }));
    e.recalc();
    const inc = performance.now() - t;
    expect(s.value(N - 1, 2)).toBe(N * (N - 1) - 2 * (N - 1));
    console.log(`full ${full.toFixed(0)} ms, incremental ${inc.toFixed(0)} ms`);
    expect(full).toBeLessThan(15000);
    expect(inc).toBeLessThan(2000);
  }, 60000);
});

describe('deep chains', () => {
  it('evaluates a 20k reverse dependency chain without stack overflow', () => {
    const b = new Workbook();
    const e = new Engine(b);
    const s = b.sheet;
    const N = 20000;
    b.transact('fill', () => {
      for (let r = 0; r < N - 1; r++) s.set(r, 0, { f: `A${r + 2}+1` });
      s.set(N - 1, 0, { v: 1 });
    });
    e.recalc();
    expect(s.value(0, 0)).toBe(N);
    b.transact('edit', () => s.set(N - 1, 0, { v: 2 }));
    e.recalc();
    expect(s.value(0, 0)).toBe(N + 1);
  }, 60000);
});
