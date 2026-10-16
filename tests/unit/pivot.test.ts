import { describe, it, expect } from 'vitest';
import { computePivot, drillRows, cellValue, pivotItems, normalizeHeaders, type PivotSource } from '../../src/model/pivot';
import type { PivotDef } from '../../src/model/types';
import { partsToSerial } from '../../src/model/dates';

const src: PivotSource = {
  headers: ['Region', 'Product', 'Qty', 'Sales'],
  rows: [
    ['East', 'Apples', 3, 30],
    ['West', 'Apples', 5, 50],
    ['East', 'Pears', 2, 40],
    ['West', 'Pears', 1, 20],
    ['East', 'Apples', 4, 45],
    ['North', 'Plums', 6, null],
  ],
};

function def(p: Partial<PivotDef>): PivotDef {
  return { id: 'p1', source: 'x', target: { sheet: 'S', r: 0, c: 0 }, rows: [], cols: [], filters: [], values: [], layout: 'compact', grandRows: true, grandCols: true, subtotals: true, ...p };
}

const grid = (cells: any[][]) => cells.map((r) => r.map(cellValue));

describe('computePivot', () => {
  it('sums by one row field with a grand total', () => {
    const r = computePivot(def({ rows: ['Region'], values: [{ field: 'Sales', fn: 'sum' }] }), src);
    expect(grid(r.cells)).toEqual([
      ['Row Labels', 'Sum of Sales'],
      ['East', 115],
      ['North', 0],
      ['West', 70],
      ['Grand Total', 185],
    ]);
  });

  it('counts, averages and multiple value fields', () => {
    const r = computePivot(def({ rows: ['Product'], values: [{ field: 'Sales', fn: 'count' }, { field: 'Qty', fn: 'average' }, { field: 'Qty', fn: 'max' }] }), src);
    const g = grid(r.cells);
    expect(g[0]).toEqual(['Row Labels', 'Count of Sales', 'Average of Qty', 'Max of Qty']);
    expect(g[1]).toEqual(['Apples', 3, 4, 5]);
    expect(g[3]).toEqual(['Plums', 0, 6, 6]);
    expect(g[4]).toEqual(['Grand Total', 5, 21 / 6, 6]);
  });

  it('row and column fields with totals in both directions', () => {
    const r = computePivot(def({ rows: ['Region'], cols: ['Product'], values: [{ field: 'Qty', fn: 'sum' }] }), src);
    const g = grid(r.cells);
    expect(g[0]).toEqual(['Sum of Qty', 'Column Labels', null, null, null]);
    expect(g[1]).toEqual(['Row Labels', 'Apples', 'Pears', 'Plums', 'Grand Total']);
    expect(g[2]).toEqual(['East', 7, 2, null, 9]);
    expect(g[4]).toEqual(['West', 5, 1, null, 6]);
    expect(g[5]).toEqual(['Grand Total', 12, 3, 6, 21]);
    expect(r.headerRows).toBe(2);
  });

  it('nested row fields: compact subtotals, tabular subtotal rows', () => {
    const d = def({ rows: ['Region', 'Product'], values: [{ field: 'Qty', fn: 'sum' }] });
    const g = grid(computePivot(d, src).cells);
    expect(g.slice(1, 4)).toEqual([['East', 9], ['Apples', 7], ['Pears', 2]]);
    const t = grid(computePivot({ ...d, layout: 'tabular' }, src).cells);
    expect(t[0]).toEqual(['Region', 'Product', 'Sum of Qty']);
    expect(t[1]).toEqual(['East', 'Apples', 7]);
    expect(t[2]).toEqual([null, 'Pears', 2]);
    expect(t[3]).toEqual(['East Total', null, 9]);
    const o = grid(computePivot({ ...d, layout: 'outline', subtotals: false }, src).cells);
    expect(o[1]).toEqual(['East', null, null]);
    expect(o[2]).toEqual([null, 'Apples', 7]);
  });

  it('show values as percent of total, row, running total and rank', () => {
    const pct = grid(computePivot(def({ rows: ['Region'], values: [{ field: 'Qty', fn: 'sum', showAs: 'percentOfTotal' }] }), src).cells);
    expect(pct[1][1]).toBeCloseTo(9 / 21);
    expect(pct[4][1]).toBeCloseTo(1);
    const row = grid(computePivot(def({ rows: ['Region'], cols: ['Product'], values: [{ field: 'Qty', fn: 'sum', showAs: 'percentOfRow' }] }), src).cells);
    expect(row[2][1]).toBeCloseTo(7 / 9);
    const run = grid(computePivot(def({ rows: ['Region'], values: [{ field: 'Qty', fn: 'sum', showAs: 'runningTotal' }] }), src).cells);
    expect(run.slice(1, 4).map((x) => x[1])).toEqual([9, 15, 21]);
    const rank = grid(computePivot(def({ rows: ['Region'], values: [{ field: 'Qty', fn: 'sum', showAs: 'rank' }] }), src).cells);
    expect(rank.slice(1, 4).map((x) => x[1])).toEqual([1, 2, 2]);
  });

  it('report filters limit the records and are listed above the table', () => {
    const r = computePivot(def({ rows: ['Product'], filters: ['Region'], filterValues: { Region: ['East'] }, values: [{ field: 'Sales', fn: 'sum' }] }), src);
    const g = grid(r.cells);
    expect(g[0]).toEqual(['Region', 'East']);
    expect(r.filterRows).toBe(2);
    expect(g[2]).toEqual(['Row Labels', 'Sum of Sales']);
    expect(g[g.length - 1]).toEqual(['Grand Total', 115]);
    expect(r.records).toBe(3);
  });

  it('stdDev, var and product', () => {
    const g = grid(computePivot(def({ values: [{ field: 'Qty', fn: 'stdDevP' }, { field: 'Qty', fn: 'var' }, { field: 'Qty', fn: 'product' }] }), src).cells);
    const xs = [3, 5, 2, 1, 4, 6], m = 3.5;
    expect(g[1][0]).toBeCloseTo(Math.sqrt(xs.reduce((a, x) => a + (x - m) ** 2, 0) / 6));
    expect(g[1][1]).toBeCloseTo(xs.reduce((a, x) => a + (x - m) ** 2, 0) / 5);
    expect(g[1][2]).toBe(720);
  });

  it('groups dates by years and quarters', () => {
    const d1 = partsToSerial(2023, 2, 10), d2 = partsToSerial(2023, 11, 1), d3 = partsToSerial(2024, 5, 5);
    const s: PivotSource = { headers: ['Date', 'Amt'], rows: [[d1, 1], [d2, 2], [d3, 4]] };
    const g = grid(computePivot(def({ rows: ['Date'], dateGroups: { Date: ['years', 'quarters'] }, values: [{ field: 'Amt', fn: 'sum' }] }), s).cells);
    expect(g.slice(1)).toEqual([[2023, 3], ['Qtr1', 1], ['Qtr4', 2], [2024, 4], ['Qtr2', 4], ['Grand Total', 7]]);
  });

  it('drills down to the source rows behind a value', () => {
    const d = def({ rows: ['Region'], cols: ['Product'], values: [{ field: 'Qty', fn: 'sum' }] });
    const r = computePivot(d, src);
    const info = r.drill.get('2,1')!;
    expect(drillRows(d, src, info.row, info.col)).toEqual([0, 4]);
    const grand = r.drill.get('5,4')!;
    expect(drillRows(d, src, grand.row, grand.col).length).toBe(6);
  });

  it('sorts descending, lists items and normalizes headers', () => {
    const g = grid(computePivot(def({ rows: ['Region'], sort: { Region: 'desc' }, values: [{ field: 'Qty', fn: 'sum' }] }), src).cells);
    expect(g.slice(1, 4).map((x) => x[0])).toEqual(['West', 'North', 'East']);
    expect(pivotItems(src, 'Region')).toEqual(['East', 'North', 'West']);
    expect(normalizeHeaders(['A', null, 'a', 3])).toEqual(['A', 'Column2', 'a2', '3']);
  });

  it('shows a placeholder for an empty layout', () => {
    const r = computePivot(def({ name: 'PivotTable1' }), src);
    expect(String(cellValue(r.cells[0][0]))).toMatch(/PivotTable Fields/);
  });
});

describe('pivot commands helpers', () => {
  it('reads a source range, places fields and recommends layouts', async () => {
    const { Workbook } = await import('../../src/model/workbook');
    const { readSource, placeField, recommendPivots } = await import('../../src/commands/pivot');
    const b = new Workbook();
    const s = b.sheet;
    b.transact('d', () => [['Region', 'Qty'], ['East', 3], ['West', 4], ['East', 5]].forEach((row, r) => row.forEach((v, c) => s.set(r, c, { v }))));
    const src = readSource(b, 'Sheet1!$A$1:$B$4', s);
    if (typeof src === 'string') throw new Error(src);
    expect(src.headers).toEqual(['Region', 'Qty']);
    expect(src.rows.length).toBe(3);
    expect(typeof readSource(b, 'Nope!A1', s)).toBe('string');
    const d = def({});
    placeField(d, 'Region', 'rows');
    placeField(d, 'Qty', 'values');
    placeField(d, 'Region', 'cols', { from: 'rows' });
    expect(d.rows).toEqual([]);
    expect(d.cols).toEqual(['Region']);
    expect(d.values).toEqual([{ field: 'Qty', fn: 'sum' }]);
    placeField(d, 'Qty', null);
    expect(d.values).toEqual([]);
    const recs = recommendPivots(src);
    expect(recs[0].title).toBe('Sum of Qty by Region');
    expect(grid(computePivot({ ...def({}), ...recs[0].def } as PivotDef, src).cells).pop()).toEqual(['Grand Total', 12]);
  });
});
