import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook, type Sheet } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell, type Range } from '../../src/model/address';
import { shiftCells } from '../../src/model/structure';
import { partsToSerial } from '../../src/model/dates';
import type { Scalar } from '../../src/model/types';
import {
  sortRange, compareForSort, criteriaPasses, customMatch, applyFilters, distinctColumnValues, advancedCriteria, splitLine, suggestBreaks,
  convertSplitField, inferFlashFill, runFlashFill, duplicateRows, removeRowsInRange, addSubtotals, removeSubtotals, changeLevel, showOutlineLevel,
  setDetail, autoOutline, consolidate, applySteps, periodRange, subtractRange, pickListValues, extendFilterRange, columnStats,
} from '../../src/commands/dataops';
import { Substituter, computeDataTable, goalSeek, simplex, solveLP, solveProblem, nelderMead, buildSheetProblem, createForecastSheet, linearize, saveSolverModel, loadSolverModel, scenarioResults, parseCells } from '../../src/commands/whatif';
import {
  descriptiveStatistics, histogram, correlationMatrix, regression, tTest, zTest, fTest, anovaSingle, anovaTwoNoRep, anovaTwoRep, movingAverage,
  exponentialSmoothing, rankAndPercentile, randomNumbers, sampling, fourier, parseComplex, formatComplex, fft, invertMatrix,
} from '../../src/commands/analysis';
import { dvStore } from '../../src/commands/data';

function book() {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  const rc = (a: string) => { const p = parseCell(a)!; return p; };
  const set = (a: string, v: any) => {
    const { r, c } = rc(a);
    b.transact('t', () => s.set(r, c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : v == null ? undefined : { v }));
  };
  const rows = (start: string, data: any[][]) => { const { r, c } = rc(start); b.transact('t', () => data.forEach((row, i) => row.forEach((v, j) => s.set(r + i, c + j, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : v == null ? undefined : { v })))); e.recalc(); };
  const get = (a: string) => { e.recalc(); const { r, c } = rc(a); return s.value(r, c); };
  const f = (a: string) => { const { r, c } = rc(a); return s.get(r, c)?.f; };
  const col = (c: number, r1: number, r2: number) => { e.recalc(); const out: Scalar[] = []; for (let r = r1; r <= r2; r++) out.push(s.value(r, c)); return out; };
  return { b, e, s, set, rows, get, f, col };
}
const R = (r1: number, c1: number, r2: number, c2: number): Range => ({ r1, c1, r2, c2 });

// --------------------------------------------------------------------------------------------

describe('sort', () => {
  it('orders values like a spreadsheet (numbers, text, logicals, errors, blanks last)', () => {
    const vals: Scalar[] = ['b', 3, null, true, 'A', 1, false];
    const asc = [...vals].sort((a, b) => compareForSort(a, b));
    expect(asc).toEqual([1, 3, 'A', 'b', false, true, null]);
    const desc = [...vals].sort((a, b) => compareForSort(a, b, true));
    expect(desc).toEqual([true, false, 'b', 'A', 3, 1, null]);
  });

  it('sorts by several keys with a header row and moves whole rows', () => {
    const t = book();
    t.rows('A1', [['Name', 'Dept', 'Pay'], ['Cy', 'B', 30], ['Al', 'A', 20], ['Bo', 'B', 10], ['Di', 'A', 40]]);
    t.b.transact('Sort', () => sortRange(t.s, R(0, 0, 4, 2), [{ index: 1 }, { index: 2, desc: true }], { header: true }));
    expect(t.col(0, 0, 4)).toEqual(['Name', 'Di', 'Al', 'Cy', 'Bo']);
    expect(t.col(2, 1, 4)).toEqual([40, 20, 30, 10]);
    t.b.undo();
    expect(t.col(0, 1, 4)).toEqual(['Cy', 'Al', 'Bo', 'Di']);
  });

  it('formulas referring to their own row keep doing so after the move', () => {
    const t = book();
    t.rows('A1', [['Item', 'Qty', 'Twice'], ['x', 3, '=B2*2'], ['y', 1, '=B3*2'], ['z', 2, '=B4*2']]);
    t.b.transact('Sort', () => sortRange(t.s, R(0, 0, 3, 2), [{ index: 1 }], { header: true }));
    expect(t.col(0, 1, 3)).toEqual(['y', 'z', 'x']);
    expect(t.f('C2')).toBe('B2*2');
    expect(t.f('C4')).toBe('B4*2');
    expect(t.col(2, 1, 3)).toEqual([2, 4, 6]);
  });

  it('moves styles and comments, supports custom lists, case sensitivity and color keys', () => {
    const t = book();
    t.rows('A1', [['Mar'], ['Jan'], ['Feb']]);
    const red = t.b.styleId({ fill: '#ff0000' });
    t.b.transact('s', () => { t.s.set(1, 0, { v: 'Jan', s: red }); t.s.comments.set(1 * 16384, { author: 'a', text: 'jan note', time: 0 }); });
    t.b.transact('Sort', () => sortRange(t.s, R(0, 0, 2, 0), [{ index: 0, list: ['Jan', 'Feb', 'Mar'] }], {}));
    expect(t.col(0, 0, 2)).toEqual(['Jan', 'Feb', 'Mar']);
    expect(t.s.get(0, 0)?.s).toBe(red);
    expect(t.s.comments.get(0)?.text).toBe('jan note');
    t.b.transact('Sort', () => sortRange(t.s, R(0, 0, 2, 0), [{ index: 0, on: 'fill', color: '#ff0000', bottom: true }], {}));
    expect(t.col(0, 0, 2)).toEqual(['Feb', 'Mar', 'Jan']);
    const u = book();
    u.rows('A1', [['B'], ['a'], ['A'], ['b']]);
    u.b.transact('Sort', () => sortRange(u.s, R(0, 0, 3, 0), [{ index: 0 }], { caseSensitive: true }));
    expect(u.col(0, 0, 3)).toEqual(['a', 'A', 'b', 'B']);
  });

  it('sorts left to right and refuses ranges with multi-row merges', () => {
    const t = book();
    t.rows('A1', [[3, 1, 2], ['c', 'a', 'b']]);
    t.b.transact('Sort', () => sortRange(t.s, R(0, 0, 1, 2), [{ index: 0 }], { byCols: true }));
    expect([t.get('A2'), t.get('B2'), t.get('C2')]).toEqual(['a', 'b', 'c']);
    t.b.transact('m', () => { t.b.recordSheet(t.s); t.s.merges.push(R(0, 0, 1, 0)); });
    expect(() => t.b.transact('Sort', () => sortRange(t.s, R(0, 0, 1, 2), [{ index: 0 }], {}))).toThrow(/merged/);
  });

  it('leaves hidden (filtered) rows in place', () => {
    const t = book();
    t.rows('A1', [[5], [4], [3], [2], [1]]);
    t.s.hiddenRows.add(2);
    t.b.transact('Sort', () => sortRange(t.s, R(0, 0, 4, 0), [{ index: 0 }], { skipHidden: true }));
    expect(t.col(0, 0, 4)).toEqual([1, 2, 3, 4, 5].filter(() => true).map((_, i) => [1, 2, 3, 4, 5][i]).map((x, i) => (i === 2 ? 3 : x)));
    expect(t.get('A3')).toBe(3);
  });
});

describe('filter', () => {
  it('matches checkbox lists, blanks and custom conditions', () => {
    expect(criteriaPasses({ values: ['Apple'] }, 'apple', 'apple')).toBe(true);
    expect(criteriaPasses({ values: ['Apple'] }, null, '')).toBe(false);
    expect(criteriaPasses({ values: ['Apple'], blanks: true }, null, '')).toBe(true);
    expect(criteriaPasses({ custom: [{ op: 'gt', value: '10' }, { op: 'lt', value: '20' }], and: true }, 15, '15')).toBe(true);
    expect(criteriaPasses({ custom: [{ op: 'gt', value: '10' }, { op: 'lt', value: '20' }], and: true }, 25, '25')).toBe(false);
    expect(criteriaPasses({ custom: [{ op: 'lt', value: '10' }, { op: 'gt', value: '20' }], and: false }, 25, '25')).toBe(true);
    expect(customMatch('eq', 'a*e', 'Apple', 'Apple')).toBe(true);
    expect(customMatch('eq', 'b?t', 'bat', 'bat')).toBe(true);
    expect(customMatch('begins', 'ap', 'Apple', 'Apple')).toBe(true);
    expect(customMatch('notContains', 'pp', 'Apple', 'Apple')).toBe(false);
    expect(customMatch('ends', 'LE', 'Apple', 'Apple')).toBe(true);
    expect(customMatch('ne', '5', 5, '5')).toBe(false);
    expect(customMatch('ge', 'b', 'c', 'c')).toBe(true);
    expect(customMatch('lt', '1/15/2024', partsToSerial(2024, 1, 10), '1/10/2024')).toBe(true);
  });

  it('top 10, average and date periods', () => {
    const stats = columnStats([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    expect(criteriaPasses({ top10: { top: true, percent: false, n: 3 } }, 8, '8', { stats })).toBe(true);
    expect(criteriaPasses({ top10: { top: true, percent: false, n: 3 } }, 7, '7', { stats })).toBe(false);
    expect(criteriaPasses({ top10: { top: false, percent: true, n: 20 } }, 2, '2', { stats })).toBe(true);
    expect(criteriaPasses({ top10: { top: false, percent: true, n: 20 } }, 3, '3', { stats })).toBe(false);
    expect(criteriaPasses({ dynamic: 'aboveAverage' }, 6, '6', { stats })).toBe(true);
    expect(criteriaPasses({ dynamic: 'belowAverage' }, 6, '6', { stats })).toBe(false);
    const today = partsToSerial(2024, 5, 15); // Wednesday
    expect(periodRange('thisMonth', today)).toEqual([partsToSerial(2024, 5, 1), partsToSerial(2024, 6, 1)]);
    expect(periodRange('lastQuarter', today)).toEqual([partsToSerial(2024, 1, 1), partsToSerial(2024, 4, 1)]);
    expect(periodRange('thisWeek', today)).toEqual([partsToSerial(2024, 5, 12), partsToSerial(2024, 5, 19)]);
    expect(customMatch('period', 'yesterday', today - 1, '', { today })).toBe(true);
    expect(customMatch('period', 'Q2', partsToSerial(2020, 6, 30), '')).toBe(true);
    expect(customMatch('period', 'M12', partsToSerial(2020, 6, 30), '')).toBe(false);
    expect(criteriaPasses({ color: { fill: '#FF0000' } }, 1, '1', { style: { fill: '#ff0000' } })).toBe(true);
    expect(criteriaPasses({ color: { fill: '' } }, 1, '1', { style: { fill: '#ff0000' } })).toBe(false);
  });

  it('applies AutoFilter and table filters by hiding rows, with undo', () => {
    const t = book();
    t.rows('A1', [['Fruit', 'Qty'], ['Apple', 5], ['Pear', 12], ['Apple', 20], ['Plum', null], ['Kiwi', 8]]);
    t.b.transact('f', () => { t.b.recordSheet(t.s); t.s.filter = { range: R(0, 0, 5, 1), columns: { 0: { values: ['Apple', 'Plum'], blanks: false }, 1: { custom: [{ op: 'ge', value: '10' }] } } }; applyFilters(t.s); });
    expect([...t.s.hiddenRows].sort()).toEqual([1, 2, 4, 5]);
    t.b.transact('f', () => { t.s.filter!.columns = { 0: { values: ['Apple'] } }; applyFilters(t.s); });
    expect([...t.s.hiddenRows].sort()).toEqual([2, 4, 5]);
    // SUBTOTAL sees the hidden rows
    t.set('D1', '=SUBTOTAL(109,B2:B6)');
    expect(t.get('D1')).toBe(25);
    t.b.undo(); t.b.undo();
    expect(t.s.hiddenRows.size).toBe(4);
    // Table filter
    const u = book();
    u.rows('A1', [['N', 'V'], ['a', 1], ['b', 2], ['c', 3]]);
    u.b.transact('t', () => { u.b.recordSheet(u.s); u.s.tables.push({ id: 't1', name: 'T1', range: R(0, 0, 3, 1), columns: [{ name: 'N' }, { name: 'V' }], style: 'x', opts: { header: true, totals: false, banded: true, bandedCols: false, firstCol: false, lastCol: false, filterButton: true }, filter: { columns: { 1: { custom: [{ op: 'gt', value: '1' }] } } } }); applyFilters(u.s); });
    expect([...u.s.hiddenRows]).toEqual([1]);
  });

  it('lists distinct values under the other columns\' filters and extends ranges', () => {
    const t = book();
    t.rows('A1', [['F', 'C'], ['b', 'x'], ['a', 'y'], ['b', 'y'], [null, 'y'], [10, 'x'], [9, 'x']]);
    const area = { r1: 1, r2: 6, c1: 0, c2: 1, columns: { 1: { values: ['x'] } } };
    const d = distinctColumnValues(t.s, area, 0);
    expect(d.values).toEqual(['9', '10', 'b']);
    expect(d.blanks).toBe(false);
    const all = distinctColumnValues(t.s, { ...area, columns: {} }, 0);
    expect(all.values).toEqual(['9', '10', 'a', 'b']);
    expect(all.blanks).toBe(true);
    t.rows('A8', [['z', 'x']]);
    expect(extendFilterRange(t.s, R(0, 0, 6, 1)).r2).toBe(7);
    expect(pickListValues(t.s, 4, 1)).toEqual(['C', 'x', 'y']);
  });

  it('evaluates Advanced Filter criteria ranges (AND across, OR down, computed criteria)', () => {
    const t = book();
    t.rows('A1', [['Name', 'Region', 'Sales'], ['Ann', 'East', 500], ['Bob', 'West', 900], ['Cal', 'East', 1200], ['Dee', 'North', 300], ['Eve', 'Westfield', 50]]);
    t.rows('F1', [['Region', 'Sales'], ['East', '>1000'], ['West', null]]);
    const test = advancedCriteria(t.s, R(0, 0, 5, 2), R(0, 5, 2, 6), t.s, (f, r, c, br, bc) => t.e.evaluateScalar(f, t.s, r, c, { r: br, c: bc }));
    expect([1, 2, 3, 4, 5].filter(test)).toEqual([2, 3, 5]);
    t.rows('H1', [['Above avg'], ['=C2>AVERAGE($C$2:$C$6)']]);
    const test2 = advancedCriteria(t.s, R(0, 0, 5, 2), R(0, 7, 1, 7), t.s, (f, r, c, br, bc) => t.e.evaluateScalar(f, t.s, r, c, { r: br, c: bc }));
    expect([1, 2, 3, 4, 5].filter(test2)).toEqual([2, 3]);
    t.rows('J1', [['Name'], ['="=Bob"']]);
    const test3 = advancedCriteria(t.s, R(0, 0, 5, 2), R(0, 9, 1, 9), t.s, (f, r, c, br, bc) => t.e.evaluateScalar(f, t.s, r, c, { r: br, c: bc }));
    expect([1, 2, 3, 4, 5].filter(test3)).toEqual([2]);
  });
});

describe('text to columns', () => {
  it('splits delimited text with qualifiers and consecutive delimiters', () => {
    expect(splitLine('a,b,,c', { mode: 'delimited', delimiters: [','] })).toEqual(['a', 'b', '', 'c']);
    expect(splitLine('a  b   c', { mode: 'delimited', delimiters: [' '], consecutive: true })).toEqual(['a', 'b', 'c']);
    expect(splitLine('"Smith, John",42,"say ""hi"""', { mode: 'delimited', delimiters: [','], qualifier: '"' })).toEqual(['Smith, John', '42', 'say "hi"']);
    expect(splitLine('a;b\tc', { mode: 'delimited', delimiters: [';', '\t'] })).toEqual(['a', 'b', 'c']);
    expect(splitLine("'x,y',z", { mode: 'delimited', delimiters: [','], qualifier: "'" })).toEqual(['x,y', 'z']);
  });
  it('splits fixed width text and suggests breaks', () => {
    expect(splitLine('ABC  123  xy', { mode: 'fixed', breaks: [5, 10] })).toEqual(['ABC  ', '123  ', 'xy']);
    expect(suggestBreaks(['Apple   10  red', 'Kiwi    7   green', 'Fig     123 blue'])).toEqual([8, 12]);
  });
  it('converts fields by column format', () => {
    expect(convertSplitField(' 42 ', 'general').v).toBe(42);
    expect(convertSplitField('007', 'text')).toEqual({ v: '007', fmt: '@' });
    expect(convertSplitField('3/4/2024', 'DMY').v).toBe(partsToSerial(2024, 4, 3));
    expect(convertSplitField('2024-03-04', 'YMD').v).toBe(partsToSerial(2024, 3, 4));
    expect(convertSplitField('Mar 4 24', 'MDY').v).toBe(partsToSerial(2024, 3, 4));
    expect(convertSplitField('TRUE', 'general').v).toBe(true);
    expect(convertSplitField('=1+1', 'general').v).toBe('=1+1');
  });
});

describe('flash fill', () => {
  const run = (ex: [string[], string][], probe: string[]) => { const p = inferFlashFill(ex.map(([src, out]) => ({ src, out }))); return p ? runFlashFill(p, probe) : null; };
  it('extracts tokens and fields', () => {
    expect(run([[['John Smith'], 'John']], ['Mary Jones'])).toBe('Mary');
    expect(run([[['John Smith'], 'Smith']], ['Mary Jones'])).toBe('Jones');
    expect(run([[['a.b@foo.com'], 'foo.com'], [['x@bar.org'], 'bar.org']], ['q@zed.net'])).toBe('zed.net');
    expect(run([[['2024-03-15'], '15']], ['1999-12-31'])).toBe('31');
    expect(run([[['Order #123 (open)'], '123']], ['Order #98765 (closed)'])).toBe('98765');
  });
  it('combines columns with literals and changes case', () => {
    expect(run([[['John', 'Smith'], 'Smith, John']], ['Mary', 'Jones'])).toBe('Jones, Mary');
    expect(run([[['John', 'Smith'], 'john.smith@example.com']], ['Mary', 'Jones'])).toBe('mary.jones@example.com');
    expect(run([[['hello world'], 'HELLO WORLD']], ['abc def'])).toBe('ABC DEF');
    expect(run([[['mary ann'], 'Mary Ann']], ['joe bob'])).toBe('Joe Bob');
  });
  it('builds initials and uses every example', () => {
    expect(run([[['john smith'], 'J.S.'], [['ann lee'], 'A.L.']], ['bob ray'])).toBe('B.R.');
    expect(run([[['John Fitzgerald Kennedy'], 'JFK']], ['Lyndon Baines Johnson'])).toBe('LBJ');
    // Two examples disambiguate "first token" vs "last token"
    expect(run([[['Ann B Cole'], 'Cole'], [['Dan Fry'], 'Fry']], ['Eve K L Moss'])).toBe('Moss');
  });
  it('returns null when no program explains the examples', () => {
    expect(inferFlashFill([{ src: ['abc'], out: 'abc!' }, { src: ['xyz'], out: 'xyz!' }])).toBeTruthy();
    expect(inferFlashFill([{ src: ['abc'], out: 'abc' }, { src: ['xyz'], out: 'qqq' }])).toBeNull();
    expect(inferFlashFill([{ src: ['abc'], out: 'zzz' }, { src: ['xyz'], out: 'qqq' }])).toBeNull();
  });
});

describe('remove duplicates', () => {
  it('compares displayed values case-insensitively and shifts cells up within the range', () => {
    const t = book();
    t.rows('A1', [['K', 'V', 'Other'], ['a', 1, 'x1'], ['A', 1, 'x2'], ['b', 2, 'x3'], ['a', 2, 'x4'], ['b', 2, 'x5']]);
    const d1 = duplicateRows(t.s, R(0, 0, 5, 1), [0, 1], true);
    expect(d1).toEqual([2, 5]);
    expect(duplicateRows(t.s, R(0, 0, 5, 1), [0], true)).toEqual([2, 4, 5]);
    t.set('E1', 'outside');
    t.b.transact('rd', () => removeRowsInRange(t.s, R(0, 0, 5, 2), d1, true));
    expect(t.col(0, 0, 5)).toEqual(['K', 'a', 'b', 'a', null, null]);
    expect(t.col(2, 1, 3)).toEqual(['x1', 'x3', 'x4']);
    expect(t.get('E1')).toBe('outside');
  });
});

describe('subtotals and outline', () => {
  const data = [['Region', 'Sales'], ['East', 10], ['East', 20], ['West', 5], ['West', 7], ['West', 1]];
  it('inserts SUBTOTAL rows, a grand total and outline levels; Remove All undoes it', () => {
    const t = book();
    t.rows('A1', data);
    let out!: Range;
    t.b.transact('Subtotal', () => { out = addSubtotals(t.s, R(0, 0, 5, 1), { groupCol: 0, fn: 9, cols: [1], replace: true, pageBreaks: false, below: true }, (at, n) => shiftCells(t.s, 'row', at, n)); });
    expect(out.r2).toBe(8);
    expect(t.col(0, 0, 8)).toEqual(['Region', 'East', 'East', 'East Total', 'West', 'West', 'West', 'West Total', 'Grand Total']);
    expect(t.f('B4')).toBe('SUBTOTAL(9,B2:B3)');
    expect(t.f('B8')).toBe('SUBTOTAL(9,B5:B7)');
    expect(t.f('B9')).toBe('SUBTOTAL(9,B2:B8)');
    expect([t.get('B4'), t.get('B8'), t.get('B9')]).toEqual([30, 13, 43]);
    expect([1, 2, 3, 4, 5, 6, 7, 8].map((r) => t.s.rowLevel.get(r) || 0)).toEqual([2, 2, 1, 2, 2, 2, 1, 0]);
    t.b.transact('l', () => { t.b.recordSheet(t.s); showOutlineLevel(t.s, 'row', 2); });
    expect([...t.s.hiddenRows].sort((a, b) => a - b)).toEqual([1, 2, 4, 5, 6]);
    expect(t.get('B9')).toBe(43);
    t.b.transact('l', () => { t.b.recordSheet(t.s); showOutlineLevel(t.s, 'row', 3); });
    expect(t.s.hiddenRows.size).toBe(0);
    let back!: Range;
    t.b.transact('Remove', () => { back = removeSubtotals(t.s, out, (at, n) => shiftCells(t.s, 'row', at, n)); });
    expect(back.r2).toBe(5);
    expect(t.col(1, 1, 5)).toEqual([10, 20, 5, 7, 1]);
    expect(t.s.rowLevel.size).toBe(0);
  });

  it('summary above data and averages', () => {
    const t = book();
    t.rows('A1', data);
    t.b.transact('Subtotal', () => { addSubtotals(t.s, R(0, 0, 5, 1), { groupCol: 0, fn: 1, cols: [1], replace: true, pageBreaks: false, below: false }, (at, n) => shiftCells(t.s, 'row', at, n)); });
    expect(t.col(0, 0, 3)).toEqual(['Region', 'Grand Average', 'East Average', 'East']);
    expect(t.get('B3')).toBe(15);
    expect(t.get('B2')).toBeCloseTo(43 / 5);
  });

  it('groups, hides and shows detail, auto outlines', () => {
    const t = book();
    t.b.transact('g', () => { t.b.recordSheet(t.s); changeLevel(t.s, 'row', 1, 3, 1); changeLevel(t.s, 'row', 2, 3, 1); });
    expect([1, 2, 3].map((r) => t.s.rowLevel.get(r))).toEqual([1, 2, 2]);
    t.b.transact('h', () => { t.b.recordSheet(t.s); setDetail(t.s, 'row', 4, false); });
    expect([...t.s.hiddenRows].sort()).toEqual([1, 2, 3]);
    t.b.transact('s', () => { t.b.recordSheet(t.s); setDetail(t.s, 'row', 4, true); });
    expect(t.s.hiddenRows.size).toBe(0);
    t.b.transact('u', () => { t.b.recordSheet(t.s); changeLevel(t.s, 'row', 1, 3, -1); });
    expect([1, 2, 3].map((r) => t.s.rowLevel.get(r) || 0)).toEqual([0, 1, 1]);
    const u = book();
    u.rows('A1', [[1], [2], ['=SUM(A1:A2)'], [3], [4], ['=SUM(A4:A5)'], ['=SUM(A1:A6)']]);
    u.b.transact('a', () => { u.b.recordSheet(u.s); autoOutline(u.s, R(0, 0, 6, 0)); });
    expect([0, 1, 2, 3, 4, 5, 6].map((r) => u.s.rowLevel.get(r) || 0)).toEqual([2, 2, 1, 2, 2, 1, 0]);
  });
});

describe('consolidate and transform', () => {
  it('consolidates by labels and by position', () => {
    const a = { values: [[null, 'Q1', 'Q2'], ['East', 1, 2], ['West', 3, 4]] as Scalar[][] };
    const b = { values: [[null, 'Q2', 'Q3'], ['West', 10, 20], ['North', 5, 6]] as Scalar[][] };
    const { out } = consolidate([a, b], 'Sum', true, true);
    expect(out).toEqual([[null, 'Q1', 'Q2', 'Q3'], ['East', 1, 2, null], ['West', 3, 14, 20], ['North', null, 5, 6]]);
    const pos = consolidate([{ values: [[1, 2], [3, 4]] }, { values: [[10, 20], [30, 40], [50, 60]] }], 'Average', false, false);
    expect(pos.out).toEqual([[5.5, 11], [16.5, 22], [50, 60]]);
    const withRefs = consolidate([{ values: [[1]], ref: () => 'S1!$A$1' }, { values: [[2]], ref: () => 'S2!$A$1' }], 'Sum', false, false);
    expect(withRefs.refs[0][0]).toEqual(['S1!$A$1', 'S2!$A$1']);
  });

  it('applies transform steps', () => {
    const data = { headers: ['Name', 'Jan', 'Feb'], rows: [['a', 1, 2], ['b', 3, null], ['a', 5, 6]] as Scalar[][] };
    const up = applySteps(data, [{ kind: 'unpivot', keep: ['Name'], attr: 'Month', value: 'Amt' }]);
    expect(up.headers).toEqual(['Name', 'Month', 'Amt']);
    expect(up.rows.length).toBe(5);
    const g = applySteps(data, [{ kind: 'groupBy', keys: ['Name'], col: 'Jan', fn: 'Sum' }, { kind: 'sort', col: 'Sum of Jan', desc: true }]);
    expect(g.rows).toEqual([['a', 6], ['b', 3]]);
    const s = applySteps({ headers: ['Full'], rows: [['x-1'], ['y-22']] }, [{ kind: 'split', col: 'Full', delimiter: '-' }, { kind: 'rename', col: 'Full.2', to: 'N' }, { kind: 'filter', col: 'N', op: 'gt', value: '5' }]);
    expect(s).toEqual({ headers: ['Full.1', 'N'], rows: [['y', 22]] });
    expect(applySteps(data, [{ kind: 'removeCols', cols: ['Feb'] }, { kind: 'removeDuplicates' }]).rows.length).toBe(3);
    expect(() => applySteps(data, [{ kind: 'sort', col: 'Nope', desc: false }])).toThrow(/wasn't found/);
  });

  it('subtracts ranges and stores validation formulas', () => {
    expect(subtractRange(R(0, 0, 4, 4), R(1, 1, 2, 2))).toEqual([R(0, 0, 0, 4), R(3, 0, 4, 4), R(1, 0, 2, 0), R(1, 3, 2, 4)]);
    expect(subtractRange(R(0, 0, 1, 1), R(5, 5, 6, 6))).toEqual([R(0, 0, 1, 1)]);
    expect(dvStore('Yes, No', 'list')).toBe('"Yes, No"');
    expect(dvStore('=$A$1:$A$3', 'list')).toBe('$A$1:$A$3');
    expect(dvStore('1/2/2024', 'date')).toBe(String(partsToSerial(2024, 1, 2)));
    expect(dvStore('', 'whole')).toBeUndefined();
  });
});

describe('what-if', () => {
  it('goal seek finds roots (pure and on a sheet)', () => {
    const r = goalSeek((x) => x * x, 9, 1);
    expect(r.ok).toBe(true);
    expect(r.x).toBeCloseTo(3, 6);
    const r2 = goalSeek((x) => Math.exp(x), 100, 0);
    expect(r2.x).toBeCloseTo(Math.log(100), 6);
    const r3 = goalSeek(() => 5, 9, 1);
    expect(r3.ok).toBe(false);
    const t = book();
    t.rows('A1', [[1000], [0.05], ['=A1*(1+A2)^10']]);
    const sub = new Substituter(t.b, t.e);
    const res = goalSeek((x) => { sub.set({ sheet: t.s, r: 0, c: 0 }, x); sub.recalc(); return t.s.value(2, 0) as number; }, 2000, 1000);
    sub.restore();
    expect(res.x).toBeCloseTo(2000 / Math.pow(1.05, 10), 4);
    expect(t.get('A1')).toBe(1000);
    expect(t.get('A3')).toBeCloseTo(1000 * Math.pow(1.05, 10), 8);
  });

  it('computes one- and two-variable data tables without changing the model', () => {
    const t = book();
    t.rows('A1', [[2], [3], ['=A1*A2']]);
    // one variable, column oriented: inputs in C2:C4, formula in D1
    t.rows('C1', [[null, '=A3', '=A3+1'], [1], [5], [10]]);
    const sub = new Substituter(t.b, t.e);
    expect(computeDataTable(sub, t.s, R(0, 2, 3, 4), null, { sheet: t.s, r: 0, c: 0 })).toEqual([[3, 4], [15, 16], [30, 31]]);
    expect(t.get('A3')).toBe(6);
    // two variables: formula at G1, row values H1:I1 -> A2, column values G2:G3 -> A1
    t.rows('G1', [['=A3', 10, 20], [1], [2]]);
    expect(computeDataTable(sub, t.s, R(0, 6, 2, 8), { sheet: t.s, r: 1, c: 0 }, { sheet: t.s, r: 0, c: 0 })).toEqual([[10, 20], [20, 40]]);
    // row oriented
    t.rows('K1', [[null, 1, 2], ['=A3']]);
    expect(computeDataTable(sub, t.s, R(0, 10, 1, 12), { sheet: t.s, r: 1, c: 0 }, null)).toEqual([[2, 4]]);
    expect(t.get('A3')).toBe(6);
  });

  it('scenario results substitute each scenario in turn', () => {
    const t = book();
    t.rows('A1', [[1], [2], ['=A1+A2']]);
    const scs = [{ name: 'Best', cells: '$A$1:$A$2', values: [10, 20] }, { name: 'Worst', cells: '$A$1,$A$2', values: [0, 1] }];
    const res = scenarioResults(new Substituter(t.b, t.e), t.b, t.s, scs, [{ sheet: t.s, r: 2, c: 0 }]);
    expect(res).toEqual({ current: [3], per: [[30], [1]] });
    expect(t.get('A3')).toBe(3);
    expect(parseCells(t.b, 'A1:B2, C3', t.s)!.length).toBe(5);
  });

  it('creates a forecast sheet with FORECAST.ETS formulas and a chart', () => {
    const t = book();
    let sh!: Sheet;
    t.b.transact('fc', () => { sh = createForecastSheet(t.b, [1, 2, 3, 4, 5, 6, 7, 8], [10, 12, 14, 16, 18, 20, 22, 24], { periods: 3, confidence: 0.95, seasonality: 'none' }); });
    expect(sh.value(0, 2)).toBe('Forecast');
    expect(sh.get(8, 2)?.f).toBe('B9');
    expect(sh.get(9, 2)?.f).toBe('FORECAST.ETS(A10,$B$2:$B$9,$A$2:$A$9,0,1)');
    expect(sh.value(11, 0)).toBe(11);
    expect(sh.charts.length).toBe(1);
    expect(sh.charts[0].series[1].values).toMatch(/\$C\$2:\$C\$12$/);
    t.e.recalc();
    expect(sh.value(9, 2) as number).toBeCloseTo(26, 0);
    expect((sh.value(9, 4) as number) >= (sh.value(9, 3) as number)).toBe(true);
  });
});

describe('solver', () => {
  it('simplex solves a small LP', () => {
    const r = simplex([3, 5], [[1, 0], [0, 2], [3, 2]], ['<=', '<=', '<='], [4, 12, 18]);
    expect(r.status).toBe('optimal');
    expect(r.x[0]).toBeCloseTo(2); expect(r.x[1]).toBeCloseTo(6); expect(r.value).toBeCloseTo(36);
    const inf = simplex([1, 1], [[1, 1], [1, 1]], ['<=', '>='], [1, 2]);
    expect(inf.status).toBe('infeasible');
    expect(simplex([1, 0], [[0, 1]], ['<='], [1]).status).toBe('unbounded');
    // >= and = rows (phase 1): min x + y st x + 2y >= 4, x - y = 1  -> x=2, y=1
    const m = simplex([-1, -1], [[1, 2], [1, -1]], ['>=', '='], [4, 1]);
    expect(m.x[0]).toBeCloseTo(2); expect(m.x[1]).toBeCloseTo(1);
  });

  it('branch and bound handles integer variables and free variables', () => {
    // max 5x + 4y st 6x + 4y <= 24, x + 2y <= 6, x,y int -> (4,0)=20 or (3,1)=19 -> 20
    const r = solveLP([5, 4], [[6, 4], [1, 2]], ['<=', '<='], [24, 6], { nonNeg: [true, true], ints: [0, 1] });
    expect(r.value).toBeCloseTo(20); expect(r.x).toEqual([4, 0]);
    // free variable: min x st x >= -5
    const f = solveLP([-1], [[1]], ['>='], [-5], { nonNeg: [false] });
    expect(f.x[0]).toBeCloseTo(-5);
  });

  it('detects linearity and solves a nonlinear problem', () => {
    expect(linearize((x) => 2 * x[0] - x[1] + 3, [1, 1])).toBeTruthy();
    expect(linearize((x) => x[0] * x[1], [1, 1])).toBeNull();
    const nl = solveProblem({ x0: [0, 0], f: (x) => (x[0] - 3) ** 2 + (x[1] + 1) ** 2 + 2, goal: 'min', cons: [], nonNeg: false });
    expect(nl.linear).toBe(false);
    expect(nl.x[0]).toBeCloseTo(3, 3); expect(nl.x[1]).toBeCloseTo(-1, 3);
    const nn = solveProblem({ x0: [1, 1], f: (x) => (x[0] - 3) ** 2 + (x[1] + 1) ** 2, goal: 'min', cons: [], nonNeg: true });
    expect(nn.x[1]).toBeCloseTo(0, 3);
    // max x*y subject to x + y <= 10 -> (5, 5)
    const c = solveProblem({ x0: [1, 1], f: (x) => x[0] * x[1], goal: 'max', cons: [{ g: (x) => x[0] + x[1] - 10, rel: '<=' }], nonNeg: true });
    expect(c.status).toBe('optimal');
    expect(c.x[0]).toBeCloseTo(5, 2); expect(c.value).toBeCloseTo(25, 2);
    const nm = nelderMead((x) => (x[0] - 1) ** 2 + 100 * (x[1] - x[0] ** 2) ** 2, [-1, 1], { maxIter: 5000 });
    expect(nm.x[0]).toBeCloseTo(1, 2);
    // value-of goal
    const v = solveProblem({ x0: [1], f: (x) => 2 * x[0] + 1, goal: 'value', target: 7, cons: [], nonNeg: true });
    expect(v.x[0]).toBeCloseTo(3);
  });

  it('builds a sheet model, stores it in hidden names and solves it', () => {
    const t = book();
    // x in B1, y in B2; objective B3 = 3x + 5y; constraints B4 = x <= 4, B5 = 2y <= 12, B6 = 3x + 2y <= 18
    t.rows('B1', [[0, 4], [0, 12], ['=3*B1+5*B2'], ['=B1'], ['=2*B2'], ['=3*B1+2*B2']]);
    const model = { objective: '$B$3', goal: 'max' as const, target: 0, vars: '$B$1:$B$2', cons: [{ lhs: '$B$4:$B$5', rel: '<=' as const, rhs: '$C$1:$C$2' }, { lhs: '$B$6', rel: '<=' as const, rhs: '18' }, { lhs: '$B$1', rel: 'int' as const, rhs: 'integer' }], nonNeg: true, method: 'auto' as const };
    t.b.transact('m', () => saveSolverModel(t.b, t.s, model));
    const loaded = loadSolverModel(t.b, t.s);
    expect(loaded.vars).toBe('$B$1:$B$2');
    expect(loaded.cons.length).toBe(3);
    expect(loaded.cons[2].rel).toBe('int');
    expect(loaded.method).toBe('auto');
    const built = buildSheetProblem(t.b, t.e, t.s, loaded);
    if (typeof built === 'string') throw new Error(built);
    const r = solveProblem(built.prob);
    built.sub.restore();
    expect(r.linear).toBe(true);
    expect(r.x[0]).toBeCloseTo(2); expect(r.x[1]).toBeCloseTo(6); expect(r.value).toBeCloseTo(36);
    expect(t.get('B3')).toBe(0);
  });
});

describe('analysis toolpak', () => {
  const xs = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  const evalF = (f: string) => { const t = book(); t.set('Z1', '=' + f); return t.get('Z1') as number; };
  it('descriptive statistics', () => {
    const out = descriptiveStatistics([{ label: 'X', data: [2, 4, 4, 4, 5, 5, 7, 9] }], { summary: true, conf: 0.95, kLarge: 2, kSmall: 1 });
    const get = (label: string) => out.m.find((r) => r[0] === label)?.[1];
    expect(out.m[0][0]).toBe('X');
    expect(get('Mean')).toBe(5);
    expect(get('Median')).toBe(4.5);
    expect(get('Mode')).toBe(4);
    expect(get('Sample Variance')).toBeCloseTo(32 / 7, 12);
    expect(get('Standard Deviation')).toBeCloseTo(Math.sqrt(32 / 7), 12);
    expect(get('Standard Error')).toBeCloseTo(Math.sqrt(32 / 7) / Math.sqrt(8), 12);
    expect(get('Kurtosis')).toBeCloseTo(evalF('KURT(2,4,4,4,5,5,7,9)'), 10);
    expect(get('Skewness')).toBeCloseTo(evalF('SKEW(2,4,4,4,5,5,7,9)'), 10);
    expect(get('Range')).toBe(7); expect(get('Minimum')).toBe(2); expect(get('Maximum')).toBe(9); expect(get('Sum')).toBe(40); expect(get('Count')).toBe(8);
    expect(get('Largest(2)')).toBe(7); expect(get('Smallest(1)')).toBe(2);
    expect(get('Confidence Level(95.0%)')).toBeCloseTo(evalF('CONFIDENCE.T(0.05,STDEV.S(2,4,4,4,5,5,7,9),8)'), 10);
    const ten = descriptiveStatistics([{ label: 'A', data: xs }, { label: 'B', data: [1, 1, 2] }]);
    expect(ten.m[2][1]).toBe(5.5);
    expect(ten.m.find((r) => r[0] === 'Kurtosis')?.[1]).toBeCloseTo(-1.2, 12);
    expect((ten.m.find((r) => r[0] === 'Mode')?.[1] as any).code).toBe('#N/A');
    expect(ten.m[2][2]).toBe('Mean'); expect(ten.m[2][3]).toBeCloseTo(4 / 3);
  });

  it('histogram, correlation and covariance', () => {
    const h = histogram([1, 2, 2, 3, 5, 8, 9], [2, 5], { cumulative: true, pareto: true });
    expect(h.m.slice(1).map((r) => r.slice(0, 3))).toEqual([[2, 3, 3 / 7], [5, 2, 5 / 7], ['More', 2, 1]]);
    expect(h.m[1].slice(4)).toEqual([2, 3, 3 / 7]);
    const auto = histogram([1, 2, 3, 4, 5, 6, 7, 8, 9], null);
    expect(auto.m.slice(1).reduce((s, r) => s + (r[1] as number), 0)).toBe(9);
    const c = correlationMatrix([{ label: 'a', data: [1, 2, 3, 4] }, { label: 'b', data: [2, 4, 5, 9] }]);
    expect(c.m[2][1]).toBeCloseTo(evalF('CORREL({1,2,3,4},{2,4,5,9})'), 12);
    expect(c.m[1][2]).toBeNull();
    const v = correlationMatrix([{ label: 'a', data: [1, 2, 3, 4] }, { label: 'b', data: [2, 4, 5, 9] }], true);
    expect(v.m[2][1]).toBeCloseTo(evalF('COVARIANCE.P({1,2,3,4},{2,4,5,9})'), 12);
    expect(v.m[1][1]).toBeCloseTo(1.25, 12);
  });

  it('regression matches LINEST', () => {
    const y = [3.1, 4.9, 7.2, 8.8, 11.1, 13.0];
    const x1 = [1, 2, 3, 4, 5, 6], x2 = [2, 1, 4, 3, 6, 5];
    const out = regression({ label: 'Y', data: y }, [{ label: 'X1', data: x1 }, { label: 'X2', data: x2 }], { residuals: true, stdResiduals: true, conf: 0.9 });
    const row = (l: string) => out.m.find((r) => r[0] === l)!;
    const t = book();
    t.rows('A1', y.map((v, i) => [v, x1[i], x2[i]]));
    t.set('E1', '=LINEST(A1:A6,B1:C6,TRUE,TRUE)');
    t.e.recalc();
    // LINEST: row 1 = coefs (x2, x1, b), row 2 = SEs, row 3 = r2, sey; row 4 = F, df; row 5 = ssreg, ssresid
    expect(row('Intercept')[1]).toBeCloseTo(t.get('G1') as number, 9);
    expect(row('X1')[1]).toBeCloseTo(t.get('F1') as number, 9);
    expect(row('X2')[1]).toBeCloseTo(t.get('E1') as number, 9);
    expect(row('X1')[2]).toBeCloseTo(t.get('F2') as number, 9);
    expect(row('R Square')[1]).toBeCloseTo(t.get('E3') as number, 10);
    expect(row('Standard Error')[1]).toBeCloseTo(t.get('F3') as number, 10);
    expect(row('Regression')[4]).toBeCloseTo(t.get('E4') as number, 6);
    expect(row('Residual')[1]).toBe(3);
    expect(row('Residual')[2]).toBeCloseTo(t.get('F5') as number, 10);
    expect(row('Observations')[1]).toBe(6);
    const pv = row('X1')[4] as number;
    expect(pv).toBeCloseTo(evalF(`T.DIST.2T(${Math.abs((row('X1')[3] as number))},3)`), 10);
    expect(out.m.some((r) => r[0] === 'RESIDUAL OUTPUT')).toBe(true);
    expect(out.m[15].slice(7)).toEqual(['Lower 90.0%', 'Upper 90.0%']);
    expect(() => regression({ label: 'Y', data: [1, 2, 3] }, [{ label: 'a', data: [1, 2, 3] }, { label: 'b', data: [2, 4, 6] }])).toThrow(/collinear/);
  });

  it('t-tests, z-test and F-test match worksheet functions', () => {
    const a = { label: 'A', data: [12, 15, 11, 18, 14, 16] }, b = { label: 'B', data: [10, 14, 9, 15, 12, 11] };
    const A = '{12,15,11,18,14,16}', B = '{10,14,9,15,12,11}';
    const val = (o: { m: Scalar[][] }, l: string) => o.m.find((r) => r[0] === l)?.[1] as number;
    const p = tTest(a, b, 'paired');
    expect(val(p, 'P(T<=t) two-tail')).toBeCloseTo(evalF(`T.TEST(${A},${B},2,1)`), 10);
    expect(val(p, 'df')).toBe(5);
    expect(val(p, 't Critical two-tail')).toBeCloseTo(evalF('T.INV.2T(0.05,5)'), 9);
    const eq = tTest(a, b, 'equal');
    expect(val(eq, 'P(T<=t) two-tail')).toBeCloseTo(evalF(`T.TEST(${A},${B},2,2)`), 10);
    expect(val(eq, 'P(T<=t) one-tail')).toBeCloseTo(evalF(`T.TEST(${A},${B},1,2)`), 10);
    const un = tTest(a, { label: 'C', data: [1, 30, 2, 25] }, 'unequal');
    expect(val(un, 'df')).toBe(3);
    const z = zTest(a, b, 6, 4);
    expect(val(z, 'z')).toBeCloseTo((14.333333333333334 - 11.833333333333334) / Math.sqrt(6 / 6 + 4 / 6), 10);
    expect(val(z, 'z Critical two-tail')).toBeCloseTo(1.959963984540054, 9);
    const f = fTest(a, b);
    expect(val(f, 'F')).toBeCloseTo(evalF(`VAR.S(${A})/VAR.S(${B})`), 12);
    expect(val(f, 'P(F<=f) one-tail') * 2).toBeCloseTo(evalF(`F.TEST(${A},${B})`), 9);
  });

  it('ANOVA tables', () => {
    const g = [{ label: 'g1', data: [6, 8, 4, 5, 3, 4] }, { label: 'g2', data: [8, 12, 9, 11, 6, 8] }, { label: 'g3', data: [13, 9, 11, 8, 7, 12] }];
    const a = anovaSingle(g);
    const bg = a.m.find((r) => r[0] === 'Between Groups')!;
    expect(bg[1]).toBeCloseTo(84, 10); expect(bg[2]).toBe(2); expect(bg[4]).toBeCloseTo(9.264705882352942, 9);
    expect(bg[5]).toBeCloseTo(evalF('F.DIST.RT(9.264705882352942,2,15)'), 10);
    expect(bg[6]).toBeCloseTo(evalF('F.INV.RT(0.05,2,15)'), 6);
    expect(a.m.find((r) => r[0] === 'Within Groups')![1]).toBeCloseTo(68, 10);
    const two = anovaTwoNoRep([[1, 2, 3], [2, 4, 7], [3, 5, 8]], ['r1', 'r2', 'r3'], ['c1', 'c2', 'c3']);
    const rows = two.m.find((r) => r[0] === 'Rows')!, cols = two.m.find((r) => r[0] === 'Columns')!, err = two.m.find((r) => r[0] === 'Error')!;
    expect(rows[1]).toBeCloseTo(158 / 9, 9); expect(cols[1]).toBeCloseTo(218 / 9, 9); expect(err[1]).toBeCloseTo(28 / 9, 9);
    const rep = anovaTwoRep([[1, 2], [3, 4], [5, 7], [6, 9]], 2, ['s1', 's2'], ['A', 'B']);
    const total = rep.m.find((r) => r[0] === 'Total' && typeof r[1] === 'number' && r.length > 2 && r[2] === 7)!;
    expect(total[1]).toBeCloseTo(49.875, 10);
    const ss = (l: string) => rep.m.find((r) => r[0] === l)![1] as number;
    expect(ss('Sample') + ss('Columns') + ss('Interaction') + ss('Within')).toBeCloseTo(49.875, 10);
    expect(ss('Within')).toBeCloseTo(6.5, 10);
  });

  it('moving average, exponential smoothing, rank and percentile', () => {
    const m = movingAverage([1, 2, 3, 4, 5, 6], 3, true);
    expect(m.m.map((r) => r[0]).slice(2)).toEqual([2, 3, 4, 5]);
    expect((m.m[0][0] as any).code).toBe('#N/A');
    expect(m.m[4][1]).toBeCloseTo(1, 10);
    const e = exponentialSmoothing([10, 20, 30, 40], 0.3);
    expect(e.m.map((r) => r[0]).slice(1)).toEqual([10, 17, 26.1]);
    const r = rankAndPercentile([{ label: 'S', data: [7, 3, 9, 3] }]);
    expect(r.m.slice(1)).toEqual([[3, 9, 1, 1], [1, 7, 2, 0.666], [2, 3, 3, 0], [4, 3, 3, 0]]);
  });

  it('random numbers, sampling and Fourier analysis', () => {
    const u = randomNumbers(2, 100, { kind: 'uniform', a: 5, b: 6 }, 42);
    expect(u.m.length).toBe(100);
    expect(u.m.flat().every((v) => (v as number) >= 5 && (v as number) < 6)).toBe(true);
    expect(randomNumbers(1, 3, { kind: 'normal', mean: 0, sd: 1 }, 7).m).toEqual(randomNumbers(1, 3, { kind: 'normal', mean: 0, sd: 1 }, 7).m);
    expect(randomNumbers(1, 0, { kind: 'patterned', from: 1, to: 3, step: 1, repeatEach: 2, repeatSeq: 1 }).m.flat()).toEqual([1, 1, 2, 2, 3, 3]);
    const bern = randomNumbers(1, 200, { kind: 'bernoulli', p: 0.5 }, 3).m.flat();
    expect(bern.every((v) => v === 0 || v === 1)).toBe(true);
    expect(randomNumbers(1, 50, { kind: 'discrete', values: [1, 2], probs: [0, 1] }, 9).m.flat().every((v) => v === 2)).toBe(true);
    expect(sampling([1, 2, 3, 4, 5, 6, 7], 'periodic', 3).m).toEqual([[3], [6]]);
    expect(sampling([4, 4, 4], 'random', 5, 1).m).toEqual([[4], [4], [4], [4], [4]]);
    const f = fourier([1, 2, 3, 4]);
    expect(f.m.map((r) => r[0])).toEqual(['10', '-2+2i', '-2', '-2-2i']);
    expect(fourier(f.m.map((r) => r[0]), true).m.map((r) => r[0])).toEqual(['1', '2', '3', '4']);
    expect(() => fourier([1, 2, 3])).toThrow(/2, 4, 8/);
    expect(parseComplex('3-4i')).toEqual([3, -4]);
    expect(parseComplex('-i')).toEqual([0, -1]);
    expect(parseComplex('2.5j')).toEqual([0, 2.5]);
    expect(parseComplex('abc')).toBeNull();
    expect(formatComplex([0, 1])).toBe('i');
    expect(formatComplex([1.5, -1])).toBe('1.5-i');
    expect(fft([[1, 0], [0, 0]]).map((c) => c[0])).toEqual([1, 1]);
    expect(invertMatrix([[2, 0], [0, 4]])).toEqual([[0.5, 0], [0, 0.25]]);
  });
});
