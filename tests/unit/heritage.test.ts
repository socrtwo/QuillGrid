import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import '../../src/commands';
import { Workbook, Sheet } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell, parseRange } from '../../src/model/address';
import { defaultPrefs } from '../../src/prefs';
import { getCommand } from '../../src/commands/registry';
import { getFn } from '../../src/calc/functions/registry';
import { CellError } from '../../src/model/types';
import { parseInput, setInput } from '../../src/commands/ops';
import { validates } from '../../src/commands/edit';
import { translateLotus, lotusEntry, lotusLabel, hasLotusSyntax, LOTUS_FUNCTIONS } from '../../src/heritage/lotus';
import { toR1C1, formulaRules, inconsistentFormulas, cellKind, explainTree, explainText, dependencyGraph, cellHistory, cellContentText, packRanges } from '../../src/heritage/audit';
import {
  matInverse, matMultiply, matSolve, matDeterminant, parseDist, sample, seeded, monteCarlo, summarize, leafInputs, sliderRange,
  matchesCondition, typeMismatches, distinctChoices, assignKeys,
} from '../../src/heritage/models';
import { sheetToCode } from '../../src/heritage/codegen';
import { easterSunday, isoWeeksInYear, refConvert, categorize } from '../../src/calc/functions/heritage';
import { convertLotusCells, deleteMatching, matrixOp, applyColumnType, runMonteCarlo, cellInfoRows, labelFor, columnTypeSpec } from '../../src/commands/heritage';
import { slashTree, commandItem, slashKeys } from '../../src/ui/slashMenu';

/** Minimal App stand-in for running commands on a Workbook. */
function fakeApp() {
  const book = new Workbook();
  const engine = new Engine(book);
  book.onChange = () => engine.recalc();
  const app: any = {
    book, engine, prefs: defaultPrefs(), recording: null,
    get sheet() { return book.sheet; },
    get tab() { return { book, engine, displayName: 'Test', mapView: false, cf: { invalidate() {} } }; },
    grid: { select: (r: number, c: number, o: any = {}) => { book.sheet.sel = { r, c, ranges: [o.range || { r1: r, c1: c, r2: r, c2: c }] }; }, scrollIntoView() {}, objects: { select() {} }, invalidate() {} },
    isReadOnly: () => false,
    exec: (id: string, arg?: any) => { if (id === 'sheet.activate') { book.activeSheet = arg; return true; } const c = getCommand(id); return c ? c.run(app, arg) : false; },
    refresh() {}, savePrefs() {}, ui: { status() {}, openPane() {}, rebuild() {}, paneOpen: () => false, togglePane() {} },
  };
  const put = (a: string, v: any, sh: Sheet = book.sheet) => { const p = parseCell(a)!; book.transact('t', () => sh.set(p.r, p.c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v })); };
  const get = (a: string, sh: Sheet = book.sheet) => { engine.recalc(); const p = parseCell(a)!; return sh.value(p.r, p.c); };
  const select = (ref: string) => { const rg = parseRange(ref)!; book.sheet.sel = { r: rg.r1, c: rg.c1, ranges: [rg] }; };
  let slot = 0;
  const ev = (f: string) => { const a = `AZ${500 + slot++}`; put(a, '=' + f); return get(a); };
  return { app, book, engine, put, get, select, ev };
}

const known = (n: string) => !!getFn(n);

describe('Lotus 1-2-3 formula translation', () => {
  it('translates ranges, operators and functions', () => {
    expect(translateLotus('@SUM(A1..A5)')).toBe('SUM(A1:A5)');
    expect(translateLotus('@SUM(A1.A5)')).toBe('SUM(A1:A5)');
    expect(translateLotus('+A1*@AVG(B1..B3)')).toBe('A1*AVERAGEA(B1:B3)');
    expect(translateLotus('@IF(A1>1#AND#B1<2#AND#C1=3,"y","n")')).toBe('IF(AND(A1>1,B1<2,C1=3),"y","n")');
    expect(translateLotus('#NOT#A1=1#OR#B1')).toBe('OR(NOT(A1=1),B1)');
    expect(translateLotus('@@("A1")')).toBe('INDIRECT("A1")');
    expect(translateLotus('@ISSTRING(A1)')).toBe('ISTEXT(A1)');
    expect(translateLotus('@ERR')).toBe('#VALUE!');
    expect(translateLotus('@TODAY')).toBe('TODAY()');
    expect(translateLotus('@PI*2')).toBe('PI()*2');
  });
  it('keeps 1-2-3 precedence', () => {
    expect(translateLotus('-2^2')).toBe('-(2^2)');
    expect(translateLotus('"a"&"b"="ab"')).toBe('"a"&("b"="ab")');
    expect(translateLotus('A1-(B1-C1)')).toBe('A1-(B1-C1)');
    expect(translateLotus('(A1+B1)*C1')).toBe('(A1+B1)*C1');
  });
  it('converts zero-based offsets and argument orders', () => {
    expect(translateLotus('@CHOOSE(A1,1,2,3)')).toBe('CHOOSE(A1+1,1,2,3)');
    expect(translateLotus('@CHOOSE(0,"x","y")')).toBe('CHOOSE(1,"x","y")');
    expect(translateLotus('@VLOOKUP(A1,$B$1..$D$9,2)')).toBe('VLOOKUP(A1,$B$1:$D$9,3,ISNUMBER(A1))');
    expect(translateLotus('@INDEX(R,1,2)')).toBe('INDEX(R,3,2)');
    expect(translateLotus('@MID(A1,0,3)')).toBe('MID(A1,1,3)');
    expect(translateLotus('@FIND("x",A1,0)*2')).toBe('(FIND("x",A1,1)-1)*2');
    expect(translateLotus('@PMT(1000,0.01,12)')).toBe('PMT(0.01,12,-1000)');
    expect(translateLotus('@IRR(0.1,A1..A5)')).toBe('IRR(A1:A5,0.1)');
    expect(translateLotus('@DSUM(DB,2,CRIT)')).toBe('DSUM(DB,3,CRIT)');
    expect(translateLotus('@YEAR(@TODAY)')).toBe('YEAR(TODAY())-1900');
  });
  it('maps sheet letters to sheets', () => {
    expect(translateLotus('A:A1..C:B2', { sheetNames: ['Sheet1', 'Sheet 2', 'Data'] })).toBe('Sheet1:Data!A1:B2');
    expect(translateLotus('+B:C3*2', { sheetNames: ['Sheet1', 'Sheet 2'] })).toBe("'Sheet 2'!C3*2");
  });
  it('rejects malformed formulas', () => {
    expect(() => translateLotus('@SUM(A1..A3')).toThrow();
    expect(() => translateLotus('@SUM')).toThrow();
  });
  it('decides which entries are 1-2-3 formulas', () => {
    expect(lotusEntry('@SUM(A1..A3)', known)).toBe('SUM(A1:A3)');
    expect(lotusEntry('+A1+@SUM(B1..B2)', known)).toBe('A1+SUM(B1:B2)');
    expect(lotusEntry('=@AVG(A1..A3)', known)).toBe('AVERAGEA(A1:A3)');
    expect(lotusEntry('=@INDEX(A1:B2,1,1)', known)).toBeNull(); // implicit intersection keeps its meaning
    expect(lotusEntry('=SUM(A1:A3)', known)).toBeNull();
    expect(lotusEntry('@home', known)).toBeNull();
    expect(lotusEntry('@john (work)', known)).toBeNull();
    expect(lotusEntry('-5', known)).toBeNull();
    expect(lotusEntry('hello', known)).toBeNull();
    expect(hasLotusSyntax('+A1..A3')).toBe(true);
    expect(LOTUS_FUNCTIONS).toContain('PUREAVG');
  });
  it('reads label prefixes', () => {
    expect(lotusLabel('"Total')).toEqual({ text: 'Total', align: 'right' });
    expect(lotusLabel('^Title')).toEqual({ text: 'Title', align: 'center' });
    expect(lotusLabel('\\-')).toEqual({ text: '-', align: 'fill' });
    expect(lotusLabel('"quoted"')).toBeNull();
    expect(lotusLabel('plain')).toBeNull();
  });
  it('computes with 1-2-3 meanings', () => {
    const { put, ev } = fakeApp();
    put('A1', 10); put('A2', 'label'); put('A3', 20);
    const t = (f: string) => ev(translateLotus(f));
    expect(t('@AVG(A1..A3)')).toBe(10); // labels count as zero
    expect(t('@PUREAVG(A1..A3)')).toBe(15);
    expect(t('@COUNT(A1..A3)')).toBe(3);
    expect(t('@PURECOUNT(A1..A3)')).toBe(2);
    expect(t('-2^2')).toBe(-4);
    expect(t('@CHOOSE(1,"a","b","c")')).toBe('b');
    expect(t('@MID("spreadsheet",6,5)')).toBe('sheet');
    expect(t('@FIND("s","lotus",0)')).toBe(4);
    expect(t('@YEAR(@DATE(99,12,31))')).toBe(99);
    expect(t('@WEEKDAY(@DATE(2024,1,1))')).toBe(0); // Monday
    expect(Math.round((t('@PMT(1000,0.01,12)') as number) * 100) / 100).toBe(88.85);
    put('C1', 'apple'); put('D1', 1); put('C2', 'banana'); put('D2', 2);
    expect(t('@VLOOKUP("banana",C1..D2,1)')).toBe(2);
    expect(t('@INDEX(C1..D2,1,0)')).toBe(1);
  });
  it('accepts 1-2-3 entry when typing, and label prefixes when turned on', () => {
    const { app, book, get, put } = fakeApp();
    put('A1', 2); put('A2', 3);
    const p = parseInput(app, '@SUM(A1..A2)', 2, 0);
    expect(p.cell.f).toBe('SUM(A1:A2)');
    book.transact('t', () => setInput(app, book.sheet, 2, 0, '@SUM(A1..A2)'));
    expect(get('A3')).toBe(5);
    expect(parseInput(app, '"Total', 0, 1).cell.v).toBe('"Total');
    app.prefs.lotusLabels = true;
    const q = parseInput(app, '"Total', 0, 1);
    expect(q.cell.v).toBe('Total');
    expect(q.align).toBe('right');
    book.transact('t', () => setInput(app, book.sheet, 0, 1, '^Mid'));
    expect(book.sheet.styleOf(0, 1).hAlign).toBe('center');
    app.prefs.lotusFormulas = false;
    expect(parseInput(app, '@SUM(A1..A2)', 2, 0).cell.v).toBe('@SUM(A1..A2)');
  });
  it('converts text cells holding 1-2-3 formulas', () => {
    const { app, put, get, select } = fakeApp();
    put('A1', 4); put('A2', 6); put('B1', '@SUM(A1..A2)'); put('B2', '+A1*2'); put('B3', '@notafunction stays');
    select('B1:B3');
    expect(convertLotusCells(app, [parseRange('B1:B3')!])).toBe(1);
    expect(get('B1')).toBe(10);
    expect(get('B2')).toBe('+A1*2'); // no 1-2-3 syntax: left alone
    expect(get('B3')).toBe('@notafunction stays');
  });
});

describe('heritage worksheet functions', () => {
  it('ISBETWEEN', () => {
    const { ev } = fakeApp();
    expect(ev('ISBETWEEN(5,1,10)')).toBe(true);
    expect(ev('ISBETWEEN(10,1,10,TRUE,FALSE)')).toBe(false);
    expect(ev('ISBETWEEN(5,10,1)')).toBe(true);
    expect(ev('ISBETWEEN("m","a","z")')).toBe(true);
    expect(ev('ISBETWEEN(1,1,10,FALSE)')).toBe(false);
  });
  it('WEIGHTAVG, SEMEAN, COUNTUNIQUE', () => {
    const { ev, put } = fakeApp();
    put('A1', 80); put('A2', 90); put('A3', 'x'); put('B1', 1); put('B2', 3); put('B3', 5);
    expect(ev('WEIGHTAVG(A1:A3,B1:B3)')).toBe(87.5);
    expect(ev('WEIGHTAVG(A1:A2,B1:B3)')).toBeInstanceOf(CellError);
    expect(ev('SEMEAN(2,4,4,4,5,5,7,9)')).toBeCloseTo(2.138089935 / Math.sqrt(8), 9);
    expect(ev('SEMEAN(1)')).toBeInstanceOf(CellError);
    put('C1', 'a'); put('C2', 'A'); put('C3', 1); put('C4', 1); put('C5', 'b');
    expect(ev('COUNTUNIQUE(C1:C6)')).toBe(3);
  });
  it('REFCONVERT and XINDEX', () => {
    const { ev, put } = fakeApp();
    expect(ev('REFCONVERT("AB")')).toBe(28);
    expect(ev('REFCONVERT(28)')).toBe('AB');
    expect(ev('REFCONVERT(0)')).toBeInstanceOf(CellError);
    expect(refConvert('xfd')).toBe(16384);
    put('A1', ''); put('B1', 'Q1'); put('C1', 'Q2'); put('A2', 'North'); put('B2', 10); put('C2', 20); put('A3', 'South'); put('B3', 30); put('C3', 40);
    expect(ev('XINDEX(A1:C3,"Q2","South")')).toBe(40);
    expect(ev('XINDEX(A1:C3,"Q3","South")')).toBeInstanceOf(CellError);
  });
  it('ISEMAIL, ISURL, EPOCHTODATE', () => {
    const { ev } = fakeApp();
    expect(ev('ISEMAIL("ann@example.com")')).toBe(true);
    expect(ev('ISEMAIL("ann@example")')).toBe(false);
    expect(ev('ISEMAIL(5)')).toBe(false);
    expect(ev('ISURL("https://example.com/a?b=1")')).toBe(true);
    expect(ev('ISURL("example.org")')).toBe(true);
    expect(ev('ISURL("not a url")')).toBe(false);
    expect(ev('EPOCHTODATE(0)')).toBe(25569);
    expect(ev('EPOCHTODATE(86400000,2)')).toBe(25570);
    expect(ev('EPOCHTODATE(1,9)')).toBeInstanceOf(CellError);
  });
  it('CATEGORIZE', () => {
    const { ev, put } = fakeApp();
    put('E1', 'uber, lyft'); put('F1', 'Transport'); put('E2', 'coffee*'); put('F2', 'Food'); put('E3', 'rent'); put('F3', 'Housing');
    expect(ev('CATEGORIZE("UBER trip 12",E1:E3,F1:F3)')).toBe('Transport');
    expect(ev('CATEGORIZE("Coffeehouse",E1:E3,F1:F3)')).toBe('Food');
    expect(ev('CATEGORIZE("Groceries",E1:E3,F1:F3,"Other")')).toBe('Other');
    expect(ev('CATEGORIZE("Groceries",E1:E3,F1:F3)')).toBeInstanceOf(CellError);
    expect(categorize('Monthly rent', [{ keywords: 'rent', category: 'Housing' }])).toBe('Housing');
  });
  it('date helpers from StarCalc', () => {
    const { ev } = fakeApp();
    expect(ev('DAYSINMONTH(DATE(2024,2,10))')).toBe(29);
    expect(ev('DAYSINYEAR(DATE(2023,6,1))')).toBe(365);
    expect(ev('ISLEAPYEAR(DATE(2000,1,1))')).toBe(true);
    expect(ev('ISLEAPYEAR(DATE(1900,3,1))')).toBe(false);
    expect(ev('WEEKSINYEAR(DATE(2020,5,5))')).toBe(53);
    expect(ev('WEEKSINYEAR(DATE(2021,5,5))')).toBe(52);
    expect(ev('EASTERSUNDAY(2024)')).toBe(ev('DATE(2024,3,31)'));
    expect(ev('EASTERSUNDAY(25)')).toBe(ev('DATE(2025,4,20)'));
    expect(easterSunday(2019)).toEqual({ m: 4, d: 21 });
    expect(isoWeeksInYear(2015)).toBe(53);
  });
});

describe('formula insight', () => {
  it('writes R1C1 notation', () => {
    expect(toR1C1('A1+$B$2*B1', 1, 1)).toBe('R[-1]C[-1]+R2C2*R[-1]C');
    expect(toR1C1('SUM(A1:A3)', 3, 0)).toBe('SUM(R[-3]C:R[-1]C)');
    expect(toR1C1("'My Sheet'!C5", 4, 2)).toBe("'My Sheet'!RC");
    expect(toR1C1('SUM(B:B)', 0, 0)).toBe('SUM(C[1])');
  });
  it('groups formulas into rules and finds the odd one out', () => {
    const { put, book } = fakeApp();
    for (let r = 1; r <= 5; r++) { put(`A${r}`, r); put(`B${r}`, r * 2); put(`C${r}`, r === 3 ? `=A${r}-B${r}` : `=A${r}+B${r}`); }
    const rules = formulaRules(book.sheet);
    expect(rules[0].r1c1).toBe('RC[-2]+RC[-1]');
    expect(rules[0].count).toBe(4);
    expect(rules[1].count).toBe(1);
    expect(inconsistentFormulas(book.sheet)).toEqual([{ r: 2, c: 2 }]);
    expect(cellKind(book.sheet, 2, 2)).toBe('odd');
    expect(cellKind(book.sheet, 0, 2)).toBe('formula');
    expect(cellKind(book.sheet, 0, 0)).toBe('number');
    expect(cellKind(book.sheet, 9, 9)).toBe('empty');
    expect(packRanges([{ r: 0, c: 0 }, { r: 1, c: 0 }, { r: 0, c: 1 }, { r: 1, c: 1 }, { r: 5, c: 0 }])).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 1 }, { r1: 5, c1: 0, r2: 5, c2: 0 }]);
  });
  it('explains a value and draws its dependencies', () => {
    const { put, get, book, engine } = fakeApp();
    put('A1', 100); put('A2', 0.2); put('B1', '=A1*A2'); put('C1', '=B1+SUM(A1:A2)'); put('D1', '=C1*2');
    expect(get('D1')).toBe(240.4);
    const tree = explainTree(engine, book.sheet, 0, 2);
    const text = explainText(tree);
    expect(text).toContain('C1 = 120.2   =B1+SUM(A1:A2)');
    expect(text).toContain('  B1 = 20   =A1*A2');
    expect(text).toContain('    A1 = 100   (input)');
    expect(text).toContain('A1:A2 = 2 values');
    const g = dependencyGraph(engine, book.sheet, 0, 2, 2);
    const labels = g.nodes.map((n) => `${n.label}@${n.level}`).sort();
    expect(labels).toContain('C1@0');
    expect(labels).toContain('B1@-1');
    expect(labels).toContain('A1@-2');
    expect(labels).toContain('D1@1');
    expect(g.edges.length).toBeGreaterThanOrEqual(4);
  });
  it('lists a cell history from undo and skips unchanged images', () => {
    const { put, book } = fakeApp();
    put('A1', 1); put('A1', 2); put('B1', 'x'); put('A1', '=1+2');
    const hist = cellHistory(book, book.sheet, 0, 0);
    expect(hist.map((e) => cellContentText(e.before))).toEqual(['2', '1', '(blank)']);
    expect(cellContentText(hist[0].after)).toBe('=1+2');
    book.transact('Sort', () => { book.recordSheet(book.sheet); book.sheet.set(5, 5, { v: 9 }); });
    expect(cellHistory(book, book.sheet, 0, 0).length).toBe(3);
  });
  it('Info Window rows and nearby labels', () => {
    const { app, put } = fakeApp();
    put('A1', 'Price'); put('B1', 12); put('B2', '=B1*2');
    const rows = Object.fromEntries(cellInfoRows(app, app.sheet, 1, 1));
    expect(rows.Formula).toBe('=B1*2');
    expect(rows.R1C1).toBe('=R[-1]C*2');
    expect(rows.Type).toBe('Formula');
    expect(labelFor(app.sheet, 0, 1)).toBe('Price');
  });
});

describe('models', () => {
  it('matrices', () => {
    const inv = matInverse([[4, 7], [2, 6]])!;
    expect(inv[0][0]).toBeCloseTo(0.6); expect(inv[0][1]).toBeCloseTo(-0.7); expect(inv[1][0]).toBeCloseTo(-0.2); expect(inv[1][1]).toBeCloseTo(0.4);
    expect(matInverse([[1, 2], [2, 4]])).toBeNull();
    expect(matMultiply([[1, 2], [3, 4]], [[5], [6]])).toEqual([[17], [39]]);
    expect(matMultiply([[1, 2]], [[1, 2]])).toBeNull();
    const x = matSolve([[2, 1], [1, 3]], [[3], [5]])!;
    expect(x[0][0]).toBeCloseTo(0.8); expect(x[1][0]).toBeCloseTo(1.4);
    expect(matDeterminant([[1, 2], [3, 4]])).toBeCloseTo(-2);
  });
  it('distributions', () => {
    expect(parseDist('10 to 20')).toEqual({ kind: 'uniform', min: 10, max: 20 });
    expect(parseDist('normal 100 15')).toEqual({ kind: 'normal', mean: 100, sd: 15 });
    expect(parseDist('triangular 1 3 8')).toEqual({ kind: 'triangular', min: 1, mode: 3, max: 8 });
    expect(parseDist('1, 2, 5')).toEqual({ kind: 'discrete', values: [1, 2, 5] });
    expect(parseDist('banana')).toBeNull();
    const rnd = seeded(7);
    for (let i = 0; i < 200; i++) {
      const u = sample({ kind: 'uniform', min: 10, max: 20 }, rnd); expect(u).toBeGreaterThanOrEqual(10); expect(u).toBeLessThanOrEqual(20);
      const t = sample({ kind: 'triangular', min: 1, mode: 3, max: 8 }, rnd); expect(t).toBeGreaterThanOrEqual(1); expect(t).toBeLessThanOrEqual(8);
    }
    const xs = Array.from({ length: 4000 }, () => sample({ kind: 'normal', mean: 50, sd: 5 }, rnd));
    const s = summarize(xs);
    expect(s.mean).toBeCloseTo(50, 0);
    expect(s.sd).toBeGreaterThan(4.5); expect(s.sd).toBeLessThan(5.5);
  });
  it('runs a Monte Carlo simulation and restores the inputs', () => {
    const { book, engine, put, get } = fakeApp();
    put('A1', 10); put('A2', 3); put('B1', '=A1*A2');
    const sh = book.sheet;
    const undo = book.undoStack.length;
    const res = monteCarlo(book, engine, [{ sheet: sh, r: 0, c: 0, dist: { kind: 'uniform', min: 1, max: 2 } }], [{ sheet: sh, r: 0, c: 1 }], 300, seeded(3));
    expect(res.values[0].length).toBe(300);
    expect(Math.min(...res.values[0])).toBeGreaterThanOrEqual(3);
    expect(Math.max(...res.values[0])).toBeLessThanOrEqual(6);
    expect(get('A1')).toBe(10);
    expect(get('B1')).toBe(30);
    expect(book.undoStack.length).toBe(undo);
  });
  it('finds the inputs behind a result, and slider ranges', () => {
    const { book, engine, put, get } = fakeApp();
    put('A1', 5); put('A2', 7); put('A3', 'text'); put('B1', '=A1*2'); put('C1', '=B1+SUM(A2:A3)');
    get('C1');
    const ins = leafInputs(engine, book.sheet, 0, 2).map((x) => `${x.r},${x.c}`).sort();
    expect(ins).toEqual(['0,0', '1,0']);
    expect(sliderRange(50)).toEqual({ min: 0, max: 100, step: 1 });
    expect(sliderRange(-3)).toEqual({ min: -6, max: 0, step: 0.06 });
    expect(sliderRange(0).max).toBe(10);
  });
  it('conditions, typed columns and accelerators', () => {
    expect(matchesCondition(5, 'gt', '3')).toBe(true);
    expect(matchesCondition('Apple', 'eq', 'apple')).toBe(true);
    expect(matchesCondition('Pineapple', 'contains', 'APP')).toBe(true);
    expect(matchesCondition(null, 'blank', '')).toBe(true);
    expect(matchesCondition('b', 'lt', 'c')).toBe(true);
    expect(matchesCondition('x', 'gt', '3')).toBe(false);
    expect(typeMismatches([1, 'a', null, 2.5], 'integer')).toEqual([1, 3]);
    expect(typeMismatches(['a@b.co', 'nope'], 'email')).toEqual([1]);
    expect(typeMismatches(['Red', 'blue', 'Green'], 'choice', ['red', 'Blue'])).toEqual([2]);
    expect(distinctChoices(['b', 'a', 'B', null, ''])).toEqual(['b', 'a']);
    expect(assignKeys(['Copy', 'Cut', 'Paste', 'Print'])).toEqual(['C', 'U', 'P', 'R']);
    expect(new Set(assignKeys(['Sort A to Z', 'Sort Z to A', 'Sort', 'Filter'])).size).toBe(4);
    expect(assignKeys(['Draw', 'Data'], [1, 0])).toEqual(['R', 'D']);
  });
});

describe('export as code', () => {
  it('produces JavaScript that computes the same values as the engine', () => {
    const { book, put, get } = fakeApp();
    put('A1', 10); put('A2', 20); put('A3', '=SUM(A1:A2)*2'); put('B1', 'Hi'); put('B2', '=B1&" "&UPPER("there")'); put('B3', '=IF(A3>50,"big","small")');
    put('C1', '=ROUND(A1/3,2)'); put('C2', '=IFERROR(1/0,"div")'); put('C3', '=-2^2'); put('C4', '=PMT(0.01,12,1000)'); put('C5', '=AVERAGE(A1:A3)'); put('C6', '=A1<>B1');
    put('D1', '=MOD(-7,3)'); put('D2', '=MAX(A1:C1)-MIN(A1:A2)'); put('D3', '=LEN(B1)+COUNTA(A1:B3)');
    const res = sheetToCode(book, book.sheet, 'js');
    expect(res.unsupported).toEqual([]);
    const out = new Function('console', res.code + '\nreturn RESULTS;')({ log() {} });
    for (const [k, v] of Object.entries(out)) {
      const want = get(k);
      if (typeof want === 'number') expect(v as number).toBeCloseTo(want, 9); else expect(v).toBe(want);
    }
    expect(Object.keys(out).length).toBe(12);
  });
  it('marks functions it cannot translate, and writes Python', () => {
    const { book, put } = fakeApp();
    put('A1', 1); put('A2', '=XLOOKUP(1,A1:A1,A1:A1)'); put('A3', '=A1*3');
    const js = sheetToCode(book, book.sheet, 'js');
    expect(js.unsupported).toEqual(['XLOOKUP']);
    const out = new Function('console', js.code + '\nreturn RESULTS;')({ log() {} });
    expect(out.A2).toBe('#unsupported: XLOOKUP');
    expect(out.A3).toBe(3);
    const py = sheetToCode(book, book.sheet, 'python');
    expect(py.code).toContain('A3 = (num(A1) * num(3))');
    expect(py.code).toContain('def SUM(*a)');
    expect(py.code).toContain('RESULTS = {');
  });
});

describe('heritage commands', () => {
  it('deletes matching records and closes the gaps', () => {
    const { app, put, get } = fakeApp();
    put('A1', 'Name'); put('B1', 'Qty'); put('C1', 'Double');
    const rows: [string, number][] = [['a', 1], ['b', 0], ['c', 3], ['d', 0], ['e', 5]];
    rows.forEach(([n, q], i) => { put(`A${i + 2}`, n); put(`B${i + 2}`, q); put(`C${i + 2}`, `=B${i + 2}*2`); });
    const removed = deleteMatching(app, parseRange('A1:C6')!, 1, 'eq', '0');
    expect(removed).toBe(2);
    expect([get('A2'), get('A3'), get('A4'), get('A5')]).toEqual(['a', 'c', 'e', null]);
    expect([get('C2'), get('C3'), get('C4')]).toEqual([2, 6, 10]);
    app.book.undo();
    expect(get('A3')).toBe('b');
  });
  it('writes matrix results as values or formulas', () => {
    const { app, put, get } = fakeApp();
    put('A1', 4); put('B1', 7); put('A2', 2); put('B2', 6);
    expect(matrixOp(app, { op: 'invert', a: 'A1:B2', out: 'D1' })).toBeNull();
    expect(get('D1')).toBeCloseTo(0.6); expect(get('E2')).toBeCloseTo(0.4);
    put('A4', 1); put('A5', 2);
    expect(matrixOp(app, { op: 'solve', a: 'A1:B2', b: 'A4:A5', out: 'G1' })).toBeNull();
    expect(4 * (get('G1') as number) + 7 * (get('G2') as number)).toBeCloseTo(1);
    expect(matrixOp(app, { op: 'determinant', a: 'A1:B2', out: 'H1', formulas: true })).toBeNull();
    expect(get('H1')).toBeCloseTo(10);
    put('A7', 'x');
    expect(matrixOp(app, { op: 'invert', a: 'A6:A7', out: 'J1' })).toMatch(/not a number/);
  });
  it('types a column with validation that judges the entered value', () => {
    const { app, put, select } = fakeApp();
    put('A1', 'Email'); put('A2', 'ann@example.com'); put('A3', 'bad address'); put('B1', 'Done'); put('B2', true);
    select('A2');
    const r = applyColumnType(app, 'email');
    expect(typeof r).toBe('object');
    expect((r as any).mismatches).toBe(1);
    const dv = app.sheet.validations.find((v: any) => v.type === 'custom')!;
    expect(dv.formula1).toBe('ISEMAIL(A2)');
    expect(validates(app, app.sheet, 3, 0, 'new@example.org', dv)).toBe(true);
    expect(validates(app, app.sheet, 3, 0, 'nope', dv)).toBe(false);
    expect(app.sheet.value(3, 0)).toBeNull(); // the check did not leave the value behind
    select('B2');
    applyColumnType(app, 'checkbox');
    expect(app.sheet.styleOf(1, 1).checkbox).toBe(true);
    expect(columnTypeSpec('rating', { r: 1, c: 0 }).dv?.formula2).toBe('5');
  });
  it('runs Monte Carlo into a results sheet', () => {
    const { app, put, book } = fakeApp();
    put('A1', 'Units'); put('B1', 100); put('A2', 'Price'); put('B2', 5); put('A3', 'Revenue'); put('B3', '=B1*B2');
    const n = book.sheets.length;
    expect(runMonteCarlo(app, { inputs: [{ ref: 'B1', dist: '80 to 120' }], outputs: ['B3'], iterations: 200, seed: 42 })).toBeNull();
    expect(book.sheets.length).toBe(n + 1);
    const out = book.sheets[book.sheets.length - 1];
    expect(out.value(0, 1)).toBe('B3 Revenue');
    const mean = out.value(1, 1) as number;
    expect(mean).toBeGreaterThan(450); expect(mean).toBeLessThan(550);
    expect(book.sheets[0].value(0, 1)).toBe(100);
    expect(runMonteCarlo(app, { inputs: [{ ref: 'B1', dist: 'whatever' }], outputs: ['B3'], iterations: 20 })).toMatch(/not a distribution/);
    expect(runMonteCarlo(app, { inputs: [], outputs: ['B1'], iterations: 20 })).toMatch(/not a formula/);
  });
  it('categorizes by rules with live formulas', async () => {
    const { app, put, get, select } = fakeApp();
    put('A1', 'Description'); put('A2', 'Uber ride'); put('A3', 'Corner coffee'); put('A4', 'Books');
    put('E1', 'uber,taxi'); put('F1', 'Travel'); put('E2', 'coffee'); put('F2', 'Food');
    select('A2');
    await app.exec('data.categorize', { rules: 'E1:F2', ifNone: 'Other' });
    expect([get('B1'), get('B2'), get('B3'), get('B4')]).toEqual(['Category', 'Travel', 'Food', 'Other']);
    expect(app.sheet.get(1, 1).f).toContain('CATEGORIZE(A2');
  });
  it('builds the slash menu from the ribbon', () => {
    const { app } = fakeApp();
    const tree = slashTree(app);
    expect(tree[0].label).toBe('File');
    expect(tree[tree.length - 1].label).toBe('Quit');
    const data = tree.find((t) => t.label === 'Data')!;
    const groups = data.children!();
    expect(groups.map((g) => g.label)).toContain('Records');
    const sort = groups.find((g) => g.label === 'Sort & Filter')!.children!();
    expect(sort[0].id).toBe('data.sortAZ');
    const ct = commandItem(app, 'data.columnType')!;
    expect(ct.children!().map((x) => x.arg)).toContain('email');
    expect(commandItem(app, 'tools.lotusEntry')!.label).toBe('1-2-3 Formula Entry (on)');
    const keys = Object.fromEntries(slashKeys(tree).map((k, i) => [tree[i].label, k]));
    expect(keys).toMatchObject({ File: 'F', Home: 'H', Data: 'D', Formulas: 'O', View: 'V', Quit: 'Q' });
    expect(keys.Draw).not.toBe('D');
  });
});
