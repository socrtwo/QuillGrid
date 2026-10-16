import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook, Sheet } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell, parseRange } from '../../src/model/address';
import { partsToSerial } from '../../src/model/dates';
import { defaultPrefs } from '../../src/prefs';
import { getCommand } from '../../src/commands/registry';
import {
  trimText, cleanText, changeCase, textToNumber, parseTextDate, ocrClean, splitText, unpivot, transpose, levenshtein, jaro, jaroWinkler,
  bestMatch, fuzzyScore, numberToWords, intToWords, toRoman, fromRoman, describe as stats, pearson, histogram, profileColumn, groupSummary,
  totalsFormulas, monthGrid, ymdDiff, businessDays, addToDate, removeAccents, extractMatches, EXTRACT_PATTERNS, compareSheetCells, guid, collapseSpaces,
} from '../../src/commands/tools';

/** Minimal App stand-in for running Tools commands on a Workbook. */
function fakeApp() {
  const book = new Workbook();
  const engine = new Engine(book);
  book.onChange = () => engine.recalc();
  const app: any = {
    book, engine, prefs: defaultPrefs(), recording: null,
    get sheet() { return book.sheet; },
    get tab() { return { book, engine, displayName: 'Test' }; },
    grid: { select: (r: number, c: number, o: any = {}) => { book.sheet.sel = { r, c, ranges: [o.range || { r1: r, c1: c, r2: r, c2: c }] }; }, objects: { select() {} }, invalidate() {} },
    isReadOnly: () => false,
    exec: (id: string, arg?: any) => { if (id === 'sheet.activate') { book.activeSheet = arg; return true; } const c = getCommand(id); return c ? c.run(app, arg) : false; },
    refresh() {}, ui: { status() {}, openPane() {}, rebuild() {} },
  };
  const put = (a: string, v: any, sh: Sheet = book.sheet) => { const p = parseCell(a)!; book.transact('t', () => sh.set(p.r, p.c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v })); };
  const get = (a: string, sh: Sheet = book.sheet) => { engine.recalc(); const p = parseCell(a)!; return sh.value(p.r, p.c); };
  const select = (ref: string) => { const rg = parseRange(ref)!; book.sheet.sel = { r: rg.r1, c: rg.c1, ranges: [rg] }; };
  return { app, book, engine, put, get, select };
}

describe('text clean-up', () => {
  it('trims like TRIM, including non-breaking spaces', () => {
    expect(trimText('  hello   world  ')).toBe('hello world');
    expect(trimText('  a  b ')).toBe('a b');
    expect(trimText('a \n  b')).toBe('a\nb');
    expect(collapseSpaces('  a   b ')).toBe(' a b ');
  });
  it('removes non-printing characters', () => {
    expect(cleanText('a\u0007b​c﻿\nd')).toBe('abcd');
    expect(cleanText('a\u0007\nb', true)).toBe('a\nb');
  });
  it('changes case', () => {
    expect(changeCase('hello WORLD', 'upper')).toBe('HELLO WORLD');
    expect(changeCase('Hello WORLD', 'lower')).toBe('hello world');
    expect(changeCase("mary-jane o'brien", 'proper')).toBe("Mary-Jane O'brien");
    expect(changeCase('THE END. a NEW start! i think', 'sentence')).toBe('The end. A new start! I think');
    expect(changeCase('Hello', 'toggle')).toBe('hELLO');
  });
  it('converts text to numbers', () => {
    expect(textToNumber('1,234')).toBe(1234);
    expect(textToNumber(' 42 ')).toBe(42);
    expect(textToNumber('(45)')).toBe(-45);
    expect(textToNumber('12.5-')).toBe(-12.5);
    expect(textToNumber('$3.50')).toBe(3.5);
    expect(textToNumber('50%')).toBe(0.5);
    expect(textToNumber('1 234 567')).toBe(1234567);
    expect(textToNumber("'007")).toBe(7);
    expect(textToNumber('abc')).toBeNull();
    expect(textToNumber('')).toBeNull();
  });
  it('removes accents and extracts patterns', () => {
    expect(removeAccents('Crème Brûlée, Łódź, Straße')).toBe('Creme Brulee, Lodz, Strasse');
    expect(extractMatches('Mail bob@example.com or ann.lee@test.co.uk', EXTRACT_PATTERNS.emails.re)).toEqual(['bob@example.com', 'ann.lee@test.co.uk']);
    expect(extractMatches('Paid 1,200.50 and -3', EXTRACT_PATTERNS.numbers.re)).toEqual(['1,200.50', '-3']);
  });
});

describe('fixing dates', () => {
  const S = (y: number, m: number, d: number) => partsToSerial(y, m, d);
  it('reads many formats', () => {
    expect(parseTextDate('2024-03-05')).toBe(S(2024, 3, 5));
    expect(parseTextDate('20240305')).toBe(S(2024, 3, 5));
    expect(parseTextDate('3/5/24')).toBe(S(2024, 3, 5));
    expect(parseTextDate('3/5/24', { order: 'dmy' })).toBe(S(2024, 5, 3));
    expect(parseTextDate('13/02/2024')).toBe(S(2024, 2, 13));
    expect(parseTextDate('05.03.2024', { order: 'dmy' })).toBe(S(2024, 3, 5));
    expect(parseTextDate('Mar 5, 2024')).toBe(S(2024, 3, 5));
    expect(parseTextDate('5 March 2024')).toBe(S(2024, 3, 5));
    expect(parseTextDate('5-Mar-24')).toBe(S(2024, 3, 5));
    expect(parseTextDate('Tuesday, March 5th, 2024')).toBe(S(2024, 3, 5));
    expect(parseTextDate('5th of March 2024')).toBe(S(2024, 3, 5));
    expect(parseTextDate('March 2024')).toBe(S(2024, 3, 1));
    expect(parseTextDate('2024 Mar 5')).toBe(S(2024, 3, 5));
    expect(parseTextDate('24/3/14', { order: 'ymd' })).toBe(S(2024, 3, 14));
  });
  it('handles two-digit years with a pivot', () => {
    expect(parseTextDate('1/1/29')).toBe(S(2029, 1, 1));
    expect(parseTextDate('1/1/30')).toBe(S(1930, 1, 1));
    expect(parseTextDate('1/1/50', { pivot: 50 })).toBe(S(2050, 1, 1));
  });
  it('keeps times and rejects impossible dates', () => {
    expect(parseTextDate('2024-03-05 14:30')).toBeCloseTo(S(2024, 3, 5) + 14.5 / 24, 8);
    expect(parseTextDate('3/5/2024 2:30 PM')).toBeCloseTo(S(2024, 3, 5) + 14.5 / 24, 8);
    expect(parseTextDate('2024-02-30')).toBeNull();
    expect(parseTextDate('hello')).toBeNull();
    expect(parseTextDate('Smarch 5, 2024')).toBeNull();
  });
  it('Fix Dates command converts text cells and sets a date format', async () => {
    const t = fakeApp();
    t.put('A1', '2024-03-05'); t.put('A2', 'Mar 6, 2024'); t.put('A3', 'not a date'); t.put('A4', 7);
    t.select('A1:A4');
    await getCommand('tools.fixDates')!.run(t.app, { order: 'auto', pivot: 29, format: 'yyyy-mm-dd' });
    expect(t.get('A1')).toBe(S(2024, 3, 5));
    expect(t.get('A2')).toBe(S(2024, 3, 6));
    expect(t.get('A3')).toBe('not a date');
    expect(t.book.sheet.styleOf(0, 0).numFmt).toBe('yyyy-mm-dd');
    t.book.undo();
    expect(t.get('A1')).toBe('2024-03-05');
  });
});

describe('OCR clean-up', () => {
  it('fixes ligatures, quotes, hyphenation and broken lines', () => {
    expect(ocrClean('The ﬁrst “quote” and it’s')).toBe('The first "quote" and it\'s');
    expect(ocrClean('exam-\nple text\nhere')).toBe('example text here');
    expect(ocrClean('para one\nstill one\n\npara two')).toBe('para one still one\npara two');
  });
  it('fixes O/0 and l/1 confusions by context', () => {
    expect(ocrClean('Total: 1O5')).toBe('Total: 105');
    expect(ocrClean('Paid $1,2O0.00 today')).toBe('Paid $1,200.00 today');
    expect(ocrClean('l23 Main St')).toBe('123 Main St');
    expect(ocrClean('ACME C0MPANY')).toBe('ACME COMPANY');
    expect(ocrClean('he1lo there')).toBe('hello there');
    expect(ocrClean('lOl Hello world')).toBe('lOl Hello world');
    expect(ocrClean('born 18O5 in Ohio', { joinLines: false })).toBe('born 1805 in Ohio');
  });
});

describe('fill blanks, split, merge, reshape', () => {
  it('fills blanks from above, with a value and with a formula', async () => {
    const t = fakeApp();
    t.put('A1', 'East'); t.put('B1', 1); t.put('B2', 2); t.put('A3', 'West'); t.put('B3', 3); t.put('B4', 4);
    t.select('A1:A4');
    await getCommand('tools.fillBlanks')!.run(t.app, { mode: 'above' });
    expect([t.get('A2'), t.get('A4')]).toEqual(['East', 'West']);
    const u = fakeApp();
    u.put('A1', 1); u.put('A3', 3); u.put('B1', 'x'); u.put('B3', 'y');
    u.select('A1:A3');
    await getCommand('tools.fillBlanks')!.run(u.app, { mode: 'value', value: '0' });
    expect(u.get('A2')).toBe(0);
    const w = fakeApp();
    w.put('A1', 10); w.put('A3', 30); w.put('B1', 'x'); w.put('B3', 'x');
    w.select('A1:A3');
    await getCommand('tools.fillBlanks')!.run(w.app, { mode: 'formula', value: '=A1*2' });
    expect(w.book.sheet.get(1, 0)?.f).toBe('A2*2');
  });
  it('splits text', () => {
    expect(splitText('a,"b,c",d', { kind: 'delimiter', delims: ',' })).toEqual(['a', 'b,c', 'd']);
    expect(splitText('a  b', { kind: 'delimiter', delims: ' ', consecutive: true })).toEqual(['a', 'b']);
    expect(splitText('ABC1234XY', { kind: 'positions', at: [3, 7] })).toEqual(['ABC', '1234', 'XY']);
    expect(splitText('firstNameHTTPServer', { kind: 'case' })).toEqual(['first', 'Name', 'HTTP', 'Server']);
    expect(splitText('12kg', { kind: 'numberText' })).toEqual(['12', 'kg']);
    expect(splitText('Room 12B', { kind: 'numberText' })).toEqual(['Room', '12', 'B']);
  });
  it('split column command writes typed parts to the right', async () => {
    const t = fakeApp();
    t.put('A1', 'Smith, John, 42'); t.put('A2', 'Doe, Jane, 37');
    t.select('A1:A2');
    await getCommand('tools.splitColumn')!.run(t.app, { mode: { kind: 'delimiter', delims: ',' }, keep: true });
    expect([t.get('B1'), t.get('C1'), t.get('D1'), t.get('D2')]).toEqual(['Smith', 'John', 42, 37]);
  });
  it('merges columns', async () => {
    const t = fakeApp();
    t.put('A1', 'Ann'); t.put('B1', 'Lee'); t.put('A2', 'Bo');
    t.select('A1:B2');
    await getCommand('tools.mergeColumns')!.run(t.app, { sep: ' ', skipBlanks: true, keep: true });
    expect([t.get('C1'), t.get('C2')]).toEqual(['Ann Lee', 'Bo']);
  });
  it('unpivots wide data', () => {
    const out = unpivot(['Name', 'Jan', 'Feb'], [['A', 1, 2], ['B', null, 4]], 1);
    expect(out).toEqual([['Name', 'Attribute', 'Value'], ['A', 'Jan', 1], ['A', 'Feb', 2], ['B', 'Feb', 4]]);
    expect(unpivot(['Name', 'Jan'], [['B', null]], 1, true).length).toBe(2);
  });
  it('transposes, reverses and shuffles in place', async () => {
    expect(transpose([[1, 2, 3], [4, 5, 6]])).toEqual([[1, 4], [2, 5], [3, 6]]);
    const t = fakeApp();
    t.put('A1', 1); t.put('B1', 2); t.put('A2', 3); t.put('B2', 4); t.put('A3', 5); t.put('B3', 6);
    t.select('A1:B3');
    await getCommand('tools.transposeInPlace')!.run(t.app);
    expect([t.get('A1'), t.get('B1'), t.get('C1'), t.get('A2'), t.get('C2'), t.get('A3')]).toEqual([1, 3, 5, 2, 6, null]);
    const u = fakeApp();
    u.put('A1', 1); u.put('A2', 2); u.put('A3', 3); u.put('B1', '=A1*10');
    u.select('A1:A3');
    await getCommand('tools.reverse')!.run(u.app, 'rows');
    expect([u.get('A1'), u.get('A2'), u.get('A3')]).toEqual([3, 2, 1]);
  });
});

describe('fuzzy matching', () => {
  it('computes distances and similarities', () => {
    expect(levenshtein('kitten', 'sitting')).toBe(3);
    expect(levenshtein('', 'abc')).toBe(3);
    expect(jaro('MARTHA', 'MARHTA')).toBeCloseTo(0.9444, 3);
    expect(jaroWinkler('MARTHA', 'MARHTA')).toBeCloseTo(0.9611, 3);
    expect(jaroWinkler('DIXON', 'DICKSONX')).toBeCloseTo(0.8133, 3);
    expect(jaroWinkler('abc', 'abc')).toBe(1);
  });
  it('finds the best match with normalization', () => {
    const m = bestMatch('Jhon Smith', ['John Smith', 'Jane Smith', 'Joan Smythe'], 'levenshtein');
    expect(m?.match).toBe('John Smith');
    expect(fuzzyScore('ACME Corp.', 'acme corp', 'jaroWinkler')).toBe(1);
    expect(fuzzyScore('Smith John', 'John Smith', 'tokens')).toBe(1);
    expect(bestMatch('x', [])).toBeNull();
  });
});

describe('numbers in words and roman numerals', () => {
  it('spells numbers', () => {
    expect(intToWords(0)).toBe('zero');
    expect(numberToWords(1234)).toBe('one thousand two hundred thirty-four');
    expect(numberToWords(-5)).toBe('minus five');
    expect(numberToWords(3.14)).toBe('three point one four');
    expect(numberToWords(1000000)).toBe('one million');
    expect(numberToWords(1234.56, 'check')).toBe('One Thousand Two Hundred Thirty-Four and 56/100');
    expect(numberToWords(99.999, 'check', 'Dollars')).toBe('One Hundred and 00/100 Dollars');
    expect(numberToWords(123, 'legal')).toBe('One Hundred and Twenty-Three');
    expect(numberToWords(2005, 'legal')).toBe('Two Thousand and Five');
    expect(numberToWords(21, 'ordinal')).toBe('twenty-first');
    expect(numberToWords(12, 'ordinal')).toBe('twelfth');
    expect(numberToWords(40, 'ordinal')).toBe('fortieth');
    expect(numberToWords(100, 'ordinal')).toBe('one hundredth');
    expect(numberToWords(1850, 'year')).toBe('eighteen fifty');
    expect(numberToWords(1905, 'year')).toBe('nineteen oh five');
    expect(numberToWords(1900, 'year')).toBe('nineteen hundred');
    expect(numberToWords(2005, 'year')).toBe('two thousand five');
    expect(numberToWords(2024, 'year')).toBe('twenty twenty-four');
  });
  it('converts roman numerals', () => {
    expect(toRoman(1994)).toBe('MCMXCIV');
    expect(toRoman(3999)).toBe('MMMCMXCIX');
    expect(fromRoman('mcmxciv')).toBe(1994);
    expect(fromRoman('IIII')).toBeNull();
    expect(fromRoman('ABC')).toBeNull();
    expect(() => toRoman(0)).toThrow();
  });
});

describe('statistics and generators', () => {
  it('describes data like the analysis tools', () => {
    const d = stats([2, 4, 4, 4, 5, 5, 7, 9]);
    expect(d.mean).toBe(5);
    expect(d.median).toBe(4.5);
    expect(d.mode).toBe(4);
    expect(d.stdev).toBeCloseTo(Math.sqrt(32 / 7), 10);
    expect(d.range).toBe(7);
    expect(d.skewness).toBeCloseTo(0.8185, 3);
    expect(d.kurtosis).toBeCloseTo(0.9406, 3);
    expect(pearson([1, 2, 3], [2, 4, 6])).toBeCloseTo(1, 10);
    expect(pearson([1, 2, 3], [3, 2, 1])).toBeCloseTo(-1, 10);
  });
  it('profiles columns and groups', () => {
    const p = profileColumn('X', [1, 2, 'a', null, '', 2]);
    expect(p).toMatchObject({ numbers: 3, text: 1, blanks: 2, distinct: 3, min: 1, max: 2 });
    expect(p.type).toBe('Mixed (mostly number)');
    expect(groupSummary(['a', 'b', 'a'], [1, 2, 3], 'sum')).toEqual([['a', 4], ['b', 2]]);
    expect(groupSummary(['a', 'b', 'a'], [1, 2, 3], 'average')).toEqual([['a', 2], ['b', 2]]);
  });
  it('bins histograms', () => {
    const hst = histogram([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 5);
    expect(hst.counts).toEqual([2, 2, 2, 2, 2]);
    expect(hst.bounds[4]).toBe(10);
  });
  it('builds totals formulas', () => {
    const rg = { r1: 1, c1: 1, r2: 3, c2: 2 };
    expect(totalsFormulas(rg, 'sum', 'rows')).toEqual([{ r: 4, c: 1, f: 'SUM(B2:B4)' }, { r: 4, c: 2, f: 'SUM(C2:C4)' }]);
    expect(totalsFormulas(rg, 'percent', 'cols')[0]).toEqual({ r: 1, c: 3, f: 'SUM(B2:C2)/SUM($B$2:$C$4)', numFmt: '0.0%' });
    expect(totalsFormulas(rg, 'running', 'rows')[1].f).toBe('SUM($B$2:C$4)');
  });
  it('calendar and date arithmetic', () => {
    const g = monthGrid(2024, 2, 0);
    expect(g[0]).toEqual([0, 0, 0, 0, 1, 2, 3]);
    expect(g.flat().filter(Boolean).length).toBe(29);
    expect(monthGrid(2024, 2, 1)[0].slice(0, 4)).toEqual([0, 0, 0, 1]);
    const a = partsToSerial(2024, 1, 1), b = partsToSerial(2024, 1, 7);
    expect(businessDays(a, b)).toBe(5);
    expect(ymdDiff(partsToSerial(1990, 5, 20), partsToSerial(2024, 3, 5))).toEqual({ y: 33, m: 9, d: 14 });
    expect(addToDate(partsToSerial(2024, 1, 31), 1, 'months')).toBe(partsToSerial(2024, 2, 29));
    expect(addToDate(partsToSerial(2024, 1, 5), 1, 'workdays')).toBe(partsToSerial(2024, 1, 8));
    expect(guid()).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
  });
  it('compares sheets cell by cell', () => {
    const b = new Workbook();
    const A = b.sheet, B = new Sheet(b, 'B');
    b.sheets.push(B);
    A.set(0, 0, { v: 1 }); B.set(0, 0, { v: 1 });
    A.set(1, 0, { v: 'x' }); B.set(1, 0, { v: 'y' });
    B.set(2, 0, { v: 5 });
    A.set(3, 0, { f: 'A1+1', v: 2 }); B.set(3, 0, { f: 'A1*2', v: 2 });
    const d = compareSheetCells(A, B, true);
    expect(d.map((x) => x.kind)).toEqual(['Changed', 'Added', 'Formula']);
  });
});

describe('commands on the workbook', () => {
  it('trim and case commands are one undo step and skip locked cells on protected sheets', async () => {
    const t = fakeApp();
    t.put('A1', '  a  b '); t.put('A2', ' c ');
    t.select('A1:A2');
    await getCommand('tools.trimSpaces')!.run(t.app);
    expect([t.get('A1'), t.get('A2')]).toEqual(['a b', 'c']);
    await getCommand('tools.changeCase')!.run(t.app, 'upper');
    expect(t.get('A1')).toBe('A B');
    t.book.undo();
    expect(t.get('A1')).toBe('a b');
    t.book.sheet.protection.on = true;
    let notice = 0;
    const exec = t.app.exec;
    t.app.exec = (id: string, a?: any) => { if (id === 'review.protectionNotice') { notice++; return; } return exec(id, a); };
    await getCommand('tools.changeCase')!.run(t.app, 'upper');
    expect(t.get('A1')).toBe('a b');
    expect(notice).toBe(1);
  });
  it('number to words writes into the next column', async () => {
    const t = fakeApp();
    t.put('A1', 21); t.put('A2', 1850);
    t.select('A1:A2');
    await getCommand('tools.numberToWords')!.run(t.app, { style: 'ordinal', dest: 'right' });
    expect([t.get('B1'), t.get('B2')]).toEqual(['twenty-first', 'one thousand eight hundred fiftieth']);
  });
  it('text to numbers converts only convertible text', async () => {
    const t = fakeApp();
    t.put('A1', '1,000'); t.put('A2', 'n/a'); t.put('A3', '=1+1');
    t.select('A1:A3');
    await getCommand('tools.textToNumbers')!.run(t.app);
    expect([t.get('A1'), t.get('A2'), t.get('A3')]).toEqual([1000, 'n/a', 2]);
  });
  it('formulas to values and remove blank rows', async () => {
    const t = fakeApp();
    t.put('A1', 2); t.put('A3', '=A1*3'); t.put('B3', 'x');
    t.select('A1:B3');
    await getCommand('tools.formulasToValues')!.run(t.app);
    expect(t.book.sheet.get(2, 0)?.f).toBeUndefined();
    expect(t.get('A3')).toBe(6);
    await getCommand('tools.deleteBlankRows')!.run(t.app);
    expect([t.get('A2'), t.get('B2'), t.get('A3')]).toEqual([6, 'x', null]);
  });
});
