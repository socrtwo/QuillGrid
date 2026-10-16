import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { DEFAULT_THEME } from '../../src/model/workbook';
import { createTable, convertTableToRange, setTotalsRow, setHeaderRow, renameTable, uniqueTableName, uniqueColumnNames, validTableName, structuredToA1, insertTableRows, colRef } from '../../src/commands/table';
import { buildSparklines, parseLink, buildLink } from '../../src/commands/insert';
import { shiftCells } from '../../src/model/structure';

function setup() {
  const book = new Workbook();
  const s = book.sheet;
  const rows: any[][] = [['Item', 'Qty', 'Price'], ['Apples', 3, 1.25], ['Pears', 5, 0.8], ['Plums', 12, 0.35]];
  book.transact('t', () => rows.forEach((row, r) => row.forEach((v, c) => s.set(r, c, { v }))));
  const engine = new Engine(book);
  engine.rebuild(); engine.recalc(true);
  return { book, s, engine };
}

describe('table creation', () => {
  it('creates a table with header names and a unique name', () => {
    const { book, s } = setup();
    let t: any;
    book.transact('c', () => { t = createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 }, { headers: true }); });
    expect(t.name).toBe('Table1');
    expect(t.columns.map((c: any) => c.name)).toEqual(['Item', 'Qty', 'Price']);
    expect(t.style).toBe('TableStyleMedium2');
    expect(t.opts.header).toBe(true);
    expect(uniqueTableName(book)).toBe('Table2');
    expect(() => book.transact('c2', () => createTable(s, { r1: 2, c1: 1, r2: 5, c2: 4 }, { headers: true }))).toThrow(/overlap/);
    book.undo();
    expect(s.tables.length).toBe(0);
  });
  it('inserts a generated header row when the data has no headers', () => {
    const { book, s } = setup();
    let t: any;
    book.transact('c', () => { t = createTable(s, { r1: 1, c1: 0, r2: 3, c2: 2 }, { headers: false }); });
    expect(t.range).toEqual({ r1: 1, c1: 0, r2: 4, c2: 2 });
    expect(s.value(1, 0)).toBe('Column1');
    expect(s.value(2, 0)).toBe('Apples');
    expect(s.value(4, 2)).toBe(0.35);
  });
  it('makes column names unique and validates table names', () => {
    expect(uniqueColumnNames(['A', '', 'A', 'a'])).toEqual(['A', 'Column2', 'A2', 'a3']);
    const { book } = setup();
    expect(validTableName(book, 'Sales_2024')).toBeNull();
    expect(validTableName(book, 'A1')).toMatch(/cell reference/);
    expect(validTableName(book, 'My Table')).toMatch(/start with/);
    expect(colRef('Unit Price')).toBe('[Unit Price]');
    expect(colRef("Q#1")).toBe("[Q'#1]");
  });
});

describe('table options and editing', () => {
  it('adds a total row with SUBTOTAL for the last numeric column', () => {
    const { book, s, engine } = setup();
    let t: any;
    book.transact('c', () => { t = createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 }, { headers: true }); });
    book.transact('tot', () => setTotalsRow(s, t, true));
    expect(t.range.r2).toBe(4);
    expect(s.value(4, 0)).toBe('Total');
    expect(s.get(4, 2)?.f).toBe('SUBTOTAL(109,Table1[Price])');
    engine.recalc(true);
    expect(s.value(4, 2)).toBeCloseTo(2.4);
    book.transact('tot2', () => setTotalsRow(s, t, false));
    expect(t.range.r2).toBe(3);
    expect(s.get(4, 2)).toBeUndefined();
  });
  it('toggles the header row and keeps column names', () => {
    const { book, s } = setup();
    let t: any;
    book.transact('c', () => { t = createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 }, { headers: true }); });
    book.transact('h', () => setHeaderRow(s, t, false));
    expect(t.range.r1).toBe(1);
    expect(s.value(0, 0)).toBeNull();
    book.transact('h2', () => setHeaderRow(s, t, true));
    expect(t.range.r1).toBe(0);
    expect(s.value(0, 1)).toBe('Qty');
  });
  it('renames tables in formulas and converts structured references', () => {
    const { book, s, engine } = setup();
    let t: any;
    book.transact('c', () => { t = createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 }, { headers: true }); s.set(6, 0, { f: 'SUM(Table1[Qty])' }); });
    engine.rebuild(); engine.recalc(true);
    expect(s.value(6, 0)).toBe(20);
    book.transact('r', () => renameTable(book, t, 'Fruit'));
    expect(s.get(6, 0)?.f).toBe('SUM(Fruit[Qty])');
    expect(structuredToA1('SUM(Fruit[Qty])', s, t, s)).toBe('SUM($B$2:$B$4)');
    expect(structuredToA1('Fruit[[#All],[Item]:[Qty]]', s, t, s)).toBe('$A$1:$B$4');
    expect(structuredToA1('[@Qty]*[@Price]', s, t, s, 2)).toBe('$B$3*$C$3');
  });
  it('converts a table to a range, baking in its look', () => {
    const { book, s, engine } = setup();
    let t: any;
    book.transact('c', () => { t = createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 }, { headers: true }); s.set(6, 0, { f: 'SUM(Table1[Qty])' }); });
    book.transact('conv', () => convertTableToRange(s, t, DEFAULT_THEME));
    expect(s.tables.length).toBe(0);
    const hs = s.styleOf(0, 0);
    expect(hs.fill).toBeTruthy(); expect(hs.bold).toBe(true);
    expect(s.get(6, 0)?.f).toBe('SUM($B$2:$B$4)');
    engine.rebuild(); engine.recalc(true);
    expect(s.value(6, 0)).toBe(20);
    book.undo();
    expect(s.tables.length).toBe(1);
  });
  it('stays valid when rows are inserted and table rows are added', () => {
    const { book, s } = setup();
    let t: any;
    book.transact('c', () => { t = createTable(s, { r1: 0, c1: 0, r2: 3, c2: 2 }, { headers: true }); });
    book.transact('ins', () => { book.recordSheet(s); shiftCells(s, 'row', 2, 2); });
    expect(s.tables[0].range).toEqual({ r1: 0, c1: 0, r2: 5, c2: 2 });
    const t2 = s.tables[0];
    book.transact('tr', () => insertTableRows(s, t2, 6, 1));
    expect(t2.range.r2).toBe(6);
    void t;
  });
});

describe('insert helpers', () => {
  it('builds one sparkline per row of the data', () => {
    const { s } = setup();
    const list = buildSparklines(s, { sheet: s, range: { r1: 1, c1: 1, r2: 3, c2: 2 } }, { r1: 1, c1: 4, r2: 3, c2: 4 }, 'line');
    expect(list.map((x) => x.data)).toEqual(['$B$2:$C$2', '$B$3:$C$3', '$B$4:$C$4']);
    expect(list.map((x) => x.cell)).toEqual([{ r: 1, c: 4 }, { r: 2, c: 4 }, { r: 3, c: 4 }]);
    expect(() => buildSparklines(s, { sheet: s, range: { r1: 1, c1: 1, r2: 3, c2: 2 } }, { r1: 1, c1: 4, r2: 7, c2: 4 }, 'line')).toThrow();
  });
  it('parses and builds hyperlinks', () => {
    expect(buildLink({ kind: 'web', url: 'example.com', sheet: '', cell: '', email: '', subject: '' })).toBe('https://example.com');
    expect(buildLink({ kind: 'place', url: '', sheet: 'My Sheet', cell: 'B4', email: '', subject: '' })).toBe("#'My Sheet'!B4");
    expect(buildLink({ kind: 'place', url: '', sheet: 'Sheet1', cell: 'Totals', email: '', subject: '' })).toBe('#Totals');
    expect(buildLink({ kind: 'email', url: '', sheet: '', cell: '', email: 'a@b.com', subject: 'Hi there' })).toBe('mailto:a@b.com?subject=Hi%20there');
    expect(parseLink("#'My Sheet'!B4", 'Sheet1')).toMatchObject({ kind: 'place', sheet: 'My Sheet', cell: 'B4' });
    expect(parseLink('mailto:a@b.com?subject=Hi%20there', 'Sheet1')).toMatchObject({ kind: 'email', email: 'a@b.com', subject: 'Hi there' });
  });
});
