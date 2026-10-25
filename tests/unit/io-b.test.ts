/**
 * Modern, XML and text spreadsheet formats (the "B" set): detection, encodings, locale-aware
 * text import, fixed width, JSON, SQLite, Gnumeric, OpenOffice.org 1.x, Excel binary (.xlsb),
 * Excel Strict, Apple Numbers, KSpread, UOF, MHTML, and fidelity of the existing importers.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { gzipSync } from 'node:zlib';
import { cellKey } from '../../src/model/address';
import type { Sheet } from '../../src/model/workbook';
import { CellError } from '../../src/model/types';
import { importFile, sniffFormat } from '../../src/io/index';
import { SNIFF_RULES_B } from '../../src/io/sniff-b';
import { detectEncoding, decodeAuto } from '../../src/io/encoding';
import { detectLocale } from '../../src/io/csv';
import '../../src/calc/functions';
import { Engine } from '../../src/calc/engine';
import JSZip from 'jszip';
import type { Workbook } from '../../src/model/workbook';

/** Recalculates every formula and returns the cells whose result differs from the value the file stored. */
export function recalcDiffs(book: Workbook): string[] {
  const cached: [Sheet, number, unknown][] = [];
  for (const sh of book.sheets) for (const [k, c] of sh.cells) if (c.f != null) cached.push([sh, k, c.v]);
  const engine = new Engine(book);
  engine.rebuild();
  engine.recalc();
  const out: string[] = [];
  for (const [sh, k, v] of cached) {
    const now = sh.cells.get(k)?.v;
    const same = typeof v === 'number' && typeof now === 'number' ? Math.abs(v - now) < 1e-9 * Math.max(1, Math.abs(v)) : v instanceof CellError || now instanceof CellError ? String(v) === String(now) : v === now;
    if (!same && v != null) out.push(`${sh.name}!${k}: stored ${String(v)}, calculated ${String(now)}`);
  }
  return out;
}

const FIX = 'tests/fixtures/io/b/';
export function fixture(name: string): ArrayBuffer {
  const b = readFileSync(FIX + name);
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}
export const ab = (u8: Uint8Array): ArrayBuffer => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
export const txt = (s: string): ArrayBuffer => ab(new TextEncoder().encode(s));

const K = (a1: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return cellKey(+m[2] - 1, c);
};
export const val = (s: Sheet, a1: string) => s.cells.get(K(a1))?.v ?? null;
export const fml = (s: Sheet, a1: string) => s.cells.get(K(a1))?.f;
export const sty = (s: Sheet, a1: string) => s.book.styles[s.cells.get(K(a1))?.s ?? 0] || {};

describe('detection rules', () => {
  it('every rule names a registered format', async () => {
    const { formatById } = await import('../../src/io/index');
    for (const r of SNIFF_RULES_B) { expect(formatById(r.format), r.id).toBeTruthy(); expect(r.category).toBe('spreadsheet'); expect(r.ext.length).toBeGreaterThan(0); }
  });
  it('identifies content regardless of the extension', () => {
    const id = (data: ArrayBuffer, name: string) => sniffFormat(data, name)?.id;
    expect(id(fixture('sample.sqlite'), 'data.bin')).toBe('sqlite');
    expect(id(fixture('sample.gnumeric'), 'mystery.dat')).toBe('gz');
    expect(id(fixture('sample.gnumeric'), 'book.gnumeric')).toBe('gnumeric');
    expect(id(txt('<?xml version="1.0"?>\n<gnm:Workbook xmlns:gnm="http://www.gnumeric.org/v10.dtd"></gnm:Workbook>'), 'x.xml')).toBe('gnumeric');
    expect(id(txt('[{"a":1},{"a":2}]'), 'x.txt')).toBe('json');
    expect(id(txt('{"a":1}\n{"a":2}\n{"a":3}\n'), 'x.log')).toBe('ndjson');
    expect(id(txt('Name      Qty   Price\nApple       3    1.25\nPear        5    2.50\nPlum       12    0.75\n'), 'report')).toBe('prn');
    expect(id(txt('a,b,c\n1,2,3\n4,5,6\n'), 'x.csv')).toBe('csv');
    expect(id(txt('TABLE\r\n0,1\r\n"EXCEL"\r\nVECTORS\r\n0,2\r\n""\r\n'), 'noext')).toBe('dif');
    expect(id(txt('MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary="b"\r\n\r\n--b\r\n'), 'page')).toBe('mht');
  });
});

describe('text encodings', () => {
  const enc = (name: string) => new Uint8Array(readFileSync(`${FIX}encodings/${name}`));
  it('detects legacy code pages and CJK encodings', () => {
    const cases: [string, string, string][] = [
      ['cp1252.txt', 'windows-1252', 'Größe'], ['cp1250.txt', 'windows-1250', 'Žluťoučký'], ['cp1251.txt', 'windows-1251', 'Москва'],
      ['cp1253.txt', 'windows-1253', 'Αθήνα'], ['cp1255.txt', 'windows-1255', 'ירושלים'], ['mac.txt', 'macintosh', 'Café crème'],
      ['cp850.txt', 'cp850', 'Ørsted'], ['cp437.txt', '', 'Zürich'], ['sjis.txt', 'shift_jis', 'こんにちは'], ['gbk.txt', 'gbk', '这是中国'],
      ['big5.txt', 'big5', '這是一個'], ['euckr.txt', 'euc-kr', '안녕하세요'], ['koi8.txt', 'koi8-r', 'Петербург'],
    ];
    for (const [f, e, sample] of cases) {
      const u8 = enc(f);
      if (e) expect(detectEncoding(u8).encoding, f).toBe(e);
      expect(decodeAuto(u8).text, f).toContain(sample);
    }
  });
  it('reads UTF-16 with and without a byte order mark', () => {
    const s = 'Name\tStadt\nMüller\tKöln\n';
    const le = new Uint8Array(s.length * 2);
    for (let i = 0; i < s.length; i++) { le[i * 2] = s.charCodeAt(i); le[i * 2 + 1] = 0; }
    expect(decodeAuto(le).text).toBe(s);
    const be = new Uint8Array(2 + s.length * 2); be[0] = 0xfe; be[1] = 0xff;
    for (let i = 0; i < s.length; i++) { be[2 + i * 2] = 0; be[3 + i * 2] = s.charCodeAt(i); }
    expect(decodeAuto(be).text).toBe(s);
  });
  it('opens a Windows-1252 semicolon CSV with decimal commas and day-first dates', async () => {
    const { book } = await importFile(fixture('encodings/cp1252.txt'), 'de.csv');
    const s = book.sheets[0];
    expect(val(s, 'A2')).toBe('Müller');
    expect(val(s, 'C2')).toBe(1.85);
    const r = await importFile(txt('Datum;Betrag;Menge\n15.03.2024;1.234,56;3\n01.04.2024;-7,5;1.000\n'), 'x.csv');
    const t = r.book.sheets[0];
    expect(val(t, 'B2')).toBe(1234.56);
    expect(val(t, 'B3')).toBe(-7.5);
    expect(val(t, 'C3')).toBe(1000);
    expect(val(t, 'A2')).toBe(45366); // 15 March 2024
    expect(val(t, 'A3')).toBe(45383); // 1 April 2024
  });
  it('detects number and date conventions', () => {
    expect(detectLocale([['1.234,56', '3,5'], ['12,25', '7']], ';').decimal).toBe(',');
    expect(detectLocale([['1,234.56', '3.5']], ',').decimal).toBe('.');
    expect(detectLocale([['25/12/2024'], ['01/02/2024']]).dateOrder).toBe('dmy');
    expect(detectLocale([['12/25/2024'], ['01/02/2024']]).dateOrder).toBe('mdy');
  });
  it('honors the sep= line and DOS code page 437', async () => {
    const r = await importFile(txt('sep=|\na|b\n1|2\n'), 'x.csv');
    expect(val(r.book.sheets[0], 'B2')).toBe(2);
    const bytes = new Uint8Array([...new TextEncoder().encode('Name,Ort\nRen'), 0x82, ...new TextEncoder().encode('e,Z'), 0x81, ...new TextEncoder().encode('rich\n')]);
    const d = await importFile(ab(bytes), 'dos.csv', { csv: { encoding: 'cp437' } });
    expect(val(d.book.sheets[0], 'A2')).toBe('Renée');
    expect(val(d.book.sheets[0], 'B2')).toBe('Zürich');
  });
});

describe('fixed width text', () => {
  it('splits aligned columns and types the values', async () => {
    const t = [
      'Item        Qty     Price  Date',
      'Apples        3      1.25  2024-03-15',
      'Big pears    12  1,234.50  2024-03-16',
      'Plums         5      0.80  2024-03-17',
    ].join('\r\n');
    const { book, format } = await importFile(txt(t), 'report.prn');
    expect(format).toBe('prn');
    const s = book.sheets[0];
    expect(val(s, 'A3')).toBe('Big pears');
    expect(val(s, 'B3')).toBe(12);
    expect(val(s, 'C3')).toBe(1234.5);
    expect(val(s, 'D2')).toBe(45366);
  });
  it('falls back to delimited text when columns do not line up', async () => {
    const { book, warnings } = await importFile(txt('a b c\n1 22 333\n4444 5 6\n'), 'x.prn');
    expect(warnings.join(' ')).toMatch(/delimited/);
    expect(book.sheets[0].cells.size).toBeGreaterThan(0);
  });
});

describe('JSON', () => {
  it('reads arrays of records with nested objects and dates', async () => {
    const data = [{ name: 'Ann', age: 31, address: { city: 'Oslo' }, tags: ['a', 'b'], joined: '2024-03-15' }, { name: 'Bo', age: 25, extra: true }];
    const { book } = await importFile(txt(JSON.stringify(data)), 'people.json');
    const s = book.sheets[0];
    expect([val(s, 'A1'), val(s, 'B1'), val(s, 'C1'), val(s, 'D1'), val(s, 'E1'), val(s, 'F1')]).toEqual(['name', 'age', 'address.city', 'tags', 'joined', 'extra']);
    expect(val(s, 'C2')).toBe('Oslo');
    expect(val(s, 'D2')).toBe('a, b');
    expect(val(s, 'E2')).toBe(45366);
    expect(sty(s, 'E2').numFmt).toBe('yyyy-mm-dd');
    expect(val(s, 'F3')).toBe(true);
    expect(s.freeze.rows).toBe(1);
  });
  it('reads row arrays, column objects, wrapped arrays and NDJSON', async () => {
    let r = await importFile(txt('[["a","b"],[1,2]]'), 'rows.json');
    expect(val(r.book.sheets[0], 'B2')).toBe(2);
    r = await importFile(txt('{"x":[1,2,3],"y":[4,5,6]}'), 'cols.json');
    expect(val(r.book.sheets[0], 'B4')).toBe(6);
    r = await importFile(txt('{"meta":{"v":1},"users":[{"id":1}],"orders":[{"id":7,"total":9.5}]}'), 'api.json');
    expect(r.book.sheets.map((s) => s.name)).toEqual(['users', 'orders']);
    expect(val(r.book.sheets[1], 'B2')).toBe(9.5);
    r = await importFile(txt('{"a":1,"b":"x"}\n{"a":2,"b":"y"}\nnot json\n'), 'log.jsonl');
    expect(val(r.book.sheets[0], 'B3')).toBe('y');
    expect(r.warnings.join(' ')).toMatch(/skipped/);
  });
  it('still opens Anne\'s Grid JSON natively', async () => {
    const { exportFile } = await import('../../src/io/index');
    const { Workbook } = await import('../../src/model/workbook');
    const b = new Workbook(); b.sheets[0].cells.set(K('A1'), { v: 'native' });
    const blob = await exportFile(b, 'json');
    const r = await importFile(txt(await blob.text()), 'book.json');
    expect(val(r.book.sheets[0], 'A1')).toBe('native');
  });
});

describe('SQLite', () => {
  it('opens every table as a sheet', async () => {
    const { book, warnings, format } = await importFile(fixture('sample.sqlite'), 'shop.db');
    expect(format).toBe('sqlite');
    expect(book.sheets.map((s) => s.name)).toEqual(['fruit', 'big table', 'kv', 'ints']);
    const f = book.sheets[0];
    expect([val(f, 'A1'), val(f, 'B1'), val(f, 'F1')]).toEqual(['id', 'name', 'pic']);
    expect([val(f, 'A2'), val(f, 'B2'), val(f, 'C2'), val(f, 'D2')]).toEqual([1, 'Apple', 1.25, 3]);
    expect(val(f, 'E3')).toBe('Crème brûlée – 日本');
    expect(val(f, 'F3')).toBe('[BLOB 3 bytes]');
    expect((val(f, 'E4') as string).length).toBe(5000); // overflow pages
    const big = book.sheets[1];
    expect(big.usedRange()!.r2).toBe(1000);
    expect(val(big, 'B1001')).toBe(1000000);
    expect([val(book.sheets[2], 'A2'), val(book.sheets[2], 'B4')]).toEqual(['a', 3]); // WITHOUT ROWID, key order
    expect(val(book.sheets[3], 'A2')).toBe('9007199254740993');
    expect(warnings.join(' ')).toMatch(/Views.*cheap/);
  });
});

describe('Gnumeric', () => {
  it('reads values, shared formulas, styles, validation, links and layout', async () => {
    const { book, format } = await importFile(fixture('sample.gnumeric'), 'sample.gnumeric');
    expect(format).toBe('gnumeric');
    expect(book.sheets.map((s) => s.name)).toEqual(['Fruit', 'Notes & More']);
    const s = book.sheets[0];
    expect([val(s, 'A2'), val(s, 'B2'), val(s, 'C2')]).toEqual(['Apple', 3, 1.25]);
    expect([fml(s, 'D2'), fml(s, 'D3'), fml(s, 'D4'), fml(s, 'D5')]).toEqual(['B2*C2', 'B3*C3', 'B4*C4', 'sum(D2:D4)']);
    expect(val(s, 'A6')).toBe(true);
    expect(val(s, 'B6')).toBeInstanceOf(CellError);
    expect(sty(s, 'A1')).toMatchObject({ bold: true, size: 12, font: 'DejaVu Sans', fill: '#ffff00', hAlign: 'center', bb: { style: 'thin' } });
    expect(sty(s, 'C2')).toMatchObject({ numFmt: '0.00', italic: true, color: '#0000ff' });
    expect(sty(s, 'E2').numFmt).toBe('yyyy-mm-dd');
    expect(s.cells.get(K('A7'))?.link).toBe('https://example.com/');
    expect(s.validations[0]).toMatchObject({ type: 'list', formula1: '"Apple,Pear,Plum"', errorMessage: 'Pick a fruit' });
    expect(s.merges).toEqual([{ r1: 7, c1: 0, r2: 7, c2: 2 }]);
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.colW.get(0)).toBe(128);
    expect(s.hiddenCols.has(5)).toBe(true);
    expect(s.rowH.get(0)).toBe(24);
    expect(s.comments.get(K('A1'))?.text).toBe('Header note');
    expect(s.tabColor).toBe('#ff8800');
    expect(s.zoom).toBe(1.25);
    expect(book.names).toContainEqual({ name: 'Rate', ref: 'Fruit!$C$2', sheet: null });
    const n = book.sheets[1];
    expect(val(n, 'A1')).toBe('Tricky <&> "quotes"');
    expect(fml(n, 'A2')).toBe('Fruit!D5*2');
    expect(n.showGrid).toBe(false);
  });
  it('opens a gzipped Gnumeric file with an unrelated name through the gzip container', async () => {
    const { book, format } = await importFile(fixture('sample.gnumeric'), 'download.bin');
    expect(format).toBe('gz');
    expect(val(book.sheets[0], 'A2')).toBe('Apple');
  });
  it('reads the old gmr namespace with Content elements', async () => {
    const xml = '<?xml version="1.0"?><gmr:Workbook xmlns:gmr="http://www.gnome.org/gnumeric/v7"><gmr:Sheets><gmr:Sheet><gmr:Name>Old</gmr:Name><gmr:Cells>'
      + '<gmr:Cell Col="0" Row="0" Style="0"><gmr:Content>Hello</gmr:Content></gmr:Cell><gmr:Cell Col="1" Row="0"><gmr:Content>=1+2</gmr:Content></gmr:Cell>'
      + '</gmr:Cells></gmr:Sheet></gmr:Sheets></gmr:Workbook>';
    const { book } = await importFile(ab(gzipSync(Buffer.from(xml))), 'old.gnumeric');
    expect(val(book.sheets[0], 'A1')).toBe('Hello');
    expect(fml(book.sheets[0], 'B1')).toBe('1+2');
  });
});

describe('OpenOffice.org 1.x Calc (.sxc)', () => {
  it('reads values, formulas, styles, merges, comments, links and settings', async () => {
    const { book, format } = await importFile(fixture('sample.sxc'), 'legacy.sxc');
    expect(format).toBe('sxc');
    expect(book.sheets.map((s) => s.name)).toEqual(['Fruit', 'Notes & More']);
    const s = book.sheets[0];
    expect([val(s, 'A2'), val(s, 'B2'), val(s, 'C2'), val(s, 'D2')]).toEqual(['Apple', 3, 1.25, 3.75]);
    expect([fml(s, 'D2'), fml(s, 'D5')]).toEqual(['B2*C2', 'SUM(D2:D4)']);
    expect(val(s, 'E2')).toBe(45366);
    expect(val(s, 'E3')).toBeCloseTo(13.5 / 24, 9);
    expect(val(s, 'E4')).toBe(true);
    expect(sty(s, 'E3').numFmt).toBe('h:mm:ss');
    expect(val(s, 'E5')).toBe('line one\nline  two');
    expect(sty(s, 'A1')).toMatchObject({ bold: true, fill: '#ffff00', hAlign: 'center', bb: { style: 'thin' } });
    expect(sty(s, 'C2')).toMatchObject({ numFmt: '0.00', italic: true, color: '#0000ff' });
    expect(sty(s, 'E2').numFmt).toBe('yyyy-mm-dd');
    expect(sty(s, 'E5')).toMatchObject({ font: 'Courier New', size: 12, wrap: true });
    expect(s.merges).toEqual([{ r1: 5, c1: 0, r2: 5, c2: 2 }]);
    expect(s.cells.get(K('D6'))?.link).toBe('https://example.com/');
    expect(s.comments.get(K('A1'))).toMatchObject({ author: 'Ann', text: 'Header note' });
    expect(s.hiddenRows.has(6)).toBe(true);
    expect(s.rowH.get(0)).toBe(30);
    expect(s.colW.get(0)).toBe(151);
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    const n = book.sheets[1];
    expect(fml(n, 'A2')).toBe('Fruit!D5*2');
    expect(fml(n, 'A3')).toBe('IF(Fruit!B2>2,"many","few")');
    expect(val(n, 'A1')).toBe('Tricky <&> "quotes"');
    expect(book.names).toContainEqual({ name: 'Rate', ref: 'Fruit!$C$2', sheet: null });
    expect(book.props.title).toBe('Fruit list');
    expect(recalcDiffs(book)).toEqual([]);
  });
});

describe('Excel binary workbook (.xlsb)', () => {
  it('reads cells, formulas, styles, layout, comments, links and names', async () => {
    const { buildXlsb } = await import('./helpers/xlsb-writer');
    const data = await buildXlsb();
    expect(sniffFormat(data, 'anything.zip')?.id).toBe('xlsb');
    const { book, format, warnings } = await importFile(data, 'book.xlsb');
    expect(format).toBe('xlsb');
    expect(warnings, warnings.join()).toEqual([]);
    expect(book.sheets.map((s) => [s.name, s.visibility])).toEqual([['Fruit', 'visible'], ['Notes & More', 'hidden']]);
    const s = book.sheets[0];
    expect([val(s, 'A2'), val(s, 'B2'), val(s, 'C2'), val(s, 'D2'), val(s, 'C4')]).toEqual(['Apple', 3, 1.25, 3.75, 0.75]);
    expect([fml(s, 'D2'), fml(s, 'D3'), fml(s, 'D4')]).toEqual(['B2*C2', 'B3*C3', 'B4*C4']);
    expect(fml(s, 'D5')).toBe('SUM(D2:D4)');
    expect(fml(s, 'E5')).toBe('ROUND(D5,1)');
    expect(fml(s, 'F5')).toBe('IF(B2>2,"many","few")');
    expect(fml(s, 'H2')).toBe("'Notes & More'!A1");
    expect(fml(s, 'H3')).toBe('TaxRate*10');
    expect(fml(s, 'H4')).toBe('SUM({1,2;3,4})');
    expect(fml(s, 'J5')).toBe('SUM(Sales[Units])');
    expect(fml(s, 'C6')).toBe('NA()');
    expect(fml(s, 'D6')).toBe('IFERROR(1/0,"x")');
    expect(s.cells.get(K('A8'))?.arr).toEqual({ r1: 7, c1: 0, r2: 8, c2: 0 });
    expect(s.cells.get(K('A9'))?.sp).toBe(K('A8'));
    expect(val(s, 'A6')).toBe(true);
    expect((val(s, 'B6') as CellError).code).toBe('#DIV/0!');
    expect(val(s, 'G6')).toBe('Bold part plain');
    expect(s.cells.get(K('G6'))?.rich?.[0]).toMatchObject({ t: 'Bold part', s: { bold: true } });
    expect(sty(s, 'A1')).toMatchObject({ bold: true, size: 12, font: 'Arial', color: '#c00000', fill: '#ffff00', hAlign: 'center', bb: { style: 'thin' } });
    expect(sty(s, 'C2')).toMatchObject({ numFmt: '#,##0.000', italic: true, underline: 'single', color: '#0000ff' });
    expect(sty(s, 'E2').numFmt).toBe('m/d/yyyy');
    expect(sty(s, 'H6').locked).toBe(false);
    expect(sty(s, 'I6')).toMatchObject({ wrap: true, vAlign: 'top' });
    expect(s.merges).toEqual([{ r1: 6, r2: 6, c1: 0, c2: 2 }]);
    expect(s.cells.get(K('D7'))).toMatchObject({ link: 'https://example.com/', linkTip: 'Visit' });
    expect(s.comments.get(K('A1'))).toMatchObject({ author: 'Ann', text: 'Header note' });
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.zoom).toBe(1.25);
    expect(s.tabColor).toBe('#ff8800');
    expect(s.colW.get(0)).toBe(140);
    expect(s.hiddenCols.has(5)).toBe(true);
    expect(s.rowH.get(0)).toBe(40);
    expect(s.hiddenRows.has(9)).toBe(true);
    expect(s.tables[0]).toMatchObject({ name: 'Sales', range: { r1: 1, c1: 9, r2: 3, c2: 10 }, columns: [{ name: 'Region' }, { name: 'Units' }], style: 'TableStyleMedium9' });
    expect(s.print).toMatchObject({ orientation: 'landscape', paper: 'A4', scale: 90, area: '$A$1:$D$5', gridlines: true, centerH: true });
    expect(s.print.header).toEqual({ left: 'Left', center: '&[Page] of &[Pages]', right: '' });
    expect(book.names).toEqual([{ name: 'TaxRate', ref: 'Fruit!$C$2', sheet: null }]);
    expect(book.props.title).toBe('Binary fruit');
    expect(val(book.sheets[1], 'A1')).toBe('Tricky <&> "quotes"');
    expect(recalcDiffs(book)).toEqual([]);
  });
});

describe('Excel binary workbook written by Excel (Apache POI sample)', () => {
  it('matches the values LibreOffice shows', async () => {
    const { book, warnings } = await importFile(fixture('third-party/poi-testVarious.xlsb'), 'testVarious.xlsb');
    const s = book.sheets[0];
    expect(s.name).toBe('mySheet1');
    const col = (c: string) => Array.from({ length: 17 }, (_, i) => val(s, `${c}${i + 1}`));
    expect(col('A').slice(0, 5)).toEqual(['String', 'integer', 'float', 'currency', 'percent']);
    expect(col('B')).toEqual(['This is a string', 13, 13.1211231321, 3.03, 0.2, 13.12131231, 123456789012345, 1234567890123450, 0.25, 42803, 'contents', 'tika_link', 4, expect.any(CellError), 0.5, '   46/1963', '  3/128']);
    expect([fml(s, 'B13'), fml(s, 'B14'), fml(s, 'B15')]).toEqual(['C13*2', 'COS(NA)', '1/C13']);
    expect(s.cells.get(K('B12'))?.link).toBe('http://tika.apache.org/');
    expect(s.comments.get(K('B11'))?.text).toMatch(/^test comment/);
    expect(sty(s, 'B5').numFmt).toBe('0%');
    expect(sty(s, 'B10').numFmt).toBe('m/d/yyyy');
    expect(warnings.join(' ')).toMatch(/VBA/);
    expect(s.drawings.length + s.charts.length).toBeGreaterThan(0);
    expect(recalcDiffs(book)).toEqual([]);
  });
});

describe('Apple Numbers', () => {
  it("reads a Numbers '09 document saved by Numbers (values, formulas, formats, merges)", async () => {
    const { book, format, warnings } = await importFile(fixture('third-party/tika-testNumbers.numbers'), 'budget.numbers');
    expect(format).toBe('numbers');
    expect(book.sheets.map((s) => s.name)).toEqual(['Account Categories', 'Transactions', 'Table 1']);
    const cat = book.sheets[0];
    expect([val(cat, 'A2'), val(cat, 'B2'), val(cat, 'A7'), val(cat, 'B7')]).toEqual(['Home', -872.4, 'Total', -2575.9]);
    expect(fml(cat, 'B2')).toBe('SUMIF(Transactions!$D$3:$D$14,A2,Transactions!E$3:E$14)');
    expect(fml(cat, 'B7')).toBe('SUM(B$2:B$6)');
    expect(sty(cat, 'B2').numFmt).toBe('€#,##0.00;(€#,##0.00)');
    const tr = book.sheets[1];
    expect([val(tr, 'E1'), val(tr, 'F1'), val(tr, 'B3'), val(tr, 'C3'), val(tr, 'D3')]).toEqual(['Beginning Balance', 4650, 40087, 'Rent', 'Home']);
    expect(sty(tr, 'B3').numFmt).toBe('m/d/yy');
    expect(tr.merges).toContainEqual({ r1: 0, c1: 0, r2: 0, c2: 1 });
    expect(tr.colW.get(2)).toBe(240);
    expect(sty(tr, 'A2')).toMatchObject({ bold: true, font: 'Cochin' });
    expect(recalcDiffs(book)).toEqual([]);
    expect(warnings.join(' ')).toMatch(/Charts/);
  });
  it("reads index.xml.gz inside the package", async () => {
    const zip = await JSZip.loadAsync(fixture('third-party/tika-testNumbers.numbers'));
    const xml = await zip.file('index.xml')!.async('uint8array');
    const z2 = new JSZip();
    z2.file('index.xml.gz', gzipSync(xml));
    z2.file('buildVersionHistory.plist', '<plist/>');
    const data = await z2.generateAsync({ type: 'arraybuffer' });
    expect(sniffFormat(data, 'x.zip')?.id).toBe('numbers');
    const { book } = await importFile(data, 'x.numbers');
    expect(val(book.sheets[0], 'B7')).toBe(-2575.9);
  });
  it('reads a Numbers 2013 document saved by Numbers (iWork archives)', async () => {
    const { book, format, warnings } = await importFile(fixture('third-party/tika-testNumbers2013.numbers'), 'budget.numbers');
    expect(format).toBe('numbers');
    const names = book.sheets.map((s) => s.name);
    expect(names).toEqual(expect.arrayContaining(['Account Categories', 'Transactions', 'Table 1']));
    const cat = book.sheets[names.indexOf('Account Categories')];
    expect([val(cat, 'A1'), val(cat, 'B2'), val(cat, 'A7'), val(cat, 'B7')]).toEqual(['Category', -872.4, 'Total', -2575.9]);
    const tr = book.sheets[names.indexOf('Transactions')];
    expect([val(tr, 'E1'), val(tr, 'F1'), val(tr, 'C3'), val(tr, 'E3'), val(tr, 'F14')]).toEqual(['Beginning Balance', 4650, 'Rent', -775, 4279.1]);
    expect(Math.floor(val(tr, 'B3') as number)).toBe(40087);
    expect(sty(tr, 'B3').numFmt).toBeTruthy();
    const t1 = book.sheets[names.indexOf('Table 1')];
    expect([val(t1, 'B2'), val(t1, 'C3'), val(t1, 'D4')]).toEqual(['Test', 0.5, 0.6]);
    expect(tr.colW.get(2)).toBe(240);
    expect(warnings.join(' ')).toMatch(/last calculated values/);
  });
  it('reads version 5 cell storage (decimal128, dates, booleans, durations)', async () => {
    const { buildNumbers } = await import('./helpers/iwa-writer');
    const data = await buildNumbers();
    expect(sniffFormat(data, 'download')?.id).toBe('numbers');
    const { book } = await importFile(data, 'stand.numbers');
    const s = book.sheets[0];
    expect(s.name).toBe('Fruit stand');
    expect([val(s, 'A2'), val(s, 'B2'), val(s, 'C2'), val(s, 'D2'), val(s, 'C3'), val(s, 'D3')]).toEqual(['Apples', 3, 1.25, 3.75, -2.5, -12.5]);
    expect([val(s, 'B4'), val(s, 'C4'), val(s, 'D4')]).toEqual([45366, true, 0.0625]);
    expect(sty(s, 'B4').numFmt).toBe('yyyy-mm-dd');
    expect(s.colW.get(0)).toBe(160);
    expect(s.hiddenCols.has(3)).toBe(true);
  });
  it('opens the preview when that is all the package holds', async () => {
    const z = new JSZip();
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 3, 8, 2, 0, 0, 0]);
    z.file('preview.jpg', png);
    z.file('Metadata/DocumentIdentifier', 'x');
    const { book, warnings } = await importFile(await z.generateAsync({ type: 'arraybuffer' }), 'locked.numbers');
    expect(book.sheets[0].drawings[0].kind).toBe('image');
    expect(warnings.join(' ')).toMatch(/only a preview/);
  });
});

describe('KSpread', () => {
  for (const f of ['sample.ksp', 'sample-old.ksp']) {
    it(`reads ${f}`, async () => {
      const data = fixture(f);
      expect(sniffFormat(data, f)?.id).toBe('ksp');
      const { book, format } = await importFile(data, f);
      expect(format).toBe('ksp');
      expect(book.sheets.map((s) => [s.name, s.visibility])).toEqual([['Fruit', 'visible'], ['Hidden', 'hidden']]);
      const s = book.sheets[0];
      expect([val(s, 'A2'), val(s, 'B2'), val(s, 'C2'), val(s, 'D2'), val(s, 'E2'), val(s, 'E3')]).toEqual(['Apple', 3, 1.25, 3.75, 45366, true]);
      expect([fml(s, 'D2'), fml(s, 'D3'), fml(s, 'D4'), fml(s, 'E4')]).toEqual(['B2*C2', 'B3*C3', 'SUM(D2:D3)', 'IF($B$2>2,"many","few")']);
      expect(sty(s, 'A1')).toMatchObject({ bold: true, size: 12, font: 'Helvetica', fill: '#ffff00', hAlign: 'center', bb: { style: 'thin' } });
      expect(sty(s, 'C2')).toMatchObject({ numFmt: '0.00', italic: true });
      expect(sty(s, 'E2').numFmt).toBe('yyyy-mm-dd');
      expect(s.merges).toEqual([{ r1: 4, c1: 0, r2: 4, c2: 2 }]);
      expect(s.comments.get(K('A1'))?.text).toBe('Header note');
      expect(s.colW.get(0)).toBe(120);
      expect(book.props).toMatchObject({ title: 'Fruit list', author: 'Ann', company: 'Fruit Co' });
      expect(recalcDiffs(book)).toEqual([]);
    });
  }
});

describe('Uniform Office Format (.uos)', () => {
  it('reads values, formulas, styles, merges and links', async () => {
    const data = fixture('sample.uos');
    expect(sniffFormat(data, 'book.xml')?.id).toBe('uos');
    const { book, format } = await importFile(data, 'book.uos');
    expect(format).toBe('uos');
    expect(book.sheets.map((s) => s.name)).toEqual(['Fruit', 'Notes & More']);
    const s = book.sheets[0];
    expect([val(s, 'A2'), val(s, 'B2'), val(s, 'C2'), val(s, 'D2'), val(s, 'E2')]).toEqual(['Apple', 3, 1.25, 3.75, 45366]);
    expect(val(s, 'E3')).toBeCloseTo(13.5 / 24, 9);
    expect([fml(s, 'D2'), fml(s, 'D5')]).toEqual(['B2*C2', 'SUM(D2:D4)']);
    expect(sty(s, 'A1')).toMatchObject({ bold: true, fill: '#ffff00', hAlign: 'center', bb: { style: 'thin' } });
    expect(sty(s, 'E2').numFmt).toBe('mm"/"dd"/"yyyy');
    expect(sty(s, 'E5')).toMatchObject({ font: 'Courier New', size: 12, wrap: true });
    expect(s.merges).toEqual([{ r1: 5, c1: 0, r2: 5, c2: 2 }]);
    expect(s.cells.get(K('D6'))?.link).toBe('https://example.com/');
    expect(s.hiddenRows.has(6)).toBe(true);
    expect(val(book.sheets[1], 'A1')).toBe('Tricky <&> "quotes"');
  });
});

describe('MIME HTML web archives (.mht)', () => {
  it('reads every sheet part of an Excel single file web page', async () => {
    const sheet = (rows: string) => `<html><body><table>${rows}</table></body></html>`;
    const qp = (s: string) => s.replace(/=/g, '=3D').replace(/<tr>/g, '=\r\n<tr>');
    const mht = [
      'MIME-Version: 1.0', 'X-Document-Type: Workbook', 'Content-Type: multipart/related; boundary="----=_NextPart_01"', '',
      'This document is a Single File Web Page.', '',
      '------=_NextPart_01', 'Content-Location: file:///C:/x/book.htm', 'Content-Transfer-Encoding: quoted-printable', 'Content-Type: text/html; charset="us-ascii"', '',
      qp('<html><frameset><frame src="book_files/sheet001.htm"></frameset></html>'), '',
      '------=_NextPart_01', 'Content-Location: file:///C:/x/book_files/tabstrip.htm', 'Content-Transfer-Encoding: quoted-printable', 'Content-Type: text/html; charset="us-ascii"', '',
      qp('<html><body><a href="sheet001.htm" target="frSheet"><font face="Arial">Sales</font></a> <a href="sheet002.htm" target="frSheet"><font face="Arial">Costs</font></a></body></html>'), '',
      '------=_NextPart_01', 'Content-Location: file:///C:/x/book_files/sheet001.htm', 'Content-Transfer-Encoding: quoted-printable', 'Content-Type: text/html; charset="windows-1252"', '',
      qp(sheet('<tr><td style="font-weight:700">Item</td><td>Qty</td></tr><tr><td>Caf=E9</td><td x:num="3">3</td></tr>')).replace('Caf=3DE9', 'Caf=E9'), '',
      '------=_NextPart_01', 'Content-Location: file:///C:/x/book_files/sheet002.htm', 'Content-Transfer-Encoding: base64', 'Content-Type: text/html; charset="utf-8"', '',
      btoa(sheet('<tr><td>Rent</td><td>775</td></tr>')), '',
      '------=_NextPart_01--', '',
    ].join('\r\n');
    const data = txt(mht);
    expect(sniffFormat(data, 'book.bin')?.id).toBe('mht');
    const { book } = await importFile(data, 'book.mht');
    expect(book.sheets.map((s) => s.name)).toEqual(['Sales', 'Costs']);
    expect([val(book.sheets[0], 'A2'), val(book.sheets[0], 'B2'), val(book.sheets[1], 'B1')]).toEqual(['Café', 3, 775]);
    expect(sty(book.sheets[0], 'A1').bold).toBe(true);
  });
});

/** Rewrites a transitional .xlsx package with the ISO/IEC 29500 Strict namespaces. */
async function toStrict(data: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(data);
  const map: [string, string][] = [
    ['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships'],
    ['http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing'],
    ['http://schemas.openxmlformats.org/drawingml/2006/chart', 'http://purl.oclc.org/ooxml/drawingml/chart'],
    ['http://schemas.openxmlformats.org/drawingml/2006/main', 'http://purl.oclc.org/ooxml/drawingml/main'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/extended-properties', 'http://purl.oclc.org/ooxml/officeDocument/extendedProperties'],
  ];
  const out = new JSZip();
  for (const name of Object.keys(zip.files)) {
    const f = zip.files[name];
    if (f.dir) continue;
    if (/\.(xml|rels)$/.test(name)) {
      let t = await f.async('string');
      for (const [a, b] of map) t = t.split(a).join(b);
      if (name === 'xl/workbook.xml') t = t.replace('<workbook ', '<workbook conformance="strict" ');
      out.file(name, t);
    } else out.file(name, await f.async('uint8array'));
  }
  return out.generateAsync({ type: 'arraybuffer' });
}

describe('Excel Strict (ISO/IEC 29500 Strict) workbooks', () => {
  it('reads cells, styles, formulas, names and charts in the Strict namespaces', async () => {
    const { Workbook } = await import('../../src/model/workbook');
    const { exportXlsx } = await import('../../src/io/xlsx-export');
    const b = new Workbook();
    const s = b.sheets[0];
    s.cells.set(K('A1'), { v: 'Item', s: b.styleId({ bold: true, fill: '#ffff00' }) });
    s.cells.set(K('B1'), { v: 'Qty' });
    s.cells.set(K('A2'), { v: 'Apple' }); s.cells.set(K('B2'), { v: 3 });
    s.cells.set(K('A3'), { v: 'Pear' }); s.cells.set(K('B3'), { v: 5 });
    s.cells.set(K('B4'), { f: 'SUM(B2:B3)', v: 8 });
    s.cells.set(K('C1'), { v: 'site', link: 'https://example.com/' });
    s.merges.push({ r1: 5, c1: 0, r2: 5, c2: 1 });
    s.charts.push({ id: 'c1', type: 'column', title: 'Qty', series: [{ name: 'Qty', values: 'Sheet1!$B$2:$B$3', categories: 'Sheet1!$A$2:$A$3' }], at: { r: 6, c: 0, dx: 0, dy: 0, w: 300, h: 200 } });
    b.names.push({ name: 'Total', ref: 'Sheet1!$B$4', sheet: null });
    const strict = await toStrict(await (await exportXlsx(b)).arrayBuffer());
    expect(sniffFormat(strict, 'strict.xlsx')?.id).toBe('xlsx');
    const { book, warnings } = await importFile(strict, 'strict.xlsx');
    const t = book.sheets[0];
    expect([val(t, 'A2'), val(t, 'B3'), fml(t, 'B4'), val(t, 'B4')]).toEqual(['Apple', 5, 'SUM(B2:B3)', 8]);
    expect(sty(t, 'A1')).toMatchObject({ bold: true, fill: '#ffff00' });
    expect(t.cells.get(K('C1'))?.link).toBe('https://example.com/');
    expect(t.merges).toEqual([{ r1: 5, c1: 0, r2: 5, c2: 1 }]);
    expect(t.charts.length, warnings.join()).toBe(1);
    expect(t.charts[0]).toMatchObject({ type: 'column', title: 'Qty' });
    expect(book.names).toContainEqual({ name: 'Total', ref: 'Sheet1!$B$4', sheet: null });
  });
});

describe('OpenDocument and .xlsx fidelity (files written by LibreOffice)', () => {
  const check = (book: Workbook, from: string) => {
    const s = book.sheets[0];
    expect(s.name).toBe('Data');
    const dv = s.validations.find((v) => v.type === 'list');
    expect(dv, from).toMatchObject({ ranges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }], formula1: '"Apple,Pear,Plum"', errorTitle: 'Oops', errorMessage: 'Pick a fruit', inputMessage: 'Choose a fruit', allowBlank: true });
    expect(s.validations.find((v) => v.type === 'whole'), from).toMatchObject({ operator: 'between', formula1: '1', formula2: '100', errorStyle: 'warning', ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }] });
    const gt = s.cf.find((r) => r.type === 'cellIs' && r.operator === 'greaterThan');
    expect(gt, from).toMatchObject({ formulas: ['10'], ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], style: { fill: '#ffcccc', color: '#cc0000' } });
    expect(s.cf.find((r) => r.type === 'colorScale'), from).toMatchObject({ colors: ['#f8696b', '#63be7b'], stops: [{ type: 'min' }, { type: 'max' }] });
    expect(s.cf.find((r) => r.type === 'beginsWith'), from).toMatchObject({ text: 'P' });
    expect(s.charts.length, from).toBe(1);
    expect(s.charts[0], from).toMatchObject({ type: 'column', title: 'Fruit qty', legend: 'bottom', gridlines: true });
    expect(s.charts[0].series[0].values.replace(/\$/g, '')).toBe('Data!B2:B4');
    expect(s.charts[0].series[0].categories?.replace(/\$/g, '')).toBe('Data!A2:A4');
    expect(recalcDiffs(book)).toEqual([]);
  };
  it('reads validity, conditional formats and charts from .ods', async () => {
    check((await importFile(fixture('lo-rich.ods'), 'rich.ods')).book, 'ods');
  });
  it('reads them from flat .fods with an inline chart', async () => {
    check((await importFile(fixture('rich.fods'), 'rich.fods')).book, 'fods');
  });
  it('turns ODF style:map conditions into conditional formats when there are no calcext rules', async () => {
    const fods = new TextDecoder().decode(new Uint8Array(fixture('rich.fods'))).replace(/<calcext:conditional-formats>[\s\S]*<\/calcext:conditional-formats>/, '');
    const s = (await importFile(txt(fods), 'maps.fods')).book.sheets[0];
    expect(s.cf).toHaveLength(1);
    expect(s.cf[0]).toMatchObject({ type: 'cellIs', operator: 'greaterThanOrEqual', formulas: ['5'], ranges: [{ r1: 1, c1: 3, r2: 1, c2: 3 }], style: { fill: '#ccffcc', color: '#006600' } });
  });
  it('reads them from the .xlsx LibreOffice wrote', async () => {
    check((await importFile(fixture('lo-rich.xlsx'), 'rich.xlsx')).book, 'xlsx');
  });
});

describe('.xlsx fidelity details', () => {
  it('reads Excel 2010 conditional formats, hyperlink tips and header/footer variants', async () => {
    const { Workbook } = await import('../../src/model/workbook');
    const { exportXlsx } = await import('../../src/io/xlsx-export');
    const b = new Workbook();
    const s0 = b.sheets[0];
    for (let r = 1; r <= 5; r++) s0.cells.set(K(`A${r}`), { v: r * 10 });
    s0.cells.set(K('B1'), { v: 'site', link: 'https://example.com/' });
    s0.cells.set(K('C1'), { v: 'x' });
    const zip = await JSZip.loadAsync(await (await exportXlsx(b)).arrayBuffer());
    let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';
    const main = '<conditionalFormatting sqref="A1:A5"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>'
      + `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ${X14}><x14:id>{11111111-0000-0000-0000-000000000001}</x14:id></ext></extLst></cfRule></conditionalFormatting>`;
    const ext = `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${X14}><x14:conditionalFormattings>`
      + '<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{11111111-0000-0000-0000-000000000001}"><x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>'
      + '<x14:conditionalFormatting><x14:cfRule type="cellIs" priority="2" operator="greaterThan" id="{2}"><xm:f>Sheet1!$A$3</xm:f><x14:dxf><font><b/></font><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>'
      + '<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="3" id="{3}"><x14:iconSet iconSet="3Triangles"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>'
      + '</x14:conditionalFormattings></ext></extLst>';
    xml = xml.replace(/(<\/sheetData>(?:<mergeCells[\s\S]*?<\/mergeCells>)?)/, `$1${main}`);
    xml = xml.replace(/<hyperlink /, '<hyperlink tooltip="Open the site" ');
    xml = xml.replace(/<headerFooter[\s\S]*?<\/headerFooter>|<headerFooter[^>]*\/>/, '');
    xml = xml.replace(/<\/worksheet>$/, '') + '<headerFooter differentOddEven="1" differentFirst="1"><oddHeader>&amp;CPage &amp;P</oddHeader><evenHeader>&amp;LEven</evenHeader><firstHeader>&amp;RFirst</firstHeader></headerFooter>' + ext + '</worksheet>';
    xml = xml.replace(/(<headerFooter[\s\S]*<\/headerFooter>)([\s\S]*)(<drawing[^>]*\/>)/, '$1$3$2');
    zip.file('xl/worksheets/sheet1.xml', xml);
    const { book } = await importFile(await zip.generateAsync({ type: 'arraybuffer' }), 'x14.xlsx');
    const s = book.sheets[0];
    const bar = s.cf.find((r) => r.type === 'dataBar')!;
    expect(bar).toMatchObject({ barColor: '#638ec6', gradientBar: false });
    expect(s.cf.filter((r) => r.type === 'dataBar')).toHaveLength(1);
    expect(s.cf.find((r) => r.type === 'cellIs')).toMatchObject({ operator: 'greaterThan', formulas: ['Sheet1!$A$3'], style: { bold: true, fill: '#ffff00' } });
    expect(s.cf.find((r) => r.type === 'iconSet')).toMatchObject({ iconSet: '3Triangles', stops: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] });
    expect(s.cells.get(K('B1'))).toMatchObject({ link: 'https://example.com/', linkTip: 'Open the site' });
    expect(s.print).toMatchObject({ differentOddEven: true, differentFirst: true, header: { center: 'Page &[Page]' }, evenHeader: { left: 'Even' }, firstHeader: { right: 'First' } });
    expect(s.cf.some((r) => 'extId' in r)).toBe(false);
  });
});

describe('HTML tables from spreadsheet programs', () => {
  it('reads LibreOffice HTML: sdval/sdnum, sheet data attributes, font wrappers, comments, sheet names', async () => {
    const { book } = await importFile(fixture('lo-fruit.html'), 'fruit.html');
    expect(book.sheets.map((s) => s.name)).toEqual(['Fruit', 'Notes & More']);
    const s = book.sheets[0];
    expect([val(s, 'A2'), val(s, 'B2'), val(s, 'C2'), val(s, 'D2'), val(s, 'E2')]).toEqual(['Apple', 3, 1.25, 3.75, 45366]);
    expect(fml(s, 'D2')).toBe('B2*C2');
    expect(sty(s, 'C2')).toMatchObject({ numFmt: '0.00', italic: true, color: '#0000ff' });
    expect(sty(s, 'E2').numFmt).toBe('mm/dd/yyyy');
    expect(sty(s, 'A1')).toMatchObject({ bold: true, fill: '#ffff00', hAlign: 'center', bb: { style: 'thin' } });
    expect(s.comments.get(K('A1'))?.text).toBe('Header note');
    expect(val(s, 'A1')).toBe('Name');
    expect(s.colW.get(0)).toBe(151);
    expect(s.merges).toContainEqual({ r1: 5, c1: 0, r2: 5, c2: 2 });
    expect(fml(book.sheets[1], 'A2')).toBe('Fruit!D5*2');
  });
  it('reads Excel HTML: class styles, mso-number-format, x:num, x:fmla, x:bool', async () => {
    const html = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><style>
      <!-- td {mso-number-format:General; font-family:Calibri}
      .xl65 {font-weight:700; background:#FFFF00; mso-pattern:black none; text-align:center; border-bottom:.5pt solid windowtext}
      .xl66 {mso-number-format:"0\\.00"; font-style:italic; color:blue}
      .xl67 {mso-number-format:"yyyy\\-mm\\-dd"} --></style></head><body><table>
      <col width=120><col width=64 span=2>
      <tr><td class=xl65>Item</td><td class=xl65>Price</td><td class=xl65>Date</td></tr>
      <tr><td>Apple</td><td class=xl66 x:num="1.25">1.25</td><td class=xl67 x:num="45366">2024-03-15</td></tr>
      <tr><td>Total</td><td class=xl66 x:num="1.25" x:fmla="=SUM(B2)">1.25</td><td x:bool="TRUE">TRUE</td></tr>
      </table></body></html>`;
    const { book } = await importFile(txt(html), 'excel.htm');
    const s = book.sheets[0];
    expect(sty(s, 'A1')).toMatchObject({ bold: true, fill: '#ffff00', hAlign: 'center', bb: { style: 'thin' } });
    expect(sty(s, 'B2')).toMatchObject({ numFmt: '0.00', italic: true, color: '#0000ff' });
    expect(sty(s, 'C2').numFmt).toBe('yyyy-mm-dd');
    expect([val(s, 'B2'), val(s, 'C2'), fml(s, 'B3'), val(s, 'C3')]).toEqual([1.25, 45366, 'SUM(B2)', true]);
    expect([s.colW.get(0), s.colW.get(2)]).toEqual([120, 64]);
  });
  it('reads Google Sheets HTML copies (data-sheets-value, R1C1 formulas)', async () => {
    const html = '<google-sheets-html-origin><table><tr><td data-sheets-value=\'{"1":3,"3":4}\'>4</td><td data-sheets-value=\'{"1":3,"3":8}\' data-sheets-formula="=R[0]C[-1]*2">8</td><td data-sheets-value=\'{"1":2,"2":"007"}\'>007</td><td data-sheets-value=\'{"1":4,"4":1}\'>TRUE</td></tr></table>';
    const { book } = await importFile(txt(html), 'clip.html');
    const s = book.sheets[0];
    expect([val(s, 'A1'), val(s, 'B1'), fml(s, 'B1'), val(s, 'C1'), val(s, 'D1')]).toEqual([4, 8, 'A1*2', '007', true]);
  });
});

describe('Legacy interchange formats written by other programs', () => {
  const fruit = (s: Sheet) => {
    expect([val(s, 'A2'), val(s, 'B2'), val(s, 'C2'), val(s, 'D2'), val(s, 'E2')]).toEqual(['Apple', 3, 1.25, 3.75, 45366]);
    expect(val(s, 'E5')).toBe('line one\nline  two');
  };
  it('reads LibreOffice XML Spreadsheet 2003 (ODF formulas)', async () => {
    const { book } = await importFile(fixture('lo-fruit-2003.xml'), 'fruit.xml');
    const s = book.sheets[0];
    fruit(s);
    expect([fml(s, 'D2'), fml(s, 'D5')]).toEqual(['B2*C2', 'SUM(D2:D4)']);
    expect(fml(book.sheets[1], 'A3')).toBe('IF(Fruit!B2>2,"many","few")');
    expect(recalcDiffs(book)).toEqual([]);
  });
  it('reads LibreOffice SYLK (A1 formulas, line breaks) and DIF (dates as text, multi-line strings)', async () => {
    const slk = (await importFile(fixture('lo-fruit.slk'), 'fruit.slk')).book;
    fruit(slk.sheets[0]);
    expect(fml(slk.sheets[0], 'D5')).toBe('SUM(D2:D4)');
    const dif = (await importFile(fixture('lo-fruit.dif'), 'fruit.dif')).book;
    expect(dif.sheets[0].name).toBe('Fruit');
    fruit(dif.sheets[0]);
  });
  it('reads Excel SYLK formats, fonts, borders and accented letters', async () => {
    const t = ['ID;PWXL;N;E', 'P;PGeneral', 'P;P0.00', 'P;EArial;M200', 'P;EArial;M240;SB', 'F;P0;DG0G8;M255', 'F;W1 1 20', 'B;Y3;X2;D0 0 2 1',
      'C;Y1;X1;K"Caf\x1bNBe"', 'F;P1;FF2R;SDM2TB;Y2;X1', 'C;Y2;X1;K1.5', 'C;Y2;X2;K3;ERC[-1]*2+R1C1', 'C;Y3;X1;K"a\x1b :b"', 'E'].join('\r\n');
    const { book } = await importFile(txt(t), 'x.slk');
    const s = book.sheets[0];
    expect(val(s, 'A1')).toBe('Café');
    expect(sty(s, 'A2')).toMatchObject({ numFmt: '0.00', hAlign: 'right', bold: true, size: 12, bt: { style: 'thin' }, bb: { style: 'thin' } });
    expect(fml(s, 'B2')).toBe('A2*2+$A$1');
    expect(val(s, 'A3')).toBe('a\nb');
    expect(s.colW.get(0)).toBe(145);
  });
  it('reads Excel XML Spreadsheet 2003 conditional formats, validation and screen tips', async () => {
    const x = `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel">
<Worksheet ss:Name="S"><Table><Row><Cell ss:HRef="https://example.com/" x:HRefScreenTip="Go"><Data ss:Type="String">site</Data></Cell><Cell><Data ss:Type="Number">5</Data></Cell></Row>
<Row><Cell><Data ss:Type="String">b</Data></Cell><Cell><Data ss:Type="Number">20</Data></Cell></Row></Table>
<ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel"><Range>R1C2:R2C2</Range><Condition><Qualifier>Greater</Qualifier><Value1>10</Value1><Format Style='color:#9C0006;background:#FFC7CE;font-weight:700'/></Condition>
<Condition><Value1>RC=5</Value1><Format Style='font-style:italic'/></Condition></ConditionalFormatting>
<DataValidation xmlns="urn:schemas-microsoft-com:office:excel"><Range>R1C2:R2C2</Range><Type>Whole</Type><Min>1</Min><Max>100</Max><ErrorStyle>Warn</ErrorStyle><ErrorMessage>1-100</ErrorMessage></DataValidation>
<DataValidation xmlns="urn:schemas-microsoft-com:office:excel"><Range>R1C1:R2C1</Range><Type>List</Type><Value>&quot;a,b&quot;</Value><InputMessage>Pick</InputMessage></DataValidation>
</Worksheet></Workbook>`;
    const s = (await importFile(txt(x), 'x.xml')).book.sheets[0];
    expect(s.cells.get(K('A1'))).toMatchObject({ link: 'https://example.com/', linkTip: 'Go' });
    expect(s.cf[0]).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulas: ['10'], ranges: [{ r1: 0, c1: 1, r2: 1, c2: 1 }], style: { color: '#9c0006', fill: '#ffc7ce', bold: true } });
    expect(s.cf[1]).toMatchObject({ type: 'expression', formulas: ['B1=5'], style: { italic: true } });
    expect(s.validations[0]).toMatchObject({ type: 'whole', operator: 'between', formula1: '1', formula2: '100', errorStyle: 'warning', errorMessage: '1-100' });
    expect(s.validations[1]).toMatchObject({ type: 'list', formula1: '"a,b"', inputMessage: 'Pick' });
  });
});

describe('recalculation check used above', () => {
  it('reports a stored result that the formula does not produce', async () => {
    const { Workbook } = await import('../../src/model/workbook');
    const b = new Workbook();
    b.sheets[0].cells.set(K('A1'), { v: 2 });
    b.sheets[0].cells.set(K('B1'), { f: 'A1*2', v: 5 });
    expect(recalcDiffs(b)).toHaveLength(1);
  });
});
