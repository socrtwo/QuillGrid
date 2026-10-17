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
    for (const r of SNIFF_RULES_B) expect(formatById(r.id), r.id).toBeTruthy();
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
  it('still opens QuillGrid JSON natively', async () => {
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
    expect(s.tables[0]).toMatchObject({ name: 'Sales', range: { r1: 0, c1: 9, r2: 2, c2: 10 }, columns: [{ name: 'Region' }, { name: 'Units' }], style: 'TableStyleMedium9' });
    expect(s.print).toMatchObject({ orientation: 'landscape', paper: 'A4', scale: 90, area: '$A$1:$D$5', gridlines: true, centerH: true });
    expect(s.print.header).toEqual({ left: 'Left', center: '&[Page] of &[Pages]', right: '' });
    expect(book.names).toEqual([{ name: 'TaxRate', ref: 'Fruit!$C$2', sheet: null }]);
    expect(book.props.title).toBe('Binary fruit');
    expect(val(book.sheets[1], 'A1')).toBe('Tricky <&> "quotes"');
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
  });
});
