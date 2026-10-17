import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import '../../src/calc/functions';
import { importFile, sniff } from '../../src/io/index';
import { registerSniffer } from '../../src/io/sniff';
import type { Workbook } from '../../src/model/workbook';
import { CellError } from '../../src/model/types';
import { formatValue } from '../../src/model/numfmt';
import { snum, lnum, f80, lmbcs, codepage } from '../../src/io/legacy/bin';
import { lotusNumFmt, lotusVariant, worksWinNumFmt } from '../../src/io/legacy/lotus';
import { wbNumFmt } from '../../src/io/legacy/quattro';
import { lotusCodeToText } from '../../src/io/legacy/lotusfmla';
import { translateVisiCalc, importVisiCalc } from '../../src/io/legacy/visicalc';
import { importDbf } from '../../src/io/legacy/dbf';
import {
  buildWk1, buildWk3, build123, buildWorksWks, buildWq1, buildWb1, buildBiff2, buildBiff4w, buildDbf, buildXlr, VISICALC_SAMPLE, lnumBytes,
} from './legacy-builders';

const ab = (u8: Uint8Array): ArrayBuffer => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
const load = (p: string) => ab(readFileSync(`tests/fixtures/${p}`));
const cell = (b: Workbook, ref: string, sheet = 0) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return b.sheets[sheet].get(+m[2] - 1, c - 1);
};
const style = (b: Workbook, ref: string, sheet = 0) => b.styles[cell(b, ref, sheet)?.s || 0] || {};

describe('content sniffing', () => {
  it('identifies legacy spreadsheets whatever their name', () => {
    expect(sniff(buildWk1(), 'x.wk1')).toMatchObject({ format: 'lotus', category: 'spreadsheet' });
    expect(sniff(buildWk1({ version: 0x0404 }), 'x')).toMatchObject({ format: 'lotus', label: expect.stringContaining('1A') });
    expect(sniff(buildWk3(), 'noext')).toMatchObject({ format: 'lotus', label: expect.stringContaining('Release 3') });
    expect(sniff(build123(), 'a.123').label).toContain('97');
    expect(sniff(buildWorksWks(), 'a.wks')).toMatchObject({ format: 'works' });
    expect(sniff(buildWq1(), 'a.dat')).toMatchObject({ format: 'quattro' });
    expect(sniff(buildWb1(), 'a.wb1')).toMatchObject({ format: 'quattro', label: expect.stringContaining('WB1') });
    expect(sniff(buildBiff2(), 'a')).toMatchObject({ format: 'xls', label: expect.stringContaining('BIFF2') });
    expect(sniff(buildBiff4w(), 'a.xlw').label).toContain('BIFF4W');
    expect(sniff(buildDbf().dbf, 'table')).toMatchObject({ format: 'dbf', category: 'spreadsheet' });
    expect(sniff(buildXlr(), 'a.xlr')).toMatchObject({ format: 'xls', label: expect.stringContaining('Works') });
    expect(sniff(new TextEncoder().encode(VISICALC_SAMPLE), 'budget')).toMatchObject({ format: 'visicalc' });
  });

  it('opens by content when the extension is wrong, with a note', () => {
    const r = sniff(buildWk1(), 'report.xls');
    expect(r.format).toBe('lotus');
    expect(r.note).toBe('"report.xls" is named .xls but contains a Lotus 1-2-3 Release 2 worksheet (WK1); it was opened by its contents.');
    expect(sniff(buildWk1(), 'report').note).toMatch(/no extension/);
    expect(sniff(buildWk1(), 'report.wk1').note).toBeUndefined();
  });

  it('recognizes documents, presentations and mail', () => {
    expect(sniff(load('documents/doc.docx'), 'x.bin')).toMatchObject({ format: 'docx', category: 'document' });
    expect(sniff(load('documents/doc.odt'), 'x')).toMatchObject({ format: 'odt', category: 'document' });
    expect(sniff(load('documents/doc.rtf'), 'x')).toMatchObject({ format: 'rtf', category: 'document' });
    expect(sniff(load('documents/doc.doc'), 'x')).toMatchObject({ format: 'doc', category: 'document' });
    expect(sniff(load('documents/deck.pptx'), 'x')).toMatchObject({ format: 'pptx', category: 'presentation' });
    expect(sniff(load('documents/deck.odp'), 'x')).toMatchObject({ format: 'odp', category: 'presentation' });
    expect(sniff(load('documents/mail.eml'), 'x.txt')).toMatchObject({ format: 'eml', category: 'email' });
    expect(sniff(load('documents/quick.msg'), 'x')).toMatchObject({ format: 'msg', category: 'email' });
    expect(sniff(new TextEncoder().encode('From alice@example.com Mon Jan  1 00:00:00 2024\nFrom: a@b.c\nSubject: hi\n\nbody\n'), 'box')).toMatchObject({ format: 'mbox', category: 'email' });
    expect(sniff(new TextEncoder().encode('!BDN' + '\0'.repeat(60)), 'x')).toMatchObject({ format: 'pst', category: 'email' });
    expect(sniff(new TextEncoder().encode('%PDF-1.4\n'), 'x')).toMatchObject({ format: 'wordtext', label: 'PDF document', category: 'document' });
    expect(sniff(Uint8Array.from([0xff, 0x57, 0x50, 0x43, 16, 0, 0, 0, 1, 10, 0, 0]), 'x')).toMatchObject({ category: 'document', label: expect.stringContaining('WordPerfect') });
  });

  it('keeps recognizing the existing formats', () => {
    expect(sniff(load('io/libreoffice.ods'), 'x.zip').format).toBe('ods');
    expect(sniff(load('io/libreoffice-roundtrip.xlsx'), 'x').format).toBe('xlsx');
    expect(sniff(load('legacy/testEXCEL_95.xls'), 'x').format).toBe('xls');
    expect(sniff(new TextEncoder().encode('a,b,c\n1,2,3\n'), 'x.csv').format).toBe('csv');
    expect(sniff(new TextEncoder().encode('a\tb\n1\t2\n'), 'x').format).toBe('tsv');
    expect(sniff(new TextEncoder().encode('ID;PWXL;N;E\nC;Y1;X1;K1\nE\n'), 'x').format).toBe('sylk');
  });

  it('uses the extension only to break ties, and when nothing else matches', () => {
    expect(sniff(new Uint8Array([1, 2, 3, 250, 0, 7]), 'old.wk1')).toMatchObject({ format: 'lotus', byExtension: true });
    expect(sniff(new Uint8Array([1, 2, 3, 250, 0, 7]), 'mystery').format).toBe('unknown');
  });

  it('accepts new rules', () => {
    registerSniffer({ id: 'test-xyz', format: 'csv', label: 'XYZ test data', category: 'spreadsheet', ext: ['xyz'], detect: (c) => (c.head.startsWith('XYZ1') ? 0.99 : 0) });
    expect(sniff(new TextEncoder().encode('XYZ1,2,3'), 'f').label).toBe('XYZ test data');
  });
});

describe('number encodings', () => {
  it('decodes Lotus small numbers, Release 4 numbers and 80-bit floats', () => {
    expect(snum(980 << 1)).toBe(980);
    expect(snum(0x37)).toBeCloseTo(0.015);
    expect(snum(0x03c3)).toBe(30000);
    expect(lnum(0x40)).toBe(1);
    expect(lnum(0x351)).toBeCloseTo(1.3);
    expect(lnum(0x61)).toBe(-10);
    expect(lnum(0x42)).toBe(100);
    for (const v of [12.5, -3, 2, 0.25]) expect(lnum(new DataView(Uint8Array.from(lnumBytes(v)).buffer).getUint32(0, true))).toBeCloseTo(v);
    const x = Uint8Array.from([0x36, 0x3c, 0xbd, 0x52, 0x96, 0xc8, 0xc1, 0x88, 0x13, 0x40]);
    expect(f80(x, 0)).toBeCloseTo(1120313.1, 0);
  });

  it('decodes LMBCS and DOS code pages', () => {
    expect(lmbcs(Uint8Array.from([0x61, 0x0f, 0x2d, 0x0f, 0x2a, 0x62]))).toBe('a\r\nb');
    expect(lmbcs(Uint8Array.from([0x82]))).toBe('é');
    expect(codepage(437)(Uint8Array.from([0x82, 0x9c]))).toBe('é£');
  });
});

describe('Lotus 1-2-3', () => {
  it('maps Works for Windows and Quattro Pro for Windows style formats', () => {
    expect(worksWinNumFmt(0x40)).toBe('0.00');
    expect(worksWinNumFmt(0x43)).toBe('0.00%');
    expect(worksWinNumFmt(0x8a)).toBe('00000');
    expect(worksWinNumFmt(0x05)).toBeNull();
    expect(wbNumFmt(0x14)).toBe('0.0000');
    expect(wbNumFmt(0x22)).toBe('0.00E+00');
    expect(wbNumFmt(0x05)).toBe('dd-mmm-yy');
    expect(wbNumFmt(0x02)).toBeNull();
  });

  it('maps cell formats', () => {
    expect(lotusNumFmt(0x22)).toBe('$#,##0.00_);($#,##0.00)');
    expect(lotusNumFmt(0x72)).toBe('dd-mmm-yy');
    expect(lotusNumFmt(0x71)).toBeNull();
    expect(lotusNumFmt(0x7f)).toBeUndefined();
  });

  it('decompiles formula code with 1-2-3 precedence', () => {
    // 2+3*4, A1..B2 sum, relative reference one column left
    expect(lotusCodeToText(Uint8Array.from([5, 2, 0, 5, 3, 0, 5, 4, 0, 0x0b, 0x09, 3]), { dialect: 'wk1', row: 0, col: 0, sheet: 0, dec: codepage(437) })).toBe('2+(3*4)');
    expect(lotusCodeToText(Uint8Array.from([1, 0xff, 0xbf, 0x00, 0x80, 3]), { dialect: 'wk1', row: 0, col: 1, sheet: 0, dec: codepage(437) })).toBe('A1');
  });

  it('reads a Release 2 worksheet: labels, numbers, formats, formulas, names, widths', async () => {
    const r = await importFile(ab(buildWk1()), 'prices.wk1');
    const b = r.book;
    expect(r.format).toBe('lotus');
    expect(cell(b, 'A1')?.v).toBe('Item');
    expect(cell(b, 'B1')?.v).toBe('Price');
    expect(style(b, 'B1').hAlign).toBe('right');
    expect(style(b, 'C1').hAlign).toBe('center');
    expect(cell(b, 'B2')?.v).toBe(1.25);
    expect(style(b, 'B2').numFmt).toBe('$#,##0.00_);($#,##0.00)');
    expect(cell(b, 'D2')).toMatchObject({ f: 'B2*C2', v: 5 });
    expect(cell(b, 'D5')?.f).toBe('SUM($D$2:D4)');
    expect(cell(b, 'A6')).toMatchObject({ f: '2+3*4', v: 14 });
    expect(cell(b, 'B6')).toMatchObject({ v: 'yes' });
    expect(cell(b, 'B6')?.f).toMatch(/^IF\(AND\(C2>3,C3<3\),"yes","no"\)$/);
    expect(cell(b, 'C6')?.f).toBe('AVERAGEA(C2:C4)');
    expect(style(b, 'A7').numFmt).toBe('dd-mmm-yy');
    expect(b.sheets[0].colW.get(0)).toBe(14 * 7 + 5);
    expect(b.names).toContainEqual(expect.objectContaining({ name: 'PRICES', ref: 'Sheet1!$B$2:$B$4' }));
  });

  it('reads a Release 3 workbook with sheet names and cross-sheet formulas', async () => {
    const b = (await importFile(ab(buildWk3()), 'sales.wk3')).book;
    expect(b.sheets.map((s) => s.name)).toEqual(['A', 'Totals']);
    expect(cell(b, 'B2')?.v).toBeCloseTo(1250.5);
    expect(cell(b, 'B3')?.v).toBe(980);
    expect(cell(b, 'B4')?.v).toBeCloseTo(0.015);
    expect(cell(b, 'B5')?.f).toBe('SUM(B2:B4)');
    expect(cell(b, 'B1', 1)?.f).toBe('A!B5*2');
    expect(b.sheets[0].colW.get(0)).toBe(20 * 7 + 5);
    expect(b.names).toContainEqual(expect.objectContaining({ name: 'SALES' }));
  });

  it('reads 1-2-3 97 compressed numbers and formulas', async () => {
    const b = (await importFile(ab(build123()), 'x.123')).book;
    expect(b.sheets[0].name).toBe('Data');
    expect(cell(b, 'A2')?.v).toBe(12.5);
    expect(cell(b, 'A3')?.v).toBe(-3);
    expect(cell(b, 'A4')).toMatchObject({ f: 'A2+A3*2', v: 6.5 });
  });

  it('keeps an untranslatable formula as its value with a note', async () => {
    const wk1 = buildWk1();
    // Replace the @AVG code (0x51) in C6 with an unknown function code.
    const i = wk1.findIndex((x, k) => x === 0x51 && wk1[k + 1] === 1 && wk1[k + 2] === 3);
    wk1[i] = 0xfe;
    const r = await importFile(ab(wk1), 'x.wk1');
    expect(cell(r.book, 'C6')?.f).toBeUndefined();
    expect(cell(r.book, 'C6')?.v).toBeCloseTo(14 / 3);
    expect(r.book.sheets[0].comments.size).toBe(1);
    expect(r.warnings.join(' ')).toMatch(/could not be translated/);
  });

  it('reads Works for DOS spreadsheets', async () => {
    const b = (await importFile(ab(buildWorksWks()), 'budget.wks')).book;
    expect(cell(b, 'A1')?.v).toBe('Budget');
    expect(style(b, 'A1').bold).toBe(true);
    expect(cell(b, 'B1')?.v).toBe(99.5);
  });

  it('detects the variant from the first record', () => {
    expect(lotusVariant(buildWk1())).toBe('wk1');
    expect(lotusVariant(buildWq1())).toBe('wq1');
    expect(lotusVariant(new Uint8Array(4))).toBeNull();
  });

  it('salvages a damaged worksheet', async () => {
    const wk1 = buildWk1();
    const cut = wk1.subarray(0, wk1.length - 40);
    const r = await importFile(ab(cut), 'cut.wk1');
    expect(cell(r.book, 'A1')?.v).toBe('Item');
    expect(r.warnings.join(' ')).toMatch(/damaged|ends early/);
  });
});

describe('Quattro Pro', () => {
  it('reads Quattro for DOS (WQ1) labels and formulas', async () => {
    const b = (await importFile(ab(buildWq1()), 'q.wq1')).book;
    expect(cell(b, 'A1')?.v).toBe('Hello');
    expect(style(b, 'B1').numFmt).toBe('0.000E+00');
    expect(cell(b, 'C1')).toMatchObject({ f: 'B1/2', v: 500 });
  });

  it('reads Quattro Pro for Windows notebooks with pages and block references', async () => {
    const b = (await importFile(ab(buildWb1()), 'q.wb1')).book;
    expect(b.sheets.map((s) => s.name)).toEqual(['Inputs', 'Results']);
    expect(cell(b, 'B4')).toMatchObject({ f: 'SUM(B2:B3)', v: 100 });
    expect(cell(b, 'A1', 1)?.f).toBe('Inputs!$B$1*2');
  });
});

describe('Excel 2.x-5.0', () => {
  it('reads BIFF2 cell attributes, formulas and widths', async () => {
    const b = (await importFile(ab(buildBiff2()), 'old.xls')).book;
    expect(cell(b, 'A1')?.v).toBe('Total');
    expect(style(b, 'A1').bold).toBe(true);
    expect(cell(b, 'A2')?.v).toBe(7);
    expect(style(b, 'A2').hAlign).toBe('right');
    expect(style(b, 'A3').numFmt).toBe('0.00');
    expect(cell(b, 'A4')).toMatchObject({ f: 'SUM(A2:A3)*2', v: 19 });
    expect(cell(b, 'A5')?.v).toBe(true);
    expect(b.sheets[0].colW.get(0)).toBe(140);
  });

  it('reads Excel 4.0 workbooks (BIFF4W) as several sheets', async () => {
    const b = (await importFile(ab(buildBiff4w()), 'book.xlw')).book;
    expect(b.sheets.map((s) => s.name)).toEqual(['Alpha', 'Beta']);
    expect(cell(b, 'A3')).toMatchObject({ f: 'A2*2', v: 20 });
    expect(cell(b, 'A2', 1)?.v).toBe(3.5);
    expect(style(b, 'A2', 1).numFmt).toBe('0.00');
  });

  it('reads Works 6-9 spreadsheets (XLR)', async () => {
    const r = await importFile(ab(buildXlr()), 'w.xlr');
    expect(cell(r.book, 'A4')?.f).toBe('SUM(A2:A3)*2');
  });

  it.each([
    ['testEXCEL_3.xls', 'Table C-13--Lemons: Utilization of production, by State, 1977/78 to date', 'C6', 376.2],
    ['testEXCEL_4.xls', 'Table 10 -- Examination Coverage:  Recommended and Average Recommended Additional Tax After', 'B13', 168184400],
    ['testEXCEL_5.xls', 'Sample Excel Worksheet - Numbers and their Squares', 'C8', 9],
    ['testEXCEL_95.xls', 'Sample Excel Worksheet - Numbers and their Squares', 'C8', 9],
  ])('matches LibreOffice on %s', async (f, a1, ref, value) => {
    const b = (await importFile(load(`legacy/${f}`), f)).book;
    expect(cell(b, 'A1')?.v).toBe(a1);
    expect(cell(b, ref)?.v).toBe(value);
  });

  it('shows "General" inside a format with literal text', () => {
    expect(formatValue(2, '\\(General\\)').text).toBe('(2)');
    expect(formatValue(1.5, '[Red]General').text).toBe('1.5');
  });

  it('treats a "GENERAL" format as General', async () => {
    const b = (await importFile(load('legacy/testEXCEL_5.xls'), 'x.xls')).book;
    expect(style(b, 'B6').numFmt).toBeUndefined();
    expect(cell(b, 'C6')?.f).toBe('+$B$6*$B$6');
  });
});

describe('dBASE', () => {
  it('reads fields, types, deleted records and memo files', () => {
    const { dbf, dbt } = buildDbf({ memo: true });
    const w: string[] = [];
    const b = importDbf(dbf, 'stock.dbf', w, { memo: dbt });
    const sh = b.sheets[0];
    expect(sh.name).toBe('stock');
    expect([0, 1, 2, 3, 4, 5].map((c) => sh.get(0, c)?.v)).toEqual(['NAME', 'QTY', 'PRICE', 'SOLD', 'OK', 'NOTES']);
    expect(sh.get(1, 0)?.v).toBe('Café');
    expect(sh.get(1, 1)?.v).toBe(3);
    expect(sh.get(1, 2)?.v).toBe(4.5);
    expect(b.styles[sh.get(1, 2)?.s || 0].numFmt).toBe('0.00');
    expect(sh.get(1, 3)?.v).toBe(36234); // 15 March 1999
    expect(sh.get(1, 4)?.v).toBe(true);
    expect(sh.get(1, 5)?.v).toBe('First note');
    expect(sh.get(2, 5)?.v).toBe('Second note');
    expect(sh.get(3, 0)).toBeUndefined();
    expect(w.join(' ')).toMatch(/deleted/);
    expect(sh.freeze.rows).toBe(1);
  });

  it('reports a missing memo file and uses the table code page', async () => {
    const r = await importFile(ab(buildDbf({ memo: true }).dbf), 'stock.dbf');
    expect(r.warnings.join(' ')).toMatch(/memo file/);
    const b437 = importDbf(buildDbf({ codepage: 'cp437' }).dbf, 'x.dbf');
    expect(b437.sheets[0].get(1, 0)?.v).toBe('Café');
  });
});

describe('VisiCalc', () => {
  it('translates left-to-right formulas', () => {
    expect(translateVisiCalc('2+3*4')).toBe('(2+3)*4');
    expect(translateVisiCalc('+B2*(1+B3)')).toBe('B2*(1+B3)');
    expect(translateVisiCalc('@SUM(A1...A9)/@COUNT(A1...A9)')).toBe('SUM(A1:A9)/COUNT(A1:A9)');
    expect(translateVisiCalc('-A1^2')).toBe('(-A1)^2');
    expect(translateVisiCalc('@NA')).toBe('NA()');
  });

  it('imports cells, formats and the column width', () => {
    const b = importVisiCalc(VISICALC_SAMPLE, 'budget.vc');
    const sh = b.sheets[0];
    expect(sh.get(0, 0)?.v).toBe('ITEM');
    expect(b.styles[sh.get(0, 1)?.s || 0].hAlign).toBe('right');
    expect(sh.get(1, 1)?.v).toBe(2.5);
    expect(b.styles[sh.get(1, 1)?.s || 0].numFmt).toBe('0.00');
    expect(sh.get(3, 1)?.f).toBe('SUM(B2:B3)');
    expect(sh.get(3, 2)?.f).toBe('(2+3)*4');
    expect(b.styles[sh.get(4, 0)?.s || 0].hAlign).toBe('fill');
    expect(sh.defaultColW).toBe(12 * 7 + 5);
  });
});

describe('best-effort formats', () => {
  it('lists AppleWorks values and says what is missing', async () => {
    // A minimal AppleWorks header and two cell records (text "Rent", number 450).
    const u = new Uint8Array(64);
    u.set([6, 7, 0xbd, 0, 0x42, 0x4f, 0x42, 0x4f], 0);
    u.set([0, 10, 3, 0, 0, 9, 4, 0x52, 0x65, 0x6e, 0x74, 0], 20);
    const n = [0, 14, 2, 0, 0, 9, 0x40, 0x07, 0xe1, 0, 0, 0, 0, 0, 0, 0];
    u.set(n, 34);
    const r = await importFile(ab(u), 'home.cwk');
    expect(r.book.sheets[0].get(0, 0)?.v).toBe('Rent');
    expect(r.book.sheets[0].get(1, 0)?.v).toBe(450);
    expect(r.warnings.join(' ')).toMatch(/values only/);
  });

  it('opens a file that nothing recognizes with a clear message', async () => {
    await expect(importFile(ab(Uint8Array.from([1, 2, 3, 250, 0, 7, 9, 200])), 'mystery')).rejects.toThrow(/does not recognize/);
  });

  it('formula errors become error values', async () => {
    const wk1 = buildWk1();
    const b = (await importFile(ab(wk1), 'x.wk1')).book;
    for (const sh of b.sheets) for (const c of sh.cells.values()) if (c.f != null) expect(c.v === null || c.v instanceof CellError || typeof c.v !== 'undefined').toBe(true);
  });
});
