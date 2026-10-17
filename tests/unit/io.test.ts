import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { Workbook, Sheet } from '../../src/model/workbook';
import { CellError } from '../../src/model/types';
import { cellKey } from '../../src/model/address';
import { importFile, exportFile, sniffFormat, formatForName, formatById, friendlyOpenError, PasswordError, FORMATS } from '../../src/io/index';
import { exportXlsx, toFileFormula } from '../../src/io/xlsx-export';
import { importXlsx, widthToPx } from '../../src/io/xlsx-import';
import { exportOds, importOds, exportFods, importFods, odfToFormula, formulaToOdf } from '../../src/io/ods';
import { parseCsv, detectDelimiter, writeCsv, importCsv, convertField } from '../../src/io/csv';
import { importHtml, sheetToMarkdown } from '../../src/io/html';
import { importSylk, importDif, importXml2003, r1c1ToA1 } from '../../src/io/other';

async function buf(b: Blob): Promise<ArrayBuffer> {
  if (typeof (b as any).arrayBuffer === 'function') return (b as any).arrayBuffer();
  return new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(fr.result as ArrayBuffer); fr.onerror = rej; fr.readAsArrayBuffer(b); });
}
async function text(b: Blob): Promise<string> { return new TextDecoder().decode(await buf(b)); }

const K = (a1: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return cellKey(+m[2] - 1, c);
};
const cell = (s: Sheet, a1: string) => s.cells.get(K(a1));

function makeBook(): Workbook {
  const b = new Workbook();
  const s = b.sheets[0];
  const put = (sh: Sheet, a1: string, c: any) => sh.cells.set(K(a1), c);
  const head = b.styleId({ bold: true, fill: '#ffff00', hAlign: 'center', bb: { style: 'thin', color: '#000000' } });
  const money = b.styleId({ numFmt: '0.00' });
  const date = b.styleId({ numFmt: 'yyyy-mm-dd' });
  const custom = b.styleId({ numFmt: '#,##0.000;[Red]-#,##0.000', italic: true, color: '#0000ff', font: 'Arial', size: 14 });
  put(s, 'A1', { v: 'Name', s: head }); put(s, 'B1', { v: 'Qty', s: head }); put(s, 'C1', { v: 'Price', s: head }); put(s, 'D1', { v: 'Total', s: head });
  [['Apple', 3, 1.25], ['Pear', 5, 2.5], ['Plum', 12, 0.75]].forEach(([n, q, p], i) => {
    const r = i + 2;
    put(s, `A${r}`, { v: n }); put(s, `B${r}`, { v: q }); put(s, `C${r}`, { v: p, s: money });
    put(s, `D${r}`, { f: `B${r}*C${r}`, v: (q as number) * (p as number), s: money });
  });
  put(s, 'A5', { v: 'Sum' }); put(s, 'D5', { f: 'SUM(D2:D4)', v: 24.25, s: money });
  put(s, 'E2', { v: 45366, s: date });
  put(s, 'F2', { v: true }); put(s, 'F3', { v: new CellError('#N/A') }); put(s, 'F4', { v: 1234.5678, s: custom });
  put(s, 'A7', { v: 'Merged title' });
  s.merges.push({ r1: 6, c1: 0, r2: 6, c2: 2 });
  s.freeze = { rows: 1, cols: 0 };
  s.colW.set(1, 100); s.rowH.set(0, 30); s.hiddenRows.add(19); s.hiddenCols.add(8);
  s.cf.push({ id: 'cf1', ranges: [{ r1: 1, c1: 3, r2: 3, c2: 3 }], type: 'cellIs', operator: 'greaterThan', formulas: ['10'], style: { fill: '#ff0000', bold: true } });
  s.cf.push({ id: 'cf2', ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], type: 'colorScale', colors: ['#f8696b', '#63be7b'], stops: [{ type: 'min' }, { type: 'max' }] });
  s.validations.push({ id: 'dv1', ranges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }], type: 'list', formula1: '"Apple,Pear,Plum"', allowBlank: true, showDropdown: true, errorStyle: 'stop', errorMessage: 'Pick a fruit' });
  // Table A10:C13
  put(s, 'A10', { v: 'Region' }); put(s, 'B10', { v: 'Sales' }); put(s, 'C10', { v: 'Units' });
  put(s, 'A11', { v: 'North' }); put(s, 'B11', { v: 100 }); put(s, 'C11', { v: 4 });
  put(s, 'A12', { v: 'South' }); put(s, 'B12', { v: 200 }); put(s, 'C12', { v: 6 });
  put(s, 'A13', { v: 'Total' }); put(s, 'B13', { f: 'SUBTOTAL(109,Sales[Sales])', v: 300 });
  s.tables.push({ id: 't1', name: 'Sales', range: { r1: 9, c1: 0, r2: 12, c2: 2 }, columns: [{ name: 'Region' }, { name: 'Sales', totalFn: 'sum' }, { name: 'Units' }], style: 'TableStyleMedium9', opts: { header: true, totals: true, banded: true, bandedCols: false, firstCol: false, lastCol: false, filterButton: true } });
  s.comments.set(K('A1'), { author: 'Ann', text: 'Header note', time: 0, note: true });
  b.names.push({ name: 'Rate', ref: 'Sheet1!$C$2', sheet: null });
  s.print.area = 'A1:D5';
  s.print.orientation = 'landscape';
  s.print.header = { left: 'Left', center: '&[Page] of &[Pages]', right: '' };
  put(s, 'G1', { v: 'site', link: 'https://example.com/' });
  put(s, 'H1', { v: 'Bold part plain', rich: [{ t: 'Bold part', s: { bold: true } }, { t: ' plain' }] });
  // Dynamic array anchor + spilled cells
  put(s, 'J1', { f: 'SEQUENCE(3)', v: 1 });
  put(s, 'J2', { v: 2, sp: K('J1') }); put(s, 'J3', { v: 3, sp: K('J1') });
  // Legacy CSE array
  put(s, 'K1', { f: 'B2:B3*2', v: 6, arr: { r1: 0, c1: 10, r2: 1, c2: 10 } }); put(s, 'K2', { v: 10, sp: K('K1') });

  const s2 = new Sheet(b, 'Data & More');
  b.sheets.push(s2);
  s2.cells.set(K('A1'), { v: 'Tricky <&> "quotes" \u0001ctrl' });
  s2.cells.set(K('A2'), { f: "Sheet1!D5*2", v: 48.5 });
  s2.cells.set(K('A3'), { f: 'XLOOKUP("Pear",Sheet1!A2:A4,Sheet1!B2:B4)', v: 5 });
  s2.cells.set(K('A4'), { f: 'IFERROR(1/0,"x")', v: 'x' });
  s2.cells.set(K('A5'), { v: '  leading and trailing  ' });
  s2.cells.set(K('A6'), { v: 'line1\nline2', s: b.styleId({ wrap: true }) });
  s2.tabColor = '#ff8800';
  s2.visibility = 'hidden';
  b.props.title = 'Round trip';
  b.props.author = 'Tester';
  b.clearHistory();
  return b;
}

describe('format registry', () => {
  it('lists the formats', () => {
    for (const id of ['xlsx', 'xlsm', 'xltx', 'ods', 'fods', 'csv', 'tsv', 'txt', 'html', 'json', 'qgw', 'qgwe', 'xls', 'sylk', 'dif', 'xml', 'md']) expect(formatById(id), id).toBeTruthy();
    expect(formatForName('a.XLSX')!.id).toBe('xlsx');
    expect(formatForName('report.slk')!.id).toBe('sylk');
    expect(FORMATS.find((f) => f.id === 'md')!.canImport).toBe(false);
  });
  it('friendly errors', () => {
    expect(friendlyOpenError('End of central directory not found. Is this a zip file ?')).toMatch(/table of contents/);
    expect(friendlyOpenError('invalid distance too far back')).toMatch(/damaged/);
  });
});

describe('xlsx', () => {
  it('formula prefixes for newer functions', () => {
    expect(toFileFormula('XLOOKUP(1,A:A,B:B)')).toBe('_xlfn.XLOOKUP(1,A:A,B:B)');
    expect(toFileFormula('FILTER(A1:A9,B1:B9>2)')).toBe('_xlfn._xlws.FILTER(A1:A9,B1:B9>2)');
    expect(toFileFormula('LET(x,2,x*STDEV.S(A1:A3))')).toBe('_xlfn.LET(_xlpm.x,2,_xlpm.x*_xlfn.STDEV.S(A1:A3))');
    expect(toFileFormula('SUM(A1,"XLOOKUP(")')).toBe('SUM(A1,"XLOOKUP(")');
  });

  it('round-trips a feature-rich workbook', async () => {
    const book = makeBook();
    const blob = await exportXlsx(book);
    const data = await buf(blob);
    expect(sniffFormat(data, 'x.bin')!.id).toBe('xlsx');
    // Raw XML checks
    const zip = await JSZip.loadAsync(data);
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const order = ['<sheetPr', '<dimension', '<sheetViews', '<sheetFormatPr', '<cols', '<sheetData', '<mergeCells', '<conditionalFormatting', '<dataValidations', '<hyperlinks', '<pageMargins', '<pageSetup', '<headerFooter', '<legacyDrawing', '<tableParts'];
    let last = -1;
    for (const tag of order) {
      const i = sheet1.indexOf(tag);
      if (tag === '<sheetPr' && i < 0) continue;
      expect(i, tag).toBeGreaterThan(last);
      last = i;
    }
    expect(sheet1).toContain('cm="1"');
    expect(sheet1).toContain('t="array" ref="J1:J3"');
    expect(sheet1).toContain('t="array" ref="K1:K2"');
    const sheet2 = await zip.file('xl/worksheets/sheet2.xml')!.async('string');
    expect(sheet2).toContain('_xlfn.XLOOKUP');
    expect(sheet2).not.toMatch(/\u0001/);
    expect(zip.file('xl/metadata.xml')).toBeTruthy();
    const wb = await zip.file('xl/workbook.xml')!.async('string');
    expect(wb).toContain('_xlnm.Print_Area');
    expect(wb).toContain('state="hidden"');
    // Every XML part parses
    for (const name of Object.keys(zip.files)) {
      if (!/\.(xml|rels|vml)$/.test(name)) continue;
      const x = await zip.file(name)!.async('string');
      const doc = new DOMParser().parseFromString(x, 'application/xml');
      expect(doc.getElementsByTagName('parsererror').length, name).toBe(0);
    }

    const back = await importXlsx(data);
    const s = back.sheets[0];
    expect(back.sheets.map((x) => x.name)).toEqual(['Sheet1', 'Data & More']);
    expect(cell(s, 'A1')!.v).toBe('Name');
    expect(back.styles[cell(s, 'A1')!.s!]).toMatchObject({ bold: true, fill: '#ffff00', hAlign: 'center' });
    expect(back.styles[cell(s, 'A1')!.s!].bb).toEqual({ style: 'thin', color: '#000000' });
    expect(cell(s, 'D2')!.f).toBe('B2*C2');
    expect(cell(s, 'D2')!.v).toBeCloseTo(3.75);
    expect(back.styles[cell(s, 'D2')!.s!].numFmt).toBe('0.00');
    expect(back.styles[cell(s, 'E2')!.s!].numFmt).toBe('yyyy-mm-dd');
    expect(back.styles[cell(s, 'F4')!.s!]).toMatchObject({ numFmt: '#,##0.000;[Red]-#,##0.000', italic: true, color: '#0000ff', font: 'Arial', size: 14 });
    expect(cell(s, 'F2')!.v).toBe(true);
    expect((cell(s, 'F3')!.v as CellError).code).toBe('#N/A');
    expect(s.merges).toEqual([{ r1: 6, c1: 0, r2: 6, c2: 2 }]);
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.colW.get(1)).toBe(100);
    expect(s.rowH.get(0)).toBe(30);
    expect(s.hiddenRows.has(19)).toBe(true);
    expect(s.hiddenCols.has(8)).toBe(true);
    expect(s.cf.length).toBe(2);
    expect(s.cf[0]).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulas: ['10'] });
    expect(s.cf[0].style).toMatchObject({ fill: '#ff0000', bold: true });
    expect(s.cf[1]).toMatchObject({ type: 'colorScale', colors: ['#f8696b', '#63be7b'] });
    expect(s.validations[0]).toMatchObject({ type: 'list', formula1: '"Apple,Pear,Plum"', allowBlank: true, showDropdown: true, errorMessage: 'Pick a fruit' });
    expect(s.tables[0]).toMatchObject({ name: 'Sales', range: { r1: 9, c1: 0, r2: 12, c2: 2 }, style: 'TableStyleMedium9' });
    expect(s.tables[0].columns.map((c) => c.name)).toEqual(['Region', 'Sales', 'Units']);
    expect(s.tables[0].columns[1].totalFn).toBe('sum');
    expect(s.tables[0].opts).toMatchObject({ header: true, totals: true, banded: true, filterButton: true });
    expect(s.comments.get(K('A1'))).toMatchObject({ author: 'Ann', text: 'Header note', note: true });
    expect(back.names).toEqual([{ name: 'Rate', ref: 'Sheet1!$C$2', sheet: null }]);
    expect(s.print.area).toBe('$A$1:$D$5');
    expect(s.print.orientation).toBe('landscape');
    expect(s.print.header).toEqual({ left: 'Left', center: '&[Page] of &[Pages]', right: '' });
    expect(cell(s, 'G1')!.link).toBe('https://example.com/');
    expect(cell(s, 'H1')!.rich).toEqual([{ t: 'Bold part', s: { bold: true } }, { t: ' plain' }]);
    expect(cell(s, 'J1')).toMatchObject({ f: 'SEQUENCE(3)', v: 1 });
    expect(cell(s, 'J1')!.arr).toBeUndefined();
    expect(cell(s, 'J3')).toMatchObject({ v: 3, sp: K('J1') });
    expect(cell(s, 'K1')!.arr).toEqual({ r1: 0, c1: 10, r2: 1, c2: 10 });
    expect(cell(s, 'K2')).toMatchObject({ v: 10, sp: K('K1') });
    const s2 = back.sheets[1];
    expect(cell(s2, 'A1')!.v).toBe('Tricky <&> "quotes" \u0001ctrl');
    expect(cell(s2, 'A2')!.f).toBe('Sheet1!D5*2');
    expect(cell(s2, 'A3')!.f).toBe('XLOOKUP("Pear",Sheet1!A2:A4,Sheet1!B2:B4)');
    expect(cell(s2, 'A4')!.v).toBe('x');
    expect(cell(s2, 'A5')!.v).toBe('  leading and trailing  ');
    expect(cell(s2, 'A6')!.v).toBe('line1\nline2');
    expect(s2.tabColor).toBe('#ff8800');
    expect(s2.visibility).toBe('hidden');
    expect(back.props.title).toBe('Round trip');
    expect(back.props.author).toBe('Tester');
  });

  it('round-trips pictures and charts', async () => {
    const b = new Workbook();
    const s = b.sheets[0];
    for (let r = 0; r < 4; r++) { s.cells.set(cellKey(r, 0), { v: `c${r}` }); s.cells.set(cellKey(r, 1), { v: r * 10 }); }
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    s.drawings.push({ id: 'i1', kind: 'image', at: { r: 5, c: 1, dx: 4, dy: 3, w: 120, h: 80 }, src: png, alt: 'dot' });
    s.drawings.push({ id: 't1', kind: 'textbox', at: { r: 1, c: 6, dx: 0, dy: 0, w: 150, h: 40 }, text: 'Hello\nWorld' });
    s.charts.push({ id: 'c1', type: 'stackedBar', title: 'Sales', series: [{ name: 'S1', values: 'Sheet1!B1:B4', categories: 'Sheet1!A1:A4', color: '#ff0000' }], at: { r: 10, c: 0, dx: 0, dy: 0, w: 480, h: 288 }, legend: 'bottom', gridlines: true });
    s.charts.push({ id: 'c2', type: 'pie', series: [{ values: 'B1:B4', categories: 'A1:A4' }], at: { r: 10, c: 9, dx: 0, dy: 0, w: 300, h: 200 } });
    s.charts.push({ id: 'c3', type: 'scatterLines', series: [{ values: 'B1:B4', categories: 'B1:B4' }], at: { r: 30, c: 0, dx: 0, dy: 0, w: 300, h: 200 } });
    s.sparklines.push({ id: 'sp1', cell: { r: 0, c: 3 }, data: 'B1:B4', type: 'column' });
    const back = await importXlsx(await buf(await exportXlsx(b)));
    const t = back.sheets[0];
    expect(t.drawings.length).toBe(2);
    const img = t.drawings.find((d) => d.kind === 'image')!;
    expect(img.src!.startsWith('data:image/png;base64,')).toBe(true);
    expect(img.at).toMatchObject({ r: 5, c: 1, dx: 4, dy: 3, w: 120, h: 80 });
    expect(img.alt).toBe('dot');
    expect(t.drawings.find((d) => d.kind === 'textbox')!.text).toBe('Hello\nWorld');
    expect(t.charts.length).toBe(3);
    expect(t.charts[0]).toMatchObject({ type: 'stackedBar', title: 'Sales', legend: 'bottom' });
    expect(t.charts[0].series[0]).toMatchObject({ values: 'Sheet1!$B$1:$B$4', categories: 'Sheet1!$A$1:$A$4', color: '#ff0000', name: 'S1' });
    expect(t.charts[0].at).toMatchObject({ r: 10, c: 0, w: 480, h: 288 });
    expect(t.charts[1].type).toBe('pie');
    expect(t.charts[2].type).toBe('scatterLines');
    expect(t.sparklines[0]).toMatchObject({ cell: { r: 0, c: 3 }, data: 'Sheet1!B1:B4', type: 'column' });
  });

  it('reads a hand-written minimal xlsx with shared formulas, inline strings and prefixes', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', '<?xml version="1.0"?><x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheets><x:sheet name="Data" sheetId="1" r:id="rId1"/></x:sheets><x:definedNames><x:definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$B$4</x:definedName><x:definedName name="Total">Data!$C$4</x:definedName></x:definedNames></x:workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/></Relationships>');
    zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData>' +
      '<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Item &amp; Co</x:t></x:is></x:c><x:c r="B1" t="inlineStr"><x:is><x:r><x:rPr><x:b/></x:rPr><x:t>Bo</x:t></x:r><x:r><x:t>ld</x:t></x:r></x:is></x:c></x:row>' +
      '<x:row r="2"><x:c r="A2"><x:v>1</x:v></x:c><x:c r="B2"><x:v>10</x:v></x:c><x:c r="C2"><x:f t="shared" ref="C2:C4" si="0">A2*B2</x:f><x:v>10</x:v></x:c></x:row>' +
      '<x:row r="3"><x:c r="A3"><x:v>2</x:v></x:c><x:c r="B3"><x:v>20</x:v></x:c><x:c r="C3"><x:f t="shared" si="0"/><x:v>40</x:v></x:c></x:row>' +
      '<x:row><x:c><x:v>3</x:v></x:c><x:c><x:v>30</x:v></x:c><x:c><x:f t="shared" si="0"/><x:v>90</x:v></x:c><x:c t="b"><x:v>1</x:v></x:c><x:c t="e"><x:v>#DIV/0!</x:v></x:c><x:c t="str"><x:f>_xlfn.CONCAT("a",_xlfn.XLOOKUP(1,A2:A4,B2:B4))</x:f><x:v>a10</x:v></x:c></x:row>' +
      '</x:sheetData></x:worksheet>');
    const book = await importXlsx(await zip.generateAsync({ type: 'arraybuffer' }));
    const s = book.sheets[0];
    expect(s.name).toBe('Data');
    expect(cell(s, 'A1')!.v).toBe('Item & Co');
    expect(cell(s, 'B1')!.v).toBe('Bold');
    expect(cell(s, 'B1')!.rich![0]).toEqual({ t: 'Bo', s: { bold: true } });
    expect(cell(s, 'C2')!.f).toBe('A2*B2');
    expect(cell(s, 'C3')!.f).toBe('A3*B3');
    expect(cell(s, 'C4')!.f).toBe('A4*B4');
    expect(cell(s, 'C4')!.v).toBe(90);
    expect(cell(s, 'D4')!.v).toBe(true);
    expect((cell(s, 'E4')!.v as CellError).code).toBe('#DIV/0!');
    expect(cell(s, 'F4')).toMatchObject({ f: 'CONCAT("a",XLOOKUP(1,A2:A4,B2:B4))', v: 'a10' });
    expect(book.names).toEqual([{ name: 'Total', ref: 'Data!$C$4', sheet: null }]);
  });

  it('imports a 100k-row sheet quickly', async () => {
    const rows: string[] = [];
    for (let r = 1; r <= 100000; r++) rows.push(`<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}" t="s"><v>${r % 3}</v></c><c r="C${r}"><f>A${r}*2</f><v>${r * 2}</v></c></row>`);
    const zip = new JSZip();
    zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Big" sheetId="1" r:id="rId1"/></sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>');
    zip.file('xl/sharedStrings.xml', '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>zero</t></si><si><t>one</t></si><si><t>two</t></si></sst>');
    zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join('')}</sheetData></worksheet>`);
    const data = await zip.generateAsync({ type: 'arraybuffer' });
    const t0 = Date.now();
    const book = await importXlsx(data);
    const ms = Date.now() - t0;
    const s = book.sheets[0];
    expect(s.cells.size).toBe(300000);
    expect(cell(s, 'B100000')!.v).toBe('one');
    expect(cell(s, 'C100000')!.f).toBe('A100000*2');
    expect(ms).toBeLessThan(10000);
  }, 60000);

  it('column width conversion', () => {
    expect(widthToPx(8.43)).toBe(59);
    expect(widthToPx(9.140625)).toBe(64);
  });
});

describe('damaged files', () => {
  it('routes a truncated .xlsx through recovery', async () => {
    const full = new Uint8Array(await buf(await exportXlsx(makeBook())));
    const cut = full.slice(0, Math.floor(full.length * 0.8)).buffer;
    let res: Awaited<ReturnType<typeof importFile>> | null = null;
    let err: Error | null = null;
    try { res = await importFile(cut, 'cut.xlsx'); } catch (e) { err = e as Error; }
    if (res) {
      expect(res.recovery).toBeTruthy();
      expect(res.book.sheets.length).toBeGreaterThan(0);
    } else {
      expect(err!.message).not.toMatch(/^End of central directory/);
    }
  });
  it('rejects unknown garbage with a readable message', async () => {
    const junk = new Uint8Array([0x50, 0x4b, 3, 4, 1, 2, 3, 4, 5]).buffer;
    await expect(importFile(junk, 'junk.xlsx')).rejects.toThrow();
  });
});

describe('ods', () => {
  it('converts formulas', () => {
    expect(odfToFormula('of:=SUM([.A1:.B5];[$Sheet2.C1])')).toBe('SUM(A1:B5,Sheet2!C1)');
    expect(odfToFormula("of:=['My Sheet'.$A$1]+[$'My Sheet'.B2:.C3]")).toBe("'My Sheet'!$A$1+'My Sheet'!B2:C3");
    expect(odfToFormula('of:=IF([.A1]>0;"a;b";{1;2|3;4})')).toBe('IF(A1>0,"a;b",{1,2;3,4})');
    expect(formulaToOdf("SUM(A1:B5,'My Sheet'!C1)")).toBe("of:=SUM([.A1:.B5];[$'My Sheet'.C1])");
    expect(formulaToOdf('SUM({1,2;3,4})')).toBe('of:=SUM({1;2|3;4})');
    const b = makeBook();
    const ctx = { book: b, sheet: b.sheets[0], r: 11 };
    expect(formulaToOdf('SUBTOTAL(109,Sales[Sales])', ctx)).toBe('of:=SUBTOTAL(109;[.$B$11:.$B$12])');
    expect(formulaToOdf('[@Units]*2', ctx)).toBe('of:=[.$C$12]*2');
    expect(formulaToOdf('COUNTA(Sales[[#All],[Region]:[Sales]])', { ...ctx, sheet: b.sheets[1] })).toBe('of:=COUNTA([$Sheet1.$A$10:.$B$13])');
  });

  it('round-trips through .ods and .fods', async () => {
    const book = makeBook();
    const data = await buf(await exportOds(book));
    expect(sniffFormat(data, 'whatever')!.id).toBe('ods');
    const zip = await JSZip.loadAsync(data);
    for (const name of ['content.xml', 'styles.xml', 'meta.xml', 'settings.xml', 'META-INF/manifest.xml']) {
      const x = await zip.file(name)!.async('string');
      expect(new DOMParser().parseFromString(x, 'application/xml').getElementsByTagName('parsererror').length, name).toBe(0);
    }
    for (const back of [await importOds(data), await importFods(await text(exportFods(book)))]) {
      const s = back.sheets[0];
      expect(back.sheets.map((x) => x.name)).toEqual(['Sheet1', 'Data & More']);
      expect(cell(s, 'A1')!.v).toBe('Name');
      expect(back.styles[cell(s, 'A1')!.s!]).toMatchObject({ bold: true, fill: '#ffff00', hAlign: 'center' });
      expect(cell(s, 'D2')!.f).toBe('B2*C2');
      expect(cell(s, 'D2')!.v).toBeCloseTo(3.75);
      expect(back.styles[cell(s, 'D2')!.s!].numFmt).toBe('0.00');
      expect(cell(s, 'E2')!.v).toBe(45366);
      expect(back.styles[cell(s, 'E2')!.s!].numFmt).toBe('yyyy-mm-dd');
      expect(cell(s, 'F2')!.v).toBe(true);
      expect(s.merges).toContainEqual({ r1: 6, c1: 0, r2: 6, c2: 2 });
      expect(s.colW.get(1)).toBe(100);
      expect(s.rowH.get(0)).toBe(30);
      expect(s.hiddenRows.has(19)).toBe(true);
      expect(s.freeze).toEqual({ rows: 1, cols: 0 });
      expect(s.comments.get(K('A1'))!.text).toBe('Header note');
      expect(back.names).toContainEqual({ name: 'Rate', ref: 'Sheet1!$C$2', sheet: null });
      expect(cell(s, 'G1')!.link).toBe('https://example.com/');
      const s2 = back.sheets[1];
      expect(cell(s2, 'A1')!.v).toBe('Tricky <&> "quotes" ctrl');
      expect(cell(s2, 'A2')!.f).toBe('Sheet1!D5*2');
      expect(cell(s2, 'A5')!.v).toBe('  leading and trailing  ');
      expect(cell(s2, 'A6')!.v).toBe('line1\nline2');
      expect(s2.visibility).toBe('hidden');
      expect(s2.tabColor).toBe('#ff8800');
      expect(back.props.title).toBe('Round trip');
    }
  });
});

describe('csv', () => {
  it('parses quoting, embedded newlines and delimiters', () => {
    const rows = parseCsv('a,"b,c","say ""hi"""\r\n1,"line1\nline2",3\n', { delimiter: ',' });
    expect(rows).toEqual([['a', 'b,c', 'say "hi"'], ['1', 'line1\nline2', '3']]);
    expect(detectDelimiter('a;b;c\n1;2;3\n4;5;6')).toBe(';');
    expect(detectDelimiter('a\tb\n1\t2')).toBe('\t');
    expect(detectDelimiter('x,"a;b;c;d",y\n1,2,3')).toBe(',');
    expect(parseCsv('﻿a,b\n\n1,2')).toEqual([['a', 'b'], [''], ['1', '2']]);
  });
  it('detects types', () => {
    expect(convertField('1,234.5').v).toBe(1234.5);
    expect(convertField('1.234,5', { decimal: ',' }).v).toBe(1234.5);
    expect(convertField('12%')).toMatchObject({ v: 0.12 });
    expect(convertField('TRUE').v).toBe(true);
    expect(convertField('2024-03-14')).toMatchObject({ v: 45365, fmt: 'yyyy-mm-dd' });
    expect(convertField('14/03/2024', { dateOrder: 'dmy' }).v).toBe(45365);
    expect(convertField('00123', {}, 'text')).toMatchObject({ v: '00123', fmt: '@' });
    expect(convertField('abc').v).toBe('abc');
  });
  it('imports and exports', async () => {
    const data = new TextEncoder().encode('Name;Amount;Zip\nA;1,5;00123\nB;2,25;04567\n').buffer;
    const book = importCsv(data, 'test.csv', { decimal: ',', columnTypes: ['general', 'general', 'text'] });
    const s = book.sheets[0];
    expect(s.name).toBe('test');
    expect(cell(s, 'B2')!.v).toBe(1.5);
    expect(cell(s, 'C2')!.v).toBe('00123');
    const out = await text(await exportFile(book, 'csv'));
    expect(out).toBe('Name,Amount,Zip\r\nA,1.5,00123\r\nB,2.25,04567\r\n');
    expect(writeCsv([['a,b', 'c"d', 'e\nf', ' g']])).toBe('"a,b","c""d","e\nf"," g"\r\n');
    const res = await importFile(new TextEncoder().encode('x\ty\n1\t2\n').buffer, 'data.tsv');
    expect(res.format).toBe('tsv');
    expect(cell(res.book.sheets[0], 'B2')!.v).toBe(2);
    const bom = await buf(await exportFile(book, 'csv', { csv: { bom: true } }));
    expect(new Uint8Array(bom).slice(0, 3)).toEqual(new Uint8Array([0xef, 0xbb, 0xbf]));
    const u16 = await buf(await exportFile(book, 'txt', { csv: { encoding: 'utf-16le' } }));
    const again = await importFile(u16, 'u.txt');
    expect(cell(again.book.sheets[0], 'A2')!.v).toBe('A');
  });
});

describe('html and markdown', () => {
  it('round-trips through html', async () => {
    const book = makeBook();
    const html = await text(await exportFile(book, 'html', { sheetIndex: 0 }));
    expect(html).toContain('colspan="3"');
    expect(html).toContain('1.25');
    const back = importHtml(html);
    const s = back.sheets[0];
    expect(cell(s, 'A1')!.v).toBe('Name');
    expect(back.styles[cell(s, 'A1')!.s!]).toMatchObject({ bold: true, fill: '#ffff00' });
    expect(cell(s, 'C2')!.v).toBe(1.25);
    expect(s.merges).toContainEqual({ r1: 6, c1: 0, r2: 6, c2: 2 });
    const t = importHtml('<table><tr><td rowspan="2">a</td><td>b</td></tr><tr><td>c</td></tr><tr><td>&lt;x&gt;</td><td><a href="http://q">1,000</a></td></tr></table>').sheets[0];
    expect(cell(t, 'B2')!.v).toBe('c');
    expect(cell(t, 'A3')!.v).toBe('<x>');
    expect(cell(t, 'B3')).toMatchObject({ v: 1000, link: 'http://q' });
    expect(t.merges).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 0 }]);
  });
  it('exports markdown', async () => {
    const b = new Workbook();
    const s = b.sheets[0];
    s.cells.set(K('A1'), { v: 'Name' }); s.cells.set(K('B1'), { v: 'N' });
    s.cells.set(K('A2'), { v: 'a|b' }); s.cells.set(K('B2'), { v: 5 });
    expect(sheetToMarkdown(s)).toBe('| Name | N |\n| --- | ---: |\n| a\\|b | 5 |\n');
    expect(await text(await exportFile(b, 'md'))).toContain('| Name | N |');
  });
});

describe('native formats', () => {
  it('round-trips json, qgw and qgwe', async () => {
    const book = makeBook();
    book.sheets[0].drawings.push({ id: 'i', kind: 'image', at: { r: 0, c: 0, dx: 0, dy: 0, w: 1, h: 1 }, src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==' });
    for (const id of ['json', 'qgw']) {
      const data = await buf(await exportFile(book, id));
      const res = await importFile(data, `x.${id}`);
      expect(res.format).toBe(id);
      const s = res.book.sheets[0];
      expect(cell(s, 'D2')!.f).toBe('B2*C2');
      expect((cell(s, 'F3')!.v as CellError).code).toBe('#N/A');
      expect(s.drawings[0].src).toBe(book.sheets[0].drawings[0].src);
      expect(res.book.sheets[1].name).toBe('Data & More');
    }
    const enc = await buf(await exportFile(book, 'qgwe', { password: 'secret' }));
    expect(sniffFormat(enc, 'noext')!.id).toBe('qgwe');
    await expect(importFile(enc, 'x.qgwe')).rejects.toBeInstanceOf(PasswordError);
    await expect(importFile(enc, 'x.qgwe', { password: 'wrong' })).rejects.toBeInstanceOf(PasswordError);
    const res = await importFile(enc, 'x.qgwe', { password: 'secret' });
    expect(res.encrypted).toBe(true);
    expect(cell(res.book.sheets[0], 'A1')!.v).toBe('Name');
  }, 30000);
});

describe('legacy formats', () => {
  it('converts R1C1', () => {
    expect(r1c1ToA1('SUM(R[-2]C:R[-1]C)', 4, 1)).toBe('SUM(B3:B4)');
    expect(r1c1ToA1('R1C1+RC[1]&"R1C1"', 2, 2)).toBe('$A$1+D3&"R1C1"');
    expect(r1c1ToA1('ROUND(R2C3,0)', 0, 0)).toBe('ROUND($C$2,0)');
  });
  it('reads SYLK', () => {
    const b = importSylk('ID;PWXL;N;E\r\nP;PGeneral\r\nP;P0.00\r\nC;Y1;X1;K"Hello;;World"\r\nC;X2;K12.5\r\nC;Y2;X2;K25;ER[-1]C*2\r\nF;P1;FF2G;X2;Y2\r\nE\r\n');
    const s = b.sheets[0];
    expect(cell(s, 'A1')!.v).toBe('Hello;World');
    expect(cell(s, 'B1')!.v).toBe(12.5);
    expect(cell(s, 'B2')).toMatchObject({ v: 25, f: 'B1*2' });
    expect(b.styles[cell(s, 'B2')!.s!].numFmt).toBe('0.00');
  });
  it('reads DIF', () => {
    const b = importDif('TABLE\r\n0,1\r\n""\r\nVECTORS\r\n0,2\r\n""\r\nTUPLES\r\n0,2\r\n""\r\nDATA\r\n0,0\r\n""\r\n-1,0\r\nBOT\r\n1,0\r\n"Name"\r\n0,42\r\nV\r\n-1,0\r\nBOT\r\n1,0\r\n"x ""y"""\r\n0,1\r\nTRUE\r\n-1,0\r\nEOD\r\n');
    const s = b.sheets[0];
    expect(cell(s, 'A1')!.v).toBe('Name');
    expect(cell(s, 'B1')!.v).toBe(42);
    expect(cell(s, 'A2')!.v).toBe('x "y"');
    expect(cell(s, 'B2')!.v).toBe(true);
  });
  it('reads XML Spreadsheet 2003', async () => {
    const xml = `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <Styles><Style ss:ID="b"><Font ss:Bold="1"/><Interior ss:Color="#FFFF00" ss:Pattern="Solid"/></Style><Style ss:ID="d"><NumberFormat ss:Format="Short Date"/></Style></Styles>
 <Names><NamedRange ss:Name="Top" ss:RefersTo="=Data!R1C1:R1C2"/></Names>
 <Worksheet ss:Name="Data"><Table><Column ss:Width="75"/>
  <Row><Cell ss:StyleID="b"><Data ss:Type="String">Head</Data></Cell><Cell ss:MergeAcross="1"><Data ss:Type="Number">2</Data></Cell></Row>
  <Row ss:Index="3" ss:Height="30"><Cell ss:Formula="=R[-2]C[1]*3"><Data ss:Type="Number">6</Data></Cell><Cell ss:StyleID="d"><Data ss:Type="DateTime">2024-03-14T00:00:00.000</Data></Cell>
   <Cell><Data ss:Type="Boolean">1</Data><Comment ss:Author="Me"><ss:Data>Note</ss:Data></Comment></Cell></Row>
 </Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><SplitHorizontal>1</SplitHorizontal></WorksheetOptions></Worksheet>
</Workbook>`;
    const bytes = new TextEncoder().encode(xml).buffer;
    expect(sniffFormat(bytes, 'book.xml')!.id).toBe('xml');
    const b = (await importFile(bytes, 'book.xml')).book;
    const s = b.sheets[0];
    expect(s.name).toBe('Data');
    expect(b.styles[cell(s, 'A1')!.s!]).toMatchObject({ bold: true, fill: '#ffff00' });
    expect(s.merges).toEqual([{ r1: 0, c1: 1, r2: 0, c2: 2 }]);
    expect(cell(s, 'A3')).toMatchObject({ f: 'B1*3', v: 6 });
    expect(cell(s, 'B3')!.v).toBe(45365);
    expect(b.styles[cell(s, 'B3')!.s!].numFmt).toBe('m/d/yyyy');
    expect(s.comments.get(K('C3'))!.text).toBe('Note');
    expect(s.colW.get(0)).toBe(100);
    expect(s.rowH.get(2)).toBe(40);
    expect(s.freeze.rows).toBe(1);
    expect(b.names[0]).toMatchObject({ name: 'Top', ref: 'Data!$A$1:$B$1' });
  });
});

describe('files written by LibreOffice', () => {
  const load = (f: string) => { const u8 = readFileSync(`tests/fixtures/io/${f}`); return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer; };
  it('reads an .xlsx LibreOffice saved from our export', async () => {
    const res = await importFile(load('libreoffice-roundtrip.xlsx'), 'lo.xlsx');
    const b = res.book;
    const s = b.sheets[0];
    expect(b.styles[cell(s, 'A1')!.s!]).toEqual({ bb: { style: 'thin' }, bold: true, fill: '#ffff00', hAlign: 'center' });
    expect(cell(s, 'D5')).toMatchObject({ f: 'SUM(D2:D4)', v: 24.25 });
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.cf.map((r) => r.type)).toEqual(['cellIs', 'colorScale']);
    expect(s.validations[0]).toMatchObject({ type: 'list', formula1: '"Apple,Pear,Plum"' });
    expect(s.validations[0].formula2).toBeUndefined();
    expect(s.tables[0].name).toBe('Sales');
    expect(s.charts.length).toBe(11);
    expect(s.charts[0].series[0]).toMatchObject({ name: 'Qty', values: 'Sheet1!$B$2:$B$4' });
    expect(s.charts[0].series[1]).toMatchObject({ nameRef: 'Sheet1!$C$1', categories: 'Sheet1!$A$2:$A$4' });
    expect(s.charts.map((c) => c.type)).toContain('pie');
    expect(s.drawings.some((d) => d.kind === 'image')).toBe(true);
    expect(s.sparklines.length).toBe(1);
    expect(s.print.header.center).toBe('&[Page] of &[Pages]');
    expect(b.sheets[1].visibility).toBe('hidden');
  });
  it('reads an .ods LibreOffice saved from our .xlsx', async () => {
    const res = await importFile(load('libreoffice.ods'), 'lo.ods');
    expect(res.format).toBe('ods');
    const b = res.book;
    const s = b.sheets[0];
    expect(b.styles[cell(s, 'A1')!.s!]).toMatchObject({ bold: true, fill: '#ffff00', hAlign: 'center' });
    expect(b.styles[cell(s, 'A1')!.s!].font).toBeUndefined();
    expect(cell(s, 'B2')).toEqual({ v: 3 });
    expect(cell(s, 'D2')).toMatchObject({ f: 'B2*C2', v: 3.75 });
    expect(b.styles[cell(s, 'D2')!.s!].numFmt).toBe('0.00');
    expect(b.styles[cell(s, 'E2')!.s!].numFmt).toBe('yyyy-mm-dd');
    expect(cell(s, 'E2')!.v).toBe(45366);
    expect(cell(s, 'F2')).toMatchObject({ v: true });
    expect(cell(s, 'F2')!.f).toBeUndefined();
    expect((cell(s, 'F3')!.v as CellError).code).toBe('#N/A');
    expect(b.styles[cell(s, 'F4')!.s!]).toMatchObject({ font: 'Arial', size: 14, italic: true, color: '#0000ff' });
    expect(cell(s, 'K2')).toMatchObject({ v: 10, sp: K('K1') });
    expect(s.merges).toEqual([{ r1: 6, c1: 0, r2: 6, c2: 2 }]);
    expect(s.hiddenRows.has(19)).toBe(true);
    expect(s.rowH.get(0)).toBe(30);
    expect(s.print.area).toBe('$A$1:$D$5');
    expect(s.drawings.filter((d) => d.kind === 'image').length).toBe(1);
    expect(s.charts.length).toBeGreaterThan(0); // charts in OpenDocument files are read
    expect(b.names).toEqual([{ name: 'Rate', ref: 'Sheet1!$C$2', sheet: null }]);
    expect(b.sheets[1].visibility).toBe('hidden');
  });
});
