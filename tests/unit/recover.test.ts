import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { recoverWorkbook, fixSheetXml, scanSheetCells, parseSharedStrings, shiftFormula, scanOdsTables, reportText } from '../../src/io/recover';
import { readBiff, importXls, XLS_ENCRYPTED_MSG, BIFF_FUNCTIONS } from '../../src/io/recover/xls';
import { isWellFormed } from '../../src/io/recover/salvage';
import type { Workbook } from '../../src/model/workbook';
import { CellError } from '../../src/model/types';

const buf = (u8: Uint8Array) => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
const enc = (s: string) => new TextEncoder().encode(s);

// ------------------------------------------------------------------------------------------
// XLSX fixture: a valid package built with JSZip, then damaged in different ways

const N = 600;
const WORDS = ['Alpha', 'Beta', 'Gamma', 'Delta'];
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNS = 'http://schemas.openxmlformats.org/package/2006/relationships';

function expectedData(i: number): [string, number, number, string] {
  return [WORDS[i % 4], i * 1.5, i * 3, `Row ${i}`];
}

function dataSheetXml(opts: { drawing?: boolean } = {}): string {
  const rows: string[] = [];
  for (let i = 1; i <= N; i++) {
    const [, b, c, d] = expectedData(i);
    rows.push(`<row r="${i}"><c r="A${i}" t="s" s="1"><v>${i % 4}</v></c><c r="B${i}"><v>${b}</v></c><c r="C${i}"><f>B${i}*2</f><v>${c}</v></c><c r="D${i}" t="inlineStr"><is><t>${d}</t></is></c></row>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS}" xmlns:r="${RNS}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>${rows.join('')}</sheetData><mergeCells count="1"><mergeCell ref="F1:G2"/></mergeCells>${opts.drawing ? '<drawing r:id="rId1"/>' : ''}</worksheet>`;
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

function sstXml(words = WORDS): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="${NS}" count="${N}" uniqueCount="${words.length}">${words.map((w) => `<si><t>${w}</t></si>`).join('')}</sst>`;
}

interface PkgOpts { skip?: string[]; replace?: Record<string, string | Uint8Array>; drawing?: boolean; sheetXml?: string }

async function buildXlsx(o: PkgOpts = {}): Promise<Uint8Array> {
  const files: [string, string | Uint8Array][] = [
    ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`],
    ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS}"><Relationship Id="rId1" Type="${RNS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS}" xmlns:r="${RNS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Summary" sheetId="2" r:id="rId2"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS}"><Relationship Id="rId1" Type="${RNS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${RNS}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${RNS}/styles" Target="styles.xml"/><Relationship Id="rId4" Type="${RNS}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`],
    ['xl/styles.xml', STYLES],
    ['xl/sharedStrings.xml', sstXml()],
    ['xl/worksheets/sheet2.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Total</t></is></c><c r="B1"><f>SUM(Data!B1:B${N})</f><v>${(1.5 * N * (N + 1)) / 2}</v></c></row></sheetData></worksheet>`],
    ['xl/worksheets/sheet1.xml', o.sheetXml ?? dataSheetXml({ drawing: o.drawing })],
  ];
  if (o.drawing) files.push(['xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS}"><Relationship Id="rId1" Type="${RNS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`]);
  const zip = new JSZip();
  for (const [n, d] of files) {
    if (o.skip?.includes(n)) continue;
    zip.file(n, o.replace?.[n] ?? d);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

/** The regular XLSX importer loads (it is developed separately and may be incomplete). */
async function importerAvailable(): Promise<boolean> {
  try { return typeof (await import('../../src/io/xlsx-import')).importXlsx === 'function'; } catch { return false; }
}

function entryData(u8: Uint8Array, want: string): { at: number; size: number } {
  for (let i = 0; i < u8.length - 30; i++) {
    if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 3 && u8[i + 3] === 4) {
      const nl = u8[i + 26] | (u8[i + 27] << 8), el = u8[i + 28] | (u8[i + 29] << 8);
      if (new TextDecoder().decode(u8.slice(i + 30, i + 30 + nl)) === want) {
        return { at: i + 30 + nl + el, size: (u8[i + 18] | (u8[i + 19] << 8) | (u8[i + 20] << 16) | (u8[i + 21] << 24)) >>> 0 };
      }
    }
  }
  throw new Error('entry not found');
}

/** Rows 1..k of the Data sheet that hold exactly the expected values. */
function checkData(book: Workbook, sheet = 'Data'): { complete: number; cells: number; wrong: number } {
  const sh = book.sheetByName(sheet);
  if (!sh) return { complete: 0, cells: 0, wrong: 0 };
  let complete = 0, cells = 0, wrong = 0, prefix = true;
  for (let i = 1; i <= N; i++) {
    const exp = expectedData(i);
    let rowOk = true;
    for (let c = 0; c < 4; c++) {
      const v = sh.value(i - 1, c);
      if (v == null) { rowOk = false; continue; }
      cells++;
      if (v !== exp[c]) { wrong++; rowOk = false; }
    }
    if (prefix && rowOk) complete++; else prefix = false;
  }
  return { complete, cells, wrong };
}

describe('XLSX recovery', () => {
  it('opens an intact workbook normally', async () => {
    const u8 = await buildXlsx();
    const r = await recoverWorkbook(buf(u8), 'intact.xlsx');
    expect(r.kind).toBe('xlsx');
    expect(r.book).not.toBeNull();
    expect(checkData(r.book!)).toEqual({ complete: N, cells: N * 4, wrong: 0 });
    if (await importerAvailable()) {
      expect(r.methods.find((m) => m.id === 'open')?.status).toBe('success');
      expect(r.methods.find((m) => m.id === 'zip')?.status).toBe('skipped');
    } else {
      // The XLSX importer is being developed separately; without it the raw worksheet scan opens the file.
      expect(r.bestMethod).toBe('rawsheet');
    }
  });

  it('Open and Repair (forceRepair) rebuilds a package that opens', async () => {
    const r = await recoverWorkbook(buf(await buildXlsx()), 'intact.xlsx', { forceRepair: true });
    expect(checkData(r.book!).complete).toBe(N);
    expect(r.repaired).toBeTruthy();
    const zip = await JSZip.loadAsync(await r.repaired!.blob.arrayBuffer());
    expect(Object.keys(zip.files)).toContain('xl/worksheets/sheet1.xml');
    expect(isWellFormed(await zip.file('[Content_Types].xml')!.async('string'))).toBe(true);
  });

  it('recovers every cell when the central directory is missing', async () => {
    const u8 = await buildXlsx();
    let cd = -1;
    for (let i = 0; i < u8.length - 4; i++) if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 1 && u8[i + 3] === 2) { cd = i; break; }
    const r = await recoverWorkbook(buf(u8.slice(0, cd)), 'nocd.xlsx');
    expect(r.methods.find((m) => m.id === 'zip')?.detail).toMatch(/local headers/);
    expect(checkData(r.book!)).toEqual({ complete: N, cells: N * 4, wrong: 0 });
    expect(r.book!.sheetByName('Summary')?.value(0, 0)).toBe('Total');
  });

  it('regenerates a missing [Content_Types].xml', async () => {
    const r = await recoverWorkbook(buf(await buildXlsx({ skip: ['[Content_Types].xml'] })), 'noct.xlsx');
    expect(checkData(r.book!).complete).toBe(N);
    if (r.repaired) {
      const zip = await JSZip.loadAsync(await r.repaired.blob.arrayBuffer());
      expect(zip.file('[Content_Types].xml')).toBeTruthy();
    }
  });

  it('regenerates the workbook part and relationships from the sheet parts', async () => {
    const r = await recoverWorkbook(buf(await buildXlsx({ skip: ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '_rels/.rels'] })), 'nowb.xlsx');
    expect(r.book!.sheets.length).toBe(2);
    // Sheet names are lost with workbook.xml; sheet1.xml comes first.
    expect(checkData(r.book!, r.book!.sheets[0].name).complete).toBe(N);
  });

  it('repairs broken sharedStrings XML', async () => {
    const bad = sstXml().replace('<si><t>Beta</t></si>', '<si><t>Beta</si>').replace('<si><t>Gamma', '<si>< & <t>Gamma');
    expect(isWellFormed(bad)).toBe(false);
    const r = await recoverWorkbook(buf(await buildXlsx({ replace: { 'xl/sharedStrings.xml': bad } })), 'sst.xlsx');
    expect(checkData(r.book!)).toEqual({ complete: N, cells: N * 4, wrong: 0 });
  });

  it('keeps the text of a truncated last shared string', async () => {
    const full = sstXml();
    const cut = full.slice(0, full.indexOf('<si><t>Delta') + 10);
    const r = await recoverWorkbook(buf(await buildXlsx({ replace: { 'xl/sharedStrings.xml': cut } })), 'sst2.xlsx');
    const sh = r.book!.sheetByName('Data')!;
    for (let i = 1; i <= N; i++) {
      const v = sh.value(i - 1, 0);
      if (i % 4 === 3) expect(String(v)).toMatch(/^Del/);
      else expect(v).toBe(WORDS[i % 4]);
    }
    expect(parseSharedStrings(cut)).toEqual({ items: ['Alpha', 'Beta', 'Gamma', 'Del'], truncated: true });
  });

  it('falls back to a minimal stylesheet when styles.xml is unusable, keeping values', async () => {
    const r = await recoverWorkbook(buf(await buildXlsx({ replace: { 'xl/styles.xml': '<styleSheet xmlns="x"><fonts><fo\u0001nt><b' } })), 'styles.xlsx');
    expect(checkData(r.book!)).toEqual({ complete: N, cells: N * 4, wrong: 0 });
  });

  it('removes dangling relationships (a drawing whose part is missing)', async () => {
    const r = await recoverWorkbook(buf(await buildXlsx({ drawing: true })), 'drawing.xlsx', { forceRepair: true });
    expect(checkData(r.book!).complete).toBe(N);
    const zip = await JSZip.loadAsync(await r.repaired!.blob.arrayBuffer());
    expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).not.toContain('<drawing');
    expect(await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string')).not.toContain('drawing1.xml');
    expect(r.repaired!.repairedParts.join(' ')).toMatch(/dangling/);
  });

  for (const pct of [0.8, 0.6]) {
    it(`recovers a prefix of rows from a file truncated at ${pct * 100}%`, async () => {
      const u8 = await buildXlsx();
      const r = await recoverWorkbook(buf(u8.slice(0, Math.floor(u8.length * pct))), `trunc${pct}.xlsx`);
      expect(r.book).not.toBeNull();
      const c = checkData(r.book!);
      expect(c.wrong).toBe(0);
      expect(c.complete).toBeGreaterThan(N * (pct - 0.5));
      expect(c.complete).toBeLessThan(N);
      // Every recovered row is complete: the sheet was cut at the last complete row.
      expect(c.cells).toBe(c.complete * 4);
      expect(r.methods.find((m) => m.id === 'zip')?.status).toBe('partial');
    });
  }

  it('recovers the rows before bytes flipped in the middle of the sheet deflate stream', async () => {
    const u8 = await buildXlsx();
    const e = entryData(u8, 'xl/worksheets/sheet1.xml');
    const bad = u8.slice();
    for (let k = 0; k < 8; k++) bad[e.at + Math.floor(e.size / 2) + k] ^= 0x5a;
    const r = await recoverWorkbook(buf(bad), 'flipped.xlsx');
    expect(r.book).not.toBeNull();
    const c = checkData(r.book!);
    expect(c.complete).toBeGreaterThan(N * 0.3);
    expect(c.wrong).toBeLessThanOrEqual(4);
    expect(r.unverified).toBe(true);
    expect(r.summary).toMatch(/verify/);
  });

  it('finds the workbook behind leading garbage and sets an appended PDF aside', async () => {
    const u8 = await buildXlsx();
    const pdf = enc('%PDF-1.4\n1 0 obj << /Type /Catalog >> endobj\ntrailer\n%%EOF\n');
    const mixed = new Uint8Array([...u8, ...pdf]);
    const r = await recoverWorkbook(buf(mixed), 'mixed.xlsx');
    expect(checkData(r.book!).complete).toBe(N);
    expect(r.foreign.some((s) => s.ext === 'pdf')).toBe(true);
  });

  it('identifies an xlsx with the wrong extension', async () => {
    const r = await recoverWorkbook(buf(await buildXlsx()), 'renamed.xls');
    expect(r.kind).toBe('xlsx');
    expect(r.notes.join(' ')).toMatch(/really/);
    expect(checkData(r.book!).complete).toBe(N);
  });
});

describe('worksheet XML repair helpers', () => {
  it('renumbers missing, duplicate and out-of-order r attributes and drops broken cells', () => {
    const xml = `<worksheet xmlns="${NS}"><sheetData><row><c><v>1</v></c><c><v>2</v></c></row><row r="5"><c r="A5"><v>3</v></c><c r="A5"><v>4</v></c></row><row r="5"><c r="B9"><v>5</v></c><c r="C6"><v>x</v><v></c><c r="D6"><v>6</v></c></row></sheetData></worksheet>`;
    const f = fixSheetXml(xml);
    expect(isWellFormed(f.xml)).toBe(true);
    expect(f.xml).toContain('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>');
    expect(f.xml).toContain('<c r="A5"><v>3</v></c><c r="B5"><v>4</v></c>');
    expect(f.xml).toContain('<row r="6"><c r="B6"><v>5</v></c><c r="D6"><v>6</v></c></row>');
    expect(f.dropped).toBe(1);
  });

  it('closes a truncated worksheet', () => {
    const f = fixSheetXml(`<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2`);
    expect(isWellFormed(f.xml)).toBe(true);
    expect(f.xml).toContain('<c r="A1"><v>1</v></c>');
  });

  it('scans cells from fragments, with shared strings and shared formulas', () => {
    const cells = scanSheetCells('<row r="3"><c r="B3" t="s"><v>1</v></c><c r="C3"><f t="shared" si="0" ref="C3:C4">B3&amp;"x"</f><v>0</v></c></row><row r="4"><c r="C4"><f t="shared" si="0"/><v>0</v></c><c t="b"><v>1</v></c><c t="e"><v>#DIV/0!</v></c>', ['zero', 'one']);
    const at = (r: number, c: number) => cells.find((x) => x.r === r && x.c === c)?.cell;
    expect(at(2, 1)?.v).toBe('one');
    expect(at(2, 2)?.f).toBe('B3&"x"');
    expect(at(3, 2)?.f).toBe('B4&"x"');
    expect(at(3, 3)?.v).toBe(true);
    expect((at(3, 4)?.v as CellError).code).toBe('#DIV/0!');
  });

  it('shifts relative references only', () => {
    expect(shiftFormula('SUM($A$1:A2)+B$3+"A1"+LOG10(C4)', 1, 2)).toBe('SUM($A$1:C3)+D$3+"A1"+LOG10(E5)');
  });
});

// ------------------------------------------------------------------------------------------
// Other spreadsheet formats

describe('ODS, text, HTML and XML 2003 recovery', () => {
  const content = (rows: number) => `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2"><office:body><office:spreadsheet><table:table table:name="Sheet A">${Array.from({ length: rows }, (_, i) => `<table:table-row><table:table-cell office:value-type="string"><text:p>Item ${i + 1}</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="${i + 1}"><text:p>${i + 1}</text:p></table:table-cell><table:table-cell table:formula="of:=[.B${i + 1}]*2" office:value-type="float" office:value="${(i + 1) * 2}"/></table:table-row>`).join('')}</table:table></office:spreadsheet></office:body></office:document-content>`;
  const ods = async (xml: string) => {
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' });
    zip.file('content.xml', xml);
    zip.file('META-INF/manifest.xml', '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>');
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  };

  it('scans OpenDocument tables directly', () => {
    const b = scanOdsTables(content(3))!;
    const sh = b.sheets[0];
    expect(sh.name).toBe('Sheet A');
    expect(sh.value(2, 0)).toBe('Item 3');
    expect(sh.value(2, 1)).toBe(3);
    expect(sh.get(2, 2)?.f).toBe('B3*2');
  });

  it('recovers an ODS whose content.xml is truncated', async () => {
    const u8 = await ods(content(400));
    const r = await recoverWorkbook(buf(u8.slice(0, Math.floor(u8.length * 0.7))), 'trunc.ods', { forceRepair: true });
    expect(r.kind).toBe('ods');
    const sh = r.book!.sheets[0];
    let k = 0;
    while (k < 400 && sh.value(k, 0) === `Item ${k + 1}` && sh.value(k, 1) === k + 1) k++;
    expect(k).toBeGreaterThan(50);
    expect(r.repaired?.kind).toBe('ods');
  });

  it('reads CSV, HTML and damaged SpreadsheetML 2003', async () => {
    const csv = await recoverWorkbook(buf(enc('a,b,c\n1,2,3\n"x,y",4,5\n')), 'data.csv', { forceRepair: true });
    expect(csv.book!.sheets[0].value(2, 0)).toBe('x,y');
    expect(csv.book!.sheets[0].value(1, 2)).toBe(3);
    const html = await recoverWorkbook(buf(enc('<html><body><table><tr><td>Name</td><td>Qty</td></tr><tr><td>Pen</td><td>4</td></tr></table>')), 'page.html', { forceRepair: true });
    expect(html.book!.sheets[0].value(1, 1)).toBe(4);
    const xml = `<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="S1"><Table><Row><Cell><Data ss:Type="Number">5</Data></Cell><Cell ss:Formula="=RC[-1]*2"><Data ss:Type="Number">10</Data></Cell></Row><Row><Cell><Data ss:Type="String">cut he`;
    const x = await recoverWorkbook(buf(enc(xml)), 'old.xml');
    expect(x.kind).toBe('xml2003');
    expect(x.book!.sheets[0].value(0, 1)).toBe(10);
    expect(x.book!.sheets[0].get(0, 1)?.f).toBe('A1*2');
    expect(x.book!.sheets[0].value(1, 0)).toBe('cut he');
  });
});

describe('last resorts and identification', () => {
  it('never throws on garbage and falls back to a strings scan', async () => {
    const junk = new Uint8Array(4000).map((_, i) => (i * 2654435761) >>> 24);
    const text = enc('Quarterly totals\tNorth\t1200\nQuarterly totals\tSouth\t900\n');
    const mixed = new Uint8Array([...junk.slice(0, 2000), ...text, ...junk.slice(2000)]);
    const r = await recoverWorkbook(buf(mixed), 'mystery.bin');
    expect(r.textOnly).toBe(true);
    expect(r.methods.find((m) => m.id === 'strings')?.status).toBe('partial');
    const vals = [...r.book!.sheets[0].cells.values()].map((c) => String(c.v));
    expect(vals.some((v) => v.includes('North') || v.includes('Quarterly'))).toBe(true);
    expect(reportText(r)).toMatch(/Strings scan/);
  });

  it('reports an empty file', async () => {
    const r = await recoverWorkbook(new ArrayBuffer(0), 'empty.xlsx');
    expect(r.book).toBeNull();
    expect(r.summary).toMatch(/empty/);
  });

  it('reports an encrypted Office package clearly', async () => {
    const r = await recoverWorkbook(buf(cfb([['EncryptionInfo', new Uint8Array(200).fill(4)], ['EncryptedPackage', new Uint8Array(5000).fill(7)]])), 'secret.xlsx');
    expect(r.kind).toBe('encrypted');
    expect(r.book).toBeNull();
    expect(r.summary).toMatch(/password-protected/);
  });
});

// ------------------------------------------------------------------------------------------
// BIFF writer (test helper): records, a BIFF8 workbook, and a compound-file container

const le16 = (n: number) => [n & 255, (n >>> 8) & 255];
const le32 = (n: number) => [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255];
const f64 = (n: number) => [...new Uint8Array(new Float64Array([n]).buffer)];
const rec = (id: number, body: number[]) => [...le16(id), ...le16(body.length), ...body];
const latin = (s: string) => [...s].map((ch) => ch.charCodeAt(0) & 255);
const u16s = (s: string) => [...s].flatMap((ch) => le16(ch.charCodeAt(0)));
const xstr = (s: string, cch8 = false) => { const hi = /[^\x00-\xff]/.test(s); return [...(cch8 ? [s.length] : le16(s.length)), hi ? 1 : 0, ...(hi ? u16s(s) : latin(s))]; };

const LONG = 'Long-' + 'abcdefghij'.repeat(3);
const STD_LINK = [0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
const URL_MONIKER = [0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];

function biff8Stream(o: { filepass?: boolean } = {}): Uint8Array {
  const g: number[] = [];
  g.push(...rec(0x0809, [...le16(0x0600), ...le16(0x0005), ...le16(0x1234), ...le16(1997), ...le32(0), ...le32(0x0106)]));
  if (o.filepass) g.push(...rec(0x002f, [...le16(1), ...le16(1), ...le16(1), ...Array(48).fill(0)]));
  g.push(...rec(0x0042, le16(1200)));
  g.push(...rec(0x0022, le16(0)));
  const font = (h: number, bold: boolean, color: number) => rec(0x0031, [...le16(h), ...le16(0), ...le16(color), ...le16(bold ? 700 : 400), ...le16(0), 0, 0, 0, 0, ...xstr('Arial', true)]);
  g.push(...font(200, false, 0x7fff), ...font(240, true, 10));
  g.push(...rec(0x041e, [...le16(164), ...xstr('0.000')]));
  const xf = (ifnt: number, ifmt: number, align: number, pattern: number, fore: number) =>
    rec(0x00e0, [...le16(ifnt), ...le16(ifmt), ...le16(1), align, 0, 0, 0, ...le32(0), ...le32((pattern << 26) >>> 0), ...le16(fore | (65 << 7))]);
  g.push(...xf(0, 0, 0x20, 0, 64), ...xf(1, 0, 0x2a, 1, 13), ...xf(0, 164, 0x20, 0, 64));
  g.push(...rec(0x01ae, [...le16(2), 0x01, 0x04]));
  g.push(...rec(0x0017, [...le16(2), ...le16(0), ...le16(1), ...le16(1), ...le16(0), ...le16(0), ...le16(0)]));
  const nm = (gr: number, chars: number[], itab: number, rgce: number[]) => rec(0x18, [...le16(gr), 0, chars.length, ...le16(rgce.length), ...le16(0), ...le16(itab), 0, 0, 0, 0, 0, ...chars, ...rgce]);
  g.push(...nm(0, latin('MyName'), 0, [0x3a, ...le16(0), ...le16(0), ...le16(0)]));
  g.push(...nm(0x20, [6], 1, [0x3b, ...le16(1), ...le16(0), ...le16(9), ...le16(0), ...le16(3)]));
  const bsAt: number[] = [];
  for (const [n, st] of [['Data', 0], ['Data 2', 1]] as [string, number][]) { bsAt.push(g.length + 4); g.push(...rec(0x85, [...le32(0), st, 0, ...xstr(n, true)])); }
  const rich = [...le16(4), 0x08, ...le16(1), ...latin('Rich'), ...le16(0), ...le16(1)];
  g.push(...rec(0xfc, [...le32(8), ...le32(6), ...xstr('Name'), ...xstr('Value'), ...xstr('Café Ω'), ...rich, ...le16(LONG.length), 0, ...latin(LONG.slice(0, 10))]));
  g.push(...rec(0x3c, [1, ...u16s(LONG.slice(10)), ...xstr('After')]));
  g.push(...rec(0x0a, []));

  const s: number[] = [];
  const bof = () => rec(0x0809, [...le16(0x0600), ...le16(0x0010), ...le16(0), ...le16(0), ...le32(0), ...le32(0)]);
  const labelsst = (r: number, c: number, x: number, i: number) => rec(0xfd, [...le16(r), ...le16(c), ...le16(x), ...le32(i)]);
  const number = (r: number, c: number, x: number, v: number) => rec(0x203, [...le16(r), ...le16(c), ...le16(x), ...f64(v)]);
  const rk = (r: number, c: number, v: number) => rec(0x27e, [...le16(r), ...le16(c), ...le16(0), ...le32(v)]);
  const formula = (r: number, c: number, cached: number[], rgce: number[], grbit = 0) => rec(0x06, [...le16(r), ...le16(c), ...le16(0), ...cached, ...le16(grbit), ...le32(0), ...le16(rgce.length), ...rgce]);
  const ref = (r: number, c: number) => [...le16(r), ...le16(c | 0xc000)];
  s.push(...bof());
  s.push(...rec(0x7d, [...le16(0), ...le16(0), ...le16(20 * 256), ...le16(0), ...le16(0), ...le16(0)]));
  s.push(...rec(0x208, [...le16(1), ...le16(0), ...le16(5), ...le16(600), ...le16(0), ...le16(0), ...le16(0x0140), ...le16(0x0f)]));
  s.push(...labelsst(0, 0, 1, 0), ...labelsst(0, 1, 0, 1));
  s.push(...labelsst(1, 0, 0, 2), ...number(1, 1, 2, 3.5));
  s.push(...labelsst(2, 0, 0, 4), ...rk(2, 1, (42 << 2) | 2));
  s.push(...labelsst(3, 0, 0, 3), ...rk(3, 1, (125 << 2) | 3));
  s.push(...labelsst(4, 0, 0, 5), ...rec(0xbd, [...le16(4), ...le16(5), ...le16(0), ...le32((10 << 2) | 2), ...le16(0), ...le32((20 << 2) | 2), ...le16(0), ...le32((30 << 2) | 2), ...le16(7)]));
  s.push(...formula(1, 2, f64(7), [0x24, ...ref(1, 1), 0x1e, ...le16(2), 0x05]));
  s.push(...formula(2, 2, [0, 0, 0, 0, 0, 0, 0xff, 0xff], [0x24, ...ref(0, 0), 0x17, 1, 0, 0x21, 0x08]), ...rec(0x207, xstr('Name!')));
  s.push(...formula(3, 2, f64(46.75), [0x25, ...le16(1), ...le16(3), ...le16(1 | 0xc000), ...le16(1 | 0xc000), 0x19, 0x10, ...le16(0)]));
  s.push(...formula(4, 2, f64(200), [0x3a, ...le16(0), ...le16(0), ...le16(0), 0x23, ...le16(1), ...le16(0), 0x03]));
  s.push(...formula(5, 2, [1, 0, 1, 0, 0, 0, 0xff, 0xff], [0x24, ...ref(1, 1), 0x1e, ...le16(1), 0x0d, 0x1d, 1, 0x1d, 0, 0x42, 3, ...le16(1)]));
  s.push(...formula(6, 2, [2, 0, 7, 0, 0, 0, 0xff, 0xff], [0x1e, ...le16(1), 0x1e, ...le16(0), 0x06]));
  s.push(...formula(7, 2, f64(4), [0x24, ...ref(1, 1), 0x1e, ...le16(0), 0x41, ...le16(27)]));
  s.push(...formula(1, 3, f64(4.5), [0x01, ...le16(1), ...le16(3)], 0x08));
  const shr = [0x2c, ...le16(0), ...le16(0xfe | 0xc000), 0x1e, ...le16(1), 0x03];
  s.push(...rec(0x4bc, [...le16(1), ...le16(3), 3, 3, 0, 3, ...le16(shr.length), ...shr]));
  s.push(...formula(2, 3, f64(43), [0x01, ...le16(1), ...le16(3)], 0x08));
  s.push(...formula(3, 3, f64(2.25), [0x01, ...le16(1), ...le16(3)], 0x08));
  s.push(...rec(0x205, [...le16(1), ...le16(4), ...le16(0), 1, 0]), ...rec(0x205, [...le16(2), ...le16(4), ...le16(0), 0x2a, 1]));
  s.push(...labelsst(0, 5, 0, 1), ...rec(0xe5, [...le16(1), ...le16(0), ...le16(1), ...le16(5), ...le16(6)]));
  s.push(...labelsst(7, 0, 0, 0));
  const url = 'https://example.com/';
  s.push(...rec(0x1b8, [...le16(7), ...le16(7), ...le16(0), ...le16(0), ...STD_LINK, ...le32(2), ...le32(3), ...URL_MONIKER, ...le32((url.length + 1) * 2), ...u16s(url), 0, 0]));
  s.push(...rec(0x5d, [...le16(0x15), ...le16(0x12), ...le16(0x19), ...le16(1), ...le16(0), ...Array(12).fill(0), ...le16(0), ...le16(0)]));
  s.push(...rec(0x1b6, [...le16(0x12), ...le16(0), 0, 0, 0, 0, 0, 0, ...le16(9), ...le16(16), ...le32(0)]), ...rec(0x3c, [0, ...latin('Note text')]), ...rec(0x3c, Array(16).fill(0)));
  s.push(...rec(0x1c, [...le16(0), ...le16(0), ...le16(0), ...le16(1), ...xstr('Paul'), 0]));
  s.push(...rec(0x23e, [...le16(0x02 | 0x04 | 0x08 | 0x10 | 0x20 | 0x80 | 0x100 | 0x200 | 0x400), ...le16(0), ...le16(0), ...le32(64), ...le16(0), ...le16(0), ...le32(0)]));
  s.push(...rec(0x41, [...le16(1), ...le16(1), ...le16(1), ...le16(1), 3, 0]));
  s.push(...rec(0x0a, []));

  const s2 = [...bof(), ...number(0, 0, 0, 100), ...rec(0x0a, [])];
  const pos = [g.length, g.length + s.length];
  bsAt.forEach((at, i) => g.splice(at, 4, ...le32(pos[i])));
  return new Uint8Array([...g, ...s, ...s2]);
}

/** Minimal compound file (512-byte sectors, no mini stream: streams are padded to 4096 bytes). */
function cfb(streams: [string, Uint8Array][]): Uint8Array {
  const SEC = 512, END = 0xfffffffe, FREE = 0xffffffff, NONE = 0xffffffff;
  const padded = streams.map(([n, d]) => { const size = Math.max(d.length, 4096); const out = new Uint8Array(Math.ceil(size / SEC) * SEC); out.set(d); return { n, out, size }; });
  const total = 2 + padded.reduce((a, p) => a + p.out.length / SEC, 0);
  const file = new Uint8Array(SEC * (total + 1));
  const dv = new DataView(file.buffer);
  [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].forEach((b, i) => { file[i] = b; });
  dv.setUint16(0x18, 0x3e, true); dv.setUint16(0x1a, 3, true); dv.setUint16(0x1c, 0xfffe, true); dv.setUint16(0x1e, 9, true); dv.setUint16(0x20, 6, true);
  dv.setUint32(0x2c, 1, true); dv.setUint32(0x30, 1, true); dv.setUint32(0x38, 4096, true); dv.setUint32(0x3c, END, true); dv.setUint32(0x44, END, true);
  for (let k = 0; k < 109; k++) dv.setUint32(0x4c + k * 4, k === 0 ? 0 : FREE, true);
  const fat = new Array(128).fill(FREE);
  fat[0] = 0xfffffffd; fat[1] = END;
  let next = 2;
  const starts: number[] = [];
  for (const p of padded) {
    const n = p.out.length / SEC;
    starts.push(next);
    for (let k = 0; k < n; k++) fat[next + k] = k === n - 1 ? END : next + k + 1;
    file.set(p.out, SEC * (next + 1));
    next += n;
  }
  fat.forEach((v, k) => dv.setUint32(SEC + k * 4, v >>> 0, true));
  const dir = (i: number, name: string, type: number, child: number, right: number, start: number, size: number) => {
    const o = SEC * 2 + i * 128;
    [...name].forEach((ch, k) => dv.setUint16(o + k * 2, ch.charCodeAt(0), true));
    dv.setUint16(o + 0x40, (name.length + 1) * 2, true);
    file[o + 0x42] = type; file[o + 0x43] = 1;
    dv.setUint32(o + 0x44, NONE, true); dv.setUint32(o + 0x48, right, true); dv.setUint32(o + 0x4c, child, true);
    dv.setUint32(o + 0x74, start, true); dv.setUint32(o + 0x78, size, true);
  };
  dir(0, 'Root Entry', 5, 1, NONE, END, 0);
  padded.forEach((p, i) => dir(i + 1, p.n, 2, NONE, i + 2 <= padded.length ? i + 2 : NONE, starts[i], p.size));
  return file;
}

function checkBiff8(book: Workbook): void {
  const sh = book.sheetByName('Data')!;
  expect(book.sheets.map((s) => s.name)).toEqual(['Data', 'Data 2']);
  expect(sh.value(0, 0)).toBe('Name');
  expect(sh.value(1, 0)).toBe('Café Ω');
  expect(sh.value(2, 0)).toBe(LONG);
  expect(sh.value(3, 0)).toBe('Rich');
  expect(sh.value(4, 0)).toBe('After');
  expect(sh.value(1, 1)).toBe(3.5);
  expect(sh.value(2, 1)).toBe(42);
  expect(sh.value(3, 1)).toBe(1.25);
  expect([sh.value(4, 5), sh.value(4, 6), sh.value(4, 7)]).toEqual([10, 20, 30]);
}

describe('Excel 97-2003 (BIFF8) reader', () => {
  const xls = cfb([['Workbook', biff8Stream()]]);

  it('reads values, shared strings across CONTINUE, formulas, names and layout', () => {
    const { book, stats, warnings } = readBiff(xls);
    expect(warnings).toEqual([]);
    expect(stats.biff).toBe(8);
    checkBiff8(book);
    const sh = book.sheetByName('Data')!;
    const f = (r: number, c: number) => sh.get(r, c)?.f;
    expect([f(1, 2), sh.value(1, 2)]).toEqual(['B2*2', 7]);
    expect([f(2, 2), sh.value(2, 2)]).toEqual(['A1&"!"', 'Name!']);
    expect([f(3, 2), sh.value(3, 2)]).toEqual(['SUM(B2:B4)', 46.75]);
    expect(f(4, 2)).toBe("'Data 2'!$A$1+MyName");
    expect([f(5, 2), sh.value(5, 2)]).toEqual(['IF(B2>1,TRUE,FALSE)', true]);
    expect(f(6, 2)).toBe('1/0');
    expect((sh.value(6, 2) as CellError).code).toBe('#DIV/0!');
    expect(f(7, 2)).toBe('ROUND(B2,0)');
    expect([f(1, 3), f(2, 3), f(3, 3)]).toEqual(['B2+1', 'B3+1', 'B4+1']);
    expect(sh.value(3, 3)).toBe(2.25);
    expect(sh.value(1, 4)).toBe(true);
    expect((sh.value(2, 4) as CellError).code).toBe('#N/A');
    expect(stats.formulasDecoded).toBe(stats.formulas);
    // Styles: bold red 12pt, centered, wrapped, yellow fill; number format 0.000
    const st = book.styles[sh.get(0, 0)!.s!];
    expect(st).toMatchObject({ bold: true, size: 12, color: '#ff0000', hAlign: 'center', wrap: true, fill: '#ffff00' });
    expect(book.styles[sh.get(1, 1)!.s!].numFmt).toBe('0.000');
    // Layout
    expect(sh.merges).toEqual([{ r1: 0, r2: 1, c1: 5, c2: 6 }]);
    expect(sh.colW.get(0)).toBe(140);
    expect(sh.rowH.get(1)).toBe(40);
    expect(sh.freeze).toEqual({ rows: 1, cols: 1 });
    expect(sh.comments.get(0)?.text).toBe('Note text');
    expect(sh.comments.get(0)?.author).toBe('Paul');
    expect(sh.get(7, 0)?.link).toBe('https://example.com/');
    expect(book.sheetByName('Data 2')!.visibility).toBe('hidden');
    expect(book.names.find((n) => n.name === 'MyName')?.ref).toBe("'Data 2'!$A$1");
    expect(sh.print.area).toBe('$A$1:$D$10');
  });

  it('has the standard function table', () => {
    expect(BIFF_FUNCTIONS.size).toBeGreaterThan(360);
    expect(BIFF_FUNCTIONS.get(1)?.[0]).toBe('IF');
    expect(BIFF_FUNCTIONS.get(102)).toEqual(['VLOOKUP', -1]);
    expect(BIFF_FUNCTIONS.get(346)).toEqual(['COUNTIF', 2]);
  });

  it('imports through importXls and opens via recoverWorkbook', async () => {
    const book = await importXls(buf(xls), 'book.xls');
    checkBiff8(book);
    const r = await recoverWorkbook(buf(xls), 'book.xls');
    expect(r.kind).toBe('xls');
    expect(r.methods.find((m) => m.id === 'open')?.status).toBe('success');
    checkBiff8(r.book!);
  });

  it('salvages a workbook whose compound-file header is destroyed', async () => {
    const bad = xls.slice();
    bad.fill(0, 0, 512);
    expect(() => readBiff(bad)).toThrow();
    const r = await recoverWorkbook(buf(bad), 'broken.xls');
    expect(r.methods.find((m) => m.id === 'biff')?.detail).toMatch(/raw-byte scan/);
    checkBiff8(r.book!);
    expect(r.book!.sheetByName('Data')!.get(2, 3)?.f).toBe('B3+1');
  });

  it('resynchronizes after a damaged record', async () => {
    const stream = biff8Stream();
    // Corrupt the length of the first LABELSST in the Data sheet.
    let at = -1;
    for (let i = 0; i < stream.length - 4; i++) if (stream[i] === 0xfd && stream[i + 1] === 0 && stream[i + 2] === 10 && stream[i + 3] === 0) { at = i; break; }
    stream[at + 2] = 0xff; stream[at + 3] = 0xff;
    const file = cfb([['Workbook', stream]]);
    expect(() => readBiff(file)).toThrow();
    const { book, stats } = readBiff(file, { lenient: true });
    expect(stats.resyncs).toBeGreaterThan(0);
    const sh = book.sheetByName('Data')!;
    expect(sh.value(1, 1)).toBe(3.5);
    expect(sh.get(3, 3)?.f).toBe('B4+1');
    const r = await recoverWorkbook(buf(file), 'resync.xls');
    expect(r.bestMethod).toBe('biff');
    expect(r.methods.find((m) => m.id === 'biff')?.status).toBe('partial');
  });

  it('recovers the data present in a truncated .xls', async () => {
    const r = await recoverWorkbook(buf(xls.slice(0, 512 * 6)), 'short.xls');
    expect(r.book).not.toBeNull();
    const sh = r.book!.sheets.find((s) => s.name === 'Data')!;
    expect(sh.value(0, 0)).toBe('Name');
  });

  it('reports encrypted workbooks (FILEPASS)', async () => {
    const file = cfb([['Workbook', biff8Stream({ filepass: true })]]);
    expect(() => readBiff(file)).toThrow(XLS_ENCRYPTED_MSG);
    const r = await recoverWorkbook(buf(file), 'locked.xls');
    expect(r.book).toBeNull();
    expect(r.summary).toBe(XLS_ENCRYPTED_MSG);
  });
});

describe('Excel 5.0/95 and BIFF2', () => {
  it('reads a BIFF5 workbook with a Windows-1252 code page', () => {
    const g: number[] = [];
    const bof5 = (dt: number) => rec(0x809, [...le16(0x0500), ...le16(dt), ...le16(0), ...le16(0)]);
    g.push(...bof5(5), ...rec(0x42, le16(1252)));
    g.push(...rec(0x31, [...le16(200), ...le16(0), ...le16(0x7fff), ...le16(400), ...le16(0), 0, 0, 0, 0, 5, ...latin('Arial')]));
    g.push(...rec(0xe0, [...le16(0), ...le16(0), ...le16(1), ...le16(0x20), ...le32(0), ...le32(0)]));
    const bs = g.length + 4;
    g.push(...rec(0x85, [...le32(0), 0, 0, 7, ...latin('Feuille')]), ...rec(0x0a, []));
    const s = [...bof5(0x10),
      ...rec(0x204, [...le16(0), ...le16(0), ...le16(0), ...le16(6), 0x43, 0x61, 0x66, 0xe9, 0x20, 0x80]),
      ...rec(0x203, [...le16(1), ...le16(1), ...le16(0), ...f64(21)]),
      ...rec(0x06, [...le16(1), ...le16(2), ...le16(0), ...f64(42), ...le16(0), ...le32(0), ...le16(8), 0x24, ...le16(1 | 0xc000), 1, 0x1e, ...le16(2), 0x05]),
      ...rec(0x0a, [])];
    g.splice(bs, 4, ...le32(g.length));
    const { book, stats } = readBiff(cfb([['Book', new Uint8Array([...g, ...s])]]));
    expect(stats.biff).toBe(5);
    const sh = book.sheets[0];
    expect(sh.name).toBe('Feuille');
    expect(sh.value(0, 0)).toBe('Café €');
    expect(sh.get(1, 2)?.f).toBe('B2*2');
    expect(sh.value(1, 2)).toBe(42);
  });

  it('reads a raw BIFF2 worksheet', async () => {
    const bytes = new Uint8Array([
      ...rec(0x0009, [...le16(2), ...le16(0x10)]),
      ...rec(0x0004, [...le16(0), ...le16(0), 0, 0, 0, 5, ...latin('Hello')]),
      ...rec(0x0003, [...le16(1), ...le16(0), 0, 0, 0, ...f64(2.5)]),
      ...rec(0x0002, [...le16(2), ...le16(0), 0, 0, 0, ...le16(7)]),
      ...rec(0x000a, []),
    ]);
    const r = await recoverWorkbook(buf(bytes), 'old.xls');
    expect(r.kind).toBe('xls');
    const sh = r.book!.sheets[0];
    expect([sh.value(0, 0), sh.value(1, 0), sh.value(2, 0)]).toEqual(['Hello', 2.5, 7]);
  });
});
