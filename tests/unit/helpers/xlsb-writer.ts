/**
 * Test helper: writes a small Excel binary workbook (.xlsb) from the record layouts published in
 * [MS-XLSB], to exercise the reader without third-party sample files. The same bytes were opened
 * in LibreOffice (which reads .xlsb) to confirm the layouts.
 */
import JSZip from 'jszip';

const enc = {
  u8: (n: number) => [n & 255],
  u16: (n: number) => [n & 255, (n >> 8) & 255],
  u32: (n: number) => [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255],
  f64: (n: number) => { const b = new DataView(new ArrayBuffer(8)); b.setFloat64(0, n, true); return [...new Uint8Array(b.buffer)]; },
  wstr: (s: string) => { const out = enc.u32(s.length); for (let i = 0; i < s.length; i++) out.push(...enc.u16(s.charCodeAt(i))); return out; },
  nwstr: (s: string | null) => (s == null ? enc.u32(0xffffffff) : enc.wstr(s)),
  /** BrtColor: type 0 auto, 1 indexed, 2 rgb, 3 theme. */
  color: (type: number, index = 0, rgb = '000000', tint = 0) => [1 | (type << 1), index, ...enc.u16(Math.round(tint * 32767) & 0xffff), parseInt(rgb.slice(0, 2), 16), parseInt(rgb.slice(2, 4), 16), parseInt(rgb.slice(4, 6), 16), 255],
  rfx: (r1: number, r2: number, c1: number, c2: number) => [...enc.u32(r1), ...enc.u32(r2), ...enc.u32(c1), ...enc.u32(c2)],
};

class Stream {
  bytes: number[] = [];
  rec(type: number, payload: number[] = []): this {
    if (type < 128) this.bytes.push(type); else this.bytes.push((type & 0x7f) | 0x80, type >> 7);
    let n = payload.length;
    do { let b = n & 0x7f; n >>= 7; if (n) b |= 0x80; this.bytes.push(b); } while (n);
    this.bytes.push(...payload);
    return this;
  }
  get u8(): Uint8Array { return new Uint8Array(this.bytes); }
}

/** Parsed-expression tokens. */
const ptg = {
  ref: (row: number, col: number, rowRel = true, colRel = true, cls = 0x20) => [0x24 | cls, ...enc.u32(row), ...enc.u16(col | (colRel ? 0x4000 : 0) | (rowRel ? 0x8000 : 0))],
  refN: (dr: number, dc: number) => [0x4c, ...enc.u32(dr), ...enc.u16((dc & 0x3fff) | 0xc000)],
  area: (r1: number, r2: number, c1: number, c2: number, rel = true) => [0x25, ...enc.u32(r1), ...enc.u32(r2), ...enc.u16(c1 | (rel ? 0xc000 : 0)), ...enc.u16(c2 | (rel ? 0xc000 : 0))],
  ref3d: (ixti: number, row: number, col: number) => [0x5a, ...enc.u16(ixti), ...enc.u32(row), ...enc.u16(col | 0xc000)],
  area3dAbs: (ixti: number, r1: number, r2: number, c1: number, c2: number) => [0x3b, ...enc.u16(ixti), ...enc.u32(r1), ...enc.u32(r2), ...enc.u16(c1), ...enc.u16(c2)],
  int: (n: number) => [0x1e, ...enc.u16(n)],
  num: (n: number) => [0x1f, ...enc.f64(n)],
  str: (s: string) => { const out = [0x17, ...enc.u16(s.length)]; for (let i = 0; i < s.length; i++) out.push(...enc.u16(s.charCodeAt(i))); return out; },
  func: (idx: number) => [0x41, ...enc.u16(idx)],
  funcVar: (n: number, idx: number) => [0x42, n, ...enc.u16(idx)],
  name: (i: number) => [0x43, ...enc.u32(i)],
  attrSum: () => [0x19, 0x10, 0, 0],
  /** Shared/array formula reference: the row is in rgce, the column in the extra data (rgcb). */
  exp: (row: number) => [0x01, ...enc.u32(row)],
  array: () => [0x60, ...new Array(14).fill(0)],
  list: (listIndex: number, colMode: number, rowType: number, c1: number, c2: number) => [0x18, 0x19, ...enc.u16(0), ...enc.u16(colMode | (rowType << 2)), ...enc.u32(listIndex), ...enc.u16(c1), ...enc.u16(c2)],
  mul: [0x05], gt: [0x0d], div: [0x06],
};

const fmla = (rgce: number[], rgcb: number[] = []) => [...enc.u32(rgce.length), ...rgce, ...enc.u32(rgcb.length), ...rgcb];
const cell = (col: number, xf: number) => [...enc.u32(col), ...enc.u32(xf & 0xffffff)];

export async function buildXlsb(): Promise<ArrayBuffer> {
  // ---- workbook.bin
  const wb = new Stream()
    .rec(131)
    .rec(153, [...enc.u32(0), ...enc.u32(0), ...enc.wstr('')])
    .rec(135)
    .rec(158, [...enc.u32(0), ...enc.u32(0), ...enc.u32(20000), ...enc.u32(12000), ...enc.u32(600), ...enc.u32(0), ...enc.u32(0), 0x78])
    .rec(136)
    .rec(143)
    .rec(156, [...enc.u32(0), ...enc.u32(1), ...enc.nwstr('rId1'), ...enc.wstr('Fruit')])
    .rec(156, [...enc.u32(1), ...enc.u32(2), ...enc.nwstr('rId2'), ...enc.wstr('Notes & More')])
    .rec(144)
    .rec(353)
    .rec(357)
    .rec(362, [...enc.u32(2), ...enc.u32(0), ...enc.u32(0), ...enc.u32(0), ...enc.u32(0), ...enc.u32(1), ...enc.u32(1)])
    .rec(354)
    // Defined names: TaxRate = Fruit!$C$2; Print_Area of Fruit = $A$1:$D$5
    .rec(39, [...enc.u32(0), 0, ...enc.u32(0xffffffff), ...enc.wstr('TaxRate'), ...fmla([0x3a, ...enc.u16(0), ...enc.u32(1), ...enc.u16(2)]), ...enc.nwstr(null)])
    .rec(39, [...enc.u32(0x20), 0, ...enc.u32(0), ...enc.wstr('_xlnm.Print_Area'), ...fmla(ptg.area3dAbs(0, 0, 4, 0, 3)), ...enc.nwstr(null)])
    .rec(157, [...enc.u32(0x1dd63), ...enc.u32(1), ...enc.u32(100), ...enc.f64(0.001), ...enc.u32(1), 0x6a, 0])
    .rec(132);

  // ---- styles.bin
  const font = (h: number, flags: number, bls: number, uls: number, color: number[], name: string) => [...enc.u16(h), ...enc.u16(flags), ...enc.u16(bls), ...enc.u16(0), uls, 2, 0, 0, ...color, 2, ...enc.wstr(name)];
  const fill = (fls: number, fg: number[], bg: number[]) => [...enc.u32(fls), ...fg, ...bg, ...new Array(48).fill(0)];
  const blxf = (dg: number, color: number[]) => [dg, 0, ...color];
  const border = (bottom: number) => [0, ...blxf(0, enc.color(0)), ...blxf(bottom, enc.color(2, 0, '000000')), ...blxf(0, enc.color(0)), ...blxf(0, enc.color(0)), ...blxf(0, enc.color(0))];
  const xf = (parent: number, fmt: number, fnt: number, fil: number, bdr: number, flags: number) => [...enc.u16(parent), ...enc.u16(fmt), ...enc.u16(fnt), ...enc.u16(fil), ...enc.u16(bdr), 0, 0, ...enc.u16(flags), 0, 0];
  const LOCKED = 0x1000, BOTTOM = 2 << 3;
  const st = new Stream()
    .rec(278)
    .rec(615, enc.u32(1)).rec(44, [...enc.u16(164), ...enc.wstr('#,##0.000')]).rec(616)
    .rec(611, enc.u32(3))
    .rec(43, font(220, 0, 400, 0, enc.color(3, 1), 'Calibri'))
    .rec(43, font(240, 0, 700, 0, enc.color(2, 0, 'c00000'), 'Arial'))
    .rec(43, font(220, 2, 400, 1, enc.color(2, 0, '0000ff'), 'Calibri'))
    .rec(612)
    .rec(603, enc.u32(3))
    .rec(45, fill(0, enc.color(1, 64), enc.color(1, 65)))
    .rec(45, fill(17, enc.color(1, 64), enc.color(1, 65)))
    .rec(45, fill(1, enc.color(2, 0, 'ffff00'), enc.color(1, 64)))
    .rec(604)
    .rec(613, enc.u32(2)).rec(46, border(0)).rec(46, border(1)).rec(614)
    .rec(626, enc.u32(1)).rec(47, xf(0xffff, 0, 0, 0, 0, LOCKED | BOTTOM)).rec(627)
    .rec(617, enc.u32(6))
    .rec(47, xf(0, 0, 0, 0, 0, LOCKED | BOTTOM))
    .rec(47, xf(0, 0, 1, 2, 1, LOCKED | BOTTOM | 2))
    .rec(47, xf(0, 164, 2, 0, 0, LOCKED | BOTTOM))
    .rec(47, xf(0, 14, 0, 0, 0, LOCKED | BOTTOM))
    .rec(47, xf(0, 0, 0, 0, 0, LOCKED | 0x40))
    .rec(47, xf(0, 0, 0, 0, 0, BOTTOM))
    .rec(618)
    .rec(619, enc.u32(1)).rec(48, [...enc.u32(0), ...enc.u16(1), 0, 0xff, ...enc.wstr('Normal')]).rec(620)
    .rec(279);

  // ---- sharedStrings.bin
  const strs = ['Name', 'Qty', 'Price', 'Total', 'Apple', 'Pear', 'Plum', 'Merged title'];
  const sst = new Stream().rec(159, [...enc.u32(9), ...enc.u32(9)]);
  for (const s of strs) sst.rec(19, [0, ...enc.wstr(s)]);
  sst.rec(19, [1, ...enc.wstr('Bold part plain'), ...enc.u32(2), ...enc.u16(0), ...enc.u16(1), ...enc.u16(9), ...enc.u16(0)]);
  sst.rec(160);

  // ---- sheet1.bin (Fruit)
  const row = (r: number, ht = 0, flags = 0) => [...enc.u32(r), ...enc.u32(0), ...enc.u16(ht || 300), 0, flags, 0, ...enc.u32(0)];
  const s1 = new Stream()
    .rec(129)
    .rec(147, [...enc.u16(0x04c9), 0x02, ...enc.color(2, 0, 'ff8800'), ...enc.u32(0xffffffff), ...enc.u32(0xffffffff), ...enc.wstr('Sheet1Code')])
    .rec(148, enc.rfx(0, 9, 0, 10))
    .rec(133)
    .rec(137, [...enc.u16(0x03dc), ...enc.u32(0), ...enc.u32(0), ...enc.u32(0), 0x40, 0, ...enc.u16(125), ...enc.u16(100), ...enc.u16(0), ...enc.u16(0), ...enc.u32(0), 0, 0])
    .rec(151, [...enc.f64(0), ...enc.f64(1), ...enc.u32(1), ...enc.u32(0), ...enc.u32(2), 0x01])
    .rec(138)
    .rec(134)
    .rec(485, [...enc.u32(0xffffffff), ...enc.u16(8), ...enc.u16(300)])
    .rec(390)
    .rec(60, [...enc.u32(0), ...enc.u32(0), ...enc.u32(20 * 256), ...enc.u32(0), ...enc.u16(2)])
    .rec(60, [...enc.u32(5), ...enc.u32(5), ...enc.u32(9 * 256), ...enc.u32(0), ...enc.u16(1)])
    .rec(391)
    .rec(145)
    // Row 1: headers (custom height 30pt)
    .rec(0, row(0, 600, 0x40))
    .rec(7, [...cell(0, 1), ...enc.u32(0)]).rec(7, [...cell(1, 1), ...enc.u32(1)]).rec(7, [...cell(2, 1), ...enc.u32(2)]).rec(7, [...cell(3, 1), ...enc.u32(3)])
    // Row 2: Apple; D2:D4 share B*C; H2 = 3D reference; J..K table
    .rec(0, row(1))
    .rec(7, [...cell(0, 0), ...enc.u32(4)])
    .rec(2, [...cell(1, 0), ...enc.u32((3 << 2) | 2)])
    .rec(5, [...cell(2, 2), ...enc.f64(1.25)])
    .rec(9, [...cell(3, 2), ...enc.f64(3.75), ...enc.u16(0), ...fmla(ptg.exp(1), enc.u32(3))])
    .rec(427, [...enc.rfx(1, 3, 3, 3), ...fmla([...ptg.refN(0, -2), ...ptg.refN(0, -1), ...ptg.mul])])
    .rec(5, [...cell(4, 3), ...enc.f64(45366)])
    .rec(8, [...cell(7, 0), ...enc.wstr('Tricky <&> "quotes"'), ...enc.u16(0), ...fmla(ptg.ref3d(1, 0, 0))])
    .rec(6, [...cell(9, 0), ...enc.wstr('Region')]).rec(6, [...cell(10, 0), ...enc.wstr('Units')])
    // Row 3
    .rec(0, row(2))
    .rec(7, [...cell(0, 0), ...enc.u32(5)])
    .rec(2, [...cell(1, 0), ...enc.u32((5 << 2) | 2)])
    .rec(5, [...cell(2, 2), ...enc.f64(2.5)])
    .rec(9, [...cell(3, 2), ...enc.f64(12.5), ...enc.u16(0), ...fmla(ptg.exp(1), enc.u32(3))])
    .rec(9, [...cell(7, 0), ...enc.f64(12.5), ...enc.u16(0), ...fmla([...ptg.name(1), ...ptg.int(10), ...ptg.mul])])
    .rec(6, [...cell(9, 0), ...enc.wstr('North')]).rec(2, [...cell(10, 0), ...enc.u32((4 << 2) | 2)])
    // Row 4
    .rec(0, row(3))
    .rec(7, [...cell(0, 0), ...enc.u32(6)])
    .rec(2, [...cell(1, 0), ...enc.u32((12 << 2) | 2)])
    .rec(2, [...cell(2, 2), ...enc.u32((75 << 2) | 2 | 1)])
    .rec(9, [...cell(3, 2), ...enc.f64(9), ...enc.u16(0), ...fmla(ptg.exp(1), enc.u32(3))])
    .rec(9, [...cell(7, 0), ...enc.f64(10), ...enc.u16(0), ...fmla([...ptg.array(), ...ptg.funcVar(1, 4)], [...enc.u32(2), ...enc.u32(2), 0, ...enc.f64(1), 0, ...enc.f64(2), 0, ...enc.f64(3), 0, ...enc.f64(4)])])
    .rec(6, [...cell(9, 0), ...enc.wstr('South')]).rec(2, [...cell(10, 0), ...enc.u32((6 << 2) | 2)])
    // Row 5: totals
    .rec(0, row(4))
    .rec(6, [...cell(0, 0), ...enc.wstr('Sum')])
    .rec(9, [...cell(3, 2), ...enc.f64(25.25), ...enc.u16(0), ...fmla([...ptg.area(1, 3, 3, 3), ...ptg.attrSum()])])
    .rec(9, [...cell(4, 0), ...enc.f64(25.3), ...enc.u16(0), ...fmla([...ptg.ref(4, 3), ...ptg.int(1), ...ptg.func(27)])])
    .rec(8, [...cell(5, 0), ...enc.wstr('many'), ...enc.u16(0), ...fmla([...ptg.ref(1, 1), ...ptg.int(2), ...ptg.gt, ...ptg.str('many'), ...ptg.str('few'), ...ptg.funcVar(3, 1)])])
    .rec(9, [...cell(9, 0), ...enc.f64(10), ...enc.u16(0), ...fmla([...ptg.list(1, 1, 0, 1, 1), ...ptg.funcVar(1, 4)])])
    // Row 6: other types
    .rec(0, row(5))
    .rec(4, [...cell(0, 0), 1])
    .rec(3, [...cell(1, 0), 0x07])
    .rec(11, [...cell(2, 0), 0x2a, ...enc.u16(0), ...fmla(ptg.func(10))])
    .rec(8, [...cell(3, 0), ...enc.wstr('x'), ...enc.u16(0), ...fmla([...ptg.int(1), ...ptg.int(0), ...ptg.div, ...ptg.str('x'), ...ptg.funcVar(2, 480)])])
    .rec(7, [...cell(6, 0), ...enc.u32(8)])
    .rec(5, [...cell(7, 5), ...enc.f64(-1234.5)]).rec(6, [...cell(8, 4), ...enc.wstr("wrapped")])
    // Row 7: merged title, hyperlink
    .rec(0, row(6))
    .rec(7, [...cell(0, 0), ...enc.u32(7)])
    .rec(6, [...cell(3, 0), ...enc.wstr('site')])
    // Row 8: array formula A8:A9 = B2:B3*2
    .rec(0, row(7))
    .rec(9, [...cell(0, 0), ...enc.f64(6), ...enc.u16(0), ...fmla(ptg.exp(7), enc.u32(0))])
    .rec(426, [...enc.rfx(7, 8, 0, 0), 0, ...fmla([...ptg.area(1, 2, 1, 1), ...ptg.int(2), ...ptg.mul])])
    .rec(0, row(8))
    .rec(9, [...cell(0, 0), ...enc.f64(10), ...enc.u16(0), ...fmla(ptg.exp(7), enc.u32(0))])
    // Row 10: hidden
    .rec(0, row(9, 300, 0x20))
    .rec(6, [...cell(0, 0), ...enc.wstr('hidden row')])
    .rec(146)
    .rec(177, enc.u32(1)).rec(176, enc.rfx(6, 6, 0, 2)).rec(178)
    .rec(494, [...enc.rfx(6, 6, 3, 3), ...enc.nwstr('rId1'), ...enc.wstr(''), ...enc.wstr('Visit'), ...enc.wstr('')])
    .rec(477, enc.u16(0x0009))
    .rec(476, [...enc.f64(0.5), ...enc.f64(0.5), ...enc.f64(1), ...enc.f64(1), ...enc.f64(0.3), ...enc.f64(0.3)])
    .rec(478, [...enc.u32(9), ...enc.u32(90), ...enc.u32(600), ...enc.u32(600), ...enc.u32(1), ...enc.u32(1), ...enc.u32(1), ...enc.u32(1), ...enc.u16(0x0002), ...enc.nwstr(null)])
    .rec(479, [...enc.u16(0), ...enc.nwstr('&LLeft&C&P of &N'), ...enc.nwstr('&RFooter'), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null)])
    .rec(660, enc.u32(1)).rec(405, enc.nwstr('rId3')).rec(661)
    .rec(130);

  // ---- sheet2.bin (Notes & More)
  const s2 = new Stream().rec(129).rec(145)
    .rec(0, row(0))
    .rec(6, [...cell(0, 0), ...enc.wstr('Tricky <&> "quotes"')])
    .rec(146).rec(130);

  // ---- comments1.bin
  const cm = new Stream().rec(628).rec(630).rec(632, enc.wstr('Ann')).rec(631).rec(633)
    .rec(635, [...enc.u32(0), ...enc.rfx(0, 0, 0, 0), ...new Array(16).fill(0)])
    .rec(637, [0, ...enc.wstr('Ann:\nHeader note')])
    .rec(636).rec(634).rec(629);

  // ---- table1.bin: Sales = J1:K3 (header row, no totals)
  const tb = new Stream()
    .rec(431, [...enc.rfx(0, 2, 9, 10), ...enc.u32(0), ...enc.u32(1), ...enc.u32(1), ...enc.u32(0), ...enc.u32(0), ...new Array(28).fill(0), ...enc.nwstr('Table1'), ...enc.nwstr('Sales'), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null)])
    .rec(433, enc.u32(2))
    .rec(435, [...enc.u32(1), ...enc.u32(0), ...new Array(16).fill(0), ...enc.nwstr('Region'), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null)]).rec(436)
    .rec(435, [...enc.u32(2), ...enc.u32(0), ...new Array(16).fill(0), ...enc.nwstr('Units'), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null), ...enc.nwstr(null)]).rec(436)
    .rec(434)
    .rec(513, [...enc.u16(0x0004), ...enc.nwstr('TableStyleMedium9')])
    .rec(432);

  const zip = new JSZip();
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (items: [string, string, string, boolean?][]) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + items.map(([id, type, target, ext]) => `<Relationship Id="${id}" Type="${type.includes('/') ? type : R + '/' + type}" Target="${target}"${ext ? ' TargetMode="External"' : ''}/>`).join('') + '</Relationships>';
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/worksheets/sheet1.bin" ContentType="application/vnd.ms-excel.worksheet"/><Override PartName="/xl/worksheets/sheet2.bin" ContentType="application/vnd.ms-excel.worksheet"/>'
    + '<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/><Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>'
    + '<Override PartName="/xl/comments1.bin" ContentType="application/vnd.ms-excel.comments"/><Override PartName="/xl/tables/table1.bin" ContentType="application/vnd.ms-excel.table"/>'
    + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
  zip.file('_rels/.rels', rels([['rId1', 'officeDocument', 'xl/workbook.bin'], ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml']]));
  zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Binary fruit</dc:title><dc:creator>Ann</dc:creator></cp:coreProperties>');
  zip.file('xl/_rels/workbook.bin.rels', rels([['rId1', 'worksheet', 'worksheets/sheet1.bin'], ['rId2', 'worksheet', 'worksheets/sheet2.bin'], ['rId3', 'styles', 'styles.bin'], ['rId4', 'sharedStrings', 'sharedStrings.bin']]));
  zip.file('xl/worksheets/_rels/sheet1.bin.rels', rels([['rId1', 'hyperlink', 'https://example.com/', true], ['rId2', 'comments', '../comments1.bin'], ['rId3', 'table', '../tables/table1.bin']]));
  zip.file('xl/workbook.bin', wb.u8);
  zip.file('xl/styles.bin', st.u8);
  zip.file('xl/sharedStrings.bin', sst.u8);
  zip.file('xl/worksheets/sheet1.bin', s1.u8);
  zip.file('xl/worksheets/sheet2.bin', s2.u8);
  zip.file('xl/comments1.bin', cm.u8);
  zip.file('xl/tables/table1.bin', tb.u8);
  return zip.generateAsync({ type: 'arraybuffer' });
}
