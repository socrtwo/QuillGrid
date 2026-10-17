/**
 * Excel Binary Workbook (.xlsb) reader.
 *
 * An .xlsb file is an OOXML package (ZIP, [Content_Types].xml, relationships) whose workbook,
 * worksheet, shared string, style and comment parts are binary record streams (BIFF12, as
 * published in the [MS-XLSB] specification) instead of XML. Each record is a variable-length
 * type (1-2 bytes, 7 bits each) and size (1-4 bytes), then its payload. Strings are
 * XLWideString (32-bit character count, UTF-16LE). Formulas are parsed expressions (Ptg tokens)
 * with 32-bit rows and 14-bit columns.
 *
 * Read here: sheets (visibility, tab color, code name), shared strings with rich runs, number
 * formats, fonts, fills, borders, cell formats and named cell styles, cells of every type with
 * cached formula results, formulas (operators, functions including future functions, names,
 * 3D references, shared and array formulas, structured table references, array constants),
 * column widths and styles, row heights and outline levels, merged cells, frozen panes, view
 * settings, hyperlinks, comments, tables, sheet protection, margins, print options, headers and
 * footers, defined names (print areas and titles), the 1904 date system and calculation
 * settings. Pictures, shapes and charts are XML parts in .xlsb too and are read by the .xlsx
 * reader's drawing code. A formula that cannot be decoded keeps its cached value.
 */
import JSZip from 'jszip';
import { Workbook, Sheet, newId } from '../model/workbook';
import { CellError, type Cell, type CellStyle, type Border, type DefinedName, type TableDef } from '../model/types';
import { cellKey, colName, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { BUILTIN_FORMATS } from '../model/numfmt';
import { BIFF_FUNCTIONS } from './recover/xls';
import { parseXml, desc, attr } from './index';
import {
  Pkg, INDEXED_COLORS, applyTint, widthToPx, ptToPx, parseTheme, parseDocProps, readSheetDrawing, readThreadedComments,
  splitHeaderFooter, PAPER, type Rel,
} from './xlsx-import';

// --------------------------------------------------------------------------------------------
// Record stream

/** Record type ids used here ([MS-XLSB] 2.3). */
export const BRT = {
  RowHdr: 0, CellBlank: 1, CellRk: 2, CellError: 3, CellBool: 4, CellReal: 5, CellSt: 6, CellIsst: 7,
  FmlaString: 8, FmlaNum: 9, FmlaBool: 10, FmlaError: 11, SSTItem: 19, Name: 39, Font: 43, Fmt: 44, Fill: 45,
  Border: 46, XF: 47, Style: 48, ColInfo: 60, CellRString: 62, BeginSheet: 129, WsView: 137, WsProp: 147, WsDim: 148,
  Pane: 151, WbProp: 153, BundleSh: 156, CalcProp: 157, BookView: 158, MergeCell: 176, ExternSheet: 362,
  SupBookSrc: 355, SupSelf: 357, SupSame: 358, ArrFmla: 426, ShrFmla: 427, BeginList: 431, BeginListCol: 435,
  ListPart: 405, Margins: 476, PrintOptions: 477, PageSetup: 478, HeaderFooter: 479, WsFmtInfo: 485, HLink: 494,
  SheetProtection: 535, BeginCellXFs: 617, EndCellXFs: 618, BeginCellStyleXFs: 626, EndCellStyleXFs: 627,
  CommentAuthor: 632, BeginComment: 635, CommentText: 637, SupAddin: 667, TableStyleClient: 513,
} as const;

export interface Rec { type: number; data: Uint8Array }

export function* records(u8: Uint8Array): Generator<Rec> {
  let p = 0;
  while (p < u8.length) {
    let type = 0, size = 0;
    for (let i = 0; i < 2; i++) { const b = u8[p++]; type |= (b & 0x7f) << (7 * i); if (!(b & 0x80)) break; }
    for (let i = 0; i < 4; i++) { const b = u8[p++]; size |= (b & 0x7f) << (7 * i); if (!(b & 0x80)) break; }
    if (p > u8.length) return;
    yield { type, data: u8.subarray(p, Math.min(u8.length, p + size)) };
    p += size;
  }
}

/** Little-endian reader over a record payload. */
export class Rd {
  dv: DataView;
  p = 0;
  constructor(public u8: Uint8Array) { this.dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength); }
  get left(): number { return this.u8.length - this.p; }
  u8v(): number { return this.u8[this.p++]; }
  u16(): number { const v = this.dv.getUint16(this.p, true); this.p += 2; return v; }
  i32(): number { const v = this.dv.getInt32(this.p, true); this.p += 4; return v; }
  u32(): number { const v = this.dv.getUint32(this.p, true); this.p += 4; return v; }
  f64(): number { const v = this.dv.getFloat64(this.p, true); this.p += 8; return v; }
  skip(n: number): this { this.p += n; return this; }
  bytes(n: number): Uint8Array { const b = this.u8.subarray(this.p, this.p + n); this.p += n; return b; }
  /** XLWideString; XLNullableWideString returns null for 0xFFFFFFFF. */
  wstr(): string { const n = this.u32(); if (n === 0xffffffff) return ''; return this.chars(n); }
  nwstr(): string | null { const n = this.u32(); if (n === 0xffffffff) return null; return this.chars(n); }
  chars(n: number): string {
    const end = Math.min(this.u8.length, this.p + n * 2);
    let s = '';
    for (let i = this.p; i + 1 < end; i += 2) s += String.fromCharCode(this.u8[i] | (this.u8[i + 1] << 8));
    this.p = end;
    return s;
  }
  rk(): number {
    const v = this.u32();
    let n: number;
    if (v & 2) n = v >> 2;
    else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, v & 0xfffffffc, true); n = b.getFloat64(0, true); }
    return v & 1 ? n / 100 : n;
  }
  rfx(): Range { const r1 = this.u32(), r2 = this.u32(), c1 = this.u32(), c2 = this.u32(); return { r1, r2: Math.min(r2, MAX_ROWS - 1), c1, c2: Math.min(c2, MAX_COLS - 1) }; }
}

const ERRS: Record<number, string> = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA' };
const errVal = (b: number) => new CellError((ERRS[b] || '#N/A') as CellError['code']);

// --------------------------------------------------------------------------------------------
// Colors and styles

interface Theme { colors: string[] }

/** BrtColor (8 bytes) -> #rrggbb, or undefined for automatic. */
function readColor(r: Rd, theme: Theme): string | undefined {
  const b0 = r.u8v(), index = r.u8v();
  const tint = r.dv.getInt16(r.p, true) / 32767; r.p += 2;
  const R = r.u8v(), G = r.u8v(), B = r.u8v(); r.p++;
  const type = b0 >> 1;
  let hex: string | undefined;
  if (type === 2) hex = '#' + [R, G, B].map((x) => x.toString(16).padStart(2, '0')).join('');
  else if (type === 1) hex = index === 64 || index === 0x51 ? undefined : index === 65 ? '#ffffff' : INDEXED_COLORS[index];
  else if (type === 3) hex = theme.colors[[1, 0, 3, 2][index] ?? index];
  if (hex && tint) hex = applyTint(hex, tint);
  return hex;
}

type FontP = Partial<Pick<CellStyle, 'font' | 'size' | 'bold' | 'italic' | 'underline' | 'strike' | 'vertAlign' | 'color'>>;

function readFont(d: Uint8Array, theme: Theme): FontP {
  const r = new Rd(d);
  const f: FontP = {};
  const height = r.u16(), grbit = r.u16(), bls = r.u16(), sss = r.u16(), uls = r.u8v();
  r.skip(3);
  const color = readColor(r, theme);
  r.skip(1);
  const name = r.wstr();
  if (height) f.size = height / 20;
  if (grbit & 0x02) f.italic = true;
  if (grbit & 0x08) f.strike = true;
  if (bls >= 600) f.bold = true;
  if (sss === 1) f.vertAlign = 'superscript'; else if (sss === 2) f.vertAlign = 'subscript';
  const ul: Record<number, CellStyle['underline']> = { 1: 'single', 2: 'double', 0x21: 'singleAccounting', 0x22: 'doubleAccounting' };
  if (ul[uls]) f.underline = ul[uls];
  if (color) f.color = color;
  if (name) f.font = name;
  return f;
}

const PATTERNS = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];

function readFill(d: Uint8Array, theme: Theme): Partial<CellStyle> {
  const r = new Rd(d);
  const fls = r.u32();
  const fg = readColor(r, theme), bg = readColor(r, theme);
  const out: Partial<CellStyle> = {};
  const type = PATTERNS[fls] || 'none';
  if (fls === 0x28) {
    // Gradient fill: stops follow the gradient header.
    const gr = new Rd(d.subarray(20));
    gr.skip(4 + 8 + 8 + 8 + 8 + 8);
    const n = gr.left >= 4 ? gr.u32() : 0;
    const stops: string[] = [];
    for (let i = 0; i < n && gr.left >= 16; i++) { const c = readColor(gr, theme); gr.skip(8); if (c) stops.push(c); }
    if (stops.length) { out.gradient = { from: stops[0], to: stops[stops.length - 1] }; out.fill = stops[0]; }
    return out;
  }
  if (type === 'none') return out;
  if (type === 'solid') { if (fg) out.fill = fg; }
  else { if (bg) out.fill = bg; out.pattern = { type, ...(fg ? { color: fg } : {}) }; }
  return out;
}

const BORDER_STYLES: (Border['style'] | null)[] = [null, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'dashDot', 'dashDot', 'mediumDashed', 'dashDot'];

function readBorder(d: Uint8Array, theme: Theme): Partial<CellStyle> {
  const r = new Rd(d);
  const flags = r.u8v();
  const side = (): Border | undefined => {
    const dg = r.u8v(); r.skip(1);
    const color = readColor(r, theme);
    const st = BORDER_STYLES[dg];
    return st ? { style: st, ...(color ? { color } : {}) } : undefined;
  };
  const top = side(), bottom = side(), left = side(), right = side(), diag = side();
  const out: Partial<CellStyle> = {};
  if (top) out.bt = top; if (bottom) out.bb = bottom; if (left) out.bl = left; if (right) out.br = right;
  if (diag) { if (flags & 1) out.diagDown = diag; if (flags & 2) out.diagUp = diag; }
  return out;
}

interface Xf { fmt: number; font: number; fill: number; border: number; trot: number; indent: number; flags: number }

function readXf(d: Uint8Array): Xf & { parent: number } {
  const r = new Rd(d);
  const parent = r.u16(), fmt = r.u16(), font = r.u16(), fill = r.u16(), border = r.u16();
  const trot = r.u8v(), indent = r.u8v(), flags = r.u16();
  return { parent, fmt, font, fill, border, trot, indent, flags };
}

const H_ALIGN: (CellStyle['hAlign'] | undefined)[] = [undefined, 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const V_ALIGN: (CellStyle['vAlign'] | undefined)[] = ['top', 'middle', undefined, 'justify', 'distributed'];

interface Styles { fonts: FontP[]; fills: Partial<CellStyle>[]; borders: Partial<CellStyle>[]; fmts: Map<number, string>; xfMap: number[] }

function readStyles(u8: Uint8Array | null, book: Workbook, theme: Theme): Styles {
  const st: Styles = { fonts: [], fills: [], borders: [], fmts: new Map(), xfMap: [0] };
  if (!u8) return st;
  const cellXfs: Xf[] = [], styleXfs: Xf[] = [], named: { name: string; xf: number; builtin: boolean }[] = [];
  let mode = 0;
  for (const { type, data } of records(u8)) {
    switch (type) {
      case BRT.Fmt: { const r = new Rd(data); const id = r.u16(); const code = r.wstr(); st.fmts.set(id, code); if (id >= 164 && !book.customFormats.includes(code)) book.customFormats.push(code); break; }
      case BRT.Font: st.fonts.push(readFont(data, theme)); break;
      case BRT.Fill: st.fills.push(readFill(data, theme)); break;
      case BRT.Border: st.borders.push(readBorder(data, theme)); break;
      case BRT.BeginCellXFs: mode = 1; break;
      case BRT.EndCellXFs: mode = 0; break;
      case BRT.BeginCellStyleXFs: mode = 2; break;
      case BRT.EndCellStyleXFs: mode = 0; break;
      case BRT.XF: if (mode === 1) cellXfs.push(readXf(data)); else if (mode === 2) styleXfs.push(readXf(data)); break;
      case BRT.Style: { const r = new Rd(data); const xf = r.u32(); const grbit = r.u16(); r.skip(2); named.push({ name: r.wstr(), xf, builtin: !!(grbit & 1) }); break; }
    }
  }
  const def = st.fonts[0] || {};
  const fmtFor = (id: number) => (id === 0 ? undefined : st.fmts.get(id) ?? BUILTIN_FORMATS[id]);
  const toStyle = (xf: Xf): CellStyle => {
    const s: CellStyle = {};
    const fmt = fmtFor(xf.fmt);
    if (fmt && fmt !== 'General') s.numFmt = fmt;
    const f = st.fonts[xf.font];
    if (f) {
      if (f.font && f.font !== def.font) s.font = f.font;
      if (f.size && f.size !== def.size) s.size = f.size;
      if (f.bold) s.bold = true; if (f.italic) s.italic = true; if (f.strike) s.strike = true;
      if (f.underline) s.underline = f.underline; if (f.vertAlign) s.vertAlign = f.vertAlign;
      if (f.color && f.color !== def.color) s.color = f.color;
    }
    Object.assign(s, xf.fill === 1 && st.fills[1]?.pattern?.type === 'gray125' ? {} : st.fills[xf.fill] || {});
    Object.assign(s, st.borders[xf.border] || {});
    const h = H_ALIGN[xf.flags & 7]; if (h) s.hAlign = h;
    const v = V_ALIGN[(xf.flags >> 3) & 7]; if (v) s.vAlign = v;
    if (xf.flags & 0x40) s.wrap = true;
    if (xf.flags & 0x100) s.shrink = true;
    if (((xf.flags >> 10) & 3) === 2) s.rtl = true;
    if (!(xf.flags & 0x1000)) s.locked = false;
    if (xf.flags & 0x2000) s.hidden = true;
    if (xf.indent) s.indent = xf.indent;
    if (xf.trot) s.rotate = xf.trot === 255 ? 255 : xf.trot > 90 ? -(xf.trot - 90) : xf.trot;
    return s;
  };
  st.xfMap = cellXfs.map((xf) => book.styleId(toStyle(xf)));
  if (!st.xfMap.length) st.xfMap = [0];
  for (const n of named) if (styleXfs[n.xf]) book.cellStyles.push({ name: n.name, style: toStyle(styleXfs[n.xf]), ...(n.builtin ? { builtin: true } : {}) });
  return st;
}

/** RichStr: plain or with formatting runs (each run a font index). */
function readRichStr(r: Rd, fonts: FontP[]): { t: string; rich?: Cell['rich'] } {
  const flags = r.u8v();
  const t = r.wstr();
  if (!(flags & 1) || r.left < 4) return { t };
  const n = r.u32();
  const runs: { ich: number; ifnt: number }[] = [];
  for (let i = 0; i < n && r.left >= 4; i++) runs.push({ ich: r.u16(), ifnt: r.u16() });
  if (!runs.length) return { t };
  const rich: NonNullable<Cell['rich']> = [];
  if (runs[0].ich > 0) rich.push({ t: t.slice(0, runs[0].ich) });
  runs.forEach((run, i) => {
    const seg = t.slice(run.ich, i + 1 < runs.length ? runs[i + 1].ich : undefined);
    const f = fonts[run.ifnt];
    rich.push(f && Object.keys(f).length ? { t: seg, s: { ...f } } : { t: seg });
  });
  return { t, rich };
}

// --------------------------------------------------------------------------------------------
// Formulas

interface FormulaCtx {
  sheets: string[];
  names: { name: string; sheet: number }[];
  xti: { sup: number; first: number; last: number }[];
  sups: { kind: 'self' | 'same' | 'addin' | 'ext'; path: string; sheets: string[]; names: string[] }[];
  tables: Map<number, { name: string; cols: string[] }>;
}

/** Functions added in Excel 2007, numbered after the BIFF8 table (index -> name, fixed argument count or -1). */
const FN_2007 = ('380 CUBEVALUE|381 CUBEMEMBER|382 CUBEMEMBERPROPERTY|383 CUBERANKEDMEMBER|384 HEX2BIN|385 HEX2DEC 1|386 HEX2OCT|387 DEC2BIN|'
  + '388 DEC2HEX|389 DEC2OCT|390 OCT2BIN|391 OCT2HEX|392 OCT2DEC 1|393 BIN2DEC 1|394 BIN2OCT|395 BIN2HEX|396 IMSUB 2|397 IMDIV 2|398 IMPOWER 2|'
  + '399 IMABS 1|400 IMSQRT 1|401 IMLN 1|402 IMLOG2 1|403 IMLOG10 1|404 IMSIN 1|405 IMCOS 1|406 IMEXP 1|407 IMARGUMENT 1|408 IMCONJUGATE 1|'
  + '409 IMAGINARY 1|410 IMREAL 1|411 COMPLEX|412 IMSUM|413 IMPRODUCT|414 SERIESSUM 4|415 FACTDOUBLE 1|416 SQRTPI 1|417 QUOTIENT 2|418 DELTA|'
  + '419 GESTEP|420 ISEVEN 1|421 ISODD 1|422 MROUND 2|423 ERF|424 ERFC 1|425 BESSELJ 2|426 BESSELK 2|427 BESSELY 2|428 BESSELI 2|429 XIRR|'
  + '430 XNPV 3|431 PRICEMAT|432 YIELDMAT|433 INTRATE|434 RECEIVED|435 DISC|436 PRICEDISC|437 YIELDDISC|438 TBILLEQ 3|439 TBILLPRICE 3|'
  + '440 TBILLYIELD 3|441 PRICE|442 YIELD|443 DOLLARDE 2|444 DOLLARFR 2|445 NOMINAL 2|446 EFFECT 2|447 CUMPRINC 6|448 CUMIPMT 6|449 EDATE 2|'
  + '450 EOMONTH 2|451 YEARFRAC|452 COUPDAYBS|453 COUPDAYS|454 COUPDAYSNC|455 COUPNCD|456 COUPNUM|457 COUPPCD|458 DURATION|459 MDURATION|'
  + '460 ODDLPRICE|461 ODDLYIELD|462 ODDFPRICE|463 ODDFYIELD|464 RANDBETWEEN 2|465 WEEKNUM|466 AMORDEGRC|467 AMORLINC|468 CONVERT 3|'
  + '469 ACCRINT|470 ACCRINTM|471 WORKDAY|472 NETWORKDAYS|473 GCD|474 MULTINOMIAL|475 LCM|476 FVSCHEDULE 2|477 CUBEKPIMEMBER|478 CUBESET|'
  + '479 CUBESETCOUNT 1|480 IFERROR 2|481 COUNTIFS|482 SUMIFS|483 AVERAGEIF|484 AVERAGEIFS').split('|');
const FUNCS = new Map<number, [string, number]>(BIFF_FUNCTIONS);
for (const e of FN_2007) { const [i, n, a] = e.split(' '); FUNCS.set(+i, [n, a ? +a : -1]); }

const OPS: Record<number, string> = { 3: '+', 4: '-', 5: '*', 6: '/', 7: '^', 8: '&', 9: '<', 10: '<=', 11: '=', 12: '>=', 13: '>', 14: '<>', 15: ' ', 16: ',', 17: ':' };

function quoteSheetName(name: string): string {
  return /^[A-Za-z_À-￿][\w.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(R\d*C\d*|TRUE|FALSE)$/i.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function numText(n: number): string {
  if (!Number.isFinite(n)) return '#NUM!';
  const s = String(n);
  return s.includes('e') ? n.toExponential().toUpperCase() : s;
}

/**
 * Decodes a BIFF12 parsed formula (rgce + extra data rgcb) into formula text without '='.
 * (r, c) is the cell for relative references (shared formulas, RefN/AreaN); throws on tokens it
 * does not know, so the caller can keep the cached value.
 */
export function decodeFormula(rgce: Uint8Array, rgcb: Uint8Array, r: number, c: number, ctx: FormulaCtx): string {
  const rd = new Rd(rgce);
  const xr = new Rd(rgcb);
  const st: string[] = [];
  const pop = () => { const x = st.pop(); if (x === undefined) throw new Error('formula stack underflow'); return x; };
  const loc = (row: number, colw: number, relative: boolean): string => {
    const rowRel = !!(colw & 0x8000), colRel = !!(colw & 0x4000);
    let col = colw & 0x3fff;
    if (relative) {
      if (rowRel) row = r + row;
      if (colRel) col = c + ((col << 18) >> 18);
    }
    if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) return '#REF!';
    return `${colRel ? '' : '$'}${colName(col)}${rowRel ? '' : '$'}${row + 1}`;
  };
  const ref = (relative: boolean) => { const row = relative ? rd.i32() : rd.u32(); const col = rd.u16(); return loc(row, col, relative); };
  const area = (relative: boolean) => {
    const r1 = relative ? rd.i32() : rd.u32(), r2 = relative ? rd.i32() : rd.u32(), c1 = rd.u16(), c2 = rd.u16();
    if (!relative && r1 === 0 && r2 >= MAX_ROWS - 1) return `${c1 & 0x4000 ? '' : '$'}${colName(c1 & 0x3fff)}:${c2 & 0x4000 ? '' : '$'}${colName(c2 & 0x3fff)}`;
    if (!relative && (c1 & 0x3fff) === 0 && (c2 & 0x3fff) >= MAX_COLS - 1) return `${c1 & 0x8000 ? '' : '$'}${r1 + 1}:${c2 & 0x8000 ? '' : '$'}${r2 + 1}`;
    return `${loc(r1, c1, relative)}:${loc(r2, c2, relative)}`;
  };
  const sheetPrefix = (ixti: number): string | null => {
    const x = ctx.xti[ixti];
    if (!x) return null;
    const sup = ctx.sups[x.sup];
    if (!sup || sup.kind === 'self' || sup.kind === 'same') {
      if (x.first < 0 || x.first >= ctx.sheets.length) return null;
      const a = ctx.sheets[x.first], b = ctx.sheets[x.last] ?? a;
      return x.first === x.last || x.last < 0 ? quoteSheetName(a) : quoteSheetName(`${a}:${b}`);
    }
    const file = sup.path.replace(/^.*[\\/]/, '');
    const sh = sup.sheets[x.first] ?? '';
    return `'[${file.replace(/'/g, "''")}]${sh.replace(/'/g, "''")}'`;
  };
  const arrayConst = (): string => {
    const rows = xr.u32(), cols = xr.u32();
    if (rows * cols > 100000) throw new Error('array constant too large');
    const out: string[] = [];
    for (let i = 0; i < rows; i++) {
      const row: string[] = [];
      for (let j = 0; j < cols; j++) {
        const t = xr.u8v();
        if (t === 0x00) row.push(numText(xr.f64()));
        else if (t === 0x01) row.push(`"${xr.wstr().replace(/"/g, '""')}"`);
        else if (t === 0x02) { row.push(xr.u8v() ? 'TRUE' : 'FALSE'); }
        else if (t === 0x04) { row.push(ERRS[xr.u8v()] || '#N/A'); xr.skip(3); }
        else throw new Error('bad array constant');
      }
      out.push(row.join(','));
    }
    return `{${out.join(';')}}`;
  };
  while (rd.left > 0) {
    const ptg = rd.u8v();
    if (ptg >= 0x03 && ptg <= 0x11) {
      const b = pop(), a = pop();
      st.push(ptg === 0x10 ? `${a},${b}` : `${a}${OPS[ptg]}${b}`);
      continue;
    }
    switch (ptg) {
      case 0x01: case 0x02: throw new Error('shared formula token');
      case 0x12: st.push(`+${pop()}`); continue;
      case 0x13: st.push(`-${pop()}`); continue;
      case 0x14: st.push(`${pop()}%`); continue;
      case 0x15: st.push(`(${pop()})`); continue;
      case 0x16: st.push(''); continue;
      case 0x17: { const n = rd.u16(); st.push(`"${rd.chars(n).replace(/"/g, '""')}"`); continue; }
      case 0x18: {
        const eptg = rd.u8v();
        if (eptg !== 0x19) throw new Error('extended token');
        // PtgList: structured reference to a table.
        rd.skip(2);
        const flags = rd.u16();
        const list = rd.u32(), cf = rd.u16(), cl = rd.u16();
        const t = ctx.tables.get(list);
        if (!t) throw new Error('unknown table');
        const colMode = flags & 3, rowType = (flags >> 2) & 0x1f;
        const items: string[] = [];
        const rowItems: Record<number, string[]> = { 0: [], 1: ['#All'], 2: ['#Headers'], 4: ['#Data'], 6: ['#Headers', '#Data'], 8: ['#Totals'], 12: ['#Data', '#Totals'], 16: ['#This Row'] };
        const esc = (s: string) => s.replace(/([[\]#'])/g, "'$1");
        const thisRow = rowType === 16;
        for (const it of rowItems[rowType] || []) if (!thisRow) items.push(`[${it}]`);
        if (colMode === 1) items.push(`[${esc(t.cols[cf] ?? '')}]`);
        else if (colMode === 2) items.push(`[${esc(t.cols[cf] ?? '')}]:[${esc(t.cols[cl] ?? '')}]`);
        if (thisRow) st.push(`${t.name}[@${colMode ? (colMode === 1 ? esc(t.cols[cf] ?? '') : `[${esc(t.cols[cf] ?? '')}]:[${esc(t.cols[cl] ?? '')}]`) : ''}]`);
        else st.push(items.length === 1 && colMode === 1 && !(rowItems[rowType] || []).length ? `${t.name}${items[0]}` : items.length ? `${t.name}[${items.join(',')}]` : `${t.name}[]`);
        continue;
      }
      case 0x19: {
        const gr = rd.u8v(), w = rd.u16();
        if (gr & 0x04) { rd.skip(2 * (w + 1)); continue; }
        if (gr & 0x10) st.push(`SUM(${pop()})`);
        continue;
      }
      case 0x1c: st.push(ERRS[rd.u8v()] || '#N/A'); continue;
      case 0x1d: st.push(rd.u8v() ? 'TRUE' : 'FALSE'); continue;
      case 0x1e: st.push(String(rd.u16())); continue;
      case 0x1f: st.push(numText(rd.f64())); continue;
    }
    if (ptg < 0x20 || ptg > 0x7f) throw new Error(`unknown token 0x${ptg.toString(16)}`);
    switch ((ptg & 0x1f) | 0x20) {
      case 0x20: rd.skip(14); st.push(arrayConst()); break;
      case 0x21: {
        const idx = rd.u16();
        const fn = FUNCS.get(idx);
        if (!fn || fn[1] < 0) throw new Error(`unknown function ${idx}`);
        const args: string[] = [];
        for (let k = 0; k < fn[1]; k++) args.unshift(pop());
        st.push(`${fn[0]}(${args.join(',')})`);
        break;
      }
      case 0x22: {
        const n = rd.u8v() & 0x7f, idx = rd.u16() & 0x7fff;
        const args: string[] = [];
        for (let k = 0; k < n; k++) args.unshift(pop());
        if (idx === 255) st.push(`${(args.shift() || '').replace(/^_xlfn\.(_xlws\.)?|^_xlws\./i, '')}(${args.join(',')})`);
        else { const fn = FUNCS.get(idx); if (!fn) throw new Error(`unknown function ${idx}`); st.push(`${fn[0]}(${args.join(',')})`); }
        break;
      }
      case 0x23: { const i = rd.u32(); const nm = ctx.names[i - 1]; if (!nm) throw new Error('unknown name'); st.push(nm.name); break; }
      case 0x24: st.push(ref(false)); break;
      case 0x25: st.push(area(false)); break;
      case 0x26: { rd.skip(6); const n = xr.u32(); xr.skip(n * 16); break; }
      case 0x27: case 0x28: rd.skip(6); break;
      case 0x29: rd.skip(2); break;
      case 0x2a: rd.skip(6); st.push('#REF!'); break;
      case 0x2b: rd.skip(12); st.push('#REF!'); break;
      case 0x2c: st.push(ref(true)); break;
      case 0x2d: st.push(area(true)); break;
      case 0x39: {
        const ixti = rd.u16(), i = rd.u32();
        const x = ctx.xti[ixti], sup = x ? ctx.sups[x.sup] : undefined;
        if (!sup || sup.kind === 'self' || sup.kind === 'same') { const nm = ctx.names[i - 1]; if (!nm) throw new Error('unknown name'); st.push(nm.name); }
        else { const nm = sup.names[i - 1]; if (nm == null) throw new Error('unknown external name'); st.push(sup.kind === 'addin' ? nm.replace(/^_xlfn\./i, '') : `[${sup.path.replace(/^.*[\\/]/, '')}]${nm}`); }
        break;
      }
      case 0x3a: { const pre = sheetPrefix(rd.u16()); const s = ref(false); st.push(pre ? `${pre}!${s}` : '#REF!'); break; }
      case 0x3b: { const pre = sheetPrefix(rd.u16()); const s = area(false); st.push(pre ? `${pre}!${s}` : '#REF!'); break; }
      case 0x3c: { const pre = sheetPrefix(rd.u16()); rd.skip(6); st.push(pre ? `${pre}!#REF!` : '#REF!'); break; }
      case 0x3d: { const pre = sheetPrefix(rd.u16()); rd.skip(12); st.push(pre ? `${pre}!#REF!` : '#REF!'); break; }
      default: throw new Error(`unknown token 0x${ptg.toString(16)}`);
    }
  }
  if (st.length !== 1) throw new Error('unbalanced formula');
  return st[0].replace(/^_xlfn\./i, '');
}

/** CellParsedFormula / NameParsedFormula: cce, rgce, cb, rgcb. */
function readParsed(r: Rd): { rgce: Uint8Array; rgcb: Uint8Array } {
  const cce = r.u32();
  const rgce = r.bytes(Math.min(cce, r.left));
  const cb = r.left >= 4 ? r.u32() : 0;
  const rgcb = r.bytes(Math.min(cb, r.left));
  return { rgce, rgcb };
}

// --------------------------------------------------------------------------------------------
// Workbook

interface SheetInfo { name: string; rel: Rel | undefined; state: number }

export async function importXlsb(data: ArrayBuffer, warnings: string[] = []): Promise<Workbook> {
  const zip = await JSZip.loadAsync(data);
  const pkg = new Pkg(zip);
  const book = new Workbook(false);
  const rootRels = await pkg.rels('');
  let wbPath = rootRels.find((r) => r.type.endsWith('/officeDocument'))?.target || 'xl/workbook.bin';
  if (!pkg.has(wbPath)) wbPath = [...pkg.files.keys()].find((k) => /(^|\/)workbook\.bin$/.test(k)) || wbPath;
  const wb = await pkg.bytes(wbPath);
  if (!wb) throw new Error('The workbook part (workbook.bin) is missing.');
  const wbRels = await pkg.rels(wbPath);
  const relOf = (t: string) => wbRels.find((r) => r.type.endsWith('/' + t));
  const warn = (m: string) => { if (!warnings.includes(m)) warnings.push(m); };

  // Theme (XML) first: colors in styles refer to it.
  const theme: Theme = { colors: [...book.theme.colors] };
  const themeRel = relOf('theme');
  if (themeRel) { try { theme.colors = parseTheme(await pkg.xml(themeRel.target), book); } catch { warn('The theme part is damaged and was skipped.'); } }

  // Workbook records.
  const sheets: SheetInfo[] = [];
  const ctx: FormulaCtx = { sheets: [], names: [], xti: [], sups: [], tables: new Map() };
  const rawNames: { name: string; sheet: number; flags: number; rgce: Uint8Array; rgcb: Uint8Array; comment: string | null }[] = [];
  let activeTab = 0;
  for (const { type, data: d } of records(wb)) {
    const r = new Rd(d);
    switch (type) {
      case BRT.BundleSh: {
        // hsState, iTabID, then the relationship id and the name. Some writers put 4 more bytes
        // before the strings: take the layout whose strings end exactly at the record end.
        const state = r.u32();
        for (const pre of [8, 12, 4]) {
          const t = new Rd(d); t.p = pre;
          if (t.left < 8) continue;
          const n1 = t.u32(); if (n1 !== 0xffffffff && n1 * 2 > t.left) continue; t.p -= 4;
          const relId = t.nwstr();
          if (t.left < 4) continue;
          const n2 = t.u32(); if (n2 * 2 !== t.left) continue; t.p -= 4;
          sheets.push({ name: t.wstr(), rel: wbRels.find((x) => x.id === relId), state });
          break;
        }
        break;
      }
      case BRT.WbProp: { const f = r.u32(); if (f & 1) book.calc.date1904 = true; break; }
      case BRT.BookView: { if (d.length >= 28) { r.skip(24); activeTab = r.u32(); } break; }
      case BRT.CalcProp: {
        if (d.length < 21) break;
        r.u32(); // recalc engine id
        const mode = r.i32(), iter = r.u32(), delta = r.f64();
        r.u32(); // user thread count
        const flags = r.u8v();
        book.calc.mode = mode === 0 ? 'manual' : mode === 2 ? 'autoNoTables' : 'auto';
        book.calc.maxIterations = iter || 100; book.calc.maxChange = delta || 0.001;
        if (flags & 0x04) book.calc.iterative = true;
        if (!(flags & 0x08)) book.calc.precisionAsDisplayed = true;
        break;
      }
      case BRT.SupSelf: ctx.sups.push({ kind: 'self', path: '', sheets: [], names: [] }); break;
      case BRT.SupSame: ctx.sups.push({ kind: 'same', path: '', sheets: [], names: [] }); break;
      case BRT.SupAddin: ctx.sups.push({ kind: 'addin', path: '', sheets: [], names: [] }); break;
      case BRT.SupBookSrc: {
        const relId = r.wstr();
        const rel = wbRels.find((x) => x.id === relId);
        ctx.sups.push({ kind: 'ext', path: rel?.target || 'external', sheets: [], names: [] });
        break;
      }
      case BRT.ExternSheet: {
        const n = r.u32();
        for (let i = 0; i < n && r.left >= 12; i++) ctx.xti.push({ sup: r.u32(), first: r.i32(), last: r.i32() });
        break;
      }
      case BRT.Name: {
        const flags = r.u32(); r.skip(1);
        const itab = r.u32();
        const name = r.wstr();
        const { rgce, rgcb } = readParsed(r);
        const comment = r.left >= 4 ? r.nwstr() : null;
        rawNames.push({ name, sheet: itab === 0xffffffff ? -1 : itab, flags, rgce, rgcb, comment });
        ctx.names.push({ name: name.replace(/^_xlnm\./, ''), sheet: itab });
        break;
      }
    }
  }
  ctx.sheets = sheets.map((s) => s.name);
  if (activeTab < sheets.length) book.activeSheet = activeTab;

  const styles = readStyles(await pkg.bytes(relOf('styles')?.target || 'xl/styles.bin'), book, theme);
  const sstBytes = await pkg.bytes(relOf('sharedStrings')?.target || 'xl/sharedStrings.bin');
  const sst: { t: string; rich?: Cell['rich'] }[] = [];
  if (sstBytes) for (const { type, data: d } of records(sstBytes)) if (type === BRT.SSTItem) sst.push(readRichStr(new Rd(d), styles.fonts));
  const persons = new Map<string, string>();
  const personRel = wbRels.find((r) => /\/person$/.test(r.type));
  if (personRel) { try { for (const p of desc(await pkg.xml(personRel.target), 'person')) persons.set(attr(p, 'id') || '', attr(p, 'displayName') || ''); } catch { /* optional */ } }

  // Tables first (formulas refer to them by id).
  const sheetRels = new Map<number, Rel[]>();
  for (let i = 0; i < sheets.length; i++) {
    const rel = sheets[i].rel;
    if (!rel) continue;
    const rels = await pkg.rels(rel.target);
    sheetRels.set(i, rels);
  }
  const tableDefs = new Map<number, TableDef[]>();
  for (const [i, rels] of sheetRels) {
    for (const rel of rels.filter((x) => x.type.endsWith('/table'))) {
      const tb = await pkg.bytes(rel.target);
      if (!tb) continue;
      const t = readTablePart(tb);
      if (!t) continue;
      ctx.tables.set(t.id, { name: t.def.name, cols: t.def.columns.map((c) => c.name) });
      if (!tableDefs.has(i)) tableDefs.set(i, []);
      tableDefs.get(i)!.push(t.def);
    }
  }

  const used = new Set<string>();
  for (let i = 0; i < sheets.length; i++) {
    const info = sheets[i];
    let name = info.name || `Sheet${i + 1}`;
    while (used.has(name.toLowerCase())) name += '_';
    used.add(name.toLowerCase());
    const sheet = new Sheet(book, name);
    if (info.state === 1) sheet.visibility = 'hidden'; else if (info.state === 2) sheet.visibility = 'veryHidden';
    book.sheets.push(sheet);
    if (!info.rel || !pkg.has(info.rel.target)) { warn(`Sheet "${name}" has no data part; it was opened empty.`); continue; }
    if (!/\.bin$/i.test(info.rel.target)) { warn(`Sheet "${name}" is a chart or dialog sheet and was opened empty.`); continue; }
    try {
      const bytes = (await pkg.bytes(info.rel.target))!;
      readSheet(bytes, sheet, i, ctx, styles, sst, warnings);
      sheet.tables.push(...(tableDefs.get(i) || []));
      const rels = sheetRels.get(i) || [];
      for (const rel of rels) {
        try {
          if (rel.type.endsWith('/comments')) { const cb = await pkg.bytes(rel.target); if (cb) readComments(cb, sheet, styles.fonts); }
          else if (rel.type.endsWith('/drawing')) await readSheetDrawing(pkg, sheet, rel.target, theme.colors, warnings);
          else if (/\/threadedComment$/.test(rel.type)) await readThreadedComments(pkg, sheet, rel.target, persons, warnings);
        } catch { warn(`Some objects on sheet "${name}" could not be read.`); }
      }
      // Hyperlink targets live in the sheet relationships.
      for (const [k, cell] of sheet.cells) {
        if (cell.link && cell.link.startsWith('rel:')) {
          const [, id, loc] = /^rel:([^#]*)(#.*)?$/.exec(cell.link) || [];
          const rel = rels.find((x) => x.id === id);
          const target = (rel ? rel.target : '') + (loc || '');
          if (target) cell.link = target; else delete cell.link;
          sheet.cells.set(k, cell);
        }
      }
    } catch (e) {
      warn(`Sheet "${name}" could not be read completely: ${e instanceof Error ? e.message : e}`);
    }
  }
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));

  // Defined names (after the sheets: print areas belong to them).
  for (const n of rawNames) {
    const sheet = n.sheet >= 0 ? book.sheets[n.sheet] : undefined;
    if (!n.rgce.length) continue;
    let ref: string;
    try { ref = decodeFormula(n.rgce, n.rgcb, 0, 0, ctx); } catch { warn('Some defined names could not be read.'); continue; }
    const bare = n.name.replace(/^_xlnm\./, '');
    if (/^(Print_Area|Print_Titles|_FilterDatabase)$/i.test(bare) || (n.flags & 0x20)) {
      if (!sheet) continue;
      const parts = ref.split(',').map((p) => p.replace(/^.*!/, ''));
      if (/^Print_Area$/i.test(bare)) sheet.print.area = parts.join(',');
      else if (/^Print_Titles$/i.test(bare)) for (const p of parts) {
        if (/^\$?\d+:\$?\d+$/.test(p)) sheet.print.titleRows = p.replace(/\$/g, '');
        else if (/^\$?[A-Z]+:\$?[A-Z]+$/i.test(p)) sheet.print.titleCols = p.replace(/\$/g, '');
      }
      continue;
    }
    const d: DefinedName = { name: n.name, ref, sheet: sheet ? sheet.name : null };
    if (n.comment) d.comment = n.comment;
    if (n.flags & 1) d.hidden = true;
    book.names.push(d);
  }
  if (wbRels.some((r) => /vbaProject$/.test(r.type))) {
    const vba = wbRels.find((r) => /vbaProject$/.test(r.type))!;
    const b = await pkg.bytes(vba.target);
    if (b) { (book as any).vbaProject = b; warn('This workbook contains VBA macros. They are kept when saved as .xlsm but are not run.'); }
  }
  if (wbRels.some((r) => /pivotCacheDefinition$/.test(r.type))) warn('PivotTables were opened as static values.');
  await parseDocProps(pkg, rootRels, book);
  book.clearHistory();
  return book;
}

// --------------------------------------------------------------------------------------------
// Worksheet

function readSheet(u8: Uint8Array, sheet: Sheet, index: number, ctx: FormulaCtx, styles: Styles, sst: { t: string; rich?: Cell['rich'] }[], warnings: string[]): void {
  const xf = (i: number) => styles.xfMap[i] || 0;
  let row = 0;
  const shared: { rg: Range; rgce: Uint8Array; rgcb: Uint8Array }[] = [];
  const pendingExp: { r: number; c: number }[] = [];
  const arrays: Range[] = [];
  let bad = 0, fitToPage = false;
  const book = sheet.book;
  for (const { type, data: d } of records(u8)) {
    const r = new Rd(d);
    switch (type) {
      case BRT.RowHdr: {
        row = r.u32();
        if (row >= MAX_ROWS) row = MAX_ROWS - 1;
        const ixfe = r.u32(), height = r.u16();
        r.skip(1);
        const f = r.u8v();
        const lvl = f & 7;
        if (f & 0x20) sheet.hiddenRows.add(row);
        if (lvl) sheet.rowLevel.set(row, lvl);
        if (height && (f & 0x40)) sheet.rowH.set(row, ptToPx(height / 20));
        const cf = d.length >= 13 ? d[12] : 0;
        if (ixfe && xf(ixfe) && (cf & 0x01 || ixfe)) { const s = xf(ixfe); if (s) sheet.rowStyle.set(row, s); }
        break;
      }
      case BRT.CellBlank: case BRT.CellRk: case BRT.CellError: case BRT.CellBool: case BRT.CellReal: case BRT.CellSt: case BRT.CellIsst:
      case BRT.FmlaString: case BRT.FmlaNum: case BRT.FmlaBool: case BRT.FmlaError: case BRT.CellRString: {
        const col = r.u32();
        const sref = r.u32() & 0xffffff;
        if (col >= MAX_COLS) break;
        const cell: Cell = {};
        const s = xf(sref);
        if (s) cell.s = s;
        switch (type) {
          case BRT.CellRk: cell.v = r.rk(); break;
          case BRT.CellError: cell.v = errVal(r.u8v()); break;
          case BRT.CellBool: cell.v = !!r.u8v(); break;
          case BRT.CellReal: cell.v = r.f64(); break;
          case BRT.CellSt: cell.v = r.wstr(); break;
          case BRT.CellIsst: { const ss = sst[r.u32()]; cell.v = ss ? ss.t : ''; if (ss?.rich) cell.rich = ss.rich; break; }
          case BRT.CellRString: { const rs = readRichStr(r, styles.fonts); cell.v = rs.t; if (rs.rich) cell.rich = rs.rich; break; }
          case BRT.FmlaString: cell.v = r.wstr(); break;
          case BRT.FmlaNum: cell.v = r.f64(); break;
          case BRT.FmlaBool: cell.v = !!r.u8v(); break;
          case BRT.FmlaError: cell.v = errVal(r.u8v()); break;
        }
        if (type >= BRT.FmlaString && type <= BRT.FmlaError) {
          r.u16(); // grbitFlags
          const { rgce, rgcb } = readParsed(r);
          if (rgce[0] === 0x01 || rgce[0] === 0x02) {
            if (rgce[0] === 0x01) pendingExp.push({ r: row, c: col });
          } else {
            try { cell.f = decodeFormula(rgce, rgcb, row, col, ctx); } catch { bad++; }
          }
        }
        if (cell.v === '' && cell.f == null) delete cell.v;
        const prev = sheet.cells.get(cellKey(row, col));
        if (prev?.sp != null) cell.sp = prev.sp; // member of an array formula read earlier
        if (cell.v != null || cell.f != null || cell.s || cell.sp != null) sheet.cells.set(cellKey(row, col), cell);
        break;
      }
      case 0x1cd: case 0x23d:
        if (!warnings.includes('Conditional formats and data validation rules in .xlsb files are not opened.')) warnings.push('Conditional formats and data validation rules in .xlsb files are not opened.');
        break;
      case BRT.ShrFmla: {
        const rg = r.rfx();
        const { rgce, rgcb } = readParsed(r);
        shared.push({ rg, rgce, rgcb });
        break;
      }
      case BRT.ArrFmla: {
        const rg = r.rfx();
        r.skip(1);
        const { rgce, rgcb } = readParsed(r);
        arrays.push(rg);
        const k = cellKey(rg.r1, rg.c1);
        const anchor = sheet.cells.get(k) || {};
        try {
          anchor.f = decodeFormula(rgce, rgcb, rg.r1, rg.c1, ctx);
          anchor.arr = rg;
          sheet.cells.set(k, anchor);
          for (let rr = rg.r1; rr <= rg.r2; rr++) for (let cc = rg.c1; cc <= rg.c2; cc++) {
            if (rr === rg.r1 && cc === rg.c1) continue;
            const k2 = cellKey(rr, cc);
            const x = sheet.cells.get(k2) || {};
            x.sp = k;
            sheet.cells.set(k2, x);
          }
        } catch { bad++; }
        break;
      }
      case BRT.ColInfo: {
        const c1 = r.u32(), c2 = Math.min(r.u32(), MAX_COLS - 1);
        const w = r.u32(), ixfe = r.u32(), f = r.u16();
        const hidden = !!(f & 1), lvl = (f >> 8) & 7;
        const px = widthToPx(w / 256);
        const s = xf(ixfe);
        for (let c = c1; c <= c2 && c < c1 + 16384; c++) {
          if (w && !(hidden && !w) && (f & 2 || px !== sheet.defaultColW)) sheet.colW.set(c, px);
          if (hidden) sheet.hiddenCols.add(c);
          if (lvl) sheet.colLevel.set(c, lvl);
          if (s) sheet.colStyle.set(c, s);
        }
        break;
      }
      case BRT.WsFmtInfo: {
        r.u32();
        const cch = r.u16(), h = r.u16();
        if (h) sheet.defaultRowH = ptToPx(h / 20);
        if (cch && cch !== 8) sheet.defaultColW = widthToPx(cch) + 5;
        break;
      }
      case BRT.MergeCell: { const rg = r.rfx(); if (rg.r2 > rg.r1 || rg.c2 > rg.c1) sheet.merges.push(rg); break; }
      case BRT.WsProp: {
        fitToPage = !!(r.u16() & 0x100);
        r.skip(1);
        const color = readColor(r, { colors: book.theme.colors });
        if (color) sheet.tabColor = color;
        r.skip(8);
        const code = r.left >= 4 ? r.wstr() : '';
        if (code) sheet.codeName = code;
        break;
      }
      case BRT.WsView: {
        const f = r.u16(); r.u32();
        const top = r.u32(), left = r.u32();
        r.skip(2);
        const scale = r.u16();
        if (f & 0x02) sheet.showFormulas = true;
        if (!(f & 0x04)) sheet.showGrid = false;
        if (!(f & 0x08)) sheet.showHeadings = false;
        if (!(f & 0x10)) sheet.showZeros = false;
        if (f & 0x20) sheet.rtl = true;
        if (f & 0x40) book.activeSheet = index;
        if (scale >= 10 && scale <= 400) sheet.zoom = scale / 100;
        if (top || left) sheet.scroll = { r: Math.min(top, MAX_ROWS - 1), c: Math.min(left, MAX_COLS - 1) };
        break;
      }
      case BRT.Pane: {
        const xs = r.f64(), ys = r.f64();
        r.skip(12);
        const f = r.u8v();
        if (f & 1) sheet.freeze = { rows: Math.round(ys), cols: Math.round(xs) };
        else if (xs || ys) sheet.split = { x: Math.round((xs / 20) * 4 / 3), y: Math.round((ys / 20) * 4 / 3) };
        break;
      }
      case BRT.HLink: {
        const rg = r.rfx();
        const relId = r.nwstr();
        const loc = r.wstr(), tip = r.wstr();
        const link = relId ? `rel:${relId}${loc ? '#' + loc : ''}` : loc ? '#' + loc : '';
        if (!link) break;
        for (let rr = rg.r1; rr <= Math.min(rg.r2, rg.r1 + 1000); rr++) for (let cc = rg.c1; cc <= Math.min(rg.c2, rg.c1 + 100); cc++) {
          const k = cellKey(rr, cc);
          const cell = { ...(sheet.cells.get(k) || {}), link };
          if (tip) cell.linkTip = tip;
          sheet.cells.set(k, cell);
        }
        break;
      }
      case BRT.SheetProtection: {
        r.u16();
        if (r.u32() === 1) sheet.protection = { on: true, allow: ['selectLocked', 'selectUnlocked'] };
        break;
      }
      case BRT.Margins: {
        if (d.length < 48) break;
        const left = r.f64(), right = r.f64(), top = r.f64(), bottom = r.f64(), header = r.f64(), footer = r.f64();
        sheet.print.margins = { left, right, top, bottom, header, footer };
        break;
      }
      case BRT.PrintOptions: {
        const f = r.u16();
        sheet.print.centerH = !!(f & 1); sheet.print.centerV = !!(f & 2); sheet.print.headings = !!(f & 4); sheet.print.gridlines = !!(f & 8);
        break;
      }
      case BRT.PageSetup: {
        if (d.length < 34) break;
        const paper = r.u32(), scale = r.u32(); r.skip(12);
        const first = r.i32(), fw = r.u32(), fh = r.u32(), f = r.u16();
        if (PAPER[paper]) sheet.print.paper = PAPER[paper];
        if (scale >= 10 && scale <= 400) sheet.print.scale = scale;
        if (f & 0x02) sheet.print.orientation = 'landscape';
        if (f & 0x01) sheet.print.order = 'overThenDown';
        if (f & 0x10) sheet.print.draft = true;
        if (f & 0x08) sheet.print.blackAndWhite = true;
        if (f & 0x80 && first) sheet.print.firstPage = first;
        if (fitToPage) { sheet.print.fitWide = fw; sheet.print.fitTall = fh; }
        break;
      }
      case BRT.HeaderFooter: {
        const f = r.u16();
        const h = r.nwstr(), ft = r.nwstr(), he = r.nwstr(), fe = r.nwstr(), hf = r.nwstr(), ff = r.nwstr();
        if (h) sheet.print.header = splitHeaderFooter(h);
        if (ft) sheet.print.footer = splitHeaderFooter(ft);
        if (f & 1) { sheet.print.differentOddEven = true; if (he) sheet.print.evenHeader = splitHeaderFooter(he); if (fe) sheet.print.evenFooter = splitHeaderFooter(fe); }
        if (f & 2) { sheet.print.differentFirst = true; if (hf) sheet.print.firstHeader = splitHeaderFooter(hf); if (ff) sheet.print.firstFooter = splitHeaderFooter(ff); }
        break;
      }
    }
  }
  // Cells that use a shared formula.
  for (const p of pendingExp) {
    const have = sheet.cells.get(cellKey(p.r, p.c));
    if (have?.f != null || have?.sp != null || arrays.some((a) => p.r >= a.r1 && p.r <= a.r2 && p.c >= a.c1 && p.c <= a.c2)) continue; // array formula cells
    const sf = shared.find((s) => p.r >= s.rg.r1 && p.r <= s.rg.r2 && p.c >= s.rg.c1 && p.c <= s.rg.c2);
    const k = cellKey(p.r, p.c);
    if (!sf) { bad++; continue; }
    try {
      const cell = sheet.cells.get(k) || {};
      cell.f = decodeFormula(sf.rgce, sf.rgcb, p.r, p.c, ctx);
      sheet.cells.set(k, cell);
    } catch { bad++; }
  }
  if (bad && !warnings.includes('Some formulas could not be read; their last calculated values were kept.')) warnings.push('Some formulas could not be read; their last calculated values were kept.');
}

function readComments(u8: Uint8Array, sheet: Sheet, fonts: FontP[]): void {
  const authors: string[] = [];
  let cur: { r: number; c: number; author: string } | null = null;
  for (const { type, data: d } of records(u8)) {
    const r = new Rd(d);
    if (type === BRT.CommentAuthor) authors.push(r.wstr());
    else if (type === BRT.BeginComment) { const a = r.u32(); const rg = r.rfx(); cur = { r: rg.r1, c: rg.c1, author: authors[a] || '' }; }
    else if (type === BRT.CommentText && cur) {
      let text = readRichStr(r, fonts).t;
      if (cur.author && text.startsWith(cur.author + ':')) text = text.slice(cur.author.length + 1).replace(/^\r?\n/, '');
      const k = cellKey(cur.r, cur.c);
      if (!sheet.comments.has(k)) sheet.comments.set(k, { author: cur.author, text, time: 0, note: true });
    }
  }
}

/** Table part (BrtBeginList and its columns). */
function readTablePart(u8: Uint8Array): { id: number; def: TableDef } | null {
  let def: TableDef | null = null, id = 0;
  for (const { type, data: d } of records(u8)) {
    const r = new Rd(d);
    if (type === BRT.BeginList) {
      const range = r.rfx();
      r.u32(); // list type
      id = r.u32();
      const header = r.u32(), totals = r.u32();
      const flags = r.u32();
      r.skip(4 * 7);
      const name = r.nwstr() || '';
      const display = r.nwstr() || name;
      def = { id: newId('tbl'), name: display || `Table${id}`, range, columns: [], style: 'TableStyleMedium2', opts: { header: header !== 0, totals: totals > 0, banded: true, bandedCols: false, firstCol: false, lastCol: false, filterButton: !(flags & 0x4) } };
    } else if (type === BRT.BeginListCol && def) {
      r.skip(4 * 6);
      const name = r.nwstr();
      const caption = r.left >= 4 ? r.nwstr() : null;
      def.columns.push({ name: caption ?? name ?? `Column${def.columns.length + 1}` });
    } else if (type === BRT.TableStyleClient && def) {
      const f = r.u16();
      def.opts.firstCol = !!(f & 1); def.opts.lastCol = !!(f & 2); def.opts.banded = !!(f & 4); def.opts.bandedCols = !!(f & 8);
      const st = r.left >= 4 ? r.nwstr() : null;
      if (st) def.style = st;
    }
  }
  if (!def) return null;
  while (def.columns.length < def.range.c2 - def.range.c1 + 1) def.columns.push({ name: `Column${def.columns.length + 1}` });
  return { id, def };
}

export { parseXml };
