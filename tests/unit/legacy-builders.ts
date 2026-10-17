/**
 * Minimal, valid legacy spreadsheet files built from the published record layouts, for tests.
 * Nothing on the test machine writes these formats, so each builder assembles the records itself.
 * The same files were checked against LibreOffice's importers while the readers were written.
 */

class Bytes {
  parts: number[] = [];
  u8(...v: number[]): this { for (const x of v) this.parts.push(x & 0xff); return this; }
  u16(v: number): this { return this.u8(v, v >> 8); }
  u16be(v: number): this { return this.u8(v >> 8, v); }
  u32(v: number): this { return this.u8(v, v >> 8, v >> 16, v >>> 24); }
  u32be(v: number): this { return this.u8(v >>> 24, v >> 16, v >> 8, v); }
  f64(v: number): this { const dv = new DataView(new ArrayBuffer(8)); dv.setFloat64(0, v, true); return this.u8(...new Uint8Array(dv.buffer)); }
  str(s: string): this { for (const ch of s) this.parts.push(ch.charCodeAt(0) & 0xff); return this; }
  cstr(s: string): this { return this.str(s).u8(0); }
  bytes(b: ArrayLike<number>): this { for (let i = 0; i < b.length; i++) this.parts.push(b[i]); return this; }
  get length(): number { return this.parts.length; }
  out(): Uint8Array { return Uint8Array.from(this.parts); }
}

/** Type/length/body records (Lotus, Quattro, Works, BIFF). */
class Records {
  b = new Bytes();
  rec(op: number, body: Bytes | number[] = []): this {
    const data = body instanceof Bytes ? body.parts : body;
    this.b.u16(op).u16(data.length).bytes(data);
    return this;
  }
  get offset(): number { return this.b.length; }
  out(): Uint8Array { return this.b.out(); }
}
const B = () => new Bytes();

// ------------------------------------------------------------------------------------------
// Lotus 1-2-3 Release 2 (WK1)

/** Relative WK1 coordinate: bit 15 set, signed 14-bit offset. */
const rel = (d: number) => 0x8000 | (d & 0x3fff);

export function buildWk1(opts: { version?: number } = {}): Uint8Array {
  const r = new Records();
  r.rec(0x00, B().u16(opts.version ?? 0x0406));
  r.rec(0x06, B().u16(0).u16(0).u16(3).u16(6));
  // WINDOW1: default format (special/general = 0x71), default width 9.
  r.rec(0x07, B().u16(0).u16(0).u8(0x71, 0).u16(9).u16(8).u16(20).u16(0).u16(0).u16(0).u16(0).u16(0).u16(0).u16(4).u16(4).u16(72).u8(0));
  r.rec(0x08, B().u16(0).u8(14));                                     // COLW1: A is 14 characters wide
  r.rec(0x0b, B().str('PRICES').bytes(new Array(10).fill(0)).u16(1).u16(1).u16(1).u16(3)); // NAME PRICES = B2..B4
  r.rec(0x0f, B().u8(0xff).u16(0).u16(0).cstr("'Item"));             // A1 label
  r.rec(0x0f, B().u8(0xff).u16(1).u16(0).cstr('"Price'));            // B1 right-aligned label
  r.rec(0x0f, B().u8(0xff).u16(2).u16(0).cstr('^Qty'));              // C1 centered label
  r.rec(0x0f, B().u8(0xff).u16(0).u16(1).cstr("'Apples"));
  r.rec(0x0e, B().u8(0x22).u16(1).u16(1).f64(1.25));                 // B2 currency, 2 decimals
  r.rec(0x0d, B().u8(0x00).u16(2).u16(1).u16(4));                    // C2 fixed 0 decimals, integer 4
  r.rec(0x0f, B().u8(0xff).u16(0).u16(2).cstr("'Pears"));
  r.rec(0x0e, B().u8(0x22).u16(1).u16(2).f64(2.5));
  r.rec(0x0d, B().u8(0x00).u16(2).u16(2).u16(2));
  r.rec(0x0f, B().u8(0xff).u16(0).u16(3).cstr("'Plums"));
  r.rec(0x0e, B().u8(0x22).u16(1).u16(3).f64(0.75));
  r.rec(0x0d, B().u8(0x00).u16(2).u16(3).u16(8));
  // D2: +B2*C2 (relative references)
  const f1 = B().u8(0x01).u16(rel(-2)).u16(rel(0)).u8(0x01).u16(rel(-1)).u16(rel(0)).u8(0x0b, 0x03);
  r.rec(0x10, B().u8(0x22).u16(3).u16(1).f64(5).u16(f1.length).bytes(f1.parts));
  // D5: @SUM(D2..D4) with an absolute first corner
  const f2 = B().u8(0x02).u16(3).u16(1).u16(rel(0)).u16(rel(-1)).u8(0x50, 1, 0x03);
  r.rec(0x10, B().u8(0x22).u16(3).u16(4).f64(5).u16(f2.length).bytes(f2.parts));
  // A6: 2+3*4 with 1-2-3 precedence = 14 ; B6: @IF(C2>3#AND#C3<3,"yes","no") ; C6: @AVG(C2..C4)
  const f3 = B().u8(0x05).u16(2).u8(0x05).u16(3).u8(0x05).u16(4).u8(0x0b, 0x09, 0x03);
  r.rec(0x10, B().u8(0xff).u16(0).u16(5).f64(14).u16(f3.length).bytes(f3.parts));
  const f4 = B().u8(0x01).u16(rel(1)).u16(rel(-4)).u8(0x05).u16(3).u8(0x13).u8(0x01).u16(rel(1)).u16(rel(-3)).u8(0x05).u16(3).u8(0x12, 0x14)
    .u8(0x06).cstr('yes').u8(0x06).cstr('no').u8(0x3b, 0x03);
  r.rec(0x10, B().u8(0xff).u16(1).u16(5).bytes([0, 0, 0, 0, 1, 0, 0xf0, 0x7f]).u16(f4.length).bytes(f4.parts));
  r.rec(0x33, B().u8(0xff).u16(1).u16(5).cstr('yes'));
  const f5 = B().u8(0x02).u16(rel(0)).u16(rel(-4)).u16(rel(0)).u16(rel(-2)).u8(0x51, 1, 0x03);
  r.rec(0x10, B().u8(0x02).u16(2).u16(5).f64(14 / 3).u16(f5.length).bytes(f5.parts));
  // A7: date 1-Jan-2000 shown D1 (special 2)
  r.rec(0x0e, B().u8(0x72).u16(0).u16(6).f64(36526));
  r.rec(0x01);
  return r.out();
}

// ------------------------------------------------------------------------------------------
// Lotus 1-2-3 Release 3 (WK3): two sheets, a cross-sheet formula, a named sheet

function f80(v: number): number[] {
  if (v === 0) return new Array(10).fill(0);
  const sign = v < 0 ? 0x8000 : 0;
  v = Math.abs(v);
  let e = Math.floor(Math.log2(v));
  let m = v / Math.pow(2, e);
  if (m >= 2) { m /= 2; e++; }
  const hi = Math.floor(m * 2 ** 31) >>> 0;
  const lo = Math.round((m * 2 ** 31 - Math.floor(m * 2 ** 31)) * 2 ** 32) >>> 0;
  const b = B().u32(lo).u32(hi).u16(sign | (e + 16383));
  return b.parts;
}

export function buildWk3(): Uint8Array {
  const r = new Records();
  r.rec(0x00, B().u16(0x1000).bytes(new Array(24).fill(0)));
  r.rec(0x1b, B().u16(14000).u16(1).cstr('Totals'));                        // sheet 1 is named Totals
  r.rec(0x07, B().u8(0, 0).u16(0).u8(0, 20));                              // sheet A: column A 20 wide
  r.rec(0x16, B().u16(0).u8(0, 0).cstr("'Region"));                       // A:A1
  r.rec(0x16, B().u16(0).u8(0, 1).cstr('"Sales'));                        // A:B1
  r.rec(0x16, B().u16(1).u8(0, 0).cstr("'North"));
  r.rec(0x17, B().u16(1).u8(0, 1).bytes(f80(1250.5)));                     // A:B2 10-byte number
  r.rec(0x16, B().u16(2).u8(0, 0).cstr("'South"));
  r.rec(0x18, B().u16(2).u8(0, 1).u16(980 << 1));                          // A:B3 small integer 980
  r.rec(0x18, B().u16(3).u8(0, 1).u16(0x37));                              // A:B4 small number 0.015
  // A:B5: @SUM(B2..B4)  (flags: all relative)
  const sum = B().u8(0x02, 0xff).u16(1).u8(0, 1).u16(3).u8(0, 1).u8(0x50, 1, 0x03);
  r.rec(0x19, B().u16(4).u8(0, 1).bytes(f80(2230.515)).bytes(sum.parts));
  r.rec(0x16, B().u16(0).u8(1, 0).cstr("'Total"));                         // Totals!A1
  // Totals!B1: +A:B5*2 (reference to another sheet; snum 2 = 4)
  const x = B().u8(0x01, 0x03).u16(4).u8(0, 1).u8(0x05).u16(4).u8(0x11, 0x03);
  r.rec(0x19, B().u16(0).u8(1, 1).bytes(f80(4461.03)).bytes(x.parts));
  r.rec(0x09, B().u16(0).str('SALES').bytes(new Array(11).fill(0)).u16(1).u8(0, 1).u16(3).u8(0, 1)); // SALES = A:B2..A:B4
  r.rec(0x01);
  return r.out();
}

// ------------------------------------------------------------------------------------------
// Lotus 1-2-3 97 (123): compressed numbers and 8-byte formula results

export function lnumBytes(v: number): number[] {
  // Integer when possible, else mantissa and a negative decimal exponent.
  const neg = v < 0 ? 0x20 : 0;
  v = Math.abs(v);
  let e = 0;
  while (!Number.isInteger(v) && e < 15) { v = Math.round(v * 10 * 1e9) / 1e9; e++; }
  const code = (v * 64) + (e ? 0x10 | e : 0) + neg;
  return B().u32(code >>> 0).parts;
}

export function build123(): Uint8Array {
  const r = new Records();
  r.rec(0x00, B().u16(0x1005).bytes(new Array(24).fill(0)));
  r.rec(0x23, B().u16(14000).u16(0).cstr('Data'));
  r.rec(0x16, B().u16(0).u8(0, 0).cstr("'Amount"));
  r.rec(0x25, B().u16(1).u8(0, 0).bytes(lnumBytes(12.5)));
  r.rec(0x25, B().u16(2).u8(0, 0).bytes(lnumBytes(-3)));
  // A4: +A2+A3*2 (4-byte number constant)
  const f = B().u8(0x01, 0x07).u16(1).u8(0, 0).u8(0x01, 0x07).u16(2).u8(0, 0).u8(0x05).bytes(lnumBytes(2)).u8(0x11, 0x0f, 0x03);
  r.rec(0x28, B().u16(3).u8(0, 0).f64(6.5).bytes(f.parts));
  r.rec(0x01);
  return r.out();
}

// ------------------------------------------------------------------------------------------
// Microsoft Works for DOS (Lotus records with Works extensions)

export function buildWorksWks(): Uint8Array {
  const r = new Records();
  r.rec(0x00, B().u16(0x0404));
  r.rec(0x5405, B().u16(0x0da1));
  r.rec(0x0f, B().u8(0xf1).u16(0).u16(0).cstr("'Budget"));
  r.rec(0x5402, B().u8(0xa5, 0x20));                                      // bold
  r.rec(0x0e, B().u8(0xa2).u16(1).u16(0).f64(99.5));                     // currency
  r.rec(0x01);
  return r.out();
}

// ------------------------------------------------------------------------------------------
// Quattro Pro for DOS (WQ1) and for Windows (WB1)

export function buildWq1(): Uint8Array {
  const r = new Records();
  r.rec(0x00, B().u16(0x5120));
  r.rec(0x0f, B().u8(0xff).u16(0).u16(0).u8(0x27, 5).str('Hello'));      // Quattro: prefix + length byte
  r.rec(0x0d, B().u8(0x93).u16(1).u16(0).u16(1000));                     // scientific, 3 decimals
  const f = B().u8(0x01).u16(rel(-1)).u16(rel(0)).u8(0x05).u16(2).u8(0x0c, 0x03);
  r.rec(0x10, B().u8(0xff).u16(2).u16(0).f64(500).u16(f.length).bytes(f.parts));
  r.rec(0x01);
  return r.out();
}

export function buildWb1(): Uint8Array {
  const r = new Records();
  r.rec(0x00, B().u16(0x1001));
  r.rec(0xca, [0]); r.rec(0xcc, B().cstr('Inputs'));
  r.rec(0x0f, B().u8(0, 0).u16(0).u16(1).cstr("'Rate"));
  r.rec(0x0e, B().u8(1, 0).u16(0).u16(1).f64(0.05));
  r.rec(0x0d, B().u8(1, 0).u16(1).u16(1).u16(40));
  r.rec(0x0d, B().u8(1, 0).u16(2).u16(1).u16(60));
  // B4: @SUM(B2..B3) : code 02 50 01 03, block reference (bit 12), relative rows/cols
  const refs = B().u16(0x1000).u8(0, 0).u16(0x6000 | ((-2) & 0x1fff)).u8(0, 0).u16(0x6000 | ((-1) & 0x1fff));
  const code = [0x02, 0x50, 1, 0x03];
  r.rec(0x10, B().u8(1, 0).u16(3).u16(1).f64(100).u16(0).u16(code.length + refs.length).u16(code.length).bytes(code).bytes(refs.parts));
  r.rec(0xcb, [0]);
  r.rec(0xca, [1]); r.rec(0xcc, B().cstr('Results'));
  // Results!A1: +$Inputs:B1*2   (absolute page 0, column 1, row 0)
  const refs2 = B().u16(0).u8(1, 0).u16(0);
  const code2 = [0x01, 0x05, 2, 0, 0x0b, 0x03];
  r.rec(0x10, B().u8(0, 1).u16(0).u16(1).f64(0.1).u16(0).u16(code2.length + refs2.length).u16(code2.length).bytes(code2).bytes(refs2.parts));
  r.rec(0xcb, [1]);
  r.rec(0x01);
  return r.out();
}

// ------------------------------------------------------------------------------------------
// Excel 2.x (BIFF2) worksheet and Excel 4.0 workbook (BIFF4W)

export function buildBiff2(): Uint8Array {
  const r = new Records();
  r.rec(0x0009, B().u16(0x0002).u16(0x0010));
  r.rec(0x0031, B().u16(200).u16(0).u8(5).str('Arial'));                 // font 0
  r.rec(0x0031, B().u16(200).u16(1).u8(5).str('Arial'));                 // font 1 bold
  r.rec(0x001e, B().u8(7).str('General'));                               // format 0
  r.rec(0x001e, B().u8(4).str('0.00'));                                  // format 1
  r.rec(0x0024, B().u8(0, 0).u16(20 * 256));                             // column A 20 characters
  r.rec(0x0004, B().u16(0).u16(0).u8(0, 0x40, 0).u8(5).str('Total'));    // A1 bold label (font 1)
  r.rec(0x0002, B().u16(1).u16(0).u8(0, 0, 3).u16(7));                   // A2 integer 7, right aligned
  r.rec(0x0003, B().u16(2).u16(0).u8(0, 1, 0).f64(2.5));                 // A3 2.5 shown 0.00
  // A4: =SUM(A2:A3)*2  -> ptgArea(0x25) ptgFuncVar(0x42, argc 1, SUM=4) ptgInt 2 ptgMul
  const fm = B().u8(0x25).u16(1 | 0xc000).u16(2 | 0xc000).u8(0, 0).u8(0x42, 1, 4).u8(0x1e).u16(2).u8(0x05);
  r.rec(0x0006, B().u16(3).u16(0).u8(0, 0, 0).f64(19).u8(0).u8(fm.length).bytes(fm.parts));
  r.rec(0x0005, B().u16(4).u16(0).u8(0, 0, 0).u8(1, 0));                 // A5 TRUE
  r.rec(0x000a);
  return r.out();
}

export function buildBiff4w(): Uint8Array {
  const sheet = (label: string, value: number, formulaOf?: string) => {
    const r = new Records();
    r.rec(0x0409, B().u16(0x0004).u16(0x0010).u16(0));
    r.rec(0x0231, B().u16(200).u16(0).u16(0x7fff).u8(5).str('Arial'));
    r.rec(0x041e, B().u16(0).u8(7).str('General'));
    r.rec(0x041e, B().u16(0).u8(4).str('0.00'));
    // XF4: font 0, format 0 / format 1
    r.rec(0x0443, B().u8(0, 0).u16(0).u8(0x20, 0).u32(0).u32(0));
    r.rec(0x0443, B().u8(0, 1).u16(0).u8(0x20, 0).u32(0).u32(0));
    r.rec(0x0204, B().u16(0).u16(0).u16(0).u16(label.length).str(label));
    r.rec(0x0203, B().u16(1).u16(0).u16(1).f64(value));
    if (formulaOf) {
      // =A2*2 : ptgRef (0x44) row 1 (relative), col 0 ; ptgInt 2 ; ptgMul
      const fm = B().u8(0x44).u16(1 | 0xc000).u8(0).u8(0x1e).u16(2).u8(0x05);
      r.rec(0x0406, B().u16(2).u16(0).u16(0).f64(value * 2).u16(0).u16(fm.length).bytes(fm.parts));
    }
    r.rec(0x000a);
    return r.out();
  };
  const s1 = sheet('First', 10, 'x'), s2 = sheet('Second', 3.5);
  const names = ['Alpha', 'Beta'];
  // Globals: BOF4 (workbook), BUNDLESHEET per sheet (stream offset of its BOF), EOF.
  const globalsLen = (4 + 6) + names.reduce((n, nm) => n + 4 + 7 + nm.length, 0) + 4;
  const g = new Records();
  g.rec(0x0409, B().u16(0x0004).u16(0x0100).u16(0));
  let pos = globalsLen;
  [s1, s2].forEach((s, i) => { g.rec(0x008f, B().u32(pos).u16(0).u8(names[i].length).str(names[i])); pos += s.length; });
  g.rec(0x000a);
  const out = new Uint8Array(globalsLen + s1.length + s2.length);
  out.set(g.out(), 0); out.set(s1, globalsLen); out.set(s2, globalsLen + s1.length);
  return out;
}

// ------------------------------------------------------------------------------------------
// dBASE III table with a memo file

export function buildDbf(opts: { codepage?: 'cp437' | 'cp1252'; memo?: boolean } = {}): { dbf: Uint8Array; dbt: Uint8Array } {
  const fields: [string, string, number, number][] = [['NAME', 'C', 12, 0], ['QTY', 'N', 5, 0], ['PRICE', 'N', 8, 2], ['SOLD', 'D', 8, 0], ['OK', 'L', 1, 0]];
  if (opts.memo) fields.push(['NOTES', 'M', 10, 0]);
  const rows: (string | number | boolean)[][] = [['Café', 3, 4.5, '19990315', true, 1], ['Widget', 12, 0.25, '20240101', false, 2], ['Deleted', 1, 1, '20000101', true, 0]];
  const rl = 1 + fields.reduce((n, f) => n + f[2], 0);
  const hl = 32 + fields.length * 32 + 1;
  const b = B().u8(opts.memo ? 0x83 : 0x03, 124, 3, 15).u32(rows.length).u16(hl).u16(rl).bytes(new Array(17).fill(0)).u8(opts.codepage === 'cp437' ? 0x01 : 0x03, 0, 0);
  for (const [nm, t, len, dec] of fields) b.str(nm.padEnd(11, '\0')).str(t).u32(0).u8(len, dec).bytes(new Array(14).fill(0));
  b.u8(0x0d);
  const enc = (s: string) => Array.from(s, (ch) => (ch === 'é' ? (opts.codepage === 'cp437' ? 0x82 : 0xe9) : ch.charCodeAt(0)));
  rows.forEach((row, k) => {
    b.u8(k === 2 ? 0x2a : 0x20);
    fields.forEach((f, i) => {
      const v = row[i];
      let s: string;
      if (f[1] === 'N') s = (f[3] ? (v as number).toFixed(f[3]) : String(v)).padStart(f[2]);
      else if (f[1] === 'L') s = v ? 'T' : 'F';
      else if (f[1] === 'M') s = v ? String(v).padStart(10) : ' '.repeat(10);
      else s = String(v).padEnd(f[2]);
      b.bytes(enc(s).slice(0, f[2]));
    });
  });
  b.u8(0x1a);
  const memo = B().u32(3).bytes(new Array(508).fill(0));
  memo.str('First note').u8(0x1a, 0x1a).bytes(new Array(500).fill(0));
  memo.str('Second note').u8(0x1a, 0x1a);
  return { dbf: b.out(), dbt: memo.out() };
}

// ------------------------------------------------------------------------------------------
// VisiCalc

export const VISICALC_SAMPLE = [
  '>A1:"ITEM', '>B1:/FR"COST', '>A2:"PENS', '>B2:/F$2.5', '>A3:"PADS', '>B3:/F$4',
  '>B4:/F$@SUM(B2...B3)', '>C4:2+3*4', '>D4:+B2*(1+B3)', '>A5:/-=', '/W1', '/GOC', '/GRA', '/GC12', '/X>A1:>A1:', '',
].join('\r\n');
