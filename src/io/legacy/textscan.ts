/**
 * Best-effort readers for spreadsheet formats whose layout is not documented: the values that can be
 * found are listed in reading order in column A, and the import says so.
 *
 *  - AppleWorks / ClarisWorks (.cwk): cell records hold a big-endian length, a type (2 = number as a
 *    68k 80-bit float, 3 = text), a style number and the value; the records are found by pattern.
 *  - SuperCalc (.cal) and binary Multiplan: readable text runs and nothing else.
 */
import { Workbook, Sheet } from '../../model/workbook';
import { u16be, codepage } from './bin';
import { stringsScan } from '../recover/salvage';

function sheetFor(title: string, sheetName: string): { book: Workbook; sh: Sheet } {
  const book = new Workbook(false);
  const sh = new Sheet(book, sheetName);
  sh.colW.set(0, 320);
  book.sheets.push(sh);
  if (title) book.props.title = title;
  return { book, sh };
}

/** 68000 extended precision (big-endian): sign and 15-bit exponent, then a 64-bit mantissa. */
function f80be(b: Uint8Array, o: number): number {
  const se = u16be(b, o);
  const exp = se & 0x7fff;
  if (exp === 0x7fff) return NaN;
  const hi = ((b[o + 2] << 24) | (b[o + 3] << 16) | (b[o + 4] << 8) | b[o + 5]) >>> 0;
  const lo = ((b[o + 6] << 24) | (b[o + 7] << 16) | (b[o + 8] << 8) | b[o + 9]) >>> 0;
  if (!exp && !hi && !lo) return 0;
  return (se & 0x8000 ? -1 : 1) * (hi * 4294967296 + lo) * Math.pow(2, exp - 16383 - 63);
}

export function importAppleWorks(data: ArrayBuffer, name = '', warnings: string[] = []): Workbook {
  const b = new Uint8Array(data);
  if (String.fromCharCode(b[4], b[5], b[6], b[7]) !== 'BOBO') throw new Error('This is not an AppleWorks or ClarisWorks document.');
  const dec = codepage(10000);
  const { book, sh } = sheetFor(name.replace(/\.[^.]+$/, ''), 'Values');
  let row = 0;
  for (let p = 8; p + 6 < b.length && row < 1_000_000;) {
    const len = u16be(b, p), type = b[p + 2];
    if (type === 3 && len >= 5 && len < 300 && b[p + 3] === 0 && b[p + 4] === 0) {
      const n = b[p + 6];
      if (n && n <= len - 4 && p + 7 + n <= b.length) {
        const t = dec(b.subarray(p + 7, p + 7 + n));
        if (/^[\x20-\x7e -￿]+$/.test(t)) { sh.setComputed(row++, 0, { v: t }); p += 2 + len; continue; }
      }
    }
    if (type === 2 && len === 14 && b[p + 3] === 0 && b[p + 4] === 0) {
      const v = f80be(b, p + 6);
      if (Number.isFinite(v) && Math.abs(v) < 1e300) { sh.setComputed(row++, 0, { v }); p += 16; continue; }
    }
    p++;
  }
  if (!row) throw new Error('No spreadsheet values were found in this AppleWorks document.');
  warnings.push('AppleWorks and ClarisWorks documents are read for their values only: text and numbers are listed in column A in the order they are stored; cell positions, formulas and formatting are not recovered.');
  return book;
}

export function importValuesScan(data: ArrayBuffer, name = '', warnings: string[] = [], program = 'this program'): Workbook {
  const text = stringsScan(new Uint8Array(data), 3);
  const { book, sh } = sheetFor(name.replace(/\.[^.]+$/, ''), 'Text');
  let row = 0;
  for (const line of text.split(/\n/)) {
    const t = line.trim();
    // Keep runs that read as words or numbers, not stray bytes of binary data.
    if (!t || !/^[\x20-\x7e\u00a0-\u024f]+$/.test(t) || !(/[A-Za-z]{3}/.test(t) || /^[-+]?\d+(\.\d+)?$/.test(t))) continue;
    const n = Number(t);
    sh.setComputed(row++, 0, { v: /^[-+]?\d+(\.\d+)?$/.test(t) && Number.isFinite(n) ? n : t });
  }
  if (!row) throw new Error(`No readable text was found in this ${program} file.`);
  warnings.push(`The ${program} file format is not documented; only its readable text was recovered (one item per row). Save the worksheet as SYLK, DIF or WKS in ${program} to keep cells and formulas.`);
  return book;
}
