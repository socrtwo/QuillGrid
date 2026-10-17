/**
 * Decompiler for the Reverse Polish formula code of Lotus 1-2-3 and the programs that copied its
 * file format (Symphony, Quattro, Microsoft Works).
 *
 * Release 1A/2.x (WKS/WK1) code is described in Lotus's public "Worksheet File Format"
 * specification (appendix B and the 1985 addendum). Release 3 and later (WK3, WK4, 123) move the
 * operators up by six, address cells as row/sheet/column with a separate relative-flags byte and
 * add more functions; that layout was established by inspecting files.
 *
 * The code is turned back into 1-2-3 formula text and then translated by the 1-2-3 formula
 * translator (src/heritage/lotus.ts), which keeps 1-2-3's meaning: zero-based offsets, argument
 * orders and its treatment of labels in statistical functions.
 */
import { colName } from '../../model/address';
import { translateLotus } from '../../heritage/lotus';
import { getFn, FUNCTIONS } from '../../calc/functions/registry';
import { f64, f80, i16, u16, u32, snum, lnum, cstr, type Decoder } from './bin';

/** [1-2-3 name, fixed argument count, or -1 when an argument count byte follows]. */
export const LOTUS_FUNCS: Record<number, [string, number]> = {
  0x1f: ['NA', 0], 0x20: ['ERR', 0], 0x21: ['ABS', 1], 0x22: ['INT', 1], 0x23: ['SQRT', 1], 0x24: ['LOG', 1], 0x25: ['LN', 1],
  0x26: ['PI', 0], 0x27: ['SIN', 1], 0x28: ['COS', 1], 0x29: ['TAN', 1], 0x2a: ['ATAN2', 2], 0x2b: ['ATAN', 1], 0x2c: ['ASIN', 1],
  0x2d: ['ACOS', 1], 0x2e: ['EXP', 1], 0x2f: ['MOD', 2], 0x30: ['CHOOSE', -1], 0x31: ['ISNA', 1], 0x32: ['ISERR', 1],
  0x33: ['FALSE', 0], 0x34: ['TRUE', 0], 0x35: ['RAND', 0], 0x36: ['DATE', 3], 0x37: ['NOW', 0], 0x38: ['PMT', 3], 0x39: ['PV', 3],
  0x3a: ['FV', 3], 0x3b: ['IF', 3], 0x3c: ['DAY', 1], 0x3d: ['MONTH', 1], 0x3e: ['YEAR', 1], 0x3f: ['ROUND', 2], 0x40: ['TIME', 3],
  0x41: ['HOUR', 1], 0x42: ['MINUTE', 1], 0x43: ['SECOND', 1], 0x44: ['ISNUMBER', 1], 0x45: ['ISSTRING', 1], 0x46: ['LENGTH', 1],
  0x47: ['VALUE', 1], 0x48: ['STRING', 2], 0x49: ['MID', 3], 0x4a: ['CHAR', 1], 0x4b: ['CODE', 1], 0x4c: ['FIND', 3],
  0x4d: ['DATEVALUE', 1], 0x4e: ['TIMEVALUE', 1], 0x4f: ['CELLPOINTER', 1], 0x50: ['SUM', -1], 0x51: ['AVG', -1], 0x52: ['COUNT', -1],
  0x53: ['MIN', -1], 0x54: ['MAX', -1], 0x55: ['VLOOKUP', 3], 0x56: ['NPV', 2], 0x57: ['VAR', -1], 0x58: ['STD', -1], 0x59: ['IRR', 2],
  0x5a: ['HLOOKUP', 3], 0x5b: ['DSUM', 3], 0x5c: ['DAVG', 3], 0x5d: ['DCOUNT', 3], 0x5e: ['DMIN', 3], 0x5f: ['DMAX', 3], 0x60: ['DVAR', 3],
  0x61: ['DSTD', 3], 0x62: ['INDEX', 3], 0x63: ['COLS', 1], 0x64: ['ROWS', 1], 0x65: ['REPEAT', 2], 0x66: ['UPPER', 1], 0x67: ['LOWER', 1],
  0x68: ['LEFT', 2], 0x69: ['RIGHT', 2], 0x6a: ['REPLACE', 4], 0x6b: ['PROPER', 1], 0x6c: ['CELL', 2], 0x6d: ['TRIM', 1], 0x6e: ['CLEAN', 1],
  0x6f: ['S', 1], 0x70: ['N', 1], 0x71: ['EXACT', 2], 0x73: ['@', 1], 0x74: ['RATE', 3], 0x75: ['TERM', 3], 0x76: ['CTERM', 3],
  0x77: ['SLN', 3], 0x78: ['SYD', 4], 0x79: ['DDB', 4],
};

/** Release 3 and later additions. */
export const LOTUS3_FUNCS: Record<number, [string, number]> = {
  0x7d: ['SUMPRODUCT', -1], 0x7f: ['DGET', -1], 0x84: ['VDB', -1], 0x85: ['DVARS', -1], 0x86: ['DSTDS', -1], 0x87: ['VARS', -1],
  0x88: ['STDS', -1], 0x8f: ['RANK', -1], 0x99: ['PUREAVG', -1], 0x9a: ['PURECOUNT', -1], 0x9b: ['PUREMAX', -1], 0x9c: ['PUREMIN', -1],
  0x9d: ['PURESTD', -1], 0x9e: ['PUREVAR', -1], 0x9f: ['PURESTDS', -1], 0xa0: ['PUREVARS', -1],
};

const WK1_OPS: Record<number, [string, 1 | 2]> = {
  0x08: ['-', 1], 0x09: ['+', 2], 0x0a: ['-', 2], 0x0b: ['*', 2], 0x0c: ['/', 2], 0x0d: ['^', 2], 0x0e: ['=', 2], 0x0f: ['<>', 2],
  0x10: ['<=', 2], 0x11: ['>=', 2], 0x12: ['<', 2], 0x13: ['>', 2], 0x14: ['#AND#', 2], 0x15: ['#OR#', 2], 0x16: ['#NOT#', 1],
  0x17: ['+', 1], 0x18: ['&', 2],
};

export type Dialect = 'wk1' | 'wk3' | 'wk4';

export interface FormulaCtx {
  dialect: Dialect;
  /** Cell holding the formula (zero-based). */
  row: number; col: number; sheet: number;
  dec: Decoder;
  /** Extra or replacement functions (Quattro, Works). */
  funcs?: Record<number, [string, number]>;
}

export class FormulaError extends Error {}

type N = { s: string; atom: boolean };
const atom = (s: string): N => ({ s, atom: !s.startsWith('-') });
const wrap = (n: N) => (n.atom ? n.s : `(${n.s})`);

function numText(v: number): string {
  if (!Number.isFinite(v)) throw new FormulaError('non-finite constant');
  if (Object.is(v, -0)) return '0';
  const s = String(Math.abs(v));
  return v < 0 ? `-${s}` : s;
}

/** Decompiles formula code to 1-2-3 formula text (without a leading +). Throws FormulaError. */
export function lotusCodeToText(code: Uint8Array, ctx: FormulaCtx): string {
  const st: N[] = [];
  const pop = () => { const x = st.pop(); if (!x) throw new FormulaError('formula stack underflow'); return x; };
  const wk1 = ctx.dialect === 'wk1';
  const opBase = wk1 ? 0 : 6;
  const funcs = { ...LOTUS_FUNCS, ...(wk1 ? {} : LOTUS3_FUNCS), ...(ctx.funcs || {}) };
  const sheetPrefix = (sh: number, rel: boolean, force: boolean) => (sh !== ctx.sheet || force ? `${rel ? '' : '$'}${colName(sh)}:` : '');
  const cell1 = (colv: number, rowv: number): string => {
    // 1A/2.x: bit 15 marks a relative coordinate stored as a signed 14-bit offset.
    const res = (v: number, base: number) => {
      if (v & 0x8000) { let d = v & 0x3fff; if (d & 0x2000) d -= 0x4000; return { n: base + d, rel: true }; }
      return { n: v & 0x3fff, rel: false };
    };
    const c = res(colv, ctx.col), r = res(rowv, ctx.row);
    if (c.n < 0 || r.n < 0 || c.n > 16383 || r.n > 1048575) throw new FormulaError('reference out of range');
    return `${c.rel ? '' : '$'}${colName(c.n)}${r.rel ? '' : '$'}${r.n + 1}`;
  };
  const cell3 = (row: number, sheet: number, col: number, flags: number, forceSheet: boolean): string => {
    const colRel = !!(flags & 1), rowRel = !!(flags & 2), shRel = !!(flags & 4);
    return `${sheetPrefix(sheet, shRel, forceSheet)}${colRel ? '' : '$'}${colName(col)}${rowRel ? '' : '$'}${row + 1}`;
  };
  let i = 0;
  while (i < code.length) {
    const op = code[i++];
    if (op === 0x03) break;
    switch (op) {
      case 0x00: {
        if (ctx.dialect === 'wk3') { st.push(atom(numText(f80(code, i)))); i += 10; }
        else { st.push(atom(numText(f64(code, i)))); i += 8; }
        continue;
      }
      case 0x01: {
        if (wk1) { st.push(atom(cell1(u16(code, i), u16(code, i + 2)))); i += 4; }
        else { st.push(atom(cell3(u16(code, i + 1), code[i + 3], code[i + 4], code[i], false))); i += 5; }
        continue;
      }
      case 0x02: {
        if (wk1) { st.push(atom(`${cell1(u16(code, i), u16(code, i + 2))}..${cell1(u16(code, i + 4), u16(code, i + 6))}`)); i += 8; }
        else {
          const fl = code[i];
          const s1 = code[i + 3], s2 = code[i + 7];
          const cross = s1 !== s2;
          const a = cell3(u16(code, i + 1), s1, code[i + 4], fl & 7, cross);
          const b = cell3(u16(code, i + 5), s2, code[i + 8], (fl >> 3) & 7, cross);
          st.push(atom(`${a}..${b}`));
          i += 9;
        }
        continue;
      }
      case 0x04: { const x = pop(); st.push(atom(`(${x.s})`)); continue; }
      case 0x05: {
        if (wk1) { st.push(atom(numText(i16(code, i)))); i += 2; }
        else if (ctx.dialect === 'wk3') { st.push(atom(numText(snum(u16(code, i))))); i += 2; }
        else { st.push(atom(numText(lnum(u32(code, i))))); i += 4; }
        continue;
      }
      case 0x06: {
        let e = i; while (e < code.length && code[e] !== 0) e++;
        const s = cstr(code, i, ctx.dec);
        st.push(atom(`"${s.replace(/"/g, '""')}"`));
        i = e + 1;
        continue;
      }
    }
    if (!wk1 && (op === 0x07 || op === 0x08)) {
      // Named range by name.
      let e = i; while (e < code.length && code[e] !== 0) e++;
      const nm = cstr(code, i, ctx.dec);
      i = e + 1;
      if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(nm)) throw new FormulaError(`range name "${nm}" cannot be written in a formula`);
      st.push(atom(nm));
      continue;
    }
    const opk = op - opBase;
    if (opk >= 0x08 && opk <= 0x18 && WK1_OPS[opk]) {
      const [sym, n] = WK1_OPS[opk];
      if (n === 1) {
        const a = pop();
        st.push(sym === '#NOT#' ? { s: `#NOT#${wrap(a)}`, atom: false } : { s: `${sym}${wrap(a)}`, atom: false });
      } else {
        const b = pop(), a = pop();
        st.push({ s: `${wrap(a)}${sym}${wrap(b)}`, atom: false });
      }
      continue;
    }
    const fn = funcs[op];
    if (!fn) throw new FormulaError(`function code 0x${op.toString(16)} is not known`);
    const [name, fixed] = fn;
    let n = fixed;
    if (fixed < 0) n = code[i++];
    if (st.length < n) throw new FormulaError('formula stack underflow');
    const args = st.splice(st.length - n, n).map((x) => x.s);
    if (name === '@') st.push(atom(`@@(${args.join(',')})`));
    else if (!args.length && ['PI', 'TRUE', 'FALSE', 'NA', 'ERR', 'NOW', 'TODAY', 'RAND'].includes(name)) st.push(atom(`@${name}`));
    else st.push(atom(`@${name}(${args.join(',')})`));
  }
  if (st.length !== 1) throw new FormulaError('unbalanced formula');
  return st[0].s;
}

/**
 * Formula code to this app's formula text, through the 1-2-3 translator. Returns the 1-2-3 text
 * as well, so an untranslatable formula can be reported in its original form.
 */
export function lotusFormula(code: Uint8Array, ctx: FormulaCtx, sheetNames: string[]): { f: string | null; lotus: string | null; error?: string } {
  let lotus: string | null = null;
  try {
    lotus = lotusCodeToText(code, ctx);
    const f = translateLotus(lotus, { sheetNames });
    const unknown = unknownFunctions(f);
    if (unknown.length) return { f: null, lotus, error: `unknown function ${unknown.join(', ')}` };
    return { f, lotus };
  } catch (e: any) {
    return { f: null, lotus, error: e?.message || String(e) };
  }
}

/** Function names used in a formula that this app does not implement (empty when the registry is not loaded). */
export function unknownFunctions(f: string): string[] {
  if (!FUNCTIONS.size) return [];
  const out: string[] = [];
  const re = /"(?:[^"]|"")*"|'(?:[^']|'')*'|([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(f))) if (m[1] && !getFn(m[1]) && !out.includes(m[1].toUpperCase())) out.push(m[1].toUpperCase());
  return out;
}
