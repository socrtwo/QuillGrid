/** Information functions. */
import { defFn } from './registry';
import {
  isErr, isLambda, isMatrix, isRef, isRefs, toScalar, num, str, ERR, type Ctx, type Value, type RefV,
} from '../values';
import type { Scalar } from '../../model/types';
import { cellName, quoteSheet, colName } from '../../model/address';
import { isDateFormat } from '../../model/numfmt';

const S = (a: Value) => a as Scalar;

const ERROR_NUMBERS: Record<string, number> = {
  '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#GETTING_DATA': 8,
  '#SPILL!': 9, '#CONNECT!': 10, '#BLOCKED!': 11, '#UNKNOWN!': 12, '#FIELD!': 13, '#CALC!': 14,
};

/** Format code reported by CELL("format"). */
export function cellFormatCode(fmt: string | undefined): string {
  const f = (fmt || 'General').trim();
  if (/^general$/i.test(f) || f === '@') return 'G';
  const first = f.split(';')[0];
  const neg = f.split(';')[1] ?? '';
  const colorNeg = /\[(red|blue|green|black|white|yellow|magenta|cyan|color\s*\d+)\]/i.test(neg);
  const paren = /\(/.test(first.replace(/"[^"]*"|\\./g, '')) || /\(/.test(neg.replace(/"[^"]*"|\\./g, '').replace(/_\(/g, ''));
  const suffix = (colorNeg ? '-' : '') + (paren ? '()' : '');
  if (isDateFormat(f)) {
    const l = first.toLowerCase().replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '');
    if (/h/.test(l) || /s/.test(l)) {
      const ap = /am\/pm|a\/p/.test(l);
      const sec = /s/.test(l);
      if (/[dy]/.test(l)) return 'D4';
      if (ap) return sec ? 'D6' : 'D7';
      return sec ? 'D8' : 'D9';
    }
    const hasY = /y/.test(l), hasD = /d/.test(l), hasM = /m/.test(l);
    if (hasD && hasM && hasY) return /mmm/.test(l) ? 'D1' : 'D4';
    if (hasD && hasM) return /mmm/.test(l) ? 'D2' : 'D5';
    if (hasM && hasY) return 'D3';
    return 'D4';
  }
  const body = first.replace(/"[^"]*"|\\.|_.|\*./g, '');
  if (/\?\/|#\s*\?/.test(body) && /\//.test(body)) return 'G';
  const decMatch = /\.([0#?]+)/.exec(body);
  const dec = decMatch ? decMatch[1].length : 0;
  if (/E[+-]/i.test(body)) return 'S' + dec + suffix;
  if (/%/.test(body)) return 'P' + dec + suffix;
  if (/[$€£¥]/.test(first) || /\[\$/.test(first)) return 'C' + dec + suffix;
  if (/[0#?],[0#?]/.test(body)) return ',' + dec + suffix;
  if (/[0#?]/.test(body)) return 'F' + dec + suffix;
  return 'G';
}

function refArg(v: Value | undefined): RefV | null {
  if (v === undefined || v === null) return null;
  if (isRef(v)) return v;
  if (isRefs(v)) return v.areas[0] ?? null;
  return null;
}

function cellInfo(type: string, ref: RefV | null, ctx: Ctx): Value {
  const sheet = ref ? ref.sheet : ctx.sheet;
  const r = ref ? ref.r1 : ctx.r, c = ref ? ref.c1 : ctx.c;
  const cell = sheet.get(r, c);
  const style = sheet.styleOf(r, c);
  const fmt = style.numFmt;
  switch (type) {
    case 'address': {
      const a = cellName(r, c, true, true);
      return sheet === ctx.sheet ? a : `${quoteSheet(sheet.name)}!${a}`;
    }
    case 'col': return c + 1;
    case 'row': return r + 1;
    case 'color': {
      const neg = (fmt || '').split(';')[1] ?? '';
      return /\[(red|blue|green|black|white|yellow|magenta|cyan|color\s*\d+)\]/i.test(neg) ? 1 : 0;
    }
    case 'contents': return ctx.cell(sheet, r, c);
    case 'filename': {
      const fn = (ctx.book as any).fileName as string | undefined;
      if (!fn) return '';
      const i = Math.max(fn.lastIndexOf('/'), fn.lastIndexOf('\\'));
      return `${fn.slice(0, i + 1)}[${fn.slice(i + 1)}]${sheet.name}`;
    }
    case 'format': return cellFormatCode(fmt);
    case 'parentheses': {
      const first = (fmt || '').split(';')[0].replace(/"[^"]*"|\\.|_./g, '');
      return /\(/.test(first) ? 1 : 0;
    }
    case 'prefix': {
      const v = ctx.cell(sheet, r, c);
      if (typeof v !== 'string') return '';
      switch (style.hAlign) {
        case 'right': return '"';
        case 'center': return '^';
        case 'fill': return '\\';
        default: return "'";
      }
    }
    case 'protect': return style.locked === false ? 0 : 1;
    case 'type': {
      const v = ctx.cell(sheet, r, c);
      if (v == null && cell?.f == null) return 'b';
      return typeof v === 'string' ? 'l' : 'v';
    }
    case 'width': {
      const px = sheet.colW.get(c);
      const w = Math.trunc(((px ?? sheet.defaultColW) - 5) / 7);
      return [[Math.max(0, w), px == null]];
    }
  }
  return ERR.value;
}

function typeCode(v: Value, ctx: Ctx): number {
  if (isLambda(v)) return 128;
  if (isRefs(v)) return 16;
  if (isMatrix(v)) return v.length === 1 && v[0].length === 1 ? typeCode(v[0][0], ctx) : 64;
  if (isRef(v)) return v.r1 === v.r2 && v.c1 === v.c2 && !v.sheets ? typeCode(ctx.cell(v.sheet, v.r1, v.c1), ctx) : 64;
  if (v == null || typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 4;
  if (isErr(v)) return 16;
  return 1;
}

defFn([
  {
    name: 'CELL', min: 1, max: 2, category: 'Information', desc: 'Returns information about the formatting, location or contents of a cell.', params: ['info_type', '[reference]'], refs: [1],
    impl: ([t, ref], ctx) => {
      const type = str(S(t)).toLowerCase();
      if (ref !== undefined && ref !== null && !isRef(ref) && !isRefs(ref)) return ERR.value;
      return cellInfo(type, refArg(ref), ctx);
    },
  },
  {
    name: 'ERROR.TYPE', min: 1, max: 1, category: 'Information', desc: 'Returns a number identifying an error value.', params: ['error_val'],
    impl: ([v]) => { const x = S(v); return isErr(x) ? ERROR_NUMBERS[x.code] ?? ERR.na : ERR.na; },
  },
  {
    name: 'INFO', min: 1, max: 1, category: 'Information', desc: 'Returns information about the current operating environment.', params: ['type_text'],
    impl: ([t], ctx) => {
      switch (str(S(t)).toLowerCase()) {
        case 'directory': return '/';
        case 'numfile': return ctx.book.sheets.length;
        case 'origin': {
          const sh = ctx.sheet;
          return `$A:$${colName(sh.scroll?.c ?? 0)}$${(sh.scroll?.r ?? 0) + 1}`;
        }
        case 'osversion': return typeof navigator !== 'undefined' && navigator.userAgent ? navigator.platform || 'Web' : 'Web';
        case 'recalc': return ctx.book.calc.mode === 'manual' ? 'Manual' : 'Automatic';
        case 'release': return '16.0';
        case 'system': return 'pcdos';
      }
      return ERR.value;
    },
  },
  { name: 'ISBLANK', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value refers to an empty cell.', params: ['value'], impl: ([v]) => S(v) === null },
  { name: 'ISERR', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is any error except #N/A.', params: ['value'], impl: ([v]) => { const x = S(v); return isErr(x) && x.code !== '#N/A'; } },
  { name: 'ISERROR', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is any error.', params: ['value'], impl: ([v]) => isErr(S(v)) },
  { name: 'ISNA', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is the #N/A error.', params: ['value'], impl: ([v]) => { const x = S(v); return isErr(x) && x.code === '#N/A'; } },
  { name: 'ISLOGICAL', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is a logical value.', params: ['value'], impl: ([v]) => typeof S(v) === 'boolean' },
  { name: 'ISNUMBER', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is a number.', params: ['value'], impl: ([v]) => typeof S(v) === 'number' },
  { name: 'ISTEXT', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is text.', params: ['value'], impl: ([v]) => typeof S(v) === 'string' },
  { name: 'ISNONTEXT', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is not text (blank cells count as non-text).', params: ['value'], impl: ([v]) => typeof S(v) !== 'string' },
  {
    name: 'ISEVEN', min: 1, max: 1, category: 'Information', desc: 'TRUE if the number is even.', params: ['number'],
    impl: ([v]) => { const x = S(v); if (typeof x === 'boolean') return ERR.value; return Math.trunc(num(x)) % 2 === 0; },
  },
  {
    name: 'ISODD', min: 1, max: 1, category: 'Information', desc: 'TRUE if the number is odd.', params: ['number'],
    impl: ([v]) => { const x = S(v); if (typeof x === 'boolean') return ERR.value; return Math.abs(Math.trunc(num(x)) % 2) === 1; },
  },
  {
    name: 'ISREF', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is a reference.', params: ['value'], refs: [0],
    impl: ([v]) => isRef(v) || isRefs(v),
  },
  {
    name: 'ISFORMULA', min: 1, max: 1, category: 'Information', desc: 'TRUE if the referenced cell contains a formula.', params: ['reference'], refs: [0],
    impl: ([v]) => {
      if (isRefs(v)) v = v.areas[0];
      if (!isRef(v)) return isErr(v as Scalar) ? (v as Scalar) : ERR.value;
      if (v.r1 === v.r2 && v.c1 === v.c2) return v.sheet.get(v.r1, v.c1)?.f != null;
      const out: Scalar[][] = [];
      for (let r = v.r1; r <= v.r2; r++) { const row: Scalar[] = []; for (let c = v.c1; c <= v.c2; c++) row.push(v.sheet.get(r, c)?.f != null); out.push(row); }
      return out;
    },
  },
  {
    name: 'N', min: 1, max: 1, category: 'Information', desc: 'Converts a value to a number (text becomes 0, TRUE becomes 1).', params: ['value'],
    impl: ([v]) => { const x = S(v); if (isErr(x)) return x; if (typeof x === 'number') return x; if (typeof x === 'boolean') return x ? 1 : 0; return 0; },
  },
  { name: 'NA', min: 0, max: 0, category: 'Information', desc: 'Returns the #N/A error value.', params: [], impl: () => ERR.na },
  {
    name: 'SHEET', min: 0, max: 1, category: 'Information', desc: 'Sheet number of a reference or sheet name.', params: ['[value]'], refs: [0],
    impl: ([v], ctx) => {
      const sheets = ctx.book.sheets;
      if (v === undefined || v === null) return sheets.indexOf(ctx.sheet) + 1;
      if (isRef(v)) return sheets.indexOf(v.sheets ? v.sheets[0] : v.sheet) + 1;
      if (isRefs(v)) return sheets.indexOf(v.areas[0].sheet) + 1;
      const x = isMatrix(v) ? toScalar(v, ctx) : S(v);
      if (isErr(x)) return x;
      if (typeof x === 'string') { const i = sheets.findIndex((s) => s.name.toLowerCase() === x.toLowerCase()); return i < 0 ? ERR.na : i + 1; }
      return ERR.na;
    },
  },
  {
    name: 'SHEETS', min: 0, max: 1, category: 'Information', desc: 'Number of sheets in a reference, or in the workbook.', params: ['[reference]'], refs: [0],
    impl: ([v], ctx) => {
      if (v === undefined || v === null) return ctx.book.sheets.length;
      if (isRef(v)) return v.sheets ? v.sheets.length : 1;
      if (isRefs(v)) return new Set(v.areas.map((a) => a.sheet)).size;
      if (isErr(v as Scalar)) return v;
      return ERR.na;
    },
  },
  {
    name: 'TYPE', min: 1, max: 1, category: 'Information', desc: 'Returns a number for the type of a value (1 number, 2 text, 4 logical, 16 error, 64 array).', params: ['value'], arrays: [0],
    impl: ([v], ctx) => typeCode(v ?? null, ctx),
  },
]);

