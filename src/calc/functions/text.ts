/** Text functions. */
import { defFn } from './registry';
import {
  fail, isErr, isLambda, isMatrix, isRef, isRefs, iterValues, toMatrix, toScalar, num, str, int, bool, parseNumeric, wildcard,
  ERR, type Ctx, type Value,
} from '../values';
import type { Matrix, Scalar } from '../../model/types';
import { formatValue, isDateFormat } from '../../model/numfmt';

const MAX_TEXT = 32767;

/** Scalar from an argument that may be a reference or array (array arguments). */
function sc(v: Value | undefined, ctx: Ctx): Scalar | undefined {
  if (v === undefined) return undefined;
  if (isRef(v) || isMatrix(v)) return toScalar(v, ctx);
  if (isRefs(v)) return fail(ERR.value);
  if (isLambda(v)) return fail(ERR.calc);
  return v;
}

/** Optional numeric argument: undefined (omitted) uses the default; blank counts as 0. */
function optInt(v: Value | undefined, def: number): number { return v === undefined ? def : int(v as Scalar); }
/** Optional numeric argument where an empty argument also uses the default. */
function optIntE(v: Scalar | undefined, def: number): number { return v === undefined || v === null ? def : int(v); }

function checkLen(s: string): string { if (s.length > MAX_TEXT) fail(ERR.value); return s; }

/** Value as text for concatenation (errors propagate). */
const txt = (v: Scalar) => str(v);

// --------------------------------------------------------------------------------------------
// Code pages

const CP1252: Record<number, number> = {
  0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160, 0x8b: 0x2039,
  0x8c: 0x0152, 0x8e: 0x017d, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02dc, 0x99: 0x2122,
  0x9a: 0x0161, 0x9b: 0x203a, 0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178,
};
const CP1252_REV = new Map<number, number>(Object.entries(CP1252).map(([k, v]) => [v, +k]));

function toHalfWidth(s: string): string {
  let out = '';
  for (const ch of s) {
    const c = ch.codePointAt(0)!;
    if (c >= 0xff01 && c <= 0xff5e) out += String.fromCharCode(c - 0xfee0);
    else if (c === 0x3000) out += ' ';
    else if (HALF_KANA.has(ch)) out += HALF_KANA.get(ch);
    else out += ch;
  }
  return out;
}

function toFullWidth(s: string): string {
  let out = '';
  const chars = [...s];
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    const c = ch.codePointAt(0)!;
    if (c >= 0x21 && c <= 0x7e) out += String.fromCharCode(c + 0xfee0);
    else if (c === 0x20) out += '　';
    else if (FULL_KANA.has(ch)) {
      // Combine voiced/semi-voiced marks with the preceding kana.
      const nx = chars[i + 1];
      const base = FULL_KANA.get(ch)!;
      if (nx === 'ﾞ' && VOICED.has(base)) { out += VOICED.get(base); i++; }
      else if (nx === 'ﾟ' && SEMI.has(base)) { out += SEMI.get(base); i++; }
      else out += base;
    } else out += ch;
  }
  return out;
}

// Half-width katakana (U+FF61..U+FF9F) and their full-width forms.
const HW = '｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ';
const FW = '。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜';
const FULL_KANA = new Map<string, string>();
const HALF_KANA = new Map<string, string>();
for (let i = 0; i < HW.length; i++) { FULL_KANA.set(HW[i], FW[i]); HALF_KANA.set(FW[i], HW[i]); }
const VOICED = new Map<string, string>();
const SEMI = new Map<string, string>();
{
  const v1 = 'カキクケコサシスセソタチツテトハヒフヘホウ', v2 = 'ガギグゲゴザジズゼゾダヂヅデドバビブベボヴ';
  for (let i = 0; i < v1.length; i++) { VOICED.set(v1[i], v2[i]); HALF_KANA.set(v2[i], HALF_KANA.get(v1[i])! + 'ﾞ'); }
  const s1 = 'ハヒフヘホ', s2 = 'パピプペポ';
  for (let i = 0; i < s1.length; i++) { SEMI.set(s1[i], s2[i]); HALF_KANA.set(s2[i], HALF_KANA.get(s1[i])! + 'ﾟ'); }
}

// --------------------------------------------------------------------------------------------
// Number formatting helpers

function roundDec(x: number, d: number): number {
  const m = Math.pow(10, d);
  const r = Math.round(+(Math.abs(x) * m).toPrecision(15)) / m;
  return x < 0 ? -r : r;
}

function fixed(x: number, decimals: number, commas: boolean): { s: string; neg: boolean } {
  const d = Math.min(decimals, 127);
  const r = roundDec(x, d);
  let s = Math.abs(r).toFixed(Math.max(0, Math.min(d, 100)));
  if (commas) { const [ip, fp] = s.split('.'); s = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (fp != null ? '.' + fp : ''); }
  return { s, neg: r < 0 && /[1-9]/.test(s) };
}

/** In date/time formats a comma is literal text; escape it so the formatter does not treat it as grouping. */
function literalCommas(fmt: string): string {
  let out = '', inQ = false, inB = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '\\' && !inQ) { out += ch + (fmt[i + 1] ?? ''); i++; continue; }
    if (ch === '"' && !inB) inQ = !inQ;
    else if (ch === '[' && !inQ) inB = true;
    else if (ch === ']' && !inQ) inB = false;
    out += ch === ',' && !inQ && !inB ? '\\,' : ch;
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// BAHTTEXT

const TH_DIGITS = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
const TH_PLACES = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];

/** Thai words for a non-negative integer (no "zero" output for 0). */
function thaiInt(n: number): string {
  if (n === 0) return '';
  if (n >= 1e6) {
    const hi = Math.floor(n / 1e6), lo = n % 1e6;
    return thaiInt(hi) + 'ล้าน' + thaiInt(lo);
  }
  const digits = String(n).split('').map(Number);
  let out = '';
  const len = digits.length;
  for (let i = 0; i < len; i++) {
    const d = digits[i], place = len - 1 - i;
    if (d === 0) continue;
    if (place === 1 && d === 1) out += 'สิบ';
    else if (place === 1 && d === 2) out += 'ยี่สิบ';
    else if (place === 0 && d === 1 && len > 1) out += 'เอ็ด';
    else out += TH_DIGITS[d] + TH_PLACES[place];
  }
  return out;
}

/** Megadigit-aware: for numbers >= 1e6, units of the lower part with a 1 use "เอ็ด" too. */
function thaiNumber(n: number): string {
  if (n >= 1e6) {
    const hi = Math.floor(n / 1e6), lo = n % 1e6;
    let low = thaiInt(lo);
    if (lo === 1) low = 'เอ็ด';
    return thaiNumber(hi) + 'ล้าน' + low;
  }
  return thaiInt(n);
}

function bahtText(x: number): string {
  const neg = x < 0;
  const a = Math.abs(x);
  let baht = Math.floor(a);
  let satang = Math.round((a - baht) * 100 + 1e-9);
  if (satang >= 100) { baht += 1; satang -= 100; }
  let out = '';
  if (baht === 0 && satang === 0) return 'ศูนย์บาทถ้วน';
  if (baht > 0) out += thaiNumber(baht) + 'บาท';
  if (satang > 0) out += thaiNumber(satang) + 'สตางค์';
  else out += 'ถ้วน';
  return (neg ? 'ลบ' : '') + out;
}

// --------------------------------------------------------------------------------------------
// Searching helpers

function searchText(find: string, within: string, start: number, ci: boolean, wild: boolean): number {
  if (start < 1 || start > within.length + 1) return fail(ERR.value);
  if (find === '') return start;
  if (wild && /[*?~]/.test(find)) {
    const re = new RegExp(wildcard(find, false).source, 'gi');
    re.lastIndex = start - 1;
    const m = re.exec(within);
    if (!m) return fail(ERR.value);
    return m.index + 1;
  }
  const i = ci ? within.toLowerCase().indexOf(find.toLowerCase(), start - 1) : within.indexOf(find, start - 1);
  if (i < 0) return fail(ERR.value);
  return i + 1;
}

/** Delimiter list from a scalar or array argument. */
function delimList(v: Value | undefined, ctx: Ctx): string[] | null {
  if (v === undefined || v === null) return null;
  const out: string[] = [];
  for (const { v: x } of iterValues(v, ctx)) {
    if (isErr(x)) fail(x);
    out.push(str(x));
  }
  return out;
}

interface Hit { pos: number; len: number }

function findDelims(text: string, delims: string[], ci: boolean): Hit[] {
  const hits: Hit[] = [];
  const t = ci ? text.toLowerCase() : text;
  const ds = delims.filter((d) => d !== '').map((d) => (ci ? d.toLowerCase() : d)).sort((a, b) => b.length - a.length);
  if (!ds.length) return hits;
  let i = 0;
  while (i < t.length) {
    const d = ds.find((x) => t.startsWith(x, i));
    if (d) { hits.push({ pos: i, len: d.length }); i += d.length; } else i++;
  }
  return hits;
}

function textAround(args: Value[], ctx: Ctx, after: boolean): Value {
  const text = str(sc(args[0], ctx) as Scalar);
  const delims = delimList(args[1], ctx);
  if (!delims || !delims.length) return ERR.value;
  const inst = optIntE(sc(args[2], ctx), 1);
  const ci = optIntE(sc(args[3], ctx), 0);
  const matchEnd = optIntE(sc(args[4], ctx), 0);
  const notFound = args[5];
  if (ci !== 0 && ci !== 1) return ERR.value;
  if (inst === 0 || Math.abs(inst) > Math.max(text.length, 1)) return ERR.value;
  const emptyDelim = delims.some((d) => d === '');
  let hits = findDelims(text, delims, ci === 1);
  if (emptyDelim) hits = inst > 0 ? [{ pos: 0, len: 0 }, ...hits] : [...hits, { pos: text.length, len: 0 }];
  if (matchEnd) {
    if (inst > 0) hits.push({ pos: text.length, len: 0 });
    else hits.unshift({ pos: 0, len: 0 });
  }
  const h = inst > 0 ? hits[inst - 1] : hits[hits.length + inst];
  if (!h) {
    if (notFound === undefined || notFound === null) return ERR.na;
    return notFound;
  }
  return after ? text.slice(h.pos + h.len) : text.slice(0, h.pos);
}

function regexFrom(pattern: string, ci: number, global = false): RegExp {
  if (ci !== 0 && ci !== 1) fail(ERR.value);
  let src = pattern;
  let flags = (global ? 'g' : '') + (ci ? 'i' : '');
  // PCRE inline flag (?i) at the start.
  if (/^\(\?i\)/.test(src)) { src = src.slice(4); if (!flags.includes('i')) flags += 'i'; }
  try { return new RegExp(src, flags + 'u'); } catch { /* fall through */ }
  try { return new RegExp(src, flags); } catch { return fail(ERR.value); }
}

// --------------------------------------------------------------------------------------------
// VALUE / NUMBERVALUE

function valueOf(v: Scalar): number {
  if (typeof v === 'number') return v;
  if (v == null) return 0;
  if (isErr(v)) return fail(v);
  if (typeof v === 'boolean') return fail(ERR.value);
  const n = parseNumeric(v);
  if (n == null) return fail(ERR.value);
  return n;
}

function numberValue(text: string, dec: string, grp: string): number {
  if (dec === grp) return fail(ERR.value);
  let t = text.replace(/\s+/g, '');
  if (t === '') return 0;
  let pct = 0;
  while (t.endsWith('%')) { pct++; t = t.slice(0, -1); }
  let neg = false;
  if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
  if (t.startsWith('-')) { neg = !neg; t = t.slice(1); } else if (t.startsWith('+')) t = t.slice(1);
  const di = t.indexOf(dec);
  if (di >= 0 && t.indexOf(dec, di + 1) >= 0) return fail(ERR.value);
  let ip = di >= 0 ? t.slice(0, di) : t, fp = di >= 0 ? t.slice(di + dec.length) : '';
  if (grp && fp.includes(grp)) return fail(ERR.value);
  if (grp) ip = ip.split(grp).join('');
  let exp = '';
  const em = /[eE][+-]?\d+$/.exec(fp || ip);
  if (em) { if (fp) fp = fp.slice(0, -em[0].length); else ip = ip.slice(0, -em[0].length); exp = em[0]; }
  if (!/^\d*$/.test(ip) || !/^\d*$/.test(fp) || (ip === '' && fp === '')) return fail(ERR.value);
  const n = Number(`${ip || '0'}.${fp || '0'}${exp}`);
  if (!Number.isFinite(n)) return fail(ERR.value);
  return (neg ? -n : n) / Math.pow(100, pct);
}

// --------------------------------------------------------------------------------------------
// ARRAYTOTEXT / VALUETOTEXT

function valueText(v: Scalar, strict: boolean): string {
  if (v == null) return '';
  if (isErr(v)) return v.code;
  if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v;
  return str(v);
}

function fmtArg(v: Scalar | undefined): boolean {
  const f = v === undefined || v === null ? 0 : int(v);
  if (f !== 0 && f !== 1) fail(ERR.value);
  return f === 1;
}

// --------------------------------------------------------------------------------------------

const S = (a: Value) => a as Scalar;

defFn([
  {
    name: 'ARRAYTOTEXT', min: 1, max: 2, category: 'Text', desc: 'Converts an array or range to a single text value.', params: ['array', '[format]'], arrays: [0],
    impl: ([a, f], ctx) => {
      const strict = fmtArg(f as Scalar | undefined);
      const m = toMatrix(a, ctx);
      if (strict) return checkLen('{' + m.map((row) => row.map((x) => valueText(x, true)).join(',')).join(';') + '}');
      return checkLen(m.flatMap((row) => row.map((x) => valueText(x, false))).join(', '));
    },
  },
  {
    name: 'VALUETOTEXT', min: 1, max: 2, category: 'Text', desc: 'Returns a value as text, optionally quoting text in strict format.', params: ['value', '[format]'],
    impl: ([v, f]) => valueText(S(v), fmtArg(f as Scalar | undefined)),
  },
  { name: 'ASC', min: 1, max: 1, category: 'Text', desc: 'Converts full-width (double-byte) characters to half-width (single-byte) characters.', params: ['text'], impl: ([t]) => toHalfWidth(str(S(t))) },
  { name: 'DBCS', min: 1, max: 1, category: 'Text', desc: 'Converts half-width (single-byte) characters to full-width (double-byte) characters.', params: ['text'], aliases: ['JIS'], impl: ([t]) => toFullWidth(str(S(t))) },
  { name: 'BAHTTEXT', min: 1, max: 1, category: 'Text', desc: 'Spells out a number in Thai words with the baht currency suffix.', params: ['number'], impl: ([n]) => bahtText(num(S(n))) },
  {
    name: 'CHAR', min: 1, max: 1, category: 'Text', desc: 'Returns the character for a code number (1-255) in the system character set.', params: ['number'],
    impl: ([n]) => { const c = int(S(n)); if (c < 1 || c > 255) return ERR.value; return String.fromCharCode(CP1252[c] ?? c); },
  },
  {
    name: 'CODE', min: 1, max: 1, category: 'Text', desc: 'Returns the character-set code of the first character of a text value.', params: ['text'],
    impl: ([t]) => {
      const s = str(S(t));
      if (!s) return ERR.value;
      const c = s.charCodeAt(0);
      if (c < 0x80 || (c >= 0xa0 && c <= 0xff)) return c;
      if (c <= 0x9f) return c;
      return CP1252_REV.get(c) ?? 63;
    },
  },
  { name: 'CLEAN', min: 1, max: 1, category: 'Text', desc: 'Removes non-printable control characters (codes 0-31) from text.', params: ['text'], impl: ([t]) => str(S(t)).replace(/[\x00-\x1f]/g, '') },
  {
    name: 'CONCAT', min: 1, max: 253, category: 'Text', desc: 'Joins the text of values, ranges and arrays.', params: ['text1', '[text2]', '...'], arrays: 'all',
    impl: (args, ctx) => {
      let s = '';
      for (const a of args) for (const { v } of iterValues(a ?? null, ctx)) { s += txt(v); if (s.length > MAX_TEXT) return ERR.value; }
      return s;
    },
  },
  {
    name: 'CONCATENATE', min: 1, max: 255, category: 'Text', desc: 'Joins several text values into one.', params: ['text1', '[text2]', '...'],
    impl: (args) => checkLen(args.map((a) => txt(S(a))).join('')),
  },
  {
    name: 'DOLLAR', min: 1, max: 2, category: 'Text', desc: 'Formats a number as currency text with the given decimals.', params: ['number', '[decimals]'], aliases: ['USDOLLAR'],
    impl: ([n, d]) => {
      const dec = optInt(d, 2);
      if (dec > 127) return ERR.value;
      const { s, neg } = fixed(num(S(n)), dec, true);
      return neg ? `($${s})` : `$${s}`;
    },
  },
  {
    name: 'FIXED', min: 1, max: 3, category: 'Text', desc: 'Rounds a number and returns it as text, with or without thousands separators.', params: ['number', '[decimals]', '[no_commas]'],
    impl: ([n, d, nc]) => {
      const dec = optInt(d, 2);
      if (dec > 127) return ERR.value;
      const { s, neg } = fixed(num(S(n)), dec, !(nc === undefined ? false : bool(S(nc))));
      return (neg ? '-' : '') + s;
    },
  },
  { name: 'EXACT', min: 2, max: 2, category: 'Text', desc: 'TRUE if two text values are identical, including case.', params: ['text1', 'text2'], impl: ([a, b]) => str(S(a)) === str(S(b)) },
  {
    name: 'FIND', min: 2, max: 3, category: 'Text', desc: 'Position of one text inside another (case-sensitive, no wildcards).', params: ['find_text', 'within_text', '[start_num]'],
    impl: ([f, w, s]) => searchText(str(S(f)), str(S(w)), optInt(s, 1), false, false),
  },
  {
    name: 'FINDB', min: 2, max: 3, category: 'Text', desc: 'Like FIND, counting bytes in double-byte locales.', params: ['find_text', 'within_text', '[start_num]'],
    impl: ([f, w, s]) => searchText(str(S(f)), str(S(w)), optInt(s, 1), false, false),
  },
  {
    name: 'SEARCH', min: 2, max: 3, category: 'Text', desc: 'Position of one text inside another (not case-sensitive, wildcards allowed).', params: ['find_text', 'within_text', '[start_num]'],
    impl: ([f, w, s]) => searchText(str(S(f)), str(S(w)), optInt(s, 1), true, true),
  },
  {
    name: 'SEARCHB', min: 2, max: 3, category: 'Text', desc: 'Like SEARCH, counting bytes in double-byte locales.', params: ['find_text', 'within_text', '[start_num]'],
    impl: ([f, w, s]) => searchText(str(S(f)), str(S(w)), optInt(s, 1), true, true),
  },
  ...(['LEFT', 'LEFTB'] as const).map((name) => ({
    name, min: 1, max: 2, category: 'Text' as const, desc: name === 'LEFT' ? 'Returns the first characters of a text value.' : 'Returns the first characters of a text value (bytes in double-byte locales).', params: ['text', name === 'LEFT' ? '[num_chars]' : '[num_bytes]'],
    impl: ([t, n]: Value[]) => { const k = optInt(n, 1); if (k < 0) return ERR.value; return str(S(t)).slice(0, k); },
  })),
  ...(['RIGHT', 'RIGHTB'] as const).map((name) => ({
    name, min: 1, max: 2, category: 'Text' as const, desc: name === 'RIGHT' ? 'Returns the last characters of a text value.' : 'Returns the last characters of a text value (bytes in double-byte locales).', params: ['text', name === 'RIGHT' ? '[num_chars]' : '[num_bytes]'],
    impl: ([t, n]: Value[]) => { const k = optInt(n, 1); if (k < 0) return ERR.value; const s = str(S(t)); return k === 0 ? '' : s.slice(Math.max(0, s.length - k)); },
  })),
  ...(['MID', 'MIDB'] as const).map((name) => ({
    name, min: 3, max: 3, category: 'Text' as const, desc: name === 'MID' ? 'Returns characters from the middle of a text value.' : 'Returns characters from the middle of a text value (bytes in double-byte locales).', params: ['text', 'start_num', name === 'MID' ? 'num_chars' : 'num_bytes'],
    impl: ([t, s, n]: Value[]) => {
      const st = int(S(s)), k = int(S(n));
      if (st < 1 || k < 0) return ERR.value;
      return str(S(t)).substr(st - 1, k);
    },
  })),
  ...(['LEN', 'LENB'] as const).map((name) => ({
    name, min: 1, max: 1, category: 'Text' as const, desc: name === 'LEN' ? 'Number of characters in a text value.' : 'Number of bytes in a text value (double-byte locales).', params: ['text'],
    impl: ([t]: Value[]) => str(S(t)).length,
  })),
  { name: 'LOWER', min: 1, max: 1, category: 'Text', desc: 'Converts text to lowercase.', params: ['text'], impl: ([t]) => str(S(t)).toLowerCase() },
  { name: 'UPPER', min: 1, max: 1, category: 'Text', desc: 'Converts text to uppercase.', params: ['text'], impl: ([t]) => str(S(t)).toUpperCase() },
  {
    name: 'PROPER', min: 1, max: 1, category: 'Text', desc: 'Capitalizes the first letter of each word and lowercases the rest.', params: ['text'],
    impl: ([t]) => {
      const s = str(S(t));
      let out = '', prevLetter = false;
      for (const ch of s) {
        const isL = /\p{L}/u.test(ch);
        out += isL ? (prevLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
        prevLetter = isL;
      }
      return out;
    },
  },
  {
    name: 'NUMBERVALUE', min: 1, max: 3, category: 'Text', desc: 'Converts text to a number using the given decimal and group separators.', params: ['text', '[decimal_separator]', '[group_separator]'],
    impl: ([t, d, g]) => {
      const v = S(t);
      if (typeof v === 'number') return v;
      if (isErr(v)) return v;
      const dec = d === undefined || d === null ? '.' : str(S(d)).charAt(0);
      const grp = g === undefined || g === null ? ',' : str(S(g)).charAt(0);
      if ((d !== undefined && d !== null && !dec) || (g !== undefined && g !== null && !grp)) return ERR.value;
      return numberValue(str(v), dec, grp);
    },
  },
  {
    name: 'PHONETIC', min: 1, max: 1, category: 'Text', desc: 'Returns the phonetic (furigana) text of a reference; without furigana, the text of the cells.', params: ['reference'], refs: [0],
    impl: ([r], ctx) => {
      if (!isRef(r) && !isRefs(r)) return isMatrix(r) ? ERR.value : typeof r === 'string' ? r : ERR.na;
      let s = '';
      for (const { v } of iterValues(r, ctx)) { if (isErr(v)) return v; if (typeof v === 'string') s += v; }
      return s;
    },
  },
  {
    name: 'REGEXTEST', min: 2, max: 3, category: 'Text', desc: 'TRUE if any part of the text matches a regular expression.', params: ['text', 'pattern', '[case_sensitivity]'],
    impl: ([t, p, c]) => regexFrom(str(S(p)), optIntE(S(c), 0)).test(str(S(t))),
  },
  {
    name: 'REGEXEXTRACT', min: 2, max: 4, category: 'Text', desc: 'Extracts text matching a regular expression: the first match, all matches, or the capture groups.', params: ['text', 'pattern', '[return_mode]', '[case_sensitivity]'],
    impl: ([t, p, m, c]) => {
      const s = str(S(t));
      const mode = optIntE(S(m), 0);
      if (mode < 0 || mode > 2) return ERR.value;
      const re = regexFrom(str(S(p)), optIntE(S(c), 0), mode === 1);
      if (mode === 0) { const x = re.exec(s); return x ? x[0] : ERR.na; }
      if (mode === 1) {
        const all: Matrix = [];
        let x: RegExpExecArray | null;
        while ((x = re.exec(s))) { all.push([x[0]]); if (x[0] === '') re.lastIndex++; }
        return all.length ? all : ERR.na;
      }
      const x = re.exec(s);
      if (!x) return ERR.na;
      if (x.length === 1) return [[x[0]]];
      return [x.slice(1).map((g) => g ?? '')];
    },
  },
  {
    name: 'REGEXREPLACE', min: 3, max: 5, category: 'Text', desc: 'Replaces text matching a regular expression (all occurrences or the nth one).', params: ['text', 'pattern', 'replacement', '[occurrence]', '[case_sensitivity]'],
    impl: ([t, p, r, o, c]) => {
      const s = str(S(t)), rep = str(S(r));
      const occ = optIntE(S(o), 0);
      const re = regexFrom(str(S(p)), optIntE(S(c), 0), true);
      if (occ === 0) return checkLen(s.replace(re, rep));
      const matches: RegExpExecArray[] = [];
      let x: RegExpExecArray | null;
      while ((x = re.exec(s))) { matches.push(x); if (x[0] === '') re.lastIndex++; }
      const k = occ > 0 ? occ - 1 : matches.length + occ;
      const hit = matches[k];
      if (!hit) return s;
      const one = new RegExp(re.source, re.flags.replace('g', ''));
      const replaced = hit[0].replace(one, rep);
      return checkLen(s.slice(0, hit.index) + replaced + s.slice(hit.index + hit[0].length));
    },
  },
  ...(['REPLACE', 'REPLACEB'] as const).map((name) => ({
    name, min: 4, max: 4, category: 'Text' as const, desc: name === 'REPLACE' ? 'Replaces part of a text value, by position, with other text.' : 'Replaces part of a text value by byte position (double-byte locales).', params: ['old_text', 'start_num', name === 'REPLACE' ? 'num_chars' : 'num_bytes', 'new_text'],
    impl: ([t, s, n, nt]: Value[]) => {
      const st = int(S(s)), k = int(S(n));
      if (st < 1 || k < 0) return ERR.value;
      const x = str(S(t));
      return checkLen(x.slice(0, st - 1) + str(S(nt)) + x.slice(st - 1 + k));
    },
  })),
  {
    name: 'REPT', min: 2, max: 2, category: 'Text', desc: 'Repeats text a given number of times.', params: ['text', 'number_times'],
    impl: ([t, n]) => {
      const k = int(S(n)), s = str(S(t));
      if (k < 0 || s.length * k > MAX_TEXT) return ERR.value;
      return s.repeat(k);
    },
  },
  {
    name: 'SUBSTITUTE', min: 3, max: 4, category: 'Text', desc: 'Replaces occurrences of old text with new text (all, or a specific instance).', params: ['text', 'old_text', 'new_text', '[instance_num]'],
    impl: ([t, o, n, i]) => {
      const s = str(S(t)), old = str(S(o)), nw = str(S(n));
      if (i !== undefined) {
        const k = int(S(i));
        if (k < 1) return ERR.value;
        if (!old) return s;
        let pos = -1;
        for (let j = 0; j < k; j++) { pos = s.indexOf(old, j === 0 ? 0 : pos + old.length); if (pos < 0) return s; }
        return checkLen(s.slice(0, pos) + nw + s.slice(pos + old.length));
      }
      if (!old) return s;
      return checkLen(s.split(old).join(nw));
    },
  },
  {
    name: 'T', min: 1, max: 1, category: 'Text', desc: 'Returns the value if it is text, otherwise empty text.', params: ['value'],
    impl: ([v]) => { const x = S(v); if (isErr(x)) return x; return typeof x === 'string' ? x : ''; },
  },
  {
    name: 'TEXT', min: 2, max: 2, category: 'Text', desc: 'Formats a number as text using a number format code.', params: ['value', 'format_text'],
    impl: ([v, f], ctx) => {
      let x = S(v);
      if (isErr(x)) return x;
      const fmt = str(S(f));
      if (typeof x === 'string') { const n = parseNumeric(x); if (n != null) x = n; }
      if (x == null) x = 0;
      if (fmt === '') return '';
      return formatValue(x, isDateFormat(fmt) ? literalCommas(fmt) : fmt, { date1904: ctx.book.calc.date1904 }).text;
    },
  },
  {
    name: 'TEXTAFTER', min: 2, max: 6, category: 'Text', desc: 'Returns the text that follows a delimiter.', params: ['text', 'delimiter', '[instance_num]', '[match_mode]', '[match_end]', '[if_not_found]'], arrays: [1, 5],
    impl: (args, ctx) => textAround(args, ctx, true),
  },
  {
    name: 'TEXTBEFORE', min: 2, max: 6, category: 'Text', desc: 'Returns the text that precedes a delimiter.', params: ['text', 'delimiter', '[instance_num]', '[match_mode]', '[match_end]', '[if_not_found]'], arrays: [1, 5],
    impl: (args, ctx) => textAround(args, ctx, false),
  },
  {
    name: 'TEXTJOIN', min: 3, max: 252, category: 'Text', desc: 'Joins values from ranges and arrays with a delimiter, optionally skipping empty ones.', params: ['delimiter', 'ignore_empty', 'text1', '[text2]', '...'], arrays: 'all',
    impl: (args, ctx) => {
      const delims = delimList(args[0], ctx) ?? [''];
      if (!delims.length) delims.push('');
      const ign = bool(sc(args[1], ctx) ?? false);
      const parts: string[] = [];
      for (const a of args.slice(2)) for (const { v } of iterValues(a ?? null, ctx)) {
        if (isErr(v)) return v;
        const s = str(v);
        if (ign && s === '') continue;
        parts.push(s);
      }
      let out = '';
      parts.forEach((p, i) => { if (i) out += delims[(i - 1) % delims.length]; out += p; });
      return checkLen(out);
    },
  },
  {
    name: 'TEXTSPLIT', min: 2, max: 6, category: 'Text', desc: 'Splits text into rows and columns at the given delimiters.', params: ['text', 'col_delimiter', '[row_delimiter]', '[ignore_empty]', '[match_mode]', '[pad_with]'], arrays: [1, 2],
    impl: ([t, cd, rd, ie, mm, pad], ctx) => {
      const s = str(S(t));
      const cols = delimList(cd, ctx), rows = delimList(rd, ctx);
      if (!cols && !rows) return ERR.value;
      if ((cols && cols.some((d) => d === '')) || (rows && rows.some((d) => d === ''))) return ERR.value;
      const ign = ie === undefined || ie === null ? false : bool(S(ie));
      const ci = optIntE(S(mm), 0);
      if (ci !== 0 && ci !== 1) return ERR.value;
      const padV: Scalar = pad === undefined ? ERR.na : S(pad);
      const split = (x: string, ds: string[] | null): string[] => {
        if (!ds || !ds.length) return [x];
        const hits = findDelims(x, ds, ci === 1);
        const out: string[] = [];
        let p = 0;
        for (const h of hits) { out.push(x.slice(p, h.pos)); p = h.pos + h.len; }
        out.push(x.slice(p));
        return out;
      };
      let lines = split(s, rows);
      if (ign) lines = lines.filter((l) => l !== '');
      let m = lines.map((l) => { let c = split(l, cols); if (ign) c = c.filter((x) => x !== ''); return c; });
      if (ign) m = m.filter((r) => r.length);
      if (!m.length) return ERR.calc;
      const w = Math.max(...m.map((r) => r.length));
      return m.map((r) => { const row: Scalar[] = [...r]; while (row.length < w) row.push(padV); return row; });
    },
  },
  {
    name: 'TRIM', min: 1, max: 1, category: 'Text', desc: 'Removes leading and trailing spaces and reduces internal runs of spaces to one.', params: ['text'],
    impl: ([t]) => str(S(t)).replace(/ +/g, ' ').replace(/^ | $/g, ''),
  },
  {
    name: 'UNICHAR', min: 1, max: 1, category: 'Text', desc: 'Returns the Unicode character for a code point.', params: ['number'],
    impl: ([n]) => {
      const c = int(S(n));
      if (c < 1 || c > 0x10ffff) return ERR.value;
      if (c >= 0xd800 && c <= 0xdfff) return ERR.na;
      return String.fromCodePoint(c);
    },
  },
  {
    name: 'UNICODE', min: 1, max: 1, category: 'Text', desc: 'Returns the Unicode code point of the first character.', params: ['text'],
    impl: ([t]) => { const s = str(S(t)); if (!s) return ERR.value; return s.codePointAt(0)!; },
  },
  { name: 'VALUE', min: 1, max: 1, category: 'Text', desc: 'Converts text that looks like a number, date or time to a number.', params: ['text'], impl: ([t]) => valueOf(S(t)) },
]);

