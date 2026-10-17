/**
 * Worksheet functions carried over from other spreadsheets, past and present: Lotus 1-2-3 (@ISBETWEEN,
 * @WEIGHTAVG, @SEMEAN, @REFCONVERT, @XINDEX), Lotus Agenda's rule-based categories (CATEGORIZE),
 * StarCalc / OpenOffice.org Calc date helpers (DAYSINMONTH, DAYSINYEAR, ISLEAPYEAR, WEEKSINYEAR,
 * EASTERSUNDAY) and online spreadsheets (COUNTUNIQUE, ISEMAIL, ISURL, EPOCHTODATE).
 * docs/HERITAGE.md lists where each one comes from.
 */
import { defFn } from './registry';
import {
  num, str, bool, compare, equals, iterValues, numbers, toMatrix, isErr, fail, ERR, type Ctx, type Value,
} from '../values';
import type { Scalar } from '../../model/types';
import { colName, colIndex, MAX_COLS } from '../../model/address';
import { serialToParts, partsToSerial, isLeap, daysInMonth } from '../../model/dates';

const S = (a: Value) => a as Scalar;
const d1904 = (ctx: Ctx) => ctx.book.calc.date1904;

/** Year, month and day of a date serial (validated). */
function ymd(v: Value, ctx: Ctx): { y: number; m: number; d: number } {
  const n = Math.floor(num(S(v)));
  if (n < 0 || n > 2958465) return fail(ERR.num);
  const p = serialToParts(n, d1904(ctx));
  return { y: p.y, m: p.m, d: p.d };
}

/** Easter Sunday (Gregorian calendar) by the anonymous Gregorian algorithm. */
export function easterSunday(y: number): { m: number; d: number } {
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31), day = ((h + l - 7 * m + 114) % 31) + 1;
  return { m: month, d: day };
}

/** Number of ISO 8601 weeks (52 or 53) in a year. */
export function isoWeeksInYear(y: number): number {
  // A year has 53 ISO weeks when it starts on a Thursday, or is a leap year starting on a Wednesday.
  const dow = (yy: number) => new Date(Date.UTC(yy, 0, 1)).getUTCDay();
  const jan1 = dow(y);
  return jan1 === 4 || (isLeap(y) && jan1 === 3) ? 53 : 52;
}

/** Lotus 1-2-3 @REFCONVERT: column letters to numbers (A -> 1) and numbers to letters (28 -> AB). */
export function refConvert(v: Scalar): string | number | null {
  if (typeof v === 'number') {
    const n = Math.trunc(v);
    return n >= 1 && n <= MAX_COLS ? colName(n - 1) : null;
  }
  if (typeof v === 'string') {
    const t = v.trim();
    if (/^\d+$/.test(t)) return refConvert(+t);
    if (!/^[A-Za-z]{1,3}$/.test(t)) return null;
    const i = colIndex(t.toUpperCase());
    return i >= 0 && i < MAX_COLS ? i + 1 : null;
  }
  return null;
}

/** First category whose keyword list (comma-separated alternatives, case-insensitive substrings, * and ? wildcards) matches the text. */
export function categorize(text: string, rules: { keywords: string; category: Scalar }[]): Scalar | undefined {
  const t = text.toLowerCase();
  for (const rule of rules) {
    for (const raw of rule.keywords.split(',')) {
      const k = raw.trim().toLowerCase();
      if (!k) continue;
      if (/[*?]/.test(k)) {
        const re = new RegExp(k.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.'), 'i');
        if (re.test(t)) return rule.category;
      } else if (t.includes(k)) return rule.category;
    }
  }
  return undefined;
}

const EMAIL_RE = /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*\.[A-Za-z]{2,}$/;
const URL_RE = /^(?:(?:https?|ftp):\/\/)?(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,}(?::\d{1,5})?(?:[/?#][^\s]*)?$/i;

export const isEmailText = (s: string) => s.length <= 254 && EMAIL_RE.test(s.trim()) && !s.includes('..');
export const isUrlText = (s: string) => URL_RE.test(s.trim());

defFn([
  {
    name: 'ISBETWEEN', min: 3, max: 5, category: 'Logical', desc: 'TRUE if a value lies between two others (inclusive unless told otherwise). Works for numbers, dates and text.',
    params: ['value', 'lower', 'upper', '[lower_inclusive]', '[upper_inclusive]'],
    impl: ([v, lo, hi, li, ui]) => {
      const x = S(v), a0 = S(lo), b0 = S(hi);
      for (const e of [x, a0, b0]) if (isErr(e)) return e;
      if (x == null) return false;
      // The bounds can be given in either order (1-2-3 @ISBETWEEN).
      const [a, b] = compare(a0, b0) <= 0 ? [a0, b0] : [b0, a0];
      const inclLo = li === undefined || li === null ? true : bool(S(li));
      const inclHi = ui === undefined || ui === null ? true : bool(S(ui));
      const ca = compare(x, a), cb = compare(x, b);
      return (inclLo ? ca >= 0 : ca > 0) && (inclHi ? cb <= 0 : cb < 0);
    },
  },
  {
    name: 'WEIGHTAVG', min: 2, max: 2, category: 'Statistical', arrays: [0, 1], desc: 'Weighted average: the sum of value × weight divided by the sum of the weights.',
    params: ['values', 'weights'],
    impl: ([vals, wts], ctx) => {
      const a = toMatrix(vals, ctx), b = toMatrix(wts, ctx);
      const xs = a.flat(), ws = b.flat();
      if (xs.length !== ws.length) return ERR.value;
      let sw = 0, sxw = 0;
      for (let i = 0; i < xs.length; i++) {
        const x = xs[i], w = ws[i];
        if (isErr(x)) return x;
        if (isErr(w)) return w;
        if (typeof x !== 'number' || typeof w !== 'number') continue;
        sw += w; sxw += x * w;
      }
      return sw === 0 ? ERR.div0 : sxw / sw;
    },
  },
  {
    name: 'SEMEAN', min: 1, max: 255, category: 'Statistical', arrays: 'all', desc: 'Standard error of the mean: the sample standard deviation divided by the square root of the count.',
    params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => {
      const xs = numbers(args, ctx);
      const n = xs.length;
      if (n < 2) return ERR.div0;
      const m = xs.reduce((s, x) => s + x, 0) / n;
      const v = xs.reduce((s, x) => s + (x - m) ** 2, 0) / (n - 1);
      return Math.sqrt(v / n);
    },
  },
  {
    name: 'REFCONVERT', min: 1, max: 1, category: 'Lookup & Reference', desc: 'Converts column letters to a column number (AB -> 28) or a column number to letters (28 -> AB).',
    params: ['column'],
    impl: ([v]) => { const x = S(v); if (isErr(x)) return x; const r = refConvert(x); return r == null ? ERR.value : r; },
  },
  {
    name: 'XINDEX', min: 3, max: 3, category: 'Lookup & Reference', arrays: [0], desc: 'The value in a table found by its column heading (top row) and row heading (left column).',
    params: ['table', 'column_heading', 'row_heading'],
    impl: ([tbl, ch, rh], ctx) => {
      const m = toMatrix(tbl, ctx);
      const col = S(ch), row = S(rh);
      if (isErr(col)) return col;
      if (isErr(row)) return row;
      if (m.length < 2 || (m[0]?.length ?? 0) < 2) return ERR.ref;
      let c = -1, r = -1;
      for (let j = 1; j < m[0].length; j++) if (m[0][j] != null && equals(m[0][j], col)) { c = j; break; }
      for (let i = 1; i < m.length; i++) if (m[i][0] != null && equals(m[i][0], row)) { r = i; break; }
      if (c < 0 || r < 0) return ERR.na;
      return m[r][c] ?? 0;
    },
  },
  {
    name: 'COUNTUNIQUE', min: 1, max: 255, category: 'Statistical', arrays: 'all', desc: 'Number of distinct values (blank cells are ignored; text is compared without regard to case).',
    params: ['value1', '[value2]', '...'],
    impl: (args, ctx) => {
      const seen = new Set<string>();
      for (const a of args) for (const { v } of iterValues(a, ctx)) {
        if (v == null || v === '') continue;
        if (isErr(v)) seen.add('e:' + v.code);
        else seen.add(typeof v === 'string' ? 's:' + v.toLowerCase() : typeof v === 'number' ? 'n:' + v : 'b:' + v);
      }
      return seen.size;
    },
  },
  {
    name: 'ISEMAIL', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is text shaped like an email address.', params: ['value'],
    impl: ([v]) => { const x = S(v); return typeof x === 'string' && isEmailText(x); },
  },
  {
    name: 'ISURL', min: 1, max: 1, category: 'Information', desc: 'TRUE if the value is text shaped like a web address.', params: ['value'],
    impl: ([v]) => { const x = S(v); return typeof x === 'string' && isUrlText(x); },
  },
  {
    name: 'EPOCHTODATE', min: 1, max: 2, category: 'Date & Time', desc: 'Converts a Unix timestamp (seconds, milliseconds or microseconds since 1970-01-01 UTC) to a date-time serial.',
    params: ['timestamp', '[unit]'],
    impl: ([t, u], ctx) => {
      const ts = num(S(t));
      const unit = u === undefined || u === null ? 1 : Math.trunc(num(S(u)));
      const div = unit === 1 ? 1 : unit === 2 ? 1e3 : unit === 3 ? 1e6 : 0;
      if (!div || ts < 0) return ERR.num;
      const serial = ts / div / 86400 + 25569 - (d1904(ctx) ? 1462 : 0);
      return serial > 2958465 ? ERR.num : serial;
    },
  },
  {
    name: 'CATEGORIZE', min: 3, max: 4, category: 'Text', arrays: [1, 2], desc: 'The category of the first keyword found in the text. Keyword cells may list alternatives separated by commas; * and ? are wildcards.',
    params: ['text', 'keywords', 'categories', '[if_none]'],
    impl: ([t, kw, cats, dflt], ctx) => {
      const x = S(t);
      if (isErr(x)) return x;
      const ks = toMatrix(kw, ctx).flat(), cs = toMatrix(cats, ctx).flat();
      if (ks.length !== cs.length) return ERR.value;
      const rules = ks.map((k, i) => ({ keywords: k == null || isErr(k) ? '' : str(k), category: cs[i] }));
      const hit = categorize(x == null ? '' : str(x), rules);
      if (hit !== undefined) return hit ?? '';
      return dflt === undefined ? ERR.na : S(dflt);
    },
  },
  {
    name: 'DAYSINMONTH', min: 1, max: 1, category: 'Date & Time', desc: 'Number of days in the month of a date.', params: ['date'],
    impl: ([d], ctx) => { const p = ymd(d, ctx); return daysInMonth(p.y, p.m); },
  },
  {
    name: 'DAYSINYEAR', min: 1, max: 1, category: 'Date & Time', desc: 'Number of days (365 or 366) in the year of a date.', params: ['date'],
    impl: ([d], ctx) => (isLeap(ymd(d, ctx).y) ? 366 : 365),
  },
  {
    name: 'ISLEAPYEAR', min: 1, max: 1, category: 'Date & Time', desc: 'TRUE if the date falls in a leap year.', params: ['date'],
    impl: ([d], ctx) => isLeap(ymd(d, ctx).y),
  },
  {
    name: 'WEEKSINYEAR', min: 1, max: 1, category: 'Date & Time', desc: 'Number of ISO 8601 weeks (52 or 53) in the year of a date.', params: ['date'],
    impl: ([d], ctx) => isoWeeksInYear(ymd(d, ctx).y),
  },
  {
    name: 'EASTERSUNDAY', min: 1, max: 1, category: 'Date & Time', desc: 'Date of Easter Sunday (Gregorian calendar) in a year. Two-digit years below 30 are 20xx, others 19xx.', params: ['year'],
    impl: ([v], ctx) => {
      let y = Math.trunc(num(S(v)));
      if (y >= 0 && y < 100) y += y < 30 ? 2000 : 1900;
      if (y < 1583 || y > 9956) return ERR.num;
      const e = easterSunday(y);
      return partsToSerial(y, e.m, e.d, d1904(ctx));
    },
  },
]);
