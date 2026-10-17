/**
 * VisiCalc worksheets (.VC), as saved by VisiCalc and Advanced VisiCalc with /S.
 *
 * A saved worksheet is a list of cell entries followed by global commands:
 *
 *   >A1:/FR"TOTAL          cell A1, right-aligned label
 *   >B3:+B1+B2             formula
 *   >C3:/F$123.5           number shown with two decimals
 *   >D9:@SUM(D1...D8)      function with a range
 *   /W1  /GOC  /GRA  /GC9   window, order, recalculation and column-width settings
 *
 * VisiCalc evaluates formulas strictly from left to right ("2+3*4" is 20), so the translation adds
 * parentheses wherever this app's precedence would change the result. Formats: /FL left, /FR right,
 * /FI integer, /F$ two decimals, /FG general, /F* bar of asterisks (kept as the number).
 */
import { Workbook, Sheet } from '../../model/workbook';
import { CellError, type Cell, type CellStyle } from '../../model/types';
import { cellKey, colIndex, MAX_COLS, MAX_ROWS } from '../../model/address';
import { charsToPx } from './bin';

const FUNCS: Record<string, string> = {
  SUM: 'SUM', MIN: 'MIN', MAX: 'MAX', COUNT: 'COUNT', AVERAGE: 'AVERAGE', NPV: 'NPV', LOOKUP: 'LOOKUP', ABS: 'ABS', INT: 'TRUNC',
  EXP: 'EXP', SQRT: 'SQRT', LN: 'LN', LOG10: 'LOG10', SIN: 'SIN', ASIN: 'ASIN', COS: 'COS', ACOS: 'ACOS', TAN: 'TAN', ATAN: 'ATAN',
  IF: 'IF', AND: 'AND', OR: 'OR', NOT: 'NOT', ISNA: 'ISNA', ISERROR: 'ISERROR', CHOOSE: 'CHOOSE', MOD: 'MOD', ROUND: 'ROUND',
};
const CONSTS: Record<string, string> = { PI: 'PI()', NA: 'NA()', ERROR: '#VALUE!', TRUE: 'TRUE', FALSE: 'FALSE' };

export class VisiCalcError extends Error {}

/** Translates a VisiCalc formula (left-to-right evaluation) to a formula without "=". */
export function translateVisiCalc(src: string): string {
  let i = 0;
  const s = src.trim();
  const peek = () => s[i];
  const skipWs = () => { while (s[i] === ' ') i++; };
  const ref = (): string | null => {
    const m = /^([A-Z]{1,2})(\d{1,3})/.exec(s.slice(i));
    if (!m) return null;
    i += m[0].length;
    return m[1] + m[2];
  };
  const operand = (): { t: string; atom: boolean } => {
    skipWs();
    const ch = peek();
    if (ch === '+' || ch === '-') { i++; const o = operand(); return { t: ch === '-' ? `-${o.atom ? o.t : `(${o.t})`}` : o.t, atom: ch === '+' && o.atom }; }
    if (ch === '(') { i++; const e = expr(); skipWs(); if (s[i] !== ')') throw new VisiCalcError('missing )'); i++; return { t: `(${e})`, atom: true }; }
    if (ch === '@') {
      const m = /^@([A-Z][A-Z0-9]*)/.exec(s.slice(i));
      if (!m) throw new VisiCalcError('bad function');
      i += m[0].length;
      const name = m[1];
      if (CONSTS[name] && s[i] !== '(') return { t: CONSTS[name], atom: true };
      const fn = FUNCS[name];
      if (!fn) throw new VisiCalcError(`@${name} is not known`);
      if (s[i] !== '(') throw new VisiCalcError(`@${name} needs arguments`);
      i++;
      const args: string[] = [];
      for (;;) {
        skipWs();
        const a = /^([A-Z]{1,2}\d{1,3})\.\.\.?([A-Z]{1,2}\d{1,3})/.exec(s.slice(i));
        if (a) { i += a[0].length; args.push(`${a[1]}:${a[2]}`); }
        else args.push(expr());
        skipWs();
        if (s[i] === ',') { i++; continue; }
        if (s[i] === ')') { i++; break; }
        throw new VisiCalcError('missing ) after arguments');
      }
      return { t: `${fn}(${args.join(',')})`, atom: true };
    }
    const r = ref();
    if (r) return { t: r, atom: true };
    const n = /^(\d+\.?\d*|\.\d+)(E[+-]?\d+)?/i.exec(s.slice(i));
    if (n) { i += n[0].length; return { t: n[0], atom: true }; }
    throw new VisiCalcError(`unexpected "${ch ?? ''}"`);
  };
  const expr = (): string => {
    let left = operand();
    for (;;) {
      skipWs();
      const m = /^(<=|>=|<>|[-+*/^=<>])/.exec(s.slice(i));
      if (!m) break;
      i += m[0].length;
      const right = operand();
      // Left to right: the result so far is one operand of the next operator.
      left = { t: `${left.atom ? left.t : `(${left.t})`}${m[1]}${right.atom ? right.t : `(${right.t})`}`, atom: false };
    }
    return left.t;
  };
  const out = expr();
  skipWs();
  if (i < s.length) throw new VisiCalcError(`unexpected text "${s.slice(i)}"`);
  return out;
}

export function importVisiCalc(text: string, name = '', warnings: string[] = []): Workbook {
  const book = new Workbook(false);
  const sh = new Sheet(book, 'Sheet1');
  book.sheets.push(sh);
  sh.defaultColW = charsToPx(9);
  const styles = new Map<string, number>();
  const sid = (st: CellStyle) => { const k = JSON.stringify(st); let id = styles.get(k); if (id == null) { id = Object.keys(st).length ? book.styleId(st) : 0; styles.set(k, id); } return id; };
  let bad = 0, cells = 0;
  let globalFmt = '';
  for (const line of text.split(/\r\n?|\n/)) {
    const g = /^\/G(C(\d+)|F([LRIG$*D]))/.exec(line);
    if (g) {
      if (g[2]) sh.defaultColW = charsToPx(+g[2]);
      if (g[3]) globalFmt = g[3];
      continue;
    }
    const m = /^>([A-Z]{1,2})(\d{1,3}):(.*)$/.exec(line);
    if (!m) continue;
    const c = colIndex(m[1]), r = +m[2] - 1;
    if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) continue;
    let body = m[3];
    let fmt = globalFmt;
    const f = /^\/F([LRIG$*D])/.exec(body);
    if (f) { fmt = f[1]; body = body.slice(3); }
    const st: CellStyle = {};
    if (fmt === 'L') st.hAlign = 'left';
    else if (fmt === 'R') st.hAlign = 'right';
    else if (fmt === 'I') st.numFmt = '0';
    else if (fmt === '$') st.numFmt = '0.00';
    const cell: Cell = {};
    if (body.startsWith('/-')) {
      // Repeating label: the character(s) fill the cell.
      cell.v = body.slice(2);
      st.hAlign = 'fill';
    } else if (body.startsWith('"')) {
      cell.v = body.slice(1);
    } else if (/^[+-]?(\d+\.?\d*|\.\d+)(E[+-]?\d+)?$/i.test(body.trim())) {
      cell.v = Number(body);
    } else if (body.trim()) {
      try { cell.f = translateVisiCalc(body); cell.v = null; }
      catch (e: any) {
        bad++;
        cell.v = new CellError('#VALUE!');
        sh.comments.set(cellKey(r, c), { author: 'Import', text: `VisiCalc formula, not translated (${e?.message || e}): ${body}`, time: Date.now(), note: true });
      }
    } else continue;
    const s = sid(st);
    if (s) cell.s = s;
    sh.setComputed(r, c, cell);
    cells++;
  }
  if (!cells) throw new Error('No VisiCalc cells were found in this file.');
  if (bad) warnings.push(`${bad} formula${bad > 1 ? 's' : ''} could not be translated (see the cell notes).`);
  if (name) book.props.title = name.replace(/\.[^.]+$/, '');
  book.activeSheet = 0;
  return book;
}
