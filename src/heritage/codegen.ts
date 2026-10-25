/**
 * Export a sheet as a program, after Resolver One, which showed every spreadsheet as the Python
 * program that computes it. The generated JavaScript or Python assigns each input cell, then
 * computes each formula cell in dependency order with a small runtime of spreadsheet functions,
 * and prints the results. Formulas that use functions outside that runtime call `unsupported()`,
 * and the header lists them, so the program never silently computes something different.
 */
import { parse, parseRefText, type Node, type RefToken } from '../calc/parser';
import { cellName, keyRow, keyCol, cellKey, MAX_COLS, MAX_ROWS } from '../model/address';
import type { Sheet, Workbook } from '../model/workbook';
import { CellError, type Scalar } from '../model/types';

export type CodeLang = 'js' | 'python';

/** Functions the generated runtime implements (same meaning in both languages). */
export const CODE_FUNCTIONS = [
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'PRODUCT', 'MEDIAN', 'SUMPRODUCT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'TRUNC',
  'ABS', 'SQRT', 'POWER', 'MOD', 'EXP', 'LN', 'LOG10', 'PI', 'SIGN', 'IF', 'AND', 'OR', 'NOT', 'IFERROR', 'CONCAT', 'CONCATENATE',
  'LEN', 'UPPER', 'LOWER', 'TRIM', 'LEFT', 'RIGHT', 'MID', 'PMT', 'TRUE', 'FALSE',
] as const;

interface CellRef { sheet: Sheet; r: number; c: number }

/** Runtime names a cell variable must not shadow (LOG10 is also a cell address). */
const RESERVED = new Set<string>([...CODE_FUNCTIONS, 'err', 'num', 'str', 'bool', 'eq', 'div', 'flat', 'nums', 'unsupported', 'math', 'RESULTS']);

const ident = (book: Workbook, home: Sheet, sheet: Sheet, r: number, c: number) => {
  const base = cellName(r, c) + (RESERVED.has(cellName(r, c)) ? '_' : '');
  if (sheet === home) return base;
  const s = sheet.name.replace(/[^A-Za-z0-9_]/g, '_');
  return `${/^[A-Za-z_]/.test(s) ? s : '_' + s}_${base}`;
};

function literal(v: Scalar, lang: CodeLang): string {
  if (v == null) return lang === 'js' ? 'null' : 'None';
  if (typeof v === 'boolean') return lang === 'js' ? String(v) : v ? 'True' : 'False';
  if (typeof v === 'number') return Number.isFinite(v) ? String(v) : '0';
  if (v instanceof CellError) return `err(${JSON.stringify(v.code)})`;
  return JSON.stringify(v);
}

const JS_RUNTIME = String.raw`// Runtime: spreadsheet semantics for the functions used below.
class SheetError extends Error {}
const err = (code) => { throw new SheetError(code); };
const unsupported = (what) => err('unsupported: ' + what);
const flat = (xs) => xs.flat(Infinity);
const num = (x) => (x == null ? 0 : typeof x === 'boolean' ? +x : typeof x === 'number' ? x : isNaN(+x) || x === '' ? err('#VALUE!') : +x);
const str = (x) => (x == null ? '' : typeof x === 'boolean' ? (x ? 'TRUE' : 'FALSE') : String(x));
const bool = (x) => (typeof x === 'string' ? (x.toUpperCase() === 'TRUE' ? true : x.toUpperCase() === 'FALSE' ? false : err('#VALUE!')) : !!num(x));
const eq = (a, b) => { if (a == null) a = typeof b === 'string' ? '' : 0; if (b == null) b = typeof a === 'string' ? '' : 0; return typeof a === 'string' && typeof b === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b; };
const nums = (xs) => flat(xs).filter((x) => typeof x === 'number');
const div = (a, b) => (num(b) === 0 ? err('#DIV/0!') : num(a) / num(b));
const roundTo = (x, d, f) => { const m = 10 ** num(d); return Math.sign(num(x)) * f(Math.abs(num(x)) * m + 1e-9 * (f === Math.round)) / m; };
const SUM = (...a) => nums(a).reduce((s, x) => s + x, 0);
const AVERAGE = (...a) => { const v = nums(a); return v.length ? SUM(v) / v.length : err('#DIV/0!'); };
const MIN = (...a) => { const v = nums(a); return v.length ? Math.min(...v) : 0; };
const MAX = (...a) => { const v = nums(a); return v.length ? Math.max(...v) : 0; };
const COUNT = (...a) => nums(a).length;
const COUNTA = (...a) => flat(a).filter((x) => x != null && x !== '').length;
const PRODUCT = (...a) => nums(a).reduce((s, x) => s * x, 1);
const MEDIAN = (...a) => { const v = nums(a).sort((x, y) => x - y); const n = v.length; return n ? (n % 2 ? v[(n - 1) / 2] : (v[n / 2 - 1] + v[n / 2]) / 2) : err('#NUM!'); };
const SUMPRODUCT = (...a) => { const v = a.map((m) => flat(m)); return v[0].reduce((s, _, i) => s + v.reduce((p, col) => p * (typeof col[i] === 'number' ? col[i] : 0), 1), 0); };
const ROUND = (x, d = 0) => roundTo(x, d, Math.round);
const ROUNDUP = (x, d = 0) => roundTo(x, d, Math.ceil);
const ROUNDDOWN = (x, d = 0) => roundTo(x, d, Math.floor);
const INT = (x) => Math.floor(num(x));
const TRUNC = (x, d = 0) => ROUNDDOWN(x, d);
const ABS = (x) => Math.abs(num(x));
const SQRT = (x) => (num(x) < 0 ? err('#NUM!') : Math.sqrt(num(x)));
const POWER = (x, y) => num(x) ** num(y);
const MOD = (x, y) => { const d = num(y); if (d === 0) err('#DIV/0!'); const r = num(x) % d; return r && Math.sign(r) !== Math.sign(d) ? r + d : r; };
const EXP = (x) => Math.exp(num(x));
const LN = (x) => (num(x) <= 0 ? err('#NUM!') : Math.log(num(x)));
const LOG10 = (x) => (num(x) <= 0 ? err('#NUM!') : Math.log10(num(x)));
const PI = () => Math.PI;
const SIGN = (x) => Math.sign(num(x));
const IF = (c, a, b) => (bool(c) ? (a ? a() : true) : (b ? b() : false));
const AND = (...a) => flat(a).filter((x) => x != null).every(bool);
const OR = (...a) => flat(a).filter((x) => x != null).some(bool);
const NOT = (x) => !bool(x);
const IFERROR = (a, b) => { try { return a(); } catch (e) { if (e instanceof SheetError) return b(); throw e; } };
const CONCAT = (...a) => flat(a).map(str).join('');
const CONCATENATE = CONCAT;
const LEN = (s) => str(s).length;
const UPPER = (s) => str(s).toUpperCase();
const LOWER = (s) => str(s).toLowerCase();
const TRIM = (s) => str(s).trim().replace(/ +/g, ' ');
const LEFT = (s, n = 1) => str(s).slice(0, num(n));
const RIGHT = (s, n = 1) => (num(n) ? str(s).slice(-num(n)) : '');
const MID = (s, i, n) => str(s).substr(num(i) - 1, num(n));
const PMT = (r, n, pv, fv = 0, t = 0) => { r = num(r); n = num(n); pv = num(pv); fv = num(fv); t = num(t); if (!r) return -(pv + fv) / n; const f = (1 + r) ** n; return -(r * (pv * f + fv)) / ((1 + r * t) * (f - 1)); };
`;

const PY_RUNTIME = String.raw`# Runtime: spreadsheet semantics for the functions used below.
import math


class SheetError(Exception):
    pass


def err(code):
    raise SheetError(code)


def unsupported(what):
    err("unsupported: " + what)


def flat(xs):
    out = []
    for x in xs:
        if isinstance(x, (list, tuple)):
            out.extend(flat(x))
        else:
            out.append(x)
    return out


def num(x):
    if x is None:
        return 0
    if isinstance(x, bool):
        return int(x)
    if isinstance(x, (int, float)):
        return x
    try:
        return float(x)
    except ValueError:
        err("#VALUE!")


def str_(x):
    if x is None:
        return ""
    if isinstance(x, bool):
        return "TRUE" if x else "FALSE"
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x)


def bool_(x):
    if isinstance(x, str):
        if x.upper() in ("TRUE", "FALSE"):
            return x.upper() == "TRUE"
        err("#VALUE!")
    return num(x) != 0


def eq(a, b):
    if a is None:
        a = "" if isinstance(b, str) else 0
    if b is None:
        b = "" if isinstance(a, str) else 0
    if isinstance(a, str) and isinstance(b, str):
        return a.lower() == b.lower()
    return a == b


def nums(xs):
    return [x for x in flat(xs) if isinstance(x, (int, float)) and not isinstance(x, bool)]


def div(a, b):
    if num(b) == 0:
        err("#DIV/0!")
    return num(a) / num(b)


def _round(x, d, f):
    m = 10 ** num(d)
    v = abs(num(x)) * m
    return math.copysign(f(v + (1e-9 if f is _half_up else 0)) / m, num(x)) if num(x) else 0.0


def _half_up(v):
    return math.floor(v + 0.5)


def SUM(*a): return sum(nums(a))
def AVERAGE(*a): v = nums(a); return sum(v) / len(v) if v else err("#DIV/0!")
def MIN(*a): v = nums(a); return min(v) if v else 0
def MAX(*a): v = nums(a); return max(v) if v else 0
def COUNT(*a): return len(nums(a))
def COUNTA(*a): return len([x for x in flat(a) if x is not None and x != ""])
def PRODUCT(*a): return math.prod(nums(a))
def MEDIAN(*a): v = sorted(nums(a)); n = len(v); return (v[(n - 1) // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2) if n else err("#NUM!")
def SUMPRODUCT(*a):
    cols = [flat(m) for m in a]
    return sum(math.prod(c[i] if isinstance(c[i], (int, float)) else 0 for c in cols) for i in range(len(cols[0])))
def ROUND(x, d=0): return _round(x, d, _half_up)
def ROUNDUP(x, d=0): return _round(x, d, math.ceil)
def ROUNDDOWN(x, d=0): return _round(x, d, math.floor)
def INT(x): return math.floor(num(x))
def TRUNC(x, d=0): return ROUNDDOWN(x, d)
def ABS(x): return abs(num(x))
def SQRT(x): return math.sqrt(num(x)) if num(x) >= 0 else err("#NUM!")
def POWER(x, y): return num(x) ** num(y)
def MOD(x, y): return num(x) % num(y) if num(y) else err("#DIV/0!")
def EXP(x): return math.exp(num(x))
def LN(x): return math.log(num(x)) if num(x) > 0 else err("#NUM!")
def LOG10(x): return math.log10(num(x)) if num(x) > 0 else err("#NUM!")
def PI(): return math.pi
def SIGN(x): return (num(x) > 0) - (num(x) < 0)
def IF(c, a=None, b=None): return (a() if a else True) if bool_(c) else (b() if b else False)
def AND(*a): return all(bool_(x) for x in flat(a) if x is not None)
def OR(*a): return any(bool_(x) for x in flat(a) if x is not None)
def NOT(x): return not bool_(x)
def IFERROR(a, b):
    try:
        return a()
    except SheetError:
        return b()
def CONCAT(*a): return "".join(str_(x) for x in flat(a))
CONCATENATE = CONCAT
def LEN(s): return len(str_(s))
def UPPER(s): return str_(s).upper()
def LOWER(s): return str_(s).lower()
def TRIM(s): return " ".join(w for w in str_(s).split(" ") if w)
def LEFT(s, n=1): return str_(s)[:int(num(n))]
def RIGHT(s, n=1): return str_(s)[-int(num(n)):] if num(n) else ""
def MID(s, i, n): return str_(s)[int(num(i)) - 1:int(num(i)) - 1 + int(num(n))]
def PMT(r, n, pv, fv=0, t=0):
    r, n, pv, fv, t = num(r), num(n), num(pv), num(fv), num(t)
    if not r:
        return -(pv + fv) / n
    f = (1 + r) ** n
    return -(r * (pv * f + fv)) / ((1 + r * t) * (f - 1))
`;

export interface CodeResult { code: string; unsupported: string[]; cells: number; formulas: number }

/** Generates a program (JavaScript or Python) that computes the sheet's formulas. */
export function sheetToCode(book: Workbook, sheet: Sheet, lang: CodeLang): CodeResult {
  const unsupported = new Set<string>();
  const order: CellRef[] = [];
  const state = new Map<string, 0 | 1 | 2>(); // 1 = visiting, 2 = done
  const exprs = new Map<string, string>();
  const keyOf = (x: CellRef) => `${x.sheet.id}:${cellKey(x.r, x.c)}`;
  const nameOf = (x: CellRef) => ident(book, sheet, x.sheet, x.r, x.c);
  const py = lang === 'python';
  const cycles: string[] = [];

  const clip = (sh: Sheet, ref: RefToken) => {
    const u = sh.usedRange(false) || { r1: 0, c1: 0, r2: 0, c2: 0 };
    return { r1: ref.r1, c1: ref.c1, r2: Math.min(ref.r2, ref.r2 >= MAX_ROWS - 1 ? u.r2 : ref.r2), c2: Math.min(ref.c2, ref.c2 >= MAX_COLS - 1 ? u.c2 : ref.c2) };
  };
  const sheetOf = (ref: RefToken, from: Sheet) => (ref.sheet ? book.sheetByName(ref.sheet) : from);

  const visit = (x: CellRef): void => {
    const k = keyOf(x);
    const st = state.get(k);
    if (st === 2) return;
    if (st === 1) { cycles.push(nameOf(x)); return; }
    state.set(k, 1);
    const cell = x.sheet.get(x.r, x.c);
    if (cell?.f != null) {
      let node: Node | null = null;
      try { node = parse(cell.f); } catch { unsupported.add('unparseable formula'); }
      exprs.set(k, node ? emit(node, x.sheet) : 'unsupported("formula")');
    }
    state.set(k, 2);
    order.push(x);
  };

  const refExpr = (ref: RefToken, from: Sheet): string => {
    const sh = sheetOf(ref, from);
    if (!sh || ref.invalid || ref.sheet2) { unsupported.add(ref.sheet2 ? '3-D references' : 'invalid references'); return 'unsupported("reference")'; }
    if (ref.kind === 'cell') { const x = { sheet: sh, r: ref.r1, c: ref.c1 }; visit(x); return nameOf(x); }
    const rg = clip(sh, ref);
    const rows: string[] = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      const row: string[] = [];
      for (let c = rg.c1; c <= rg.c2; c++) {
        const cell = sh.get(r, c);
        if (!cell || (cell.v == null && cell.f == null)) { row.push(py ? 'None' : 'null'); continue; }
        const x = { sheet: sh, r, c };
        visit(x);
        row.push(nameOf(x));
      }
      rows.push(`[${row.join(', ')}]`);
    }
    return `[${rows.join(', ')}]`;
  };

  const lazy = (s: string) => (py ? `lambda: ${s}` : `() => ${s}`);

  function emit(n: Node, from: Sheet): string {
    switch (n.t) {
      case 'num': return String(n.v);
      case 'str': return JSON.stringify(n.v);
      case 'bool': return py ? (n.v ? 'True' : 'False') : String(n.v);
      case 'err': return `err(${JSON.stringify(n.v)})`;
      case 'ref': return refExpr(n.ref, from);
      case 'name': {
        const dn = book.names.find((d) => d.name.toUpperCase() === n.name.toUpperCase() && (!d.sheet || d.sheet === from.name));
        const ref = dn ? parseRefText(dn.ref.replace(/^=/, '')) : null;
        if (ref) return refExpr(ref, from);
        unsupported.add(`name ${n.name}`);
        return `unsupported(${JSON.stringify('name ' + n.name)})`;
      }
      case 'un': return `(-num(${emit(n.a, from)}))`;
      case 'pct': return `(num(${emit(n.a, from)}) / 100)`;
      case 'isect': return emit(n.a, from);
      case 'bin': {
        const a = emit(n.a, from), b = emit(n.b, from);
        switch (n.op) {
          case '+': case '-': case '*': return `(num(${a}) ${n.op} num(${b}))`;
          case '/': return `div(${a}, ${b})`;
          case '^': return `(num(${a}) ** num(${b}))`;
          case '&': return py ? `(str_(${a}) + str_(${b}))` : `(str(${a}) + str(${b}))`;
          case '=': return `eq(${a}, ${b})`;
          case '<>': return py ? `(not eq(${a}, ${b}))` : `(!eq(${a}, ${b}))`;
          case '<': case '>': case '<=': case '>=': return `(num(${a}) ${n.op} num(${b}))`;
        }
        unsupported.add(`operator ${n.op}`);
        return `unsupported(${JSON.stringify('operator ' + n.op)})`;
      }
      case 'fn': {
        const name = n.name.toUpperCase();
        if (name === 'TRUE' || name === 'FALSE') return py ? (name === 'TRUE' ? 'True' : 'False') : name.toLowerCase();
        if (!(CODE_FUNCTIONS as readonly string[]).includes(name)) {
          unsupported.add(name);
          return `unsupported(${JSON.stringify(name)})`;
        }
        const args = n.args.map((x) => (x.t === 'missing' ? (py ? 'None' : 'null') : emit(x, from)));
        if (name === 'IF') return `IF(${[args[0], ...args.slice(1).map(lazy)].join(', ')})`;
        if (name === 'IFERROR') return `IFERROR(${args.map(lazy).join(', ')})`;
        return `${name}(${args.join(', ')})`;
      }
      default:
        unsupported.add(n.t === 'arr' ? 'array constants' : n.t === 'sref' ? 'structured references' : n.t === 'spill' ? 'spill references' : n.t);
        return `unsupported(${JSON.stringify(n.t)})`;
    }
  }

  const keys = [...sheet.cells.keys()].sort((a, b) => keyRow(a) - keyRow(b) || keyCol(a) - keyCol(b));
  for (const k of keys) {
    const cell = sheet.cells.get(k)!;
    if (cell.v == null && cell.f == null) continue;
    if (cell.sp != null && cell.f == null) continue; // spilled results belong to their formula
    visit({ sheet, r: keyRow(k), c: keyCol(k) });
  }

  const c = py ? '#' : '//';
  const lines: string[] = [];
  lines.push(`${c} ${sheet.name}: generated by Anne\'s Grid (Export as Code).`);
  lines.push(`${c} Each cell is a variable; formulas are computed in dependency order.`);
  if (unsupported.size) lines.push(`${c} Not translated (these cells raise an error when run): ${[...unsupported].sort().join(', ')}`);
  if (cycles.length) lines.push(`${c} Circular references found at: ${[...new Set(cycles)].join(', ')}`);
  lines.push('', py ? PY_RUNTIME : JS_RUNTIME, `${c} Cells`);
  const outputs: string[] = [];
  let nf = 0;
  for (const x of order) {
    const k = keyOf(x);
    const name = nameOf(x);
    const e = exprs.get(k);
    if (e == null) { lines.push(`${py ? '' : 'const '}${name} = ${literal(x.sheet.value(x.r, x.c), lang)}`); continue; }
    nf++;
    if (x.sheet === sheet) outputs.push(name);
    if (py) {
      lines.push(`try:\n    ${name} = ${e}\nexcept SheetError as e:\n    ${name} = "#" + str(e).lstrip("#")`);
    } else lines.push(`let ${name}; try { ${name} = ${e}; } catch (e) { if (!(e instanceof SheetError)) throw e; ${name} = '#' + e.message.replace(/^#/, ''); }`);
  }
  lines.push('', `${c} Results`);
  if (py) {
    lines.push(`RESULTS = {${outputs.map((o) => `${JSON.stringify(o)}: ${o}`).join(', ')}}`);
    lines.push('if __name__ == "__main__":', '    for k, v in RESULTS.items():', '        print(k, "=", v)');
  } else {
    lines.push(`const RESULTS = { ${outputs.join(', ')} };`);
    lines.push('console.log(RESULTS);');
  }
  return { code: lines.join('\n') + '\n', unsupported: [...unsupported].sort(), cells: order.length, formulas: nf };
}
