/**
 * Database functions (DSUM, DAVERAGE, DGET, ...).
 *
 * `database` is a range whose first row holds field names; `field` is a field name or 1-based column
 * number; `criteria` is a range whose first row holds field names and whose following rows are
 * conditions. Conditions in one row are AND'ed and rows are OR'ed. Each condition cell uses the
 * COUNTIF syntax (">5", "<>x", "a*", ...); plain text matches values that begin with it, as in
 * advanced filters, and "=text" matches exactly.
 *
 * Computed criteria (a formula under a blank header or a header that is not a field name) are
 * evaluated for each record by re-evaluating the criteria cell's formula with its relative
 * references shifted from the first record to the current one. This needs the criteria to be a
 * worksheet reference; computed criteria inside array constants or other sheets' formulas that
 * reference unqualified cells on their own sheet are evaluated against the formula's sheet only.
 */
import { defFn, type FnDef } from './registry';
import {
  str, fail, Fail, parseNumeric, isRef, toMatrix, toScalar, criteria as makeCriteria, ERR, CellError,
  type Ctx, type Value, type RefV,
} from '../values';
import { parse, shiftFormula, mapRefs } from '../parser';
import type { Matrix, Scalar } from '../../model/types';

const S = (v: Value): Scalar => v as Scalar;

/** Condition test for one criteria cell under the D-function rules. */
function condition(c: Scalar): (v: Scalar) => boolean {
  if (typeof c === 'string' && c !== '' && !/^(<=|>=|<>|=|<|>)/.test(c)) {
    const isSpecial = parseNumeric(c) != null || /^(TRUE|FALSE)$/i.test(c) || c.startsWith('#');
    if (!isSpecial) return makeCriteria(c.endsWith('*') ? c : c + '*');
  }
  return makeCriteria(c);
}

interface Db { m: Matrix; heads: string[] }

function readDb(v: Value, ctx: Ctx): Db {
  const m = toMatrix(v, ctx);
  if (m.length < 1) fail(ERR.value);
  const heads = m[0].map((h) => (h == null ? '' : h instanceof CellError ? '' : str(h).toLowerCase()));
  return { m, heads };
}

function fieldIndex(db: Db, f: Scalar): number {
  if (f instanceof CellError) return fail(f);
  if (typeof f === 'number') {
    const i = Math.trunc(f);
    if (i < 1 || i > db.heads.length) fail(ERR.value);
    return i - 1;
  }
  if (typeof f === 'boolean' || f == null) return fail(ERR.value);
  const i = db.heads.indexOf(f.toLowerCase());
  if (i < 0) fail(ERR.value);
  return i;
}

/** Indices (0-based, excluding the header) of the records that satisfy the criteria. */
function matchingRows(db: Db, critV: Value, ctx: Ctx, dbV: Value): number[] {
  const cm = toMatrix(critV, ctx);
  if (cm.length < 2) fail(ERR.value);
  const cheads = cm[0];
  const critRef = isRef(critV) ? critV : null;
  const dbRef = isRef(dbV) ? dbV : null;
  type Test = (row: Scalar[], rec: number) => boolean;
  const rowTests: Test[][] = [];
  for (let r = 1; r < cm.length; r++) {
    const tests: Test[] = [];
    for (let c = 0; c < cheads.length; c++) {
      const h = cheads[c];
      const name = h == null || h instanceof CellError ? '' : str(h).toLowerCase();
      const col = name ? db.heads.indexOf(name) : -1;
      const formula = critRef ? critRef.sheet.get(critRef.r1 + r, critRef.c1 + c)?.f : undefined;
      const cell = cm[r][c];
      if (col >= 0) {
        if (cell == null || cell === '') continue;
        const t = condition(cell);
        tests.push((row) => t(row[col]));
      } else if (formula && critRef) {
        tests.push(computed(formula, critRef, dbRef, ctx));
      } else if (cell == null || cell === '') {
        continue;
      } else {
        // A condition under an unknown header: only a constant TRUE lets records through.
        const ok = cell === true;
        tests.push(() => ok);
      }
    }
    rowTests.push(tests);
  }
  const out: number[] = [];
  for (let i = 1; i < db.m.length; i++) {
    const row = db.m[i];
    if (rowTests.some((tests) => tests.every((t) => t(row, i - 1)))) out.push(i - 1);
  }
  return out;
}

/** Computed criterion: re-evaluates the criteria formula with relative references moved to each record. */
function computed(formula: string, critRef: RefV, dbRef: RefV | null, ctx: Ctx): (row: Scalar[], rec: number) => boolean {
  let src = formula;
  if (critRef.sheet !== ctx.sheet) {
    const name = critRef.sheet.name;
    src = mapRefs(src, (ref) => (ref.sheet ? null : { ...ref, sheet: name }));
  }
  const cache = new Map<number, boolean>();
  return (_row, rec) => {
    if (!dbRef) return false;
    let hit = cache.get(rec);
    if (hit !== undefined) return hit;
    try {
      const node = parse(shiftFormula(src, rec, 0));
      const v = toScalar(ctx.evalNode(node), ctx);
      hit = typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : false;
    } catch (e) {
      if (!(e instanceof Fail) && !(e instanceof SyntaxError) && (e as Error)?.constructor?.name !== 'FormulaError') throw e;
      hit = false;
    }
    cache.set(rec, hit);
    return hit;
  };
}

type Agg = (vals: Scalar[], count: number) => Scalar;

function dfn(name: string, desc: string, agg: Agg, fieldOptional = false): FnDef {
  return {
    name, min: 3, max: 3, category: 'Database', desc, params: ['database', 'field', 'criteria'], arrays: [0, 2],
    impl: (a, ctx) => {
      const db = readDb(a[0], ctx);
      const f = S(a[1]);
      const rows = matchingRows(db, a[2], ctx, a[0]);
      if (fieldOptional && f == null) return rows.length;
      const col = fieldIndex(db, f);
      return agg(rows.map((i) => db.m[i + 1][col]), rows.length);
    },
  };
}

function nums(vals: Scalar[]): number[] {
  const out: number[] = [];
  for (const v of vals) {
    if (v instanceof CellError) fail(v);
    if (typeof v === 'number') out.push(v);
  }
  return out;
}

function variance(xs: number[], sample: boolean): number {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) fail(ERR.div0);
  const mean = xs.reduce((s, x) => s + x, 0) / n;
  const ss = xs.reduce((s, x) => s + (x - mean) * (x - mean), 0);
  return ss / (sample ? n - 1 : n);
}

defFn([
  dfn('DAVERAGE', 'Average of a field over the records that match the criteria.', (v) => {
    const xs = nums(v);
    if (!xs.length) fail(ERR.div0);
    return xs.reduce((s, x) => s + x, 0) / xs.length;
  }),
  dfn('DCOUNT', 'Counts numeric values of a field over the records that match the criteria.', (v) => v.filter((x) => typeof x === 'number').length, true),
  dfn('DCOUNTA', 'Counts non-blank values of a field over the records that match the criteria.', (v) => v.filter((x) => x != null && x !== '').length, true),
  dfn('DGET', 'Returns the single value of a field from the one record that matches the criteria.', (v) => {
    if (v.length === 0) fail(ERR.value);
    if (v.length > 1) fail(ERR.num);
    return v[0] ?? 0;
  }),
  dfn('DMAX', 'Largest value of a field over the records that match the criteria.', (v) => { const xs = nums(v); return xs.length ? Math.max(...xs) : 0; }),
  dfn('DMIN', 'Smallest value of a field over the records that match the criteria.', (v) => { const xs = nums(v); return xs.length ? Math.min(...xs) : 0; }),
  dfn('DPRODUCT', 'Product of a field over the records that match the criteria.', (v) => { const xs = nums(v); return xs.length ? xs.reduce((p, x) => p * x, 1) : 0; }),
  dfn('DSTDEV', 'Sample standard deviation of a field over the records that match the criteria.', (v) => Math.sqrt(variance(nums(v), true))),
  dfn('DSTDEVP', 'Population standard deviation of a field over the records that match the criteria.', (v) => Math.sqrt(variance(nums(v), false))),
  dfn('DSUM', 'Sum of a field over the records that match the criteria.', (v) => nums(v).reduce((s, x) => s + x, 0)),
  dfn('DVAR', 'Sample variance of a field over the records that match the criteria.', (v) => variance(nums(v), true)),
  dfn('DVARP', 'Population variance of a field over the records that match the criteria.', (v) => variance(nums(v), false)),
]);
