/** Dynamic array (array-shaping) functions. */
import { defFn, getFn } from './registry';
import {
  fail, isErr, isLambda, isMatrix, isRef, isRefs, toMatrix, toScalar, num, int, bool, compare, transpose, flat,
  ERR, type Ctx, type Value, type RefV, type LambdaV,
} from '../values';
import type { Matrix, Scalar } from '../../model/types';
import { MAX_COLS, MAX_ROWS } from '../../model/address';

const MAX_CELLS = MAX_ROWS * 64;

/** Scalar from an argument that may be a reference or array. */
function sc(v: Value | undefined, ctx: Ctx): Scalar | undefined {
  if (v === undefined) return undefined;
  if (isRef(v) || isMatrix(v)) return toScalar(v, ctx);
  if (isRefs(v)) return fail(ERR.value);
  if (isLambda(v)) return fail(ERR.calc);
  return v;
}

/** Optional integer: omitted or empty uses the default. */
function optInt(v: Value | undefined, ctx: Ctx, def: number): number {
  const x = sc(v, ctx);
  if (x === undefined || x === null) return def;
  return int(x);
}

function optBool(v: Value | undefined, ctx: Ctx, def: boolean): boolean {
  const x = sc(v, ctx);
  if (x === undefined || x === null) return def;
  return bool(x);
}

function mat(v: Value | undefined, ctx: Ctx): Matrix {
  if (v === undefined) return [[null]];
  if (isRefs(v)) return fail(ERR.value);
  if (isLambda(v)) return fail(ERR.calc);
  const m = toMatrix(v, ctx);
  if (!m.length || !m[0].length) return fail(ERR.calc);
  return m;
}

/** Integer list from a scalar or array argument. */
function intList(v: Value, ctx: Ctx): number[] {
  return flat(mat(v, ctx)).map((x) => { if (isErr(x)) fail(x); return int(x); });
}

const S = (a: Value) => a as Scalar;

function pad(m: Matrix, rows: number, cols: number, with_: Scalar): Matrix {
  const out: Matrix = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) row.push(r < m.length && c < m[r].length ? m[r][c] : with_);
    out.push(row);
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Sorting

const typeRank = (v: Scalar) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);

/** Sort comparison: numbers < text < logicals < errors (reversed when descending); blanks always last. */
export function sortCompare(a: Scalar, b: Scalar, order: number): number {
  const ba = a == null, bb = b == null;
  if (ba || bb) return ba && bb ? 0 : ba ? 1 : -1;
  const ra = typeRank(a), rb = typeRank(b);
  if (ra !== rb) return (ra - rb) * order;
  if (ra === 3) return 0;
  return compare(a, b) * order;
}

/** Key for UNIQUE/grouping: type-aware and case-insensitive for text. */
function keyOf(v: Scalar): string {
  if (v == null) return 'b';
  if (typeof v === 'number') return 'n' + v;
  if (typeof v === 'string') return 's' + v.toLowerCase();
  if (typeof v === 'boolean') return 'l' + v;
  return 'e' + v.code;
}
const rowKey = (r: Scalar[]) => r.map(keyOf).join('\u0001');

// --------------------------------------------------------------------------------------------
// GROUPBY / PIVOTBY

/** Aggregation function from a LAMBDA, or an eta-reduced function name (which arrives as #NAME?). */
function aggregator(v: Value | undefined, ctx: Ctx): (m: Matrix) => Scalar {
  let fn: LambdaV | null = null;
  if (v && isLambda(v)) fn = v;
  else if (isErr(v as Scalar) && (v as any).code === '#NAME?' && (v as any).detail && getFn((v as any).detail)) {
    fn = { k: 'lambda', params: ['_X'], body: { t: 'fn', name: String((v as any).detail).toUpperCase(), args: [{ t: 'name', name: '_X' }] }, scope: null };
  }
  if (!fn) return fail(isErr(v as Scalar) ? (v as Scalar as any) : ERR.value);
  const f = fn;
  return (m) => {
    const r = ctx.callLambda(f, [m]);
    if (isLambda(r)) return ERR.calc;
    if (isRefs(r)) return ERR.value;
    if (isMatrix(r) || isRef(r)) {
      const mm = toMatrix(r, ctx);
      if (mm.length !== 1 || mm[0].length !== 1) return ERR.calc;
      return mm[0][0];
    }
    return r as Scalar;
  };
}

interface Prepared { fields: Matrix; values: Matrix; fieldHdr: Scalar[]; valueHdr: Scalar[]; show: boolean; filter: boolean[] | null }

function prepare(fieldsV: Value, valuesV: Value, headersV: Value | undefined, filterV: Value | undefined, ctx: Ctx, extra: Matrix[] = []): Prepared & { extra: Matrix[]; extraHdr: Scalar[][] } {
  let fields = mat(fieldsV, ctx), values = mat(valuesV, ctx);
  if (fields.length !== values.length || extra.some((e) => e.length !== values.length)) return fail(ERR.value);
  const n0 = values.length;
  let fh = sc(headersV, ctx);
  let has: boolean, show: boolean;
  if (fh === undefined || fh === null) {
    has = typeof values[0][0] === 'string' && values.length > 1 && typeof values[1][0] === 'number';
    show = has && (fields[0].length > 1 || extra.some((e) => e[0].length > 1));
  } else {
    const h = int(fh);
    if (h < 0 || h > 3) return fail(ERR.value);
    has = h === 1 || h === 3;
    show = h === 2 || h === 3;
  }
  let fieldHdr: Scalar[], valueHdr: Scalar[], extraHdr: Scalar[][];
  if (has) {
    fieldHdr = fields[0]; valueHdr = values[0]; extraHdr = extra.map((e) => e[0]);
    fields = fields.slice(1); values = values.slice(1); extra = extra.map((e) => e.slice(1));
  } else {
    fieldHdr = fields[0].map((_, i) => `Row Field ${i + 1}`);
    valueHdr = values[0].map((_, i) => `Value ${i + 1}`);
    extraHdr = extra.map((e) => e[0].map((_, i) => `Col Field ${i + 1}`));
  }
  let filter: boolean[] | null = null;
  if (filterV !== undefined && filterV !== null) {
    let fm = mat(filterV, ctx);
    if (fm[0].length !== 1) return fail(ERR.value);
    if (fm.length === n0 && has) fm = fm.slice(1);
    if (fm.length !== values.length) return fail(ERR.value);
    filter = fm.map((r) => { const x = r[0]; if (isErr(x)) fail(x); return x == null ? false : bool(x); });
  }
  return { fields, values, fieldHdr, valueHdr, show, filter, extra, extraHdr };
}

interface Group { key: Scalar[]; rows: number[] }

function groupRows(fields: Matrix, idx: number[]): Group[] {
  const map = new Map<string, Group>();
  for (const i of idx) {
    const k = rowKey(fields[i]);
    let g = map.get(k);
    if (!g) { g = { key: fields[i], rows: [] }; map.set(k, g); }
    g.rows.push(i);
  }
  return [...map.values()];
}

const colOf = (m: Matrix, rows: number[], c: number): Matrix => rows.map((i) => [m[i][c]]);

function groupBy(args: Value[], ctx: Ctx): Value {
  const [rfV, valV, fnV, hdrV, depthV, sortV, filterV, relV] = args;
  const agg = aggregator(fnV, ctx);
  const p = prepare(rfV, valV, hdrV, filterV, ctx);
  const depth = optInt(depthV, ctx, 1);
  const rel = optInt(relV, ctx, 0);
  if (rel !== 0 && rel !== 1) return ERR.value;
  const nf = p.fields[0].length, nv = p.values[0].length;
  const idx = p.values.map((_, i) => i).filter((i) => !p.filter || p.filter[i]);
  if (!idx.length) return ERR.calc;
  const groups = groupRows(p.fields, idx).map((g) => ({ ...g, aggs: Array.from({ length: nv }, (_, c) => agg(colOf(p.values, g.rows, c))) }));
  // Sorting
  const sortKeys = sortV === undefined || sc(sortV, ctx) === null ? Array.from({ length: nf }, (_, i) => i + 1) : intList(sortV, ctx);
  for (const k of sortKeys) if (k === 0 || Math.abs(k) > nf + nv) return ERR.value;
  const subLevels = rel === 0 ? Math.min(Math.abs(depth) - 1, nf - 1) : 0;
  const cmp = (a: typeof groups[0], b: typeof groups[0]): number => {
    // Keep hierarchy levels that carry subtotals contiguous.
    for (let l = 0; l < subLevels; l++) {
      const sk = sortKeys.find((k) => Math.abs(k) === l + 1);
      const c = sortCompare(a.key[l], b.key[l], sk && sk < 0 ? -1 : 1);
      if (c) return c;
    }
    for (const k of sortKeys) {
      const i = Math.abs(k) - 1;
      const x = i < nf ? a.key[i] : a.aggs[i - nf], y = i < nf ? b.key[i] : b.aggs[i - nf];
      const c = sortCompare(x, y, k < 0 ? -1 : 1);
      if (c) return c;
    }
    return 0;
  };
  groups.sort(cmp);
  const out: Matrix = [];
  if (p.show) out.push([...p.fieldHdr, ...p.valueHdr]);
  const top = depth < 0;
  const totalRow = (): Scalar[] => ['Total', ...Array(nf - 1).fill(null), ...Array.from({ length: nv }, (_, c) => agg(colOf(p.values, idx, c)))];
  if (top && depth !== 0) out.push(totalRow());
  const emit = (gs: typeof groups, level: number) => {
    if (level >= subLevels) { for (const g of gs) out.push([...g.key, ...g.aggs]); return; }
    // Split into contiguous runs by the key at this level.
    let i = 0;
    while (i < gs.length) {
      let j = i;
      while (j < gs.length && keyOf(gs[j].key[level]) === keyOf(gs[i].key[level])) j++;
      const run = gs.slice(i, j);
      const rows = run.flatMap((g) => g.rows);
      const sub: Scalar[] = [...gs[i].key.slice(0, level + 1), ...Array(nf - level - 1).fill(null), ...Array.from({ length: nv }, (_, c) => agg(colOf(p.values, rows, c)))];
      if (top) out.push(sub);
      emit(run, level + 1);
      if (!top) out.push(sub);
      i = j;
    }
  };
  emit(groups, 0);
  if (!top && depth !== 0) out.push(totalRow());
  return out;
}

function pivotBy(args: Value[], ctx: Ctx): Value {
  const [rfV, cfV, valV, fnV, hdrV, rDepthV, rSortV, cDepthV, cSortV, filterV, relV] = args;
  const agg = aggregator(fnV, ctx);
  const cfm = mat(cfV, ctx);
  const p = prepare(rfV, valV, hdrV, filterV, ctx, [cfm]);
  const cf = p.extra[0];
  const rDepth = optInt(rDepthV, ctx, 1), cDepth = optInt(cDepthV, ctx, 1);
  const rel = optInt(relV, ctx, 0);
  if (rel !== 0 && rel !== 1) return ERR.value;
  const nrf = p.fields[0].length, ncf = cf[0].length, nv = p.values[0].length;
  const idx = p.values.map((_, i) => i).filter((i) => !p.filter || p.filter[i]);
  if (!idx.length) return ERR.calc;
  const sortGroups = (gs: Group[], spec: Value | undefined, nk: number) => {
    const keys = spec === undefined || sc(spec, ctx) === null ? Array.from({ length: nk }, (_, i) => i + 1) : intList(spec, ctx);
    for (const k of keys) if (k === 0 || Math.abs(k) > nk) fail(ERR.value);
    gs.sort((a, b) => { for (const k of keys) { const c = sortCompare(a.key[Math.abs(k) - 1], b.key[Math.abs(k) - 1], k < 0 ? -1 : 1); if (c) return c; } return 0; });
    return gs;
  };
  const rgroups = sortGroups(groupRows(p.fields, idx), rSortV, nrf);
  const cgroups = sortGroups(groupRows(cf, idx), cSortV, ncf);
  const inCol = new Map<number, number>();
  cgroups.forEach((g, gi) => g.rows.forEach((i) => inCol.set(i, gi)));
  const cell = (rows: number[], c: number): Scalar => (rows.length ? agg(colOf(p.values, rows, c)) : null);
  const colTotals = cDepth !== 0, rowTotals = rDepth !== 0;
  const colBlocks: { label: Scalar[]; rows: (r: number[]) => number[] }[] = cgroups.map((g, gi) => ({ label: g.key, rows: (rs: number[]) => rs.filter((i) => inCol.get(i) === gi) }));
  const totalBlock = { label: ['Total', ...Array(ncf - 1).fill(null)] as Scalar[], rows: (rs: number[]) => rs };
  const blocks = colTotals ? (cDepth < 0 ? [totalBlock, ...colBlocks] : [...colBlocks, totalBlock]) : colBlocks;
  const out: Matrix = [];
  for (let l = 0; l < ncf; l++) {
    const row: Scalar[] = Array(nrf).fill(null);
    for (const b of blocks) for (let v = 0; v < nv; v++) row.push(v === 0 ? b.label[l] : null);
    out.push(row);
  }
  if (nv > 1 || p.show) {
    const row: Scalar[] = p.show ? [...p.fieldHdr] : Array(nrf).fill(null);
    for (let b = 0; b < blocks.length; b++) for (let v = 0; v < nv; v++) row.push(nv > 1 || p.show ? p.valueHdr[v] : null);
    out.push(row);
  }
  const dataRow = (label: Scalar[], rows: number[]): Scalar[] => {
    const row = [...label];
    for (const b of blocks) { const rs = b.rows(rows); for (let v = 0; v < nv; v++) row.push(cell(rs, v)); }
    return row;
  };
  const totalR = () => dataRow(['Total', ...Array(nrf - 1).fill(null)], idx);
  if (rowTotals && rDepth < 0) out.push(totalR());
  for (let gi = 0; gi < rgroups.length; gi++) out.push(dataRow(rgroups[gi].key, rgroups[gi].rows));
  if (rowTotals && rDepth > 0) out.push(totalR());
  return out;
}

// --------------------------------------------------------------------------------------------

function chooseLines(args: Value[], ctx: Ctx, byCol: boolean): Value {
  let m = mat(args[0], ctx);
  if (byCol) m = transpose(m);
  const n = m.length;
  const out: Matrix = [];
  for (const a of args.slice(1)) {
    for (const k of intList(a, ctx)) {
      if (k === 0 || Math.abs(k) > n) return ERR.value;
      out.push(m[k > 0 ? k - 1 : n + k].slice());
    }
  }
  return byCol ? transpose(out) : out;
}

function stack(args: Value[], ctx: Ctx, horizontal: boolean): Value {
  const mats = args.map((a) => mat(a ?? null, ctx)).map((m) => (horizontal ? transpose(m) : m));
  const w = Math.max(...mats.map((m) => m[0].length));
  const out: Matrix = [];
  for (const m of mats) for (const r of m) { const row = r.slice(); while (row.length < w) row.push(ERR.na); out.push(row); }
  return horizontal ? transpose(out) : out;
}

function toLine(args: Value[], ctx: Ctx): Scalar[] {
  const m = mat(args[0], ctx);
  const ignore = optInt(args[1], ctx, 0);
  if (ignore < 0 || ignore > 3) return fail(ERR.value);
  const byCol = optBool(args[2], ctx, false);
  const src = byCol ? transpose(m) : m;
  const out: Scalar[] = [];
  for (const r of src) for (const x of r) {
    if ((ignore & 1) && x == null) continue;
    if ((ignore & 2) && isErr(x)) continue;
    out.push(x);
  }
  if (!out.length) return fail(ERR.calc);
  return out;
}

function wrap(args: Value[], ctx: Ctx, byCols: boolean): Value {
  const m = mat(args[0], ctx);
  if (m.length !== 1 && m[0].length !== 1) return ERR.value;
  const items = flat(m);
  const n = int(S(sc(args[1], ctx) ?? null));
  if (n < 1) return ERR.num;
  const padV: Scalar = args[2] === undefined ? ERR.na : (sc(args[2], ctx) ?? null);
  const lines: Matrix = [];
  for (let i = 0; i < items.length; i += n) lines.push(items.slice(i, i + n));
  const padded = pad(lines, lines.length, n, padV);
  return byCols ? transpose(padded) : padded;
}

function trimRange(args: Value[], ctx: Ctx): Value {
  const [rv, trV, tcV] = args;
  const tr = optInt(trV, ctx, 3), tc = optInt(tcV, ctx, 3);
  if (tr < 0 || tr > 3 || tc < 0 || tc > 3) return ERR.value;
  if (isRef(rv)) {
    const b = ctx.usedBounds(rv.sheet);
    let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
    rv.sheet.forEachIn({ r1: rv.r1, c1: rv.c1, r2: Math.min(rv.r2, b.r2), c2: Math.min(rv.c2, b.c2) }, (r, c, cell) => {
      if (cell.v == null && cell.f == null) return;
      if (r < r1) r1 = r; if (r > r2) r2 = r; if (c < c1) c1 = c; if (c > c2) c2 = c;
    });
    if (r2 < 0) return ERR.ref;
    const out: RefV = { k: 'ref', sheet: rv.sheet, r1: tr & 1 ? r1 : rv.r1, r2: tr & 2 ? r2 : rv.r2, c1: tc & 1 ? c1 : rv.c1, c2: tc & 2 ? c2 : rv.c2 };
    return out;
  }
  const m = mat(rv, ctx);
  let r1 = 0, r2 = m.length - 1, c1 = 0, c2 = m[0].length - 1;
  const rowBlank = (r: number) => m[r].every((x) => x == null);
  const colBlank = (c: number) => m.every((row) => row[c] == null);
  if (tr & 1) while (r1 <= r2 && rowBlank(r1)) r1++;
  if (tr & 2) while (r2 >= r1 && rowBlank(r2)) r2--;
  if (tc & 1) while (c1 <= c2 && colBlank(c1)) c1++;
  if (tc & 2) while (c2 >= c1 && colBlank(c2)) c2--;
  if (r1 > r2 || c1 > c2) return ERR.ref;
  return m.slice(r1, r2 + 1).map((r) => r.slice(c1, c2 + 1));
}

defFn([
  {
    name: 'SEQUENCE', min: 1, max: 4, category: 'Dynamic Array', desc: 'Generates an array of sequential numbers.', params: ['rows', '[columns]', '[start]', '[step]'],
    impl: ([r, c, s, st]) => {
      const rows = r === null ? 1 : int(S(r)), cols = c === undefined || c === null ? 1 : int(S(c));
      const start = s === undefined || s === null ? 1 : num(S(s)), step = st === undefined || st === null ? 1 : num(S(st));
      if (rows < 0 || cols < 0) return ERR.value;
      if (rows === 0 || cols === 0) return ERR.calc;
      if (rows > MAX_ROWS || cols > MAX_COLS || rows * cols > MAX_CELLS) return ERR.num;
      const out: Matrix = [];
      for (let i = 0; i < rows; i++) { const row: Scalar[] = []; for (let j = 0; j < cols; j++) row.push(start + (i * cols + j) * step); out.push(row); }
      return out;
    },
  },
  {
    name: 'FILTER', min: 2, max: 3, category: 'Dynamic Array', desc: 'Keeps the rows or columns of an array that meet a condition.', params: ['array', 'include', '[if_empty]'], arrays: [0, 1, 2],
    impl: ([a, inc, ie], ctx) => {
      const m = mat(a, ctx), im = mat(inc, ctx);
      const rows = m.length, cols = m[0].length;
      let byRow: boolean;
      if (im[0].length === 1 && im.length === rows) byRow = true;
      else if (im.length === 1 && im[0].length === cols) byRow = false;
      else return ERR.value;
      const flags = flat(im).map((x) => {
        if (isErr(x)) fail(x);
        if (x == null) return false;
        if (typeof x === 'string') return fail(ERR.value);
        return bool(x);
      });
      let out: Matrix;
      if (byRow) out = m.filter((_, i) => flags[i]);
      else out = m.map((r) => r.filter((_, j) => flags[j]));
      if (!out.length || !out[0].length) {
        if (ie === undefined || ie === null) return ERR.calc;
        return ie;
      }
      return out;
    },
  },
  {
    name: 'SORT', min: 1, max: 4, category: 'Dynamic Array', desc: 'Sorts the rows (or columns) of an array.', params: ['array', '[sort_index]', '[sort_order]', '[by_col]'], arrays: [0, 1, 2],
    impl: ([a, si, so, bc], ctx) => {
      const byCol = optBool(bc, ctx, false);
      let m = mat(a, ctx);
      if (byCol) m = transpose(m);
      const idx = si === undefined || sc(si, ctx) === null && !isMatrix(si) ? [1] : intList(si, ctx);
      const ord = so === undefined || sc(so, ctx) === null && !isMatrix(so) ? [1] : intList(so, ctx);
      const width = m[0].length;
      for (const k of idx) if (k < 1 || k > width) return ERR.value;
      for (const o of ord) if (o !== 1 && o !== -1) return ERR.value;
      if (ord.length > 1 && ord.length !== idx.length) return ERR.value;
      const order = (i: number) => ord[ord.length === 1 ? 0 : i];
      const sorted = m.map((r, i) => ({ r, i })).sort((x, y) => {
        for (let k = 0; k < idx.length; k++) { const c = sortCompare(x.r[idx[k] - 1], y.r[idx[k] - 1], order(k)); if (c) return c; }
        return x.i - y.i;
      }).map((x) => x.r);
      return byCol ? transpose(sorted) : sorted;
    },
  },
  {
    name: 'SORTBY', min: 2, max: 253, category: 'Dynamic Array', desc: 'Sorts an array by the values of other arrays.', params: ['array', 'by_array1', '[sort_order1]', '...'], arrays: 'all',
    impl: (args, ctx) => {
      const m = mat(args[0], ctx);
      const rows = m.length, cols = m[0].length;
      const keys: { vals: Scalar[]; order: number }[] = [];
      let dir: 'row' | 'col' | null = null;
      for (let i = 1; i < args.length; i += 2) {
        const bm = mat(args[i], ctx);
        let d: 'row' | 'col';
        if (bm[0].length === 1 && bm.length === rows && !(bm.length === 1 && dir === 'col')) d = 'row';
        else if (bm.length === 1 && bm[0].length === cols) d = 'col';
        else return ERR.value;
        if (dir && dir !== d) return ERR.value;
        dir = d;
        const o = args[i + 1] === undefined ? 1 : optInt(args[i + 1], ctx, 1);
        if (o !== 1 && o !== -1) return ERR.value;
        keys.push({ vals: flat(bm), order: o });
      }
      const n = dir === 'col' ? cols : rows;
      const perm = Array.from({ length: n }, (_, i) => i).sort((x, y) => {
        for (const k of keys) { const c = sortCompare(k.vals[x], k.vals[y], k.order); if (c) return c; }
        return x - y;
      });
      if (dir === 'col') return m.map((r) => perm.map((j) => r[j]));
      return perm.map((i) => m[i].slice());
    },
  },
  {
    name: 'UNIQUE', min: 1, max: 3, category: 'Dynamic Array', desc: 'Returns the distinct rows or columns of an array.', params: ['array', '[by_col]', '[exactly_once]'], arrays: [0],
    impl: ([a, bc, eo], ctx) => {
      const byCol = optBool(bc, ctx, false), once = optBool(eo, ctx, false);
      let m = mat(a, ctx);
      if (byCol) m = transpose(m);
      const counts = new Map<string, { row: Scalar[]; n: number }>();
      for (const r of m) { const k = rowKey(r); const e = counts.get(k); if (e) e.n++; else counts.set(k, { row: r, n: 1 }); }
      const out = [...counts.values()].filter((e) => !once || e.n === 1).map((e) => e.row.slice());
      if (!out.length) return ERR.calc;
      return byCol ? transpose(out) : out;
    },
  },
  {
    name: 'TAKE', min: 2, max: 3, category: 'Dynamic Array', desc: 'Returns rows or columns from the start or end of an array.', params: ['array', 'rows', '[columns]'], arrays: [0],
    impl: ([a, r, c], ctx) => {
      const m = mat(a, ctx);
      const nr = r === null ? null : int(S(r)), nc = c === undefined || c === null ? null : int(S(c));
      if (nr === 0 || nc === 0) return ERR.calc;
      let out = m;
      if (nr != null) out = nr > 0 ? out.slice(0, nr) : out.slice(Math.max(0, out.length + nr));
      if (nc != null) out = out.map((row) => (nc > 0 ? row.slice(0, nc) : row.slice(Math.max(0, row.length + nc))));
      return out.map((row) => row.slice());
    },
  },
  {
    name: 'DROP', min: 2, max: 3, category: 'Dynamic Array', desc: 'Removes rows or columns from the start or end of an array.', params: ['array', 'rows', '[columns]'], arrays: [0],
    impl: ([a, r, c], ctx) => {
      const m = mat(a, ctx);
      const nr = r === null ? 0 : int(S(r)), nc = c === undefined || c === null ? 0 : int(S(c));
      let out = nr >= 0 ? m.slice(nr) : m.slice(0, Math.max(0, m.length + nr));
      out = out.map((row) => (nc >= 0 ? row.slice(nc) : row.slice(0, Math.max(0, row.length + nc))));
      if (!out.length || !out[0].length) return ERR.calc;
      return out;
    },
  },
  {
    name: 'EXPAND', min: 2, max: 4, category: 'Dynamic Array', desc: 'Expands an array to the given size, padding new cells.', params: ['array', 'rows', '[columns]', '[pad_with]'], arrays: [0],
    impl: ([a, r, c, p], ctx) => {
      const m = mat(a, ctx);
      const rows = r === null ? m.length : int(S(r)), cols = c === undefined || c === null ? m[0].length : int(S(c));
      if (rows < m.length || cols < m[0].length) return ERR.value;
      if (rows > MAX_ROWS || cols > MAX_COLS) return ERR.num;
      return pad(m, rows, cols, p === undefined ? ERR.na : S(p));
    },
  },
  {
    name: 'CHOOSECOLS', min: 2, max: 254, category: 'Dynamic Array', desc: 'Returns the given columns of an array.', params: ['array', 'col_num1', '[col_num2]', '...'], arrays: 'all',
    impl: (args, ctx) => chooseLines(args, ctx, true),
  },
  {
    name: 'CHOOSEROWS', min: 2, max: 254, category: 'Dynamic Array', desc: 'Returns the given rows of an array.', params: ['array', 'row_num1', '[row_num2]', '...'], arrays: 'all',
    impl: (args, ctx) => chooseLines(args, ctx, false),
  },
  {
    name: 'HSTACK', min: 1, max: 254, category: 'Dynamic Array', desc: 'Joins arrays side by side.', params: ['array1', '[array2]', '...'], arrays: 'all',
    impl: (args, ctx) => stack(args, ctx, true),
  },
  {
    name: 'VSTACK', min: 1, max: 254, category: 'Dynamic Array', desc: 'Stacks arrays on top of each other.', params: ['array1', '[array2]', '...'], arrays: 'all',
    impl: (args, ctx) => stack(args, ctx, false),
  },
  {
    name: 'TOCOL', min: 1, max: 3, category: 'Dynamic Array', desc: 'Returns an array as a single column.', params: ['array', '[ignore]', '[scan_by_column]'], arrays: [0],
    impl: (args, ctx) => toLine(args, ctx).map((x) => [x]),
  },
  {
    name: 'TOROW', min: 1, max: 3, category: 'Dynamic Array', desc: 'Returns an array as a single row.', params: ['array', '[ignore]', '[scan_by_column]'], arrays: [0],
    impl: (args, ctx) => [toLine(args, ctx)],
  },
  {
    name: 'WRAPCOLS', min: 2, max: 3, category: 'Dynamic Array', desc: 'Wraps a row or column of values into columns of a given height.', params: ['vector', 'wrap_count', '[pad_with]'], arrays: [0],
    impl: (args, ctx) => wrap(args, ctx, true),
  },
  {
    name: 'WRAPROWS', min: 2, max: 3, category: 'Dynamic Array', desc: 'Wraps a row or column of values into rows of a given width.', params: ['vector', 'wrap_count', '[pad_with]'], arrays: [0],
    impl: (args, ctx) => wrap(args, ctx, false),
  },
  {
    name: 'TRANSPOSE', min: 1, max: 1, category: 'Dynamic Array', desc: 'Swaps the rows and columns of an array.', params: ['array'], arrays: [0],
    impl: ([a], ctx) => transpose(mat(a, ctx)),
  },
  {
    name: 'GROUPBY', min: 3, max: 8, category: 'Dynamic Array', desc: 'Groups rows by the values of fields and aggregates the related values.', params: ['row_fields', 'values', 'function', '[field_headers]', '[total_depth]', '[sort_order]', '[filter_array]', '[field_relationship]'], arrays: 'all',
    impl: (args, ctx) => groupBy(args, ctx),
  },
  {
    name: 'PIVOTBY', min: 4, max: 11, category: 'Dynamic Array', desc: 'Summarizes values by row and column fields into a cross-tabulation.', params: ['row_fields', 'col_fields', 'values', 'function', '[field_headers]', '[row_total_depth]', '[row_sort_order]', '[col_total_depth]', '[col_sort_order]', '[filter_array]', '[relationship]'], arrays: 'all',
    impl: (args, ctx) => pivotBy(args, ctx),
  },
  {
    name: 'TRIMRANGE', min: 1, max: 3, category: 'Dynamic Array', desc: 'Trims empty outer rows and columns from a range or array.', params: ['range', '[trim_rows]', '[trim_cols]'], refs: [0], returnsRef: true,
    impl: (args, ctx) => trimRange(args, ctx),
  },
]);

