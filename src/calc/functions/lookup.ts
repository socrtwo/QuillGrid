/** Lookup & Reference functions. */
import { defFn } from './registry';
import {
  fail, isErr, isLambda, isMatrix, isRef, isRefs, toMatrix, toScalar, num, str, int, bool, wildcard, compare, flat, dims,
  ERR, type Ctx, type Value, type RefV,
} from '../values';
import type { Matrix, Scalar } from '../../model/types';
import { MAX_COLS, MAX_ROWS, colName, quoteSheet } from '../../model/address';

const S = (a: Value) => a as Scalar;

/** Scalar from an argument that may be a reference or array. */
export function sc(v: Value | undefined, ctx: Ctx): Scalar | undefined {
  if (v === undefined) return undefined;
  if (isRef(v) || isMatrix(v)) return toScalar(v, ctx);
  if (isRefs(v)) return fail(ERR.value);
  if (isLambda(v)) return fail(ERR.calc);
  return v;
}

// --------------------------------------------------------------------------------------------
// Lookup engine shared by MATCH, VLOOKUP, HLOOKUP, LOOKUP, XLOOKUP and XMATCH

type Kind = 0 | 1 | 2 | 3; // number, text, boolean, other
const kindOf = (v: Scalar): Kind | -1 => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : -1);

/** Comparison of same-kind values (text case-insensitive). */
function cmpSame(a: Scalar, b: Scalar): number { return compare(a, b); }

/** Exact equality used by lookups: same kind, text case-insensitive. */
function eqLookup(x: Scalar, v: Scalar): boolean {
  if (v == null) return false;
  return kindOf(x) === kindOf(v) && kindOf(v) >= 0 && cmpSame(x, v) === 0;
}

function exactMatcher(v: Scalar, wild: boolean): (x: Scalar) => boolean {
  if (wild && typeof v === 'string' && /[*?~]/.test(v)) {
    const re = wildcard(v);
    return (x) => typeof x === 'string' && re.test(x);
  }
  return (x) => eqLookup(x, v);
}

/**
 * Spreadsheet binary search. asc: returns the last index whose value is <= v; desc: the last index
 * whose value is >= v. Values of a different type (and blanks/errors) are skipped by probing towards
 * the lower bound, as spreadsheets do. Returns -1 when no such index exists.
 */
export function bsearch(list: Scalar[], v: Scalar, desc = false): number {
  const k = kindOf(v);
  if (k < 0) return -1;
  let lo = 0, hi = list.length - 1, ans = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let j = mid;
    while (j >= lo && kindOf(list[j]) !== k) j--;
    if (j < lo) { lo = mid + 1; continue; }
    const c = cmpSame(list[j], v);
    if (desc ? c >= 0 : c <= 0) { ans = j; lo = mid + 1; } else hi = j - 1;
  }
  return ans;
}

/** Next index after i holding a value of the same kind as v (for "next larger" after a binary search). */
function nextSameKind(list: Scalar[], i: number, v: Scalar): number {
  const k = kindOf(v);
  for (let j = i + 1; j < list.length; j++) if (kindOf(list[j]) === k) return j;
  return -1;
}

/**
 * XLOOKUP/XMATCH search. matchMode: 0 exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcard.
 * searchMode: 1 first-to-last, -1 last-to-first, 2 binary ascending, -2 binary descending.
 */
export function xfind(list: Scalar[], v: Scalar, matchMode: number, searchMode: number): number {
  if (![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) return fail(ERR.value);
  if (Math.abs(searchMode) === 2) {
    if (matchMode === 2) return fail(ERR.value);
    const desc = searchMode === -2;
    const i = bsearch(list, v, desc);
    const exact = i >= 0 && eqLookup(list[i], v);
    if (exact) return i;
    if (matchMode === 0) return -1;
    if (!desc) return matchMode === -1 ? i : nextSameKind(list, i, v);
    return matchMode === 1 ? i : nextSameKind(list, i, v);
  }
  const n = list.length;
  const order = searchMode === 1 ? (i: number) => i : (i: number) => n - 1 - i;
  const match = exactMatcher(v, matchMode === 2);
  let best = -1;
  const k = kindOf(v);
  for (let t = 0; t < n; t++) {
    const i = order(t);
    const x = list[i];
    if (match(x)) return i;
    if (matchMode === -1 || matchMode === 1) {
      if (kindOf(x) !== k || k < 0) continue;
      const c = cmpSame(x, v);
      if (matchMode === -1 && c < 0 && (best < 0 || cmpSame(x, list[best]) > 0)) best = i;
      if (matchMode === 1 && c > 0 && (best < 0 || cmpSame(x, list[best]) < 0)) best = i;
    }
  }
  return best;
}

/** MATCH-style search: 0 exact (wildcards), 1 largest <= (sorted ascending), -1 smallest >= (sorted descending). */
export function matchIndex(list: Scalar[], v: Scalar, type: number): number {
  if (type === 0) {
    const m = exactMatcher(v, true);
    return list.findIndex(m);
  }
  return bsearch(list, v, type < 0);
}

/** Lookup value from an argument (errors propagate). */
function lookupValue(v: Value | undefined, ctx: Ctx): Scalar {
  const x = sc(v, ctx) ?? null;
  if (isErr(x)) fail(x);
  return x;
}

// --------------------------------------------------------------------------------------------
// References

/** A sub-range of a reference, relative to its top-left corner. */
function subRef(ref: RefV, r0: number, c0: number, rows: number, cols: number): RefV {
  return { k: 'ref', sheet: ref.sheet, r1: ref.r1 + r0, c1: ref.c1 + c0, r2: ref.r1 + r0 + rows - 1, c2: ref.c1 + c0 + cols - 1 };
}

/** Rows/cols of a reference, clipping whole-column/row references to the used range for reading. */
function refDims(ref: RefV): { rows: number; cols: number } { return { rows: ref.r2 - ref.r1 + 1, cols: ref.c2 - ref.c1 + 1 }; }

function indexImpl(args: Value[], ctx: Ctx): Value {
  const [arr, rn, cn, an] = args;
  const rowGiven = rn !== undefined, colGiven = cn !== undefined;
  const rowN = rowGiven ? int(S(rn)) : 0;
  const colN = colGiven ? int(S(cn)) : 0;
  if (rowN < 0 || colN < 0) return ERR.value;
  // Reference form
  if (isRef(arr) || isRefs(arr)) {
    const areas = isRefs(arr) ? arr.areas : [arr];
    const area = an === undefined || an === null ? 1 : int(S(an));
    if (area < 1 || area > areas.length) return ERR.ref;
    const ref = areas[area - 1];
    const { rows, cols } = refDims(ref);
    let r = rowN, c = colN;
    if (!colGiven && rows === 1 && cols > 1) { c = rowN; r = 0; if (c === 0) return ref; return c > cols ? ERR.ref : subRef(ref, 0, c - 1, 1, 1); }
    if (r > rows || c > cols) return ERR.ref;
    if (r === 0 && c === 0) return ref;
    if (r === 0) return subRef(ref, 0, c - 1, rows, 1);
    if (c === 0) return cols === 1 ? subRef(ref, r - 1, 0, 1, 1) : subRef(ref, r - 1, 0, 1, cols);
    return subRef(ref, r - 1, c - 1, 1, 1);
  }
  // Array form
  if (an !== undefined && an !== null && int(S(an)) !== 1) return ERR.ref;
  const m = toMatrix(arr ?? null, ctx);
  const rows = m.length, cols = m[0]?.length ?? 0;
  let r = rowN, c = colN;
  if (!colGiven && rows === 1) { c = rowN; r = 1; if (c === 0) return m; }
  else if (!colGiven && cols === 1) c = 1;
  if (r > rows || c > cols) return ERR.ref;
  if (r === 0 && c === 0) return m;
  if (r === 0) return m.map((row) => [row[c - 1]]);
  if (c === 0) return [m[r - 1].slice()];
  return m[r - 1][c - 1];
}

/** Reads one cell of a reference with extension past its end in the reference's orientation (LOOKUP result_vector). */
function vectorAt(v: Value, i: number, ctx: Ctx): Scalar {
  if (isRef(v)) {
    const { rows, cols } = refDims(v);
    if (rows === 1) return v.c1 + i < MAX_COLS ? ctx.cell(v.sheet, v.r1, v.c1 + i) : ERR.ref;
    if (cols === 1) return v.r1 + i < MAX_ROWS ? ctx.cell(v.sheet, v.r1 + i, v.c1) : ERR.ref;
  }
  const items = flat(toMatrix(v, ctx));
  return i < items.length ? items[i] : ERR.na;
}

function column(m: Matrix, c: number): Scalar[] { return m.map((r) => r[c]); }

/** Lookup vector from a 1-D range/array; null when the argument is two-dimensional. */
function vector(v: Value, ctx: Ctx): { list: Scalar[]; horizontal: boolean } | null {
  const d = isRef(v) ? refDims(v) : dims(v);
  if (d.rows !== 1 && d.cols !== 1) return null;
  const m = toMatrix(v, ctx);
  if (d.rows === 1) return { list: (m[0] ?? []).slice(), horizontal: true };
  return { list: column(m, 0), horizontal: false };
}

function xlookup(args: Value[], ctx: Ctx): Value {
  const [lv, la, ra, nf, mm, sm] = args;
  const v = sc(lv, ctx) ?? null;
  if (isErr(v)) return v;
  const matchMode = mm === undefined || sc(mm, ctx) == null ? 0 : int(sc(mm, ctx)!);
  const searchMode = sm === undefined || sc(sm, ctx) == null ? 1 : int(sc(sm, ctx)!);
  if (isRefs(la) || isRefs(ra) || isLambda(la) || isLambda(ra)) return ERR.value;
  const vec = vector(la, ctx);
  if (!vec) return ERR.value;
  const n = vec.list.length;
  // The return array must match the lookup vector along the lookup direction (uncliped shape for refs).
  const full = isRef(ra) ? refDims(ra) : dims(ra);
  const laFull = isRef(la) ? refDims(la) : dims(la);
  const lookN = vec.horizontal ? laFull.cols : laFull.rows;
  if ((vec.horizontal ? full.cols : full.rows) !== lookN) return ERR.value;
  const i = xfind(vec.list, v, matchMode, searchMode);
  if (i < 0 || i >= n) {
    if (nf === undefined || nf === null) return ERR.na;
    return nf;
  }
  if (isRef(ra)) return vec.horizontal ? subRef(ra, 0, i, full.rows, 1) : subRef(ra, i, 0, 1, full.cols);
  const m = toMatrix(ra, ctx);
  if (vec.horizontal) return m.length === 1 ? m[0][i] : m.map((row) => [row[i]]);
  return m[i].length === 1 ? m[i][0] : [m[i].slice()];
}

function a1Address(r: number, c: number, abs: number): string {
  const ar = abs === 1 || abs === 2, ac = abs === 1 || abs === 3;
  return `${ac ? '$' : ''}${colName(c)}${ar ? '$' : ''}${r + 1}`;
}

function r1c1Address(row: number, col: number, abs: number): string {
  const ar = abs === 1 || abs === 2, ac = abs === 1 || abs === 3;
  return `R${ar ? row : `[${row}]`}C${ac ? col : `[${col}]`}`;
}

function lookupTable(args: Value[], ctx: Ctx, horizontal: boolean): Value {
  const [lv, tbl, idx, rl] = args;
  const v = lookupValue(lv, ctx);
  const k = int(S(idx));
  if (k < 1) return ERR.value;
  if (isRefs(tbl) || isLambda(tbl)) return ERR.value;
  const approx = rl === undefined ? true : bool(S(rl));
  const m = toMatrix(tbl, ctx);
  const size = horizontal ? m.length : (m[0]?.length ?? 0);
  const full = isRef(tbl) ? refDims(tbl) : dims(tbl);
  if (k > (horizontal ? full.rows : full.cols)) return ERR.ref;
  const keys = horizontal ? (m[0] ?? []).slice() : column(m, 0);
  const i = approx ? bsearch(keys, v) : matchIndex(keys, v, 0);
  if (i < 0) return ERR.na;
  if (k > size) return null;
  return horizontal ? (m[k - 1]?.[i] ?? null) : (m[i]?.[k - 1] ?? null);
}

defFn([
  {
    name: 'ADDRESS', min: 2, max: 5, category: 'Lookup & Reference', desc: 'Builds a cell address as text from row and column numbers.', params: ['row_num', 'column_num', '[abs_num]', '[a1]', '[sheet_text]'],
    impl: ([r, c, a, s, sh]) => {
      const row = int(S(r)), col = int(S(c));
      const abs = a === undefined || a === null ? 1 : int(S(a));
      const a1 = s === undefined || s === null ? true : bool(S(s));
      if (abs < 1 || abs > 4) return ERR.value;
      const absR = abs === 1 || abs === 2, absC = abs === 1 || abs === 3;
      if (a1 || absR) { if (row < 1 || row > MAX_ROWS) return ERR.value; } else if (Math.abs(row) >= MAX_ROWS) return ERR.value;
      if (a1 || absC) { if (col < 1 || col > MAX_COLS) return ERR.value; } else if (Math.abs(col) >= MAX_COLS) return ERR.value;
      const addr = a1 ? a1Address(row - 1, col - 1, abs) : r1c1Address(row, col, abs);
      if (sh === undefined || sh === null) return addr;
      const name = str(S(sh));
      return `${name === '' ? '' : quoteSheet(name)}!${addr}`;
    },
  },
  {
    name: 'AREAS', min: 1, max: 1, category: 'Lookup & Reference', desc: 'Number of areas in a reference.', params: ['reference'], refs: [0],
    impl: ([r]) => (isRefs(r) ? r.areas.length : isRef(r) ? 1 : isErr(r as Scalar) ? (r as Scalar) : ERR.value),
  },
  {
    name: 'COLUMN', min: 0, max: 1, category: 'Lookup & Reference', desc: 'Column number of a reference (or of the formula cell).', params: ['[reference]'], refs: [0],
    impl: ([r], ctx) => {
      if (r === undefined || r === null) return ctx.c + 1;
      if (isRefs(r)) return ERR.value;
      if (!isRef(r)) return isErr(r as Scalar) ? (r as Scalar) : ERR.value;
      if (r.c1 === r.c2) return r.c1 + 1;
      const row: Scalar[] = [];
      for (let c = r.c1; c <= r.c2; c++) row.push(c + 1);
      return [row];
    },
  },
  {
    name: 'ROW', min: 0, max: 1, category: 'Lookup & Reference', desc: 'Row number of a reference (or of the formula cell).', params: ['[reference]'], refs: [0],
    impl: ([r], ctx) => {
      if (r === undefined || r === null) return ctx.r + 1;
      if (isRefs(r)) return ERR.value;
      if (!isRef(r)) return isErr(r as Scalar) ? (r as Scalar) : ERR.value;
      if (r.r1 === r.r2) return r.r1 + 1;
      const out: Matrix = [];
      for (let x = r.r1; x <= r.r2; x++) out.push([x + 1]);
      return out;
    },
  },
  {
    name: 'COLUMNS', min: 1, max: 1, category: 'Lookup & Reference', desc: 'Number of columns in a reference or array.', params: ['array'], arrays: [0],
    impl: ([a]) => { if (isRefs(a)) return ERR.ref; if (isLambda(a)) return ERR.value; if (isErr(a as Scalar)) return a; return dims(a ?? null).cols; },
  },
  {
    name: 'ROWS', min: 1, max: 1, category: 'Lookup & Reference', desc: 'Number of rows in a reference or array.', params: ['array'], arrays: [0],
    impl: ([a]) => { if (isRefs(a)) return ERR.ref; if (isLambda(a)) return ERR.value; if (isErr(a as Scalar)) return a; return dims(a ?? null).rows; },
  },
  {
    name: 'FORMULATEXT', min: 1, max: 1, category: 'Lookup & Reference', desc: 'Returns the formula of a referenced cell as text.', params: ['reference'], refs: [0],
    impl: ([r]) => {
      if (isRefs(r)) r = r.areas[0];
      if (!isRef(r)) return isErr(r as Scalar) ? (r as Scalar) : ERR.na;
      const cell = r.sheet.get(r.r1, r.c1);
      if (cell?.f == null) return ERR.na;
      return '=' + cell.f;
    },
  },
  { name: 'GETPIVOTDATA', min: 2, max: 254, category: 'Lookup & Reference', desc: 'Returns data stored in a PivotTable report.', params: ['data_field', 'pivot_table', '[field1]', '[item1]', '...'], arrays: 'all', impl: () => ERR.ref },
  {
    name: 'HYPERLINK', min: 1, max: 2, category: 'Lookup & Reference', desc: 'Creates a link; the cell shows the friendly name (or the link text).', params: ['link_location', '[friendly_name]'],
    impl: ([l, f]) => {
      if (f !== undefined && f !== null) return S(f);
      const x = S(l);
      if (isErr(x)) return x;
      return str(x);
    },
  },
  {
    name: 'IMAGE', min: 1, max: 5, category: 'Lookup & Reference', desc: 'Places an image from a URL in the cell (shows the source text here).', params: ['source', '[alt_text]', '[sizing]', '[height]', '[width]'],
    impl: ([s]) => { const x = S(s); if (isErr(x)) return x; return str(x); },
  },
  {
    name: 'INDEX', min: 2, max: 4, category: 'Lookup & Reference', desc: 'Value or reference at a row and column of a range or array.', params: ['array', 'row_num', '[column_num]', '[area_num]'], arrays: [0], returnsRef: true,
    impl: (args, ctx) => indexImpl(args, ctx),
  },
  {
    name: 'INDIRECT', min: 1, max: 2, category: 'Lookup & Reference', desc: 'Reference given by a text string (A1 or R1C1 style).', params: ['ref_text', '[a1]'], volatile: true, returnsRef: true,
    impl: ([t, a], ctx) => {
      ctx.volatile();
      const x = S(t);
      if (isErr(x)) return x;
      const text = str(x).trim();
      if (!text) return ERR.ref;
      const a1 = a === undefined || a === null ? true : bool(S(a));
      const v = ctx.resolveRef(text, a1);
      if (isRef(v) || isRefs(v)) return v;
      return ERR.ref;
    },
  },
  {
    name: 'LOOKUP', min: 2, max: 3, category: 'Lookup & Reference', desc: 'Looks up a value in a sorted vector or array and returns the matching item.', params: ['lookup_value', 'lookup_vector', '[result_vector]'], arrays: [1, 2],
    impl: ([lv, lvec, rvec], ctx) => {
      const v = lookupValue(lv, ctx);
      if (isRefs(lvec) || isLambda(lvec)) return ERR.value;
      const m = toMatrix(lvec, ctx);
      const rows = m.length, cols = m[0]?.length ?? 0;
      if (rvec === undefined) {
        // Array form: search the first row (wide) or column (tall/square), return from the last.
        const wide = cols > rows;
        const keys = wide ? m[0] : column(m, 0);
        const i = bsearch(keys, v);
        if (i < 0) return ERR.na;
        return wide ? m[rows - 1][i] : m[i][cols - 1];
      }
      const keys = rows === 1 ? m[0] : cols === 1 ? column(m, 0) : null;
      if (!keys) return ERR.na;
      const i = bsearch(keys, v);
      if (i < 0) return ERR.na;
      if (isRefs(rvec) || isLambda(rvec)) return ERR.value;
      return vectorAt(rvec, i, ctx);
    },
  },
  {
    name: 'MATCH', min: 2, max: 3, category: 'Lookup & Reference', desc: 'Relative position of a value in a row or column.', params: ['lookup_value', 'lookup_array', '[match_type]'], arrays: [1],
    impl: ([lv, la, mt], ctx) => {
      const v = lookupValue(lv, ctx);
      if (isRefs(la) || isLambda(la)) return ERR.na;
      const vec = vector(la ?? null, ctx);
      if (!vec) return ERR.na;
      const t = mt === undefined ? 1 : num(S(mt));
      const type = t > 0 ? 1 : t < 0 ? -1 : 0;
      const i = matchIndex(vec.list, v, type);
      return i < 0 ? ERR.na : i + 1;
    },
  },
  {
    name: 'VLOOKUP', min: 3, max: 4, category: 'Lookup & Reference', desc: 'Looks for a value in the first column of a table and returns a value from the same row.', params: ['lookup_value', 'table_array', 'col_index_num', '[range_lookup]'], arrays: [1],
    impl: (args, ctx) => lookupTable(args, ctx, false),
  },
  {
    name: 'HLOOKUP', min: 3, max: 4, category: 'Lookup & Reference', desc: 'Looks for a value in the first row of a table and returns a value from the same column.', params: ['lookup_value', 'table_array', 'row_index_num', '[range_lookup]'], arrays: [1],
    impl: (args, ctx) => lookupTable(args, ctx, true),
  },
  {
    name: 'XLOOKUP', min: 3, max: 6, category: 'Lookup & Reference', desc: 'Searches a range or array and returns the item from a second range at the match.', params: ['lookup_value', 'lookup_array', 'return_array', '[if_not_found]', '[match_mode]', '[search_mode]'], arrays: [1, 2, 3], returnsRef: true,
    impl: (args, ctx) => xlookup(args, ctx),
  },
  {
    name: 'XMATCH', min: 2, max: 4, category: 'Lookup & Reference', desc: 'Relative position of a value in a row or column, with match and search modes.', params: ['lookup_value', 'lookup_array', '[match_mode]', '[search_mode]'], arrays: [1],
    impl: ([lv, la, mm, sm], ctx) => {
      const v = lookupValue(lv, ctx);
      if (isRefs(la) || isLambda(la)) return ERR.value;
      const vec = vector(la ?? null, ctx);
      if (!vec) return ERR.value;
      const matchMode = mm === undefined || mm === null ? 0 : int(S(mm));
      const searchMode = sm === undefined || sm === null ? 1 : int(S(sm));
      const i = xfind(vec.list, v, matchMode, searchMode);
      return i < 0 ? ERR.na : i + 1;
    },
  },
  {
    name: 'OFFSET', min: 3, max: 5, category: 'Lookup & Reference', desc: 'Reference shifted from a starting reference by rows and columns, with an optional size.', params: ['reference', 'rows', 'cols', '[height]', '[width]'], refs: [0], volatile: true, returnsRef: true,
    impl: ([ref, rs, cs, h, w], ctx) => {
      ctx.volatile();
      let base: RefV;
      if (isRef(ref)) base = ref;
      else if (isRefs(ref)) return ERR.value;
      else return isErr(ref as Scalar) ? (ref as Scalar) : ERR.value;
      const dr = int(S(rs)), dc = int(S(cs));
      const { rows, cols } = refDims(base);
      const hh = h === undefined || h === null ? rows : int(S(h));
      const ww = w === undefined || w === null ? cols : int(S(w));
      if (hh === 0 || ww === 0) return ERR.ref;
      let r1 = base.r1 + dr, c1 = base.c1 + dc;
      let r2 = r1 + hh - 1, c2 = c1 + ww - 1;
      if (hh < 0) { r2 = r1; r1 = r1 + hh + 1; }
      if (ww < 0) { c2 = c1; c1 = c1 + ww + 1; }
      if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) return ERR.ref;
      return { k: 'ref', sheet: base.sheet, r1, c1, r2, c2 } as RefV;
    },
  },
  { name: 'RTD', min: 3, max: 255, category: 'Lookup & Reference', desc: 'Retrieves real-time data from a COM automation server (not available).', params: ['progID', 'server', 'topic1', '[topic2]', '...'], impl: () => ERR.na },
]);

