/**
 * Model and data helpers behind the heritage commands:
 * - matrices (Lotus 1-2-3 /Data Matrix: invert and multiply, plus solving linear systems),
 * - Monte Carlo simulation over input cells (uncertain inputs as in Causal and risk add-ins),
 * - the input cells a result depends on (What-If Sliders),
 * - record conditions (1-2-3 /Data Query Delete), typed columns (Airtable / Grist field types),
 * - the keyboard menu tree and its accelerator letters (1-2-3's / menu).
 */
import type { Sheet, Workbook } from '../model/workbook';
import type { Engine } from '../calc/engine';
import type { Cell, Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellKey } from '../model/address';

// ---------------------------------------------------------------------------------------------
// Matrices

export type Mat = number[][];

/** Gauss-Jordan inverse with partial pivoting; null when singular. */
export function matInverse(a: Mat): Mat | null {
  const n = a.length;
  if (!n || a.some((row) => row.length !== n)) return null;
  const m = a.map((row, i) => [...row, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let col = 0; col < n; col++) {
    let piv = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(m[r][col]) > Math.abs(m[piv][col])) piv = r;
    if (Math.abs(m[piv][col]) < 1e-12) return null;
    [m[col], m[piv]] = [m[piv], m[col]];
    const p = m[col][col];
    for (let j = 0; j < 2 * n; j++) m[col][j] /= p;
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const f = m[r][col];
      if (f) for (let j = 0; j < 2 * n; j++) m[r][j] -= f * m[col][j];
    }
  }
  return m.map((row) => row.slice(n).map((x) => (Math.abs(x) < 1e-15 ? 0 : x)));
}

export function matMultiply(a: Mat, b: Mat): Mat | null {
  if (!a.length || !b.length || a[0].length !== b.length) return null;
  return a.map((row) => b[0].map((_, j) => row.reduce((s, x, k) => s + x * b[k][j], 0)));
}

/** Solves A·x = b (b may have several columns). */
export function matSolve(a: Mat, b: Mat): Mat | null {
  const inv = matInverse(a);
  return inv && b.length === a.length ? matMultiply(inv, b) : null;
}

export function matDeterminant(a: Mat): number | null {
  const n = a.length;
  if (!n || a.some((row) => row.length !== n)) return null;
  const m = a.map((r) => [...r]);
  let det = 1;
  for (let col = 0; col < n; col++) {
    let piv = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(m[r][col]) > Math.abs(m[piv][col])) piv = r;
    if (Math.abs(m[piv][col]) < 1e-15) return 0;
    if (piv !== col) { [m[col], m[piv]] = [m[piv], m[col]]; det = -det; }
    det *= m[col][col];
    for (let r = col + 1; r < n; r++) { const f = m[r][col] / m[col][col]; for (let j = col; j < n; j++) m[r][j] -= f * m[col][j]; }
  }
  return det;
}

// ---------------------------------------------------------------------------------------------
// Monte Carlo

export type Dist =
  | { kind: 'uniform'; min: number; max: number }
  | { kind: 'normal'; mean: number; sd: number }
  | { kind: 'triangular'; min: number; mode: number; max: number }
  | { kind: 'discrete'; values: number[] };

/** Seeded generator (mulberry32) so simulations can be repeated. */
export function seeded(seed = 1): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

export function sample(d: Dist, rnd: () => number): number {
  switch (d.kind) {
    case 'uniform': return d.min + (d.max - d.min) * rnd();
    case 'normal': {
      let u = 0; while (u === 0) u = rnd();
      return d.mean + d.sd * Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * rnd());
    }
    case 'triangular': {
      const { min: a, mode: c, max: b } = d;
      if (b <= a) return a;
      const u = rnd(), f = (c - a) / (b - a);
      return u < f ? a + Math.sqrt(u * (b - a) * (c - a)) : b - Math.sqrt((1 - u) * (b - a) * (b - c));
    }
    case 'discrete': return d.values.length ? d.values[Math.floor(rnd() * d.values.length)] : 0;
  }
}

/** Reads a distribution typed as text: "10 to 20" (uniform), "normal 100 15", "triangular 1 3 8", "1, 2, 5" (pick one). */
export function parseDist(text: string): Dist | null {
  const t = text.trim().toLowerCase();
  const nums = (t.match(/-?\d+(?:\.\d+)?(?:e[-+]?\d+)?/g) || []).map(Number);
  if (/^(normal|norm|gauss)/.test(t) && nums.length === 2 && nums[1] >= 0) return { kind: 'normal', mean: nums[0], sd: nums[1] };
  if (/^tri/.test(t) && nums.length === 3) { const [a, m, b] = nums; return a <= m && m <= b ? { kind: 'triangular', min: a, mode: m, max: b } : null; }
  if (/\bto\b|\.\.|–/.test(t) && nums.length === 2) return { kind: 'uniform', min: Math.min(nums[0], nums[1]), max: Math.max(nums[0], nums[1]) };
  if (/^(uniform|between)/.test(t) && nums.length === 2) return { kind: 'uniform', min: Math.min(nums[0], nums[1]), max: Math.max(nums[0], nums[1]) };
  if (/,/.test(t) && nums.length >= 2) return { kind: 'discrete', values: nums };
  return null;
}

export interface CellPos { sheet: Sheet; r: number; c: number }
export interface Summary { n: number; errors: number; mean: number; sd: number; min: number; max: number; p5: number; p25: number; p50: number; p75: number; p95: number }

export function percentile(sorted: number[], p: number): number {
  if (!sorted.length) return NaN;
  const i = (sorted.length - 1) * p, lo = Math.floor(i), hi = Math.ceil(i);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (i - lo);
}

export function summarize(xs: number[], errors = 0): Summary {
  const s = [...xs].sort((a, b) => a - b);
  const n = s.length;
  const mean = n ? s.reduce((a, b) => a + b, 0) / n : NaN;
  const sd = n > 1 ? Math.sqrt(s.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1)) : 0;
  return { n, errors, mean, sd, min: s[0] ?? NaN, max: s[n - 1] ?? NaN, p5: percentile(s, 0.05), p25: percentile(s, 0.25), p50: percentile(s, 0.5), p75: percentile(s, 0.75), p95: percentile(s, 0.95) };
}

/**
 * Runs the model `iterations` times: each input cell gets a value drawn from its distribution,
 * the workbook recalculates (volatile functions such as RAND draw again too), and the outputs are
 * recorded. The input cells are restored afterwards; nothing is added to the undo history.
 */
export function monteCarlo(book: Workbook, engine: Engine, inputs: (CellPos & { dist: Dist })[], outputs: CellPos[], iterations: number, rnd: () => number = Math.random): { values: number[][]; errors: number[] } {
  const saved = inputs.map((x) => x.sheet.get(x.r, x.c));
  const values = outputs.map(() => [] as number[]);
  const errors = outputs.map(() => 0);
  const full = book.calc.mode === 'manual';
  const put = (x: CellPos, cell: Cell | undefined) => {
    const k = cellKey(x.r, x.c);
    if (cell) x.sheet.cells.set(k, cell); else x.sheet.cells.delete(k);
    book.noteChange(x.sheet, k);
  };
  try {
    for (let i = 0; i < iterations; i++) {
      inputs.forEach((x, j) => put(x, { ...(saved[j]?.s ? { s: saved[j]!.s } : {}), v: sample(x.dist, rnd) }));
      engine.recalc(full);
      outputs.forEach((o, j) => {
        const v = o.sheet.value(o.r, o.c);
        if (typeof v === 'number' && Number.isFinite(v)) values[j].push(v); else errors[j]++;
      });
    }
  } finally {
    inputs.forEach((x, j) => put(x, saved[j]));
    engine.recalc(full);
  }
  return { values, errors };
}

/** Constant number cells that feed a formula, directly or through other formulas (ranges up to 64 cells are looked into). */
export function leafInputs(engine: Engine, sheet: Sheet, r: number, c: number, max = 24): CellPos[] {
  const out: CellPos[] = [];
  const seen = new Set<string>();
  const stack: CellPos[] = [{ sheet, r, c }];
  while (stack.length && out.length < max) {
    const x = stack.pop()!;
    const id = `${x.sheet.id}:${x.r}:${x.c}`;
    if (seen.has(id)) continue;
    seen.add(id);
    const cell = x.sheet.get(x.r, x.c);
    if (cell?.f == null) {
      if (typeof cell?.v === 'number' && !(x.sheet === sheet && x.r === r && x.c === c)) out.push(x);
      continue;
    }
    const precs = engine.precedentsOf(x.sheet, x.r, x.c);
    for (let i = precs.length - 1; i >= 0; i--) {
      const p = precs[i];
      if ((p.r2 - p.r1 + 1) * (p.c2 - p.c1 + 1) > 64) continue;
      for (let rr = p.r2; rr >= p.r1; rr--) for (let cc = p.c2; cc >= p.c1; cc--) stack.push({ sheet: p.sheet, r: rr, c: cc });
    }
  }
  return out;
}

/** Slider bounds for a value: 0 to twice the value (or the mirror for negatives), and a step of 1/100 of the span. */
export function sliderRange(v: number): { min: number; max: number; step: number } {
  if (!v) return { min: -10, max: 10, step: 0.1 };
  const a = Math.abs(v);
  const mag = 10 ** Math.floor(Math.log10(a * 2));
  const span = Math.ceil((a * 2) / mag) * mag;
  const min = v > 0 ? 0 : -span, max = v > 0 ? span : 0;
  const step = +(span / 100).toPrecision(2);
  return { min, max, step: Number.isInteger(v) && step >= 1 ? Math.round(step) : step };
}

// ---------------------------------------------------------------------------------------------
// Records

export type CondOp = 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'contains' | 'begins' | 'blank' | 'nonblank' | 'error';

export const COND_LABELS: Record<CondOp, string> = {
  eq: 'equals', ne: 'does not equal', gt: 'is greater than', ge: 'is at least', lt: 'is less than', le: 'is at most',
  contains: 'contains', begins: 'begins with', blank: 'is blank', nonblank: 'is not blank', error: 'is an error',
};

export function matchesCondition(v: Scalar, op: CondOp, target: string): boolean {
  const blank = v == null || v === '';
  if (op === 'blank') return blank;
  if (op === 'nonblank') return !blank;
  if (op === 'error') return v instanceof CellError;
  if (v instanceof CellError) return false;
  const tn = target.trim() !== '' && !isNaN(+target) ? +target : null;
  const vs = v == null ? '' : String(v).toLowerCase();
  const ts = target.toLowerCase();
  switch (op) {
    case 'contains': return vs.includes(ts);
    case 'begins': return vs.startsWith(ts);
    case 'eq': return tn != null && typeof v === 'number' ? v === tn : vs === ts;
    case 'ne': return !(tn != null && typeof v === 'number' ? v === tn : vs === ts);
  }
  if (tn != null && typeof v === 'number') return op === 'gt' ? v > tn : op === 'ge' ? v >= tn : op === 'lt' ? v < tn : v <= tn;
  if (tn != null) return false;
  const cmp = vs.localeCompare(ts);
  return op === 'gt' ? cmp > 0 : op === 'ge' ? cmp >= 0 : op === 'lt' ? cmp < 0 : cmp <= 0;
}

// ---------------------------------------------------------------------------------------------
// Typed columns

export type ColumnType = 'text' | 'number' | 'integer' | 'currency' | 'percent' | 'date' | 'checkbox' | 'choice' | 'email' | 'url' | 'rating';

export const COLUMN_TYPES: { id: ColumnType; label: string; icon: string }[] = [
  { id: 'text', label: 'Text', icon: 'notes' },
  { id: 'number', label: 'Number', icon: 'pin' },
  { id: 'integer', label: 'Whole Number', icon: 'tag' },
  { id: 'currency', label: 'Currency', icon: 'payments' },
  { id: 'percent', label: 'Percent', icon: 'percent' },
  { id: 'date', label: 'Date', icon: 'calendar_today' },
  { id: 'checkbox', label: 'Checkbox', icon: 'check_box' },
  { id: 'choice', label: 'Choice (drop-down)', icon: 'arrow_drop_down_circle' },
  { id: 'email', label: 'Email', icon: 'alternate_email' },
  { id: 'url', label: 'Web Address', icon: 'link' },
  { id: 'rating', label: 'Rating (1-5)', icon: 'star' },
];

/** Distinct non-blank values in display order (for Choice columns). */
export function distinctChoices(values: Scalar[], max = 200): string[] {
  const seen = new Map<string, string>();
  for (const v of values) {
    if (v == null || v === '' || v instanceof CellError) continue;
    const s = String(v);
    if (!seen.has(s.toLowerCase())) seen.set(s.toLowerCase(), s);
    if (seen.size >= max) break;
  }
  return [...seen.values()];
}

/** Values that do not fit a column type (row offsets within the column). */
export function typeMismatches(values: Scalar[], type: ColumnType, choices: string[] = []): number[] {
  const bad: number[] = [];
  const set = new Set(choices.map((c) => c.toLowerCase()));
  values.forEach((v, i) => {
    if (v == null || v === '') return;
    let ok = true;
    switch (type) {
      case 'number': case 'currency': case 'percent': case 'date': ok = typeof v === 'number'; break;
      case 'integer': ok = typeof v === 'number' && Number.isInteger(v); break;
      case 'rating': ok = typeof v === 'number' && Number.isInteger(v) && v >= 1 && v <= 5; break;
      case 'checkbox': ok = typeof v === 'boolean'; break;
      case 'choice': ok = set.has(String(v).toLowerCase()); break;
      case 'email': ok = typeof v === 'string' && /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v.trim()); break;
      case 'url': ok = typeof v === 'string' && /^(https?:\/\/|www\.)\S+$/i.test(v.trim()); break;
      case 'text': ok = true; break;
    }
    if (!ok) bad.push(i);
  });
  return bad;
}

// ---------------------------------------------------------------------------------------------
// Slash menu

/**
 * Accelerator letters for a list of labels: 1-2-3 picked menu words with distinct first letters; when
 * two share one, the next unused letter of the label is taken, then a digit. `priority` (indices,
 * most important first) decides who gets a contested letter; by default it is the list order.
 */
export function assignKeys(labels: string[], priority: number[] = labels.map((_, i) => i)): string[] {
  const used = new Set<string>();
  const out: string[] = labels.map(() => '');
  const order = [...priority.filter((i) => i >= 0 && i < labels.length), ...labels.map((_, i) => i).filter((i) => !priority.includes(i))];
  // First pass: first letters, in priority order.
  for (const i of order) {
    const k = (/[A-Za-z0-9]/.exec(labels[i])?.[0] ?? '').toUpperCase();
    if (k && !used.has(k)) { used.add(k); out[i] = k; }
  }
  for (const i of order) {
    if (out[i]) continue;
    const l = labels[i];
    // Then the first letter of a later word, then any letter.
    const words = l.split(/[\s&/-]+/).slice(1).map((w) => (/[A-Za-z0-9]/.exec(w)?.[0] ?? '').toUpperCase());
    const letters = [...words, ...[...l.toUpperCase()].filter((ch) => /[A-Z0-9]/.test(ch))];
    const k = letters.find((ch) => ch && !used.has(ch)) ?? '123456789'.split('').find((d) => !used.has(d)) ?? '';
    if (k) used.add(k);
    out[i] = k;
  }
  return out;
}
