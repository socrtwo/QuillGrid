/**
 * Statistical functions: descriptive statistics, conditional aggregates, regression, probability
 * distributions and hypothesis tests, plus the legacy (Compatibility) names.
 *
 * The numerical core (gamma, incomplete gamma/beta, normal CDF/inverse, root finding) is exported so the
 * Math & Trig library (AGGREGATE, SUBTOTAL) can share it.
 */
import { defFn, type FnDef } from './registry';
import {
  bool, fail, isRef, isRefs, iterValues, toMatrix, num, criteria, numbers, ERR, CellError,
  type Ctx, type Value, type RefV,
} from '../values';
import type { Matrix, Scalar } from '../../model/types';

// ============================================================================================
// Numerical core

const LN_SQRT_2PI = 0.91893853320467274178;
const SQRT2 = Math.SQRT2;
const EPS = 1e-16;
const FPMIN = 1e-300;

const LANCZOS = [
  0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
  12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
];

const FACTS: number[] = [1];
for (let i = 1; i <= 170; i++) FACTS[i] = FACTS[i - 1] * i;

/** n! for integers 0..170 (Infinity beyond). */
export function factorial(n: number): number { return n <= 170 ? FACTS[n] : Infinity; }

/** ln|Γ(x)| via Lanczos (g = 7, n = 9) with reflection for x < 0.5. */
export function gammaln(x: number): number {
  if (Number.isInteger(x) && x > 0 && x <= 171) return Math.log(FACTS[x - 1]);
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - gammaln(1 - x);
  if (x > 1e7) return (x - 0.5) * Math.log(x) - x + LN_SQRT_2PI + 1 / (12 * x);
  const z = x - 1;
  let a = LANCZOS[0];
  const t = z + 7.5;
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (z + i);
  return LN_SQRT_2PI + (z + 0.5) * Math.log(t) - t + Math.log(a);
}

/** Γ(x); NaN at the poles (0, -1, -2...). */
export function gamma(x: number): number {
  if (Number.isInteger(x)) return x <= 0 ? NaN : x <= 171 ? FACTS[x - 1] : Infinity;
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
  if (x > 171.7) return Infinity;
  const z = x - 1;
  let a = LANCZOS[0];
  const t = z + 7.5;
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (z + i);
  const p = Math.pow(t, (z + 0.5) / 2);
  return Math.sqrt(2 * Math.PI) * p * (p * Math.exp(-t)) * a;
}

/** Regularized lower incomplete gamma P(a, x). */
export function gammaP(a: number, x: number): number {
  if (x <= 0) return 0;
  if (!Number.isFinite(x)) return 1;
  return x < a + 1 ? gser(a, x) : 1 - gcf(a, x);
}

/** Regularized upper incomplete gamma Q(a, x) = 1 - P(a, x), computed without cancellation. */
export function gammaQ(a: number, x: number): number {
  if (x <= 0) return 1;
  if (!Number.isFinite(x)) return 0;
  return x < a + 1 ? 1 - gser(a, x) : gcf(a, x);
}

function gser(a: number, x: number): number {
  let ap = a, del = 1 / a, sum = del;
  for (let n = 0; n < 100000; n++) {
    ap += 1;
    del *= x / ap;
    sum += del;
    if (Math.abs(del) < Math.abs(sum) * EPS) break;
  }
  return sum * Math.exp(-x + a * Math.log(x) - gammaln(a));
}

function gcf(a: number, x: number): number {
  let b = x + 1 - a, c = 1 / FPMIN, d = 1 / b, h = d;
  for (let i = 1; i < 100000; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = b + an / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return Math.exp(-x + a * Math.log(x) - gammaln(a)) * h;
}

function betacf(x: number, a: number, b: number): number {
  const qab = a + b, qap = a + 1, qam = a - 1;
  let c = 1, d = 1 - (qab * x) / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= 100000; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return h;
}

/**
 * Regularized incomplete beta I_x(a, b) (or its complement when `upper`). `y` = 1 - x may be supplied
 * when it is known more accurately than 1 - x.
 */
export function betaInc(x: number, a: number, b: number, upper = false, y = 1 - x): number {
  if (x <= 0) return upper ? 1 : 0;
  if (y <= 0) return upper ? 0 : 1;
  const lbt = gammaln(a + b) - gammaln(a) - gammaln(b) + a * Math.log(x) + b * Math.log(y);
  const bt = Math.exp(lbt);
  if (x < (a + 1) / (a + b + 2)) {
    const v = (bt * betacf(x, a, b)) / a;
    return upper ? 1 - v : v;
  }
  const v = (bt * betacf(y, b, a)) / b;
  return upper ? v : 1 - v;
}

/** Complementary error function via the incomplete gamma function. */
export function erfc(x: number): number {
  if (x >= 0) return gammaQ(0.5, x * x);
  return 2 - gammaQ(0.5, x * x);
}
export function erf(x: number): number {
  return x >= 0 ? gammaP(0.5, x * x) : -gammaP(0.5, x * x);
}

/** Standard normal CDF. */
export function normCdf(z: number): number { return 0.5 * erfc(-z / SQRT2); }
/** Standard normal upper tail 1 - Φ(z). */
export function normSf(z: number): number { return 0.5 * erfc(z / SQRT2); }
export function normPdf(z: number): number { return Math.exp(-0.5 * z * z - LN_SQRT_2PI); }

/** Standard normal quantile: Acklam's rational approximation refined with Halley steps. */
export function normInv(p: number): number {
  if (p <= 0) return -Infinity;
  if (p >= 1) return Infinity;
  if (p === 0.5) return 0;
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const plow = 0.02425;
  let x: number;
  if (p < plow) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= 1 - plow) {
    const q = p - 0.5, r = q * q;
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  const q = 1 - p;
  for (let i = 0; i < 3; i++) {
    const e = p < 0.5 ? normCdf(x) - p : q - normSf(x);
    const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
    const nx = x - u / (1 + (x * u) / 2);
    if (!Number.isFinite(nx)) break;
    x = nx;
  }
  return x;
}

/** Brent's root finder for f on [a, b] where f(a), f(b) bracket a root. */
export function brent(f: (x: number) => number, a: number, b: number, tol = 1e-15, maxIter = 500): number {
  let fa = f(a), fb = f(b);
  if (fa === 0) return a;
  if (fb === 0) return b;
  if (fa * fb > 0) return NaN;
  let c = a, fc = fa, d = b - a, e = d;
  for (let i = 0; i < maxIter; i++) {
    if (fb * fc > 0) { c = a; fc = fa; d = b - a; e = d; }
    if (Math.abs(fc) < Math.abs(fb)) { a = b; b = c; c = a; fa = fb; fb = fc; fc = fa; }
    const tol1 = 2 * Number.EPSILON * Math.abs(b) + 0.5 * tol * Math.max(Math.abs(b), 1e-300);
    const xm = 0.5 * (c - b);
    if (Math.abs(xm) <= tol1 || fb === 0) return b;
    if (Math.abs(e) >= tol1 && Math.abs(fa) > Math.abs(fb)) {
      const s = fb / fa;
      let p: number, q: number;
      if (a === c) { p = 2 * xm * s; q = 1 - s; }
      else {
        const qq = fa / fc, r = fb / fc;
        p = s * (2 * xm * qq * (qq - r) - (b - a) * (r - 1));
        q = (qq - 1) * (r - 1) * (s - 1);
      }
      if (p > 0) q = -q; else p = -p;
      if (2 * p < Math.min(3 * xm * q - Math.abs(tol1 * q), Math.abs(e * q))) { e = d; d = p / q; }
      else { d = xm; e = d; }
    } else { d = xm; e = d; }
    a = b; fa = fb;
    b += Math.abs(d) > tol1 ? d : xm > 0 ? tol1 : -tol1;
    fb = f(b);
  }
  return b;
}

/** Inverts a monotone function g (increasing or decreasing) on [lo, ∞) or [lo, hi] for target p. */
export function invert(g: (x: number) => number, p: number, lo: number, hi: number, maxHi = 1e300): number {
  const h = (x: number) => g(x) - p;
  let fhi = h(hi);
  const flo = h(lo);
  let guard = 0;
  while (flo * fhi > 0 && hi < maxHi && guard++ < 2000) { lo = hi; hi *= 2; fhi = h(hi); }
  return brent(h, lo, hi);
}

// --- Distributions -------------------------------------------------------------------------

export function tCdf(t: number, df: number): number {
  const x = df / (df + t * t), y = (t * t) / (df + t * t);
  const tail = 0.5 * betaInc(x, df / 2, 0.5, false, y);
  return t > 0 ? 1 - tail : tail;
}
/** Upper tail P(T > t). */
export function tSf(t: number, df: number): number { return tCdf(-t, df); }
export function tPdf(t: number, df: number): number {
  return Math.exp(gammaln((df + 1) / 2) - gammaln(df / 2) - 0.5 * Math.log(df * Math.PI) - ((df + 1) / 2) * Math.log1p((t * t) / df));
}
export function tInv(p: number, df: number): number {
  if (p === 0.5) return 0;
  if (p < 0.5) return -tInv(1 - p, df);
  // p > 0.5: solve upper tail = 1 - p for precision.
  const q = 1 - p;
  if (df === 1) return Math.tan(Math.PI * (p - 0.5));
  if (df === 2) return (2 * p - 1) / Math.sqrt(2 * p * q);
  return invert((x) => -tSf(x, df), -q, 0, 4);
}

export function chiCdf(x: number, df: number): number { return gammaP(df / 2, x / 2); }
export function chiSf(x: number, df: number): number { return gammaQ(df / 2, x / 2); }
export function chiPdf(x: number, df: number): number {
  if (x < 0) return 0;
  if (x === 0) return df === 2 ? 0.5 : df < 2 ? Infinity : 0;
  return Math.exp((df / 2 - 1) * Math.log(x) - x / 2 - (df / 2) * Math.LN2 - gammaln(df / 2));
}

export function gammaCdf(x: number, a: number, b: number): number { return gammaP(a, x / b); }
export function gammaPdf(x: number, a: number, b: number): number {
  if (x < 0) return 0;
  if (x === 0) return a === 1 ? 1 / b : a < 1 ? Infinity : 0;
  return Math.exp((a - 1) * Math.log(x) - x / b - gammaln(a) - a * Math.log(b));
}
export function gammaInv(p: number, a: number, b: number): number {
  if (p === 0) return 0;
  const guess = Math.max(a * b, 1e-3);
  return invert((x) => gammaP(a, x / b), p, 0, guess);
}

export function fCdf(x: number, d1: number, d2: number): number {
  if (x <= 0) return 0;
  const den = d1 * x + d2;
  return betaInc((d1 * x) / den, d1 / 2, d2 / 2, false, d2 / den);
}
export function fSf(x: number, d1: number, d2: number): number {
  if (x <= 0) return 1;
  const den = d1 * x + d2;
  return betaInc((d1 * x) / den, d1 / 2, d2 / 2, true, d2 / den);
}
export function fPdf(x: number, d1: number, d2: number): number {
  if (x < 0) return 0;
  if (x === 0) return d1 === 2 ? 1 : d1 < 2 ? Infinity : 0;
  const lb = gammaln(d1 / 2) + gammaln(d2 / 2) - gammaln((d1 + d2) / 2);
  return Math.exp((d1 / 2) * Math.log(d1 / d2) + (d1 / 2 - 1) * Math.log(x) - ((d1 + d2) / 2) * Math.log1p((d1 * x) / d2) - lb);
}

export function betaPdf(x: number, a: number, b: number): number {
  if (x < 0 || x > 1) return 0;
  if (x === 0) return a === 1 ? b : a < 1 ? Infinity : 0;
  if (x === 1) return b === 1 ? a : b < 1 ? Infinity : 0;
  return Math.exp((a - 1) * Math.log(x) + (b - 1) * Math.log1p(-x) + gammaln(a + b) - gammaln(a) - gammaln(b));
}

export function combin(n: number, k: number): number {
  if (k < 0 || k > n) return 0;
  k = Math.min(k, n - k);
  if (n <= 170) return Math.round(FACTS[n] / FACTS[k] / FACTS[n - k]);
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return Number.isFinite(r) ? Math.round(r) : Math.exp(gammaln(n + 1) - gammaln(k + 1) - gammaln(n - k + 1));
}

/** ln C(n, k) */
function lnCombin(n: number, k: number): number { return gammaln(n + 1) - gammaln(k + 1) - gammaln(n - k + 1); }

export function binomPmf(k: number, n: number, p: number): number {
  if (k < 0 || k > n) return 0;
  if (p === 0) return k === 0 ? 1 : 0;
  if (p === 1) return k === n ? 1 : 0;
  const c = combin(n, k);
  if (Number.isFinite(c) && c < 1e300) {
    const v = c * Math.pow(p, k) * Math.pow(1 - p, n - k);
    if (v > 1e-280) return v;
  }
  return Math.exp(lnCombin(n, k) + k * Math.log(p) + (n - k) * Math.log1p(-p));
}

export function binomCdf(k: number, n: number, p: number): number {
  if (k < 0) return 0;
  if (k >= n) return 1;
  if (n > 5000) return betaInc(p, k + 1, n - k, true);
  let s = 0;
  for (let i = 0; i <= k; i++) s += binomPmf(i, n, p);
  return Math.min(1, s);
}

function poissonPmf(k: number, m: number): number {
  if (m === 0) return k === 0 ? 1 : 0;
  if (k <= 170 && m < 700) return (Math.exp(-m) * Math.pow(m, k)) / FACTS[k];
  return Math.exp(-m + k * Math.log(m) - gammaln(k + 1));
}

function hypgeomPmf(s: number, n: number, S: number, N: number): number {
  return Math.exp(lnCombin(S, s) + lnCombin(N - S, n - s) - lnCombin(N, n));
}

// ============================================================================================
// Argument helpers

type Args = Value[];
const sc = (a: Args, i: number): Scalar => a[i] as Scalar;
/** Optional numeric argument: omitted or empty -> default. */
const optN = (a: Args, i: number, def: number): number => (i >= a.length || a[i] == null ? def : num(sc(a, i)));
const optB = (a: Args, i: number, def: boolean): boolean => (i >= a.length || a[i] == null ? def : bool(sc(a, i)));
const tr = Math.trunc;
const chkNum = (x: number): number => (Number.isFinite(x) ? x : fail(ERR.num));

/** Numbers from aggregate arguments (text/booleans in references skipped, direct args coerced). */
export const nums = (args: Value[], ctx: Ctx): number[] => numbers(args, ctx);
/** Numbers for the *A variants: text in references counts as 0, booleans as 1/0. */
export const numsA = (args: Value[], ctx: Ctx): number[] => numbers(args, ctx, { countText0: true });

/** Numbers from a single array/range argument, ignoring text and booleans, propagating errors. */
export function arrNums(v: Value, ctx: Ctx): number[] {
  const out: number[] = [];
  for (const { v: x, fromRef } of iterValues(v, ctx)) {
    if (typeof x === 'number') out.push(x);
    else if (x instanceof CellError) fail(x);
    else if (!fromRef && x != null) out.push(num(x));
  }
  return out;
}

/** All cell values of an argument as a flat list (row-major). */
function flatVals(v: Value, ctx: Ctx): Scalar[] {
  const out: Scalar[] = [];
  for (const { v: x } of iterValues(v, ctx)) out.push(x);
  return out;
}

/** Paired numeric data from two arrays of equal size (pairs where either is non-numeric are skipped). */
export function pairs(a: Value, b: Value, ctx: Ctx, naOnSize = true): [number[], number[]] {
  const xa = flatVals(a, ctx), xb = flatVals(b, ctx);
  if (xa.length !== xb.length) { if (naOnSize) fail(ERR.na); }
  const x: number[] = [], y: number[] = [];
  const n = Math.min(xa.length, xb.length);
  for (let i = 0; i < n; i++) {
    const p = xa[i], q = xb[i];
    if (p instanceof CellError) fail(p);
    if (q instanceof CellError) fail(q);
    if (typeof p === 'number' && typeof q === 'number') { x.push(p); y.push(q); }
  }
  return [x, y];
}

// ============================================================================================
// Descriptive statistics helpers (exported for AGGREGATE/SUBTOTAL)

export const sum = (xs: number[]): number => { let s = 0; for (const x of xs) s += x; return s; };
export const mean = (xs: number[]): number => (xs.length ? sum(xs) / xs.length : fail(ERR.div0));
export function devsq(xs: number[]): number {
  const m = sum(xs) / xs.length;
  let s = 0;
  for (const x of xs) s += (x - m) * (x - m);
  return s;
}
export function variance(xs: number[], sample: boolean): number {
  const n = xs.length;
  if (n === 0 || (sample && n < 2)) return fail(ERR.div0);
  return devsq(xs) / (sample ? n - 1 : n);
}
export const stdev = (xs: number[], sample: boolean): number => Math.sqrt(variance(xs, sample));
export function median(xs: number[]): number {
  if (!xs.length) return fail(ERR.num);
  const s = [...xs].sort((a, b) => a - b), n = s.length;
  return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
}
export function modeSngl(xs: number[]): number {
  const m = modes(xs);
  return m.length ? m[0] : fail(ERR.na);
}
/** Values occurring most often (count > 1), in order of first appearance. */
export function modes(xs: number[]): number[] {
  const counts = new Map<number, number>();
  for (const x of xs) counts.set(x, (counts.get(x) || 0) + 1);
  let best = 1;
  for (const c of counts.values()) if (c > best) best = c;
  if (best < 2) return [];
  const out: number[] = [];
  const seen = new Set<number>();
  for (const x of xs) if (!seen.has(x) && counts.get(x) === best) { seen.add(x); out.push(x); }
  return out;
}
export function kth(xs: number[], k: number, largest: boolean): number {
  const n = xs.length;
  k = Math.ceil(k - 1e-12);
  if (!n || k < 1 || k > n) return fail(ERR.num);
  const s = [...xs].sort((a, b) => (largest ? b - a : a - b));
  return s[k - 1];
}
export function percentileInc(xs: number[], k: number): number {
  const n = xs.length;
  if (!n || k < 0 || k > 1) return fail(ERR.num);
  const s = [...xs].sort((a, b) => a - b);
  const h = (n - 1) * k, lo = Math.floor(h);
  return lo + 1 < n ? s[lo] + (h - lo) * (s[lo + 1] - s[lo]) : s[lo];
}
export function percentileExc(xs: number[], k: number): number {
  const n = xs.length;
  if (!n || k <= 0 || k >= 1) return fail(ERR.num);
  const h = (n + 1) * k;
  if (h < 1 - 1e-12 || h > n + 1e-12) return fail(ERR.num);
  const s = [...xs].sort((a, b) => a - b);
  const lo = Math.max(1, Math.min(n, Math.floor(h)));
  return lo < n ? s[lo - 1] + (h - lo) * (s[lo] - s[lo - 1]) : s[n - 1];
}
export function quartileInc(xs: number[], q: number): number {
  q = tr(q);
  if (q < 0 || q > 4) return fail(ERR.num);
  return percentileInc(xs, q / 4);
}
export function quartileExc(xs: number[], q: number): number {
  q = tr(q);
  if (q <= 0 || q >= 4) return fail(ERR.num);
  return percentileExc(xs, q / 4);
}

function percentRank(xs: number[], x: number, sig: number, exc: boolean): number {
  sig = tr(sig);
  if (sig < 1) return fail(ERR.num);
  const n = xs.length;
  if (!n) return fail(ERR.num);
  const s = [...xs].sort((a, b) => a - b);
  if (x < s[0] || x > s[n - 1]) return fail(ERR.na);
  let r: number;
  if (n === 1 && !exc) r = 1;
  else {
    let less = 0;
    while (less < n && s[less] < x) less++;
    const div = exc ? n + 1 : n - 1;
    const base = exc ? 1 : 0;
    if (less < n && s[less] === x) r = (less + base) / div;
    else {
      const i = less - 1; // s[i] < x < s[i + 1]
      r = (i + base + (x - s[i]) / (s[i + 1] - s[i])) / div;
    }
  }
  const f = Math.pow(10, sig);
  return Math.floor(+(r * f).toPrecision(15)) / f;
}

function rank(x: number, xs: number[], asc: boolean, avg: boolean): number {
  let better = 0, eq = 0;
  for (const v of xs) {
    if (v === x) eq++;
    else if (asc ? v < x : v > x) better++;
  }
  if (!eq) return fail(ERR.na);
  return avg ? better + (eq + 1) / 2 : better + 1;
}

function kurt(xs: number[]): number {
  const n = xs.length;
  if (n < 4) return fail(ERR.div0);
  const m = mean(xs), s = stdev(xs, true);
  if (s === 0) return fail(ERR.div0);
  let q = 0;
  for (const x of xs) q += ((x - m) / s) ** 4;
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * q - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
}

function skew(xs: number[], pop: boolean): number {
  const n = xs.length;
  if (pop ? n < 1 : n < 3) return fail(ERR.div0);
  const m = mean(xs), s = pop ? stdev(xs, false) : stdev(xs, true);
  if (s === 0) return fail(ERR.div0);
  let c = 0;
  for (const x of xs) c += ((x - m) / s) ** 3;
  return pop ? c / n : (n / ((n - 1) * (n - 2))) * c;
}

// ============================================================================================
// Conditional aggregation (…IF / …IFS)

interface Grid { rows: number; cols: number; ur: number; uc: number; get(i: number, j: number): Scalar }

export function grid(v: Value, ctx: Ctx): Grid {
  if (isRef(v)) {
    const rows = v.r2 - v.r1 + 1, cols = v.c2 - v.c1 + 1;
    const b = ctx.usedBounds(v.sheet);
    const ur = Math.max(0, Math.min(rows, b.r2 - v.r1 + 1)), uc = Math.max(0, Math.min(cols, b.c2 - v.c1 + 1));
    return { rows, cols, ur, uc, get: (i, j) => ctx.cell(v.sheet, v.r1 + i, v.c1 + j) };
  }
  if (isRefs(v)) return fail(ERR.value);
  const m = toMatrix(v, ctx);
  const rows = m.length, cols = m[0]?.length ?? 0;
  return { rows, cols, ur: rows, uc: cols, get: (i, j) => m[i]?.[j] ?? null };
}

/** A sum/average range resized to the shape of the criteria range (as SUMIF does). */
export function resized(v: Value, rows: number, cols: number, ctx: Ctx): Grid {
  if (isRef(v)) return grid({ ...v, r2: v.r1 + rows - 1, c2: v.c1 + cols - 1 } as RefV, ctx);
  return grid(v, ctx);
}

/**
 * Walks equally-sized criteria ranges and calls `hit(i, j)` for each position where every criterion
 * matches. Returns the number of matching positions outside the used range (all blank).
 */
function walkIfs(ranges: Grid[], tests: ((v: Scalar) => boolean)[], hit: (i: number, j: number) => void): number {
  const { rows, cols } = ranges[0];
  for (const g of ranges) if (g.rows !== rows || g.cols !== cols) fail(ERR.value);
  let ur = 0, uc = 0;
  for (const g of ranges) { ur = Math.max(ur, g.ur); uc = Math.max(uc, g.uc); }
  for (let i = 0; i < ur; i++) for (let j = 0; j < uc; j++) {
    let ok = true;
    for (let k = 0; k < ranges.length && ok; k++) ok = tests[k](ranges[k].get(i, j));
    if (ok) hit(i, j);
  }
  const extra = rows * cols - ur * uc;
  return extra > 0 && tests.every((t) => t(null)) ? extra : 0;
}

function critOf(v: Value): (x: Scalar) => boolean {
  return criteria(v as Scalar);
}

/** Collects the numeric values of `target` at positions where all criteria match. */
export function ifsValues(target: Grid | null, pairsArgs: Value[], ctx: Ctx): { vals: number[]; count: number } {
  const ranges: Grid[] = [], tests: ((v: Scalar) => boolean)[] = [];
  for (let i = 0; i < pairsArgs.length; i += 2) { ranges.push(grid(pairsArgs[i], ctx)); tests.push(critOf(pairsArgs[i + 1])); }
  if (target && (target.rows !== ranges[0].rows || target.cols !== ranges[0].cols)) fail(ERR.value);
  const vals: number[] = [];
  let count = 0;
  const extra = walkIfs(ranges, tests, (i, j) => {
    count++;
    if (target) {
      const v = target.get(i, j);
      if (v instanceof CellError) fail(v);
      if (typeof v === 'number') vals.push(v);
    }
  });
  return { vals, count: count + extra };
}

/** arrays indices for IFS functions: range positions every other argument starting at `start`. */
export function ifsArrays(first: number[], start: number): number[] {
  const out = [...first];
  for (let i = start; i < 255; i += 2) out.push(i);
  return out;
}

// ============================================================================================
// Regression

interface Fit {
  coef: number[]; // per x column (0 for dropped)
  b: number;
  kept: boolean[];
  n: number;
  k: number;
  df: number;
  ssreg: number;
  ssresid: number;
  r2: number;
  sey: number;
  se: number[];
  seb: number;
}

function matNums(m: Matrix): number[][] {
  return m.map((row) => row.map((v) => {
    if (v instanceof CellError) return fail(v);
    if (typeof v !== 'number') return fail(ERR.value);
    return v;
  }));
}

/** Arranges known_y / known_x into a y vector and an n x k design matrix (without the constant). */
function regData(yv: Value, xv: Value | undefined, ctx: Ctx): { y: number[]; X: number[][]; byRow: boolean; ymat: number[][] } {
  const ym = matNums(toMatrix(yv, ctx));
  const yr = ym.length, yc = ym[0]?.length ?? 0;
  const y = ym.flat();
  const n = y.length;
  const byRow = yr === 1 && yc > 1;
  let X: number[][];
  if (xv == null) X = y.map((_, i) => [i + 1]);
  else {
    const xm = matNums(toMatrix(xv, ctx));
    const xr = xm.length, xc = xm[0]?.length ?? 0;
    if (yc === 1 && yr > 1 || (yr === 1 && yc === 1)) {
      if (xr === n) X = xm.map((r) => [...r]);
      else if (xr === 1 && xc === n && n > 1) X = xm[0].map((v) => [v]);
      else return fail(ERR.ref);
    } else if (yr === 1) {
      if (xc === n) { X = []; for (let i = 0; i < n; i++) X.push(xm.map((r) => r[i])); }
      else return fail(ERR.ref);
    } else {
      if (xr !== yr || xc !== yc) return fail(ERR.ref);
      X = xm.flat().map((v) => [v]);
    }
  }
  return { y, X, byRow, ymat: ym };
}

function fitLinear(y: number[], X: number[][], useConst: boolean): Fit {
  const n = y.length, k = X[0]?.length ?? 0;
  const xbar = new Array(k).fill(0);
  let ybar = 0;
  if (useConst) {
    for (let i = 0; i < n; i++) { ybar += y[i]; for (let j = 0; j < k; j++) xbar[j] += X[i][j]; }
    ybar /= n;
    for (let j = 0; j < k; j++) xbar[j] /= n;
  }
  // Modified Gram-Schmidt on centered columns; drop columns that are (nearly) collinear.
  const cols: number[][] = [];
  for (let j = 0; j < k; j++) cols.push(X.map((r) => r[j] - xbar[j]));
  const Q: number[][] = [];
  const R: number[][] = []; // R[q][j] for kept q index
  const keptIdx: number[] = [];
  const kept = new Array(k).fill(false);
  for (let j = 0; j < k; j++) {
    const v = cols[j].slice();
    const norm0 = Math.sqrt(v.reduce((s, x) => s + x * x, 0));
    const rcol = new Array(Q.length).fill(0);
    for (let pass = 0; pass < 2; pass++) {
      for (let q = 0; q < Q.length; q++) {
        let d = 0;
        for (let i = 0; i < n; i++) d += Q[q][i] * v[i];
        rcol[q] += d;
        for (let i = 0; i < n; i++) v[i] -= d * Q[q][i];
      }
    }
    const nv = Math.sqrt(v.reduce((s, x) => s + x * x, 0));
    if (norm0 === 0 || nv <= 1e-12 * norm0 || keptIdx.length >= n - (useConst ? 1 : 0)) continue;
    kept[j] = true;
    keptIdx.push(j);
    for (let i = 0; i < n; i++) v[i] /= nv;
    Q.push(v);
    R.push(rcol.concat([nv]));
  }
  const p = keptIdx.length;
  // R as p x p upper triangular: Rm[a][b], a <= b
  const Rm: number[][] = Array.from({ length: p }, () => new Array(p).fill(0));
  for (let b = 0; b < p; b++) for (let a = 0; a <= b; a++) Rm[a][b] = R[b][a];
  const yc = y.map((v) => v - ybar);
  const qty = Q.map((q) => q.reduce((s, x, i) => s + x * yc[i], 0));
  const beta = new Array(p).fill(0);
  for (let a = p - 1; a >= 0; a--) {
    let s = qty[a];
    for (let b = a + 1; b < p; b++) s -= Rm[a][b] * beta[b];
    beta[a] = s / Rm[a][a];
  }
  const coef = new Array(k).fill(0);
  keptIdx.forEach((j, a) => { coef[j] = beta[a]; });
  let b0 = 0;
  if (useConst) { b0 = ybar; for (let j = 0; j < k; j++) b0 -= coef[j] * xbar[j]; }
  let ssresid = 0, sstot = 0;
  for (let i = 0; i < n; i++) {
    let yh = b0;
    for (let j = 0; j < k; j++) yh += coef[j] * X[i][j];
    ssresid += (y[i] - yh) ** 2;
    sstot += useConst ? (y[i] - ybar) ** 2 : y[i] ** 2;
  }
  const ssreg = sstot - ssresid;
  const df = n - p - (useConst ? 1 : 0);
  const r2 = sstot === 0 ? 1 : ssreg / sstot;
  const s2 = df > 0 ? ssresid / df : NaN;
  const sey = Math.sqrt(s2);
  // (R'R)^-1 = Rinv Rinv'
  const Rinv: number[][] = Array.from({ length: p }, () => new Array(p).fill(0));
  for (let c = 0; c < p; c++) {
    for (let a = p - 1; a >= 0; a--) {
      let s = a === c ? 1 : 0;
      for (let b = a + 1; b < p; b++) s -= Rm[a][b] * Rinv[b][c];
      Rinv[a][c] = s / Rm[a][a];
    }
  }
  const cov = (a: number, b: number) => { let s = 0; for (let l = 0; l < p; l++) s += Rinv[a][l] * Rinv[b][l]; return s; };
  const se = new Array(k).fill(0);
  keptIdx.forEach((j, a) => { se[j] = Math.sqrt(s2 * cov(a, a)); });
  let seb = NaN;
  if (useConst) {
    let quad = 0;
    for (let a = 0; a < p; a++) for (let b = 0; b < p; b++) quad += xbar[keptIdx[a]] * cov(a, b) * xbar[keptIdx[b]];
    seb = Math.sqrt(s2 * (1 / n + quad));
  }
  return { coef, b: b0, kept, n, k, df, ssreg, ssresid, r2, sey, se, seb };
}

function linestOut(fit: Fit, useConst: boolean, stats: boolean, tf: (x: number) => number): Matrix {
  const k = fit.k;
  const row1: Scalar[] = [];
  for (let j = k - 1; j >= 0; j--) row1.push(tf(fit.coef[j]));
  row1.push(useConst ? tf(fit.b) : tf(0));
  if (!stats) return [row1];
  const nn = (x: number): Scalar => (Number.isFinite(x) ? x : ERR.num);
  const row2: Scalar[] = [];
  for (let j = k - 1; j >= 0; j--) row2.push(nn(fit.se[j]));
  row2.push(useConst ? nn(fit.seb) : ERR.na);
  const pad = (r: Scalar[]) => { while (r.length < k + 1) r.push(ERR.na); return r; };
  const p = fit.kept.filter(Boolean).length;
  const F = fit.df > 0 && fit.ssresid > 0 && p > 0 ? (fit.ssreg / p) / (fit.ssresid / fit.df) : NaN;
  return [row1, row2, pad([nn(fit.r2), nn(fit.sey)]), pad([nn(F), fit.df]), pad([fit.ssreg, fit.ssresid])];
}

function predict(fit: Fit, xs: number[][], tf: (x: number) => number): number[] {
  return xs.map((r) => { let v = fit.b; for (let j = 0; j < fit.k; j++) v += fit.coef[j] * r[j]; return tf(v); });
}

/** TREND / GROWTH */
function trendLike(args: Args, ctx: Ctx, log: boolean): Value {
  const yv = args[0];
  const xv = args.length > 1 && args[1] != null ? args[1] : undefined;
  const useConst = optB(args, 3, true);
  const d = regData(yv, xv, ctx);
  let y = d.y;
  if (log) { if (y.some((v) => v <= 0)) return ERR.num; y = y.map(Math.log); }
  const fit = fitLinear(y, d.X, useConst);
  const k = d.X[0].length;
  let newX: number[][], shape: [number, number];
  if (args.length > 2 && args[2] != null) {
    const nm = matNums(toMatrix(args[2], ctx));
    const r = nm.length, c = nm[0]?.length ?? 0;
    if (k === 1) { newX = nm.flat().map((v) => [v]); shape = [r, c]; }
    else if (!d.byRow) { if (c !== k) return ERR.ref; newX = nm; shape = [r, 1]; }
    else { if (r !== k) return ERR.ref; newX = []; for (let i = 0; i < c; i++) newX.push(nm.map((row) => row[i])); shape = [1, c]; }
  } else {
    newX = d.X;
    shape = k === 1 ? [d.ymat.length, d.ymat[0].length] : d.byRow ? [1, d.y.length] : [d.y.length, 1];
  }
  const out = predict(fit, newX, log ? Math.exp : (x) => x);
  const m: Matrix = [];
  let idx = 0;
  for (let r = 0; r < shape[0]; r++) { const row: Scalar[] = []; for (let c = 0; c < shape[1]; c++) row.push(out[idx++]); m.push(row); }
  return m;
}

// ============================================================================================
// Exponential smoothing (FORECAST.ETS family) — additive error/trend/seasonality, best effort.

interface EtsModel { alpha: number; beta: number; gamma: number; period: number; level: number; trend: number; season: number[]; sse: number; fitted: number[]; y: number[]; step: number; t0: number; tLast: number; mae: number; rmse: number; smape: number; mase: number }

function etsPrepare(values: Value, timeline: Value, ctx: Ctx, completion: number, aggregation: number): { y: number[]; step: number; t0: number } {
  const vs = flatVals(values, ctx), ts = flatVals(timeline, ctx);
  if (vs.length !== ts.length) return fail(ERR.na);
  const pts: [number, number | null][] = [];
  for (let i = 0; i < ts.length; i++) {
    const t = ts[i], v = vs[i];
    if (t instanceof CellError) fail(t);
    if (v instanceof CellError) fail(v);
    if (typeof t !== 'number') return fail(ERR.value);
    pts.push([t, typeof v === 'number' ? v : null]);
  }
  if (pts.length < 3) return fail(ERR.num);
  pts.sort((a, b) => a[0] - b[0]);
  // aggregate duplicates
  const groups: [number, number[]][] = [];
  for (const [t, v] of pts) {
    const last = groups[groups.length - 1];
    if (last && last[0] === t) { if (v != null) last[1].push(v); }
    else groups.push([t, v != null ? [v] : []]);
  }
  const agg = (xs: number[]): number | null => {
    if (!xs.length) return null;
    switch (aggregation) {
      case 2: return xs.length;
      case 3: return xs.length;
      case 4: return Math.max(...xs);
      case 5: return median(xs);
      case 6: return Math.min(...xs);
      case 7: return sum(xs);
      default: return mean(xs);
    }
  };
  const times = groups.map((g) => g[0]);
  let step = Infinity;
  for (let i = 1; i < times.length; i++) step = Math.min(step, times[i] - times[i - 1]);
  if (!(step > 0)) return fail(ERR.num);
  for (let i = 1; i < times.length; i++) {
    const q = (times[i] - times[i - 1]) / step;
    if (Math.abs(q - Math.round(q)) > 1e-9) return fail(ERR.num);
  }
  const n = Math.round((times[times.length - 1] - times[0]) / step) + 1;
  const y: (number | null)[] = new Array(n).fill(null);
  for (const [t, xs] of groups) y[Math.round((t - times[0]) / step)] = agg(xs);
  // completion: 1 (default) = interpolate, 0 = zeros
  const known = y.map((v, i) => (v == null ? -1 : i)).filter((i) => i >= 0);
  if (known.length < 2) return fail(ERR.num);
  for (let i = 0; i < n; i++) {
    if (y[i] != null) continue;
    if (completion === 0) { y[i] = 0; continue; }
    let a = i - 1; while (a >= 0 && y[a] == null) a--;
    let b = i + 1; while (b < n && y[b] == null) b++;
    if (a < 0) y[i] = y[b]!;
    else if (b >= n) y[i] = y[a]!;
    else y[i] = y[a]! + ((y[b]! - y[a]!) * (i - a)) / (b - a);
  }
  return { y: y as number[], step, t0: times[0] };
}

function detectSeason(y: number[]): number {
  const n = y.length;
  if (n < 6) return 1;
  // detrend by first differences of a linear fit
  const xs = y.map((_, i) => i);
  const xm = mean(xs), ym = mean(y);
  let sxy = 0, sxx = 0;
  for (let i = 0; i < n; i++) { sxy += (xs[i] - xm) * (y[i] - ym); sxx += (xs[i] - xm) ** 2; }
  const sl = sxy / sxx;
  const r = y.map((v, i) => v - (ym + sl * (i - xm)));
  const r0 = r.reduce((s, v) => s + v * v, 0);
  if (r0 === 0) return 1;
  const acf = (lag: number) => { let s = 0; for (let i = lag; i < n; i++) s += r[i] * r[i - lag]; return s / r0; };
  const maxLag = Math.min(Math.floor(n / 2), 8760);
  let best = 1, bestV = 0;
  const vals: number[] = [];
  for (let lag = 1; lag <= maxLag; lag++) vals.push(acf(lag));
  for (let lag = 2; lag <= maxLag; lag++) {
    const v = vals[lag - 1];
    const prev = vals[lag - 2], next = lag < maxLag ? vals[lag] : -Infinity;
    if (v > prev && v >= next && v > bestV && v > 2 / Math.sqrt(n)) { best = lag; bestV = v; }
  }
  return best;
}

function etsRun(y: number[], alpha: number, beta: number, gam: number, m: number): { sse: number; level: number; trend: number; season: number[]; fitted: number[] } {
  const n = y.length;
  let level: number, trend: number;
  let season: number[] = [];
  if (m > 1 && n >= 2 * m) {
    const c1 = mean(y.slice(0, m)), c2 = mean(y.slice(m, 2 * m));
    level = c1; trend = (c2 - c1) / m;
    season = y.slice(0, m).map((v) => v - c1);
  } else {
    level = y[0]; trend = n > 1 ? y[1] - y[0] : 0;
    m = 1;
  }
  let sse = 0;
  const fitted: number[] = [];
  for (let t = 0; t < n; t++) {
    const s = m > 1 ? season[t % m] : 0;
    const f = level + trend + s;
    fitted.push(f);
    if (t > 0) sse += (y[t] - f) ** 2;
    const prevL = level;
    level = alpha * (y[t] - s) + (1 - alpha) * (level + trend);
    trend = beta * (level - prevL) + (1 - beta) * trend;
    if (m > 1) season[t % m] = gam * (y[t] - level) + (1 - gam) * s;
  }
  return { sse, level, trend, season, fitted };
}

function etsFit(values: Value, timeline: Value, ctx: Ctx, seasonality: number | null, completion: number, aggregation: number): EtsModel {
  const { y, step, t0 } = etsPrepare(values, timeline, ctx, completion, aggregation);
  let m: number;
  if (seasonality == null || seasonality === 1) m = detectSeason(y);
  else if (seasonality === 0) m = 1;
  else { m = tr(seasonality); if (m < 0 || m > 8760) fail(ERR.num); }
  if (m > 1 && y.length < 2 * m) m = 1;
  const grid1 = [0.001, 0.05, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 0.999];
  let best = { a: 0.5, b: 0.1, g: 0.1, sse: Infinity };
  for (const a of grid1) for (const b of [0.001, 0.05, 0.1, 0.2, 0.4]) for (const g of m > 1 ? [0.001, 0.05, 0.1, 0.25, 0.5] : [0]) {
    const r = etsRun(y, a, b, g, m);
    if (r.sse < best.sse) best = { a, b, g, sse: r.sse };
  }
  // coordinate refinement
  let stepSize = 0.05;
  for (let it = 0; it < 60 && stepSize > 1e-5; it++) {
    let improved = false;
    for (const key of ['a', 'b', 'g'] as const) {
      if (key === 'g' && m <= 1) continue;
      for (const dir of [-1, 1]) {
        const cand = { ...best };
        cand[key] = Math.min(0.999, Math.max(0.001, cand[key] + dir * stepSize));
        const r = etsRun(y, cand.a, cand.b, cand.g, m);
        if (r.sse < best.sse) { best = { ...cand, sse: r.sse }; improved = true; }
      }
    }
    if (!improved) stepSize /= 2;
  }
  const r = etsRun(y, best.a, best.b, best.g, m);
  const n = y.length;
  let mae = 0, smape = 0, naive = 0;
  for (let t = 1; t < n; t++) {
    const e = y[t] - r.fitted[t];
    mae += Math.abs(e);
    const den = (Math.abs(y[t]) + Math.abs(r.fitted[t])) / 2;
    smape += den ? Math.abs(e) / den : 0;
    naive += Math.abs(y[t] - y[t - 1]);
  }
  mae /= n - 1; smape /= n - 1; naive /= n - 1;
  return {
    alpha: best.a, beta: best.b, gamma: m > 1 ? best.g : 0, period: m, level: r.level, trend: r.trend, season: r.season,
    sse: r.sse, fitted: r.fitted, y, step, t0, tLast: t0 + (n - 1) * step,
    mae, rmse: Math.sqrt(r.sse / (n - 1)), smape, mase: naive ? mae / naive : 0,
  };
}

function etsForecast(md: EtsModel, target: number): number {
  const h = (target - md.tLast) / md.step;
  if (h <= 0) return fail(ERR.num);
  const n = md.y.length;
  const s = (i: number) => (md.period > 1 ? md.season[(((n + i - 1) % md.period) + md.period) % md.period] : 0);
  const hi = Math.ceil(h), lo = Math.floor(h);
  const at = (k: number) => md.level + k * md.trend + s(k);
  if (hi === lo) return at(h);
  const a = lo === 0 ? md.y[n - 1] : at(lo);
  return a + (h - lo) * (at(hi) - a);
}

function etsArgs(args: Args, ctx: Ctx, from: number): { seas: number | null; comp: number; agg: number } {
  const seas = from < args.length && args[from] != null ? num(sc(args, from)) : null;
  const comp = optN(args, from + 1, 1);
  const agg = optN(args, from + 2, 1);
  if (comp !== 0 && comp !== 1) fail(ERR.num);
  if (agg < 1 || agg > 7) fail(ERR.num);
  void ctx;
  return { seas, comp, agg };
}

// ============================================================================================
// Definitions

const S = 'Statistical' as const;

function dist(name: string, params: string[], desc: string, min: number, impl: (a: Args, ctx: Ctx) => Value): FnDef {
  return { name, min, max: params.length, category: S, desc, params, impl };
}

const defs: FnDef[] = [
  // --- Counting & averages -----------------------------------------------------------------
  {
    name: 'COUNT', min: 1, max: 255, category: S, arrays: 'all', desc: 'Counts how many arguments are numbers.', params: ['value1', '[value2]', '...'],
    impl: (args, ctx) => {
      let n = 0;
      for (const a of args) for (const { v, fromRef } of iterValues(a, ctx)) {
        if (typeof v === 'number') n++;
        else if (!fromRef && (typeof v === 'boolean' || (typeof v === 'string' && !isNaN(tryNum(v))))) n++;
      }
      return n;
    },
  },
  {
    name: 'COUNTA', min: 1, max: 255, category: S, arrays: 'all', desc: 'Counts how many arguments are not empty.', params: ['value1', '[value2]', '...'],
    impl: (args, ctx) => {
      let n = 0;
      for (const a of args) for (const { v, fromRef } of iterValues(a, ctx)) if (v != null || !fromRef) n++;
      return n;
    },
  },
  {
    name: 'COUNTBLANK', min: 1, max: 1, category: S, arrays: [0], desc: 'Counts empty cells (and cells with empty text) in a range.', params: ['range'],
    impl: ([a], ctx) => {
      if (isRef(a)) {
        const g = grid(a, ctx);
        let n = g.rows * g.cols - g.ur * g.uc;
        for (let i = 0; i < g.ur; i++) for (let j = 0; j < g.uc; j++) { const v = g.get(i, j); if (v == null || v === '') n++; }
        return n;
      }
      let n = 0;
      for (const { v } of iterValues(a, ctx)) if (v == null || v === '') n++;
      return n;
    },
  },
  {
    name: 'COUNTIF', min: 2, max: 2, category: S, arrays: [0], desc: 'Counts cells in a range that meet a condition.', params: ['range', 'criteria'],
    impl: ([r, c], ctx) => ifsValues(null, [r, c], ctx).count,
  },
  {
    name: 'COUNTIFS', min: 2, max: 254, category: S, arrays: ifsArrays([], 0), desc: 'Counts cells that meet every one of several conditions.', params: ['criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]', '...'],
    impl: (args, ctx) => (args.length % 2 ? ERR.value : ifsValues(null, args, ctx).count),
  },
  {
    name: 'AVERAGE', min: 1, max: 255, category: S, arrays: 'all', desc: 'Arithmetic mean of the arguments.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => mean(nums(args, ctx)),
  },
  {
    name: 'AVERAGEA', min: 1, max: 255, category: S, arrays: 'all', desc: 'Mean of the arguments, counting text as 0 and TRUE as 1.', params: ['value1', '[value2]', '...'],
    impl: (args, ctx) => mean(numsA(args, ctx)),
  },
  {
    name: 'AVERAGEIF', min: 2, max: 3, category: S, arrays: [0, 2], desc: 'Mean of the cells that meet a condition.', params: ['range', 'criteria', '[average_range]'],
    impl: (args, ctx) => {
      const rg = grid(args[0], ctx);
      const target = args.length > 2 && args[2] != null ? resized(args[2], rg.rows, rg.cols, ctx) : rg;
      const { vals } = ifsValues(target, [args[0], args[1]], ctx);
      return vals.length ? sum(vals) / vals.length : ERR.div0;
    },
  },
  {
    name: 'AVERAGEIFS', min: 3, max: 255, category: S, arrays: ifsArrays([0], 1), desc: 'Mean of the cells that meet every one of several conditions.', params: ['average_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]', '...'],
    impl: (args, ctx) => {
      if (args.length % 2 === 0) return ERR.value;
      const { vals } = ifsValues(grid(args[0], ctx), args.slice(1), ctx);
      return vals.length ? sum(vals) / vals.length : ERR.div0;
    },
  },
  {
    name: 'MAXIFS', min: 3, max: 255, category: S, arrays: ifsArrays([0], 1), desc: 'Largest value among cells that meet every condition.', params: ['max_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]', '...'],
    impl: (args, ctx) => {
      if (args.length % 2 === 0) return ERR.value;
      const { vals } = ifsValues(grid(args[0], ctx), args.slice(1), ctx);
      return vals.length ? Math.max(...vals) : 0;
    },
  },
  {
    name: 'MINIFS', min: 3, max: 255, category: S, arrays: ifsArrays([0], 1), desc: 'Smallest value among cells that meet every condition.', params: ['min_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]', '...'],
    impl: (args, ctx) => {
      if (args.length % 2 === 0) return ERR.value;
      const { vals } = ifsValues(grid(args[0], ctx), args.slice(1), ctx);
      return vals.length ? Math.min(...vals) : 0;
    },
  },
  {
    name: 'AVEDEV', min: 1, max: 255, category: S, arrays: 'all', desc: 'Mean of the absolute deviations of values from their mean.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => {
      const xs = nums(args, ctx);
      if (!xs.length) return ERR.num;
      const m = mean(xs);
      return sum(xs.map((x) => Math.abs(x - m))) / xs.length;
    },
  },
  {
    name: 'DEVSQ', min: 1, max: 255, category: S, arrays: 'all', desc: 'Sum of squared deviations from the mean.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => { const xs = nums(args, ctx); return xs.length ? devsq(xs) : ERR.num; },
  },
  {
    name: 'GEOMEAN', min: 1, max: 255, category: S, arrays: 'all', desc: 'Geometric mean of positive numbers.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => {
      const xs = nums(args, ctx);
      if (!xs.length || xs.some((x) => x <= 0)) return ERR.num;
      return Math.exp(sum(xs.map(Math.log)) / xs.length);
    },
  },
  {
    name: 'HARMEAN', min: 1, max: 255, category: S, arrays: 'all', desc: 'Harmonic mean of positive numbers.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => {
      const xs = nums(args, ctx);
      if (!xs.length || xs.some((x) => x <= 0)) return ERR.num;
      return xs.length / sum(xs.map((x) => 1 / x));
    },
  },
  {
    name: 'TRIMMEAN', min: 2, max: 2, category: S, arrays: [0], desc: 'Mean of the interior of a data set, excluding a percentage of extreme values.', params: ['array', 'percent'],
    impl: ([a, p], ctx) => {
      const xs = arrNums(a, ctx), pc = num(p as Scalar);
      if (pc < 0 || pc >= 1 || !xs.length) return ERR.num;
      const cut = Math.floor((xs.length * pc) / 2);
      const s = [...xs].sort((x, y) => x - y).slice(cut, xs.length - cut);
      return mean(s);
    },
  },
  {
    name: 'MEDIAN', min: 1, max: 255, category: S, arrays: 'all', desc: 'Middle value of the given numbers.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => median(nums(args, ctx)),
  },
  {
    name: 'MODE.SNGL', min: 1, max: 255, category: S, arrays: 'all', desc: 'Most frequently occurring value.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => modeSngl(nums(args, ctx)),
  },
  {
    name: 'MODE.MULT', min: 1, max: 255, category: S, arrays: 'all', desc: 'Vertical array of the most frequently occurring values.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => { const m = modes(nums(args, ctx)); return m.length ? m.map((x) => [x]) : ERR.na; },
  },
  {
    name: 'MAX', min: 1, max: 255, category: S, arrays: 'all', desc: 'Largest number in the arguments.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => { const xs = nums(args, ctx); let m = -Infinity; for (const x of xs) if (x > m) m = x; return xs.length ? m : 0; },
  },
  {
    name: 'MIN', min: 1, max: 255, category: S, arrays: 'all', desc: 'Smallest number in the arguments.', params: ['number1', '[number2]', '...'],
    impl: (args, ctx) => { const xs = nums(args, ctx); let m = Infinity; for (const x of xs) if (x < m) m = x; return xs.length ? m : 0; },
  },
  {
    name: 'MAXA', min: 1, max: 255, category: S, arrays: 'all', desc: 'Largest value, counting text as 0 and logical values as 1/0.', params: ['value1', '[value2]', '...'],
    impl: (args, ctx) => { const xs = numsA(args, ctx); let m = -Infinity; for (const x of xs) if (x > m) m = x; return xs.length ? m : 0; },
  },
  {
    name: 'MINA', min: 1, max: 255, category: S, arrays: 'all', desc: 'Smallest value, counting text as 0 and logical values as 1/0.', params: ['value1', '[value2]', '...'],
    impl: (args, ctx) => { const xs = numsA(args, ctx); let m = Infinity; for (const x of xs) if (x < m) m = x; return xs.length ? m : 0; },
  },
  {
    name: 'LARGE', min: 2, max: 2, category: S, arrays: [0], desc: 'k-th largest value in a data set.', params: ['array', 'k'],
    impl: ([a, k], ctx) => kth(arrNums(a, ctx), num(k as Scalar), true),
  },
  {
    name: 'SMALL', min: 2, max: 2, category: S, arrays: [0], desc: 'k-th smallest value in a data set.', params: ['array', 'k'],
    impl: ([a, k], ctx) => kth(arrNums(a, ctx), num(k as Scalar), false),
  },
  {
    name: 'PERCENTILE.INC', min: 2, max: 2, category: S, arrays: [0], desc: 'k-th percentile (k from 0 to 1 inclusive).', params: ['array', 'k'],
    impl: ([a, k], ctx) => percentileInc(arrNums(a, ctx), num(k as Scalar)),
  },
  {
    name: 'PERCENTILE.EXC', min: 2, max: 2, category: S, arrays: [0], desc: 'k-th percentile (k strictly between 0 and 1).', params: ['array', 'k'],
    impl: ([a, k], ctx) => percentileExc(arrNums(a, ctx), num(k as Scalar)),
  },
  {
    name: 'QUARTILE.INC', min: 2, max: 2, category: S, arrays: [0], desc: 'Quartile (0 to 4) of a data set, inclusive method.', params: ['array', 'quart'],
    impl: ([a, q], ctx) => quartileInc(arrNums(a, ctx), num(q as Scalar)),
  },
  {
    name: 'QUARTILE.EXC', min: 2, max: 2, category: S, arrays: [0], desc: 'Quartile (1 to 3) of a data set, exclusive method.', params: ['array', 'quart'],
    impl: ([a, q], ctx) => quartileExc(arrNums(a, ctx), num(q as Scalar)),
  },
  {
    name: 'PERCENTRANK.INC', min: 2, max: 3, category: S, arrays: [0], desc: 'Rank of a value as a percentage (0 to 1 inclusive) of a data set.', params: ['array', 'x', '[significance]'],
    impl: (a, ctx) => percentRank(arrNums(a[0], ctx), num(sc(a, 1)), optN(a, 2, 3), false),
  },
  {
    name: 'PERCENTRANK.EXC', min: 2, max: 3, category: S, arrays: [0], desc: 'Rank of a value as a percentage (exclusive of 0 and 1) of a data set.', params: ['array', 'x', '[significance]'],
    impl: (a, ctx) => percentRank(arrNums(a[0], ctx), num(sc(a, 1)), optN(a, 2, 3), true),
  },
  {
    name: 'RANK.EQ', min: 2, max: 3, category: S, arrays: [1], desc: 'Rank of a number in a list; ties share the top rank.', params: ['number', 'ref', '[order]'],
    impl: (a, ctx) => rank(num(sc(a, 0)), arrNums(a[1], ctx), optN(a, 2, 0) !== 0, false),
  },
  {
    name: 'RANK.AVG', min: 2, max: 3, category: S, arrays: [1], desc: 'Rank of a number in a list; ties get the average rank.', params: ['number', 'ref', '[order]'],
    impl: (a, ctx) => rank(num(sc(a, 0)), arrNums(a[1], ctx), optN(a, 2, 0) !== 0, true),
  },
  {
    name: 'FREQUENCY', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Counts how many values fall into each bin; returns a vertical array.', params: ['data_array', 'bins_array'],
    impl: ([d, b], ctx) => {
      const xs = arrNums(d, ctx), bins = arrNums(b, ctx);
      if (!bins.length) return [[xs.length]];
      const order = bins.map((v, i) => [v, i] as [number, number]).sort((p, q) => p[0] - q[0]);
      const counts = new Array(bins.length + 1).fill(0);
      for (const x of xs) {
        let placed = false;
        for (const [v, i] of order) if (x <= v) { counts[i]++; placed = true; break; }
        if (!placed) counts[bins.length]++;
      }
      return counts.map((c) => [c]);
    },
  },
  // --- Spread ------------------------------------------------------------------------------
  { name: 'STDEV.S', min: 1, max: 255, category: S, arrays: 'all', desc: 'Standard deviation of a sample.', params: ['number1', '[number2]', '...'], impl: (a, ctx) => stdev(nums(a, ctx), true) },
  { name: 'STDEV.P', min: 1, max: 255, category: S, arrays: 'all', desc: 'Standard deviation of an entire population.', params: ['number1', '[number2]', '...'], impl: (a, ctx) => stdev(nums(a, ctx), false) },
  { name: 'STDEVA', min: 1, max: 255, category: S, arrays: 'all', desc: 'Sample standard deviation, counting text as 0 and logical values as 1/0.', params: ['value1', '[value2]', '...'], impl: (a, ctx) => stdev(numsA(a, ctx), true) },
  { name: 'STDEVPA', min: 1, max: 255, category: S, arrays: 'all', desc: 'Population standard deviation, counting text as 0 and logical values as 1/0.', params: ['value1', '[value2]', '...'], impl: (a, ctx) => stdev(numsA(a, ctx), false) },
  { name: 'VAR.S', min: 1, max: 255, category: S, arrays: 'all', desc: 'Variance of a sample.', params: ['number1', '[number2]', '...'], impl: (a, ctx) => variance(nums(a, ctx), true) },
  { name: 'VAR.P', min: 1, max: 255, category: S, arrays: 'all', desc: 'Variance of an entire population.', params: ['number1', '[number2]', '...'], impl: (a, ctx) => variance(nums(a, ctx), false) },
  { name: 'VARA', min: 1, max: 255, category: S, arrays: 'all', desc: 'Sample variance, counting text as 0 and logical values as 1/0.', params: ['value1', '[value2]', '...'], impl: (a, ctx) => variance(numsA(a, ctx), true) },
  { name: 'VARPA', min: 1, max: 255, category: S, arrays: 'all', desc: 'Population variance, counting text as 0 and logical values as 1/0.', params: ['value1', '[value2]', '...'], impl: (a, ctx) => variance(numsA(a, ctx), false) },
  { name: 'KURT', min: 1, max: 255, category: S, arrays: 'all', desc: 'Excess kurtosis of a data set.', params: ['number1', '[number2]', '...'], impl: (a, ctx) => kurt(nums(a, ctx)) },
  { name: 'SKEW', min: 1, max: 255, category: S, arrays: 'all', desc: 'Sample skewness of a distribution.', params: ['number1', '[number2]', '...'], impl: (a, ctx) => skew(nums(a, ctx), false) },
  { name: 'SKEW.P', min: 1, max: 255, category: S, arrays: 'all', desc: 'Population skewness of a distribution.', params: ['number1', '[number2]', '...'], impl: (a, ctx) => skew(nums(a, ctx), true) },
  // --- Paired data -------------------------------------------------------------------------
  {
    name: 'CORREL', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Correlation coefficient of two data sets.', params: ['array1', 'array2'],
    impl: ([a, b], ctx) => correl(...pairs(a, b, ctx)),
  },
  {
    name: 'PEARSON', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Pearson product-moment correlation coefficient.', params: ['array1', 'array2'],
    impl: ([a, b], ctx) => correl(...pairs(a, b, ctx)),
  },
  {
    name: 'RSQ', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Square of the Pearson correlation coefficient.', params: ['known_ys', 'known_xs'],
    impl: ([a, b], ctx) => correl(...pairs(a, b, ctx)) ** 2,
  },
  {
    name: 'COVARIANCE.P', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Population covariance of two data sets.', params: ['array1', 'array2'],
    impl: ([a, b], ctx) => covar(...pairs(a, b, ctx), false),
  },
  {
    name: 'COVARIANCE.S', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Sample covariance of two data sets.', params: ['array1', 'array2'],
    impl: ([a, b], ctx) => covar(...pairs(a, b, ctx), true),
  },
  {
    name: 'SLOPE', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Slope of the linear regression line.', params: ['known_ys', 'known_xs'],
    impl: ([a, b], ctx) => { const [y, x] = pairs(a, b, ctx); return lin(x, y).slope; },
  },
  {
    name: 'INTERCEPT', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Point where the linear regression line crosses the y-axis.', params: ['known_ys', 'known_xs'],
    impl: ([a, b], ctx) => { const [y, x] = pairs(a, b, ctx); return lin(x, y).icpt; },
  },
  {
    name: 'STEYX', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Standard error of the predicted y for each x in a regression.', params: ['known_ys', 'known_xs'],
    impl: ([a, b], ctx) => {
      const [y, x] = pairs(a, b, ctx);
      const n = x.length;
      if (n < 3) return ERR.div0;
      const { sxx, sxy, syy } = lin(x, y);
      return Math.sqrt(Math.max(0, syy - (sxy * sxy) / sxx) / (n - 2));
    },
  },
  {
    name: 'FORECAST.LINEAR', min: 3, max: 3, category: S, arrays: [1, 2], desc: 'Predicts a value along a linear trend fitted to known data.', params: ['x', 'known_ys', 'known_xs'],
    impl: ([x, a, b], ctx) => { const [y, xs] = pairs(a, b, ctx); const l = lin(xs, y); return l.icpt + l.slope * num(x as Scalar); },
  },
  {
    name: 'LINEST', min: 1, max: 4, category: S, arrays: [0, 1], desc: 'Least-squares line (or plane) statistics for known data.', params: ['known_ys', '[known_xs]', '[const]', '[stats]'],
    impl: (args, ctx) => {
      const d = regData(args[0], args.length > 1 && args[1] != null ? args[1] : undefined, ctx);
      const useConst = optB(args, 2, true), stats = optB(args, 3, false);
      return linestOut(fitLinear(d.y, d.X, useConst), useConst, stats, (x) => x);
    },
  },
  {
    name: 'LOGEST', min: 1, max: 4, category: S, arrays: [0, 1], desc: 'Exponential curve statistics fitted to known data.', params: ['known_ys', '[known_xs]', '[const]', '[stats]'],
    impl: (args, ctx) => {
      const d = regData(args[0], args.length > 1 && args[1] != null ? args[1] : undefined, ctx);
      if (d.y.some((v) => v <= 0)) return ERR.num;
      const useConst = optB(args, 2, true), stats = optB(args, 3, false);
      const fit = fitLinear(d.y.map(Math.log), d.X, useConst);
      const out = linestOut(fit, useConst, stats, Math.exp);
      return out;
    },
  },
  {
    name: 'TREND', min: 1, max: 4, category: S, arrays: [0, 1, 2], desc: 'Values along a linear trend fitted to known data.', params: ['known_ys', '[known_xs]', '[new_xs]', '[const]'],
    impl: (args, ctx) => trendLike(args, ctx, false),
  },
  {
    name: 'GROWTH', min: 1, max: 4, category: S, arrays: [0, 1, 2], desc: 'Values along an exponential growth curve fitted to known data.', params: ['known_ys', '[known_xs]', '[new_xs]', '[const]'],
    impl: (args, ctx) => trendLike(args, ctx, true),
  },
  {
    name: 'FORECAST.ETS', min: 3, max: 6, category: S, arrays: [1, 2], desc: 'Forecasts a future value using exponential triple smoothing.', params: ['target_date', 'values', 'timeline', '[seasonality]', '[data_completion]', '[aggregation]'],
    impl: (args, ctx) => {
      const o = etsArgs(args, ctx, 3);
      return etsForecast(etsFit(args[1], args[2], ctx, o.seas, o.comp, o.agg), num(sc(args, 0)));
    },
  },
  {
    name: 'FORECAST.ETS.CONFINT', min: 3, max: 7, category: S, arrays: [1, 2], desc: 'Confidence interval half-width for a FORECAST.ETS prediction.', params: ['target_date', 'values', 'timeline', '[confidence_level]', '[seasonality]', '[data_completion]', '[aggregation]'],
    impl: (args, ctx) => {
      const conf = optN(args, 3, 0.95);
      if (conf <= 0 || conf >= 1) return ERR.num;
      const o = etsArgs(args, ctx, 4);
      const md = etsFit(args[1], args[2], ctx, o.seas, o.comp, o.agg);
      const h = Math.ceil((num(sc(args, 0)) - md.tLast) / md.step);
      if (h <= 0) return ERR.num;
      let v = 1;
      for (let j = 1; j < h; j++) v += (md.alpha * (1 + j * md.beta) + (md.period > 1 && j % md.period === 0 ? md.gamma * (1 - md.alpha) : 0)) ** 2;
      return normInv(1 - (1 - conf) / 2) * md.rmse * Math.sqrt(v);
    },
  },
  {
    name: 'FORECAST.ETS.SEASONALITY', min: 2, max: 4, category: S, arrays: [0, 1], desc: 'Length of the repetitive pattern FORECAST.ETS detects.', params: ['values', 'timeline', '[data_completion]', '[aggregation]'],
    impl: (args, ctx) => {
      const comp = optN(args, 2, 1), agg = optN(args, 3, 1);
      const p = etsFit(args[0], args[1], ctx, null, comp, agg).period;
      return p === 1 ? 0 : p;
    },
  },
  {
    name: 'FORECAST.ETS.STAT', min: 3, max: 6, category: S, arrays: [0, 1], desc: 'A statistic of the FORECAST.ETS model (alpha, beta, gamma, MASE, SMAPE, MAE, RMSE, step).', params: ['values', 'timeline', 'statistic_type', '[seasonality]', '[data_completion]', '[aggregation]'],
    impl: (args, ctx) => {
      const t = tr(num(sc(args, 2)));
      const o = etsArgs(args, ctx, 3);
      const md = etsFit(args[0], args[1], ctx, o.seas, o.comp, o.agg);
      switch (t) {
        case 1: return md.alpha; case 2: return md.beta; case 3: return md.gamma; case 4: return md.mase;
        case 5: return md.smape; case 6: return md.mae; case 7: return md.rmse; case 8: return md.step;
        default: return ERR.num;
      }
    },
  },
  // --- Tests -------------------------------------------------------------------------------
  {
    name: 'T.TEST', min: 4, max: 4, category: S, arrays: [0, 1], desc: "Probability associated with a Student's t-test.", params: ['array1', 'array2', 'tails', 'type'],
    impl: (a, ctx) => ttest(a, ctx),
  },
  {
    name: 'F.TEST', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Two-tailed probability that the variances of two data sets are not significantly different.', params: ['array1', 'array2'],
    impl: ([a, b], ctx) => {
      const x = arrNums(a, ctx), y = arrNums(b, ctx);
      if (x.length < 2 || y.length < 2) return ERR.div0;
      const v1 = variance(x, true), v2 = variance(y, true);
      if (v1 === 0 || v2 === 0) return ERR.div0;
      const f = v1 / v2;
      const p = fCdf(f, x.length - 1, y.length - 1);
      return 2 * Math.min(p, 1 - p);
    },
  },
  {
    name: 'CHISQ.TEST', min: 2, max: 2, category: S, arrays: [0, 1], desc: 'Chi-squared test for independence of observed vs. expected values.', params: ['actual_range', 'expected_range'],
    impl: ([a, e], ctx) => {
      const am = toMatrix(a, ctx), em = toMatrix(e, ctx);
      const r = am.length, c = am[0]?.length ?? 0;
      if (em.length !== r || (em[0]?.length ?? 0) !== c) return ERR.na;
      let chi = 0;
      for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) {
        const x = am[i][j], y = em[i][j];
        if (x instanceof CellError) return x;
        if (y instanceof CellError) return y;
        if (typeof x !== 'number' || typeof y !== 'number') continue;
        if (y === 0) return ERR.div0;
        chi += (x - y) ** 2 / y;
      }
      const df = r > 1 && c > 1 ? (r - 1) * (c - 1) : r * c - 1;
      if (df < 1) return ERR.na;
      return chiSf(chi, df);
    },
  },
  {
    name: 'Z.TEST', min: 2, max: 3, category: S, arrays: [0], desc: 'One-tailed P-value of a z-test.', params: ['array', 'x', '[sigma]'],
    impl: (a, ctx) => {
      const xs = arrNums(a[0], ctx);
      const x = num(sc(a, 1));
      if (!xs.length) return ERR.na;
      const s = a.length > 2 && a[2] != null ? num(sc(a, 2)) : stdev(xs, true);
      if (s <= 0) return s === 0 ? ERR.div0 : ERR.num;
      return normSf((mean(xs) - x) / (s / Math.sqrt(xs.length)));
    },
  },
  {
    name: 'PROB', min: 3, max: 4, category: S, arrays: [0, 1], desc: 'Probability that values in a range lie between two limits.', params: ['x_range', 'prob_range', 'lower_limit', '[upper_limit]'],
    impl: (a, ctx) => {
      const xs = flatVals(a[0], ctx), ps = flatVals(a[1], ctx);
      if (xs.length !== ps.length) return ERR.na;
      const lo = num(sc(a, 2));
      const hi = a.length > 3 && a[3] != null ? num(sc(a, 3)) : lo;
      let total = 0, s = 0;
      for (let i = 0; i < xs.length; i++) {
        const x = xs[i], p = ps[i];
        if (x instanceof CellError) return x;
        if (p instanceof CellError) return p;
        if (typeof p !== 'number' || typeof x !== 'number') continue;
        if (p < 0 || p > 1) return ERR.num;
        total += p;
        if (x >= lo && x <= hi) s += p;
      }
      if (Math.abs(total - 1) > 1e-7) return ERR.num;
      return s;
    },
  },
  // --- Distributions -----------------------------------------------------------------------
  dist('NORM.DIST', ['x', 'mean', 'standard_dev', 'cumulative'], 'Normal distribution for a mean and standard deviation.', 4, (a) => {
    const x = num(sc(a, 0)), m = num(sc(a, 1)), s = num(sc(a, 2)), cum = bool(sc(a, 3));
    if (s <= 0) return ERR.num;
    const z = (x - m) / s;
    return cum ? normCdf(z) : normPdf(z) / s;
  }),
  dist('NORM.INV', ['probability', 'mean', 'standard_dev'], 'Inverse of the normal cumulative distribution.', 3, (a) => {
    const p = num(sc(a, 0)), m = num(sc(a, 1)), s = num(sc(a, 2));
    if (p <= 0 || p >= 1 || s <= 0) return ERR.num;
    return m + s * normInv(p);
  }),
  dist('NORM.S.DIST', ['z', 'cumulative'], 'Standard normal distribution (mean 0, standard deviation 1).', 2, (a) => {
    const z = num(sc(a, 0));
    return bool(sc(a, 1)) ? normCdf(z) : normPdf(z);
  }),
  dist('NORM.S.INV', ['probability'], 'Inverse of the standard normal cumulative distribution.', 1, (a) => {
    const p = num(sc(a, 0));
    if (p <= 0 || p >= 1) return ERR.num;
    return normInv(p);
  }),
  dist('LOGNORM.DIST', ['x', 'mean', 'standard_dev', 'cumulative'], 'Lognormal distribution of x, where ln(x) is normally distributed.', 4, (a) => {
    const x = num(sc(a, 0)), m = num(sc(a, 1)), s = num(sc(a, 2)), cum = bool(sc(a, 3));
    if (x <= 0 || s <= 0) return ERR.num;
    const z = (Math.log(x) - m) / s;
    return cum ? normCdf(z) : normPdf(z) / (s * x);
  }),
  dist('LOGNORM.INV', ['probability', 'mean', 'standard_dev'], 'Inverse of the lognormal cumulative distribution.', 3, (a) => {
    const p = num(sc(a, 0)), m = num(sc(a, 1)), s = num(sc(a, 2));
    if (p <= 0 || p >= 1 || s <= 0) return ERR.num;
    return Math.exp(m + s * normInv(p));
  }),
  dist('STANDARDIZE', ['x', 'mean', 'standard_dev'], 'Normalized value (z-score) of x.', 3, (a) => {
    const x = num(sc(a, 0)), m = num(sc(a, 1)), s = num(sc(a, 2));
    if (s <= 0) return ERR.num;
    return (x - m) / s;
  }),
  dist('PHI', ['x'], 'Density of the standard normal distribution at x.', 1, (a) => normPdf(num(sc(a, 0)))),
  dist('GAUSS', ['z'], 'Probability that a standard normal value lies between the mean and z standard deviations from it.', 1, (a) => normCdf(num(sc(a, 0))) - 0.5),
  dist('CONFIDENCE.NORM', ['alpha', 'standard_dev', 'size'], 'Half-width of a confidence interval for a population mean using the normal distribution.', 3, (a) => {
    const al = num(sc(a, 0)), s = num(sc(a, 1)), n = tr(num(sc(a, 2)));
    if (al <= 0 || al >= 1 || s <= 0 || n < 1) return ERR.num;
    return (normInv(1 - al / 2) * s) / Math.sqrt(n);
  }),
  dist('CONFIDENCE.T', ['alpha', 'standard_dev', 'size'], "Half-width of a confidence interval for a population mean using Student's t.", 3, (a) => {
    const al = num(sc(a, 0)), s = num(sc(a, 1)), n = tr(num(sc(a, 2)));
    if (al <= 0 || al >= 1 || s <= 0 || n < 1) return ERR.num;
    if (n === 1) return ERR.div0;
    return (tInv(1 - al / 2, n - 1) * s) / Math.sqrt(n);
  }),
  dist('T.DIST', ['x', 'deg_freedom', 'cumulative'], "Left-tailed Student's t-distribution.", 3, (a) => {
    const x = num(sc(a, 0)), df = tr(num(sc(a, 1))), cum = bool(sc(a, 2));
    if (df < 1) return ERR.num;
    return cum ? tCdf(x, df) : tPdf(x, df);
  }),
  dist('T.DIST.2T', ['x', 'deg_freedom'], "Two-tailed Student's t-distribution.", 2, (a) => {
    const x = num(sc(a, 0)), df = tr(num(sc(a, 1)));
    if (df < 1 || x < 0) return ERR.num;
    return 2 * tSf(x, df);
  }),
  dist('T.DIST.RT', ['x', 'deg_freedom'], "Right-tailed Student's t-distribution.", 2, (a) => {
    const x = num(sc(a, 0)), df = tr(num(sc(a, 1)));
    if (df < 1) return ERR.num;
    return tSf(x, df);
  }),
  dist('T.INV', ['probability', 'deg_freedom'], "Left-tailed inverse of Student's t-distribution.", 2, (a) => {
    const p = num(sc(a, 0)), df = tr(num(sc(a, 1)));
    if (p <= 0 || p >= 1 || df < 1) return ERR.num;
    return tInv(p, df);
  }),
  dist('T.INV.2T', ['probability', 'deg_freedom'], "Two-tailed inverse of Student's t-distribution.", 2, (a) => {
    const p = num(sc(a, 0)), df = tr(num(sc(a, 1)));
    if (p <= 0 || p > 1 || df < 1) return ERR.num;
    return Math.abs(tInv(1 - p / 2, df));
  }),
  dist('CHISQ.DIST', ['x', 'deg_freedom', 'cumulative'], 'Left-tailed chi-squared distribution.', 3, (a) => {
    const x = num(sc(a, 0)), df = tr(num(sc(a, 1))), cum = bool(sc(a, 2));
    if (x < 0 || df < 1 || df > 1e10) return ERR.num;
    return cum ? chiCdf(x, df) : chiPdf(x, df);
  }),
  dist('CHISQ.DIST.RT', ['x', 'deg_freedom'], 'Right-tailed chi-squared distribution.', 2, (a) => {
    const x = num(sc(a, 0)), df = tr(num(sc(a, 1)));
    if (x < 0 || df < 1 || df > 1e10) return ERR.num;
    return chiSf(x, df);
  }),
  dist('CHISQ.INV', ['probability', 'deg_freedom'], 'Inverse of the left-tailed chi-squared distribution.', 2, (a) => {
    const p = num(sc(a, 0)), df = tr(num(sc(a, 1)));
    if (p < 0 || p >= 1 || df < 1 || df > 1e10) return ERR.num;
    if (p === 0) return 0;
    return invert((x) => chiCdf(x, df), p, 0, Math.max(df, 1));
  }),
  dist('CHISQ.INV.RT', ['probability', 'deg_freedom'], 'Inverse of the right-tailed chi-squared distribution.', 2, (a) => {
    const p = num(sc(a, 0)), df = tr(num(sc(a, 1)));
    if (p <= 0 || p > 1 || df < 1 || df > 1e10) return ERR.num;
    if (p === 1) return 0;
    return invert((x) => -chiSf(x, df), -p, 0, Math.max(df, 1));
  }),
  dist('F.DIST', ['x', 'deg_freedom1', 'deg_freedom2', 'cumulative'], 'Left-tailed F probability distribution.', 4, (a) => {
    const x = num(sc(a, 0)), d1 = tr(num(sc(a, 1))), d2 = tr(num(sc(a, 2))), cum = bool(sc(a, 3));
    if (x < 0 || d1 < 1 || d2 < 1) return ERR.num;
    return cum ? fCdf(x, d1, d2) : fPdf(x, d1, d2);
  }),
  dist('F.DIST.RT', ['x', 'deg_freedom1', 'deg_freedom2'], 'Right-tailed F probability distribution.', 3, (a) => {
    const x = num(sc(a, 0)), d1 = tr(num(sc(a, 1))), d2 = tr(num(sc(a, 2)));
    if (x < 0 || d1 < 1 || d2 < 1) return ERR.num;
    return fSf(x, d1, d2);
  }),
  dist('F.INV', ['probability', 'deg_freedom1', 'deg_freedom2'], 'Inverse of the left-tailed F distribution.', 3, (a) => {
    const p = num(sc(a, 0)), d1 = tr(num(sc(a, 1))), d2 = tr(num(sc(a, 2)));
    if (p < 0 || p > 1 || d1 < 1 || d2 < 1) return ERR.num;
    if (p === 0) return 0;
    if (p === 1) return ERR.num;
    return invert((x) => fCdf(x, d1, d2), p, 0, 1);
  }),
  dist('F.INV.RT', ['probability', 'deg_freedom1', 'deg_freedom2'], 'Inverse of the right-tailed F distribution.', 3, (a) => {
    const p = num(sc(a, 0)), d1 = tr(num(sc(a, 1))), d2 = tr(num(sc(a, 2)));
    if (p <= 0 || p > 1 || d1 < 1 || d2 < 1) return ERR.num;
    if (p === 1) return 0;
    return invert((x) => -fSf(x, d1, d2), -p, 0, 1);
  }),
  dist('BETA.DIST', ['x', 'alpha', 'beta', 'cumulative', '[A]', '[B]'], 'Beta distribution, optionally scaled to the interval [A, B].', 4, (a) => {
    const x = num(sc(a, 0)), al = num(sc(a, 1)), be = num(sc(a, 2)), cum = bool(sc(a, 3));
    const A = optN(a, 4, 0), B = optN(a, 5, 1);
    if (al <= 0 || be <= 0 || x < A || x > B || A === B) return ERR.num;
    const z = (x - A) / (B - A);
    return cum ? betaInc(z, al, be) : betaPdf(z, al, be) / (B - A);
  }),
  dist('BETA.INV', ['probability', 'alpha', 'beta', '[A]', '[B]'], 'Inverse of the cumulative beta distribution.', 3, (a) => betaInvF(a)),
  dist('BINOM.DIST', ['number_s', 'trials', 'probability_s', 'cumulative'], 'Individual-term binomial distribution probability.', 4, (a) => {
    const k = tr(num(sc(a, 0))), n = tr(num(sc(a, 1))), p = num(sc(a, 2)), cum = bool(sc(a, 3));
    if (k < 0 || k > n || p < 0 || p > 1) return ERR.num;
    return cum ? binomCdf(k, n, p) : binomPmf(k, n, p);
  }),
  dist('BINOM.DIST.RANGE', ['trials', 'probability_s', 'number_s', '[number_s2]'], 'Probability that the number of successes lies in a range.', 3, (a) => {
    const n = tr(num(sc(a, 0))), p = num(sc(a, 1)), s1 = tr(num(sc(a, 2)));
    const s2 = a.length > 3 && a[3] != null ? tr(num(sc(a, 3))) : s1;
    if (n < 0 || p < 0 || p > 1 || s1 < 0 || s1 > n || s2 < s1 || s2 > n) return ERR.num;
    let s = 0;
    for (let k = s1; k <= s2; k++) s += binomPmf(k, n, p);
    return s;
  }),
  dist('BINOM.INV', ['trials', 'probability_s', 'alpha'], 'Smallest number of successes whose cumulative binomial probability is at least alpha.', 3, (a) => binomInv(a)),
  dist('NEGBINOM.DIST', ['number_f', 'number_s', 'probability_s', 'cumulative'], 'Negative binomial distribution.', 4, (a) => {
    const f = tr(num(sc(a, 0))), s = tr(num(sc(a, 1))), p = num(sc(a, 2)), cum = bool(sc(a, 3));
    if (f < 0 || s < 1 || p < 0 || p > 1) return ERR.num;
    if (cum) return betaInc(p, s, f + 1);
    return negbinPmf(f, s, p);
  }),
  dist('HYPGEOM.DIST', ['sample_s', 'number_sample', 'population_s', 'number_pop', 'cumulative'], 'Hypergeometric distribution.', 5, (a) => {
    const s = tr(num(sc(a, 0))), n = tr(num(sc(a, 1))), S2 = tr(num(sc(a, 2))), N = tr(num(sc(a, 3))), cum = bool(sc(a, 4));
    const e = hypgeomCheck(s, n, S2, N);
    if (e) return e;
    if (!cum) return hypgeomPmf(s, n, S2, N);
    let t = 0;
    for (let i = Math.max(0, n - N + S2); i <= s; i++) t += hypgeomPmf(i, n, S2, N);
    return Math.min(1, t);
  }),
  dist('POISSON.DIST', ['x', 'mean', 'cumulative'], 'Poisson distribution.', 3, (a) => {
    const x = tr(num(sc(a, 0))), m = num(sc(a, 1)), cum = bool(sc(a, 2));
    if (x < 0 || m < 0) return ERR.num;
    if (!cum) return poissonPmf(x, m);
    if (m === 0) return 1;
    if (x < 1000) { let s = 0; for (let i = 0; i <= x; i++) s += poissonPmf(i, m); return Math.min(1, s); }
    return gammaQ(x + 1, m);
  }),
  dist('EXPON.DIST', ['x', 'lambda', 'cumulative'], 'Exponential distribution.', 3, (a) => {
    const x = num(sc(a, 0)), l = num(sc(a, 1)), cum = bool(sc(a, 2));
    if (x < 0 || l <= 0) return ERR.num;
    return cum ? -Math.expm1(-l * x) : l * Math.exp(-l * x);
  }),
  dist('GAMMA.DIST', ['x', 'alpha', 'beta', 'cumulative'], 'Gamma distribution.', 4, (a) => {
    const x = num(sc(a, 0)), al = num(sc(a, 1)), be = num(sc(a, 2)), cum = bool(sc(a, 3));
    if (x < 0 || al <= 0 || be <= 0) return ERR.num;
    return cum ? gammaCdf(x, al, be) : gammaPdf(x, al, be);
  }),
  dist('GAMMA.INV', ['probability', 'alpha', 'beta'], 'Inverse of the cumulative gamma distribution.', 3, (a) => {
    const p = num(sc(a, 0)), al = num(sc(a, 1)), be = num(sc(a, 2));
    if (p < 0 || p >= 1 || al <= 0 || be <= 0) return ERR.num;
    return gammaInv(p, al, be);
  }),
  dist('GAMMA', ['number'], 'Value of the gamma function.', 1, (a) => {
    const x = num(sc(a, 0));
    if (x <= 0 && Number.isInteger(x)) return ERR.num;
    return chkNum(gamma(x));
  }),
  dist('GAMMALN', ['x'], 'Natural logarithm of the gamma function.', 1, (a) => {
    const x = num(sc(a, 0));
    if (x <= 0) return ERR.num;
    return gammaln(x);
  }),
  dist('GAMMALN.PRECISE', ['x'], 'Natural logarithm of the gamma function.', 1, (a) => {
    const x = num(sc(a, 0));
    if (x <= 0) return ERR.num;
    return gammaln(x);
  }),
  dist('WEIBULL.DIST', ['x', 'alpha', 'beta', 'cumulative'], 'Weibull distribution.', 4, (a) => {
    const x = num(sc(a, 0)), al = num(sc(a, 1)), be = num(sc(a, 2)), cum = bool(sc(a, 3));
    if (x < 0 || al <= 0 || be <= 0) return ERR.num;
    const t = Math.pow(x / be, al);
    return cum ? -Math.expm1(-t) : (al / Math.pow(be, al)) * Math.pow(x, al - 1) * Math.exp(-t);
  }),
  dist('FISHER', ['x'], 'Fisher transformation of x.', 1, (a) => {
    const x = num(sc(a, 0));
    if (x <= -1 || x >= 1) return ERR.num;
    return 0.5 * Math.log((1 + x) / (1 - x));
  }),
  dist('FISHERINV', ['y'], 'Inverse of the Fisher transformation.', 1, (a) => Math.tanh(num(sc(a, 0)))),
  dist('PERMUT', ['number', 'number_chosen'], 'Number of ordered arrangements of k items chosen from n.', 2, (a) => {
    const n = tr(num(sc(a, 0))), k = tr(num(sc(a, 1)));
    if (n < 0 || k < 0 || n < k) return ERR.num;
    let r = 1;
    for (let i = 0; i < k; i++) r *= n - i;
    return chkNum(r);
  }),
  dist('PERMUTATIONA', ['number', 'number_chosen'], 'Number of arrangements of k items chosen from n with repetition.', 2, (a) => {
    const n = tr(num(sc(a, 0))), k = tr(num(sc(a, 1)));
    if (n < 0 || k < 0) return ERR.num;
    return chkNum(Math.pow(n, k));
  }),
];

function tryNum(s: string): number {
  try { return num(s); } catch { return NaN; }
}

function correl(x: number[], y: number[]): number {
  const n = x.length;
  if (n < 2) return fail(ERR.div0);
  const { sxx, syy, sxy } = lin(x, y, true);
  if (sxx === 0 || syy === 0) return fail(ERR.div0);
  return sxy / Math.sqrt(sxx * syy);
}

function covar(x: number[], y: number[], sample: boolean): number {
  const n = x.length;
  if (n === 0 || (sample && n < 2)) return fail(ERR.div0);
  const mx = mean(x), my = mean(y);
  let s = 0;
  for (let i = 0; i < n; i++) s += (x[i] - mx) * (y[i] - my);
  return s / (sample ? n - 1 : n);
}

function lin(x: number[], y: number[], noFail = false): { slope: number; icpt: number; sxx: number; syy: number; sxy: number } {
  const n = x.length;
  if (!n) return fail(ERR.div0);
  const mx = mean(x), my = mean(y);
  let sxx = 0, syy = 0, sxy = 0;
  for (let i = 0; i < n; i++) { sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; sxy += (x[i] - mx) * (y[i] - my); }
  if (sxx === 0 && !noFail) return fail(ERR.div0);
  const slope = sxy / sxx;
  return { slope, icpt: my - slope * mx, sxx, syy, sxy };
}

function ttest(a: Args, ctx: Ctx): Value {
  const tails = tr(num(sc(a, 2))), type = tr(num(sc(a, 3)));
  if (tails !== 1 && tails !== 2) return ERR.num;
  if (type < 1 || type > 3) return ERR.num;
  let t: number, df: number;
  if (type === 1) {
    const [x, y] = pairs(a[0], a[1], ctx);
    const n = x.length;
    if (n < 2) return ERR.div0;
    const d = x.map((v, i) => v - y[i]);
    const sd = stdev(d, true);
    if (sd === 0) return ERR.div0;
    t = mean(d) / (sd / Math.sqrt(n));
    df = n - 1;
  } else {
    const x = arrNums(a[0], ctx), y = arrNums(a[1], ctx);
    const n1 = x.length, n2 = y.length;
    if (n1 < 2 || n2 < 2) return ERR.div0;
    const v1 = variance(x, true), v2 = variance(y, true);
    const m1 = mean(x), m2 = mean(y);
    if (type === 2) {
      df = n1 + n2 - 2;
      const sp = ((n1 - 1) * v1 + (n2 - 1) * v2) / df;
      const den = Math.sqrt(sp * (1 / n1 + 1 / n2));
      if (den === 0) return ERR.div0;
      t = (m1 - m2) / den;
    } else {
      const q1 = v1 / n1, q2 = v2 / n2;
      if (q1 + q2 === 0) return ERR.div0;
      t = (m1 - m2) / Math.sqrt(q1 + q2);
      df = (q1 + q2) ** 2 / (q1 * q1 / (n1 - 1) + q2 * q2 / (n2 - 1));
    }
  }
  return tails * tSf(Math.abs(t), df);
}

function betaInvF(a: Args): Value {
  const p = num(sc(a, 0)), al = num(sc(a, 1)), be = num(sc(a, 2));
  const A = optN(a, 3, 0), B = optN(a, 4, 1);
  if (p <= 0 || p > 1 || al <= 0 || be <= 0 || A >= B) return ERR.num;
  if (p === 1) return B;
  const z = brent((x) => betaInc(x, al, be) - p, 0, 1, 1e-15, 1000);
  return A + z * (B - A);
}

function binomInv(a: Args): Value {
  const n = tr(num(sc(a, 0))), p = num(sc(a, 1)), al = num(sc(a, 2));
  if (n < 0 || p < 0 || p > 1 || al < 0 || al > 1) return ERR.num;
  let s = 0;
  for (let k = 0; k <= n; k++) {
    s += binomPmf(k, n, p);
    if (s >= al * (1 - 1e-15)) return k;
  }
  return n;
}

function negbinPmf(f: number, s: number, p: number): number {
  const c = combin(f + s - 1, s - 1);
  const v = c * Math.pow(p, s) * Math.pow(1 - p, f);
  if (Number.isFinite(v) && v > 1e-280) return v;
  return Math.exp(lnCombin(f + s - 1, s - 1) + s * Math.log(p) + f * Math.log1p(-p));
}

function hypgeomCheck(s: number, n: number, S2: number, N: number): CellError | null {
  if (s < 0 || s > n || s > S2 || n <= 0 || n > N || S2 <= 0 || S2 > N || N <= 0) return ERR.num;
  if (s < n - N + S2) return ERR.num;
  return null;
}

defFn(defs);

// --- Legacy / Compatibility names ---------------------------------------------------------

const byName = new Map(defs.map((d) => [d.name, d]));
function compat(name: string, base: string, desc?: string): FnDef {
  const b = byName.get(base)!;
  return { ...b, name, category: 'Compatibility', desc: desc ?? `${b.desc} (Older version of ${base}.)`, aliases: undefined };
}

defFn([
  compat('BETAINV', 'BETA.INV'),
  compat('BINOMDIST', 'BINOM.DIST'),
  compat('CHIDIST', 'CHISQ.DIST.RT'),
  compat('CHIINV', 'CHISQ.INV.RT'),
  compat('CHITEST', 'CHISQ.TEST'),
  compat('CONFIDENCE', 'CONFIDENCE.NORM'),
  compat('COVAR', 'COVARIANCE.P'),
  compat('CRITBINOM', 'BINOM.INV'),
  compat('EXPONDIST', 'EXPON.DIST'),
  compat('FDIST', 'F.DIST.RT'),
  compat('FINV', 'F.INV.RT'),
  compat('FTEST', 'F.TEST'),
  compat('GAMMADIST', 'GAMMA.DIST'),
  compat('GAMMAINV', 'GAMMA.INV'),
  compat('LOGINV', 'LOGNORM.INV'),
  compat('MODE', 'MODE.SNGL'),
  compat('NORMDIST', 'NORM.DIST'),
  compat('NORMINV', 'NORM.INV'),
  compat('NORMSINV', 'NORM.S.INV'),
  compat('PERCENTILE', 'PERCENTILE.INC'),
  compat('PERCENTRANK', 'PERCENTRANK.INC'),
  compat('POISSON', 'POISSON.DIST'),
  compat('QUARTILE', 'QUARTILE.INC'),
  compat('RANK', 'RANK.EQ'),
  compat('STDEV', 'STDEV.S'),
  compat('STDEVP', 'STDEV.P'),
  compat('TINV', 'T.INV.2T'),
  compat('TTEST', 'T.TEST'),
  compat('VAR', 'VAR.S'),
  compat('VARP', 'VAR.P'),
  compat('WEIBULL', 'WEIBULL.DIST'),
  compat('ZTEST', 'Z.TEST'),
  { ...byName.get('FORECAST.LINEAR')!, name: 'FORECAST', desc: 'Predicts a value along a linear trend fitted to known data.' },
  {
    name: 'BETADIST', min: 3, max: 5, category: 'Compatibility', desc: 'Cumulative beta distribution. (Older version of BETA.DIST.)', params: ['x', 'alpha', 'beta', '[A]', '[B]'],
    impl: (a) => {
      const x = num(sc(a, 0)), al = num(sc(a, 1)), be = num(sc(a, 2));
      const A = optN(a, 3, 0), B = optN(a, 4, 1);
      if (al <= 0 || be <= 0 || x < A || x > B || A === B) return ERR.num;
      return betaInc((x - A) / (B - A), al, be);
    },
  },
  {
    name: 'HYPGEOMDIST', min: 4, max: 4, category: 'Compatibility', desc: 'Hypergeometric probability. (Older version of HYPGEOM.DIST.)', params: ['sample_s', 'number_sample', 'population_s', 'number_pop'],
    impl: (a) => {
      const s = tr(num(sc(a, 0))), n = tr(num(sc(a, 1))), S2 = tr(num(sc(a, 2))), N = tr(num(sc(a, 3)));
      return hypgeomCheck(s, n, S2, N) ?? hypgeomPmf(s, n, S2, N);
    },
  },
  {
    name: 'LOGNORMDIST', min: 3, max: 3, category: 'Compatibility', desc: 'Cumulative lognormal distribution. (Older version of LOGNORM.DIST.)', params: ['x', 'mean', 'standard_dev'],
    impl: (a) => {
      const x = num(sc(a, 0)), m = num(sc(a, 1)), s = num(sc(a, 2));
      if (x <= 0 || s <= 0) return ERR.num;
      return normCdf((Math.log(x) - m) / s);
    },
  },
  {
    name: 'NEGBINOMDIST', min: 3, max: 3, category: 'Compatibility', desc: 'Negative binomial probability. (Older version of NEGBINOM.DIST.)', params: ['number_f', 'number_s', 'probability_s'],
    impl: (a) => {
      const f = tr(num(sc(a, 0))), s = tr(num(sc(a, 1))), p = num(sc(a, 2));
      if (f < 0 || s < 1 || p < 0 || p > 1) return ERR.num;
      return negbinPmf(f, s, p);
    },
  },
  {
    name: 'NORMSDIST', min: 1, max: 1, category: 'Compatibility', desc: 'Cumulative standard normal distribution. (Older version of NORM.S.DIST.)', params: ['z'],
    impl: (a) => normCdf(num(sc(a, 0))),
  },
  {
    name: 'TDIST', min: 3, max: 3, category: 'Compatibility', desc: "One- or two-tailed Student's t-distribution. (Older version of T.DIST.2T and T.DIST.RT.)", params: ['x', 'deg_freedom', 'tails'],
    impl: (a) => {
      const x = num(sc(a, 0)), df = tr(num(sc(a, 1))), tails = tr(num(sc(a, 2)));
      if (x < 0 || df < 1 || (tails !== 1 && tails !== 2)) return ERR.num;
      return tails * tSf(x, df);
    },
  },
]);

