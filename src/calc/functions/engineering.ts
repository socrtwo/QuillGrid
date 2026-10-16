/**
 * Engineering functions: Bessel functions, base conversion, bitwise operations, complex numbers,
 * error functions and CONVERT with the full unit table (SI and binary prefixes).
 *
 * Complex numbers are text in the form "x+yi" / "x+yj"; results keep the suffix of the inputs and
 * are formatted with up to 15 significant digits, omitting zero parts.
 */
import { defFn, type FnDef } from './registry';
import { num, str, fail, iterValues, numToText, ERR, CellError, type Value } from '../values';
import type { Scalar } from '../../model/types';

const S = (v: Value): Scalar => v as Scalar;
const has = (a: Value[], i: number): boolean => i < a.length && a[i] != null;

/** Number argument that rejects booleans given as text-like inputs the way engineering functions do. */
function numArg(v: Value): number {
  const x = S(v);
  if (typeof x === 'boolean') fail(ERR.value);
  return num(x);
}

// --------------------------------------------------------------------------------------------
// Bessel functions

/** Gauss-Legendre nodes/weights on [-1, 1]. */
const GL = (() => {
  const n = 20;
  const x: number[] = [], w: number[] = [];
  for (let i = 1; i <= n; i++) {
    let z = Math.cos((Math.PI * (i - 0.25)) / (n + 0.5));
    let pp = 0;
    for (let it = 0; it < 100; it++) {
      let p1 = 1, p2 = 0;
      for (let j = 1; j <= n; j++) { const p3 = p2; p2 = p1; p1 = ((2 * j - 1) * z * p2 - (j - 1) * p3) / j; }
      pp = (n * (z * p1 - p2)) / (z * z - 1);
      const z1 = z;
      z = z1 - p1 / pp;
      if (Math.abs(z - z1) < 1e-16) break;
    }
    x.push(z); w.push(2 / ((1 - z * z) * pp * pp));
  }
  return { x, w };
})();

/** Composite Gauss-Legendre integral of f over [a, b] with `panels` equal panels. */
function integrate(f: (t: number) => number, a: number, b: number, panels: number): number {
  const h = (b - a) / panels;
  let s = 0;
  for (let p = 0; p < panels; p++) {
    const c = a + (p + 0.5) * h;
    for (let i = 0; i < GL.x.length; i++) s += GL.w[i] * f(c + (GL.x[i] * h) / 2);
  }
  return (s * h) / 2;
}

function besselJ(x: number, n: number): number {
  if (Math.abs(x) <= 12 || n >= Math.abs(x)) {
    // Power series.
    const h = x / 2;
    let t = 1;
    for (let k = 1; k <= n; k++) t *= h / k;
    let s = t;
    for (let k = 0; k < 500; k++) {
      t *= -(h * h) / ((k + 1) * (n + k + 1));
      s += t;
      if (Math.abs(t) < 1e-17 * Math.abs(s)) break;
    }
    return s;
  }
  // Trapezoid rule on the periodic integral representation (spectrally accurate).
  const M = 2 * Math.ceil(Math.abs(x) + n) + 64;
  let s = 0;
  for (let k = 0; k < M; k++) { const t = (2 * Math.PI * k) / M; s += Math.cos(n * t - x * Math.sin(t)); }
  return s / M;
}

function besselI(x: number, n: number): number {
  const h = x / 2;
  let t = 1;
  for (let k = 1; k <= n; k++) t *= h / k;
  let s = t;
  for (let k = 0; k < 5000; k++) {
    t *= (h * h) / ((k + 1) * (n + k + 1));
    s += t;
    if (Math.abs(t) < 1e-17 * Math.abs(s)) break;
  }
  return s;
}

function besselK(x: number, n: number): number {
  // K_n(x) = ∫_0^∞ exp(-x cosh t) cosh(n t) dt, trapezoid rule (doubly exponential decay).
  const h = Math.min(0.1, 0.15 / Math.sqrt(x));
  const peak = Math.asinh(n / x);
  const g = (t: number) => (Math.exp(-x * Math.cosh(t) + n * t) + Math.exp(-x * Math.cosh(t) - n * t)) / 2;
  let s = g(0) / 2;
  for (let k = 1; k < 1e6; k++) {
    const t = k * h;
    const v = g(t);
    s += v;
    if (t > peak && v <= 1e-18 * s) break;
  }
  return s * h;
}

function besselY(x: number, n: number): number {
  // DLMF 10.9.7: Y_n(x) = 1/π ∫_0^π sin(x sinθ − nθ) dθ − 1/π ∫_0^∞ (e^{nt} + (−1)^n e^{−nt}) e^{−x sinh t} dt
  const p1 = Math.ceil(x + n) + 8;
  const i1 = integrate((th) => Math.sin(x * Math.sin(th) - n * th), 0, Math.PI, p1);
  let T = 1;
  while (x * Math.sinh(T) - n * T < 60 && T < 1000) T += 0.5;
  const sg = n % 2 ? -1 : 1;
  const i2 = integrate((t) => Math.exp(n * t - x * Math.sinh(t)) + sg * Math.exp(-n * t - x * Math.sinh(t)), 0, T, Math.ceil(T / 0.25));
  return (i1 - i2) / Math.PI;
}

function besselArgs(a: Value[]): [number, number] {
  const x = numArg(a[0]), n = Math.trunc(numArg(a[1]));
  if (n < 0) fail(ERR.num);
  return [x, n];
}

// --------------------------------------------------------------------------------------------
// Base conversion

interface Base { radix: number; bits: number; re: RegExp }
const BIN: Base = { radix: 2, bits: 10, re: /^[01]{0,10}$/ };
const OCT: Base = { radix: 8, bits: 30, re: /^[0-7]{0,10}$/ };
const HEX: Base = { radix: 16, bits: 40, re: /^[0-9a-f]{0,10}$/i };

function baseText(v: Value): string {
  const x = S(v);
  if (typeof x === 'boolean') fail(ERR.value);
  if (x instanceof CellError) fail(x);
  if (typeof x === 'number') {
    if (x < 0 || !Number.isInteger(x)) fail(ERR.num);
    return numToText(x);
  }
  return str(x);
}

function parseBase(v: Value, b: Base): number {
  const t = baseText(v);
  if (!b.re.test(t)) fail(ERR.num);
  if (t === '') return 0;
  let n = parseInt(t, b.radix);
  if (t.length === 10 && n >= Math.pow(2, b.bits - 1)) n -= Math.pow(2, b.bits);
  return n;
}

function toBase(n: number, b: Base, places: Value | undefined): string {
  const half = Math.pow(2, b.bits - 1);
  if (n < -half || n >= half) fail(ERR.num);
  if (n < 0) return (Math.pow(2, b.bits) + n).toString(b.radix).toUpperCase();
  const s = n.toString(b.radix).toUpperCase();
  if (places == null) return s;
  const p = Math.trunc(numArg(places));
  if (p <= 0 || p > 10 || s.length > p) fail(ERR.num);
  return s.padStart(p, '0');
}

const BASES: Record<string, Base> = { BIN, OCT, HEX };
const BASE_LABEL: Record<string, string> = { BIN: 'binary', OCT: 'octal', HEX: 'hexadecimal', DEC: 'decimal' };

const baseFns: FnDef[] = [];
for (const from of ['BIN', 'OCT', 'HEX', 'DEC']) {
  for (const to of ['BIN', 'OCT', 'HEX', 'DEC']) {
    if (from === to) continue;
    const withPlaces = to !== 'DEC';
    baseFns.push({
      name: `${from}2${to}`, min: 1, max: withPlaces ? 2 : 1, category: 'Engineering',
      desc: `Converts a ${BASE_LABEL[from]} number to ${BASE_LABEL[to]}.`,
      params: withPlaces ? ['number', '[places]'] : ['number'],
      impl: (a) => {
        let n: number;
        if (from === 'DEC') {
          n = Math.trunc(numArg(a[0]));
        } else n = parseBase(a[0], BASES[from]);
        if (to === 'DEC') return n;
        return toBase(n, BASES[to], has(a, 1) ? a[1] : undefined);
      },
    });
  }
}

// --------------------------------------------------------------------------------------------
// Bitwise (48-bit unsigned integers)

const BIT_MAX = Math.pow(2, 48);

function bitArg(v: Value): number {
  const n = numArg(v);
  if (n < 0 || n >= BIT_MAX || !Number.isInteger(n)) fail(ERR.num);
  return n;
}

function bitOp(a: number, b: number, op: (x: number, y: number) => number): number {
  const D = 0x1000000;
  const ah = Math.floor(a / D), al = a % D, bh = Math.floor(b / D), bl = b % D;
  return (op(ah, bh) >>> 0) * D + (op(al, bl) >>> 0);
}

function shift(n: number, s: number): number {
  if (Math.abs(s) > 53) fail(ERR.num);
  const r = s >= 0 ? n * Math.pow(2, s) : Math.floor(n / Math.pow(2, -s));
  if (r >= BIT_MAX) fail(ERR.num);
  return r;
}

// --------------------------------------------------------------------------------------------
// Complex numbers

interface Cx { re: number; im: number; sfx: '' | 'i' | 'j' }

const NUM = String.raw`(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?`;
const RE_REAL = new RegExp(`^[+-]?${NUM}$`);
const RE_IMAG = new RegExp(`^([+-]?)(${NUM})?([ij])$`);
const RE_BOTH = new RegExp(`^([+-]?${NUM})([+-])(${NUM})?([ij])$`);

function parseCx(v: Scalar): Cx {
  if (v instanceof CellError) return fail(v);
  if (typeof v === 'boolean') return fail(ERR.value);
  if (v == null) return { re: 0, im: 0, sfx: '' };
  if (typeof v === 'number') return { re: v, im: 0, sfx: '' };
  const t = v;
  if (t === '') return { re: 0, im: 0, sfx: '' };
  if (RE_REAL.test(t)) return { re: parseFloat(t), im: 0, sfx: '' };
  let m = RE_IMAG.exec(t);
  if (m) return { re: 0, im: (m[1] === '-' ? -1 : 1) * (m[2] ? parseFloat(m[2]) : 1), sfx: m[3] as 'i' | 'j' };
  m = RE_BOTH.exec(t);
  if (m) return { re: parseFloat(m[1]), im: (m[2] === '-' ? -1 : 1) * (m[3] ? parseFloat(m[3]) : 1), sfx: m[4] as 'i' | 'j' };
  return fail(ERR.num);
}

function cxArg(v: Value): Cx { return parseCx(S(v)); }

/** Shared suffix of several operands; mixing "i" and "j" is an error. */
function suffix(...xs: Cx[]): 'i' | 'j' {
  let s: '' | 'i' | 'j' = '';
  for (const x of xs) {
    if (!x.sfx) continue;
    if (s && s !== x.sfx) fail(ERR.value);
    s = x.sfx;
  }
  return s || 'i';
}

function fmtNum(x: number): string {
  if (x === 0) return '0';
  const r = +x.toPrecision(15);
  if (Math.abs(r) >= 1e15) {
    const [m, e] = r.toExponential(14).split('e');
    const mm = m.replace(/\.?0+$/, '');
    return `${mm}E${+e < 0 ? '-' : '+'}${String(Math.abs(+e)).padStart(2, '0')}`;
  }
  return numToText(r);
}

function fmtCx(re: number, im: number, sfx: 'i' | 'j'): string {
  if (!Number.isFinite(re) || !Number.isFinite(im)) fail(ERR.num);
  const r = +re.toPrecision(15), i = +im.toPrecision(15);
  if (i === 0) return fmtNum(r);
  const it = i === 1 ? '' : i === -1 ? '-' : fmtNum(i);
  if (r === 0) return `${it}${sfx}`;
  return `${fmtNum(r)}${i > 0 ? '+' : ''}${it}${sfx}`;
}

type C2 = [number, number];
const cmul = (a: C2, b: C2): C2 => [a[0] * b[0] - a[1] * b[1], a[0] * b[1] + a[1] * b[0]];
function cdiv(a: C2, b: C2): C2 {
  const d = b[0] * b[0] + b[1] * b[1];
  if (d === 0) fail(ERR.num);
  return [(a[0] * b[0] + a[1] * b[1]) / d, (a[1] * b[0] - a[0] * b[1]) / d];
}
const csin = ([a, b]: C2): C2 => [Math.sin(a) * Math.cosh(b), Math.cos(a) * Math.sinh(b)];
const ccos = ([a, b]: C2): C2 => [Math.cos(a) * Math.cosh(b), -Math.sin(a) * Math.sinh(b)];
const csinh = ([a, b]: C2): C2 => [Math.sinh(a) * Math.cos(b), Math.cosh(a) * Math.sin(b)];
const ccosh = ([a, b]: C2): C2 => [Math.cosh(a) * Math.cos(b), Math.sinh(a) * Math.sin(b)];
function cln([a, b]: C2): C2 {
  if (a === 0 && b === 0) fail(ERR.num);
  return [Math.log(Math.hypot(a, b)), Math.atan2(b, a)];
}

function cxUnary(name: string, desc: string, f: (z: C2) => C2): FnDef {
  return {
    name, min: 1, max: 1, category: 'Engineering', desc, params: ['inumber'],
    impl: (a) => { const z = cxArg(a[0]); const r = f([z.re, z.im]); return fmtCx(r[0], r[1], suffix(z)); },
  };
}

/** All complex operands from a possibly-range argument (blank cells skipped). */
function cxList(args: Value[], ctx: Parameters<NonNullable<FnDef['impl']>>[1]): Cx[] {
  const out: Cx[] = [];
  for (const a of args) for (const { v, fromRef } of iterValues(a, ctx)) { if (fromRef && (v == null || v === '')) continue; out.push(parseCx(v)); }
  return out;
}

// --------------------------------------------------------------------------------------------
// Error function

function erf(x: number): number {
  const ax = Math.abs(x);
  if (ax >= 2) return Math.sign(x) * (1 - erfcCF(ax));
  // erf(x) = 2/√π · e^{-x²} · Σ 2ⁿ x^{2n+1} / (1·3·…·(2n+1))  (all terms positive)
  let t = ax, s = ax;
  for (let n = 1; n < 200; n++) {
    t *= (2 * ax * ax) / (2 * n + 1);
    s += t;
    if (t < 1e-17 * s) break;
  }
  return (Math.sign(x) * 2 / Math.sqrt(Math.PI)) * Math.exp(-ax * ax) * s;
}

/** erfc for x >= 2 by continued fraction (modified Lentz). */
function erfcCF(x: number): number {
  if (x > 27) return 0;
  // erfc(x) = e^{-x²}/√π · 1/(x + (1/2)/(x + 1/(x + (3/2)/(x + 2/(x + …)))))
  const tiny = 1e-300;
  let f = x, C = x, D = 0;
  for (let k = 1; k < 500; k++) {
    const an = k / 2;
    D = x + an * D; if (D === 0) D = tiny; D = 1 / D;
    C = x + an / C; if (C === 0) C = tiny;
    const delta = C * D;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-16) break;
  }
  return Math.exp(-x * x) / Math.sqrt(Math.PI) / f;
}

function erfc(x: number): number {
  if (x >= 2) return erfcCF(x);
  if (x <= -2) return 2 - erfcCF(-x);
  return 1 - erf(x);
}

// --------------------------------------------------------------------------------------------
// CONVERT

type Cat = 'mass' | 'length' | 'time' | 'pressure' | 'force' | 'energy' | 'power' | 'magnetism' | 'temperature' | 'volume' | 'area' | 'information' | 'speed';
interface Unit { cat: Cat; f: number; prefix?: number; binary?: boolean }

const LBM = 0.45359237; // kg is not the base; masses are in grams
const IN = 0.0254, FT = 0.3048, YD = 0.9144, MI = 1609.344, NMI = 1852;
const LY = 9460730472580800, PC = 3.08567758149137e16, ANG = 1e-10;
const PICA_PT = IN / 72, PICA = IN / 6, SURVEY_MI = (5280 * 1200) / 3937;
const GRAV = 9.80665;
const LBF = LBM * GRAV;
const GAL = 231 * IN * IN * IN;

const UNITS: Record<string, Unit> = {};
function u(names: string[], cat: Cat, f: number, prefix?: number, binary?: boolean): void {
  for (const n of names) UNITS[n] = { cat, f, prefix, binary };
}
// Mass (grams)
u(['g'], 'mass', 1, 1);
u(['sg'], 'mass', (LBM * 1000 * GRAV) / FT);
u(['lbm'], 'mass', LBM * 1000);
u(['u'], 'mass', 1.660538782e-24, 1);
u(['ozm'], 'mass', (LBM * 1000) / 16);
u(['grain'], 'mass', (LBM * 1000) / 7000);
u(['cwt', 'shweight'], 'mass', LBM * 1000 * 100);
u(['uk_cwt', 'lcwt', 'hweight'], 'mass', LBM * 1000 * 112);
u(['stone'], 'mass', LBM * 1000 * 14);
u(['ton'], 'mass', LBM * 1000 * 2000);
u(['uk_ton', 'LTON', 'brton'], 'mass', LBM * 1000 * 2240);
// Length (meters)
u(['m'], 'length', 1, 1);
u(['mi'], 'length', MI);
u(['Nmi'], 'length', NMI);
u(['in'], 'length', IN);
u(['ft'], 'length', FT);
u(['yd'], 'length', YD);
u(['ang'], 'length', ANG, 1);
u(['ell'], 'length', 45 * IN);
u(['ly'], 'length', LY, 1);
u(['parsec', 'pc'], 'length', PC, 1);
u(['Picapt', 'Pica'], 'length', PICA_PT);
u(['pica'], 'length', PICA);
u(['survey_mi'], 'length', SURVEY_MI);
// Time (seconds)
u(['yr'], 'time', 365.25 * 86400);
u(['day', 'd'], 'time', 86400);
u(['hr'], 'time', 3600);
u(['mn', 'min'], 'time', 60);
u(['sec', 's'], 'time', 1, 1);
// Pressure (pascals)
u(['Pa', 'p'], 'pressure', 1, 1);
u(['atm', 'at'], 'pressure', 101325, 1);
u(['mmHg'], 'pressure', 101325 / 760, 1);
u(['psi'], 'pressure', LBF / (IN * IN));
u(['Torr'], 'pressure', 101325 / 760);
// Force (newtons)
u(['N'], 'force', 1, 1);
u(['dyn', 'dy'], 'force', 1e-5, 1);
u(['lbf'], 'force', LBF);
u(['pond'], 'force', GRAV / 1000, 1);
// Energy (joules)
u(['J'], 'energy', 1, 1);
u(['e'], 'energy', 1e-7, 1);
u(['c'], 'energy', 4.184, 1);
u(['cal'], 'energy', 4.1868, 1);
u(['eV', 'ev'], 'energy', 1.602176487e-19, 1);
u(['HPh', 'hh'], 'energy', ((LBF * FT * 550) * 3600));
u(['Wh', 'wh'], 'energy', 3600, 1);
u(['flb'], 'energy', LBF * FT);
u(['BTU', 'btu'], 'energy', 1055.05585262);
// Power (watts)
u(['HP', 'h'], 'power', LBF * FT * 550);
u(['PS'], 'power', 75 * GRAV);
u(['W', 'w'], 'power', 1, 1);
// Magnetism (tesla)
u(['T'], 'magnetism', 1, 1);
u(['ga'], 'magnetism', 1e-4, 1);
// Temperature (handled specially; factor unused)
u(['C', 'cel', 'F', 'fah', 'Rank', 'Reau'], 'temperature', 1);
u(['K', 'kel'], 'temperature', 1, 1);
// Volume (cubic meters)
u(['tsp'], 'volume', GAL / 768);
u(['tspm'], 'volume', 5e-6);
u(['tbs'], 'volume', GAL / 256);
u(['oz'], 'volume', GAL / 128);
u(['cup'], 'volume', GAL / 16);
u(['pt', 'us_pt'], 'volume', GAL / 8);
u(['uk_pt'], 'volume', 4.54609e-3 / 8);
u(['qt'], 'volume', GAL / 4);
u(['uk_qt'], 'volume', 4.54609e-3 / 4);
u(['gal'], 'volume', GAL);
u(['uk_gal'], 'volume', 4.54609e-3);
u(['l', 'L', 'lt'], 'volume', 1e-3, 1);
u(['ang3', 'ang^3'], 'volume', ANG ** 3, 3);
u(['barrel'], 'volume', 42 * GAL);
u(['bushel'], 'volume', 2150.42 * IN * IN * IN);
u(['ft3', 'ft^3'], 'volume', FT ** 3);
u(['in3', 'in^3'], 'volume', IN ** 3);
u(['ly3', 'ly^3'], 'volume', LY ** 3);
u(['m3', 'm^3'], 'volume', 1, 3);
u(['mi3', 'mi^3'], 'volume', MI ** 3);
u(['yd3', 'yd^3'], 'volume', YD ** 3);
u(['Nmi3', 'Nmi^3'], 'volume', NMI ** 3);
u(['Picapt3', 'Picapt^3', 'Pica3', 'Pica^3'], 'volume', PICA_PT ** 3);
u(['GRT', 'regton'], 'volume', 100 * FT ** 3);
u(['MTON'], 'volume', 40 * FT ** 3);
// Area (square meters)
u(['uk_acre'], 'area', 4840 * YD * YD);
u(['us_acre'], 'area', 43560 * (1200 / 3937) ** 2);
u(['ang2', 'ang^2'], 'area', ANG ** 2, 2);
u(['ar'], 'area', 100, 1);
u(['ft2', 'ft^2'], 'area', FT ** 2);
u(['ha'], 'area', 10000);
u(['in2', 'in^2'], 'area', IN ** 2);
u(['ly2', 'ly^2'], 'area', LY ** 2);
u(['m2', 'm^2'], 'area', 1, 2);
u(['Morgen'], 'area', 2500);
u(['mi2', 'mi^2'], 'area', MI ** 2);
u(['Nmi2', 'Nmi^2'], 'area', NMI ** 2);
u(['Picapt2', 'Pica2', 'Pica^2', 'Picapt^2'], 'area', PICA_PT ** 2);
u(['yd2', 'yd^2'], 'area', YD ** 2);
// Information (bits)
u(['bit'], 'information', 1, 1, true);
u(['byte'], 'information', 8, 1, true);
// Speed (meters per second)
u(['admkn'], 'speed', (6080 * FT) / 3600);
u(['kn'], 'speed', NMI / 3600);
u(['m/h', 'm/hr'], 'speed', 1 / 3600, 1);
u(['m/s', 'm/sec'], 'speed', 1, 1);
u(['mph'], 'speed', MI / 3600);

const PREFIX: Record<string, number> = {
  Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 1e1, e: 1e1,
  d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24,
};
const BIN_PREFIX: Record<string, number> = { Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10 };

interface Resolved { unit: Unit; name: string; scale: number }

function resolveUnit(s: string): Resolved | null {
  if (UNITS[s]) return { unit: UNITS[s], name: s, scale: 1 };
  for (const [p, v] of Object.entries(BIN_PREFIX)) {
    if (s.startsWith(p)) { const base = UNITS[s.slice(p.length)]; if (base?.binary) return { unit: base, name: s.slice(p.length), scale: v }; }
  }
  for (const p of ['da', ...Object.keys(PREFIX).filter((k) => k !== 'da')]) {
    if (s.length > p.length && s.startsWith(p)) {
      const name = s.slice(p.length), base = UNITS[name];
      if (base?.prefix) return { unit: base, name, scale: Math.pow(PREFIX[p], base.prefix) };
    }
  }
  return null;
}

function toKelvin(x: number, n: string): number {
  switch (n) {
    case 'C': case 'cel': return x + 273.15;
    case 'F': case 'fah': return ((x - 32) * 5) / 9 + 273.15;
    case 'Rank': return (x * 5) / 9;
    case 'Reau': return x * 1.25 + 273.15;
    default: return x;
  }
}
function fromKelvin(k: number, n: string): number {
  switch (n) {
    case 'C': case 'cel': return k - 273.15;
    case 'F': case 'fah': return ((k - 273.15) * 9) / 5 + 32;
    case 'Rank': return (k * 9) / 5;
    case 'Reau': return (k - 273.15) * 0.8;
    default: return k;
  }
}

function convert(x: number, from: string, to: string): number {
  const a = resolveUnit(from), b = resolveUnit(to);
  if (!a || !b || a.unit.cat !== b.unit.cat) return fail(ERR.na);
  if (a.unit.cat === 'temperature') return fromKelvin(toKelvin(x * a.scale, a.name), b.name) / b.scale;
  if (a.unit === b.unit && a.scale === b.scale) return x;
  return (x * a.unit.f * a.scale) / (b.unit.f * b.scale);
}

// --------------------------------------------------------------------------------------------
// Definitions

const E = 'Engineering' as const;
const def = (d: Omit<FnDef, 'category'>): FnDef => ({ ...d, category: E });

defFn([
  def({
    name: 'BESSELI', min: 2, max: 2, desc: 'Modified Bessel function of the first kind, In(x).', params: ['x', 'n'],
    impl: (a) => { const [x, n] = besselArgs(a); return besselI(x, n); },
  }),
  def({
    name: 'BESSELJ', min: 2, max: 2, desc: 'Bessel function of the first kind, Jn(x).', params: ['x', 'n'],
    impl: (a) => { const [x, n] = besselArgs(a); return besselJ(x, n); },
  }),
  def({
    name: 'BESSELK', min: 2, max: 2, desc: 'Modified Bessel function of the second kind, Kn(x).', params: ['x', 'n'],
    impl: (a) => { const [x, n] = besselArgs(a); if (x <= 0) fail(ERR.num); return besselK(x, n); },
  }),
  def({
    name: 'BESSELY', min: 2, max: 2, desc: 'Bessel function of the second kind (Weber/Neumann), Yn(x).', params: ['x', 'n'],
    impl: (a) => { const [x, n] = besselArgs(a); if (x <= 0) fail(ERR.num); return besselY(x, n); },
  }),
  ...baseFns,
  def({
    name: 'BITAND', min: 2, max: 2, desc: 'Bitwise AND of two non-negative integers.', params: ['number1', 'number2'],
    impl: (a) => bitOp(bitArg(a[0]), bitArg(a[1]), (x, y) => x & y),
  }),
  def({
    name: 'BITOR', min: 2, max: 2, desc: 'Bitwise OR of two non-negative integers.', params: ['number1', 'number2'],
    impl: (a) => bitOp(bitArg(a[0]), bitArg(a[1]), (x, y) => x | y),
  }),
  def({
    name: 'BITXOR', min: 2, max: 2, desc: 'Bitwise exclusive OR of two non-negative integers.', params: ['number1', 'number2'],
    impl: (a) => bitOp(bitArg(a[0]), bitArg(a[1]), (x, y) => x ^ y),
  }),
  def({
    name: 'BITLSHIFT', min: 2, max: 2, desc: 'Shifts the bits of a number left (a negative amount shifts right).', params: ['number', 'shift_amount'],
    impl: (a) => shift(bitArg(a[0]), Math.trunc(numArg(a[1]))),
  }),
  def({
    name: 'BITRSHIFT', min: 2, max: 2, desc: 'Shifts the bits of a number right (a negative amount shifts left).', params: ['number', 'shift_amount'],
    impl: (a) => shift(bitArg(a[0]), -Math.trunc(numArg(a[1]))),
  }),
  def({
    name: 'COMPLEX', min: 2, max: 3, desc: 'Builds a complex number text from real and imaginary coefficients.',
    params: ['real_num', 'i_num', '[suffix]'],
    impl: (a) => {
      const re = numArg(a[0]), im = numArg(a[1]);
      const s = has(a, 2) ? str(S(a[2])) : 'i';
      if (s !== 'i' && s !== 'j' && s !== '') fail(ERR.value);
      return fmtCx(re, im, (s || 'i') as 'i' | 'j');
    },
  }),
  def({
    name: 'CONVERT', min: 3, max: 3, desc: 'Converts a number from one measurement unit to another.',
    params: ['number', 'from_unit', 'to_unit'],
    impl: (a) => convert(numArg(a[0]), str(S(a[1])), str(S(a[2]))),
  }),
  def({
    name: 'DELTA', min: 1, max: 2, desc: 'Returns 1 if two numbers are equal, otherwise 0.', params: ['number1', '[number2]'],
    impl: (a) => (numArg(a[0]) === (has(a, 1) ? numArg(a[1]) : 0) ? 1 : 0),
  }),
  def({
    name: 'ERF', min: 1, max: 2, desc: 'Error function integrated between lower_limit and upper_limit (or from 0).',
    params: ['lower_limit', '[upper_limit]'],
    impl: (a) => {
      const lo = numArg(a[0]);
      if (!has(a, 1)) return erf(lo);
      const hi = numArg(a[1]);
      return erf(hi) - erf(lo);
    },
  }),
  def({
    name: 'ERF.PRECISE', min: 1, max: 1, desc: 'Error function integrated from 0 to x.', params: ['x'],
    impl: (a) => erf(numArg(a[0])),
  }),
  def({
    name: 'ERFC', min: 1, max: 1, desc: 'Complementary error function integrated from x to infinity.', params: ['x'],
    impl: (a) => erfc(numArg(a[0])),
  }),
  def({
    name: 'ERFC.PRECISE', min: 1, max: 1, desc: 'Complementary error function integrated from x to infinity.', params: ['x'],
    impl: (a) => erfc(numArg(a[0])),
  }),
  def({
    name: 'GESTEP', min: 1, max: 2, desc: 'Returns 1 if a number is greater than or equal to a threshold step, otherwise 0.', params: ['number', '[step]'],
    impl: (a) => (numArg(a[0]) >= (has(a, 1) ? numArg(a[1]) : 0) ? 1 : 0),
  }),
  def({
    name: 'IMABS', min: 1, max: 1, desc: 'Absolute value (modulus) of a complex number.', params: ['inumber'],
    impl: (a) => { const z = cxArg(a[0]); return Math.hypot(z.re, z.im); },
  }),
  def({
    name: 'IMAGINARY', min: 1, max: 1, desc: 'Imaginary coefficient of a complex number.', params: ['inumber'],
    impl: (a) => cxArg(a[0]).im,
  }),
  def({
    name: 'IMREAL', min: 1, max: 1, desc: 'Real coefficient of a complex number.', params: ['inumber'],
    impl: (a) => cxArg(a[0]).re,
  }),
  def({
    name: 'IMARGUMENT', min: 1, max: 1, desc: 'Argument (angle in radians) of a complex number.', params: ['inumber'],
    impl: (a) => { const z = cxArg(a[0]); if (z.re === 0 && z.im === 0) fail(ERR.div0); return Math.atan2(z.im, z.re); },
  }),
  cxUnary('IMCONJUGATE', 'Complex conjugate of a complex number.', ([a, b]) => [a, -b]),
  cxUnary('IMCOS', 'Cosine of a complex number.', ccos),
  cxUnary('IMCOSH', 'Hyperbolic cosine of a complex number.', ccosh),
  cxUnary('IMCOT', 'Cotangent of a complex number.', (z) => cdiv(ccos(z), csin(z))),
  cxUnary('IMCSC', 'Cosecant of a complex number.', (z) => cdiv([1, 0], csin(z))),
  cxUnary('IMCSCH', 'Hyperbolic cosecant of a complex number.', (z) => cdiv([1, 0], csinh(z))),
  cxUnary('IMEXP', 'Exponential of a complex number.', ([a, b]) => { const e = Math.exp(a); return [e * Math.cos(b), e * Math.sin(b)]; }),
  cxUnary('IMLN', 'Natural logarithm of a complex number.', cln),
  cxUnary('IMLOG10', 'Base-10 logarithm of a complex number.', (z) => { const l = cln(z); return [l[0] / Math.LN10, l[1] / Math.LN10]; }),
  cxUnary('IMLOG2', 'Base-2 logarithm of a complex number.', (z) => { const l = cln(z); return [l[0] / Math.LN2, l[1] / Math.LN2]; }),
  cxUnary('IMSEC', 'Secant of a complex number.', (z) => cdiv([1, 0], ccos(z))),
  cxUnary('IMSECH', 'Hyperbolic secant of a complex number.', (z) => cdiv([1, 0], ccosh(z))),
  cxUnary('IMSIN', 'Sine of a complex number.', csin),
  cxUnary('IMSINH', 'Hyperbolic sine of a complex number.', csinh),
  cxUnary('IMSQRT', 'Square root of a complex number.', ([a, b]) => {
    const r = Math.sqrt(Math.hypot(a, b)), t = Math.atan2(b, a) / 2;
    return [r * Math.cos(t), r * Math.sin(t)];
  }),
  cxUnary('IMTAN', 'Tangent of a complex number.', (z) => cdiv(csin(z), ccos(z))),
  def({
    name: 'IMDIV', min: 2, max: 2, desc: 'Quotient of two complex numbers.', params: ['inumber1', 'inumber2'],
    impl: (a) => { const x = cxArg(a[0]), y = cxArg(a[1]); const sx = suffix(x, y); const r = cdiv([x.re, x.im], [y.re, y.im]); return fmtCx(r[0], r[1], sx); },
  }),
  def({
    name: 'IMPOWER', min: 2, max: 2, desc: 'Complex number raised to a real power.', params: ['inumber', 'number'],
    impl: (a) => {
      const z = cxArg(a[0]), n = numArg(a[1]);
      const r = Math.hypot(z.re, z.im);
      if (r === 0) { if (n <= 0) fail(ERR.num); return '0'; }
      const t = Math.atan2(z.im, z.re), rn = Math.pow(r, n);
      return fmtCx(rn * Math.cos(n * t), rn * Math.sin(n * t), suffix(z));
    },
  }),
  def({
    name: 'IMSUB', min: 2, max: 2, desc: 'Difference of two complex numbers.', params: ['inumber1', 'inumber2'],
    impl: (a) => { const x = cxArg(a[0]), y = cxArg(a[1]); return fmtCx(x.re - y.re, x.im - y.im, suffix(x, y)); },
  }),
  def({
    name: 'IMSUM', min: 1, max: 255, arrays: 'all', desc: 'Sum of complex numbers.', params: ['inumber1', '[inumber2]', '...'],
    impl: (a, ctx) => {
      const xs = cxList(a, ctx);
      let re = 0, im = 0;
      for (const x of xs) { re += x.re; im += x.im; }
      return fmtCx(re, im, suffix(...xs));
    },
  }),
  def({
    name: 'IMPRODUCT', min: 1, max: 255, arrays: 'all', desc: 'Product of complex numbers.', params: ['inumber1', '[inumber2]', '...'],
    impl: (a, ctx) => {
      const xs = cxList(a, ctx);
      let p: C2 = [1, 0];
      for (const x of xs) p = cmul(p, [x.re, x.im]);
      return fmtCx(p[0], p[1], suffix(...xs));
    },
  }),
]);
