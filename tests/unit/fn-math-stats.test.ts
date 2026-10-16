import { describe, it, expect } from 'vitest';
import '../../src/calc/functions/logical';
import '../../src/calc/functions/math';
import '../../src/calc/functions/stats';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell } from '../../src/model/address';
import { getFn, FUNCTIONS } from '../../src/calc/functions/registry';
import { CellError } from '../../src/model/types';

type V = number | string | boolean | null | CellError;

function setup(data: Record<string, unknown> = {}) {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  const put = (a: string, v: unknown, sh = s) => {
    const p = parseCell(a)!;
    b.transact('t', () => sh.set(p.r, p.c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v: v as V }));
  };
  for (const [k, v] of Object.entries(data)) put(k, v);
  const val = (a: string, sh = s): V => { e.recalc(); const p = parseCell(a)!; return sh.value(p.r, p.c) as V; };
  let slot = 0;
  /** Evaluates a formula in a scratch cell far from the data. */
  const ev = (f: string): V => {
    const addr = `AZ${1000 + slot++ * 40}`;
    put(addr, '=' + f);
    return val(addr);
  };
  /** Evaluates a formula and reads its spilled result. */
  const arr = (f: string, rows: number, cols: number): V[][] => {
    const r0 = 1000 + slot++ * 40;
    put(`AZ${r0}`, '=' + f);
    e.recalc();
    const p = parseCell(`AZ${r0}`)!;
    const out: V[][] = [];
    for (let r = 0; r < rows; r++) { const row: V[] = []; for (let c = 0; c < cols; c++) row.push(s.value(p.r + r, p.c + c) as V); out.push(row); }
    return out;
  };
  return { b, e, s, put, val, ev, arr };
}

const code = (v: V) => (v instanceof CellError ? v.code : v);
function close(v: V, expected: number, digits = 8) {
  expect(typeof v, `got ${String(code(v))}`).toBe('number');
  const tol = Math.max(Math.abs(expected), 1e-300) * Math.pow(10, -digits);
  expect(Math.abs((v as number) - expected), `got ${v}, expected ${expected}`).toBeLessThanOrEqual(tol);
}

describe('registry metadata', () => {
  it('every math/stat function has desc and params', () => {
    const mine = [...FUNCTIONS.values()].filter((f) => ['Math & Trig', 'Statistical', 'Compatibility'].includes(f.category));
    expect(mine.length).toBeGreaterThan(200);
    for (const f of mine) {
      expect(f.desc.length, f.name).toBeGreaterThan(5);
      expect(Array.isArray(f.params), f.name).toBe(true);
      expect(f.params.length >= f.min || f.max > 200, f.name).toBe(true);
    }
  });
  it('has all required names', () => {
    const names = `ABS ACOS ACOSH ACOT ACOTH AGGREGATE ARABIC ASIN ASINH ATAN ATAN2 ATANH BASE CEILING CEILING.MATH CEILING.PRECISE ISO.CEILING COMBIN COMBINA COS COSH COT COTH CSC CSCH DECIMAL DEGREES EVEN EXP FACT FACTDOUBLE FLOOR FLOOR.MATH FLOOR.PRECISE GCD INT LCM LN LOG LOG10 MDETERM MINVERSE MMULT MOD MROUND MULTINOMIAL MUNIT ODD PI POWER PRODUCT QUOTIENT RADIANS RAND RANDARRAY RANDBETWEEN ROMAN ROUND ROUNDDOWN ROUNDUP SEC SECH SERIESSUM SIGN SIN SINH SQRT SQRTPI SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 TAN TANH TRUNC PERCENTOF
AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CONFIDENCE.NORM CONFIDENCE.T CORREL COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS COVARIANCE.P COVARIANCE.S DEVSQ EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FISHER FISHERINV FORECAST FORECAST.LINEAR FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FREQUENCY GAMMA GAMMA.DIST GAMMA.INV GAMMALN GAMMALN.PRECISE GAUSS GEOMEAN GROWTH HARMEAN HYPGEOM.DIST INTERCEPT KURT LARGE LINEST LOGEST LOGNORM.DIST LOGNORM.INV MAX MAXA MAXIFS MEDIAN MIN MINA MINIFS MODE.MULT MODE.SNGL NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV PEARSON PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC PERMUT PERMUTATIONA PHI POISSON.DIST PROB QUARTILE.EXC QUARTILE.INC RANK.AVG RANK.EQ RSQ SKEW SKEW.P SLOPE SMALL STANDARDIZE STDEV.P STDEV.S STDEVA STDEVPA STEYX T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TREND TRIMMEAN VAR.P VAR.S VARA VARPA WEIBULL.DIST Z.TEST
BETADIST BETAINV BINOMDIST CHIDIST CHIINV CHITEST CONFIDENCE COVAR CRITBINOM EXPONDIST FDIST FINV FTEST GAMMADIST GAMMAINV HYPGEOMDIST LOGINV LOGNORMDIST MODE NEGBINOMDIST NORMDIST NORMINV NORMSDIST NORMSINV PERCENTILE PERCENTRANK POISSON QUARTILE RANK STDEV STDEVP TDIST TINV TTEST VAR VARP WEIBULL ZTEST`.split(/\s+/);
    const missing = names.filter((n) => !getFn(n));
    expect(missing).toEqual([]);
    expect(getFn('RAND')!.volatile).toBe(true);
    expect(getFn('STDEV')!.category).toBe('Compatibility');
  });
});

describe('Math & Trig: scalar functions', () => {
  it('basic arithmetic and logs', () => {
    const t = setup();
    expect(t.ev('ABS(-3.5)')).toBe(3.5);
    expect(t.ev('SIGN(-2)')).toBe(-1);
    close(t.ev('PI()'), Math.PI);
    expect(t.ev('SQRT(16)')).toBe(4);
    expect(code(t.ev('SQRT(-1)'))).toBe('#NUM!');
    close(t.ev('SQRTPI(2)'), 2.50662827463, 11);
    close(t.ev('EXP(1)'), Math.E);
    expect(code(t.ev('LN(0)'))).toBe('#NUM!');
    expect(t.ev('LOG(8,2)')).toBe(3);
    expect(t.ev('LOG(100)')).toBe(2);
    expect(code(t.ev('LOG(10,1)'))).toBe('#DIV/0!');
    expect(t.ev('LOG10(1000)')).toBe(3);
    expect(t.ev('POWER(2,10)')).toBe(1024);
    expect(code(t.ev('POWER(0,0)'))).toBe('#NUM!');
    expect(code(t.ev('POWER(-8,1/3)'))).toBe('#NUM!');
    expect(t.ev('DEGREES(PI())')).toBe(180);
    close(t.ev('RADIANS(270)'), 4.71238898, 8);
    expect(code(t.ev('ABS("abc")'))).toBe('#VALUE!');
    expect(t.ev('ABS("-4")')).toBe(4);
    expect(t.ev('ABS(TRUE)')).toBe(1);
  });
  it('trigonometry', () => {
    const t = setup();
    close(t.ev('SIN(PI()/6)'), 0.5);
    close(t.ev('COS(PI()/3)'), 0.5);
    close(t.ev('TAN(PI()/4)'), 1);
    expect(code(t.ev('SIN(2^27)'))).toBe('#NUM!');
    close(t.ev('ACOT(2)'), 0.4636476090008061);
    close(t.ev('ACOTH(6)'), 0.168236118);
    expect(code(t.ev('ACOTH(0.5)'))).toBe('#NUM!');
    expect(code(t.ev('ACOS(2)'))).toBe('#NUM!');
    close(t.ev('ACOSH(10)'), 2.993222846);
    expect(code(t.ev('ATANH(1)'))).toBe('#NUM!');
    close(t.ev('ATAN2(1,1)'), Math.PI / 4);
    expect(code(t.ev('ATAN2(0,0)'))).toBe('#DIV/0!');
    expect(code(t.ev('COT(0)'))).toBe('#DIV/0!');
    close(t.ev('COT(30)'), -0.156119952);
    close(t.ev('COTH(2)'), 1.037314721);
    close(t.ev('CSC(15)'), 1.537780562);
    close(t.ev('CSCH(1.5)'), 0.469642441);
    close(t.ev('SEC(45)'), 1.903594407);
    close(t.ev('SECH(45)'), 5.73e-20, 2);
    close(t.ev('SINH(1)'), 1.175201194);
    close(t.ev('ASINH(-2.5)'), -1.647231146);
  });
  it('rounding', () => {
    const t = setup();
    expect(t.ev('ROUND(2.15,1)')).toBe(2.2);
    expect(t.ev('ROUND(2.149,1)')).toBe(2.1);
    expect(t.ev('ROUND(-1.475,2)')).toBe(-1.48);
    expect(t.ev('ROUND(21.5,-1)')).toBe(20);
    expect(t.ev('ROUND(626.3,-3)')).toBe(1000);
    expect(t.ev('ROUND(1.005,2)')).toBe(1.01);
    expect(t.ev('ROUNDDOWN(3.14159,3)')).toBe(3.141);
    expect(t.ev('ROUNDDOWN(-3.14159,1)')).toBe(-3.1);
    expect(t.ev('ROUNDDOWN(31415.92654,-2)')).toBe(31400);
    expect(t.ev('ROUNDUP(3.2,0)')).toBe(4);
    expect(t.ev('ROUNDUP(-3.14159,1)')).toBe(-3.2);
    expect(t.ev('ROUNDUP(0.1*3,1)')).toBe(0.3);
    expect(t.ev('TRUNC(8.9)')).toBe(8);
    expect(t.ev('TRUNC(-8.9)')).toBe(-8);
    expect(t.ev('INT(-8.9)')).toBe(-9);
    expect(t.ev('EVEN(1.5)')).toBe(2);
    expect(t.ev('EVEN(-1)')).toBe(-2);
    expect(t.ev('ODD(2)')).toBe(3);
    expect(t.ev('ODD(-1)')).toBe(-1);
    expect(t.ev('ODD(0)')).toBe(1);
    expect(t.ev('MROUND(10,3)')).toBe(9);
    expect(t.ev('MROUND(-10,-3)')).toBe(-9);
    expect(t.ev('MROUND(1.3,0.2)')).toBe(1.4);
    expect(code(t.ev('MROUND(5,-2)'))).toBe('#NUM!');
  });
  it('CEILING / FLOOR family', () => {
    const t = setup();
    expect(t.ev('CEILING(2.5,1)')).toBe(3);
    expect(t.ev('CEILING(-2.5,-2)')).toBe(-4);
    expect(t.ev('CEILING(-2.5,2)')).toBe(-2);
    expect(code(t.ev('CEILING(1.5,-1)'))).toBe('#NUM!');
    expect(t.ev('CEILING(0.234,0.01)')).toBe(0.24);
    expect(t.ev('FLOOR(3.7,2)')).toBe(2);
    expect(t.ev('FLOOR(-2.5,-2)')).toBe(-2);
    expect(t.ev('FLOOR(-2.5,2)')).toBe(-4);
    expect(code(t.ev('FLOOR(3.7,-2)'))).toBe('#NUM!');
    expect(t.ev('FLOOR(1.58,0.1)')).toBe(1.5);
    expect(t.ev('FLOOR(0.3,0.1)')).toBe(0.3);
    expect(code(t.ev('FLOOR(1,0)'))).toBe('#DIV/0!');
    expect(t.ev('CEILING.MATH(24.3,5)')).toBe(25);
    expect(t.ev('CEILING.MATH(6.7)')).toBe(7);
    expect(t.ev('CEILING.MATH(-8.1,2)')).toBe(-8);
    expect(t.ev('CEILING.MATH(-5.5,2,-1)')).toBe(-6);
    expect(t.ev('FLOOR.MATH(24.3,5)')).toBe(20);
    expect(t.ev('FLOOR.MATH(-8.1,2)')).toBe(-10);
    expect(t.ev('FLOOR.MATH(-5.5,2,-1)')).toBe(-4);
    expect(t.ev('CEILING.PRECISE(-4.6,-2)')).toBe(-4);
    expect(t.ev('ISO.CEILING(4.3,-2)')).toBe(6);
    expect(t.ev('FLOOR.PRECISE(-3.2,-1)')).toBe(-4);
    expect(t.ev('FLOOR.PRECISE(3.2)')).toBe(3);
  });
  it('integer functions and combinatorics', () => {
    const t = setup();
    expect(t.ev('MOD(3,2)')).toBe(1);
    expect(t.ev('MOD(-3,2)')).toBe(1);
    expect(t.ev('MOD(3,-2)')).toBe(-1);
    expect(t.ev('MOD(-3,-2)')).toBe(-1);
    expect(code(t.ev('MOD(1,0)'))).toBe('#DIV/0!');
    expect(t.ev('QUOTIENT(-10,3)')).toBe(-3);
    expect(t.ev('GCD(24,36)')).toBe(12);
    expect(t.ev('GCD(5,0)')).toBe(5);
    expect(code(t.ev('GCD(-1,2)'))).toBe('#NUM!');
    expect(t.ev('LCM(24,36)')).toBe(72);
    expect(t.ev('LCM(1.5,4)')).toBe(4);
    expect(t.ev('FACT(5)')).toBe(120);
    expect(t.ev('FACT(1.9)')).toBe(1);
    expect(code(t.ev('FACT(-1)'))).toBe('#NUM!');
    expect(t.ev('FACTDOUBLE(6)')).toBe(48);
    expect(t.ev('FACTDOUBLE(7)')).toBe(105);
    expect(t.ev('FACTDOUBLE(-1)')).toBe(1);
    expect(t.ev('COMBIN(8,2)')).toBe(28);
    expect(code(t.ev('COMBIN(2,3)'))).toBe('#NUM!');
    expect(t.ev('COMBINA(4,3)')).toBe(20);
    expect(t.ev('COMBINA(10,3)')).toBe(220);
    expect(t.ev('MULTINOMIAL(2,3,4)')).toBe(1260);
    expect(t.ev('PERMUT(100,3)')).toBe(970200);
    expect(t.ev('PERMUTATIONA(3,2)')).toBe(9);
  });
  it('number systems', () => {
    const t = setup();
    expect(t.ev('ROMAN(499,0)')).toBe('CDXCIX');
    expect(t.ev('ROMAN(499,1)')).toBe('LDVLIV');
    expect(t.ev('ROMAN(499,2)')).toBe('XDIX');
    expect(t.ev('ROMAN(499,3)')).toBe('VDIV');
    expect(t.ev('ROMAN(499,4)')).toBe('ID');
    expect(t.ev('ROMAN(2013)')).toBe('MMXIII');
    expect(t.ev('ROMAN(1999,TRUE)')).toBe('MCMXCIX');
    expect(code(t.ev('ROMAN(4000)'))).toBe('#VALUE!');
    expect(t.ev('ARABIC("LVII")')).toBe(57);
    expect(t.ev('ARABIC("mcmxii")')).toBe(1912);
    expect(t.ev('ARABIC("-MMXI")')).toBe(-2011);
    expect(t.ev('ARABIC("")')).toBe(0);
    expect(code(t.ev('ARABIC("ABC")'))).toBe('#VALUE!');
    expect(t.ev('BASE(7,2)')).toBe('111');
    expect(t.ev('BASE(100,16)')).toBe('64');
    expect(t.ev('BASE(15,2,10)')).toBe('0000001111');
    expect(code(t.ev('BASE(10,1)'))).toBe('#NUM!');
    expect(t.ev('DECIMAL("FF",16)')).toBe(255);
    expect(t.ev('DECIMAL("111",2)')).toBe(7);
    expect(t.ev('DECIMAL("zap",36)')).toBe(45745);
    expect(code(t.ev('DECIMAL("12",2)'))).toBe('#NUM!');
  });
  it('random functions are in range', () => {
    const t = setup();
    const r = t.ev('RAND()') as number;
    expect(r).toBeGreaterThanOrEqual(0); expect(r).toBeLessThan(1);
    const rb = t.ev('RANDBETWEEN(5,7)') as number;
    expect([5, 6, 7]).toContain(rb);
    expect(code(t.ev('RANDBETWEEN(7,5)'))).toBe('#NUM!');
    const m = t.arr('RANDARRAY(3,2,10,20,TRUE)', 3, 2).flat() as number[];
    for (const x of m) { expect(Number.isInteger(x)).toBe(true); expect(x).toBeGreaterThanOrEqual(10); expect(x).toBeLessThanOrEqual(20); }
    expect(code(t.ev('RANDARRAY(2,2,5,1)'))).toBe('#VALUE!');
  });
  it('element-wise lifting of scalar functions', () => {
    const t = setup({ A1: 1.25, A2: -2.5, A3: 3.75 });
    expect(t.arr('ROUND(A1:A3,0)', 3, 1)).toEqual([[1], [-3], [4]]);
    expect(t.arr('ABS({-1,2,-3})', 1, 3)).toEqual([[1, 2, 3]]);
    expect(t.ev('SUM(ABS(A1:A3))')).toBe(7.5);
  });
});

describe('Math & Trig: aggregates', () => {
  it('SUM / PRODUCT / SUMSQ coercion rules', () => {
    const t = setup({ A1: 1, A2: '2', A3: true, A4: 4, A5: null });
    expect(t.ev('SUM(A1:A5)')).toBe(5);              // text and booleans in refs ignored
    expect(t.ev('SUM(1,"2",TRUE)')).toBe(4);          // direct args coerced
    expect(code(t.ev('SUM(1,"x")'))).toBe('#VALUE!');
    expect(t.ev('SUM({1,"2",TRUE})')).toBe(1);        // arrays behave like refs
    expect(t.ev('PRODUCT(A1:A5,2)')).toBe(8);
    expect(t.ev('PRODUCT(A5)')).toBe(0);
    expect(t.ev('SUMSQ(3,4)')).toBe(25);
    t.put('B1', '=1/0');
    expect(code(t.ev('SUM(A1,B1)'))).toBe('#DIV/0!');
  });
  it('SUM over 3D references', () => {
    const t = setup({ A1: 1 });
    const s2 = t.b.sheet.constructor as any;
    const sheet2 = new s2(t.b, 'Sheet2');
    const sheet3 = new s2(t.b, 'Sheet3');
    t.b.sheets.push(sheet2, sheet3);
    t.put('A1', 10, sheet2);
    t.put('A1', 100, sheet3);
    expect(t.ev('SUM(Sheet1:Sheet3!A1)')).toBe(111);
  });
  it('SUMIF / SUMIFS with criteria and wildcards', () => {
    const t = setup({
      A1: 'apple', B1: 10, C1: 1,
      A2: 'banana', B2: 20, C2: 2,
      A3: 'apricot', B3: 30, C3: 3,
      A4: 'cherry', B4: 40, C4: 4,
      A5: 'a*b', B5: 50, C5: 5,
      A6: null, B6: 60, C6: 6,
    });
    expect(t.ev('SUMIF(B1:B6,">25")')).toBe(180);
    expect(t.ev('SUMIF(A1:A6,"ap*",B1:B6)')).toBe(40);
    expect(t.ev('SUMIF(A1:A6,"?????",B1:B6)')).toBe(10);
    expect(t.ev('SUMIF(A1:A6,"a~*b",B1:B6)')).toBe(50);
    expect(t.ev('SUMIF(A1:A6,"<>apple",B1:B6)')).toBe(200);
    expect(t.ev('SUMIF(A1:A6,"",B1:B6)')).toBe(60);
    expect(t.ev('SUMIF(A1:A6,"APPLE",B1)')).toBe(10);   // sum_range resized from its top-left
    expect(t.ev('SUMIF(C1:C6,3,B1:B6)')).toBe(30);
    expect(t.ev('SUMIF(C1:C6,"<=2",B1:B6)')).toBe(30);
    expect(t.ev('SUMIFS(B1:B6,A1:A6,"a*",C1:C6,">1")')).toBe(80);
    expect(t.ev('SUMIFS(B1:B6,C1:C6,">=2",C1:C6,"<5")')).toBe(90);
    expect(code(t.ev('SUMIFS(B1:B6,A1:A5,"a*")'))).toBe('#VALUE!');
    expect(t.arr('SUMIF(A1:A6,{"apple","cherry"},B1:B6)', 1, 2)).toEqual([[10, 40]]);
  });
  it('SUMPRODUCT and pairwise sums', () => {
    const t = setup({ A1: 3, A2: 8, A3: 1, B1: 4, B2: 6, B3: 9, C1: 'x' });
    expect(t.ev('SUMPRODUCT(A1:A3,B1:B3)')).toBe(69);
    expect(t.ev('SUMPRODUCT((A1:A3>2)*1,B1:B3)')).toBe(10);
    expect(code(t.ev('SUMPRODUCT(A1:A3,B1:B2)'))).toBe('#VALUE!');
    expect(t.ev('SUMPRODUCT(A1:A3,C1:C3)')).toBe(0);
    expect(t.ev('SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})')).toBe(-55);
    expect(t.ev('SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})')).toBe(521);
    expect(t.ev('SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})')).toBe(79);
    expect(code(t.ev('SUMXMY2({1,2},{1,2,3})'))).toBe('#N/A');
    close(t.ev('SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666667,-0.001388889})'), 0.707103, 6);
    close(t.ev('PERCENTOF(A1:A2,A1:A3)'), 11 / 12, 13);
    expect(code(t.ev('PERCENTOF(1,0)'))).toBe('#DIV/0!');
  });
  it('matrix functions', () => {
    const t = setup();
    expect(t.ev('MDETERM({3,6,1;1,1,0;3,10,2})')).toBe(1);
    expect(t.ev('MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})')).toBe(88);
    expect(code(t.ev('MDETERM({1,2,3;4,5,6})'))).toBe('#VALUE!');
    expect(t.arr('MINVERSE({4,-1;2,0})', 2, 2)).toEqual([[0, 0.5], [-1, 2]]);
    expect(code(t.ev('MINVERSE({1,2;2,4})'))).toBe('#NUM!');
    expect(t.arr('MMULT({1,3;7,2},{2,0;0,2})', 2, 2)).toEqual([[2, 6], [14, 4]]);
    expect(code(t.ev('MMULT({1,2},{1,2})'))).toBe('#VALUE!');
    expect(t.arr('MUNIT(3)', 3, 3)).toEqual([[1, 0, 0], [0, 1, 0], [0, 0, 1]]);
  });
  it('SUBTOTAL skips nested subtotals and hidden rows', () => {
    const t = setup({ A1: 10, A2: 20, A3: '=SUBTOTAL(9,A1:A2)', A4: 5, A5: 7, A6: '=SUBTOTAL(9,A4:A5)', A7: 'text' });
    expect(t.val('A3')).toBe(30);
    expect(t.ev('SUBTOTAL(9,A1:A7)')).toBe(42);
    expect(t.ev('SUM(A1:A7)')).toBe(84);
    expect(t.ev('SUBTOTAL(1,A1:A7)')).toBe(10.5);
    expect(t.ev('SUBTOTAL(2,A1:A7)')).toBe(4);
    expect(t.ev('SUBTOTAL(3,A1:A7)')).toBe(5);
    expect(t.ev('SUBTOTAL(4,A1:A7)')).toBe(20);
    expect(t.ev('SUBTOTAL(5,A1:A7)')).toBe(5);
    expect(t.ev('SUBTOTAL(6,A1:A7)')).toBe(7000);
    close(t.ev('SUBTOTAL(7,A1:A7)'), 6.658328118);
    // Manually hidden row: 9 includes, 109 excludes
    t.s.hiddenRows.add(1); // row 2 (A2 = 20)
    t.b.transact('touch', () => t.s.set(9, 9, { v: 1 }));
    expect(t.ev('SUBTOTAL(9,A1:A2,A4:A5)')).toBe(42);
    expect(t.ev('SUBTOTAL(109,A1:A2,A4:A5)')).toBe(22);
    expect(t.ev('SUBTOTAL(102,A1:A7)')).toBe(3);
    expect(code(t.ev('SUBTOTAL(12,A1:A2)'))).toBe('#VALUE!');
  });
  it('SUBTOTAL excludes rows hidden by an AutoFilter even for 1-11', () => {
    const t = setup({ A1: 'h', A2: 1, A3: 2, A4: 4 });
    t.s.filter = { range: { r1: 0, c1: 0, r2: 3, c2: 0 }, columns: {} };
    t.s.hiddenRows.add(2);
    expect(t.ev('SUBTOTAL(9,A2:A4)')).toBe(5);
    expect(t.ev('SUBTOTAL(109,A2:A4)')).toBe(5);
  });
  it('AGGREGATE options', () => {
    const t = setup({ A1: 4, A2: '=1/0', A3: 8, A4: '=SUBTOTAL(9,A1)', A5: 2, A6: 6 });
    expect(code(t.ev('SUM(A1:A6)'))).toBe('#DIV/0!');
    expect(t.ev('AGGREGATE(9,6,A1:A6)')).toBe(24);          // errors ignored, nested kept
    expect(t.ev('AGGREGATE(9,3,A1:A6)')).toBe(20);          // errors + nested ignored
    expect(code(t.ev('AGGREGATE(9,4,A1:A6)'))).toBe('#DIV/0!');
    expect(code(t.ev('AGGREGATE(9,0,A1:A6)'))).toBe('#DIV/0!');
    expect(t.ev('AGGREGATE(4,6,A1:A6)')).toBe(8);
    expect(t.ev('AGGREGATE(14,6,A1:A6,2)')).toBe(6);        // LARGE
    expect(t.ev('AGGREGATE(15,7,A1:A6,1)')).toBe(2);        // SMALL
    expect(t.ev('AGGREGATE(12,6,A1:A6)')).toBe(4);          // MEDIAN of 4,8,4,2,6
    expect(t.ev('AGGREGATE(16,6,A1:A6,0.5)')).toBe(4);
    expect(t.ev('AGGREGATE(17,6,A1:A6,4)')).toBe(8);
    expect(t.ev('AGGREGATE(14,6,A1:A6/(A1:A6>3),3)')).toBe(4); // array form
    expect(t.arr('AGGREGATE(14,6,A1:A6,{1,2})', 1, 2)).toEqual([[8, 6]]);
    t.s.hiddenRows.add(2); // A3 = 8
    t.b.transact('touch', () => t.s.set(9, 9, { v: 1 }));
    expect(t.ev('AGGREGATE(9,3,A1:A6)')).toBe(12);          // hidden, errors, nested ignored
    expect(t.ev('AGGREGATE(9,7,A1:A6)')).toBe(16);          // options 4-7 keep nested subtotals
    expect(t.ev('AGGREGATE(9,6,A1:A6)')).toBe(24);
    expect(code(t.ev('AGGREGATE(20,0,A1:A6)'))).toBe('#VALUE!');
    expect(code(t.ev('AGGREGATE(14,6,A1:A6)'))).toBe('#VALUE!');
  });
});

describe('Statistical: descriptive', () => {
  const D = '{1345,1301,1368,1322,1310,1370,1318,1350,1303,1299}';
  it('counting', () => {
    const t = setup({ A1: 1, A2: 'x', A3: true, A4: null, A5: '', A6: '=1/0', A7: 3 });
    expect(t.ev('COUNT(A1:A7)')).toBe(2);
    expect(t.ev('COUNT(1,"2",TRUE,"x")')).toBe(3);
    expect(t.ev('COUNTA(A1:A7)')).toBe(6);
    expect(t.ev('COUNTBLANK(A1:A7)')).toBe(2);
    expect(t.ev('COUNTBLANK(A1:A20)')).toBe(15);
  });
  it('COUNTIF / COUNTIFS / AVERAGEIF(S) / MAXIFS / MINIFS', () => {
    const t = setup({
      A1: 'east', B1: 10, A2: 'west', B2: 25, A3: 'East', B3: 35, A4: 'north', B4: 'n/a', A5: 'eastern', B5: 50,
    });
    expect(t.ev('COUNTIF(A1:A5,"east")')).toBe(2);
    expect(t.ev('COUNTIF(A1:A5,"east*")')).toBe(3);
    expect(t.ev('COUNTIF(B1:B5,">20")')).toBe(3);
    expect(t.ev('COUNTIF(B1:B5,"<>10")')).toBe(4);
    expect(t.ev('COUNTIF(A1:A10,"")')).toBe(5);
    expect(t.ev('COUNTIFS(A1:A5,"east*",B1:B5,">=35")')).toBe(2);
    expect(t.ev('AVERAGEIF(A1:A5,"east",B1:B5)')).toBe(22.5);
    expect(t.ev('AVERAGEIF(B1:B5,">20")')).toBeCloseTo(36.666666667, 8);
    expect(code(t.ev('AVERAGEIF(A1:A5,"south",B1:B5)'))).toBe('#DIV/0!');
    expect(t.ev('AVERAGEIFS(B1:B5,A1:A5,"*e*",B1:B5,">10")')).toBeCloseTo(36.666666667, 8);
    expect(t.ev('MAXIFS(B1:B5,A1:A5,"east*")')).toBe(50);
    expect(t.ev('MINIFS(B1:B5,A1:A5,"east*")')).toBe(10);
    expect(t.ev('MAXIFS(B1:B5,A1:A5,"zzz")')).toBe(0);
  });
  it('averages, extremes and spread', () => {
    const t = setup({ A1: 10, A2: 7, A3: 9, A4: 27, A5: 2, A6: 'x', A7: true });
    expect(t.ev('AVERAGE(A1:A7)')).toBe(11);
    expect(t.ev('AVERAGE(A1:A5,5)')).toBe(10);
    expect(code(t.ev('AVERAGE(A6)'))).toBe('#DIV/0!');
    close(t.ev('AVERAGEA(A1:A7)'), 56 / 7);
    expect(t.ev('MAX(A1:A7)')).toBe(27);
    expect(t.ev('MAX(A6)')).toBe(0);
    expect(t.ev('MIN(A1:A7)')).toBe(2);
    expect(t.ev('MAXA(A6:A7)')).toBe(1);
    expect(t.ev('MINA(A1:A7)')).toBe(0);
    expect(t.ev('MEDIAN(1,2,3,4,5,6)')).toBe(3.5);
    expect(t.ev('MEDIAN(A1:A5)')).toBe(9);
    close(t.ev(`STDEV.S(${D})`), 27.46391572);
    close(t.ev(`STDEV.P(${D})`), 26.05455814);
    close(t.ev(`VAR.S(${D})`), 754.2666667);
    close(t.ev(`VAR.P(${D})`), 678.84);
    close(t.ev(`STDEV(${D})`), 27.46391572);
    close(t.ev(`VARP(${D})`), 678.84);
    expect(code(t.ev('STDEV.S(1)'))).toBe('#DIV/0!');
    close(t.ev('STDEVA(A1:A7)'), Math.sqrt(t.ev('VARA(A1:A7)') as number));
    close(t.ev('VARPA(1,2,3)'), 2 / 3);
    close(t.ev('AVEDEV(4,5,6,7,5,4,3)'), 1.020408163);
    expect(t.ev('DEVSQ(4,5,8,7,11,4,3)')).toBe(48);
    close(t.ev('GEOMEAN(4,5,8,7,11,4,3)'), 5.476986969);
    close(t.ev('HARMEAN(4,5,8,7,11,4,3)'), 5.028375962);
    expect(code(t.ev('GEOMEAN(1,-1)'))).toBe('#NUM!');
    close(t.ev('KURT(3,4,5,2,3,4,5,6,4,7)'), -0.151799637);
    close(t.ev('SKEW(3,4,5,2,3,4,5,6,4,7)'), 0.359543071);
    close(t.ev('SKEW.P(3,4,5,2,3,4,5,6,4,7)'), 0.303193339);
    expect(code(t.ev('KURT(1,2,3)'))).toBe('#DIV/0!');
    close(t.ev('TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)'), 3.777777778);
  });
  it('mode, rank, k-th and percentiles', () => {
    const t = setup();
    expect(t.ev('MODE.SNGL(5.6,4,4,3,2,4)')).toBe(4);
    expect(code(t.ev('MODE(1,2,3)'))).toBe('#N/A');
    expect(t.arr('MODE.MULT({1,2,3,4,3,2,1,2,3})', 2, 1)).toEqual([[2], [3]]);
    expect(t.ev('LARGE({3,5,3,5,4},3)')).toBe(4);
    expect(t.ev('SMALL({3,4,5,2,3,4,6,4,7},4)')).toBe(4);
    expect(code(t.ev('SMALL({1,2},3)'))).toBe('#NUM!');
    expect(t.arr('LARGE({1,9,5},{1,2,3})', 1, 3)).toEqual([[9, 5, 1]]);
    expect(t.ev('RANK.EQ(7,{7,3.5,3.5,1,2})')).toBe(1);
    expect(t.ev('RANK.EQ(3.5,{7,3.5,3.5,1,2},1)')).toBe(3);
    expect(t.ev('RANK.AVG(94,{89,88,92,101,94,97,95})')).toBe(4);
    expect(t.ev('RANK.AVG(3.5,{7,3.5,3.5,1,2})')).toBe(2.5);
    expect(code(t.ev('RANK(10,{1,2})'))).toBe('#N/A');
    expect(t.ev('PERCENTILE.INC({1,3,2,4},0.3)')).toBe(1.9);
    expect(t.ev('PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.25)')).toBe(2.5);
    expect(code(t.ev('PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.01)'))).toBe('#NUM!');
    expect(code(t.ev('PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},2)'))).toBe('#NUM!');
    expect(t.ev('QUARTILE.INC({1,2,4,7,8,9,10,12},1)')).toBe(3.5);
    expect(t.ev('QUARTILE({1,2,4,7,8,9,10,12},4)')).toBe(12);
    expect(t.ev('QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)')).toBe(15);
    expect(t.ev('QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},3)')).toBe(43);
    expect(code(t.ev('QUARTILE.EXC({1,2,3},0)'))).toBe('#NUM!');
    const pr = '{13,12,11,8,4,3,2,1,1,1}';
    expect(t.ev(`PERCENTRANK.INC(${pr},2)`)).toBe(0.333);
    expect(t.ev(`PERCENTRANK.INC(${pr},4)`)).toBe(0.555);
    expect(t.ev(`PERCENTRANK.INC(${pr},8)`)).toBe(0.666);
    expect(t.ev(`PERCENTRANK.INC(${pr},5)`)).toBe(0.583);
    expect(t.ev(`PERCENTRANK(${pr},5,1)`)).toBe(0.5);
    expect(code(t.ev(`PERCENTRANK.INC(${pr},20)`))).toBe('#N/A');
    const pe = '{1,2,3,6,6,6,7,8,9}';
    expect(t.ev(`PERCENTRANK.EXC(${pe},7)`)).toBe(0.7);
    expect(t.ev(`PERCENTRANK.EXC(${pe},5.43)`)).toBe(0.381);
    expect(t.ev(`PERCENTRANK.EXC(${pe},5.43,1)`)).toBe(0.3);
  });
  it('FREQUENCY', () => {
    const t = setup();
    expect(t.arr('FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89})', 4, 1)).toEqual([[1], [2], [4], [2]]);
    expect(t.arr('FREQUENCY({1,2,3,4,5},{4,2})', 3, 1)).toEqual([[2], [2], [1]]);
  });
});

describe('Statistical: paired data and regression', () => {
  it('correlation and covariance', () => {
    const t = setup();
    close(t.ev('CORREL({3,2,4,5,6},{9,7,12,15,17})'), 0.997054486);
    close(t.ev('PEARSON({9,7,5,3,1},{10,6,1,5,3})'), 30 / Math.sqrt(1840));
    close(t.ev('RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})'), 0.057950192);
    expect(t.ev('COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})')).toBe(5.2);
    close(t.ev('COVARIANCE.S({2,4,8},{5,11,12})'), 9.666666667);
    expect(t.ev('COVAR({3,2,4,5,6},{9,7,12,15,17})')).toBe(5.2);
    expect(code(t.ev('CORREL({1,2,3},{1,2})'))).toBe('#N/A');
    expect(code(t.ev('CORREL({1,1,1},{1,2,3})'))).toBe('#DIV/0!');
  });
  it('simple linear regression', () => {
    const t = setup();
    const Y = '{2,3,9,1,8,7,5}', X = '{6,5,11,7,5,4,4}';
    close(t.ev(`SLOPE(${Y},${X})`), 0.305555556);
    close(t.ev('INTERCEPT({2,3,9,1,8},{6,5,11,7,5})'), 0.048387097);
    close(t.ev(`STEYX(${Y},${X})`), 3.305718950);
    close(t.ev('FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})'), 10.607253);
    close(t.ev('FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})'), 10.607253, 7);
  });
  it('LINEST, LOGEST, TREND, GROWTH', () => {
    const t = setup();
    expect(t.arr('LINEST({1,9,5,7},{0,4,2,3})', 1, 2)).toEqual([[2, 1]]);
    // Single-variable stats must agree with SLOPE/INTERCEPT/RSQ/STEYX.
    const Y = '{3,5,4,7,9,8,12}', X = '{1,2,3,4,5,6,7}';
    const st = t.arr(`LINEST(${Y},${X},TRUE,TRUE)`, 5, 2);
    close(st[0][0], t.ev(`SLOPE(${Y},${X})`) as number);
    close(st[0][1], t.ev(`INTERCEPT(${Y},${X})`) as number);
    close(st[2][0], t.ev(`RSQ(${Y},${X})`) as number);
    close(st[2][1], t.ev(`STEYX(${Y},${X})`) as number);
    expect(st[3][1]).toBe(5);
    close((st[4][0] as number) + (st[4][1] as number), t.ev(`DEVSQ(${Y})`) as number);
    // se of slope = steyx / sqrt(sxx); sxx = 28
    close(st[1][0], (st[2][1] as number) / Math.sqrt(28));
    expect(code(st[2][1 + 0] as V)).not.toBe('#N/A');
    expect(code(t.arr(`LINEST(${Y},${X},TRUE,TRUE)`, 3, 2)[2][1])).toBe(code(st[2][1]));
    // F statistic: (ssreg/1)/(ssresid/df)
    close(st[3][0], (st[4][0] as number) / ((st[4][1] as number) / 5));
    // Two regressors with exact relationship y = 1 + 2 x1 + 3 x2
    t.put('A1', 1); t.put('B1', 2); t.put('C1', 1 + 2 * 1 + 3 * 2);
    t.put('A2', 2); t.put('B2', 1); t.put('C2', 1 + 4 + 3);
    t.put('A3', 3); t.put('B3', 5); t.put('C3', 1 + 6 + 15);
    t.put('A4', 4); t.put('B4', 3); t.put('C4', 1 + 8 + 9);
    t.put('A5', 5); t.put('B5', 4); t.put('C5', 1 + 10 + 12);
    const m = t.arr('LINEST(C1:C5,A1:B5)', 1, 3) as number[][];
    close(m[0][0], 3); close(m[0][1], 2); close(m[0][2], 1);
    const nc = t.arr('LINEST({2,4,6.5},{1,2,3},FALSE)', 1, 2) as number[][];
    close(nc[0][0], (2 + 8 + 19.5) / 14);
    expect(nc[0][1]).toBe(0);
    // LOGEST: y = 2 * 3^x
    const lg = t.arr('LOGEST({6,18,54,162},{1,2,3,4})', 1, 2) as number[][];
    close(lg[0][0], 3); close(lg[0][1], 2);
    expect(code(t.ev('LOGEST({1,-1},{1,2})'))).toBe('#NUM!');
    expect(t.arr('TREND({1,3,5,7},{1,2,3,4},{5,6})', 1, 2)).toEqual([[9, 11]]);
    expect(t.arr('TREND({1;3;5;7})', 4, 1)).toEqual([[1], [3], [5], [7]]);
    const g = t.arr('GROWTH({6,18,54,162},{1,2,3,4},{5})', 1, 1) as number[][];
    close(g[0][0], 486);
  });
  it('FORECAST.ETS family (best effort)', () => {
    const t = setup();
    const vals = '{10,20,30,10,21,31,11,22,32,12,23,33}';
    const tl = '{1,2,3,4,5,6,7,8,9,10,11,12}';
    expect(t.ev(`FORECAST.ETS.SEASONALITY(${vals},${tl})`)).toBe(3);
    const f = t.ev(`FORECAST.ETS(13,${vals},${tl})`) as number;
    expect(f).toBeGreaterThan(10); expect(f).toBeLessThan(17);
    const f2 = t.ev(`FORECAST.ETS(15,${vals},${tl})`) as number;
    expect(f2).toBeGreaterThan(30); expect(f2).toBeLessThan(38);
    expect(t.ev(`FORECAST.ETS.STAT(${vals},${tl},8)`)).toBe(1);
    expect(t.ev(`FORECAST.ETS.CONFINT(13,${vals},${tl})`) as number).toBeGreaterThan(0);
    const lin = t.ev('FORECAST.ETS(6,{2,4,6,8,10},{1,2,3,4,5},0)') as number;
    close(lin, 12, 2);
    expect(code(t.ev('FORECAST.ETS(6,{1,2,3},{1,2,3.7})'))).toBe('#NUM!');
  });
});

describe('Statistical: distributions (Excel reference values)', () => {
  const t = setup();
  const cases: [string, number, number?][] = [
    ['NORM.DIST(42,40,1.5,TRUE)', 0.908788780274132],
    ['NORM.DIST(42,40,1.5,FALSE)', 0.109340049783996],
    ['NORMDIST(42,40,1.5,TRUE)', 0.908788780274132],
    ['NORM.INV(0.908789,40,1.5)', 42.000002, 6],
    ['NORM.S.DIST(1.333333,TRUE)', 0.908788726],
    ['NORMSDIST(1.333333)', 0.908788726],
    ['NORM.S.DIST(1.333333,FALSE)', 0.164010148],
    ['NORM.S.INV(0.908789)', 1.333334673],
    ['NORM.S.INV(0.001)', -3.090232306167813],
    ['NORM.S.INV(1E-10)', -6.361340902404056],
    ['NORM.S.DIST(-8,TRUE)', 6.22096057427178e-16],
    ['T.DIST(60,1,TRUE)', 0.994695326],
    ['T.DIST(8,3,FALSE)', 0.000736906],
    ['T.DIST.2T(1.959999998,60)', 0.054644930],
    ['T.DIST.RT(1.959999998,60)', 0.027322465],
    ['TDIST(1.959999998,60,2)', 0.054644930],
    ['T.INV(0.75,2)', 0.816496581],
    ['T.INV.2T(0.546449,60)', 0.606533076],
    ['TINV(0.05,10)', 2.228138852],
    ['CHISQ.DIST(0.5,1,TRUE)', 0.520499878],
    ['CHISQ.DIST(2,3,FALSE)', 0.207553749],
    ['CHISQ.DIST.RT(18.307,10)', 0.050000589],
    ['CHIDIST(18.307,10)', 0.050000589],
    ['CHISQ.INV(0.93,1)', 3.283020287],
    ['CHISQ.INV(0.6,2)', 1.832581464],
    ['CHISQ.INV.RT(0.050001,10)', 18.306973460],
    ['F.DIST(15.2069,6,4,TRUE)', 0.9900000430027937],
    ['F.DIST(15.2069,6,4,FALSE)', 0.0012237917087831735],
    ['F.DIST.RT(15.2069,6,4)', 0.0099999569972063, 7],
    ['F.INV(0.01,6,4)', 0.109309914],
    ['F.INV.RT(0.01,6,4)', 15.206864870],
    ['FINV(0.01,6,4)', 15.206864870],
    ['BETA.DIST(2,8,10,TRUE,1,3)', 0.685470581],
    ['BETA.DIST(2,8,10,FALSE,1,3)', 1.483764648],
    ['BETADIST(2,8,10,1,3)', 0.685470581],
    ['BETA.INV(0.685470581,8,10,1,3)', 2],
    ['BINOM.DIST(6,10,0.5,FALSE)', 0.205078125, 12],
    ['BINOM.DIST(6,10,0.5,TRUE)', 0.828125, 12],
    ['BINOM.DIST.RANGE(60,0.75,48)', 0.083974967],
    ['BINOM.DIST.RANGE(60,0.75,45,50)', 0.523629793],
    ['NEGBINOM.DIST(10,5,0.25,TRUE)', 0.3135141],
    ['NEGBINOM.DIST(10,5,0.25,FALSE)', 0.055048660],
    ['NEGBINOMDIST(10,5,0.25)', 0.055048660],
    ['HYPGEOM.DIST(1,4,8,20,TRUE)', 0.465428277],
    ['HYPGEOM.DIST(1,4,8,20,FALSE)', 0.363261094],
    ['HYPGEOMDIST(1,4,8,20)', 0.363261094],
    ['POISSON.DIST(2,5,TRUE)', 0.124652019],
    ['POISSON.DIST(2,5,FALSE)', 0.084224337],
    ['EXPON.DIST(0.2,10,TRUE)', 0.864664717],
    ['EXPON.DIST(0.2,10,FALSE)', 1.353352832],
    ['GAMMA.DIST(10.00001131,9,2,FALSE)', 0.032639],
    ['GAMMA.DIST(10.00001131,9,2,TRUE)', 0.068094],
    ['GAMMA.INV(0.068094,9,2)', 10.0000112],
    ['GAMMA(2.5)', 1.329340388],
    ['GAMMA(-0.75)', -4.834146544],
    ['GAMMALN(4)', 1.791759469],
    ['GAMMALN.PRECISE(4.5)', 2.453736571],
    ['WEIBULL.DIST(105,20,100,TRUE)', 0.9295813900692769, 12],
    ['WEIBULL.DIST(105,20,100,FALSE)', 0.035588864],
    ['LOGNORM.DIST(4,3.5,1.2,TRUE)', 0.039083556],
    ['LOGNORM.DIST(4,3.5,1.2,FALSE)', 0.017617597],
    ['LOGNORMDIST(4,3.5,1.2)', 0.039083556],
    ['LOGNORM.INV(0.039084,3.5,1.2)', 4.000025219],
    ['CONFIDENCE.NORM(0.05,2.5,50)', 0.692951912],
    ['CONFIDENCE.T(0.05,1,50)', 0.284196855],
    ['FISHER(0.75)', 0.972955075],
    ['FISHERINV(0.972955)', 0.749999967],
    ['PHI(0.75)', 0.301137432],
    ['GAUSS(2)', 0.477249868],
    ['STANDARDIZE(42,40,1.5)', 1.333333333],
  ];
  for (const [f, v, digits] of cases) {
    it(f, () => {
      const got = t.ev(f);
      const d = digits ?? (String(v).replace(/^-?0\.0*|\.|e.*$/g, '').length >= 9 ? 7 : 4);
      close(got, v, d);
    });
  }
  it('discrete inverses', () => {
    expect(t.ev('BINOM.INV(6,0.5,0.75)')).toBe(4);
    expect(t.ev('CRITBINOM(6,0.5,0.75)')).toBe(4);
  });
  it('inverse round trips at high precision', () => {
    for (const [p, df] of [[0.01, 3], [0.3, 7], [0.975, 30], [0.999, 2.5]]) {
      close(t.ev(`T.DIST(T.INV(${p},${df}),${df},TRUE)`), p, 12);
      close(t.ev(`CHISQ.DIST(CHISQ.INV(${p},${df}),${df},TRUE)`), p, 12);
      close(t.ev(`F.DIST(F.INV(${p},${df},5),${df},5,TRUE)`), p, 12);
      close(t.ev(`BETA.DIST(BETA.INV(${p},${df},2),${df},2,TRUE)`), p, 12);
      close(t.ev(`GAMMA.DIST(GAMMA.INV(${p},${df},3),${df},3,TRUE)`), p, 12);
      close(t.ev(`NORM.S.DIST(NORM.S.INV(${p}),TRUE)`), p, 13);
    }
  });
  it('distribution error cases', () => {
    expect(code(t.ev('NORM.DIST(1,0,0,TRUE)'))).toBe('#NUM!');
    expect(code(t.ev('NORM.INV(1,0,1)'))).toBe('#NUM!');
    expect(code(t.ev('T.DIST(1,0,TRUE)'))).toBe('#NUM!');
    expect(code(t.ev('T.DIST.2T(-1,2)'))).toBe('#NUM!');
    expect(code(t.ev('CHISQ.DIST(-1,2,TRUE)'))).toBe('#NUM!');
    expect(code(t.ev('BINOM.DIST(11,10,0.5,TRUE)'))).toBe('#NUM!');
    expect(code(t.ev('GAMMA(0)'))).toBe('#NUM!');
    expect(code(t.ev('GAMMA(-2)'))).toBe('#NUM!');
    expect(code(t.ev('FISHER(1)'))).toBe('#NUM!');
    expect(code(t.ev('BETA.DIST(4,8,10,TRUE,1,3)'))).toBe('#NUM!');
    expect(code(t.ev('CONFIDENCE.T(0.05,1,1)'))).toBe('#DIV/0!');
    expect(code(t.ev('NORM.DIST("a",0,1,TRUE)'))).toBe('#VALUE!');
    expect(code(t.ev('HYPGEOM.DIST(5,4,8,20,TRUE)'))).toBe('#NUM!');
  });
});

describe('Statistical: hypothesis tests', () => {
  const t = setup();
  it('T.TEST', () => {
    const A = '{3,4,5,8,9,1,2,4,5}', B = '{6,19,3,2,14,4,5,17,1}';
    close(t.ev(`T.TEST(${A},${B},2,1)`), 0.196015785);
    close(t.ev(`TTEST(${A},${B},1,1)`), 0.196015785 / 2);
    close(t.ev(`T.TEST(${A},${B},2,2)`), 0.191995887);
    close(t.ev(`T.TEST(${A},${B},2,3)`), 0.20229392336868413, 6); // Welch, fractional df
    expect(code(t.ev(`T.TEST(${A},${B},3,1)`))).toBe('#NUM!');
  });
  it('F.TEST, CHISQ.TEST, Z.TEST, PROB', () => {
    close(t.ev('F.TEST({6,7,9,15,21},{20,28,31,38,40})'), 0.648317846);
    close(t.ev('CHISQ.TEST({58,35;11,25;10,23},{45.35,47.65;17.56,18.44;16.09,16.91})'), 0.000308192, 6);
    close(t.ev('Z.TEST({3,6,7,8,6,5,4,2,1,9},4)'), 0.090574197);
    close(t.ev('Z.TEST({3,6,7,8,6,5,4,2,1,9},6)'), 0.863043389);
    close(t.ev('ZTEST({3,6,7,8,6,5,4,2,1,9},4,2)'), 0.04099516050019156, 12);
    close(t.ev('PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)'), 0.1);
    close(t.ev('PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)'), 0.8);
    expect(code(t.ev('PROB({0,1},{0.2,0.3},1)'))).toBe('#NUM!');
  });
});

describe('whole-column references and edge cases', () => {
  it('criteria functions clip whole columns but count blanks beyond the used range', () => {
    const t = setup({ A1: 'x', A2: 'y', A3: 'x', B1: 1, B2: 2, B3: 3 });
    expect(t.ev('SUMIF(A:A,"x",B:B)')).toBe(4);
    expect(t.ev('COUNTIF(A:A,"x")')).toBe(2);
    expect(t.ev('COUNTIFS(A:A,"x",B:B,">1")')).toBe(1);
    expect(t.ev('COUNTIF(A1:A10,"<>x")')).toBe(8);
    expect(t.ev('SUM(B:B)')).toBe(6);
    expect(t.ev('SUBTOTAL(9,B:B)')).toBe(6);
  });
  it('LINEST drops a collinear regressor', () => {
    const t = setup();
    // x2 = 2*x1, y = 3 + 2*x1
    const m = t.arr('LINEST({5;7;9;11},{1,2;2,4;3,6;4,8})', 1, 3) as number[][];
    close(m[0][2], 3);
    // one of the two coefficients carries the slope and the other is 0
    expect(Math.abs(m[0][0] * 2 + m[0][1] - 2)).toBeLessThan(1e-9);
    expect(m[0][0] === 0 || m[0][1] === 0).toBe(true);
  });
  it('errors propagate through aggregates', () => {
    const t = setup({ A1: 1, A2: '=NA()', A3: 3 });
    t.put('A2', '=1/0');
    expect(code(t.ev('AVERAGE(A1:A3)'))).toBe('#DIV/0!');
    expect(code(t.ev('MAX(A1:A3)'))).toBe('#DIV/0!');
    expect(t.ev('COUNT(A1:A3)')).toBe(2);
    expect(t.ev('COUNTA(A1:A3)')).toBe(3);
    expect(code(t.ev('LARGE(A1:A3,1)'))).toBe('#DIV/0!');
    expect(code(t.ev('CORREL(A1:A3,A1:A3)'))).toBe('#DIV/0!');
  });
});
