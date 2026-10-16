import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { CellError } from '../../src/model/types';

function book() {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  const pos = (a: string) => { const m = /^([A-Z]+)(\d+)$/.exec(a)!; return { r: +m[2] - 1, c: m[1].charCodeAt(0) - 65 }; };
  const set = (a: string, v: any) => {
    const { r, c } = pos(a);
    b.transact('t', () => s.set(r, c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v }));
  };
  const get = (a: string) => { e.recalc(); const { r, c } = pos(a); return s.value(r, c); };
  /** Fills a block starting at `a` with rows of values. */
  const fill = (a: string, rows: any[][]) => {
    const { r, c } = pos(a);
    b.transact('t', () => rows.forEach((row, i) => row.forEach((v, j) => s.set(r + i, c + j, v == null ? undefined : typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v }))));
  };
  return { b, e, s, set, get, fill };
}

const T = book();
/** Evaluates a formula in a scratch cell. */
function ev(f: string): any {
  T.set('Z1000', '=' + f);
  return T.get('Z1000');
}
const err = (v: any) => (v instanceof CellError ? v.code : `not an error: ${v}`);
const near = (f: string, want: number, digits = 7) => {
  const v = ev(f);
  if (typeof v !== 'number') throw new Error(`${f} -> ${v instanceof CellError ? v.code : v}`);
  const tol = Math.max(Math.abs(want), 1) * Math.pow(10, -digits);
  if (Math.abs(v - want) > tol) throw new Error(`${f} = ${v}, expected ${want}`);
};

describe('financial: time value of money', () => {
  it('PMT/PV/FV/NPER/RATE', () => {
    near('PMT(0.08/12,10,10000)', -1037.03208935915, 12);
    near('PMT(0.06/12,18*12,0,50000)', -129.081160867991, 11);
    near('PMT(0,10,1000)', -100);
    near('PV(0.08/12,12*20,500,,0)', -59777.1458511878, 11);
    near('FV(0.06/12,10,-200,-500,1)', 2581.40337406012, 11);
    near('FV(0.12/12,12,-1000)', 12682.5030131662, 11);
    near('NPER(0.12/12,-100,-1000,10000,1)', 59.6738656742946, 11);
    near('NPER(0.01,-100,-1000,10000)', 60.0821228537617, 11);
    near('NPER(0.01,-100,-1000)', -9.57859403981306, 11);
    near('RATE(4*12,-200,8000)', 0.00770147248820165, 10);
    near('RATE(4*12,-200,8000)*12', 0.0924176698584198, 10);
    expect(err(ev('PMT(0.1,0,1000)'))).toBe('#NUM!');
  });
  it('IPMT/PPMT/ISPMT/CUMIPMT/CUMPRINC', () => {
    near('IPMT(0.1/12,1,3*12,8000)', -66.6666666666667, 11);
    near('IPMT(0.1,3,3,8000)', -292.447129909366, 11);
    near('PPMT(0.1/12,1,2*12,2000)', -75.6231860083663, 11);
    near('PPMT(0.08,10,10,200000)', -27598.0534624214, 11);
    near('ISPMT(0.1/12,1,3*12,8000000)', -64814.8148148148, 11);
    near('CUMIPMT(0.09/12,30*12,125000,13,24,0)', -11135.2321307508, 9);
    near('CUMIPMT(0.09/12,30*12,125000,1,1,0)', -937.5, 11);
    near('CUMPRINC(0.09/12,30*12,125000,13,24,0)', -934.107123420868, 9);
    near('CUMPRINC(0.09/12,30*12,125000,1,1,0)', -68.2782711809784, 9);
    expect(err(ev('IPMT(0.1,4,3,8000)'))).toBe('#NUM!');
    expect(err(ev('CUMIPMT(0.09/12,360,125000,13,24,2)'))).toBe('#NUM!');
  });
  it('NPV/IRR/MIRR/XNPV/XIRR', () => {
    near('NPV(0.1,-10000,3000,4200,6800)', 1188.44341233522, 11);
    near('NPV(0.08,{8000,9200,10000,12000,14500})+-40000', 1922.06155493237, 10);
    near('IRR({-70000,12000,15000,18000,21000,26000})', 0.0866309480365316, 9);
    near('IRR({-70000,12000,15000,18000,21000})', -0.0212448482734117, 9);
    near('IRR({-70000,12000,15000},-0.1)', -0.443506941334748, 9);
    near('MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)', 0.126094130365916, 10);
    near('MIRR({-120000,39000,30000,21000},0.1,0.12)', -0.0480446552704767, 9);
    near('XNPV(0.09,{-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})', 2086.64760203195, 10);
    near('XIRR({-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904},0.1)', 0.373362533578871, 8);
    near('XIRR({-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})', 0.373362533578871, 8);
    expect(err(ev('IRR({1,2,3})'))).toBe('#NUM!');
    expect(err(ev('XIRR({-10,5,6},{39448,39400,39904})'))).toBe('#NUM!');
  });
  it('EFFECT/NOMINAL/PDURATION/RRI/FVSCHEDULE', () => {
    near('EFFECT(0.0525,4)', 0.0535426673707584, 12);
    near('NOMINAL(0.053543,4)', 0.05250032, 7);
    near('PDURATION(0.025,2000,2200)', 3.85986616262266, 11);
    near('RRI(96,10000,11000)', Math.pow(1.1, 1 / 96) - 1, 12);
    near('RRI(96,10000,11000)', 0.0009933, 4);
    near('FVSCHEDULE(1,{0.09,0.11,0.1})', 1.33089, 12);
    expect(err(ev('EFFECT(0,4)'))).toBe('#NUM!');
  });
});

describe('financial: depreciation', () => {
  it('SLN/SYD/DB/DDB/VDB', () => {
    near('SLN(30000,7500,10)', 2250);
    near('SYD(30000,7500,10,1)', 4090.90909090909, 11);
    near('SYD(30000,7500,10,10)', 409.090909090909, 11);
    near('DB(1000000,100000,6,1,7)', 186083.333333333, 11);
    near('DB(1000000,100000,6,2,7)', 259639.416666667, 11);
    near('DB(1000000,100000,6,7,7)', 15845.10, 6);
    near('DDB(2400,300,10*365,1)', 1.31506849315068, 11);
    near('DDB(2400,300,10*12,1,2)', 40, 11);
    near('DDB(2400,300,10,1,2)', 480, 11);
    near('DDB(2400,300,10,2,1.5)', 306, 11);
    near('DDB(2400,300,10,10)', 22.1225472000002, 9);
    near('VDB(2400,300,10*365,0,1)', 1.31506849315068, 11);
    near('VDB(2400,300,10*12,0,1)', 40, 11);
    near('VDB(2400,300,10,0,1)', 480, 11);
    near('VDB(2400,300,10*12,6,18)', 396.31, 5);
    near('VDB(2400,300,10*12,6,18,1.5)', 311.81, 5);
    near('VDB(2400,300,10,0,0.875,1.5)', 315, 11);
  });
  it('AMORDEGRC/AMORLINC', () => {
    near('AMORDEGRC(2400,39679,39813,300,1,0.15,1)', 776);
    near('AMORLINC(2400,39679,39813,300,1,0.15,1)', 360);
    expect(err(ev('AMORLINC(2400,39679,39813,300,1,0.15,2)'))).toBe('#NUM!');
  });
});

describe('financial: securities', () => {
  it('coupon dates', () => {
    expect(ev('COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,1)')).toBe(71);
    expect(ev('COUPDAYS(DATE(2011,1,25),DATE(2011,11,15),2,1)')).toBe(181);
    expect(ev('COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,15),2,1)')).toBe(110);
    expect(ev('COUPNCD(DATE(2011,1,25),DATE(2011,11,15),2,1)')).toBe(40678);
    expect(ev('COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)')).toBe(4);
    expect(ev('COUPPCD(DATE(2011,1,25),DATE(2011,11,15),2,1)')).toBe(40497);
    expect(ev('COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,0)')).toBe(70);
    expect(ev('COUPDAYS(DATE(2011,1,25),DATE(2011,11,15),2,0)')).toBe(180);
    expect(ev('COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,15),2,0)')).toBe(110);
    // End-of-month maturity keeps month-end coupon dates.
    expect(ev('COUPPCD(DATE(2011,3,15),DATE(2012,8,31),2,1)')).toBe(ev('DATE(2011,2,28)'));
    expect(err(ev('COUPDAYS(DATE(2012,1,1),DATE(2011,1,1),2,1)'))).toBe('#NUM!');
    expect(err(ev('COUPDAYS(DATE(2011,1,1),DATE(2012,1,1),3,1)'))).toBe('#NUM!');
    expect(err(ev('COUPDAYS(DATE(2011,1,1),DATE(2012,1,1),2,5)'))).toBe('#NUM!');
  });
  it('PRICE/YIELD/DURATION/MDURATION', () => {
    near('PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)', 94.6343616213221, 10);
    near('YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)', 0.0650000068807314, 8);
    near('DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)', 10.9191452815919, 9);
    near('MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)', 5.73566981391884, 9);
    // Single remaining coupon uses the closed form.
    near('YIELD(DATE(2008,2,15),DATE(2008,5,15),0.0575,99.5,100,2,0)', (0.019375 / 1.009375) * 4, 12);
  });
  it('discounted and at-maturity securities', () => {
    near('ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)', 16.6666666666667, 11);
    near('ACCRINT(DATE(2008,3,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,FALSE)', 15.5555555555556, 11);
    near('ACCRINT(DATE(2008,3,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,TRUE)', 15.5555555555556, 11);
    near('ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)', 20.5479452054795, 11);
    near('DISC(DATE(2018,7,1),DATE(2048,1,1),97.975,100,1)', 0.000686384169029533, 9);
    near('INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)', 0.05768, 11);
    near('PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)', 99.7958333333333, 11);
    near('PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)', 99.9844988755569, 10);
    near('RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)', 1014584.6544071, 9);
    near('YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)', 0.0528225719868565, 10);
    near('YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)', 0.0609543336915754, 10);
  });
  it('Treasury bills', () => {
    near('TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)', (365 * 0.0914) / (360 - 0.0914 * 62), 12);
    near('TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)', 98.45, 11);
    near('TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)', (1.55 / 98.45) * (360 / 62), 12);
    expect(err(ev('TBILLPRICE(DATE(2008,3,31),DATE(2009,6,1),0.09)'))).toBe('#NUM!');
  });
  it('odd periods', () => {
    near('ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)', 113.597717474079, 9);
    near('ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)', 0.0772455415972989, 8);
    near('ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)', 99.8782860147213, 10);
    near('ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)', 0.0451922356291692, 9);
    // Round trip price -> yield.
    near('ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),100,2,1)', 0.0625, 10);
  });
  it('DOLLARDE/DOLLARFR', () => {
    near('DOLLARDE(1.02,16)', 1.125, 12);
    near('DOLLARDE(1.1,32)', 1.3125, 12);
    near('DOLLARFR(1.125,16)', 1.02, 12);
    near('DOLLARFR(1.125,32)', 1.04, 12);
    expect(err(ev('DOLLARDE(1.02,0)'))).toBe('#DIV/0!');
    expect(err(ev('DOLLARDE(1.02,-1)'))).toBe('#NUM!');
  });
  it('lifts scalar arguments over arrays', () => {
    const v = ev('PMT({0.01;0.02},12,1000)');
    expect(v).toBeCloseTo(-88.8487886783417, 9);
    expect(T.get('Z1001')).toBeCloseTo(-94.5595966229514, 9);
  });
});

/** Parses a complex-number result text into [re, im]. */
function cx(s: any): [number, number] {
  if (typeof s !== 'string') throw new Error(`not a complex text: ${s instanceof CellError ? s.code : s}`);
  const m = /^([+-]?[\d.]+(?:E[+-]?\d+)?)?(?:([+-]?)([\d.]+(?:E[+-]?\d+)?)?([ij]))?$/i.exec(s);
  if (!m) throw new Error(`bad complex ${s}`);
  const re = m[1] && !m[4] ? +m[1] : m[4] && m[1] && !/^[+-]/.test(m[2] || '') && m[2] === '' && !m[3] ? 0 : m[1] ? +m[1] : 0;
  let im = 0;
  if (m[4]) im = (m[2] === '-' ? -1 : 1) * (m[3] ? +m[3] : 1);
  return [re, im];
}
const cxNear = (f: string, re: number, im: number, digits = 12) => {
  const [r, i] = cx(ev(f));
  const tol = (x: number) => Math.max(Math.abs(x), 1) * Math.pow(10, -digits);
  if (Math.abs(r - re) > tol(re) || Math.abs(i - im) > tol(im)) throw new Error(`${f} = ${ev(f)}, expected ${re}${im >= 0 ? '+' : ''}${im}i`);
};

describe('engineering: Bessel', () => {
  it('matches reference values', () => {
    near('BESSELI(1.5,1)', 0.981666428, 8);
    // Exact J2(1.9); older documentation quotes 0.329925829 from a lower-precision approximation.
    near('BESSELJ(1.9,2)', 0.329925727692387, 12);
    near('BESSELJ(1.9,2)', 0.329925829, 6);
    near('BESSELK(1.5,1)', 0.277387804, 8);
    near('BESSELY(2.5,1)', 0.145918138, 8);
    near('BESSELJ(1,0)', 0.765197686557967, 13);
    near('BESSELJ(10,1)', 0.0434727461688614, 12);
    near('BESSELJ(20,0)', 0.167024664340583, 12);
    near('BESSELJ(-1,1)', -0.440050585744934, 13);
    near('BESSELY(1,0)', 0.088256964215677, 12);
    near('BESSELY(10,0)', 0.0556711672835994, 12);
    near('BESSELY(1,5)', -260.405866625812, 11);
    near('BESSELI(1,0)', 1.26606587775201, 13);
    near('BESSELI(10,0)', 2815.71662846625, 13);
    near('BESSELK(1,0)', 0.421024438240708, 12);
    near('BESSELK(10,1)', 1.86487734538256e-5, 11);
    near('BESSELJ(1.9,2.9)', 0.329925727692387, 12);
    expect(err(ev('BESSELY(0,1)'))).toBe('#NUM!');
    expect(err(ev('BESSELK(-1,1)'))).toBe('#NUM!');
    expect(err(ev('BESSELJ(1,-1)'))).toBe('#NUM!');
    expect(err(ev('BESSELJ("x",1)'))).toBe('#VALUE!');
  });
});

describe('engineering: bases and bits', () => {
  it('converts between bases', () => {
    expect(ev('BIN2DEC(1100100)')).toBe(100);
    expect(ev('BIN2DEC(1111111111)')).toBe(-1);
    expect(ev('BIN2HEX(11111011,4)')).toBe('00FB');
    expect(ev('BIN2HEX(1110)')).toBe('E');
    expect(ev('BIN2HEX(1111111111)')).toBe('FFFFFFFFFF');
    expect(ev('BIN2OCT(1001,3)')).toBe('011');
    expect(ev('BIN2OCT(1111111111)')).toBe('7777777777');
    expect(ev('DEC2BIN(9,4)')).toBe('1001');
    expect(ev('DEC2BIN(-100)')).toBe('1110011100');
    expect(ev('DEC2HEX(100,4)')).toBe('0064');
    expect(ev('DEC2HEX(-54)')).toBe('FFFFFFFFCA');
    expect(ev('DEC2HEX(28)')).toBe('1C');
    expect(ev('DEC2OCT(58,3)')).toBe('072');
    expect(ev('DEC2OCT(-100)')).toBe('7777777634');
    expect(ev('HEX2BIN("F",8)')).toBe('00001111');
    expect(ev('HEX2BIN("B7")')).toBe('10110111');
    expect(ev('HEX2BIN("FFFFFFFE00")')).toBe('1000000000');
    expect(ev('HEX2DEC("A5")')).toBe(165);
    expect(ev('HEX2DEC("FFFFFFFF5B")')).toBe(-165);
    expect(ev('HEX2DEC("3DA408B9")')).toBe(1034160313);
    expect(ev('HEX2OCT("F",3)')).toBe('017');
    expect(ev('HEX2OCT("3B4E")')).toBe('35516');
    expect(ev('HEX2OCT("FFFFFFFF00")')).toBe('7777777400');
    expect(ev('OCT2BIN(3,3)')).toBe('011');
    expect(ev('OCT2BIN(7777777000)')).toBe('1000000000');
    expect(ev('OCT2DEC(54)')).toBe(44);
    expect(ev('OCT2DEC(7777777533)')).toBe(-165);
    expect(ev('OCT2HEX(100,4)')).toBe('0040');
    expect(ev('OCT2HEX(7777777533)')).toBe('FFFFFFFF5B');
    expect(ev('DEC2BIN(3.9)')).toBe('11');
    expect(err(ev('DEC2BIN(512)'))).toBe('#NUM!');
    expect(err(ev('DEC2BIN(9,2)'))).toBe('#NUM!');
    expect(err(ev('BIN2DEC(102)'))).toBe('#NUM!');
    expect(err(ev('BIN2DEC(10000000000)'))).toBe('#NUM!');
    expect(err(ev('HEX2BIN("FFFFFFF000")'))).toBe('#NUM!');
    expect(err(ev('DEC2HEX(1,-1)'))).toBe('#NUM!');
  });
  it('bitwise operations', () => {
    expect(ev('BITAND(1,5)')).toBe(1);
    expect(ev('BITAND(13,25)')).toBe(9);
    expect(ev('BITOR(23,10)')).toBe(31);
    expect(ev('BITXOR(5,3)')).toBe(6);
    expect(ev('BITLSHIFT(4,2)')).toBe(16);
    expect(ev('BITRSHIFT(13,2)')).toBe(3);
    expect(ev('BITLSHIFT(13,-2)')).toBe(3);
    expect(ev('BITAND(2^47+1,2^47+3)')).toBe(2 ** 47 + 1);
    expect(ev('BITXOR(2^47,1)')).toBe(2 ** 47 + 1);
    expect(err(ev('BITAND(-1,1)'))).toBe('#NUM!');
    expect(err(ev('BITAND(1.5,1)'))).toBe('#NUM!');
    expect(err(ev('BITAND(2^48,1)'))).toBe('#NUM!');
    expect(err(ev('BITLSHIFT(2^47,1)'))).toBe('#NUM!');
  });
  it('DELTA/GESTEP/ERF', () => {
    expect(ev('DELTA(5,4)')).toBe(0);
    expect(ev('DELTA(5,5)')).toBe(1);
    expect(ev('DELTA(0.5)')).toBe(0);
    expect(ev('GESTEP(5,4)')).toBe(1);
    expect(ev('GESTEP(-4,-5)')).toBe(1);
    expect(ev('GESTEP(-1)')).toBe(0);
    near('ERF(0.745)', 0.70792892, 8);
    near('ERF(1)', 0.842700792949715, 14);
    near('ERF(0.1)', 0.112462916018285, 14);
    near('ERF(3)', 0.999977909503001, 14);
    near('ERF(-1)', -0.842700792949715, 14);
    near('ERF(0,1)', 0.842700792949715, 14);
    near('ERF(1,2)', 0.995322265018953 - 0.842700792949715, 13);
    near('ERF.PRECISE(0.745)', 0.70792892, 8);
    near('ERFC(1)', 0.157299207050285, 13);
    near('ERFC.PRECISE(1)', 0.157299207050285, 13);
    expect(ev('ERFC(3)')).toBeCloseTo(2.20904969985854e-5, 18);
    expect(ev('ERFC(5)')).toBeCloseTo(1.53745979442803e-12, 24);
    near('ERFC(-1)', 1.84270079294971, 13);
  });
});

describe('engineering: CONVERT', () => {
  it('converts units', () => {
    near('CONVERT(1,"lbm","kg")', 0.45359237, 14);
    near('CONVERT(68,"F","C")', 20, 13);
    near('CONVERT(6,"tsp","tbs")', 2, 14);
    near('CONVERT(6,"gal","l")', 22.712470704, 13);
    near('CONVERT(6,"mi","km")', 9.656064, 14);
    near('CONVERT(6,"km","mi")', 6000 / 1609.344, 14);
    near('CONVERT(6,"in","ft")', 0.5, 14);
    near('CONVERT(6,"cm","in")', 2.36220472440945, 13);
    near('CONVERT(1,"byte","bit")', 8);
    near('CONVERT(100,"ft^2","m^2")', 9.290304, 14);
    near('CONVERT(1,"km^2","m^2")', 1e6, 14);
    near('CONVERT(1,"cm3","m3")', 1e-6, 14);
    near('CONVERT(1,"kibyte","byte")', 1024);
    near('CONVERT(1,"Gibyte","Mibyte")', 1024);
    near('CONVERT(1,"kbit","bit")', 1000);
    near('CONVERT(1,"hr","mn")', 60);
    near('CONVERT(1,"yr","day")', 365.25);
    near('CONVERT(1,"atm","mmHg")', 760, 13);
    near('CONVERT(1,"atm","Pa")', 101325);
    near('CONVERT(1,"psi","Pa")', 6894.75729316836, 13);
    near('CONVERT(100,"C","K")', 373.15, 14);
    near('CONVERT(0,"C","F")', 32, 14);
    near('CONVERT(0,"K","Rank")', 0);
    near('CONVERT(80,"Reau","C")', 100, 14);
    near('CONVERT(1,"HP","W")', 745.69987158227, 13);
    near('CONVERT(1,"Nmi","m")', 1852);
    near('CONVERT(1,"ly","m")', 9460730472580800, 14);
    near('CONVERT(1,"BTU","J")', 1055.05585262, 14);
    near('CONVERT(1,"cal","J")', 4.1868, 14);
    near('CONVERT(1,"lbf","N")', 4.4482216152605, 13);
    near('CONVERT(1,"mph","m/s")', 0.44704, 14);
    near('CONVERT(1,"kn","m/h")', 1852, 14);
    near('CONVERT(1,"T","ga")', 10000, 14);
    near('CONVERT(1,"us_acre","m2")', 4046.87260987425, 13);
    near('CONVERT(1,"ha","ar")', 100, 14);
    near('CONVERT(1,"mm","m")', 0.001, 14);
    near('CONVERT(1,"ton","lbm")', 2000, 14);
    near('CONVERT(1,"uk_gal","l")', 4.54609, 14);
    expect(err(ev('CONVERT(2.5,"ft","sec")'))).toBe('#N/A');
    expect(err(ev('CONVERT(1,"xyz","m")'))).toBe('#N/A');
    expect(err(ev('CONVERT(1,"kft","m")'))).toBe('#N/A');
    expect(err(ev('CONVERT(1,"LBM","kg")'))).toBe('#N/A');
    expect(err(ev('CONVERT("a","m","ft")'))).toBe('#VALUE!');
  });
});

describe('engineering: complex numbers', () => {
  it('formats and parses', () => {
    expect(ev('COMPLEX(3,4)')).toBe('3+4i');
    expect(ev('COMPLEX(3,4,"j")')).toBe('3+4j');
    expect(ev('COMPLEX(0,1)')).toBe('i');
    expect(ev('COMPLEX(0,-1)')).toBe('-i');
    expect(ev('COMPLEX(1,0)')).toBe('1');
    expect(ev('COMPLEX(0,0)')).toBe('0');
    expect(ev('COMPLEX(-2.5,0.5)')).toBe('-2.5+0.5i');
    expect(err(ev('COMPLEX(1,1,"k")'))).toBe('#VALUE!');
    expect(ev('IMREAL("6-9i")')).toBe(6);
    expect(ev('IMAGINARY("3+4i")')).toBe(4);
    expect(ev('IMAGINARY("0-j")')).toBe(-1);
    expect(ev('IMAGINARY(4)')).toBe(0);
    expect(ev('IMAGINARY("-2.5j")')).toBe(-2.5);
    expect(ev('IMAGINARY("i")')).toBe(1);
    expect(ev('IMREAL("1.5E2+3i")')).toBe(150);
    expect(err(ev('IMREAL("3+4")'))).toBe('#NUM!');
    expect(err(ev('IMREAL("i4")'))).toBe('#NUM!');
    expect(err(ev('IMREAL(TRUE)'))).toBe('#VALUE!');
  });
  it('arithmetic', () => {
    expect(ev('IMABS("5+12i")')).toBe(13);
    near('IMARGUMENT("3+4i")', 0.927295218001612, 14);
    expect(err(ev('IMARGUMENT("0")'))).toBe('#DIV/0!');
    expect(ev('IMCONJUGATE("3+4i")')).toBe('3-4i');
    expect(ev('IMCONJUGATE("3-4j")')).toBe('3+4j');
    expect(ev('IMDIV("-238+240i","10+24i")')).toBe('5+12i');
    expect(err(ev('IMDIV("1+i","0")'))).toBe('#NUM!');
    expect(ev('IMPRODUCT("3+4i","5-3i")')).toBe('27+11i');
    expect(ev('IMPRODUCT("1+2i",30)')).toBe('30+60i');
    expect(ev('IMSUB("13+4i","5+3i")')).toBe('8+i');
    expect(ev('IMSUM("3+4i","5-3i")')).toBe('8+i');
    expect(ev('IMSUM("3+4j","5-3j")')).toBe('8+j');
    expect(ev('IMSUM({"1+i","2+2i";"3","4i"})')).toBe('6+7i');
    expect(err(ev('IMSUM("1+i","1+j")'))).toBe('#VALUE!');
    cxNear('IMPOWER("2+3i",3)', -46, 9, 13);
    cxNear('IMSQRT("1+i")', 1.09868411346781, 0.455089860562227, 14);
    cxNear('IMEXP("1+i")', 1.46869393991589, 2.28735528717884, 14);
    cxNear('IMLN("3+4i")', 1.6094379124341, 0.927295218001612, 14);
    cxNear('IMLOG10("3+4i")', 0.698970004336019, 0.402719196273373, 14);
    cxNear('IMLOG2("3+4i")', 2.32192809488736, 1.33780421245098, 14);
    cxNear('IMSIN("4+3i")', -7.61923172032141, -6.548120040911, 13);
    cxNear('IMCOS("1+i")', 0.833730025131149, -0.988897705762865, 14);
    cxNear('IMSINH("-4+3i")', 27.0168132580039, 3.85373803791938, 13);
    cxNear('IMCOSH("4+3i")', -27.0349456030742, 3.85115333481178, 13);
    cxNear('IMTAN("4+3i")', 0.00490825806749606, 1.00070953606723, 13);
    cxNear('IMCOT("4+3i")', 0.0049011823943045, -0.999266927805902, 13);
    cxNear('IMCSC("4+3i")', -0.0754898329158637, 0.0648774713706355, 13);
    cxNear('IMCSCH("1+i")', 0.303931001628426, -0.621518017170428, 13);
    cxNear('IMSEC("4+3i")', -0.0652940278579471, -0.0752249603027732, 13);
    cxNear('IMSECH("4+3i")', -0.0362534969158689, -0.00516434460775318, 13);
    expect(err(ev('IMLN("0")'))).toBe('#NUM!');
    // Results keep 15 significant digits.
    expect(ev('IMSQRT("1+i")')).toBe('1.09868411346781+0.455089860562227i');
    expect(ev('IMSQRT("-4")')).toMatch(/^1\.22464679914735E-16\+2i$|^2i$/);
  });
});

describe('database functions', () => {
  const t = book();
  t.fill('A1', [
    ['Tree', 'Height', 'Age', 'Yield', 'Profit', 'Height'],
    ['="=Apple"', '>10', null, null, null, '<16'],
    ['="=Pear"'],
    [],
    ['Tree', 'Height', 'Age', 'Yield', 'Profit'],
    ['Apple', 18, 20, 14, 105],
    ['Pear', 12, 12, 10, 96],
    ['Cherry', 13, 14, 9, 105],
    ['Apple', 14, 15, 10, 75],
    ['Pear', 9, 8, 8, 76.8],
    ['Apple', 8, 9, 6, 45],
  ]);
  t.fill('H1', [['Tree', null, 'Big'], ['Ch', null, '=E6>100']]);
  t.fill('H4', [['Tree'], ['=Ch']]);
  t.fill('H7', [['Profit', 'Age'], ['>100', '<15'], [null, '>18']]);
  const dv = (f: string) => { t.set('Z100', '=' + f); return t.get('Z100'); };
  it('aggregates matching records', () => {
    expect(dv('DCOUNT(A5:E11,"Age",A1:F2)')).toBe(1);
    expect(dv('DCOUNTA(A5:E11,"Profit",A1:F2)')).toBe(1);
    expect(dv('DMAX(A5:E11,"Profit",A1:A3)')).toBe(105);
    expect(dv('DMIN(A5:E11,"Profit",A1:B2)')).toBe(75);
    expect(dv('DAVERAGE(A5:E11,"Yield",A1:B2)')).toBe(12);
    expect(dv('DAVERAGE(A5:E11,3,A5:E11)')).toBe(13);
    expect(dv('DSUM(A5:E11,"Profit",A1:A2)')).toBe(225);
    expect(dv('DSUM(A5:E11,"Profit",A1:F2)')).toBe(75);
    expect(dv('DPRODUCT(A5:E11,"Yield",A1:B2)')).toBe(140);
    expect(dv('DVAR(A5:E11,"Yield",A1:A3)')).toBeCloseTo(8.8, 12);
    expect(dv('DVARP(A5:E11,"Yield",A1:A3)')).toBeCloseTo(7.04, 12);
    expect(dv('DSTDEV(A5:E11,"Yield",A1:A3)')).toBeCloseTo(2.96647939483827, 12);
    expect(dv('DSTDEVP(A5:E11,"Yield",A1:A3)')).toBeCloseTo(2.65329983228432, 12);
    expect(err(dv('DGET(A5:E11,"Yield",A1:A3)'))).toBe('#NUM!');
    expect(dv('DGET(A5:E11,"Yield",A1:F2)')).toBe(10);
    expect(err(dv('DGET(A5:E11,"Yield",H4:H5)'))).toBe('#VALUE!');
    expect(err(dv('DSUM(A5:E11,"Nope",A1:A2)'))).toBe('#VALUE!');
    expect(err(dv('DSUM(A5:E11,9,A1:A2)'))).toBe('#VALUE!');
  });
  it('text criteria match prefixes; "=" matches exactly', () => {
    expect(dv('DSUM(A5:E11,"Profit",H1:H2)')).toBe(105);
    expect(dv('DSUM(A5:E11,"Profit",H4:H5)')).toBe(0);
  });
  it('rows are OR-ed and columns AND-ed', () => {
    // (Profit > 100 AND Age < 15) OR Age > 18
    expect(dv('DCOUNT(A5:E11,"Profit",H7:I9)')).toBe(2);
    expect(dv('DCOUNTA(A5:E11,"Tree",H7:I9)')).toBe(2);
    expect(dv('DCOUNT(A5:E11,,H7:I9)')).toBe(2);
  });
  it('computed criteria', () => {
    expect(dv('DCOUNT(A5:E11,"Profit",J1:J2)')).toBe(2);
    expect(dv('DSUM(A5:E11,"Yield",J1:J2)')).toBe(23);
  });
});

describe('web functions', () => {
  it('ENCODEURL', () => {
    expect(ev('ENCODEURL("http://contoso.sharepoint.com/Finance/Profit and Loss Statement.xlsx")'))
      .toBe('http%3A%2F%2Fcontoso.sharepoint.com%2FFinance%2FProfit%20and%20Loss%20Statement.xlsx');
    expect(ev('ENCODEURL("a(b)!*~_.-")')).toBe('a%28b%29%21%2A~_.-');
    expect(ev('ENCODEURL("é")')).toBe('%C3%A9');
  });
  it('FILTERXML', () => {
    expect(ev('FILTERXML("<a><b>1</b><b>x</b></a>","//b")')).toBe(1);
    expect(T.get('Z1001')).toBe('x');
    expect(ev('FILTERXML("<a><b>hello</b></a>","/a/b")')).toBe('hello');
    expect(ev('FILTERXML("<a><b n=""7""/></a>","//b/@n")')).toBe(7);
    expect(err(ev('FILTERXML("<a><b>1</b></a>","//c")'))).toBe('#VALUE!');
    expect(err(ev('FILTERXML("<a><b>","//b")'))).toBe('#VALUE!');
  });
  it('WEBSERVICE is not supported', () => {
    expect(err(ev('WEBSERVICE("https://example.com")'))).toBe('#VALUE!');
  });
});
