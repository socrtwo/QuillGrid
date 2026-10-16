import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { formatValue } from '../../src/model/numfmt';

function book() {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  const set = (a: string, v: any) => {
    const m = /^([A-Z]+)(\d+)$/.exec(a)!;
    const c = m[1].charCodeAt(0) - 65, r = +m[2] - 1;
    b.transact('t', () => s.set(r, c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v }));
  };
  const get = (a: string) => { e.recalc(); const m = /^([A-Z]+)(\d+)$/.exec(a)!; return s.value(+m[2] - 1, m[1].charCodeAt(0) - 65); };
  return { b, e, s, set, get };
}

describe('engine', () => {
  it('arithmetic and precedence', () => {
    const t = book();
    t.set('A1', 2); t.set('A2', '=-A1^2'); t.set('A3', '=1+2*3-4/2'); t.set('A4', '=0.1+0.2=0.3'); t.set('A5', '=50%*2'); t.set('A6', '="a"&1&TRUE');
    expect(t.get('A2')).toBe(4); expect(t.get('A3')).toBe(5); expect(t.get('A4')).toBe(true); expect(t.get('A5')).toBe(1); expect(t.get('A6')).toBe('a1TRUE');
  });
  it('dependencies recalc', () => {
    const t = book();
    t.set('A1', 1); t.set('B1', '=A1*10'); t.set('C1', '=B1+1');
    expect(t.get('C1')).toBe(11);
    t.set('A1', 5);
    expect(t.get('C1')).toBe(51);
  });
  it('spills arrays and blocks', () => {
    const t = book();
    t.set('A1', '={1,2;3,4}');
    expect(t.get('B2')).toBe(4);
    t.set('B1', 'x');
    expect((t.get('A1') as any).code).toBe('#SPILL!');
    t.b.transact('c', () => t.s.set(0, 1, undefined));
    expect(t.get('A1')).toBe(1);
    expect(t.get('B1')).toBe(2);
  });
  it('LET and LAMBDA', () => {
    const t = book();
    t.set('A1', '=LET(x,2,y,3,x*y)'); t.set('A2', '=LAMBDA(a,b,a+b)(3,4)'); t.set('A3', '=IF(A1>5,"big","small")');
    expect(t.get('A1')).toBe(6); expect(t.get('A2')).toBe(7); expect(t.get('A3')).toBe('big');
  });
  it('circular refs give 0', () => {
    const t = book();
    t.set('A1', '=B1+1'); t.set('B1', '=A1+1');
    expect(t.get('A1')).toBeTypeOf('number');
  });
  it('undo/redo', () => {
    const t = book();
    t.set('A1', 1); t.set('A1', 2);
    t.b.undo(); expect(t.get('A1')).toBe(1);
    t.b.redo(); expect(t.get('A1')).toBe(2);
  });
  it('number formats', () => {
    expect(formatValue(1234.5, '#,##0.00').text).toBe('1,234.50');
    expect(formatValue(-1234.5, '#,##0.00').text).toBe('-1,234.50');
    expect(formatValue(0.256, '0.0%').text).toBe('25.6%');
    expect(formatValue(45366, 'yyyy-mm-dd').text).toBe('2024-03-15');
    expect(formatValue(45366.5, 'm/d/yyyy h:mm AM/PM').text).toBe('3/15/2024 12:00 PM');
    expect(formatValue(1.5, '# ?/?').text).toBe('1 1/2');
    expect(formatValue(12345.678, '0.00E+00').text).toBe('1.23E+04');
    expect(formatValue(5, '00000').text).toBe('00005');
    expect(formatValue(-5, '#,##0;[Red](#,##0)').text).toBe('(5)');
    expect(formatValue(-5, '#,##0;[Red](#,##0)').color).toBe('#ff0000');
    expect(formatValue(0, '#,##0;-#,##0;"zero"').text).toBe('zero');
    expect(formatValue('hi', '@').text).toBe('hi');
    expect(formatValue(1234567, '#,##0,"K"').text).toBe('1,235K');
    expect(formatValue(5551234567, '[<=9999999]###-####;(###) ###-####').text).toBe('(555) 123-4567');
    expect(formatValue(1/3, 'General').text).toBe('0.333333333');
    expect(formatValue(1.5, '[h]:mm:ss').text).toBe('36:00:00');
  });
});

describe('LET scope in lazy functions', () => {
  it('IF over an array inside LET sees LET names', () => {
    const b = new Workbook(); const e = new Engine(b);
    const m = e.evaluateMatrix('LET(g,SEQUENCE(2,2),IF(g>2,g,0))', b.sheet);
    expect(m).toEqual([[0, 0], [3, 4]]);
    expect(e.evaluateScalar('LET(x,5,IFERROR(x/0,x))', b.sheet, 0, 0)).toBe(5);
    expect(e.evaluateScalar('LET(f,LAMBDA(a,a*2),y,3,f(y)+y)', b.sheet, 0, 0)).toBe(9);
  });
});
