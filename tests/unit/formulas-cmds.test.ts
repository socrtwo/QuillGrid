import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { getFn } from '../../src/calc/functions';
import {
  FormulaStepper, validateDefinedName, sanitizeName, namesFromSelection, guessLabelEdges, splitArgs, callSpanAt, paramNames, findIssues, literalText,
} from '../../src/commands/formulas';

function book() {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  const set = (a: string, v: any) => {
    const m = /^([A-Z]+)(\d+)$/.exec(a)!;
    const c = m[1].charCodeAt(0) - 65, r = +m[2] - 1;
    b.transact('t', () => s.set(r, c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v }));
    e.recalc();
  };
  return { b, e, s, set };
}

describe('Evaluate Formula stepper', () => {
  it('replaces references, then operators, innermost first', () => {
    const t = book();
    t.set('A1', 2); t.set('A2', 3); t.set('B1', '=A1+A2*4');
    const st = new FormulaStepper(t.e, t.s, 0, 1);
    const seen: string[] = [];
    while (!st.done) { const p = st.next!; seen.push(st.text.slice(p.s, p.e)); st.step(); }
    expect(seen).toEqual(['A1', 'A2', '3*4', '2+12']);
    expect(st.text).toBe('14');
  });

  it('evaluates function calls once their arguments are constants and keeps ranges', () => {
    const t = book();
    t.set('A1', 1); t.set('A2', 2); t.set('A3', 3); t.set('B1', '=SUM(A1:A3)*ROUND(A1/3,2)');
    const st = new FormulaStepper(t.e, t.s, 0, 1);
    const seen: string[] = [];
    while (!st.done) { seen.push(st.text.slice(st.next!.s, st.next!.e)); st.step(); }
    expect(seen[0]).toBe('SUM(A1:A3)');
    expect(seen).toContain('A1');
    expect(seen).toContain('1/3');
    expect(st.text).toBe('1.98');
  });

  it('follows only the chosen IF branch and keeps ROW references', () => {
    const t = book();
    t.set('A1', 5); t.set('B1', '=IF(A1>3,ROW(A1)*10,1/0)');
    const st = new FormulaStepper(t.e, t.s, 0, 1);
    const seen: string[] = [];
    while (!st.done) { seen.push(st.text.slice(st.next!.s, st.next!.e)); st.step(); }
    expect(seen).not.toContain('1/0');
    expect(seen).toContain('ROW(A1)');
    expect(st.text).toBe('10');
  });

  it('handles text, negatives and offers Step In for formula cells', () => {
    const t = book();
    t.set('A1', -4); t.set('A2', '=A1*2'); t.set('B1', '="x"&A2&ABS(A1)');
    const st = new FormulaStepper(t.e, t.s, 0, 1);
    st.step(); // "x" is a literal, so A2 is the first reducible part
    expect(st.text.startsWith('"x"&-8')).toBe(true);
    expect(st.run()).toBe('"x-84"');
    const st2 = new FormulaStepper(t.e, t.s, 0, 1);
    expect(st2.next!.stepIn).toEqual({ sheet: t.s, r: 1, c: 0 });
    expect(literalText('a"b')).toBe('"a""b"');
  });
});

describe('defined names', () => {
  it('validates like the spreadsheet', () => {
    const t = book();
    t.b.names.push({ name: 'Rate', ref: 'Sheet1!$A$1', sheet: null });
    expect(validateDefinedName(t.b, 'Tax Rate', null)).toMatch(/spaces/);
    expect(validateDefinedName(t.b, 'A1', null)).toMatch(/cell reference/);
    expect(validateDefinedName(t.b, 'R1C1', null)).toMatch(/cell reference/);
    expect(validateDefinedName(t.b, 'r', null)).toMatch(/cell reference/);
    expect(validateDefinedName(t.b, '1abc', null)).toMatch(/begin/);
    expect(validateDefinedName(t.b, 'rate', null)).toMatch(/already exists/);
    expect(validateDefinedName(t.b, 'rate', 'Sheet1')).toBeNull();
    expect(validateDefinedName(t.b, 'Tax_Rate.2024', null)).toBeNull();
    expect(validateDefinedName(t.b, 'Rate', null, t.b.names[0])).toBeNull();
    expect(sanitizeName('Sales 2024')).toBe('Sales_2024');
    expect(sanitizeName('Q1')).toBe('Q1_');
    expect(sanitizeName('2024')).toBe('_2024');
  });

  it('creates names from the labels in a selection', () => {
    const t = book();
    t.set('A1', 'Item'); t.set('B1', 'Jan Sales'); t.set('C1', 'Feb');
    t.set('A2', 'Apples'); t.set('B2', 1); t.set('C2', 2);
    t.set('A3', 'Pears'); t.set('B3', 3); t.set('C3', 4);
    const rg = { r1: 0, c1: 0, r2: 2, c2: 2 };
    expect(guessLabelEdges(t.s, rg)).toEqual({ top: true, left: true });
    const list = namesFromSelection(t.s, rg, { top: true, left: true });
    expect(list.map((x) => [x.name, x.ref])).toEqual([
      ['Jan_Sales', 'Sheet1!$B$2:$B$3'], ['Feb', 'Sheet1!$C$2:$C$3'], ['Apples', 'Sheet1!$B$2:$C$2'], ['Pears', 'Sheet1!$B$3:$C$3'],
    ]);
    const top = namesFromSelection(t.s, rg, { top: true });
    expect(top[0]).toMatchObject({ name: 'Item', ref: 'Sheet1!$A$2:$A$3' });
    t.b.names.push(...list.map((x) => ({ name: x.name, ref: x.ref, sheet: null })));
    t.set('D1', '=SUM(Jan_Sales)+SUM(Pears)');
    expect(t.s.value(0, 3)).toBe(11);
  });
});

describe('function wizard helpers', () => {
  it('splits arguments and finds the call at the caret', () => {
    expect(splitArgs('A1,"a,b",SUM(1,2),{1,2;3,4}')).toEqual(['A1', '"a,b"', 'SUM(1,2)', '{1,2;3,4}']);
    const t = '=ROUND(SUM(A1:A3),2)+1';
    const sp = callSpanAt(t, t.indexOf('A3'), 'ROUND')!;
    expect(t.slice(sp.start, sp.end)).toBe('ROUND(SUM(A1:A3),2)');
    expect(sp.args).toEqual(['SUM(A1:A3)', '2']);
    const inner = callSpanAt(t, t.indexOf('A3'))!;
    expect(inner.name).toBe('SUM');
  });

  it('names variadic parameters', () => {
    expect(paramNames(getFn('SUM')!, 4)).toEqual(['number1', 'number2', 'number3', 'number4']);
    expect(paramNames(getFn('SUMIFS')!, 7)).toEqual(['sum_range', 'criteria_range1', 'criteria1', 'criteria_range2', 'criteria2', 'criteria_range3', 'criteria3']);
  });
});

describe('error checking', () => {
  it('finds errors, inconsistent formulas and numbers stored as text', () => {
    const t = book();
    t.set('A1', 1); t.set('A2', 2); t.set('A3', 3); t.set('A4', 4);
    t.set('B1', '=A1*2'); t.set('B2', '=A2*3'); t.set('B3', '=A3*2');
    t.set('C1', '=1/0'); t.set('D1', '42');
    t.b.transact('t', () => t.s.set(0, 3, { v: '42' }));
    const issues = findIssues(t.s);
    expect(issues.map((x) => [x.r, x.c, x.kind])).toEqual([[0, 2, 'error'], [0, 3, 'textNumber'], [1, 1, 'inconsistent']]);
  });
});
