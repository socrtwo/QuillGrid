import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook, Sheet } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { shiftCells, moveBlock, renameSheet, insertCellsShift } from '../../src/model/structure';
import { parseCell } from '../../src/model/address';

function setup() {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  const put = (sh: Sheet, a: string, v: any) => { const p = parseCell(a)!; b.transact('t', () => sh.set(p.r, p.c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v })); };
  const f = (sh: Sheet, a: string) => { const p = parseCell(a)!; return sh.get(p.r, p.c)?.f; };
  const v = (sh: Sheet, a: string) => { e.recalc(); const p = parseCell(a)!; return sh.value(p.r, p.c); };
  return { b, e, s, put, f, v };
}

describe('structure', () => {
  it('insert rows shifts references', () => {
    const t = setup();
    t.put(t.s, 'A1', 1); t.put(t.s, 'A2', 2); t.put(t.s, 'A3', '=SUM(A1:A2)'); t.put(t.s, 'B1', '=$A$2*2');
    t.b.transact('ins', () => shiftCells(t.s, 'row', 1, 2));
    expect(t.f(t.s, 'A5')).toBe('SUM(A1:A4)');
    expect(t.f(t.s, 'B1')).toBe('$A$4*2');
    expect(t.v(t.s, 'A5')).toBe(3);
  });
  it('delete rows makes #REF!', () => {
    const t = setup();
    t.put(t.s, 'A1', 5); t.put(t.s, 'B1', '=A1*2'); t.put(t.s, 'C3', '=A1+A2');
    t.b.transact('del', () => shiftCells(t.s, 'row', 0, -1));
    expect(t.f(t.s, 'B1')).toBeUndefined();
    expect(t.f(t.s, 'C2')).toBe('#REF!+A1');
  });
  it('delete inside a range shrinks it', () => {
    const t = setup();
    for (let i = 1; i <= 5; i++) t.put(t.s, `A${i}`, i);
    t.put(t.s, 'B1', '=SUM(A1:A5)');
    t.b.transact('del', () => shiftCells(t.s, 'row', 2, -2));
    expect(t.f(t.s, 'B1')).toBe('SUM(A1:A3)');
    expect(t.v(t.s, 'B1')).toBe(1 + 2 + 5);
  });
  it('insert columns and undo', () => {
    const t = setup();
    t.put(t.s, 'A1', 1); t.put(t.s, 'C1', '=A1');
    t.b.transact('ins', () => shiftCells(t.s, 'col', 1, 1));
    expect(t.f(t.s, 'D1')).toBe('A1');
    t.b.undo();
    expect(t.f(t.s, 'C1')).toBe('A1');
    t.b.redo();
    expect(t.f(t.s, 'D1')).toBe('A1');
  });
  it('other sheets follow', () => {
    const t = setup();
    const s2 = new Sheet(t.b, 'Data'); t.b.sheets.push(s2);
    t.put(s2, 'A1', 7); t.put(t.s, 'A1', '=Data!A1*2');
    t.b.transact('ins', () => shiftCells(s2, 'row', 0, 3));
    expect(t.f(t.s, 'A1')).toBe('Data!A4*2');
    expect(t.v(t.s, 'A1')).toBe(14);
  });
  it('move block keeps references pointing to moved cells', () => {
    const t = setup();
    t.put(t.s, 'A1', 3); t.put(t.s, 'B1', '=A1+1'); t.put(t.s, 'C1', '=A1*10');
    t.b.transact('mv', () => moveBlock(t.s, { r1: 0, c1: 0, r2: 0, c2: 1 }, 4, 0));
    expect(t.f(t.s, 'B5')).toBe('A5+1');
    expect(t.f(t.s, 'C1')).toBe('A5*10');
    expect(t.v(t.s, 'C1')).toBe(30);
  });
  it('rename sheet updates formulas', () => {
    const t = setup();
    const s2 = new Sheet(t.b, 'Data'); t.b.sheets.push(s2);
    t.put(s2, 'A1', 2); t.put(t.s, 'A1', '=Data!A1+1');
    t.b.transact('rn', () => renameSheet(t.b, s2, 'My Data'));
    expect(t.f(t.s, 'A1')).toBe("'My Data'!A1+1");
    expect(t.v(t.s, 'A1')).toBe(3);
  });
  it('insert cells shifting down', () => {
    const t = setup();
    t.put(t.s, 'A1', 1); t.put(t.s, 'A2', 2); t.put(t.s, 'B1', 9);
    t.b.transact('ins', () => insertCellsShift(t.s, { r1: 0, c1: 0, r2: 0, c2: 0 }, 'down'));
    expect(t.v(t.s, 'A2')).toBe(1); expect(t.v(t.s, 'A3')).toBe(2); expect(t.v(t.s, 'B1')).toBe(9);
  });
});
