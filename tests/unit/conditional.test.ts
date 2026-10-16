import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { CFEvaluator } from '../../src/grid/cf';
import {
  quickRule, insertRule, clearRulesIn, subtractRange, valueToFormula, formulaToValue, dataBarRule, colorScaleRule, iconSetRule, rulesFor,
  describeRule, parseRangesText, rangesText, FORMAT_PRESETS, normalizePriorities,
} from '../../src/commands/conditional';

function setup() {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  b.transact('t', () => { for (let i = 0; i < 10; i++) s.set(i, 0, { v: i + 1 }); s.set(0, 1, { v: 'apple pie' }); s.set(1, 1, { v: 'Banana' }); s.set(2, 1, { v: 'apple' }); s.set(3, 1, { v: 'Apple' }); });
  e.rebuild(); e.recalc(true);
  const cf = new CFEvaluator(e);
  let v = 0;
  const at = (r: number, c: number) => { cf.invalidate(); return cf.evaluate(s, r, c, ++v); };
  return { b, e, s, cf, at };
}

const colA = { r1: 0, c1: 0, r2: 9, c2: 0 };
const colB = { r1: 0, c1: 1, r2: 9, c2: 1 };
const red = FORMAT_PRESETS[0].style;

describe('conditional formatting rules', () => {
  it('greater than / between quick rules evaluate through CFEvaluator', () => {
    const t = setup();
    insertRule(t.s, quickRule('greaterThan', [colA], { value: '7' }, red));
    expect(t.at(7, 0)?.style?.fill).toBe('#ffc7ce');
    expect(t.at(6, 0)).toBeNull();
    t.s.cf = [];
    insertRule(t.s, quickRule('between', [colA], { value: 3, value2: '=$A$5' }, { color: '#006100' }));
    expect(t.at(2, 0)?.style?.color).toBe('#006100');
    expect(t.at(4, 0)?.style?.color).toBe('#006100');
    expect(t.at(5, 0)).toBeNull();
  });

  it('text contains, duplicates and top/bottom', () => {
    const t = setup();
    insertRule(t.s, quickRule('text', [colB], { text: 'apple' }, red));
    expect(t.at(0, 1)).not.toBeNull();
    expect(t.at(1, 1)).toBeNull();
    t.s.cf = [];
    insertRule(t.s, quickRule('duplicate', [colB], {}, red));
    expect(t.at(2, 1)).not.toBeNull(); // apple / Apple
    expect(t.at(0, 1)).toBeNull();
    t.s.cf = [];
    insertRule(t.s, quickRule('top', [colA], { rank: 3 }, red));
    expect([7, 8, 9].every((r) => t.at(r, 0))).toBe(true);
    expect(t.at(6, 0)).toBeNull();
    t.s.cf = [];
    insertRule(t.s, quickRule('bottomPct', [colA], { rank: 20 }, red));
    expect(t.at(1, 0)).not.toBeNull();
    expect(t.at(2, 0)).toBeNull();
    t.s.cf = [];
    insertRule(t.s, quickRule('above', [colA], {}, red));
    expect(t.at(5, 0)).not.toBeNull();
    expect(t.at(4, 0)).toBeNull();
  });

  it('formula rules use relative references from the top-left cell', () => {
    const t = setup();
    insertRule(t.s, { ranges: [{ r1: 0, c1: 0, r2: 9, c2: 1 }], type: 'expression', formulas: ['MOD($A1,2)=0'], style: { bold: true } });
    expect(t.at(1, 1)?.style?.bold).toBe(true); // row 2: A2 = 2
    expect(t.at(2, 1)).toBeNull();
  });

  it('data bars, color scales and icon sets', () => {
    const t = setup();
    insertRule(t.s, dataBarRule([colA], '#638ec6', true));
    expect(t.at(9, 0)?.bar?.pct).toBeCloseTo(1);
    t.s.cf = [];
    insertRule(t.s, colorScaleRule([colA], ['#ff0000', '#ffffff', '#00ff00']));
    expect(t.at(0, 0)?.fill).toBe('#ff0000');
    expect(t.at(9, 0)?.fill).toBe('#00ff00');
    t.s.cf = [];
    insertRule(t.s, iconSetRule([colA], '3Arrows'));
    expect(t.at(0, 0)?.icon?.index).toBe(0);
    expect(t.at(9, 0)?.icon?.index).toBe(2);
  });

  it('new rules take priority and stop-if-true is honored', () => {
    const t = setup();
    insertRule(t.s, { ranges: [colA], type: 'cellIs', operator: 'greaterThan', formulas: ['0'], style: { fill: '#111111' } });
    const second = insertRule(t.s, { ranges: [colA], type: 'cellIs', operator: 'greaterThan', formulas: ['5'], style: { fill: '#222222', bold: true } });
    expect(second.priority).toBe(1);
    expect(t.at(7, 0)?.style).toEqual({ fill: '#222222', bold: true });
    second.stopIfTrue = true;
    expect(t.at(7, 0)?.style).toEqual({ fill: '#222222', bold: true });
    expect(normalizePriorities(t.s.cf).map((r) => r.priority)).toEqual([1, 2]);
  });

  it('clears rules from part of a range', () => {
    const t = setup();
    insertRule(t.s, quickRule('greaterThan', [colA], { value: '0' }, red));
    expect(subtractRange(colA, { r1: 3, c1: 0, r2: 4, c2: 5 })).toEqual([{ r1: 0, c1: 0, r2: 2, c2: 0 }, { r1: 5, c1: 0, r2: 9, c2: 0 }]);
    clearRulesIn(t.s, [{ r1: 3, c1: 0, r2: 4, c2: 0 }]);
    expect(t.at(3, 0)).toBeNull();
    expect(t.at(5, 0)).not.toBeNull();
    expect(rulesFor(t.s, [{ r1: 4, c1: 0, r2: 4, c2: 0 }])).toEqual([]);
    clearRulesIn(t.s, [colA]);
    expect(t.s.cf.length).toBe(0);
  });

  it('value conversions, descriptions and ranges text', () => {
    expect(valueToFormula('12')).toBe('12');
    expect(valueToFormula('=$B$1')).toBe('$B$1');
    expect(valueToFormula('Late')).toBe('"Late"');
    expect(formulaToValue('"Late"')).toBe('Late');
    expect(formulaToValue('$B$1')).toBe('=$B$1');
    expect(describeRule({ id: 'x', ranges: [colA], type: 'cellIs', operator: 'between', formulas: ['1', '5'] })).toBe('Cell Value between 1 and 5');
    expect(rangesText([colA, { r1: 0, c1: 2, r2: 0, c2: 2 }])).toBe('=$A$1:$A$10,$C$1');
    expect(parseRangesText('=$A$1:$A$10,C1')).toEqual([colA, { r1: 0, c1: 2, r2: 0, c2: 2 }]);
    expect(parseRangesText('nonsense!!')).toBeNull();
  });
});
