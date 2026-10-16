import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell } from '../../src/model/address';
import { def, getCommand } from '../../src/commands/registry';
import type { DrawingObject } from '../../src/model/types';
import {
  createMacroApi, runJs, moduleProcs, toRelativeSteps, fromRelativeSteps, parseStepText, runRecorded, rowsToObjects, jsonToRows,
  controlAction, controlValue, controlItems, type MacroHost,
} from '../../src/commands/developer';

function host() {
  const book = new Workbook();
  const engine = new Engine(book);
  book.onChange = () => engine.recalc();
  const execs: { id: string; arg?: any }[] = [];
  const logs: any[] = [];
  const msgs: string[] = [];
  const h: MacroHost = {
    book,
    get sheet() { return book.sheet; },
    exec: (id, arg) => { execs.push({ id, arg }); return true; },
    msgBox: (t) => { msgs.push(t); },
    inputBox: () => '3',
    log: (...a) => logs.push(a),
    recalc: () => engine.recalc(),
  };
  const v = (a: string, s = book.sheet) => { engine.recalc(); const p = parseCell(a)!; return s.value(p.r, p.c); };
  return { book, engine, h, execs, logs, msgs, v };
}

describe('JavaScript macro API', () => {
  it('reads and writes cells, formulas and formats', () => {
    const t = host();
    const api = createMacroApi(t.h);
    api.cell('A1').value = 5;
    api.cell('A2').value = '=A1*2';
    expect(api.cell('A2').value).toBe(10);
    expect(api.cell('A2').formula).toBe('=A1*2');
    api.cell('B1').formula = '=SUM(A1:A2)';
    expect(t.v('B1')).toBe(15);
    api.cell('C1').value = 0.25;
    api.cell('C1').format = '0%';
    expect(api.cell('C1').format).toBe('0%');
    expect(api.cell('C1').text).toBe('25%');
    expect(api.cell('A1').offset(1, 0).address).toBe('A2');
    api.cell('D1').value = new Date(Date.UTC(2024, 0, 2));
    expect(typeof t.v('D1')).toBe('number');
  });
  it('gets and sets 2-D range values and works across sheets', () => {
    const t = host();
    const api = createMacroApi(t.h);
    api.range('A1:B2').values = [[1, 2], [3, 'x']];
    expect(api.range('A1:B2').values).toEqual([[1, 2], [3, 'x']]);
    api.range('C1:C3').values = [[7]] as any;
    expect(api.range('C1:C3').values).toEqual([[7], [7], [7]]);
    const s2 = api.addSheet('Data');
    s2.cell('A1').value = 'hello';
    expect(api.sheets).toEqual(['Sheet1', 'Data']);
    expect(api.sheet('Data').cell('A1').value).toBe('hello');
    expect(api.cell("Data!A1").value).toBe('hello');
    api.range('A1:B2').clear();
    expect(api.range('A1:B2').values).toEqual([[null, null], [null, null]]);
    expect(() => api.cell('not a ref!!')).toThrow();
  });
  it('select, exec, msgBox and log go to the host', () => {
    const t = host();
    const api = createMacroApi(t.h);
    api.select('B2:C3');
    expect(api.selection).toBe('B2:C3');
    expect(api.activeCell).toBe('B2');
    def({ id: 'test.noop', label: 'noop', tab: 'Test', run: () => 1 });
    api.exec('test.noop', 4);
    expect(t.execs).toEqual([{ id: 'test.noop', arg: 4 }]);
    expect(() => api.exec('no.such.command')).toThrow(/Unknown command/);
    api.msgBox('hi'); api.log('a', 1);
    expect(t.msgs).toEqual(['hi']);
    expect(t.logs).toEqual([['a', 1]]);
  });
  it('refuses locked cells on protected sheets', () => {
    const t = host();
    t.book.sheet.protection.on = true;
    const api = createMacroApi(t.h);
    expect(() => { api.cell('A1').value = 1; }).toThrow(/locked/);
  });
  it('runs module code as one undo step and calls named procedures', async () => {
    const t = host();
    const code = `function Fill() { for (let i = 1; i <= 3; i++) api.cell('A' + i).value = i * 10; }\nasync function Ask() { const n = Number(await api.inputBox('How many?')); api.cell('B1').value = n; return n; }`;
    expect(moduleProcs(code)).toEqual(['Fill', 'Ask']);
    await runJs(t.h, code, 'Fill');
    expect([t.v('A1'), t.v('A2'), t.v('A3')]).toEqual([10, 20, 30]);
    const before = t.book.undoStack.length;
    t.book.undo();
    expect(t.v('A1')).toBeNull();
    expect(t.v('A3')).toBeNull();
    expect(before).toBeGreaterThan(0);
    expect(await runJs(t.h, code, 'Ask')).toBe(3);
    expect(t.v('B1')).toBe(3);
    await expect(runJs(t.h, 'this is not js(')).rejects.toThrow(/Syntax error/);
    await expect(runJs(t.h, code, 'Missing')).rejects.toThrow(/No function named Missing/);
  });
});

describe('recorded macros', () => {
  it('converts positions to relative offsets and back', () => {
    const steps = [{ id: 'nav.select', arg: { ranges: [{ r1: 2, c1: 1, r2: 3, c2: 1 }], r: 2, c: 1 } }, { id: 'cell.enter', arg: { r: 2, c: 1, text: 'x' } }, { id: 'fmt.bold' }];
    const rel = toRelativeSteps(steps, { r: 2, c: 1 });
    expect(rel[0].arg).toEqual({ ranges: [{ r1: 0, c1: 0, r2: 1, c2: 0 }], r: 0, c: 0 });
    expect(rel[1].arg).toEqual({ r: 0, c: 0, text: 'x' });
    expect(rel[2]).toEqual({ id: 'fmt.bold' });
    const abs = fromRelativeSteps(rel, { r: 10, c: 5 });
    expect(abs[1].arg).toEqual({ r: 10, c: 5, text: 'x' });
    expect(fromRelativeSteps([{ id: 'cell.enter', arg: { r: -3, c: 0, text: 'y' } }], { r: 1, c: 0 })[0].arg.r).toBe(0);
  });
  it('parses edited step text', () => {
    expect(parseStepText('fmt.bold\n// comment\ncell.enter {"r":0,"c":0,"text":"hi"}\n')).toEqual([{ id: 'fmt.bold' }, { id: 'cell.enter', arg: { r: 0, c: 0, text: 'hi' } }]);
    expect(() => parseStepText('cell.enter {bad json')).toThrow(/JSON/);
  });
  it('replays steps through exec, relative to the active cell', async () => {
    const calls: { id: string; arg?: any }[] = [];
    def({ id: 'test.put', label: 'Put', tab: 'Test', run: () => true });
    const book = new Workbook();
    book.sheet.sel = { r: 5, c: 2, ranges: [{ r1: 5, c1: 2, r2: 5, c2: 2 }] };
    const app: any = { get sheet() { return book.sheet; }, book, exec: (id: string, arg?: any) => { calls.push({ id, arg }); return getCommand(id)!.run(app, arg); } };
    def({ id: 'cell.enter', label: 'Enter', tab: 'Test', internal: true, run: () => true });
    const n = await runRecorded(app, { name: 'M', created: 0, relative: true, steps: [{ id: 'cell.enter', arg: { r: 1, c: 0, text: 'a' } }, { id: 'test.put' }, { id: 'unknown.cmd' }] });
    expect(n).toBe(2);
    expect(calls).toEqual([{ id: 'cell.enter', arg: { r: 6, c: 2, text: 'a' } }, { id: 'test.put', arg: undefined }]);
  });
});

describe('form controls', () => {
  function withControl(ctl: DrawingObject['control'], extra: DrawingObject[] = []) {
    const book = new Workbook();
    const s = book.sheet;
    const d: DrawingObject = { id: 'c1', kind: 'shape', at: { r: 0, c: 0, dx: 0, dy: 0, w: 80, h: 20 }, control: ctl };
    s.drawings.push(d, ...extra);
    return { book, s, d };
  }
  it('check box toggles its linked cell', () => {
    const { s, d } = withControl({ type: 'checkbox', link: '$C$1' });
    controlAction(s, d);
    expect(s.value(0, 2)).toBe(true);
    controlAction(s, d);
    expect(s.value(0, 2)).toBe(false);
    expect(controlValue(s, d.control!)).toBe(false);
  });
  it('spin button and scroll bar step within limits', () => {
    const { s, d } = withControl({ type: 'spin', link: 'D1', min: 0, max: 3, step: 2 });
    controlAction(s, d, 'up'); controlAction(s, d, 'up'); controlAction(s, d, 'up');
    expect(s.value(0, 3)).toBe(3);
    controlAction(s, d, 'down');
    expect(s.value(0, 3)).toBe(1);
    const sc = withControl({ type: 'scroll', link: 'E1', min: 0, max: 100, step: 1, page: 10 });
    controlAction(sc.s, sc.d, 'pageUp');
    expect(sc.s.value(0, 4)).toBe(10);
  });
  it('option buttons write their index; lists write the chosen item number', () => {
    const second: DrawingObject = { id: 'c2', kind: 'shape', at: { r: 1, c: 0, dx: 0, dy: 0, w: 80, h: 20 }, control: { type: 'option', link: '$F$1' } };
    const { s } = withControl({ type: 'option', link: '$F$1' }, [second]);
    controlAction(s, second);
    expect(s.value(0, 5)).toBe(2);
    const l = withControl({ type: 'list', link: 'G1', input: 'A1:A3' });
    l.s.set(0, 0, { v: 'Red' }); l.s.set(1, 0, { v: 'Green' }); l.s.set(2, 0, { v: 'Blue' });
    expect(controlItems(l.s, l.d.control!)).toEqual(['Red', 'Green', 'Blue']);
    controlAction(l.s, l.d, 3);
    expect(l.s.value(0, 6)).toBe(3);
  });
  it('a check box without a link keeps its own state (undoable)', () => {
    const { book, s, d } = withControl({ type: 'checkbox' });
    controlAction(s, d);
    expect(s.drawings[0].control!.checked).toBe(true);
    book.undo();
    expect(s.drawings[0].control!.checked).toBeFalsy();
  });
});

describe('JSON import and export', () => {
  it('rows to objects with unique keys', () => {
    expect(rowsToObjects([['Name', 'Age', 'Name', null], ['Ann', 30, 'x', 1], [null, null, null, null], ['Bo', 4, 'y', 2]])).toEqual([
      { Name: 'Ann', Age: 30, Name_2: 'x', Column4: 1 }, { Name: 'Bo', Age: 4, Name_2: 'y', Column4: 2 },
    ]);
  });
  it('JSON to rows, flattening nested objects', () => {
    expect(jsonToRows([{ a: 1, b: { c: 'x', d: true } }, { a: 2, e: [1, 2] }])).toEqual([['a', 'b.c', 'b.d', 'e'], [1, 'x', true, null], [2, null, null, '1, 2']]);
    expect(jsonToRows({ results: [{ id: 1 }] })).toEqual([['id'], [1]]);
    expect(jsonToRows([[1, 2], [3, 4]])).toEqual([[1, 2], [3, 4]]);
    expect(() => jsonToRows(5)).toThrow();
  });
});
