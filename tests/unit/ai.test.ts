import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { CellError } from '../../src/model/types';
import {
  extractFormula, validateFormula, parseJsonLoose, parseTable, parseRowAnswers, rowBatchPrompt, batchItems, runRowJob, tableText,
  cellContext, chartFromSpec, cfRulesFromSpec, validationFromSpec, cleanOutput, answerValue, refsIn,
} from '../../src/ai/apply';
import { FEATURES, BASE_SYSTEM } from '../../src/ai/features';
import type { ChatRequest, ChatResult } from '../../src/ai/client';

function sheetWith(rows: any[][]) {
  const book = new Workbook();
  const engine = new Engine(book);
  const s = book.sheet;
  rows.forEach((row, r) => row.forEach((v, c) => { if (v != null) s.set(r, c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v }); }));
  engine.rebuild(); engine.recalc(true);
  return { book, engine, s };
}

const reply = (text: string, extra: Partial<ChatResult> = {}): ChatResult => ({ text, model: 'mock', provider: 'mock', providerName: 'Mock', ...extra });

describe('formula extraction', () => {
  it('finds formulas in typical answers', () => {
    expect(extractFormula('=SUM(A1:A10)\nAdds the column.')).toEqual({ formula: '=SUM(A1:A10)', explanation: 'Adds the column.' });
    expect(extractFormula('Here is the formula:\n```excel\n=SUMIFS(C:C,A:A,"East")\n```\nThis sums East.')?.formula).toBe('=SUMIFS(C:C,A:A,"East")');
    expect(extractFormula('Use `=AVERAGE(B2:B9)` to get the mean.')?.formula).toBe('=AVERAGE(B2:B9)');
    expect(extractFormula('Formula: =IF(A1>10, "High", "Low") - returns High or Low')?.formula).toBe('=IF(A1>10, "High", "Low")');
    expect(extractFormula('XLOOKUP(E2,A:A,B:B,"Not found")')?.formula).toBe('=XLOOKUP(E2,A:A,B:B,"Not found")');
    expect(extractFormula('<think>hmm</think>\n=A1+B1 which adds them')?.formula).toBe('=A1+B1');
    expect(extractFormula('=A1 + B1 * 2')?.formula).toBe('=A1 + B1 * 2');
    expect(extractFormula('I cannot help with that.')).toBeNull();
  });
  it('validates with the parser and engine', () => {
    const { engine, s } = sheetWith([['Item', 'Qty'], ['a', 2], ['b', 3]]);
    const ok = validateFormula(engine, s, 5, 0, '=SUM(B2:B3)');
    expect(ok).toMatchObject({ ok: true, value: 5 });
    expect(validateFormula(engine, s, 5, 0, '=SUMM(B2:B3)')).toMatchObject({ ok: false, error: 'Unknown function SUMM' });
    expect(validateFormula(engine, s, 5, 0, '=SUM(B2:B3').ok).toBe(false);
    expect(validateFormula(engine, s, 5, 0, '=IF(1)').error).toMatch(/IF takes/);
    const err = validateFormula(engine, s, 5, 0, '=1/0');
    expect(err.ok).toBe(true);
    expect(err.resultError).toBe('#DIV/0!');
    const spill = validateFormula(engine, s, 5, 0, '=SEQUENCE(3)');
    expect(spill.matrix?.length).toBe(3);
  });
  it('lists references in a formula', () => {
    expect(refsIn('SUM(A1:B3)+$C$4*Sheet2!D5')).toEqual(['A1:B3', '$C$4', 'Sheet2!D5']);
  });
});

describe('JSON and table parsing', () => {
  it('parses loose JSON', () => {
    expect(parseJsonLoose('```json\n[1, 2, 3,]\n```')).toEqual([1, 2, 3]);
    expect(parseJsonLoose('Sure! {"a": 1} hope that helps')).toEqual({ a: 1 });
    expect(parseJsonLoose('no json here')).toBeUndefined();
  });
  it('reads tables in several shapes', () => {
    expect(parseTable('[{"State":"Ohio","Year":1803},{"State":"Maine","Year":1820,"Note":"x"}]')).toEqual([['State', 'Year', 'Note'], ['Ohio', 1803, null], ['Maine', 1820, 'x']]);
    expect(parseTable('{"columns":["a","b"],"rows":[[1,2],[3,4]]}')).toEqual([['a', 'b'], [1, 2], [3, 4]]);
    expect(parseTable('{"rows":[{"x":1}]}')).toEqual([['x'], [1]]);
    expect(parseTable('| Name | Age |\n|---|---|\n| Ann | 30 |\n| **Bo** | 4 |')).toEqual([['Name', 'Age'], ['Ann', '30'], ['Bo', '4']]);
    expect(parseTable('a\tb\n1\t2')).toEqual([['a', 'b'], ['1', '2']]);
    expect(parseTable('just words')).toBeNull();
  });
  it('cleans chatty output', () => {
    expect(cleanOutput('Sure, here is the answer:\nHello')).toBe('Hello');
    expect(cleanOutput('"Quoted"')).toBe('Quoted');
    expect(answerValue('42')).toBe(42);
    expect(answerValue('0.75', true)).toBe(0.75);
    expect(answerValue('Spain')).toBe('Spain');
  });
});

describe('row-by-row prompts and answers', () => {
  it('builds a numbered batch prompt', () => {
    const p = rowBatchPrompt('Translate into French.', ['Hello', 'Good\nbye'], { examples: [{ input: 'Yes', output: 'Oui' }] });
    expect(p).toContain('Task: Translate into French.');
    expect(p).toContain('an array of exactly 2 strings');
    expect(p).toContain('1. Hello\n2. Good ⏎ bye');
    expect(p).toContain('Input: Yes\nOutput: Oui');
    expect(rowBatchPrompt('Extract.', ['x'], { keys: ['First name', 'Last name'] })).toContain('"First name", "Last name"');
  });
  it('reads answers as JSON arrays, objects or numbered lines', () => {
    expect(parseRowAnswers('["a", "b"]', 2)).toEqual(['a', 'b']);
    expect(parseRowAnswers('{"results": ["a", 2]}', 2)).toEqual(['a', '2']);
    expect(parseRowAnswers('[{"first name":"Ann","Last name":"Lee"}]', 1, ['First name', 'Last name'])).toEqual([{ 'First name': 'Ann', 'Last name': 'Lee' }]);
    expect(parseRowAnswers('[["Ann","Lee"]]', 1, ['First', 'Last'])).toEqual([{ First: 'Ann', Last: 'Lee' }]);
    expect(parseRowAnswers('1. Positive\n2) Negative\n3: "Mixed"', 3)).toEqual(['Positive', 'Negative', 'Mixed']);
    expect(parseRowAnswers('["only one"]', 2)).toBeNull();
    expect(parseRowAnswers('Just this', 1)).toEqual(['Just this']);
  });
  it('batches by count and size', () => {
    expect(batchItems(['a', 'b', 'c', 'd', 'e'], 2)).toEqual([[0, 1], [2, 3], [4]]);
    expect(batchItems(['x'.repeat(10), 'y'.repeat(10), 'z'], 10, 15)).toEqual([[0], [1, 2]]);
  });
  it('runs a job with a mocked model, retrying a misaligned batch item by item', async () => {
    const seen: ChatRequest[] = [];
    const mock = async (req: ChatRequest) => {
      seen.push(req);
      const user = String(req.messages[1].content);
      const items = [...user.matchAll(/^(\d+)\. (.*)$/gm)].map((m) => m[2]);
      if (items.length === 3) return reply('["only", "two"]');
      return reply(JSON.stringify(items.map((x) => x.toUpperCase())));
    };
    const progress: number[] = [];
    const out = await runRowJob({ chat: mock, items: ['a', 'b', 'c', 'd'], instruction: 'Uppercase.', system: BASE_SYSTEM, batchSize: 3, onProgress: (d) => progress.push(d) });
    expect(out).toEqual(['A', 'B', 'C', 'D']);
    expect(seen.length).toBe(5); // batch of 3 (bad), 3 singles, batch of 1
    expect(progress).toEqual([3, 4]);
    expect(seen[0].messages[0].role).toBe('system');
  });
  it('stops when aborted and leaves unanswered items null', async () => {
    const ctl = new AbortController();
    const mock = async () => { ctl.abort(); return reply('["x"]'); };
    const out = await runRowJob({ chat: mock, items: ['a', 'b'], instruction: 'x', system: '', batchSize: 1, signal: ctl.signal });
    expect(out).toEqual(['x', null]);
  });
  it('treats a truncated answer as missing', async () => {
    const out = await runRowJob({ chat: async () => reply('["A"]', { truncated: true }), items: ['a'], instruction: 'x', system: '' });
    expect(out).toEqual([null]);
  });
});

describe('context and specs', () => {
  const data = [['Region', 'Month', 'Sales'], ['East', 'Jan', 100], ['West', 'Jan', 80], ['East', 'Feb', 120], ['West', 'Feb', '=B9']];
  it('describes a table for prompts', () => {
    const { s } = sheetWith(data);
    const t = tableText(s, { r1: 0, c1: 0, r2: 4, c2: 2 }, { header: true, formulas: true });
    expect(t.headers).toEqual(['Region', 'Month', 'Sales']);
    expect(t.text).toContain('Columns: A "Region" (text), B "Month" (text), C "Sales" (number)');
    expect(t.text).toContain('2\tEast\tJan\t100');
    expect(t.text).toContain('=B9 →');
    const small = tableText(s, { r1: 0, c1: 0, r2: 4, c2: 2 }, { header: true, maxRows: 2 });
    expect(small.truncated).toBe(true);
    expect(small.text).toContain('rows not shown');
    const c = cellContext(s, 4, 2);
    expect(c).toContain('Formula: =B9');
    expect(c).toContain('B9: B9=(empty)');
  });
  it('builds charts from header names', () => {
    const { s } = sheetWith(data);
    const ch = chartFromSpec(s, { r1: 0, c1: 0, r2: 4, c2: 2 }, ['Region', 'Month', 'Sales'], true, { type: 'line', title: 'Sales', x: 'Month', y: ['sales'] })!;
    expect(ch.type).toBe('line');
    expect(ch.series).toHaveLength(1);
    expect(ch.series[0].values).toBe('Sheet1!$C$2:$C$5');
    expect(ch.series[0].categories).toBe('Sheet1!$B$2:$B$5');
    expect(ch.at.c).toBe(4);
    expect(chartFromSpec(s, { r1: 0, c1: 0, r2: 4, c2: 2 }, ['Region', 'Month', 'Sales'], true, { type: 'weird', x: 'Region' })!.type).toBe('column');
  });
  it('validates conditional formatting and list specs', () => {
    const def = [{ r1: 1, c1: 2, r2: 4, c2: 2 }];
    const rules = cfRulesFromSpec([
      { type: 'cellIs', operator: 'greaterThan', formulas: ['=100'], style: { fill: '#ffc7ce', color: 'red' } },
      { type: 'expression', formulas: ['$A2="East"'], range: 'A2:C5' },
      { type: 'colorScale', colors: ['#f00', 'nope'] },
      { type: 'madeUp' },
      { type: 'cellIs', formulas: ['SUM('] },
      { type: 'containsText' },
    ], def);
    expect(rules.map((r) => r.type)).toEqual(['cellIs', 'expression', 'colorScale']);
    expect(rules[0]).toMatchObject({ operator: 'greaterThan', formulas: ['100'], style: { fill: '#ffc7ce' }, ranges: def });
    expect(rules[1].ranges).toEqual([{ r1: 1, c1: 0, r2: 4, c2: 2 }]);
    expect(rules[2].colors).toEqual(['#f8696b', '#ffeb84', '#63be7b']);
    const dv = validationFromSpec({ items: ['Open', 'Done', 'Done', 'A, B'], title: 'Status' }, def)!;
    expect(dv.formula1).toBe('"Open,Done,A  B"');
    expect(dv.inputTitle).toBe('Status');
    expect(validationFromSpec({ items: [] }, def)).toBeNull();
  });
});

describe('feature catalog', () => {
  it('has unique ids and builds prompts', () => {
    const ids = FEATURES.map((f) => f.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.length).toBeGreaterThanOrEqual(30);
    for (const f of FEATURES) {
      if (f.output === 'rows') { expect(f.rows).toBeTruthy(); expect(typeof f.rows!.instruction(f.options?.[0]?.value, 'x')).toBe('string'); continue; }
      const p = f.prompt({ text: 'DATA', arg: f.options?.[0]?.value, extra: 'REQUEST', lang: 'auto', address: 'B2' });
      expect(p.length).toBeGreaterThan(20);
    }
    const wf = FEATURES.find((f) => f.id === 'ai.writeFormula')!;
    expect(wf.prompt({ text: 'TABLE', extra: 'total sales', lang: 'auto', address: 'D10' })).toMatch(/cell D10[\s\S]*total sales[\s\S]*starting with "="[\s\S]*TABLE/);
  });
  it('error values are not numbers', () => {
    expect(new CellError('#N/A').toString()).toBe('#N/A');
  });
});

describe('ribbon coverage', () => {
  it('defines every AI, Tools, Developer and Draw ribbon command', async () => {
    await import('../../src/commands/ai');
    await import('../../src/commands/tools');
    await import('../../src/commands/developer');
    await import('../../src/commands/draw');
    const { DEFAULT_TABS } = await import('../../src/ui/ribbonLayout');
    const { getCommand, allCommands } = await import('../../src/commands/registry');
    const mine = DEFAULT_TABS.flatMap((t) => t.groups.flatMap((g) => g.items)).filter((id) => /^(ai|tools|dev|draw)\./.test(id));
    expect(mine.length).toBeGreaterThan(40);
    expect(mine.filter((id) => !getCommand(id))).toEqual([]);
    expect(allCommands().filter((c) => c.tab === 'AI').length + 1).toBeGreaterThanOrEqual(40);
  });
});

describe('SSE splitting', () => {
  it('handles every blank-line separator', async () => {
    const { takeSseBlocks } = await import('../../src/ai/client');
    for (const sep of ['\n\n', '\r\n\r\n', '\r\n\n', '\n\r\n']) {
      const [blocks, rest] = takeSseBlocks(`data: {"a":1}${sep}data: {"b":2}${sep}data: {"c"`);
      expect(blocks).toEqual(['data: {"a":1}', 'data: {"b":2}']);
      expect(rest).toBe('data: {"c"');
    }
  });
});
