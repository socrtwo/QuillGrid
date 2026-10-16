/**
 * AI tab commands. Every feature in src/ai/features.ts becomes a registered command, so each one
 * can go on the Quick Access Bar, a custom tab or a keyboard shortcut like any other command.
 *
 * Nothing is sent anywhere until the user runs an AI command with a provider set up (the free
 * on-device model needs no account). Output that changes cells is always previewed first and
 * applied as one undoable step. Quick Analysis works without any AI provider.
 */
import { def, getCommand, type MenuEntry } from './registry';
import type { App } from '../app';
import type { Sheet } from '../model/workbook';
import type { Scalar, CFRule, CellComment } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, cellName, colName, rangeName, type Range } from '../model/address';
import { rangeRefText } from '../model/structure';
import { FEATURES, BASE_SYSTEM, type AiFeature } from '../ai/features';
import { aiSettings, saveAiSettings, resolveProvider } from '../ai/store';
import { providerById } from '../ai/providers';
import { chat, AiError, type ChatMessage } from '../ai/client';
import { tableText, cellContext, cleanOutput, extractFormula, validateFormula, parseJsonLoose, parseTable, runRowJob, chartFromSpec, cfRulesFromSpec, validationFromSpec, writeTable, answerValue, type TableContext, type RowAnswer } from '../ai/apply';
import { aiBusy, aiIdle, paneResult, showFormulaPreview, showChangesPreview, showTablePreview, insertTableAtCursor, stopAi, aiModelName, type ChangeRow, type PaneAction } from '../ui/aiPane';
import { openDialog, toast, confirmDialog, promptDialog } from '../ui/dialog';
import { h, icon } from '../ui/dom';
import { activeRange, clipToUsed, dataRegion, hasHeaderRow, currentRegion, cellText, setInput } from './ops';
import { addSheetWithRows, checkWritable, canChange, totalsFormulas, uniqueSheet } from './tools';
import { Sheet as SheetClass, newId } from '../model/workbook';

// ============================================================================================
// Shared plumbing

const isCloud = () => { const p = providerById(aiSettings().active || ''); return !!p && p.kind !== 'local' && p.kind !== 'builtin'; };

/** Makes sure a provider is set up (offering the setup wizard) and, for cloud services, that the user knows where data goes. */
export async function ensureAiReady(app: App, sendsData: boolean, reason = 'This command needs an AI provider. Choose one below; free options, including one that runs on this device, are at the top.'): Promise<boolean> {
  const s = aiSettings();
  if (!s.enabled) { toast('AI features are turned off (AI ▸ AI Settings).', { action: 'Settings', onAction: () => app.exec('ai.settings') }); return false; }
  if (!(s.active && (await resolveProvider()))) {
    const { openAiSetup } = await import('../ui/aiWizard');
    if (!(await openAiSetup(app, { reason }))) return false;
  }
  if (!sendsData || !isCloud()) return true;
  const id = aiSettings().active!;
  if (aiSettings().privacyAck.includes(id)) return true;
  const p = providerById(id)!;
  const ok = await confirmDialog('Sending Data to an AI Service',
    `AI commands send the cells they need directly from this device to ${p.name}. QuillLattice has no server and keeps no copy. ${p.name}'s terms and privacy policy apply${p.free === 'free' ? ', and free tiers may use what you send to improve their models' : ''}. For confidential data, use the on-device model. Continue?`, 'Continue', 'Cancel');
  if (ok) saveAiSettings((x) => { x.privacyAck.push(id); });
  return ok;
}

async function confirmLargeSend(cells: number): Promise<boolean> {
  const s = aiSettings();
  if (!s.confirmDocument || !isCloud() || cells < 5000) return true;
  return confirmDialog('Send This Much Data?', `This sends about ${cells.toLocaleString()} cells to ${providerById(s.active!)?.name}. Select a smaller range to send less.`, 'Send', 'Cancel');
}

function systemPrompt(f?: AiFeature): string {
  const s = aiSettings();
  const parts = [BASE_SYSTEM];
  if (f?.system) parts.push(f.system);
  if (s.language && s.language !== 'auto') parts.push(`Write generated text in ${s.language} unless translating.`);
  return parts.join('\n\n');
}

function fail(e: unknown, app: App): void {
  if (e instanceof AiError && e.code === 'aborted') { toast('Stopped'); return; }
  const msg = e instanceof Error ? e.message : String(e);
  if (e instanceof AiError && (e.code === 'auth' || e.code === 'no-provider')) toast(msg, { action: 'Fix', onAction: () => app.exec('ai.providers'), ms: 8000 });
  else toast(msg, { ms: 8000 });
  console.warn('AI:', e);
}

const localLimit = () => (providerById(aiSettings().active || '')?.kind === 'local' ? 4000 : 14000);

/** The data region an AI command reads, with header detection. */
function regionContext(app: App, f?: AiFeature): TableContext | null {
  const s = app.sheet;
  let rg: Range | null;
  if (f?.input === 'sheet') rg = s.usedRange(false);
  else {
    const sel = activeRange(app);
    rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? dataRegion(app) : clipToUsed(s, sel);
    if (f?.id === 'ai.writeFormula' && rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = s.usedRange(false);
  }
  if (!rg) return null;
  return tableText(s, rg, { header: hasHeaderRow(s, rg), maxChars: localLimit(), formulas: f?.input === 'sheet' || f?.output === 'formula' });
}

const plainLabel = (f: AiFeature) => f.label.replace(/…$/, '');

// ============================================================================================
// Feature runner

export async function runFeature(app: App, f: AiFeature, arg?: string): Promise<void> {
  if (f.options && arg == null) {
    const r = await openDialog({ title: plainLabel(f), fields: [{ key: 'v', label: 'Choose', type: 'select', options: f.options.map((o) => ({ value: o.value, label: o.label })), value: f.options[0].value }] });
    if (!r) return;
    arg = r.v;
  }
  const sends = f.input !== 'none' || f.output === 'macro';
  if (!(await ensureAiReady(app, sends, `“${plainLabel(f)}” needs an AI provider. Choose one below; free options, including one that runs on this device, are at the top.`))) return;
  const tab = app.tab;
  const s = app.sheet;
  const { r, c } = s.sel;

  // Input checks before asking anything
  let ctx: TableContext | null = null;
  let text = '';
  if (f.input === 'cell') {
    const cell = s.get(r, c);
    if (!cell?.f) { toast(`Select a cell with a formula first (${cellName(r, c)} has none).`); return; }
    const around = dataRegion(app);
    ctx = tableText(s, currentRegion(s, r, c), { header: hasHeaderRow(s, around), maxChars: 4000, maxRows: 15 });
    text = cellContext(s, r, c, ctx);
  } else if (f.input === 'region' || f.input === 'selection' || f.input === 'sheet') {
    if (!s.usedRange(false)) { toast('The sheet is empty. Enter or paste some data first.'); return; }
    ctx = regionContext(app, f);
    if (!ctx && f.output !== 'formula') { toast('Select some data first.'); return; }
    text = ctx?.text || '';
  } else if (f.output === 'macro' && s.usedRange(false)) {
    text = tableText(s, s.usedRange(false)!, { header: true, maxChars: 2500, maxRows: 10 }).text;
  } else if (f.output === 'validation') {
    const rg = clipToUsed(s, activeRange(app));
    text = tableText(s, rg, { maxChars: 1500, maxRows: 40 }).text;
  }
  if (ctx && !(await confirmLargeSend((ctx.range.r2 - ctx.range.r1 + 1) * (ctx.range.c2 - ctx.range.c1 + 1)))) return;

  if (f.output === 'rows') { await runRows(app, f, arg); return; }

  let extra = '';
  if (f.ask) {
    const v = await promptDialog(f.ask.title, f.ask.label, '', f.ask.multiline ? ({ type: 'textarea', rows: 4, placeholder: f.ask.placeholder } as any) : ({ placeholder: f.ask.placeholder } as any));
    if (v == null || (!v.trim() && !f.ask.optional)) return;
    extra = v.trim();
  }
  if (f.output === 'validation' || f.output === 'cf') text = tableText(s, clipToUsed(s, activeRange(app)), { maxChars: 2000, maxRows: 30 }).text;
  const address = f.output === 'cf' || f.output === 'validation' ? s.sel.ranges.map((x) => rangeName(clipToUsed(s, x, 0))).join(',') : cellName(r, c);
  const userText = f.prompt({ text, arg, extra, lang: aiSettings().language, sheet: s.name, address });
  const messages: ChatMessage[] = [{ role: 'system', content: systemPrompt(f) }, { role: 'user', content: userText }];
  const base = { messages, tier: f.tier, temperature: f.temperature, maxTokens: f.maxTokens };
  const title = plainLabel(f) + (arg ? `: ${arg}` : '');

  try {
    switch (f.output) {
      case 'formula': {
        const original = f.input === 'cell' ? '=' + s.get(r, c)!.f : undefined;
        await showFormulaPreview(app, {
          title, sheet: s, r, c, original,
          generate: async (onToken, signal) => {
            const first = await chat({ ...base, signal, onToken: (_d, full) => onToken(full) });
            const fx = extractFormula(first.text);
            const check = fx ? validateFormula(tab.engine, s, r, c, fx.formula) : null;
            if (fx && check?.ok && !(f.id === 'ai.fixFormula' && check.resultError)) return first.text;
            // One retry with the problem explained.
            const problem = !fx ? 'No formula starting with "=" was found in your answer.' : check && !check.ok ? `That formula is not valid here: ${check.error}.` : `That formula still returns ${check?.resultError}.`;
            onToken(`${first.text}\n\n(Checking… ${problem} Asking again.)`);
            const second = await chat({ ...base, signal, messages: [...messages, { role: 'assistant', content: first.text }, { role: 'user', content: `${problem} Write a corrected formula. ${'Respond with the formula on the first line, starting with "=".'}` }] });
            return extractFormula(second.text) ? second.text : first.text;
          },
        });
        return;
      }
      case 'pane': {
        const out = paneResult(app, title);
        const b = aiBusy(`${plainLabel(f)}…`);
        try {
          const res = await chat({ ...base, signal: b.signal, onToken: (_d, full) => out.update(full) });
          out.done(cleanOutput(res.text));
          if (res.truncated) toast('The answer was cut off (length limit); check its end.');
        } catch (e) { out.fail(e instanceof AiError && e.code === 'aborted' ? 'Stopped' : (e as Error).message); throw e; } finally { aiIdle(b); }
        return;
      }
      default: {
        const b = aiBusy(`${plainLabel(f)}…`);
        let raw = '';
        try { raw = (await chat({ ...base, signal: b.signal })).text; } finally { aiIdle(b); }
        if (!app.docs.includes(tab)) { toast('The workbook was closed before the answer arrived.'); return; }
        await applyStructured(app, f, raw, ctx, title);
      }
    }
  } catch (e) {
    fail(e, app);
  }
}

/** Outputs that arrive as JSON (or code) and become workbook objects after a preview. */
async function applyStructured(app: App, f: AiFeature, raw: string, ctx: TableContext | null, title: string): Promise<void> {
  const s = app.sheet;
  const showRaw = (why: string) => { paneResult(app, title).done(cleanOutput(raw)); toast(`${why} The answer is shown in the AI pane.`, { ms: 6000 }); };
  switch (f.output) {
    case 'table': {
      const rows = parseTable(raw);
      if (!rows || rows.length < 2) { showRaw('The answer was not a table.'); return; }
      const where = await showTablePreview(title, rows);
      if (where === 'here') await insertTableAtCursor(app, rows, `AI: ${plainLabel(f)}`);
      else if (where === 'sheet') {
        if (app.book.structureLocked) { toast('The workbook structure is protected.'); return; }
        const book = app.book;
        let made!: Sheet;
        book.transact(`AI: ${plainLabel(f)}`, () => {
          book.recordBook();
          made = new SheetClass(book, uniqueSheet(book, f.id === 'ai.dataDictionary' ? `Columns of ${s.name}` : (f.id === 'ai.buildTemplate' ? title.replace(/^Build Template: /, '') : 'AI Table')));
          book.sheets.splice(book.activeSheet + 1, 0, made);
          writeTable(app, made, 0, 0, rows);
          made.freeze = { rows: 1, cols: 0 };
          for (let c = 0; c < rows[0].length; c++) made.colW.set(c, Math.min(260, Math.max(64, 8 * Math.max(...rows.slice(0, 50).map((x) => String(x[c] ?? '').length)))));
          book.activeSheet = book.sheets.indexOf(made);
        });
        app.exec('sheet.activate', book.sheets.indexOf(made));
      }
      return;
    }
    case 'chart': {
      if (!ctx) return;
      const spec = parseJsonLoose(raw);
      const chart = chartFromSpec(s, ctx.range, ctx.headers, ctx.header, spec);
      if (!chart) { showRaw('No chart could be built from the answer.'); return; }
      if (!checkWritableObjects(app)) return;
      app.book.transact('AI: Suggest Chart', () => { app.book.recordSheet(s); s.charts.push(chart); });
      app.grid.objects.select(chart.id);
      toast(spec?.reason ? `${chart.title}: ${String(spec.reason).slice(0, 160)}` : `Added a ${chart.type} chart`, { ms: 6000 });
      return;
    }
    case 'cf': {
      const ranges = s.sel.ranges.map((x) => clipToUsed(s, x));
      const rules = cfRulesFromSpec(parseJsonLoose(raw), ranges);
      if (!rules.length) { showRaw('No usable formatting rules were found.'); return; }
      const ok = await openDialog({ title: 'Apply Conditional Formatting?', content: h('div', null, h('p', { class: 'ql-help' }, `${rules.length} rule${rules.length === 1 ? '' : 's'}:`), h('ul', null, ...rules.map((x) => h('li', null, describeRule(x))))), ok: 'Apply', cancel: 'Discard' });
      if (!ok) return;
      if (!checkWritable(app, ranges)) return;
      app.book.transact('AI: Conditional Formatting', () => { app.book.recordSheet(s); const p = s.cf.length; rules.forEach((x, i) => { x.priority = p + i + 1; }); s.cf.push(...rules); });
      toast(`${rules.length} rule${rules.length === 1 ? '' : 's'} added (Home ▸ Conditional Formatting ▸ Manage Rules to edit)`);
      return;
    }
    case 'validation': {
      const ranges = s.sel.ranges.map((x) => clipToUsed(s, x, 0));
      const spec = parseJsonLoose(raw);
      const dv = validationFromSpec(spec, ranges);
      if (!dv) { showRaw('No list could be made from the answer.'); return; }
      const items = dv.formula1!.slice(1, -1).split(',');
      const ok = await openDialog({ title: 'Add Drop-Down List?', content: h('div', null, h('p', { class: 'ql-help' }, `${items.length} choices for ${ranges.map(rangeName).join(', ')}:`), h('div', { class: 'ql-ai-chips' }, ...items.map((t) => h('span', { class: 'ql-chip' }, t)))), ok: 'Add', cancel: 'Discard' });
      if (!ok) return;
      if (!checkWritable(app, ranges)) return;
      app.book.transact('AI: Drop-Down List', () => { app.book.recordSheet(s); s.validations.push(dv); });
      toast('Drop-down list added');
      return;
    }
    case 'macro': {
      let code = cleanOutput(raw).replace(/^```[\w-]*\n?|```$/g, '').trim();
      if (!/function\s+\w+\s*\(/.test(code)) code = `function Main() {\n${code}\n}`;
      const book = app.book;
      let i = 1;
      while (book.modules.some((m) => m.name === `AIMacro${i}`)) i++;
      const name = `AIMacro${i}`;
      book.transact('AI: Write Macro', () => { book.recordBook(); book.modules = [...book.modules, { name, code: `// Written by ${aiModelName()} — check it before running.\n${code}\n` }]; });
      app.exec('dev.macroEditor', name);
      return;
    }
    case 'analysis': {
      if (!ctx) return;
      const j = parseJsonLoose(raw);
      if (!j || !Array.isArray(j.insights)) { showRaw('The analysis was not in the expected form.'); return; }
      const lines = ['Insights', ...j.insights.map((x: any) => `• ${String(x)}`)];
      if (Array.isArray(j.questions) && j.questions.length) lines.push('', 'Questions to explore', ...j.questions.map((x: any) => `• ${String(x)}`));
      const actions: PaneAction[] = [];
      const tabRef = app.tab, sheetRef = s, c = ctx;
      if (j.chart) actions.push({ label: 'Add Chart', icon: 'add_chart', run: () => {
        if (app.tab !== tabRef || !app.book.sheets.includes(sheetRef)) { toast('Go back to the analyzed sheet first.'); return; }
        const chart = chartFromSpec(sheetRef, c.range, c.headers, c.header, j.chart);
        if (!chart) { toast('The suggested chart does not match the data.'); return; }
        app.book.transact('AI: Add Chart', () => { app.book.recordSheet(sheetRef); sheetRef.charts.push(chart); });
        app.grid.objects.select(chart.id);
      } });
      if (j.pivot && Array.isArray(j.pivot.values) && c.header) actions.push({ label: 'Create PivotTable', icon: 'pivot_table_chart', run: () => {
        const known = (xs: any) => (Array.isArray(xs) ? xs.map(String).filter((x: string) => c.headers.includes(x)) : []);
        const fns = ['sum', 'count', 'average', 'max', 'min'];
        const values = j.pivot.values.filter((v: any) => v && c.headers.includes(String(v.field))).map((v: any) => ({ field: String(v.field), fn: fns.includes(v.fn) ? v.fn : 'sum' }));
        app.exec('pivot.create', { source: rangeRefText(sheetRef, c.range), dest: 'new', rows: known(j.pivot.rows), cols: known(j.pivot.columns), filters: [], values });
      } });
      actions.push({ label: 'Insights to New Sheet', icon: 'note_add', run: () => {
        addSheetWithRows(app, 'Insights', [[{ v: `Insights about ${sheetRef.name}!${rangeName(c.range)}`, style: { bold: true, size: 13 } }], [`Generated by ${aiModelName()} on ${new Date().toLocaleDateString()}. AI can be wrong; check the numbers.`], [], ...j.insights.map((x: any) => [`• ${String(x)}`]), ...(j.questions?.length ? [[], [{ v: 'Questions to explore', style: { bold: true } }], ...j.questions.map((x: any) => [`• ${String(x)}`])] : [])], { colW: [720] });
      } });
      paneResult(app, title).done(lines.join('\n'), actions);
      return;
    }
    case 'anomalies': {
      if (!ctx) return;
      const j = parseJsonLoose(raw);
      const list = (Array.isArray(j) ? j : Array.isArray(j?.anomalies) ? j.anomalies : []).filter((x: any) => x && Number.isFinite(Number(x.row)));
      const rows = list.map((x: any) => ({ r: Number(x.row) - 1, col: String(x.column ?? ''), reason: String(x.reason ?? '') })).filter((x: any) => x.r >= ctx.range.r1 && x.r <= ctx.range.r2);
      if (!rows.length) { toast('Nothing unusual was found.'); return; }
      const colOf = (name: string) => { const i = ctx.headers.findIndex((hh) => hh.toLowerCase() === name.toLowerCase()); if (i >= 0) return ctx.range.c1 + i; const m = /^[A-Z]{1,3}$/i.exec(name); if (m) { const k = ctx.headers.findIndex((_x, j2) => colName(ctx.range.c1 + j2) === name.toUpperCase()); if (k >= 0) return ctx.range.c1 + k; } return ctx.range.c1; };
      const preview: ChangeRow[] = rows.map((x: any) => ({ label: cellName(x.r, colOf(x.col)), before: cellText(s, x.r, colOf(x.col)), after: x.reason, changed: true }));
      const picked = await showChangesPreview('Unusual Rows', preview, { columns: ['Value', 'Why it stands out'], applyLabel: 'Highlight and Add Notes', note: 'Selected rows are highlighted and the reason is added as a note.' });
      if (!picked) return;
      if (!checkWritable(app, [ctx.range])) return;
      const author = `AI (${aiModelName()})`;
      app.book.transact('AI: Find Anomalies', () => {
        app.book.recordSheet(s);
        for (const i of picked) {
          const x = rows[i];
          const cc = colOf(x.col);
          for (let c = ctx.range.c1; c <= ctx.range.c2; c++) { const cur = s.get(x.r, c); if (canChange(s, x.r, c)) s.set(x.r, c, { ...(cur || {}), s: app.book.derivedStyle(cur?.s ?? 0, { fill: c === cc ? '#f8cbad' : '#fde9d9' }) }); }
          const note: CellComment = { author, text: x.reason, time: Date.now(), note: true };
          s.comments.set(cellKey(x.r, cc), note);
        }
      });
      return;
    }
    case 'headers': {
      if (!ctx) return;
      const j = parseJsonLoose(raw);
      const heads: string[] = Array.isArray(j?.headers) ? j.headers.map((x: any) => String(x).slice(0, 60)) : Array.isArray(j) ? j.map(String) : [];
      if (!heads.length) { showRaw('No headers were suggested.'); return; }
      const rg = ctx.range;
      const replace = !!j?.hasHeaders || ctx.header;
      const preview: ChangeRow[] = heads.slice(0, rg.c2 - rg.c1 + 1).map((t, i) => ({ label: colName(rg.c1 + i), before: replace ? cellText(s, rg.r1, rg.c1 + i) : '', after: t, changed: !replace || cellText(s, rg.r1, rg.c1 + i) !== t }));
      const picked = await showChangesPreview(replace ? 'Rename Headers' : 'Add Header Row', preview, { columns: ['Current', 'Suggested'], note: replace ? '' : 'A new row is inserted above the data.' });
      if (!picked?.length) return;
      let row = rg.r1;
      if (!replace) {
        const above = rg.r1 > 0 && Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => s.get(rg.r1 - 1, rg.c1 + i)).every((x) => !x || (x.v == null && x.f == null));
        if (above) row = rg.r1 - 1;
        else { app.grid.select(rg.r1, rg.c1); app.exec('cells.insertRows', 'force'); row = rg.r1; }
      }
      if (!checkWritable(app, [{ r1: row, c1: rg.c1, r2: row, c2: rg.c2 }])) return;
      app.book.transact('AI: Headers', () => { for (const i of picked) { const cur = s.get(row, rg.c1 + i); s.set(row, rg.c1 + i, { ...(cur?.s ? { s: cur.s } : {}), v: heads[i] }); const x = s.get(row, rg.c1 + i)!; s.set(row, rg.c1 + i, { ...x, s: app.book.derivedStyle(x.s ?? 0, { bold: true }) }); } });
      return;
    }
  }
}

function checkWritableObjects(app: App): boolean {
  const s = app.sheet;
  if (s.protection.on && !s.protection.allow.includes('objects')) { app.exec('review.protectionNotice'); return false; }
  return true;
}

function describeRule(x: CFRule): string {
  const where = x.ranges.map(rangeName).join(', ');
  const st = x.style ? ` → ${[x.style.fill ? `fill ${x.style.fill}` : '', x.style.color ? `text ${x.style.color}` : '', x.style.bold ? 'bold' : ''].filter(Boolean).join(', ')}` : '';
  switch (x.type) {
    case 'cellIs': return `${where}: value ${x.operator} ${x.formulas?.join(' and ')}${st}`;
    case 'expression': return `${where}: when =${x.formulas?.[0]}${st}`;
    case 'colorScale': return `${where}: color scale ${x.colors?.join(' – ')}`;
    case 'dataBar': return `${where}: data bars ${x.barColor}`;
    case 'iconSet': return `${where}: icon set ${x.iconSet}`;
    case 'top10': return `${where}: ${x.bottom ? 'bottom' : 'top'} ${x.rank}${x.percent ? '%' : ''}${st}`;
    case 'containsText': return `${where}: contains “${x.text}”${st}`;
    default: return `${where}: ${x.type}${x.timePeriod ? ` (${x.timePeriod})` : ''}${st}`;
  }
}

// --------------------------------------------------------------------------------------------
// Row-by-row features

interface RowItem { r: number; c: number; text: string }

async function runRows(app: App, f: AiFeature, arg?: string, custom?: { instruction: string; mode: 'replace' | 'column'; header?: string }): Promise<void> {
  const spec = custom ? { item: custom.mode === 'replace' ? 'cell' as const : 'row' as const, mode: custom.mode, header: () => custom.header || 'AI', instruction: () => custom.instruction } : f.rows!;
  const s = app.sheet;
  const tab = app.tab;
  const sel = clipToUsed(s, activeRange(app));
  let extra = '';
  if (f.ask && !custom) {
    const v = await promptDialog(f.ask.title, f.ask.label, '', f.ask.multiline ? ({ type: 'textarea', rows: 3, placeholder: f.ask.placeholder } as any) : ({ placeholder: f.ask.placeholder } as any));
    if (v == null || (!v.trim() && !f.ask.optional)) return;
    extra = v.trim();
  }
  const items: RowItem[] = [];
  let outCol = -1;
  let headerRow = -1;
  let examples: { input: string; output: string }[] | undefined;
  const keys = 'keys' in spec && spec.keys ? spec.keys(arg, extra) : undefined;
  const textOnly = 'textOnly' in spec && spec.textOnly;

  if (spec.item === 'cell') {
    const rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? dataRegion(app) : sel;
    const header = rg.r2 > rg.r1 && hasHeaderRow(s, rg) ? rg.r1 : -1;
    const cols = spec.mode === 'replace' ? [rg.c1, rg.c2] : [s.sel.c >= rg.c1 && s.sel.c <= rg.c2 ? s.sel.c : rg.c1, s.sel.c >= rg.c1 && s.sel.c <= rg.c2 ? s.sel.c : rg.c1];
    for (let r = rg.r1; r <= rg.r2; r++) {
      if (r === header || (s.hiddenRows.has(r) && s.filter)) continue;
      for (let c = cols[0]; c <= cols[1]; c++) {
        const x = s.get(r, c);
        if (!x || x.f != null || x.v == null || x.v === '' || x.v instanceof CellError) continue;
        if (textOnly && typeof x.v !== 'string') continue;
        if (spec.mode === 'replace' && !canChange(s, r, c)) continue;
        items.push({ r, c, text: cellText(s, r, c) });
      }
    }
    if (spec.mode !== 'replace') { const reg = currentRegion(s, rg.r1, cols[0]); outCol = Math.max(reg.c2, rg.c2) + 1; headerRow = header >= 0 ? header : -1; }
  } else {
    // Rows of the data region; the target column is an empty selected column next to the data, or the first column after it.
    const empty = (c: number, r1: number, r2: number) => { for (let r = r1; r <= r2; r++) { const x = s.get(r, c); if (x && (x.v != null || x.f != null)) return false; } return true; };
    let region = currentRegion(s, sel.r1, sel.c1);
    if (sel.c1 > 0 && (region.r1 === region.r2 && region.c1 === region.c2) && !s.get(sel.r1, sel.c1)) region = currentRegion(s, sel.r1, sel.c1 - 1);
    const header = region.r2 > region.r1 && hasHeaderRow(s, region) ? region.r1 : -1;
    const bodyR1 = header >= 0 ? header + 1 : region.r1;
    if (f.id === 'ai.fillByExample') {
      outCol = sel.c1;
      if (outCol > region.c2 + 1 || outCol < region.c1) { toast('Select the column where you typed the examples, next to your data.'); return; }
    } else if (sel.c1 === sel.c2 && sel.c1 >= region.c1 && sel.c1 <= region.c2 + 1 && empty(sel.c1, bodyR1, region.r2)) outCol = sel.c1;
    else outCol = region.c2 + 1;
    const r2 = Math.max(region.r2, sel.r2 > region.r2 && sel.r2 - region.r2 < 2 ? sel.r2 : region.r2);
    const colsIn: number[] = [];
    for (let c = region.c1; c <= Math.max(region.c2, outCol - 1); c++) if (c !== outCol) colsIn.push(c);
    const label = (c: number) => (header >= 0 ? cellText(s, header, c) || colName(c) : colName(c));
    const rowText = (r: number) => colsIn.map((c) => `${label(c)}: ${cellText(s, r, c)}`).filter((t) => !/: $/.test(t)).join('; ');
    headerRow = header;
    if (f.id === 'ai.fillByExample') {
      examples = [];
      for (let r = bodyR1; r <= r2; r++) {
        if (s.hiddenRows.has(r) && s.filter) continue;
        const target = cellText(s, r, outCol);
        if (target) { if (examples.length < 12) examples.push({ input: rowText(r), output: target }); }
        else if (rowText(r)) items.push({ r, c: outCol, text: rowText(r) });
      }
      if (!examples.length) { toast(`Type two or three examples in column ${colName(outCol)} first; the AI fills the remaining rows the same way.`, { ms: 7000 }); return; }
    } else {
      for (let r = bodyR1; r <= r2; r++) { if (s.hiddenRows.has(r) && s.filter) continue; const t = rowText(r); if (t) items.push({ r, c: outCol, text: t }); }
    }
  }
  if (!items.length) { toast(textOnly ? 'There is no text in the selection.' : 'There is nothing to work on in the selection.'); return; }
  if (items.length > 5000) { toast('Select at most 5,000 cells or rows at a time.'); return; }
  if (!(await confirmLargeSend(items.reduce((a, x) => a + 1 + Math.floor(x.text.length / 40), 0)))) return;

  const instruction = spec.instruction(arg, extra);
  const b = aiBusy(`${plainLabel(f)}… 0 of ${items.length}`);
  let answers: RowAnswer[] = [];
  try {
    answers = await runRowJob({
      chat, items: items.map((x) => x.text), instruction, system: systemPrompt(f), keys, examples, tier: f.tier, temperature: f.temperature, signal: b.signal,
      batchSize: providerById(aiSettings().active || '')?.kind === 'local' ? 5 : 20,
      onProgress: (done, total) => { b.label(`${plainLabel(f)}… ${done} of ${total}`); b.progress(done, total); },
    });
  } catch (e) {
    if (!(e instanceof AiError && e.code === 'aborted')) { aiIdle(b); fail(e, app); return; }
  } finally { aiIdle(b); }
  const got = answers.filter((a) => a != null).length;
  if (!got) { toast(b.signal.aborted ? 'Stopped' : 'The AI returned no usable answers.'); return; }
  if (!app.docs.includes(tab)) return;

  const show = (a: RowAnswer) => (a == null ? '' : typeof a === 'string' ? a : Object.entries(a).map(([k, v]) => `${k}: ${v}`).join(' · '));
  const preview: ChangeRow[] = items.map((x, i) => {
    const a = answers[i];
    const before = spec.mode === 'replace' ? x.text : spec.mode === 'columns' ? '' : cellText(s, x.r, x.c);
    const after = show(a);
    return { label: cellName(x.r, spec.mode === 'replace' ? x.c : outCol), before: spec.mode === 'replace' ? before : x.text.slice(0, 120), after, changed: a != null && (spec.mode !== 'replace' || after !== before) };
  });
  const missing = items.length - got;
  const picked = await showChangesPreview(plainLabel(f) + (arg ? `: ${arg}` : ''), preview, {
    columns: spec.mode === 'replace' ? ['Before', 'After'] : ['Row', spec.mode === 'columns' ? (keys || []).join(' · ') : 'New value'],
    note: `${missing ? `${missing} item${missing === 1 ? '' : 's'} got no answer and stay unchanged. ` : ''}${spec.mode !== 'replace' ? `Results go to column ${colName(outCol)}${spec.mode === 'columns' && keys && keys.length > 1 ? `:${colName(outCol + keys.length - 1)}` : ''}.` : ''}`,
  });
  if (!picked?.length) return;
  if (app.sheet !== s) { const i = app.book.sheets.indexOf(s); if (i >= 0) app.exec('sheet.activate', i); }
  const width = spec.mode === 'columns' ? (keys?.length || 1) : 1;
  if (spec.mode !== 'replace') {
    const target = { r1: Math.min(...items.map((x) => x.r), headerRow >= 0 ? headerRow : Infinity), c1: outCol, r2: Math.max(...items.map((x) => x.r)), c2: outCol + width - 1 };
    if (!checkWritable(app, [target])) return;
    let busy = false;
    for (const i of picked) for (let k = 0; k < width && !busy; k++) { const x = s.get(items[i].r, outCol + k); if (x && (x.v != null || x.f != null) && !(f.id === 'ai.fillByExample')) busy = true; }
    if (busy && !(await confirmDialog('Replace Contents?', `Some cells in ${rangeName(target)} already have data. Replace them?`, 'Replace'))) return;
  }
  const writeAnswer = (r: number, c: number, text: string, numeric?: boolean) => {
    if (!canChange(s, r, c)) return;
    if (/^\d{4}-\d{2}(-\d{2})?$/.test(text.trim())) { setInput(app, s, r, c, text.trim()); return; }
    const v = answerValue(text, numeric);
    const cur = s.get(r, c);
    s.set(r, c, { ...(cur?.s ? { s: cur.s } : {}), ...(cur?.link ? { link: cur.link } : {}), ...(v === '' ? {} : { v }) });
  };
  app.book.transact(`AI: ${plainLabel(f)}`, () => {
    if (spec.mode !== 'replace' && headerRow >= 0) {
      const heads = spec.mode === 'columns' ? keys || [] : ['header' in spec && spec.header ? spec.header(arg, extra) : 'AI'];
      heads.forEach((t, k) => { const x = s.get(headerRow, outCol + k); if (!x || (x.v == null && x.f == null)) { const st = s.get(headerRow, outCol - 1)?.s; s.set(headerRow, outCol + k, { ...(st ? { s: st } : {}), v: t }); } });
    }
    for (const i of picked) {
      const x = items[i], a = answers[i];
      if (a == null) continue;
      if (typeof a === 'string') writeAnswer(x.r, spec.mode === 'replace' ? x.c : outCol, a, 'numeric' in spec ? spec.numeric : false);
      else (keys || []).forEach((k, j) => writeAnswer(x.r, outCol + j, a[k] ?? '', /score|amount|quantity/i.test(k)));
    }
  });
  toast(`${picked.length} cell${picked.length === 1 ? '' : 's'} updated. Undo (Ctrl+Z) reverses all of them.`);
}

// ============================================================================================
// Quick Analysis (no AI needed)

type QA = { label: string; icon: string; run: (app: App, rg: Range) => void | Promise<void> };

function qaRange(app: App): Range | null {
  const s = app.sheet;
  const sel = activeRange(app);
  const rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? dataRegion(app) : clipToUsed(s, sel);
  if (!s.usedRange(false) || (rg.r1 === rg.r2 && rg.c1 === rg.c2 && s.value(rg.r1, rg.c1) == null)) return null;
  return rg;
}

function numbersArea(s: Sheet, rg: Range): Range {
  // Skip a header row and a label column for formatting and totals.
  let r1 = rg.r1, c1 = rg.c1;
  if (rg.r2 > rg.r1 && hasHeaderRow(s, rg)) r1++;
  if (rg.c2 > rg.c1) { let text = 0; for (let r = r1; r <= rg.r2; r++) if (typeof s.value(r, rg.c1) === 'string') text++; if (text > (rg.r2 - r1 + 1) / 2) c1++; }
  return { r1, c1, r2: rg.r2, c2: rg.c2 };
}

function addCf(app: App, rg: Range, rule: Omit<CFRule, 'id' | 'ranges'>, label: string): void {
  const s = app.sheet;
  if (!checkWritable(app, [rg])) return;
  const area = numbersArea(s, rg);
  app.book.transact(label, () => { app.book.recordSheet(s); s.cf.push({ id: newId('cf'), ranges: [area], priority: s.cf.length + 1, ...rule } as CFRule); });
}

function addTotals(app: App, rg: Range, kind: 'sum' | 'average' | 'count' | 'percent' | 'running', dir: 'rows' | 'cols'): void {
  const s = app.sheet;
  const area = numbersArea(s, rg);
  const hasNum = (r1: number, c1: number, r2: number, c2: number) => { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (typeof s.value(r, c) === 'number') return true; return false; };
  const cells = totalsFormulas(area, kind, dir).filter((x) => (dir === 'rows' ? hasNum(area.r1, x.c, area.r2, x.c) : hasNum(x.r, area.c1, x.r, area.c2)));
  if (!cells.length) { toast('There are no numbers to total in the selection.'); return; }
  const target = dir === 'rows' ? { r1: area.r2 + 1, c1: area.c1, r2: area.r2 + 1, c2: area.c2 } : { r1: area.r1, c1: area.c2 + 1, r2: area.r2, c2: area.c2 + 1 };
  if (!checkWritable(app, [target])) return;
  const label = { sum: 'Sum', average: 'Average', count: 'Count', percent: '% Total', running: 'Running Total' }[kind];
  app.book.transact(`Quick Analysis: ${label}`, () => {
    for (const x of cells) {
      const cur = s.get(x.r, x.c);
      const st = app.book.derivedStyle(cur?.s ?? s.get(dir === 'rows' ? area.r2 : x.r, dir === 'rows' ? x.c : area.c2)?.s ?? 0, { bold: true, ...(x.numFmt ? { numFmt: x.numFmt } : {}), ...(dir === 'rows' ? { bt: { style: 'thin' as const } } : {}) });
      s.set(x.r, x.c, { f: x.f, s: st });
    }
    // Label next to / above the totals when there is room
    if (dir === 'rows' && area.c1 > rg.c1) { const lc = area.c1 - 1, x = s.get(area.r2 + 1, lc); if (!x || (x.v == null && x.f == null)) s.set(area.r2 + 1, lc, { v: label, s: app.book.styleId({ bold: true }) }); }
    if (dir === 'cols' && area.r1 > rg.r1) { const x = s.get(area.r1 - 1, area.c2 + 1); if (!x || (x.v == null && x.f == null)) s.set(area.r1 - 1, area.c2 + 1, { v: label, s: s.get(area.r1 - 1, area.c2)?.s ?? app.book.styleId({ bold: true }) }); }
  });
}

function runOther(app: App, id: string, rg: Range, arg?: any): void {
  if (!getCommand(id)) { toast('That command is not available in this build.'); return; }
  app.grid.select(rg.r1, rg.c1, { range: rg, keepActive: false });
  app.exec(id, arg);
}

const QA_TABS: Record<string, QA[]> = {
  Formatting: [
    { label: 'Data Bars', icon: 'data_thresholding', run: (app, rg) => addCf(app, rg, { type: 'dataBar', barColor: '#638ec6', gradientBar: true }, 'Data Bars') },
    { label: 'Color Scale', icon: 'gradient', run: (app, rg) => addCf(app, rg, { type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] }, 'Color Scale') },
    { label: 'Icon Set', icon: 'arrow_upward', run: (app, rg) => addCf(app, rg, { type: 'iconSet', iconSet: '3Arrows' }, 'Icon Set') },
    { label: 'Greater Than', icon: 'keyboard_double_arrow_right', run: async (app, rg) => {
      const s = app.sheet, area = numbersArea(s, rg);
      const xs: number[] = []; for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) { const v = s.value(r, c); if (typeof v === 'number') xs.push(v); }
      const avg = xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0;
      const v = await promptDialog('Greater Than', 'Format cells that are greater than:', String(+avg.toFixed(2)));
      if (v == null || v.trim() === '') return;
      addCf(app, rg, { type: 'cellIs', operator: 'greaterThan', formulas: [v.trim().replace(/^=/, '')], style: { fill: '#ffc7ce', color: '#9c0006' } }, 'Greater Than');
    } },
    { label: 'Top 10%', icon: 'vertical_align_top', run: (app, rg) => addCf(app, rg, { type: 'top10', rank: 10, percent: true, style: { fill: '#ffc7ce', color: '#9c0006' } }, 'Top 10%') },
    { label: 'Duplicate Values', icon: 'content_copy', run: (app, rg) => addCf(app, rg, { type: 'duplicateValues', style: { fill: '#ffc7ce', color: '#9c0006' } }, 'Duplicate Values') },
    { label: 'Clear Format', icon: 'format_color_reset', run: (app, rg) => {
      const s = app.sheet;
      const hit = (x: CFRule) => x.ranges.some((q) => q.r1 <= rg.r2 && q.r2 >= rg.r1 && q.c1 <= rg.c2 && q.c2 >= rg.c1);
      const n = s.cf.filter(hit).length;
      if (!n) { toast('There are no conditional formatting rules here.'); return; }
      if (!checkWritable(app, [rg])) return;
      app.book.transact('Clear Rules', () => { app.book.recordSheet(s); s.cf = s.cf.filter((x) => !hit(x)); });
      toast(`${n} rule${n === 1 ? '' : 's'} removed`);
    } },
  ],
  Charts: [
    { label: 'Column', icon: 'bar_chart', run: (app, rg) => runOther(app, 'insert.chartColumn', rg) },
    { label: 'Line', icon: 'show_chart', run: (app, rg) => runOther(app, 'insert.chartLine', rg) },
    { label: 'Pie', icon: 'pie_chart', run: (app, rg) => runOther(app, 'insert.chartPie', rg) },
    { label: 'Bar', icon: 'align_horizontal_left', run: (app, rg) => runOther(app, 'insert.chartBar', rg) },
    { label: 'Scatter', icon: 'scatter_plot', run: (app, rg) => runOther(app, 'insert.chartScatter', rg) },
    { label: 'More Charts', icon: 'add_chart', run: (app, rg) => runOther(app, getCommand('insert.recommendedCharts') ? 'insert.recommendedCharts' : 'insert.chartDialog', rg) },
  ],
  Totals: [
    { label: 'Sum', icon: 'functions', run: (app, rg) => addTotals(app, rg, 'sum', 'rows') },
    { label: 'Average', icon: 'avg_pace', run: (app, rg) => addTotals(app, rg, 'average', 'rows') },
    { label: 'Count', icon: 'tag', run: (app, rg) => addTotals(app, rg, 'count', 'rows') },
    { label: '% Total', icon: 'percent', run: (app, rg) => addTotals(app, rg, 'percent', 'rows') },
    { label: 'Running Total', icon: 'trending_up', run: (app, rg) => addTotals(app, rg, 'running', 'rows') },
    { label: 'Sum (Columns)', icon: 'functions', run: (app, rg) => addTotals(app, rg, 'sum', 'cols') },
    { label: 'Average (Columns)', icon: 'avg_pace', run: (app, rg) => addTotals(app, rg, 'average', 'cols') },
    { label: 'Count (Columns)', icon: 'tag', run: (app, rg) => addTotals(app, rg, 'count', 'cols') },
    { label: '% Total (Columns)', icon: 'percent', run: (app, rg) => addTotals(app, rg, 'percent', 'cols') },
    { label: 'Running Total (Columns)', icon: 'trending_up', run: (app, rg) => addTotals(app, rg, 'running', 'cols') },
  ],
  Tables: [
    { label: 'Table', icon: 'table', run: (app, rg) => runOther(app, 'insert.table', rg) },
    { label: 'PivotTable', icon: 'pivot_table_chart', run: (app, rg) => runOther(app, getCommand('pivot.recommended') ? 'pivot.recommended' : 'pivot.create', rg) },
    { label: 'Blank PivotTable', icon: 'table_view', run: (app, rg) => runOther(app, 'pivot.create', rg, { source: rangeRefText(app.sheet, rg), dest: 'new' }) },
  ],
  Sparklines: [
    { label: 'Line', icon: 'show_chart', run: (app, rg) => runOther(app, 'insert.sparkLine', rg) },
    { label: 'Column', icon: 'bar_chart', run: (app, rg) => runOther(app, 'insert.sparkColumn', rg) },
    { label: 'Win/Loss', icon: 'stacked_bar_chart', run: (app, rg) => runOther(app, 'insert.sparkWinLoss', rg) },
  ],
};

let qaTab = 'Formatting';

async function quickAnalysis(app: App, arg?: { tab?: string; action?: string }): Promise<void> {
  const rg = qaRange(app);
  if (!rg) { toast('Select a range with data for Quick Analysis.'); return; }
  if (arg?.tab && arg.action) {
    const a = QA_TABS[arg.tab]?.find((x) => x.label === arg.action);
    if (a) await a.run(app, rg);
    return;
  }
  const tabs = h('div', { class: 'ql-qa-tabs', role: 'tablist' });
  const tiles = h('div', { class: 'ql-qa-tiles', role: 'tabpanel' });
  let pick: QA | null = null;
  let close: (() => void) | null = null;
  const draw = () => {
    tabs.replaceChildren(...Object.keys(QA_TABS).map((t) => h('button', { class: 'ql-qa-tab' + (t === qaTab ? ' ql-on' : ''), role: 'tab', 'aria-selected': String(t === qaTab), type: 'button', onclick: () => { qaTab = t; draw(); } }, t)));
    tiles.replaceChildren(...QA_TABS[qaTab].map((a) => h('button', { class: 'ql-qa-tile', type: 'button', onclick: () => { pick = a; close?.(); } }, icon(a.icon), h('span', null, a.label))));
  };
  draw();
  const help = h('p', { class: 'ql-help' }, `${rangeName(rg)} · ${{ Formatting: 'Highlight interesting values with conditional formatting.', Charts: 'Recommended charts help you visualize the data.', Totals: 'Formulas that calculate totals for you.', Tables: 'Tables help you sort, filter and summarize data.', Sparklines: 'Mini charts placed in single cells.' }[qaTab]}`);
  const body = h('div', { class: 'ql-qa' }, tabs, tiles, help);
  tabs.addEventListener('click', () => { help.textContent = `${rangeName(rg)} · ${{ Formatting: 'Highlight interesting values with conditional formatting.', Charts: 'Recommended charts help you visualize the data.', Totals: 'Formulas that calculate totals for you.', Tables: 'Tables help you sort, filter and summarize data.', Sparklines: 'Mini charts placed in single cells.' }[qaTab]}`; });
  await openDialog({ title: 'Quick Analysis', content: body, ok: false, cancel: 'Close', onOpen: (c) => { close = () => c.close(null); } });
  if (pick) {
    if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
    await (pick as QA).run(app, rg);
  }
}

// ============================================================================================
// Prompt library

async function managePrompts(app: App): Promise<void> {
  for (;;) {
    const s = aiSettings();
    const r = await openDialog({
      title: 'Prompt Library', wide: true,
      fields: [
        { key: 'id', label: 'Saved prompts', type: 'list', options: s.prompts.map((p) => ({ value: p.id, label: p.name })), value: s.prompts[0]?.id },
        { key: 'help', type: 'info', html: '<p class="ql-help">Saved prompts appear under AI ▸ My Prompts and can go on the Quick Access Bar or a shortcut. They run on the selected cells, each row, or the sheet.</p>' },
      ],
      ok: 'Close', cancel: false,
      buttons: [
        { label: 'Run', primary: true, onClick: (v, c) => c.close({ act: 'run', id: v.id }) },
        { label: 'New…', onClick: (_v, c) => c.close({ act: 'new' }) },
        { label: 'Edit…', onClick: (v, c) => c.close({ act: 'edit', id: v.id }) },
        { label: 'Delete', onClick: (v, c) => c.close({ act: 'delete', id: v.id }) },
        { label: 'Add Examples', onClick: (_v, c) => c.close({ act: 'examples' }) },
      ],
    });
    if (!r || !r.act) return;
    if (r.act === 'run' && r.id) { await app.exec('ai.runPrompt', r.id); return; }
    if (r.act === 'delete' && r.id) { saveAiSettings((x) => { x.prompts = x.prompts.filter((p) => p.id !== r.id); }); continue; }
    if (r.act === 'examples') {
      saveAiSettings((x) => {
        const add = (name: string, prompt: string, input: 'selection' | 'document' | 'none', output: 'replace' | 'insert' | 'pane' | 'comment' | 'newdoc') => { if (!x.prompts.some((p) => p.name === name)) x.prompts.push({ id: `p${Date.now().toString(36)}${x.prompts.length}`, name, prompt, input, output }); };
        add('Genealogy: standardize place names', 'Rewrite the place name in the modern form "Town, County, State/Province, Country", keeping historical names in parentheses.', 'selection', 'replace');
        add('Expense category', 'Assign an expense category (Food, Travel, Lodging, Office, Software, Utilities, Other) to this row.', 'selection', 'insert');
        add('Weekly report from this sheet', 'Write a short weekly status report from this data: highlights, problems, next steps.', 'document', 'pane');
        add('Product descriptions', 'Write a one-sentence product description for this row.', 'selection', 'insert');
      });
      continue;
    }
    const cur = r.act === 'edit' ? s.prompts.find((p) => p.id === r.id) : undefined;
    if (r.act === 'edit' && !cur) continue;
    const e = await openDialog({
      title: cur ? 'Edit Prompt' : 'New Prompt', wide: true, fields: [
        { key: 'name', label: 'Name', type: 'text', value: cur?.name || '' },
        { key: 'prompt', label: 'Instruction', type: 'textarea', rows: 5, value: cur?.prompt || '', placeholder: 'e.g. Rewrite the address in USPS format' },
        { key: 'mode', label: 'Works on', type: 'select', half: true, value: cur ? (cur.output === 'replace' ? 'cells' : cur.output === 'insert' ? 'rows' : cur.input === 'document' ? 'sheet' : cur.input === 'none' ? 'none' : 'region') : 'cells', options: [
          { value: 'cells', label: 'Each selected cell (preview, then replace)' },
          { value: 'rows', label: 'Each row (results in a new column)' },
          { value: 'region', label: 'The selected data (answer in the AI pane)' },
          { value: 'sheet', label: 'The whole sheet (answer in the AI pane)' },
          { value: 'none', label: 'Nothing (just the instruction)' },
          { value: 'table', label: 'Make a table on a new sheet' },
        ] },
      ],
      validate: (v) => (!String(v.name).trim() || !String(v.prompt).trim() ? 'Enter a name and an instruction.' : null),
    });
    if (!e) continue;
    const map: Record<string, { input: 'selection' | 'document' | 'none'; output: 'replace' | 'insert' | 'pane' | 'comment' | 'newdoc' }> = {
      cells: { input: 'selection', output: 'replace' }, rows: { input: 'selection', output: 'insert' }, region: { input: 'selection', output: 'pane' },
      sheet: { input: 'document', output: 'pane' }, none: { input: 'none', output: 'pane' }, table: { input: 'none', output: 'newdoc' },
    };
    saveAiSettings((x) => {
      const p = { id: cur?.id || `p${Date.now().toString(36)}`, name: String(e.name).trim(), prompt: String(e.prompt).trim(), ...map[e.mode] };
      x.prompts = cur ? x.prompts.map((q) => (q.id === cur.id ? p : q)) : [...x.prompts, p];
    });
  }
}

async function runSavedPrompt(app: App, id?: string): Promise<void> {
  const s = aiSettings();
  if (!s.prompts.length) { await managePrompts(app); return; }
  let p = s.prompts.find((x) => x.id === id);
  if (!p) {
    const r = await openDialog({ title: 'Run Prompt', fields: [{ key: 'p', label: 'Prompt', type: 'list', rows: 8, options: s.prompts.map((x) => ({ value: x.id, label: x.name })), value: s.prompts[0].id }] });
    if (!r?.p) return;
    p = s.prompts.find((x) => x.id === r.p);
  }
  if (!p) return;
  const sp = p;
  const f: AiFeature = {
    id: 'ai.custom', label: sp.name, icon: 'play_circle', group: 'Assistant', tip: '', tier: 'quality', temperature: 0.4,
    input: sp.input === 'document' ? 'sheet' : sp.input === 'none' ? 'none' : 'region',
    output: sp.output === 'newdoc' ? 'table' : 'pane',
    prompt: ({ text }) => `${sp.prompt}${sp.output === 'newdoc' ? '\nRespond with JSON only: an array of objects with the same keys (the keys become headers).' : ''}${text ? `\n\n${text}` : ''}`,
  };
  if (sp.output === 'replace' || sp.output === 'insert') {
    if (!(await ensureAiReady(app, true))) return;
    if (!app.sheet.usedRange(false)) { toast('The sheet is empty.'); return; }
    await runRows(app, f, undefined, { instruction: sp.prompt, mode: sp.output === 'replace' ? 'replace' : 'column', header: sp.name.slice(0, 30) });
    return;
  }
  await runFeature(app, f);
}

// ============================================================================================
// Registration

const featureCommands = FEATURES.map((f) => ({
  id: f.id, label: f.label, icon: f.icon, tab: 'AI', group: f.group, tip: f.tip, keys: f.keys, popular: f.popular, origin: f.origin, keywords: `ai ${f.keywords || ''}`,
  kind: (f.options ? 'menu' : 'button') as 'menu' | 'button',
  edits: !['pane'].includes(f.output),
  run: (app: App, arg?: any) => runFeature(app, f, typeof arg === 'string' ? arg : undefined),
  menu: f.options ? (): MenuEntry[] => f.options!.map((o) => ({ id: f.id, arg: o.value, label: o.label })) : undefined,
}));

def(featureCommands);

def([
  {
    id: 'ai.quickAnalysis', label: 'Quick Analysis', icon: 'bolt', tab: 'Home', group: 'Analysis', keys: ['Mod-q'], popular: true, edits: true,
    tip: 'Formatting, charts, totals, tables and sparklines for the selected data, in one place (no AI needed)', keywords: 'quick analysis totals data bars sparkline',
    run: (app, arg?: { tab?: string; action?: string }) => quickAnalysis(app, arg),
  },
  {
    id: 'ai.chat', label: 'AI Assistant', icon: 'chat_bubble', tab: 'AI', group: 'Assistant', kind: 'toggle', popular: true, keys: ['Mod-Alt-a'], noRecord: true, edits: false,
    tip: 'Chat about this workbook; answers can be put into the active cell or inserted as a table with one tap', keywords: 'chat assistant copilot ask',
    run: (app) => app.ui.togglePane('ai'), isActive: (app) => !!app.ui?.paneOpen?.('ai'),
  },
  { id: 'ai.prompts', label: 'Prompt Library', icon: 'library_books', tab: 'AI', group: 'Assistant', noRecord: true, edits: false, tip: 'Save your own AI instructions and run them on cells, rows or the sheet', run: (app) => managePrompts(app) },
  {
    id: 'ai.runPrompt', label: 'My Prompts', icon: 'play_circle', tab: 'AI', group: 'Assistant', kind: 'menu', edits: true, tip: 'Run one of your saved AI prompts',
    run: (app, id?: string) => runSavedPrompt(app, typeof id === 'string' ? id : undefined),
    menu: () => { const s = aiSettings(); return [...s.prompts.map((p) => ({ id: 'ai.runPrompt', arg: p.id, label: p.name })), ...(s.prompts.length ? [{ separator: true }] : []), { id: 'ai.prompts', label: 'Manage Prompts…' }]; },
  },
  {
    id: 'ai.setup', label: 'Set Up AI', icon: 'auto_awesome', tab: 'AI', group: 'Setup', popular: true, noRecord: true, edits: false,
    tip: 'Choose a free or paid AI: on this device, free cloud tiers, your own computer, or a paid key. Opens the right page in your browser and picks up the key.',
    keywords: 'api key openai anthropic gemini mistral groq openrouter ollama wizard',
    run: async (app) => { const { openAiSetup } = await import('../ui/aiWizard'); await openAiSetup(app); },
  },
  { id: 'ai.providers', label: 'Providers & Models', icon: 'tune', tab: 'AI', group: 'Setup', noRecord: true, edits: false, tip: 'Switch AI provider, choose models, test or remove keys', run: async (app) => { const { openAiProviders } = await import('../ui/aiWizard'); await openAiProviders(app); } },
  {
    id: 'ai.chooseProvider', label: 'Active AI', icon: 'swap_horiz', tab: 'AI', group: 'Setup', kind: 'menu', noRecord: true, edits: false, tip: 'Switch between the AI providers you have set up',
    run: async (app, id?: string) => {
      if (!id) { const { openAiProviders } = await import('../ui/aiWizard'); await openAiProviders(app); return; }
      if (!aiSettings().providers[id]) return;
      saveAiSettings((s) => { s.active = id; });
      toast(`AI: ${providerById(id)?.name}`);
    },
    menu: () => { const s = aiSettings(); const items: MenuEntry[] = Object.keys(s.providers).map((id) => ({ id: 'ai.chooseProvider', arg: id, label: `${providerById(id)?.name || id} · ${s.providers[id].model}`, checked: s.active === id })); return [...items, ...(items.length ? [{ separator: true }] : []), { id: 'ai.setup', label: 'Add a Provider…' }]; },
  },
  { id: 'ai.settings', label: 'AI Settings', icon: 'settings_suggest', tab: 'AI', group: 'Setup', noRecord: true, edits: false, tip: 'Turn AI on or off, privacy prompts, language for generated text, usage', run: async (app) => { const { openAiSettings } = await import('../ui/aiWizard'); await openAiSettings(app); } },
  { id: 'ai.stop', label: 'Stop AI', icon: 'stop_circle', tab: 'AI', group: 'Assistant', noRecord: true, edits: false, tip: 'Stop the AI request in progress', run: () => stopAi() },
  {
    id: 'ai.privacy', label: 'AI & Privacy', icon: 'shield', tab: 'AI', group: 'Setup', noRecord: true, edits: false, tip: 'Where your data goes when you use AI',
    run: async () => {
      const s = aiSettings();
      const p = providerById(s.active || '');
      const esc = (t: string) => t.replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]!));
      await openDialog({ title: 'AI & Privacy', ok: 'Close', cancel: false, fields: [{ key: 'i', type: 'info', html: [
        '<p><b>QuillLattice has no server.</b> AI requests go directly from this device to the provider you choose, using your own key. Nothing is sent until you run an AI command or ask in the AI pane.</p>',
        `<p><b>Active provider:</b> ${p ? esc(p.name) : 'none'}${p && (p.kind === 'local' || p.kind === 'builtin') ? ' — runs on this device; your data does not leave it.' : p ? ' — the cells a command needs (the selection, the data around it, or the sheet when you ask) are sent to this service; its terms apply.' : ''}</p>`,
        '<p><b>Keys</b> are stored on this device only (the Android Keystore-encrypted store in the app, or this browser\'s IndexedDB on the web). They are never placed in workbooks, drafts or exported settings.</p>',
        '<p><b>Changes are previewed.</b> AI output that changes cells is shown first and applied as one step you can undo.</p>',
        '<p><b>Free tiers</b> of some services may use prompts to improve their models. Use the on-device model for confidential data.</p>',
      ].join('') }] });
    },
  },
  {
    id: 'ai.clearLocal', label: 'Delete Downloaded Models', icon: 'delete', tab: 'AI', group: 'Setup', noRecord: true, edits: false, tip: 'Free the storage used by the on-device model',
    run: async () => {
      const { clearLocalModels, localStorageEstimate } = await import('../ai/local');
      const before = await localStorageEstimate();
      if (!(await confirmDialog('Delete Downloaded Models', `Delete the on-device AI model files? ${before}`, 'Delete'))) return;
      const n = await clearLocalModels();
      toast(n ? 'Downloaded model files deleted' : 'No downloaded model files were found');
    },
  },
]);

void (null as unknown as Scalar);
