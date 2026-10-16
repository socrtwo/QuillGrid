/**
 * AI pane (chat with workbook context, results with "insert into cell" actions), the busy
 * indicator with progress and Stop, and the preview dialogs used before AI output changes cells:
 * formulas (validated, with the result), cell changes (before/after with per-row choice) and tables.
 */
import { h, icon, clear } from './dom';
import { openDialog, toast, confirmDialog } from './dialog';
import { registerPane, type PaneDef } from './panes';
import type { App } from '../app';
import type { Sheet } from '../model/workbook';
import type { Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellName, rangeName } from '../model/address';
import { formatValue } from '../model/numfmt';
import { chat, AiError, type ChatMessage } from '../ai/client';
import { aiSettings } from '../ai/store';
import { providerById } from '../ai/providers';
import { cleanOutput, extractFormula, parseTable, tableText, validateFormula, writeTable, type FormulaCheck } from '../ai/apply';
import { BASE_SYSTEM } from '../ai/features';
import { dataRegion, hasHeaderRow, clipToUsed, setInput } from '../commands/ops';

// ------------------------------------------------------------------------------------------
// Busy indicator with progress and Stop

export interface Busy { signal: AbortSignal; abort(): void; label(t: string): void; progress(done: number, total: number): void }

let busyEl: HTMLElement | null = null;
let current: (Busy & { ctl: AbortController }) | null = null;

export function aiBusy(label: string): Busy {
  current?.abort();
  const ctl = new AbortController();
  busyEl?.remove();
  const text = h('span', { class: 'ql-ai-busy-label' }, label);
  const bar = h('progress', { class: 'ql-ai-progress', max: '100' }) as HTMLProgressElement;
  bar.removeAttribute('value');
  busyEl = h('div', { class: 'ql-ai-busy', role: 'status', 'aria-live': 'polite' }, h('span', { class: 'ql-spin' }), text, bar, h('button', { class: 'ql-btn ql-sm', type: 'button', onclick: () => ctl.abort() }, 'Stop'));
  document.body.appendChild(busyEl);
  const b = {
    ctl, signal: ctl.signal,
    abort: () => ctl.abort(),
    label: (t: string) => { text.textContent = t; },
    progress: (done: number, total: number) => { bar.max = Math.max(1, total); bar.value = done; },
  };
  current = b;
  ctl.signal.addEventListener('abort', () => aiIdle(b));
  return b;
}

export function aiIdle(b?: Busy): void {
  if (b && b !== current) return;
  busyEl?.remove();
  busyEl = null;
  current = null;
}

export function stopAi(): void { current?.abort(); }

export function aiModelName(): string {
  const s = aiSettings();
  const id = s.active || '';
  return s.providers[id]?.model || providerById(id)?.name || 'AI';
}

export const fmtScalar = (v: Scalar, numFmt?: string) => (v instanceof CellError ? v.code : v == null ? '' : formatValue(v, numFmt).text);

// ------------------------------------------------------------------------------------------
// Inserting into cells

/** Puts AI text into the active cell: a formula when the text holds one, else the text itself. */
export function insertIntoCell(app: App, text: string): void {
  const s = app.sheet;
  const { r, c } = s.sel;
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
  if (!app.exec('cell.canEdit', { r, c })) return;
  const f = extractFormula(text);
  const value = f && (text.trim().startsWith('=') || /```/.test(text) || text.trim().split('\n').length <= 3) ? f.formula : cleanOutput(text).replace(/\n+/g, ' ').slice(0, 32000);
  app.book.transact('AI: Insert', () => setInput(app, s, r, c, value));
}

/** Writes a table at the active cell after checking for existing data. */
export async function insertTableAtCursor(app: App, rows: Scalar[][], title = 'AI: Insert Table'): Promise<boolean> {
  const s = app.sheet;
  const { r, c } = s.sel;
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return false; }
  const width = Math.max(...rows.map((x) => x.length));
  const rg = { r1: r, c1: c, r2: r + rows.length - 1, c2: c + width - 1 };
  if (s.protection.on) { app.exec('review.protectionNotice'); return false; }
  let busy = false;
  for (let i = rg.r1; i <= rg.r2 && !busy; i++) for (let j = rg.c1; j <= rg.c2; j++) { const x = s.get(i, j); if (x && (x.v != null || x.f != null)) { busy = true; break; } }
  if (busy && !(await confirmDialog('Replace Contents?', `There's already data in ${rangeName(rg)}. Replace it?`, 'Replace'))) return false;
  app.book.transact(title, () => writeTable(app, s, r, c, rows));
  app.grid.select(r, c, { range: rg });
  return true;
}

// ------------------------------------------------------------------------------------------
// Formula preview

export interface FormulaPreview {
  title: string;
  sheet: Sheet;
  r: number; c: number;
  /** Streams the raw answer; resolves with the final raw text. */
  generate: (onToken: (full: string) => void, signal: AbortSignal) => Promise<string>;
  original?: string;
}

/** Shows a generated formula, validated against the parser and engine, with its result; Insert writes it into the target cell. */
export async function showFormulaPreview(app: App, o: FormulaPreview): Promise<void> {
  const ctl = new AbortController();
  const tab = app.tab;
  const stream = h('pre', { class: 'ql-ai-out' }, '…');
  const fIn = h('textarea', { class: 'ql-ai-formula', rows: 3, spellcheck: 'false', 'aria-label': 'Formula' }) as HTMLTextAreaElement;
  const status = h('div', { class: 'ql-ai-status' });
  const result = h('div', { class: 'ql-ai-result' });
  const expl = h('div', { class: 'ql-ai-expl ql-help' });
  const form = h('div', { class: 'ql-ai-formula-wrap', style: 'display:none' }, h('label', { class: 'ql-label' }, `Formula for ${cellName(o.r, o.c)}`), fIn, status, result, expl);
  let check: FormulaCheck | null = null;
  let done = false;
  const recheck = () => {
    const f = fIn.value.trim();
    check = f ? validateFormula(tab.engine, o.sheet, o.r, o.c, f.startsWith('=') ? f : '=' + f) : null;
    clear(status); clear(result);
    if (!check) return;
    if (!check.ok) { status.append(icon('error'), ` ${check.error}`); status.className = 'ql-ai-status ql-bad'; return; }
    status.className = 'ql-ai-status ' + (check.resultError ? 'ql-warn' : 'ql-good');
    status.append(icon(check.resultError ? 'warning' : 'check_circle'), check.resultError ? ` The formula is valid but currently returns ${check.resultError}.` : ' The formula is valid.');
    const m = check.matrix || [[check.value ?? null]];
    const fmt = o.sheet.styleOf(o.r, o.c).numFmt;
    if (m.length === 1 && m[0].length === 1) result.append(h('span', { class: 'ql-label' }, 'Result: '), h('b', null, fmtScalar(m[0][0], fmt) || '(empty)'));
    else {
      const t = h('table', { class: 'ql-stats' });
      m.slice(0, 6).forEach((row) => t.appendChild(h('tr', null, ...row.slice(0, 6).map((v) => h('td', null, fmtScalar(v, fmt))), row.length > 6 ? h('td', null, '…') : null)));
      result.append(h('div', { class: 'ql-label' }, `Result (spills ${m.length} × ${m[0].length}):`), t);
    }
  };
  fIn.addEventListener('input', recheck);
  const body = h('div', { class: 'ql-ai-preview' },
    o.original ? h('details', null, h('summary', null, 'Current formula'), h('code', null, o.original)) : null,
    stream, form,
    h('p', { class: 'ql-help ql-ai-disclaimer' }, 'AI can be wrong. Check the result against a few rows before relying on it.'));
  const run = async () => {
    done = false; stream.style.display = ''; form.style.display = 'none'; stream.textContent = '…';
    try {
      const raw = await o.generate((full) => { stream.textContent = full; }, ctl.signal);
      const f = extractFormula(raw);
      stream.style.display = 'none'; form.style.display = '';
      fIn.value = f?.formula || '';
      expl.textContent = f?.explanation || (f ? '' : cleanOutput(raw));
      if (!f) { clear(status); status.className = 'ql-ai-status ql-bad'; status.append(icon('error'), ' No formula was found in the answer.'); }
      else recheck();
    } catch (e: any) {
      stream.textContent = e instanceof AiError && e.code === 'aborted' ? 'Stopped.' : `⚠ ${e?.message || e}`;
    }
    done = true;
  };
  void run();
  await openDialog({
    title: o.title, content: body, wide: true, ok: false, cancel: 'Discard',
    buttons: [
      { label: 'Insert', primary: true, onClick: (_v, c) => {
        if (!done) return;
        const f = fIn.value.trim();
        if (!f) return;
        if (check && !check.ok) { toast(`Fix the formula first: ${check.error}`); return; }
        if (!app.docs.includes(tab)) { c.close(null); return; }
        if (tab !== app.tab) { toast('The workbook changed; the formula was copied instead.'); void navigator.clipboard?.writeText(f); c.close(null); return; }
        if (!app.exec('cell.canEdit', { r: o.r, c: o.c })) return;
        c.close(null);
        if (app.sheet !== o.sheet) { const i = app.book.sheets.indexOf(o.sheet); if (i >= 0) app.exec('sheet.activate', i); }
        app.grid.select(o.r, o.c);
        app.exec('cell.enter', { r: o.r, c: o.c, text: f.startsWith('=') ? f : '=' + f });
      } },
      { label: 'Copy', onClick: async () => { try { await navigator.clipboard.writeText(fIn.value); toast('Copied'); } catch { fIn.select(); toast('Press Ctrl+C to copy'); } } },
      { label: 'Retry', onClick: () => { if (done) void run(); } },
    ],
  });
  ctl.abort();
}

// ------------------------------------------------------------------------------------------
// Cell changes preview

export interface ChangeRow { label: string; before: string; after: string; changed: boolean }

/** Before/after table with a check box per row. Resolves with the indexes to apply, or null. */
export async function showChangesPreview(title: string, rows: ChangeRow[], opts: { note?: string; applyLabel?: string; columns?: [string, string] } = {}): Promise<number[] | null> {
  const changed = rows.map((r, i) => (r.changed ? i : -1)).filter((i) => i >= 0);
  if (!changed.length) { toast('No changes were suggested.'); return null; }
  const LIMIT = 1000;
  const boxes: HTMLInputElement[] = [];
  const table = h('table', { class: 'ql-stats ql-ai-diff' }, h('tr', null, h('th', null, h('input', { type: 'checkbox', checked: true, 'aria-label': 'Select all', onchange: (e: Event) => boxes.forEach((b) => (b.checked = (e.target as HTMLInputElement).checked)) })), h('th', null, 'Cell'), h('th', null, opts.columns?.[0] || 'Before'), h('th', null, opts.columns?.[1] || 'After')));
  for (const i of changed.slice(0, LIMIT)) {
    const r = rows[i];
    const cb = h('input', { type: 'checkbox', checked: true, 'aria-label': `Apply to ${r.label}` }) as HTMLInputElement;
    (cb as any)._i = i;
    boxes.push(cb);
    table.appendChild(h('tr', null, h('td', null, cb), h('td', null, r.label), h('td', { class: 'ql-ai-before' }, r.before), h('td', { class: 'ql-ai-after' }, r.after)));
  }
  const body = h('div', { class: 'ql-ai-preview' },
    h('p', { class: 'ql-help' }, `${changed.length} of ${rows.length} cell${rows.length === 1 ? '' : 's'} would change.${changed.length > LIMIT ? ` The first ${LIMIT} are listed; the rest are applied too.` : ''}${opts.note ? ' ' + opts.note : ''}`),
    h('div', { class: 'ql-ai-diff-wrap' }, table),
    h('p', { class: 'ql-help ql-ai-disclaimer' }, 'AI can be wrong. Untick anything you do not want; the changes are applied as one step you can undo.'));
  const r = await openDialog({ title, content: body, wide: true, ok: opts.applyLabel || 'Apply', cancel: 'Discard' });
  if (!r) return null;
  const picked = boxes.filter((b) => b.checked).map((b) => (b as any)._i as number);
  return [...picked, ...changed.slice(LIMIT)];
}

/** Table preview: Insert at the active cell, or on a new sheet. */
export async function showTablePreview(title: string, rows: Scalar[][]): Promise<'here' | 'sheet' | null> {
  const t = h('table', { class: 'ql-stats' });
  rows.slice(0, 60).forEach((row, i) => t.appendChild(h('tr', null, ...row.map((v) => h(i === 0 ? 'th' : 'td', null, v instanceof CellError ? v.code : v == null ? '' : String(v))))));
  const body = h('div', { class: 'ql-ai-preview' },
    h('p', { class: 'ql-help' }, `${rows.length - 1} row${rows.length === 2 ? '' : 's'} × ${Math.max(...rows.map((x) => x.length))} columns${rows.length > 60 ? ' (first 60 rows shown)' : ''}.`),
    h('div', { class: 'ql-ai-diff-wrap' }, t),
    h('p', { class: 'ql-help ql-ai-disclaimer' }, 'AI can be wrong; facts marked "(verify)" need checking.'));
  let choice: 'here' | 'sheet' | null = null;
  await openDialog({ title, content: body, wide: true, ok: false, cancel: 'Discard', buttons: [
    { label: 'New Sheet', primary: true, onClick: (_v, c) => { choice = 'sheet'; c.close(null); } },
    { label: 'At Active Cell', onClick: (_v, c) => { choice = 'here'; c.close(null); } },
    { label: 'Copy', onClick: async () => { try { await navigator.clipboard.writeText(rows.map((r) => r.map((v) => (v == null ? '' : String(v))).join('\t')).join('\n')); toast('Copied'); } catch { toast('Clipboard is not available'); } } },
  ] });
  return choice;
}

// ------------------------------------------------------------------------------------------
// AI pane

export interface PaneAction { label: string; icon?: string; run: () => void }
interface Turn { role: 'user' | 'assistant'; text: string; title?: string; actions?: PaneAction[]; error?: boolean }
const history: Turn[] = [];
let paneEl: HTMLElement | null = null;
let paneApp: App | null = null;

/** Shows a (streaming) result in the AI pane. */
export function paneResult(app: App, title: string): { update: (t: string) => void; done: (t: string, actions?: PaneAction[]) => void; fail: (msg: string) => void } {
  if (!app.ui.paneOpen('ai')) app.ui.openPane('ai');
  const turn: Turn = { role: 'assistant', text: '', title };
  history.push(turn);
  drawPane();
  return {
    update: (t) => { turn.text = t; drawTurnText(turn); },
    done: (t, actions) => { turn.text = t; turn.actions = actions; drawPane(); },
    fail: (msg) => { turn.text = `⚠ ${msg}`; turn.error = true; drawPane(); },
  };
}

function drawTurnText(turn: Turn): void {
  const el = paneEl?.querySelector(`[data-turn="${history.indexOf(turn)}"] .ql-ai-msg-text`);
  if (el) { el.textContent = turn.text || '…'; const log = paneEl?.querySelector('.ql-ai-log'); if (log) log.scrollTop = log.scrollHeight; }
}

function turnActions(app: App, t: Turn): HTMLElement {
  const row = h('div', { class: 'ql-row ql-ai-msg-actions' });
  for (const a of t.actions || []) row.appendChild(h('button', { class: 'ql-btn ql-sm', type: 'button', onclick: a.run }, a.icon ? icon(a.icon) : null, ` ${a.label}`));
  if (!t.error) {
    const f = extractFormula(t.text);
    row.appendChild(h('button', { class: 'ql-btn ql-sm', type: 'button', title: f ? `Insert ${f.formula} into the active cell` : 'Insert this text into the active cell', onclick: () => insertIntoCell(app, t.text) }, icon(f ? 'function' : 'input'), f ? ' Insert Formula' : ' Insert in Cell'));
    const tbl = parseTable(t.text);
    if (tbl && tbl.length >= 2 && Math.max(...tbl.map((r) => r.length)) >= 2) row.appendChild(h('button', { class: 'ql-btn ql-sm', type: 'button', title: 'Insert the table at the active cell', onclick: () => void insertTableAtCursor(app, tbl) }, icon('table'), ' Insert Table'));
    row.appendChild(h('button', { class: 'ql-btn ql-sm', type: 'button', 'aria-label': 'Copy', title: 'Copy', onclick: async () => { try { await navigator.clipboard.writeText(t.text); toast('Copied'); } catch { toast('Clipboard is not available'); } } }, icon('content_copy')));
  }
  return row;
}

function drawPane(): void {
  if (!paneEl || !paneApp) return;
  const app = paneApp;
  const log = paneEl.querySelector('.ql-ai-log') as HTMLElement | null;
  if (!log) return;
  clear(log);
  if (!history.length) {
    log.append(h('p', { class: 'ql-help' }, aiSettings().active ? 'Ask about this workbook, or ask for a formula, a table or an explanation. Answers have buttons to put them into the active cell.' : 'AI is not set up yet. Nothing is sent anywhere until you choose a provider and ask something.'));
    if (!aiSettings().active) log.append(h('button', { class: 'ql-btn ql-primary', type: 'button', onclick: () => app.exec('ai.setup') }, 'Set Up AI…'));
    for (const q of ['Write a formula that totals column C where column A is "East"', 'What trends do you see in this data?', 'Explain the formula in the active cell']) {
      log.append(h('button', { class: 'ql-ai-suggest', type: 'button', onclick: () => { const ta = paneEl?.querySelector('textarea'); if (ta) { (ta as HTMLTextAreaElement).value = q; (ta as HTMLTextAreaElement).focus(); } } }, q));
    }
  }
  history.forEach((t, i) => {
    log.append(h('div', { class: `ql-ai-msg ql-ai-${t.role}`, 'data-turn': String(i) },
      t.title ? h('div', { class: 'ql-ai-msg-title' }, t.title) : null,
      h('div', { class: 'ql-ai-msg-text' }, t.text || '…'),
      t.role === 'assistant' && t.text ? turnActions(app, t) : null));
  });
  log.scrollTop = log.scrollHeight;
}

type Include = 'region' | 'sheet' | 'none';

function contextFor(app: App, include: Include): string {
  if (include === 'none') return '';
  const s = app.sheet;
  const book = app.book;
  const head = `Workbook "${app.tab.displayName}" with sheets: ${book.sheets.map((x) => x.name).join(', ')}. Active sheet "${s.name}", active cell ${cellName(s.sel.r, s.sel.c)}${s.get(s.sel.r, s.sel.c)?.f ? ` (formula =${s.get(s.sel.r, s.sel.c)!.f})` : ''}, selection ${s.sel.ranges.map(rangeName).join(',')}.`;
  let rg = include === 'sheet' ? s.usedRange(false) : dataRegion(app);
  if (!rg) return head + ' The sheet is empty.';
  rg = clipToUsed(s, rg);
  const t = tableText(s, rg, { header: hasHeaderRow(s, rg), maxChars: include === 'sheet' ? 30000 : 12000, formulas: true });
  return `${head}\n${include === 'sheet' ? 'The used range of the sheet' : 'The data around the selection'}:\n${t.text}`;
}

async function ask(app: App, question: string, include: Include): Promise<void> {
  const { ensureAiReady } = await import('../commands/ai');
  if (!(await ensureAiReady(app, include !== 'none'))) return;
  history.push({ role: 'user', text: question });
  const turn: Turn = { role: 'assistant', text: '' };
  history.push(turn);
  drawPane();
  const msgs: ChatMessage[] = [{ role: 'system', content: `${BASE_SYSTEM} You are chatting with the user in a side panel of the spreadsheet. Markdown is fine. When you give a formula, put it alone on its own line starting with "=". When you give a table, use a Markdown table.` }];
  const ctx = contextFor(app, include);
  if (ctx) msgs.push({ role: 'user', content: ctx }, { role: 'assistant', content: 'Understood.' });
  for (const t of history.slice(0, -2).slice(-10)) if (!t.error) msgs.push({ role: t.role, content: t.text });
  msgs.push({ role: 'user', content: question });
  const b = aiBusy('AI is answering…');
  try {
    const r = await chat({ messages: msgs, signal: b.signal, onToken: (_d, full) => { turn.text = full; drawTurnText(turn); } });
    turn.text = cleanOutput(r.text);
  } catch (e: any) {
    turn.error = !(e instanceof AiError && e.code === 'aborted');
    turn.text = turn.error ? `⚠ ${e?.message || e}` : turn.text || '(stopped)';
  } finally { aiIdle(b); }
  drawPane();
}

export const aiPaneDef: PaneDef = {
  id: 'ai',
  title: 'AI Assistant',
  render(app, el) {
    paneEl = el;
    paneApp = app;
    el.classList.add('ql-ai-pane');
    const input = h('textarea', { class: 'ql-input', rows: 3, placeholder: 'Ask about this workbook, or ask for a formula or a table…', 'aria-label': 'Message' }) as HTMLTextAreaElement;
    const include = h('select', { class: 'ql-input ql-sm', 'aria-label': 'Context to send' },
      h('option', { value: 'region' }, 'Send the data around the selection'), h('option', { value: 'sheet' }, 'Send the whole sheet'), h('option', { value: 'none' }, 'Send no workbook data')) as HTMLSelectElement;
    const send = () => { const q = input.value.trim(); if (!q) return; input.value = ''; void ask(app, q, include.value as Include); };
    input.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); } });
    const s = aiSettings();
    const prov = s.active ? providerById(s.active) : null;
    el.append(
      h('div', { class: 'ql-ai-head ql-help' }, prov ? `${prov.name} · ${s.providers[s.active!]?.model || ''}` : 'No AI provider', ' ', h('button', { class: 'ql-link', type: 'button', onclick: () => app.exec('ai.providers') }, 'Change')),
      h('div', { class: 'ql-ai-log' }),
      h('div', { class: 'ql-ai-compose' }, input, h('div', { class: 'ql-row' }, include, h('span', { class: 'ql-spacer' }),
        h('button', { class: 'ql-iconbtn', type: 'button', title: 'Clear conversation', 'aria-label': 'Clear conversation', onclick: () => { history.length = 0; drawPane(); } }, icon('delete_sweep')),
        h('button', { class: 'ql-btn ql-primary', type: 'button', onclick: send }, icon('send'), ' Send'))));
    drawPane();
  },
  onClose() { paneEl = null; },
};

registerPane(aiPaneDef);

/** Show a long result in a read-only dialog. */
export async function showAiText(title: string, text: string): Promise<void> {
  await openDialog({ title, content: h('pre', { class: 'ql-ai-out' }, text), wide: true, ok: 'Close', cancel: false });
}
