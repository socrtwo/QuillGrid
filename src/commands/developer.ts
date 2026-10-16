/**
 * Developer tab: recorded macros, JavaScript macro modules with a small object model, macro
 * security, form controls (buttons, check boxes, lists, spinners…) and JSON import/export.
 *
 * Recorded macros are lists of command steps (prefs.macros) replayed with app.exec. JavaScript
 * macros live in the workbook (book.modules) and run with `new Function('api', code)` after the
 * macro security setting allows it. The API object (createMacroApi) works on any Workbook, so it is
 * tested without the UI.
 */
import { def, getCommand, allCommands, type MenuEntry } from './registry';
import type { App } from '../app';
import type { Workbook, Sheet } from '../model/workbook';
import type { Cell, Scalar, DrawingObject, FormControl, FormControlType } from '../model/types';
import { CellError, errorFromCode } from '../model/types';
import { cellName, rangeName, parseRange, type Range } from '../model/address';
import { resolveRangeText, rangeRefText } from '../model/structure';
import { jsDateToSerial } from '../model/dates';
import { newId } from '../model/workbook';
import type { Macro } from '../prefs';
import { openDialog, toast, confirmDialog, promptDialog, alertDialog, showMenu } from '../ui/dialog';
import { h, icon, clear, download, pickFile } from '../ui/dom';
import { activeCell, dataRegion, hasHeaderRow, cellText, parseInput } from './ops';
import { objectHooks } from '../grid/objects';
import type { Grid } from '../grid/grid';
import { canChange, checkWritable, writeOut } from './tools';
import { isDateFormat } from '../model/numfmt';

// ============================================================================================
// Settings kept on this device

const SEC_KEY = 'quilllattice.macroSecurity';
const WARN_KEY = 'quilllattice.macroWarned';
const REL_KEY = 'quilllattice.macroRelative';

export type MacroSecurity = 'disable' | 'ask' | 'enable';

const lsGet = (k: string): string | null => { try { return localStorage.getItem(k); } catch { return null; } };
const lsSet = (k: string, v: string): void => { try { localStorage.setItem(k, v); } catch { /* unavailable */ } };

export function macroSecurity(): MacroSecurity {
  const v = lsGet(SEC_KEY);
  return v === 'disable' || v === 'enable' ? v : 'ask';
}

let relativeRefs = lsGet(REL_KEY) === '1';
let designMode = false;
/** Workbooks whose macros the user allowed for this session (security "ask"). */
const trusted = new WeakSet<Workbook>();

/** Asks once per workbook (setting "ask") and warns once ever that macros run code. */
async function allowMacros(app: App): Promise<boolean> {
  const sec = macroSecurity();
  if (sec === 'disable') { toast('Macros are disabled (Developer ▸ Macro Security).', { action: 'Change', onAction: () => app.exec('dev.macroSecurity') }); return false; }
  if (!lsGet(WARN_KEY)) {
    const ok = await confirmDialog('Macros Can Run Code', 'JavaScript macros can read and change everything in your workbooks and can use the network like a web page. Run macros only from sources you trust.', 'I Understand', 'Cancel');
    if (!ok) return false;
    lsSet(WARN_KEY, '1');
  }
  if (sec === 'enable' || trusted.has(app.book)) return true;
  const ok = await confirmDialog('Enable Macros?', `“${app.tab.displayName}” contains macros. Enable them for this session?`, 'Enable Macros', 'Cancel');
  if (ok) trusted.add(app.book);
  return ok;
}

// ============================================================================================
// Macro API

export interface MacroHost {
  book: Workbook;
  /** Active sheet. */
  readonly sheet: Sheet;
  exec(id: string, arg?: any): unknown;
  select?(sheet: Sheet, r: number, c: number, range: Range): void;
  msgBox?(text: string): Promise<void> | void;
  inputBox?(prompt: string, value?: string): Promise<string | null> | string | null;
  log?(...args: any[]): void;
  recalc?(): void;
  /** Called before every write or command when stepping through a macro (trace). */
  trace?(what: string): void;
}

function toScalar(v: any): Scalar {
  if (v == null) return null;
  if (v instanceof Date) return jsDateToSerial(v);
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v === 'boolean' || typeof v === 'string') return v;
  if (v instanceof CellError) return v;
  return String(v);
}

function outValue(v: Scalar): any { return v instanceof CellError ? v.code : v; }

/**
 * The object passed to JavaScript macros as `api`:
 *   api.cell('B2').value / .formula / .format / .text      get and set
 *   api.range('A1:C5').values (2-D array, get and set), .formulas, .clear(), .address
 *   api.sheet('Data').cell('A1') … ; api.sheets; api.activeSheet; api.addSheet(name)
 *   api.select('A1:B3'); api.selection; api.activeCell
 *   api.exec('fmt.bold'); api.msgBox('Done'); await api.inputBox('Name?'); api.log(...)
 */
export function createMacroApi(host: MacroHost) {
  const book = host.book;
  const fresh = () => { if (book.pending.cells.size || book.pending.structural) host.recalc?.(); };
  const write = (label: string, fn: () => void) => { host.trace?.(label); book.transact('Macro', fn); };
  const locate = (addr: string, sheet?: Sheet): { sheet: Sheet; range: Range } => {
    const home = sheet || host.sheet;
    const res = resolveRangeText(book, String(addr), home) || (parseRange(String(addr)) ? { sheet: home, range: parseRange(String(addr))! } : null);
    if (!res) throw new Error(`Not a valid cell or range: ${addr}`);
    return res;
  };
  const guard = (s: Sheet, r: number, c: number) => { if (!canChange(s, r, c)) throw new Error(`Cell ${cellName(r, c)} on "${s.name}" is locked (the sheet is protected).`); };
  const setValue = (s: Sheet, r: number, c: number, raw: any) => {
    guard(s, r, c);
    const cur = s.get(r, c);
    if (typeof raw === 'string' && raw.startsWith('=') && raw.length > 1) { s.set(r, c, { ...(cur?.s ? { s: cur.s } : {}), f: raw.slice(1) }); return; }
    const v = toScalar(raw);
    const next: Cell = { ...(cur?.s ? { s: cur.s } : {}) };
    if (v != null && v !== '') next.v = v;
    if (raw instanceof Date) next.s = book.derivedStyle(cur?.s ?? 0, { numFmt: 'yyyy-mm-dd' });
    s.set(r, c, next.v == null && !next.s ? undefined : next);
  };

  const cellObj = (addr: string, sheet?: Sheet) => {
    const { sheet: s, range } = locate(addr, sheet);
    const r = range.r1, c = range.c1;
    return {
      get address() { return cellName(r, c); },
      get sheet() { return s.name; },
      get value(): any { fresh(); return outValue(s.value(r, c)); },
      set value(v: any) { write(`Set ${cellName(r, c)}`, () => setValue(s, r, c, v)); },
      get formula(): string { const x = s.get(r, c); return x?.f != null ? '=' + x.f : x?.v == null ? '' : String(outValue(x.v)); },
      set formula(f: string) { write(`Formula ${cellName(r, c)}`, () => { guard(s, r, c); const cur = s.get(r, c); const t = String(f ?? ''); if (!t) { s.set(r, c, cur?.s ? { s: cur.s } : undefined); return; } s.set(r, c, t.startsWith('=') ? { ...(cur?.s ? { s: cur.s } : {}), f: t.slice(1) } : { ...(cur?.s ? { s: cur.s } : {}), v: toScalar(t) }); }); },
      get format(): string { return s.styleOf(r, c).numFmt || 'General'; },
      set format(fmt: string) { write(`Format ${cellName(r, c)}`, () => { guard(s, r, c); const cur = s.get(r, c); s.set(r, c, { ...(cur || {}), s: book.derivedStyle(cur?.s ?? 0, { numFmt: fmt || undefined }) }); }); },
      get text(): string { fresh(); return cellText(s, r, c); },
      get style() { return { ...s.styleOf(r, c) }; },
      setStyle(patch: Record<string, any>) { write(`Style ${cellName(r, c)}`, () => { guard(s, r, c); const cur = s.get(r, c); s.set(r, c, { ...(cur || {}), s: book.derivedStyle(cur?.s ?? 0, patch) }); }); },
      offset(dr: number, dc: number) { return cellObj(cellName(Math.max(0, r + dr), Math.max(0, c + dc)), s); },
    };
  };

  const rangeObj = (addr: string, sheet?: Sheet) => {
    const { sheet: s, range: rg } = locate(addr, sheet);
    const rows = rg.r2 - rg.r1 + 1, cols = rg.c2 - rg.c1 + 1;
    if (rows * cols > 5_000_000) throw new Error('Range is too large for a macro (5 million cells at most)');
    return {
      get address() { return rangeName(rg); },
      get sheet() { return s.name; },
      rows, cols,
      get values(): any[][] { fresh(); const out: any[][] = []; for (let r = rg.r1; r <= rg.r2; r++) { const row: any[] = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(outValue(s.value(r, c))); out.push(row); } return out; },
      set values(m: any[][]) {
        if (!Array.isArray(m)) throw new Error('values must be a 2-D array');
        write(`Set ${rangeName(rg)}`, () => {
          const grid = m.length && !Array.isArray(m[0]) ? (rows === 1 ? [m] : m.map((x) => [x])) : m;
          for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
            const row = grid[i % grid.length] as any[];
            setValue(s, rg.r1 + i, rg.c1 + j, row ? row[j % row.length] : null);
          }
        });
      },
      get formulas(): string[][] { const out: string[][] = []; for (let r = rg.r1; r <= rg.r2; r++) { const row: string[] = []; for (let c = rg.c1; c <= rg.c2; c++) { const x = s.get(r, c); row.push(x?.f != null ? '=' + x.f : x?.v == null ? '' : String(outValue(x.v))); } out.push(row); } return out; },
      clear() { write(`Clear ${rangeName(rg)}`, () => { for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const x = s.get(r, c); if (x && canChange(s, r, c)) s.set(r, c, x.s ? { s: x.s } : undefined); } }); },
      cell(i: number, j: number) { return cellObj(cellName(rg.r1 + i, rg.c1 + j), s); },
      forEach(fn: (cell: ReturnType<typeof cellObj>, i: number, j: number) => void) { for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) fn(cellObj(cellName(rg.r1 + i, rg.c1 + j), s), i, j); },
    };
  };

  const sheetObj = (s: Sheet) => ({
    get name() { return s.name; },
    cell: (a: string) => cellObj(a, s),
    range: (a: string) => rangeObj(a, s),
    get usedRange() { const u = s.usedRange(false); return u ? rangeName(u) : null; },
    activate() { host.trace?.(`Activate ${s.name}`); const i = book.sheets.indexOf(s); if (i >= 0) host.exec('sheet.activate', i); },
  });

  return {
    cell: (addr: string) => cellObj(addr),
    range: (addr: string) => rangeObj(addr),
    sheet: (name?: string) => { const s = name == null ? host.sheet : book.sheetByName(String(name)); if (!s) throw new Error(`No sheet named ${name}`); return sheetObj(s); },
    get sheets(): string[] { return book.sheets.map((s) => s.name); },
    get activeSheet() { return sheetObj(host.sheet); },
    addSheet(name?: string) {
      host.trace?.('Add sheet');
      let made: Sheet | null = null;
      book.transact('Macro', () => {
        book.recordBook();
        const S = host.sheet.constructor as any;
        made = new S(book, name && !book.sheetByName(name) ? name : book.uniqueSheetName());
        book.sheets.push(made!);
      });
      return sheetObj(made!);
    },
    select(addr: string) { const { sheet: s, range } = locate(addr); host.trace?.(`Select ${rangeName(range)}`); if (host.select) host.select(s, range.r1, range.c1, range); else { s.sel = { r: range.r1, c: range.c1, ranges: [range] }; } },
    get selection(): string { const s = host.sheet; return s.sel.ranges.map(rangeName).join(','); },
    get activeCell(): string { const s = host.sheet; return cellName(s.sel.r, s.sel.c); },
    exec(id: string, arg?: any) { host.trace?.(`Run ${id}`); if (!getCommand(id)) throw new Error(`Unknown command: ${id}`); return host.exec(id, arg); },
    commands: () => allCommands().map((c) => c.id),
    msgBox: (text: any) => (host.msgBox ? host.msgBox(String(text)) : undefined),
    inputBox: (prompt: any, value?: any) => (host.inputBox ? host.inputBox(String(prompt), value == null ? '' : String(value)) : null),
    log: (...a: any[]) => host.log?.(...a),
    toast: (text: any) => toast(String(text)),
    /** Date serial for a JavaScript Date or ISO text (for comparisons with date cells). */
    dateSerial: (d: Date | string) => jsDateToSerial(d instanceof Date ? d : new Date(d)),
  };
}

export type MacroApi = ReturnType<typeof createMacroApi>;

/** Function names declared in a module ("function Name(" at the start of a line). */
export function moduleProcs(code: string): string[] {
  return [...code.matchAll(/^\s*(?:async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/gm)].map((m) => m[1]);
}

/**
 * Runs JavaScript macro code with the API. The synchronous part (everything before the first
 * await) is one undoable step. `proc` calls a function declared in the code.
 */
export function runJs(host: MacroHost, code: string, proc?: string): Promise<any> {
  const api = createMacroApi(host);
  const body = `"use strict";\nreturn (async () => {\n${code}\n;${proc ? `\nif (typeof ${proc} !== 'function') throw new Error('No function named ${proc}');\nreturn await ${proc}();` : ''}\n})();`;
  let fn: (api: MacroApi) => Promise<any>;
  try { fn = new Function('api', body) as any; } catch (e) { return Promise.reject(new Error(`Syntax error: ${(e as Error).message}`)); }
  return host.book.transact('Run Macro', () => fn(api));
}

function appHost(app: App, log?: (...a: any[]) => void, trace?: (s: string) => void): MacroHost {
  return {
    get book() { return app.book; },
    get sheet() { return app.sheet; },
    exec: (id, arg) => app.exec(id, arg),
    select: (s, r, c, range) => { const i = app.book.sheets.indexOf(s); if (s !== app.sheet && i >= 0) app.exec('sheet.activate', i); app.grid.select(r, c, { range }); },
    msgBox: (t) => alertDialog('Macro', t),
    inputBox: (p, v) => promptDialog('Macro', p, v),
    log: log || ((...a) => console.log('[macro]', ...a)),
    recalc: () => app.engine.recalc(),
    trace,
  } as MacroHost;
}

// ============================================================================================
// Recorded macros

/** Commands whose arguments contain cell positions that relative macros shift. */
const POSITIONAL = new Set(['nav.select', 'cell.enter']);

function shiftArg(arg: any, dr: number, dc: number): any {
  if (!arg || typeof arg !== 'object') return arg;
  const out: any = Array.isArray(arg) ? [] : {};
  for (const [k, v] of Object.entries(arg)) {
    if ((k === 'r' || k === 'r1' || k === 'r2') && typeof v === 'number') out[k] = v + dr;
    else if ((k === 'c' || k === 'c1' || k === 'c2') && typeof v === 'number') out[k] = v + dc;
    else if (v && typeof v === 'object') out[k] = shiftArg(v, dr, dc);
    else out[k] = v;
  }
  return out;
}

/** Converts absolute positions in recorded steps to offsets from the starting active cell. */
export function toRelativeSteps(steps: Macro['steps'], start: { r: number; c: number }): Macro['steps'] {
  return steps.map((s) => (POSITIONAL.has(s.id) && s.arg ? { id: s.id, arg: shiftArg(s.arg, -start.r, -start.c) } : s));
}

/** Converts relative steps back to positions from the current active cell. Offsets that land off the sheet are clamped to it. */
export function fromRelativeSteps(steps: Macro['steps'], at: { r: number; c: number }): Macro['steps'] {
  return steps.map((s) => {
    if (!(POSITIONAL.has(s.id) && s.arg)) return s;
    const a = shiftArg(s.arg, at.r, at.c);
    const clampAll = (o: any) => { if (!o || typeof o !== 'object') return; for (const [k, v] of Object.entries(o)) { if (typeof v === 'number' && /^(r|c)[12]?$/.test(k)) o[k] = Math.max(0, v); else clampAll(v); } };
    clampAll(a);
    return { id: s.id, arg: a };
  });
}

/** Replays a recorded macro. step: ask before each step (Step Into). Returns the number of steps run. */
export async function runRecorded(app: App, m: Macro, opts: { step?: boolean } = {}): Promise<number> {
  const steps = m.relative ? fromRelativeSteps(m.steps, activeCell(app)) : m.steps;
  let n = 0;
  for (let i = 0; i < steps.length; i++) {
    const step = steps[i];
    if (step.id === '__script') { await runJs(appHost(app), String(step.arg || '')); n++; continue; }
    const c = getCommand(step.id);
    if (!c) { toast(`Macro step skipped: unknown command ${step.id}`); continue; }
    if (opts.step) {
      const ok = await confirmDialog(`Step ${i + 1} of ${steps.length}`, `${c.label}${step.arg !== undefined ? `\n${JSON.stringify(step.arg).slice(0, 300)}` : ''}`, 'Run Step', 'Stop');
      if (!ok) break;
    }
    const r = app.exec(step.id, step.arg);
    if (r instanceof Promise) await r;
    n++;
  }
  return n;
}

let recMeta: { start: { r: number; c: number }; description: string; key: string; relative: boolean } | null = null;

async function startRecording(app: App): Promise<void> {
  if (app.recording) { await stopRecording(app); return; }
  const taken = new Set([...app.prefs.macros.map((m) => m.name.toLowerCase())]);
  let i = 1;
  while (taken.has(`macro${i}`)) i++;
  const r = await openDialog({
    title: 'Record Macro', fields: [
      { key: 'name', label: 'Macro name', type: 'text', value: `Macro${i}` },
      { key: 'key', label: 'Shortcut key (optional)', type: 'text', value: '', placeholder: 'e.g. Mod-Shift-M' },
      { key: 'desc', label: 'Description', type: 'textarea', rows: 2, value: '' },
      { key: 'relative', label: 'Use relative references (replay from whichever cell is active)', type: 'checkbox', value: relativeRefs },
    ],
    validate: (v) => (!/^[A-Za-z_][\w ]{0,63}$/.test(String(v.name).trim()) ? 'Macro names start with a letter and use letters, digits, spaces and underscores.' : null),
  });
  if (!r) return;
  relativeRefs = !!r.relative;
  lsSet(REL_KEY, relativeRefs ? '1' : '0');
  recMeta = { start: activeCell(app), description: String(r.desc || ''), key: String(r.key || '').trim(), relative: relativeRefs };
  app.recording = { name: String(r.name).trim(), steps: [] };
  app.recordingPaused = false;
  toast('Recording… use Stop Recording (or the REC button on the status bar) when finished.');
  app.refresh();
}

async function stopRecording(app: App): Promise<void> {
  const rec = app.recording;
  if (!rec) { toast('No macro is being recorded.'); return; }
  app.recording = null;
  app.recordingPaused = false;
  const meta = recMeta || { start: { r: 0, c: 0 }, description: '', key: '', relative: false };
  recMeta = null;
  if (!rec.steps.length) { toast('Nothing was recorded.'); app.refresh(); return; }
  const m: Macro = { name: rec.name, steps: meta.relative ? toRelativeSteps(rec.steps, meta.start) : rec.steps, description: meta.description || undefined, key: meta.key || undefined, created: Date.now(), relative: meta.relative || undefined };
  saveMacro(app, m);
  toast(`Macro “${m.name}” saved (${m.steps.length} step${m.steps.length === 1 ? '' : 's'}).`);
}

function saveMacro(app: App, m: Macro, oldName?: string): void {
  const prev = app.prefs.macros.find((x) => x.name === (oldName || m.name));
  if (prev?.key) delete app.prefs.keymap[prev.key];
  app.prefs.macros = [...app.prefs.macros.filter((x) => x.name !== (oldName || m.name) && x.name !== m.name), m];
  if (m.key) app.prefs.keymap[m.key] = `macro:${m.name}`;
  app.savePrefs();
  app.ui.rebuild?.();
}

// ============================================================================================
// Running macros by name

interface MacroEntry { kind: 'recorded' | 'module' | 'proc'; name: string; label: string; module?: string; proc?: string; macro?: Macro }

function listMacros(app: App): MacroEntry[] {
  const out: MacroEntry[] = app.prefs.macros.map((m) => ({ kind: 'recorded', name: m.name, label: `${m.name}  (recorded${m.key ? ', ' + m.key : ''})`, macro: m }));
  for (const mod of app.book.modules) {
    const procs = moduleProcs(mod.code);
    if (!procs.length) out.push({ kind: 'module', name: mod.name, label: `${mod.name}  (module)`, module: mod.name });
    for (const p of procs) out.push({ kind: 'proc', name: `${mod.name}.${p}`, label: `${mod.name}.${p}`, module: mod.name, proc: p });
  }
  return out;
}

async function runMacroByName(app: App, name: string, opts: { step?: boolean } = {}): Promise<void> {
  const rec = app.prefs.macros.find((m) => m.name === name);
  if (rec) {
    if (rec.steps.some((s) => s.id === '__script') && !(await allowMacros(app))) return;
    const n = await runRecorded(app, rec, opts);
    if (opts.step) toast(`${n} step${n === 1 ? '' : 's'} run`);
    return;
  }
  const [modName, proc] = name.includes('.') ? name.split('.', 2) : [name, undefined];
  const mod = app.book.modules.find((m) => m.name === modName) || (proc ? undefined : app.book.modules.find((m) => moduleProcs(m.code).includes(name)));
  if (!mod) { toast(`Macro “${name}” was not found.`); return; }
  const procName = proc || (mod.name !== name ? name : undefined);
  if (!(await allowMacros(app))) return;
  const traceLines: string[] = [];
  try {
    await runJs(appHost(app, undefined, opts.step ? (s) => traceLines.push(s) : undefined), mod.code, procName);
    if (opts.step) await alertDialog('Macro Trace', traceLines.length ? traceLines.join('\n') : 'The macro made no changes.');
  } catch (e) {
    await alertDialog('Macro Error', `${procName ? `${mod.name}.${procName}` : mod.name}: ${(e as Error).message}`);
  }
}

// ============================================================================================
// Macros dialog and editor

async function macrosDialog(app: App): Promise<void> {
  for (;;) {
    const list = listMacros(app);
    const r = await openDialog({
      title: 'Macros', wide: true, ok: 'Close', cancel: false,
      fields: [
        { key: 'name', label: 'Macro name', type: 'list', rows: 10, options: list.map((m) => ({ value: m.name, label: m.label })), value: list[0]?.name },
        { key: 'info', type: 'info', html: list.length ? '<p class="ql-help">Recorded macros replay commands; module macros are JavaScript in this workbook (Developer ▸ Macro Editor).</p>' : '<p class="ql-help">No macros yet. Record one with Developer ▸ Record Macro, or write one in the Macro Editor.</p>' },
      ],
      buttons: [
        { label: 'Run', primary: true, onClick: (v, c) => c.close({ act: 'run', name: v.name }) },
        { label: 'Step Into', onClick: (v, c) => c.close({ act: 'step', name: v.name }) },
        { label: 'Edit', onClick: (v, c) => c.close({ act: 'edit', name: v.name }) },
        { label: 'Create', onClick: (_v, c) => c.close({ act: 'create' }) },
        { label: 'Delete', onClick: (v, c) => c.close({ act: 'delete', name: v.name }) },
        { label: 'Options…', onClick: (v, c) => c.close({ act: 'options', name: v.name }) },
      ],
    });
    if (!r || !r.act) return;
    const entry = list.find((m) => m.name === r.name);
    if (r.act === 'create') { await macroEditor(app); return; }
    if (!entry) { toast('Select a macro first.'); continue; }
    if (r.act === 'run' || r.act === 'step') { await runMacroByName(app, entry.name, { step: r.act === 'step' }); return; }
    if (r.act === 'edit') {
      if (entry.kind === 'recorded') await editRecorded(app, entry.macro!);
      else await macroEditor(app, entry.module);
      continue;
    }
    if (r.act === 'delete') {
      if (!(await confirmDialog('Delete Macro', `Delete the macro “${entry.name}”?`, 'Delete'))) continue;
      if (entry.kind === 'recorded') {
        if (entry.macro!.key) delete app.prefs.keymap[entry.macro!.key];
        app.prefs.macros = app.prefs.macros.filter((m) => m !== entry.macro);
        app.savePrefs();
        app.ui.rebuild?.();
      } else if (!app.isReadOnly()) {
        app.book.transact('Delete Macro', () => {
          app.book.recordBook();
          if (entry.kind === 'module' || !entry.proc) app.book.modules = app.book.modules.filter((m) => m.name !== entry.module);
          else app.book.modules = app.book.modules.map((m) => (m.name === entry.module ? { ...m, code: removeProc(m.code, entry.proc!) } : m));
        });
      }
      continue;
    }
    if (r.act === 'options') {
      if (entry.kind !== 'recorded') { toast('Shortcut keys and descriptions belong to recorded macros. For module macros, assign a shortcut to “macro:Module.Name” in Customize Keyboard.'); continue; }
      const m = entry.macro!;
      const o = await openDialog({ title: 'Macro Options', fields: [
        { key: 'name', label: 'Name', type: 'text', value: m.name },
        { key: 'key', label: 'Shortcut key', type: 'text', value: m.key || '', placeholder: 'e.g. Mod-Shift-M' },
        { key: 'desc', label: 'Description', type: 'textarea', rows: 3, value: m.description || '' },
        { key: 'relative', label: 'Replay relative to the active cell', type: 'checkbox', value: !!m.relative },
      ] });
      if (!o) continue;
      saveMacro(app, { ...m, name: String(o.name).trim() || m.name, key: String(o.key || '').trim() || undefined, description: String(o.desc || '') || undefined, relative: !!o.relative || undefined }, m.name);
    }
  }
}

function removeProc(code: string, proc: string): string {
  const re = new RegExp(`^\\s*(?:async\\s+)?function\\s+${proc}\\s*\\(`, 'm');
  const m = re.exec(code);
  if (!m) return code;
  let i = code.indexOf('{', m.index), depth = 0;
  for (; i < code.length; i++) { if (code[i] === '{') depth++; else if (code[i] === '}' && --depth === 0) break; }
  return (code.slice(0, m.index) + code.slice(i + 1)).replace(/\n{3,}/g, '\n\n');
}

async function editRecorded(app: App, m: Macro): Promise<void> {
  const r = await openDialog({
    title: `Edit Macro: ${m.name}`, wide: true,
    help: 'One step per line: a command ID, optionally followed by a space and a JSON argument. Command IDs are listed under Help ▸ Command Reference.',
    fields: [{ key: 'steps', label: 'Steps', type: 'textarea', rows: 14, value: m.steps.map((s) => (s.arg === undefined ? s.id : `${s.id} ${JSON.stringify(s.arg)}`)).join('\n') }],
    validate: (v) => { try { parseStepText(String(v.steps)); return null; } catch (e) { return (e as Error).message; } },
  });
  if (!r) return;
  saveMacro(app, { ...m, steps: parseStepText(String(r.steps)) });
  toast('Macro saved');
}

export function parseStepText(text: string): Macro['steps'] {
  const steps: Macro['steps'] = [];
  text.split(/\r?\n/).forEach((line, i) => {
    const t = line.trim();
    if (!t || t.startsWith('//')) return;
    const sp = t.indexOf(' ');
    const id = sp < 0 ? t : t.slice(0, sp);
    if (!/^[\w.:-]+$/.test(id)) throw new Error(`Line ${i + 1}: "${id}" is not a command ID`);
    if (sp < 0) { steps.push({ id }); return; }
    try { steps.push({ id, arg: JSON.parse(t.slice(sp + 1)) }); } catch { throw new Error(`Line ${i + 1}: the argument is not valid JSON`); }
  });
  return steps;
}

const SAMPLE_MODULE = `// JavaScript macros for this workbook. Each function is a macro you can run from
// Developer ▸ Macros, a button, the Quick Access Bar or a shortcut ("macro:Module1.Name").
//
// api.cell('A1').value / .formula / .format     api.range('A1:C3').values  (2-D array)
// api.sheet('Data').cell('B2')    api.select('A1:B5')    api.exec('fmt.bold')
// api.msgBox('text')    await api.inputBox('Question?')    api.log(...)

function HighlightNegatives() {
  const r = api.range(api.selection.split(',')[0]);
  r.forEach((cell) => { if (typeof cell.value === 'number' && cell.value < 0) cell.setStyle({ color: '#c00000', bold: true }); });
}

async function FillSeries() {
  const n = Number(await api.inputBox('How many rows?', '10'));
  const start = api.activeCell;
  for (let i = 0; i < n; i++) api.cell(start).offset(i, 0).value = i + 1;
}
`;

/** Full-screen JavaScript editor for the workbook's macro modules. */
async function macroEditor(app: App, moduleName?: string): Promise<void> {
  const book = app.book;
  let mods = book.modules.map((m) => ({ ...m }));
  if (!mods.length) mods = [{ name: 'Module1', code: SAMPLE_MODULE }];
  let cur = Math.max(0, mods.findIndex((m) => m.name === moduleName));
  let dirty = false;
  const listEl = h('div', { class: 'ql-macro-mods', role: 'listbox', 'aria-label': 'Modules' });
  const code = h('textarea', { class: 'ql-macro-code', spellcheck: 'false', autocapitalize: 'off', autocomplete: 'off', 'aria-label': 'Macro code' }) as HTMLTextAreaElement;
  const out = h('pre', { class: 'ql-macro-out', 'aria-live': 'polite' });
  const procSel = h('select', { class: 'ql-input', 'aria-label': 'Macro to run' }) as HTMLSelectElement;
  const log = (...a: any[]) => { out.textContent += a.map((x) => (typeof x === 'string' ? x : JSON.stringify(x))).join(' ') + '\n'; out.scrollTop = out.scrollHeight; };
  const refreshProcs = () => {
    clear(procSel);
    procSel.appendChild(h('option', { value: '' }, '(whole module)'));
    for (const p of moduleProcs(code.value)) procSel.appendChild(h('option', { value: p }, p));
  };
  const drawList = () => {
    clear(listEl);
    mods.forEach((m, i) => listEl.appendChild(h('button', { class: 'ql-nav-item' + (i === cur ? ' ql-on' : ''), type: 'button', onclick: () => { mods[cur].code = code.value; cur = i; load(); } }, icon('code'), ' ', m.name)));
  };
  const load = () => { code.value = mods[cur]?.code ?? ''; drawList(); refreshProcs(); };
  code.addEventListener('input', () => { dirty = true; mods[cur].code = code.value; });
  code.addEventListener('change', refreshProcs);
  code.addEventListener('keydown', (e) => {
    if (e.key === 'Tab' && !e.ctrlKey && !e.altKey) {
      e.preventDefault();
      const s = code.selectionStart, en = code.selectionEnd;
      code.setRangeText('  ', s, en, 'end');
      dirty = true; mods[cur].code = code.value;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); save(); }
    if (e.key === 'F5') { e.preventDefault(); void run(); }
  });
  const save = () => {
    if (app.isReadOnly()) { toast('This workbook is read-only.'); return false; }
    mods[cur].code = code.value;
    book.transact('Edit Macros', () => { book.recordBook(); book.modules = mods.map((m) => ({ name: m.name, code: m.code })); });
    dirty = false;
    toast('Macros saved in the workbook');
    return true;
  };
  const run = async () => {
    mods[cur].code = code.value;
    if (!(await allowMacros(app))) return;
    out.textContent = '';
    const t0 = performance.now();
    try {
      const res = await runJs(appHost(app, log), code.value, procSel.value || undefined);
      log(`✓ Finished in ${Math.round(performance.now() - t0)} ms${res !== undefined ? ` → ${JSON.stringify(res)}` : ''}`);
    } catch (e) { log(`✗ ${(e as Error).message}`); }
  };
  const toolbar = h('div', { class: 'ql-row ql-macro-bar' },
    h('button', { class: 'ql-btn ql-primary', type: 'button', onclick: () => void run(), title: 'Run (F5)' }, icon('play_arrow'), ' Run'), procSel,
    h('button', { class: 'ql-btn', type: 'button', onclick: () => save(), title: 'Save (Ctrl+S)' }, icon('save'), ' Save'),
    h('span', { class: 'ql-spacer' }),
    h('button', { class: 'ql-btn', type: 'button', onclick: async () => {
      const name = await promptDialog('New Module', 'Module name', `Module${mods.length + 1}`);
      if (!name?.trim()) return;
      if (mods.some((m) => m.name === name.trim())) { toast('A module with that name exists.'); return; }
      mods[cur].code = code.value; mods.push({ name: name.trim().replace(/\W+/g, '_'), code: '// New module\n\nfunction Main() {\n  \n}\n' }); cur = mods.length - 1; dirty = true; load();
    } }, icon('add'), ' Module'),
    h('button', { class: 'ql-btn', type: 'button', onclick: async () => {
      const name = await promptDialog('Rename Module', 'Module name', mods[cur].name);
      if (!name?.trim()) return;
      mods[cur].name = name.trim().replace(/\W+/g, '_'); dirty = true; drawList();
    } }, 'Rename'),
    h('button', { class: 'ql-btn', type: 'button', onclick: async () => {
      if (!(await confirmDialog('Delete Module', `Delete ${mods[cur].name}?`, 'Delete'))) return;
      mods.splice(cur, 1); if (!mods.length) mods.push({ name: 'Module1', code: '' }); cur = 0; dirty = true; load();
    } }, 'Delete'),
  );
  const body = h('div', { class: 'ql-macro-editor' }, toolbar, h('div', { class: 'ql-macro-main' }, listEl, code), h('div', { class: 'ql-label' }, 'Output'), out,
    h('p', { class: 'ql-help' }, 'api: cell(a).value/.formula/.format/.text/.setStyle()/.offset(), range(a).values/.formulas/.clear()/.forEach(), sheet(name), sheets, addSheet(), select(a), selection, activeCell, exec(id, arg), commands(), msgBox(t), inputBox(p) (await), log(), toast().'));
  load();
  await openDialog({
    title: 'Macro Editor', full: true, content: body, ok: 'Save and Close', cancel: 'Close',
    onOpen: () => setTimeout(() => code.focus(), 50),
  }).then(async (r) => {
    if (r) save();
    else if (dirty && (await confirmDialog('Save Changes?', 'Save the changes to the macro modules in this workbook?', 'Save', "Don't Save"))) save();
  });
}

// ============================================================================================
// Form controls

const CONTROL_TYPES: { type: FormControlType; label: string; icon: string; w: number; h: number }[] = [
  { type: 'button', label: 'Button', icon: 'smart_button', w: 96, h: 28 },
  { type: 'checkbox', label: 'Check Box', icon: 'check_box', w: 110, h: 22 },
  { type: 'option', label: 'Option Button', icon: 'radio_button_checked', w: 110, h: 22 },
  { type: 'combo', label: 'Combo Box (Drop-Down)', icon: 'arrow_drop_down_circle', w: 130, h: 24 },
  { type: 'list', label: 'List Box', icon: 'list', w: 130, h: 90 },
  { type: 'spin', label: 'Spin Button', icon: 'unfold_more', w: 22, h: 40 },
  { type: 'scroll', label: 'Scroll Bar', icon: 'linear_scale', w: 150, h: 20 },
  { type: 'label', label: 'Label', icon: 'label', w: 90, h: 20 },
  { type: 'group', label: 'Group Box', icon: 'select_all', w: 160, h: 110 },
];

function linkTarget(sheet: Sheet, ref?: string): { sheet: Sheet; r: number; c: number } | null {
  if (!ref) return null;
  const res = resolveRangeText(sheet.book, ref, sheet);
  return res ? { sheet: res.sheet, r: res.range.r1, c: res.range.c1 } : null;
}

export function controlItems(sheet: Sheet, ctl: FormControl): string[] {
  if (!ctl.input) return [];
  const res = resolveRangeText(sheet.book, ctl.input, sheet);
  if (!res) return [];
  const out: string[] = [];
  const rg = res.range;
  for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 999); r++) for (let c = rg.c1; c <= rg.c2; c++) out.push(cellText(res.sheet, r, c));
  return out;
}

/** Current state of a control: linked cell value, or its own state. */
export function controlValue(sheet: Sheet, ctl: FormControl): Scalar {
  const t = linkTarget(sheet, ctl.link);
  if (t) return t.sheet.value(t.r, t.c);
  if (ctl.type === 'checkbox' || ctl.type === 'option') return !!ctl.checked;
  return ctl.value ?? null;
}

function optionIndex(sheet: Sheet, d: DrawingObject): number {
  const grp = sheet.drawings.filter((x) => x.control?.type === 'option' && (x.control.group || '') === (d.control!.group || '') && (x.control.link || '') === (d.control!.link || ''));
  return grp.indexOf(d) + 1;
}

function renderControl(d: DrawingObject, g: Grid): HTMLElement | null {
  const ctl = d.control;
  if (!ctl) return null;
  const s = g.sheet;
  const v = controlValue(s, ctl);
  const label = ctl.label ?? CONTROL_TYPES.find((t) => t.type === ctl.type)?.label ?? '';
  const box = h('div', { class: `ql-ctl ql-ctl-${ctl.type}` + (designMode ? ' ql-ctl-design' : ''), style: `width:${d.at.w}px;height:${d.at.h}px` });
  switch (ctl.type) {
    case 'button': box.appendChild(h('span', null, label)); break;
    case 'label': box.appendChild(h('span', null, label)); break;
    case 'group': box.appendChild(h('span', { class: 'ql-ctl-caption' }, label)); break;
    case 'checkbox': {
      const on = v === true || v === 1 || (typeof v === 'string' && v.toUpperCase() === 'TRUE');
      box.append(h('span', { class: 'ql-ctl-box' + (on ? ' ql-on' : '') }, on ? '✓' : v instanceof CellError ? '■' : ''), h('span', null, label));
      break;
    }
    case 'option': {
      const on = ctl.link ? Number(v) === optionIndex(s, d) : !!ctl.checked;
      box.append(h('span', { class: 'ql-ctl-radio' + (on ? ' ql-on' : '') }), h('span', null, label));
      break;
    }
    case 'combo': {
      const items = controlItems(s, ctl);
      const i = typeof v === 'number' ? Math.round(v) : 0;
      box.append(h('span', { class: 'ql-ctl-value' }, items[i - 1] ?? ''), h('span', { class: 'ql-ctl-drop' }, '▾'));
      break;
    }
    case 'list': {
      const items = controlItems(s, ctl);
      const i = typeof v === 'number' ? Math.round(v) : 0;
      items.forEach((t, k) => box.appendChild(h('div', { class: 'ql-ctl-item' + (k + 1 === i ? ' ql-on' : '') }, t || ' ')));
      break;
    }
    case 'spin': box.append(h('span', { class: 'ql-ctl-spin-up' }, '▲'), h('span', { class: 'ql-ctl-spin-down' }, '▼')); break;
    case 'scroll': {
      const min = ctl.min ?? 0, max = ctl.max ?? 100;
      const x = typeof v === 'number' ? v : min;
      const pct = max > min ? Math.max(0, Math.min(1, (x - min) / (max - min))) : 0;
      const vertical = d.at.h > d.at.w;
      box.classList.toggle('ql-vertical', vertical);
      box.append(h('span', { class: 'ql-ctl-arrow' }, vertical ? '▲' : '◀'), h('span', { class: 'ql-ctl-track' }, h('span', { class: 'ql-ctl-thumb', style: vertical ? `top:calc(${pct * 100}% - ${pct * 14}px)` : `left:calc(${pct * 100}% - ${pct * 14}px)` })), h('span', { class: 'ql-ctl-arrow' }, vertical ? '▼' : '▶'));
      break;
    }
  }
  return box;
}

/** Applies a control interaction (Developer ▸ control click). part: 'up'/'down'/'pageUp'/'pageDown' for spinners and scroll bars, an item index for lists. */
export function controlAction(sheet: Sheet, d: DrawingObject, part?: string | number): { changed: boolean } {
  const ctl = d.control!;
  const book = sheet.book;
  const t = linkTarget(sheet, ctl.link);
  const setLinked = (v: Scalar) => {
    if (t) { if (!canChange(t.sheet, t.r, t.c)) throw new Error('The linked cell is locked on a protected sheet.'); const cur = t.sheet.get(t.r, t.c); t.sheet.set(t.r, t.c, { ...(cur?.s ? { s: cur.s } : {}), ...(v == null ? {} : { v }) }); }
    else {
      book.recordSheet(sheet);
      if (typeof v === 'boolean') ctl.checked = v; else if (typeof v === 'number') ctl.value = v;
    }
  };
  const cur = controlValue(sheet, ctl);
  let changed = false;
  book.transact(`Control: ${ctl.label || ctl.type}`, () => {
    switch (ctl.type) {
      case 'checkbox': setLinked(!(cur === true || cur === 1 || String(cur).toUpperCase() === 'TRUE')); changed = true; break;
      case 'option': {
        if (t) setLinked(optionIndex(sheet, d));
        else {
          book.recordSheet(sheet);
          for (const x of sheet.drawings) if (x.control?.type === 'option' && (x.control.group || '') === (ctl.group || '')) x.control.checked = x === d;
        }
        changed = true; break;
      }
      case 'combo': case 'list': if (typeof part === 'number') { setLinked(part); changed = true; } break;
      case 'spin': case 'scroll': {
        const min = ctl.min ?? 0, max = ctl.max ?? 100, step = ctl.step || 1, page = ctl.page || 10;
        const x = typeof cur === 'number' ? cur : min;
        const delta = part === 'up' ? step : part === 'down' ? -step : part === 'pageUp' ? page : part === 'pageDown' ? -page : 0;
        const next = Math.max(min, Math.min(max, x + delta));
        if (next !== cur) { setLinked(next); changed = true; }
        break;
      }
    }
  });
  return { changed };
}

function findControl(app: App, id: string): DrawingObject | undefined {
  return app.sheet.drawings.find((d) => d.id === id && d.control);
}

objectHooks.render.push(renderControl);
objectHooks.pointerdown.push((g, id, e) => {
  if (designMode) return false;
  const d = g.sheet.drawings.find((x) => x.id === id);
  if (!d?.control) return false;
  e.preventDefault();
  const el = (e.currentTarget as HTMLElement) || (e.target as HTMLElement);
  const rect = el.getBoundingClientRect();
  const fx = (e.clientX - rect.left) / Math.max(1, rect.width), fy = (e.clientY - rect.top) / Math.max(1, rect.height);
  let part: string | number | undefined;
  const t = d.control.type;
  if (t === 'spin') part = fy < 0.5 ? 'up' : 'down';
  else if (t === 'scroll') {
    const vertical = d.at.h > d.at.w;
    const f = vertical ? fy : fx, len = vertical ? rect.height : rect.width;
    const arrow = 16 * (len / (vertical ? d.at.h : d.at.w));
    const p = f * len;
    // Up/left arrow decreases, down/right increases (a vertical scroll bar grows downward, like a list)
    part = p < arrow ? 'down' : p > len - arrow ? 'up' : f < 0.5 ? 'pageDown' : 'pageUp';
  } else if (t === 'list') {
    const items = controlItems(g.sheet, d.control);
    const itemEls = el.querySelectorAll('.ql-ctl-item');
    const hit = [...itemEls].findIndex((x) => { const r = x.getBoundingClientRect(); return e.clientY >= r.top && e.clientY < r.bottom; });
    part = hit >= 0 && hit < items.length ? hit + 1 : undefined;
    if (part == null) return true;
  }
  g.app.exec('dev.controlClick', { id, part, x: e.clientX, y: e.clientY });
  return true;
});
objectHooks.dblclick.push((g, id) => {
  const d = g.sheet.drawings.find((x) => x.id === id);
  if (!d?.control) return false;
  if (designMode) g.app.exec('dev.properties', { id });
  return true;
});

async function insertControl(app: App, type?: FormControlType): Promise<void> {
  if (!type) {
    const r = await openDialog({ title: 'Insert Control', fields: [{ key: 't', label: 'Control', type: 'list', rows: 9, options: CONTROL_TYPES.map((t) => ({ value: t.type, label: t.label })), value: 'button' }] });
    if (!r?.t) return;
    type = r.t as FormControlType;
  }
  const spec = CONTROL_TYPES.find((t) => t.type === type)!;
  const s = app.sheet;
  if (s.protection.on && !s.protection.allow.includes('objects')) { app.exec('review.protectionNotice'); return; }
  const { r, c } = activeCell(app);
  const ctl: FormControl = { type, label: ['button', 'checkbox', 'option', 'label', 'group'].includes(type) ? `${spec.label.split(' ')[0]} ${s.drawings.filter((d) => d.control?.type === type).length + 1}` : undefined };
  if (type === 'spin' || type === 'scroll') Object.assign(ctl, { min: 0, max: 100, step: 1, page: 10 });
  if (type !== 'button' && type !== 'label' && type !== 'group') ctl.link = rangeRefText(s, { r1: r, c1: c + 2, r2: r, c2: c + 2 }, true, s);
  if (type === 'combo' || type === 'list') {
    const reg = dataRegion(app);
    ctl.input = rangeRefText(s, { r1: reg.r1, c1: reg.c1, r2: reg.r2, c2: reg.c1 }, true, s);
  }
  const d: DrawingObject = { id: newId('ctl'), kind: 'shape', shape: 'rect', at: { r, c, dx: 2, dy: 2, w: spec.w, h: spec.h }, alt: spec.label, control: ctl };
  app.book.transact(`Insert ${spec.label}`, () => { app.book.recordSheet(s); s.drawings.push(d); });
  designMode = true;
  app.grid.objects.select(d.id);
  const p = await controlProperties(app, d.id);
  if (p) toast('Turn off Design Mode to use the control.', { action: 'Turn Off', onAction: () => app.exec('dev.designMode') });
}

async function controlProperties(app: App, id?: string): Promise<boolean> {
  const s = app.sheet;
  const sel = id || app.grid.objects.selected;
  const d = sel ? s.drawings.find((x) => x.id === sel) : undefined;
  if (!d?.control) { toast('Select a form control first: turn on Design Mode, then click the control.'); return false; }
  const ctl = d.control;
  const t = ctl.type;
  const macros = [{ value: '', label: '(none)' }, ...listMacros(app).map((m) => ({ value: m.name, label: m.name }))];
  const r = await openDialog({
    title: `Format Control: ${CONTROL_TYPES.find((x) => x.type === t)?.label}`, fields: [
      { key: 'label', label: 'Text', type: 'text', value: ctl.label || '', visibleIf: () => ['button', 'checkbox', 'option', 'label', 'group'].includes(t) },
      { key: 'link', label: 'Cell link', type: 'text', value: ctl.link || '', placeholder: '$C$1', visibleIf: () => !['button', 'label', 'group'].includes(t) },
      { key: 'input', label: 'Input range (list items)', type: 'text', value: ctl.input || '', placeholder: '$A$1:$A$10', visibleIf: () => t === 'combo' || t === 'list' },
      { key: 'min', label: 'Minimum value', type: 'number', value: ctl.min ?? 0, half: true, visibleIf: () => t === 'spin' || t === 'scroll' },
      { key: 'max', label: 'Maximum value', type: 'number', value: ctl.max ?? 100, half: true, visibleIf: () => t === 'spin' || t === 'scroll' },
      { key: 'step', label: 'Incremental change', type: 'number', value: ctl.step ?? 1, half: true, visibleIf: () => t === 'spin' || t === 'scroll' },
      { key: 'page', label: 'Page change', type: 'number', value: ctl.page ?? 10, half: true, visibleIf: () => t === 'scroll' },
      { key: 'group', label: 'Option group name', type: 'text', value: ctl.group || '', visibleIf: () => t === 'option' },
      { key: 'macro', label: t === 'button' ? 'Macro to run when clicked' : 'Macro to run after a change', type: 'select', options: macros, value: ctl.macro || '' },
    ],
    validate: (v) => {
      for (const k of ['link', 'input']) if (v[k] && !resolveRangeText(app.book, String(v[k]), s)) return `“${v[k]}” is not a valid reference.`;
      if ((t === 'spin' || t === 'scroll') && Number(v.min) >= Number(v.max)) return 'The maximum must be greater than the minimum.';
      return null;
    },
  });
  if (!r) return false;
  app.book.transact('Format Control', () => {
    app.book.recordSheet(s);
    const x = s.drawings.find((y) => y.id === d.id);
    if (!x?.control) return;
    const c = x.control;
    if (r.label != null) c.label = String(r.label);
    c.link = String(r.link || '').trim() || undefined;
    c.input = String(r.input || '').trim() || undefined;
    if (t === 'spin' || t === 'scroll') { c.min = Number(r.min) || 0; c.max = Number(r.max); c.step = Math.max(1e-9, Number(r.step) || 1); c.page = Number(r.page) || 10; }
    c.group = String(r.group || '').trim() || undefined;
    c.macro = r.macro || undefined;
    x.alt = c.label || x.alt;
  });
  return true;
}

// ============================================================================================
// JSON import/export

/** Rows (first row = headers) to an array of objects. Empty header cells become Column1, Column2… */
export function rowsToObjects(rows: Scalar[][], opts: { dates?: (r: number, c: number, v: number) => string | null } = {}): Record<string, any>[] {
  if (!rows.length) return [];
  const used = new Set<string>();
  const keys = rows[0].map((h0, i) => {
    let k = h0 == null || h0 === '' ? `Column${i + 1}` : String(h0 instanceof CellError ? h0.code : h0);
    const base = k;
    let n = 2;
    while (used.has(k)) k = `${base}_${n++}`;
    used.add(k);
    return k;
  });
  return rows.slice(1).filter((r) => r.some((v) => v != null && v !== '')).map((row, i) => {
    const o: Record<string, any> = {};
    keys.forEach((k, j) => {
      const v = row[j] ?? null;
      o[k] = v instanceof CellError ? v.code : typeof v === 'number' && opts.dates ? opts.dates(i + 1, j, v) ?? v : v;
    });
    return o;
  });
}

/** JSON data (array of objects, array of arrays, or an object holding one) to rows with a header row. Nested objects are flattened with dotted keys. */
export function jsonToRows(data: any): Scalar[][] {
  let arr: any = data;
  if (arr && !Array.isArray(arr) && typeof arr === 'object') {
    const inner = Object.values(arr).find((v) => Array.isArray(v));
    arr = inner || [arr];
  }
  if (!Array.isArray(arr)) throw new Error('The JSON must be an array of records.');
  if (arr.every((x: any) => Array.isArray(x))) return arr.map((row: any[]) => row.map(scalarOf));
  const flat = arr.map((x: any) => flatten(x));
  const keys: string[] = [];
  for (const o of flat) for (const k of Object.keys(o)) if (!keys.includes(k)) keys.push(k);
  return [keys, ...flat.map((o: any) => keys.map((k) => scalarOf(o[k])))];
}

function flatten(x: any, prefix = '', out: Record<string, any> = {}): Record<string, any> {
  if (x == null || typeof x !== 'object' || Array.isArray(x)) { out[prefix || 'value'] = Array.isArray(x) ? (x.every((y) => y == null || typeof y !== 'object') ? x.join(', ') : JSON.stringify(x)) : x; return out; }
  for (const [k, v] of Object.entries(x)) {
    const key = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === 'object' && !Array.isArray(v) && Object.keys(v).length && prefix.split('.').length < 3) flatten(v, key, out);
    else out[key] = Array.isArray(v) ? (v.every((y) => y == null || typeof y !== 'object') ? v.join(', ') : JSON.stringify(v)) : v && typeof v === 'object' ? JSON.stringify(v) : v;
  }
  return out;
}

function scalarOf(v: any): Scalar {
  if (v == null) return null;
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  if (typeof v === 'string') return errorFromCode(v) || v;
  return JSON.stringify(v);
}

async function exportJson(app: App): Promise<void> {
  const s = app.sheet;
  const rg = dataRegion(app);
  if (!s.usedRange(false)) { toast('There is no data to export.'); return; }
  const rows: Scalar[][] = [];
  for (let r = rg.r1; r <= rg.r2; r++) { if (s.hiddenRows.has(r) && s.filter) continue; const row: Scalar[] = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(s.value(r, c)); rows.push(row); }
  const header = hasHeaderRow(s, rg);
  const toIso = (v: number) => { const d = new Date(Math.round((v - 25569) * 86400000)); return isNaN(+d) ? null : d.toISOString().replace(/T00:00:00\.000Z$/, '').replace(/\.000Z$/, 'Z'); };
  const make = (useHeader: boolean, isoDates: boolean, pretty: boolean) => {
    const src = useHeader ? rows : [rows[0].map((_, i) => `Column${i + 1}`), ...rows];
    const objs = rowsToObjects(src as Scalar[][], { dates: isoDates ? (i, j, v) => (isDateFormat(s.styleOf(rg.r1 + (useHeader ? i : i - 1), rg.c1 + j).numFmt) ? toIso(v) : null) : undefined });
    return JSON.stringify(objs, null, pretty ? 2 : undefined);
  };
  const ta = h('textarea', { class: 'ql-macro-code', rows: 14, readonly: 'true', 'aria-label': 'JSON' }) as HTMLTextAreaElement;
  const r = await openDialog({
    title: 'Export as JSON', wide: true, help: `${rangeName(rg)} on ${s.name}`, content: ta, ok: 'Download', cancel: 'Close',
    fields: [
      { key: 'header', label: 'First row has the field names', type: 'checkbox', value: header, half: true },
      { key: 'iso', label: 'Dates as ISO text (2024-03-05)', type: 'checkbox', value: true, half: true },
      { key: 'pretty', label: 'Indented', type: 'checkbox', value: true, half: true },
    ],
    buttons: [{ label: 'Copy', onClick: async () => { try { await navigator.clipboard.writeText(ta.value); toast('JSON copied'); } catch { ta.select(); toast('Press Ctrl+C to copy'); } } }],
    onChange: (v) => { ta.value = make(!!v.header, !!v.iso, !!v.pretty); },
  });
  if (!r) return;
  download(new Blob([ta.value], { type: 'application/json' }), `${s.name.replace(/[\\/:*?"<>|]+/g, '_')}.json`);
}

async function importJson(app: App, arg?: string | any[]): Promise<void> {
  let data: any = arg;
  if (data == null) {
    const r = await openDialog({ title: 'Import JSON', wide: true, help: 'Paste JSON (an array of records) or choose a .json file. The data is written at the active cell with field names in the first row.', fields: [
      { key: 'file', label: 'File', type: 'file', accept: '.json,application/json' },
      { key: 'text', label: 'Or paste JSON', type: 'textarea', rows: 8, value: '' },
    ] });
    if (!r) return;
    const text = r.file ? await (r.file as File).text() : String(r.text || '');
    if (!text.trim()) return;
    try { data = JSON.parse(text); } catch (e) { toast(`That is not valid JSON: ${(e as Error).message}`); return; }
  } else if (typeof data === 'string') { try { data = JSON.parse(data); } catch { toast('That is not valid JSON.'); return; } }
  let rows: Scalar[][];
  try { rows = jsonToRows(data); } catch (e) { toast((e as Error).message); return; }
  if (!rows.length) { toast('The JSON has no records.'); return; }
  const s = app.sheet;
  const { r: r0, c: c0 } = activeCell(app);
  const width = Math.max(...rows.map((x) => x.length));
  const target = { r1: r0, c1: c0, r2: r0 + rows.length - 1, c2: c0 + width - 1 };
  if (!checkWritable(app, [target])) return;
  let busy = false;
  for (let r = target.r1; r <= Math.min(target.r2, target.r1 + 2000) && !busy; r++) for (let c = target.c1; c <= target.c2; c++) { const x = s.get(r, c); if (x && (x.v != null || x.f != null)) { busy = true; break; } }
  if (busy && !(await confirmDialog('Replace Contents?', `There's already data in ${rangeName(target)}. Replace it?`, 'Replace'))) return;
  app.book.transact('Import JSON', () => {
    rows.forEach((row, i) => row.forEach((v, j) => {
      if (i === 0) { writeOut(s, r0, c0 + j, v == null ? null : { v, style: { bold: true } }); return; }
      if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}(T[\d:.]+Z?)?$/.test(v)) { const p = parseInput(app, v.replace('T', ' ').replace(/Z$/, '').replace(/\.\d+$/, ''), r0 + i, c0 + j, s); if (p.cell.v != null && typeof p.cell.v === 'number') { writeOut(s, r0 + i, c0 + j, { v: p.cell.v, style: { numFmt: p.format || 'yyyy-mm-dd' } }); return; } }
      writeOut(s, r0 + i, c0 + j, v);
    }));
  });
  app.grid.select(r0, c0, { range: target });
  toast(`${rows.length - 1} record${rows.length === 2 ? '' : 's'} imported`);
}

// ============================================================================================
// Commands

const D = 'Developer';

def([
  {
    id: 'dev.macroEditor', label: 'Macro Editor', icon: 'code_blocks', tab: D, group: 'Code', keys: ['Alt-F11'], noRecord: true, popular: true,
    tip: 'Write JavaScript macros stored in this workbook, with a small API for cells, ranges and commands', keywords: 'vba visual basic script code editor javascript module',
    run: (app, arg?: string) => macroEditor(app, typeof arg === 'string' ? arg : undefined),
  },
  { id: 'dev.viewCode', label: 'View Code', icon: 'code', tab: D, group: 'Controls', noRecord: true, tip: 'Open the macro code of this workbook (or of the selected button’s macro)', run: (app) => { const d = app.grid.objects.selected ? findControl(app, app.grid.objects.selected) : undefined; const mod = d?.control?.macro?.split('.')[0]; return macroEditor(app, mod); } },
  {
    id: 'dev.macros', label: 'Macros', icon: 'smart_button', tab: D, group: 'Code', keys: ['Alt-F8'], noRecord: true, popular: true,
    tip: 'Run, step through, edit or delete macros; set their shortcut keys', keywords: 'run macro list',
    run: (app) => macrosDialog(app),
  },
  {
    id: 'dev.recordMacro', label: 'Record Macro', icon: 'fiber_manual_record', tab: D, group: 'Code', kind: 'toggle', noRecord: true,
    tip: 'Record the commands you use and replay them later; click again to stop', keywords: 'record macro recorder',
    run: (app) => startRecording(app), isActive: (app) => !!app.recording,
  },
  { id: 'dev.stopRecording', label: 'Stop Recording', icon: 'stop_circle', tab: D, group: 'Code', noRecord: true, tip: 'Stop recording and save the macro', run: (app) => stopRecording(app), isEnabled: (app) => !!app.recording },
  { id: 'dev.pauseRecording', label: 'Pause Recording', icon: 'pause_circle', tab: D, group: 'Code', kind: 'toggle', noRecord: true, tip: 'Pause or resume the macro recorder', run: (app) => { if (!app.recording) { toast('No macro is being recorded.'); return; } app.recordingPaused = !app.recordingPaused; toast(app.recordingPaused ? 'Recording paused' : 'Recording resumed'); }, isActive: (app) => app.recordingPaused },
  {
    id: 'dev.relativeRefs', label: 'Use Relative References', icon: 'my_location', tab: D, group: 'Code', kind: 'toggle', noRecord: true,
    tip: 'Record cell selections relative to the starting cell, so the macro works from wherever it is run',
    run: (app) => { relativeRefs = !relativeRefs; lsSet(REL_KEY, relativeRefs ? '1' : '0'); if (recMeta) recMeta.relative = relativeRefs; toast(relativeRefs ? 'Relative references on' : 'Absolute references'); app.refresh(); },
    isActive: () => relativeRefs,
  },
  {
    id: 'dev.macroSecurity', label: 'Macro Security', icon: 'gpp_maybe', tab: D, group: 'Code', noRecord: true, kind: 'menu',
    tip: 'Choose whether JavaScript macros may run: disabled, ask for each workbook, or enabled', keywords: 'trust center security enable macros',
    run: async (app, arg?: MacroSecurity) => {
      if (!arg) {
        const r = await openDialog({ title: 'Macro Security', fields: [{ key: 'sec', label: 'JavaScript macros', type: 'radio', value: macroSecurity(), options: [{ value: 'disable', label: 'Disable all macros' }, { value: 'ask', label: 'Ask before enabling macros in each workbook (recommended)' }, { value: 'enable', label: 'Enable all macros (not recommended; macros can run any code)' }] }, { key: 'i', type: 'info', html: '<p class="ql-help">Recorded macros only replay commands and always run. This setting applies to this device.</p>' }] });
        if (!r) return;
        arg = r.sec;
      }
      lsSet(SEC_KEY, arg!);
      toast(`Macro security: ${{ disable: 'all macros disabled', ask: 'ask for each workbook', enable: 'all macros enabled' }[arg!]}`);
      app.refresh();
    },
    menu: (): MenuEntry[] => [{ id: 'dev.macroSecurity', arg: 'disable', label: 'Disable All Macros', checked: macroSecurity() === 'disable' }, { id: 'dev.macroSecurity', arg: 'ask', label: 'Ask for Each Workbook', checked: macroSecurity() === 'ask' }, { id: 'dev.macroSecurity', arg: 'enable', label: 'Enable All Macros', checked: macroSecurity() === 'enable' }],
  },
  {
    id: 'dev.runMacro', label: 'Run Macro', icon: 'play_circle', tab: D, group: 'Code', internal: true, noRecord: true,
    run: async (app, name?: string) => {
      if (!name) {
        const list = listMacros(app);
        if (!list.length) { toast('There are no macros. Record one or open the Macro Editor.'); return; }
        const r = await openDialog({ title: 'Run Macro', fields: [{ key: 'n', label: 'Macro', type: 'list', rows: 8, options: list.map((m) => ({ value: m.name, label: m.label })), value: list[0].name }] });
        if (!r?.n) return;
        name = r.n;
      }
      await runMacroByName(app, name!);
    },
  },
  {
    id: 'dev.insertControl', label: 'Insert Control', icon: 'add_box', tab: D, group: 'Controls', kind: 'menu', edits: true,
    tip: 'Add a button, check box, option button, drop-down, list box, spin button, scroll bar, label or group box',
    keywords: 'form control button checkbox combo list spinner scroll',
    run: (app, arg?: FormControlType) => insertControl(app, arg),
    menu: () => CONTROL_TYPES.map((t) => ({ id: 'dev.insertControl', arg: t.type, label: t.label, icon: t.icon })),
  },
  {
    id: 'dev.designMode', label: 'Design Mode', icon: 'design_services', tab: D, group: 'Controls', kind: 'toggle', noRecord: true,
    tip: 'Select, move and resize controls instead of using them',
    run: (app) => { designMode = !designMode; if (!designMode && app.grid.objects.selected && findControl(app, app.grid.objects.selected)) app.grid.objects.select(null); app.grid.objects.position(); toast(designMode ? 'Design Mode on: click a control to select it' : 'Design Mode off: controls respond to clicks'); app.grid.invalidate(); },
    isActive: () => designMode,
  },
  { id: 'dev.properties', label: 'Properties', icon: 'tune', tab: D, group: 'Controls', edits: true, tip: 'Cell link, input range, limits and macro of the selected control', keywords: 'format control properties link', run: (app, arg?: { id: string }) => controlProperties(app, arg?.id) },
  {
    id: 'dev.controlClick', label: 'Use Control', icon: 'touch_app', tab: D, group: 'Controls', internal: true, edits: true,
    run: async (app, arg: { id: string; part?: string | number; x?: number; y?: number }) => {
      const d = findControl(app, arg?.id);
      if (!d?.control) return;
      const ctl = d.control;
      if (ctl.type === 'button') { if (ctl.macro) await runMacroByName(app, ctl.macro); else toast('Assign a macro to this button (Developer ▸ Design Mode, then double-click it).'); return; }
      if (ctl.type === 'label' || ctl.type === 'group') return;
      let part = arg.part;
      if (ctl.type === 'combo' && part == null) {
        const items = controlItems(app.sheet, ctl);
        if (!items.length) { toast('Set the input range of this drop-down (Properties).'); return; }
        const cur = controlValue(app.sheet, ctl);
        part = await new Promise<number | undefined>((resolve) => {
          showMenu(items.map((t, i) => ({ id: 'pick', arg: i + 1, label: t || '(blank)', checked: cur === i + 1 })), { x: arg.x ?? 100, y: arg.y ?? 100, onPick: (e) => resolve(e.arg) });
          setTimeout(() => resolve(undefined), 120000);
        });
        if (part == null) return;
      }
      try { if (controlAction(app.sheet, d, part).changed && ctl.macro) await runMacroByName(app, ctl.macro); } catch (e) { toast((e as Error).message); }
    },
  },
  { id: 'dev.exportJson', label: 'Export JSON', icon: 'data_object', tab: D, group: 'Data', noRecord: true, tip: 'Save the selection or the current table as a JSON array of records, using the header row as field names', keywords: 'json export api records', run: (app) => exportJson(app) },
  { id: 'dev.importJson', label: 'Import JSON', icon: 'file_open', tab: D, group: 'Data', edits: true, tip: 'Read a JSON array of records into the sheet at the active cell, with field names as headers', keywords: 'json import api records', run: (app, arg?: string | any[]) => importJson(app, arg) },
  {
    id: 'dev.exportMacros', label: 'Export Macros', icon: 'file_export', tab: D, group: 'Code', noRecord: true, tip: 'Save recorded macros and workbook modules to a file',
    run: (app) => { download(new Blob([JSON.stringify({ macros: app.prefs.macros, modules: app.book.modules }, null, 2)], { type: 'application/json' }), 'macros.json'); },
  },
  {
    id: 'dev.importMacros', label: 'Import Macros', icon: 'upload_file', tab: D, group: 'Code', noRecord: true, tip: 'Load recorded macros and modules from a file exported by QuillLattice',
    run: async (app) => {
      const [f] = await pickFile('.json,application/json');
      if (!f) return;
      try {
        const j = JSON.parse(await f.text());
        const ms = (Array.isArray(j.macros) ? j.macros : []).filter((m: any) => m && typeof m.name === 'string' && Array.isArray(m.steps) && m.steps.every((s: any) => s && typeof s.id === 'string'));
        const mods = (Array.isArray(j.modules) ? j.modules : []).filter((m: any) => m && typeof m.name === 'string' && typeof m.code === 'string');
        for (const m of ms) saveMacro(app, { ...m, created: Number(m.created) || Date.now() });
        if (mods.length && !app.isReadOnly()) app.book.transact('Import Macros', () => { app.book.recordBook(); app.book.modules = [...app.book.modules.filter((x) => !mods.some((y: any) => y.name === x.name)), ...mods.map((m: any) => ({ name: m.name, code: m.code }))]; });
        toast(`${ms.length} recorded macro${ms.length === 1 ? '' : 's'} and ${mods.length} module${mods.length === 1 ? '' : 's'} imported`);
      } catch (e) { toast(`Could not import: ${(e as Error).message}`); }
    },
  },
  {
    id: 'dev.commandIds', label: 'Copy Command ID List', icon: 'content_copy', tab: D, group: 'Code', noRecord: true, tip: 'Copy every command ID and name, for macros (api.exec) and settings files',
    run: async () => { try { await navigator.clipboard.writeText(allCommands().map((c) => `${c.id}\t${c.label}\t${c.tab}`).join('\n')); toast('Command list copied'); } catch { toast('The clipboard is not available.'); } },
  },
]);
