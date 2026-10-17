/** Application preferences (per device), persisted to localStorage. */

export interface CustomGroup { id: string; label: string; items: string[] }
export interface CustomTab { id: string; label: string; groups: CustomGroup[]; hidden?: boolean }
export interface RecentEntry { key: string; name: string; when: number; uri?: string; format?: string; pinned?: boolean }
export interface Macro { name: string; steps: { id: string; arg?: any }[]; description?: string; key?: string; created: number; relative?: boolean }
export interface CustomList { items: string[] }

export const DEFAULT_REPLACEMENTS: Record<string, string> = {
  '(c)': '©', '(r)': '®', '(tm)': '™', '...': '…', '->': '→', '<-': '←', '=>': '⇒', '<=>': '⇔', '+-': '±', '!=': '≠',
  teh: 'the', adn: 'and', recieve: 'receive', seperate: 'separate', occured: 'occurred', untill: 'until', wich: 'which', thier: 'their', accomodate: 'accommodate', definately: 'definitely',
};

export interface Prefs {
  author: string;
  initials: string;
  qat: string[];
  qatPosition: 'top' | 'bottom';
  qatLabels: boolean;
  customTabs: CustomTab[];
  hiddenTabs: string[];
  tabOrder: string[];
  keymap: Record<string, string>;
  autosave: boolean;
  autosaveSeconds: number;
  replacements: Record<string, string>;
  autoCorrect: boolean;
  theme: 'system' | 'light' | 'dark';
  darkSheet: boolean;
  ribbon: 'full' | 'tabs' | 'hidden';
  spellcheck: boolean;
  recent: RecentEntry[];
  macros: Macro[];
  dictionary: string[];
  ignoredWords: string[];
  defaultFont: string;
  defaultSize: number;
  language: string;
  spellDialect: 'any' | 'document' | 'en-US' | 'en-GB' | 'en-CA' | 'en-AU';
  confirmClose: boolean;
  defaultSaveFormat: string;
  readAloudRate: number;
  highContrastUI: boolean;
  uiScale: number;
  statusItems: string[];
  clipboardHistory: number;
  pinnedCommands: string[];
  firstRun: boolean;
  /** Spreadsheet editing options */
  newSheets: number;
  enterMoves: boolean;
  enterDirection: 'down' | 'right' | 'up' | 'left';
  editInCell: boolean;
  autoComplete: boolean;
  fillHandle: boolean;
  extendFormats: boolean;
  autoPercent: boolean;
  flashFill: boolean;
  showFormulaBar: boolean;
  showFunctionTips: boolean;
  dateOrder: 'mdy' | 'dmy' | 'ymd';
  decimalSeparator: string;
  thousandsSeparator: string;
  useSystemSeparators: boolean;
  fixedDecimals: number;
  r1c1: boolean;
  zoomOnRoll: boolean;
  customLists: CustomList[];
  miniToolbar: boolean;
  quickAnalysis: boolean;
  touchHandles: boolean;
  /** Touch/Mouse Mode: more space between commands for fingers. */
  touchMode: boolean;
  /** Hide Quick Access Bar. */
  qatHidden: boolean;
  /** Lotus 1-2-3 formula entry: @SUM(A1..A5), +A1*@AVG(B1..B3), #AND# are translated when typed. */
  lotusFormulas: boolean;
  /** Lotus 1-2-3 label prefixes: "right, ^centered, \repeating (' is always left-aligned text). */
  lotusLabels: boolean;
  /** The / key opens the slash menu when a cell is selected and not being edited. */
  slashKey: boolean;
}

export const DEFAULT_QAT = ['file.save', 'edit.undo', 'edit.redo', 'file.saveAs', 'file.print', 'view.cmdSearch'];

/**
 * The Quick Access Bar drop-down checklist: the desktop spreadsheet's standard choices (AutoSave, New,
 * Open, Save, Email, Quick Print, Print Preview and Print, Spelling, Undo, Redo, Sort Ascending,
 * Sort Descending, Touch/Mouse Mode), then frequently used extras.
 */
export const QAT_STANDARD = ['file.autosave', 'file.new', 'file.open', 'file.save', 'file.email', 'file.quickPrint', 'file.print', 'file.printPreview', 'review.spelling', 'edit.undo', 'edit.redo', 'data.sortAZ', 'data.sortZA', 'view.touchMode'];
export const QAT_EXTRAS = ['file.saveAs', 'file.share', 'data.filter', 'edit.autoSum', 'edit.find', 'formulas.calcNow', 'insert.chartColumn', 'view.freezePanes', 'fmt.borders', 'view.cmdSearch'];
export const QAT_CHOICES = [...QAT_STANDARD, ...QAT_EXTRAS];
/** Kept first, in this order, by addAllQatChoices. */
export const QAT_PINNED_FIRST = ['file.save', 'file.print'];

/**
 * Adds every drop-down choice to the bar: Save and Print first, then the remaining choices sorted by
 * label (case-insensitive), then anything else already on the bar, in its existing order.
 */
export function addAllQatChoices(current: string[], labelOf: (id: string) => string | undefined, choices: string[] = QAT_CHOICES): string[] {
  const known = choices.filter((id) => labelOf(id) !== undefined);
  const first = QAT_PINNED_FIRST.filter((id) => known.includes(id));
  const rest = known.filter((id) => !first.includes(id))
    .sort((a, b) => labelOf(a)!.localeCompare(labelOf(b)!, undefined, { sensitivity: 'base' }));
  const head = [...first, ...rest];
  return [...head, ...current.filter((id) => id === '|' || !head.includes(id))];
}

export const BUILTIN_LISTS: CustomList[] = [
  { items: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'] },
  { items: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'] },
  { items: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'] },
  { items: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'] },
  { items: ['Q1', 'Q2', 'Q3', 'Q4'] },
];

export function defaultPrefs(): Prefs {
  return {
    author: 'Author',
    initials: 'A',
    qat: [...DEFAULT_QAT],
    qatPosition: 'top',
    qatLabels: false,
    customTabs: [],
    hiddenTabs: [],
    tabOrder: [],
    keymap: {},
    autosave: true,
    autosaveSeconds: 10,
    replacements: { ...DEFAULT_REPLACEMENTS },
    autoCorrect: true,
    theme: 'system',
    darkSheet: false,
    ribbon: 'full',
    spellcheck: true,
    recent: [],
    macros: [],
    dictionary: [],
    ignoredWords: [],
    defaultFont: '',
    defaultSize: 0,
    language: 'en-US',
    spellDialect: 'any',
    confirmClose: true,
    defaultSaveFormat: 'xlsx',
    readAloudRate: 1,
    highContrastUI: false,
    uiScale: 1,
    statusItems: ['average', 'count', 'numCount', 'min', 'max', 'sum', 'calc', 'views', 'zoom'],
    clipboardHistory: 24,
    pinnedCommands: [],
    firstRun: true,
    newSheets: 1,
    enterMoves: true,
    enterDirection: 'down',
    editInCell: true,
    autoComplete: true,
    fillHandle: true,
    extendFormats: true,
    autoPercent: true,
    flashFill: true,
    showFormulaBar: true,
    showFunctionTips: true,
    dateOrder: 'mdy',
    decimalSeparator: '.',
    thousandsSeparator: ',',
    useSystemSeparators: true,
    fixedDecimals: 0,
    r1c1: false,
    zoomOnRoll: false,
    customLists: [],
    miniToolbar: true,
    quickAnalysis: true,
    touchHandles: true,
    touchMode: false,
    qatHidden: false,
    lotusFormulas: true,
    lotusLabels: false,
    slashKey: false,
  };
}

const KEY = 'quilllattice.prefs.v1';

export function loadPrefs(): Prefs {
  const d = defaultPrefs();
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return d;
    const p = sanitizePrefs(JSON.parse(raw));
    return { ...d, ...p, replacements: p.replacements || d.replacements };
  } catch {
    return d;
  }
}

const ENUMS: Partial<Record<keyof Prefs, readonly string[]>> = {
  qatPosition: ['top', 'bottom'], theme: ['system', 'light', 'dark'], ribbon: ['full', 'tabs', 'hidden'],
  enterDirection: ['down', 'right', 'up', 'left'], dateOrder: ['mdy', 'dmy', 'ymd'],
  spellDialect: ['any', 'document', 'en-US', 'en-GB', 'en-CA', 'en-AU'],
};
const isObj = (v: unknown): v is Record<string, any> => !!v && typeof v === 'object' && !Array.isArray(v);
const strings = (v: unknown): string[] => (Array.isArray(v) ? v.filter((x): x is string => typeof x === 'string') : []);

/** Keeps only preferences of the expected shape (stored, or pasted into Import Settings). Unknown keys are dropped. */
export function sanitizePrefs(raw: unknown): Partial<Prefs> {
  if (!isObj(raw)) return {};
  const d = defaultPrefs() as Record<string, any>;
  const out: Record<string, any> = {};
  for (const [k, v] of Object.entries(raw)) {
    if (!(k in d)) continue;
    const def = d[k];
    if (Array.isArray(def) ? Array.isArray(v) : isObj(def) ? isObj(v) : typeof v === typeof def && (typeof v !== 'number' || Number.isFinite(v))) out[k] = v;
  }
  for (const [k, ok] of Object.entries(ENUMS)) if (k in out && !ok!.includes(out[k])) delete out[k];
  for (const k of ['qat', 'hiddenTabs', 'tabOrder', 'dictionary', 'ignoredWords', 'statusItems', 'pinnedCommands']) if (k in out) out[k] = strings(out[k]);
  for (const k of ['keymap', 'replacements']) if (k in out) out[k] = Object.fromEntries(Object.entries(out[k]).filter(([, x]) => typeof x === 'string'));
  if (out.customTabs) {
    out.customTabs = out.customTabs.filter((t: any) => isObj(t) && typeof t.id === 'string' && Array.isArray(t.groups))
      .map((t: any) => ({ ...t, label: String(t.label ?? t.id), groups: t.groups.filter(isObj).map((g: any) => ({ ...g, id: String(g.id ?? ''), label: String(g.label ?? ''), items: strings(g.items) })) }));
  }
  if (out.recent) out.recent = out.recent.filter((r: any) => isObj(r) && typeof r.key === 'string' && typeof r.name === 'string');
  if (out.macros) out.macros = out.macros.filter((m: any) => isObj(m) && typeof m.name === 'string' && Array.isArray(m.steps) && m.steps.every((s: any) => isObj(s) && typeof s.id === 'string'));
  if (out.customLists) out.customLists = out.customLists.filter((l: any) => isObj(l) && Array.isArray(l.items)).map((l: any) => ({ items: strings(l.items) }));
  return out as Partial<Prefs>;
}

export function savePrefs(p: Prefs): void {
  try { localStorage.setItem(KEY, JSON.stringify(p)); } catch { /* storage full or unavailable */ }
}
