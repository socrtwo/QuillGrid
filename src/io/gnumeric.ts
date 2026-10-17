/**
 * Gnumeric workbooks (.gnumeric, .gnm): gzip-compressed (or plain) XML in the Gnumeric
 * namespace (http://www.gnumeric.org/v10.dtd; older files use the gmr prefix and
 * http://www.gnome.org/gnumeric/v2..v8 namespaces).
 *
 * Reads sheets, cells (strings, numbers, booleans, errors, formulas including shared expressions
 * by ExprID and array formulas), style regions (fonts, colors, fills, borders, alignment, number
 * formats, protection), validations, hyperlinks and input messages attached to style regions,
 * conditional formats, column widths and row heights (hidden, outline levels), merged regions,
 * frozen panes, cell comments, workbook and sheet names, tab colors, zoom, gridlines and the
 * 1904 date system.
 *
 * Gnumeric stores formulas in an Excel-compatible A1 syntax with ',' separators; cached results
 * are not stored, so formulas are calculated when the workbook opens.
 */
import { Workbook, Sheet, newId } from '../model/workbook';
import { CellError, errorFromCode, type Cell, type CellStyle, type Border, type CFRule, type DataValidation } from '../model/types';
import { cellKey, parseRange, splitSheetRef, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { shiftFormula } from '../calc/parser';
import { parseXml, kids, kid, desc, attr, decodeText } from './index';
import { gunzip, isGzip } from './gzip';

/** Gnumeric color "RRRR:GGGG:BBBB" (16-bit hex channels) -> #rrggbb. */
export function gnmColor(s: string | null): string | undefined {
  if (!s) return undefined;
  const p = s.split(':');
  if (p.length !== 3) return undefined;
  return '#' + p.map((x) => {
    const n = parseInt(x, 16);
    const v = x.length > 2 ? n >> 8 : n;
    return Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0');
  }).join('');
}

const H_ALIGN: Record<string, CellStyle['hAlign']> = {
  GNM_HALIGN_LEFT: 'left', GNM_HALIGN_RIGHT: 'right', GNM_HALIGN_CENTER: 'center', GNM_HALIGN_FILL: 'fill', GNM_HALIGN_JUSTIFY: 'justify',
  GNM_HALIGN_CENTER_ACROSS_SELECTION: 'centerContinuous', GNM_HALIGN_DISTRIBUTED: 'distributed',
  HALIGN_LEFT: 'left', HALIGN_RIGHT: 'right', HALIGN_CENTER: 'center', HALIGN_FILL: 'fill', HALIGN_JUSTIFY: 'justify', HALIGN_CENTER_ACROSS_SELECTION: 'centerContinuous',
  2: 'left', 4: 'right', 8: 'center', 16: 'fill', 32: 'justify', 64: 'centerContinuous', 128: 'distributed',
};
const V_ALIGN: Record<string, CellStyle['vAlign']> = {
  GNM_VALIGN_TOP: 'top', GNM_VALIGN_CENTER: 'middle', GNM_VALIGN_JUSTIFY: 'justify', GNM_VALIGN_DISTRIBUTED: 'distributed',
  VALIGN_TOP: 'top', VALIGN_CENTER: 'middle', VALIGN_JUSTIFY: 'justify', 1: 'top', 4: 'middle', 8: 'justify', 16: 'distributed',
};
const BORDER: Border['style'][] = ['thin', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'dashDot', 'dashDot', 'mediumDashed', 'dashDot'];

const on = (v: string | null) => v === '1' || v === 'true' || v === 'TRUE';

function readStyle(st: Element): { style: CellStyle; extras: Element[] } {
  const s: CellStyle = {};
  const ha = attr(st, 'HAlign');
  if (ha && H_ALIGN[ha]) s.hAlign = H_ALIGN[ha];
  const va = attr(st, 'VAlign');
  if (va && V_ALIGN[va]) s.vAlign = V_ALIGN[va];
  if (on(attr(st, 'WrapText'))) s.wrap = true;
  if (on(attr(st, 'ShrinkToFit'))) s.shrink = true;
  const rot = +(attr(st, 'Rotation') || 0);
  if (rot) s.rotate = rot === -1 ? 255 : rot > 90 && rot <= 180 ? 90 : rot > 180 ? Math.max(-90, rot - 360) : rot;
  const ind = +(attr(st, 'Indent') || 0); if (ind) s.indent = ind;
  if (attr(st, 'Locked') === '0') s.locked = false;
  if (on(attr(st, 'Hidden'))) s.hidden = true;
  const fore = gnmColor(attr(st, 'Fore'));
  if (fore && fore !== '#000000') s.color = fore;
  const shade = +(attr(st, 'Shade') || 0);
  const back = gnmColor(attr(st, 'Back'));
  if (shade === 1 && back) s.fill = back;
  else if (shade > 1 && back) { s.fill = back; s.pattern = { type: shade <= 3 ? 'gray125' : 'lightGray', ...(gnmColor(attr(st, 'PatternColor')) ? { color: gnmColor(attr(st, 'PatternColor')) } : {}) }; }
  const fmt = attr(st, 'Format');
  if (fmt && fmt !== 'General') s.numFmt = fmt;
  const font = kid(st, 'Font');
  if (font) {
    const name = (font.textContent || '').trim();
    if (name && name !== 'Sans') s.font = name;
    const size = +(attr(font, 'Unit') || 0);
    if (size && size !== 10) s.size = size;
    if (on(attr(font, 'Bold'))) s.bold = true;
    if (on(attr(font, 'Italic'))) s.italic = true;
    if (on(attr(font, 'StrikeThrough'))) s.strike = true;
    const u = attr(font, 'Underline');
    if (u === '1' || u === 'single' || u === 'UNDERLINE_SINGLE') s.underline = 'single';
    else if (u === '2' || u === 'double' || u === 'UNDERLINE_DOUBLE') s.underline = 'double';
    else if (u === '3') s.underline = 'singleAccounting';
    else if (u === '4') s.underline = 'doubleAccounting';
    const sc = attr(font, 'Script');
    if (sc === '1') s.vertAlign = 'superscript'; else if (sc === '-1') s.vertAlign = 'subscript';
  }
  const sb = kid(st, 'StyleBorder');
  if (sb) {
    const side = (n: string): Border | undefined => {
      const e = kid(sb, n);
      const k = +(attr(e, 'Style') || 0);
      if (!e || !k) return undefined;
      const color = gnmColor(attr(e, 'Color'));
      return { style: BORDER[k] || 'thin', ...(color && color !== '#000000' ? { color } : {}) };
    };
    const t = side('Top'), b = side('Bottom'), l = side('Left'), r = side('Right'), d = side('Diagonal'), rd = side('Rev-Diagonal');
    if (t) s.bt = t; if (b) s.bb = b; if (l) s.bl = l; if (r) s.br = r; if (d) s.diagUp = d; if (rd) s.diagDown = rd;
  }
  return { style: s, extras: kids(st).filter((e) => ['Validation', 'HyperLink', 'InputMsg', 'Condition', 'Conditions'].includes(e.localName)) };
}

const VT_DV: Record<string, DataValidation['type']> = { 0: 'any', 1: 'whole', 2: 'decimal', 3: 'list', 4: 'date', 5: 'time', 6: 'textLength', 7: 'custom' };
const OPS: DataValidation['operator'][] = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];

/** Gnumeric expressions are A1 with ',' (some locales ';'); strip a leading '='. */
function expr(s: string | null | undefined): string {
  return (s || '').trim().replace(/^=/, '');
}

function readExtras(sheet: Sheet, rg: Range, extras: Element[], warnings: string[]): void {
  for (const e of extras) {
    switch (e.localName) {
      case 'HyperLink': {
        const target = attr(e, 'target') || '';
        const type = attr(e, 'type') || '';
        if (!target) break;
        const link = /CurWB/.test(type) ? '#' + target : /Email/.test(type) && !/^mailto:/i.test(target) ? 'mailto:' + target : target;
        const tip = attr(e, 'tip');
        for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 1000); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) {
          const k = cellKey(r, c);
          sheet.cells.set(k, { ...(sheet.cells.get(k) || {}), link, ...(tip ? { linkTip: tip } : {}) });
        }
        break;
      }
      case 'Validation': {
        // Gnumeric's "list" validation is type 3 (in), custom/expression is 7 in newer files.
        const typeAttr = attr(e, 'Type') || '0';
        const type = /^GNM_VALIDATION_TYPE_/.test(typeAttr) ? ({ GNM_VALIDATION_TYPE_ANY: 'any', GNM_VALIDATION_TYPE_AS_INT: 'whole', GNM_VALIDATION_TYPE_AS_NUMBER: 'decimal', GNM_VALIDATION_TYPE_IN_LIST: 'list', GNM_VALIDATION_TYPE_AS_DATE: 'date', GNM_VALIDATION_TYPE_AS_TIME: 'time', GNM_VALIDATION_TYPE_TEXT_LENGTH: 'textLength', GNM_VALIDATION_TYPE_CUSTOM: 'custom' } as Record<string, DataValidation['type']>)[typeAttr] || 'any' : VT_DV[typeAttr] || 'any';
        const f1 = expr(kid(e, 'Expression0')?.textContent), f2 = expr(kid(e, 'Expression1')?.textContent);
        // A custom rule whose formula is a quoted, comma-separated literal is a list.
        const v: DataValidation = { id: newId('dv'), ranges: [rg], type: type === 'custom' && /^"[^"]*,[^"]*"$/.test(f1) ? 'list' : type };
        const op = +(attr(e, 'Operator') || 0);
        if (!['any', 'list', 'custom'].includes(v.type)) v.operator = OPS[op] || 'between';
        if (f1) v.formula1 = f1;
        if (f2 && (v.operator === 'between' || v.operator === 'notBetween')) v.formula2 = f2;
        v.allowBlank = attr(e, 'AllowBlank') !== 'false';
        if (v.type === 'list') v.showDropdown = attr(e, 'UseDropdown') !== 'false';
        const es = +(attr(e, 'Style') || 1);
        v.errorStyle = es === 2 ? 'warning' : es === 3 ? 'information' : 'stop';
        if (attr(e, 'Title')) v.errorTitle = attr(e, 'Title')!;
        if (attr(e, 'Message')) v.errorMessage = attr(e, 'Message')!;
        sheet.validations.push(v);
        break;
      }
      case 'InputMsg': {
        const dv = sheet.validations.find((x) => x.ranges.some((r) => r.r1 === rg.r1 && r.c1 === rg.c1 && r.r2 === rg.r2 && r.c2 === rg.c2));
        const target = dv || (sheet.validations[sheet.validations.push({ id: newId('dv'), ranges: [rg], type: 'any', allowBlank: true }) - 1]);
        if (attr(e, 'Title')) target.inputTitle = attr(e, 'Title')!;
        if (attr(e, 'Message')) target.inputMessage = attr(e, 'Message')!;
        break;
      }
      case 'Condition': case 'Conditions': {
        const conds = e.localName === 'Condition' ? [e] : kids(e, 'Condition');
        for (const c of conds) {
          const rule = readCondition(c, rg);
          if (rule) sheet.cf.push(rule); else if (!warnings.includes('Some conditional formats were not opened.')) warnings.push('Some conditional formats were not opened.');
        }
        break;
      }
    }
  }
}

/** Gnumeric conditional style operators (GnmStyleCondOp). */
const COND_OPS: Record<string, CFRule['operator']> = { 0: 'between', 1: 'notBetween', 2: 'equal', 3: 'notEqual', 4: 'greaterThan', 5: 'lessThan', 6: 'greaterThanOrEqual', 7: 'lessThanOrEqual' };

function readCondition(c: Element, rg: Range): CFRule | null {
  const op = attr(c, 'Operator') || '0';
  const st = kid(c, 'Style');
  const style = st ? readStyle(st).style : {};
  const exprs = kids(c).filter((x) => /^Expression\d$/.test(x.localName)).map((x) => expr(x.textContent));
  const r: CFRule = { id: newId('cf'), ranges: [rg], type: 'cellIs', style };
  if (COND_OPS[op]) { r.operator = COND_OPS[op]; r.formulas = exprs; return r; }
  // GnmStyleCondOp: 8 custom expression; 0x10.. text and blank/error tests.
  const text = () => exprs[0]?.replace(/^"|"$/g, '').replace(/""/g, '"');
  switch (op) {
    case '8': r.type = 'expression'; r.formulas = exprs; return r;
    case '16': r.type = 'containsText'; r.text = text(); return r;
    case '17': r.type = 'notContainsText'; r.text = text(); return r;
    case '18': r.type = 'beginsWith'; r.text = text(); return r;
    case '20': r.type = 'endsWith'; r.text = text(); return r;
    case '22': r.type = 'containsErrors'; return r;
    case '23': r.type = 'notContainsErrors'; return r;
    case '24': r.type = 'containsBlanks'; return r;
    case '25': r.type = 'notContainsBlanks'; return r;
  }
  return null;
}

function cellValue(vt: string | null, text: string): Cell['v'] {
  switch (vt) {
    case '10': return null;
    case '20': return /^(true|1)$/i.test(text.trim());
    case '30': case '40': { const n = Number(text); return Number.isFinite(n) ? n : text; }
    case '50': return errorFromCode(text.trim()) || new CellError('#VALUE!');
    case '60': return text;
    case '70': case '80': return text;
  }
  if (text === '') return null;
  const n = Number(text);
  return Number.isFinite(n) && /\d/.test(text) ? n : text;
}

export async function importGnumeric(data: ArrayBuffer | Uint8Array, warnings: string[] = []): Promise<Workbook> {
  let u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (isGzip(u8)) u8 = await gunzip(u8);
  return readGnumericXml(decodeText(u8), warnings);
}

export function readGnumericXml(text: string, warnings: string[] = []): Workbook {
  const doc = parseXml(text, 'Gnumeric workbook');
  const root = doc.documentElement;
  if (root.localName !== 'Workbook') throw new Error('This is not a Gnumeric workbook.');
  const book = new Workbook(false);
  // 1904 date system: workbook attribute or the newer DateConvention.
  for (const a of desc(root, 'Attribute')) {
    const n = kid(a, 'name')?.textContent || '';
    if (/date.?1904|DateConvention/i.test(n) && /true|1904/i.test(kid(a, 'value')?.textContent || '')) book.calc.date1904 = true;
  }
  if (/1904/.test(attr(kid(root, 'DateConvention'), 'value') || kid(root, 'DateConvention')?.textContent || '')) book.calc.date1904 = true;
  const calc = kid(root, 'Calculation');
  if (calc) {
    if (on(attr(calc, 'ManualRecalc'))) book.calc.mode = 'manual';
    if (on(attr(calc, 'EnableIteration'))) book.calc.iterative = true;
    const mi = +(attr(calc, 'MaxIterations') || 0); if (mi) book.calc.maxIterations = mi;
    const tol = +(attr(calc, 'IterationTolerance') || 0); if (tol) book.calc.maxChange = tol;
  }
  const sheetEls = kids(kid(root, 'Sheets'), 'Sheet');
  for (const se of sheetEls) {
    let name = (kid(se, 'Name')?.textContent || '').trim() || `Sheet${book.sheets.length + 1}`;
    while (book.sheetByName(name)) name += '_';
    const sheet = new Sheet(book, name);
    book.sheets.push(sheet);
    readSheet(book, sheet, se, warnings);
  }
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
  // Names: workbook level and per sheet.
  const addNames = (el: Element | null, sheet: string | null) => {
    for (const n of kids(el, 'Name')) {
      const nm = (kid(n, 'name')?.textContent || '').trim();
      const val = expr(kid(n, 'value')?.textContent);
      if (!nm || !val || /^#REF!$/.test(val)) continue;
      if (/^Print_Area$/i.test(nm)) { const sh = sheet ? book.sheetByName(sheet) : null; if (sh) sh.print.area = splitSheetRef(val).ref; continue; }
      book.names.push({ name: nm, ref: val, sheet });
    }
  };
  addNames(kid(root, 'Names'), null);
  sheetEls.forEach((se, i) => addNames(kid(se, 'Names'), book.sheets[i]?.name ?? null));
  const ui = kid(root, 'UIData');
  const sel = +(attr(ui, 'SelectedTab') || 0);
  if (sel >= 0 && sel < book.sheets.length) book.activeSheet = sel;
  const meta = desc(root, 'document-meta')[0];
  if (meta) {
    const t = (n: string) => (desc(meta, n)[0]?.textContent || '').trim();
    book.props.title = t('title'); book.props.subject = t('subject'); book.props.author = t('initial-creator') || t('creator'); book.props.comments = t('description');
  }
  book.clearHistory();
  return book;
}

function readSheet(book: Workbook, sheet: Sheet, se: Element, warnings: string[]): void {
  const vis = attr(se, 'Visibility') || '';
  if (/VERY_HIDDEN/.test(vis)) sheet.visibility = 'veryHidden'; else if (/HIDDEN/.test(vis)) sheet.visibility = 'hidden';
  if (on(attr(se, 'HideGrid'))) sheet.showGrid = false;
  if (on(attr(se, 'HideZero'))) sheet.showZeros = false;
  if (on(attr(se, 'DisplayFormulas'))) sheet.showFormulas = true;
  if (on(attr(se, 'RTL'))) sheet.rtl = true;
  const tab = gnmColor(attr(se, 'TabColor')); if (tab) sheet.tabColor = tab;
  const z = +(kid(se, 'Zoom')?.textContent || 0); if (z >= 0.1 && z <= 4) sheet.zoom = z;
  // Column and row geometry (points).
  const cols = kid(se, 'Cols');
  const dcw = +(attr(cols, 'DefaultSizePts') || 0); if (dcw) sheet.defaultColW = Math.round((dcw * 4) / 3);
  for (const ci of kids(cols, 'ColInfo')) {
    const no = +(attr(ci, 'No') || 0), count = Math.max(1, +(attr(ci, 'Count') || 1));
    const w = +(attr(ci, 'Unit') || 0);
    for (let c = no; c < Math.min(MAX_COLS, no + count, no + 1024); c++) {
      if (w && Math.round((w * 4) / 3) !== sheet.defaultColW) sheet.colW.set(c, Math.round((w * 4) / 3));
      if (on(attr(ci, 'Hidden'))) sheet.hiddenCols.add(c);
      const lvl = +(attr(ci, 'OutlineLevel') || 0); if (lvl) sheet.colLevel.set(c, lvl);
    }
  }
  const rows = kid(se, 'Rows');
  const drh = +(attr(rows, 'DefaultSizePts') || 0); if (drh) sheet.defaultRowH = Math.round((drh * 4) / 3);
  for (const ri of kids(rows, 'RowInfo')) {
    const no = +(attr(ri, 'No') || 0), count = Math.max(1, +(attr(ri, 'Count') || 1));
    const h = +(attr(ri, 'Unit') || 0);
    for (let r = no; r < Math.min(MAX_ROWS, no + count, no + 100000); r++) {
      if (h && Math.round((h * 4) / 3) !== sheet.defaultRowH) sheet.rowH.set(r, Math.round((h * 4) / 3));
      if (on(attr(ri, 'Hidden'))) sheet.hiddenRows.add(r);
      const lvl = +(attr(ri, 'OutlineLevel') || 0); if (lvl) sheet.rowLevel.set(r, lvl);
    }
  }
  // Cells.
  const shared = new Map<string, { f: string; r: number; c: number }>();
  for (const ce of kids(kid(se, 'Cells'), 'Cell')) {
    const r = +(attr(ce, 'Row') ?? -1), c = +(attr(ce, 'Col') ?? -1);
    if (!(r >= 0 && r < MAX_ROWS && c >= 0 && c < MAX_COLS)) continue;
    const content = kid(ce, 'Content');
    const text = content ? content.textContent || '' : kids(ce).length ? '' : ce.textContent || '';
    const cell: Cell = {};
    const id = attr(ce, 'ExprID');
    if (text.startsWith('=') && text.length > 1) {
      cell.f = text.slice(1);
      if (id) shared.set(id, { f: cell.f, r, c });
      const rs = +(attr(ce, 'Rows') || 0), cs = +(attr(ce, 'Cols') || 0);
      if (rs > 1 || cs > 1) cell.arr = { r1: r, c1: c, r2: r + Math.max(1, rs) - 1, c2: c + Math.max(1, cs) - 1 };
    } else if (id && !text) {
      const m = shared.get(id);
      if (m) cell.f = shiftFormula(m.f, r - m.r, c - m.c);
    } else {
      const v = cellValue(attr(ce, 'ValueType'), text);
      if (v != null && v !== '') cell.v = v;
      const vf = attr(ce, 'ValueFormat');
      if (vf && vf !== 'General') cell.s = book.styleId({ numFmt: vf });
    }
    if (cell.v != null || cell.f != null || cell.s) sheet.cells.set(cellKey(r, c), cell);
  }
  // Array formula results are calculated; mark the covered cells.
  for (const [k, cell] of sheet.cells) {
    if (!cell.arr) continue;
    const a = cell.arr;
    for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) { const k2 = cellKey(r, c); if (k2 !== k && !sheet.cells.get(k2)?.f) sheet.cells.set(k2, { sp: k }); }
  }
  // Style regions: applied after the cells so per-cell value formats can be merged in.
  const cache = new Map<Element, { sid: number; extras: Element[] }>();
  const regions = kids(kid(se, 'Styles'), 'StyleRegion');
  for (const sr of regions) {
    const st = kid(sr, 'Style');
    if (!st) continue;
    const rg: Range = { r1: +(attr(sr, 'startRow') || 0), c1: +(attr(sr, 'startCol') || 0), r2: Math.min(MAX_ROWS - 1, +(attr(sr, 'endRow') || 0)), c2: Math.min(MAX_COLS - 1, +(attr(sr, 'endCol') || 0)) };
    let entry = cache.get(st);
    if (!entry) { const { style, extras } = readStyle(st); entry = { sid: book.styleId(style), extras }; cache.set(st, entry); }
    if (entry.extras.length) readExtras(sheet, rg, entry.extras, warnings);
    applyStyle(sheet, rg, entry.sid);
  }
  // Merges.
  for (const m of kids(kid(se, 'MergedRegions'), 'Merge')) {
    const rg = parseRange(m.textContent || '');
    if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) sheet.merges.push(rg);
  }
  // Frozen panes and the cursor.
  const fp = desc(se, 'FreezePanes')[0];
  if (fp) {
    const a = parseRange(attr(fp, 'FrozenTopLeft') || 'A1'), b = parseRange(attr(fp, 'UnfrozenTopLeft') || 'A1');
    if (a && b) sheet.freeze = { rows: Math.max(0, b.r1 - a.r1), cols: Math.max(0, b.c1 - a.c1) };
  }
  const sels = kid(se, 'Selections');
  if (sels) {
    const r = +(attr(sels, 'CursorRow') || 0), c = +(attr(sels, 'CursorCol') || 0);
    sheet.sel = { r, c, ranges: [{ r1: r, c1: c, r2: r, c2: c }] };
  }
  // Comments.
  for (const cm of desc(kid(se, 'Objects'), 'CellComment')) {
    const rg = parseRange(attr(cm, 'ObjectBound') || '');
    if (!rg) continue;
    const text = attr(cm, 'Text') ?? (cm.textContent || '');
    sheet.comments.set(cellKey(rg.r1, rg.c1), { author: attr(cm, 'Author') || '', text, time: 0, note: true });
  }
  if (desc(kid(se, 'Objects'), 'SheetObjectGraph').length || desc(kid(se, 'Objects'), 'GnmGraph').length) {
    if (!warnings.includes('Charts in Gnumeric files are not opened.')) warnings.push('Charts in Gnumeric files are not opened.');
  }
  const prot = kid(se, 'SheetProtection') || kid(se, 'Protection');
  if (prot && on(attr(prot, 'Protected') ?? attr(prot, 'protected') ?? '1')) sheet.protection = { on: true, allow: ['selectLocked', 'selectUnlocked'] };
}

/** Applies a style to a region: every existing cell in it, and whole rows/columns for big regions. */
function applyStyle(sheet: Sheet, rg: Range, sid: number): void {
  const book = sheet.book;
  const merge = (cur: number | undefined): number => {
    if (!cur) return sid;
    if (!sid) return cur;
    // A cell's own value format (ValueFormat) wins over the region format.
    const own = book.styles[cur];
    return own?.numFmt && Object.keys(own).length === 1 ? book.derivedStyle(sid, { numFmt: own.numFmt }) : sid;
  };
  const wholeCols = rg.r1 === 0 && rg.r2 >= 65535;
  const wholeRows = rg.c1 === 0 && rg.c2 >= 255;
  if (wholeCols && wholeRows) {
    // Workbook-wide default style: only cells that exist take it (it is usually the default).
    if (sid) for (const [k, cell] of sheet.cells) cell.s = merge(cell.s);
    else for (const cell of sheet.cells.values()) if (cell.s && Object.keys(book.styles[cell.s] || {}).length > 1) delete cell.s;
    return;
  }
  if (wholeCols && rg.c2 - rg.c1 < 256) for (let c = rg.c1; c <= rg.c2; c++) { if (sid) sheet.colStyle.set(c, sid); else sheet.colStyle.delete(c); }
  else if (wholeRows && rg.r2 - rg.r1 < 10000) for (let r = rg.r1; r <= rg.r2; r++) { if (sid) sheet.rowStyle.set(r, sid); else sheet.rowStyle.delete(r); }
  const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
  if (area <= 20000) {
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      const k = cellKey(r, c);
      const cell = sheet.cells.get(k);
      if (cell) { const s = merge(cell.s); if (s) cell.s = s; else delete cell.s; }
      else if (sid && !wholeCols && !wholeRows) sheet.cells.set(k, { s: sid });
    }
  } else {
    for (const [k, cell] of sheet.cells) {
      const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS;
      if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) { const s = merge(cell.s); if (s) cell.s = s; else delete cell.s; }
    }
  }
}
