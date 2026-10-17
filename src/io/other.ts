/**
 * Legacy interchange formats (import only):
 *   SYLK (.slk)                 Multiplan/Excel symbolic link: values, R1C1 formulas, formats, widths
 *   DIF  (.dif)                 Data Interchange Format: values only
 *   XML Spreadsheet 2003 (.xml) Excel's XML format: values, R1C1 formulas, styles, merges, comments,
 *                               column widths, row heights, names, frozen panes, hidden sheets
 */
import { Workbook, Sheet, newId } from '../model/workbook';
import { CellError, errorFromCode, type Cell, type CellStyle, type Border, type CFRule, type DataValidation } from '../model/types';
import { cellKey, colName, parseRange, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { partsToSerial, parseDateTime } from '../model/dates';
import { parseXml, kids, kid, desc, attr, normColor } from './index';
import { odfToFormula } from './ods';

// --------------------------------------------------------------------------------------------
// R1C1 -> A1

/** Converts R1C1-style references in a formula (relative to row r, col c; zero-based) to A1. */
export function r1c1ToA1(f: string, r: number, c: number): string {
  const conv = (seg: string) => seg.replace(/(?<![A-Za-z0-9_.$\]])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(\[])/g, (_m, rs: string | undefined, cs: string | undefined) => {
    const part = (spec: string | undefined, base: number): { n: number; abs: boolean } => {
      if (!spec) return { n: base, abs: false };
      if (spec.startsWith('[')) return { n: base + parseInt(spec.slice(1, -1), 10), abs: false };
      return { n: parseInt(spec, 10) - 1, abs: true };
    };
    const R = part(rs, r), C = part(cs, c);
    if (R.n < 0 || C.n < 0 || R.n >= MAX_ROWS || C.n >= MAX_COLS) return '#REF!';
    return `${C.abs ? '$' : ''}${colName(C.n)}${R.abs ? '$' : ''}${R.n + 1}`;
  });
  // Whole rows / columns: R1:R3, C2:C4
  const rows = (seg: string) => seg.replace(/(?<![A-Za-z0-9_.$])R(\[-?\d+\]|\d+):R(\[-?\d+\]|\d+)(?![A-Za-z0-9_(\[])/g, (_m, a: string, b: string) => {
    const p = (s: string) => (s.startsWith('[') ? `${r + parseInt(s.slice(1, -1), 10) + 1}` : `$${s}`);
    return `${p(a)}:${p(b)}`;
  }).replace(/(?<![A-Za-z0-9_.$])C(\[-?\d+\]|\d+):C(\[-?\d+\]|\d+)(?![A-Za-z0-9_(\[])/g, (_m, a: string, b: string) => {
    const p = (s: string) => (s.startsWith('[') ? colName(c + parseInt(s.slice(1, -1), 10)) : '$' + colName(parseInt(s, 10) - 1));
    return `${p(a)}:${p(b)}`;
  });
  let out = '', i = 0, start = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') {
      out += conv(rows(f.slice(start, i)));
      let j = i + 1;
      while (j < f.length) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1);
      i = start = j + 1;
      continue;
    }
    i++;
  }
  return out + conv(rows(f.slice(start)));
}

// --------------------------------------------------------------------------------------------
// SYLK

function splitSylk(line: string): string[] {
  const out: string[] = [];
  let cur = '';
  for (let i = 0; i < line.length; i++) {
    if (line[i] === ';') {
      if (line[i + 1] === ';') { cur += ';'; i++; continue; }
      out.push(cur); cur = '';
      continue;
    }
    cur += line[i];
  }
  out.push(cur);
  return out;
}

function sylkValue(k: string): Cell['v'] {
  if (k.startsWith('"')) return k.replace(/^"|"$/g, '').replace(/""/g, '"');
  if (k === 'TRUE' || k === 'FALSE') return k === 'TRUE';
  if (k.startsWith('#')) return errorFromCode(k) || new CellError('#VALUE!');
  const n = Number(k);
  return Number.isFinite(n) ? n : k;
}

/** SYLK escapes: ESC " :" is a line break; ESC N + accent code + letter is an accented letter. */
const SYLK_ACCENTS: Record<string, string> = { A: '̀', B: '́', C: '̂', D: '̃', H: '̈', J: '̊', K: '̧' };
export function sylkUnescape(s: string): string {
  if (s.indexOf('\x1b') < 0) return s;
  return s.replace(/\x1b ?:/g, '\n').replace(/\x1bN([A-K])([A-Za-z])/g, (_m, acc: string, ch: string) => (ch + (SYLK_ACCENTS[acc] || '')).normalize('NFC')).replace(/\x1b./g, '');
}

/** SYLK F-field format letter + digits -> format code. */
function sylkFieldFormat(code: string): { fmt?: string; align?: CellStyle['hAlign'] } {
  const m = /^([DCEFG$*%])(\d*)([DCGLRX-])?/.exec(code);
  if (!m) return {};
  const n = +(m[2] || 0);
  const dec = n > 0 ? '.' + '0'.repeat(Math.min(n, 15)) : '';
  const fmt = m[1] === 'F' ? '0' + dec : m[1] === '$' ? '$#,##0' + (dec || '.00') : m[1] === '%' ? '0' + dec + '%' : m[1] === 'E' ? '0' + (dec || '.00') + 'E+00' : m[1] === 'C' ? '#,##0' + dec : undefined;
  const align = ({ L: 'left', R: 'right', C: 'center', X: 'fill' } as Record<string, CellStyle['hAlign']>)[m[3] || ''];
  return { fmt, align };
}

export function importSylk(text: string): Workbook {
  const book = new Workbook(false);
  const sheet = new Sheet(book, 'Sheet1');
  book.sheets.push(sheet);
  const formats: string[] = [];
  const fonts: { font?: string; size?: number; bold?: boolean; italic?: boolean }[] = [];
  let row = 0, col = 0;
  const colFmt = new Map<number, CellStyle>(), rowFmt = new Map<number, CellStyle>();
  const lines = text.split(/\r\n|\r|\n/);
  for (const raw of lines) {
    if (!raw) continue;
    const f = splitSylk(raw);
    const rt = f[0];
    if (rt === 'E') break;
    if (rt === 'P') {
      const p = f.slice(1).find((x) => x.startsWith('P'));
      if (p != null) formats.push(sylkUnescape(p.slice(1)).replace(/\\/g, ''));
      const e = f.slice(1).find((x) => x.startsWith('E') || x.startsWith('F'));
      if (e != null && p == null) {
        const font: (typeof fonts)[number] = { font: e.slice(1).replace(/^-/, '') };
        for (const x of f.slice(1)) {
          if (x[0] === 'M') font.size = (parseInt(x.slice(1), 10) || 200) / 20;
          if (x[0] === 'S') { if (x.includes('B')) font.bold = true; if (x.includes('I')) font.italic = true; }
        }
        fonts.push(font);
      }
      continue;
    }
    if (rt !== 'C' && rt !== 'F') continue;
    let value: Cell['v'] | undefined, expr: string | undefined, fmt: string | undefined;
    let style: CellStyle | null = null;
    let hasX = false, hasY = false, colRange: [number, number] | null = null, rowRange: number | null = null;
    for (const fld of f.slice(1)) {
      const t = fld[0], rest = fld.slice(1);
      switch (t) {
        case 'Y': row = (parseInt(rest, 10) || 1) - 1; hasY = true; break;
        case 'X': col = (parseInt(rest, 10) || 1) - 1; hasX = true; break;
        case 'K': value = sylkValue(sylkUnescape(rest)); break;
        case 'E': expr = sylkUnescape(rest); break;
        case 'P': if (rt === 'F') fmt = formats[parseInt(rest, 10)]; break;
        case 'F': if (rt === 'F') { const ff = sylkFieldFormat(rest); if (ff.fmt && !fmt) fmt = ff.fmt; if (ff.align) { style = style || {}; style.hAlign = ff.align; } } break;
        case 'S': if (rt === 'F') {
          style = style || {};
          const fm = /M(\d+)/.exec(rest);
          const font = fm ? fonts[+fm[1] - 1] : undefined;
          if (font) { if (font.font) style.font = font.font; if (font.size && font.size !== 10) style.size = font.size; if (font.bold) style.bold = true; if (font.italic) style.italic = true; }
          const letters = rest.replace(/M\d+/, '');
          if (letters.includes('D')) style.bold = true;
          if (letters.includes('I')) style.italic = true;
          const b: Border = { style: 'thin' };
          if (letters.includes('T')) style.bt = b; if (letters.includes('B')) style.bb = b; if (letters.includes('L')) style.bl = b; if (letters.includes('R')) style.br = b;
          if (letters.includes('S')) style.pattern = { type: 'gray125' };
        } break;
        case 'W': if (rt === 'F') {
          const [a, b, w] = rest.trim().split(/\s+/).map(Number);
          if (a && b && w) for (let c = a - 1; c <= b - 1 && c < MAX_COLS; c++) sheet.colW.set(c, Math.round(w * 7 + 5));
        } break;
        case 'C': if (rt === 'F') { const n = parseInt(rest, 10); if (n) colRange = [n - 1, n - 1]; } break;
        case 'R': if (rt === 'F') { const n = parseInt(rest, 10); if (n) rowRange = n - 1; } break;
        case 'M': if (rt === 'F' && !hasX && !hasY) { const h = parseInt(rest, 10); if (h && rowRange != null) sheet.rowH.set(rowRange, Math.round((h / 20) * 4 / 3)); } break;
      }
    }
    if (rt === 'F' && !hasX && !hasY) {
      const st: CellStyle = { ...(style || {}) };
      if (fmt && fmt !== 'General') st.numFmt = fmt;
      if (Object.keys(st).length) { if (colRange) colFmt.set(colRange[0], st); else if (rowRange != null) rowFmt.set(rowRange, st); }
      continue;
    }
    if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) continue;
    const k = cellKey(row, col);
    if (rt === 'C' && (value !== undefined || expr)) {
      const cell: Cell = { ...(sheet.cells.get(k) || {}) };
      if (value !== undefined) cell.v = value;
      if (expr) cell.f = r1c1ToA1(expr.replace(/^=/, ''), row, col);
      sheet.cells.set(k, cell);
    } else if (rt === 'F' && (fmt || style)) {
      const st: CellStyle = { ...(style || {}) };
      if (fmt && fmt !== 'General') st.numFmt = fmt;
      const sid = book.styleId(st);
      if (sid) sheet.cells.set(k, { ...(sheet.cells.get(k) || {}), s: sid });
    }
  }
  for (const [c, st] of colFmt) { const id = book.styleId(st); if (id) sheet.colStyle.set(c, id); }
  for (const [r, st] of rowFmt) { const id = book.styleId(st); if (id) sheet.rowStyle.set(r, id); }
  book.clearHistory();
  return book;
}

// --------------------------------------------------------------------------------------------
// DIF

/**
 * DIF (Data Interchange Format, VisiCalc 1979; also written by Lotus, Excel, LibreOffice):
 * header items (TABLE, VECTORS, TUPLES, LABEL, ...), then DATA with BOT-separated tuples of
 * type,number / string pairs. Handles the common variants: numbers with V, TRUE/FALSE, NA and
 * ERROR indicators, values written as text (dates), strings spanning lines and the
 * column-major orientation some writers use (one vector per row).
 */
export function importDif(text: string): Workbook {
  const book = new Workbook(false);
  const lines = text.split(/\r\n|\r|\n/);
  let i = 0, title = '';
  let vectors = 0, tuples = 0;
  // Header: topic line, then "vector,value", then a string.
  while (i + 2 < lines.length) {
    const topic = lines[i].trim().toUpperCase();
    const [, nv] = lines[i + 1].split(',');
    const str = lines[i + 2].trim().replace(/^"|"$/g, '');
    if (topic === 'DATA') { i += 3; break; }
    if (topic === 'TABLE') title = str;
    if (topic === 'VECTORS') vectors = +nv || 0;
    if (topic === 'TUPLES') tuples = +nv || 0;
    i += 3;
  }
  const sheet = new Sheet(book, (title && title !== 'EXCEL' ? title : 'Sheet1').replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet1');
  book.sheets.push(sheet);
  let row = -1, col = 0;
  const readString = (): string => {
    let t = lines[i] ?? '';
    i++;
    if (t.trimStart().startsWith('"')) {
      const closed = (x: string) => (x.match(/"/g) || []).length % 2 === 0 && x.trim().endsWith('"');
      while (!closed(t) && i < lines.length) { t += '\n' + lines[i]; i++; }
      const s = t.trim();
      return s.slice(1, s.endsWith('"') && s.length > 1 ? -1 : undefined).replace(/""/g, '"');
    }
    return t.trim();
  };
  while (i < lines.length) {
    const head = lines[i].trim();
    i++;
    if (!head) continue;
    const comma = head.indexOf(',');
    const type = parseInt(head.slice(0, comma < 0 ? undefined : comma), 10);
    const num = comma < 0 ? '' : head.slice(comma + 1).trim();
    if (type === -1) {
      const kw = readString().toUpperCase();
      if (kw === 'BOT') { row++; col = 0; continue; }
      if (kw === 'EOD') break;
      continue;
    }
    if (row < 0) row = 0;
    if (type === 0) {
      const ind = readString().toUpperCase();
      let cell: Cell | null = null;
      if (ind === 'V') {
        const n = Number(num);
        if (num !== '' && Number.isFinite(n)) cell = { v: n };
        else if (num) {
          const d = parseDateTime(num.replace(/^"|"$/g, ''));
          cell = d ? { v: d.serial, s: book.styleId({ numFmt: d.format }) } : { v: num };
        }
      } else if (ind === 'TRUE' || ind === 'FALSE') cell = { v: ind === 'TRUE' };
      else if (ind === 'NA') cell = { v: new CellError('#N/A') };
      else if (ind === 'ERROR') cell = { v: new CellError('#VALUE!') };
      if (cell && row < MAX_ROWS && col < MAX_COLS) sheet.cells.set(cellKey(row, col), cell);
      col++;
    } else if (type === 1) {
      const s = readString();
      if (s !== '' && row < MAX_ROWS && col < MAX_COLS) sheet.cells.set(cellKey(row, col), { v: s, ...(s.includes('\n') ? { s: book.styleId({ wrap: true }) } : {}) });
      col++;
    } else i++;
  }
  // Some writers put each column in a tuple (VECTORS = rows): transpose when the shape says so.
  const used = sheet.usedRange();
  if (used && vectors && tuples && used.r2 + 1 === vectors && used.c2 + 1 === tuples && vectors !== tuples) {
    const cells = [...sheet.cells];
    sheet.cells.clear();
    for (const [k, cell] of cells) sheet.cells.set(cellKey(k % MAX_COLS, Math.floor(k / MAX_COLS)), cell);
  }
  book.clearHistory();
  return book;
}

// --------------------------------------------------------------------------------------------
// XML Spreadsheet 2003

const NAMED_FORMATS: Record<string, string> = {
  'General Number': 'General', Fixed: '0.00', Standard: '#,##0.00', Percent: '0.00%', Scientific: '0.00E+00', Currency: '$#,##0.00_);[Red]($#,##0.00)',
  'Euro Currency': '[$€-x-euro2] #,##0.00', 'General Date': 'm/d/yyyy h:mm', 'Long Date': 'dddd, mmmm d, yyyy', 'Medium Date': 'd-mmm-yy', 'Short Date': 'm/d/yyyy',
  'Long Time': 'h:mm:ss AM/PM', 'Medium Time': 'h:mm AM/PM', 'Short Time': 'h:mm', 'Yes/No': '"Yes";"Yes";"No"', 'True/False': '"True";"True";"False"', 'On/Off': '"On";"On";"Off"',
  '@': '@',
};

function xmlBorder(b: Element): Border | undefined {
  const ls = attr(b, 'LineStyle');
  if (!ls || ls === 'None') return undefined;
  const w = +(attr(b, 'Weight') || 1);
  const style: Border['style'] = ls === 'Double' ? 'double' : ls === 'Dash' ? (w >= 2 ? 'mediumDashed' : 'dashed') : ls === 'Dot' ? 'dotted' : /DashDot/.test(ls) ? 'dashDot' : w >= 3 ? 'thick' : w >= 2 ? 'medium' : w === 0 ? 'hair' : 'thin';
  const color = normColor(attr(b, 'Color')) || undefined;
  return { style, ...(color ? { color } : {}) };
}

export function importXml2003(text: string, warnings: string[] = []): Workbook {
  const doc = parseXml(text, 'XML Spreadsheet');
  const root = doc.documentElement;
  if (root.localName !== 'Workbook') throw new Error('This XML file is not an Excel XML Spreadsheet');
  const book = new Workbook(false);
  // Styles
  const raw = new Map<string, { el: Element; parent: string | null }>();
  for (const st of kids(kid(root, 'Styles'), 'Style')) raw.set(attr(st, 'ID') || '', { el: st, parent: attr(st, 'Parent') });
  const resolved = new Map<string, CellStyle>();
  const resolve = (id: string, depth = 0): CellStyle => {
    if (resolved.has(id)) return resolved.get(id)!;
    const r = raw.get(id);
    if (!r || depth > 8) return {};
    const st: CellStyle = r.parent ? { ...resolve(r.parent, depth + 1) } : {};
    const font = kid(r.el, 'Font');
    if (font) {
      if (attr(font, 'Bold') === '1') st.bold = true;
      if (attr(font, 'Italic') === '1') st.italic = true;
      if (attr(font, 'StrikeThrough') === '1') st.strike = true;
      const u = attr(font, 'Underline'); if (u && u !== 'None') st.underline = u === 'Double' ? 'double' : 'single';
      const va = attr(font, 'VerticalAlign'); if (va === 'Superscript') st.vertAlign = 'superscript'; else if (va === 'Subscript') st.vertAlign = 'subscript';
      const c = normColor(attr(font, 'Color')); if (c && c !== '#000000') st.color = c;
      const sz = +(attr(font, 'Size') || 0); if (sz && sz !== 11 && sz !== 10) st.size = sz;
      const fn = attr(font, 'FontName'); if (fn && fn !== 'Calibri' && fn !== 'Arial') st.font = fn;
    }
    const interior = kid(r.el, 'Interior');
    if (interior) {
      const c = normColor(attr(interior, 'Color'));
      const pat = attr(interior, 'Pattern');
      if (c && pat !== 'None') st.fill = c;
    }
    const al = kid(r.el, 'Alignment');
    if (al) {
      const h = attr(al, 'Horizontal');
      const hm: Record<string, CellStyle['hAlign']> = { Left: 'left', Center: 'center', Right: 'right', Justify: 'justify', Fill: 'fill', CenterAcrossSelection: 'centerContinuous', Distributed: 'distributed' };
      if (h && hm[h]) st.hAlign = hm[h];
      const v = attr(al, 'Vertical');
      const vm: Record<string, CellStyle['vAlign']> = { Top: 'top', Center: 'middle', Bottom: 'bottom', Justify: 'justify', Distributed: 'distributed' };
      if (v && vm[v]) st.vAlign = vm[v];
      if (attr(al, 'WrapText') === '1') st.wrap = true;
      if (attr(al, 'ShrinkToFit') === '1') st.shrink = true;
      const ind = +(attr(al, 'Indent') || 0); if (ind) st.indent = ind;
      const rot = +(attr(al, 'Rotate') || 0); if (rot) st.rotate = rot;
    }
    for (const b of kids(kid(r.el, 'Borders'), 'Border')) {
      const bd = xmlBorder(b);
      if (!bd) continue;
      switch (attr(b, 'Position')) {
        case 'Top': st.bt = bd; break; case 'Bottom': st.bb = bd; break; case 'Left': st.bl = bd; break; case 'Right': st.br = bd; break;
        case 'DiagonalLeft': st.diagDown = bd; break; case 'DiagonalRight': st.diagUp = bd; break;
      }
    }
    const nf = kid(r.el, 'NumberFormat');
    const fmt = attr(nf, 'Format');
    if (fmt) { const code = NAMED_FORMATS[fmt] ?? fmt; if (code !== 'General') st.numFmt = code; else delete st.numFmt; }
    const pr = kid(r.el, 'Protection');
    if (pr && attr(pr, 'Protected') === '0') st.locked = false;
    if (pr && attr(pr, 'HideFormula') === '1') st.hidden = true;
    resolved.set(id, st);
    return st;
  };
  const sidCache = new Map<string, number>();
  const sid = (id: string | null) => {
    if (!id || id === 'Default') return 0;
    let s = sidCache.get(id);
    if (s == null) { s = book.styleId(resolve(id)); sidCache.set(id, s); }
    return s;
  };

  for (const ws of kids(root, 'Worksheet')) {
    let name = attr(ws, 'Name') || `Sheet${book.sheets.length + 1}`;
    while (book.sheetByName(name)) name += '_';
    const sheet = new Sheet(book, name);
    book.sheets.push(sheet);
    const table = kid(ws, 'Table');
    if (table) {
      const dcw = +(attr(table, 'DefaultColumnWidth') || 0); if (dcw) sheet.defaultColW = Math.round((dcw * 4) / 3);
      const drh = +(attr(table, 'DefaultRowHeight') || 0); if (drh) sheet.defaultRowH = Math.round((drh * 4) / 3);
      let c = 0;
      for (const col of kids(table, 'Column')) {
        const idx = attr(col, 'Index'); if (idx) c = +idx - 1;
        const span = +(attr(col, 'Span') || 0);
        const w = +(attr(col, 'Width') || 0);
        const hidden = attr(col, 'Hidden') === '1';
        const s = sid(attr(col, 'StyleID'));
        for (let k = 0; k <= span && c < MAX_COLS; k++, c++) {
          if (w) sheet.colW.set(c, Math.round((w * 4) / 3));
          if (hidden) sheet.hiddenCols.add(c);
          if (s) sheet.colStyle.set(c, s);
        }
      }
      let r = 0;
      for (const rowEl of kids(table, 'Row')) {
        const idx = attr(rowEl, 'Index'); if (idx) r = +idx - 1;
        const span = +(attr(rowEl, 'Span') || 0);
        const h = +(attr(rowEl, 'Height') || 0);
        if (h && attr(rowEl, 'AutoFitHeight') !== '1') sheet.rowH.set(r, Math.round((h * 4) / 3));
        else if (h && Math.abs(Math.round((h * 4) / 3) - sheet.defaultRowH) > 2) sheet.rowH.set(r, Math.round((h * 4) / 3));
        if (attr(rowEl, 'Hidden') === '1') sheet.hiddenRows.add(r);
        const rs = sid(attr(rowEl, 'StyleID')); if (rs) sheet.rowStyle.set(r, rs);
        let c2 = 0;
        for (const cellEl of kids(rowEl, 'Cell')) {
          const ci = attr(cellEl, 'Index'); if (ci) c2 = +ci - 1;
          if (c2 >= MAX_COLS) break;
          const cell: Cell = {};
          const s = sid(attr(cellEl, 'StyleID')); if (s) cell.s = s;
          const data = kid(cellEl, 'Data') || kid(cellEl, 'SmartTags') && desc(cellEl, 'Data')[0];
          if (data) {
            const t = attr(data, 'Type');
            const txt = data.textContent || '';
            if (t === 'Number') { const n = Number(txt); cell.v = Number.isFinite(n) ? n : txt; }
            else if (t === 'Boolean') cell.v = txt.trim() === '1' || txt.trim().toUpperCase() === 'TRUE';
            else if (t === 'Error') cell.v = errorFromCode(txt.trim()) || new CellError('#VALUE!');
            else if (t === 'DateTime') {
              const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(txt);
              cell.v = m ? partsToSerial(+m[1], +m[2], +m[3]) + ((+(m[4] || 0)) * 3600 + (+(m[5] || 0)) * 60 + (+(m[6] || 0))) / 86400 : txt;
              if (m && +m[1] <= 1899) cell.v = ((+(m[4] || 0)) * 3600 + (+(m[5] || 0)) * 60 + (+(m[6] || 0))) / 86400;
            } else cell.v = txt;
          }
          const formula = attr(cellEl, 'Formula');
          if (formula) cell.f = /^of:/i.test(formula) ? odfToFormula(formula) : r1c1ToA1(formula.replace(/^=/, ''), r, c2);
          const arr = attr(cellEl, 'ArrayRange');
          if (arr && formula) {
            const a = r1c1ToA1(arr, r, c2).replace(/\$/g, '');
            const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a);
            if (m) {
              const toC = (s2: string) => [...s2].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
              cell.arr = { r1: +m[2] - 1, c1: toC(m[1]), r2: +(m[4] || m[2]) - 1, c2: toC(m[3] || m[1]) };
            }
          }
          const href = attr(cellEl, 'HRef'); if (href) cell.link = href;
          const tip = attr(cellEl, 'HRefScreenTip'); if (href && tip) cell.linkTip = tip;
          const ma = +(attr(cellEl, 'MergeAcross') || 0), md = +(attr(cellEl, 'MergeDown') || 0);
          if (ma || md) sheet.merges.push({ r1: r, c1: c2, r2: r + md, c2: c2 + ma });
          const cm = kid(cellEl, 'Comment');
          if (cm) sheet.comments.set(cellKey(r, c2), { author: attr(cm, 'Author') || '', text: (kid(cm, 'Data')?.textContent || '').trim(), time: 0, note: true });
          if (cell.v != null || cell.f != null || cell.s || cell.link) sheet.cells.set(cellKey(r, c2), cell);
          c2 += 1 + ma;
        }
        r += 1 + span;
      }
    }
    // Worksheet options (x: namespace)
    const opts = kid(ws, 'WorksheetOptions');
    if (opts) {
      if (kid(opts, 'FreezePanes')) {
        const sh = +(kid(opts, 'SplitHorizontal')?.textContent || 0), sv = +(kid(opts, 'SplitVertical')?.textContent || 0);
        sheet.freeze = { rows: sh, cols: sv };
      }
      if (kid(opts, 'Selected')) book.activeSheet = book.sheets.length - 1;
      const vis = kid(opts, 'Visible')?.textContent;
      if (vis === 'SheetHidden') sheet.visibility = 'hidden'; else if (vis === 'SheetVeryHidden') sheet.visibility = 'veryHidden';
      if (kid(opts, 'DoNotDisplayGridlines')) sheet.showGrid = false;
      const zoom = +(kid(opts, 'Zoom')?.textContent || 0); if (zoom) sheet.zoom = zoom / 100;
      if (kid(opts, 'DisplayRightToLeft')) sheet.rtl = true;
      if (kid(opts, 'ProtectContents')?.textContent === 'True') sheet.protection = { on: true, allow: ['selectLocked', 'selectUnlocked'] };
    }
    readXmlRules(ws, sheet);
    for (const nr of kids(kid(ws, 'Names'), 'NamedRange')) addName(book, nr, sheet.name);
    if (kid(ws, 'AutoFilter')) {
      const rng = attr(kid(ws, 'AutoFilter'), 'Range');
      if (rng) {
        const a = r1c1ToA1(rng, 0, 0).replace(/\$/g, '');
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(a);
        const toC = (s2: string) => [...s2].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
        if (m) sheet.filter = { range: { r1: +m[2] - 1, c1: toC(m[1]), r2: +m[4] - 1, c2: toC(m[3]) }, columns: {} };
      }
    }
  }
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
  for (const nr of kids(kid(root, 'Names'), 'NamedRange')) addName(book, nr, null);
  // Document properties
  const dp = kid(root, 'DocumentProperties');
  if (dp) {
    const t = (n: string) => kid(dp, n)?.textContent || '';
    book.props.title = t('Title'); book.props.subject = t('Subject'); book.props.author = t('Author'); book.props.lastModifiedBy = t('LastAuthor');
    book.props.company = t('Company'); book.props.manager = t('Manager'); book.props.keywords = t('Keywords'); book.props.comments = t('Description');
    const c = Date.parse(t('Created')); if (Number.isFinite(c)) book.props.created = c;
  }
  book.clearHistory();
  return book;
}

const XML_OPS: Record<string, DataValidation['operator']> = {
  Between: 'between', NotBetween: 'notBetween', Equal: 'equal', NotEqual: 'notEqual', Greater: 'greaterThan', Less: 'lessThan', GreaterOrEqual: 'greaterThanOrEqual', LessOrEqual: 'lessThanOrEqual',
};

/** "R2C1:R4C1,R6C1" (absolute R1C1) -> ranges. */
function xmlRanges(s: string): Range[] {
  return s.split(',').map((p) => parseRange(r1c1ToA1(p.trim(), 0, 0).replace(/\$/g, ''))).filter((x): x is Range => !!x);
}

/** Excel XML conditional formats and data validation (urn:schemas-microsoft-com:office:excel). */
function readXmlRules(ws: Element, sheet: Sheet): void {
  const t = (el: Element | null, n: string) => kid(el, n)?.textContent?.trim() ?? '';
  for (const cf of kids(ws, 'ConditionalFormatting')) {
    const ranges = xmlRanges(t(cf, 'Range'));
    if (!ranges.length) continue;
    const base = ranges[0];
    for (const cond of kids(cf, 'Condition')) {
      const q = t(cond, 'Qualifier');
      const f = (n: string) => { const v = t(cond, n); return v ? r1c1ToA1(v.replace(/^=/, ''), base.r1, base.c1) : ''; };
      const rule: CFRule = { id: newId('cf'), ranges, type: q ? 'cellIs' : 'expression', priority: sheet.cf.length + 1 };
      if (q) { rule.operator = XML_OPS[q] || 'equal'; rule.formulas = [f('Value1'), ...(t(cond, 'Value2') ? [f('Value2')] : [])]; }
      else rule.formulas = [f('Value1')];
      const css = attr(kid(cond, 'Format'), 'Style') || '';
      const st: Partial<CellStyle> = {};
      const prop = (n: string) => new RegExp(`(?:^|;)\\s*${n}\\s*:\\s*([^;]+)`, 'i').exec(css)?.[1]?.trim();
      const c = normColor(prop('color') || ''); if (c) st.color = c;
      const bg = normColor(prop('background') || prop('background-color') || ''); if (bg) st.fill = bg;
      if (/bold|[6-9]00/.test(prop('font-weight') || '')) st.bold = true;
      if (/italic/.test(prop('font-style') || '')) st.italic = true;
      if (/underline/.test(prop('text-underline-style') || prop('text-decoration') || '')) st.underline = 'single';
      if (/line-through/.test(prop('text-line-through') || prop('text-decoration') || '')) st.strike = true;
      rule.style = st;
      sheet.cf.push(rule);
    }
  }
  for (const dv of kids(ws, 'DataValidation')) {
    const ranges = xmlRanges(t(dv, 'Range'));
    if (!ranges.length) continue;
    const base = ranges[0];
    const conv = (v: string) => (v ? r1c1ToA1(v.replace(/^=/, ''), base.r1, base.c1) : '');
    const typeMap: Record<string, DataValidation['type']> = { Whole: 'whole', Decimal: 'decimal', List: 'list', Date: 'date', Time: 'time', TextLength: 'textLength', Custom: 'custom' };
    const v: DataValidation = { id: newId('dv'), ranges, type: typeMap[t(dv, 'Type')] || 'any' };
    const q = t(dv, 'Qualifier');
    const min = t(dv, 'Min'), max = t(dv, 'Max'), value = t(dv, 'Value') || t(dv, 'CellRangeList');
    if (v.type === 'list' || v.type === 'custom') v.formula1 = conv(value);
    else if (v.type !== 'any') {
      v.operator = XML_OPS[q] || (min || max ? 'between' : 'equal');
      if (v.operator === 'between' || v.operator === 'notBetween') { v.formula1 = conv(min || value); v.formula2 = conv(max); }
      else v.formula1 = conv(value || min || max);
    }
    v.allowBlank = !kid(dv, 'IgnoreBlanks') || t(dv, 'IgnoreBlanks') !== 'False';
    if (v.type === 'list') v.showDropdown = !kid(dv, 'ComboHide');
    if (t(dv, 'InputTitle')) v.inputTitle = t(dv, 'InputTitle');
    if (t(dv, 'InputMessage')) v.inputMessage = t(dv, 'InputMessage');
    if (t(dv, 'ErrorTitle')) v.errorTitle = t(dv, 'ErrorTitle');
    if (t(dv, 'ErrorMessage')) v.errorMessage = t(dv, 'ErrorMessage');
    const es = t(dv, 'ErrorStyle');
    v.errorStyle = es === 'Warn' ? 'warning' : es === 'Info' ? 'information' : 'stop';
    sheet.validations.push(v);
  }
}

function addName(book: Workbook, nr: Element, sheet: string | null): void {
  const name = attr(nr, 'Name') || '';
  const ref = r1c1ToA1((attr(nr, 'RefersTo') || '').replace(/^=/, ''), 0, 0);
  if (!name || !ref) return;
  if (/^_FilterDatabase$/i.test(name)) return;
  if (/^Print_Area$/i.test(name) && sheet) { const sh = book.sheetByName(sheet); if (sh) sh.print.area = ref.replace(/^.*!/, ''); return; }
  if (/^Print_Titles$/i.test(name) && sheet) {
    const sh = book.sheetByName(sheet);
    if (sh) for (const part of ref.split(',')) {
      const r = part.replace(/^.*!/, '').replace(/\$/g, '');
      if (/^\d+:\d+$/.test(r)) sh.print.titleRows = r; else if (/^[A-Z]+:[A-Z]+$/i.test(r)) sh.print.titleCols = r;
    }
    return;
  }
  book.names.push({ name, ref, sheet });
}
