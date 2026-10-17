/**
 * Uniform Office Format spreadsheets (UOF 1.0, Chinese national standard GB/T 20916; .uos),
 * as written by RedOffice, Yozo/EIOffice and LibreOffice's UOF filter.
 *
 * A single XML document with Chinese element names: uof:UOF > uof:式样集 (styles: 单元格式样 with
 * 字体格式, 对齐格式, 边框, 填充, 数字格式) and uof:电子表格 > 表:主体 > 表:工作表 (sheets) with
 * 表:工作表内容 > 表:列 / 表:行 > 表:单元格 > 表:数据 (数据类型, 数据数值, text in 字:句/字:文本串).
 * Read: values, formulas (表:公式), styles, column widths, row heights, hidden rows and columns,
 * merges, hyperlinks, comments and hidden sheets.
 */
import { Workbook, Sheet } from '../model/workbook';
import { CellError, errorFromCode, type Cell, type CellStyle, type Border } from '../model/types';
import { cellKey, MAX_COLS, MAX_ROWS } from '../model/address';
import { partsToSerial } from '../model/dates';
import { parseXml, kids, kid, desc, normColor } from './index';

/** Attribute by local name, whatever the prefix (UOF uses several namespaces for the same names). */
function at(el: Element | null | undefined, local: string): string | null {
  if (!el) return null;
  for (const a of Array.from(el.attributes)) if (a.localName === local) return a.value;
  return null;
}

const H: Record<string, CellStyle['hAlign']> = { left: 'left', center: 'center', right: 'right', fill: 'fill', justify: 'justify', distributed: 'distributed', 'center-across-selection': 'centerContinuous' };
const V: Record<string, CellStyle['vAlign']> = { top: 'top', center: 'middle', bottom: 'bottom', justify: 'justify', distributed: 'distributed' };

function border(el: Element | null): Border | undefined {
  const t = at(el, '类型');
  if (!el || !t || t === 'none') return undefined;
  const w = parseFloat(at(el, '宽度') || '') || 0;
  const style: Border['style'] = t === 'double' ? 'double' : /dash/.test(t) ? 'dashed' : /dot/.test(t) ? 'dotted' : w >= 2.5 ? 'thick' : w >= 1.5 ? 'medium' : 'thin';
  const c = normColor(at(el, '颜色'));
  return { style, ...(c && c !== '#000000' ? { color: c } : {}) };
}

/** UOF format codes write date and time letters in upper case ("YYYY-MM-DD"); the code reader expects lower case. */
function formatCode(code: string): string {
  let out = '', q = false;
  for (const ch of code) {
    if (ch === '"') q = !q;
    out += !q && /[YMDHS]/.test(ch) ? ch.toLowerCase() : ch;
  }
  return out.replace(/am\/pm/gi, 'AM/PM');
}

function readStyle(el: Element): CellStyle {
  const s: CellStyle = {};
  const font = kid(el, '字体格式');
  if (font) {
    if (at(kid(font, '粗体'), '值') === 'true') s.bold = true;
    if (at(kid(font, '斜体'), '值') === 'true') s.italic = true;
    const u = kid(font, '下划线'); if (u && at(u, '类型') && at(u, '类型') !== 'none') s.underline = at(u, '类型') === 'double' ? 'double' : 'single';
    const del = kid(font, '删除线'); if (del && at(del, '类型') && at(del, '类型') !== 'none') s.strike = true;
    const f = kid(font, '字体');
    if (f) {
      const name = at(f, '西文字体引用') || at(f, '中文字体引用');
      if (name) s.font = name.replace(/_/g, ' ');
      const sz = parseFloat(at(f, '字号') || '');
      if (sz && sz !== 10) s.size = sz;
      const c = normColor(at(f, '颜色'));
      if (c && c !== '#000000') s.color = c;
    }
    const pos = kid(font, '上下标'); if (pos) { const t = pos.textContent || ''; if (/sup/.test(t)) s.vertAlign = 'superscript'; else if (/sub/.test(t)) s.vertAlign = 'subscript'; }
  }
  const al = kid(el, '对齐格式');
  if (al) {
    const h = H[(kid(al, '水平对齐方式')?.textContent || '').trim()]; if (h) s.hAlign = h;
    const v = V[(kid(al, '垂直对齐方式')?.textContent || '').trim()]; if (v && v !== 'bottom') s.vAlign = v;
    if (at(kid(al, '自动换行'), '值') === 'true') s.wrap = true;
    if (at(kid(al, '缩小字体填充'), '值') === 'true') s.shrink = true;
    const ind = parseFloat(kid(al, '缩进')?.textContent || ''); if (ind > 0) s.indent = Math.round(ind);
    const rot = parseFloat(kid(al, '文字旋转角度')?.textContent || ''); if (rot) s.rotate = Math.max(-90, Math.min(90, rot));
  }
  for (const b of kids(el, '边框')) {
    const l = border(kid(b, '左')), t = border(kid(b, '上')), r = border(kid(b, '右')), d = border(kid(b, '下'));
    if (l) s.bl = l; if (t) s.bt = t; if (r) s.br = r; if (d) s.bb = d;
    const d1 = border(kid(b, '对角线1')), d2 = border(kid(b, '对角线2'));
    if (d1) s.diagDown = d1; if (d2) s.diagUp = d2;
  }
  const fill = kid(el, '填充');
  if (fill) { const c = normColor(desc(fill, '颜色')[0]?.textContent || ''); if (c) s.fill = c; }
  const nf = kid(el, '数字格式');
  if (nf) {
    const code = at(nf, '格式码'), cat = at(nf, '分类名称');
    if (code && code !== 'General' && code !== 'boolean' && cat !== 'general') s.numFmt = formatCode(code);
  }
  return s;
}

/** Length in the document unit -> px. */
function len(v: string | null, unit: string): number | null {
  const n = parseFloat(v || '');
  if (!(n > 0)) return null;
  switch (unit) {
    case 'inch': return n * 96;
    case 'cm': return (n * 96) / 2.54;
    case 'mm': return (n * 96) / 25.4;
    default: return (n * 4) / 3; // points
  }
}

export function importUof(text: string, warnings: string[] = []): Workbook {
  const doc = parseXml(text, 'UOF document');
  const root = doc.documentElement;
  if (root.localName !== 'UOF') throw new Error('This is not a Uniform Office Format document.');
  const book = new Workbook(false);
  const styles = new Map<string, number>();
  for (const st of desc(root, '单元格式样')) {
    const id = at(st, '标识符');
    if (id && !styles.has(id)) styles.set(id, id === 'Default' ? 0 : book.styleId(readStyle(st)));
  }
  const sheetEl = desc(root, '电子表格')[0];
  if (!sheetEl) throw new Error('This UOF document has no spreadsheet body.');
  const unit = (desc(sheetEl, '度量单位')[0]?.textContent || 'pt').trim();
  const rc = at(desc(sheetEl, '是否RC引用')[0], '值') === 'true';
  for (const ws of desc(sheetEl, '工作表')) {
    let name = at(ws, '名称') || at(ws, '标识符') || `Sheet${book.sheets.length + 1}`;
    while (book.sheetByName(name)) name += '_';
    const sheet = new Sheet(book, name);
    book.sheets.push(sheet);
    if (at(ws, '隐藏') === 'true') sheet.visibility = 'hidden';
    const view = desc(ws, '视图')[0];
    if (view && at(kid(view, '网格'), '值') === 'false') { /* LibreOffice writes false for the default grid; keep gridlines */ }
    const tab = normColor(desc(ws, '标签背景色')[0]?.textContent || '');
    if (tab && tab !== '#ffffff') sheet.tabColor = tab;
    const content = desc(ws, '工作表内容')[0];
    if (!content) continue;
    let c0 = 0;
    for (const col of kids(content, '列')) {
      const n = parseInt(at(col, '列号') || '', 10);
      if (n > 0) c0 = n - 1;
      const span = Math.max(1, parseInt(at(col, '跨度') || '1', 10) || 1);
      const w = len(at(col, '列宽'), unit);
      for (let c = c0; c < c0 + span && c < MAX_COLS; c++) {
        if (w) sheet.colW.set(c, Math.round(w));
        if (at(col, '隐藏') === 'true') sheet.hiddenCols.add(c);
      }
      c0 += span;
    }
    let r = 0;
    for (const row of kids(content, '行')) {
      const n = parseInt(at(row, '行号') || '', 10);
      if (n > 0) r = n - 1;
      if (r >= MAX_ROWS) break;
      const h = len(at(row, '行高'), 'pt');
      if (h && Math.abs(Math.round(h) - sheet.defaultRowH) > 1) sheet.rowH.set(r, Math.round(h));
      if (at(row, '隐藏') === 'true') sheet.hiddenRows.add(r);
      let c = 0;
      for (const ce of kids(row, '单元格')) {
        const cn = parseInt(at(ce, '列号') || '', 10);
        if (cn > 0) c = cn - 1;
        if (c >= MAX_COLS) break;
        const cell: Cell = {};
        const sid = styles.get(at(ce, '式样引用') || '') || 0;
        if (sid) cell.s = sid;
        const data = kid(ce, '数据');
        if (data) {
          const type = at(data, '数据类型') || 'text';
          const raw = at(data, '数据数值');
          const shown = desc(data, '文本串').map((t) => t.textContent || '').join('') || (data.textContent || '').trim();
          switch (type) {
            case 'number': case 'percentage': case 'currency': { const v = Number(raw ?? shown); cell.v = Number.isFinite(v) ? v : shown; break; }
            case 'date': {
              const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(raw || '');
              cell.v = m ? partsToSerial(+m[1], +m[2], +m[3]) + ((+(m[4] || 0)) * 3600 + (+(m[5] || 0)) * 60 + (+(m[6] || 0))) / 86400 : shown;
              break;
            }
            case 'time': { const m = /^-?P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/.exec(raw || ''); cell.v = m ? +(m[1] || 0) + ((+(m[2] || 0)) * 3600 + (+(m[3] || 0)) * 60 + (+(m[4] || 0))) / 86400 : shown; break; }
            case 'boolean': cell.v = /^(true|1)$/i.test(raw ?? shown); break;
            case 'error': cell.v = errorFromCode(shown) || new CellError('#VALUE!'); break;
            default: if (shown) cell.v = shown;
          }
        }
        const f = at(ce, '公式') || at(data, '公式') || kid(data, '公式')?.textContent;
        if (f && f.trim().length > 1) cell.f = rc ? f.trim().replace(/^=/, '') : f.trim().replace(/^=/, '');
        const link = at(ce, '超链接引用'); if (link) cell.link = link;
        const mc = parseInt(at(ce, '合并列数') || '0', 10), mr = parseInt(at(ce, '合并行数') || '0', 10);
        if (mc > 0 || mr > 0) sheet.merges.push({ r1: r, c1: c, r2: r + Math.max(0, mr), c2: c + Math.max(0, mc) });
        const note = kid(ce, '批注');
        if (note) sheet.comments.set(cellKey(r, c), { author: at(note, '作者') || '', text: desc(note, '文本串').map((t) => t.textContent || '').join('\n'), time: 0, note: true });
        if (cell.v != null || cell.f != null || cell.s || cell.link) sheet.cells.set(cellKey(r, c), cell);
        c += 1 + Math.max(0, mc); // covered cells of a merge are not written
      }
      r += 1;
    }
  }
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
  const meta = desc(root, '元数据')[0];
  if (meta) {
    const t = (n: string) => (desc(meta, n)[0]?.textContent || '').trim();
    book.props.title = t('标题'); book.props.author = t('作者') || t('创建者'); book.props.subject = t('主题');
  }
  void warnings;
  book.clearHistory();
  return book;
}
