/** Built-in workbook templates. All content is original. */
import { Workbook, type Sheet } from './model/workbook';
import type { CellStyle } from './model/types';
import { cellKey, parseCell, type Range } from './model/address';
import { DEFAULT_TABLE_OPTS } from './model/tableStyles';
import { jsDateToSerial } from './model/dates';

export interface TemplateDef { id: string; name: string; category: string; description: string; icon?: string; build: () => Workbook }

type CellSpec = string | number | boolean | null | { v?: string | number | boolean; f?: string; s?: CellStyle };

/** Writes rows of values starting at addr. Strings starting with '=' are formulas. */
function put(sheet: Sheet, addr: string, rows: CellSpec[][], style?: CellStyle): void {
  const a = parseCell(addr)!;
  rows.forEach((row, i) => row.forEach((spec, j) => {
    if (spec === null || spec === undefined) return;
    const r = a.r + i, c = a.c + j;
    let cell: any = {};
    let st = style;
    if (typeof spec === 'object') { if (spec.f) cell.f = spec.f; if (spec.v !== undefined) cell.v = spec.v; if (spec.s) st = { ...(st || {}), ...spec.s }; }
    else if (typeof spec === 'string' && spec.startsWith('=')) cell.f = spec.slice(1);
    else cell.v = spec;
    if (st) cell.s = sheet.book.styleId(st);
    sheet.cells.set(cellKey(r, c), cell);
  }));
}

function styleRange(sheet: Sheet, rg: string, st: CellStyle): void {
  const [p, q] = rg.split(':');
  const a = parseCell(p)!, b = parseCell(q || p)!;
  for (let r = a.r; r <= b.r; r++) for (let c = a.c; c <= b.c; c++) {
    const k = cellKey(r, c);
    const cur = sheet.cells.get(k) || {};
    sheet.cells.set(k, { ...cur, s: sheet.book.derivedStyle(cur.s ?? 0, st) });
  }
}

function widths(sheet: Sheet, ws: number[]): void { ws.forEach((w, i) => sheet.colW.set(i, w)); }

function table(sheet: Sheet, name: string, rg: Range, cols: string[], style = 'TableStyleMedium2', totals = false): void {
  sheet.tables.push({ id: 't' + name, name, range: rg, columns: cols.map((n) => ({ name: n })), style, opts: { ...DEFAULT_TABLE_OPTS, totals } });
}

const TITLE: CellStyle = { size: 18, bold: true, color: '#1f3b57' };
const HEAD: CellStyle = { bold: true, fill: '#1f3b57', color: '#ffffff' };
const MONEY = '$#,##0.00';
const today = () => Math.floor(jsDateToSerial(new Date()));

export const TEMPLATES: TemplateDef[] = [
  { id: 'blank', name: 'Blank workbook', category: 'General', description: 'An empty workbook', icon: 'grid_on', build: () => new Workbook() },
  {
    id: 'budget', name: 'Monthly budget', category: 'General', description: 'Income, expenses and what is left', icon: 'savings',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Budget';
      put(s, 'A1', [['Monthly Budget']], TITLE);
      put(s, 'A3', [['Income', 'Planned', 'Actual', 'Difference']], HEAD);
      put(s, 'A4', [['Salary', 4200, 4200, '=C4-B4'], ['Other income', 300, 250, '=C5-B5'], [{ v: 'Total income', s: { bold: true } }, '=SUM(B4:B5)', '=SUM(C4:C5)', '=C6-B6']]);
      put(s, 'A8', [['Expenses', 'Planned', 'Actual', 'Difference']], HEAD);
      const items = ['Housing', 'Utilities', 'Groceries', 'Transportation', 'Insurance', 'Health', 'Savings', 'Entertainment', 'Other'];
      const planned = [1400, 220, 600, 300, 180, 120, 500, 150, 100];
      items.forEach((it, i) => put(s, `A${9 + i}`, [[it, planned[i], null, `=B${9 + i}-C${9 + i}`]]));
      put(s, 'A18', [[{ v: 'Total expenses', s: { bold: true } }, '=SUM(B9:B17)', '=SUM(C9:C17)', '=B18-C18']]);
      put(s, 'A20', [[{ v: 'Balance', s: { bold: true, size: 13 } }, '=B6-B18', '=C6-C18']]);
      styleRange(s, 'B4:D20', { numFmt: MONEY });
      s.cf.push({ id: 'cf1', ranges: [{ r1: 3, c1: 3, r2: 19, c2: 3 }], type: 'cellIs', operator: 'lessThan', formulas: ['0'], style: { color: '#c00000' } });
      widths(s, [160, 100, 100, 100]);
      return b;
    },
  },
  {
    id: 'invoice', name: 'Invoice', category: 'Business', description: 'Itemized invoice with tax and total', icon: 'receipt_long',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Invoice';
      put(s, 'A1', [['INVOICE']], { ...TITLE, size: 24 });
      put(s, 'A3', [['Your Company'], ['Street address'], ['City, State ZIP']]);
      put(s, 'E3', [['Invoice #', 1001], ['Date', { v: today(), s: { numFmt: 'mmmm d, yyyy' } }], ['Due', { f: 'F4+30', s: { numFmt: 'mmmm d, yyyy' } }]]);
      put(s, 'A7', [['Bill to:'], ['Customer name'], ['Address']]);
      put(s, 'A11', [['Description', 'Qty', 'Unit price', 'Amount']], HEAD);
      for (let i = 0; i < 8; i++) put(s, `A${12 + i}`, [[null, null, null, `=IF(B${12 + i}="","",B${12 + i}*C${12 + i})`]]);
      put(s, 'A12', [['Consulting (hours)', 10, 95]]);
      put(s, 'C21', [['Subtotal', '=SUM(D12:D19)'], ['Tax rate', { v: 0.07, s: { numFmt: '0.0%' } }], ['Tax', '=D21*D22'], [{ v: 'Total', s: { bold: true } }, { f: 'D21+D23', s: { bold: true } }]]);
      styleRange(s, 'C12:D19', { numFmt: MONEY }); styleRange(s, 'D21', { numFmt: MONEY }); styleRange(s, 'D23:D24', { numFmt: MONEY });
      widths(s, [260, 60, 100, 110, 80, 120]);
      return b;
    },
  },
  {
    id: 'loan', name: 'Loan amortization', category: 'Business', description: 'Payment schedule with interest and principal', icon: 'account_balance',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Loan';
      put(s, 'A1', [['Loan Amortization Schedule']], TITLE);
      put(s, 'A3', [['Loan amount', { v: 250000, s: { numFmt: MONEY } }], ['Annual rate', { v: 0.065, s: { numFmt: '0.00%' } }], ['Years', 30], ['Payments per year', 12], ['Payment', { f: '-PMT(B4/B6,B5*B6,B3)', s: { numFmt: MONEY, bold: true } }]]);
      put(s, 'A9', [['No.', 'Payment', 'Interest', 'Principal', 'Balance']], HEAD);
      put(s, 'A10', [['=SEQUENCE(B5*B6)']]);
      put(s, 'B10', [['=IF(A10#="","",$B$7)']]);
      put(s, 'C10', [['=IPMT($B$4/$B$6,A10#,$B$5*$B$6,-$B$3)']]);
      put(s, 'D10', [['=PPMT($B$4/$B$6,A10#,$B$5*$B$6,-$B$3)']]);
      put(s, 'E10', [['=$B$3-SCAN(0,D10#,LAMBDA(a,x,a+x))']]);
      s.colStyle.set(1, b.styleId({ numFmt: MONEY })); s.colStyle.set(2, b.styleId({ numFmt: MONEY })); s.colStyle.set(3, b.styleId({ numFmt: MONEY })); s.colStyle.set(4, b.styleId({ numFmt: MONEY }));
      s.freeze = { rows: 9, cols: 0 };
      widths(s, [130, 110, 110, 110, 120]);
      return b;
    },
  },
  {
    id: 'expenses', name: 'Expense tracker', category: 'General', description: 'Log expenses in a table with category totals', icon: 'payments',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Expenses';
      put(s, 'A1', [['Date', 'Description', 'Category', 'Amount', 'Paid with']]);
      const t = today();
      put(s, 'A2', [[t - 9, 'Groceries', 'Food', 84.12, 'Card'], [t - 7, 'Bus pass', 'Transport', 45, 'Cash'], [t - 4, 'Electric bill', 'Utilities', 96.4, 'Bank'], [t - 2, 'Dinner out', 'Food', 52.3, 'Card']]);
      styleRange(s, 'A2:A5', { numFmt: 'm/d/yyyy' }); styleRange(s, 'D2:D5', { numFmt: MONEY });
      table(s, 'Expenses', { r1: 0, c1: 0, r2: 5, c2: 4 }, ['Date', 'Description', 'Category', 'Amount', 'Paid with'], 'TableStyleMedium2', true);
      s.tables[0].columns[3].totalFn = 'sum';
      put(s, 'A6', [['Total', null, null, '=SUBTOTAL(109,Expenses[Amount])']]);
      put(s, 'G1', [['Category', 'Total']], HEAD);
      put(s, 'G2', [['=SORT(UNIQUE(Expenses[Category]))', '=SUMIFS(Expenses[Amount],Expenses[Category],G2#)']]);
      styleRange(s, 'H2:H12', { numFmt: MONEY });
      s.validations.push({ id: 'dv1', ranges: [{ r1: 1, c1: 2, r2: 500, c2: 2 }], type: 'list', formula1: '"Food,Transport,Utilities,Housing,Health,Entertainment,Other"', showDropdown: true, allowBlank: true });
      widths(s, [90, 200, 110, 90, 90, 20, 110, 90]);
      return b;
    },
  },
  {
    id: 'gradebook', name: 'Gradebook', category: 'Education', description: 'Scores, averages and letter grades', icon: 'school',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Grades';
      put(s, 'A1', [['Student', 'Test 1', 'Test 2', 'Project', 'Final', 'Average', 'Grade']], HEAD);
      const names = ['Avery', 'Blake', 'Casey', 'Devon', 'Emerson', 'Finley'];
      names.forEach((n, i) => { const r = i + 2; put(s, `A${r}`, [[n, 70 + ((i * 7) % 25), 75 + ((i * 11) % 20), 80 + ((i * 5) % 18), 68 + ((i * 13) % 30), `=AVERAGE(B${r}:E${r})`, `=IFS(F${r}>=90,"A",F${r}>=80,"B",F${r}>=70,"C",F${r}>=60,"D",TRUE,"F")`]]); });
      styleRange(s, 'F2:F7', { numFmt: '0.0' });
      put(s, 'A9', [[{ v: 'Class average', s: { bold: true } }, '=AVERAGE(B2:B7)', '=AVERAGE(C2:C7)', '=AVERAGE(D2:D7)', '=AVERAGE(E2:E7)', '=AVERAGE(F2:F7)']]);
      styleRange(s, 'B9:F9', { numFmt: '0.0' });
      s.cf.push({ id: 'cf1', ranges: [{ r1: 1, c1: 5, r2: 6, c2: 5 }], type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] });
      widths(s, [120, 70, 70, 70, 70, 80, 60]);
      return b;
    },
  },
  {
    id: 'timesheet', name: 'Weekly timesheet', category: 'Business', description: 'Hours per day with overtime', icon: 'schedule',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Timesheet';
      put(s, 'A1', [['Weekly Timesheet']], TITLE);
      put(s, 'A3', [['Day', 'Time in', 'Lunch out', 'Lunch in', 'Time out', 'Hours']], HEAD);
      ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'].forEach((d, i) => {
        const r = 4 + i;
        put(s, `A${r}`, [[d, i < 5 ? { v: 9 / 24, s: { numFmt: 'h:mm AM/PM' } } : null, i < 5 ? { v: 12 / 24, s: { numFmt: 'h:mm AM/PM' } } : null, i < 5 ? { v: 12.5 / 24, s: { numFmt: 'h:mm AM/PM' } } : null, i < 5 ? { v: 17.5 / 24, s: { numFmt: 'h:mm AM/PM' } } : null, { f: `IF(B${r}="",0,((C${r}-B${r})+(E${r}-D${r}))*24)`, s: { numFmt: '0.00' } }]]);
      });
      put(s, 'E12', [['Total', { f: 'SUM(F4:F10)', s: { numFmt: '0.00', bold: true } }], ['Overtime', { f: 'MAX(0,F12-40)', s: { numFmt: '0.00' } }]]);
      widths(s, [100, 90, 90, 90, 90, 70]);
      return b;
    },
  },
  {
    id: 'project', name: 'Project tracker', category: 'Business', description: 'Tasks with owners, dates, status and a timeline', icon: 'view_timeline',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Project';
      put(s, 'A1', [['Task', 'Owner', 'Start', 'Due', 'Status', '% Done']]);
      const t = today();
      const rows = [['Plan', 'Sam', t, t + 5, 'Done', 1], ['Design', 'Lee', t + 3, t + 14, 'In progress', 0.6], ['Build', 'Kim', t + 10, t + 30, 'Not started', 0], ['Test', 'Ray', t + 25, t + 40, 'Not started', 0], ['Launch', 'Sam', t + 41, t + 42, 'Not started', 0]];
      put(s, 'A2', rows);
      styleRange(s, 'C2:D6', { numFmt: 'd-mmm' }); styleRange(s, 'F2:F6', { numFmt: '0%' });
      table(s, 'Tasks', { r1: 0, c1: 0, r2: 5, c2: 5 }, ['Task', 'Owner', 'Start', 'Due', 'Status', '% Done'], 'TableStyleMedium9');
      s.validations.push({ id: 'dv1', ranges: [{ r1: 1, c1: 4, r2: 200, c2: 4 }], type: 'list', formula1: '"Not started,In progress,Blocked,Done"', showDropdown: true });
      s.cf.push({ id: 'cf1', ranges: [{ r1: 1, c1: 5, r2: 200, c2: 5 }], type: 'dataBar', barColor: '#5aa33f' });
      s.cf.push({ id: 'cf2', ranges: [{ r1: 1, c1: 3, r2: 200, c2: 3 }], type: 'expression', formulas: ['AND($D2<TODAY(),$E2<>"Done")'], style: { color: '#c00000', bold: true } });
      widths(s, [140, 80, 80, 80, 100, 70]);
      return b;
    },
  },
  {
    id: 'inventory', name: 'Inventory list', category: 'Business', description: 'Stock levels with reorder flags', icon: 'inventory_2',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Inventory';
      put(s, 'A1', [['SKU', 'Item', 'Qty', 'Reorder at', 'Unit cost', 'Value', 'Reorder?']]);
      const items = [['A-100', 'Widget', 120, 50, 2.5], ['A-101', 'Gadget', 35, 40, 7.25], ['B-200', 'Sprocket', 410, 100, 0.9], ['C-300', 'Gizmo', 12, 20, 15]];
      items.forEach((it, i) => put(s, `A${i + 2}`, [[...it, `=C${i + 2}*E${i + 2}`, `=IF(C${i + 2}<=D${i + 2},"Yes","")`]]));
      styleRange(s, 'E2:F5', { numFmt: MONEY });
      table(s, 'Inventory', { r1: 0, c1: 0, r2: 4, c2: 6 }, ['SKU', 'Item', 'Qty', 'Reorder at', 'Unit cost', 'Value', 'Reorder?'], 'TableStyleLight9');
      s.cf.push({ id: 'cf1', ranges: [{ r1: 1, c1: 6, r2: 4, c2: 6 }], type: 'containsText', text: 'Yes', style: { fill: '#ffc7ce', color: '#9c0006' } });
      widths(s, [70, 140, 60, 80, 80, 90, 80]);
      return b;
    },
  },
  {
    id: 'genealogy', name: 'Genealogy research log', category: 'Research', description: 'Sources searched, findings and follow-ups for family history', icon: 'family_history',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Research Log';
      put(s, 'A1', [['Date', 'Ancestor', 'Objective', 'Repository', 'Source / citation', 'Record type', 'Result', 'Notes', 'Follow-up']]);
      put(s, 'A2', [[today(), 'Surname, Given (b. 1850)', 'Find birth record', 'County archive', 'Register of births, vol. 3, p. 41', 'Vital record', 'Found', 'Parents named', 'Order certified copy']]);
      styleRange(s, 'A2:A200', { numFmt: 'yyyy-mm-dd' });
      table(s, 'ResearchLog', { r1: 0, c1: 0, r2: 1, c2: 8 }, ['Date', 'Ancestor', 'Objective', 'Repository', 'Source / citation', 'Record type', 'Result', 'Notes', 'Follow-up'], 'TableStyleMedium4');
      s.validations.push({ id: 'dv1', ranges: [{ r1: 1, c1: 6, r2: 500, c2: 6 }], type: 'list', formula1: '"Found,Not found,Partial,To check"', showDropdown: true });
      s.validations.push({ id: 'dv2', ranges: [{ r1: 1, c1: 5, r2: 500, c2: 5 }], type: 'list', formula1: '"Vital record,Census,Church,Probate,Land,Military,Immigration,Newspaper,Cemetery,DNA,Other"', showDropdown: true });
      widths(s, [90, 170, 150, 130, 230, 110, 80, 200, 150]);
      s.freeze = { rows: 1, cols: 0 };
      const p = new (s.constructor as any)(b, 'Individuals') as Sheet; b.sheets.push(p);
      put(p, 'A1', [['ID', 'Surname', 'Given names', 'Sex', 'Birth date', 'Birth place', 'Death date', 'Death place', 'Father ID', 'Mother ID', 'Spouse ID', 'Sources']]);
      table(p, 'Individuals', { r1: 0, c1: 0, r2: 1, c2: 11 }, ['ID', 'Surname', 'Given names', 'Sex', 'Birth date', 'Birth place', 'Death date', 'Death place', 'Father ID', 'Mother ID', 'Spouse ID', 'Sources'], 'TableStyleLight11');
      widths(p, [50, 110, 140, 40, 100, 150, 100, 150, 70, 70, 70, 200]);
      return b;
    },
  },
  {
    id: 'reading', name: 'Reading and publishing tracker', category: 'Research', description: 'Titles, editions, word counts and status', icon: 'menu_book',
    build: () => {
      const b = new Workbook(); const s = b.sheet; s.name = 'Titles';
      put(s, 'A1', [['Title', 'Author', 'Edition', 'Format', 'Words', 'Pages', 'Status', 'ISBN', 'Notes']]);
      put(s, 'A2', [['Example Title', 'Author Name', '1st', 'Print', 85000, '=ROUND(E2/275,0)', 'Typesetting', '', '']]);
      styleRange(s, 'E2:E200', { numFmt: '#,##0' });
      table(s, 'Titles', { r1: 0, c1: 0, r2: 1, c2: 8 }, ['Title', 'Author', 'Edition', 'Format', 'Words', 'Pages', 'Status', 'ISBN', 'Notes'], 'TableStyleMedium7');
      widths(s, [200, 140, 70, 80, 80, 60, 110, 130, 200]);
      return b;
    },
  },
  {
    id: 'calendar', name: 'Year calendar', category: 'General', description: 'A printable calendar for the current year', icon: 'calendar_month',
    build: () => {
      const b = new Workbook(); const s = b.sheet;
      const y = new Date().getFullYear();
      s.name = String(y);
      put(s, 'A1', [[{ v: y, s: TITLE }]]);
      for (let m = 0; m < 12; m++) {
        const r0 = 2 + Math.floor(m / 3) * 9, c0 = (m % 3) * 8;
        const ref = (r: number, c: number) => `${String.fromCharCode(65 + c0 + c)}${r0 + r + 1}`;
        put(s, ref(0, 0), [[{ f: `DATE(${y},${m + 1},1)`, s: { numFmt: 'mmmm', bold: true, size: 12 } }]]);
        put(s, ref(1, 0), [['S', 'M', 'T', 'W', 'T', 'F', 'S']], { bold: true, hAlign: 'center', color: '#7a8a99' });
        put(s, ref(2, 0), [[{ f: `LET(d,DATE(${y},${m + 1},1),s,d-WEEKDAY(d)+1,g,SEQUENCE(6,7,s),IF(MONTH(g)=${m + 1},g,""))`, s: { numFmt: 'd', hAlign: 'center' } }]]);
        styleRange(s, `${ref(2, 0)}:${ref(7, 6)}`, { numFmt: 'd', hAlign: 'center' });
      }
      for (let c = 0; c < 24; c++) s.colW.set(c, c % 8 === 7 ? 16 : 28);
      return b;
    },
  },
];

export function templateById(id: string): TemplateDef | undefined { return TEMPLATES.find((t) => t.id === id); }
