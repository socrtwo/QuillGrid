/**
 * Table style palette: Light 1–21, Medium 1–28 and Dark 1–11, generated from the theme accents so the
 * names match the ones used in .xlsx files (TableStyleMedium2 and so on) and look consistent.
 */
import type { TableDef } from './types';
import type { ThemeDef } from './workbook';
import { mix } from '../grid/cf';

export interface TableLook {
  headerFill?: string; headerColor?: string; headerBold?: boolean;
  band1?: string; band2?: string;
  colBand1?: string;
  border?: string; innerBorder?: string;
  totalFill?: string; totalBorder?: string;
  firstColFill?: string; lastColFill?: string;
  text?: string;
}

export const TABLE_STYLE_NAMES: string[] = [
  ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
  ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
  ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
];

function accent(theme: ThemeDef, i: number): string {
  // index 0 = neutral (dark1/black), 1..6 = accents
  if (i === 0) return '#404040';
  return theme.colors[3 + i] || '#2f6fb0';
}

export function tableLook(name: string, theme: ThemeDef): TableLook {
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name || '');
  if (!m) return tableLook('TableStyleMedium2', theme);
  const fam = m[1], n = +m[2];
  const a = accent(theme, (n - 1) % 7);
  const light = (t: number) => mix(a, '#ffffff', t);
  const dark = (t: number) => mix(a, '#000000', t);
  if (fam === 'Light') {
    const g = Math.floor((n - 1) / 7); // 0: banded no header fill, 1: bordered, 2: header line
    if (g === 0) return { headerBold: true, headerColor: n === 1 ? '#000000' : dark(0.25), band1: light(0.82), border: a, text: n === 1 ? '#000000' : dark(0.25) };
    if (g === 1) return { headerFill: a, headerColor: '#ffffff', headerBold: true, band1: light(0.82), border: a, innerBorder: light(0.6) };
    return { headerBold: true, band1: light(0.82), border: a, innerBorder: a };
  }
  if (fam === 'Medium') {
    const g = Math.floor((n - 1) / 7);
    if (g === 0) return { headerFill: a, headerColor: '#ffffff', headerBold: true, band1: light(0.8), band2: '#ffffff', border: light(0.4), innerBorder: light(0.4), totalBorder: a };
    if (g === 1) return { headerFill: a, headerColor: '#ffffff', headerBold: true, band1: light(0.65), band2: light(0.82), innerBorder: '#ffffff', totalFill: a };
    if (g === 2) return { headerFill: '#000000', headerColor: '#ffffff', headerBold: true, band1: light(0.65), band2: light(0.82), border: '#000000', totalFill: light(0.4) };
    return { headerFill: a, headerColor: '#ffffff', headerBold: true, band1: light(0.7), band2: light(0.85), border: a, innerBorder: light(0.5) };
  }
  const g = Math.floor((n - 1) / 7);
  if (g === 0) return { headerFill: '#000000', headerColor: '#ffffff', headerBold: true, band1: dark(0.25), band2: a, text: '#ffffff', totalFill: dark(0.5) };
  return { headerFill: dark(0.4), headerColor: '#ffffff', headerBold: true, band1: light(0.6), band2: light(0.8), totalFill: dark(0.4) };
}

/** Fill/font/border overrides for a cell inside a table. */
export function tableCellLook(t: TableDef, r: number, c: number, theme: ThemeDef): { fill?: string; color?: string; bold?: boolean; borderBottom?: string; borderTop?: string } | null {
  const rg = t.range;
  if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2) return null;
  const look = tableLook(t.style, theme);
  const o = t.opts;
  const isHeader = o.header && r === rg.r1;
  const isTotal = o.totals && r === rg.r2;
  if (isHeader) return { fill: look.headerFill, color: look.headerColor, bold: look.headerBold, borderBottom: look.border };
  if (isTotal) return { fill: look.totalFill, bold: true, borderTop: look.totalBorder || look.border, color: look.totalFill ? '#ffffff' : undefined };
  const dataRow = r - rg.r1 - (o.header ? 1 : 0);
  const out: { fill?: string; color?: string; bold?: boolean; borderBottom?: string } = {};
  if (o.banded) out.fill = dataRow % 2 === 0 ? look.band1 : look.band2;
  if (o.bandedCols && (c - rg.c1) % 2 === 0) out.fill = look.band1;
  if ((o.firstCol && c === rg.c1) || (o.lastCol && c === rg.c2)) out.bold = true;
  if (look.text) out.color = look.text;
  if (look.innerBorder) out.borderBottom = look.innerBorder;
  return out;
}

export const DEFAULT_TABLE_OPTS = { header: true, totals: false, banded: true, bandedCols: false, firstCol: false, lastCol: false, filterButton: true };
