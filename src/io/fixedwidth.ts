/**
 * Fixed-width text (.prn, Excel's "Formatted Text (Space delimited)", report listings).
 *
 * Column boundaries are found where every non-blank line has a space (sniff-b.ts
 * fixedWidthColumns); each slice is trimmed and typed like a CSV field (numbers, dates,
 * booleans), using the decimal and date conventions detected from the data. A file whose
 * columns do not line up is read as delimited text instead (space-delimited .prn files).
 */
import type { Workbook } from '../model/workbook';
import type { CsvOptions } from './index';
import { decodeAuto } from './encoding';
import { fixedWidthColumns } from './sniff-b';
import { rowsToBook, importCsv, detectLocale } from './csv';

/** Splits lines at the given column starts (trimmed fields). */
export function splitFixedWidth(lines: string[], starts: number[]): string[][] {
  return lines.map((l) => starts.map((s, k) => l.slice(s, k + 1 < starts.length ? starts[k + 1] : undefined).trim()));
}

export function importFixedWidth(data: ArrayBuffer, name: string, opts: CsvOptions = {}, warnings: string[] = []): Workbook {
  const { text } = decodeAuto(new Uint8Array(data), opts.encoding);
  const lines = text.replace(/\t/g, '        ').split(/\r\n|\r|\n/);
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  const sample = lines.slice(0, 2000);
  const starts = fixedWidthColumns(sample);
  if (starts.length < 2) {
    warnings.push('The columns of this text file do not line up; it was read as delimited text.');
    return importCsv(data, name, opts);
  }
  // Lines after the sample may hold text in a gap; keep the boundaries but never lose text.
  const rows = splitFixedWidth(lines, starts);
  const o: CsvOptions = { ...opts };
  if (o.decimal == null || o.dateOrder == null) {
    const loc = detectLocale(rows, ' ');
    if (o.decimal == null) { o.decimal = loc.decimal; o.thousands ??= loc.thousands; }
    if (o.dateOrder == null && loc.dateOrder) o.dateOrder = loc.dateOrder;
  }
  const base = name.replace(/^.*[\\/]/, '').replace(/\.[^.]*$/, '') || 'Sheet1';
  const book = rowsToBook(rows, o, base);
  // Column widths follow the text layout (about 7 px per character).
  const sheet = book.sheets[0];
  starts.forEach((s, k) => {
    const w = (k + 1 < starts.length ? starts[k + 1] : Math.max(...sample.map((l) => l.length))) - s;
    sheet.colW.set(k, Math.max(40, Math.min(400, Math.round(w * 7 + 10))));
  });
  return book;
}
