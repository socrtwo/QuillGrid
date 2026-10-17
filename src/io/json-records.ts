/**
 * JSON data as spreadsheets: arrays of records (objects), arrays of rows (arrays), column
 * objects ({"name": [...], "age": [...]}), keyed records ({"id1": {...}}), objects that wrap one
 * or more record arrays ({"data": [...]}, each array its own sheet), and NDJSON / JSON Lines.
 *
 * Nested objects flatten to dotted column names ("address.city"); arrays of scalars inside a
 * record join with ", "; deeper structures are kept as JSON text. ISO 8601 date strings become
 * dates. The header row is bold and frozen.
 */
import { Workbook, Sheet } from '../model/workbook';
import { cellKey, MAX_COLS, MAX_ROWS } from '../model/address';
import type { Cell } from '../model/types';
import { partsToSerial } from '../model/dates';

type Json = null | boolean | number | string | Json[] | { [k: string]: Json };
const isObj = (v: Json): v is { [k: string]: Json } => !!v && typeof v === 'object' && !Array.isArray(v);
const isScalar = (v: Json) => v === null || typeof v !== 'object';

function flatten(v: Json, prefix: string, out: Map<string, Json>, depth = 0): void {
  if (isObj(v) && depth < 6) {
    const keys = Object.keys(v);
    if (!keys.length) { out.set(prefix, ''); return; }
    for (const k of keys) flatten(v[k], prefix ? `${prefix}.${k}` : k, out, depth + 1);
    return;
  }
  out.set(prefix || 'value', v);
}

const ISO = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?(Z|[+-]\d{2}:?\d{2})?)?$/;

function toCell(v: Json, dateStyle: () => number, dateTimeStyle: () => number): Cell | null {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? { v } : { v: String(v) };
  if (typeof v === 'boolean') return { v };
  if (typeof v === 'string') {
    const m = ISO.exec(v);
    if (m && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31) {
      const d = partsToSerial(+m[1], +m[2], +m[3]);
      if (m[4] == null) return { v: d, s: dateStyle() };
      return { v: d + ((+m[4]) * 3600 + (+m[5]) * 60 + (+(m[6] || 0))) / 86400, s: dateTimeStyle() };
    }
    return { v };
  }
  if (Array.isArray(v) && v.every(isScalar)) return { v: v.map((x) => (x === null ? '' : String(x))).join(', ') };
  return { v: JSON.stringify(v) };
}

interface Table { name: string; rows: Json[] }

/** Finds the tables in a parsed JSON document. */
function findTables(doc: Json, name: string): Table[] {
  if (Array.isArray(doc)) return [{ name, rows: doc }];
  if (isObj(doc)) {
    const keys = Object.keys(doc);
    const arrays = keys.filter((k) => Array.isArray(doc[k]) && (doc[k] as Json[]).length && (doc[k] as Json[]).some((x) => isObj(x) || Array.isArray(x)));
    if (arrays.length) return arrays.map((k) => ({ name: k, rows: doc[k] as Json[] }));
    // Columns: {"a": [1,2], "b": [3,4]}
    const cols = keys.filter((k) => Array.isArray(doc[k]));
    if (cols.length >= 1 && cols.length === keys.length) {
      const n = Math.max(...cols.map((k) => (doc[k] as Json[]).length));
      const rows: Json[] = [];
      for (let i = 0; i < n; i++) { const r: { [k: string]: Json } = {}; for (const k of cols) r[k] = (doc[k] as Json[])[i] ?? null; rows.push(r); }
      return [{ name, rows }];
    }
    // Keyed records: {"id1": {...}, "id2": {...}}
    if (keys.length > 1 && keys.every((k) => isObj(doc[k]))) return [{ name, rows: keys.map((k) => ({ key: k, ...(doc[k] as object) }) as Json) }];
    return [{ name, rows: [doc] }];
  }
  return [{ name, rows: [doc] }];
}

function fillSheet(sheet: Sheet, rows: Json[], warnings: string[]): void {
  const book = sheet.book;
  let ds = 0, dts = 0, hs = 0;
  const dateStyle = () => ds || (ds = book.styleId({ numFmt: 'yyyy-mm-dd' }));
  const dateTimeStyle = () => dts || (dts = book.styleId({ numFmt: 'yyyy-mm-dd hh:mm:ss' }));
  const put = (r: number, c: number, cell: Cell | null) => { if (cell && r < MAX_ROWS && c < MAX_COLS) sheet.cells.set(cellKey(r, c), cell); };
  if (rows.length > MAX_ROWS - 1) { warnings.push(`Only the first ${MAX_ROWS - 1} records fit on a sheet.`); rows = rows.slice(0, MAX_ROWS - 1); }
  const objs = rows.filter(isObj).length;
  if (objs >= rows.length / 2 && objs > 0) {
    const cols: string[] = [];
    const index = new Map<string, number>();
    const flat = rows.map((row) => {
      const m = new Map<string, Json>();
      if (isObj(row)) flatten(row, '', m); else m.set('value', row);
      for (const k of m.keys()) if (!index.has(k)) { index.set(k, cols.length); cols.push(k); }
      return m;
    });
    hs = book.styleId({ bold: true });
    cols.forEach((k, c) => put(0, c, { v: k, s: hs }));
    flat.forEach((m, i) => { for (const [k, v] of m) put(i + 1, index.get(k)!, toCell(v, dateStyle, dateTimeStyle)); });
    sheet.freeze = { rows: 1, cols: 0 };
    cols.forEach((k, c) => sheet.colW.set(c, Math.max(64, Math.min(320, 8 * Math.max(k.length, ...flat.slice(0, 200).map((m) => String(m.get(k) ?? '').length)) + 12))));
    return;
  }
  rows.forEach((row, r) => {
    if (Array.isArray(row)) row.forEach((v, c) => put(r, c, toCell(v, dateStyle, dateTimeStyle)));
    else put(r, 0, toCell(row, dateStyle, dateTimeStyle));
  });
}

function build(tables: Table[], warnings: string[]): Workbook {
  const book = new Workbook(false);
  for (const t of tables) {
    let nm = (t.name || 'Sheet1').replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet1';
    while (book.sheetByName(nm)) nm = nm.slice(0, 28) + '_' + (book.sheets.length + 1);
    const sheet = new Sheet(book, nm);
    book.sheets.push(sheet);
    fillSheet(sheet, t.rows, warnings);
  }
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
  book.clearHistory();
  return book;
}

const baseName = (name: string) => name.replace(/^.*[\\/]/, '').replace(/\.[^.]*$/, '') || 'Sheet1';

export function importJsonRecords(text: string, name = '', warnings: string[] = []): Workbook {
  const src = text.replace(/^﻿/, '');
  let doc: Json;
  try { doc = JSON.parse(src); } catch (e) {
    // Several JSON values, one per line (NDJSON without the extension).
    if (src.trim().split(/\r?\n/).length > 1) return importNdjson(src, name, warnings);
    throw new Error(`This JSON file is not valid: ${e instanceof Error ? e.message : e}`);
  }
  return build(findTables(doc, baseName(name)), warnings);
}

export function importNdjson(text: string, name = '', warnings: string[] = []): Workbook {
  const rows: Json[] = [];
  let bad = 0;
  for (const line of text.replace(/^﻿/, '').split(/\r?\n/)) {
    const t = line.trim();
    if (!t) continue;
    try { rows.push(JSON.parse(t)); } catch { bad++; }
  }
  if (bad) warnings.push(`${bad} line${bad > 1 ? 's were' : ' was'} not valid JSON and ${bad > 1 ? 'were' : 'was'} skipped.`);
  return build([{ name: baseName(name), rows }], warnings);
}

