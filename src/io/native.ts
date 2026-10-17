/**
 * Native QuillGrid formats:
 *   .json  the workbook model as JSON (Workbook.toJSON, errors as {"#err": code})
 *   .qgw   a ZIP with workbook.json, pictures extracted to media/
 *   .qgwe  the same JSON encrypted with a password (PBKDF2 + AES-256-GCM, src/util/crypto.ts)
 */
import JSZip from 'jszip';
import { Workbook, Sheet, errReplacer } from '../model/workbook';
import { cellKey } from '../model/address';
import type { Scalar } from '../model/types';
import { encryptJson, decryptJson } from '../util/crypto';
import { parseDataUrl, bytesToBase64, extForMime, mimeForExt, PasswordError } from './index';

const plain = (book: Workbook): any => JSON.parse(JSON.stringify(book.toJSON(), errReplacer));

function finish(book: Workbook): Workbook {
  book.clearHistory();
  return book;
}

// --------------------------------------------------------------------------------------------
// JSON

export function exportJson(book: Workbook): Blob {
  return new Blob([JSON.stringify(book.toJSON(), errReplacer)], { type: 'application/json' });
}

export function importJson(text: string): Workbook {
  const j = JSON.parse(text.replace(/^﻿/, ''));
  if (j && typeof j === 'object' && !Array.isArray(j) && (j.format === 'quillgrid' || Array.isArray(j.sheets))) return finish(Workbook.fromJSON(j.workbook || j));
  if (j && j.qgw && j.workbook) return finish(Workbook.fromJSON(j.workbook));
  // Generic JSON data: an array of rows (arrays) or records (objects).
  const rows: any[] = Array.isArray(j) ? j : Array.isArray(j?.data) ? j.data : Array.isArray(j?.rows) ? j.rows : [j];
  const book = new Workbook(false);
  const sheet = new Sheet(book, 'Sheet1');
  book.sheets.push(sheet);
  const val = (x: any): Scalar => (x == null ? null : typeof x === 'number' || typeof x === 'boolean' || typeof x === 'string' ? x : JSON.stringify(x));
  if (rows.every((r) => Array.isArray(r))) {
    rows.forEach((r: any[], i) => r.forEach((x, c) => { const v = val(x); if (v != null) sheet.cells.set(cellKey(i, c), { v }); }));
  } else {
    const keys: string[] = [];
    for (const r of rows) if (r && typeof r === 'object') for (const k of Object.keys(r)) if (!keys.includes(k)) keys.push(k);
    const bold = book.styleId({ bold: true });
    keys.forEach((k, c) => sheet.cells.set(cellKey(0, c), { v: k, s: bold }));
    rows.forEach((r, i) => {
      if (r && typeof r === 'object') keys.forEach((k, c) => { const v = val(r[k]); if (v != null) sheet.cells.set(cellKey(i + 1, c), { v }); });
      else { const v = val(r); if (v != null) sheet.cells.set(cellKey(i + 1, 0), { v }); }
    });
  }
  return finish(book);
}

// --------------------------------------------------------------------------------------------
// .qgw (zip)

export async function exportQlw(book: Workbook): Promise<Blob> {
  const zip = new JSZip();
  zip.file('mimetype', 'application/x-quillgrid', { compression: 'STORE' });
  const json = plain(book);
  let n = 0;
  for (const sh of json.sheets || []) {
    for (const d of sh.drawings || []) {
      if (typeof d.src !== 'string' || !d.src.startsWith('data:')) continue;
      const img = parseDataUrl(d.src);
      if (!img) continue;
      const name = `media/image${++n}.${extForMime(img.mime)}`;
      zip.file(name, img.bytes);
      d.src = name;
    }
  }
  const vba: Uint8Array | undefined = (book as any).vbaProject;
  if (vba instanceof Uint8Array) zip.file('vbaProject.bin', vba);
  zip.file('workbook.json', JSON.stringify({ qgw: 1, app: 'QuillGrid', saved: new Date().toISOString(), workbook: json }));
  const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  return new Blob([data as BlobPart], { type: 'application/x-quillgrid' });
}

export async function importQlw(data: ArrayBuffer): Promise<Workbook> {
  const zip = await JSZip.loadAsync(data);
  const f = zip.file('workbook.json');
  if (!f) throw new Error('workbook.json is missing from this QuillGrid file');
  const env = JSON.parse(await f.async('string'));
  const json = env.workbook || env;
  for (const sh of json.sheets || []) {
    for (const d of sh.drawings || []) {
      if (typeof d.src !== 'string' || !d.src.startsWith('media/')) continue;
      const m = zip.file(d.src);
      if (!m) { d.src = ''; continue; }
      const bytes = await m.async('uint8array');
      d.src = `data:${mimeForExt(d.src.slice(d.src.lastIndexOf('.') + 1))};base64,${bytesToBase64(bytes)}`;
    }
  }
  const book = Workbook.fromJSON(json);
  const vba = zip.file('vbaProject.bin');
  if (vba) (book as any).vbaProject = await vba.async('uint8array');
  return finish(book);
}

// --------------------------------------------------------------------------------------------
// .qgwe (encrypted)

export async function exportQlwe(book: Workbook, password: string): Promise<Blob> {
  const bytes = await encryptJson({ qgw: 1, app: 'QuillGrid', workbook: plain(book) }, password);
  return new Blob([bytes as BlobPart], { type: 'application/x-quillgrid-encrypted' });
}

export async function importQlwe(data: ArrayBuffer, password: string): Promise<Workbook> {
  let env: any;
  try {
    env = await decryptJson(data, password);
  } catch (e) {
    if (e instanceof Error && /Not an encrypted/.test(e.message)) throw e;
    throw new PasswordError('The password is incorrect.', true);
  }
  return finish(Workbook.fromJSON(env.workbook || env));
}
