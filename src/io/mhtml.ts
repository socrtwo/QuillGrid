/**
 * MIME HTML web archives (.mht, .mhtml): Excel's "Single File Web Page", and pages saved by
 * browsers. The archive is a multipart/related MIME message; the HTML part (quoted-printable or
 * base64 encoded) is read as an HTML table workbook. Excel puts each sheet in its own part,
 * listed in a frameset, so every HTML part with a table becomes one or more sheets, named after
 * the sheet tabs when Excel's tab list is present.
 */
import { Workbook, Sheet } from '../model/workbook';
import { importHtml } from './html';
import { base64ToBytes, decodeText } from './index';

interface Part { headers: Map<string, string>; body: string }

function parseHeaders(block: string): Map<string, string> {
  const h = new Map<string, string>();
  for (const line of block.replace(/\r?\n[ \t]+/g, ' ').split(/\r?\n/)) {
    const i = line.indexOf(':');
    if (i > 0) h.set(line.slice(0, i).trim().toLowerCase(), line.slice(i + 1).trim());
  }
  return h;
}

function decodeQuotedPrintable(s: string): Uint8Array {
  const t = s.replace(/=\r?\n/g, '');
  const out: number[] = [];
  for (let i = 0; i < t.length; i++) {
    const ch = t.charCodeAt(i);
    if (ch === 61 && /^[0-9A-Fa-f]{2}$/.test(t.slice(i + 1, i + 3))) { out.push(parseInt(t.slice(i + 1, i + 3), 16)); i += 2; }
    else if (ch < 256) out.push(ch);
    else out.push(...new TextEncoder().encode(t[i]));
  }
  return new Uint8Array(out);
}

export function splitMime(text: string): { headers: Map<string, string>; parts: Part[] } {
  const sep = /\r?\n\r?\n/.exec(text);
  const headers = parseHeaders(sep ? text.slice(0, sep.index) : text);
  const body = sep ? text.slice(sep.index + sep[0].length) : '';
  const m = /boundary="?([^";\r\n]+)"?/i.exec(headers.get('content-type') || '');
  if (!m) return { headers, parts: [{ headers, body }] };
  const parts: Part[] = [];
  for (const chunk of body.split('--' + m[1]).slice(1)) {
    if (/^--/.test(chunk)) break;
    const c = chunk.replace(/^\r?\n/, '');
    const s2 = /\r?\n\r?\n/.exec(c);
    if (!s2) continue;
    parts.push({ headers: parseHeaders(c.slice(0, s2.index)), body: c.slice(s2.index + s2[0].length) });
  }
  return { headers, parts };
}

function partText(p: Part): string {
  const enc = (p.headers.get('content-transfer-encoding') || '').toLowerCase();
  const cs = /charset="?([^";]+)"?/i.exec(p.headers.get('content-type') || '')?.[1];
  const bytes = enc === 'base64' ? base64ToBytes(p.body) : enc === 'quoted-printable' ? decodeQuotedPrintable(p.body) : new Uint8Array([...p.body].map((ch) => ch.charCodeAt(0) & 255));
  if (cs) { try { return new TextDecoder(cs).decode(bytes); } catch { /* unknown charset */ } }
  return decodeText(bytes);
}

export function importMhtml(data: ArrayBuffer, warnings: string[] = []): Workbook {
  const raw = new TextDecoder('latin1').decode(new Uint8Array(data));
  const { parts } = splitMime(raw);
  const html = parts.filter((p) => /text\/html/i.test(p.headers.get('content-type') || ''));
  if (!html.length) throw new Error('This web archive has no HTML page.');
  // Excel's tab list (tabstrip.htm) names the sheet parts.
  const names = new Map<string, string>();
  for (const p of html) {
    const t = partText(p);
    for (const m of t.matchAll(/<a\s+href="?([^"\s>]+)"?[^>]*>(?:<[^>]+>)*([^<]+)/gi)) if (/sheet\d+\.htm/i.test(m[1])) names.set(m[1].replace(/^.*\//, '').toLowerCase(), m[2].trim());
  }
  const book = new Workbook(false);
  for (const p of html) {
    const text = partText(p);
    if (!/<table/i.test(text)) continue;
    const loc = (p.headers.get('content-location') || '').replace(/^.*\//, '').toLowerCase();
    const part = importHtml(text, warnings);
    part.sheets.forEach((s, i) => {
      let name = (names.get(loc) && part.sheets.length === 1 ? names.get(loc)! : s.name) || 'Sheet';
      if (i === 0 && names.get(loc)) name = names.get(loc)!;
      while (book.sheetByName(name)) name = name.slice(0, 28) + '_' + (book.sheets.length + 1);
      // Move the sheet with its styles into the combined workbook.
      const json = s.toJSON();
      const ns = Sheet.fromJSON(book, { ...json, id: undefined, name });
      for (const cell of ns.cells.values()) if (cell.s) cell.s = book.styleId(part.styles[cell.s] || {});
      book.sheets.push(ns);
    });
  }
  if (!book.sheets.length) {
    warnings.push('The web archive has no tables; its text was placed in column A.');
    return importHtml(partText(html[0]), warnings);
  }
  book.clearHistory();
  return book;
}
