import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { importFile } from '../../src/io/index';
import { rtfBlocks, textToBlocks, blocksToBook, pdfText } from '../../src/io/documents';
import { parseEml, splitMbox, decodeWords, parseMsg } from '../../src/io/email';
import type { Workbook } from '../../src/model/workbook';
import { buildMsg, buildPpt } from './legacy-builders';

const ab = (u8: Uint8Array): ArrayBuffer => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
const load = (p: string) => ab(readFileSync(`tests/fixtures/documents/${p}`));
const text = (b: Workbook, sheet: number, r: number, c = 0) => b.sheets[sheet].get(r, c)?.v;

describe('word-processor documents', () => {
  it.each(['doc.docx', 'doc.odt', 'doc.rtf', 'doc.doc'])('%s: text on one sheet, each table on its own', async (f) => {
    const r = await importFile(load(f), f);
    const b = r.book;
    expect(b.sheets.map((s) => s.name)).toEqual(['Text', 'Table 1', 'Table 2']);
    expect(text(b, 0, 0)).toBe('Quarterly report');
    expect(text(b, 0, 1)).toBe('Sales grew in every region.');
    expect([0, 1, 2].map((c) => text(b, 1, 0, c))).toEqual(['Region', 'Q1', 'Q2']);
    expect(text(b, 1, 1, 1)).toBe(1200);
    expect(text(b, 1, 1, 2)).toBe(1350.5);
    expect(text(b, 1, 2, 0)).toBe('South and East');
    expect(text(b, 2, 1, 1)).toBe('d');
  });

  it('keeps merged table cells (DOCX, ODT)', async () => {
    for (const f of ['doc.docx', 'doc.odt']) {
      const b = (await importFile(load(f), f)).book;
      expect(b.sheets[1].merges).toContainEqual({ r1: 2, c1: 0, r2: 2, c2: 1 });
      expect(text(b, 1, 2, 2)).toBe(900);
    }
  });

  it('parses RTF tables, escapes and Unicode', () => {
    const blocks = rtfBlocks('{\\rtf1\\ansi\\ansicpg1252 {\\fonttbl{\\f0 Arial;}}Caf\\\'e9 \\u8364? price\\par\\trowd\\cellx1000\\cellx2000 A\\cell B\\cell\\row\\trowd 1\\cell 2\\cell\\row\\pard After\\par}');
    expect(blocks[0]).toEqual({ kind: 'para', text: 'Café € price' });
    expect(blocks[1]).toEqual({ kind: 'table', rows: [[{ text: 'A' }, { text: 'B' }], [{ text: '1' }, { text: '2' }]] });
    expect(blocks[2]).toEqual({ kind: 'para', text: 'After' });
  });

  it('turns runs of tab-separated lines into tables', () => {
    const b = blocksToBook(textToBlocks('Title\nName\tQty\nPens\t3\nEnd'), 't');
    expect(b.sheets.map((s) => s.name)).toEqual(['Text', 'Table 1']);
    expect(b.sheets[1].get(1, 1)?.v).toBe(3);
  });

  it('extracts PDF text from uncompressed content streams', () => {
    const pdf = '%PDF-1.4\n1 0 obj << /Length 60 >>\nstream\nBT /F1 12 Tf 72 700 Td (Hello) Tj 0 -14 Td (World \\(2\\)) Tj ET\nendstream\nendobj\n%%EOF';
    expect(pdfText(new TextEncoder().encode(pdf))).toContain('Hello\nWorld (2)');
  });
});

describe('presentations', () => {
  it.each(['deck.pptx', 'deck.odp'])('%s: a sheet per slide with text, tables and notes', async (f) => {
    const b = (await importFile(load(f), f)).book;
    expect(b.sheets.map((s) => s.name)).toEqual(['Slide 1', 'Slide 2']);
    expect(text(b, 0, 0)).toBe('Plan');
    expect(text(b, 0, 1)).toBe('First point');
    expect(text(b, 0, 2)).toBe('Second point');
    expect(text(b, 0, 4)).toBe('Notes');
    expect(text(b, 0, 5)).toBe('Speaker note here');
    expect(text(b, 1, 0)).toBe('Numbers');
    expect([text(b, 1, 2, 0), text(b, 1, 2, 1), text(b, 1, 3, 1)]).toEqual(['Item', 'Cost', 2.5]);
  });

  it('PowerPoint 97-2003: slide titles and body text', async () => {
    const r = await importFile(ab(buildPpt()), 'plan.ppt');
    expect(r.book.sheets[0].name).toBe('Slide 1');
    expect(text(r.book, 0, 0)).toBe('Launch plan');
    expect([1, 2, 3].map((k) => text(r.book, 0, k))).toEqual(['Build', 'Test', 'Ship']);
  });
});

describe('e-mail', () => {
  it('decodes encoded words', () => {
    expect(decodeWords('=?UTF-8?B?SsO8cmdlbg==?= <j@example.com>')).toBe('Jürgen <j@example.com>');
    expect(decodeWords('=?iso-8859-1?Q?Caf=E9_ouvert?=')).toBe('Café ouvert');
  });

  it('EML: headers, body and a CSV attachment as a sheet', async () => {
    const m = parseEml(new Uint8Array(load('mail.eml')));
    expect(m.from).toBe('Jürgen <j@example.com>');
    expect(m.subject).toBe('Weekly numbers');
    expect(m.text).toContain('attached — see below');
    expect(m.attachments.map((a) => a.name)).toEqual(['numbers.csv']);
    const b = (await importFile(load('mail.eml'), 'mail.eml')).book;
    expect(b.sheets.map((s) => s.name)).toEqual(['Message', 'numbers']);
    expect(text(b, 0, 0)).toBe('From');
    expect(text(b, 0, 3, 0)).toBe('Subject');
    expect(text(b, 0, 3, 1)).toBe('Weekly numbers');
    const d = new Date(Date.UTC(2025, 9, 14, 7, 30));
    expect(b.sheets[0].get(2, 1)?.v).toBeCloseTo((d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569, 6); // shown in local time
    expect(text(b, 1, 2, 1)).toBe(5);
  });

  it('MBOX: a row per message', async () => {
    const box = 'From a@x Mon Jan  1 00:00:00 2024\nFrom: a@x\nSubject: One\nDate: Mon, 1 Jan 2024 10:00:00 +0000\n\nFirst\n>From the start\n\nFrom b@x Tue Jan  2 00:00:00 2024\nFrom: b@x\nSubject: Two\n\nSecond\n';
    expect(splitMbox(new TextEncoder().encode(box)).length).toBe(2);
    const b = (await importFile(ab(new TextEncoder().encode(box)), 'inbox.mbox')).book;
    expect(b.sheets[0].name).toBe('Messages');
    expect([text(b, 0, 1, 3), text(b, 0, 2, 3)]).toEqual(['One', 'Two']);
    expect(String(text(b, 0, 1, 5))).toContain('From the start');
  });

  it('Outlook MSG: properties and attachments', async () => {
    const m = parseMsg(buildMsg());
    expect(m).toMatchObject({ subject: 'Budget for March', from: 'Ada Lovelace <ada@example.com>', to: 'Charles Babbage' });
    expect(m.date?.toISOString()).toBe('2024-02-29T12:00:00.000Z');
    const b = (await importFile(ab(buildMsg()), 'budget.msg')).book;
    expect(b.sheets.map((s) => s.name)).toEqual(['Message', 'figures']);
    expect(text(b, 1, 1, 1)).toBe(1200);
    const q = parseMsg(new Uint8Array(load('quick.msg')));
    expect(q.subject).toBe('Test the content transformer');
    expect(q.text).toContain('quick brown fox');
  });
});
