import { describe, it, expect } from 'vitest';
import { createHash } from 'node:crypto';
import { Workbook } from '../../src/model/workbook';
import { defaultPrefs } from '../../src/prefs';
import { getCommand } from '../../src/commands/registry';
import { cellKey } from '../../src/model/address';
import { SpellSession, spellTargets, wordsIn, replaceInCell, dialectFor, type SpellChecker } from '../../src/proof/proofing';
import { hashPassword, verifyPassword, legacyPasswordHash, workbookStats, accessibilityIssues, contrastRatio } from '../../src/commands/review';
import '../../src/commands/pagelayout';

function fakeApp() {
  const book = new Workbook();
  const app: any = {
    book,
    get sheet() { return book.sheet; },
    prefs: { ...defaultPrefs(), author: 'Tester' },
    tab: { id: 't1', readOnly: false, title: 'Book', fileName: '', viewMode: 'normal' },
    isReadOnly: () => false,
    exec(id: string, arg?: any) { return getCommand(id)!.run(app, arg); },
    ui: { paneOpen: () => true, openPane() {}, togglePane() {}, anchorFor: () => document.body },
    grid: { editing: false, select() {}, scrollIntoView() {}, invalidate() {}, objects: { showAllComments: false, hoverComment() {}, position() {} } },
    savePrefs() {},
    refresh() {},
  };
  return app;
}

const stub = (bad: string[]): SpellChecker => ({
  async check(words) { return new Set(words.filter((w) => bad.includes(w))); },
  async suggest() { return ['the']; },
  async add() {},
});

describe('spelling walk', () => {
  it('finds words in text, skipping codes, acronyms and addresses', () => {
    expect(wordsIn("Teh cat's NASA A4 http://x.example/teh www.wrold.com hello").map((w) => w.word)).toEqual(['Teh', 'cat', 'hello']);
    expect(dialectFor('document', 'en-GB')).toBe('en-GB');
    expect(dialectFor('any', 'en-US')).toBe('any');
  });

  it('walks text cells, skipping formulas, with change all and ignore', async () => {
    const b = new Workbook();
    const s = b.sheet;
    b.transact('t', () => {
      s.set(0, 0, { v: 'teh cat' });
      s.set(0, 1, { v: 'hello wrold' });
      s.set(1, 0, { f: 'A1&"teh"', v: 'teh catteh' });
      s.set(2, 0, { v: 'teh again' });
      s.set(3, 0, { v: 42 });
    });
    const targets = spellTargets(s, null, { r: 0, c: 0 }, false);
    expect(targets.map((t) => [t.r, t.c])).toEqual([[0, 0], [0, 1], [2, 0]]);
    const write = (t: any, patch: any) => { b.transact('Spelling', () => t.sheet.patch(t.r, t.c, patch)); return true; };
    const ses = new SpellSession(targets, stub(['teh', 'wrold']), { dict: 'any', ignored: new Set(), dictionary: new Set(), write });
    await ses.start();
    const a = await ses.next();
    expect(a).toMatchObject({ word: 'teh', start: 0, end: 3 });
    ses.changeAll('the');
    expect(s.value(0, 0)).toBe('the cat');
    const bIssue = await ses.next();
    expect(bIssue).toMatchObject({ word: 'wrold', target: { r: 0, c: 1 } });
    ses.ignoreOnce();
    expect(await ses.next()).toBeNull();
    expect(s.value(2, 0)).toBe('the again');
    expect(ses.changed).toBe(2);
    // Undo restores the original text.
    b.undo();
    expect(s.value(2, 0)).toBe('teh again');
  });

  it('ignore all, the personal dictionary and wrapping to the start', async () => {
    const b = new Workbook();
    const s = b.sheet;
    b.transact('t', () => { s.set(0, 0, { v: 'zorp one' }); s.set(5, 0, { v: 'blarg zorp' }); });
    const ignored = new Set<string>(), dictionary = new Set<string>();
    const ses = new SpellSession(spellTargets(s, null, { r: 5, c: 0 }, false), stub(['zorp', 'blarg']), { dict: 'any', ignored, dictionary, write: () => true });
    expect((await ses.next())!.word).toBe('blarg');
    await ses.addToDictionary();
    expect(dictionary.has('blarg')).toBe(true);
    expect((await ses.next())!.word).toBe('zorp');
    ses.ignoreAll();
    expect(await ses.next()).toBeNull();
    ses.extend(spellTargets(s, null).filter((t) => t.r < 5));
    expect(await ses.next()).toBeNull(); // "zorp" is ignored everywhere now
    expect(ignored.has('zorp')).toBe(true);
  });

  it('keeps rich text runs when replacing inside one run', () => {
    const p = replaceInCell({ v: 'Big teh end', rich: [{ t: 'Big ', s: { bold: true } }, { t: 'teh end' }] }, 4, 7, 'the');
    expect(p.v).toBe('Big the end');
    expect(p.rich![1].t).toBe('the end');
  });
});

describe('protection passwords', () => {
  it('hashes with a salt and verifies', async () => {
    const h1 = await hashPassword('secret', 1000);
    expect(h1).toMatch(/^SHA-256:[^:]+:1000:[^:]+$/);
    expect(h1).not.toContain('secret');
    expect(await hashPassword('secret', 1000)).not.toBe(h1); // random salt
    expect(await verifyPassword('secret', h1)).toBe(true);
    expect(await verifyPassword('Secret', h1)).toBe(false);
    expect(await verifyPassword('anything', undefined)).toBe(true);
  });

  it('verifies spreadsheet-file hashes (SHA-512 iterated and legacy)', async () => {
    const salt = Buffer.from('0123456789abcdef');
    const pw = Buffer.from('pass', 'utf16le');
    let hv = createHash('sha512').update(Buffer.concat([salt, pw])).digest();
    for (let i = 0; i < 50; i++) { const it = Buffer.alloc(4); it.writeUInt32LE(i); hv = createHash('sha512').update(Buffer.concat([hv, it])).digest(); }
    const stored = `SHA-512:${salt.toString('base64')}:50:${hv.toString('base64')}`;
    expect(await verifyPassword('pass', stored)).toBe(true);
    expect(await verifyPassword('pass2', stored)).toBe(false);
    expect(legacyPasswordHash('password')).toBe('83AF');
    expect(await verifyPassword('password', '83af')).toBe(true);
  });
});

describe('comment and note commands', () => {
  it('creates, replies, resolves and deletes comments with undo', async () => {
    const app = fakeApp();
    const s = app.sheet;
    await app.exec('review.newComment', { r: 0, c: 0, text: 'Check this' });
    const k = cellKey(0, 0);
    expect(s.comments.get(k)).toMatchObject({ author: 'Tester', text: 'Check this', replies: [] });
    await app.exec('review.replyComment', { r: 0, c: 0, text: 'Done' });
    expect(s.comments.get(k).replies).toHaveLength(1);
    app.exec('review.resolveComment', { r: 0, c: 0 });
    expect(s.comments.get(k).resolved).toBe(true);
    app.exec('review.deleteComment', { r: 0, c: 0 });
    expect(s.comments.has(k)).toBe(false);
    app.book.undo();
    expect(s.comments.get(k).resolved).toBe(true);
  });

  it('notes are separate from comment threads', async () => {
    const app = fakeApp();
    const s = app.sheet;
    await app.exec('review.newNote', { r: 1, c: 1, text: 'A note' });
    expect(s.comments.get(cellKey(1, 1))).toMatchObject({ note: true, text: 'A note', author: 'Tester' });
    await app.exec('review.editNote', { r: 1, c: 1, text: 'Changed' });
    expect(s.comments.get(cellKey(1, 1)).text).toBe('Changed');
    await app.exec('review.newNote', { r: 1, c: 1, text: 'ignored?' });
    expect(s.comments.get(cellKey(1, 1)).text).toBe('ignored?'); // a second New Note edits the existing note
    app.exec('review.showAllNotes');
    expect(app.grid.objects.showAllComments).toBe(true);
    app.exec('review.deleteNote', { r: 1, c: 1 });
    expect(s.comments.size).toBe(0);
  });

  it('refuses comment changes on a protected sheet unless objects are allowed', async () => {
    const app = fakeApp();
    app.sheet.protection = { on: true, allow: ['selectLocked', 'selectUnlocked'] };
    await app.exec('review.newComment', { r: 0, c: 0, text: 'x' });
    expect(app.sheet.comments.size).toBe(0);
    app.sheet.protection.allow.push('objects');
    await app.exec('review.newComment', { r: 0, c: 0, text: 'x' });
    expect(app.sheet.comments.size).toBe(1);
  });

  it('allow edit ranges unlock locked cells on a protected sheet', () => {
    const app = fakeApp();
    const s = app.sheet;
    s.protection = { on: true, allow: [] };
    s.editRanges = [{ title: 'Inputs', ref: 'A1:B2' }];
    expect(app.exec('review.canEditRange', { r: 1, c: 1 })).toBe(true);
    expect(app.exec('review.canEditRange', { r: 5, c: 5 })).toBe(false);
  });
});

describe('statistics, accessibility and page layout commands', () => {
  it('counts workbook contents', () => {
    const app = fakeApp();
    const b = app.book, s = app.sheet;
    b.transact('t', () => { s.set(0, 0, { v: 1 }); s.set(4, 2, { f: 'A1*2', v: 2 }); b.recordSheet(s); s.comments.set(0, { author: 'a', text: 'n', time: 0, note: true }); });
    const st = workbookStats(b);
    expect(st).toMatchObject({ sheets: 1, cells: 2, formulas: 1, notes: 1, comments: 0 });
    expect(st.perSheet[0].lastCell).toBe('C5');
  });

  it('finds accessibility issues', () => {
    const app = fakeApp();
    const b = app.book, s = app.sheet;
    b.transact('t', () => {
      b.recordSheet(s);
      s.set(0, 0, { v: 'low', s: b.styleId({ color: '#777777', fill: '#888888' }) });
      s.set(1, 0, { v: 'click here', link: 'https://example.com' });
      s.merges.push({ r1: 3, c1: 0, r2: 3, c2: 2 });
      s.drawings.push({ id: 'p1', kind: 'image', at: { r: 5, c: 0, dx: 0, dy: 0, w: 50, h: 50 }, src: 'data:image/png;base64,' });
    });
    const rules = accessibilityIssues(b).map((i) => i.rule);
    expect(rules).toEqual(expect.arrayContaining(['alt', 'merged', 'contrast', 'link']));
    expect(contrastRatio('#000000', '#ffffff')).toBeCloseTo(21, 0);
  });

  it('page layout commands change print settings undoably', () => {
    const app = fakeApp();
    const s = app.sheet;
    app.exec('layout.orientation', 'landscape');
    app.exec('layout.margins', 'narrow');
    app.exec('layout.fitWidth', 1);
    app.exec('layout.printGridlines');
    expect(s.print).toMatchObject({ orientation: 'landscape', fitWide: 1, gridlines: true });
    expect(s.print.margins.left).toBe(0.25);
    app.exec('layout.scale', 80);
    expect(s.print).toMatchObject({ scale: 80, fitWide: undefined });
    s.sel = { r: 4, c: 2, ranges: [{ r1: 4, c1: 2, r2: 4, c2: 2 }] };
    app.exec('layout.breaks', 'insert');
    expect([s.rowBreaks, s.colBreaks]).toEqual([[4], [2]]);
    app.exec('layout.breaks', 'reset');
    expect(s.rowBreaks).toEqual([]);
    s.sel = { r: 0, c: 0, ranges: [{ r1: 0, c1: 0, r2: 9, c2: 3 }] };
    app.exec('layout.printArea', 'set');
    expect(s.print.area).toBe('$A$1:$D$10');
    app.exec('layout.themes', 'Harbor');
    expect(app.book.theme.name).toBe('Harbor');
    expect(app.book.styles[0].font).toBe('Open Sans');
    app.book.undo();
    expect(app.book.theme.name).toBe('Quill');
    expect(app.book.styles[0].font).toBeUndefined();
    app.book.undo();
    expect(s.print.area).toBeUndefined();
  });
});

describe('sha256', () => {
  it('matches node crypto', async () => {
    const { sha256 } = await import('../../src/commands/review');
    for (const n of [0, 3, 36, 55, 56, 64, 200]) {
      const d = new Uint8Array(n).map((_x, i) => (i * 7 + 3) & 255);
      expect(Buffer.from(sha256(d)).toString('hex')).toBe(createHash('sha256').update(d).digest('hex'));
    }
  });
  it('hashes a password at the default strength quickly', async () => {
    const t = Date.now();
    await hashPassword('secret');
    expect(Date.now() - t).toBeLessThan(3000);
  });
});
