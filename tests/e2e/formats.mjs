// Opens legacy spreadsheets, documents, slides and mail through the real File > Open path, including
// files with a wrong extension and with none, and checks what lands in the grid.
// Usage: npm run build && npm run e2e:formats   (QL_URL to test a running server; QL_SHOTS=dir for screenshots)
import { spawn } from 'node:child_process';
import { existsSync, readFileSync, mkdirSync } from 'node:fs';
import { chromium } from 'playwright';

const PORT = 4311;
let server = null;
let url = process.env.QL_URL;
if (!url) {
  server = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'ignore', detached: true });
  url = `http://localhost:${PORT}/`;
  for (let i = 0; i < 60; i++) {
    try { if ((await fetch(url)).ok) break; } catch { /* not up yet */ }
    await new Promise((r) => setTimeout(r, 500));
  }
}
const shots = process.env.QL_SHOTS || '';
if (shots) mkdirSync(shots, { recursive: true });

const exe = existsSync('/opt/pw-browsers/chromium') ? { executablePath: '/opt/pw-browsers/chromium' } : {};
const browser = await chromium.launch(exe);
const failures = [];
const check = (name, got, want) => { if (JSON.stringify(got) !== JSON.stringify(want)) failures.push(`${name}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`); };

const cases = [
  { file: 'legacy/prices.wk1', as: 'prices.wk1', expect: { sheets: ['Sheet1'], A1: 'Item', D2: 5, formulaD2: 'B2*C2' } },
  { file: 'legacy/prices.wk1', as: 'prices.xls', note: /named \.xls/, expect: { A1: 'Item', B6: 'yes' } },
  { file: 'legacy/prices.wk1', as: 'prices', note: /no extension/, expect: { A1: 'Item' } },
  { file: 'legacy/sales.wk3', as: 'sales.wk3', expect: { sheets: ['A', 'Totals'], B5: 2230.515 } },
  { file: 'legacy/notebook.wb1', as: 'notebook.dat', expect: { sheets: ['Inputs', 'Results'], B4: 100 } },
  { file: 'legacy/stock.dbf', as: 'stock.dbf', expect: { A1: 'NAME', A2: 'Café' } },
  { file: 'legacy/budget.vc', as: 'budget.vc', expect: { A1: 'ITEM', C4: 20 } },
  { file: 'legacy/testEXCEL_4.xls', as: 'irs.xls', expect: { B13: 168184400 } },
  { file: 'documents/doc.docx', as: 'report.docx', dialog: 'Import into Anne\'s Grid', expect: { sheets: ['Text', 'Table 1', 'Table 2'], A1: 'Quarterly report' } },
  { file: 'documents/doc.docx', as: 'report.xlsx', dialog: 'Import into Anne\'s Grid', expect: { sheets: ['Text', 'Table 1', 'Table 2'] } },
  { file: 'documents/deck.pptx', as: 'deck.pptx', dialog: 'Import into Anne\'s Grid', expect: { sheets: ['Slide 1', 'Slide 2'], A1: 'Plan' } },
  { file: 'documents/mail.eml', as: 'mail', dialog: 'Import into Anne\'s Grid', expect: { sheets: ['Message', 'numbers'], B4: 'Weekly numbers' } },
  { file: 'documents/quick.msg', as: 'quick.msg', dialog: 'Import into Anne\'s Grid', expect: { A1: 'From' } },
  // Formats of the modern-format set, through the same detector
  { file: 'io/b/third-party/tika-testNumbers2013.numbers', as: 'figures', note: /no extension/, expect: {} },
  { file: 'io/b/sample.sqlite', as: 'data.bin', expect: {} },
];

try {
  const p = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  const errs = [];
  p.on('pageerror', (e) => errs.push('pageerror: ' + e.message));
  p.on('console', (m) => { if (m.type() === 'error') errs.push('console: ' + m.text()); });
  await p.addInitScript(() => localStorage.setItem('quillgrid.prefs.v1', JSON.stringify({ firstRun: false })));
  await p.goto(url, { waitUntil: 'networkidle' });
  await p.waitForSelector('.ql-grid-scroller');
  await p.waitForTimeout(500);

  for (const c of cases) {
    const tabsBefore = await p.evaluate(() => ql.docs.length);
    const [chooser] = await Promise.all([p.waitForEvent('filechooser'), p.evaluate(() => { void ql.exec('file.open'); })]);
    await chooser.setFiles({ name: c.as, mimeType: 'application/octet-stream', buffer: readFileSync(`tests/fixtures/${c.file}`) });
    if (c.dialog) {
      const btn = p.locator('.ql-dialog button', { hasText: c.dialog });
      await btn.waitFor({ timeout: 5000 });
      if (shots) await p.screenshot({ path: `${shots}/${c.as}-dialog.png` });
      await btn.click();
    }
    await p.waitForFunction((n) => ql.docs.length > n || (ql.active?.fileName || '').length > 0, tabsBefore, { timeout: 10000 }).catch(() => undefined);
    await p.waitForTimeout(700);
    const toast = await p.evaluate(() => Array.from(document.querySelectorAll('.ql-toast')).map((t) => t.textContent).join(' | '));
    const got = await p.evaluate((refs) => {
      const b = ql.active.book;
      const out = { sheets: b.sheets.map((s) => s.name), title: ql.active.title };
      for (const ref of refs) {
        const m = /^([A-Z]+)(\d+)$/.exec(ref.replace(/^formula/, ''));
        let col = 0;
        for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
        const cell = b.sheets[0].get(+m[2] - 1, col - 1);
        out[ref] = ref.startsWith('formula') ? cell?.f ?? null : cell?.v ?? null;
      }
      return out;
    }, Object.keys(c.expect).filter((k) => k !== 'sheets'));
    for (const [k, v] of Object.entries(c.expect)) check(`${c.as} ${k}`, got[k], v);
    const filled = await p.evaluate(() => ql.active.book.sheets.some((s) => s.cells.size > 0));
    if (!filled) failures.push(`${c.as}: the workbook is empty`);
    if (c.note && !c.note.test(toast)) failures.push(`${c.as}: note missing from toast "${toast}"`);
    if (shots) await p.screenshot({ path: `${shots}/${c.as}.png` });
    await p.evaluate(() => document.querySelectorAll('.ql-toast').forEach((t) => t.remove()));
  }
  if (errs.length) failures.push(...errs);
} finally {
  await browser.close();
  if (server) { try { process.kill(-server.pid); } catch { /* already gone */ } }
}

if (failures.length) {
  console.error('Formats e2e failed:\n  ' + failures.join('\n  '));
  process.exit(1);
}
console.log(`Formats e2e passed (${cases.length} files)`);
