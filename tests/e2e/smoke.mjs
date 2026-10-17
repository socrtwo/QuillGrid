// Browser smoke test: serves dist/ with `vite preview`, types a small workbook, and checks the results.
// Usage: npm run build && npm run e2e   (set QL_URL to test an already-running server instead)
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { chromium } from 'playwright';

const PORT = 4179;
let server = null;
let url = process.env.QL_URL;
if (!url) {
  // Own process group, so that stopping it also stops the vite process npx starts (a server left
  // running would answer the next run with an old build).
  server = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'ignore', detached: true });
  url = `http://localhost:${PORT}/`;
  for (let i = 0; i < 60; i++) {
    try { if ((await fetch(url)).ok) break; } catch { /* not up yet */ }
    await new Promise((r) => setTimeout(r, 500));
  }
}

const exe = existsSync('/opt/pw-browsers/chromium') ? { executablePath: '/opt/pw-browsers/chromium' } : {};
const browser = await chromium.launch(exe);
const failures = [];
const check = (name, got, want) => { if (JSON.stringify(got) !== JSON.stringify(want)) failures.push(`${name}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`); };

try {
  const p = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  const errs = [];
  p.on('pageerror', (e) => errs.push('pageerror: ' + e.message));
  p.on('console', (m) => { if (m.type() === 'error') errs.push('console: ' + m.text()); });
  await p.addInitScript(() => localStorage.setItem('quillgrid.prefs.v1', JSON.stringify({ firstRun: false })));
  await p.goto(url, { waitUntil: 'networkidle' });
  await p.waitForSelector('.ql-grid-scroller');
  await p.waitForTimeout(500);

  await p.click('.ql-grid-scroller', { position: { x: 120, y: 40 } });
  await p.evaluate(() => ql.grid.select(0, 0));
  for (const h of ['Item', 'Qty', 'Price']) { await p.keyboard.type(h); await p.keyboard.press('Tab'); }
  await p.keyboard.type('Total'); await p.keyboard.press('Enter');
  const rows = [['Apples', '3', '1.25'], ['Pears', '5', '0.8'], ['Plums', '12', '0.35']];
  for (const [i, r] of rows.entries()) {
    await p.evaluate((row) => ql.grid.select(row, 0), i + 1);
    for (const v of r) { await p.keyboard.type(v); await p.keyboard.press('Tab'); }
    await p.keyboard.type(`=B${i + 2}*C${i + 2}`); await p.keyboard.press('Enter');
  }
  await p.evaluate(() => ql.grid.select(4, 3));
  await p.keyboard.type('=SUM(D2:D4)'); await p.keyboard.press('Enter');
  await p.evaluate(() => ql.grid.select(6, 0));
  await p.keyboard.type('=SEQUENCE(3,2)'); await p.keyboard.press('Enter');
  await p.evaluate(() => ql.grid.select(10, 0));
  await p.keyboard.type('=LET(x,4,IF(x>3,x*2,0))'); await p.keyboard.press('Enter');

  const vals = await p.evaluate(() => [ql.sheet.value(1, 3), ql.sheet.value(4, 3), ql.sheet.value(8, 1), ql.sheet.value(10, 0)]);
  check('values', vals, [3.75, 11.95, 6, 8]);

  // Undo/redo round trip.
  await p.evaluate(() => ql.exec('edit.undo'));
  check('undo', await p.evaluate(() => ql.sheet.value(10, 0)), null);
  await p.evaluate(() => ql.exec('edit.redo'));
  check('redo', await p.evaluate(() => ql.sheet.value(10, 0)), 8);

  // Every ribbon tab renders without errors.
  const tabs = await p.locator('button.ql-tab[data-tab]:visible').count();
  for (let i = 0; i < tabs; i++) {
    await p.locator('button.ql-tab[data-tab]:visible').nth(i).click();
    await p.waitForTimeout(50);
  }
  if (errs.length) failures.push(...errs.slice(0, 10));
  if (process.argv[2]) await p.screenshot({ path: process.argv[2] });
} finally {
  await browser.close();
  if (server) { try { process.kill(-server.pid); } catch { server.kill(); } }
}

if (failures.length) { console.error('E2E smoke FAILED\n' + failures.join('\n')); process.exit(1); }
console.log('E2E smoke passed');
