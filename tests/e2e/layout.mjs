// Layout audit: opens every command that shows a window, menu, pane or Backstage page on emulated phones
// (portrait, small, landscape, and portrait with the on-screen keyboard up) with injected system-bar insets,
// and checks geometric invariants:
//   - every visible interactive element of an open overlay lies inside the safe viewport (not under the
//     status/navigation bars, the keyboard, or off-screen);
//   - sibling controls do not overlap each other, and button text does not spill out of its button;
//   - nothing is wider than the viewport (no horizontal page scroll, no sideways-scrolling dialog bodies);
//   - the app chrome (title bar, tabs, ribbon, formula bar, sheet tabs, bottom Quick Access Bar, status
//     bar) stays inside the safe viewport without overlapping;
//   - scroll cues: every scroll container keeps its own background, and "More" chips sit inside the
//     container they belong to;
// plus scenario checks for context menus at the edges, the cell editor's autocomplete near the bottom,
// the expanded formula bar, the command palette with the keyboard up, toasts versus dialog buttons, and
// the Tour.
//
// Usage: npm run build && npm run e2e:layout
//   QL_URL=http://localhost:4192/  test an already-running server (dev or preview) instead
//   QL_SHOTS=/some/dir             save a screenshot of each failing case
//   QL_ONLY=regex                  only sweep matching command ids
//   QL_PROFILES=phone,small        only run these profiles
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync } from 'node:fs';
import { chromium } from 'playwright';

const PORT = 4191;
let server = null;
let url = process.env.QL_URL;
if (!url) {
  // A server left on the port would be tested instead of the current build.
  if (await fetch(`http://localhost:${PORT}/`).then(() => true, () => false)) { console.error(`Port ${PORT} is already in use; stop that server or set QL_URL.`); process.exit(1); }
  // Own process group, so that stopping it also stops the vite process npx starts.
  server = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'ignore', detached: true });
  url = `http://localhost:${PORT}/`;
  for (let i = 0; i < 60; i++) {
    try { if ((await fetch(url)).ok) break; } catch { /* not up yet */ }
    await new Promise((r) => setTimeout(r, 500));
  }
}
const SHOTS = process.env.QL_SHOTS || '';
if (SHOTS) mkdirSync(SHOTS, { recursive: true });
const ONLY = process.env.QL_ONLY ? new RegExp(process.env.QL_ONLY) : null;

/** Emulated devices. Insets are what Capacitor reports as --safe-area-inset-*; `keyboard` shrinks the window
 *  the way Android's adjustResize does when the on-screen keyboard opens. */
const PROFILES = [
  { name: 'phone', width: 412, height: 915, insets: { top: 24, bottom: 48 } },
  { name: 'small', width: 360, height: 640, insets: { top: 24, bottom: 48 } },
  { name: 'tiny', width: 320, height: 568, insets: { top: 20, bottom: 34 } },
  { name: 'landscape', width: 915, height: 412, insets: { top: 24, bottom: 48, right: 48 } },
  { name: 'keyboard', width: 412, height: 915, keyboard: 336, insets: { top: 24, bottom: 48 } },
].filter((p) => !process.env.QL_PROFILES || process.env.QL_PROFILES.split(',').includes(p.name));

// Commands that leave the page, open native pickers, print, or replace the whole workbook.
const SKIP = /^(file\.(open|openEncrypted|openRepair|recoverText|print|quickPrint|close|exit|new|newFromTemplate|recover|openRecent|email|share|saveAs|saveAll|save|exportPdf|saveCopy|revert)|view\.(fullScreen|newWindow)|ai\.|dev\.(run|stop|record|runMacro)|edit\.(cut|copy|paste)|help\.(website|feedback|report))/;

/** Runs in the page: returns a list of problems for everything currently open. */
function auditInPage(opts) {
  const W = innerWidth, H = innerHeight;
  const probe = document.createElement('div');
  probe.style.cssText = 'position:fixed;visibility:hidden;pointer-events:none;padding:var(--sat) var(--sar) var(--sab) var(--sal)';
  document.body.appendChild(probe);
  const ps = getComputedStyle(probe);
  const ins = { t: parseFloat(ps.paddingTop) || 0, b: parseFloat(ps.paddingBottom) || 0, l: parseFloat(ps.paddingLeft) || 0, r: parseFloat(ps.paddingRight) || 0 };
  probe.remove();
  const vv = window.visualViewport;
  const safe = { l: ins.l, t: ins.t, r: W - ins.r, b: Math.min(H - ins.b, vv ? vv.offsetTop + vv.height : H) };
  const out = [];
  const TOL = 0.75;
  const label = (el) => {
    const cls = (el.className && typeof el.className === 'string' ? el.className.trim().split(/\s+/).filter((c) => !c.startsWith('sc-'))[0] : '') || '';
    const txt = (el.getAttribute('aria-label') || el.getAttribute('title') || el.textContent || el.value || '').trim().replace(/\s+/g, ' ').slice(0, 28);
    return `${el.tagName.toLowerCase()}${cls ? '.' + cls : ''}${txt ? ` "${txt}"` : ''}`;
  };
  const shown = (el) => {
    for (let n = el; n && n !== document.documentElement; n = n.parentElement) {
      const st = getComputedStyle(n);
      if (st.display === 'none' || st.visibility === 'hidden' || +st.opacity === 0) return false;
    }
    const r = el.getBoundingClientRect();
    return r.width > 0.5 && r.height > 0.5;
  };
  /** The part of el that is not clipped away by scrolling/overflow ancestors. */
  const clipped = (el) => {
    const r0 = el.getBoundingClientRect();
    let r = { l: r0.left, t: r0.top, r: r0.right, b: r0.bottom };
    for (let n = el; n.parentElement && n !== document.body; n = n.parentElement) {
      if (getComputedStyle(n).position === 'fixed') break;
      const a = n.parentElement;
      const st = getComputedStyle(a);
      if (st.overflowX === 'visible' && st.overflowY === 'visible') continue;
      const ar = a.getBoundingClientRect();
      const cl = ar.left + a.clientLeft, ct = ar.top + a.clientTop;
      r = { l: Math.max(r.l, cl), t: Math.max(r.t, ct), r: Math.min(r.r, cl + a.clientWidth), b: Math.min(r.b, ct + a.clientHeight) };
    }
    return r;
  };
  const inside = (r) => r.l >= safe.l - TOL && r.t >= safe.t - TOL && r.r <= safe.r + TOL && r.b <= safe.b + TOL;
  const fmt = (r) => `[${Math.round(r.l)},${Math.round(r.t)} → ${Math.round(r.r)},${Math.round(r.b)}]`;
  const SAFE = `safe ${fmt(safe)}`;
  const INTERACTIVE = 'button, input:not([type=hidden]), select, textarea, a[href], [role=option], [role=menuitem], [role=tab], [contenteditable=true]';

  // Open surfaces.
  const surfaces = [];
  for (const o of document.querySelectorAll('.ql-overlay')) if (shown(o)) for (const box of o.querySelectorAll(':scope > .ql-dialog, :scope > .ql-palette')) surfaces.push(['dialog', box]);
  for (const m of document.querySelectorAll('.ql-menu')) if (shown(m)) surfaces.push(['menu', m]);
  const bs = document.querySelector('.ql-backstage-open .ql-backstage');
  if (bs && shown(bs)) surfaces.push(['backstage', bs]);
  const pane = document.querySelector('.ql-pane-open .ql-pane');
  if (pane && shown(pane)) surfaces.push(['pane', pane]);
  for (const t of document.querySelectorAll('.ql-toast')) if (shown(t)) surfaces.push(['toast', t]);

  // The topmost dialog's controls must not be covered by a toast.
  const toasts = [...document.querySelectorAll('.ql-toast')].filter(shown).map((t) => t.getBoundingClientRect());
  const topOverlay = [...document.querySelectorAll('.ql-overlay')].filter(shown).pop();
  if (toasts.length) {
    const under = topOverlay ? topOverlay.querySelectorAll('.ql-dialog-foot button') : document.querySelectorAll('.ql-sheetbar button, .ql-qat-bottom button, .ql-status button');
    for (const b of under) {
      if (!shown(b)) continue;
      const r = b.getBoundingClientRect();
      for (const t of toasts) if (Math.min(r.right, t.right) - Math.max(r.left, t.left) > 2 && Math.min(r.bottom, t.bottom) - Math.max(r.top, t.top) > 2) out.push(`toast covers button ${label(b)}`);
    }
  }

  for (const [kind, s] of surfaces) {
    const name = `${kind} ${label(s).slice(0, 40)}`;
    // Fixed parts must be fully inside the safe area.
    const parts = kind === 'dialog' ? [...s.querySelectorAll(':scope > .ql-dialog-head, :scope > .ql-dialog-foot, :scope > .ql-palette-input')]
      : kind === 'backstage' ? [...s.querySelectorAll(':scope > .ql-bs-nav, :scope > .ql-bs-body')]
      : kind === 'pane' ? [...s.querySelectorAll(':scope > .ql-pane-head, :scope > .ql-pane-body')] : [s];
    for (const p of parts) {
      if (!shown(p)) continue;
      const r0 = p.getBoundingClientRect();
      const r = { l: r0.left, t: r0.top, r: r0.right, b: r0.bottom };
      if (!inside(r)) out.push(`${name}: ${label(p)} ${fmt(r)} outside ${SAFE}`);
    }
    const els = [...s.querySelectorAll(INTERACTIVE)].filter(shown);
    if (s.matches(INTERACTIVE) && shown(s)) els.push(s);
    for (const el of els) {
      const r = clipped(el);
      if (r.r - r.l < 1 || r.b - r.t < 1) continue; // scrolled out of view
      if (!inside(r)) out.push(`${name}: ${label(el)} ${fmt(r)} outside ${SAFE}`);
      // Text spilling out of a button that does not clip it.
      if (el.tagName === 'BUTTON' && getComputedStyle(el).overflowX === 'visible' && el.scrollWidth > el.clientWidth + 2) out.push(`${name}: ${label(el)} content ${el.scrollWidth}px wider than the button (${el.clientWidth}px)`);
    }
    // Sibling controls must not overlap.
    const byParent = new Map();
    for (const el of els) { const k = el.parentElement; if (!byParent.has(k)) byParent.set(k, []); byParent.get(k).push(el); }
    for (const sibs of byParent.values()) {
      for (let i = 0; i < sibs.length; i++) for (let j = i + 1; j < sibs.length; j++) {
        const a = sibs[i].getBoundingClientRect(), b = sibs[j].getBoundingClientRect();
        const ow = Math.min(a.right, b.right) - Math.max(a.left, b.left), oh = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
        if (ow > 1.5 && oh > 1.5) out.push(`${name}: ${label(sibs[i])} overlaps ${label(sibs[j])} (${Math.round(ow)}x${Math.round(oh)})`);
      }
    }
    // Content wider than a dialog body or pane makes it scroll sideways.
    for (const b of s.querySelectorAll('.ql-dialog-body, .ql-pane-body, .ql-bs-body')) {
      if (shown(b) && b.scrollWidth > b.clientWidth + 1) out.push(`${name}: ${label(b)} scrolls sideways (${b.scrollWidth}px content in ${b.clientWidth}px)`);
    }
  }

  // No horizontal page overflow.
  for (const el of [document.documentElement, document.body]) if (el.scrollWidth > W + 1) out.push(`page is ${el.scrollWidth}px wide in a ${W}px viewport`);

  // Scroll cues: chips stay within the safe area and on their container; containers keep their own background.
  for (const chip of document.querySelectorAll('.sc-chip:not(.sc-hidden)')) {
    const r0 = chip.getBoundingClientRect();
    const r = { l: r0.left, t: r0.top, r: r0.right, b: r0.bottom };
    if (!inside(r)) out.push(`scroll cue chip ${fmt(r)} outside ${SAFE}`);
  }
  if (opts.checkBackgrounds) {
    const opaque = (c) => c && c !== 'transparent' && !/rgba\(.*,\s*0\)$/.test(c);
    const bgOf = (el) => { for (let n = el; n; n = n.parentElement) { const c = getComputedStyle(n).backgroundColor; if (opaque(c)) return c; } return 'rgb(255, 255, 255)'; };
    for (const el of document.querySelectorAll('.sc-scroll')) {
      if (!shown(el)) continue;
      const withCue = getComputedStyle(el).backgroundColor;
      el.classList.remove('sc-scroll');
      const own = bgOf(el);
      el.classList.add('sc-scroll');
      if (withCue !== own) out.push(`scroll cue changes the background of ${label(el)}: ${own} → ${withCue}`);
      // The edge shadows must not be overridden by a more specific background rule (an inline style is the
      // author's explicit choice and is left alone).
      if (!el.style.background && !/gradient/.test(getComputedStyle(el).backgroundImage)) out.push(`scroll cue shadows overridden on ${label(el)}`);
    }
  }

  // App chrome: inside the safe area and not overlapping each other.
  if (opts.chrome) {
    const chrome = ['.ql-titlebar', '.ql-tabs', '.ql-ribbon', '.ql-formulabar', '.ql-sheetbar', '.ql-qat-bottom', '.ql-status'].map((q) => document.querySelector(q)).filter((e) => e && shown(e));
    for (const c of chrome) {
      const r0 = c.getBoundingClientRect();
      const r = { l: r0.left, t: r0.top, r: r0.right, b: r0.bottom };
      if (!inside(r)) out.push(`chrome ${label(c)} ${fmt(r)} outside ${SAFE}`);
    }
    for (let i = 0; i < chrome.length; i++) for (let j = i + 1; j < chrome.length; j++) {
      const a = chrome[i].getBoundingClientRect(), b = chrome[j].getBoundingClientRect();
      if (Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top) > 1 && Math.min(a.right, b.right) - Math.max(a.left, b.left) > 1) out.push(`chrome ${label(chrome[i])} overlaps ${label(chrome[j])}`);
    }
    const grid = document.querySelector('.ql-grid');
    if (grid && grid.getBoundingClientRect().height < 60) out.push(`grid area only ${Math.round(grid.getBoundingClientRect().height)}px tall`);
    // Quick Access Bar buttons keep their size (the bar scrolls) rather than squeezing their labels.
    for (const l of document.querySelectorAll('.ql-qat .ql-cmd-label')) if (shown(l) && l.scrollWidth > l.clientWidth + 1) out.push(`Quick Access Bar label "${l.textContent}" is cut off (${l.clientWidth}px of ${l.scrollWidth}px)`);
    for (const t of document.querySelectorAll('.ql-qat .ql-qat-customize')) if (shown(t)) { const r = clipped(t); if (r.r - r.l < t.getBoundingClientRect().width - 1) out.push('Quick Access Bar customize button is scrolled out of view'); }
  }
  return { surfaces: surfaces.filter(([k]) => k !== 'toast').length, problems: [...new Set(out)] };
}

async function setup(browser, prof) {
  const ctx = await browser.newContext({ viewport: { width: prof.width, height: prof.height }, hasTouch: true, isMobile: true, deviceScaleFactor: 1 });
  const p = await ctx.newPage();
  p.errors = [];
  p.on('pageerror', (e) => p.errors.push(e.message));
  await p.addInitScript((ins) => {
    localStorage.setItem('quilllattice.prefs.v1', JSON.stringify({ firstRun: false }));
    window.print = () => {}; window.open = () => null;
    HTMLInputElement.prototype.showPicker = function () {};
    const apply = () => { const d = document.documentElement.style; for (const [k, v] of Object.entries(ins)) d.setProperty(`--safe-area-inset-${k}`, `${v}px`); };
    // The parser may replace the root element this script first sees: set the insets again once it is there.
    if (document.documentElement) apply();
    addEventListener('DOMContentLoaded', apply);
    // A file input clicked by a command would open a native picker; swallow it.
    const click = HTMLInputElement.prototype.click;
    HTMLInputElement.prototype.click = function () { if (this.type !== 'file') click.call(this); };
  }, prof.insets);
  await p.goto(url, { waitUntil: 'networkidle' });
  await p.waitForSelector('.ql-grid-scroller');
  await p.waitForTimeout(600);
  // Some data, so data commands have something to work on, and a selection.
  await p.evaluate(() => {
    const s = ql.sheet;
    const rows = [['Item', 'Qty', 'Price', 'Date'], ['Apples', 3, 1.25, 45000], ['Pears', 5, 0.8, 45001], ['Plums', 12, 0.35, 45002], ['Figs', 7, 2.1, 45003]];
    rows.forEach((r, i) => r.forEach((v, j) => s.set(i, j, { v })));
    ql.grid.select(0, 0); ql.refresh();
  });
  const got = await p.evaluate(() => getComputedStyle(document.querySelector('.ql-app')).paddingTop);
  if (got !== `${prof.insets.top || 0}px`) throw new Error(`[${prof.name}] inset injection failed: app padding-top is ${got}`);
  if (prof.keyboard) await p.setViewportSize({ width: prof.width, height: prof.height - prof.keyboard });
  p.prefs0 = await p.evaluate(() => JSON.stringify(ql.prefs));
  return p;
}

async function reset(p) {
  for (let i = 0; i < 3; i++) await p.keyboard.press('Escape').catch(() => {});
  await p.evaluate((prefs0) => {
    document.querySelectorAll('.ql-overlay, .ql-menu, .ql-toast').forEach((e) => e.remove());
    try { ql.ui.hideBackstage(); } catch { /* */ }
    try { ql.ui.closePane(); } catch { /* */ }
    try { if (ql.grid.editing) ql.grid.cancelEdit(); } catch { /* */ }
    document.body.classList.remove('ql-zen', 'ql-fullscreen');
    if (JSON.stringify(ql.prefs) !== prefs0) { Object.assign(ql.prefs, JSON.parse(prefs0)); ql.ui.applyPrefs(); ql.ui.rebuild(); }
    if ((ql.sheet.zoom || 1) !== 1) ql.grid.setZoom(1);
    ql.grid.select(0, 0);
  }, p.prefs0).catch(() => {});
}

const failures = [];
async function check(p, prof, what, opts = {}) {
  const res = await p.evaluate(auditInPage, { chrome: !!opts.chrome, checkBackgrounds: !!opts.checkBackgrounds }).catch((e) => ({ surfaces: 0, problems: [`audit failed: ${e.message}`] }));
  if (res.problems.length) {
    for (const pr of res.problems) failures.push(`[${prof.name}] ${what}: ${pr}`);
    if (SHOTS) await p.screenshot({ path: `${SHOTS}/${prof.name}-${what.replace(/[^\w.-]+/g, '_').slice(0, 60)}.png` }).catch(() => {});
  }
  return res;
}

async function sweep(browser, prof) {
  const p = await setup(browser, prof);
  const stats = { prof: prof.name, commands: 0, withUi: 0, menus: 0 };
  await check(p, prof, 'start', { chrome: true });
  const ids = (await p.evaluate(() => window.__qlAllIds)).filter((id) => !SKIP.test(id) && (!ONLY || ONLY.test(id)));
  for (const id of ids) {
    stats.commands++;
    // A command that reloads the page (none should) shows up as a page error rather than stopping the sweep.
    const before = await p.evaluate(() => performance.timeOrigin).catch(() => 0);
    await p.evaluate((id) => { try { const r = ql.exec(id); if (r && r.catch) r.catch(() => {}); } catch { /* */ } }, id).catch(() => {});
    await p.waitForTimeout(140);
    const res = await check(p, prof, id, { checkBackgrounds: true, chrome: true });
    if (res.surfaces) stats.withUi++;
    if ((await p.evaluate(() => performance.timeOrigin).catch(() => 0)) !== before) {
      p.errors.push(`${id} reloaded the page`);
      await p.waitForSelector('.ql-grid-scroller');
      await p.waitForTimeout(400);
    }
    await reset(p);
    // Drop-down commands: open their menu from the ribbon (or wherever the command is shown).
    const hasMenu = await p.evaluate((id) => { const c = window.__qlGetCommand?.(id); if (!c?.menu) return false; try { ql.ui.openCommandMenu(c, ql.ui.anchorFor(id)); } catch { return false; } return true; }, id).catch(() => false);
    if (hasMenu) {
      stats.menus++;
      await p.waitForTimeout(60);
      await check(p, prof, `${id} (menu)`, { checkBackgrounds: true });
      await reset(p);
    }
  }
  await p.context().close();
  if (p.errors.length) failures.push(...[...new Set(p.errors)].slice(0, 5).map((e) => `[${prof.name}] page error: ${e}`));
  // Scenarios start from a fresh page, clear of whatever the commands above changed.
  const q = await setup(browser, prof);
  await scenarios(q, prof);
  await q.context().close();
  if (q.errors.length) failures.push(...[...new Set(q.errors)].slice(0, 5).map((e) => `[${prof.name}] page error: ${e}`));
  return stats;
}

async function scenarios(p, prof) {
  const W = prof.width, H = prof.height - (prof.keyboard || 0);
  // Quick Access Bar drop-down and the More menu.
  await p.locator('.ql-qat-customize:visible').first().click();
  await p.waitForTimeout(150);
  await check(p, prof, 'qat menu', { checkBackgrounds: true });
  // The menu is taller than the screen: a "More" cue must be showing.
  const cue = await p.evaluate(() => { const m = document.querySelector('.ql-menu'); return !!m && (m.scrollHeight <= m.clientHeight + 6 || !!document.querySelector('.sc-chip:not(.sc-hidden)')); });
  if (!cue) failures.push(`[${prof.name}] qat menu: more choices below the fold but no scroll cue`);
  await reset(p);
  await p.locator('.ql-titlebar [aria-label=More]').click();
  await p.waitForTimeout(100);
  await check(p, prof, 'more menu');
  await reset(p);

  // Context menus near every edge.
  for (const [x, y] of [[W - 2, H - 2], [2, H - 2], [W - 2, 2], [W / 2, H / 2]]) {
    await p.evaluate(([x, y]) => ql.ui.gridContextMenu(x, y, { area: 'cell' }), [x, y]);
    await p.waitForTimeout(60);
    await check(p, prof, `context menu at ${Math.round(x)},${Math.round(y)}`);
    await reset(p);
  }
  await p.evaluate(([x, y]) => ql.ui.sheetMenu(x, y), [W - 4, H - 4]);
  await check(p, prof, 'sheet tab menu');
  await reset(p);

  // Customize Quick Access Bar: the list boxes show the scroll cue and "Add All A–Z" can be found.
  await p.evaluate(() => { ql.exec('file.customizeQat'); });
  await p.waitForTimeout(250);
  await check(p, prof, 'customize qat', { checkBackgrounds: true });
  const qat = await p.evaluate(() => {
    const lists = [...document.querySelectorAll('.ql-dialog .ql-cmdlist')];
    const btn = [...document.querySelectorAll('.ql-cust-mid button')].find((b) => /Add All/.test(b.getAttribute('aria-label') || ''));
    return { cued: lists.every((l) => l.classList.contains('sc-scroll')), label: btn ? btn.innerText.trim() : null };
  });
  if (!qat.cued) failures.push(`[${prof.name}] customize qat: list boxes have no scroll cue`);
  if (!qat.label || !/Add All/.test(qat.label)) failures.push(`[${prof.name}] customize qat: "Add All A–Z" has no visible label (got ${JSON.stringify(qat.label)})`);
  await reset(p);

  // Command palette (types, so on a phone the keyboard is up).
  await p.evaluate(() => ql.ui.commandPalette());
  await p.waitForTimeout(100);
  await p.keyboard.type('format');
  await p.waitForTimeout(100);
  await check(p, prof, 'command palette', { checkBackgrounds: true });
  await reset(p);

  // Backstage pages.
  for (const page of ['home', 'new', 'open', 'info', 'saveas', 'export', 'print', 'share']) {
    await p.evaluate((pg) => ql.ui.showBackstage(pg), page);
    await p.waitForTimeout(80);
    await check(p, prof, `backstage ${page}`, { checkBackgrounds: true });
  }
  await reset(p);

  // Cell editor autocomplete near the bottom edge of the grid.
  const row = await p.evaluate(() => {
    const g = ql.grid; const sc = document.querySelector('.ql-grid-scroller').getBoundingClientRect();
    for (let r = 0; r < 200; r++) { const rc = g.cellRect(r, 1); if (rc.y + rc.h > g.height - 30) return Math.max(0, r - 1); }
    return 20;
  });
  await p.evaluate((r) => { ql.focus(); ql.grid.select(r, 1); }, row);
  await p.keyboard.type('=SU');
  await p.waitForTimeout(150);
  const ac = await p.evaluate(() => {
    const pop = document.querySelector('.ql-ac-popup.ql-show');
    if (!pop) return `no autocomplete list (editing: ${ql.grid.editing}, focus: ${document.activeElement?.className})`;
    const r = pop.getBoundingClientRect(), g = document.querySelector('.ql-grid').getBoundingClientRect();
    const vv = visualViewport;
    const bottom = Math.min(g.bottom, vv ? vv.offsetTop + vv.height : innerHeight);
    if (r.bottom > bottom + 1 || r.top < g.top - 1 || r.right > g.right + 1) return `autocomplete [${Math.round(r.left)},${Math.round(r.top)} → ${Math.round(r.right)},${Math.round(r.bottom)}] outside the grid [${Math.round(g.left)},${Math.round(g.top)} → ${Math.round(g.right)},${Math.round(bottom)}]`;
    const ed = document.querySelector('.ql-cell-editor').getBoundingClientRect();
    if (Math.min(r.bottom, ed.bottom) - Math.max(r.top, ed.top) > 1 && Math.min(r.right, ed.right) - Math.max(r.left, ed.left) > 1) return 'autocomplete covers the cell being edited';
    return '';
  });
  if (ac) failures.push(`[${prof.name}] cell editor near the bottom: ${ac}`);
  await reset(p);

  // Expanded formula bar keeps the grid usable.
  await p.evaluate(() => document.querySelector('.ql-formulabar').classList.add('ql-fb-expanded'));
  await p.waitForTimeout(50);
  await check(p, prof, 'formula bar expanded', { chrome: true });
  await p.evaluate(() => document.querySelector('.ql-formulabar').classList.remove('ql-fb-expanded'));

  // Quick Access Bar at the bottom, with labels, next to the sheet tabs and status bar.
  await p.evaluate(() => { ql.prefs.qatPosition = 'bottom'; ql.prefs.qatLabels = true; ql.ui.applyPrefs(); ql.ui.rebuild(); });
  await p.waitForTimeout(80);
  await check(p, prof, 'quick access bar at the bottom', { chrome: true });
  // A toast must cover neither the bottom chrome's buttons nor a dialog's buttons.
  await p.evaluate(() => ql.exec('__layout_test_toast__'));
  await p.waitForTimeout(80);
  await check(p, prof, 'toast over the bottom bars');
  await reset(p);
  await p.evaluate(() => { ql.exec('file.customizeQat'); });
  await p.waitForTimeout(200);
  await p.evaluate(() => ql.exec('__layout_test_toast__'));
  await p.waitForTimeout(80);
  await check(p, prof, 'toast over a dialog');
  await reset(p);

  // Dark theme: containers keep their backgrounds and the cue's shadows.
  await p.evaluate(() => { ql.prefs.theme = 'dark'; ql.ui.applyPrefs(); });
  await p.evaluate(() => ql.ui.showBackstage('new'));
  await p.waitForTimeout(80);
  await check(p, prof, 'dark backstage', { checkBackgrounds: true });
  await reset(p);
  await p.evaluate(() => { ql.prefs.theme = 'dark'; ql.ui.applyPrefs(); ql.exec('file.customizeQat'); });
  await p.waitForTimeout(200);
  await check(p, prof, 'dark customize qat', { checkBackgrounds: true });
  await reset(p);
  await p.evaluate(() => { ql.prefs.theme = 'dark'; ql.ui.applyPrefs(); });
  await p.locator('.ql-qat-customize:visible').first().click();
  await p.waitForTimeout(120);
  await check(p, prof, 'dark qat menu', { checkBackgrounds: true });
  await reset(p);

  // Tour.
  await p.evaluate(() => { void ql.ui.tour(); });
  await p.waitForTimeout(150);
  for (let i = 0; i < 5; i++) {
    await check(p, prof, `tour step ${i + 1}`);
    await p.evaluate(() => document.querySelector('.ql-dialog .ql-primary')?.click());
    await p.waitForTimeout(160);
  }
  await reset(p);

  // Chart handles inside the grid and big enough to touch.
  await p.evaluate(() => { ql.grid.select(0, 0); });
  await p.evaluate(async () => { try { await ql.exec('insert.chartCombo'); } catch { /* */ } });
  await p.waitForTimeout(250);
  const handles = await p.evaluate(() => {
    const hs = [...document.querySelectorAll('.ql-object.ql-selected .ql-handle')];
    if (!hs.length) return '';
    const small = hs.filter((h) => h.getBoundingClientRect().width < 17);
    return small.length ? `${small.length} chart handles smaller than a touch target (${small.map((h) => Math.round(h.getBoundingClientRect().width)).join(', ')}px)` : '';
  });
  if (handles) failures.push(`[${prof.name}] chart: ${handles}`);
  await reset(p);
}

const exe = existsSync('/opt/pw-browsers/chromium') ? { executablePath: '/opt/pw-browsers/chromium' } : {};
const browser = await chromium.launch(exe);
let stats = [];
try {
  stats = await Promise.all(PROFILES.map((prof) => sweep(browser, prof)));
} finally {
  await browser.close();
  if (server) { try { process.kill(-server.pid); } catch { server.kill(); } }
}
for (const s of stats) console.log(`${s.prof}: ${s.commands} commands, ${s.withUi} showed a window/menu/pane, ${s.menus} drop-down menus`);
if (failures.length) {
  console.error(`Layout audit FAILED (${failures.length} problems)\n` + failures.join('\n'));
  process.exit(1);
}
console.log('Layout audit passed');
