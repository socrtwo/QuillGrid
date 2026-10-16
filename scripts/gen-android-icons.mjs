#!/usr/bin/env node
// Generates the QuillLattice launcher icon in every format the Android app needs
// from a single set of vector shapes (defined below in a 108x108 adaptive-icon
// coordinate space; the 66dp safe zone is the circle centered at 54,54 r=33).
//
// Outputs:
//   resources/icon.svg                                   rounded-square app icon (desktop, PWA, README)
//   resources/icon-square.svg                            opaque full-bleed square (iOS, PWA maskable)
//   resources/icon-foreground.svg                        adaptive foreground layer only
//   public/favicon.svg                                   browser tab icon (tighter crop)
//   resources/icon-512.png                               store / README icon
//   android/app/src/main/res/drawable/ic_launcher_foreground.xml   (vector)
//   android/app/src/main/res/drawable/ic_launcher_background.xml   (vector)
//   android/app/src/main/res/drawable/ic_launcher_monochrome.xml   (vector, themed icons)
//   android/app/src/main/res/values/ic_launcher_background.xml     (background / splash color)
//   android/app/src/main/res/mipmap-*/ic_launcher.png, ic_launcher_round.png  (legacy, API 24-25)
//
// The artwork: a cream worksheet with a header row and column, a highlighted active cell with its
// fill handle, and a mint quill whose gold nib writes into that cell, on a deep green field.
//
// Usage: node scripts/gen-android-icons.mjs [--no-png]
// PNG rendering uses Playwright's Chromium. Set CHROMIUM_PATH to override the browser binary.

import { mkdirSync, writeFileSync, existsSync, rmSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const res = join(root, 'android/app/src/main/res');

const GREEN = '#1E5F3A'; // brand
const GREEN_MID = '#2E8B57';
const MINT = '#8FD9B6';
const CREAM = '#FBF6E9';
const HEADER = '#D5EADC';
const HEADER_ACTIVE = '#9CCFB0';
const GRID = '#7FAF92';
const INK = '#5E8C72';
const GOLD = '#F0C25A';

// --- Artwork (108x108 viewport) -------------------------------------------
const n = (v) => +v.toFixed(2);
const rect = (x, y, w, h) => `M${n(x)},${n(y)} L${n(x + w)},${n(y)} L${n(x + w)},${n(y + h)} L${n(x)},${n(y + h)} Z`;
const rrect = (x, y, w, h, r) =>
  `M${n(x + r)},${n(y)} L${n(x + w - r)},${n(y)} Q${n(x + w)},${n(y)} ${n(x + w)},${n(y + r)} L${n(x + w)},${n(y + h - r)} ` +
  `Q${n(x + w)},${n(y + h)} ${n(x + w - r)},${n(y + h)} L${n(x + r)},${n(y + h)} Q${n(x)},${n(y + h)} ${n(x)},${n(y + h - r)} ` +
  `L${n(x)},${n(y + r)} Q${n(x)},${n(y)} ${n(x + r)},${n(y)} Z`;

// Worksheet: x 30..68, y 38..76; header row 5.5 high, header column 5.5 wide; 3 data columns x 4 rows.
const SX = 30, SY = 38, SW = 38, SH = 38, HD = 5.5;
const colX = [SX + HD, SX + HD + (SW - HD) / 3, SX + HD + (2 * (SW - HD)) / 3, SX + SW];
const rowY = [SY + HD, SY + HD + (SH - HD) / 4, SY + HD + (2 * (SH - HD)) / 4, SY + HD + (3 * (SH - HD)) / 4, SY + SH];
const AC = 1, AR = 1; // active cell (column B, row 2)
const sheet = rrect(SX, SY, SW, SH, 2.6);
const sheetShadow = rrect(SX + 1.1, SY + 1.6, SW, SH, 2.6);
const headerRow = `M${SX + 2.6},${SY} L${SX + SW - 2.6},${SY} Q${SX + SW},${SY} ${SX + SW},${SY + 2.6} L${SX + SW},${SY + HD} L${SX},${SY + HD} L${SX},${SY + 2.6} Q${SX},${SY} ${SX + 2.6},${SY} Z`;
const headerCol = `M${SX},${SY + HD} L${SX + HD},${SY + HD} L${SX + HD},${SY + SH} L${SX + 2.6},${SY + SH} Q${SX},${SY + SH} ${SX},${SY + SH - 2.6} Z`;
const headerActive = `${rect(colX[AC], SY, colX[AC + 1] - colX[AC], HD)} ${rect(SX, rowY[AR], HD, rowY[AR + 1] - rowY[AR])}`;
const LW = 0.7;
const gridLines = [
  ...colX.slice(0, 3).map((x) => rect(x - LW / 2, SY, LW, SH)),
  ...rowY.slice(0, 4).map((y) => rect(SX, y - LW / 2, SW, LW)),
].join(' ');
// Right-aligned "numbers" in a few cells.
const bar = (c, r, w) => rect(colX[c + 1] - 1.8 - w, (rowY[r] + rowY[r + 1]) / 2 - 0.8, w, 1.6);
const values = [bar(0, 0, 5), bar(2, 0, 4), bar(0, 1, 3.5), bar(0, 2, 6), bar(1, 2, 4.5), bar(2, 2, 3), bar(0, 3, 4), bar(1, 3, 6), bar(2, 3, 5)].join(' ');
// Active cell: thick border (outer minus inner, evenOdd) and fill handle.
const ax = colX[AC], ay = rowY[AR], aw = colX[AC + 1] - ax, ah = rowY[AR + 1] - ay, B = 1.5;
const activeBorder = `${rect(ax - B / 2, ay - B / 2, aw + B, ah + B)} ${rect(ax + B / 2, ay + B / 2, aw - B, ah - B)}`;
const fillHandle = rect(ax + aw - 1.6, ay + ah - 1.6, 3.2, 3.2);

// Quill: drawn in a local frame (u along the shaft from the nib tip, v across it), then placed with
// its tip in the active cell, rising up and to the right at 48 degrees.
const TIP = [ax + aw * 0.45, ay + ah * 0.62];
const ANG = (48 * Math.PI) / 180;
const A = [Math.cos(ANG), -Math.sin(ANG)];
const N = [Math.sin(ANG), Math.cos(ANG)];
const P = (u, v) => `${n(TIP[0] + u * A[0] + v * N[0])},${n(TIP[1] + u * A[1] + v * N[1])}`;
const local = (cmds) =>
  cmds
    .map(([op, ...pts]) => (op === 'Z' ? 'Z' : `${op}${pts.map(([u, v]) => P(u, v)).join(' ')}`))
    .join(' ');
const vane = local([
  ['M', [8.8, 0]],
  ['C', [13, -7.5], [27, -10.5], [40, -1.8]],
  ['C', [35, 1.6], [31.5, 3.2], [28, 4.3]],
  ['L', [25.6, 2.6]],
  ['L', [25.2, 5.1]],
  ['C', [19.5, 6.2], [13.5, 4.6], [8.8, 0]],
  ['Z'],
]);
// Bare end of the shaft past the vane.
const shaft = local([['M', [7.2, -0.5]], ['L', [42.5, -2.2]], ['L', [42.5, -1.4]], ['L', [7.2, 0.5]], ['Z']]);
// Barb lines on the vane (thin slivers angled back toward the nib).
const barb = (u0, side, len) => {
  const s = side;
  return [['M', [u0, -0.3 * s]], ['L', [u0 + len * 0.55, -len * s]], ['L', [u0 + len * 0.55 + 0.7, -len * s + 0.25 * s]], ['L', [u0 + 0.9, -0.3 * s]], ['Z']];
};
const barbs = local([...barb(14, 1, 4.6), ...barb(20, 1, 6), ...barb(27, 1, 5.6), ...barb(33, 1, 3.6), ...barb(15.5, -1, 3.2), ...barb(30, -1, 2.2)]);
// Nib with a slit and a breather hole (evenOdd cut-outs).
const nibOuter = local([['M', [0, 0]], ['C', [2.5, -1.1], [5.5, -2.2], [8.8, -2]], ['L', [8.8, 2]], ['C', [5.5, 2.2], [2.5, 1.1], [0, 0]], ['Z']]);
const nibSlit = local([['M', [1.4, -0.22]], ['L', [5.4, -0.22]], ['L', [5.4, 0.22]], ['L', [1.4, 0.22]], ['Z']]);
const nibHole = local([['M', [6.2, -0.75]], ['L', [7.7, -0.75]], ['L', [7.7, 0.75]], ['L', [6.2, 0.75]], ['Z']]);
const nib = `${nibOuter} ${nibSlit} ${nibHole}`;

const fg = [
  { d: sheetShadow, fill: '#000000', alpha: 0.2 },
  { d: sheet, fill: CREAM },
  { d: headerRow, fill: HEADER },
  { d: headerCol, fill: HEADER },
  { d: headerActive, fill: HEADER_ACTIVE },
  { d: gridLines, fill: GRID },
  { d: values, fill: INK, alpha: 0.55 },
  { d: activeBorder, fill: GREEN, evenOdd: true },
  { d: fillHandle, fill: GREEN },
  { d: vane, fill: MINT },
  { d: barbs, fill: GREEN_MID, alpha: 0.7 },
  { d: shaft, fill: CREAM },
  { d: nib, fill: GOLD, evenOdd: true },
];
// Monochrome (Android 13+ themed icon): silhouettes only. The sheet is an outline with its grid,
// the active cell is a frame, and the quill is solid with its shaft cut out of the vane.
const sheetFrame = `${sheet} ${rrect(SX + 1.6, SY + 1.6, SW - 3.2, SH - 3.2, 1.4)}`;
const mono = [
  { d: sheetFrame, fill: '#FFFFFF', evenOdd: true },
  { d: gridLines, fill: '#FFFFFF' },
  { d: activeBorder, fill: '#FFFFFF', evenOdd: true },
  { d: fillHandle, fill: '#FFFFFF' },
  // even-odd: where the shaft crosses the vane it is cut out, so it still reads as a line
  { d: `${vane} ${shaft}`, fill: '#FFFFFF', evenOdd: true },
  { d: nib, fill: '#FFFFFF', evenOdd: true },
];
const bg = [
  { d: 'M0,0 L108,0 L108,108 L0,108 Z', fill: GREEN },
  // faint lattice in the field, then a soft diagonal glow
  { d: Array.from({ length: 11 }, (_, i) => `${rect(i * 10 + 3.6, 0, 0.5, 108)} ${rect(0, i * 10 + 3.6, 108, 0.5)}`).join(' '), fill: '#FFFFFF', alpha: 0.05 },
  { d: 'M0,0 L108,0 L0,108 Z', fill: GREEN_MID, alpha: 0.45 },
];

// --- Writers -----------------------------------------------------------------
const svgPaths = (layers) =>
  layers
    .map(
      (l) =>
        `  <path d="${l.d}" fill="${l.fill}"${l.alpha != null ? ` fill-opacity="${l.alpha}"` : ''}${l.evenOdd ? ' fill-rule="evenodd"' : ''}/>`,
    )
    .join('\n');

function svg(layers, { size = 108, viewBox = '0 0 108 108', clip = null } = {}) {
  const [vx, vy, vw] = viewBox.split(' ').map(Number);
  const clipDef =
    clip === 'circle'
      ? `<clipPath id="c"><circle cx="${vx + vw / 2}" cy="${vy + vw / 2}" r="${vw / 2}"/></clipPath>`
      : clip === 'rounded'
        ? `<clipPath id="c"><rect x="${vx}" y="${vy}" width="${vw}" height="${vw}" rx="${n(vw * 0.18)}"/></clipPath>`
        : '';
  const open = clip ? `<defs>${clipDef}</defs>\n<g clip-path="url(#c)">` : '<g>';
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="${viewBox}">\n${open}\n${svgPaths(layers)}\n</g>\n</svg>\n`;
}

const hex = (c, alpha) => {
  const h = c.replace('#', '').toUpperCase();
  if (alpha == null) return `#${h}`;
  const a = Math.round(alpha * 255).toString(16).padStart(2, '0').toUpperCase();
  return `#${a}${h}`;
};
function vector(layers, comment) {
  const paths = layers
    .map(
      (l) =>
        `    <path\n        android:fillColor="${hex(l.fill, l.alpha)}"${l.evenOdd ? '\n        android:fillType="evenOdd"' : ''}\n        android:pathData="${l.d}" />`,
    )
    .join('\n');
  return `<?xml version="1.0" encoding="utf-8"?>\n<!-- ${comment} Generated by scripts/gen-android-icons.mjs; edit the script, not this file. -->\n<vector xmlns:android="http://schemas.android.com/apk/res/android"\n    android:width="108dp"\n    android:height="108dp"\n    android:viewportWidth="108"\n    android:viewportHeight="108">\n${paths}\n</vector>\n`;
}

function write(path, content) {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, content);
  console.log('wrote', path.replace(root + '/', ''));
}

write(join(root, 'resources/icon.svg'), svg([...bg, ...fg], { size: 1024, viewBox: '14 14 80 80', clip: 'rounded' }));
write(join(root, 'resources/icon-square.svg'), svg([...bg, ...fg], { size: 1024, viewBox: '12 12 84 84' }));
write(join(root, 'resources/icon-foreground.svg'), svg(fg, { size: 1024 }));
write(join(root, 'public/favicon.svg'), svg([...bg, ...fg], { size: 64, viewBox: '22 22 64 64', clip: 'rounded' }));
write(join(res, 'drawable/ic_launcher_foreground.xml'), vector(fg, 'QuillLattice adaptive icon foreground.'));
write(join(res, 'drawable/ic_launcher_background.xml'), vector(bg, 'QuillLattice adaptive icon background.'));
write(join(res, 'drawable/ic_launcher_monochrome.xml'), vector(mono, 'QuillLattice themed (monochrome) icon.'));
write(
  join(res, 'values/ic_launcher_background.xml'),
  `<?xml version="1.0" encoding="utf-8"?>\n<resources>\n    <color name="ic_launcher_background">${GREEN}</color>\n    <color name="splash_background">${GREEN}</color>\n</resources>\n`,
);

// The Capacitor template ships a drawable-v24 foreground that would shadow ours,
// plus PNG foregrounds referenced from mipmap; remove them.
for (const p of ['drawable-v24/ic_launcher_foreground.xml']) {
  const f = join(res, p);
  if (existsSync(f)) rmSync(f);
}

if (process.argv.includes('--no-png')) process.exit(0);

// --- PNG rendering -------------------------------------------------------------
const { chromium } = await import('playwright');
const launchOpts = { headless: true };
if (process.env.CHROMIUM_PATH) launchOpts.executablePath = process.env.CHROMIUM_PATH;
let browser;
try {
  browser = await chromium.launch(launchOpts);
} catch (err) {
  // Fall back to any locally installed Playwright Chromium build.
  const { readdirSync } = await import('node:fs');
  const base = process.env.PLAYWRIGHT_BROWSERS_PATH || '/opt/pw-browsers';
  const cand = existsSync(base)
    ? readdirSync(base)
        .filter((d) => d.startsWith('chromium-'))
        .flatMap((d) => [join(base, d, 'chrome-linux64/chrome'), join(base, d, 'chrome-linux/chrome')])
        .filter(existsSync)
    : [];
  if (!cand.length) throw err;
  browser = await chromium.launch({ ...launchOpts, executablePath: cand[0] });
}
const pageCtx = await browser.newPage();

async function renderPng(svgText, px, out) {
  await pageCtx.setViewportSize({ width: px, height: px });
  const sized = svgText.replace(/width="\d+" height="\d+"/, `width="${px}" height="${px}"`);
  await pageCtx.setContent(
    `<!doctype html><html><head><style>html,body{margin:0;padding:0;background:transparent}svg{display:block}</style></head><body>${sized}</body></html>`,
  );
  const buf = await pageCtx.screenshot({ omitBackground: true, clip: { x: 0, y: 0, width: px, height: px } });
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, buf);
  console.log('wrote', out.replace(root + '/', ''), `${px}px`);
}

// Legacy icons: crop to the central 72x72 of the 108 canvas (what launchers show).
const legacySquare = svg([...bg, ...fg], { viewBox: '18 18 72 72', clip: 'rounded' });
const legacyRound = svg([...bg, ...fg], { viewBox: '18 18 72 72', clip: 'circle' });
const densities = { mdpi: 48, hdpi: 72, xhdpi: 96, xxhdpi: 144, xxxhdpi: 192 };
for (const [d, px] of Object.entries(densities)) {
  await renderPng(legacySquare, px, join(res, `mipmap-${d}/ic_launcher.png`));
  await renderPng(legacyRound, px, join(res, `mipmap-${d}/ic_launcher_round.png`));
  const oldFg = join(res, `mipmap-${d}/ic_launcher_foreground.png`);
  if (existsSync(oldFg)) rmSync(oldFg);
}
await renderPng(svg([...bg, ...fg], { viewBox: '14 14 80 80', clip: 'rounded' }), 512, join(root, 'resources/icon-512.png'));

await browser.close();
