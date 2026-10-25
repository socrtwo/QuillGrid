// Anne's Office suite icon family: same tile, same white quill feather, app-specific white card.
// Writes <app>/icon.svg (full tile), <app>/icon-foreground.svg (adaptive-icon foreground, transparent),
// <app>/icon-background.txt (solid color for adaptive backgrounds) and preview.html.
import { writeFileSync, mkdirSync } from 'node:fs';
const OUT = new URL('.', import.meta.url).pathname;
export const APPS = {
  QuillPen:   { color: '#2A57B8', dark: '#173A86', light: '#3E74D9', kind: 'pen' },
  QuillGrid:  { color: '#1E7A46', dark: '#12532F', light: '#2E9A5C', kind: 'grid' },
  QuillPost:  { color: '#0D7C86', dark: '#075961', light: '#1A9BA6', kind: 'post' },
  QuillStage: { color: '#C2502A', dark: '#8F3618', light: '#DE6A3E', kind: 'stage' },
};
function card(a) {
  const c = a.color;
  switch (a.kind) {
    case 'pen': return `<rect x="112" y="128" width="232" height="272" rx="24" fill="#fff"/>
      <g fill="${c}"><rect x="146" y="178" width="164" height="18" rx="9"/><rect x="146" y="222" width="164" height="18" rx="9"/><rect x="146" y="266" width="120" height="18" rx="9"/><rect x="146" y="310" width="150" height="18" rx="9"/><rect x="146" y="354" width="90" height="18" rx="9" opacity=".55"/></g>`;
    case 'grid': return `<rect x="112" y="128" width="232" height="272" rx="24" fill="#fff"/>
      <g fill="${c}"><rect x="144" y="162" width="168" height="34" rx="8"/></g>
      <g fill="${c}" opacity=".85">${[0,1,2,3].map(r=>[0,1,2].map(k=>`<rect x="${144+k*58}" y="${210+r*44}" width="50" height="34" rx="7"/>`).join('')).join('')}</g>`;
    case 'post': return `<rect x="88" y="168" width="280" height="200" rx="24" fill="#fff"/>
      <path d="M104 188 L228 286 L352 188" fill="none" stroke="${c}" stroke-width="22" stroke-linecap="round" stroke-linejoin="round"/>
      <path d="M104 352 L190 280 M352 352 L266 280" stroke="${c}" stroke-width="14" stroke-linecap="round" opacity=".55"/>`;
    case 'stage': return `<rect x="80" y="160" width="296" height="186" rx="22" fill="#fff"/>
      <g fill="${c}"><rect x="118" y="270" width="36" height="46" rx="6"/><rect x="170" y="236" width="36" height="80" rx="6"/><rect x="222" y="206" width="36" height="110" rx="6"/></g>
      <path d="M290 222 L336 252 L290 282 Z" fill="${c}" opacity=".8"/>
      <rect x="216" y="346" width="24" height="44" fill="#fff"/><rect x="172" y="384" width="112" height="20" rx="10" fill="#fff"/>`;
  }
}
// The shared quill: a curved feather from the upper right down to a nib at the lower left of the card's right edge.
function quill(a) {
  return `<g>
    <path d="M436 70 C 376 84 316 140 290 214 C 274 260 268 300 270 338 C 300 300 330 280 360 262 C 330 262 312 262 300 258 C 336 240 368 218 392 188 C 368 190 350 190 336 186 C 378 160 410 120 436 70 Z" fill="#fff" stroke="${a.dark}" stroke-width="8" stroke-linejoin="round"/>
    <path d="M430 78 C 360 150 300 250 262 380" fill="none" stroke="${a.dark}" stroke-width="9" stroke-linecap="round"/>
    <path d="M262 380 L252 420 L276 392 Z" fill="${a.dark}"/>
  </g>`;
}
function tile(a, id) {
  return `<defs><linearGradient id="g${id}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${a.light}"/><stop offset="1" stop-color="${a.dark}"/></linearGradient></defs>
  <rect width="512" height="512" rx="116" fill="url(#g${id})"/>`;
}
const svg = (inner, vb = '0 0 512 512') => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${vb}" width="512" height="512">${inner}</svg>\n`;
let preview = '<body style="margin:0;padding:24px;background:#e8eaee;display:flex;gap:28px;flex-wrap:wrap;font-family:sans-serif">';
for (const [name, a] of Object.entries(APPS)) {
  mkdirSync(OUT + name, { recursive: true });
  const id = name.toLowerCase();
  const full = svg(tile(a, id) + card(a) + quill(a));
  // Adaptive icons: the system masks a 108dp square; keep artwork inside the central 66dp safe circle (the quill tip is 259px from the center, 0.6 keeps it within 156px).
  const fg = svg(`<g transform="translate(256 256) scale(.6) translate(-256 -256)">${card(a)}${quill(a)}</g>`);
  writeFileSync(OUT + name + '/icon.svg', full);
  writeFileSync(OUT + name + '/icon-foreground.svg', fg);
  writeFileSync(OUT + name + '/icon-background.txt', a.color + '\n');
  preview += `<figure style="margin:0;text-align:center"><img src="${name}/icon.svg" width="160"><br><img src="${name}/icon.svg" width="48"> <img src="${name}/icon.svg" width="24"><figcaption>${name}</figcaption></figure>`;
  preview += `<figure style="margin:0;text-align:center"><div style="width:160px;height:160px;border-radius:50%;background:${a.color};overflow:hidden"><img src="${name}/icon-foreground.svg" width="160"></div><figcaption>adaptive (round mask)</figcaption></figure>`;
}
writeFileSync(OUT + 'preview.html', preview + '</body>');
console.log('ok');
