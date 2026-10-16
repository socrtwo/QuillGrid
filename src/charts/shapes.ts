/** Shape renderer (SVG). Placeholder until the full shape library lands. */
import type { DrawingObject } from '../model/types';
import { escHtml } from '../util/sanitize';

export function renderShapeSvg(d: DrawingObject): string {
  const w = d.at.w, h = d.at.h;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><rect x="1" y="1" width="${w - 2}" height="${h - 2}" fill="${escHtml(d.fill || '#4a90c2')}" stroke="${escHtml(d.stroke || '#1f3b57')}"/>${d.text ? `<text x="${w / 2}" y="${h / 2}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="14">${escHtml(d.text)}</text>` : ''}</svg>`;
}
