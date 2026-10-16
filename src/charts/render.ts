/** Chart renderer (SVG). Placeholder until the full renderer lands. */
import type { BookTab } from '../app';
import type { ChartDef } from '../model/types';
import { escHtml } from '../util/sanitize';

export function renderChartSvg(tab: BookTab, chart: ChartDef, w: number, h: number): string {
  void tab;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><rect width="100%" height="100%" fill="#fff" stroke="#ccc"/><text x="${w / 2}" y="${h / 2}" text-anchor="middle" font-family="sans-serif" font-size="14">${escHtml(chart.title || chart.type)}</text></svg>`;
}
