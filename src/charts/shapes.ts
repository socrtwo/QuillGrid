/**
 * Shape library and renderer (SVG) for floating drawings: preset shapes (OOXML preset names, so
 * shapes round-trip through .xlsx), text boxes, Decorative Text, icons, diagrams, embedded-file
 * icons, slicers and equations.
 *
 * Drawing data comes from documents, so every color, number and text value is normalized or escaped
 * before it reaches the markup.
 */
import type { DrawingObject, CellStyle } from '../model/types';
import { escHtml, safeColor, safeNum } from '../util/sanitize';

export const SHAPE_PALETTE = ['#2f6fb0', '#e07b28', '#7a8a99', '#e3b505', '#4a90c2', '#5aa33f', '#1f3b57', '#9e480e', '#636363', '#997300'];

export interface ShapeInfo { id: string; label: string; group: string }

const G_LINES = 'Lines', G_RECT = 'Rectangles', G_BASIC = 'Basic Shapes', G_ARROWS = 'Block Arrows', G_EQ = 'Equation Shapes', G_FLOW = 'Flowchart', G_STARS = 'Stars and Banners', G_CALLOUTS = 'Callouts';

export const SHAPES: ShapeInfo[] = [
  { id: 'line', label: 'Line', group: G_LINES }, { id: 'straightConnector1', label: 'Line Arrow', group: G_LINES }, { id: 'doubleArrowLine', label: 'Line Arrow: Double', group: G_LINES },
  { id: 'elbowConnector', label: 'Connector: Elbow', group: G_LINES }, { id: 'curvedConnector', label: 'Connector: Curved', group: G_LINES },
  { id: 'rect', label: 'Rectangle', group: G_RECT }, { id: 'roundRect', label: 'Rectangle: Rounded Corners', group: G_RECT }, { id: 'snip1Rect', label: 'Rectangle: Single Corner Snipped', group: G_RECT },
  { id: 'snip2SameRect', label: 'Rectangle: Top Corners Snipped', group: G_RECT }, { id: 'round1Rect', label: 'Rectangle: Single Corner Rounded', group: G_RECT },
  { id: 'ellipse', label: 'Oval', group: G_BASIC }, { id: 'triangle', label: 'Isosceles Triangle', group: G_BASIC }, { id: 'rtTriangle', label: 'Right Triangle', group: G_BASIC },
  { id: 'parallelogram', label: 'Parallelogram', group: G_BASIC }, { id: 'trapezoid', label: 'Trapezoid', group: G_BASIC }, { id: 'diamond', label: 'Diamond', group: G_BASIC },
  { id: 'pentagon', label: 'Regular Pentagon', group: G_BASIC }, { id: 'hexagon', label: 'Hexagon', group: G_BASIC }, { id: 'heptagon', label: 'Heptagon', group: G_BASIC },
  { id: 'octagon', label: 'Octagon', group: G_BASIC }, { id: 'decagon', label: 'Decagon', group: G_BASIC }, { id: 'dodecagon', label: 'Dodecagon', group: G_BASIC },
  { id: 'pie', label: 'Partial Circle', group: G_BASIC }, { id: 'chord', label: 'Chord', group: G_BASIC }, { id: 'teardrop', label: 'Teardrop', group: G_BASIC },
  { id: 'frame', label: 'Frame', group: G_BASIC }, { id: 'halfFrame', label: 'Half Frame', group: G_BASIC }, { id: 'corner', label: 'L-Shape', group: G_BASIC },
  { id: 'plus', label: 'Cross', group: G_BASIC }, { id: 'plaque', label: 'Plaque', group: G_BASIC }, { id: 'can', label: 'Cylinder', group: G_BASIC },
  { id: 'cube', label: 'Cube', group: G_BASIC }, { id: 'bevel', label: 'Bevel', group: G_BASIC }, { id: 'donut', label: 'Circle: Hollow', group: G_BASIC },
  { id: 'noSmoking', label: '"No" Symbol', group: G_BASIC }, { id: 'blockArc', label: 'Block Arc', group: G_BASIC }, { id: 'foldedCorner', label: 'Rectangle: Folded Corner', group: G_BASIC },
  { id: 'smileyFace', label: 'Smiley Face', group: G_BASIC }, { id: 'heart', label: 'Heart', group: G_BASIC }, { id: 'lightningBolt', label: 'Lightning Bolt', group: G_BASIC },
  { id: 'sun', label: 'Sun', group: G_BASIC }, { id: 'moon', label: 'Moon', group: G_BASIC }, { id: 'cloud', label: 'Cloud', group: G_BASIC },
  { id: 'arc', label: 'Arc', group: G_BASIC }, { id: 'bracketPair', label: 'Double Bracket', group: G_BASIC }, { id: 'bracePair', label: 'Double Brace', group: G_BASIC },
  { id: 'leftBracket', label: 'Left Bracket', group: G_BASIC }, { id: 'rightBracket', label: 'Right Bracket', group: G_BASIC },
  { id: 'leftBrace', label: 'Left Brace', group: G_BASIC }, { id: 'rightBrace', label: 'Right Brace', group: G_BASIC },
  { id: 'rightArrow', label: 'Arrow: Right', group: G_ARROWS }, { id: 'leftArrow', label: 'Arrow: Left', group: G_ARROWS }, { id: 'upArrow', label: 'Arrow: Up', group: G_ARROWS },
  { id: 'downArrow', label: 'Arrow: Down', group: G_ARROWS }, { id: 'leftRightArrow', label: 'Arrow: Left-Right', group: G_ARROWS }, { id: 'upDownArrow', label: 'Arrow: Up-Down', group: G_ARROWS },
  { id: 'quadArrow', label: 'Arrow: Quad', group: G_ARROWS }, { id: 'bentArrow', label: 'Arrow: Bent', group: G_ARROWS }, { id: 'uturnArrow', label: 'Arrow: U-Turn', group: G_ARROWS },
  { id: 'notchedRightArrow', label: 'Arrow: Notched Right', group: G_ARROWS }, { id: 'stripedRightArrow', label: 'Arrow: Striped Right', group: G_ARROWS },
  { id: 'homePlate', label: 'Arrow: Pentagon', group: G_ARROWS }, { id: 'chevron', label: 'Arrow: Chevron', group: G_ARROWS },
  { id: 'mathPlus', label: 'Plus Sign', group: G_EQ }, { id: 'mathMinus', label: 'Minus Sign', group: G_EQ }, { id: 'mathMultiply', label: 'Multiplication Sign', group: G_EQ },
  { id: 'mathDivide', label: 'Division Sign', group: G_EQ }, { id: 'mathEqual', label: 'Equals', group: G_EQ }, { id: 'mathNotEqual', label: 'Not Equal', group: G_EQ },
  { id: 'flowChartProcess', label: 'Flowchart: Process', group: G_FLOW }, { id: 'flowChartAlternateProcess', label: 'Flowchart: Alternate Process', group: G_FLOW },
  { id: 'flowChartDecision', label: 'Flowchart: Decision', group: G_FLOW }, { id: 'flowChartInputOutput', label: 'Flowchart: Data', group: G_FLOW },
  { id: 'flowChartPredefinedProcess', label: 'Flowchart: Predefined Process', group: G_FLOW }, { id: 'flowChartInternalStorage', label: 'Flowchart: Internal Storage', group: G_FLOW },
  { id: 'flowChartDocument', label: 'Flowchart: Document', group: G_FLOW }, { id: 'flowChartMultidocument', label: 'Flowchart: Multidocument', group: G_FLOW },
  { id: 'flowChartTerminator', label: 'Flowchart: Terminator', group: G_FLOW }, { id: 'flowChartPreparation', label: 'Flowchart: Preparation', group: G_FLOW },
  { id: 'flowChartManualInput', label: 'Flowchart: Manual Input', group: G_FLOW }, { id: 'flowChartManualOperation', label: 'Flowchart: Manual Operation', group: G_FLOW },
  { id: 'flowChartConnector', label: 'Flowchart: Connector', group: G_FLOW }, { id: 'flowChartOffpageConnector', label: 'Flowchart: Off-page Connector', group: G_FLOW },
  { id: 'flowChartPunchedCard', label: 'Flowchart: Card', group: G_FLOW }, { id: 'flowChartDelay', label: 'Flowchart: Delay', group: G_FLOW },
  { id: 'flowChartMerge', label: 'Flowchart: Merge', group: G_FLOW }, { id: 'flowChartExtract', label: 'Flowchart: Extract', group: G_FLOW },
  { id: 'flowChartSort', label: 'Flowchart: Sort', group: G_FLOW }, { id: 'flowChartCollate', label: 'Flowchart: Collate', group: G_FLOW },
  { id: 'flowChartOnlineStorage', label: 'Flowchart: Stored Data', group: G_FLOW }, { id: 'flowChartMagneticDisk', label: 'Flowchart: Magnetic Disk', group: G_FLOW },
  { id: 'flowChartDisplay', label: 'Flowchart: Display', group: G_FLOW }, { id: 'flowChartSummingJunction', label: 'Flowchart: Summing Junction', group: G_FLOW },
  { id: 'flowChartOr', label: 'Flowchart: Or', group: G_FLOW },
  { id: 'irregularSeal1', label: 'Explosion: 8 Points', group: G_STARS }, { id: 'irregularSeal2', label: 'Explosion: 14 Points', group: G_STARS },
  { id: 'star4', label: 'Star: 4 Points', group: G_STARS }, { id: 'star5', label: 'Star: 5 Points', group: G_STARS }, { id: 'star6', label: 'Star: 6 Points', group: G_STARS },
  { id: 'star7', label: 'Star: 7 Points', group: G_STARS }, { id: 'star8', label: 'Star: 8 Points', group: G_STARS }, { id: 'star10', label: 'Star: 10 Points', group: G_STARS },
  { id: 'star12', label: 'Star: 12 Points', group: G_STARS }, { id: 'star16', label: 'Star: 16 Points', group: G_STARS }, { id: 'star24', label: 'Star: 24 Points', group: G_STARS },
  { id: 'star32', label: 'Star: 32 Points', group: G_STARS }, { id: 'ribbon2', label: 'Ribbon: Tilted Up', group: G_STARS }, { id: 'wave', label: 'Wave', group: G_STARS },
  { id: 'horizontalScroll', label: 'Scroll: Horizontal', group: G_STARS },
  { id: 'wedgeRectCallout', label: 'Speech Bubble: Rectangle', group: G_CALLOUTS }, { id: 'wedgeRoundRectCallout', label: 'Speech Bubble: Rounded Rectangle', group: G_CALLOUTS },
  { id: 'wedgeEllipseCallout', label: 'Speech Bubble: Oval', group: G_CALLOUTS }, { id: 'cloudCallout', label: 'Thought Bubble: Cloud', group: G_CALLOUTS },
  { id: 'borderCallout1', label: 'Callout: Line', group: G_CALLOUTS }, { id: 'accentCallout1', label: 'Callout: Accent Bar', group: G_CALLOUTS },
];

export const SHAPE_GROUPS = [...new Set(SHAPES.map((s) => s.group))];

/** Older or alternate names (ODS import, earlier versions) mapped to preset names. */
const ALIASES: Record<string, string> = {
  rectangle: 'rect', rounded: 'roundRect', oval: 'ellipse', circle: 'ellipse', rtriangle: 'rtTriangle', cross: 'plus', arrowRight: 'rightArrow', arrowLeft: 'leftArrow',
  arrowUp: 'upArrow', arrowDown: 'downArrow', arrowLR: 'leftRightArrow', smiley: 'smileyFace', lightning: 'lightningBolt', arrowLine: 'straightConnector1',
  brace: 'leftBrace', bracket: 'leftBracket', flowProcess: 'flowChartProcess', flowDecision: 'flowChartDecision', flowTerminator: 'flowChartTerminator',
  flowData: 'flowChartInputOutput', flowDocument: 'flowChartDocument', flowDatabase: 'flowChartMagneticDisk', burst: 'irregularSeal1', callout: 'wedgeRectCallout',
  thought: 'cloudCallout', banner: 'ribbon2', scroll: 'horizontalScroll', textbox: 'rect', textBox: 'rect', star: 'star5', flowChartAlternate: 'flowChartAlternateProcess',
};

export function shapeLabel(id: string | undefined): string {
  const n = ALIASES[id || ''] || id || 'rect';
  return SHAPES.find((s) => s.id === n)?.label || n;
}

// --------------------------------------------------------------------------------------------
// Geometry

interface Geo { body: string; open?: boolean; text?: { x: number; y: number; w: number; h: number }; extra?: string }

const f = (n: number) => (Math.round(n * 100) / 100).toString();

/** Builds geometry for a preset shape inside a box of w × h (margin m for the stroke). */
export function shapeGeometry(id: string, w: number, h: number, m = 2): Geo {
  const W = Math.max(1, w - 2 * m), H = Math.max(1, h - 2 * m);
  const X = (u: number) => m + u * W, Y = (v: number) => m + v * H;
  const pt = (u: number, v: number) => `${f(X(u))},${f(Y(v))}`;
  const poly = (pts: number[][]) => `<path d="M${pts.map(([u, v]) => pt(u, v)).join(' L')} Z"/>`;
  const path = (d: string) => `<path d="${d}"/>`;
  const ell = (cu: number, cv: number, ru: number, rv: number, extra = '') => `<ellipse cx="${f(X(cu))}" cy="${f(Y(cv))}" rx="${f(ru * W)}" ry="${f(rv * H)}"${extra}/>`;
  const regular = (n: number, rot = -Math.PI / 2) => poly(Array.from({ length: n }, (_, i) => { const a = rot + (i * 2 * Math.PI) / n; return [0.5 + 0.5 * Math.cos(a), 0.5 + 0.5 * Math.sin(a)]; }));
  const star = (n: number, inner: number) => poly(Array.from({ length: n * 2 }, (_, i) => { const r = i % 2 ? inner : 0.5; const a = -Math.PI / 2 + (i * Math.PI) / n; return [0.5 + r * Math.cos(a), 0.5 + r * Math.sin(a)]; }));
  const inset = (a: number, b = a) => ({ x: X(a), y: Y(b), w: W * (1 - 2 * a), h: H * (1 - 2 * b) });
  const arcPt = (cx: number, cy: number, rx: number, ry: number, deg: number) => { const a = (deg * Math.PI) / 180; return `${f(cx + rx * Math.cos(a))},${f(cy + ry * Math.sin(a))}`; };
  const cx = X(0.5), cy = Y(0.5), rx = W / 2, ry = H / 2;
  switch (ALIASES[id] || id) {
    case 'line': return { body: `<line x1="${f(X(0))}" y1="${f(Y(0.5))}" x2="${f(X(1))}" y2="${f(Y(0.5))}"/>`, open: true };
    case 'straightConnector1': return { body: `<line x1="${f(X(0))}" y1="${f(Y(0.5))}" x2="${f(X(1))}" y2="${f(Y(0.5))}" marker-end="url(#MK)"/>`, open: true, extra: 'end' };
    case 'doubleArrowLine': return { body: `<line x1="${f(X(0))}" y1="${f(Y(0.5))}" x2="${f(X(1))}" y2="${f(Y(0.5))}" marker-end="url(#MK)" marker-start="url(#MK)"/>`, open: true, extra: 'both' };
    case 'elbowConnector': return { body: `<path d="M${pt(0, 0)} H${f(X(0.5))} V${f(Y(1))} H${f(X(1))}" marker-end="url(#MK)"/>`, open: true, extra: 'end' };
    case 'curvedConnector': return { body: `<path d="M${pt(0, 0)} C${pt(0.6, 0)} ${pt(0.4, 1)} ${pt(1, 1)}" marker-end="url(#MK)"/>`, open: true, extra: 'end' };
    case 'rect': case 'flowChartProcess': return { body: `<rect x="${f(X(0))}" y="${f(Y(0))}" width="${f(W)}" height="${f(H)}"/>`, text: inset(0.05) };
    case 'roundRect': case 'flowChartAlternateProcess': { const r = Math.min(W, H) * 0.16; return { body: `<rect x="${f(X(0))}" y="${f(Y(0))}" width="${f(W)}" height="${f(H)}" rx="${f(r)}"/>`, text: inset(0.06) }; }
    case 'snip1Rect': { const s = Math.min(W, H) * 0.2; return { body: path(`M${f(X(0))},${f(Y(0))} H${f(X(1) - s)} L${f(X(1))},${f(Y(0) + s)} V${f(Y(1))} H${f(X(0))} Z`), text: inset(0.06) }; }
    case 'snip2SameRect': { const s = Math.min(W, H) * 0.2; return { body: path(`M${f(X(0) + s)},${f(Y(0))} H${f(X(1) - s)} L${f(X(1))},${f(Y(0) + s)} V${f(Y(1))} H${f(X(0))} V${f(Y(0) + s)} Z`), text: inset(0.06) }; }
    case 'round1Rect': { const s = Math.min(W, H) * 0.2; return { body: path(`M${f(X(0))},${f(Y(0))} H${f(X(1) - s)} A${f(s)},${f(s)} 0 0 1 ${f(X(1))},${f(Y(0) + s)} V${f(Y(1))} H${f(X(0))} Z`), text: inset(0.06) }; }
    case 'ellipse': case 'flowChartConnector': return { body: ell(0.5, 0.5, 0.5, 0.5), text: inset(0.15) };
    case 'triangle': return { body: poly([[0.5, 0], [1, 1], [0, 1]]), text: { x: X(0.25), y: Y(0.45), w: W * 0.5, h: H * 0.5 } };
    case 'flowChartExtract': return { body: poly([[0.5, 0], [1, 1], [0, 1]]), text: { x: X(0.25), y: Y(0.45), w: W * 0.5, h: H * 0.5 } };
    case 'flowChartMerge': return { body: poly([[0, 0], [1, 0], [0.5, 1]]), text: { x: X(0.25), y: Y(0.05), w: W * 0.5, h: H * 0.5 } };
    case 'rtTriangle': return { body: poly([[0, 0], [1, 1], [0, 1]]), text: { x: X(0.05), y: Y(0.5), w: W * 0.5, h: H * 0.45 } };
    case 'parallelogram': case 'flowChartInputOutput': return { body: poly([[0.2, 0], [1, 0], [0.8, 1], [0, 1]]), text: inset(0.2, 0.08) };
    case 'trapezoid': return { body: poly([[0.2, 0], [0.8, 0], [1, 1], [0, 1]]), text: inset(0.2, 0.08) };
    case 'flowChartManualOperation': return { body: poly([[0, 0], [1, 0], [0.8, 1], [0.2, 1]]), text: inset(0.2, 0.08) };
    case 'diamond': case 'flowChartDecision': return { body: poly([[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]), text: inset(0.25) };
    case 'pentagon': return { body: regular(5), text: inset(0.2) };
    case 'hexagon': return { body: poly([[0.25, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0.25, 1], [0, 0.5]]), text: inset(0.2, 0.1) };
    case 'flowChartPreparation': return { body: poly([[0.2, 0], [0.8, 0], [1, 0.5], [0.8, 1], [0.2, 1], [0, 0.5]]), text: inset(0.2, 0.1) };
    case 'heptagon': return { body: regular(7), text: inset(0.2) };
    case 'octagon': return { body: poly([[0.29, 0], [0.71, 0], [1, 0.29], [1, 0.71], [0.71, 1], [0.29, 1], [0, 0.71], [0, 0.29]]), text: inset(0.12) };
    case 'decagon': return { body: regular(10, 0), text: inset(0.15) };
    case 'dodecagon': return { body: regular(12, Math.PI / 12), text: inset(0.15) };
    case 'pie': return { body: path(`M${f(cx)},${f(cy)} L${arcPt(cx, cy, rx, ry, 0)} A${f(rx)},${f(ry)} 0 1 1 ${arcPt(cx, cy, rx, ry, -90)} Z`), text: inset(0.2) };
    case 'chord': return { body: path(`M${arcPt(cx, cy, rx, ry, 45)} A${f(rx)},${f(ry)} 0 1 1 ${arcPt(cx, cy, rx, ry, -90)} Z`), text: inset(0.2) };
    case 'teardrop': return { body: path(`M${f(cx)},${f(Y(0))} H${f(X(1))} V${f(cy)} A${f(rx)},${f(ry)} 0 1 1 ${f(cx)},${f(Y(0))} Z`), text: inset(0.18) };
    case 'frame': { const t = Math.min(W, H) * 0.12; return { body: `<path fill-rule="evenodd" d="M${f(X(0))},${f(Y(0))} H${f(X(1))} V${f(Y(1))} H${f(X(0))} Z M${f(X(0) + t)},${f(Y(0) + t)} V${f(Y(1) - t)} H${f(X(1) - t)} V${f(Y(0) + t)} Z"/>`, text: inset(0.15) }; }
    case 'halfFrame': { const t = Math.min(W, H) * 0.15; return { body: path(`M${f(X(0))},${f(Y(0))} H${f(X(1))} L${f(X(1) - t)},${f(Y(0) + t)} H${f(X(0) + t)} V${f(Y(1) - t)} L${f(X(0))},${f(Y(1))} Z`) }; }
    case 'corner': return { body: poly([[0, 0], [0.4, 0], [0.4, 0.6], [1, 0.6], [1, 1], [0, 1]]) };
    case 'plus': return { body: poly([[0.3, 0], [0.7, 0], [0.7, 0.3], [1, 0.3], [1, 0.7], [0.7, 0.7], [0.7, 1], [0.3, 1], [0.3, 0.7], [0, 0.7], [0, 0.3], [0.3, 0.3]]), text: inset(0.3) };
    case 'plaque': { const r = Math.min(W, H) * 0.16; return { body: path(`M${f(X(0) + r)},${f(Y(0))} H${f(X(1) - r)} A${f(r)},${f(r)} 0 0 0 ${f(X(1))},${f(Y(0) + r)} V${f(Y(1) - r)} A${f(r)},${f(r)} 0 0 0 ${f(X(1) - r)},${f(Y(1))} H${f(X(0) + r)} A${f(r)},${f(r)} 0 0 0 ${f(X(0))},${f(Y(1) - r)} V${f(Y(0) + r)} A${f(r)},${f(r)} 0 0 0 ${f(X(0) + r)},${f(Y(0))} Z`), text: inset(0.12) }; }
    case 'can': case 'flowChartMagneticDisk': { const e = H * 0.12; return { body: path(`M${f(X(0))},${f(Y(0) + e)} A${f(rx)},${f(e)} 0 0 1 ${f(X(1))},${f(Y(0) + e)} V${f(Y(1) - e)} A${f(rx)},${f(e)} 0 0 1 ${f(X(0))},${f(Y(1) - e)} Z`) + `<path fill="none" d="M${f(X(0))},${f(Y(0) + e)} A${f(rx)},${f(e)} 0 0 0 ${f(X(1))},${f(Y(0) + e)}"/>`, text: { x: X(0.05), y: Y(0) + 2.5 * e, w: W * 0.9, h: H - 3.5 * e } }; }
    case 'cube': { const d = Math.min(W, H) * 0.25; return { body: path(`M${f(X(0))},${f(Y(0) + d)} L${f(X(0) + d)},${f(Y(0))} H${f(X(1))} V${f(Y(1) - d)} L${f(X(1) - d)},${f(Y(1))} H${f(X(0))} Z`) + `<path fill="none" d="M${f(X(0))},${f(Y(0) + d)} H${f(X(1) - d)} V${f(Y(1))} M${f(X(1) - d)},${f(Y(0) + d)} L${f(X(1))},${f(Y(0))}"/><path fill="#000" fill-opacity="0.12" stroke="none" d="M${f(X(1) - d)},${f(Y(0) + d)} L${f(X(1))},${f(Y(0))} V${f(Y(1) - d)} L${f(X(1) - d)},${f(Y(1))} Z"/>`, text: { x: X(0), y: Y(0) + d, w: W - d, h: H - d } }; }
    case 'bevel': { const d = Math.min(W, H) * 0.12; return { body: `<rect x="${f(X(0))}" y="${f(Y(0))}" width="${f(W)}" height="${f(H)}"/><rect x="${f(X(0) + d)}" y="${f(Y(0) + d)}" width="${f(W - 2 * d)}" height="${f(H - 2 * d)}" fill="#fff" fill-opacity="0.25"/><path fill="none" d="M${f(X(0))},${f(Y(0))} L${f(X(0) + d)},${f(Y(0) + d)} M${f(X(1))},${f(Y(0))} L${f(X(1) - d)},${f(Y(0) + d)} M${f(X(1))},${f(Y(1))} L${f(X(1) - d)},${f(Y(1) - d)} M${f(X(0))},${f(Y(1))} L${f(X(0) + d)},${f(Y(1) - d)}"/>`, text: inset(0.12) }; }
    case 'donut': return { body: `<path fill-rule="evenodd" d="M${f(X(0))},${f(cy)} A${f(rx)},${f(ry)} 0 1 1 ${f(X(1))},${f(cy)} A${f(rx)},${f(ry)} 0 1 1 ${f(X(0))},${f(cy)} Z M${f(X(0.25))},${f(cy)} A${f(rx / 2)},${f(ry / 2)} 0 1 0 ${f(X(0.75))},${f(cy)} A${f(rx / 2)},${f(ry / 2)} 0 1 0 ${f(X(0.25))},${f(cy)} Z"/>` };
    case 'noSmoking': { const k = 0.7; return { body: `<path fill-rule="evenodd" d="M${f(X(0))},${f(cy)} A${f(rx)},${f(ry)} 0 1 1 ${f(X(1))},${f(cy)} A${f(rx)},${f(ry)} 0 1 1 ${f(X(0))},${f(cy)} Z M${arcPt(cx, cy, rx * k, ry * k, -135 + 8)} A${f(rx * k)},${f(ry * k)} 0 0 1 ${arcPt(cx, cy, rx * k, ry * k, 45 - 8)} Z M${arcPt(cx, cy, rx * k, ry * k, 45 + 8)} A${f(rx * k)},${f(ry * k)} 0 0 1 ${arcPt(cx, cy, rx * k, ry * k, 225 - 8)} Z"/>` }; }
    case 'blockArc': { const k = 0.55; return { body: path(`M${arcPt(cx, cy, rx, ry, 180)} A${f(rx)},${f(ry)} 0 0 1 ${arcPt(cx, cy, rx, ry, 0)} L${arcPt(cx, cy, rx * k, ry * k, 0)} A${f(rx * k)},${f(ry * k)} 0 0 0 ${arcPt(cx, cy, rx * k, ry * k, 180)} Z`) }; }
    case 'arc': return { body: `<path d="M${arcPt(cx, cy, rx, ry, -90)} A${f(rx)},${f(ry)} 0 0 1 ${arcPt(cx, cy, rx, ry, 0)}"/>`, open: true };
    case 'foldedCorner': { const d = Math.min(W, H) * 0.18; return { body: path(`M${f(X(0))},${f(Y(0))} H${f(X(1))} V${f(Y(1) - d)} L${f(X(1) - d)},${f(Y(1))} H${f(X(0))} Z`) + `<path fill="#000" fill-opacity="0.15" d="M${f(X(1))},${f(Y(1) - d)} L${f(X(1) - d * 0.8)},${f(Y(1) - d * 0.8)} L${f(X(1) - d)},${f(Y(1))} Z"/>`, text: inset(0.06) }; }
    case 'smileyFace': return { body: ell(0.5, 0.5, 0.5, 0.5) + `<g fill="currentColor" stroke="none">${ell(0.34, 0.37, 0.06, 0.07)}${ell(0.66, 0.37, 0.06, 0.07)}</g><path fill="none" d="M${pt(0.27, 0.62)} Q${pt(0.5, 0.86)} ${pt(0.73, 0.62)}"/>` };
    case 'heart': return { body: path(`M${pt(0.5, 0.95)} C${pt(0.05, 0.6)} ${pt(-0.02, 0.18)} ${pt(0.25, 0.08)} C${pt(0.4, 0.03)} ${pt(0.5, 0.15)} ${pt(0.5, 0.25)} C${pt(0.5, 0.15)} ${pt(0.6, 0.03)} ${pt(0.75, 0.08)} C${pt(1.02, 0.18)} ${pt(0.95, 0.6)} ${pt(0.5, 0.95)} Z`), text: inset(0.25, 0.25) };
    case 'lightningBolt': return { body: poly([[0.38, 0], [0.62, 0.28], [0.52, 0.33], [0.8, 0.6], [0.7, 0.64], [1, 1], [0.48, 0.7], [0.58, 0.65], [0.2, 0.42], [0.32, 0.36], [0, 0.15]]) };
    case 'sun': { let rays = ''; for (let i = 0; i < 8; i++) { const a = (i * Math.PI) / 4; const a1 = a - 0.16, a2 = a + 0.16; rays += `M${f(cx + rx * 0.62 * Math.cos(a1))},${f(cy + ry * 0.62 * Math.sin(a1))} L${f(cx + rx * Math.cos(a))},${f(cy + ry * Math.sin(a))} L${f(cx + rx * 0.62 * Math.cos(a2))},${f(cy + ry * 0.62 * Math.sin(a2))} Z `; } return { body: path(rays) + ell(0.5, 0.5, 0.26, 0.26), text: inset(0.3) }; }
    case 'moon': return { body: path(`M${pt(0.75, 0)} A${f(rx * 1.0)},${f(ry)} 0 1 0 ${pt(0.75, 1)} A${f(rx * 0.75)},${f(ry * 0.85)} 0 1 1 ${pt(0.75, 0)} Z`) };
    case 'cloud': case 'cloudCallout': {
      const c = path(`M${pt(0.22, 0.85)} C${pt(0.0, 0.86)} ${pt(0.0, 0.5)} ${pt(0.16, 0.48)} C${pt(0.08, 0.2)} ${pt(0.38, 0.06)} ${pt(0.5, 0.24)} C${pt(0.6, 0.02)} ${pt(0.92, 0.1)} ${pt(0.86, 0.42)} C${pt(1.02, 0.5)} ${pt(1.0, 0.86)} ${pt(0.78, 0.85)} C${pt(0.7, 1.0)} ${pt(0.32, 1.0)} ${pt(0.22, 0.85)} Z`);
      if ((ALIASES[id] || id) === 'cloud') return { body: c, text: inset(0.2) };
      const k = 0.78;
      const cc = path(`M${pt(0.22, 0.85 * k)} C${pt(0.0, 0.86 * k)} ${pt(0.0, 0.5 * k)} ${pt(0.16, 0.48 * k)} C${pt(0.08, 0.2 * k)} ${pt(0.38, 0.06 * k)} ${pt(0.5, 0.24 * k)} C${pt(0.6, 0.02 * k)} ${pt(0.92, 0.1 * k)} ${pt(0.86, 0.42 * k)} C${pt(1.02, 0.5 * k)} ${pt(1.0, 0.86 * k)} ${pt(0.78, 0.85 * k)} C${pt(0.7, k)} ${pt(0.32, k)} ${pt(0.22, 0.85 * k)} Z`);
      return { body: cc + ell(0.2, 0.86, 0.05, 0.04) + ell(0.12, 0.95, 0.03, 0.025), text: { x: X(0.2), y: Y(0.15), w: W * 0.6, h: H * 0.5 } };
    }
    case 'bracketPair': return { body: `<path d="M${pt(0.12, 0)} H${f(X(0))} V${f(Y(1))} H${f(X(0.12))} M${pt(0.88, 0)} H${f(X(1))} V${f(Y(1))} H${f(X(0.88))}"/>`, open: true, text: inset(0.1) };
    case 'bracePair': return { body: `<path d="M${pt(0.12, 0)} Q${pt(0.04, 0)} ${pt(0.04, 0.1)} V${f(Y(0.42))} Q${pt(0.04, 0.5)} ${pt(0, 0.5)} Q${pt(0.04, 0.5)} ${pt(0.04, 0.58)} V${f(Y(0.9))} Q${pt(0.04, 1)} ${pt(0.12, 1)} M${pt(0.88, 0)} Q${pt(0.96, 0)} ${pt(0.96, 0.1)} V${f(Y(0.42))} Q${pt(0.96, 0.5)} ${pt(1, 0.5)} Q${pt(0.96, 0.5)} ${pt(0.96, 0.58)} V${f(Y(0.9))} Q${pt(0.96, 1)} ${pt(0.88, 1)}"/>`, open: true, text: inset(0.1) };
    case 'leftBracket': return { body: `<path d="M${pt(1, 0)} Q${pt(0, 0)} ${pt(0, 0.1)} V${f(Y(0.9))} Q${pt(0, 1)} ${pt(1, 1)}"/>`, open: true };
    case 'rightBracket': return { body: `<path d="M${pt(0, 0)} Q${pt(1, 0)} ${pt(1, 0.1)} V${f(Y(0.9))} Q${pt(1, 1)} ${pt(0, 1)}"/>`, open: true };
    case 'leftBrace': return { body: `<path d="M${pt(1, 0)} Q${pt(0.5, 0)} ${pt(0.5, 0.1)} V${f(Y(0.4))} Q${pt(0.5, 0.5)} ${pt(0, 0.5)} Q${pt(0.5, 0.5)} ${pt(0.5, 0.6)} V${f(Y(0.9))} Q${pt(0.5, 1)} ${pt(1, 1)}"/>`, open: true };
    case 'rightBrace': return { body: `<path d="M${pt(0, 0)} Q${pt(0.5, 0)} ${pt(0.5, 0.1)} V${f(Y(0.4))} Q${pt(0.5, 0.5)} ${pt(1, 0.5)} Q${pt(0.5, 0.5)} ${pt(0.5, 0.6)} V${f(Y(0.9))} Q${pt(0.5, 1)} ${pt(0, 1)}"/>`, open: true };
    case 'rightArrow': return { body: poly([[0, 0.25], [0.6, 0.25], [0.6, 0], [1, 0.5], [0.6, 1], [0.6, 0.75], [0, 0.75]]), text: { x: X(0), y: Y(0.25), w: W * 0.8, h: H * 0.5 } };
    case 'leftArrow': return { body: poly([[1, 0.25], [0.4, 0.25], [0.4, 0], [0, 0.5], [0.4, 1], [0.4, 0.75], [1, 0.75]]), text: { x: X(0.2), y: Y(0.25), w: W * 0.8, h: H * 0.5 } };
    case 'upArrow': return { body: poly([[0.25, 1], [0.25, 0.4], [0, 0.4], [0.5, 0], [1, 0.4], [0.75, 0.4], [0.75, 1]]), text: { x: X(0.25), y: Y(0.3), w: W * 0.5, h: H * 0.7 } };
    case 'downArrow': return { body: poly([[0.25, 0], [0.25, 0.6], [0, 0.6], [0.5, 1], [1, 0.6], [0.75, 0.6], [0.75, 0]]), text: { x: X(0.25), y: Y(0), w: W * 0.5, h: H * 0.7 } };
    case 'leftRightArrow': return { body: poly([[0, 0.5], [0.22, 0], [0.22, 0.25], [0.78, 0.25], [0.78, 0], [1, 0.5], [0.78, 1], [0.78, 0.75], [0.22, 0.75], [0.22, 1]]), text: { x: X(0.15), y: Y(0.25), w: W * 0.7, h: H * 0.5 } };
    case 'upDownArrow': return { body: poly([[0.5, 0], [1, 0.22], [0.75, 0.22], [0.75, 0.78], [1, 0.78], [0.5, 1], [0, 0.78], [0.25, 0.78], [0.25, 0.22], [0, 0.22]]), text: { x: X(0.25), y: Y(0.15), w: W * 0.5, h: H * 0.7 } };
    case 'quadArrow': return { body: poly([[0.5, 0], [0.68, 0.18], [0.58, 0.18], [0.58, 0.42], [0.82, 0.42], [0.82, 0.32], [1, 0.5], [0.82, 0.68], [0.82, 0.58], [0.58, 0.58], [0.58, 0.82], [0.68, 0.82], [0.5, 1], [0.32, 0.82], [0.42, 0.82], [0.42, 0.58], [0.18, 0.58], [0.18, 0.68], [0, 0.5], [0.18, 0.32], [0.18, 0.42], [0.42, 0.42], [0.42, 0.18], [0.32, 0.18]]) };
    case 'bentArrow': return { body: path(`M${pt(0, 1)} V${f(Y(0.5))} Q${pt(0, 0.25)} ${pt(0.3, 0.25)} H${f(X(0.72))} V${f(Y(0))} L${pt(1, 0.37)} L${pt(0.72, 0.74)} V${f(Y(0.5))} H${f(X(0.35))} Q${pt(0.26, 0.5)} ${pt(0.26, 0.6)} V${f(Y(1))} Z`) };
    case 'uturnArrow': return { body: path(`M${pt(0, 1)} V${f(Y(0.35))} Q${pt(0, 0)} ${pt(0.35, 0)} H${f(X(0.5))} Q${pt(0.82, 0)} ${pt(0.82, 0.35)} V${f(Y(0.55))} H${f(X(1))} L${pt(0.7, 0.85)} L${pt(0.4, 0.55)} H${f(X(0.58))} V${f(Y(0.38))} Q${pt(0.58, 0.25)} ${pt(0.45, 0.25)} H${f(X(0.38))} Q${pt(0.25, 0.25)} ${pt(0.25, 0.38)} V${f(Y(1))} Z`) };
    case 'notchedRightArrow': return { body: poly([[0, 0.25], [0.6, 0.25], [0.6, 0], [1, 0.5], [0.6, 1], [0.6, 0.75], [0, 0.75], [0.15, 0.5]]), text: { x: X(0.15), y: Y(0.25), w: W * 0.65, h: H * 0.5 } };
    case 'stripedRightArrow': return { body: poly([[0.18, 0.25], [0.6, 0.25], [0.6, 0], [1, 0.5], [0.6, 1], [0.6, 0.75], [0.18, 0.75]]) + `<rect x="${f(X(0))}" y="${f(Y(0.25))}" width="${f(W * 0.04)}" height="${f(H * 0.5)}"/><rect x="${f(X(0.08))}" y="${f(Y(0.25))}" width="${f(W * 0.06)}" height="${f(H * 0.5)}"/>`, text: { x: X(0.18), y: Y(0.25), w: W * 0.6, h: H * 0.5 } };
    case 'homePlate': return { body: poly([[0, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0, 1]]), text: inset(0.05) };
    case 'chevron': return { body: poly([[0, 0], [0.7, 0], [1, 0.5], [0.7, 1], [0, 1], [0.3, 0.5]]), text: { x: X(0.3), y: Y(0.1), w: W * 0.4, h: H * 0.8 } };
    case 'mathPlus': return { body: poly([[0.38, 0.1], [0.62, 0.1], [0.62, 0.38], [0.9, 0.38], [0.9, 0.62], [0.62, 0.62], [0.62, 0.9], [0.38, 0.9], [0.38, 0.62], [0.1, 0.62], [0.1, 0.38], [0.38, 0.38]]) };
    case 'mathMinus': return { body: poly([[0.1, 0.38], [0.9, 0.38], [0.9, 0.62], [0.1, 0.62]]) };
    case 'mathMultiply': return { body: poly([[0.2, 0.08], [0.5, 0.38], [0.8, 0.08], [0.92, 0.2], [0.62, 0.5], [0.92, 0.8], [0.8, 0.92], [0.5, 0.62], [0.2, 0.92], [0.08, 0.8], [0.38, 0.5], [0.08, 0.2]]) };
    case 'mathDivide': return { body: poly([[0.1, 0.42], [0.9, 0.42], [0.9, 0.58], [0.1, 0.58]]) + ell(0.5, 0.22, 0.09, 0.09) + ell(0.5, 0.78, 0.09, 0.09) };
    case 'mathEqual': return { body: poly([[0.1, 0.25], [0.9, 0.25], [0.9, 0.42], [0.1, 0.42]]) + poly([[0.1, 0.58], [0.9, 0.58], [0.9, 0.75], [0.1, 0.75]]) };
    case 'mathNotEqual': return { body: poly([[0.1, 0.25], [0.9, 0.25], [0.9, 0.42], [0.1, 0.42]]) + poly([[0.1, 0.58], [0.9, 0.58], [0.9, 0.75], [0.1, 0.75]]) + poly([[0.62, 0.05], [0.72, 0.1], [0.38, 0.95], [0.28, 0.9]]) };
    case 'flowChartPredefinedProcess': return { body: `<rect x="${f(X(0))}" y="${f(Y(0))}" width="${f(W)}" height="${f(H)}"/><path fill="none" d="M${pt(0.12, 0)} V${f(Y(1))} M${pt(0.88, 0)} V${f(Y(1))}"/>`, text: inset(0.14, 0.05) };
    case 'flowChartInternalStorage': return { body: `<rect x="${f(X(0))}" y="${f(Y(0))}" width="${f(W)}" height="${f(H)}"/><path fill="none" d="M${pt(0.15, 0)} V${f(Y(1))} M${pt(0, 0.2)} H${f(X(1))}"/>`, text: { x: X(0.17), y: Y(0.22), w: W * 0.8, h: H * 0.75 } };
    case 'flowChartDocument': return { body: path(`M${pt(0, 0)} H${f(X(1))} V${f(Y(0.82))} C${pt(0.75, 0.65)} ${pt(0.5, 1.05)} ${pt(0, 0.88)} Z`), text: inset(0.05, 0.05) };
    case 'flowChartMultidocument': return { body: path(`M${pt(0.1, 0)} H${f(X(1))} V${f(Y(0.7))} L${pt(0.95, 0.7)} V${f(Y(0.08))} H${f(X(0.1))} Z`) + path(`M${pt(0.05, 0.06)} H${f(X(0.95))} V${f(Y(0.76))} L${pt(0.9, 0.76)} V${f(Y(0.14))} H${f(X(0.05))} Z`) + path(`M${pt(0, 0.12)} H${f(X(0.9))} V${f(Y(0.85))} C${pt(0.68, 0.72)} ${pt(0.45, 1.05)} ${pt(0, 0.92)} Z`), text: { x: X(0.02), y: Y(0.15), w: W * 0.85, h: H * 0.7 } };
    case 'flowChartTerminator': return { body: `<rect x="${f(X(0))}" y="${f(Y(0))}" width="${f(W)}" height="${f(H)}" rx="${f(Math.min(H / 2, W / 2))}"/>`, text: inset(0.12, 0.05) };
    case 'flowChartManualInput': return { body: poly([[0, 0.25], [1, 0], [1, 1], [0, 1]]), text: { x: X(0.05), y: Y(0.25), w: W * 0.9, h: H * 0.72 } };
    case 'flowChartOffpageConnector': return { body: poly([[0, 0], [1, 0], [1, 0.75], [0.5, 1], [0, 0.75]]), text: inset(0.05) };
    case 'flowChartPunchedCard': return { body: poly([[0.2, 0], [1, 0], [1, 1], [0, 1], [0, 0.2]]), text: inset(0.08) };
    case 'flowChartDelay': return { body: path(`M${pt(0, 0)} H${f(X(0.5))} A${f(rx)},${f(ry)} 0 0 1 ${pt(0.5, 1)} H${f(X(0))} Z`), text: inset(0.08, 0.1) };
    case 'flowChartSort': return { body: poly([[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]) + `<path fill="none" d="M${pt(0, 0.5)} H${f(X(1))}"/>` };
    case 'flowChartCollate': return { body: poly([[0, 0], [1, 0], [0, 1], [1, 1]]) };
    case 'flowChartOnlineStorage': return { body: path(`M${pt(0.17, 0)} H${f(X(1))} A${f(W * 0.17)},${f(ry)} 0 0 0 ${pt(1, 1)} H${f(X(0.17))} A${f(W * 0.17)},${f(ry)} 0 0 1 ${pt(0.17, 0)} Z`), text: { x: X(0.15), y: Y(0.05), w: W * 0.65, h: H * 0.9 } };
    case 'flowChartDisplay': return { body: path(`M${pt(0, 0.5)} L${pt(0.17, 0)} H${f(X(0.83))} A${f(W * 0.17)},${f(ry)} 0 0 1 ${pt(0.83, 1)} H${f(X(0.17))} Z`), text: { x: X(0.17), y: Y(0.05), w: W * 0.66, h: H * 0.9 } };
    case 'flowChartSummingJunction': return { body: ell(0.5, 0.5, 0.5, 0.5) + `<path fill="none" d="M${pt(0.15, 0.15)} L${pt(0.85, 0.85)} M${pt(0.85, 0.15)} L${pt(0.15, 0.85)}"/>` };
    case 'flowChartOr': return { body: ell(0.5, 0.5, 0.5, 0.5) + `<path fill="none" d="M${pt(0.5, 0)} V${f(Y(1))} M${pt(0, 0.5)} H${f(X(1))}"/>` };
    case 'irregularSeal1': return { body: poly([[0.5, 0.06], [0.6, 0.28], [0.85, 0.12], [0.75, 0.38], [1, 0.45], [0.78, 0.58], [0.92, 0.85], [0.64, 0.72], [0.55, 1], [0.44, 0.74], [0.18, 0.92], [0.26, 0.64], [0, 0.55], [0.22, 0.42], [0.08, 0.15], [0.36, 0.3]]), text: inset(0.3) };
    case 'irregularSeal2': return { body: poly([[0.45, 0], [0.55, 0.22], [0.7, 0.08], [0.72, 0.28], [0.95, 0.2], [0.82, 0.4], [1, 0.52], [0.8, 0.6], [0.9, 0.82], [0.68, 0.74], [0.62, 1], [0.5, 0.78], [0.36, 0.95], [0.32, 0.74], [0.08, 0.82], [0.2, 0.6], [0, 0.48], [0.2, 0.4], [0.06, 0.18], [0.3, 0.26]]), text: inset(0.3) };
    case 'star4': return { body: star(4, 0.15), text: inset(0.35) };
    case 'star5': return { body: star(5, 0.2), text: inset(0.33) };
    case 'star6': return { body: star(6, 0.29), text: inset(0.3) };
    case 'star7': return { body: star(7, 0.3), text: inset(0.3) };
    case 'star8': return { body: star(8, 0.38), text: inset(0.25) };
    case 'star10': return { body: star(10, 0.4), text: inset(0.22) };
    case 'star12': return { body: star(12, 0.4), text: inset(0.22) };
    case 'star16': return { body: star(16, 0.42), text: inset(0.2) };
    case 'star24': return { body: star(24, 0.44), text: inset(0.18) };
    case 'star32': return { body: star(32, 0.45), text: inset(0.18) };
    case 'ribbon2': return { body: poly([[0, 0.2], [0.15, 0.2], [0.15, 0], [0.85, 0], [0.85, 0.2], [1, 0.2], [0.9, 0.5], [1, 0.8], [0.85, 0.8], [0.15, 0.8], [0, 0.8], [0.1, 0.5]]) + `<path fill="#000" fill-opacity="0.15" d="M${pt(0.15, 0.62)} V${f(Y(0.8))} H${f(X(0.85))} V${f(Y(0.62))} Z"/>`, text: { x: X(0.16), y: Y(0.02), w: W * 0.68, h: H * 0.6 } };
    case 'wave': return { body: path(`M${pt(0, 0.12)} C${pt(0.3, -0.08)} ${pt(0.6, 0.32)} ${pt(1, 0.12)} V${f(Y(0.88))} C${pt(0.7, 1.08)} ${pt(0.4, 0.68)} ${pt(0, 0.88)} Z`), text: inset(0.05, 0.2) };
    case 'horizontalScroll': return { body: path(`M${pt(0.06, 0.12)} H${f(X(1))} V${f(Y(0.92))} H${f(X(0.06))} Z`) + ell(0.06, 0.52, 0.06, 0.4) + `<path fill="none" d="M${pt(0.94, 0.12)} Q${pt(1, 0.06)} ${pt(0.94, 0)} H${f(X(0.12))}"/>`, text: { x: X(0.14), y: Y(0.15), w: W * 0.82, h: H * 0.74 } };
    case 'wedgeRectCallout': return { body: path(`M${pt(0, 0)} H${f(X(1))} V${f(Y(0.72))} H${f(X(0.42))} L${pt(0.2, 1)} L${pt(0.24, 0.72)} H${f(X(0))} Z`), text: { x: X(0.05), y: Y(0.04), w: W * 0.9, h: H * 0.64 } };
    case 'wedgeRoundRectCallout': { const r = Math.min(W, H * 0.72) * 0.16; return { body: path(`M${f(X(0) + r)},${f(Y(0))} H${f(X(1) - r)} Q${pt(1, 0)} ${f(X(1))},${f(Y(0) + r)} V${f(Y(0.72) - r)} Q${pt(1, 0.72)} ${f(X(1) - r)},${f(Y(0.72))} H${f(X(0.42))} L${pt(0.2, 1)} L${pt(0.24, 0.72)} H${f(X(0) + r)} Q${pt(0, 0.72)} ${f(X(0))},${f(Y(0.72) - r)} V${f(Y(0) + r)} Q${pt(0, 0)} ${f(X(0) + r)},${f(Y(0))} Z`), text: { x: X(0.06), y: Y(0.05), w: W * 0.88, h: H * 0.62 } }; }
    case 'wedgeEllipseCallout': return { body: path(`M${arcPt(X(0.5), Y(0.4), rx, H * 0.4, 110)} A${f(rx)},${f(H * 0.4)} 0 1 1 ${arcPt(X(0.5), Y(0.4), rx, H * 0.4, 130)} L${pt(0.12, 1)} Z`), text: { x: X(0.15), y: Y(0.1), w: W * 0.7, h: H * 0.6 } };
    case 'borderCallout1': return { body: `<rect x="${f(X(0.3))}" y="${f(Y(0))}" width="${f(W * 0.7)}" height="${f(H * 0.7)}"/><path fill="none" d="M${pt(0.3, 0.35)} L${pt(0, 1)}"/>`, text: { x: X(0.33), y: Y(0.03), w: W * 0.64, h: H * 0.64 } };
    case 'accentCallout1': return { body: `<rect x="${f(X(0.32))}" y="${f(Y(0))}" width="${f(W * 0.68)}" height="${f(H * 0.7)}"/><path fill="none" d="M${pt(0.3, 0)} V${f(Y(0.7))} M${pt(0.3, 0.35)} L${pt(0, 1)}"/>`, text: { x: X(0.35), y: Y(0.03), w: W * 0.62, h: H * 0.64 } };
    default: return { body: `<rect x="${f(X(0))}" y="${f(Y(0))}" width="${f(W)}" height="${f(H)}"/>`, text: inset(0.05) };
  }
}

// --------------------------------------------------------------------------------------------
// Paint helpers (fills, effects, text)

/** Unique, markup-safe id fragment for a drawing (several SVGs share one document). */
const sid = (id: string) => String(id || 'x').replace(/[^A-Za-z0-9_-]/g, '');

/** A paint value: a color, "none", or "gradient:#from:#to[:angle]". Returns [defs, paint]. */
function paint(v: string | undefined, fallback: string, id: string): [string, string] {
  const s = String(v ?? fallback);
  if (s === 'none' || s === 'transparent') return ['', 'none'];
  const m = /^gradient:([^:]+):([^:]+)(?::(-?\d+))?$/.exec(s);
  if (m) {
    const a = safeColor(m[1], fallback), b = safeColor(m[2], '#ffffff');
    const ang = ((Number(m[3]) || 90) * Math.PI) / 180;
    const x2 = 0.5 + 0.5 * Math.cos(ang), y2 = 0.5 + 0.5 * Math.sin(ang);
    return [`<linearGradient id="${id}" x1="${f(1 - x2)}" y1="${f(1 - y2)}" x2="${f(x2)}" y2="${f(y2)}"><stop offset="0" stop-color="${a}"/><stop offset="1" stop-color="${b}"/></linearGradient>`, `url(#${id})`];
  }
  return ['', safeColor(s, fallback)];
}

/** SVG filter for drawing effects stored in fx (shadow, glow, soft edges). */
function effectFilter(fx: Record<string, any> | undefined, id: string): [string, string] {
  if (!fx) return ['', ''];
  const parts: string[] = [];
  if (fx.shadow) parts.push(`<feDropShadow dx="${safeNum(fx.shadowX ?? 3, 3, -40, 40)}" dy="${safeNum(fx.shadowY ?? 3, 3, -40, 40)}" stdDeviation="${safeNum(fx.shadowBlur ?? 2.5, 2.5, 0, 30)}" flood-color="${safeColor(fx.shadowColor, '#000000')}" flood-opacity="0.38"/>`);
  if (fx.glow) parts.push(`<feMorphology in="SourceAlpha" operator="dilate" radius="2" result="g1"/><feGaussianBlur in="g1" stdDeviation="3" result="g2"/><feFlood flood-color="${safeColor(fx.glow, '#ffc000')}" flood-opacity="0.7"/><feComposite in2="g2" operator="in" result="g3"/><feMerge><feMergeNode in="g3"/><feMergeNode in="SourceGraphic"/></feMerge>`);
  if (fx.soft) parts.push(`<feGaussianBlur stdDeviation="${safeNum(fx.soft, 2, 0.5, 20)}"/>`);
  if (!parts.length) return ['', ''];
  // Chain: each primitive applies to the previous result; drop shadow + glow + blur combine in order.
  return [`<filter id="${id}" x="-30%" y="-30%" width="160%" height="160%">${parts.join('')}</filter>`, ` filter="url(#${id})"`];
}

export function fontFamily(font: Partial<CellStyle> | undefined, dflt = 'Carlito, Calibri, Arial, sans-serif'): string {
  const n = font?.font;
  return n ? `${String(n).replace(/["<>&;{}]/g, '')}, ${dflt}` : dflt;
}

/** Word-wraps text to a width using an average glyph width estimate (no DOM measuring). */
export function wrapLines(text: string, maxW: number, size: number, bold = false): string[] {
  const out: string[] = [];
  const cw = size * (bold ? 0.56 : 0.52);
  for (const para of String(text).split('\n')) {
    if (maxW <= cw) { out.push(para); continue; }
    const words = para.split(/(\s+)/).filter((x) => x.length);
    let line = '';
    for (const wd of words) {
      const t = line + wd;
      if (t.trimEnd().length * cw > maxW && line.trim()) { out.push(line.trimEnd()); line = wd.trimStart(); }
      else line = t;
      while (line.length * cw > maxW && line.length > 1) { const k = Math.max(1, Math.floor(maxW / cw)); out.push(line.slice(0, k)); line = line.slice(k); }
    }
    out.push(line.trimEnd());
  }
  return out;
}

interface TextOpts { x: number; y: number; w: number; h: number; align?: 'left' | 'center' | 'right'; valign?: 'top' | 'middle' | 'bottom'; color?: string; font?: Partial<CellStyle>; size?: number; fill?: string; stroke?: string; strokeWidth?: number; extra?: string }

export function textBlock(text: string, o: TextOpts): string {
  if (!text) return '';
  const size = safeNum(o.size ?? o.font?.size ?? 14, 14, 4, 400) * (o.size ? 1 : 4 / 3);
  const bold = !!o.font?.bold;
  const lines = wrapLines(text, Math.max(8, o.w), size, bold);
  const lh = size * 1.2;
  const total = lines.length * lh;
  const va = o.valign || 'middle';
  const top = va === 'top' ? o.y : va === 'bottom' ? o.y + o.h - total : o.y + (o.h - total) / 2;
  const align = o.align || 'center';
  const tx = align === 'left' ? o.x : align === 'right' ? o.x + o.w : o.x + o.w / 2;
  const anchor = align === 'left' ? 'start' : align === 'right' ? 'end' : 'middle';
  const fill = o.fill ?? safeColor(o.color ?? o.font?.color, '#000000');
  const deco = [o.font?.underline ? 'underline' : '', o.font?.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  return `<text x="${f(tx)}" y="${f(top)}" font-family="${escHtml(fontFamily(o.font))}" font-size="${f(size)}"${bold ? ' font-weight="bold"' : ''}${o.font?.italic ? ' font-style="italic"' : ''}${deco ? ` text-decoration="${deco}"` : ''} text-anchor="${anchor}" fill="${fill}"${o.stroke ? ` stroke="${o.stroke}" stroke-width="${o.strokeWidth ?? 1}" paint-order="stroke"` : ''}${o.extra || ''}>${lines.map((l, i) => `<tspan x="${f(tx)}" y="${f(top + i * lh + size * 0.92)}">${escHtml(l)}</tspan>`).join('')}</text>`;
}

function svgOpen(w: number, h: number, extra = ''): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${f(w)}" height="${f(h)}" viewBox="0 0 ${f(w)} ${f(h)}" overflow="visible"${extra}>`;
}

function flipTransform(fx: Record<string, any> | undefined, w: number, h: number): string {
  if (!fx?.flipH && !fx?.flipV) return '';
  return ` transform="translate(${fx.flipH ? w : 0},${fx.flipV ? h : 0}) scale(${fx.flipH ? -1 : 1},${fx.flipV ? -1 : 1})"`;
}

// --------------------------------------------------------------------------------------------
// Decorative Text presets

export interface WordArtStyle { id: string; label: string; fill: string; outline?: string; outlineW?: number; shadow?: boolean; glow?: string; bold?: boolean; font?: string }
export const WORDART_STYLES: WordArtStyle[] = [
  { id: 'wa1', label: 'Fill: Black, Text Color 1; Shadow', fill: '#222222', shadow: true, bold: true },
  { id: 'wa2', label: 'Fill: Blue, Accent 1; Shadow', fill: '#2f6fb0', shadow: true, bold: true },
  { id: 'wa3', label: 'Fill: Orange, Accent 2; Outline: White', fill: '#e07b28', outline: '#ffffff', outlineW: 1.5, shadow: true, bold: true },
  { id: 'wa4', label: 'Fill: White; Outline: Blue, Accent 1; Glow', fill: '#ffffff', outline: '#2f6fb0', outlineW: 2, glow: '#4a90c2', bold: true },
  { id: 'wa5', label: 'Gradient Fill: Blue to Purple', fill: 'gradient:#2f6fb0:#7b3fa0:0', bold: true, shadow: true },
  { id: 'wa6', label: 'Gradient Fill: Gold', fill: 'gradient:#e3b505:#9e480e:90', outline: '#7f6000', outlineW: 0.8, bold: true },
  { id: 'wa7', label: 'Outline Only: Black', fill: 'none', outline: '#222222', outlineW: 1.6, bold: true },
  { id: 'wa8', label: 'Fill: Green, Accent 6; Glow: Gold', fill: '#5aa33f', glow: '#e3b505', bold: true },
  { id: 'wa9', label: 'Fill: Gray; Hard Shadow', fill: '#7a8a99', outline: '#1f3b57', outlineW: 1, shadow: true, bold: true },
  { id: 'wa10', label: 'Gradient Fill: Sunset', fill: 'gradient:#ff6a00:#ee0979:90', bold: true, shadow: true },
  { id: 'wa11', label: 'Fill: Red; Outline: Dark Red', fill: '#c00000', outline: '#600000', outlineW: 1.2, bold: true },
  { id: 'wa12', label: 'Script: Navy', fill: '#1f3b57', font: 'Georgia', bold: false, shadow: true },
];

function wordArtSvg(d: DrawingObject, w: number, h: number): string {
  const id = sid(d.id);
  const fx = (d as any).fx as Record<string, any> | undefined;
  const text = String(d.text ?? 'Your Text Here');
  const lines = text.split('\n');
  const longest = Math.max(1, ...lines.map((l) => l.length));
  const auto = Math.min((h * 0.8) / lines.length / 1.15, (w * 0.95) / (longest * 0.58));
  const size = Math.max(6, d.font?.size ? Math.min(d.font.size * (4 / 3), auto * 1.6) : auto);
  const [gdef, fill] = paint(d.font?.color || fx?.textFill || '#2f6fb0', '#2f6fb0', `g${id}`);
  const outline = fx?.textOutline && fx.textOutline !== 'none' ? safeColor(fx.textOutline, '#000000') : '';
  const [filt, fattr] = effectFilter({ shadow: fx?.shadow, glow: fx?.glow, shadowX: 2, shadowY: 3, shadowBlur: 1.5 }, `e${id}`);
  const lh = size * 1.15;
  const top = (h - lines.length * lh) / 2;
  const fam = fontFamily(d.font, 'Carlito, Calibri, Arial, sans-serif');
  const tspans = lines.map((l, i) => `<tspan x="${f(w / 2)}" y="${f(top + i * lh + size * 0.88)}">${escHtml(l)}</tspan>`).join('');
  return `${svgOpen(w, h)}<defs>${gdef}${filt}</defs><g${flipTransform(fx, w, h)}><text${fattr} x="${f(w / 2)}" text-anchor="middle" font-family="${escHtml(fam)}" font-size="${f(size)}"${d.font?.bold !== false ? ' font-weight="bold"' : ''}${d.font?.italic ? ' font-style="italic"' : ''} fill="${fill}"${outline ? ` stroke="${outline}" stroke-width="${safeNum(fx?.textOutlineW ?? 1.2, 1.2, 0.2, 12)}" paint-order="stroke" stroke-linejoin="round"` : ''}>${tspans}</text></g></svg>`;
}

// --------------------------------------------------------------------------------------------
// Equations (linear format → SVG text with superscripts/subscripts)

const EQ_WORDS: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π',
  rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω', Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  sum: '∑', prod: '∏', int: '∫', iint: '∬', oint: '∮', infty: '∞', inf: '∞', sqrt: '√', cbrt: '∛', pm: '±', mp: '∓', times: '×', div: '÷', cdot: '·', le: '≤', leq: '≤', ge: '≥', geq: '≥',
  ne: '≠', neq: '≠', approx: '≈', equiv: '≡', propto: '∝', partial: '∂', nabla: '∇', in: '∈', notin: '∉', subset: '⊂', supset: '⊃', cup: '∪', cap: '∩', forall: '∀', exists: '∃',
  rightarrow: '→', to: '→', leftarrow: '←', Rightarrow: '⇒', Leftrightarrow: '⇔', degree: '°', deg: '°', hbar: 'ℏ', ell: 'ℓ', emptyset: '∅', angle: '∠', perp: '⊥', parallel: '∥', ldots: '…', cdots: '⋯',
};

/** Converts linear-format math ("x^2 + \alpha_1", "sqrt(x)", "a/b") into display tokens. */
export function equationRuns(src: string): { t: string; k: 'n' | 'sup' | 'sub' }[] {
  let s = String(src).replace(/\\([A-Za-z]+)/g, (m, w: string) => EQ_WORDS[w] ?? m);
  s = s.replace(/<=/g, '≤').replace(/>=/g, '≥').replace(/!=/g, '≠').replace(/->/g, '→').replace(/\+-/g, '±').replace(/\*/g, '·');
  const out: { t: string; k: 'n' | 'sup' | 'sub' }[] = [];
  let i = 0;
  const push = (t: string, k: 'n' | 'sup' | 'sub') => { if (!t) return; const last = out[out.length - 1]; if (last && last.k === k) last.t += t; else out.push({ t, k }); };
  while (i < s.length) {
    const ch = s[i];
    if ((ch === '^' || ch === '_') && i + 1 < s.length) {
      const k = ch === '^' ? 'sup' : 'sub';
      i++;
      let arg = '';
      if (s[i] === '(' || s[i] === '{') {
        const close = s[i] === '(' ? ')' : '}';
        let depth = 0, j = i;
        for (; j < s.length; j++) { if (s[j] === s[i]) depth++; else if (s[j] === close && --depth === 0) break; }
        arg = s.slice(i + 1, j); i = j + 1;
      } else {
        const m = /^[A-Za-z0-9.αβγδεθλμπσφωΣΠ∞+-]/.exec(s.slice(i));
        if (m) { const mm = /^-?[A-Za-z0-9.]+|^./.exec(s.slice(i))!; arg = mm[0]; i += arg.length; }
      }
      push(arg, k);
      continue;
    }
    push(ch, 'n');
    i++;
  }
  return out;
}

function equationSvg(d: DrawingObject, w: number, h: number): string {
  const runs = equationRuns(d.text || '');
  const base = safeNum(d.font?.size ?? 16, 16, 6, 200) * (4 / 3);
  const color = safeColor(d.font?.color, '#000000');
  const [fdef, fill] = paint(d.fill ?? 'none', 'none', `g${sid(d.id)}`);
  const stroke = d.stroke && d.stroke !== 'none' ? safeColor(d.stroke, '#7f7f7f') : 'none';
  const spans = runs.map((r) => r.k === 'n' ? `<tspan font-size="${f(base)}" dy="0">${escHtml(r.t)}</tspan>` : `<tspan font-size="${f(base * 0.65)}" baseline-shift="${r.k === 'sup' ? 'super' : 'sub'}">${escHtml(r.t)}</tspan>`).join('');
  return `${svgOpen(w, h)}<defs>${fdef}</defs><rect x="0.5" y="0.5" width="${f(w - 1)}" height="${f(h - 1)}" fill="${fill}" stroke="${stroke}"/><text x="${f(w / 2)}" y="${f(h / 2 + base * 0.32)}" text-anchor="middle" font-family="Cambria Math, STIX Two Math, Cambria, Georgia, serif" font-style="italic" fill="${color}">${spans}</text></svg>`;
}

// --------------------------------------------------------------------------------------------
// Diagrams (lists, processes, cycles, hierarchies, matrices, pyramids)

export const DIAGRAMS: { kind: string; label: string; help: string; group: string }[] = [
  { kind: 'list', label: 'Vertical Bullet List', help: 'Items, one per line', group: 'List' },
  { kind: 'blocks', label: 'Basic Block List', help: 'Items, one per line', group: 'List' },
  { kind: 'process', label: 'Basic Process', help: 'Steps in order, one per line', group: 'Process' },
  { kind: 'chevrons', label: 'Chevron Process', help: 'Steps in order, one per line', group: 'Process' },
  { kind: 'timeline', label: 'Timeline', help: '"Date: event", one per line', group: 'Process' },
  { kind: 'cycle', label: 'Basic Cycle', help: 'Stages of a repeating cycle', group: 'Cycle' },
  { kind: 'radial', label: 'Radial Cycle', help: 'First line is the center', group: 'Cycle' },
  { kind: 'hierarchy', label: 'Organization Chart', help: 'Indent child lines with spaces or a leading "-" per level', group: 'Hierarchy' },
  { kind: 'matrix', label: 'Basic Matrix', help: 'Four quadrants, one per line', group: 'Matrix' },
  { kind: 'venn', label: 'Basic Venn', help: '2 or 3 overlapping sets', group: 'Relationship' },
  { kind: 'pyramid', label: 'Basic Pyramid', help: 'Top to bottom', group: 'Pyramid' },
  { kind: 'funnel', label: 'Funnel', help: 'Top to bottom', group: 'Pyramid' },
];

export function diagramSvg(kind: string, text: string, w: number, h: number, colors: string[] = SHAPE_PALETTE, textColor = '#ffffff'): string {
  const raw = String(text || '').split('\n').filter((s) => s.trim());
  const items = kind === 'hierarchy' ? raw : raw.map((s) => s.trim());
  const pal = colors.length ? colors.map((c, i) => safeColor(c, SHAPE_PALETTE[i % SHAPE_PALETTE.length])) : SHAPE_PALETTE;
  const col = (i: number) => pal[i % pal.length];
  const tc = safeColor(textColor, '#ffffff');
  const n = Math.max(1, items.length);
  let body = '';
  const label = (t: string, x: number, y: number, maxW: number, size = 13, color = tc) => textBlock(t, { x: x - maxW / 2, y: y - 40, w: maxW, h: 80, size, fill: color });
  switch (kind) {
    case 'process': {
      const gap = 26, bw = (w - (n - 1) * gap) / n, bh = Math.min(h - 10, 90);
      items.forEach((t, i) => {
        const x = i * (bw + gap), y = (h - bh) / 2;
        body += `<rect x="${f(x)}" y="${f(y)}" width="${f(bw)}" height="${f(bh)}" rx="8" fill="${col(i)}"/>` + label(t, x + bw / 2, y + bh / 2, bw - 10);
        if (i < n - 1) body += `<polygon points="${f(x + bw + 5)},${f(h / 2 - 9)} ${f(x + bw + gap - 5)},${f(h / 2)} ${f(x + bw + 5)},${f(h / 2 + 9)}" fill="#9aa5b1"/>`;
      });
      break;
    }
    case 'chevrons': {
      const p = 22, bw = (w + (n - 1) * 12) / n, bh = Math.min(h - 10, 80);
      items.forEach((t, i) => {
        const x = i * (bw - 12), y = (h - bh) / 2;
        body += `<polygon points="${f(x)},${f(y)} ${f(x + bw - p)},${f(y)} ${f(x + bw)},${f(y + bh / 2)} ${f(x + bw - p)},${f(y + bh)} ${f(x)},${f(y + bh)} ${f(x + (i ? p : 0))},${f(y + bh / 2)}" fill="${col(i)}"/>` + label(t, x + bw / 2 + (i ? p / 4 : 0), y + bh / 2, Math.max(30, bw - 2 * p));
      });
      break;
    }
    case 'cycle': {
      const cx = w / 2, cy = h / 2, R = Math.max(20, Math.min(w, h) / 2 - 42), r = Math.max(16, Math.min(58, (Math.PI * R) / n - 8));
      body += `<circle cx="${f(cx)}" cy="${f(cy)}" r="${f(R)}" fill="none" stroke="#b8c2cc" stroke-width="3" stroke-dasharray="7 6"/>`;
      items.forEach((t, i) => {
        const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
        const x = cx + R * Math.cos(a), y = cy + R * Math.sin(a);
        body += `<circle cx="${f(x)}" cy="${f(y)}" r="${f(r)}" fill="${col(i)}"/>` + label(t, x, y, r * 1.7, 12);
      });
      break;
    }
    case 'radial': {
      const cx = w / 2, cy = h / 2, R = Math.max(20, Math.min(w, h) / 2 - 42);
      const [center, ...rest] = items;
      rest.forEach((t, i) => {
        const a = -Math.PI / 2 + (i * 2 * Math.PI) / Math.max(1, rest.length);
        const x = cx + R * Math.cos(a), y = cy + R * Math.sin(a);
        body += `<line x1="${f(cx)}" y1="${f(cy)}" x2="${f(x)}" y2="${f(y)}" stroke="#b8c2cc" stroke-width="3"/><circle cx="${f(x)}" cy="${f(y)}" r="36" fill="${col(i + 1)}"/>` + label(t, x, y, 64, 11);
      });
      body += `<circle cx="${f(cx)}" cy="${f(cy)}" r="50" fill="${col(0)}"/>` + label(center || '', cx, cy, 90, 14);
      break;
    }
    case 'hierarchy': {
      interface T { text: string; kids: T[]; depth: number }
      const root: T = { text: '', kids: [], depth: -1 };
      const stack: T[] = [root];
      for (const line of items) {
        const m = /^(\s*)(-*)\s*(.*)$/.exec(line)!;
        const depth = Math.floor(m[1].replace(/\t/g, '  ').length / 2) + m[2].length;
        const node: T = { text: m[3], kids: [], depth };
        while (stack.length > 1 && stack[stack.length - 1].depth >= depth) stack.pop();
        stack[stack.length - 1].kids.push(node);
        stack.push(node);
      }
      const levels: T[][] = [];
      const walk = (t: T, d: number) => { (levels[d] ||= []).push(t); t.kids.forEach((k) => walk(k, d + 1)); };
      root.kids.forEach((k) => walk(k, 0));
      const lh = h / Math.max(1, levels.length);
      const bh = Math.min(44, lh - 16);
      const pos = new Map<T, { x: number; y: number }>();
      levels.forEach((lv, li) => {
        const bw = Math.min(150, w / lv.length - 12);
        lv.forEach((t, i) => {
          const x = (w / lv.length) * (i + 0.5), y = lh * (li + 0.5);
          pos.set(t, { x, y });
          body += `<rect x="${f(x - bw / 2)}" y="${f(y - bh / 2)}" width="${f(bw)}" height="${f(bh)}" rx="6" fill="${col(li)}"/>` + label(t.text, x, y, bw - 8, 12);
        });
      });
      let lines = '';
      const link = (t: T) => t.kids.forEach((k) => { const a = pos.get(t)!, b = pos.get(k)!; lines += `<path d="M${f(a.x)},${f(a.y + bh / 2)} V${f((a.y + b.y) / 2)} H${f(b.x)} V${f(b.y - bh / 2)}" fill="none" stroke="#9aa5b1" stroke-width="2"/>`; link(k); });
      root.kids.forEach(link);
      body = lines + body;
      break;
    }
    case 'pyramid': case 'funnel': {
      const lh = h / n;
      items.forEach((t, i) => {
        const k = kind === 'pyramid' ? i : n - 1 - i;
        const a0 = (w / 2) * (k / n), b0 = (w / 2) * ((k + 1) / n);
        const [a, b] = kind === 'pyramid' ? [a0, b0] : [b0, a0];
        const y = i * lh;
        body += `<polygon points="${f(w / 2 - a)},${f(y)} ${f(w / 2 + a)},${f(y)} ${f(w / 2 + b)},${f(y + lh - 3)} ${f(w / 2 - b)},${f(y + lh - 3)}" fill="${col(i)}"/>` + label(t, w / 2, y + lh / 2, Math.max(60, (a + b) * 0.9), 12);
      });
      break;
    }
    case 'venn': {
      const r = Math.min(w, h) * 0.3;
      const centers = n >= 3 ? [[w / 2 - r * 0.6, h / 2 - r * 0.35], [w / 2 + r * 0.6, h / 2 - r * 0.35], [w / 2, h / 2 + r * 0.6]] : [[w / 2 - r * 0.6, h / 2], [w / 2 + r * 0.6, h / 2]];
      centers.forEach((c, i) => { if (items[i] !== undefined) body += `<circle cx="${f(c[0])}" cy="${f(c[1])}" r="${f(r)}" fill="${col(i)}" fill-opacity="0.55"/>`; });
      centers.forEach((c, i) => { if (items[i] !== undefined) body += label(items[i], c[0] + (c[0] - w / 2) * 0.5, c[1] + (c[1] - h / 2) * 0.5, r, 13, '#1a1a1a'); });
      break;
    }
    case 'matrix': {
      [[0, 0], [1, 0], [0, 1], [1, 1]].forEach(([qx, qy], i) => {
        const x = qx * (w / 2), y = qy * (h / 2);
        body += `<rect x="${f(x + 3)}" y="${f(y + 3)}" width="${f(w / 2 - 6)}" height="${f(h / 2 - 6)}" rx="6" fill="${col(i)}"/>` + label(items[i] || '', x + w / 4, y + h / 4, w / 2 - 20);
      });
      break;
    }
    case 'blocks': {
      const cols = Math.ceil(Math.sqrt(n)), rows = Math.ceil(n / cols);
      const bw = w / cols, bh = h / rows;
      items.forEach((t, i) => { const x = (i % cols) * bw, y = Math.floor(i / cols) * bh; body += `<rect x="${f(x + 4)}" y="${f(y + 4)}" width="${f(bw - 8)}" height="${f(bh - 8)}" rx="4" fill="${col(i)}"/>` + label(t, x + bw / 2, y + bh / 2, bw - 16); });
      break;
    }
    case 'timeline': {
      const y = h / 2;
      body += `<line x1="8" y1="${f(y)}" x2="${f(w - 8)}" y2="${f(y)}" stroke="#7a8a99" stroke-width="4"/>`;
      items.forEach((t, i) => {
        const x = 30 + (i * (w - 60)) / Math.max(1, n - 1);
        const [date, ...ev] = t.split(':');
        const up = i % 2 === 0;
        body += `<circle cx="${f(x)}" cy="${f(y)}" r="8" fill="${col(i)}"/><line x1="${f(x)}" y1="${f(y)}" x2="${f(x)}" y2="${f(up ? y - 34 : y + 34)}" stroke="${col(i)}" stroke-width="2"/>`;
        body += label(date.trim(), x, up ? y - 46 : y + 46, 110, 12, '#222222') + label(ev.join(':').trim(), x, up ? y - 70 : y + 70, 120, 11, '#555555');
      });
      break;
    }
    default: { // list
      const lh = h / n;
      items.forEach((t, i) => { body += `<rect x="0" y="${f(i * lh + 3)}" width="14" height="${f(lh - 6)}" fill="${col(i)}"/><rect x="18" y="${f(i * lh + 3)}" width="${f(w - 18)}" height="${f(lh - 6)}" fill="${col(i)}" fill-opacity="0.18"/>` + textBlock(t, { x: 28, y: i * lh, w: w - 40, h: lh, size: 13, align: 'left', fill: '#1a1a1a' }); });
    }
  }
  return `${svgOpen(w, h)}${body}</svg>`;
}

// --------------------------------------------------------------------------------------------
// Slicers (filter buttons over a table column)

export interface SlicerFx { table: string; column: string; items: string[]; selected: string[]; timeline?: boolean; cols?: number }

function slicerSvg(d: DrawingObject, w: number, h: number): string {
  const fx = ((d as any).fx || {}) as Partial<SlicerFx>;
  const items = Array.isArray(fx.items) ? fx.items.map(String) : [];
  const sel = new Set(Array.isArray(fx.selected) ? fx.selected.map(String) : []);
  const accent = safeColor(d.fill, '#2f6fb0');
  const head = 30;
  let body = `<rect x="0.5" y="0.5" width="${f(w - 1)}" height="${f(h - 1)}" rx="4" fill="#ffffff" stroke="#b8c2cc"/>`;
  body += textBlock(String(d.text || fx.column || 'Slicer'), { x: 10, y: 4, w: w - 40, h: head - 6, size: 13, align: 'left', fill: '#1a1a1a', font: { bold: true } });
  // "clear filter" glyph
  body += `<g stroke="${sel.size ? '#c00000' : '#b8c2cc'}" stroke-width="1.6" fill="none"><path d="M${f(w - 26)},9 h12 l-5,6 v6 l-2,1 v-7 z"/><path d="M${f(w - 12)},16 l5,5 m0,-5 l-5,5"/></g>`;
  const cols = Math.max(1, Math.min(6, Number(fx.cols) || 1));
  const bh = 26, gap = 4;
  const bw = (w - 16 - (cols - 1) * gap) / cols;
  items.forEach((it, i) => {
    const x = 8 + (i % cols) * (bw + gap), y = head + 4 + Math.floor(i / cols) * (bh + gap);
    if (y + bh > h - 4) return;
    const on = !sel.size || sel.has(it);
    body += `<rect x="${f(x)}" y="${f(y)}" width="${f(bw)}" height="${bh}" rx="3" fill="${on ? accent : '#f2f4f7'}" fill-opacity="${on ? 0.9 : 1}" stroke="${on ? accent : '#d0d7de'}"/>`;
    body += textBlock(it === '' ? '(blank)' : it, { x: x + 6, y, w: bw - 12, h: bh, size: 12, fill: on ? '#ffffff' : '#57606a', align: fx.timeline ? 'center' : 'left' });
  });
  return `${svgOpen(w, h)}${body}</svg>`;
}

// --------------------------------------------------------------------------------------------
// Main entry

/** Renders any non-picture drawing as SVG markup. */
export function renderShapeSvg(d: DrawingObject): string {
  const w = safeNum(d.at?.w, 120, 1, 8000), h = safeNum(d.at?.h, 80, 1, 8000);
  const fx = (d as any).fx as Record<string, any> | undefined;
  const id = sid(d.id);
  if (d.kind === 'wordart') return wordArtSvg(d, w, h);
  if (d.kind === 'textbox' && d.shape === 'equation') return equationSvg(d, w, h);
  const shape = d.shape || (d.kind === 'textbox' ? 'rect' : 'rect');
  if (d.kind === 'icon' || shape === 'icon') {
    const name = String(d.text || 'star').replace(/[^a-z0-9_]/gi, '');
    const size = Math.min(w, h);
    const [gdef, fill] = paint(d.fill || '#2f6fb0', '#2f6fb0', `g${id}`);
    return `${svgOpen(w, h)}<defs>${gdef}</defs><g${flipTransform(fx, w, h)}><text x="${f(w / 2)}" y="${f(h / 2 + size / 2)}" text-anchor="middle" font-family="Material Symbols Outlined" font-size="${f(size)}" fill="${fill}">${name}</text></g></svg>`;
  }
  if (shape.startsWith('diagram:')) return diagramSvg(shape.slice(8), d.text || '', w, h, Array.isArray(fx?.colors) ? fx!.colors : SHAPE_PALETTE, safeColor(d.font?.color, '#ffffff'));
  if (shape === 'slicer') return slicerSvg(d, w, h);
  if (shape === 'embed') {
    const name = String(d.text || 'File');
    const ext = (/\.([A-Za-z0-9]{1,6})$/.exec(name)?.[1] || 'FILE').toUpperCase();
    const s = Math.min(w, h - 22);
    const x0 = (w - s * 0.78) / 2;
    return `${svgOpen(w, h)}<path d="M${f(x0)},2 H${f(x0 + s * 0.55)} L${f(x0 + s * 0.78)},${f(2 + s * 0.22)} V${f(s)} H${f(x0)} Z" fill="#ffffff" stroke="#57606a" stroke-width="1.5"/><path d="M${f(x0 + s * 0.55)},2 V${f(2 + s * 0.22)} H${f(x0 + s * 0.78)}" fill="none" stroke="#57606a" stroke-width="1.5"/><rect x="${f(x0 + 4)}" y="${f(s * 0.55)}" width="${f(s * 0.78 - 8)}" height="${f(s * 0.22)}" fill="${safeColor(d.fill, '#2f6fb0')}"/><text x="${f(w / 2)}" y="${f(s * 0.55 + s * 0.16)}" text-anchor="middle" font-family="Carlito, Arial, sans-serif" font-size="${f(Math.max(8, s * 0.13))}" font-weight="bold" fill="#ffffff">${escHtml(ext)}</text>${textBlock(name, { x: 0, y: s + 2, w, h: h - s - 2, size: 11, fill: '#1a1a1a', valign: 'top' })}</svg>`;
  }
  const isText = d.kind === 'textbox';
  const sw = safeNum(d.strokeWidth ?? (isText ? 0.75 : 1.5), 1.5, 0, 50);
  const m = Math.max(1, sw / 2 + 0.5);
  const geo = shapeGeometry(shape, w, h, m);
  const [fdef, fill] = paint(geo.open ? 'none' : d.fill ?? (isText ? '#ffffff' : '#4a90c2'), isText ? '#ffffff' : '#4a90c2', `g${id}`);
  const stroke = d.stroke === 'none' ? 'none' : safeColor(d.stroke, isText ? '#7f7f7f' : geo.open ? '#2f6fb0' : '#1f3b57');
  const dash = fx?.dash === 'dash' ? ` stroke-dasharray="${f(sw * 4)} ${f(sw * 3)}"` : fx?.dash === 'dot' ? ` stroke-dasharray="${f(sw)} ${f(sw * 2)}"` : '';
  const [filt, fattr] = effectFilter(fx, `e${id}`);
  const marker = geo.extra ? `<marker id="MK${id}" viewBox="0 0 10 10" markerWidth="${f(Math.max(4, 8 / Math.max(1, sw / 2)))}" markerHeight="${f(Math.max(4, 8 / Math.max(1, sw / 2)))}" refX="8" refY="5" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="${stroke === 'none' ? '#2f6fb0' : stroke}"/></marker>` : '';
  const body = geo.body.replace(/url\(#MK\)/g, `url(#MK${id})`);
  const tr = geo.text || { x: m, y: m, w: w - 2 * m, h: h - 2 * m };
  const font = d.font || {};
  const pad = isText ? 6 : 4;
  const align = font.hAlign === 'left' || font.hAlign === 'right' || font.hAlign === 'center' ? font.hAlign : isText ? 'left' : 'center';
  const valign = font.vAlign === 'top' || font.vAlign === 'bottom' || font.vAlign === 'middle' ? font.vAlign : isText ? 'top' : 'middle';
  const textOutline = fx?.textOutline && fx.textOutline !== 'none' ? safeColor(fx.textOutline, '#000000') : undefined;
  let textFill: string | undefined;
  let tdef = '';
  if (fx?.textFill) { const p = paint(fx.textFill, '#000000', `t${id}`); tdef = p[0]; textFill = p[1]; }
  const text = d.text ? textBlock(d.text, { x: tr.x + pad, y: tr.y + pad / 2, w: Math.max(4, tr.w - 2 * pad), h: Math.max(4, tr.h - pad), align, valign, font, color: font.color || (isText ? '#000000' : '#ffffff'), fill: textFill, stroke: textOutline, strokeWidth: 0.8 }) : '';
  return `${svgOpen(w, h)}<defs>${fdef}${tdef}${filt}${marker}</defs><g${fattr}><g${flipTransform(fx, w, h)} fill="${fill}" stroke="${stroke}" stroke-width="${f(sw)}" stroke-linejoin="round" color="${stroke === 'none' ? '#1f3b57' : stroke}"${dash}>${body}</g></g>${text}</svg>`;
}

/** Small preview of a gallery shape (menus and pickers). */
export function shapePreviewSvg(id: string, size = 28): string {
  const geo = shapeGeometry(id, size, size, 2);
  return `${svgOpen(size, size)}<g fill="${geo.open ? 'none' : '#cfe0f3'}" stroke="#2f6fb0" stroke-width="1.4" stroke-linejoin="round" color="#2f6fb0">${geo.body.replace(/ marker-(end|start)="url\(#MK\)"/g, '')}</g></svg>`;
}
