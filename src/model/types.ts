/** Core value and workbook types. */
import type { Range } from './address';

export type ErrorCode = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#SPILL!' | '#CALC!' | '#GETTING_DATA' | '#CIRC!';

export class CellError {
  constructor(public code: ErrorCode, public detail = '') {}
  toString(): string { return this.code; }
}

export const ERR = {
  null: new CellError('#NULL!'),
  div0: new CellError('#DIV/0!'),
  value: new CellError('#VALUE!'),
  ref: new CellError('#REF!'),
  name: new CellError('#NAME?'),
  num: new CellError('#NUM!'),
  na: new CellError('#N/A'),
  spill: new CellError('#SPILL!'),
  calc: new CellError('#CALC!'),
  circ: new CellError('#CIRC!'),
} as const;

export const ERROR_CODES: ErrorCode[] = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA', '#CIRC!'];
export function errorFromCode(code: string): CellError | null {
  const c = code.toUpperCase();
  return (ERROR_CODES as string[]).includes(c) ? new CellError(c as ErrorCode) : null;
}

export type Scalar = number | string | boolean | null | CellError;
export type Matrix = Scalar[][];

export const isError = (v: unknown): v is CellError => v instanceof CellError;

export interface Border { style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair' | 'mediumDashed' | 'dashDot'; color?: string }

/** Cell formatting. Every property is optional; absent means the workbook default. */
export interface CellStyle {
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  strike?: boolean;
  vertAlign?: 'superscript' | 'subscript';
  color?: string;
  fill?: string;
  pattern?: { type: string; color?: string };
  gradient?: { from: string; to: string; angle?: number };
  numFmt?: string;
  hAlign?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vAlign?: 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
  wrap?: boolean;
  shrink?: boolean;
  indent?: number;
  rotate?: number; // -90..90, or 255 for vertical stacked
  bt?: Border; bb?: Border; bl?: Border; br?: Border;
  diagUp?: Border; diagDown?: Border;
  locked?: boolean;   // default true
  hidden?: boolean;   // hide formula when protected
  rtl?: boolean;
  /** In-cell checkbox: TRUE/FALSE values are drawn as a checked/unchecked box. */
  checkbox?: boolean;
}

export interface Cell {
  /** Value: a constant, or the cached result of the formula. */
  v?: Scalar;
  /** Formula text without the leading '='. */
  f?: string;
  /** Style index into Workbook.styles. 0 = default. */
  s?: number;
  /** For spilled cells: key of the anchor cell whose array result fills this cell. */
  sp?: number;
  /** For array (CSE) formula anchors: the range the legacy array formula occupies. */
  arr?: Range;
  /** Rich text runs (for constants), each a slice of the string value with its own font overrides. */
  rich?: { t: string; s?: Partial<CellStyle> }[];
  link?: string;
}

export interface CommentReply { author: string; text: string; time: number }
export interface CellComment { author: string; text: string; time: number; resolved?: boolean; replies?: CommentReply[]; /** legacy note */ note?: boolean; w?: number; h?: number }

export interface DataValidation {
  id: string;
  ranges: Range[];
  type: 'any' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
  operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropdown?: boolean;
  inputTitle?: string; inputMessage?: string;
  errorStyle?: 'stop' | 'warning' | 'information';
  errorTitle?: string; errorMessage?: string;
}

export type CFType = 'cellIs' | 'expression' | 'colorScale' | 'dataBar' | 'iconSet' | 'top10' | 'aboveAverage' | 'duplicateValues' | 'uniqueValues' | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith' | 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors' | 'timePeriod';

export interface CFRule {
  id: string;
  ranges: Range[];
  type: CFType;
  operator?: DataValidation['operator'];
  formulas?: string[];
  text?: string;
  rank?: number; percent?: boolean; bottom?: boolean;
  above?: boolean; equalAverage?: boolean; stdDev?: number;
  timePeriod?: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth';
  style?: Partial<CellStyle>;
  colors?: string[]; // color scale: 2 or 3 stops
  stops?: { type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula'; value?: number | string }[];
  barColor?: string; gradientBar?: boolean; showValue?: boolean;
  iconSet?: string; reverseIcons?: boolean;
  stopIfTrue?: boolean;
  priority?: number;
}

export interface AutoFilter {
  range: Range;
  /** Column offset within the range -> criteria. */
  columns: Record<number, FilterCriteria>;
  sort?: { col: number; desc: boolean };
}

export interface FilterCriteria {
  values?: string[];          // display values to keep (checkbox list)
  blanks?: boolean;
  custom?: { op: string; value: string }[]; and?: boolean;
  top10?: { top: boolean; percent: boolean; n: number };
  color?: { fill?: string; font?: string };
  dynamic?: 'aboveAverage' | 'belowAverage';
}

export interface TableStyleOpts { header: boolean; totals: boolean; banded: boolean; bandedCols: boolean; firstCol: boolean; lastCol: boolean; filterButton: boolean }

export interface TableDef {
  id: string;
  name: string;
  range: Range;          // includes header and totals rows
  columns: { name: string; totalFn?: string; totalFormula?: string; calcFormula?: string }[];
  style: string;
  opts: TableStyleOpts;
  /** Header-button filter criteria (column offset within the table -> criteria) and the last sort. */
  filter?: { columns: Record<number, FilterCriteria>; sort?: { col: number; desc: boolean } };
}

export interface ChartSeries { name?: string; nameRef?: string; values: string; categories?: string; color?: string; type?: ChartType; secondary?: boolean; /** Bubble sizes (bubble charts). */ sizes?: string }
export type ChartType = 'column' | 'stackedColumn' | 'percentColumn' | 'bar' | 'stackedBar' | 'percentBar' | 'line' | 'lineMarkers' | 'stackedLine' | 'area' | 'stackedArea' | 'percentArea' | 'pie' | 'doughnut' | 'scatter' | 'scatterLines' | 'bubble' | 'radar' | 'filledRadar' | 'histogram' | 'pareto' | 'boxWhisker' | 'waterfall' | 'funnel' | 'treemap' | 'sunburst' | 'stock' | 'combo';

export interface ChartDef {
  id: string;
  type: ChartType;
  title?: string;
  series: ChartSeries[];
  /** Anchor: top-left cell plus offset, size in px at 100% zoom. */
  at: { r: number; c: number; dx: number; dy: number; w: number; h: number };
  legend?: 'right' | 'bottom' | 'top' | 'left' | 'none';
  dataLabels?: boolean;
  xTitle?: string; yTitle?: string;
  gridlines?: boolean;
  style?: number;
  colors?: string[];
  sourceRange?: string;
  seriesInRows?: boolean;
  trendline?: 'linear' | 'exponential' | 'polynomial' | 'movingAverage' | null;
  yMin?: number; yMax?: number; logScale?: boolean;
  /** Alternative text (accessibility). */
  alt?: string;
}

export interface DrawingObject {
  id: string;
  kind: 'image' | 'shape' | 'textbox' | 'wordart' | 'icon' | 'ink' | 'signature';
  at: { r: number; c: number; dx: number; dy: number; w: number; h: number };
  src?: string;          // image data URL
  shape?: string;        // shape name
  text?: string;
  fill?: string; stroke?: string; strokeWidth?: number;
  rotate?: number;
  alt?: string;
  font?: Partial<CellStyle>;
  svg?: string;
  lockAspect?: boolean;
  /** Original picture data before effects/crop were applied (Reset Picture). */
  orig?: string;
  /** Effects and object settings: shadow, glow, flips, picture adjustments, slicer state, diagram colors... */
  fx?: Record<string, any>;
}

export interface Sparkline { id: string; cell: { r: number; c: number }; data: string; type: 'line' | 'column' | 'winloss'; color?: string; markers?: boolean; high?: boolean; low?: boolean; negative?: boolean }

export interface PrintSettings {
  orientation: 'portrait' | 'landscape';
  paper: string;  // 'Letter', 'A4', ...
  margins: { top: number; bottom: number; left: number; right: number; header: number; footer: number }; // inches
  scale: number;  // percent
  fitWide?: number; fitTall?: number;
  area?: string;
  titleRows?: string; titleCols?: string;
  gridlines: boolean;
  headings: boolean;
  centerH: boolean; centerV: boolean;
  header: { left: string; center: string; right: string };
  footer: { left: string; center: string; right: string };
  order: 'downThenOver' | 'overThenDown';
  blackAndWhite?: boolean;
  draft?: boolean;
  firstPage?: number;
  /** How comments and notes print (default none). */
  comments?: 'none' | 'atEnd' | 'asDisplayed';
  /** How cell errors print (default as displayed). */
  errors?: 'displayed' | 'blank' | 'dashes' | 'na';
  /** Header/footer options: a different first page, different odd and even pages, and their texts. */
  differentFirst?: boolean; differentOddEven?: boolean;
  firstHeader?: { left: string; center: string; right: string }; firstFooter?: { left: string; center: string; right: string };
  evenHeader?: { left: string; center: string; right: string }; evenFooter?: { left: string; center: string; right: string };
  scaleWithDoc?: boolean; alignWithMargins?: boolean;
}

/** A range users may edit on a protected sheet (Allow Edit Ranges), optionally with its own password. */
export interface AllowEditRange { title: string; ref: string; passwordHash?: string }

export interface SheetProtection { on: boolean; passwordHash?: string; allow: string[] }

export interface Scenario { name: string; cells: string; values: Scalar[]; comment?: string }

export interface PivotDef {
  id: string;
  source: string;  // "Sheet1!A1:F200" or table name
  target: { sheet: string; r: number; c: number };
  rows: string[]; cols: string[]; filters: string[];
  values: { field: string; fn: 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevP' | 'var' | 'varP'; showAs?: 'normal' | 'percentOfTotal' | 'percentOfRow' | 'percentOfCol' | 'runningTotal' | 'rank'; /** Custom caption. */ name?: string; numFmt?: string }[];
  filterValues?: Record<string, string[]>;
  layout: 'compact' | 'outline' | 'tabular';
  grandRows: boolean; grandCols: boolean; subtotals: boolean;
  style?: string;
  lastRange?: Range;
  /** PivotTable name ("PivotTable1"). */
  name?: string;
  /** Date grouping per field (row and column fields holding date serials). */
  dateGroups?: Record<string, ('years' | 'quarters' | 'months' | 'days')[]>;
  /** Label sort order per field (ascending by default). */
  sort?: Record<string, 'asc' | 'desc'>;
}

export interface DefinedName { name: string; ref: string; sheet?: string | null; comment?: string; hidden?: boolean }

export interface NamedCellStyle { name: string; style: CellStyle; builtin?: boolean }
