/**
 * Default tab and group arrangement. Groups reference command ids only, so users can rebuild any of it
 * with Customize Tabs. Contextual tabs appear when their object (table, chart, picture, shape,
 * PivotTable) is selected.
 */
import type { App } from '../app';

export interface RibbonGroup { id: string; label: string; items: string[]; launcher?: string }
export interface RibbonTab { id: string; label: string; groups: RibbonGroup[]; contextual?: (app: App) => boolean; accent?: string; custom?: boolean }

const inTable = (app: App) => { const s = app.sheet; const { r, c } = s.sel; return s.tables.some((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2); };
const inPivot = (app: App) => { const s = app.sheet; const { r, c } = s.sel; return s.pivots.some((p) => p.lastRange && r >= p.lastRange.r1 && r <= p.lastRange.r2 && c >= p.lastRange.c1 && c <= p.lastRange.c2); };
const selObj = (app: App) => app.tab?.grid?.objects.selectedObject();

export const DEFAULT_TABS: RibbonTab[] = [
  {
    id: 'home', label: 'Home', groups: [
      { id: 'clipboard', label: 'Clipboard', items: ['edit.paste', 'edit.cut', 'edit.copy', 'edit.painter'], launcher: 'view.clipboardPane' },
      { id: 'font', label: 'Font', items: ['fmt.font', 'fmt.size', 'fmt.grow', 'fmt.shrink', 'fmt.bold', 'fmt.italic', 'fmt.underline', 'fmt.strike', 'fmt.borders', 'fmt.fill', 'fmt.color'], launcher: 'fmt.fontDialog' },
      { id: 'alignment', label: 'Alignment', items: ['fmt.alignTop', 'fmt.alignMiddle', 'fmt.alignBottom', 'fmt.orientation', 'fmt.alignLeft', 'fmt.alignCenter', 'fmt.alignRight', 'fmt.indentDec', 'fmt.indentInc', 'fmt.wrap', 'fmt.merge'], launcher: 'fmt.alignmentDialog' },
      { id: 'number', label: 'Number', items: ['fmt.numberFormat', 'fmt.accounting', 'fmt.percent', 'fmt.comma', 'fmt.decInc', 'fmt.decDec'], launcher: 'fmt.numberDialog' },
      { id: 'styles', label: 'Styles', items: ['cond.menu', 'table.formatAsTable', 'fmt.cellStyles'] },
      { id: 'cells', label: 'Cells', items: ['cells.insert', 'cells.delete', 'cells.format'] },
      { id: 'editing', label: 'Editing', items: ['edit.autoSum', 'edit.fill', 'edit.clear', 'data.sortFilterMenu', 'edit.findSelect'] },
      { id: 'analysis', label: 'Analysis', items: ['ai.analyzeData', 'ai.quickAnalysis'] },
    ],
  },
  {
    id: 'insert', label: 'Insert', groups: [
      { id: 'tables', label: 'Tables', items: ['insert.pivotTable', 'insert.recommendedPivot', 'insert.table'] },
      { id: 'illustrations', label: 'Illustrations', items: ['insert.picture', 'insert.shapes', 'insert.icons', 'insert.diagram', 'insert.screenshot'] },
      { id: 'charts', label: 'Charts', items: ['insert.recommendedCharts', 'insert.chartColumn', 'insert.chartLine', 'insert.chartPie', 'insert.chartBar', 'insert.chartArea', 'insert.chartScatter', 'insert.chartStatistic', 'insert.chartWaterfall', 'insert.chartHierarchy', 'insert.chartCombo', 'insert.chartRadar'], launcher: 'insert.chartDialog' },
      { id: 'sparklines', label: 'Sparklines', items: ['insert.sparkLine', 'insert.sparkColumn', 'insert.sparkWinLoss'] },
      { id: 'filters', label: 'Filters', items: ['insert.slicer', 'insert.timeline'] },
      { id: 'links', label: 'Links', items: ['insert.link'] },
      { id: 'comments', label: 'Comments', items: ['review.newComment', 'review.newNote'] },
      { id: 'text', label: 'Text', items: ['insert.textBox', 'insert.headerFooter', 'insert.wordArt', 'insert.signatureLine', 'insert.object'] },
      { id: 'symbols', label: 'Symbols', items: ['insert.equation', 'insert.symbol', 'insert.emoji', 'insert.checkbox'] },
    ],
  },
  {
    id: 'draw', label: 'Draw', groups: [
      { id: 'tools', label: 'Drawing Tools', items: ['draw.select', 'draw.lasso', 'draw.pen', 'draw.pencil', 'draw.highlighter', 'draw.eraser'] },
      { id: 'convert', label: 'Convert', items: ['draw.inkToShape', 'draw.inkToMath'] },
      { id: 'replay', label: 'Replay', items: ['draw.replay'] },
    ],
  },
  {
    id: 'pagelayout', label: 'Page Layout', groups: [
      { id: 'themes', label: 'Themes', items: ['layout.themes', 'layout.themeColors', 'layout.themeFonts', 'layout.themeEffects'] },
      { id: 'pageSetup', label: 'Page Setup', items: ['layout.margins', 'layout.orientation', 'layout.size', 'layout.printArea', 'layout.breaks', 'layout.background', 'layout.printTitles'], launcher: 'layout.pageSetup' },
      { id: 'scale', label: 'Scale to Fit', items: ['layout.fitWidth', 'layout.fitHeight', 'layout.scale'] },
      { id: 'sheetOptions', label: 'Sheet Options', items: ['view.gridlines', 'layout.printGridlines', 'view.headings', 'layout.printHeadings'] },
      { id: 'arrange', label: 'Arrange', items: ['object.bringForward', 'object.sendBackward', 'view.selectionPane', 'object.align', 'object.group', 'object.rotate'] },
    ],
  },
  {
    id: 'formulas', label: 'Formulas', groups: [
      { id: 'library', label: 'Function Library', items: ['formulas.insertFunction', 'edit.autoSum', 'formulas.recent', 'formulas.financial', 'formulas.logical', 'formulas.text', 'formulas.dateTime', 'formulas.lookup', 'formulas.mathTrig', 'formulas.more'] },
      { id: 'names', label: 'Defined Names', items: ['formulas.nameManager', 'formulas.defineName', 'formulas.useInFormula', 'formulas.createFromSelection'] },
      { id: 'auditing', label: 'Formula Auditing', items: ['formulas.tracePrecedents', 'formulas.traceDependents', 'formulas.removeArrows', 'view.showFormulas', 'formulas.errorChecking', 'formulas.evaluate', 'formulas.watchWindow'] },
      { id: 'calculation', label: 'Calculation', items: ['formulas.calcOptions', 'formulas.calcNow', 'formulas.calcSheet'] },
      { id: 'insight', label: 'Insight', items: ['formulas.explainValue', 'formulas.dependencyDiagram', 'formulas.formulaRules', 'formulas.selectInconsistent', 'formulas.infoWindow', 'formulas.cellHistory'] },
    ],
  },
  {
    id: 'data', label: 'Data', groups: [
      { id: 'getData', label: 'Get & Transform Data', items: ['data.getData', 'data.fromText', 'data.fromWeb', 'data.fromTable', 'data.fromClipboard'] },
      { id: 'queries', label: 'Queries & Connections', items: ['data.refreshAll', 'data.connections', 'data.editLinks'] },
      { id: 'types', label: 'Data Types', items: ['data.dataTypes'] },
      { id: 'sortFilter', label: 'Sort & Filter', items: ['data.sortAZ', 'data.sortZA', 'data.sort', 'data.filter', 'data.clearFilter', 'data.reapply', 'data.advancedFilter'] },
      { id: 'tools', label: 'Data Tools', items: ['data.textToColumns', 'data.flashFill', 'data.removeDuplicates', 'data.validation', 'data.consolidate', 'data.relationships'] },
      { id: 'records', label: 'Records', items: ['data.dataForm', 'data.columnType', 'data.categorize', 'data.queryDelete'] },
      { id: 'forecast', label: 'Forecast', items: ['data.whatIf', 'data.forecastSheet', 'data.whatIfSliders', 'data.monteCarlo'] },
      { id: 'outline', label: 'Outline', items: ['data.group', 'data.ungroup', 'data.subtotal', 'data.showDetail', 'data.hideDetail'], launcher: 'data.outlineSettings' },
      { id: 'analyze', label: 'Analyze', items: ['data.analysisTools', 'data.solver', 'data.matrix'] },
    ],
  },
  {
    id: 'review', label: 'Review', groups: [
      { id: 'proofing', label: 'Proofing', items: ['review.spelling', 'review.dictionaries', 'review.thesaurus', 'review.workbookStats'] },
      { id: 'accessibility', label: 'Accessibility', items: ['review.accessibility'] },
      { id: 'insights', label: 'Insights', items: ['review.smartLookup'] },
      { id: 'language', label: 'Language', items: ['review.translate'] },
      { id: 'comments', label: 'Comments', items: ['review.newComment', 'review.deleteComment', 'review.prevComment', 'review.nextComment', 'review.showComments'] },
      { id: 'notes', label: 'Notes', items: ['review.notes'] },
      { id: 'protect', label: 'Protect', items: ['review.protectSheet', 'review.protectWorkbook', 'review.allowEditRanges', 'review.markFinal'] },
      { id: 'ink', label: 'Ink', items: ['review.hideInk'] },
    ],
  },
  {
    id: 'view', label: 'View', groups: [
      { id: 'sheetView', label: 'Sheet View', items: ['view.sheetViewMenu'] },
      { id: 'views', label: 'Workbook Views', items: ['view.normal', 'view.pageBreak', 'view.pageLayout', 'view.customViews', 'view.chartSlideShow'] },
      { id: 'show', label: 'Show', items: ['view.formulaBar', 'view.gridlines', 'view.headings', 'view.ruler', 'view.mapView'] },
      { id: 'zoom', label: 'Zoom', items: ['view.zoom', 'view.zoom100', 'view.zoomSelection'] },
      { id: 'window', label: 'Window', items: ['view.newWindow', 'view.arrangeAll', 'view.freezePanes', 'view.split', 'view.hideWindow', 'view.unhideWindow', 'view.switchWindows'] },
      { id: 'appearance', label: 'Appearance', items: ['view.darkSheet', 'view.fullScreen', 'view.navigator'] },
      { id: 'macros', label: 'Macros', items: ['dev.macros'] },
    ],
  },
  {
    id: 'tools', label: 'Tools', groups: [
      { id: 'cleanup', label: 'Clean Up', items: ['tools.trimSpaces', 'tools.cleanText', 'tools.changeCase', 'tools.textToNumbers', 'tools.fixDates', 'tools.ocrCleanup', 'tools.fillBlanks'] },
      { id: 'transform', label: 'Transform', items: ['tools.splitColumn', 'tools.mergeColumns', 'tools.unpivot', 'tools.transposeInPlace', 'tools.reverse', 'tools.shuffle'] },
      { id: 'compare', label: 'Compare', items: ['tools.compareSheets', 'tools.findDuplicates', 'tools.fuzzyMatch'] },
      { id: 'analyze', label: 'Analyze', items: ['tools.dataProfile', 'tools.descriptiveStats', 'tools.histogram', 'tools.correlation'] },
      { id: 'generate', label: 'Generate', items: ['tools.randomData', 'tools.sequence', 'tools.calendar'] },
      { id: 'classic', label: 'Classic', items: ['tools.slashMenu', 'tools.slashKey', 'tools.lotusEntry', 'tools.lotusLabels', 'tools.convertLotus', 'tools.exportCode'] },
    ],
  },
  {
    id: 'ai', label: 'AI', groups: [
      { id: 'formulas', label: 'Formulas', items: ['ai.writeFormula', 'ai.explainFormula', 'ai.fixFormula'] },
      { id: 'data', label: 'Data', items: ['ai.analyzeData', 'ai.fillColumn', 'ai.classify', 'ai.extract', 'ai.cleanData'] },
      { id: 'generate', label: 'Generate', items: ['ai.generateTable', 'ai.suggestChart', 'ai.summarize'] },
      { id: 'text', label: 'Text', items: ['ai.translateCells', 'ai.rewriteCells', 'ai.sentiment'] },
      { id: 'assistant', label: 'Assistant', items: ['ai.chat', 'ai.prompts'] },
      { id: 'setup', label: 'Setup', items: ['ai.setup', 'ai.providers', 'ai.settings'] },
    ],
  },
  {
    id: 'developer', label: 'Developer', groups: [
      { id: 'code', label: 'Code', items: ['dev.macroEditor', 'dev.macros', 'dev.recordMacro', 'dev.relativeRefs', 'dev.macroSecurity'] },
      { id: 'controls', label: 'Controls', items: ['dev.insertControl', 'dev.designMode', 'dev.properties'] },
      { id: 'xml', label: 'Data', items: ['dev.exportJson', 'dev.importJson'] },
    ],
  },
  {
    id: 'help', label: 'Help', groups: [
      { id: 'help', label: 'Help', items: ['help.help', 'help.shortcuts', 'help.commandReference', 'help.functionReference', 'help.tour', 'help.feedback', 'help.about'] },
    ],
  },
  // Contextual tabs
  {
    id: 'tableDesign', label: 'Table Design', contextual: inTable, accent: '#2e7d32', groups: [
      { id: 'properties', label: 'Properties', items: ['table.name', 'table.resize'] },
      { id: 'tools', label: 'Tools', items: ['table.summarizePivot', 'table.removeDuplicates', 'table.convertToRange', 'insert.slicer'] },
      { id: 'options', label: 'Table Style Options', items: ['table.headerRow', 'table.totalRow', 'table.bandedRows', 'table.firstColumn', 'table.lastColumn', 'table.bandedColumns', 'table.filterButton'] },
      { id: 'styles', label: 'Table Styles', items: ['table.style'] },
    ],
  },
  {
    id: 'pivotAnalyze', label: 'PivotTable Analyze', contextual: inPivot, accent: '#2e7d32', groups: [
      { id: 'pivot', label: 'PivotTable', items: ['pivot.options', 'pivot.fields'] },
      { id: 'data', label: 'Data', items: ['pivot.refresh', 'pivot.changeSource'] },
      { id: 'actions', label: 'Actions', items: ['pivot.clear', 'pivot.move', 'pivot.delete'] },
      { id: 'layout', label: 'Layout', items: ['pivot.layout', 'pivot.subtotals', 'pivot.grandTotals'] },
    ],
  },
  {
    id: 'chartDesign', label: 'Chart Design', contextual: (app) => selObj(app)?.kind === 'chart', accent: '#7b3fa0', groups: [
      { id: 'layouts', label: 'Chart Layouts', items: ['chart.addElement', 'chart.quickLayout'] },
      { id: 'styles', label: 'Chart Styles', items: ['chart.colors', 'chart.style'] },
      { id: 'data', label: 'Data', items: ['chart.switchRowCol', 'chart.selectData'] },
      { id: 'type', label: 'Type', items: ['chart.changeType'] },
      { id: 'location', label: 'Location', items: ['chart.move', 'chart.saveImage'] },
    ],
  },
  {
    id: 'pictureFormat', label: 'Picture Format', contextual: (app) => { const o = selObj(app); return o?.kind === 'drawing' && o.def.kind === 'image'; }, accent: '#c06000', groups: [
      { id: 'adjust', label: 'Adjust', items: ['picture.removeBg', 'picture.corrections', 'picture.color', 'picture.effects', 'picture.change', 'picture.reset'] },
      { id: 'styles', label: 'Picture Styles', items: ['picture.border', 'picture.shadow'] },
      { id: 'accessibility', label: 'Accessibility', items: ['picture.altText'] },
      { id: 'arrange', label: 'Arrange', items: ['object.bringForward', 'object.sendBackward', 'object.align', 'object.rotate'] },
      { id: 'size', label: 'Size', items: ['picture.crop', 'picture.size'] },
    ],
  },
  {
    id: 'shapeFormat', label: 'Shape Format', contextual: (app) => { const o = selObj(app); return o?.kind === 'drawing' && o.def.kind !== 'image'; }, accent: '#c06000', groups: [
      { id: 'insert', label: 'Insert Shapes', items: ['insert.shapes', 'drawing.editText'] },
      { id: 'styles', label: 'Shape Styles', items: ['drawing.fill', 'drawing.outline', 'drawing.effects'] },
      { id: 'wordart', label: 'Decorative Text Styles', items: ['drawing.textFill', 'drawing.textOutline'] },
      { id: 'arrange', label: 'Arrange', items: ['object.bringForward', 'object.sendBackward', 'object.align', 'object.group', 'object.rotate'] },
      { id: 'size', label: 'Size', items: ['drawing.size'] },
    ],
  },
];

export const BACKSTAGE_ITEMS = ['file.new', 'file.open', 'file.save', 'file.saveAs', 'file.print', 'file.share', 'file.exportPdf', 'file.properties', 'file.versions', 'file.close', 'file.options'];

export function tabsFor(app: App): RibbonTab[] {
  const p = app.prefs;
  const order = p.tabOrder.length ? p.tabOrder : [];
  const builtIn = DEFAULT_TABS.filter((t) => !p.hiddenTabs.includes(t.id));
  const custom: RibbonTab[] = p.customTabs.filter((t) => !t.hidden).map((t) => ({ id: t.id, label: t.label, groups: t.groups, custom: true }));
  const all = [...builtIn, ...custom];
  if (!order.length) return all;
  return [...all].sort((a, b) => {
    const ia = order.indexOf(a.id), ib = order.indexOf(b.id);
    return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
  });
}

export function idsInTabs(): Set<string> {
  const s = new Set<string>();
  for (const t of DEFAULT_TABS) for (const g of t.groups) for (const id of g.items) s.add(id);
  return s;
}
