# Anne's Grid Command Reference

Anne's Grid ships **1139 commands**. Every one can be placed on the Quick Access Bar, added to a custom tab, bound to a keyboard shortcut, recorded in a macro, or found with command search (Alt+Q).

665 of them are "Commands Not in the Tabs" — reachable through the Quick Access Bar chooser, Customize Tabs, the command search, or menus — so nothing the app can do is hidden.

This file is generated from the source by `npm run catalog`.

## Default tab layout

- **Home**: Clipboard · Font · Alignment · Number · Styles · Cells · Editing · Analysis
- **Insert**: Tables · Illustrations · Charts · Sparklines · Filters · Links · Comments · Text · Symbols
- **Draw**: Drawing Tools · Convert · Replay
- **Page Layout**: Themes · Page Setup · Scale to Fit · Sheet Options · Arrange
- **Formulas**: Function Library · Defined Names · Formula Auditing · Calculation · Insight
- **Data**: Get & Transform Data · Queries & Connections · Data Types · Sort & Filter · Data Tools · Records · Forecast · Outline · Analyze
- **Review**: Proofing · Accessibility · Insights · Language · Comments · Notes · Protect · Ink
- **View**: Sheet View · Workbook Views · Show · Zoom · Window · Appearance · Macros
- **Tools**: Clean Up · Transform · Compare · Analyze · Generate · Classic
- **AI**: Formulas · Data · Generate · Text · Assistant · Setup · Providers
- **Developer**: Code · Controls · Data
- **Help**: Help
- **Utilities**: My Tools & Shortcuts · Display & Navigation · Select · Sheets · Range · Fill · Columns & Rows · Objects & Comments · Format & Layout · Numbers & Dates · Text · Formulas · Web Data · Info & List · File & System · Import · Export · Start · Options
- **Plus**: View · Range · Combine · Convert · Insert · Delete · Lookup · Filter · Print · Security · Workbook · Inquire · Query
- **Extras**: Genealogy · Publishing · Finance · Data Quality · Dates · Geography · Workflow
- **Table Design** (contextual): Properties · Tools · Table Style Options · Table Styles
- **PivotTable Analyze** (contextual): PivotTable · Data · Actions · Layout
- **Chart Design** (contextual): Chart Layouts · Chart Styles · Data · Type · Location
- **Picture Format** (contextual): Adjust · Picture Styles · Accessibility · Arrange · Size
- **Shape Format** (contextual): Insert Shapes · Shape Styles · Decorative Text Styles · Arrange · Size

## File (54)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Close All Workbooks | Close |  | `file.closeAll` | not in tabs |
| Close Workbook | Close | Ctrl+w, Ctrl+F4 | `file.close` | not in tabs |
| Exit | Close |  | `file.exit` | not in tabs |
| Export as Anne's Grid JSON (.json) | Export |  | `file.export.json` | not in tabs |
| Export as Anne's Grid Workbook (.qgw) | Export |  | `file.export.qgw` | not in tabs |
| Export as CSV (Comma delimited) (.csv) | Export |  | `file.export.csv` | not in tabs |
| Export as Encrypted Anne's Grid Workbook (.qgwe) | Export |  | `file.export.qgwe` | not in tabs |
| Export as Excel Macro-Enabled Workbook (.xlsm) | Export |  | `file.export.xlsm` | not in tabs |
| Export as Excel Template (.xltx) | Export |  | `file.export.xltx` | not in tabs |
| Export as Excel Workbook (.xlsx) | Export |  | `file.export.xlsx` | not in tabs |
| Export as Flat OpenDocument Spreadsheet (.fods) | Export |  | `file.export.fods` | not in tabs |
| Export as Markdown table (.md) | Export |  | `file.export.md` | not in tabs |
| Export as OpenDocument Spreadsheet (.ods) | Export |  | `file.export.ods` | not in tabs |
| Export as PDF | Export |  | `file.exportPdf` | not in tabs |
| Export as Tab-separated values (.tsv) | Export |  | `file.export.tsv` | not in tabs |
| Export as Text (Tab delimited) (.txt) | Export |  | `file.export.txt` | not in tabs |
| Export as Web Page (.html) | Export |  | `file.export.html` | not in tabs |
| Encrypt with Password… | Info |  | `file.encrypt` | not in tabs |
| Inspect Workbook… | Info |  | `file.inspect` | not in tabs |
| Properties… | Info |  | `file.properties` | not in tabs |
| Recover Unsaved Workbooks | Info |  | `file.recover` | not in tabs |
| Revert to Saved | Info |  | `file.revert` | not in tabs |
| Take Snapshot | Info | Ctrl+Alt+5 | `file.snapshot` | not in tabs |
| Version History | Info |  | `file.versions` | not in tabs; inspired by Snapshots (Scrivener) / version history (online editors) |
| New from Template… | New |  | `file.newFromTemplate` | not in tabs |
| New Workbook | New | Ctrl+n | `file.new` | not in tabs |
| Clear Unpinned Recent Workbooks | Open |  | `file.clearRecent` | not in tabs |
| Manual Recovery Options | Open |  | `file.recoveryHelp` | not in tabs |
| Open and Repair… | Open |  | `file.openRepair` | not in tabs |
| Open Encrypted Workbook… | Open |  | `file.openEncrypted` | not in tabs |
| Open Recent | Open |  | `file.openRecent` | has options menu; not in tabs |
| Open… | Open | Ctrl+o, Ctrl+F12 | `file.open` | not in tabs |
| Recover Data from Any File… | Open |  | `file.recoverText` | not in tabs |
| Recovery Report | Open |  | `file.recoveryReport` | not in tabs |
| AutoCorrect Options… | Options |  | `file.autocorrectOptions` | not in tabs |
| Custom Dictionary… | Options |  | `file.customDictionary` | not in tabs |
| Custom Lists… | Options |  | `file.customLists` | not in tabs |
| Customize Keyboard… | Options |  | `file.customizeKeyboard` | not in tabs |
| Customize Quick Access Bar… | Options |  | `file.customizeQat` | not in tabs |
| Customize Tabs… | Options |  | `file.customizeRibbon` | not in tabs |
| Export / Import Settings… | Options |  | `file.exportSettings` | not in tabs |
| Options… | Options |  | `file.options` | not in tabs |
| Reset All Customizations | Options |  | `file.resetCustomizations` | not in tabs |
| Print Preview | Print | Ctrl+F2 | `file.printPreview` | not in tabs |
| Print… | Print | Ctrl+p | `file.print` | not in tabs |
| Quick Print | Print |  | `file.quickPrint` | not in tabs |
| AutoSave | Save |  | `file.autosave` | not in tabs |
| Save | Save | Ctrl+s, Shift+F12 | `file.save` | not in tabs |
| Save All | Save |  | `file.saveAll` | not in tabs |
| Save as Template | Save |  | `file.saveTemplate` | not in tabs |
| Save As… | Save | F12 | `file.saveAs` | not in tabs |
| Copy Sheet as Markdown | Share |  | `file.copyMarkdown` | not in tabs |
| Email as Attachment | Share |  | `file.email` | not in tabs |
| Share… | Share |  | `file.share` | not in tabs |

## Home (179)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Align Left | Alignment |  | `fmt.alignLeft` |  |
| Align Right | Alignment |  | `fmt.alignRight` |  |
| Bottom Align | Alignment |  | `fmt.alignBottom` |  |
| Center | Alignment |  | `fmt.alignCenter` |  |
| Center Across Selection | Alignment |  | `fmt.alignCenterAcross` | not in tabs |
| Decrease Indent | Alignment | Ctrl+Alt+Shift+Tab | `fmt.indentDec` |  |
| Distributed | Alignment |  | `fmt.alignDistributed` | not in tabs |
| Fill (repeat) | Alignment |  | `fmt.alignFill` | not in tabs |
| Format Cells: Alignment… | Alignment |  | `fmt.alignmentDialog` | not in tabs |
| Increase Indent | Alignment | Ctrl+Alt+Tab | `fmt.indentInc` |  |
| Justify | Alignment |  | `fmt.alignJustify` | not in tabs |
| Merge & Center | Alignment |  | `fmt.merge` | has options menu |
| Merge Across | Alignment |  | `fmt.mergeAcross` | not in tabs |
| Merge Cells | Alignment |  | `fmt.mergeCells` | not in tabs |
| Middle Align | Alignment |  | `fmt.alignMiddle` |  |
| Orientation | Alignment |  | `fmt.orientation` | has options menu |
| Right-to-Left Text | Alignment |  | `fmt.rtl` | not in tabs |
| Shrink to Fit | Alignment |  | `fmt.shrinkToFit` | not in tabs |
| Top Align | Alignment |  | `fmt.alignTop` |  |
| Unmerge Cells | Alignment |  | `fmt.unmerge` | not in tabs |
| Wrap Text | Alignment |  | `fmt.wrap` |  |
| Quick Analysis | Analysis | Ctrl+q | `ai.quickAnalysis` |  |
| AutoFit Column Width | Cells |  | `format.autofitCols` | not in tabs |
| AutoFit Row Height | Cells |  | `format.autofitRows` | not in tabs |
| Column Width… | Cells |  | `format.colWidthDialog` | not in tabs |
| Default Width… | Cells |  | `format.defaultWidth` | not in tabs |
| Delete | Cells |  | `cells.delete` | has options menu |
| Delete Cells… | Cells |  | `cells.deleteDialog` | not in tabs |
| Delete Sheet Columns | Cells |  | `cells.deleteCols` | not in tabs |
| Delete Sheet Rows | Cells | Ctrl+- | `cells.deleteRows` | not in tabs |
| Format | Cells |  | `cells.format` | has options menu |
| Format Cells… | Cells | Ctrl+1 | `fmt.cellsDialog` | not in tabs |
| Hide Columns | Cells | Ctrl+0 | `format.hideCols` | not in tabs |
| Hide Rows | Cells | Ctrl+9 | `format.hideRows` | not in tabs |
| Insert | Cells |  | `cells.insert` | has options menu |
| Insert Cells… | Cells |  | `cells.insertDialog` | not in tabs |
| Insert Sheet Columns | Cells |  | `cells.insertCols` | not in tabs |
| Insert Sheet Rows | Cells | Ctrl+Shift+= | `cells.insertRows` | not in tabs |
| Lock Cell | Cells |  | `format.lockCell` | not in tabs |
| Row Height… | Cells |  | `format.rowHeightDialog` | not in tabs |
| Unhide Columns | Cells | Ctrl+Shift+0, Ctrl+) | `format.unhideCols` | not in tabs |
| Unhide Rows | Cells | Ctrl+Shift+9, Ctrl+( | `format.unhideRows` | not in tabs |
| Clipboard Pane | Clipboard |  | `view.clipboardPane` | not in tabs |
| Copy | Clipboard | Ctrl+c, Ctrl+Insert | `edit.copy` |  |
| Copy as Picture | Clipboard |  | `edit.copyAsPicture` | not in tabs |
| Cut | Clipboard | Ctrl+x, Shift+Delete | `edit.cut` |  |
| Format Painter | Clipboard |  | `edit.painter` |  |
| Paste | Clipboard | Ctrl+v, Shift+Insert | `edit.paste` | has options menu |
| Paste Column Widths | Clipboard |  | `edit.pasteColumnWidths` | not in tabs |
| Paste Comments and Notes | Clipboard |  | `edit.pasteComments` | not in tabs |
| Paste Formatting | Clipboard |  | `edit.pasteFormats` | not in tabs |
| Paste Formulas | Clipboard |  | `edit.pasteFormulas` | not in tabs |
| Paste Formulas & Number Formatting | Clipboard |  | `edit.pasteFormulasNumFmt` | not in tabs |
| Paste Keeping Source Column Widths | Clipboard |  | `edit.pasteKeepWidths` | not in tabs |
| Paste Link | Clipboard |  | `edit.pasteLink` | not in tabs |
| Paste Special… | Clipboard | Ctrl+Alt+v | `edit.pasteSpecial` | not in tabs |
| Paste Transposed | Clipboard |  | `edit.pasteTranspose` | not in tabs |
| Paste Validation | Clipboard |  | `edit.pasteValidation` | not in tabs |
| Paste Values | Clipboard | Ctrl+Shift+v | `edit.pasteValues` | not in tabs |
| Paste Values & Number Formatting | Clipboard |  | `edit.pasteValuesNumFmt` | not in tabs |
| Paste without Borders | Clipboard |  | `edit.pasteNoBorders` | not in tabs |
| AutoSum | Editing | Alt+= | `edit.autoSum` | has options menu |
| Clear | Editing |  | `edit.clear` | has options menu |
| Clear All | Editing |  | `edit.clearAll` | not in tabs |
| Clear Comments and Notes | Editing |  | `edit.clearComments` | not in tabs |
| Clear Contents | Editing |  | `edit.clearContents` | not in tabs |
| Clear Formats | Editing |  | `edit.clearFormats` | not in tabs |
| Clear Hyperlinks | Editing |  | `edit.clearLinks` | not in tabs |
| Copy Formula from Cell Above | Editing | Ctrl+' | `edit.copyAbove` | not in tabs |
| Copy Value from Cell Above | Editing | Ctrl+Shift+" | `edit.copyValueAbove` | not in tabs |
| Edit Active Cell | Editing |  | `edit.editCell` | not in tabs |
| Fill | Editing |  | `edit.fill` | has options menu |
| Fill Down | Editing | Ctrl+d | `edit.fillDown` | not in tabs |
| Fill Left | Editing |  | `edit.fillLeft` | not in tabs |
| Fill Right | Editing | Ctrl+r | `edit.fillRight` | not in tabs |
| Fill Up | Editing |  | `edit.fillUp` | not in tabs |
| Find & Select | Editing |  | `edit.findSelect` | has options menu |
| Find Next | Editing | Shift+F4 | `edit.findNext` | not in tabs |
| Find… | Editing | Ctrl+f, Shift+F5 | `edit.find` | not in tabs |
| Go To Special… | Editing |  | `nav.gotoSpecial` | not in tabs |
| Go To… | Editing | Ctrl+g, F5 | `nav.goto` | not in tabs |
| Insert Current Date | Editing | Ctrl+; | `edit.insertDate` | not in tabs |
| Insert Current Time | Editing | Ctrl+Shift+; | `edit.insertTime` | not in tabs |
| Justify | Editing |  | `edit.fillJustify` | not in tabs |
| Pick from Drop-down List | Editing | Alt+ArrowDown | `edit.pickFromList` | not in tabs |
| Remove Hyperlinks | Editing |  | `edit.removeLinks` | not in tabs |
| Replace… | Editing | Ctrl+h | `edit.replace` | not in tabs |
| Select All | Editing | Ctrl+a, Ctrl+Shift+Space | `nav.selectAll` | not in tabs |
| Select Conditional Formatting | Editing |  | `nav.selectCF` | not in tabs |
| Select Constants | Editing |  | `nav.selectConstants` | not in tabs |
| Select Current Region | Editing | Ctrl+Shift+8, Ctrl+* | `nav.currentRegion` | not in tabs |
| Select Data Validation | Editing |  | `nav.selectValidation` | not in tabs |
| Select Entire Column | Editing | Ctrl+Space | `nav.selectColumn` | not in tabs |
| Select Entire Row | Editing | Shift+Space | `nav.selectRow` | not in tabs |
| Select Formulas | Editing |  | `nav.selectFormulas` | not in tabs |
| Select Notes | Editing |  | `nav.selectComments` | not in tabs |
| Select Objects | Editing |  | `nav.selectObjects` | not in tabs |
| Select Visible Cells | Editing | Alt+; | `nav.selectVisible` | not in tabs |
| Series… | Editing |  | `edit.fillSeries` | not in tabs |
| Sort & Filter | Editing |  | `data.sortFilterMenu` | has options menu |
| All Borders | Font |  | `fmt.border.all` | not in tabs |
| Bold | Font | Ctrl+b, Ctrl+2 | `fmt.bold` |  |
| Border Line Color | Font |  | `fmt.borderColor` | not in tabs |
| Border Line Style | Font |  | `fmt.borderStyle` | not in tabs |
| Borders | Font |  | `fmt.borders` | has options menu |
| Bottom Border | Font |  | `fmt.border.bottom` | not in tabs |
| Bottom Double Border | Font |  | `fmt.border.bottomDouble` | not in tabs |
| Clear Formats | Font |  | `fmt.clearFormats` | not in tabs |
| Decrease Font Size | Font | Ctrl+Shift+, | `fmt.shrink` |  |
| Double Underline | Font |  | `fmt.doubleUnderline` | not in tabs |
| Fill Color | Font |  | `fmt.fill` |  |
| Font | Font | Ctrl+Shift+f | `fmt.font` |  |
| Font Color | Font |  | `fmt.color` |  |
| Font Size | Font |  | `fmt.size` |  |
| Format Cells: Font… | Font | Ctrl+Shift+p | `fmt.fontDialog` | not in tabs |
| Increase Font Size | Font | Ctrl+Shift+. | `fmt.grow` |  |
| Italic | Font | Ctrl+i, Ctrl+3 | `fmt.italic` |  |
| Left Border | Font |  | `fmt.border.left` | not in tabs |
| No Border | Font | Ctrl+Shift+- | `fmt.border.none` | not in tabs |
| No Fill | Font |  | `fmt.noFill` | not in tabs |
| Outside Borders | Font | Ctrl+Shift+7 | `fmt.border.outside` | not in tabs |
| Right Border | Font |  | `fmt.border.right` | not in tabs |
| Strikethrough | Font | Ctrl+5 | `fmt.strike` |  |
| Subscript | Font |  | `fmt.subscript` | not in tabs |
| Superscript | Font |  | `fmt.superscript` | not in tabs |
| Thick Outside Borders | Font |  | `fmt.border.thickOutside` | not in tabs |
| Top Border | Font |  | `fmt.border.top` | not in tabs |
| Underline | Font | Ctrl+u, Ctrl+4 | `fmt.underline` | has options menu |
| Accounting Format | Number |  | `fmt.num.accounting` | not in tabs |
| Accounting Number Format | Number |  | `fmt.accounting` | has options menu |
| Comma Style | Number |  | `fmt.comma` |  |
| Currency Format | Number | Ctrl+Shift+4, Ctrl+$ | `fmt.num.currency` | not in tabs |
| Decrease Decimal | Number |  | `fmt.decDec` |  |
| Format Cells: Number… | Number |  | `fmt.numberDialog` | not in tabs |
| Fraction Format | Number |  | `fmt.num.fraction` | not in tabs |
| General Format | Number | Ctrl+Shift+`, Ctrl+Shift+~ | `fmt.num.general` | not in tabs |
| Increase Decimal | Number |  | `fmt.decInc` |  |
| ISO Date Format | Number |  | `fmt.num.isoDate` | not in tabs |
| Long Date Format | Number |  | `fmt.num.longDate` | not in tabs |
| Number Format | Number |  | `fmt.numberFormat` |  |
| Number Format | Number |  | `fmt.num.number` | not in tabs |
| Number with Separator Format | Number | Ctrl+Shift+1, Ctrl+! | `fmt.num.numberComma` | not in tabs |
| Percent Style | Number |  | `fmt.percent` |  |
| Percentage Format | Number | Ctrl+Shift+5, Ctrl+% | `fmt.num.percentage` | not in tabs |
| Scientific Format | Number | Ctrl+Shift+6, Ctrl+^ | `fmt.num.scientific` | not in tabs |
| Short Date Format | Number | Ctrl+Shift+3, Ctrl+# | `fmt.num.shortDate` | not in tabs |
| Text Format | Number |  | `fmt.num.text` | not in tabs |
| Time Format | Number | Ctrl+Shift+2, Ctrl+@ | `fmt.num.time` | not in tabs |
| Above Average | Styles |  | `cond.aboveAverage` | not in tabs |
| Below Average | Styles |  | `cond.belowAverage` | not in tabs |
| Bottom 10 Items | Styles |  | `cond.bottom10` | not in tabs |
| Bottom 10% | Styles |  | `cond.bottom10Percent` | not in tabs |
| Cell Styles | Styles |  | `fmt.cellStyles` | has options menu |
| Clear Rules from Entire Sheet | Styles |  | `cond.clearSheet` | not in tabs |
| Clear Rules from Selected Cells | Styles |  | `cond.clearSelection` | not in tabs |
| Clear Rules from This PivotTable | Styles |  | `cond.clearPivot` | not in tabs |
| Clear Rules from This Table | Styles |  | `cond.clearTable` | not in tabs |
| Color Scales | Styles |  | `cond.colorScale` | has options menu; not in tabs |
| Conditional Formatting | Styles |  | `cond.menu` | has options menu |
| Data Bars | Styles |  | `cond.dataBar` | has options menu; not in tabs |
| Format as Table | Styles |  | `table.formatAsTable` | has options menu |
| Highlight a Date Occurring | Styles |  | `cond.dateOccurring` | not in tabs |
| Highlight Between | Styles |  | `cond.between` | not in tabs |
| Highlight Duplicate Values | Styles |  | `cond.duplicateValues` | not in tabs |
| Highlight Equal To | Styles |  | `cond.equalTo` | not in tabs |
| Highlight Greater Than | Styles |  | `cond.greaterThan` | not in tabs |
| Highlight Less Than | Styles |  | `cond.lessThan` | not in tabs |
| Highlight Text that Contains | Styles |  | `cond.textContains` | not in tabs |
| Icon Sets | Styles |  | `cond.iconSet` | has options menu; not in tabs |
| Manage Rules | Styles |  | `cond.manage` | not in tabs |
| New Cell Style… | Styles |  | `fmt.newCellStyle` | not in tabs |
| New Formatting Rule | Styles |  | `cond.newRule` | not in tabs |
| New Table Style | Styles |  | `table.newStyle` | not in tabs |
| Style… | Styles | Alt+' | `fmt.styleDialog` | not in tabs |
| Top 10 Items | Styles |  | `cond.top10` | not in tabs |
| Top 10% | Styles |  | `cond.top10Percent` | not in tabs |
| Redo | Undo | Ctrl+y, Ctrl+Shift+z | `edit.redo` | not in tabs |
| Repeat | Undo | F4 | `edit.repeat` | not in tabs |
| Undo | Undo | Ctrl+z | `edit.undo` | not in tabs |

## Insert (42)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Area Chart | Charts |  | `insert.chartArea` | has options menu |
| Bar Chart | Charts |  | `insert.chartBar` | has options menu |
| Column Chart | Charts |  | `insert.chartColumn` | has options menu |
| Combo Chart | Charts |  | `insert.chartCombo` | has options menu |
| Hierarchy Chart | Charts |  | `insert.chartHierarchy` | has options menu |
| Insert Chart | Charts |  | `insert.chartDialog` | not in tabs |
| Line Chart | Charts |  | `insert.chartLine` | has options menu |
| Pie Chart | Charts |  | `insert.chartPie` | has options menu |
| Radar Chart | Charts |  | `insert.chartRadar` | has options menu |
| Recommended Charts | Charts |  | `insert.recommendedCharts` |  |
| Scatter Chart | Charts |  | `insert.chartScatter` | has options menu |
| Statistic Chart | Charts |  | `insert.chartStatistic` | has options menu |
| Waterfall Chart | Charts |  | `insert.chartWaterfall` | has options menu |
| Clear Slicer Filter | Filters |  | `insert.slicerClear` | not in tabs |
| Slicer | Filters |  | `insert.slicer` |  |
| Timeline | Filters |  | `insert.timeline` |  |
| Diagram | Illustrations |  | `insert.diagram` |  |
| Icons | Illustrations |  | `insert.icons` |  |
| Pictures | Illustrations |  | `insert.picture` | has options menu |
| Screenshot | Illustrations |  | `insert.screenshot` |  |
| Shapes | Illustrations |  | `insert.shapes` | has options menu |
| Copy Link | Links |  | `insert.copyLink` | not in tabs |
| Link | Links | Ctrl+k | `insert.link` | has options menu |
| Open Link | Links |  | `insert.openLink` | not in tabs |
| Remove Link | Links |  | `insert.removeLink` | not in tabs |
| Column Sparkline | Sparklines |  | `insert.sparkColumn` |  |
| Line Sparkline | Sparklines |  | `insert.sparkLine` |  |
| Win/Loss Sparkline | Sparklines |  | `insert.sparkWinLoss` |  |
| Checkbox | Symbols |  | `insert.checkbox` | has options menu |
| Emoji | Symbols | Ctrl+. | `insert.emoji` |  |
| Equation | Symbols |  | `insert.equation` |  |
| Symbol | Symbols |  | `insert.symbol` | has options menu |
| PivotTable | Tables |  | `insert.pivotTable` |  |
| PivotTable | Tables |  | `pivot.create` | not in tabs |
| Recommended PivotTables | Tables |  | `insert.recommendedPivot` |  |
| Recommended PivotTables | Tables |  | `pivot.recommended` | not in tabs |
| Table | Tables | Ctrl+t, Ctrl+l | `insert.table` |  |
| Decorative Text | Text |  | `insert.wordArt` | has options menu |
| Header & Footer | Text |  | `insert.headerFooter` |  |
| Object | Text |  | `insert.object` |  |
| Signature Line | Text |  | `insert.signatureLine` |  |
| Text Box | Text |  | `insert.textBox` |  |

## Draw (11)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Convert Strokes to Shapes | Convert |  | `draw.autoShapes` | not in tabs |
| Ink to Math | Convert |  | `draw.inkToMath` |  |
| Ink to Shape | Convert |  | `draw.inkToShape` |  |
| Delete All Ink | Drawing Tools |  | `draw.deleteInk` | not in tabs |
| Eraser | Drawing Tools |  | `draw.eraser` |  |
| Highlighter | Drawing Tools |  | `draw.highlighter` | has options menu |
| Lasso Select | Drawing Tools |  | `draw.lasso` |  |
| Pen | Drawing Tools |  | `draw.pen` | has options menu |
| Pencil | Drawing Tools |  | `draw.pencil` | has options menu |
| Select | Drawing Tools |  | `draw.select` |  |
| Ink Replay | Replay |  | `draw.replay` |  |

## Page Layout (33)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Align | Arrange |  | `object.align` | has options menu |
| Bring Forward | Arrange |  | `object.bringForward` | has options menu |
| Bring to Front | Arrange |  | `object.bringToFront` | not in tabs |
| Delete Object | Arrange |  | `object.delete` | not in tabs |
| Flip Horizontal | Arrange |  | `object.flipH` | not in tabs |
| Flip Vertical | Arrange |  | `object.flipV` | not in tabs |
| Group | Arrange |  | `object.group` | has options menu |
| Rotate | Arrange |  | `object.rotate` | has options menu |
| Selection Pane | Arrange |  | `view.selectionPane` |  |
| Send Backward | Arrange |  | `object.sendBackward` | has options menu |
| Send to Back | Arrange |  | `object.sendToBack` | not in tabs |
| Background | Page Setup |  | `layout.background` |  |
| Breaks | Page Setup |  | `layout.breaks` | has options menu |
| Header & Footer… | Page Setup |  | `layout.headerFooter` | not in tabs |
| Margins | Page Setup |  | `layout.margins` | has options menu |
| Orientation | Page Setup |  | `layout.orientation` | has options menu |
| Page Setup… | Page Setup |  | `layout.pageSetup` | not in tabs |
| Print Area | Page Setup |  | `layout.printArea` | has options menu |
| Print Titles | Page Setup |  | `layout.printTitles` |  |
| Size | Page Setup |  | `layout.size` | has options menu |
| Height | Scale to Fit |  | `layout.fitHeight` |  |
| Scale | Scale to Fit |  | `layout.scale` |  |
| Scaling | Scale to Fit |  | `layout.scaling` | has options menu; not in tabs |
| Width | Scale to Fit |  | `layout.fitWidth` |  |
| Center on Page | Sheet Options |  | `layout.centerOnPage` | not in tabs |
| Print Gridlines | Sheet Options |  | `layout.printGridlines` |  |
| Print Headings | Sheet Options |  | `layout.printHeadings` |  |
| Print in Black and White | Sheet Options |  | `layout.blackAndWhite` | not in tabs |
| Sheet Right-to-Left | Sheet Options |  | `view.rightToLeft` | not in tabs |
| Colors | Themes |  | `layout.themeColors` | has options menu |
| Effects | Themes |  | `layout.themeEffects` | has options menu |
| Fonts | Themes |  | `layout.themeFonts` | has options menu |
| Themes | Themes |  | `layout.themes` | has options menu |

## Formulas (36)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Calculate Full | Calculation | Ctrl+Alt+F9 | `formulas.calcFull` | not in tabs |
| Calculate Now | Calculation | F9 | `formulas.calcNow` |  |
| Calculate Sheet | Calculation | Shift+F9 | `formulas.calcSheet` |  |
| Calculation Options | Calculation |  | `formulas.calcOptions` | has options menu |
| Create from Selection | Defined Names | Ctrl+Shift+F3 | `formulas.createFromSelection` |  |
| Define Name | Defined Names |  | `formulas.defineName` |  |
| Name Manager | Defined Names | Ctrl+F3 | `formulas.nameManager` |  |
| Paste Name | Defined Names | F3 | `formulas.pasteName` | not in tabs |
| Use in Formula | Defined Names |  | `formulas.useInFormula` | has options menu |
| Add Watch | Formula Auditing |  | `formulas.addWatch` | not in tabs |
| Circular References | Formula Auditing |  | `formulas.circularRefs` | has options menu; not in tabs |
| Error Checking | Formula Auditing |  | `formulas.errorChecking` | has options menu |
| Evaluate Formula | Formula Auditing |  | `formulas.evaluate` |  |
| Remove Arrows | Formula Auditing |  | `formulas.removeArrows` | has options menu |
| Show Formulas | Formula Auditing | Ctrl+` | `view.showFormulas` |  |
| Trace Dependents | Formula Auditing |  | `formulas.traceDependents` |  |
| Trace Error | Formula Auditing |  | `formulas.traceError` | not in tabs |
| Trace Precedents | Formula Auditing |  | `formulas.tracePrecedents` |  |
| Watch Window | Formula Auditing |  | `formulas.watchWindow` |  |
| Date & Time | Function Library |  | `formulas.dateTime` | has options menu |
| Financial | Function Library |  | `formulas.financial` | has options menu |
| Function Help | Function Library |  | `formulas.functionHelp` | not in tabs |
| Function Reference | Function Library |  | `formulas.functionBrowser` | not in tabs |
| Insert Function | Function Library | Shift+F3 | `formulas.insertFunction` |  |
| Logical | Function Library |  | `formulas.logical` | has options menu |
| Lookup & Reference | Function Library |  | `formulas.lookup` | has options menu |
| Math & Trig | Function Library |  | `formulas.mathTrig` | has options menu |
| More Functions | Function Library |  | `formulas.more` | has options menu |
| Recently Used | Function Library |  | `formulas.recent` | has options menu |
| Text | Function Library |  | `formulas.text` | has options menu |
| Cell History | Insight |  | `formulas.cellHistory` | inspired by Cell edit history in online spreadsheets |
| Dependency Diagram | Insight |  | `formulas.dependencyDiagram` | inspired by Javelin diagram view |
| Explain Value | Insight |  | `formulas.explainValue` | inspired by Javelin's model views; explainable models in Causal |
| Formula Rules | Insight |  | `formulas.formulaRules` | inspired by Lotus Improv formula pane; Multiplan R1C1 notation |
| Info Window | Insight |  | `formulas.infoWindow` | inspired by Excel for Macintosh Info window |
| Select Inconsistent Formulas | Insight |  | `formulas.selectInconsistent` | inspired by Multiplan R1C1 notation (copies of a formula read identically) |

## Data (56)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Data Analysis | Analyze |  | `data.analysisTools` |  |
| Matrix | Analyze |  | `data.matrix` | has options menu; inspired by Lotus 1-2-3 /Data Matrix |
| Solver | Analyze |  | `data.solver` |  |
| Circle Invalid Data | Data Tools |  | `data.circleInvalid` | not in tabs |
| Clear Validation Circles | Data Tools |  | `data.clearCircles` | not in tabs |
| Consolidate | Data Tools |  | `data.consolidate` |  |
| Data Validation | Data Tools |  | `data.validation` | has options menu |
| Flash Fill | Data Tools | Ctrl+E | `data.flashFill` |  |
| Pick From Drop-down List | Data Tools |  | `data.pickFromList` | not in tabs |
| Relationships | Data Tools |  | `data.relationships` |  |
| Remove Duplicates | Data Tools |  | `data.removeDuplicates` |  |
| Text to Columns | Data Tools |  | `data.textToColumns` |  |
| Convert Text to Dates | Data Types |  | `data.convertToDate` | not in tabs |
| Convert Text to Numbers | Data Types |  | `data.convertToNumber` | not in tabs |
| Convert to Text | Data Types |  | `data.convertToText` | not in tabs |
| Data Types | Data Types |  | `data.dataTypes` | has options menu |
| Data Table | Forecast |  | `data.dataTable` | not in tabs |
| Forecast Sheet | Forecast |  | `data.forecastSheet` |  |
| Goal Seek | Forecast |  | `data.goalSeek` | not in tabs |
| Monte Carlo Simulation | Forecast |  | `data.monteCarlo` | inspired by Uncertain inputs in Causal; spreadsheet risk-analysis add-ins |
| Scenario Manager | Forecast |  | `data.scenarios` | not in tabs |
| What-If Analysis | Forecast |  | `data.whatIf` | has options menu |
| What-If Sliders | Forecast |  | `data.whatIfSliders` | inspired by Live model controls in Causal and reactive notebooks |
| From Clipboard | Get & Transform Data |  | `data.fromClipboard` |  |
| From Table/Range | Get & Transform Data |  | `data.fromTable` |  |
| From Text/CSV | Get & Transform Data |  | `data.fromText` |  |
| From Web | Get & Transform Data |  | `data.fromWeb` |  |
| From Workbook | Get & Transform Data |  | `data.fromWorkbook` | not in tabs |
| Get Data | Get & Transform Data |  | `data.getData` | has options menu |
| Auto Outline | Outline |  | `data.autoOutline` | not in tabs |
| Clear Outline | Outline |  | `data.clearOutline` | not in tabs |
| Group | Outline | Alt+Shift+ArrowRight | `data.group` | has options menu |
| Hide Detail | Outline |  | `data.hideDetail` |  |
| Outline Settings | Outline |  | `data.outlineSettings` | not in tabs |
| Show Detail | Outline |  | `data.showDetail` |  |
| Show Outline Level | Outline |  | `data.outlineLevel` | has options menu; not in tabs |
| Subtotal | Outline |  | `data.subtotal` |  |
| Ungroup | Outline | Alt+Shift+ArrowLeft | `data.ungroup` | has options menu |
| Edit Links | Queries & Connections |  | `data.editLinks` |  |
| Queries & Connections | Queries & Connections |  | `data.connections` |  |
| Refresh All | Queries & Connections | Ctrl+Alt+F5 | `data.refreshAll` |  |
| Categorize by Rules | Records |  | `data.categorize` | inspired by Lotus Agenda categories |
| Column Type | Records |  | `data.columnType` | has options menu; inspired by Airtable and Grist field types |
| Data Form | Records |  | `data.dataForm` | inspired by Lotus Symphony FORM window; 1-2-3 /Range Input |
| Delete Matching Records | Records |  | `data.queryDelete` | inspired by Lotus 1-2-3 /Data Query Delete |
| Advanced Filter | Sort & Filter |  | `data.advancedFilter` |  |
| Clear Filter | Sort & Filter |  | `data.clearFilter` |  |
| Custom Sort | Sort & Filter |  | `data.sort` |  |
| Filter | Sort & Filter | Ctrl+Shift+L | `data.filter` |  |
| Filter by Selected Cell's Color | Sort & Filter |  | `data.filterByColor` | not in tabs |
| Filter by Selected Cell's Value | Sort & Filter |  | `data.filterBySelected` | not in tabs |
| Put Selected Cell Color on Top | Sort & Filter |  | `data.sortByColor` | not in tabs |
| Put Selected Font Color on Top | Sort & Filter |  | `data.sortByFontColor` | not in tabs |
| Reapply | Sort & Filter | Ctrl+Alt+L | `data.reapply` |  |
| Sort A to Z | Sort & Filter |  | `data.sortAZ` |  |
| Sort Z to A | Sort & Filter |  | `data.sortZA` |  |

## Review (32)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Check Accessibility | Accessibility |  | `review.accessibility` |  |
| Delete | Comments |  | `review.deleteComment` |  |
| Delete All Comments in Sheet | Comments |  | `review.deleteAllComments` | not in tabs |
| Edit Comment | Comments |  | `review.editComment` | not in tabs |
| New Comment | Comments | Ctrl+Alt+m | `review.newComment` |  |
| Next Comment | Comments |  | `review.nextComment` |  |
| Previous Comment | Comments |  | `review.prevComment` |  |
| Reply to Comment | Comments |  | `review.replyComment` | not in tabs |
| Resolve Thread | Comments |  | `review.resolveComment` | not in tabs |
| Show Comments | Comments |  | `review.showComments` |  |
| Hide Ink | Ink |  | `review.hideInk` |  |
| Smart Lookup | Insights |  | `review.smartLookup` |  |
| Spelling Language | Language |  | `review.language` | has options menu; not in tabs |
| Translate | Language |  | `review.translate` |  |
| Convert Notes to Comments | Notes |  | `review.convertNotes` | not in tabs |
| Delete Note | Notes |  | `review.deleteNote` | not in tabs |
| Edit Note | Notes |  | `review.editNote` | not in tabs |
| New Note | Notes | Shift+F2 | `review.newNote` |  |
| Next Note | Notes |  | `review.nextNote` | not in tabs |
| Notes | Notes |  | `review.notes` | has options menu |
| Previous Note | Notes |  | `review.prevNote` | not in tabs |
| Show All Notes | Notes |  | `review.showAllNotes` | not in tabs |
| Show/Hide Note | Notes |  | `review.showNote` | not in tabs |
| Spelling | Proofing | F7 | `review.spelling` |  |
| Spelling Dictionaries… | Proofing |  | `review.dictionaries` |  |
| Spelling Options… | Proofing |  | `review.spellingOptions` | not in tabs |
| Thesaurus | Proofing | Shift+F7 | `review.thesaurus` |  |
| Workbook Statistics | Proofing |  | `review.workbookStats` |  |
| Allow Edit Ranges | Protect |  | `review.allowEditRanges` |  |
| Mark as Final | Protect |  | `review.markFinal` |  |
| Protect Sheet | Protect |  | `review.protectSheet` |  |
| Protect Workbook | Protect |  | `review.protectWorkbook` |  |

## View (39)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Add All Quick Access Choices (A–Z) | Appearance |  | `view.qatAddAll` | not in tabs |
| Collapse the Ribbon | Appearance | Ctrl+F1 | `view.collapseRibbon` | not in tabs |
| Dark Cells | Appearance |  | `view.darkSheet` |  |
| Full Screen | Appearance | Ctrl+Shift+F1 | `view.fullScreen` |  |
| Navigator | Appearance |  | `view.navigator` | inspired by LibreOffice Calc Navigator |
| Quick Access Bar at the Bottom | Appearance |  | `view.qatAtBottom` | not in tabs |
| Search Commands | Appearance | Alt+q | `view.cmdSearch` | not in tabs |
| Show Quick Access Bar | Appearance |  | `view.showQat` | not in tabs |
| Touch/Mouse Mode | Appearance |  | `view.touchMode` | not in tabs |
| Sheet View | Sheet View |  | `view.sheetViewMenu` | has options menu |
| Formula Bar | Show |  | `view.formulaBar` |  |
| Gridline Color… | Show |  | `view.gridColor` | not in tabs |
| Gridlines | Show |  | `view.gridlines` |  |
| Headings | Show |  | `view.headings` |  |
| Map View | Show |  | `view.mapView` | inspired by Lotus 1-2-3 Release 3 map view |
| Ruler | Show |  | `view.ruler` |  |
| Show Zero Values | Show |  | `view.showZeros` | not in tabs |
| Arrange All | Window |  | `view.arrangeAll` |  |
| Freeze First Column | Window |  | `view.freezeFirstCol` | not in tabs |
| Freeze Panes | Window |  | `view.freezePanes` | has options menu |
| Freeze Top Row | Window |  | `view.freezeTopRow` | not in tabs |
| Hide Window | Window |  | `view.hideWindow` |  |
| New Window | Window |  | `view.newWindow` |  |
| Next Workbook | Window | Ctrl+F6 | `view.nextWindow` | not in tabs |
| Split | Window |  | `view.split` |  |
| Switch Windows | Window |  | `view.switchWindows` | has options menu |
| Unfreeze Panes | Window |  | `view.unfreeze` | not in tabs |
| Unhide Window… | Window |  | `view.unhideWindow` |  |
| Chart Slide Show | Workbook Views |  | `view.chartSlideShow` | inspired by Quattro Pro slide shows; 1-2-3 /Graph View |
| Custom Views… | Workbook Views |  | `view.customViews` |  |
| Normal | Workbook Views |  | `view.normal` |  |
| Page Break Preview | Workbook Views |  | `view.pageBreak` |  |
| Page Layout | Workbook Views |  | `view.pageLayout` |  |
| Show Page Breaks | Workbook Views |  | `view.showPageBreaks` | not in tabs |
| 100% | Zoom |  | `view.zoom100` |  |
| Zoom In | Zoom | Ctrl+Alt+= | `view.zoomIn` | not in tabs |
| Zoom Out | Zoom | Ctrl+Alt+- | `view.zoomOut` | not in tabs |
| Zoom to Selection | Zoom |  | `view.zoomSelection` |  |
| Zoom… | Zoom |  | `view.zoom` |  |

## Tools (42)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Correlation Matrix | Analyze |  | `tools.correlation` |  |
| Data Profile | Analyze |  | `tools.dataProfile` | inspired by Data-preparation tools (column profile) |
| Descriptive Statistics | Analyze |  | `tools.descriptiveStats` |  |
| Histogram | Analyze |  | `tools.histogram` |  |
| Summarize by Category… | Analyze |  | `tools.groupSummary` | not in tabs; inspired by Categories in tabular spreadsheet apps |
| / Key Opens Menu | Classic |  | `tools.slashKey` | inspired by Lotus 1-2-3 (and the menu-key option spreadsheets offered 1-2-3 users) |
| 1-2-3 Formula Entry | Classic |  | `tools.lotusEntry` | inspired by Lotus 1-2-3 @functions; transition formula entry in later spreadsheets |
| 1-2-3 Label Prefixes | Classic |  | `tools.lotusLabels` | inspired by Lotus 1-2-3 label prefixes |
| Convert 1-2-3 Formulas | Classic |  | `tools.convertLotus` | inspired by Lotus 1-2-3 |
| Export as Code | Classic |  | `tools.exportCode` | has options menu; inspired by Resolver One (every spreadsheet is a Python program) |
| Slash Menu | Classic | Alt+/ | `tools.slashMenu` | inspired by Lotus 1-2-3 / command menu |
| Change Case | Clean Up |  | `tools.changeCase` | has options menu; inspired by Word processors (Change Case) |
| Clean Up OCR Text | Clean Up |  | `tools.ocrCleanup` | inspired by Transcription and genealogy tools |
| Convert Text to Numbers | Clean Up |  | `tools.textToNumbers` |  |
| Fill Blank Cells | Clean Up |  | `tools.fillBlanks` |  |
| Fix Dates | Clean Up |  | `tools.fixDates` |  |
| Remove Accents | Clean Up |  | `tools.removeAccents` | not in tabs |
| Remove Blank Rows | Clean Up |  | `tools.deleteBlankRows` | not in tabs |
| Remove Characters… | Clean Up |  | `tools.removeChars` | not in tabs |
| Remove Non-Printing Characters | Clean Up |  | `tools.cleanText` |  |
| Remove Repeated Spaces | Clean Up |  | `tools.removeDupSpaces` | not in tabs |
| Trim Spaces | Clean Up |  | `tools.trimSpaces` | inspired by Data clean-up add-ins (TRIM in place) |
| Compare Sheets… | Compare |  | `tools.compareSheets` | inspired by Spreadsheet comparison tools |
| Find Duplicates… | Compare |  | `tools.findDuplicates` |  |
| Fuzzy Match… | Compare |  | `tools.fuzzyMatch` | inspired by Fuzzy lookup add-ins |
| Calendar… | Generate |  | `tools.calendar` |  |
| Date Calculator… | Generate |  | `tools.dateCalculator` | not in tabs; inspired by Desktop calculator date mode |
| Random Data… | Generate |  | `tools.randomData` | inspired by Fill ▸ Random Number in other spreadsheets |
| Sequence… | Generate |  | `tools.sequence` |  |
| Add Prefix or Suffix… | Transform |  | `tools.addText` | not in tabs |
| Convert Formulas to Values | Transform |  | `tools.formulasToValues` | not in tabs |
| Extract with Pattern | Transform |  | `tools.extract` | has options menu; not in tabs |
| Merge Columns… | Transform |  | `tools.mergeColumns` |  |
| Number to Words | Transform |  | `tools.numberToWords` | has options menu; not in tabs; inspired by Check-writing and genealogy document tools |
| Reverse Order | Transform |  | `tools.reverse` | has options menu |
| Roman Numerals | Transform |  | `tools.romanNumerals` | has options menu; not in tabs |
| Round Values… | Transform |  | `tools.roundValues` | not in tabs |
| Shuffle Rows | Transform |  | `tools.shuffle` | inspired by Spreadsheet add-ins (Randomize range) |
| Split Column… | Transform |  | `tools.splitColumn` | inspired by Text to Columns; data tools in other spreadsheets |
| Transpose in Place | Transform |  | `tools.transposeInPlace` |  |
| Unmerge and Fill | Transform |  | `tools.unmergeFill` | not in tabs |
| Unpivot Columns… | Transform |  | `tools.unpivot` | inspired by Query editors (Unpivot) |

## AI (46)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| AI Assistant | Assistant | Ctrl+Alt+a | `ai.chat` |  |
| My Prompts | Assistant |  | `ai.runPrompt` | has options menu; not in tabs |
| Prompt Library | Assistant |  | `ai.prompts` |  |
| Stop AI | Assistant |  | `ai.stop` | not in tabs |
| Analyze Data | Data |  | `ai.analyzeData` |  |
| Ask About Data… | Data |  | `ai.askData` | not in tabs |
| Categorize… | Data |  | `ai.classify` |  |
| Clean Data | Data |  | `ai.cleanData` | has options menu |
| Describe Columns | Data |  | `ai.dataDictionary` | not in tabs |
| Extract | Data |  | `ai.extract` | has options menu |
| Fill by Example | Data |  | `ai.fillByExample` | not in tabs; inspired by Pattern fill in spreadsheets, extended with AI |
| Fill Column with AI… | Data |  | `ai.fillColumn` |  |
| Find Anomalies | Data |  | `ai.detectAnomalies` | not in tabs |
| Suggest Headers | Data |  | `ai.suggestHeaders` | not in tabs |
| Convert Formula | Formulas |  | `ai.convertFormula` | has options menu; not in tabs |
| Explain Formula | Formulas |  | `ai.explainFormula` |  |
| Fix Formula | Formulas |  | `ai.fixFormula` |  |
| Review Formulas | Formulas |  | `ai.reviewFormulas` | not in tabs |
| Simplify Formula | Formulas |  | `ai.simplifyFormula` | not in tabs |
| Write Formula… | Formulas | Ctrl+Alt+f | `ai.writeFormula` |  |
| Build Template | Generate |  | `ai.buildTemplate` | has options menu; not in tabs |
| Conditional Formatting… | Generate |  | `ai.cfRules` | not in tabs |
| Drop-Down List… | Generate |  | `ai.validationList` | not in tabs |
| Generate Table… | Generate |  | `ai.generateTable` |  |
| Sample Data… | Generate |  | `ai.sampleData` | not in tabs |
| Suggest Chart | Generate |  | `ai.suggestChart` |  |
| Suggest PivotTable | Generate |  | `ai.suggestPivot` | not in tabs |
| Summarize | Generate |  | `ai.summarize` | has options menu |
| Write Macro… | Generate |  | `ai.writeMacro` | not in tabs |
| Choose AIs… | Providers |  | `ai.ensembleMembers` |  |
| Several AIs | Providers |  | `ai.ensembleMode` | has options menu |
| Show AI Discussion | Providers |  | `ai.showDeliberation` |  |
| Active AI | Setup |  | `ai.chooseProvider` | has options menu; not in tabs |
| AI & Privacy | Setup |  | `ai.privacy` | not in tabs |
| AI Settings | Setup |  | `ai.settings` |  |
| Delete Downloaded Models | Setup |  | `ai.clearLocal` | not in tabs |
| Providers & Models | Setup |  | `ai.providers` |  |
| Set Up AI | Setup |  | `ai.setup` |  |
| Detect Language | Text |  | `ai.detectLanguage` | not in tabs |
| Explain This Sheet | Text |  | `ai.explainSheet` | not in tabs |
| Extract Keywords | Text |  | `ai.keywords` | not in tabs |
| Proofread Cells | Text |  | `ai.proofreadCells` | not in tabs |
| Rewrite Cells | Text |  | `ai.rewriteCells` | has options menu |
| Sentiment | Text |  | `ai.sentiment` |  |
| Summarize Each Cell | Text |  | `ai.summarizeCells` | not in tabs |
| Translate Cells | Text |  | `ai.translateCells` | has options menu |

## Developer (16)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Copy Command ID List | Code |  | `dev.commandIds` | not in tabs |
| Export Macros | Code |  | `dev.exportMacros` | not in tabs |
| Import Macros | Code |  | `dev.importMacros` | not in tabs |
| Macro Editor | Code | Alt+F11 | `dev.macroEditor` |  |
| Macro Security | Code |  | `dev.macroSecurity` | has options menu |
| Macros | Code | Alt+F8 | `dev.macros` |  |
| Pause Recording | Code |  | `dev.pauseRecording` | not in tabs |
| Record Macro | Code |  | `dev.recordMacro` |  |
| Stop Recording | Code |  | `dev.stopRecording` | not in tabs |
| Use Relative References | Code |  | `dev.relativeRefs` |  |
| Design Mode | Controls |  | `dev.designMode` |  |
| Insert Control | Controls |  | `dev.insertControl` | has options menu |
| Properties | Controls |  | `dev.properties` |  |
| View Code | Controls |  | `dev.viewCode` | not in tabs |
| Export JSON | Data |  | `dev.exportJson` |  |
| Import JSON | Data |  | `dev.importJson` |  |

## Help (8)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| About Anne's Grid | Help |  | `help.about` |  |
| Command Reference | Help |  | `help.commandReference` |  |
| Function Reference | Help |  | `help.functionReference` |  |
| Help | Help | F1 | `help.help` |  |
| Keyboard Shortcuts | Help | Ctrl+/ | `help.shortcuts` |  |
| Report an Issue | Help |  | `help.feedback` |  |
| Show the Quick Tour | Help |  | `help.tour` |  |
| Third-Party Notices | Help |  | `help.notices` | not in tabs |

## Utilities (309)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| AutoFit Columns on All Sheets | Columns & Rows |  | `util.autofitAllSheets` | not in tabs; inspired by ASAP Utilities |
| Columns & Rows | Columns & Rows |  | `util.menuColRow` | has options menu |
| Combine Rows with the Same Key… | Columns & Rows |  | `util.joinRowsByKey` | not in tabs; inspired by Kutools (Advanced Combine Rows) / Ablebits (Combine Rows) |
| Delete Empty Columns | Columns & Rows |  | `util.deleteEmptyCols` | not in tabs; inspired by ASAP Utilities / Kutools / Ablebits |
| Delete Empty Rows | Columns & Rows |  | `util.deleteEmptyRows` | inspired by ASAP Utilities / Kutools / Ablebits |
| Delete Hidden Columns | Columns & Rows |  | `util.deleteHiddenCols` | not in tabs; inspired by ASAP Utilities / Kutools |
| Delete Hidden Rows | Columns & Rows |  | `util.deleteHiddenRows` | not in tabs; inspired by ASAP Utilities / Kutools (Delete Hidden Rows) |
| Duplicate Rows… | Columns & Rows |  | `util.duplicateRows` | not in tabs; inspired by ASAP Utilities / Kutools (Duplicate Rows) |
| Flatten a Range into One Column | Columns & Rows |  | `util.tableToColumn` | has options menu; not in tabs; inspired by ASAP Utilities / Kutools (Transform Range) |
| Hide Empty Columns | Columns & Rows |  | `util.hideEmptyCols` | not in tabs; inspired by ASAP Utilities |
| Hide Empty Rows | Columns & Rows |  | `util.hideEmptyRows` | not in tabs; inspired by ASAP Utilities |
| Insert a Blank Row Where a Value Changes… | Columns & Rows |  | `util.insertRowAtChange` | not in tabs; inspired by ASAP Utilities / Kutools (Distinguish Differences) |
| Insert Blank Columns Every Nth Column… | Columns & Rows |  | `util.insertColsEvery` | not in tabs; inspired by ASAP Utilities / Kutools |
| Insert Blank Rows Every Nth Row… | Columns & Rows |  | `util.insertRowsEvery` | inspired by ASAP Utilities / Kutools (Insert Blank Rows & Columns) |
| Make Selected Columns the Same Width | Columns & Rows |  | `util.equalizeColWidths` | not in tabs; inspired by ASAP Utilities |
| Make Selected Rows the Same Height | Columns & Rows |  | `util.equalizeRowHeights` | not in tabs; inspired by ASAP Utilities |
| Move Columns Left | Columns & Rows | Ctrl+Alt+ArrowLeft | `util.moveColLeft` | not in tabs; inspired by ASAP Utilities / Kutools |
| Move Columns Right | Columns & Rows | Ctrl+Alt+ArrowRight | `util.moveColRight` | not in tabs |
| Move Rows Down | Columns & Rows | Ctrl+Alt+ArrowDown | `util.moveRowDown` | not in tabs |
| Move Rows Up | Columns & Rows | Ctrl+Alt+ArrowUp | `util.moveRowUp` | not in tabs |
| Remove Blank Cells (Shift Up or Left) | Columns & Rows |  | `util.compactCells` | has options menu; not in tabs; inspired by ASAP Utilities |
| Reorder Columns… | Columns & Rows |  | `util.reorderColumns` | inspired by ASAP Utilities (Change the column order) |
| Reshape a Column into a Table… | Columns & Rows |  | `util.columnToTable` | not in tabs; inspired by ASAP Utilities / Kutools (Transform Range) |
| Set Column Width in cm, inches or points… | Columns & Rows |  | `util.colWidthUnits` | not in tabs; inspired by ASAP Utilities / Kutools (Adjust Cell Size) |
| Set Row Height in cm, inches or points… | Columns & Rows |  | `util.rowHeightUnits` | not in tabs; inspired by ASAP Utilities / Kutools |
| Sort Columns by Header | Columns & Rows |  | `util.sortColumnsByHeader` | not in tabs; inspired by ASAP Utilities (Sort columns by header) |
| Unhide All Rows and Columns | Columns & Rows |  | `util.unhideAll` | not in tabs; inspired by ASAP Utilities |
| Back to Previous Sheet | Display & Navigation |  | `util.previousSheetVisited` | not in tabs |
| Center the Active Cell on Screen | Display & Navigation |  | `util.centerActiveCell` | not in tabs |
| Display & Navigation | Display & Navigation |  | `util.menuDisplay` | has options menu |
| Display Settings… | Display & Navigation |  | `util.displaySettings` | inspired by ASAP Utilities (Vision Control) |
| Go to First Sheet | Display & Navigation | Ctrl+Alt+PageUp | `util.firstSheet` | not in tabs |
| Go to Last Sheet | Display & Navigation | Ctrl+Alt+PageDown | `util.lastSheet` | not in tabs |
| Hide/Show All Objects | Display & Navigation |  | `util.hideAllObjects` | not in tabs; inspired by ASAP Utilities |
| Reset the Last Cell (Trim Used Range) | Display & Navigation |  | `util.resetLastCell` | inspired by ASAP Utilities (Reset Excel's last cell) |
| Reset the Last Cell on All Sheets | Display & Navigation |  | `util.resetLastCellAll` | not in tabs |
| Same View Settings on All Sheets | Display & Navigation |  | `util.sameViewAllSheets` | not in tabs; inspired by ASAP Utilities |
| Same Zoom on All Sheets | Display & Navigation |  | `util.zoomAllSheets` | not in tabs; inspired by ASAP Utilities |
| Scroll to the Top-Left (Keep Selection) | Display & Navigation |  | `util.scrollToTop` | not in tabs |
| Toggle Automatic/Manual Calculation | Display & Navigation |  | `util.toggleCalcMode` | not in tabs; inspired by ASAP Utilities |
| Toggle Zoom | Display & Navigation | Ctrl+Alt+z | `util.toggleZoom` | inspired by ASAP Utilities |
| Toggle Zoom Levels… | Display & Navigation |  | `util.toggleZoomLevels` | not in tabs |
| Workbook List… | Display & Navigation |  | `util.windowList` | not in tabs |
| Copy Values Only (Plain Text) | Export | Ctrl+Alt+Shift+c | `util.copyValuesOnly` | inspired by ASAP Utilities (Copy clean values) |
| Export | Export |  | `util.menuExport` | has options menu |
| Export Selection as a LaTeX Table | Export |  | `util.exportSelectionLatex` | not in tabs; inspired by Publishing workflows |
| Export Selection as an HTML Page | Export |  | `util.exportSelectionHtml` | not in tabs |
| Export Selection as CSV | Export |  | `util.exportSelectionCsv` | inspired by ASAP Utilities (Export selection as CSV) |
| Export Selection as Fixed-Width Text | Export |  | `util.exportSelectionFixed` | not in tabs |
| Export Selection as JSON Records | Export |  | `util.exportSelectionJson` | inspired by ASAP Utilities |
| Export Selection as Markdown | Export |  | `util.exportSelectionMd` | not in tabs |
| Export Selection as SQL Statements | Export |  | `util.exportSelectionSql` | not in tabs |
| Export Selection as Tab-Separated Text | Export |  | `util.exportSelectionTsv` | not in tabs |
| Export Selection as XML | Export |  | `util.exportSelectionXml` | not in tabs |
| Export Selection with a Custom Separator… | Export |  | `util.exportSelectionCustom` | not in tabs; inspired by ASAP Utilities |
| Export the Sheet as PDF | Export |  | `util.exportSheetPdf` | not in tabs |
| Save Each Chart as an Image | Export |  | `util.exportChartsImages` | not in tabs; inspired by ASAP Utilities / Kutools (Export Graphics) |
| Clean Up the Workbook (Reduce File Size) | File & System |  | `util.cleanWorkbook` | inspired by ASAP Utilities / Inquire (Clean Excess Cell Formatting) |
| Close All Other Workbooks | File & System |  | `util.closeOthers` | not in tabs; inspired by ASAP Utilities |
| Copy the File Name and Path | File & System |  | `util.copyPath` | not in tabs; inspired by ASAP Utilities / Kutools (Copy Full Path) |
| File & System | File & System |  | `util.menuFile` | has options menu |
| File Name and Date in the Header and Footer | File & System |  | `util.fileNameHeader` | not in tabs; inspired by ASAP Utilities |
| Insert the File Name in the Active Cell | File & System |  | `util.insertFileName` | has options menu; not in tabs; inspired by ASAP Utilities / Kutools (Insert Workbook Information) |
| List the Files in a Folder… | File & System |  | `util.listFolder` | inspired by ASAP Utilities / Kutools (Filename List) |
| Open the Folder That Contains This File | File & System |  | `util.openContainingFolder` | not in tabs; inspired by ASAP Utilities / Kutools (Open Containing Folder) |
| Remove All Macros from the Workbook | File & System |  | `util.removeAllMacros` | not in tabs; inspired by ASAP Utilities |
| Save a Backup Copy with a Timestamp | File & System |  | `util.saveBackup` | inspired by ASAP Utilities (Save file and create backup) |
| Save and Close | File & System | Ctrl+Alt+w | `util.saveAndClose` | not in tabs |
| Save Each Sheet as a Separate File… | File & System |  | `util.exportSheetsFiles` | not in tabs; inspired by ASAP Utilities / Kutools (Split Workbook) |
| Fill | Fill |  | `util.menuFill` | has options menu |
| Fill Down to the Last Row of the Data | Fill | Ctrl+Alt+d | `util.fillDownToLast` | inspired by ASAP Utilities (Quick fill down) |
| Fill Each Column with Its First Cell | Fill |  | `util.fillFirstCell` | not in tabs; inspired by ASAP Utilities |
| Fill Selection with a Value or Formula… | Fill |  | `util.fillWith` | inspired by ASAP Utilities |
| Fill the Next Column with Occurrence Counts | Fill |  | `util.fillCounts` | not in tabs; inspired by ASAP Utilities |
| Fill with a Repeating List… | Fill |  | `util.fillCycle` | not in tabs; inspired by ASAP Utilities (Fill with cyclic values) |
| Fill with Letters (A, B, C…) | Fill |  | `util.fillLetters` | not in tabs |
| Fill with Row Numbers (1, 2, 3…) | Fill |  | `util.fillRowNumbers` | not in tabs; inspired by ASAP Utilities |
| Fill with the Current Date and Time | Fill |  | `util.fillNow` | not in tabs |
| Fill with Today's Date | Fill |  | `util.fillToday` | not in tabs; inspired by ASAP Utilities |
| Color Every Nth Row or Column… | Format & Layout |  | `util.bandRows` | inspired by ASAP Utilities / Kutools (Alternate Row/Column Shading) |
| Color Groups of Rows by Value… | Format & Layout |  | `util.bandByGroup` | not in tabs; inspired by ASAP Utilities / Kutools (Distinguish Differences) |
| Delete Unused Cell Styles | Format & Layout |  | `util.deleteUnusedStyles` | not in tabs; inspired by ASAP Utilities / Inquire (Clean Excess Cell Formatting) |
| Fit the Sheet on One Page | Format & Layout |  | `util.fitOnePage` | has options menu; not in tabs; inspired by ASAP Utilities |
| Format & Layout | Format & Layout |  | `util.menuFormat` | has options menu |
| Format Cells by Content… | Format & Layout |  | `util.formatByContent` | inspired by ASAP Utilities (Apply formatting only to cells with…) |
| Freeze Conditional Formatting | Format & Layout |  | `util.freezeCF` | inspired by ASAP Utilities (Freeze conditional formatting) |
| Outline Each Block of Data | Format & Layout |  | `util.borderBlocks` | not in tabs; inspired by ASAP Utilities |
| Print Area = Used Range on All Sheets | Format & Layout |  | `util.printAreaAllSheets` | not in tabs |
| Remove All Borders | Format & Layout |  | `util.removeBorders` | not in tabs |
| Remove All Fill Colors | Format & Layout |  | `util.removeFills` | not in tabs |
| Remove All Font Colors | Format & Layout |  | `util.removeFontColors` | not in tabs |
| Remove Conditional Formatting from the Workbook | Format & Layout |  | `util.removeCFBook` | not in tabs; inspired by ASAP Utilities |
| Remove Formatting Parts… | Format & Layout |  | `util.removeFormattingParts` | not in tabs; inspired by ASAP Utilities |
| Repeat the First Row When Printing, on All Sheets | Format & Layout |  | `util.printTitlesAll` | not in tabs; inspired by ASAP Utilities |
| Set the Font on All Sheets… | Format & Layout |  | `util.fontAllSheets` | not in tabs; inspired by ASAP Utilities (Change font on all sheets) |
| Show Negative Numbers in Red | Format & Layout |  | `util.negativeRed` | not in tabs; inspired by ASAP Utilities |
| Unwrap Text on the Whole Sheet | Format & Layout |  | `util.unwrapSheet` | not in tabs |
| Break Links to Other Workbooks | Formulas |  | `util.breakExternalLinks` | not in tabs; inspired by ASAP Utilities (Remove all external links) |
| Change Reference Style… | Formulas |  | `util.refStyle` | has options menu; inspired by ASAP Utilities / Kutools (Convert Refers) / Ablebits |
| Color Cells by Content Type | Formulas |  | `util.colorByContentType` | not in tabs; inspired by Financial modelling convention |
| Convert All Formulas in the Workbook to Values | Formulas |  | `util.formulasToValuesBook` | not in tabs; inspired by ASAP Utilities |
| Convert Formulas to Text | Formulas |  | `util.formulasToText` | not in tabs; inspired by ASAP Utilities / Kutools (Convert Formula to Text) |
| Convert Text to Formulas | Formulas |  | `util.textToFormulas` | not in tabs; inspired by ASAP Utilities / Kutools (Convert Text to Formula) |
| Copy Formulas Exactly | Formulas | Ctrl+Alt+c | `util.copyExact` | not in tabs; inspired by ASAP Utilities / Kutools (Exact Copy) |
| Formulas | Formulas |  | `util.menuFormulas` | has options menu |
| List All Formulas | Formulas |  | `util.formulasReport` | inspired by ASAP Utilities / Kutools |
| Paste Formulas Exactly | Formulas | Ctrl+Alt+v | `util.pasteExact` | not in tabs; inspired by ASAP Utilities / Kutools (Exact Copy) |
| Replace Names with Their References | Formulas |  | `util.namesToRefs` | not in tabs; inspired by ASAP Utilities / Kutools (Replace Range Names) |
| Replace References with Names | Formulas |  | `util.refsToNames` | not in tabs; inspired by ASAP Utilities (Apply names) |
| Shift All References in Formulas… | Formulas |  | `util.shiftRefs` | not in tabs; inspired by ASAP Utilities |
| Show Formulas as Text in the Next Column | Formulas |  | `util.formulaTextRight` | not in tabs; inspired by ASAP Utilities |
| Wrap Formulas in a Function… | Formulas |  | `util.wrapFormulas` | inspired by ASAP Utilities / Kutools (Formula Helper) |
| Import | Import |  | `util.menuImport` | has options menu |
| Import a dBASE File (.dbf)… | Import |  | `util.importDbf` | not in tabs; inspired by ASAP Utilities (Import dBASE) |
| Import a Delimited Text File with a Custom Separator… | Import |  | `util.importDelimited` | not in tabs; inspired by ASAP Utilities |
| Import and Stack Files into One Sheet… | Import |  | `util.importStacked` | inspired by ASAP Utilities / Ablebits (Combine Sheets) / Power Tools |
| Import Files as New Sheets… | Import |  | `util.importAsSheets` | inspired by ASAP Utilities (Import several files) / Kutools (Combine) |
| Open Several Files at Once… | Import |  | `util.openMany` | not in tabs |
| Paste the Clipboard as a New Sheet | Import |  | `util.pasteAsNewSheet` | not in tabs |
| Find Problems (Errors, Broken Links, Narrow Columns) | Info & List |  | `util.findProblems` | inspired by ASAP Utilities (Detect and list problems) |
| Info & List | Info & List |  | `util.menuInfo` | has options menu |
| List All Defined Names | Info & List |  | `util.listNames` | not in tabs; inspired by ASAP Utilities (List all named ranges) |
| List All Hyperlinks | Info & List |  | `util.listHyperlinks` | not in tabs; inspired by ASAP Utilities |
| List All Merged Cells | Info & List |  | `util.listMerged` | not in tabs |
| List Document Properties | Info & List |  | `util.listProperties` | not in tabs; inspired by ASAP Utilities |
| List Duplicate Values with Counts | Info & List |  | `util.listDuplicates` | not in tabs; inspired by ASAP Utilities (Count duplicates) |
| List Each Value with Its Count | Info & List |  | `util.valueCounts` | inspired by ASAP Utilities / Kutools (Count by value) |
| List Fonts and Colors Used | Info & List |  | `util.listFonts` | not in tabs; inspired by ASAP Utilities |
| Screen and App Information | Info & List |  | `util.screenInfo` | not in tabs; inspired by ASAP Utilities (Show screen settings) |
| Selection Information | Info & List |  | `util.selectionInfo` | inspired by ASAP Utilities (Information about selected cells) |
| Workbook and Sheet Statistics | Info & List |  | `util.workbookInfo` | not in tabs |
| About the Utilities | My Tools & Shortcuts |  | `util.about` |  |
| Add to My Tools… | My Tools & Shortcuts |  | `util.addToMyTools` | not in tabs |
| Edit My Tools… | My Tools & Shortcuts |  | `util.editMyTools` | not in tabs |
| Find and Run a Tool… | My Tools & Shortcuts | Ctrl+Alt+f | `util.findTool` | inspired by ASAP Utilities (Find and run a tool) |
| My Tools | My Tools & Shortcuts |  | `util.myTools` | has options menu; inspired by ASAP Utilities (My Tools & Shortcuts) |
| Start Last Tool Again | My Tools & Shortcuts | Ctrl+Alt+r | `util.lastTool` | inspired by ASAP Utilities (Start last tool again) |
| Tip of the Day | My Tools & Shortcuts |  | `util.tipOfDay` |  |
| Tool Usage Report | My Tools & Shortcuts |  | `util.usageReport` | inspired by ASAP Utilities (Time and money saved report) |
| Utilities Options… | My Tools & Shortcuts |  | `util.options` | inspired by ASAP Utilities Options |
| Add to or Subtract from Dates… | Numbers & Dates |  | `util.shiftDates` | inspired by ASAP Utilities / Ablebits (Date & Time Wizard) |
| Apply a Calculation to the Numbers… | Numbers & Dates |  | `util.numberOperation` | inspired by ASAP Utilities / Kutools (Operation Tools) |
| Change the Sign of Numbers | Numbers & Dates |  | `util.changeSign` | not in tabs; inspired by ASAP Utilities / Kutools (Change Sign of Values) |
| Convert Dates to Text… | Numbers & Dates |  | `util.datesToText` | not in tabs; inspired by ASAP Utilities |
| Convert Decimal Hours to Times | Numbers & Dates |  | `util.hoursToTime` | not in tabs |
| Convert Number Base… | Numbers & Dates |  | `util.numberBase` | not in tabs; inspired by Kutools (Number Systems Conversion) |
| Convert Times to Decimal Hours | Numbers & Dates |  | `util.timeToHours` | has options menu; not in tabs; inspired by Kutools (Convert Time) |
| Convert Units… | Numbers & Dates |  | `util.convertUnits` | not in tabs; inspired by Kutools (Unit Conversion) |
| Divide by 100 (Percent Points to Fractions) | Numbers & Dates |  | `util.percentToFraction` | not in tabs |
| Insert SUM Below Each Column | Numbers & Dates |  | `util.sumBelow` | has options menu; not in tabs; inspired by ASAP Utilities / Kutools (AutoSum variants) |
| Make All Numbers Positive | Numbers & Dates |  | `util.absValues` | not in tabs; inspired by Kutools (Change Sign of Values) |
| Numbers & Dates | Numbers & Dates |  | `util.menuNumbers` | has options menu |
| Pad with Leading Zeros… | Numbers & Dates |  | `util.padZeros` | inspired by ASAP Utilities / Kutools |
| Remove Leading Zeros | Numbers & Dates |  | `util.removeLeadingZeros` | not in tabs; inspired by Kutools (Remove Leading Zeros) |
| Split Dates into Year, Month and Day Columns | Numbers & Dates |  | `util.splitDates` | not in tabs; inspired by ASAP Utilities |
| Store Numbers as Text | Numbers & Dates |  | `util.numbersToText` | not in tabs; inspired by ASAP Utilities / Kutools (Convert between Text and Number) |
| Weekday Names Beside Dates | Numbers & Dates |  | `util.weekdayNames` | not in tabs |
| Change the Author of All Comments… | Objects & Comments |  | `util.commentAuthor` | not in tabs; inspired by ASAP Utilities / Kutools (Format Comment) |
| Copy Comments into Cells | Objects & Comments |  | `util.commentsToCells` | has options menu; inspired by ASAP Utilities / Kutools (Convert Comment & Cell) |
| Delete All Charts | Objects & Comments |  | `util.deleteAllCharts` | not in tabs; inspired by ASAP Utilities |
| Delete All Comments in the Workbook | Objects & Comments |  | `util.deleteCommentsBook` | not in tabs; inspired by ASAP Utilities |
| Delete All Objects in the Workbook | Objects & Comments |  | `util.deleteAllObjectsBook` | not in tabs; inspired by ASAP Utilities |
| Delete All Objects on the Sheet | Objects & Comments |  | `util.deleteAllObjects` | not in tabs; inspired by ASAP Utilities / Kutools (Delete Illustrations & Objects) |
| Delete All Pictures | Objects & Comments |  | `util.deleteAllPictures` | not in tabs; inspired by ASAP Utilities |
| Delete All Shapes and Text Boxes | Objects & Comments |  | `util.deleteAllShapes` | not in tabs; inspired by ASAP Utilities |
| Find and Replace in Comments… | Objects & Comments |  | `util.commentsFindReplace` | not in tabs; inspired by Kutools (Find/Replace Comment Text) |
| Fit Pictures to Their Cells | Objects & Comments |  | `util.fitPicturesToCells` | not in tabs; inspired by ASAP Utilities / Kutools |
| List All Comments | Objects & Comments |  | `util.commentsReport` | inspired by ASAP Utilities / Kutools (Create Comment List) |
| List All Objects | Objects & Comments |  | `util.objectsReport` | not in tabs; inspired by ASAP Utilities |
| Make Comments from Cell Values | Objects & Comments |  | `util.cellsToComments` | has options menu; not in tabs; inspired by ASAP Utilities / Kutools |
| Objects & Comments | Objects & Comments |  | `util.menuObjects` | has options menu |
| Resize All Comments to Fit | Objects & Comments |  | `util.commentsResize` | not in tabs; inspired by ASAP Utilities / Kutools (AutoFit Comment) |
| Resolve All Comment Threads | Objects & Comments |  | `util.resolveAllComments` | not in tabs |
| Advanced Find and Replace… | Range |  | `util.findReplaceAdvanced` | inspired by ASAP Utilities / Ablebits (Advanced Find and Replace) |
| Advanced Sort… | Range |  | `util.sortAdvanced` | not in tabs; inspired by ASAP Utilities / Ablebits (Advanced Sort) |
| Color Duplicate Values | Range |  | `util.colorDuplicates` | not in tabs; inspired by ASAP Utilities (Color each duplicate with a unique color) |
| Compare Two Ranges… | Range |  | `util.compareRanges` | not in tabs; inspired by Kutools (Compare Cells) / Ablebits (Compare Two Sheets) |
| Convert Merged Cells to Center Across Selection | Range |  | `util.centerAcrossInstead` | not in tabs; inspired by ASAP Utilities |
| Copy a Multi-Area Selection | Range |  | `util.copyMultiSelection` | not in tabs; inspired by ASAP Utilities / Kutools (Copy Ranges) |
| Count and Sum by Color… | Range |  | `util.countSumByColor` | inspired by ASAP Utilities / Ablebits (Count & Sum by Color) |
| Delete Defined Names… | Range |  | `util.deleteNames` | not in tabs; inspired by ASAP Utilities |
| Highlight Selection (Temporary) | Range |  | `util.highlightCells` | not in tabs; inspired by ASAP Utilities (Mark cells) |
| Join Selected Cells into One… | Range |  | `util.mergeToOneCell` | not in tabs; inspired by ASAP Utilities / Kutools (Combine without losing data) |
| Lock Cells with Content, Unlock the Rest | Range |  | `util.lockNonEmpty` | not in tabs; inspired by ASAP Utilities |
| Lock Formula Cells Only | Range |  | `util.lockFormulas` | not in tabs; inspired by ASAP Utilities |
| Lock the Selected Cells | Range |  | `util.lockSelection` | not in tabs |
| Move Selection to a New Sheet | Range |  | `util.moveToNewSheet` | not in tabs |
| Range | Range |  | `util.menuRange` | has options menu |
| Remove Highlights | Range |  | `util.removeHighlights` | not in tabs |
| Replace Values from a List… | Range |  | `util.replaceFromList` | inspired by ASAP Utilities (Replace values in selection with values from a list) |
| Sort by Last Word (Surname) | Range |  | `util.sortByLastWord` | not in tabs; inspired by Genealogy and mailing-list practice |
| Sort by Text Length | Range |  | `util.sortByLength` | not in tabs |
| Sort in Natural Order | Range |  | `util.sortNatural` | not in tabs |
| Swap Two Ranges… | Range |  | `util.swapRanges` | not in tabs; inspired by Kutools / Ablebits (Swap Ranges) |
| Unlock the Selected Cells | Range |  | `util.unlockSelection` | not in tabs |
| Conditional Column Select, Hide or Delete… | Select |  | `util.condCols` | not in tabs; inspired by ASAP Utilities |
| Conditional Row Select, Hide or Delete… | Select |  | `util.condRows` | inspired by ASAP Utilities |
| Deselect Cells… | Select |  | `util.deselect` | not in tabs; inspired by ASAP Utilities |
| Extend Selection to the Last Used Cell | Select |  | `util.extendToEnd` | not in tabs; inspired by ASAP Utilities |
| Extend Selection to the Last Used Column | Select | Ctrl+Alt+Shift+ArrowRight | `util.extendToLastCol` | not in tabs; inspired by ASAP Utilities |
| Extend Selection to the Last Used Row | Select | Ctrl+Alt+Shift+ArrowDown | `util.extendToLastRow` | not in tabs; inspired by ASAP Utilities |
| Invert the Selection | Select |  | `util.invertSelection` | not in tabs; inspired by ASAP Utilities |
| Remember the Selection | Select |  | `util.rememberSelection` | not in tabs; inspired by ASAP Utilities |
| Reselect Previous Selection | Select | Ctrl+Alt+Backspace | `util.reselectPrevious` | not in tabs; inspired by ASAP Utilities |
| Reselect Remembered Selection | Select |  | `util.reselectRemembered` | not in tabs |
| Select | Select |  | `util.menuSelect` | has options menu |
| Select Blank Cells | Select |  | `util.selectBlanks` | not in tabs |
| Select Bold Cells | Select |  | `util.selectBold` | not in tabs |
| Select Cells by Content and Format… | Select |  | `util.selectByContent` | inspired by ASAP Utilities (Select cells based on content, formatting and more) |
| Select Cells with Comments or Notes | Select |  | `util.selectComments` | not in tabs |
| Select Cells with Data Validation | Select |  | `util.selectValidation` | not in tabs |
| Select Cells with Dates | Select |  | `util.selectDates` | not in tabs |
| Select Cells with Errors | Select |  | `util.selectErrors` | not in tabs |
| Select Cells with Hyperlinks | Select |  | `util.selectHyperlinks` | not in tabs |
| Select Cells with Numbers | Select |  | `util.selectNumbers` | not in tabs |
| Select Cells with Text | Select |  | `util.selectText` | not in tabs |
| Select Cells with Text Longer Than… | Select |  | `util.selectLongText` | not in tabs |
| Select Cells with the Largest or Smallest Value | Select |  | `util.selectMaxMin` | has options menu; not in tabs; inspired by Kutools (Select Cells with Max & Min Value) |
| Select Cells with the Same Fill Color | Select |  | `util.selectSameFill` | not in tabs; inspired by ASAP Utilities |
| Select Cells with the Same Font | Select |  | `util.selectSameFont` | not in tabs |
| Select Cells with the Same Font Color | Select |  | `util.selectSameFontColor` | not in tabs; inspired by ASAP Utilities |
| Select Cells with the Same Formatting | Select |  | `util.selectSameFormat` | not in tabs; inspired by ASAP Utilities |
| Select Cells with the Same Number Format | Select |  | `util.selectSameNumFmt` | not in tabs; inspired by ASAP Utilities |
| Select Cells with the Same Value | Select |  | `util.selectSameValue` | not in tabs; inspired by ASAP Utilities |
| Select Cells with TRUE/FALSE | Select |  | `util.selectBooleans` | not in tabs |
| Select Cells with Wrapped Text | Select |  | `util.selectWrapped` | not in tabs |
| Select Duplicate Values | Select |  | `util.selectDuplicates` | has options menu; not in tabs; inspired by Kutools (Select Duplicate & Unique Cells) |
| Select from Here to the End of the Column Data | Select |  | `util.selectToColumnEnd` | not in tabs |
| Select from Here to the End of the Row Data | Select |  | `util.selectToRowEnd` | not in tabs |
| Select from Here to the First Column | Select |  | `util.selectToRowStart` | not in tabs |
| Select from Here to the Top | Select |  | `util.selectToColumnStart` | not in tabs |
| Select Italic Cells | Select |  | `util.selectItalic` | not in tabs |
| Select Locked Cells | Select |  | `util.selectLocked` | not in tabs |
| Select Merged Cells | Select |  | `util.selectMerged` | not in tabs |
| Select Negative Numbers | Select |  | `util.selectNegative` | not in tabs |
| Select Non-Blank Cells | Select |  | `util.selectNonBlanks` | not in tabs |
| Select Positive Numbers | Select |  | `util.selectPositive` | not in tabs |
| Select Random Cells or Rows… | Select |  | `util.selectRandom` | not in tabs; inspired by ASAP Utilities / Kutools (Select Range Randomly) |
| Select Rows with a Blank Cell | Select |  | `util.selectRowsWithBlanks` | not in tabs |
| Select the Print Area | Select |  | `util.selectPrintArea` | not in tabs |
| Select the Used Range | Select |  | `util.selectUsedRange` | not in tabs |
| Select Unlocked Cells | Select |  | `util.selectUnlocked` | not in tabs |
| Select, Hide or Delete Every Nth Row or Column… | Select |  | `util.selectEveryNth` | inspired by ASAP Utilities / Kutools (Select Interval Rows & Columns) |
| Copy Page Setup to Other Sheets… | Sheets |  | `util.pageSetupToSheets` | not in tabs; inspired by ASAP Utilities / Kutools (Copy Page Setup) |
| Copy Sheets to a New Workbook… | Sheets |  | `util.sheetsToNewWorkbook` | not in tabs; inspired by ASAP Utilities / Ablebits (Copy Sheets) |
| Copy This Sheet Several Times… | Sheets |  | `util.copySheetTimes` | not in tabs; inspired by ASAP Utilities / Kutools (Copy Worksheets) |
| Create a Sheet Index (Table of Contents) | Sheets |  | `util.sheetIndex` | inspired by ASAP Utilities / Kutools (Create List of Sheet Names) |
| Delete All Sheets Except This One | Sheets |  | `util.deleteOtherSheets` | not in tabs; inspired by ASAP Utilities / Kutools (Delete All Inactive Sheets) |
| Delete Empty Sheets | Sheets |  | `util.deleteEmptySheets` | not in tabs; inspired by ASAP Utilities / Kutools |
| Delete Hidden Sheets | Sheets |  | `util.deleteHiddenSheets` | not in tabs; inspired by Kutools (Delete All Hidden Sheets) |
| Delete Sheets by Name… | Sheets |  | `util.deleteSheetsByName` | not in tabs |
| Find and Replace in Sheet Names… | Sheets |  | `util.renameSheetsReplace` | not in tabs; inspired by ASAP Utilities / Kutools (Rename Multiple Worksheets) |
| Freeze Panes on All Sheets Like This One | Sheets |  | `util.freezeAllSheets` | not in tabs; inspired by Kutools (Freeze Panes Multiple Worksheets) |
| Hide All Sheets Except This One | Sheets |  | `util.hideOtherSheets` | not in tabs; inspired by ASAP Utilities |
| Hide Sheets… | Sheets |  | `util.hideSheets` | not in tabs; inspired by ASAP Utilities |
| Insert Multiple Sheets… | Sheets |  | `util.insertSheets` | inspired by ASAP Utilities / Kutools (Create Sequence Worksheets) |
| Insert the Sheet Name in the Active Cell | Sheets |  | `util.insertSheetName` | has options menu; not in tabs; inspired by ASAP Utilities / Kutools (Insert Workbook Information) |
| Insert the Sheet Names in Cells | Sheets |  | `util.listSheetNames` | not in tabs; inspired by ASAP Utilities |
| Move Sheet Left | Sheets | Ctrl+Alt+Shift+ArrowLeft | `util.moveSheetLeft` | not in tabs |
| Move Sheet Right | Sheets | Ctrl+Alt+Shift+ArrowRight | `util.moveSheetRight` | not in tabs |
| Move Sheet to the Beginning | Sheets |  | `util.moveSheetFirst` | not in tabs |
| Move Sheet to the End | Sheets |  | `util.moveSheetLast` | not in tabs |
| Protect Several Sheets… | Sheets |  | `util.protectSheets` | not in tabs; inspired by ASAP Utilities / Kutools (Protect Worksheet) |
| Rename Sheets from a Cell… | Sheets |  | `util.renameSheetsFromCells` | not in tabs; inspired by ASAP Utilities |
| Reverse Sheet Order | Sheets |  | `util.reverseSheets` | not in tabs |
| Same Header and Footer on All Sheets | Sheets |  | `util.headerFooterAll` | not in tabs; inspired by ASAP Utilities |
| Sheet Overview Report | Sheets |  | `util.sheetsReport` | not in tabs |
| Sheets | Sheets |  | `util.menuSheets` | has options menu |
| Sort Sheets… | Sheets |  | `util.sortSheets` | inspired by ASAP Utilities / Kutools (Sort Sheets) |
| Tab Colors for Several Sheets… | Sheets |  | `util.tabColors` | not in tabs |
| Unfreeze Panes on All Sheets | Sheets |  | `util.unfreezeAllSheets` | not in tabs; inspired by Kutools |
| Unhide All Sheets | Sheets |  | `util.unhideAllSheets` | not in tabs; inspired by ASAP Utilities / Kutools |
| Unprotect Several Sheets… | Sheets |  | `util.unprotectSheets` | not in tabs; inspired by ASAP Utilities / Kutools (Unprotect Worksheet) |
| A Random Spreadsheet Saying | Start |  | `util.cheer` | not in tabs; inspired by ASAP Utilities (Start ▸ random funny messages) |
| Calculator… | Start |  | `util.calculator` | inspired by ASAP Utilities (Start ▸ Calculator) / Kutools |
| Scratchpad… | Start |  | `util.notepad` | inspired by ASAP Utilities (Start ▸ Notepad) |
| Start | Start |  | `util.menuStart` | has options menu |
| View the Clipboard | Start |  | `util.clipboardView` | not in tabs; inspired by ASAP Utilities (Start ▸ Clipboard viewer) |
| Extract a Word, or Text Between Markers… | Text |  | `util.extractWord` | inspired by ASAP Utilities / Ablebits (Extract Text) |
| Insert Text at a Position… | Text |  | `util.insertText` | inspired by ASAP Utilities / Kutools (Add Text) |
| Keep Only Digits, Letters… | Text |  | `util.keepOnly` | not in tabs; inspired by ASAP Utilities / Ablebits (Extract Text) |
| Pad Text to a Width… | Text |  | `util.padText` | not in tabs; inspired by ASAP Utilities |
| Remove All Spaces | Text |  | `util.removeAllSpaces` | not in tabs; inspired by Kutools (Remove Spaces) |
| Remove Characters by Position… | Text |  | `util.removeByPosition` | inspired by ASAP Utilities / Kutools (Remove by Position) / Ablebits |
| Remove Leading Spaces Only | Text |  | `util.removeLeadingSpaces` | not in tabs; inspired by ASAP Utilities |
| Remove Trailing Spaces Only | Text |  | `util.removeTrailingSpaces` | not in tabs; inspired by ASAP Utilities |
| Replace a Separator with Line Breaks… | Text |  | `util.separatorToLineBreaks` | not in tabs |
| Replace Line Breaks with Spaces | Text |  | `util.lineBreaksToSpaces` | not in tabs; inspired by ASAP Utilities |
| Reverse the Characters | Text |  | `util.reverseText` | not in tabs; inspired by ASAP Utilities / Kutools (Reverse Text Order) |
| Reverse the Word Order | Text |  | `util.reverseWords` | not in tabs; inspired by Kutools |
| Swap Text Around a Separator… | Text |  | `util.swapAround` | not in tabs; inspired by Genealogy and mailing-list practice |
| Text | Text |  | `util.menuText` | has options menu |
| Text Length Beside Each Cell | Text |  | `util.textLengths` | has options menu; not in tabs; inspired by ASAP Utilities / Kutools (Count Words) |
| Trim Spaces in the Whole Workbook | Text |  | `util.trimWorkbook` | not in tabs; inspired by ASAP Utilities |
| Copy Selection as a BBCode Table | Web Data |  | `util.copyAsBbcode` | not in tabs |
| Copy Selection as a Markdown Table | Web Data |  | `util.copyAsMarkdown` | not in tabs |
| Copy Selection as a Wiki Table | Web Data |  | `util.copyAsWiki` | not in tabs |
| Copy Selection as an HTML Table | Web Data |  | `util.copyAsHtml` | inspired by ASAP Utilities (Export selection as HTML) |
| Decode HTML Entities | Web Data |  | `util.decodeEntities` | not in tabs; inspired by ASAP Utilities |
| Decode URL Encoding | Web Data |  | `util.urlDecode` | not in tabs; inspired by ASAP Utilities (URL decode) |
| Encode as HTML Entities | Web Data |  | `util.encodeEntities` | not in tabs |
| Extract Link Addresses | Web Data |  | `util.extractUrls` | has options menu; not in tabs; inspired by ASAP Utilities / Kutools (Convert Hyperlinks) |
| Make Links from Text and Address Columns | Web Data |  | `util.linksFromTwoColumns` | not in tabs |
| Open All Selected Links | Web Data |  | `util.openLinks` | not in tabs; inspired by ASAP Utilities |
| Remove HTML Tags | Web Data |  | `util.stripHtml` | inspired by ASAP Utilities (Remove HTML tags) |
| Split Web Addresses into Parts | Web Data |  | `util.urlParts` | not in tabs |
| Turn Web and Email Addresses into Links | Web Data |  | `util.activateLinks` | inspired by ASAP Utilities / Kutools (Convert Hyperlinks) |
| URL-Encode Text | Web Data |  | `util.urlEncode` | not in tabs |
| Web Data | Web Data |  | `util.menuWeb` | has options menu |

## Plus (82)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Combine | Combine |  | `plus.menuCombine` | has options menu |
| Combine Columns (Keep the Originals) | Combine |  | `plus.combineColumnsKeep` | not in tabs; inspired by Kutools (Combine Rows, Columns or Cells without Losing Data) |
| Combine Sheets… | Combine |  | `util.combineSheets` | inspired by Ablebits (Combine Sheets / Consolidate Sheets) / Kutools (Combine) / Power Tools |
| Merge Two Tables by Key… | Combine |  | `plus.mergeTables` | inspired by Ablebits (Merge Two Tables) / Kutools (Tables Merge) |
| Split Rows into Sheets by Column… | Combine |  | `util.splitToSheets` | inspired by Kutools (Split Data) / Ablebits (Split Table) |
| Convert | Convert |  | `plus.menuConvert` | has options menu |
| Cross-Tab to List, or List to Cross-Tab | Convert |  | `plus.tableToList` | has options menu; inspired by Ablebits (Convert Range) / Kutools (Transpose Table Dimensions) |
| Full-Width to Half-Width Characters | Convert |  | `plus.halfWidth` | not in tabs; inspired by Kutools (Full Width to Half Width) |
| Half-Width to Full-Width Characters | Convert |  | `plus.fullWidth` | not in tabs; inspired by Kutools |
| Split Names… | Convert |  | `plus.splitNames` | inspired by Ablebits (Split Names) / Kutools (Split Names) |
| Subscript the Numbers in Chemical Formulas | Convert |  | `plus.chemicalSubscripts` | not in tabs; inspired by Kutools (Subscript Numbers in Chemical Formulas) |
| Superscript Ordinal Suffixes (1st, 2nd…) | Convert |  | `plus.superscriptOrdinals` | not in tabs; inspired by Kutools (Super/Subscript) |
| Delete | Delete |  | `plus.menuDelete` | has options menu |
| Delete All Form Controls | Delete |  | `plus.deleteFormControls` | not in tabs; inspired by Kutools (Delete Web Controls) |
| Delete Blank Rows on All Sheets | Delete |  | `plus.deleteBlankRowsBook` | inspired by Kutools (Delete Blank Rows) |
| Delete the Visible (Filtered) Rows | Delete |  | `plus.deleteVisibleRows` | not in tabs; inspired by Kutools (Delete Hidden/Visible Rows) |
| Remove All Hyperlinks in the Workbook | Delete |  | `plus.removeLinksBook` | not in tabs; inspired by Kutools (Remove Hyperlinks) |
| Remove Check Boxes | Delete |  | `plus.removeCheckboxes` | not in tabs; inspired by Kutools (Batch Delete Check Boxes) |
| Remove Data Validation | Delete |  | `plus.removeValidation` | has options menu; inspired by Kutools (Clear Data Validation Restrictions) |
| Filter | Filter |  | `plus.menuFilter` | has options menu |
| Filter Rows by Several Conditions… | Filter |  | `plus.superFilter` | inspired by Kutools (Super Filter / Special Filter) |
| Show All Rows Again | Filter |  | `plus.showAllRows` |  |
| Show Only Rows with a Bold Key Cell | Filter |  | `plus.filterBold` | not in tabs; inspired by Kutools (Filter Bold) |
| Show Only Rows with a Comment | Filter |  | `plus.filterComments` | not in tabs; inspired by Kutools (Filter Comment) |
| Show Only Rows with a Formula | Filter |  | `plus.filterFormulas` | not in tabs; inspired by Kutools (Filter Formula) |
| Show Only Rows with the Active Cell's Fill | Filter |  | `plus.filterColor` | not in tabs; inspired by Kutools (Filter by Color) |
| Subtotal Every N Rows… | Filter |  | `plus.subtotalEvery` | not in tabs; inspired by Kutools (Paging Subtotals) |
| Cell Relationship Report | Inquire |  | `plus.cellRelationship` | not in tabs; inspired by Inquire (Cell Relationship) |
| Compare with Another Workbook File… | Inquire |  | `plus.compareFiles` | inspired by Inquire (Compare Files) / Ablebits (Compare Two Sheets) |
| Inquire | Inquire |  | `plus.menuInquire` | has options menu |
| Sheet Relationship Report | Inquire |  | `plus.sheetRelationships` | not in tabs; inspired by Inquire (Worksheet Relationship) |
| Workbook Analysis Report | Inquire |  | `plus.workbookAnalysis` | inspired by Inquire (Workbook Analysis) |
| Add Bullets | Insert |  | `plus.bullets` | has options menu; not in tabs; inspired by Kutools (Insert Bullet) |
| Add Numbering (1. 2. 3.) | Insert |  | `plus.numbering` | has options menu; not in tabs; inspired by Kutools (Insert Numbering) |
| Check Boxes in Every Selected Cell | Insert |  | `plus.batchCheckboxes` | inspired by Kutools (Batch Insert Check Boxes) |
| Find Missing Sequence Numbers | Insert |  | `plus.findMissing` | has options menu; not in tabs; inspired by Kutools (Find Missing Sequence Number) |
| Insert | Insert |  | `plus.menuInsert` | has options menu |
| Insert a Page Break Every N Rows… | Insert |  | `plus.pageBreakEvery` | not in tabs; inspired by Kutools (Insert Page Break Every Row) |
| Insert a Watermark… | Insert |  | `plus.watermark` | not in tabs; inspired by Kutools (Insert Watermark) |
| Insert Random Passwords | Insert |  | `plus.insertRandomPassword` | not in tabs |
| List All Combinations… | Insert |  | `plus.combinations` | not in tabs; inspired by Kutools (List All Combinations) |
| Pick a Date… | Insert | Ctrl+Alt+; | `plus.datePicker` | inspired by Kutools (Date Picker) / XLTools (Popup Calendar) |
| Repeat the Header Row Every N Rows… | Insert |  | `plus.titleRows` | inspired by Kutools (Insert Titles Rows) |
| Lookup | Lookup |  | `plus.menuLookup` | has options menu |
| Lookup All Matches Joined | Lookup |  | `plus.lookupMany` | not in tabs |
| Lookup from the Bottom | Lookup |  | `plus.lookupFromBottom` | not in tabs |
| Lookup Wizard… | Lookup |  | `plus.superLookup` | inspired by Kutools (Super LOOKUP) / Ablebits (Vlookup Wizard) |
| Border Around Each Printed Page | Print |  | `plus.borderEachPage` | inspired by Kutools (Add Border to Each Page) |
| Print | Print |  | `plus.menuPrint` | has options menu |
| Print Only the Page of the Active Cell | Print |  | `plus.printCurrentPage` | not in tabs; inspired by Kutools (Print Current Page) |
| Print the Selected Areas Only | Print |  | `plus.printSelections` | inspired by Kutools (Print Multiple Selection Wizard) |
| Remove All Manual Page Breaks | Print |  | `plus.removePageBreaks` | not in tabs |
| Distinct Values of the Column (SQL) | Query |  | `plus.sqlDistinct` | not in tabs |
| Query | Query |  | `plus.menuQuery` | has options menu |
| SQL Query on the Sheets… | Query |  | `plus.sqlQuery` | inspired by XLTools (SQL Queries) |
| Color the Drop-Down Choices | Range |  | `plus.coloredDropdown` | not in tabs; inspired by Kutools (Colored Drop-down List) |
| Drop-Down List from the Column's Values | Range |  | `plus.dropdownFromColumn` | inspired by Kutools (Drop-down List) / Ablebits |
| Flip Range Horizontally | Range |  | `plus.flipH` | not in tabs; inspired by Kutools / Ablebits (Flip) |
| Flip Range Vertically | Range |  | `plus.flipV` | not in tabs; inspired by Kutools / Ablebits (Flip) |
| Merge Cells with the Same Value | Range |  | `plus.mergeSameCells` | not in tabs; inspired by Kutools (Merge Same Cells) |
| Prevent Typing… | Range |  | `plus.preventTyping` | has options menu; inspired by Kutools (Prevent Typing / Prevent Duplicate Entries) |
| Range | Range |  | `plus.menuRange` | has options menu |
| Range Helper: Add, Subtract or Intersect… | Range |  | `plus.rangeHelper` | inspired by Kutools (Select Range Helper) |
| Select Unmerged Cells | Range |  | `plus.selectUnmerged` | not in tabs; inspired by Kutools |
| Create Email Links from Rows… | Security |  | `plus.mailingLinks` | not in tabs; inspired by Kutools (Create Mailing List / Send Emails) |
| Decrypt the Selected Cells… | Security |  | `plus.decryptCells` | inspired by Kutools (Decrypt Cells) |
| Encrypt the Selected Cells… | Security |  | `plus.encryptCells` | inspired by Kutools (Encrypt Cells) |
| Hide the Formulas When Protected | Security |  | `plus.hideFormulas` | not in tabs; inspired by Kutools (Hide Formulas) |
| Security | Security |  | `plus.menuSecurity` | has options menu |
| Show the Formulas When Protected | Security |  | `plus.unhideFormulas` | not in tabs |
| Delete a Snippet… | View |  | `plus.deleteSnippet` | not in tabs |
| Formula Editor Pane | View |  | `plus.formulaEditor` | inspired by Kutools (Enhanced Edit Bar) / Excel Labs (Advanced Formula Environment) |
| Insert a Saved Snippet… | View |  | `plus.snippets` | has options menu; inspired by Kutools (AutoText) |
| Reading Layout | View | Ctrl+Alt+h | `plus.readingLayout` | inspired by Kutools (Reading Layout) |
| Reading Layout Color… | View |  | `plus.readingLayoutColor` | not in tabs |
| Same Selection on All Sheets | View |  | `plus.syncSelection` | not in tabs; inspired by Kutools (Synchronize Worksheets) |
| View | View |  | `plus.menuView` | has options menu |
| Workbook and Sheet Visibility… | View |  | `plus.sheetVisibility` | not in tabs; inspired by Kutools (Hide/Unhide Workbooks and Sheets) |
| Convert Files to Another Format… | Workbook |  | `plus.formatConverter` | inspired by Kutools (Format Converter) |
| Pin This Workbook to Recent | Workbook |  | `plus.autoOpenList` | not in tabs; inspired by Kutools (Auto Open Those Workbooks Next Time) |
| Split the Workbook into Files… | Workbook |  | `plus.splitWorkbook` | inspired by Kutools (Split Workbook) |
| Workbook | Workbook |  | `plus.menuWorkbook` | has options menu |

## Extras (61)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Base64 Encode / Decode | Data Quality |  | `extra.base64` | has options menu; not in tabs |
| Cluster Similar Values (Fingerprints)… | Data Quality |  | `extra.clusterValues` |  |
| Data Quality | Data Quality |  | `extra.menuQuality` | has options menu |
| Find Columns with Mixed Types | Data Quality |  | `extra.mixedTypes` | not in tabs |
| Hash or Checksum Beside Each Cell | Data Quality |  | `extra.hashColumn` | has options menu; not in tabs |
| Highlight Outliers (IQR Rule) | Data Quality |  | `extra.outliers` |  |
| Normalize Unicode (NFC / NFKC) | Data Quality |  | `extra.unicodeNormalize` | has options menu; not in tabs |
| Reveal Invisible Characters | Data Quality |  | `extra.showInvisible` | not in tabs |
| Z-Score Column | Data Quality |  | `extra.zScores` | not in tabs |
| Date Parts Columns (Weekday, Week, Quarter…) | Dates |  | `extra.dateColumns` |  |
| Dates | Dates |  | `extra.menuDates` | has options menu |
| Days Until / Since Each Date | Dates |  | `extra.daysUntil` | not in tabs |
| Easter Date Beside Years | Dates |  | `extra.easter` | not in tabs |
| Holiday List for a Year… | Dates |  | `extra.holidays` |  |
| Move Dates off Weekends | Dates |  | `extra.businessDayAdjust` | has options menu; not in tabs |
| Timeline Sheet from Dates and Events | Dates |  | `extra.timeline` | not in tabs |
| Aging Buckets from Dates | Finance |  | `extra.agingBuckets` | not in tabs |
| Benford's Law Check | Finance |  | `extra.benford` |  |
| Finance | Finance |  | `extra.menuFinance` | has options menu |
| Fiscal Year and Quarter Columns… | Finance |  | `extra.fiscalYear` | not in tabs |
| Index to 100 Column | Finance |  | `extra.indexTo100` | not in tabs |
| Loan Amortization Schedule… | Finance |  | `extra.amortization` |  |
| Moving Average Column… | Finance |  | `extra.movingAverage` | not in tabs |
| Percent Change Column | Finance |  | `extra.percentChange` | not in tabs |
| Round to the Nearest 0.05, 0.25, 5… | Finance |  | `extra.roundToIncrement` | not in tabs |
| Running Balance Column | Finance |  | `extra.runningBalance` |  |
| Split Signed Amounts into Debit and Credit | Finance |  | `extra.splitDebitCredit` | has options menu; not in tabs |
| Validate Card/ID Numbers (Luhn) | Finance |  | `extra.luhn` | not in tabs |
| Age Between Two Date Columns | Genealogy |  | `extra.ageAtDate` |  |
| Build an Alphabetical Name Index | Genealogy |  | `extra.familyIndex` | not in tabs |
| Convert Julian ⇄ Gregorian Dates | Genealogy |  | `extra.julianGregorian` | has options menu; not in tabs |
| Genealogy | Genealogy |  | `extra.menuGenealogy` | has options menu |
| Names to "Surname, Given" Order | Genealogy |  | `extra.surnameFirst` | not in tabs |
| Parse Genealogy Dates (ABT, BEF, BET…)… | Genealogy |  | `extra.gedcomDates` |  |
| Phonetic Keys Beside Names | Genealogy |  | `extra.metaphone` | not in tabs |
| Quaker Month Numbers to Month Names | Genealogy |  | `extra.quakerDates` | not in tabs |
| Soundex Codes Beside Names | Genealogy |  | `extra.soundex` |  |
| Decimal Degrees to Degrees-Minutes-Seconds | Geography |  | `extra.decimalToDms` | has options menu; not in tabs |
| Degrees-Minutes-Seconds to Decimal | Geography |  | `extra.dmsToDecimal` |  |
| Distance and Bearing Between Coordinates | Geography |  | `extra.distance` |  |
| Geography | Geography |  | `extra.menuGeo` | has options menu |
| Map Links from Coordinates or Addresses | Geography |  | `extra.mapLinks` | not in tabs |
| Check and Convert ISBNs | Publishing |  | `extra.isbn` |  |
| Fix Spacing Around Punctuation | Publishing |  | `extra.fixSpacing` | not in tabs |
| Fraction Characters (1/2 → ½) | Publishing |  | `extra.fractionGlyphs` | not in tabs |
| Make URL Slugs | Publishing |  | `extra.slugify` | not in tabs |
| Points ⇄ Picas, Millimeters, Inches, Pixels | Publishing |  | `extra.pointsConvert` | has options menu; not in tabs |
| Publishing | Publishing |  | `extra.menuPublishing` | has options menu |
| Reading Ease Score Beside Text | Publishing |  | `extra.readingEase` | not in tabs |
| Smart Quotes and Dashes | Publishing |  | `extra.smartPunctuation` |  |
| Straight Quotes and Dashes | Publishing |  | `extra.straightPunctuation` | not in tabs |
| Thin Spaces as Thousands Separators | Publishing |  | `extra.thinSpaces` | not in tabs |
| Title Case (Style-Guide Rules) | Publishing |  | `extra.titleCase` |  |
| Add a To-Do Check Box Column | Workflow |  | `extra.todoColumn` | not in tabs |
| Lock the Rows Marked Done | Workflow |  | `extra.lockFinished` | not in tabs |
| Put the Last Tool on the Quick Access Bar | Workflow |  | `extra.selectionToQat` | not in tabs |
| Stamp with User and Time | Workflow | Ctrl+Alt+s | `extra.auditStamp` |  |
| Undo History as a Sheet | Workflow |  | `extra.changeLog` | not in tabs |
| What Changed Since the Last Snapshot | Workflow |  | `extra.compareSnapshot` |  |
| Workflow | Workflow |  | `extra.menuWorkflow` | has options menu |
| Zebra Rows for Printing | Workflow |  | `extra.zebraPrint` | not in tabs |

## Sheet Tabs (14)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Activate Sheet… |  |  | `sheet.goto` | not in tabs |
| Delete Sheet |  |  | `sheet.delete` | not in tabs |
| Duplicate Sheet |  |  | `sheet.duplicate` | not in tabs |
| Fill Across Worksheets… |  |  | `sheet.fillAcross` | not in tabs |
| Hide Sheet |  |  | `sheet.hide` | not in tabs |
| Insert Sheet |  | Shift+F11, Alt+Shift+F1 | `sheet.insert` | not in tabs |
| Move or Copy Sheet… |  |  | `sheet.moveCopy` | not in tabs |
| Next Sheet |  | Ctrl+PageDown | `sheet.next` | not in tabs |
| Previous Sheet |  | Ctrl+PageUp | `sheet.prev` | not in tabs |
| Rename Sheet |  |  | `sheet.rename` | not in tabs |
| Save This Sheet as CSV… |  |  | `sheet.exportCsv` | not in tabs |
| Select All Sheets |  |  | `sheet.selectAll` | not in tabs |
| Tab Color |  |  | `sheet.tabColor` | not in tabs |
| Unhide Sheet… |  |  | `sheet.unhide` | not in tabs |

## Navigation (4)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Follow Hyperlink |  |  | `nav.followLink` | not in tabs |
| Go to Beginning of Sheet |  | Ctrl+Home | `nav.home` | not in tabs |
| Go to Last Used Cell |  | Ctrl+End | `nav.end` | not in tabs |
| Show Active Cell |  | Ctrl+Backspace | `nav.activeCell` | not in tabs |

## Table Design (contextual) (25)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Resize Table | Properties |  | `table.resize` |  |
| Table Name | Properties |  | `table.name` |  |
| Update Column Names from Headers | Properties |  | `table.syncHeaders` | not in tabs |
| Delete Table Columns | Rows & Columns |  | `table.deleteColumns` | not in tabs |
| Delete Table Rows | Rows & Columns |  | `table.deleteRows` | not in tabs |
| Insert Table Columns to the Left | Rows & Columns |  | `table.insertColumnLeft` | not in tabs |
| Insert Table Columns to the Right | Rows & Columns |  | `table.insertColumnRight` | not in tabs |
| Insert Table Rows Above | Rows & Columns |  | `table.insertRowAbove` | not in tabs |
| Insert Table Rows Below | Rows & Columns |  | `table.insertRowBelow` | not in tabs |
| Select Entire Table | Select |  | `table.selectAll` | not in tabs |
| Select Table Column Data | Select |  | `table.selectColumn` | not in tabs |
| Select Table Row | Select |  | `table.selectRow` | not in tabs |
| Banded Columns | Table Style Options |  | `table.bandedColumns` |  |
| Banded Rows | Table Style Options |  | `table.bandedRows` |  |
| Filter Button | Table Style Options |  | `table.filterButton` |  |
| First Column | Table Style Options |  | `table.firstColumn` |  |
| Header Row | Table Style Options |  | `table.headerRow` |  |
| Last Column | Table Style Options |  | `table.lastColumn` |  |
| Total Row | Table Style Options |  | `table.totalRow` |  |
| Total Row Function | Table Style Options |  | `table.totalFunction` | has options menu; not in tabs |
| Table Styles | Table Styles |  | `table.style` | has options menu |
| Convert to Range | Tools |  | `table.convertToRange` |  |
| Fill Calculated Column | Tools |  | `table.calculatedColumn` | not in tabs |
| Remove Duplicates | Tools |  | `table.removeDuplicates` |  |
| Summarize with PivotTable | Tools |  | `table.summarizePivot` |  |

## PivotTable Analyze (contextual) (13)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Clear All | Actions |  | `pivot.clear` |  |
| Delete PivotTable | Actions |  | `pivot.delete` |  |
| Move PivotTable | Actions |  | `pivot.move` |  |
| Change Data Source | Data |  | `pivot.changeSource` |  |
| Refresh | Data | Alt+F5 | `pivot.refresh` | has options menu |
| Refresh All PivotTables | Data |  | `pivot.refreshAll` | not in tabs |
| Show Details | Data |  | `pivot.drillDown` | not in tabs |
| Grand Totals | Layout |  | `pivot.grandTotals` | has options menu |
| Report Layout | Layout |  | `pivot.layout` | has options menu |
| Subtotals | Layout |  | `pivot.subtotals` | has options menu |
| Field List | PivotTable |  | `pivot.fields` |  |
| Options | PivotTable |  | `pivot.options` |  |
| Value Field Settings | PivotTable |  | `pivot.valueSettings` | not in tabs |

## Chart Design (contextual) (11)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Add Chart Element | Chart Layouts |  | `chart.addElement` | has options menu |
| Edit Chart | Chart Layouts |  | `chart.edit` | not in tabs |
| Quick Layout | Chart Layouts |  | `chart.quickLayout` | has options menu |
| Change Colors | Chart Styles |  | `chart.colors` | has options menu |
| Chart Styles | Chart Styles |  | `chart.style` | has options menu |
| Select Data | Data |  | `chart.selectData` |  |
| Switch Row/Column | Data |  | `chart.switchRowCol` |  |
| Delete Chart | Location |  | `chart.delete` | not in tabs |
| Move Chart | Location |  | `chart.move` |  |
| Save as Picture | Location |  | `chart.saveImage` |  |
| Change Chart Type | Type |  | `chart.changeType` |  |

## Picture Format (contextual) (12)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Alt Text | Accessibility |  | `picture.altText` |  |
| Change Picture | Adjust |  | `picture.change` |  |
| Color | Adjust |  | `picture.color` | has options menu |
| Corrections | Adjust |  | `picture.corrections` | has options menu |
| Picture Effects | Adjust |  | `picture.effects` | has options menu |
| Remove Background | Adjust |  | `picture.removeBg` |  |
| Reset Picture | Adjust |  | `picture.reset` | has options menu |
| Picture Border | Picture Styles |  | `picture.border` | has options menu |
| Picture Shadow | Picture Styles |  | `picture.shadow` |  |
| Crop | Size |  | `picture.crop` | has options menu |
| Format Picture | Size |  | `picture.format` | not in tabs |
| Size and Properties | Size |  | `picture.size` |  |

## Shape Format (contextual) (10)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Change Shape | Insert Shapes |  | `drawing.changeShape` | not in tabs |
| Edit Text | Insert Shapes |  | `drawing.editText` |  |
| Format Object | Shape Styles |  | `object.formatPane` | not in tabs |
| Shape Effects | Shape Styles |  | `drawing.effects` | has options menu |
| Shape Fill | Shape Styles |  | `drawing.fill` | has options menu |
| Shape Outline | Shape Styles |  | `drawing.outline` | has options menu |
| Size | Size |  | `drawing.size` |  |
| Decorative Text Styles | WordArt Styles |  | `drawing.wordArtStyle` | has options menu; not in tabs |
| Text Fill | WordArt Styles |  | `drawing.textFill` | has options menu |
| Text Outline | WordArt Styles |  | `drawing.textOutline` | has options menu |

## Sparkline (4)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Clear Sparklines | Group |  | `sparkline.clear` | not in tabs |
| Edit Sparklines | Sparkline |  | `sparkline.edit` | not in tabs |
| Sparkline Color | Style |  | `sparkline.color` | not in tabs |
| Sparkline Type | Type |  | `sparkline.type` | has options menu; not in tabs |
