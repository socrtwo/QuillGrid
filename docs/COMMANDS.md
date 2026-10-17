# QuillLattice Command Reference

QuillLattice ships **687 commands**. Every one can be placed on the Quick Access Bar, added to a custom tab, bound to a keyboard shortcut, recorded in a macro, or found with command search (Alt+Q).

349 of them are "Commands Not in the Tabs" — reachable through the Quick Access Bar chooser, Customize Tabs, the command search, or menus — so nothing the app can do is hidden.

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
| Export as CSV (Comma delimited) (.csv) | Export |  | `file.export.csv` | not in tabs |
| Export as Encrypted QuillLattice Workbook (.qlwe) | Export |  | `file.export.qlwe` | not in tabs |
| Export as Excel Macro-Enabled Workbook (.xlsm) | Export |  | `file.export.xlsm` | not in tabs |
| Export as Excel Template (.xltx) | Export |  | `file.export.xltx` | not in tabs |
| Export as Excel Workbook (.xlsx) | Export |  | `file.export.xlsx` | not in tabs |
| Export as Flat OpenDocument Spreadsheet (.fods) | Export |  | `file.export.fods` | not in tabs |
| Export as Markdown table (.md) | Export |  | `file.export.md` | not in tabs |
| Export as OpenDocument Spreadsheet (.ods) | Export |  | `file.export.ods` | not in tabs |
| Export as PDF | Export |  | `file.exportPdf` | not in tabs |
| Export as QuillLattice JSON (.json) | Export |  | `file.export.json` | not in tabs |
| Export as QuillLattice Workbook (.qlw) | Export |  | `file.export.qlw` | not in tabs |
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
| About QuillLattice | Help |  | `help.about` |  |
| Command Reference | Help |  | `help.commandReference` |  |
| Function Reference | Help |  | `help.functionReference` |  |
| Help | Help | F1 | `help.help` |  |
| Keyboard Shortcuts | Help | Ctrl+/ | `help.shortcuts` |  |
| Report an Issue | Help |  | `help.feedback` |  |
| Show the Quick Tour | Help |  | `help.tour` |  |
| Third-Party Notices | Help |  | `help.notices` | not in tabs |

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
