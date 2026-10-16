# QuillLattice Command Reference

QuillLattice ships **268 commands**. Every one can be placed on the Quick Access Bar, added to a custom tab, bound to a keyboard shortcut, recorded in a macro, or found with command search (Alt+Q).

197 of them are "Commands Not in the Tabs" — reachable through the Quick Access Bar chooser, Customize Tabs, the command search, or menus — so nothing the app can do is hidden.

This file is generated from the source by `npm run catalog`.

## Default tab layout

- **Home**: Clipboard · Font · Alignment · Number · Styles · Cells · Editing · Analysis
- **Insert**: Tables · Illustrations · Charts · Sparklines · Filters · Links · Comments · Text · Symbols
- **Draw**: Drawing Tools · Convert · Replay
- **Page Layout**: Themes · Page Setup · Scale to Fit · Sheet Options · Arrange
- **Formulas**: Function Library · Defined Names · Formula Auditing · Calculation
- **Data**: Get & Transform Data · Queries & Connections · Data Types · Sort & Filter · Data Tools · Forecast · Outline · Analyze
- **Review**: Proofing · Accessibility · Insights · Language · Comments · Notes · Protect · Ink
- **View**: Sheet View · Workbook Views · Show · Zoom · Window · Appearance · Macros
- **Tools**: Clean Up · Transform · Compare · Analyze · Generate
- **AI**: Formulas · Data · Generate · Text · Assistant · Setup
- **Developer**: Code · Controls · Data
- **Help**: Help
- **Table Design** (contextual): Properties · Tools · Table Style Options · Table Styles
- **PivotTable Analyze** (contextual): PivotTable · Data · Actions · Layout
- **Chart Design** (contextual): Chart Layouts · Chart Styles · Data · Type · Location
- **Picture Format** (contextual): Adjust · Picture Styles · Accessibility · Arrange · Size
- **Shape Format** (contextual): Insert Shapes · Shape Styles · WordArt Styles · Arrange · Size

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

## Home (152)

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
| Cell Styles | Styles |  | `fmt.cellStyles` | has options menu |
| New Cell Style… | Styles |  | `fmt.newCellStyle` | not in tabs |
| Style… | Styles | Alt+' | `fmt.styleDialog` | not in tabs |
| Redo | Undo | Ctrl+y, Ctrl+Shift+z | `edit.redo` | not in tabs |
| Repeat | Undo | F4 | `edit.repeat` | not in tabs |
| Undo | Undo | Ctrl+z | `edit.undo` | not in tabs |

## Page Layout (2)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Selection Pane | Arrange |  | `view.selectionPane` |  |
| Sheet Right-to-Left | Sheet Options |  | `view.rightToLeft` | not in tabs |

## Formulas (1)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Show Formulas | Formula Auditing | Ctrl+` | `view.showFormulas` |  |

## View (33)

| Command | Group | Shortcut | ID | Notes |
|---|---|---|---|---|
| Collapse the Ribbon | Appearance | Ctrl+F1 | `view.collapseRibbon` | not in tabs |
| Dark Cells | Appearance |  | `view.darkSheet` |  |
| Full Screen | Appearance | Ctrl+Shift+F1 | `view.fullScreen` |  |
| Navigator | Appearance |  | `view.navigator` | inspired by LibreOffice Calc Navigator |
| Search Commands | Appearance | Alt+q | `view.cmdSearch` | not in tabs |
| Sheet View | Sheet View |  | `view.sheetViewMenu` | has options menu |
| Formula Bar | Show |  | `view.formulaBar` |  |
| Gridline Color… | Show |  | `view.gridColor` | not in tabs |
| Gridlines | Show |  | `view.gridlines` |  |
| Headings | Show |  | `view.headings` |  |
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
