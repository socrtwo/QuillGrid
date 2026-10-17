# Heritage: what QuillGrid takes from other spreadsheets

Spreadsheets have been reinvented many times since 1979. Many of the programs below are gone, but several of their ideas never made it into today's mainstream spreadsheets. This page surveys them, product by product, and records for each notable feature whether QuillGrid:

- **has** it already (the existing command id or function),
- **added** it in this release (the new command id or function), or
- **skips** it, and why.

Every command id can be found with command search (Alt+Q), placed on the Quick Access Bar or a custom tab, bound to a key, or run from the Slash Menu. Commands added from this survey carry an "inspired by" note in [COMMANDS.md](COMMANDS.md).

## Added in this release

| Added | Where | Inspired by |
|---|---|---|
| Slash Menu: the whole ribbon as a keyboard menu, typed by letters | `tools.slashMenu` (Alt+/), `tools.slashKey` (make / open it) | Lotus 1-2-3 / menu; the menu-key option spreadsheets offered 1-2-3 users |
| 1-2-3 formula entry: `@SUM(A1..A5)`, `+A1*@AVG(B1..B3)`, `#AND#`, sheet letters, with 1-2-3's meanings | `tools.lotusEntry` (on by default), `tools.convertLotus` | Lotus 1-2-3 @functions; transition formula entry |
| 1-2-3 label prefixes `"` right, `^` centered, `\` repeating | `tools.lotusLabels` | Lotus 1-2-3 |
| Map View: cells drawn as their kind (`"` text, `#` number, `+` formula, `≠` formula unlike its neighbors) | `view.mapView` | Lotus 1-2-3 Release 3 /Worksheet Window Map |
| Export the sheet as a Python or JavaScript program | `tools.exportCode` | Resolver One |
| Explain Value: an outline of how a result is built | `formulas.explainValue` | Javelin's model views; Causal |
| Dependency Diagram of what feeds a cell and what it feeds | `formulas.dependencyDiagram` | Javelin diagram view |
| Formula Rules: each distinct formula once, with its ranges; inconsistent formulas | `formulas.formulaRules`, `formulas.selectInconsistent` | Lotus Improv formula pane; Multiplan R1C1 |
| Info Window: everything about the active cell | `formulas.infoWindow` | Excel for Macintosh Info window |
| Cell History with Restore | `formulas.cellHistory` | Cell edit history in online spreadsheets |
| Chart Slide Show | `view.chartSlideShow` | Quattro Pro slide shows; 1-2-3 /Graph View |
| Data Form: one record at a time | `data.dataForm` | Lotus Symphony FORM; 1-2-3 /Range Input |
| Column Type: typed columns with formats and validation | `data.columnType` | Airtable and Grist field types |
| Categorize by Rules, and the CATEGORIZE function | `data.categorize` | Lotus Agenda |
| Delete Matching Records | `data.queryDelete` | 1-2-3 /Data Query Delete |
| Matrix: invert, multiply, solve, determinant into the sheet | `data.matrix` | 1-2-3 /Data Matrix |
| Monte Carlo Simulation from inputs given as ranges | `data.monteCarlo` | Causal; risk-analysis add-ins |
| What-If Sliders for every input behind a result | `data.whatIfSliders` | Causal; reactive notebooks |
| Functions ISBETWEEN, WEIGHTAVG, SEMEAN, REFCONVERT, XINDEX | worksheet functions | Lotus 1-2-3 97 and Millennium @functions |
| Function CATEGORIZE | worksheet function | Lotus Agenda |
| Functions DAYSINMONTH, DAYSINYEAR, ISLEAPYEAR, WEEKSINYEAR, EASTERSUNDAY | worksheet functions | StarCalc / OpenOffice.org Calc |
| Functions COUNTUNIQUE, ISEMAIL, ISURL, EPOCHTODATE | worksheet functions | Online spreadsheets |

Also fixed along the way: custom data validation formulas now judge the value being typed (they used to see the cell's previous contents), which the Email and Web Address column types rely on.

## The survey

| Product | Feature | QuillGrid | Command / notes |
|---|---|---|---|
| **VisiCalc** (1979) | The electronic grid, A1 references, recalculation on every change | Has | The core of the app |
| | /Replicate with relative or fixed references | Has | `edit.fillDown`, `edit.fillRight`, fill handle; F4 cycles $ while editing |
| | /Title and /Window | Has | `view.freezePanes`, `view.split` |
| | /Global Order (recalculate by rows or by columns) | Skips | Recalculation follows dependencies, so the order setting has no purpose |
| | DIF interchange files | Has (import) | Opens .dif; save as CSV or .xlsx instead |
| **SuperCalc** (1980) | /Arrange (sort), /Zap, /Protect | Has | `data.sort`, `edit.clearAll`, `review.protectSheet` |
| | Graphs in the spreadsheet (SuperCalc3) | Has | Insert ▸ Charts |
| | Iterative calculation of circular models | Has | `formulas.calcOptions` |
| **Multiplan** (1982) | R1C1 references | Has / Added | R1C1 headings in Options; INDIRECT(…, FALSE); R1C1 formulas shown in `formulas.infoWindow` and `formulas.formulaRules` |
| | Named cells and areas | Has | `formulas.defineName`, `formulas.createFromSelection` |
| | Links between worksheets | Has | External links kept with their last values (`data.editLinks`) |
| | SYLK files | Has (import) | Opens .slk |
| **Lotus 1-2-3** (1983–2013: Release 2, 3, 1-2-3/G, 97, Millennium) | / command menu with first-letter accelerators | Added | `tools.slashMenu`, `tools.slashKey` |
| | @functions, `..` ranges, `#AND#`/`#OR#`/`#NOT#`, sheet letters | Added | `tools.lotusEntry`, `tools.convertLotus`; 1-2-3 offsets (@CHOOSE, @INDEX, @VLOOKUP, @MID, @FIND), argument orders (@PMT, @PV, @FV, @TERM, @CTERM, @RATE, @IRR), @YEAR from 1900, @WEEKDAY from Monday, @AVG/@COUNT/@MAX/@MIN/@STD/@VAR counting labels as zero and their @PURE versions, @DSUM family field offsets, @@, @CELL, @CELLPOINTER, @ISSTRING, @ISRANGE, @LENGTH, @REPEAT, @STRING |
| | Label prefixes ' " ^ \ | Added | `tools.lotusLabels` (' already meant text) |
| | /Range Value, /Range Transpose | Has | `tools.formulasToValues`, `edit.pasteValues`, `tools.transposeInPlace`, `edit.pasteTranspose` |
| | /Range Name Create, Labels, Table | Has | `formulas.defineName`, `formulas.createFromSelection`, `formulas.pasteName` (paste list) |
| | /Range Justify, /Range Search | Has | `edit.fillJustify`, `edit.replace` |
| | /Range Format Text (show formulas), Hidden, +/- bar format | Has / Skips | `view.showFormulas`, the `;;;` format; the +/- bar format is superseded by `cond.dataBar` and REPT |
| | /Data Fill with start, step and stop | Has | `edit.fillSeries` |
| | /Data Table 1 and 2 | Has | `data.dataTable` |
| | /Data Table 3 (three-variable, across sheets) | Skips | Rarely used; Scenario Manager, Monte Carlo and per-sheet data tables cover it |
| | /Data Sort, /Data Query Find, Extract, Unique | Has | `data.sort`, `data.filter`, `data.advancedFilter` |
| | /Data Query Delete | Added | `data.queryDelete` |
| | /Data Distribution | Has | FREQUENCY, `tools.histogram` |
| | /Data Matrix Invert and Multiply | Added | `data.matrix` (also Solve and Determinant) |
| | /Data Regression | Has | `data.analysisTools` ▸ Regression, LINEST |
| | /Data Parse | Has | `data.textToColumns` |
| | /Worksheet Window Map (Release 3) | Added | `view.mapView` |
| | /Worksheet Window Perspective (three sheets at once) | Skips | `view.newWindow` with `view.arrangeAll` shows sheets side by side |
| | /Worksheet Learn (keystroke macros), macro language | Has | `dev.recordMacro`, `dev.macroEditor` |
| | /Worksheet Status | Has | `review.workbookStats` |
| | /Graph, /Graph Name, /Graph View | Has / Added | Insert ▸ Charts; `view.chartSlideShow` |
| | /File Combine Add and Subtract | Has | `edit.pasteSpecial` (Add, Subtract), `data.consolidate` |
| | /File Xtract | Has | `sheet.moveCopy`, `sheet.exportCsv` |
| | 3-D worksheets (Release 3) | Has | Sheets and Sheet1:Sheet3!A1 references |
| | Backsolver and Solver | Has | `data.goalSeek`, `data.solver` |
| | Version Manager (Release 4, 97) | Has | `data.scenarios`, `file.versions` |
| | SmartIcons | Has | Quick Access Bar, custom tabs |
| | @ISBETWEEN, @WEIGHTAVG, @SEMEAN, @REFCONVERT, @XINDEX | Added | Worksheet functions of the same names |
| | @DQUERY (external database queries), @COORD | Skips | No database drivers in a local app; ADDRESS does what @COORD did |
| | .wk1/.wk3/.wk4/.123 files | Skips (for now) | Binary formats with compiled formulas need their own reader; the formula translator above is the groundwork |
| **Lotus Symphony** (1984) | Several windows on one worksheet | Has | `view.newWindow`, `view.split` |
| | FORM environment (record entry screens) | Added | `data.dataForm` |
| | DOC and COMM environments | Skips | QuillPen is the word processor; modem communications are obsolete |
| **Lotus Improv** (1991), **Flexisheet**, **Quantrix** | Formulas kept apart from the grid as readable rules | Added | `formulas.formulaRules`; names, tables and LET/LAMBDA give named formulas |
| | Categories (dimensions) dragged to pivot a model | Has | `insert.pivotTable`, `pivot.fields`; a separate multidimensional model engine is out of scope |
| | Saved views | Has | `view.customViews` |
| **Lotus Agenda** (1988) | Items filed into categories automatically by their text | Added | `data.categorize`, CATEGORIZE |
| **Quattro / Quattro Pro** (1987–) | Notebook sheet tabs, block names | Has | Sheet tabs, `formulas.defineName` |
| | SpeedBar, SpeedFill | Has | Quick Access Bar, fill handle, `edit.fillSeries` |
| | Optimizer, Solver, Consolidator | Has | `data.solver`, `data.consolidate` |
| | Graph slide shows, Graphs page | Added / Has | `view.chartSlideShow`; charts listed in `view.navigator` |
| | Object Inspector (properties on right-click) | Has | Context menus, `object.formatPane` |
| | Group mode (edit several sheets at once) | Has | `sheet.selectAll`, `sheet.fillAcross` |
| | 1-2-3-compatible menu | Added | `tools.slashMenu` |
| **Framework** (1984) | Frames arranged in an outline | Skips | One document type here; `view.navigator` outlines a workbook |
| | FRED language | Has | Macro language (`dev.macroEditor`) |
| **Javelin** (1985) | Diagram view of the model | Added | `formulas.dependencyDiagram` |
| | Formulas view; explaining variables | Added | `formulas.formulaRules`, `formulas.explainValue` |
| | Time-series variables | Skips | Date series and FORECAST.ETS cover the common cases |
| **Excel 1.0 for Macintosh** (1985) | Info window | Added | `formulas.infoWindow` |
| | Transition formula entry and menu key for 1-2-3 users | Added | `tools.lotusEntry`, `tools.slashKey` |
| | Transition navigation keys | Skips | The modern keys are what users know now |
| | Macro sheets (XLM) | Skips | VBA and XLM are preserved, not run; macros use JavaScript |
| **Boeing Calc** (1985) | Three-dimensional spreadsheets | Has | Sheets and 3-D references |
| **AppleWorks / ClarisWorks** | Spreadsheet, paint and text in one document | Skips | Charts, shapes and text boxes cover the spreadsheet part |
| | Fill Special | Has | `edit.fillSeries` |
| **Microsoft Works** | Easy Calc formula wizard | Has | `formulas.insertFunction`, `ai.writeFormula` |
| | Templates | Has | `file.newFromTemplate` |
| **Wingz** (1988) | HyperScript, buttons and controls on the sheet | Has | Macro language, `dev.insertControl` |
| | 3-D charts | Skips | Flat charts are read more accurately |
| **Mesa** (NeXT, OS/2) | Multi-page workbooks, links between documents | Has | Sheets; external links kept with their values |
| **StarCalc / OpenOffice.org Calc** | DAYSINMONTH, DAYSINYEAR, ISLEAPYEAR, WEEKSINYEAR, EASTERSUNDAY | Added | Worksheet functions of the same names |
| | Navigator | Has | `view.navigator` |
| **Resolver One** (2008) | The spreadsheet shown as the Python program that computes it | Added | `tools.exportCode` (Python or JavaScript) |
| | Python objects in cells | Skips | Would need a Python runtime in the app; LET and LAMBDA cover named logic |
| **Google Sheets** (2006–) | Cell edit history | Added | `formulas.cellHistory` (this session, as far back as Undo) |
| | ISBETWEEN, COUNTUNIQUE, ISEMAIL, ISURL, EPOCHTODATE | Added | Worksheet functions of the same names |
| | Named functions | Has | LAMBDA in the Name Manager |
| | SPARKLINE() | Has (as objects) | `insert.sparkLine`; the in-cell function form is not provided |
| | QUERY() | Skips | A second query language; FILTER, SORT, GROUPBY and PIVOTBY do the same |
| | IMPORTRANGE, GOOGLEFINANCE, live collaboration | Skips | Need online services; comments and version history work offline |
| | Explore | Has | `ai.analyzeData`, `ai.quickAnalysis` |
| **Airtable** (2012), **Grist** (2020) | Typed fields with pick lists and checkboxes | Added | `data.columnType` |
| | Record cards | Added | `data.dataForm` |
| | Linked records | Has | XLOOKUP, `data.relationships` |
| | Trigger formulas (values stamped when a row is entered) | Skips | Needs entry hooks per column; `edit.insertDate` and macros cover stamping |
| **Notion databases** | Board and gallery views of a table | Skips | Not a spreadsheet view; filters and PivotTables cover analysis |
| **Observable** and reactive notebooks | Reactive recalculation | Has | The engine recalculates only what depends on a change |
| | Sliders bound to inputs | Added | `data.whatIfSliders` |
| **Causal** (2019) | Inputs as ranges with Monte Carlo results | Added | `data.monteCarlo` |
| | Explaining a number from its drivers | Added | `formulas.explainValue`, `formulas.dependencyDiagram` |
| | Sliders | Added | `data.whatIfSliders` |
| **Pyspread**, **Quadratic** (2023) | Python, JavaScript or SQL code in cells | Skips | A language runtime in every workbook is large and needs sandboxing; macros use JavaScript and `tools.exportCode` turns a sheet into code |
| **Ideas without one home** | Sheet comparison | Has | `tools.compareSheets` |
| | Units-aware cells (5 m × 3 s) | Skips | Doing it robustly needs dimensions in the engine's type system; CONVERT handles conversions |
| | A library of ready-made LAMBDA functions | Skips (for now) | LAMBDA names work today; a curated library can come later |
| | Audit map of a sheet | Added | `view.mapView` marks formulas that break their neighbors' pattern |
