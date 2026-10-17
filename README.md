# QuillGrid

*Part of the Quill suite: [QuillPen](https://github.com/socrtwo/QuillPen) (documents), **QuillGrid** (spreadsheets), [QuillPost](https://github.com/socrtwo/QuillPost) (email) and [QuillStage](https://github.com/socrtwo/QuillStage) (presentations). Formerly named QuillLattice; its .qlw/.qlwe workbooks, settings and drafts carry over.*

An original, open-source spreadsheet for Android (and the web, desktop and iOS) whose goal is simple: **no missing commands**. Every feature is a command, and every command can be put on the Quick Access Bar, on your own tab, on a keyboard shortcut, or in a macro, which is the full-library customization that mobile and online spreadsheet apps leave out. It is the companion to the [QuillPen](https://github.com/socrtwo/QuillPen) word processor and shares its interface and customization system.

- **687 commands** across File, Home, Insert, Draw, Page Layout, Formulas, Data, Review, View, Tools, **AI**, Developer and Help, plus contextual Table Design, PivotTable Analyze, Chart Design, Picture Format and Shape Format tabs. See [docs/COMMANDS.md](docs/COMMANDS.md).
- **522 worksheet functions**, including dynamic arrays (FILTER, SORT, UNIQUE, SEQUENCE, XLOOKUP, TEXTSPLIT, VSTACK, GROUPBY, PIVOTBY…), LET and LAMBDA with MAP/REDUCE/SCAN/BYROW/BYCOL, and the full statistical, financial, engineering and database libraries.
- **Opens damaged workbooks** with a recovery engine built on the author's [Word Recovery](https://github.com/socrtwo/wordrecovery-SF) / Immortal Unzipper heritage: byte-level ZIP salvage, XML repair, raw worksheet scanning, and record-level salvage of Excel 97–2003 .xls files. See [docs/RECOVERY.md](docs/RECOVERY.md).
- **AI tab** with formula writing and explanation, data analysis, fill-by-instruction, classification, extraction and cleaning. It offers a free on-device model (no account) and a setup wizard for free and paid providers. See [docs/AI.md](docs/AI.md).
- **Quick Access Bar chooser** with the desktop-style categories: *Popular Commands*, *Commands Not in the Tabs*, *All Commands*, each tab, and *Macros*. You can add separators, reorder, reset, and import or export the layout. Long-press any ribbon button to add it, and the bar can sit at the top or at the bottom near the keyboard. Its drop-down has the full desktop checklist (AutoSave, New, Open, Save, Email, Quick Print, Print Preview and Print, Spelling, Undo, Redo, Sort Ascending/Descending, Touch/Mouse Mode) and **Add All Choices**, which puts Save and Print first and everything else in alphabetical order.
- **Customize Tabs:** create tabs and groups, put any command in them, rename, reorder or hide built-in tabs.
- **Customize Keyboard:** rebind any command, with conflict display.
- **Command search** (Alt+Q or the 🔍 button) across the whole library.

## Install

Each [release](../../releases) has builds for:
- **Windows** (installer for x64 or ARM, or portable);
- **macOS** (universal);
- **Linux** (AppImage or .deb for x64, arm64 and armv7l, which covers **Raspberry Pi OS**);
- **ChromeOS** (Android app, Linux app or web app);
- **Android** (APK);
- **iOS/iPadOS** (an unsigned .ipa you sign with your Apple ID);
- the **web** (static files; offline-capable and installable).

[docs/INSTALL.md](docs/INSTALL.md) has the steps for each. To build it yourself, including the APK on the phone itself with Termux, see [docs/BUILD.md](docs/BUILD.md).

On Android, QuillGrid registers for *Open with* and *Share to* on .xlsx, .xlsm, .xls, .ods, .csv, .tsv and .txt files, and for *Open with* on Lotus, Quattro, Works, dBASE, SYLK and DIF files. It uses the Android system file picker, so files on the device, SD card, Google Drive, Dropbox and similar providers all work. **Save** writes back to the same file.

## Features

**Cells and formulas.** Type a value and press Enter; numbers, currency, percentages, fractions, dates and times are recognized as you type. Formulas start with `=`:
- click or drag over cells (or use the arrow keys) to insert references, each highlighted in its own color on the grid;
- function names autocomplete with a description, and an argument tip follows the caret;
- F4 cycles absolute and relative references; F9 evaluates a selected part of a formula in place;
- formulas that return several values spill into neighboring cells (`A1#` refers to the spill range), and Ctrl+Shift+Enter enters a legacy array formula;
- structured table references (`Sales[Amount]`, `[@Qty]`), defined names (workbook or sheet scope), 3-D references (`Sheet1:Sheet3!B2`), R1C1 references in INDIRECT, and iterative calculation for circular references.

Recalculation is incremental: dependencies are recorded as formulas are evaluated, so only what a change affects is recalculated. Manual calculation mode, Calculate Now (F9) and Calculate Sheet are available.

**Grid.** A fast canvas grid handles a million rows: frozen panes, merged cells, wrapped and rotated text, text overflow into empty neighbors, hidden rows and columns, outline groups, gridline color, zoom from 10% to 400%, and right-to-left sheets. The fill handle continues numbers, dates, custom lists (days, months, quarters and your own) and text with numbers, and copies formulas with adjusted references; double-click it to fill down beside existing data. Drag a selection border to move cells (Ctrl to copy).

**Formatting.** Fonts, sizes, bold/italic/underline (single, double, accounting), strikethrough, superscript and subscript, font and fill colors, patterns and gradients, all border styles, alignment (including Center Across Selection, Fill and Distributed), indent, orientation, wrap and shrink to fit. The number-format engine supports every built-in category and custom codes with sections, conditions, colors, fractions, scientific notation, elapsed time and text placeholders. Cell styles, Format Painter (double-tap to keep it on), Format as Table with 60 table styles, and conditional formatting (highlight rules, top/bottom, above average, duplicates, formulas, data bars, color scales and icon sets, with a Rules Manager).

**Data.** Sort by values, colors or icons with multiple levels and custom lists; AutoFilter and table filters with value lists, text/number/date filters and top 10; Advanced Filter; Text to Columns; Flash Fill; Remove Duplicates; Data Validation with drop-down lists, input messages and error alerts; Consolidate; Group, Subtotal and outline levels; Goal Seek, Scenario Manager, Data Tables and a Solver; Forecast Sheet; an analysis toolkit (descriptive statistics, histogram, regression, t-tests, ANOVA, correlation, moving average, sampling and more); and Get Data from text files, the web, tables and the clipboard with a transform step list.

**Insert.** PivotTables with a field list (rows, columns, values, filters; several layouts, subtotals and "show values as"), charts (column, bar, line, area, pie, doughnut, scatter, bubble, radar, histogram, Pareto, box and whisker, waterfall, funnel, treemap, sunburst, stock and combo), sparklines, pictures, 40+ shapes, icons, diagrams, text boxes, decorative text, hyperlinks, symbols, in-cell checkboxes, threaded comments and notes.

**Formulas tab.** Insert Function with an argument wizard that evaluates as you type, Name Manager, Create from Selection, Trace Precedents and Dependents with arrows on the grid, Error Checking, step-by-step Evaluate Formula, Watch Window, and calculation options.

**Page Layout and printing.** Themes, margins, orientation, paper size, print areas, page breaks, print titles, fit to pages or scale, gridlines and headings, headers and footers with fields, Page Break Preview and Page Layout views. Printing and PDF go through the system print service ("Save as PDF").

**Review.** Spelling for cell text (offline Hunspell dictionaries for every variety of English, and downloadable dictionaries for 90+ languages), thesaurus, workbook statistics, an accessibility checker, threaded comments with replies and resolve, notes, Protect Sheet (with the full list of allowed actions) and Protect Workbook, Allow Edit Ranges, and Mark as Final.

**Developer.** Macro recorder (records commands, selections and typing; relative or absolute), a JavaScript macro editor with a workbook API, form controls (buttons, check boxes, lists, spin buttons) linked to cells, and JSON import/export.

**Files.** Open and save .xlsx, .xlsm (macros kept, not run) and .xltx, OpenDocument .ods and .fods, CSV/TSV/text (with a Text Import Wizard for delimiters, encodings and column types), HTML, and the native .qgw format. Excel 97–2003 .xls files open with values, formulas and formatting. SYLK, DIF and XML Spreadsheet 2003 files open too. AES-256 encrypted .qgwe is also supported.

**Older spreadsheets, documents and mail.** Files are recognized by their contents, so a wrong or missing extension does not matter (a short note says when that happened). Spreadsheets from earlier programs open with their values and formulas, translated to keep their original meaning: Excel 2.x, 3.0, 4.0 (including 4.0 workbooks) and 5.0/95, Lotus 1-2-3 (.wks, .wk1, .wk3, .wk4, .123) and Symphony, Quattro and Quattro Pro (.wq1, .wq2, .wb1-.wb3), Microsoft Works (.wks, .xlr), dBASE/FoxPro/Clipper tables (.dbf) and VisiCalc; AppleWorks, SuperCalc, binary Multiplan and Quattro Pro .qpw files give up their values or text. Word-processor documents (.docx, .doc, .odt, .rtf, WordPerfect and more), presentations (.pptx, .ppt, .odp) and e-mail (.eml, .mbox, .msg) can be handed to QuillPen, QuillStage or QuillPost, or imported: each table becomes a sheet, slides become sheets, and a message becomes its header fields and body, with CSV and spreadsheet attachments as extra sheets. See [docs/FORMATS.md](docs/FORMATS.md) for every format, version and fidelity level. Autosave drafts with crash recovery, version history snapshots, a document inspector, templates, and several open workbooks in tabs.

### Innovations from other spreadsheets

| Feature | Inspired by |
|---|---|
| Navigator pane (sheets, names, tables, objects, comments) | LibreOffice Calc |
| Fill Across Worksheets, Fill Justify, Go To Special | Classic desktop spreadsheets |
| Data profile, fuzzy matching, compare sheets, unpivot, split/merge columns | Data-wrangling tools |
| OCR clean-up for cells (broken lines, hyphenation, O/0 and l/1 confusions, ligatures) | Publishing and OCR workflows |
| Number to words, Roman numerals, date calculator | Accounting and genealogy record keeping |
| Slash Menu (Alt+/), 1-2-3 formula entry (`@SUM(A1..A5)`, `#AND#`) and label prefixes, Map View | Lotus 1-2-3 |
| Data Form, Delete Matching Records, Matrix (invert, multiply, solve) | Lotus Symphony, 1-2-3 /Data commands |
| Formula Rules, Dependency Diagram, Explain Value, Info Window | Lotus Improv, Multiplan, Javelin, Excel for Macintosh |
| Chart Slide Show; Categorize by Rules | Quattro Pro; Lotus Agenda |
| Export a sheet as a Python or JavaScript program | Resolver One |
| Column types; Monte Carlo Simulation and What-If Sliders; Cell History | Airtable and Grist; Causal; online spreadsheets |
| ISBETWEEN, WEIGHTAVG, SEMEAN, REFCONVERT, XINDEX, COUNTUNIQUE, ISEMAIL, ISURL, EPOCHTODATE and date helpers | 1-2-3, online spreadsheets, StarCalc |
| Command search, keyboard rebinding for every command, QAT at the bottom near the keyboard | QuillPen |
| Metric-compatible open fonts substituted for Calibri, Cambria, Times New Roman, Arial, Courier New, Georgia and Garamond | Document-fidelity practice in open-source suites |

[docs/HERITAGE.md](docs/HERITAGE.md) surveys VisiCalc, SuperCalc, Multiplan, Lotus 1-2-3, Symphony, Improv, Agenda, Quattro Pro, Framework, Javelin, Excel 1.0, Boeing Calc, Wingz, Resolver One and today's database-style and model-style spreadsheets, feature by feature: what QuillGrid already had, what was added, and what was left out and why.

**Recovery.** When a workbook is damaged, QuillGrid recovers it instead of refusing to open it. The pipeline identifies the real file type, salvages ZIP entries from their local headers with ImmortalInflate, checks CRCs, repairs XML (cutting sheet data at the last complete row and regenerating missing package parts), scans raw worksheet XML for cell values when the package cannot be rebuilt, reads Excel 97–2003 records even when the container is damaged, and falls back to text and strings scans. The recovered workbook opens as a new, unsaved file with a report of what each method found.

**AI.** The AI tab writes formulas from a plain-language request (checked by the formula engine before it is offered), explains and fixes formulas, analyzes data, fills columns by instruction or example, classifies, extracts and cleans data, generates tables, suggests charts, summarizes, translates and rewrites cells, and has a chat pane with workbook context. With two or more providers set up, a command can go to several AIs at once: one AI combines their answers, or they deliberate in the manner of a Quaker meeting (unite, stand aside or block) until they agree, with the discussion shown in the AI pane. Quick Analysis (formatting, charts, totals, tables, sparklines for the selection) works without AI.

Providers:
- An on-device model (free, private, offline after one download).
- Chrome's built-in model.
- Free tiers: OpenRouter, Gemini, Groq, Mistral, Cerebras, SambaNova, Z.ai, NVIDIA, Cloudflare.
- Ollama or LM Studio on your computer.
- Paid keys: OpenAI, Anthropic, DeepSeek, xAI, Perplexity, Together, Fireworks, Hugging Face, Moonshot, Qwen, Azure, Bedrock, or any OpenAI-compatible service.

Keys are stored encrypted on the device (Android Keystore on Android, the system keychain on desktop; a plain browser keeps them in site storage), and nothing is sent until you use an AI command.

## Known limitations (0.1)

- **Native builds are untested on devices.** The web build, 497 unit tests, a browser smoke test (`npm run e2e`) and a phone layout audit (`npm run e2e:layout`: every window, menu and pane on emulated phones with system bars and the on-screen keyboard) pass; the Android, iOS and desktop packages build from the same code but have not yet been exercised on real hardware.
- **VBA is preserved, not run.** .xlsm files keep their VBA project when saved as .xlsm; macros run only in QuillGrid's JavaScript macro language.
- **PivotTables from other programs open as static values.** PivotTables made in QuillGrid are live, but they are written to .xlsx as values, not as pivot caches.
- **OpenDocument charts and embedded objects** are not opened or saved; cell data, formulas, styles, merges and conditional formats are.
- **.xls is import-only.** Excel 97–2003 files open with values, formulas and formatting; save them as .xlsx.
- **Legacy formats are import-only**, and some open with less than everything: Lotus 1-2-3 Release 4+ files keep values and formulas but not number formats; .qpw, AppleWorks, SuperCalc and binary Multiplan files give text or values only. See [docs/FORMATS.md](docs/FORMATS.md).
- **No network functions in cells.** WEBSERVICE returns #VALUE!; STOCKHISTORY and linked data types have no data source. ENCODEURL and FILTERXML work.
- **External workbook links** are kept as their last values; they are not refreshed from the other file.
- **Solver** uses simplex with branch and bound for linear models (integer and binary constraints supported) and a penalty Nelder–Mead search for nonlinear ones, which finds local optima only; the AllDifferent constraint is not supported.
- **Printing** goes through the system print dialog; there is no direct PDF writer, so page breaks can differ slightly between platforms.
- **Performance:** a 50,000-cell dependency chain recalculates in about 0.1 s incrementally and 2.5 s in full on a desktop; very large workbooks on low-end phones will be slower.

## Build

```bash
npm ci
npm run dev          # web app at http://localhost:5173
npm test             # unit tests (vitest)
npm run build        # production web build in dist/ (first writes public/third-party-notices.txt)
npm run android:build   # build + cap sync + gradle assembleRelease (needs Android SDK)
npm run desktop         # run the desktop app (Electron); desktop:dist builds installers
npm run ios:sync        # web build into the Xcode project (ios/)
npm run catalog         # regenerate docs/COMMANDS.md from the command registry
```

CI (`.github/workflows/android.yml`) type-checks, tests, builds the web app, and produces an APK and AAB on every push. The **Release** workflow (`.github/workflows/release.yml`) builds every platform and publishes them together; run it from the Actions tab or push a tag `v*`. See [docs/BUILD.md](docs/BUILD.md) and [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## Project language

The interface, code comments and documentation use American English. Spelling checks accept every variety of English by default.

## Legal

QuillGrid is an independent, clean-room project released under the MIT License. It is not affiliated with, endorsed by, or sponsored by Microsoft, Google, Apple, The Document Foundation, any AI provider, or any other vendor. Product names appear only to describe file-format compatibility or the origin of an idea (nominative use). Excel is a trademark of Microsoft Corporation. The interface, icons (Material Symbols, Apache 2.0), templates, styles and text are original or openly licensed. See [docs/LEGAL.md](docs/LEGAL.md).
