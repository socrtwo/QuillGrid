# Architecture

Anne's Grid is a TypeScript web application. The Android and iOS releases package it with Capacitor, the desktop releases with Electron, and the same build runs in a browser as an installable, offline-capable web app. It shares its application shell (command registry, Quick Access Bar, customizable tabs, dialogs, AI setup, dictionary downloads, recovery engine and platform layer) with its sibling, Anne's Pen.

```
src/
  main.ts                 entry: fonts, functions, commands, App + Shell, launch-file handling, back button
  app.ts                  App (workbook tabs, exec(), recording, autosave) and BookTab (workbook, engine, grid)
  prefs.ts                per-device preferences: QAT, custom tabs, keymap, macros, custom lists, editing options
  templates.ts            built-in workbook templates
  model/
    address.ts            A1 addressing, ranges, cell keys
    types.ts              values, errors, cell styles, validation, conditional formats, tables, charts, pivots…
    workbook.ts           Workbook and Sheet: sparse cells, interned styles, transactions with undo/redo
    structure.ts          insert/delete rows and columns, move ranges, rename sheets — with reference adjustment
    numfmt.ts             number-format engine (General, decimals, currency, dates, fractions, sections, colors)
    dates.ts              date serials (1900 and 1904 systems) and typed date/time parsing
    tableStyles.ts        table style palette generated from theme accents
  calc/
    parser.ts             tokenizer and parser (precedence, ranges, intersections, unions, structured refs, arrays)
    values.ts             runtime values, coercion rules and helpers used by functions
    engine.ts             evaluator, dynamic dependency tracking, incremental recalculation, spilling, iteration
    functions/            ~520 worksheet functions by category (math, stats, text, date, lookup, dynamic arrays,
                          logical/LAMBDA, financial, engineering, database, information, web, heritage)
  grid/
    grid.ts               canvas grid: virtual scrolling, frozen panes, selection, fill handle, drag-move, touch
    render.ts             painter: fills, gridlines, text with overflow, borders, merges, data bars, icons, headers
    editor.ts             in-cell and formula-bar editing, reference pointing and highlights, autocomplete, tips
    objects.ts            floating charts, pictures, shapes, text boxes and comment pop-ups
    axis.ts               row/column geometry with sparse custom sizes
    cf.ts                 conditional-format evaluation
  charts/                 SVG chart renderer and shape library
  commands/
    registry.ts           Command interface and registry (the single source for QAT, tabs, keys, macros, search)
    ops.ts                shared editing helpers (selection, styles, input parsing, fill series, clipboard text)
    *.ts                  command modules by area (file, edit, format, cells, view, insert, charts, table,
                          formulas, pivot, conditional, data, page layout, review, tools, ai, developer, draw,
                          heritage: features from earlier spreadsheets, see docs/HERITAGE.md)
  heritage/               logic behind the heritage commands: 1-2-3 formula translation, R1C1 and formula
                          rules, explain/dependency graphs, cell history, matrices, Monte Carlo, code export
  io/                     format converters: xlsx/xlsm/xltx, ods/fods, csv/tsv, html, sylk, dif, xml 2003, native
    recover/              recovery engine (ZIP salvage, ImmortalInflate, XML repair, BIFF reader and salvage)
  print/                  page breaks, print rendering and preview
  proof/                  spelling worker (nspell / Hunspell) and dictionary downloads
  ai/                     providers, client, settings, on-device model, spreadsheet AI features
  ui/                     shell, ribbon layout, dialogs, panes, backstage, customize dialogs, wizards
  platform/native.ts      JS side of the QuillNative Capacitor plugin and the Electron bridge, with browser fallbacks
android/ ios/ electron/   native packaging
```

## Key design decisions

- **Commands are data.** The ribbon, the Quick Access Bar, context menus, keyboard shortcuts, macros and search all refer to command ids, so every feature can be placed anywhere. `npm run catalog` regenerates `docs/COMMANDS.md`.
- **Transactions record before-images.** Every change runs inside `Workbook.transact()`. Cell edits record the previous cell; structural edits record a snapshot of the sheet or workbook. Undo restores the before-images and captures the current state as the redo record, so no command needs inverse logic.
- **Dependencies are recorded while evaluating.** Each formula remembers the ranges it actually read. INDIRECT, OFFSET, names and IF branches are therefore tracked exactly. A change marks dependents dirty transitively; dirty formulas are evaluated on demand when read, which yields a correct order without a global sort. Ranges are indexed by column blocks so large pastes stay fast.
- **Dynamic arrays spill.** A formula returning an array writes its values into the cells below and to the right, unless they are occupied (`#SPILL!`). `A1#` refers to the spill range. Legacy Ctrl+Shift+Enter arrays are supported too.
- **The grid is a canvas.** A transparent native scroller above the canvas provides scrolling, touch panning and inertia. Only visible cells are painted, so million-row sheets scroll smoothly. Charts and pictures are DOM/SVG objects in an overlay.
- **Styles are interned.** Cells store a style index; identical formats share one entry, which keeps memory small and makes .xlsx style tables straightforward.
- **Recovery is a pipeline.** Damaged files go through identification, normal open, ZIP salvage, XML repair, raw sheet scanning and BIFF record salvage; the best result is opened as a new, unsaved workbook with a report.
