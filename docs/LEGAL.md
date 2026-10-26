# Clean-room, trademark and licensing notes

This file records how Anne's Grid provides a complete spreadsheet command set while avoiding other vendors' trademarks and copyrights. It summarizes the project's practice; it is not legal advice.

## What is reproduced, and why it is safe

- **Functionality, function names and command names.** Spreadsheet operations (Sort, Filter, Merge & Center, Freeze Panes, PivotTable, Goal Seek) and worksheet function names (SUM, VLOOKUP, XLOOKUP, NETWORKDAYS.INTL) are functional interfaces. Workbooks written by any spreadsheet program use these names in their formulas, and compatibility requires the same names and semantics. Functionality is not protected by copyright, and descriptive names are not trademarks for those functions.
- **Tab and group organization.** Grouping commands under tabs such as Home, Insert, Formulas and Data is a common, functional arrangement. Anne's Grid's grouping, ordering, labels, tooltips and descriptions were written for this project.
- **File formats.** .xlsx (ECMA-376 / ISO/IEC 29500 Office Open XML), .ods (ISO/IEC 26300 OpenDocument), CSV (RFC 4180), HTML, SYLK, DIF and the legacy binary .xls format (published by Microsoft as the [MS-XLS] Open Specification) are public specifications. The converters were written from those specifications and tested with open-source tools (LibreOffice).
- **Function semantics.** The worksheet functions were implemented from their published behavior and standard numerical methods. Function descriptions in the Insert Function dialog are original text.

## What is deliberately not reproduced

- No vendor logos, icons, splash screens, color schemes or trade dress. Command icons are Google's Material Symbols (Apache 2.0). The app icon (the Anne's Office suite tile: a white quill over a spreadsheet card) and visual design are original.
- No vendor product names in the app's name or interface. Trademarked feature names are replaced with generic terms:

| Vendor term | Anne's Grid term |
|---|---|
| WordArt | Decorative Text |
| SmartArt | Diagram |
| Flash Fill | Flash Fill (a generic description of fill-by-example) |
| Power Query / Get & Transform | Get Data and Transform |
| Ideas / Analyze Data / Copilot | Analyze Data (AI tab) |
| Tell Me / Search | Search Commands |
| Ribbon | Tabs / ribbon used generically |

- No copied help text, templates, sample workbooks, clip art or dictionaries. Templates and their sample data are original. The spelling dictionaries are SCOWL English (MIT and BSD-style) and other freely licensed Hunspell dictionaries downloaded on request.
- No proprietary fonts are bundled. Workbooks that name Calibri, Cambria, Times New Roman, Arial, Courier New, Georgia or Garamond render with metric-compatible open fonts: Carlito, Caladea, Tinos, Arimo, Cousine, Gelasio and EB Garamond (SIL OFL / Apache 2.0). A locally installed original is preferred when present.
- References to other products ("inspired by LibreOffice Calc's Navigator") are nominative use describing where an idea came from. They do not imply endorsement.

## Third-party components

| Component | License |
|---|---|
| JSZip | MIT (dual MIT/GPLv3; used under MIT) |
| nspell | MIT |
| hunspell-asm | MIT |
| dictionary-en, dictionary-en-gb, dictionary-en-ca, dictionary-en-au (SCOWL) | MIT and BSD-style |
| Material Symbols | Apache 2.0 |
| Capacitor | MIT |
| Electron (desktop builds) | MIT |
| Fonts: Carlito, Caladea, Gelasio, EB Garamond, Courier Prime, Libre Baskerville, Lato, Merriweather, Open Sans | SIL Open Font License 1.1 |
| Fonts: Tinos, Arimo, Cousine | Apache 2.0 |

Online services are called only when the user invokes the matching command: Datamuse (thesaurus), the chosen AI provider (AI tab), Google Translate and web search (opened in the browser), and dictionary downloads (Review ▸ Spelling Dictionaries).
