# File formats

## Modern, XML and text spreadsheet formats

Fidelity tiers: **full** (everything the format stores that QuillGrid can represent),
**high** (values, formulas and most formatting), **values + formulas**, **values only**.
Detection rules for these formats are in `src/io/sniff-b.ts` (`SNIFF_RULES_B`); each importer
is its own module under `src/io/`. Every format here opens by content, whatever the file is
called; the extension only breaks ties.

| Program | Versions | Era | Extensions | Detection signature | Fidelity | Limitations |
|---|---|---|---|---|---|---|
| Microsoft Excel (Office Open XML) | 2007 and later, Excel for Mac 2008+, Excel Online and Google Sheets exports | 2007- | .xlsx, .xlsm, .xltx, .xltm | ZIP with `[Content_Types].xml` and `xl/` parts | full | Pivot tables open as values; SmartArt, slicers, sparkline axes, form controls' VBA and Power Query connections are not opened; VBA is kept but not run |
| Microsoft Excel (ISO/IEC 29500 Strict) | 2010 and later ("Strict Open XML Spreadsheet") | 2010- | .xlsx | Same package; `purl.oclc.org/ooxml` namespaces | full | As above |
| Microsoft Excel Binary Workbook | 2007 and later | 2007- | .xlsb | ZIP package with `xl/workbook.bin` | high | Conditional formats and data validation stored in BIFF12 records are not opened (a note says so); pivot tables open as values; structured references and table parts follow the published record layouts but were checked only against files built from them |
| Microsoft Excel XML Spreadsheet | 2002, 2003, Excel 2004 for Mac | 2001-2004 | .xml | `<Workbook>` in `urn:schemas-microsoft-com:office:spreadsheet`, `<?mso-application progid="Excel.Sheet"?>` | high | Charts, pictures and PivotTables are not part of the format; conditional formats with icon sets did not exist yet |
| Microsoft Excel web page | 97-2019 | 1997- | .htm, .html | `<table>` with `urn:schemas-microsoft-com:office:excel` attributes | high | Style sheet classes, `mso-number-format`, `x:num`, `x:fmla`, `x:bool`, `x:err`, column widths, merges; charts saved as pictures are not opened |
| Microsoft Excel single file web page | 2000-2019 | 1999- | .mht, .mhtml | MIME header with `multipart/related` | high | Each sheet part opens as a sheet named from the tab strip; as Excel web page |
| LibreOffice / OpenOffice Calc (OpenDocument) | OpenOffice.org 2+, LibreOffice, Calligra, Google Sheets and Excel exports | 2005- | .ods, .ots | ZIP with `mimetype` `application/vnd.oasis.opendocument.spreadsheet` | full | Charts open with series, titles, legend and axes (3D and stock styling are simplified); pivot tables (data pilots) open as values; form controls and macros are not opened |
| LibreOffice Calc flat XML | OpenOffice.org 2+, LibreOffice | 2005- | .fods | `<office:document>` with the spreadsheet MIME type | full | As .ods; charts are read from the inline objects |
| StarOffice / OpenOffice.org Calc | StarOffice 6-7, OpenOffice.org 1.0-1.1 | 2001-2005 | .sxc, .stc | ZIP with `mimetype` `application/vnd.sun.xml.calc` | high | Converted to OpenDocument and read by the .ods reader; charts (embedded objects) and DataPilots are not opened |
| Uniform Office Format (UOF) | RedOffice, Yozo/EIOffice, WPS, LibreOffice UOF filter | 2007- | .uos, .uof | `<uof:UOF>` root with `vnd.uof.spreadsheet` | values + formulas | Values, formulas, cell styles, widths, heights, merges, links and comments; charts and conditional formats are not opened |
| Gnumeric | 1.0-1.12 | 2001- | .gnumeric, .gnm | gzip with a Gnumeric name, or XML in `http://www.gnumeric.org/v10.dtd` (old files: `gmr:` and `www.gnome.org/gnumeric`) | high | Formulas are recalculated when the file opens (Gnumeric stores no results); charts, solver and scenario settings are not opened |
| KSpread (KOffice) | 1.x | 1998-2009 | .ksp | KoStore ZIP with mimetype `application/x-kspread`, KOffice 1.0-1.2 tar.gz store, or `<spreadsheet>` maindoc.xml | values + formulas | Basic formatting, widths, heights, merges and comments; charts and embedded objects are not opened. Calligra Sheets 2+ saves .ods |
| Apple Numbers | '09 (1.x-2.x) | 2007-2012 | .numbers | ZIP or package with `index.xml[.gz]`, or `<ls:document>` XML | high | Each table opens as a sheet named after it; formulas are converted (references to other tables, whole-column ranges); charts are not opened; pop-up menus and checkboxes open as their values |
| Apple Numbers | 3.0 and later (2013-) | 2013- | .numbers | ZIP or package with `Index/*.iwa` (or an inner `Index.zip`) | values only | Values, dates, durations, rich text, widths, heights, hidden columns, date formats; formulas open as their last values (the expression trees are not decoded); cell styles, merges and charts are not opened. A password-protected document or one holding only a preview opens the preview picture with a note |
| Google Sheets | | 2006- | .xlsx, .ods, .csv, .html | (exports) | as those formats | HTML copied from Google Sheets keeps values (`data-sheets-value`) and formulas (`data-sheets-formula`) |
| Delimited text (CSV, TSV) | any | 1972- | .csv, .tsv, .tab, .txt | No other rule matches; delimiter found by consistency across lines (`,` `;` tab `|`), Excel's `sep=` line | values only | Encoding detected: UTF-8, UTF-16 (with or without byte order mark), Windows-1250 to 1258, Mac OS Roman, DOS 437 and 850, KOI8-R, Shift-JIS, EUC-JP, GBK, Big5, EUC-KR; decimal comma and day-first dates detected from the data |
| Fixed-width text | Excel "Formatted Text (Space delimited)", report listings | 1980s- | .prn, .txt | Blank columns line up in every line and no delimiter is used consistently | values only | Columns that do not line up fall back to delimited text |
| SYLK | Multiplan, Excel, LibreOffice | 1981- | .slk, .sylk | `ID;` first line | high | Values, R1C1 or A1 formulas, number formats, fonts, alignment, borders, widths; escapes for accents and line breaks |
| DIF | VisiCalc, Lotus, Excel, LibreOffice | 1980- | .dif | `TABLE` / `0,1` header with `VECTORS`, `TUPLES` or `DATA` | values only | Dates written as text become dates; column-major files are transposed |
| JSON | records, row arrays, column objects, wrapped arrays | 2001- | .json | Top-level `[` or `{` | values only | Nested objects flatten to dotted columns; arrays of scalars join; ISO dates become dates; each array of records inside an object becomes a sheet. QuillGrid's own JSON is still recognized |
| JSON Lines | NDJSON | 2013- | .ndjson, .jsonl | Several lines that each parse as JSON | values only | Lines that are not JSON are skipped with a note |
| SQLite | 3.x | 2004- | .sqlite, .sqlite3, .db, .db3 | `SQLite format 3\0` | values only | Read-only, by a small reader of the published file format (no SQL engine): every table opens as a sheet; WITHOUT ROWID tables are supported; views and virtual tables have no stored rows and are listed in a note; changes still in a `-wal` file are not seen; blobs show their size |
| gzip | any of the above compressed | | .gz | `1F 8B 08` | as the contents | The contents are identified again by what they are |

Not supported, and why:

- **PlanPerfect** (WordPerfect Corporation, .pln, 1980s-1990s): no published description of the
  file format exists to build a reader from, and no sample files were available to study.
- **Lotus 1-2-3 / SmartSuite XML** and **Quattro Pro XML**: no such spreadsheet XML formats were
  shipped by these programs; their binary formats are covered in the legacy section.
- **Parquet and Arrow**: data-engineering column stores, not spreadsheet program files.
- **Apple Numbers formula expressions (2013+)**: stored as calculation-engine trees whose layout
  is not published; the last calculated values are shown instead.
