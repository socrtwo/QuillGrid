# File formats

QuillGrid opens files by what they contain, not by their name. A file called `budget.xls` that
is really a Lotus 1-2-3 worksheet opens as a 1-2-3 worksheet, and a file with no extension at all
opens like any other. When the contents and the extension disagree, or the extension is missing,
a short note says what the file turned out to be.

## How files are identified

`src/io/sniff.ts` holds a table of detection rules. Each rule looks at the bytes (signatures,
record structure, ZIP and compound-file directories, text patterns) and returns a confidence
between 0 and 1. The extension is used only to break a tie between rules that score about the
same, and as the last resort when no rule recognizes the content. The rules of the modern, XML
and text formats (`SNIFF_RULES_B` in `src/io/sniff-b.ts`) are registered in the same table, so
there is one detector. Other modules add rules with `registerSniffer(rule)`; a rule returns `{ format, confidence, category, label }` where category is
`spreadsheet`, `document`, `presentation`, `email` or `unknown`.

The result decides what happens next:

- **spreadsheet**: the matching importer opens it.
- **document, presentation or email**: QuillGrid offers to hand the file to the Quill app made
  for it (QuillPen for documents, QuillStage for presentations, QuillPost for mail) and always
  offers to import it instead (see below). On Android the sibling app is launched by package
  name; on the desktop the file is opened with the program registered for its type; a web page
  and iOS cannot hand files to other apps, so the dialog explains how to open it there.
- **recognized but not readable** (for example an Outlook .pst data file): a message names the
  format instead of showing garbage.

## Legacy and binary spreadsheet formats

Fidelity tiers: **full** (everything the format stores that QuillGrid can represent),
**high** (values, formulas and most formatting), **values + formulas**, **values only**.

| Program | Versions | Era | Extensions | Detection signature | Fidelity | Limitations |
|---|---|---|---|---|---|---|
| Microsoft Excel | 2.x (BIFF2) | 1987 | .xls | BOF record 0x0009 at byte 0 | high | Font colors (FONTCOLOR) and IXFE-indexed styles are not read |
| Microsoft Excel | 3.0 (BIFF3) | 1990 | .xls | BOF 0x0209 | high | Cell background colors of BIFF3 XF records are not read |
| Microsoft Excel | 4.0 (BIFF4), 4.0 workbook (BIFF4W) | 1992 | .xls, .xlw | BOF 0x0409; workbook when the BOF type is 0x0100, sheets listed in BUNDLESHEET (0x008F) records | high | Charts and macro sheets are skipped |
| Microsoft Excel | 5.0 and 95 (BIFF5/BIFF7) | 1993-1995 | .xls | OLE2 compound file with a `Book` stream | high | Same reader as Excel 97-2003; charts, pictures and VBA are not opened |
| Microsoft Excel | 97-2003 (BIFF8) | 1997-2003 | .xls, .xlt | OLE2 compound file with a `Workbook` stream | high | See the README |
| Microsoft Works | 6-9 spreadsheet | 2001-2007 | .xlr | OLE2 compound file whose stream starts with a BIFF BOF record | high | As Excel |
| Microsoft Works | 2-4 for DOS spreadsheet | 1987-1994 | .wks | Lotus BOF 0x0404 followed by 0x54xx records | high | Font colors are not read |
| Microsoft Works | 3-4 for Windows spreadsheet | 1993-1995 | .wks | First record 0x00FF with version 0x0404 | high | Number formats, bold, italic, underline and alignment are kept; font colors are not read |
| Lotus 1-2-3 | Release 1A | 1983 | .wks | BOF record (type 0, 2 bytes) version 0x0404 | full | |
| Lotus 1-2-3 | Release 2.x | 1985-1993 | .wk1 | BOF version 0x0406 | full | Allways/WYSIWYG .fmt files are not applied |
| Lotus Symphony | 1.0-2.x | 1984-1990 | .wrk, .wr1 | BOF version 0x0405/0x0406 | values + formulas | Symphony windows, forms and documents are ignored |
| Lotus 1-2-3 | Release 3.x | 1989-1992 | .wk3 | BOF (26 bytes) version 0x1000-0x1002 | high | Several sheets, sheet names, number formats (record 0x13 row runs), column widths, hidden columns, range names and cell notes are kept; fonts and colors (and .fm3 files) are not applied |
| Lotus 1-2-3 | Release 4 and 5 for Windows | 1993-1994 | .wk4 | BOF version 0x1003-0x1004 | values + formulas | As Release 3, but number formats stored in the Release 4 style records are not decoded |
| Lotus 1-2-3 | 97, Millennium Edition | 1997-2002 | .123 | BOF version 0x1005-0x1007 | values + formulas | As Release 3; pictures and charts are skipped |
| Lotus 1-2-3 | formatting files | 1989-1993 | .fmt, .fm3 | BOF version 0x8000-0x8007 | recognized | Not applied; open the worksheet itself |
| Quattro, Quattro Pro for DOS | 1-5 | 1987-1993 | .wq1, .wq2 | BOF version 0x5120 / 0x5121 | high | Quattro-only functions are kept as values with a note |
| Quattro Pro for Windows | 1-6 | 1992-1996 | .wb1, .wb2, .wb3 | BOF (2 bytes) version 0x1001-0x1007 | high | Several pages with their names; formulas with page and block references; number formats, fonts and alignment from the notebook styles; colors, borders and user-defined formats are not decoded |
| Quattro Pro | 7 and later | 1996- | .qpw | OLE2 compound file with `NativeContent_MAIN` or `PerfectOffice_MAIN` | values only | Text only, one item per row: the record layout is not public |
| dBASE, FoxBASE, FoxPro, Visual FoxPro, Clipper | dBASE II to 7, FoxPro 2, VFP 3-9 | 1980-2007 | .dbf (.dbt, .fpt memos) | Version byte, valid date, header and record sizes that match the file length | full | Memo text needs the .dbt/.fpt file passed with the table (`importDbf(..., { memo })`); opened alone, memo cells say "(memo)" |
| VisiCalc | VisiCalc, Advanced VisiCalc | 1979-1983 | .vc | Text lines `>A1:` with cell contents, then `/W`, `/G` commands | full | Formulas are translated with VisiCalc's left-to-right evaluation |
| Multiplan | 1-4 (SYLK) | 1982-1990 | .slk, .mp | `ID;PMP` | high | Read by the SYLK importer |
| Multiplan | binary | 1982-1990 | .mp | `0C ED` at byte 0 | values only | Text only: the format is not documented |
| SuperCalc | 3-5 | 1983-1990 | .cal | "SuperCalc" in the header | values only | Text only: the format is not documented |
| AppleWorks, ClarisWorks | 4-6 | 1993-2004 | .cwk | `BOBO` at byte 4 | values only | Text and numbers in storage order in column A; cell positions, formulas and formatting are lost |

Formulas are decompiled to their original 1-2-3 syntax and translated by the 1-2-3 formula
translator (`src/heritage/lotus.ts`), so 1-2-3 semantics are kept: zero-based offsets in @CHOOSE,
@INDEX, @VLOOKUP, @MID and @FIND, argument orders of the financial functions, and labels counting
as zero in @AVG, @COUNT and the other statistical functions. A formula that uses a function
QuillGrid does not have keeps its last calculated value and gets a cell note with the original
formula; the import notes list how many there were.

Damaged files are read again leniently (records with impossible lengths are skipped and reading
resynchronizes on the next plausible record), then by the recovery pipeline, whose strings scan
keeps whatever text and numbers can be found.

Sources: the record layouts of 1-2-3 Release 1A and 2 come from Lotus's public-domain
"Worksheet File Format" specification. The BIFF2-5 layouts come from publicly available descriptions of
the Excel binary formats. The Release 3+ and Quattro Pro for Windows layouts were established
by inspecting files and checking the readers' results against LibreOffice's importers on the same
files; no code from other spreadsheet projects was used.

## Documents, presentations and e-mail

| Kind | Formats | What the import produces |
|---|---|---|
| Word processing | DOCX/DOCM/DOTX, Word 2003 XML, ODT/OTT/FODT, SXW, RTF, DOC (Word 97-2003 and 6/95), Word for DOS, Windows Write, WordPerfect, Works, EPUB, PDF, Apple Pages, others | A **Text** sheet with one paragraph per row in column A, and one sheet per table (**Table 1**, ...) with merged cells, bold header cells and numbers as values. DOC, WordPerfect and PDF are read for their text: runs of tab-separated lines become tables |
| Presentations | PPTX/PPSX/POTX, ODP/FODP, SXI, PPT | One sheet per slide: the title (bold), the text in column A, tables below the text, speaker notes at the end. PPT is read for the text of its placeholders |
| E-mail | EML, EMLX, MHT, MBOX, MSG | A **Message** sheet with From, To, Cc, Date, Subject and Attachments, then the body one line per row; MBOX gives a **Messages** sheet with a row per message. CSV and TSV attachments become sheets, spreadsheet attachments (XLSX, XLS, ODS, WK1, DBF, ...) are opened and their sheets added, and tables in an HTML body are added as sheets |
| Recognized only | PST/OST, DBX, winmail.dat, HWP 5, Access | A message names the format |

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
