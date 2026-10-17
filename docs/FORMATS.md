# File formats

QuillGrid opens files by what they contain, not by their name. A file called `budget.xls` that
is really a Lotus 1-2-3 worksheet opens as a 1-2-3 worksheet, and a file with no extension at all
opens like any other. When the contents and the extension disagree, or the extension is missing,
a short note says what the file turned out to be.

## How files are identified

`src/io/sniff.ts` holds a table of detection rules. Each rule looks at the bytes (signatures,
record structure, ZIP and compound-file directories, text patterns) and returns a confidence
between 0 and 1. The extension is used only to break a tie between rules that score about the
same, and as the last resort when no rule recognizes the content. Other modules add rules with
`registerSniffer(rule)`; a rule returns `{ format, confidence, category, label }` where category is
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
| Microsoft Works | 3-4 for Windows spreadsheet | 1993-1995 | .wks | First record 0x00FF with version 0x0404 | values + formulas | Bold, italic, underline and alignment are kept; number formats are not decoded |
| Lotus 1-2-3 | Release 1A | 1983 | .wks | BOF record (type 0, 2 bytes) version 0x0404 | full | |
| Lotus 1-2-3 | Release 2.x | 1985-1993 | .wk1 | BOF version 0x0406 | full | Allways/WYSIWYG .fmt files are not applied |
| Lotus Symphony | 1.0-2.x | 1984-1990 | .wrk, .wr1 | BOF version 0x0405/0x0406 | values + formulas | Symphony windows, forms and documents are ignored |
| Lotus 1-2-3 | Release 3.x | 1989-1992 | .wk3 | BOF (26 bytes) version 0x1000-0x1002 | values + formulas | Several sheets, sheet names, column widths, hidden columns, range names and cell notes are kept; number formats and fonts (in the record 0x13 runs and .fm3 files) are not applied |
| Lotus 1-2-3 | Release 4 and 5 for Windows | 1993-1994 | .wk4 | BOF version 0x1003-0x1004 | values + formulas | As Release 3 |
| Lotus 1-2-3 | 97, Millennium Edition | 1997-2002 | .123 | BOF version 0x1005-0x1007 | values + formulas | As Release 3; pictures and charts are skipped |
| Lotus 1-2-3 | formatting files | 1989-1993 | .fmt, .fm3 | BOF version 0x8000-0x8007 | recognized | Not applied; open the worksheet itself |
| Quattro, Quattro Pro for DOS | 1-5 | 1987-1993 | .wq1, .wq2 | BOF version 0x5120 / 0x5121 | high | Quattro-only functions are kept as values with a note |
| Quattro Pro for Windows | 1-6 | 1992-1996 | .wb1, .wb2, .wb3 | BOF (2 bytes) version 0x1001-0x1007 | values + formulas | Several pages with their names; formulas with page and block references; styles and number formats are not decoded |
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
| Word processing | DOCX/DOCM/DOTX, Word 2003 XML, ODT/OTT/FODT, SXW, RTF, DOC (Word 97-2003 and 6/95), Word for DOS, Windows Write, WordPerfect, Works, EPUB, PDF, others | A **Text** sheet with one paragraph per row in column A, and one sheet per table (**Table 1**, ...) with merged cells, bold header cells and numbers as values. DOC, WordPerfect and PDF are read for their text: runs of tab-separated lines become tables |
| Presentations | PPTX/PPSX/POTX, ODP/FODP, SXI, PPT | One sheet per slide: the title (bold), the text in column A, tables below the text, speaker notes at the end. PPT is read for the text of its placeholders |
| E-mail | EML, EMLX, MHT, MBOX, MSG | A **Message** sheet with From, To, Cc, Date, Subject and Attachments, then the body one line per row; MBOX gives a **Messages** sheet with a row per message. CSV and TSV attachments become sheets, spreadsheet attachments (XLSX, XLS, ODS, WK1, DBF, ...) are opened and their sheets added, and tables in an HTML body are added as sheets |
| Recognized only | PST/OST, DBX, winmail.dat, Pages/Keynote/Numbers 2013+, HWP 5, Access | A message names the format |
