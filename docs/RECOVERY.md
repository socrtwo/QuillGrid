# Opening damaged workbooks

QuillLattice includes a recovery engine for damaged spreadsheets. It descends from the recovery tools by Paul D. Pruitt (socrtwo): the byte-level ZIP salvage and ImmortalInflate decoder of [Immortal Unzipper / Immortal Zip](https://socrtwo.github.io/Immortal-Zip/), the XML healing and strings scan of [Word Recovery (socrtwo/wordrecovery-SF)](https://github.com/socrtwo/wordrecovery-SF), and the S2 File Identifier. QuillForge, the sibling word processor, uses the same generic pieces for documents. The spreadsheet-specific methods follow the strategies of the author's Excel tools ([S2 Recovery Tools for Microsoft Excel](https://socrtwo.github.io/corruptexcelrec-SF/), [Excel Recovery Addin](https://socrtwo.github.io/excelrcvryaddin-SF/)), reimplemented here from the file format specifications.

Recovery runs automatically when an .xlsx, .xlsm, .xltx, .ods or .xls file fails to open normally. You can also start it on purpose:

- **File ▸ Open ▸ Open and Repair…** runs every recovery method, even on a file that opens, and shows the report.
- **File ▸ Open ▸ Recover Data from Any File…** opens the manual recovery guide; its **Recover a file…** button runs every method on any file, including renamed, truncated or unknown data.
- **File ▸ Recovery Report** shows how the current workbook was recovered.
- **File ▸ Manual Recovery Options** is a guide for each platform: AutoRecover folders, version history, Shadow Copies, Time Machine, command-line tools, and the other S2 recovery tools.

A recovered workbook opens as "*name* (Recovered)". **Save** asks for a new name, so the damaged original is never overwritten.

## The pipeline

`recoverWorkbook()` in `src/io/recover/index.ts` runs the methods in order. Each one that produces a workbook is a candidate, and the candidate with the most recovered cells wins. Ties go to the normal importer, then to the repaired package. The engine never throws for damaged input: when nothing works, the report has no workbook and says why.

| Step | Method | Heritage |
|---|---|---|
| a | **Identification:** the S2 File Identifier finds the real type from the bytes, whatever the extension says (xlsx, xlsm, ods, flat ods, xls including raw BIFF2–4, csv, html, Excel 2003 XML, QuillLattice JSON). It separates concatenated or embedded data (for example a PDF fused onto a workbook) and reports encrypted Office packages (a compound file with `EncryptionInfo`/`EncryptedPackage`) | Word Recovery (`s2-file-id.js`) |
| b | **Normal open:** the regular importer for the identified type. If it opens the file but the package has CRC mismatches or malformed parts, the repair methods run as well, and the normal result competes at a reduced weight, because a lenient importer can drop data silently | New |
| c | **ZIP salvage:** walk the `PK\x03\x04` local headers and ignore the central directory. Inflate each entry with **ImmortalInflate**, a decoder that never throws, at bit shifts 0–47. Check every CRC-32 (damaged deflate data often inflates without an error) and resynchronize on the next deflate block after damage | Immortal Unzipper, Word Recovery |
| d | **XML repair and rebuild:** see below. The repaired parts are re-zipped into a clean package and opened with the normal importer, so formatting survives wherever its parts did | Word Recovery (healing), new (spreadsheet fixes) |
| e | **Raw worksheet scan:** scan the worksheet XML, even fragments, for `<c r=".." t="..">` cells with `<v>`, `<f>` and inline strings. Shared strings are mapped through whatever survived of `sharedStrings.xml`. Shared formulas are expanded by shifting the master formula's relative references. The workbook is built directly, which rescues data from badly truncated files | New |
| f | **Excel 97–2003 record salvage:** the BIFF reader in lenient mode; see below | New |
| g | **OpenDocument salvage:** heal `content.xml`, cut it at the last complete `</table:table-row>` before damage, rebuild the package (`mimetype` stored first, manifest regenerated) and open it. Otherwise a raw scan of `<table:table-cell>` elements, including repeated rows and columns, ODF formulas and merged cells | Word Recovery (healing), new |
| h | **Text fallbacks:** delimited text (delimiter detection, quoted fields), HTML tables (colspan and rowspan become merges), Excel 2003 XML (healed with the tag-soup rebuilder, R1C1 formulas converted to A1) | New |
| i | **Strings scan:** printable ASCII and UTF-16LE runs from the raw bytes, or the text of whatever package parts survived. Lines that share a delimiter are split into columns. The report labels the result as text-only recovery | Word Recovery |

### XML repair (step d)

- Strip characters that XML forbids; heal with `healXMLStrict` (the original Word Recovery healer), then rebuild tag soup with `balanceXml` when healing is not enough.
- **Damage cut:** in a worksheet known to be damaged (CRC mismatch or a decode error), keep the consistent prefix up to the last complete `</row>` before the first inconsistency (`findDamageCut`). LZ77 copies damage forward, so everything after that point is unreliable. Cells decoded after the damage that claim rows below the cut go on a separate sheet named "*sheet* (unverified)", and the report says so.
- Close `sheetData` and `worksheet` in a truncated part; drop `<c>` elements whose markup is broken or whose value does not match their type; renumber rows and cells whose `r` attribute is missing, duplicated, out of order or in the wrong row.
- `sharedStrings.xml`: rebuild from the complete `<si>` items, keep the text of a truncated last item, exclude phonetic runs, and correct `count`/`uniqueCount`.
- `styles.xml`: heal it; if it still cannot serve every style index the sheets use, replace it with a minimal valid stylesheet and remove the cells' style references. Values are kept; formatting, including date formats, is lost, and the report says so.
- Relationships: remove relationships whose targets are missing (drawings, comments, tables, VML) and the elements that point at them (`<drawing>`, `<legacyDrawing>`, `<tablePart>`, …). Remove the stale `calcChain.xml`.
- Package structure: regenerate `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml` and `xl/_rels/workbook.xml.rels` from the parts found. Worksheets are discovered by name (`xl/worksheets/sheet*.xml`), sheet names come from the surviving workbook part when possible, and orphan worksheet parts are added.

The rebuilt package is offered as **Save repaired package (.xlsx)** in the report.

## Excel 97–2003 (.xls)

`src/io/recover/xls.ts` is a full BIFF8 reader used both for normal opening (`importXls`) and for salvage (`readBiff(bytes, { lenient: true })`). It reads the compound file with the CFB reader shared with Word 97 recovery (`readCfb` in `doc97.ts`) and walks the `Workbook` (BIFF8) or `Book` (BIFF5) stream.

- **Globals:** BOF/EOF, BOUNDSHEET (names, hidden and very hidden sheets, chart and macro sheets), SST with CONTINUE records (strings split across records, with the 8-bit/16-bit flag re-read at each boundary, rich-text and phonetic runs skipped), FONT, FORMAT, XF, PALETTE, DATEMODE (1904 dates), CODEPAGE, NAME (defined names, print areas and print titles), SUPBOOK, EXTERNSHEET, EXTERNNAME, PROTECT/PASSWORD. FILEPASS (encryption) stops with a clear message.
- **Cells:** LABELSST, LABEL, RSTRING, NUMBER, RK, MULRK, BOOLERR, BLANK, MULBLANK, and FORMULA with its cached number, string (from the following STRING record), Boolean or error.
- **Formulas:** the Ptg token stream is turned back into formula text: references and areas (absolute and relative, whole rows and columns), `ptgRefN`/`ptgAreaN` in shared formulas (SHRFMLA) relative to each cell, array formulas (ARRAY), 3D references through EXTERNSHEET/SUPBOOK (including external workbooks), numbers, strings, Booleans, errors, array constants, all operators, parentheses, missing arguments, `ptgAttr` (SUM, IF, CHOOSE, spaces), `ptgFunc`/`ptgFuncVar` with the standard table of about 370 built-in functions, defined names (`ptgName`), and add-in or newer functions (`ptgNameX`, with the `_xlfn.` prefix removed). A formula that cannot be decoded keeps its cached value.
- **Formatting and layout:** number formats, font name, size, bold, italic, underline, strikethrough, superscript and subscript, font color (default palette or PALETTE), horizontal and vertical alignment, wrap, shrink, indent, rotation, borders, solid and pattern fills, locked and hidden flags, column widths and hidden columns, row heights and hidden rows, outline levels, merged cells, frozen panes (WINDOW2/PANE), zoom, gridlines and zero display, comments (NOTE with OBJ/TXO text), hyperlinks (HLINK URL, file and location monikers), and sheet protection.
- **Excel 5.0/95 (BIFF5/7):** the same records with byte strings decoded through the CODEPAGE record (Windows-125x, Mac Roman, Shift-JIS and others). The Windows-1252 range 0x80–0x9F is mapped explicitly, because some runtimes decode that label as ISO-8859-1.
- **BIFF2–4:** single-sheet files are read best effort (cells, simple formulas, fonts and formats).

**Salvage (step f):** in lenient mode, a record with an impossible length, or an unknown identifier followed by garbage, starts a resynchronization: the reader scans forward for a position where a known record header leads to two more known headers, and continues there. When the compound-file header, FAT or directory is damaged, the raw file bytes are scanned for BIFF BOF records (`09 08` with version 0x0600 or 0x0500) and the record stream is walked from there. When the compound file is damaged, both results are compared and the one with more cells is kept. Excel writes streams in contiguous sectors, so this recovers most workbooks whose container is broken.

## What was tested

`tests/unit/recover.test.ts` builds its fixtures in the test: a valid two-sheet XLSX made with JSZip (600 rows of shared strings, numbers, formulas and inline strings, a merge and a frozen pane), a BIFF8 workbook written record by record inside a minimal compound file, a BIFF5 workbook, a raw BIFF2 sheet, and an ODS. It then damages them and checks the recovered values:

- a missing central directory; a missing `[Content_Types].xml`; missing workbook part and relationships; a dangling drawing relationship (every value recovered);
- malformed `sharedStrings.xml` (stray `<` and `&`, a lost end tag) and a truncated last shared string;
- an unusable `styles.xml` (values kept, minimal stylesheet);
- packages truncated at 80% and 60% (a complete prefix of rows, no wrong values, the sheet cut at the last complete row);
- bytes flipped in the middle of a worksheet's deflate stream (the rows before the damage; the result is marked unverified);
- leading and trailing foreign data, a wrong extension, an empty file, random bytes, and an encrypted OOXML package;
- the BIFF8 reader: split SST strings with a width change at the CONTINUE boundary, rich strings, RK/MULRK, every cached-result type, formulas with references, areas, `ptgAttr` SUM, 3D references, names, IF via `ptgFuncVar`, ROUND via `ptgFunc`, shared formulas, styles, merges, widths, heights, frozen panes, a comment, a hyperlink, a hidden sheet and a print area;
- `.xls` salvage after the compound-file header is destroyed, after a record length is corrupted (resynchronization), and from a truncated file; a FILEPASS (encrypted) workbook;
- a truncated ODS, CSV, HTML and truncated Excel 2003 XML.

The tests pass both with and without the regular XLSX and ODS importers available, because the raw scans build workbooks directly.

## Limits

- **Damaged deflate data:** cells after the damage point in the same deflate stream are unreliable. They are kept on a separate "(unverified)" sheet when their addresses do not collide with the recovered rows, and the report asks you to check them. Data in later, undamaged deflate blocks is recovered when the decoder can resynchronize. The last rows before the cut may still contain damaged values.
- **Styles:** when `styles.xml` is replaced, number formats are lost, so dates appear as serial numbers until you format them again.
- **.xls:** charts, pictures, pivot tables, conditional formatting, data validation and VBA are not read. BIFF2–4 support is best effort, and BIFF5 external and add-in names are resolved best effort.
- **Password-protected files** (encrypted .xls via FILEPASS, or encrypted .xlsx packages) are detected and reported; they cannot be decrypted.
- **Strings scan:** text-only recovery loses cell positions, formulas and formatting.
