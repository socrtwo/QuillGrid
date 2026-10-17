# Test fixtures

| Folder / file | Source | License |
|---|---|---|
| `io/libreoffice.ods`, `io/libreoffice-roundtrip.xlsx` | Written for these tests with LibreOffice | Same as the project (MIT) |
| `legacy/testEXCEL_3.xls`, `legacy/testEXCEL_4.xls`, `legacy/testEXCEL_5.xls`, `legacy/testEXCEL_95.xls` | Apache POI test data (`test-data/spreadsheet/`, https://github.com/apache/poi): Excel 3.0 (BIFF3), Excel 4.0 (BIFF4), Excel 5.0 and Excel 95 (BIFF5) worksheets | Apache License 2.0 |
| `documents/quick.msg` | Apache POI test data (`test-data/hsmf/quick.msg`) | Apache License 2.0 |
| `documents/doc.docx`, `doc.odt`, `doc.rtf`, `doc.doc` | Written for these tests: a small HTML page converted with LibreOffice 24.2 | Same as the project (MIT) |
| `documents/deck.pptx`, `deck.odp` | Written for these tests with python-pptx; the ODP converted with LibreOffice 24.2 | Same as the project (MIT) |
| `documents/mail.eml` | Written for these tests | Same as the project (MIT) |

The other legacy formats (Lotus 1-2-3 WKS/WK1/WK3/123, Works, Quattro WQ1/WB1, Excel 2.x and 4.0
workbooks, dBASE, VisiCalc, Outlook MSG, PowerPoint 97-2003) are built record by record by
`tests/unit/legacy-builders.ts`; nothing on a build machine writes them. While the readers were
written, the same files and LibreOffice's own sample files were opened in LibreOffice 24.2 and the
values compared cell by cell.
