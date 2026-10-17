# Test fixtures

## io/

- `libreoffice-roundtrip.xlsx`, `libreoffice.ods`: QuillGrid exports re-saved by LibreOffice.

## io/b/ (modern, XML and text spreadsheet formats)

Original test data written for this project:

- `sample.gnumeric` — hand-written Gnumeric XML (gzip), following the Gnumeric file format.
- `sample.sxc` — OpenOffice.org 1.0 Calc package built by `make_fixtures.py`.
- `sample.ksp`, `sample-old.ksp` — KSpread documents (KoStore ZIP and KOffice 1.x tar.gz) built by `make_fixtures.py`.
- `sample.uos` — Uniform Office Format spreadsheet written by LibreOffice 24.2 from `sample.sxc`, with two
  formulas added and the second sheet's row numbers corrected by hand (LibreOffice's UOF export drops
  formulas and numbers rows on from the previous sheet).
- `sample.sqlite` — SQLite database built by `make_fixtures.py` with Python's sqlite3 module.
- `encodings/*.txt` — short CSV texts encoded in legacy code pages (Windows-125x, Mac OS Roman,
  DOS 437/850, KOI8-R, Shift-JIS, GBK, Big5, EUC-KR).

The .xlsb and Numbers 2013+ (IWA) workbooks are generated at test time by
`tests/unit/helpers/xlsb-writer.ts` and `tests/unit/helpers/iwa-writer.ts` from the published
format descriptions. The .gnumeric, .sxc and .xlsb fixtures were also opened in
LibreOffice 24.2 (`soffice --headless --convert-to csv`) to confirm that another implementation
reads the same values.

### Third-party files (io/b/third-party/)

Real files saved by the original applications, redistributed under their licenses:

| File | Source | License |
|---|---|---|
| `poi-testVarious.xlsb` | Apache POI, `test-data/spreadsheet/testVarious.xlsb` (https://github.com/apache/poi) | Apache License 2.0 |
| `tika-testNumbers.numbers` | Apache Tika, `tika-parser-apple-module/src/test/resources/test-documents/testNumbers.numbers` (https://github.com/apache/tika) | Apache License 2.0 |
| `tika-testNumbers2013.numbers` | Apache Tika, same folder, `testNumbers2013.numbers` | Apache License 2.0 |
