# File formats (plan)

Anne's Stage will open presentations from every current and historical presentation program it can
read faithfully, recognizing each file by its **content, not its extension**, the same way Anne's Pen,
Anne's Grid and Anne's Post do. A misnamed file, or one with no extension, opens with the right reader,
and the person is told when the content and the name disagree.

Fidelity tiers used below:
- **Full**: everything the format can express that the app models (layouts, masters, themes, text
  formatting, shapes, images, tables, charts, transitions, animations, notes).
- **High**: slides, masters, text and shape formatting, images, tables and notes; some effects or
  animations approximated.
- **Content**: slides with text, basic formatting, images and simple shapes, laid out as in the original.
- **Text**: slide titles, text and notes only (used when the format is undocumented or damaged).

## Current formats

| Format | Program(s) | Extensions | Detection | Planned tier |
|---|---|---|---|---|
| PresentationML (OOXML) | PowerPoint 2007–365, Google Slides, Keynote and others (export) | .pptx .pptm .potx .potm .ppsx .ppsm | ZIP with `ppt/presentation.xml` | Full |
| OOXML Strict | PowerPoint 2010+ (Strict) | .pptx | ZIP; Strict namespaces | Full |
| PowerPoint 97–2003 binary | PowerPoint 97–2003, many clones | .ppt .pps .pot | CFB with `PowerPoint Document` stream (MS-PPT) | High |
| OpenDocument Presentation | LibreOffice, OpenOffice, Collabora, Calligra | .odp .otp .fodp | ZIP `mimetype` / flat XML root | Full |
| Apple Keynote | Keynote 6+ (2013 on) | .key | ZIP or bundle with `Index/*.iwa` (Snappy + protobuf) | Content (Text where only previews parse) |
| Apple Keynote '09 and earlier | Keynote 1–5 | .key | Bundle/ZIP with `index.apxl(.gz)` XML | High |
| Kingsoft WPS Presentation | WPS Office | .dps .dpt | CFB like .ppt | High |
| Hancom Show | Hancom Office | .show .hpt | CFB or OWPML-style XML | Content |
| PDF | any (handouts, exported decks) | .pdf | `%PDF-` | Content: one slide per page, text kept editable where the PDF has it |
| Markdown slides | Marp, Slidev, reveal.js, Deckset, Remark | .md | front matter or `---` slide separators | High |
| HTML slide decks | reveal.js, impress.js, S5, Shower, Google Slides "publish" | .html | known framework markers | Content |
| LaTeX Beamer | Beamer | .tex | `\documentclass{beamer}` | Content (frames, titles, lists, images; math as images) |
| Images and image sets | any | .png .jpg .svg .gif .webp, ZIP of images | magic numbers | Full (one image per slide) |

## Historical formats

| Format | Program(s) | Era | Extensions | Detection | Planned tier |
|---|---|---|---|---|---|
| PowerPoint 4.0 / 95 binary | PowerPoint for Windows 4.0 and 95, Mac 4.0 | 1994–1996 | .ppt .pps | CFB with older `PowerPoint Document` records / Mac resource layout | Content |
| PowerPoint 2.0/3.0 | PowerPoint for Windows 2.0–3.0, Mac 1–3 | 1987–1992 | .ppt | legacy header | Text (best effort) |
| StarOffice / OpenOffice.org 1.x | StarOffice 6–7, OOo 1.x | 2002–2005 | .sxi .sti | ZIP `mimetype` application/vnd.sun.xml.impress | High |
| StarImpress binary | StarOffice 3–5 | 1995–2001 | .sdd .sdp | StarDivision CFB streams | Text (best effort) |
| Lotus Freelance Graphics | Freelance for Windows 2–9.x (SmartSuite) | 1991–2013 | .prz .pre .smc .mas | Freelance file signature | Content |
| Harvard Graphics | Harvard Graphics 2.x–98 | 1986–2002 | .ch3 .prs .sy3 .cht .hgm | Harvard signatures | Content (charts become native charts where the data is present) |
| Corel Presentations | Corel Presentations 7–X9 (WordPerfect Office) | 1996– | .shw .shw3 .mst | WordPerfect-family prefix (`\xFFWPC`) with Presentations product code | Content |
| Aldus / Adobe Persuasion | Persuasion 1–4 (Mac/Windows) | 1988–1997 | .prs .per | Persuasion header | Text (best effort) |
| Astound / Gold Disk | Astound, Gold Disk ShowPartner | 1990s | .asd .shw | product signatures | Text (best effort) |
| Microsoft Works / Word "slide" outlines | outline-only decks | 1990s | .rtf .doc | via Anne's Pen's readers: Heading 1 = slide title | Content |
| ClarisWorks / AppleWorks slide shows | ClarisWorks 4–AppleWorks 6 | 1995–2007 | .cwk | AppleWorks header, presentation document type | Text (best effort) |
| MagicPoint | MagicPoint | 1997– | .mgp | `%` directives | High |
| Sozi | Sozi (SVG zooming presentations) | 2010– | .svg | Sozi namespace | Content (frames become slides) |
| Prezi portable exports | Prezi Classic portable | 2009–2019 | .pez / folder | Prezi XML | Content (frames become slides, path order kept) |

Formats that are proprietary and undocumented are attempted only clean-room from public information;
where that is not enough, Anne's Stage recognizes the file, explains why it cannot be opened, and
offers a text salvage.

## Files from other kinds of programs

Anne's Stage recognizes documents, spreadsheets and email by content too, and offers to open them in
the sister app (Anne's Pen, Anne's Grid, Anne's Post) when it is installed, or to import them:
- documents (.docx, .odt, .rtf, .doc, .wpd, .pages …): each Heading 1 starts a slide (an outline
  deck), with its text, lists and images; or the whole document as one slide per page;
- spreadsheets (.xlsx, .ods, .csv, .xls, .numbers …): a table or a chart per selected range;
- email (.eml, .msg): a slide with the message's header and body, attachments listed.

## Damaged files

Every reader feeds the suite's recovery pipeline: byte-level ZIP salvage, XML repair, slide-by-slide
salvage (a broken slide is replaced by a placeholder, the others open), media harvest, and for .ppt
a record-level scan of the PowerPoint Document stream, with a recovery report.
