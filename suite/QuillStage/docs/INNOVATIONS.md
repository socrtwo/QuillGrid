# Innovations survey

> **Plan.** Anne's Stage has no code yet. This document surveys presentation software and presentation methods, past and present, records what each one introduced, and names the Anne's Stage feature and command id that will adopt the idea, or says why it is skipped. Command ids match [COMMANDS-PLAN.md](COMMANDS-PLAN.md). Product names are used nominatively, to say where an idea came from; see [LEGAL.md](LEGAL.md).

Dates are given to the year where they are well established and as decades or "later versions" where sources differ. Corrections are welcome.

## How to read this

Each entry has two parts:
- **Introduced:** what the product or method brought to presentations (not necessarily first ever, but where the idea became widely known).
- **Anne's Stage:** the feature that adopts it, with its planned command ids and milestone (M1–M8, see [ROADMAP.md](ROADMAP.md)), or the reason it is skipped.

A summary table is at the end.

---

## 1. The 1980s: from overheads to screen shows

### VCN ExecuVision (early 1980s, IBM PC and Apple II)

- **Introduced:** one of the first personal-computer programs made for business presentations: text and simple graphics composed on screen, a clip-art library, and a sequenced screen show instead of a printed transparency.
- **Anne's Stage:** the screen show itself is the core of the app. The clip-art idea is adopted only with openly licensed sources: **Insert ▸ Icons** (`insert.icons`, Material Symbols) and **Insert ▸ Online Pictures** (`insert.onlinePictures`, searching openly licensed collections and recording each image's license and attribution in the slide notes). M2, M4.

### Living Videotext MORE (1986, Macintosh)

- **Introduced:** an outliner that turned an outline directly into bullet charts and tree charts, so the structure of a talk came first and the slides followed from it.
- **Anne's Stage:** **Outline view** is a first-class authoring view (`view.outline`) with promote, demote, move up and move down (`outline.promote`, `outline.demote`, `outline.moveUp`, `outline.moveDown`), and **Outline to Diagram** (`diagram.fromOutline`) turns a bullet list into an org chart, tree or process diagram. M1, M3.

### Harvard Graphics (Software Publishing Corporation, 1986, DOS)

- **Introduced:** chart-centric business presentations: a gallery of text charts, bar, line, pie and organization charts filled from a data form; screen shows with transitions on DOS; later versions added an advisor that critiqued color, text amount and consistency.
- **Anne's Stage:**
  - **Chart Gallery** with data entry in an embedded grid shared with Anne's Grid (`insert.chart`, `chart.editData`), and org charts as diagrams (`insert.diagram`). M3.
  - **Design Checker** (`review.designCheck`): too many words per slide, text too small for the room size, low contrast, too many fonts or colors, inconsistent title positions, off-master objects, slides over the time budget. It runs on demand and never blocks. M5.

### Forethought Presenter / PowerPoint 1.0 (1987, Macintosh) and PowerPoint 2.0 (1988)

- **Introduced:** the slide as a page with a consistent master (title and body styles shared by every slide); a slide sorter; speaker notes pages and audience handouts as first-class outputs, alongside the slides themselves. 1.0 printed black-and-white overheads; 2.0 added color and 35 mm slide output through a slide-imaging service.
- **Anne's Stage:** the core model: **Slide Master and layouts** (`view.slideMaster`, `master.insertLayout`), **Slide Sorter** (`view.slideSorter`), **Notes Page** (`view.notesPage`), **Handout Master** (`view.handoutMaster`) and **Print ▸ Handouts / Notes Pages** (`file.print` with output choices, `file.export.handoutsPdf`). M1, M2.

### Lotus Freelance Plus / Freelance Graphics (DOS 1986; Windows 1991)

- **Introduced:** *SmartMasters*: templates that combined a look with guided placeholders ("Click here to type a title"), and later content topics: pre-written outlines for common talks (business plan, project status, sales pitch). Its page sorter and screen-show buttons were early examples of interactive slides.
- **Anne's Stage:**
  - Guided placeholders are part of every layout (`master.insertPlaceholder`).
  - **Content Starters** (`file.newFromStarter`): original, openly licensed outlines for common talks (project update, research talk, lesson, pitch, retrospective, wedding toast), each slide with a prompt in the notes about what to say. M4.
  - **Action buttons** on slides (`insert.actionButton`, `insert.action`). M3.

### Aldus Persuasion (1988, Macintosh)

- **Introduced:** truly outline-driven authoring: the outline and the slides stayed in sync in both directions, and *AutoTemplates* chose a layout from the outline structure (a title with bullets got a bullet layout, a title with a data table got a chart layout). Background and title-placement rules came from the template, not from each slide.
- **Anne's Stage:**
  - Two-way outline sync (Outline view and the Outline pane in Normal view edit the same text). M1.
  - **Layout by Outline Level** (`design.autoLayoutRules`): rules in the theme map outline shapes to layouts: a title only → Section Header; title plus 1–6 bullets → Title and Content; title plus a table pasted from Anne's Grid → Title and Chart; title plus one picture → Picture with Caption. Rules can be switched off per slide (`slide.lockLayout`). M4.

### Gold Disk ShowPartner (late 1980s–early 1990s; ShowPartner F/X)

- **Introduced:** desktop-video-style screen shows on home and business computers: scripted sequences of images and text with animated wipes, dissolves and timing, playable unattended.
- **Anne's Stage:** **Kiosk mode** (`show.setUp` → *Browsed at a kiosk*; loops, resets after inactivity, keyboard lockout) and **Rehearse Timings** (`show.rehearseTimings`). Standalone runtime executables are skipped for security; the HTML show player (`file.export.htmlShow`) replaces them. M2, M5.

---

## 2. The 1990s: multimedia and the office suites

### Astound (Gold Disk, 1994)

- **Introduced:** a multimedia presentation tool with a per-object **timeline**: every object had entrance, emphasis and exit tracks that could overlap, follow motion paths, and react to clicks; sound and video were placed on the same timeline; later versions published to the web.
- **Anne's Stage:** the **Animation Timeline** (`anim.timeline`): a track per animated object, drag bars to change delay and duration, and drag to reorder, synchronized with the Animation Pane (`anim.pane`). Media playback is an animation step on the same timeline. M5.

### Corel Presentations / WordPerfect Presentations (1993–1996 onward)

- **Introduced:** drawing and bitmap editing inside the presentation program (no round trip to a paint program), a slide outliner, and *Show on the Go*: a self-running copy of the show for computers without the program.
- **Anne's Stage:**
  - **Edit Picture** (`picture.editPixels`): a small built-in raster editor (crop, brush, eraser, fill, clone, blur region, red-eye) with undo shared with the deck. M6.
  - **Package Presentation** (`file.package`): one folder or .zip with the deck, linked media, embedded fonts where their licenses allow, and the HTML show player. M5.

### PowerPoint 3.0–2003 (1990–2003)

- **Introduced:** the AutoContent Wizard (an interview that produced a starter outline), custom shows (named subsets of slides in a different order), hyperlinks and action buttons, Pack and Go / Package for CD, the Rehearse Timings toolbar, the Presenter View for a second monitor (early 2000s), and grayscale and black-and-white previews for printing.
- **Anne's Stage:**
  - **Custom Slide Show** (`show.customShows`). M2.
  - **Rehearse Timings** and **Use Timings** (`show.rehearseTimings`, `show.useTimings`). M2.
  - **Presenter View** (`show.presenterView`). M2.
  - **Color / Grayscale / Black and White** (`view.color`, `view.grayscale`, `view.blackWhite`). M2.
  - The AutoContent Wizard returns as Content Starters (above) and as the AI tab's **Draft Presentation** (`ai.draftDeck`). M4, M7.

---

## 3. The 2000s: new platforms, typography and the web

### Microsoft Producer (2003, free add-in)

- **Introduced:** a timeline that synchronized recorded video or narration with slide changes, and publishing to the web as a page with the video, the current slide and a clickable table of contents.
- **Anne's Stage:**
  - **Recording timeline** (`record.timeline`): after recording, slide changes, builds and ink appear as markers on the narration waveform and can be dragged to re-synchronize without re-recording. M6.
  - **Export Web Lecture** (`file.export.webLecture`): a self-contained HTML page with the narration video, synchronized slides, a slide list for jumping and captions. M6.

### Apple Keynote (2003)

- **Introduced:** typography-led themes and cinematic, GPU-rendered transitions; later versions added *Magic Move* (objects that appear on consecutive slides glide to their new position, size and color); a **Build Order** panel that lists every build on the slide in one sequence across objects; *Action* builds; a customizable **presenter display**; **Rehearse Slideshow**; *Instant Alpha* (remove a color region from a picture); a phone remote; *Keynote Live* (2016, watch a show from a link); and live camera video and multi-presenter shows (2020s).
- **Anne's Stage:**
  - **Smooth Morph** transition (`trans.smoothMorph`): matches objects between consecutive slides by name, then by content, then by type, and tweens position, size, rotation, color, corner radius and text; character- and word-level text morphing. Object naming for matching is in the Selection Pane (`home.selectionPane`). M5.
  - **Build Order** panel (`anim.buildOrder`): every build on the slide in one list, drag to reorder, with start (*on click*, *with previous*, *after previous*) and delay per row. It is a view of the same animation sequence as the Animation Pane. M5.
  - **Presenter View layouts** (`show.presenterLayout`): choose and arrange current slide, next slide or next build, notes, clock, elapsed time, countdown and a ready indicator; notes font size and color; layouts saved per deck. M2.
  - **Rehearse** (`show.rehearse`): presenter view on a single screen with timing per slide. M2.
  - **Remove Color** (`picture.removeColor`): click a color and drag to widen the tolerance, like a magic wand. M4.
  - **Phone Remote** (`show.remote`): pair a phone by QR code (local network or WebRTC); it shows notes, next slide and timer and has next, previous, black screen and laser pointer. M6.
  - **Broadcast Show** (`show.broadcast`): viewers follow along in a browser from a link. M8 (needs a relay; self-hostable).
  - **Live Camera** objects (`insert.liveCamera`), below under PowerPoint Cameo. M6.

### LaTeX Beamer (Till Tantau, 2003)

- **Introduced:** slides typeset with LaTeX: overlay specifications (`<2->`, `<1-3>`) that reveal or change any element on given steps, mathematical typesetting, a handout mode that collapses overlays, notes on a second screen, and PDF as a universal show format.
- **Anne's Stage:**
  - **Equations** with LaTeX input and an equation editor (`insert.equation`, `equation.linear`, `equation.professional`). M3.
  - **Reveal by Step** (`anim.revealSteps`): give each paragraph or object a step range in a compact syntax (`2-`, `1-3`, `4`), which is turned into ordinary appear and disappear animations. M5.
  - **Export PDF with build steps** (`file.export.pdf` with *One page per build step*) and **Export Handout PDF** (`file.export.handoutsPdf`, builds collapsed). M2.
  - Importing `.tex` sources is skipped: LaTeX documents are programs, not data; export to Beamer (`file.export.beamer`) is a stretch goal. M8.

### S5, impress.js and Sozi (2004–2012, HTML and SVG shows)

- **Introduced:** slide shows that run in any browser (S5, by Eric Meyer, 2004), 3-D CSS camera moves between positioned steps (impress.js), and zooming presentations built from a single SVG drawing (Sozi, open source).
- **Anne's Stage:** **Export HTML Show** (`file.export.htmlShow`): one file, no server, keyboard and touch navigation, notes window, works offline. Canvas mode (below) borrows the camera-path idea. M5.

### Google Slides (2007, as the presentation part of Google Docs)

- **Introduced:** real-time co-editing with presence cursors in the browser, comments with assignments, automatic version history, linked charts that refresh from a spreadsheet, and audience **Q&A** during a show (2016): viewers submit and vote on questions from a short link.
- **Anne's Stage:**
  - **Version History** (`file.versions`) and **Take Snapshot** (`file.snapshot`), as in the sister apps. M1.
  - **Comments** with replies, mentions and resolve (`review.newComment`). M3.
  - **Linked Chart from Anne's Grid** (`insert.linkedChart`, `chart.refreshData`). M6.
  - **Audience Q&A** (`show.qa`): questions arrive in presenter view; the presenter can show one on screen. M6.
  - Live co-editing is planned for M8 (CRDT document sync with a self-hostable relay); it is not required for any other feature.

### LibreOffice Impress (StarOffice / OpenOffice heritage; LibreOffice from 2010)

- **Introduced (as an open-source suite):** a **Presenter Console** for the second screen, *Fontwork* decorative text, the **Navigator** listing slides and named objects, **Impress Remote** apps for phones, and the **Presentation Minimizer** that compresses images and removes unused masters.
- **Anne's Stage:**
  - **Presenter View** (above) and **Phone Remote** (`show.remote`).
  - **Decorative Text** (`insert.decorativeText`, `shape.textTransform`). M4.
  - **Navigator pane** (`view.navigator`): slides, sections, named objects, media, links and comments in one tree; double-click to go there. M3.
  - **Compress Presentation** (`file.compress`): recompress pictures to a target resolution, delete cropped areas, remove unused masters and layouts, transcode media, with a before/after size report. M5.
  - OpenDocument .odp is a first-class format (open and save). M2.

### Prezi (2009)

- **Introduced:** a **zooming canvas** instead of a deck: content is laid out on one large surface, and the show is a camera path that pans, zooms and rotates between frames; an **overview** shows the whole canvas, and topics can be visited out of order. Later, Prezi Video placed the presenter's camera with the content.
- **Anne's Stage:**
  - **Canvas mode** (`view.canvas`, `canvas.addFrame`, `canvas.editPath`): a presentation can be a single infinite canvas whose frames become steps; the camera eases between them. Every frame is still exported as a normal slide to .pptx, so files stay compatible. M7.
  - **Zoom links** (`insert.summaryZoom`, `insert.sectionZoom`, `insert.slideZoom`) give a zoom-out overview inside an ordinary deck. M5.
  - Canvas mode is not the default; a linear deck is easier to edit and print, and many viewers find constant zooming tiring. The Design Checker warns about large rotations.

---

## 4. The 2010s: browser decks, design automation and audiences

### reveal.js (2011), Marp (2016 onward) and Slidev (2021)

- **Introduced:**
  - reveal.js: HTML presentations with **vertical slides** (two-dimensional navigation: main topics across, details down), *fragments* (step-by-step reveals), a speaker-notes window, Markdown content, **code highlighting with step-through line highlights**, and *auto-animate*.
  - Marp: write the deck as **Markdown** with directives (`<!-- theme: x -->`), `---` between slides, and themes in CSS.
  - Slidev: Markdown slides for developers with live, editable code blocks, drawing on slides, recording, and a presenter mode.
- **Anne's Stage:**
  - **Import / Export Markdown Slides** (`file.import.markdown`, `file.export.markdown`): `---` separates slides, `#` is the title, lists become bullets, `Note:` or `<!-- notes -->` blocks become speaker notes, front matter sets theme and size. M4.
  - **Vertical slide stacks** (`slide.stackBelow`, `show.navigate2d`): a slide can have child slides; in a show, Down enters the stack and Right skips it. In .pptx a stack is saved as a section with consecutive slides so other programs still play it linearly. M7.
  - **Code block** object (`insert.codeBlock`): syntax highlighting for common languages, line numbers, and **Step Through Lines** (`code.stepLines`): highlight ranges `1-3 | 5 | 7-9` advanced like builds. **Run Code** (`code.run`) runs JavaScript in a sandbox and shows the output (off by default). M5.

### Haiku Deck (2012, iPad) and the "Presentation Zen" style

- **Introduced:** one full-bleed image and a few words per slide, image search built into the editor, and a hard limit on text that pushed users away from bullet lists.
- **Anne's Stage:**
  - **Big Picture** layouts (full-bleed image with a title band) in every built-in theme. M4.
  - **Find Images** (`insert.onlinePictures`) searches openly licensed collections and records attribution. M4.
  - **Word budget** in the Design Checker (`review.designCheck`), configurable per deck. M5.

### Emaze (2010s)

- **Introduced:** web templates with video and 3-D backgrounds, and presentations that play in any browser from a link.
- **Anne's Stage:** **Video background** for slides and layouts (`design.backgroundVideo`, muted and looped, with a still-frame fallback for printing and PDF). Hosted sharing is skipped; the HTML export can be put on any static host. M6.

### Canva (2013)

- **Introduced:** drag-and-drop design for non-designers with very large template and asset libraries, a **brand kit** (locked colors, fonts and logos), one-click **resize** of a design to other formats, presenter recording ("talking presentations"), and live audience questions.
- **Anne's Stage:**
  - **Brand Kit** (`design.brandKit`): theme colors, fonts, logos and approved images saved once and applied to any deck; off-brand colors are flagged by the Design Checker. M6.
  - **Resize Presentation** (`design.resizeTo`): 16:9, 4:3, 16:10, A4/Letter, 1:1 square and 9:16 vertical story, with layouts that reflow placeholders instead of only scaling (*Maximize*, *Ensure Fit*, *Reflow*). M4.
  - Recording: see Record tab. Audience questions: see Q&A.
  - Copying a stock library is skipped (licensing); openly licensed search is used instead.

### Zoho Show (2007/2008)

- **Introduced:** a browser presentation editor with collaboration, broadcasting a show to remote viewers, and apps that present on a TV.
- **Anne's Stage:** **Cast to Display** (`show.cast`): use the platform's cast or external-display support (Android external displays and casting, AirPlay on iOS and macOS, the web Presentation API) to put the show on a TV while the phone shows presenter view. M6. Broadcast: see Keynote Live.

### Microsoft Sway (2014–2015)

- **Introduced:** **responsive storylines**: content as cards that reflow for any screen, read by scrolling rather than paging, with automatic design.
- **Anne's Stage:** **Export Story** (`file.export.story`): a responsive, scrolling HTML page made from the deck (titles become section headings, slides become cards, builds become scroll-triggered reveals, notes can be included as body text), and **Story view** (`view.story`) to preview it. M6.

### Ludus (2016–2017)

- **Introduced:** a presentation tool for designers: embed almost anything (web pages, videos, design files, code) and **smart blocks**: reusable components that update everywhere when edited.
- **Anne's Stage:**
  - **Web Embed** (`insert.webEmbed`): a live web view in the show where the platform allows, with a captured snapshot used in edit mode, print, PDF and other programs. M6.
  - **Linked Components** (`slide.createComponent`, `slide.insertComponent`, `slide.detachComponent`): a group of objects saved once and reused across slides and decks; editing the source updates every instance. Saved to .pptx as ordinary shapes with tags so other programs see normal content. M6.

### Mentimeter (2014) and Slido (2012)

- **Introduced:** audience interaction from phones with a short join code: live **polls**, **word clouds**, **quizzes** with leaderboards, rating scales, and **Q&A** with up-voting, results updating on the slide in real time.
- **Anne's Stage:** **Interactive slides** (`insert.poll`, `insert.wordCloud`, `insert.quiz`, `insert.rating`, `show.qa`, `show.showResults`, `show.resetResults`). Audience devices join by QR code or short code. Transport, in order of preference: direct WebRTC with a self-hostable signaling relay; on desktop, a local-network server for rooms without internet. Results are stored in the deck (and can be exported to CSV for Anne's Grid). Without a connection, the slides show a static placeholder. M6.

### WPS Presentation (Kingsoft WPS Office)

- **Introduced:** a free, compatibility-focused office suite on mobile and desktop with **document tabs**, built-in PDF tools, and one-click "beautify" restyling.
- **Anne's Stage:** **Document tabs** (several open decks in one window, as in the sister apps), **Insert PDF as Slides** (`insert.pdfPages`), and **Design Ideas** (`design.designIdeas`). M2, M4.

### Focusky (2010s)

- **Introduced:** Prezi-like zooming presentations made offline on the desktop, exported to HTML5, executables and video.
- **Anne's Stage:** canvas mode works offline, and canvas decks export to HTML (`file.export.htmlShow`) and video (`file.export.video`) like linear decks. Executable export is skipped. M7.

### PowerPoint in the Microsoft 365 era (2015 onward)

- **Introduced:**
  - **Designer / Design Ideas** (2015): layout suggestions computed from the slide's content (pictures, lists, dates, numbers).
  - **Morph** (2015–2016): object tweening between consecutive slides.
  - **Zoom** (2017): summary, section and slide zooms that jump around a deck and return.
  - **Presenter Coach / Rehearse with Coach** (2019): private rehearsal feedback on pace, filler words, repetitive language, reading from the slide, and inclusive language, with a report.
  - **Live captions and subtitles** (2019): speech recognition during a show, with translation into another language.
  - **Cameo** (2021–2022): a live camera feed placed on slides as a shaped object.
  - **Recording Studio** (2022 revision): record with camera and teleprompter, pen and laser recorded with the narration, re-record a single slide, export to video.
  - Ink replay, ink to shape and ink to math; Accessibility Checker; Reuse Slides; Present Live / Live Presentations (audience follows on phones).
- **Anne's Stage:**
  - **Design Ideas** (`design.designIdeas`): rule-based suggestions offline (no account), optionally ranked by AI. M4.
  - **Smooth Morph** (`trans.smoothMorph`). M5.
  - **Zoom links** (`insert.summaryZoom`, `insert.sectionZoom`, `insert.slideZoom`, `zoom.returnToZoom`). M5.
  - **Rehearsal Coach** (`show.rehearsalCoach`): on-device speech recognition where available; pace in words per minute, filler words (per language), long pauses, monotone pitch, reading the slide verbatim, time per slide against a budget; a report saved to the deck's notes or exported. Never uploads audio unless the user chooses a cloud speech provider. M6.
  - **Live Subtitles** (`show.subtitles`, `show.subtitleSettings`): on-device recognition when the platform has it, otherwise the chosen provider; translation through the AI provider. M6.
  - **Live Camera** (`insert.liveCamera`): a camera feed as a shape with crop, border and effects; persists in recordings. M6.
  - **Record tab** with teleprompter notes, camera, pen and laser capture, re-record slide, and export (`record.fromBeginning`, `record.fromCurrent`, `record.teleprompter`, `file.export.video`). M6.
  - **Ink** commands (`ink.replay`, `ink.toShape`, `ink.toMath`). M6.
  - **Check Accessibility** (`review.checkAccessibility`) with alt text, reading order, contrast and missing titles. M3.
  - **Audience Link** (`show.audienceLink`): viewers follow on their phones with subtitles in their own language. M8.

### Apple Keynote Live (2016)

- **Introduced:** present to remote viewers who watch from a link in a browser, without any software.
- **Anne's Stage:** **Broadcast Show** (`show.broadcast`), M8, using the same relay as audience interaction.

---

## 5. The 2020s: AI-first decks, canvases and video

### Pitch (2019, public 2021)

- **Introduced:** collaborative decks with **slide status and assignees** (who owns which slide, what is ready), live video embedded in slides, reusable templates per team, and presentation analytics on shared links.
- **Anne's Stage:**
  - **Slide Status** (`slide.status`: *Draft, In review, Ready*) and **Assign Slide** (`slide.assign`), shown in the slide sorter and filterable. Stored in .pptx as slide tags. M6.
  - Link analytics are skipped: they require hosting and tracking that do not fit a privacy-first, offline app.

### Tome (2022) and Gamma (2022–2023)

- **Introduced:** **AI-first decks**: type a prompt or paste a document and get a complete, designed draft; **cards** of flexible length instead of fixed-size slides; nested cards; one-click restyle of the whole deck. (Tome later shifted its focus away from general presentations.)
- **Anne's Stage:**
  - AI tab: **Draft Presentation** from a prompt (`ai.draftDeck`), **Document to Deck** from an Anne's Pen document, Markdown, PDF or web page (`ai.docToDeck`), always outline first so the user approves the structure before slides are made. M7.
  - **Restyle Deck** (`design.restyle`): apply another theme and re-run layout rules on every slide. M4.
  - Cards are adopted as an export (Story, above), not as the document model, so .pptx and .odp stay faithful.

### PowerPoint Cameo and Recording Studio (2021–2022)

See the Microsoft 365 entry above (`insert.liveCamera`, Record tab).

### KineMaster-style mobile video editing (2013 onward)

- **Introduced:** practical video editing on phones: trim, split, speed, crop, fades, volume envelopes and captions on a touch timeline.
- **Anne's Stage:** **Media editing** on inserted video and audio (`media.trim`, `media.split`, `media.speed`, `media.crop`, `media.fadeIn`, `media.fadeOut`, `media.volumeEnvelope`, `media.insertCaptions`, `media.addBookmark`), non-destructive and saved as trim and bookmark settings that .pptx understands; **Export Trimmed Media** (`media.exportTrimmed`) bakes the result where the platform has an encoder. M6.

### Apple Freeform (2022), Miro (2011, as RealtimeBoard) and Whimsical (2017)

- **Introduced:** infinite collaborative canvases with sticky notes and connectors (Freeform, Miro), and fast keyboard-driven diagramming: press Tab to add a child in a mind map, Enter for a sibling, with smart connectors and automatic layout (Whimsical).
- **Anne's Stage:**
  - **Brainstorm board** in canvas mode (`canvas.brainstorm`): sticky notes (`insert.stickyNote`), freehand ink and connectors; **Board to Slides** (`canvas.toSlides`) turns clusters into slides. M7.
  - **Keyboard diagramming** in the Diagram object: Tab adds a child, Enter a sibling, Shift+Tab promotes; mind map, tree, flowchart and timeline layouts with automatic arrangement (`diagram.autoLayout`). M3, M5.
  - **Smart connectors** that stay glued to shapes and route around them (`shape.connector`, `shape.rerouteConnectors`). M3.

---

## 6. Presentation methods

### Plain-text decks (the Paul Graham / Y Combinator demo-day style)

- **Introduced:** the minimal style long recommended to founders at Y Combinator demo days (the accelerator Paul Graham co-founded): plain dark text on a light background, large type, one point per slide, no transitions, legible from the back of the room.
- **Anne's Stage:**
  - **Plain** theme family (black on white, white on black; one sans, one serif). M2.
  - **Text to Deck** (`file.import.textDeck`): paste or open plain text; each paragraph (or each line) becomes a slide with the text fitted large. M4.

### The Takahashi method (Masayoshi Takahashi, early 2000s)

- **Introduced:** very large text, often a single word or short phrase per slide, many slides, no pictures; developed in the Japanese Ruby community.
- **Anne's Stage:** **Fit Text to Slide** (`fmt.fitTextToSlide`): set the font size so the text fills the slide; a Takahashi layout in the Plain themes; **Split into Slides** (`slide.splitByParagraph`) makes one slide per line. M4.

### The Lessig method (Lawrence Lessig, early 2000s)

- **Introduced:** rapid-fire slides, often several per sentence, each showing one word, quote or image, exactly synchronized with speech.
- **Anne's Stage:** fast authoring with Text to Deck and Split into Slides; **Rehearse Timings** with tenth-of-a-second precision; and **Follow Script** (`show.followScript`, stretch): advance automatically when speech recognition reaches the end of a slide's script in the notes. M8.

### Pecha Kucha (Astrid Klein and Mark Dytham, Tokyo, 2003) and Ignite (Brady Forrest and Bre Pettis, Seattle, 2006)

- **Introduced:** fixed formats: Pecha Kucha is 20 slides × 20 seconds; Ignite is 20 slides × 15 seconds (five minutes), both auto-advancing so the speaker must keep pace.
- **Anne's Stage:** **Timed Format** (`show.timedFormat`): Pecha Kucha, Ignite or custom N × S; sets automatic advance on every slide, checks the slide count, and shows a per-slide countdown ring in presenter view and on the Phone Remote. M5.

### Assertion–evidence (Michael Alley)

- **Introduced:** a research-backed structure for technical talks: a full-sentence headline that states the claim, supported by visual evidence instead of bullets.
- **Anne's Stage:** an **Assertion–Evidence** layout in every built-in theme (two-line sentence title, large content area) and a Design Checker rule that suggests it for bullet-heavy technical slides. M4.

### Edward Tufte's critique (*The Cognitive Style of PowerPoint*, 2003)

- **Introduced:** an influential critique of low-resolution, bullet-nested slides and decorative chart elements, arguing for high data density and handouts that carry the detail.
- **Anne's Stage:** the Design Checker warns about bullet nesting deeper than two levels and chart decoration (3-D effects on 2-D data, heavy gridlines); **Sparkline** text objects (`insert.sparkline`); and **Export Handout Document** (`file.export.handoutDoc`) sends slides plus full notes to Anne's Pen as a readable document. M5, M6.

---

## Summary

| Idea | From | Anne's Stage command(s) | Milestone |
|---|---|---|---|
| Masters, slide sorter, notes pages, handouts | PowerPoint 1.0/2.0 | `view.slideMaster`, `view.slideSorter`, `view.notesPage`, `view.handoutMaster` | M1–M2 |
| Outline-first authoring | MORE, Persuasion | `view.outline`, `outline.promote`, `outline.demote` | M1 |
| Layout from outline structure | Persuasion AutoTemplates | `design.autoLayoutRules`, `slide.lockLayout` | M4 |
| Guided placeholders, content starters | Freelance SmartMasters | `master.insertPlaceholder`, `file.newFromStarter` | M4 |
| Chart gallery, design advisor | Harvard Graphics | `insert.chart`, `review.designCheck` | M3, M5 |
| Kiosk shows | ShowPartner | `show.setUp` (kiosk) | M2 |
| Animation timeline | Astound | `anim.timeline` | M5 |
| Built-in picture editing, portable show | Corel Presentations | `picture.editPixels`, `file.package` | M5–M6 |
| Custom shows, presenter view, rehearse timings | PowerPoint 97–2003 | `show.customShows`, `show.presenterView`, `show.rehearseTimings` | M2 |
| Narration timeline, web lecture | Producer | `record.timeline`, `file.export.webLecture` | M6 |
| Object morphing | Keynote, PowerPoint | `trans.smoothMorph` | M5 |
| Build order panel | Keynote | `anim.buildOrder` | M5 |
| Presenter display layouts, rehearse | Keynote | `show.presenterLayout`, `show.rehearse` | M2 |
| Remove color from picture | Keynote | `picture.removeColor` | M4 |
| Phone remote | Keynote, Impress | `show.remote` | M6 |
| Overlay steps, equations, PDF with steps | Beamer | `anim.revealSteps`, `insert.equation`, `file.export.pdf` | M2–M5 |
| HTML show | S5, reveal.js | `file.export.htmlShow` | M5 |
| Version history, comments, linked charts, Q&A | Google Slides | `file.versions`, `review.newComment`, `insert.linkedChart`, `show.qa` | M1–M6 |
| Navigator, minimizer, decorative text | Impress | `view.navigator`, `file.compress`, `insert.decorativeText` | M3–M5 |
| Zooming canvas, overview | Prezi, Focusky, Sozi | `view.canvas`, `canvas.addFrame`, `insert.summaryZoom` | M5, M7 |
| Markdown slides, vertical stacks, code steps | reveal.js, Marp, Slidev | `file.import.markdown`, `slide.stackBelow`, `code.stepLines` | M4–M7 |
| Big-picture layouts, image search | Haiku Deck | `insert.onlinePictures` | M4 |
| Video backgrounds | Emaze | `design.backgroundVideo` | M6 |
| Brand kit, resize | Canva | `design.brandKit`, `design.resizeTo` | M4, M6 |
| Cast to TV | Zoho Show | `show.cast` | M6 |
| Responsive story | Sway | `file.export.story`, `view.story` | M6 |
| Web embeds, reusable components | Ludus | `insert.webEmbed`, `slide.createComponent` | M6 |
| Polls, word clouds, quizzes, Q&A | Mentimeter, Slido | `insert.poll`, `insert.wordCloud`, `insert.quiz`, `show.qa` | M6 |
| Document tabs, PDF to slides, one-click restyle | WPS | `insert.pdfPages`, `design.restyle` | M2–M4 |
| Design ideas, zoom links, coach, subtitles, camera, recording | PowerPoint (365 era) | `design.designIdeas`, `insert.summaryZoom`, `show.rehearsalCoach`, `show.subtitles`, `insert.liveCamera`, `record.fromBeginning` | M4–M6 |
| Broadcast | Keynote Live | `show.broadcast` | M8 |
| Slide status and owners | Pitch | `slide.status`, `slide.assign` | M6 |
| AI drafts, document to deck | Tome, Gamma | `ai.draftDeck`, `ai.docToDeck` | M7 |
| Video trim, split, speed | KineMaster | `media.trim`, `media.split`, `media.speed` | M6 |
| Infinite canvas, keyboard diagramming | Freeform, Miro, Whimsical | `canvas.brainstorm`, `diagram.autoLayout`, `shape.connector` | M3–M7 |
| Plain-text decks, Takahashi, Lessig | Methods | `file.import.textDeck`, `fmt.fitTextToSlide`, `slide.splitByParagraph` | M4 |
| Pecha Kucha and Ignite timers | Methods | `show.timedFormat` | M5 |
| Assertion–evidence, data density | Alley, Tufte | layouts, `review.designCheck`, `insert.sparkline` | M4–M6 |

## Considered and skipped

| Idea | Why |
|---|---|
| Opening Keynote (.key) and Prezi files | The formats are not publicly specified. A Keynote package's embedded preview images may be offered as *Import Slide Images* (best effort, M8); nothing else is planned. |
| Link analytics and hosted sharing | Require tracking and a hosted service; contrary to the offline, privacy-first design. Static HTML export can be hosted anywhere. |
| Self-running executables (Show on the Go, Pack and Go Viewer, Focusky EXE) | Security risk and platform-specific; the HTML show player covers the need. |
| Running VBA macros in .pptm | Security; VBA is preserved on save, and Anne's Stage macros use the suite's JavaScript macro language. |
| Bundled stock photo, clip-art or template libraries from other vendors | Copyright; Anne's Stage ships original templates and searches openly licensed collections. |
| Cards as the document model | Breaks round-tripping with .pptx and .odp; cards are offered as an export. |
| Flash, Shockwave and other legacy plug-in media | Obsolete; such objects open as their poster image. |
