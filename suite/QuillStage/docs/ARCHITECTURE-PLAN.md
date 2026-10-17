# Architecture plan

> **This is a plan, not a description of existing code.** QuillStage has no code yet. The design below mirrors what QuillPen and QuillGrid already ship, so that the suite shares one shell and one set of conventions. Details will change as the code is written; this file will then be replaced by `docs/ARCHITECTURE.md`.

## Summary

QuillStage will be a TypeScript web application built with **Vite** and tested with **vitest**. The Android and iOS releases will package it with **Capacitor 8**, the Windows, macOS and Linux releases with **Electron**, and the same build will run in a browser as an installable, offline-capable **PWA**. It will share its application shell with QuillPen and QuillGrid: command registry, Quick Access Bar, customizable tabs and keyboard, dialogs, command search, macro recorder, AI setup and providers, spelling and dictionary downloads, the recovery engine's generic parts, and the platform layer.

## Intended source layout

```
src/
  main.ts                 entry: fonts, command registry, App + Shell, launch-file handling, back button
  app.ts                  App (presentation tabs, exec(), recording, autosave) and DeckTab (deck, view state, history)
  prefs.ts                per-device preferences: QAT (QAT_STANDARD, QAT_EXTRAS, addAllQatChoices), custom tabs,
                          keymap, macros, AutoCorrect, presenter layouts, AI settings (keys are not stored here)
  templates/              original themes, layouts, content starters, sample decks
  model/
    types.ts              Deck, Master, Layout, Slide, Section, CustomShow, shapes, text, tables, charts,
                          diagrams, media, comments, timing (transition and animation trees), tags
    deck.ts               Deck with transactions and before-image undo/redo
    theme.ts              theme colors, fonts, effects; color resolution (scheme colors, tints, shades, alpha)
    geometry.ts           preset shape geometry (the ~180 standard presets) and custom paths
    text.ts               rich text: paragraphs, runs, bullets, levels, fields; autofit (shrink, resize)
    inherit.ts            property inheritance: master → layout → slide → shape → paragraph → run
    timing.ts             animation model (sequences, triggers, effects) and its mapping to a timeline
  render/
    slide.ts              renders a slide to DOM/SVG at any scale (edit, thumbnail, show, print)
    text.ts               text layout: line breaking, hyphenation, autofit measurement, columns, RTL
    effects.ts            shadows, glow, soft edges, reflection, bevel (SVG filters; canvas fallback)
    thumbs.ts             thumbnail cache (OffscreenCanvas in a worker where available)
    transitions/          transition engine: CSS/Web Animations for 2-D effects, WebGL for 3-D effects,
                          Smooth Morph (object matching and tweening)
    animate.ts            animation player: builds a timeline from model/timing.ts and runs it
  edit/
    canvas.ts             editing surface: selection, handles, rotation, crop, smart guides, snapping, touch
    textEdit.ts           in-place text editing (contenteditable island bound to the model)
    outline.ts            Outline view editor (two-way with slide text)
    sorter.ts             Slide Sorter, sections, drag and drop
    master.ts             master and layout editing
    table.ts, chart.ts, diagram.ts, media.ts, ink.ts, equation.ts, code.ts
  show/
    player.ts             slide show: navigation, builds, timings, kiosk, custom shows, 2-D stacks, zoom links
    presenter.ts          presenter view: layouts, notes, next build, timers, grid, ink, laser, subtitles
    displays.ts           second-screen strategies per platform (see below)
    remote.ts             phone remote pairing (QR code; local network or WebRTC)
    audience.ts           polls, word clouds, quizzes, Q&A (WebRTC with a self-hostable signaling relay)
    coach.ts              rehearsal coach (speech recognition, pace, fillers, pauses, report)
    record.ts             recording studio: camera, microphone, ink and laser capture, per-slide takes
  canvas/                 canvas mode (frames, camera path, brainstorm board) — milestone M7
  charts/                 chart renderer (shared code with QuillGrid's SVG chart renderer)
  commands/
    registry.ts           Command interface and registry: the single source for QAT, tabs, keys, macros, search
    *.ts                  command modules by area (file, edit, fmt, slide, insert, draw, design, trans, anim,
                          show, record, review, view, master, shape, picture, table, chart, diagram, media,
                          tools, ai, dev, help)
  io/
    pptx/                 PresentationML import and export (.pptx, .pptm, .ppsx, .potx), DrawingML, charts,
                          diagrams (data model + drawing fallback), media, comments, custom XML, tags
    ppt/                  PowerPoint 97–2003 binary import (from the published [MS-PPT] specification)
    odp/                  OpenDocument presentation import and export (.odp, .otp, .fodp)
    markdown.ts           Markdown slides import and export
    outline.ts            text, RTF and .docx outlines
    qsp.ts                native JSON package (.qsp) and AES-256 encrypted (.qspe)
    export/               PDF, images, video, GIF, HTML show, story, web lecture, handouts, Beamer
    recover/              recovery engine (shared ZIP salvage, ImmortalInflate, XML healing, file identifier,
                          plus presentation-specific methods)
  proof/                  spelling worker (nspell / Hunspell) and dictionary downloads — shared
  ai/                     providers, client, settings, on-device model, presentation AI features — shared core
  ui/                     shell, ribbon layout, dialogs, panes, backstage, customize dialogs, wizards — shared
  platform/native.ts      JS side of the QuillNative Capacitor plugin and the Electron bridge, browser fallbacks
android/ ios/ electron/   native packaging
tests/
  unit/                   vitest: model, inheritance, io round trips, recovery, commands, timing
  e2e/                    browser smoke test (Playwright), as in the sister apps
```

## Key design decisions

- **Commands are data.** The ribbon, Quick Access Bar, context menus, keyboard shortcuts, macros and search all refer to command ids, so every feature can be placed anywhere. A command declares its id, label, icon, tab and group (or *not in tabs*), shortcut, enablement test and `run(app, arg?)`. `npm run catalog` will generate `docs/COMMANDS.md`, and a unit test will fail if any command lacks a label or icon or if a tab refers to an unknown id. This is how "no missing commands" is enforced: the planned inventory in [COMMANDS-PLAN.md](COMMANDS-PLAN.md) becomes a checklist test.
- **The model follows PresentationML, not the other way round.** The in-memory model uses the same concepts as Office Open XML (masters, layouts, placeholders with inheritance, theme-relative colors, the timing tree), so .pptx round trips do not lose structure. OpenDocument and Markdown are mapped onto it. Unknown .pptx parts and extension lists are kept and written back unchanged where they remain valid.
- **Transactions record before-images,** as in QuillGrid: every change runs inside `Deck.transact()`. Slide edits record the slide's previous state; structural edits record the slide list, sections and masters. Undo restores before-images, so no command needs inverse logic, and macros are recorded at the command level.
- **One renderer for every surface.** Edit view, thumbnails, slide show, presenter view, print, PDF, image and video export use the same slide renderer, so what you see is what you present and export. Text is laid out by QuillStage (not left to the browser) wherever line breaks must match between screen and export; the browser's text shaping is used inside each line.
- **Inheritance is resolved lazily.** Properties are looked up through shape → layout placeholder → master placeholder → master text styles → theme, with a per-slide cache invalidated by transactions.
- **Fonts.** Decks that name common proprietary fonts render with metric-compatible open fonts (Carlito, Caladea, Tinos, Arimo, Cousine, Gelasio, EB Garamond), as in the sister apps; installed originals are preferred; embedded fonts in .pptx are used when their embedding permissions allow.
- **Recovery is a pipeline** (below), shared in its generic parts with QuillPen and QuillGrid.

## Document formats

| Format | Plan |
|---|---|
| .pptx / .pptm / .ppsx / .potx | Read and write with JSZip and a streaming XML reader. Presentation, slides, layouts, masters, themes, notes, handouts, comments (classic and modern), charts (with embedded workbook), diagrams (data model, with the cached drawing used when a layout is unknown), media, tables, custom shows, sections, tags, custom XML. VBA project kept byte-for-byte in .pptm. |
| .ppt (97–2003) | Import only: compound-file reader shared with QuillPen's Word 97 recovery; record walker for the `PowerPoint Document` stream (persist directory, slides, masters, text atoms, shapes via the Office Drawing records), pictures from the `Pictures` stream. |
| .odp / .otp / .fodp | Read and write: draw pages, master pages, styles, presentation:animations (mapped to the timing model), transitions, notes, charts as tables of values plus rendering. |
| Markdown | `---` slide separator, title, bullets, images, code blocks, notes blocks, front matter for theme and size. |
| Native .qsp / .qspe | JSON package (zip) with the model and media; encrypted variant uses AES-256-GCM with a password-derived key, as in QuillPen and QuillGrid. |

## Rendering and the slide show

- **Edit surface:** slides render as DOM/SVG in a scaled container; handles and guides are an overlay; touch gestures for move, resize, rotate, crop and pinch zoom.
- **Thumbnails:** rendered at low resolution off the main thread where possible and cached per slide revision.
- **Transitions:** 2-D effects with the Web Animations API; 3-D effects (cube, gallery, vortex, fracture…) with WebGL on snapshots of the outgoing and incoming slides; Smooth Morph matches objects (name, then text, then type and geometry) and interpolates transforms, fills, outlines, geometry and text.
- **Animations:** the timing tree is compiled into a timeline of effects with triggers; the same timeline drives the show, the Animation Pane, the Build Order panel, the Animation Timeline and video export.
- **Presenter view and second screens:**

| Platform | Audience screen | Presenter screen |
|---|---|---|
| Windows, macOS, Linux (Electron) | A second `BrowserWindow`, full screen on the chosen display | The main window |
| Web | Window Management API where available, otherwise a second window placed by the user; the Presentation API for cast receivers | The opener window; state synchronized with `BroadcastChannel` |
| Android | The platform's external-display presentation support (HDMI, USB-C, wireless displays and casting) via the QuillNative plugin | The phone or tablet screen |
| iOS / iPadOS | An external-display window scene (AirPlay or cable) | The device screen |
| One screen (any platform) | — | Rehearse mode: presenter view alone |

- **Phone remote and audience features** use WebRTC data channels. Pairing is by QR code. A small signaling relay (to be published as a separate, self-hostable project) is needed only to connect devices that are not on the same network; on desktop, a local-network mode runs a tiny server in the Electron main process.

## Export

| Output | Plan |
|---|---|
| PDF | Vector PDF writer built on the shared renderer (text as text, with embedded open fonts; pictures as images). Slides, one page per build step, handouts (1–9 per page), notes pages, outline. A dependency such as pdf-lib (MIT) will be evaluated against writing a minimal writer. |
| Images | PNG and JPEG via canvas at chosen resolution; SVG from the renderer. |
| Video | Frames rendered to canvas and encoded with WebCodecs, muxed to WebM or MP4 (MIT-licensed muxers); narration and media mixed with Web Audio. Where WebCodecs is missing, a slower MediaRecorder path. |
| Animated GIF | Palette quantization and GIF encoding in a worker (MIT-licensed encoder). |
| HTML show | One self-contained file: deck JSON, media as data or sibling files, and a compact player with keyboard, touch, notes window and timings. |
| Story | Responsive scrolling HTML generated from the deck. |
| Web lecture | HTML show plus synchronized narration video and a slide list. |
| Handouts / outline | .docx/.odt through QuillPen's converters (shared package), .rtf, .md, .txt. |

## Recovery pipeline

QuillStage inherits the suite's recovery engine, which descends from the author's tools: byte-level ZIP salvage and the ImmortalInflate decoder of [Immortal Unzipper / Immortal Zip](https://socrtwo.github.io/Immortal-Zip/), XML healing and strings scans from [Word Recovery](https://github.com/socrtwo/wordrecovery-SF), and the S2 File Identifier. Recovery runs automatically when a file fails to open normally, and on purpose from **File ▸ Open ▸ Open and Repair**. The recovered deck opens as "*name* (Recovered)", unsaved, with a report; **Save** asks for a new name so the damaged original is never overwritten.

| Step | Method |
|---|---|
| a | **Identification** from the bytes: pptx/pptm/ppsx/potx, odp/fodp, ppt (compound file), qsp; concatenated or embedded data; encrypted packages reported. |
| b | **Normal open**; if CRC mismatches or malformed parts are found, the repair steps run too and compete. |
| c | **ZIP salvage** from local headers, ImmortalInflate at bit shifts, CRC checks, resynchronization on the next deflate block. |
| d | **XML repair and rebuild:** heal each `slideN.xml`; cut at the last complete shape (`</p:sp>`, `</p:pic>`, `</p:graphicFrame>`, `</p:grpSp>`) before damage; drop shapes whose relationships point at missing parts; rebuild `presentation.xml`'s slide list from the slides found (in numeric order, or the order recovered from `docProps/app.xml` titles); regenerate `[Content_Types].xml`, `_rels/.rels`, `presentation.xml.rels`; replace an unusable master or layout with a minimal one and re-point slides to it; drop a damaged timing tree rather than the slide. The rebuilt package can be saved as a clean .pptx. |
| e | **Slide-by-slide salvage:** each slide that can be parsed on its own becomes a slide in a new deck, so one broken slide never costs the others. |
| f | **Media harvest:** every picture, video and audio file that survives in `ppt/media/` (or the .ppt `Pictures` stream, or by signature scan of raw bytes: PNG, JPEG, GIF, EMF/WMF, MP4, MP3) is recovered; orphaned media go on "Recovered media" slides. |
| g | **Text scan:** `<a:t>` runs from any slide or notes fragment, grouped by slide and shape, rebuilt as title-and-content slides. |
| h | **.ppt record salvage:** lenient record walker that resynchronizes on known record headers; text atoms and pictures recovered even when the compound file's directory is damaged (raw scan for the stream's records). |
| i | **OpenDocument salvage:** heal `content.xml`, cut at the last complete `</draw:page>`, rebuild the package (mimetype first, manifest regenerated), otherwise raw scan of `draw:page` text and images. |
| j | **Strings scan:** printable ASCII and UTF-16LE runs as a last resort; labeled text-only recovery. |

Tests will build valid fixtures in code, damage them (truncation, flipped bytes, missing central directory, missing parts, dangling relationships, malformed XML), and check what is recovered, as QuillGrid's recovery tests do.

## AI

**Several AIs on one request** (same design as QuillPen, QuillGrid and QuillPost; `src/ai/ensemble.ts`):
the person can choose 2–6 configured providers. In *master* mode every member answers in parallel and
one chosen AI writes the final answer from the anonymized, shuffled responses. In *Quaker consensus*
mode a clerk drafts a minute from the members' answers; each member then unites with it, stands aside
(recorded concern) or blocks (principled objection); the clerk revises the minute for blocks, for up to
the configured number of rounds, and the outcome (unity, unity with stand-asides, or no unity) is shown
with the discussion. Only the final answer is ever applied to the presentation.

The AI layer is shared with the suite: a provider registry, an OpenAI-compatible client with streaming, per-provider adapters where needed, and a setup wizard (**AI ▸ Set Up AI**) grouped as *Free · on device*, *Free · your computer* (Ollama, LM Studio), *Free tier*, *Free trial* and *Paid*, which links to each provider's key page, tests the key, and fetches the model list.

- **On-device model:** a small open instruction model run with Transformers.js in a Web Worker (WebGPU when available, WebAssembly otherwise), downloaded once and then offline. Small models are good at rewriting, notes, alt text and outlines; drafting whole decks benefits from a larger local or hosted model, and the wizard says so.
- **Privacy:** nothing is sent until an AI command runs; only the slide, selection or deck text the command needs is sent; **AI & Privacy** shows what each command sends. Keys are stored encrypted with the Android Keystore or the OS keychain (Electron `safeStorage`); in a plain browser they stay in site storage, which is stated plainly.
- **Results** that change the deck are applied as one undoable step; generated decks are always outline first.

## Spelling and languages

Offline Hunspell dictionaries for every variety of English ship with the app (SCOWL, MIT and BSD-style). **Review ▸ Spelling Dictionaries** is a wizard that lists 90+ languages, shows each dictionary's size and license, downloads it on request, and lets the user set proofing languages per deck and per text selection. Spell checking runs in a worker (nspell / hunspell-asm), shared with the sister apps. Speech features (dictation, subtitles, coach) use on-device recognition where the platform offers it.

## Platforms and builds

| Target | Packaging | Notes |
|---|---|---|
| Web / PWA | Vite build in `dist/`, service worker, web app manifest | Offline-capable and installable; ChromeOS can install it |
| Android | Capacitor 8 (`android/`) | APK and AAB; *Open with* / *Share to* for presentation types; QuillNative plugin for file access, external displays, keystore |
| iOS / iPadOS | Capacitor 8 (`ios/`) | Unsigned .ipa from CI, signed by the user; Files and Share integration; external display scene |
| Windows | Electron + electron-builder | Installer (x64, arm64) and portable |
| macOS | Electron + electron-builder | Universal .dmg and .zip |
| Linux | Electron + electron-builder | AppImage and .deb for x64, arm64 and armv7l (Raspberry Pi OS) |
| ChromeOS | Android APK, Linux .deb, or the PWA | |

## Continuous integration and releases

As in QuillGrid:
- **CI** (`.github/workflows/android.yml`): on every push and pull request (docs-only changes skipped), type-check, unit tests, web build, then a Capacitor sync and Gradle build of the APK and AAB, uploaded as artifacts; signed with a release keystore from repository secrets when present, otherwise debug-signed with a warning.
- **Release** (`.github/workflows/release.yml`): on a tag `v*`, a manual run, or a commit whose message contains `[release]`: web zip, desktop installers for every OS and architecture, an unsigned iOS .ipa and the Android APK/AAB, published together as one GitHub release with SHA-256 checksums and a per-platform table; a failed platform is listed instead of blocking the others.
- Android `versionCode` derives from time so it always increases across workflows.

These workflows are added in milestone M1, together with the first code, so that CI never runs against an empty repository.

## Dependencies to evaluate

All must be MIT, BSD, Apache 2.0, ISC or similar permissive licenses, recorded in a generated third-party notices file as in the sister apps.

| Need | Candidate |
|---|---|
| ZIP | JSZip (MIT) |
| Spelling | nspell (MIT), hunspell-asm (MIT), SCOWL dictionaries |
| Equations | KaTeX (MIT) |
| Code highlighting | highlight.js (BSD-3-Clause) or Shiki (MIT) |
| PDF import | pdf.js (Apache 2.0) |
| PDF export | pdf-lib (MIT) or an in-house writer |
| Video muxing | mp4-muxer / webm-muxer (MIT) |
| GIF | gifenc (MIT) |
| On-device AI | Transformers.js (Apache 2.0) |
| Collaboration (M8) | Yjs (MIT) |
| Icons | Material Symbols (Apache 2.0) |
| Fonts | Carlito, Caladea, Gelasio, EB Garamond and others (SIL OFL 1.1); Tinos, Arimo, Cousine (Apache 2.0) |
| Native | Capacitor (MIT), Electron (MIT), electron-builder (MIT) |
