# QuillStage

An original, open-source presentation app for Android, iOS, Windows, macOS, Linux, ChromeOS and the web whose goal is the same as the rest of the Quill suite: **no missing commands**. QuillStage will give you a familiar slide workflow (outline, slide sorter, masters, transitions, animations, presenter view) and open and save .pptx and .odp files faithfully, then add the best ideas from forty years of presentation software: outline-driven authoring, a zoomable canvas mode, Markdown slides, build-order editing, rehearsal coaching, live polls and Q&A, and an optional AI tab that can run entirely on your device. Every feature is a command, and every command can be put on the Quick Access Bar, on your own tab, on a keyboard shortcut, or in a macro, and found with command search. When a presentation is damaged, QuillStage recovers what it can instead of refusing to open it.

> **Status: planning — no code yet.** This repository holds the plan: the feature survey, the command inventory, the intended architecture and the roadmap. Nothing here is built or released. Issues and discussion about the plan are welcome; see [CONTRIBUTING.md](CONTRIBUTING.md).

## The plan in brief

- **A complete command set.** Every command a desktop presentation program offers, tab by tab (File, Home, Insert, Draw, Design, Transitions, Animations, Slide Show, Record, Review, View, Help, and the contextual Shape, Picture, Table, Chart, Diagram, Media and Ink tabs), plus commands adopted from other presentation tools. See [docs/COMMANDS-PLAN.md](docs/COMMANDS-PLAN.md).
- **No missing commands.** Commands are data. The ribbon, the Quick Access Bar, custom tabs, context menus, keyboard shortcuts, macros and command search (Alt+Q) all refer to command ids, so anything QuillStage can do can be placed anywhere, including the many commands desktop programs keep out of their tabs.
- **Quick Access Bar** with the full desktop drop-down checklist (AutoSave, New, Open, Save, Email, Quick Print, Print Preview and Print, Spelling, Undo, Redo, Present from Beginning, Touch/Mouse Mode, and the items older versions offered) and **Add All Choices**, which puts Save and Print first and everything else in alphabetical order. The bar can sit at the top or at the bottom near the keyboard. See [docs/QAT.md](docs/QAT.md).
- **Customize Tabs and Keyboard:** create tabs and groups, put any command in them, rename, reorder or hide built-in tabs, and rebind any command with conflict display. Layouts can be exported and imported.
- **Innovations from other presentation software:** outline-first authoring with automatic layouts, a build-order panel, Smooth Morph transitions, zoom links and a zoomable canvas mode, Markdown and plain-text decks, vertical slide stacks, code slides with step-through highlighting, presenter rehearsal with timing and a rehearsal coach, live polls, word clouds, quizzes and audience Q&A, a phone remote, Pecha Kucha and Ignite timers, and more. Each idea is credited in [docs/INNOVATIONS.md](docs/INNOVATIONS.md).
- **Presenter view** on a second screen (or the same screen on a phone) with notes, the next slide or build, a timer, a slide grid, pen, highlighter and laser pointer, and live subtitles.
- **Export** to PDF (with or without build steps), images, video, animated GIF, handouts, a self-contained HTML show, a responsive scrolling story, and the outline to QuillPen.
- **AI tab (optional).** Draft a deck from a prompt or a document, outline to slides, design ideas, speaker notes, rewrite and shorten text, alt text, translation and a rehearsal coach. A setup wizard covers free and paid providers, local servers, and a **free on-device model** that needs no account and works offline after one download. Several AIs can work on one request: one AI combines the others' answers, or they deliberate Quaker-style (unite, stand aside or block) until they agree. Nothing is sent anywhere until you run an AI command.
- **Spelling in many languages.** Offline dictionaries for every variety of English, and a **Spelling Dictionaries** wizard that downloads freely licensed dictionaries for 90+ other languages when you need them.
- **Corrupt-file recovery.** Damaged .pptx, .ppsx, .potx, .odp and .ppt files go through a recovery pipeline descended from the author's [Immortal Unzipper / Immortal Zip](https://socrtwo.github.io/Immortal-Zip/) and [Word Recovery](https://github.com/socrtwo/wordrecovery-SF) tools: file identification from the bytes, byte-level ZIP salvage, XML repair, slide-by-slide rebuilding, media harvesting and text scans. The recovered deck opens as a new, unsaved file with a report.

## File formats (planned)

| Format | Open | Save | Notes |
|---|---|---|---|
| PowerPoint presentation, show, template (.pptx, .ppsx, .potx) | Yes | Yes | Office Open XML (PresentationML) |
| Macro-enabled (.pptm, .ppsm, .potm) | Yes | Yes | VBA kept, never run |
| PowerPoint 97–2003 (.ppt, .pps, .pot) | Yes | No | Import from the published binary specification; save as .pptx |
| OpenDocument presentation (.odp, .otp, .fodp) | Yes | Yes | ISO/IEC 26300 |
| Markdown slides (.md) | Yes | Yes | `---` separates slides; front matter for theme and size |
| Plain-text outline (.txt), outline (.rtf, .docx) | Yes | Yes | Headings become slide titles |
| PDF | As image slides | Yes | Export with or without build steps; handouts and notes pages |
| Images (.png, .jpg, .svg) | Insert | Yes | One file per slide |
| Video (.webm, .mp4 where the platform has an encoder), animated GIF | Insert | Yes | Timings, transitions, animations, narration |
| HTML | No | Yes | Self-contained player or responsive scrolling story |
| QuillStage presentation (.qsp) and encrypted (.qspe) | Yes | Yes | Native JSON package; AES-256 for .qspe |

## Platforms (planned)

Each release will have builds for:
- **Android** (APK and AAB);
- **iOS/iPadOS** (an unsigned .ipa you sign with your Apple ID);
- **Windows** (installer for x64 or ARM, or portable);
- **macOS** (universal);
- **Linux** (AppImage or .deb for x64, arm64 and armv7l, which covers **Raspberry Pi OS**);
- **ChromeOS** (Android app, Linux app or web app);
- the **web** (static files; offline-capable and installable).

On Android, QuillStage will register for *Open with* and *Share to* on presentation files and use the system file picker, so files on the device, SD card and cloud providers all work. **Save** writes back to the same file.

CI will type-check, test and build the web app and an APK on every push, and a **Release** workflow will build every platform and publish them together from a tag `v*`, a manual run, or a `[release]` commit, as in the sister apps. See [docs/ARCHITECTURE-PLAN.md](docs/ARCHITECTURE-PLAN.md) and [docs/ROADMAP.md](docs/ROADMAP.md).

## The Quill suite

| App | What it is | Repository |
|---|---|---|
| **QuillPen** | Word processor | https://github.com/socrtwo/QuillPen |
| **QuillGrid** | Spreadsheet | https://github.com/socrtwo/QuillGrid |
| **QuillPost** | Email | https://github.com/socrtwo/QuillPost |
| **QuillStage** | Presentations (this project) | https://github.com/socrtwo/QuillStage |

The apps share one application shell: the command registry, Quick Access Bar, customizable tabs and keyboard, command search, macro recorder, AI setup wizard and providers, spelling dictionaries, recovery engine and platform layer. QuillStage plans to link charts from QuillGrid workbooks, send outlines and handouts to QuillPen, and share decks through QuillPost.

## Documentation

| File | Contents |
|---|---|
| [docs/INNOVATIONS.md](docs/INNOVATIONS.md) | Survey of presentation software, past and present, and what QuillStage adopts from each |
| [docs/QAT.md](docs/QAT.md) | The Quick Access Bar plan and the drop-down choices by version |
| [docs/COMMANDS-PLAN.md](docs/COMMANDS-PLAN.md) | The planned command inventory, tab by tab, with command ids |
| [docs/ARCHITECTURE-PLAN.md](docs/ARCHITECTURE-PLAN.md) | Intended architecture: shell, document model, rendering, presenter view, export, AI, recovery, builds |
| [docs/ROADMAP.md](docs/ROADMAP.md) | Milestones |
| [docs/LEGAL.md](docs/LEGAL.md) | Clean-room, trademark and licensing notes |

## Project language

The interface, code comments and documentation use American English. Spelling checks will accept every variety of English by default.

## License

[MIT](LICENSE). Copyright (c) 2026 Paul D Pruitt.

## Legal

QuillStage is an independent, clean-room project released under the MIT License. It is not affiliated with, endorsed by, or sponsored by Microsoft, Apple, Google, Prezi, Canva, The Document Foundation, any AI provider, or any other vendor. Product names appear only to describe file-format compatibility or the origin of an idea (nominative use). PowerPoint is a trademark of Microsoft Corporation; Keynote is a trademark of Apple Inc.; other names belong to their owners. The interface, icons, themes, templates and text will be original or openly licensed. See [docs/LEGAL.md](docs/LEGAL.md).
