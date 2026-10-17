# Command inventory plan

> **Plan.** QuillStage has no code yet. This is the planned command inventory: every command of a desktop presentation program, tab by tab, with the command id QuillStage will use, followed by the commands adopted from other software in [INNOVATIONS.md](INNOVATIONS.md). When the app exists, `npm run catalog` will generate `docs/COMMANDS.md` from the command registry, and this plan will be retired.

Every command will be usable on the Quick Access Bar, on custom tabs, on a keyboard shortcut, in a macro and in command search (Alt+Q). Commands marked *not in tabs* have no default place on a tab but are reachable everywhere else, so nothing is hidden.

## Conventions

- **ids** are `area.name` in camelCase, as in QuillPen and QuillGrid. Where the sister apps already define a command (`file.save`, `edit.undo`, `review.spelling`, `view.cmdSearch`, `view.qatAddAll`…), QuillStage reuses the id so that shared settings, macros and QAT layouts carry over.
- **Galleries** (themes, transitions, animations, shape styles) are one command with an argument, written `id(arg)`, plus one command per item where it is useful to place a single item on the QAT, written as a range: `trans.fade`, `trans.push`, ….
- **Shortcuts** are planned defaults. Many follow long-standing desktop conventions because users expect them; all can be rebound.
- **Vendor feature names** are replaced with generic terms; see [LEGAL.md](LEGAL.md). The left column says what the user sees in QuillStage.

| Namespace | Area |
|---|---|
| `file.` | File screen (Backstage): new, open, save, export, print, share, info, options |
| `edit.` | Clipboard, undo, find, select |
| `fmt.` | Character and paragraph formatting |
| `slide.` / `section.` / `outline.` | Slides, sections, outline editing |
| `insert.` | Insert tab |
| `draw.` / `ink.` | Draw tab and ink |
| `design.` | Design tab: themes, variants, size, background, design ideas |
| `trans.` | Transitions |
| `anim.` | Animations |
| `show.` | Slide Show |
| `record.` | Record |
| `review.` | Review |
| `view.` / `master.` | View, masters |
| `arrange.` | Order, align, group, rotate (shared by all contextual tabs) |
| `shape.`, `picture.`, `table.`, `chart.`, `diagram.`, `media.`, `model3d.`, `zoom.`, `equation.`, `code.`, `canvas.` | Contextual tabs and objects |
| `ai.` | AI tab |
| `tools.`, `dev.`, `help.` | Tools, Developer, Help |

---

## File (Backstage)

| Command | Group | id | Shortcut | Notes |
|---|---|---|---|---|
| New Presentation | New | `file.new` | Ctrl+N | not in tabs |
| New from Template… | New | `file.newFromTemplate` | | not in tabs |
| New from Content Starter… | New | `file.newFromStarter` | | Freelance-style guided outlines |
| Open… | Open | `file.open` | Ctrl+O, Ctrl+F12 | |
| Open Recent | Open | `file.openRecent` | | QAT legacy choice "Open Recent File" |
| Clear Unpinned Recent Presentations | Open | `file.clearRecent` | | |
| Open and Repair… | Open | `file.openRepair` | | runs every recovery method |
| Recover Data from Any File… | Open | `file.recoverAny` | | |
| Recovery Report | Open | `file.recoveryReport` | | |
| Manual Recovery Options | Open | `file.recoveryHelp` | | |
| Open Encrypted Presentation… | Open | `file.openEncrypted` | | .qspe |
| Save | Save | `file.save` | Ctrl+S, Shift+F12 | |
| Save As… | Save | `file.saveAs` | F12 | |
| Save a Copy… | Save | `file.saveCopy` | | |
| AutoSave | Save | `file.autosave` | | toggle |
| Revert to Saved | Info | `file.revert` | | |
| Properties… | Info | `file.properties` | | title, author, tags, custom properties |
| Protect: Always Open Read-Only | Info | `file.alwaysReadOnly` | | |
| Protect: Encrypt with Password… | Info | `file.encrypt` | | AES-256 .qspe; standard OOXML encryption for .pptx (M8) |
| Protect: Add a Digital Signature | Info | `file.sign` | | M8 |
| Protect: Mark as Final | Info | `file.markFinal` | | |
| Inspect Presentation… | Info | `file.inspect` | | hidden slides, notes, comments, off-slide content, invisible objects, custom XML, properties |
| Check Accessibility | Info | `review.checkAccessibility` | | also on Review |
| Check Compatibility | Info | `file.checkCompatibility` | | features lost in .ppt/.odp |
| Version History | Info | `file.versions` | | |
| Take Snapshot | Info | `file.snapshot` | Ctrl+Alt+5 | |
| Recover Unsaved Presentations | Info | `file.recoverUnsaved` | | |
| Compress Media… | Info | `file.compressMedia` | | |
| Optimize Media Compatibility | Info | `file.optimizeMedia` | | transcode to widely playable codecs |
| Compress Presentation… | Info | `file.compress` | | Impress Minimizer-style; pictures, masters, media |
| Embed Fonts in File | Info | `file.embedFonts` | | only fonts whose licenses permit embedding |
| Print… / Print Preview and Print | Print | `file.print` | Ctrl+P | slides, handouts (1–9 per page), notes pages, outline; frames; grayscale |
| Quick Print | Print | `file.quickPrint` | | |
| Share… | Share | `file.share` | | system share sheet |
| Email as Attachment | Share | `file.email` | | uses QuillPost when installed |
| Email as PDF | Share | `file.emailPdf` | | |
| Present Online / Broadcast Show | Share | `show.broadcast` | | M8 |
| Export as PDF… | Export | `file.export.pdf` | | options: slides, handouts, notes, outline, one page per build step |
| Export Handouts PDF | Export | `file.export.handoutsPdf` | | |
| Create a Video… | Export | `file.export.video` | | resolution, timings, narration, captions |
| Create Animated GIF… | Export | `file.export.gif` | | |
| Export Slides as Pictures… | Export | `file.export.images` | | PNG, JPEG, SVG; current or all |
| Export HTML Show | Export | `file.export.htmlShow` | | self-contained player |
| Export Story (responsive web page) | Export | `file.export.story` | | Sway-style |
| Export Web Lecture | Export | `file.export.webLecture` | | Producer-style |
| Export Handout Document (to QuillPen) | Export | `file.export.handoutDoc` | | .docx/.odt with slides and notes |
| Export Outline… | Export | `file.export.outline` | | .rtf, .docx, .md, .txt |
| Export Markdown Slides | Export | `file.export.markdown` | | |
| Export as PowerPoint Presentation (.pptx) | Export | `file.export.pptx` | | |
| Export as PowerPoint Show (.ppsx) | Export | `file.export.ppsx` | | |
| Export as PowerPoint Template (.potx) | Export | `file.export.potx` | | |
| Export as Macro-Enabled Presentation (.pptm) | Export | `file.export.pptm` | | VBA preserved |
| Export as OpenDocument Presentation (.odp) | Export | `file.export.odp` | | |
| Export as Flat OpenDocument (.fodp) | Export | `file.export.fodp` | | |
| Export as QuillStage Presentation (.qsp / .qspe) | Export | `file.export.qsp`, `file.export.qspe` | | |
| Export LaTeX Beamer (.tex) | Export | `file.export.beamer` | | stretch, M8 |
| Package Presentation… | Export | `file.package` | | Package for CD equivalent |
| Import Markdown Slides… | Import | `file.import.markdown` | | |
| Text to Deck… | Import | `file.import.textDeck` | | each paragraph a slide |
| Slides from Outline… | Import | `insert.slidesFromOutline` | | also on Insert |
| Close | Close | `file.close` | Ctrl+W, Ctrl+F4 | |
| Close All | Close | `file.closeAll` | | |
| Exit | Close | `file.exit` | | desktop |
| Options… | Options | `file.options` | | |
| Customize Quick Access Bar… | Options | `file.customizeQat` | | |
| Customize Tabs… | Options | `file.customizeRibbon` | | |
| Customize Keyboard… | Options | `file.customizeKeyboard` | | |
| AutoCorrect Options… | Options | `file.autocorrectOptions` | | |
| Custom Dictionary… | Options | `file.customDictionary` | | |
| Export / Import Settings… | Options | `file.exportSettings` | | shared format across the suite |
| Reset All Customizations | Options | `file.resetCustomizations` | | |
| Account / About | Options | `help.about` | | |

## Home

| Command | Group | id | Shortcut | Notes |
|---|---|---|---|---|
| Paste | Clipboard | `edit.paste` | Ctrl+V | |
| Paste: Use Destination Theme | Clipboard | `edit.pasteDestTheme` | | |
| Paste: Keep Source Formatting | Clipboard | `edit.pasteKeepSource` | | |
| Paste as Picture | Clipboard | `edit.pastePicture` | | |
| Paste: Keep Text Only | Clipboard | `edit.pasteText` | Ctrl+Shift+Alt+V | |
| Paste Special… | Clipboard | `edit.pasteSpecial` | Ctrl+Alt+V | |
| Cut | Clipboard | `edit.cut` | Ctrl+X | |
| Copy | Clipboard | `edit.copy` | Ctrl+C | |
| Duplicate | Clipboard | `edit.duplicate` | Ctrl+D | objects, or slides in the thumbnails |
| Format Painter | Clipboard | `edit.formatPainter` | | double-tap to keep on |
| Copy Formatting / Paste Formatting | Clipboard | `edit.copyFormat`, `edit.pasteFormat` | Ctrl+Shift+C, Ctrl+Shift+V | |
| Clipboard Pane | Clipboard | `edit.clipboardPane` | | |
| Undo | — | `edit.undo` | Ctrl+Z | |
| Redo / Repeat | — | `edit.redo` | Ctrl+Y | |
| New Slide | Slides | `slide.new` | Ctrl+M | with layout gallery |
| New Slide with Layout | Slides | `slide.newWithLayout(layout)` | | one command per layout via argument |
| Duplicate Slide | Slides | `slide.duplicate` | | |
| Delete Slide | Slides | `slide.delete` | | |
| Layout | Slides | `slide.layout(layout)` | | |
| Reset Slide | Slides | `slide.reset` | | |
| Reuse Slides… | Slides | `slide.reuse` | | from another .pptx/.odp/.qsp |
| Slides from Outline… | Slides | `insert.slidesFromOutline` | | |
| Add Section | Slides | `section.add` | | |
| Rename Section | Slides | `section.rename` | | |
| Remove Section | Slides | `section.remove` | | |
| Remove Section & Slides | Slides | `section.removeWithSlides` | | |
| Remove All Sections | Slides | `section.removeAll` | | |
| Move Section Up / Down | Slides | `section.moveUp`, `section.moveDown` | | |
| Collapse All / Expand All Sections | Slides | `section.collapseAll`, `section.expandAll` | | |
| Hide Slide | Slides | `slide.hide` | | also on Slide Show |
| Font | Font | `fmt.font` | Ctrl+Shift+F | |
| Font Size | Font | `fmt.fontSize` | Ctrl+Shift+P | |
| Increase Font Size | Font | `fmt.growFont` | Ctrl+Shift+> | |
| Decrease Font Size | Font | `fmt.shrinkFont` | Ctrl+Shift+< | |
| Clear All Formatting | Font | `fmt.clearFormatting` | Ctrl+Space | |
| Bold | Font | `fmt.bold` | Ctrl+B | |
| Italic | Font | `fmt.italic` | Ctrl+I | |
| Underline (styles) | Font | `fmt.underline` | Ctrl+U | |
| Text Shadow | Font | `fmt.shadow` | | |
| Strikethrough / Double Strikethrough | Font | `fmt.strike`, `fmt.doubleStrike` | | |
| Superscript / Subscript | Font | `fmt.superscript`, `fmt.subscript` | Ctrl+Shift+=, Ctrl+= | |
| Character Spacing | Font | `fmt.charSpacing` | | Very Tight … Very Loose, custom |
| Change Case | Font | `fmt.changeCase` | Shift+F3 | sentence, lower, upper, capitalize, toggle |
| Small Caps / All Caps | Font | `fmt.smallCaps`, `fmt.allCaps` | | |
| Font Color | Font | `fmt.fontColor` | | |
| Text Highlight Color | Font | `fmt.highlight` | | |
| Font Dialog… | Font | `fmt.fontDialog` | Ctrl+T | |
| Replace Fonts… | Editing | `edit.replaceFonts` | | |
| Bullets | Paragraph | `fmt.bullets` | | custom bullets, picture bullets |
| Numbering | Paragraph | `fmt.numbering` | | |
| Decrease / Increase List Level | Paragraph | `fmt.outdent`, `fmt.indent` | Shift+Tab, Tab | |
| Line Spacing | Paragraph | `fmt.lineSpacing` | | |
| Align Left / Center / Right / Justify | Paragraph | `fmt.alignLeft`, `fmt.alignCenter`, `fmt.alignRight`, `fmt.justify` | Ctrl+L, Ctrl+E, Ctrl+R, Ctrl+J | |
| Distributed | Paragraph | `fmt.distribute` | | |
| Columns | Paragraph | `fmt.columns` | | |
| Text Direction | Paragraph | `fmt.textDirection` | | horizontal, rotate 90/270, stacked |
| Align Text (Top / Middle / Bottom) | Paragraph | `fmt.verticalAlign` | | |
| Convert to Diagram | Paragraph | `diagram.fromOutline` | | SmartArt equivalent |
| Paragraph Dialog… | Paragraph | `fmt.paragraphDialog` | | indents, spacing, tabs |
| Tabs… | Paragraph | `fmt.tabs` | | |
| Right-to-Left / Left-to-Right | Paragraph | `fmt.rtl`, `fmt.ltr` | | |
| Shapes gallery | Drawing | `insert.shape(type)` | | |
| Arrange | Drawing | `arrange.*` | | see Arrange below |
| Quick Styles | Drawing | `shape.quickStyle(style)` | | |
| Shape Fill / Outline / Effects | Drawing | `shape.fill`, `shape.outline`, `shape.effects` | | |
| Find… | Editing | `edit.find` | Ctrl+F | |
| Replace… | Editing | `edit.replace` | Ctrl+H | regex option |
| Select All | Editing | `edit.selectAll` | Ctrl+A | |
| Select Objects (marquee) | Editing | `edit.selectObjects` | | |
| Selection Pane | Editing | `home.selectionPane` | Alt+F10 | rename, show/hide, reorder |
| Select Same Type / Same Fill / Same Outline | Editing | `edit.selectSimilar(kind)` | | not in tabs |
| Dictate | Voice | `tools.dictate` | Alt+` | on-device speech where available |
| Design Ideas | Designer | `design.designIdeas` | | also on Design |

## Arrange (shared by Home ▸ Drawing and every Format tab)

| Command | id | Shortcut |
|---|---|---|
| Bring to Front / Bring Forward | `arrange.bringToFront`, `arrange.bringForward` | Ctrl+Shift+], Ctrl+] |
| Send to Back / Send Backward | `arrange.sendToBack`, `arrange.sendBackward` | Ctrl+Shift+[, Ctrl+[ |
| Group / Ungroup / Regroup | `arrange.group`, `arrange.ungroup`, `arrange.regroup` | Ctrl+G, Ctrl+Shift+G |
| Align Left / Center / Right / Top / Middle / Bottom | `arrange.alignLeft` … `arrange.alignBottom` | |
| Distribute Horizontally / Vertically | `arrange.distributeH`, `arrange.distributeV` | |
| Align to Slide / Align Selected Objects | `arrange.alignToSlide`, `arrange.alignToSelection` | |
| Rotate Right 90° / Left 90° / Flip Vertical / Flip Horizontal | `arrange.rotateRight`, `arrange.rotateLeft`, `arrange.flipV`, `arrange.flipH` | Alt+Right, Alt+Left |
| More Rotation Options… | `arrange.rotateDialog` | |
| Lock / Unlock Position | `arrange.lock`, `arrange.unlock` | |
| Nudge (1 px / grid) | `arrange.nudge(dir)` | arrows, Ctrl+arrows |
| Match Size / Match Position (from clipboard) | `arrange.matchSize`, `arrange.pastePosition` | not in tabs |

## Insert

| Command | Group | id | Shortcut | Notes |
|---|---|---|---|---|
| New Slide | Slides | `slide.new` | Ctrl+M | |
| Reuse Slides… | Slides | `slide.reuse` | | |
| Table | Tables | `insert.table` | | grid picker and dialog |
| Draw Table | Tables | `insert.drawTable` | | |
| Spreadsheet Table (QuillGrid) | Tables | `insert.gridObject` | | embedded workbook range |
| Pictures from This Device… | Images | `insert.picture` | | |
| Online Pictures (openly licensed)… | Images | `insert.onlinePictures` | | records license and attribution |
| Screenshot / Screen Clipping | Images | `insert.screenshot` | | desktop and web where allowed |
| Photo Album… | Images | `insert.photoAlbum` | | new deck from many pictures, captions, frames |
| Insert PDF Pages as Slides… | Images | `insert.pdfPages` | | |
| Live Camera | Camera | `insert.liveCamera` | | Cameo equivalent |
| Shapes | Illustrations | `insert.shape(type)` | | full preset geometry set (about 180 shapes) |
| Icons | Illustrations | `insert.icons` | | Material Symbols |
| 3D Models… | Illustrations | `insert.model3d` | | glTF/GLB; M8 |
| Diagram… | Illustrations | `insert.diagram` | | SmartArt equivalent: list, process, cycle, hierarchy, relationship, matrix, pyramid, picture |
| Chart… | Illustrations | `insert.chart` | | column, bar, line, area, pie, doughnut, scatter, bubble, radar, stock, surface, histogram, box and whisker, waterfall, funnel, treemap, sunburst, map (M8), combo |
| Linked Chart from QuillGrid… | Illustrations | `insert.linkedChart` | | |
| Sparkline | Illustrations | `insert.sparkline` | | |
| Poll / Word Cloud / Quiz / Rating | Interactive | `insert.poll`, `insert.wordCloud`, `insert.quiz`, `insert.rating` | | replaces Forms |
| Summary Zoom / Section Zoom / Slide Zoom | Links | `insert.summaryZoom`, `insert.sectionZoom`, `insert.slideZoom` | | |
| Link… | Links | `insert.link` | Ctrl+K | URL, slide, file, email |
| Action… | Links | `insert.action` | | on click / mouse over: go to, run macro, play sound, open file |
| Action Buttons | Links | `insert.actionButton(kind)` | | back, forward, home, end, return, info, sound, video, custom |
| Comment | Comments | `review.newComment` | Ctrl+Alt+M | |
| Text Box | Text | `insert.textBox` | | horizontal, vertical |
| Header & Footer… | Text | `insert.headerFooter` | | date, slide number, footer; notes and handouts |
| Decorative Text | Text | `insert.decorativeText` | | WordArt equivalent |
| Date & Time… | Text | `insert.dateTime` | | |
| Slide Number | Text | `insert.slideNumber` | | |
| Object… | Text | `insert.object` | | embed or link a file; OLE objects open as their picture |
| Code Block | Text | `insert.codeBlock` | | |
| Sticky Note | Text | `insert.stickyNote` | | |
| Equation | Symbols | `insert.equation` | Alt+= | LaTeX and linear input |
| Ink Equation | Symbols | `ink.toMath` | | |
| Symbol… | Symbols | `insert.symbol` | | |
| Emoji | Symbols | `insert.emoji` | | |
| Video from This Device… | Media | `insert.video` | | |
| Online Video / Web Embed… | Media | `insert.webEmbed` | | |
| Audio from This Device… | Media | `insert.audio` | | |
| Record Audio… | Media | `insert.recordAudio` | | |
| Screen Recording… | Media | `insert.screenRecording` | | desktop and web; Android with MediaProjection |
| Linked Component… | Components | `slide.insertComponent` | | Ludus smart-block equivalent |

## Draw

| Command | Group | id | Notes |
|---|---|---|---|
| Select | Tools | `draw.select` | |
| Lasso Select | Tools | `draw.lasso` | |
| Eraser (stroke / point / segment) | Tools | `draw.eraser` | |
| Pen / Pencil / Highlighter (with color, thickness, effects) | Pens | `draw.pen`, `draw.pencil`, `draw.highlighter` | |
| Add Pen | Pens | `draw.addPen` | |
| Draw with Touch | Tools | `draw.withTouch` | |
| Ruler (straightedge) | Stencils | `draw.ruler` | |
| Protractor | Stencils | `draw.protractor` | |
| Ink to Shape | Convert | `ink.toShape` | |
| Ink to Text | Convert | `ink.toText` | handwriting recognition where available |
| Ink to Math | Convert | `ink.toMath` | |
| Ink Replay / Rewind / Forward | Replay | `ink.replay`, `ink.rewind`, `ink.forward` | |
| Hide Ink / Show Ink | Ink | `ink.toggleVisibility` | |
| Delete All Ink on Slide | Ink | `ink.deleteAll` | |
| Format Background | Background | `design.formatBackground` | |

## Design

| Command | Group | id | Notes |
|---|---|---|---|
| Themes gallery | Themes | `design.theme(name)` | original themes only |
| Browse for Themes… | Themes | `design.browseThemes` | from .potx, .thmx, .otp, .qsp |
| Save Current Theme… | Themes | `design.saveTheme` | |
| Apply to Selected Slides / All Slides | Themes | `design.applyThemeSelected`, `design.applyThemeAll` | |
| Variants gallery | Variants | `design.variant(n)` | |
| Theme Colors / Customize Colors… | Variants | `design.colors`, `design.customColors` | |
| Theme Fonts / Customize Fonts… | Variants | `design.fonts`, `design.customFonts` | |
| Theme Effects | Variants | `design.effects` | |
| Background Styles | Variants | `design.backgroundStyles` | |
| Slide Size: Standard 4:3 / Widescreen 16:9 / Custom… | Customize | `design.slideSize(preset)` | |
| Resize Presentation (Maximize / Ensure Fit / Reflow) | Customize | `design.resizeTo(preset)` | Canva-style reflow |
| Format Background… | Customize | `design.formatBackground` | solid, gradient, picture, texture, pattern, hide background graphics |
| Background Video… | Customize | `design.backgroundVideo` | |
| Apply Background to All | Customize | `design.backgroundApplyAll` | |
| Reset Background | Customize | `design.backgroundReset` | |
| Design Ideas | Designer | `design.designIdeas` | |
| Restyle Deck | Designer | `design.restyle` | |
| Layout by Outline Level… | Designer | `design.autoLayoutRules` | Persuasion AutoTemplates |
| Brand Kit… | Brand | `design.brandKit` | |

## Transitions

| Command | Group | id | Notes |
|---|---|---|---|
| Preview | Preview | `trans.preview` | |
| Transition gallery | Transition to This Slide | `trans.apply(effect)` | |
| None | Subtle | `trans.none` | |
| Smooth Morph | Subtle | `trans.smoothMorph` | Morph / Magic Move equivalent |
| Fade, Push, Wipe, Split, Reveal, Cut, Random Bars, Shape, Uncover, Cover, Flash, Strips | Subtle | `trans.fade`, `trans.push`, `trans.wipe`, `trans.split`, `trans.reveal`, `trans.cut`, `trans.randomBars`, `trans.shape`, `trans.uncover`, `trans.cover`, `trans.flash`, `trans.strips` | |
| Fall Over, Drape, Curtains, Wind, Prestige, Fracture, Crush, Peel Off, Page Curl, Airplane, Origami, Dissolve, Checkerboard, Blinds, Clock, Ripple, Honeycomb, Glitter, Vortex, Shred, Switch, Flip, Gallery, Cube, Doors, Box, Comb, Zoom, Random | Exciting | `trans.fallOver` … `trans.random` | one id per effect, same naming pattern |
| Pan, Ferris Wheel, Conveyor, Rotate, Window, Orbit, Fly Through | Dynamic Content | `trans.pan` … `trans.flyThrough` | |
| Effect Options | Transition to This Slide | `trans.effectOptions(option)` | direction, shape, by object/word/character for Smooth Morph |
| Sound | Timing | `trans.sound(sound)` | original sound set; loop until next sound |
| Duration | Timing | `trans.duration` | |
| Apply To All | Timing | `trans.applyAll` | |
| Advance Slide: On Mouse Click | Timing | `trans.advanceOnClick` | |
| Advance Slide: After (time) | Timing | `trans.advanceAfter` | |
| Remove Transitions from All Slides | Timing | `trans.removeAll` | not in tabs |

## Animations

| Command | Group | id | Notes |
|---|---|---|---|
| Preview | Preview | `anim.preview` | |
| Animation gallery | Animation | `anim.apply(effect)` | replaces existing |
| Entrance: Appear, Fade, Fly In, Float In, Split, Wipe, Shape, Wheel, Random Bars, Grow & Turn, Zoom, Swivel, Bounce (and the extended set: Blinds, Box, Checkerboard, Circle, Diamond, Dissolve In, Peek In, Plus, Strips, Wedge, Expand, Compress, Rise Up, Center Revolve, Spinner, Stretch, Basic Swivel, Boomerang, Credits, Curve Up, Drop, Flip, Pinwheel, Spiral In, Thread, Whip…) | Animation | `anim.in.appear` … `anim.in.whip` | one id per effect |
| Emphasis: Pulse, Color Pulse, Teeter, Spin, Grow/Shrink, Desaturate, Darken, Lighten, Transparency, Object Color, Complementary Color, Line Color, Fill Color, Brush Color, Font Color, Underline, Bold Flash, Bold Reveal, Wave | Animation | `anim.em.pulse` … `anim.em.wave` | |
| Exit: (mirror of Entrance) | Animation | `anim.out.disappear` … | |
| Motion Paths: Lines, Arcs, Turns, Shapes, Loops, Custom Path | Animation | `anim.path.line` … `anim.path.custom` | |
| Media actions: Play, Pause, Stop, Seek | Animation | `anim.media.play`, `anim.media.pause`, `anim.media.stop`, `anim.media.seek` | |
| Effect Options | Animation | `anim.effectOptions(option)` | direction, sequence (as one object, all at once, by paragraph) |
| Effect Options Dialog… | Animation | `anim.effectDialog` | sound, after animation, animate text by word/letter, smooth start/end, bounce end, repeat, rewind |
| Add Animation | Advanced Animation | `anim.add(effect)` | adds without replacing |
| Animation Pane | Advanced Animation | `anim.pane` | |
| Trigger: On Click of / On Bookmark | Advanced Animation | `anim.trigger` | |
| Animation Painter | Advanced Animation | `anim.painter` | Alt+Shift+C |
| Start: On Click / With Previous / After Previous | Timing | `anim.startOnClick`, `anim.startWithPrevious`, `anim.startAfterPrevious` | |
| Duration / Delay | Timing | `anim.duration`, `anim.delay` | |
| Move Earlier / Move Later | Timing | `anim.moveEarlier`, `anim.moveLater` | |
| Remove Animation / Remove All on Slide | Timing | `anim.remove`, `anim.removeAll` | |
| Build Order | Advanced Animation | `anim.buildOrder` | Keynote-style panel |
| Animation Timeline | Advanced Animation | `anim.timeline` | Astound-style tracks |
| Reveal by Step… | Advanced Animation | `anim.revealSteps` | Beamer overlays |

## Slide Show

| Command | Group | id | Shortcut | Notes |
|---|---|---|---|---|
| From Beginning / Present from Beginning | Start Slide Show | `show.fromStart` | F5 | |
| From Current Slide | Start Slide Show | `show.fromCurrent` | Shift+F5 | |
| Presenter View | Start Slide Show | `show.presenterView` | Alt+F5 | |
| Rehearse (presenter view on one screen) | Start Slide Show | `show.rehearse` | | |
| Custom Slide Show… | Start Slide Show | `show.customShows` | | create, edit, copy, play |
| Broadcast Show | Start Slide Show | `show.broadcast` | | M8 |
| Audience Link | Start Slide Show | `show.audienceLink` | | M8 |
| Cast to Display | Start Slide Show | `show.cast` | | |
| Set Up Slide Show… | Set Up | `show.setUp` | | presented by a speaker / browsed in a window / kiosk; loop; without narration/animation; pen and laser colors; slides range; custom show; advance; monitor; resolution |
| Hide Slide | Set Up | `slide.hide` | | |
| Rehearse Timings | Set Up | `show.rehearseTimings` | | |
| Use Timings | Set Up | `show.useTimings` | | |
| Clear Timings (current / all) | Set Up | `show.clearTimings` | | not in tabs |
| Play Narrations | Set Up | `show.playNarrations` | | |
| Show Media Controls | Set Up | `show.mediaControls` | | |
| Timed Format (Pecha Kucha / Ignite / custom) | Set Up | `show.timedFormat` | | |
| Monitor | Monitors | `show.monitor` | | choose display |
| Use Presenter View | Monitors | `show.usePresenterView` | | |
| Presenter View Layout… | Monitors | `show.presenterLayout` | | |
| Always Use Subtitles | Captions & Subtitles | `show.subtitles` | | |
| Subtitle Settings… | Captions & Subtitles | `show.subtitleSettings` | | spoken language, subtitle language, position, size |
| Rehearsal Coach | Rehearse | `show.rehearsalCoach` | | Presenter Coach equivalent |
| Phone Remote… | Remote | `show.remote` | | |
| Audience Q&A | Audience | `show.qa` | | |
| Show Results / Reset Results | Audience | `show.showResults`, `show.resetResults` | | polls, quizzes |
| Follow Script | Rehearse | `show.followScript` | | stretch, M8 |

### In-show commands (also bindable)

| Command | id | Default keys |
|---|---|---|
| Next / Previous | `show.next`, `show.previous` | N, Space, Right, Down, Enter, Page Down / P, Backspace, Left, Up, Page Up |
| Go to Slide (number then Enter) | `show.goTo` | number, Enter |
| First / Last Slide | `show.first`, `show.last` | Home / End |
| See All Slides (grid) | `show.grid` | G, Ctrl+S |
| Zoom In / Out on slide | `show.zoomIn`, `show.zoomOut` | + / − |
| Black / White Screen | `show.black`, `show.white` | B or . / W or , |
| Pen / Highlighter / Laser Pointer / Arrow | `show.pen`, `show.highlighter`, `show.laser`, `show.arrow` | Ctrl+P, Ctrl+I, Ctrl+L, Ctrl+A |
| Erase All Ink | `show.eraseInk` | E |
| Hide Pointer / Show Pointer | `show.hidePointer` | Ctrl+H |
| Pause / Resume Automatic Show | `show.pause` | S |
| Toggle Subtitles | `show.toggleSubtitles` | J |
| Enter Stack / Leave Stack (vertical slides) | `show.navigate2d` | Down / Up |
| Return to Zoom | `zoom.returnToZoom` | |
| End Show | `show.end` | Esc |
| Keep Ink Annotations on Exit | `show.keepInk` | asked at exit |

## Record

| Command | Group | id | Notes |
|---|---|---|---|
| Record from Beginning | Record | `record.fromBeginning` | |
| Record from Current Slide | Record | `record.fromCurrent` | |
| Camera On / Off, Microphone, Camera Preview | Record | `record.camera`, `record.microphone`, `record.cameraPreview` | |
| Teleprompter | Record | `record.teleprompter` | notes scroll near the camera |
| Background Blur (camera) | Record | `record.backgroundBlur` | on-device segmentation where available |
| Re-record Slide | Record | `record.rerecordSlide` | |
| Clear Recording on Current Slide / All Slides | Record | `record.clearSlide`, `record.clearAll` | |
| Recording Timeline | Edit | `record.timeline` | Producer-style |
| Screenshot | Content | `insert.screenshot` | |
| Screen Recording | Content | `insert.screenRecording` | |
| Audio | Content | `insert.recordAudio` | |
| Save as Show (.ppsx) | Save | `file.export.ppsx` | |
| Export to Video | Save | `file.export.video` | |
| Rehearsal Coach | Practice | `show.rehearsalCoach` | |

## Review

| Command | Group | id | Shortcut | Notes |
|---|---|---|---|---|
| Spelling | Proofing | `review.spelling` | F7 | |
| Spelling Dictionaries… | Proofing | `review.dictionaries` | | wizard: download 90+ languages |
| Thesaurus | Proofing | `review.thesaurus` | Shift+F7 | |
| Word Count / Presentation Statistics | Proofing | `review.statistics` | | words, slides, notes, estimated speaking time |
| Check Accessibility | Accessibility | `review.checkAccessibility` | | |
| Alt Text | Accessibility | `review.altText` | | |
| Reading Order Pane | Accessibility | `review.readingOrder` | | |
| Design Checker | Accessibility | `review.designCheck` | | |
| Look Up | Insights | `review.lookUp` | | opens a web search for the selection |
| Translate | Language | `review.translate` | | AI provider or browser |
| Language… | Language | `review.language` | | proofing language for selection |
| New Comment | Comments | `review.newComment` | Ctrl+Alt+M | |
| Delete Comment / All on Slide / All in Presentation | Comments | `review.deleteComment`, `review.deleteSlideComments`, `review.deleteAllComments` | | |
| Previous / Next Comment | Comments | `review.prevComment`, `review.nextComment` | | |
| Show Comments (pane) | Comments | `review.commentsPane` | | |
| Compare… | Compare | `review.compare` | | merge another version, shown as revisions |
| Accept / Reject (change, slide, all) | Compare | `review.accept`, `review.reject` | | |
| Previous / Next Change | Compare | `review.prevChange`, `review.nextChange` | | |
| Reviewing Pane | Compare | `review.reviewPane` | | |
| End Review | Compare | `review.endReview` | | |
| Hide Ink | Ink | `ink.toggleVisibility` | | |
| Protect Presentation… | Protect | `review.protect` | | editing restrictions; Mark as Final |
| Slide Status / Assign Slide | Workflow | `slide.status`, `slide.assign` | | Pitch-style |

## View

| Command | Group | id | Shortcut | Notes |
|---|---|---|---|---|
| Normal | Presentation Views | `view.normal` | | |
| Outline View | Presentation Views | `view.outline` | | |
| Slide Sorter | Presentation Views | `view.slideSorter` | | |
| Notes Page | Presentation Views | `view.notesPage` | | |
| Reading View | Presentation Views | `view.reading` | | |
| Story View | Presentation Views | `view.story` | | |
| Canvas View | Presentation Views | `view.canvas` | | M7 |
| Slide Master | Master Views | `view.slideMaster` | | |
| Handout Master | Master Views | `view.handoutMaster` | | |
| Notes Master | Master Views | `view.notesMaster` | | |
| Ruler | Show | `view.ruler` | | |
| Gridlines | Show | `view.gridlines` | Shift+F9 | |
| Guides | Show | `view.guides` | Alt+F9 | add, remove, color |
| Smart Guides | Show | `view.smartGuides` | | |
| Grid and Guides… | Show | `view.gridSettings` | | snap to grid, snap to objects, spacing |
| Notes | Show | `view.notes` | | notes pane |
| Thumbnails Pane | Show | `view.thumbnails` | | |
| Navigator | Show | `view.navigator` | | |
| Zoom… | Zoom | `view.zoom` | | |
| Fit to Window | Zoom | `view.fit` | | |
| Zoom In / Out | Zoom | `view.zoomIn`, `view.zoomOut` | Ctrl+mouse wheel, pinch | Ctrl+= is Subscript |
| Color / Grayscale / Black and White | Color/Grayscale | `view.color`, `view.grayscale`, `view.blackWhite` | | |
| New Window | Window | `view.newWindow` | | desktop |
| Arrange All / Cascade | Window | `view.arrangeAll`, `view.cascade` | | desktop |
| Move Split | Window | `view.moveSplit` | | |
| Switch Windows / Next Presentation Tab | Window | `view.switchWindow` | Ctrl+Tab | |
| Macros | Macros | `dev.macros` | Alt+F8 | |
| Show Quick Access Bar / Hide | Appearance | `view.qatShow`, `view.qatHide` | | |
| Quick Access Bar Position | Appearance | `view.qatPosition` | | top, below ribbon, bottom |
| Quick Access Bar Labels | Appearance | `view.qatLabels` | | |
| Add All Quick Access Choices (A–Z) | Appearance | `view.qatAddAll` | | not in tabs |
| Collapse the Ribbon | Appearance | `view.collapseRibbon` | Ctrl+F1 | |
| Touch/Mouse Mode | Appearance | `view.touchMode` | | |
| Dark Mode / Light Mode / System | Appearance | `view.theme` | | |
| Full Screen (editing) | Appearance | `view.fullScreen` | | |
| Search Commands | — | `view.cmdSearch` | Alt+Q | |

### Slide Master tab (contextual)

| Command | Group | id |
|---|---|---|
| Insert Slide Master | Edit Master | `master.insertMaster` |
| Insert Layout | Edit Master | `master.insertLayout` |
| Delete | Edit Master | `master.delete` |
| Rename | Edit Master | `master.rename` |
| Preserve | Edit Master | `master.preserve` |
| Master Layout… | Master Layout | `master.masterLayout` |
| Insert Placeholder (Content, Content vertical, Text, Text vertical, Picture, Chart, Table, Diagram, Media, Online Image) | Master Layout | `master.insertPlaceholder(kind)` |
| Title / Footers (show on layout) | Master Layout | `master.showTitle`, `master.showFooters` |
| Themes, Colors, Fonts, Effects | Edit Theme | `design.theme`, `design.colors`, `design.fonts`, `design.effects` |
| Background Styles / Hide Background Graphics | Background | `design.backgroundStyles`, `master.hideBackgroundGraphics` |
| Slide Size | Size | `design.slideSize` |
| Close Master View | Close | `master.close` |

### Handout Master and Notes Master tabs (contextual)

| Command | id |
|---|---|
| Handout Orientation / Slide Orientation / Slides per Page | `master.handoutOrientation`, `master.slideOrientation`, `master.slidesPerPage` |
| Placeholders: Header, Date, Footer, Page Number, Slide Image, Body | `master.placeholderToggle(kind)` |
| Notes Page Orientation | `master.notesOrientation` |
| Close Master View | `master.close` |

## Help

| Command | id |
|---|---|
| Help | `help.help` (F1) |
| Keyboard Shortcuts | `help.shortcuts` |
| What's New | `help.whatsNew` |
| Report a Problem / Feedback | `help.feedback` (opens a GitHub issue form) |
| Check for Updates | `help.updates` (desktop) |
| Third-Party Notices | `help.notices` |
| About QuillStage | `help.about` |

## Developer (shown on request, as in the sister apps)

| Command | id | Notes |
|---|---|---|
| Macros… | `dev.macros` | run, edit, delete |
| Record Macro / Stop Recording | `dev.recordMacro`, `dev.stopRecording` | records commands, selections and typing |
| Macro Editor | `dev.macroEditor` | JavaScript with a presentation API |
| Macro Security… | `dev.macroSecurity` | |
| View VBA Project (read-only) | `dev.viewVba` | shown for .pptm; never run |
| Controls: Label, Text Box, Button, Check Box, Option Button, List Box, Combo Box, Toggle Button, Spin Button, Scroll Bar, Image | `dev.control(kind)` | work in shows with macros |
| Properties | `dev.controlProperties` | |
| Add-ins… | `dev.addins` | M8 |
| Document Tags… | `dev.tags` | view and edit .pptx custom tags |
| Export / Import Presentation JSON | `dev.exportJson`, `dev.importJson` | |

## Contextual tabs

### Shape Format

| Command | Group | id |
|---|---|---|
| Insert Shapes gallery | Insert Shapes | `insert.shape(type)` |
| Edit Shape: Change Shape | Insert Shapes | `shape.change(type)` |
| Edit Points / Close Path / Open Path | Insert Shapes | `shape.editPoints`, `shape.closePath`, `shape.openPath` |
| Reroute Connectors | Insert Shapes | `shape.rerouteConnectors` |
| Connector (straight, elbow, curved) | Insert Shapes | `shape.connector(kind)` |
| Text Box | Insert Shapes | `insert.textBox` |
| Merge Shapes: Union, Combine, Fragment, Intersect, Subtract | Insert Shapes | `shape.union`, `shape.combine`, `shape.fragment`, `shape.intersect`, `shape.subtract` |
| Shape Styles gallery / Shape Fill / Shape Outline / Shape Effects (shadow, reflection, glow, soft edges, bevel, 3-D rotation) | Shape Styles | `shape.quickStyle`, `shape.fill`, `shape.outline`, `shape.effects(kind)` |
| Format Shape Pane | Shape Styles | `shape.formatPane` |
| Set as Default Shape / Default Text Box | Shape Styles | `shape.setDefault`, `shape.setDefaultTextBox` |
| Decorative Text Styles / Text Fill / Text Outline / Text Effects / Transform | Decorative Text Styles | `shape.textStyle`, `shape.textFill`, `shape.textOutline`, `shape.textEffects`, `shape.textTransform` |
| Text Autofit: Do Not Autofit / Shrink Text on Overflow / Resize Shape to Fit Text | Text | `shape.autofit(mode)` |
| Text Margins / Wrap Text in Shape | Text | `shape.textMargins`, `shape.wrapText` |
| Alt Text / Mark as Decorative | Accessibility | `review.altText`, `shape.decorative` |
| Arrange | Arrange | `arrange.*` |
| Height / Width / Size and Position… / Lock Aspect Ratio | Size | `shape.height`, `shape.width`, `shape.sizeDialog`, `shape.lockAspect` |
| Save as Picture… | — | `shape.saveAsPicture` (not in tabs) |
| Create / Insert / Detach Linked Component | Components | `slide.createComponent`, `slide.insertComponent`, `slide.detachComponent` |

### Picture Format

| Command | Group | id |
|---|---|---|
| Remove Background (mark keep / mark remove) | Adjust | `picture.removeBackground` |
| Remove Color | Adjust | `picture.removeColor` |
| Corrections (sharpen/soften, brightness/contrast, presets) | Adjust | `picture.corrections` |
| Color (saturation, tone, recolor, Set Transparent Color) | Adjust | `picture.color`, `picture.setTransparentColor` |
| Artistic Effects | Adjust | `picture.artisticEffects` |
| Transparency | Adjust | `picture.transparency` |
| Compress Pictures… | Adjust | `picture.compress` |
| Change Picture (from file / online / clipboard) | Adjust | `picture.change` |
| Reset Picture / Reset Picture & Size | Adjust | `picture.reset`, `picture.resetSize` |
| Edit Picture (pixels) | Adjust | `picture.editPixels` |
| Picture Styles gallery / Border / Effects / Layout (to Diagram) | Picture Styles | `picture.style`, `picture.border`, `picture.effects`, `picture.layout` |
| Alt Text | Accessibility | `review.altText` |
| Arrange | Arrange | `arrange.*` |
| Crop / Crop to Shape / Aspect Ratio / Fill / Fit | Size | `picture.crop`, `picture.cropToShape`, `picture.cropAspect`, `picture.cropFill`, `picture.cropFit` |
| Height / Width / Size dialog | Size | `shape.height`, `shape.width`, `shape.sizeDialog` |
| Save as Picture… | — | `shape.saveAsPicture` |

### Graphics Format (icons and SVG)

| Command | id |
|---|---|
| Change Graphic | `graphic.change` |
| Convert to Shape | `graphic.toShape` |
| Graphics Fill / Outline / Effects | `shape.fill`, `shape.outline`, `shape.effects` |

### Table Design

| Command | Group | id |
|---|---|---|
| Header Row / Total Row / Banded Rows / First Column / Last Column / Banded Columns | Table Style Options | `table.headerRow`, `table.totalRow`, `table.bandedRows`, `table.firstColumn`, `table.lastColumn`, `table.bandedColumns` |
| Table Styles gallery / Clear Table | Table Styles | `table.style`, `table.clearStyle` |
| Shading / Borders / Effects | Table Styles | `table.shading`, `table.borders`, `table.effects` |
| Decorative Text Styles | Decorative Text Styles | `shape.textStyle` |
| Pen Style / Pen Weight / Pen Color / Draw Table / Eraser | Draw Borders | `table.penStyle`, `table.penWeight`, `table.penColor`, `insert.drawTable`, `table.eraser` |

### Table Layout

| Command | Group | id |
|---|---|---|
| Select Cell / Column / Row / Table | Table | `table.selectCell`, `table.selectColumn`, `table.selectRow`, `table.selectTable` |
| View Gridlines | Table | `table.viewGridlines` |
| Delete Columns / Rows / Table | Rows & Columns | `table.deleteColumns`, `table.deleteRows`, `table.delete` |
| Insert Above / Below / Left / Right | Rows & Columns | `table.insertAbove`, `table.insertBelow`, `table.insertLeft`, `table.insertRight` |
| Merge Cells / Split Cells | Merge | `table.merge`, `table.split` |
| Height / Width / Distribute Rows / Distribute Columns | Cell Size | `table.rowHeight`, `table.colWidth`, `table.distributeRows`, `table.distributeColumns` |
| Align (9 positions) / Text Direction / Cell Margins | Alignment | `table.align(pos)`, `table.textDirection`, `table.cellMargins` |
| Table Size / Lock Aspect Ratio | Table Size | `table.size`, `shape.lockAspect` |
| Sort Rows (by column) | Data | `table.sort` (not in tabs) |
| Convert Table to Text / Text to Table | Data | `table.toText`, `table.fromText` (not in tabs) |
| Arrange | Arrange | `arrange.*` |

### Chart Design

| Command | Group | id |
|---|---|---|
| Add Chart Element (axes, axis titles, chart title, data labels, data table, error bars, gridlines, legend, lines, trendline, up/down bars) | Chart Layouts | `chart.addElement(kind)` |
| Quick Layout | Chart Layouts | `chart.quickLayout` |
| Change Colors | Chart Styles | `chart.colors` |
| Chart Styles gallery | Chart Styles | `chart.style` |
| Switch Row/Column | Data | `chart.switchRowCol` |
| Select Data… | Data | `chart.selectData` |
| Edit Data (in grid / in QuillGrid) | Data | `chart.editData`, `chart.editInGrid` |
| Refresh Data | Data | `chart.refreshData` |
| Change Chart Type | Type | `chart.changeType` |
| Save as Template… | Type | `chart.saveTemplate` |
| Animate Chart by Series / by Category | Animation | `chart.animateBy(kind)` (sequence option of `anim.effectOptions`) |

### Chart Format

| Command | Group | id |
|---|---|---|
| Chart Elements selector / Format Selection / Reset to Match Style | Current Selection | `chart.selectElement`, `chart.formatSelection`, `chart.resetStyle` |
| Insert Shapes / Shape Styles / Decorative Text Styles / Arrange / Size | — | as Shape Format |

### Diagram Design (SmartArt equivalent)

| Command | Group | id |
|---|---|---|
| Add Shape (after, before, above, below, assistant) | Create Graphic | `diagram.addShape(pos)` |
| Add Bullet | Create Graphic | `diagram.addBullet` |
| Text Pane | Create Graphic | `diagram.textPane` |
| Promote / Demote / Move Up / Move Down | Create Graphic | `diagram.promote`, `diagram.demote`, `diagram.moveUp`, `diagram.moveDown` |
| Right to Left | Create Graphic | `diagram.rtl` |
| Org Chart Layout (standard, both, left hanging, right hanging) | Create Graphic | `diagram.orgLayout` |
| Layouts gallery | Layouts | `diagram.layout(name)` |
| Change Colors / Styles gallery | Styles | `diagram.colors`, `diagram.style` |
| Reset Graphic | Reset | `diagram.reset` |
| Convert to Text / Convert to Shapes | Reset | `diagram.toText`, `diagram.toShapes` |
| Auto Layout (mind map, tree, flowchart, timeline) | Layouts | `diagram.autoLayout` |
| Convert Bullets to Diagram | — | `diagram.fromOutline` |

### Diagram Format

| Command | id |
|---|---|
| Edit in 2-D | `diagram.edit2d` |
| Change Shape / Larger / Smaller | `diagram.changeShape`, `diagram.larger`, `diagram.smaller` |
| Shape Styles, Decorative Text Styles, Arrange, Size | as Shape Format |

### Video Format

| Command | Group | id |
|---|---|---|
| Play | Preview | `media.play` |
| Corrections / Color / Poster Frame (current frame, from file) / Reset Design | Adjust | `media.corrections`, `media.color`, `media.posterFrame`, `media.resetDesign` |
| Video Styles / Video Shape / Video Border / Video Effects | Video Styles | `media.style`, `media.shape`, `media.border`, `media.effects` |
| Alt Text | Accessibility | `review.altText` |
| Arrange / Crop / Size | Arrange, Size | `arrange.*`, `media.crop`, `shape.sizeDialog` |

### Video Playback

| Command | Group | id |
|---|---|---|
| Play | Preview | `media.play` |
| Add Bookmark / Remove Bookmark | Bookmarks | `media.addBookmark`, `media.removeBookmark` |
| Trim Video | Editing | `media.trim` |
| Split / Speed | Editing | `media.split`, `media.speed` |
| Fade In / Fade Out durations | Editing | `media.fadeIn`, `media.fadeOut` |
| Volume (low, medium, high, mute) / Volume Envelope | Video Options | `media.volume`, `media.volumeEnvelope` |
| Start: In Click Sequence / Automatically / When Clicked On | Video Options | `media.start(mode)` |
| Play Full Screen | Video Options | `media.fullScreen` |
| Hide While Not Playing | Video Options | `media.hideWhileNotPlaying` |
| Loop until Stopped | Video Options | `media.loop` |
| Rewind after Playing | Video Options | `media.rewind` |
| Insert Captions / Remove Captions | Caption Options | `media.insertCaptions`, `media.removeCaptions` |
| Generate Captions | Caption Options | `media.generateCaptions` (speech recognition) |
| Export Trimmed Media… | Editing | `media.exportTrimmed` |

### Audio Format and Audio Playback

| Command | id |
|---|---|
| Play, Bookmarks, Trim Audio, Fade In/Out, Volume, Start, Loop, Rewind | as Video Playback (`media.*`) |
| Play Across Slides | `media.playAcrossSlides` |
| Hide During Show | `media.hideDuringShow` |
| No Style / Play in Background | `media.noStyle`, `media.playInBackground` |
| Audio icon styles, Arrange, Size | as Picture Format |

### Live Camera Format (Cameo equivalent)

| Command | id |
|---|---|
| Preview / Turn Camera On or Off | `camera.preview`, `camera.toggle` |
| Choose Camera | `camera.select` |
| Camera Styles / Shape / Border / Effects | `camera.style`, `camera.shape`, `camera.border`, `camera.effects` |
| Background Blur | `record.backgroundBlur` |
| Arrange, Size | `arrange.*`, `shape.sizeDialog` |

### Zoom (contextual, for zoom links)

| Command | id |
|---|---|
| Change Image | `zoom.changeImage` |
| Reset | `zoom.reset` |
| Zoom Styles / Background | `zoom.style`, `zoom.background` |
| Return to Zoom | `zoom.returnToZoom` |
| Zoom Transition / Duration | `zoom.transition`, `zoom.duration` |
| Edit Summary (choose sections) | `zoom.editSummary` |

### 3D Model (M8)

| Command | id |
|---|---|
| 3D Model Views gallery | `model3d.view` |
| Pan & Zoom | `model3d.panZoom` |
| Reset Model | `model3d.reset` |
| Play Embedded Animation / Scenes | `model3d.play`, `model3d.scene` |

### Equation

| Command | id |
|---|---|
| Professional / Linear / LaTeX | `equation.professional`, `equation.linear`, `equation.latex` |
| Normal Text | `equation.normalText` |
| Symbols and Structures galleries | `equation.symbol`, `equation.structure` |
| Ink Equation | `ink.toMath` |

### Ink (contextual, when ink is selected)

| Command | id |
|---|---|
| Pens gallery / Color / Thickness | `draw.pen`, `ink.color`, `ink.thickness` |
| Convert to Shape / Text / Math | `ink.toShape`, `ink.toText`, `ink.toMath` |
| Replay | `ink.replay` |
| Arrange / Size | `arrange.*`, `shape.sizeDialog` |

### Code Block (new)

| Command | id |
|---|---|
| Language | `code.language` |
| Line Numbers | `code.lineNumbers` |
| Highlight Theme | `code.theme` |
| Step Through Lines… | `code.stepLines` |
| Run Code (sandbox) | `code.run` |
| Copy as Plain Text | `code.copyText` |

### Canvas (new, M7)

| Command | id |
|---|---|
| Add Frame / Remove Frame | `canvas.addFrame`, `canvas.removeFrame` |
| Edit Path | `canvas.editPath` |
| Overview Step | `canvas.overview` |
| Brainstorm Board | `canvas.brainstorm` |
| Board to Slides | `canvas.toSlides` |
| Convert Deck to Canvas / Canvas to Deck | `canvas.fromDeck`, `canvas.toDeck` |

### Interactive (new, M6)

| Command | id |
|---|---|
| Edit Question and Choices | `poll.edit` |
| Result Chart Type | `poll.chartType` |
| Allow Multiple Answers / Anonymous | `poll.multiple`, `poll.anonymous` |
| Quiz: Correct Answer / Timer / Leaderboard | `quiz.correct`, `quiz.timer`, `quiz.leaderboard` |
| Export Results to QuillGrid (.csv/.xlsx) | `poll.exportResults` |

## Slides, sections and outline (thumbnails and Outline view)

| Command | id | Shortcut | Notes |
|---|---|---|---|
| Move Slide Up / Down / to Start / to End | `slide.moveUp`, `slide.moveDown`, `slide.moveStart`, `slide.moveEnd` | Ctrl+Up, Ctrl+Down, Ctrl+Shift+Up, Ctrl+Shift+Down | |
| Go to Slide… | `slide.goTo` | | |
| Copy Slide to Another Presentation | `slide.copyTo` | | |
| Lock Layout | `slide.lockLayout` | | |
| Split into Slides (by paragraph) | `slide.splitByParagraph` | | |
| Merge Slides | `slide.merge` | | |
| Add Child Slide (vertical stack) | `slide.stackBelow` | | |
| Slide Status / Assign | `slide.status`, `slide.assign` | | |
| Promote / Demote | `outline.promote`, `outline.demote` | Alt+Shift+Left, Alt+Shift+Right | |
| Move Up / Move Down | `outline.moveUp`, `outline.moveDown` | Alt+Shift+Up, Alt+Shift+Down | |
| Collapse / Expand / Collapse All / Expand All | `outline.collapse`, `outline.expand`, `outline.collapseAll`, `outline.expandAll` | Alt+Shift+−, Alt+Shift+= | |
| Show Text Formatting | `outline.showFormatting` | | |

## Tools (new tab, as in the sister apps)

| Command | id | Notes |
|---|---|---|
| Fit Text to Slide | `fmt.fitTextToSlide` | Takahashi |
| Find Off-Slide Content | `tools.offSlide` | |
| Find Unused Masters and Layouts | `tools.unusedLayouts` | |
| Remove Unused Masters and Layouts | `tools.removeUnusedLayouts` | |
| Reset All Slides to Layout | `tools.resetAll` | |
| Speaker Notes to Script (QuillPen) | `tools.notesToDoc` | |
| Estimate Speaking Time | `tools.speakingTime` | from notes word count |
| Batch Replace Images | `tools.replaceImages` | |
| Extract All Media… | `tools.extractMedia` | also used by recovery |
| Slide Numbering Options… | `tools.numbering` | start number, skip title slide |
| Compare Fonts Used | `tools.fontsUsed` | substitutions report |
| OCR Clean-up for Text | `tools.ocrCleanup` | shared with the suite |

## AI tab

| Command | Group | id | Notes |
|---|---|---|---|
| Set Up AI… | Setup | `ai.setup` | wizard: on-device, local server, free tier, paid |
| Active AI | Setup | `ai.active` | |
| Providers & Models… | Setup | `ai.providers` | |
| Several AIs ▸ Single AI / Master Combines Answers / Quaker Consensus | Setup | `ai.ensembleMode` | menu; see docs/ENSEMBLE.md |
| Choose AIs… | Setup | `ai.ensembleMembers` | 2–6 providers, master/clerk, rounds |
| Show AI Discussion | Setup | `ai.showDeliberation` | toggle |
| AI & Privacy | Setup | `ai.privacy` | |
| Delete Downloaded Models | Setup | `ai.deleteModels` | |
| Stop AI | Setup | `ai.stop` | |
| Draft Presentation… | Create | `ai.draftDeck` | outline first, then slides |
| Document to Deck… | Create | `ai.docToDeck` | |
| Outline to Slides | Create | `ai.outlineToSlides` | |
| Add Slide About… | Create | `ai.addSlide` | |
| Suggest Images (openly licensed) | Create | `ai.suggestImages` | search terms, never generated by default |
| Generate Image… | Create | `ai.generateImage` | only with a provider that offers it; labeled as generated |
| Design Ideas (AI-ranked) | Design | `ai.designIdeas` | |
| Rewrite / Shorten / Expand / Change Tone | Text | `ai.rewrite`, `ai.shorten`, `ai.expand`, `ai.tone` | |
| Bullets to Sentence Headline | Text | `ai.assertionHeadline` | assertion–evidence |
| Proofread | Text | `ai.proofread` | |
| Translate Slide / Deck | Text | `ai.translate` | keeps layout |
| Write Speaker Notes | Notes | `ai.writeNotes` | |
| Notes to Slide Text / Slide Text to Notes | Notes | `ai.notesToSlide`, `ai.slideToNotes` | |
| Summarize Deck | Notes | `ai.summarize` | |
| Generate Alt Text | Accessibility | `ai.altText` | |
| Suggest Chart from Data | Data | `ai.suggestChart` | |
| Generate Quiz from Deck | Audience | `ai.quizFromDeck` | |
| Anticipate Questions | Audience | `ai.anticipateQuestions` | |
| Rehearsal Feedback | Practice | `ai.rehearsalFeedback` | after a Rehearsal Coach run |
| Write Macro | Developer | `ai.writeMacro` | shown for review, never run automatically |
| AI Assistant | Assistant | `ai.assistant` | chat pane with deck context |
| Prompt Library / My Prompts | Assistant | `ai.prompts`, `ai.myPrompts` | |

## Commands adopted from other software

Collected from [INNOVATIONS.md](INNOVATIONS.md) for easy review; each also appears above in its tab.

| Command | id | Origin of the idea |
|---|---|---|
| Content Starters | `file.newFromStarter` | Freelance SmartMasters |
| Layout by Outline Level | `design.autoLayoutRules` | Persuasion AutoTemplates |
| Design Checker | `review.designCheck` | Harvard Graphics advisor; Tufte |
| Animation Timeline | `anim.timeline` | Astound |
| Edit Picture | `picture.editPixels` | Corel Presentations |
| Package Presentation | `file.package` | Corel Show on the Go, Pack and Go |
| Recording Timeline, Web Lecture | `record.timeline`, `file.export.webLecture` | Producer |
| Smooth Morph | `trans.smoothMorph` | Keynote, PowerPoint |
| Build Order | `anim.buildOrder` | Keynote |
| Presenter View Layout, Rehearse | `show.presenterLayout`, `show.rehearse` | Keynote |
| Remove Color | `picture.removeColor` | Keynote |
| Phone Remote | `show.remote` | Keynote, Impress |
| Broadcast Show | `show.broadcast` | Keynote Live |
| Reveal by Step | `anim.revealSteps` | Beamer |
| Export PDF with build steps | `file.export.pdf` | Beamer |
| Export HTML Show | `file.export.htmlShow` | S5, reveal.js |
| Audience Q&A | `show.qa` | Google Slides, Slido |
| Linked Chart | `insert.linkedChart` | Google Slides |
| Navigator | `view.navigator` | Impress |
| Compress Presentation | `file.compress` | Impress Minimizer |
| Canvas View, frames, path | `view.canvas`, `canvas.*` | Prezi, Focusky, Sozi |
| Zoom links | `insert.summaryZoom`, `insert.sectionZoom`, `insert.slideZoom` | PowerPoint, Prezi |
| Markdown slides | `file.import.markdown`, `file.export.markdown` | Marp, reveal.js, Slidev |
| Vertical stacks | `slide.stackBelow`, `show.navigate2d` | reveal.js |
| Code Block, Step Through Lines, Run Code | `insert.codeBlock`, `code.stepLines`, `code.run` | reveal.js, Slidev |
| Video Background | `design.backgroundVideo` | Emaze |
| Brand Kit, Resize Presentation | `design.brandKit`, `design.resizeTo` | Canva |
| Cast to Display | `show.cast` | Zoho Show |
| Story export and view | `file.export.story`, `view.story` | Sway |
| Web Embed, Linked Components | `insert.webEmbed`, `slide.createComponent` | Ludus |
| Poll, Word Cloud, Quiz, Rating | `insert.poll`, `insert.wordCloud`, `insert.quiz`, `insert.rating` | Mentimeter, Slido |
| Insert PDF Pages, Restyle Deck | `insert.pdfPages`, `design.restyle` | WPS, Gamma |
| Design Ideas | `design.designIdeas` | PowerPoint |
| Rehearsal Coach | `show.rehearsalCoach` | PowerPoint Presenter Coach |
| Live Subtitles | `show.subtitles` | PowerPoint |
| Live Camera | `insert.liveCamera` | PowerPoint Cameo, Keynote |
| Record tab, Teleprompter | `record.*` | PowerPoint Recording Studio |
| Slide Status, Assign | `slide.status`, `slide.assign` | Pitch |
| Draft Presentation, Document to Deck | `ai.draftDeck`, `ai.docToDeck` | Tome, Gamma |
| Media trim, split, speed, envelope | `media.trim`, `media.split`, `media.speed`, `media.volumeEnvelope` | KineMaster |
| Brainstorm Board, Sticky Note, Board to Slides | `canvas.brainstorm`, `insert.stickyNote`, `canvas.toSlides` | Freeform, Miro |
| Keyboard diagramming, Auto Layout | `diagram.autoLayout` | Whimsical |
| Text to Deck | `file.import.textDeck` | Plain-text decks |
| Fit Text to Slide, Split into Slides | `fmt.fitTextToSlide`, `slide.splitByParagraph` | Takahashi, Lessig |
| Timed Format | `show.timedFormat` | Pecha Kucha, Ignite |
| Bullets to Sentence Headline | `ai.assertionHeadline` | Assertion–evidence |
| Sparkline, Handout Document | `insert.sparkline`, `file.export.handoutDoc` | Tufte |
| Follow Script | `show.followScript` | Lessig method |

## Suite-wide commands carried over

These exist in QuillPen and QuillGrid and keep their ids: `view.cmdSearch`, `view.qatAddAll`, `view.qatPosition`, `view.touchMode`, `file.customizeQat`, `file.customizeRibbon`, `file.customizeKeyboard`, `file.exportSettings`, `file.versions`, `file.snapshot`, `file.openRepair`, `file.recoveryReport`, `review.spelling`, `review.dictionaries`, `review.thesaurus`, `dev.recordMacro`, `dev.macroEditor`, and the AI setup commands (`ai.setup`, `ai.providers`, `ai.privacy`, `ai.active`, `ai.stop`, `ai.deleteModels`).
