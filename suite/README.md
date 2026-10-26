# Anne's Office suite: shared assets

Files shared by the four Anne's Office apps, kept here until the suite has a repository of its own.

| Path | What it is |
|---|---|
| `icons/` | The app icon family. `node icons/gen.mjs` writes each app's `icon.svg` (full tile), `icon-foreground.svg` (Android adaptive foreground, inside the 66 dp safe circle) and `icon-background.txt`. Each app copies its three files and regenerates its platform icons. |
| `scrollcue/` | The "more choices below" cues: thin scrollbars, edge shadows and a "More" chip, framework-free. Each app keeps a copy; fixes go here first. |
| `EMBED.md` | Protocol for editing Anne's Post messages in an embedded QuillPen. |
| `ENSEMBLE.md` | Several AIs on one request: master synthesis or Quaker consensus. |
| `QuillStage/` | The starting files for the Anne's Stage repository (plans only, no code). |
| `termux-suite-setup.sh` | Renames the GitHub repositories, sets descriptions and topics, creates and pushes QuillStage. Run in Termux after `gh auth login`. |
| `rename.py` | The script that renamed QuillForge, QuillLattice and Quillbox to Anne's Pen, Anne's Grid and QuillPost. Kept for reference; it names the old spellings on purpose. |
| `QuillStage.meta.json` | Name, description, topics and keywords for QuillStage. |
