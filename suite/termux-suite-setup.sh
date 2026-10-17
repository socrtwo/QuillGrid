#!/data/data/com.termux/files/usr/bin/bash
# Quill suite: rename the GitHub repositories, set descriptions/topics, and create QuillStage.
# Run in Termux (works in any bash with git + GitHub CLI):
#   pkg install -y gh git
#   gh auth login            # GitHub.com, HTTPS, log in with a browser; grant "repo" scope
#   gh auth setup-git
#   bash termux-suite-setup.sh
# Safe to re-run: each step checks whether it was already done.
set -euo pipefail
OWNER=socrtwo
BRANCH=ccr-90803495-l7seer   # branch that carries the renamed code in each repository

SUITE="Part of the Quill suite: QuillPen (documents), QuillGrid (spreadsheets), QuillPost (email) and QuillStage (presentations)."
DESC_PEN="An original, open-source word processing project focused on modern writing, editing, and collaboration experiences. Built as a clean-room alternative with a familiar document workflow and innovative features. $SUITE"
DESC_GRID="An original, open-source spreadsheet project focused on modern calculation, analysis, and data-cleanup experiences. Built as a clean-room alternative with a familiar workbook workflow and innovative features. $SUITE"
DESC_POST="An original, open-source email project focused on modern mail, calendar, and contact experiences. Built as a clean-room alternative with a familiar mailbox workflow and innovative features. $SUITE"
DESC_STAGE="An original, open-source presentation project focused on modern storytelling, design, and delivery experiences. Built as a clean-room alternative with a familiar slide workflow and innovative features. $SUITE"

exists() { gh repo view "$OWNER/$1" --json name >/dev/null 2>&1; }

rename() { # old new
  if exists "$2"; then echo "✓ $2 already exists"; elif exists "$1"; then gh repo rename "$2" -R "$OWNER/$1" --yes && echo "✓ renamed $1 → $2"; else echo "✗ neither $1 nor $2 found"; fi
}

echo "== Renaming repositories"
rename quill-lattice QuillGrid
rename QuillForge    QuillPen
rename quillbox      QuillPost

echo "== Descriptions and topics"
gh repo edit "$OWNER/QuillPen"  --description "$DESC_PEN"  --add-topic word-processor,docx,odt,office-suite,quill-suite,android,ios,electron,pwa,capacitor,clean-room,open-source,offline-first,ai
gh repo edit "$OWNER/QuillGrid" --description "$DESC_GRID" --add-topic spreadsheet,xlsx,ods,csv,office-suite,quill-suite,android,ios,electron,pwa,capacitor,clean-room,open-source,offline-first,ai
gh repo edit "$OWNER/QuillPost" --description "$DESC_POST" --add-topic email-client,imap,pop3,smtp,calendar,contacts,office-suite,quill-suite,android,ios,kotlin,ktor,clean-room,open-source,ai

echo "== QuillStage"
if ! exists QuillStage; then
  gh repo create "$OWNER/QuillStage" --public --description "$DESC_STAGE"
fi
gh repo edit "$OWNER/QuillStage" --description "$DESC_STAGE" --add-topic presentation,slides,slideshow,pptx,odp,office-suite,quill-suite,android,ios,electron,pwa,capacitor,clean-room,open-source,ai
if [ -z "$(gh api "repos/$OWNER/QuillStage/commits?per_page=1" --jq '.[0].sha' 2>/dev/null || true)" ]; then
  WORK="$(mktemp -d)"
  gh repo clone "$OWNER/QuillGrid" "$WORK/grid" -- --depth 1 --branch "$BRANCH"
  cp -r "$WORK/grid/suite/QuillStage" "$WORK/QuillStage"
  rm -rf "$WORK/QuillStage/.git"
  cd "$WORK/QuillStage"
  git init -q -b main
  git add -A
  git -c user.name="$(gh api user --jq .login)" -c user.email="$(gh api user --jq '.id|tostring')+$(gh api user --jq .login)@users.noreply.github.com" \
    commit -q -m "QuillStage: project plan, innovations survey, command and Quick Access Bar plans"
  git remote add origin "https://github.com/$OWNER/QuillStage.git"
  git push -u origin main
  cd - >/dev/null
  rm -rf "$WORK"
  echo "✓ QuillStage pushed"
else
  echo "✓ QuillStage already has commits"
fi

echo "== Visibility (QuillGrid is private; QuillPen and QuillPost are public)"
echo "   To make QuillGrid public like the others, run:"
echo "   gh repo edit $OWNER/QuillGrid --visibility public --accept-visibility-change-consequences"

echo "== Local clones on this device (if any): point them at the new names"
for pair in "quill-lattice:QuillGrid" "QuillForge:QuillPen" "quillbox:QuillPost"; do
  old="${pair%%:*}"; new="${pair##*:}"
  for dir in "$HOME/$old" "$HOME/$new" "$HOME/storage/shared/$old"; do
    if [ -d "$dir/.git" ]; then git -C "$dir" remote set-url origin "https://github.com/$OWNER/$new.git" && echo "✓ $dir → $new"; fi
  done
done
echo "Done."
