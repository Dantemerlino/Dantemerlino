#!/usr/bin/env bash
# Rebuild the published PDFs from the local Word sources (which are gitignored).
# Uses Microsoft Word itself, so the PDF matches what Word shows. The first run
# may show Word's "Grant File Access" dialog: grant this folder once.
# Usage: ./make-pdf.sh          convert only
#        ./make-pdf.sh --push   convert, then commit and push the PDFs
set -euo pipefail

cd "$(dirname "$0")"

DOCS=(Dante_Merlino_CV Dante_Merlino_Summary)

[[ -d "/Applications/Microsoft Word.app" ]] || { echo "Microsoft Word not found" >&2; exit 1; }

for name in "${DOCS[@]}"; do
  [[ -f "$name.docx" ]] || { echo "missing $name.docx" >&2; exit 1; }
  rm -f "$name.pdf"
  osascript - "$PWD/$name.docx" "$PWD/$name.pdf" <<'APPLESCRIPT'
on run {src, dst}
  with timeout of 120 seconds
    tell application "Microsoft Word"
      open (POSIX file src)
      set d to active document
      save as d file name dst file format format PDF
      close d saving no
    end tell
  end timeout
end run
APPLESCRIPT
  [[ -s "$name.pdf" ]] || { echo "conversion failed: $name.docx" >&2; exit 1; }
  echo "built $name.pdf"
done

if [[ "${1:-}" == "--push" ]]; then
  git add "${DOCS[@]/%/.pdf}"
  if git diff --cached --quiet; then
    echo "PDFs unchanged; nothing to push"
  else
    git commit -m "Update CV and summary PDFs"
    git push
  fi
fi
