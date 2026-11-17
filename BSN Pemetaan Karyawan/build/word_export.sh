#!/bin/zsh
# usage: word_export.sh file.docx -> updates the table of contents, saves the docx, writes file.pdf next to it
# Word is sandboxed, so it works on a short-named copy inside its container folder (long names fail to open).
D=~/Library/Containers/com.microsoft.Word/Data/tmp_render; rm -rf $D; mkdir -p $D
SRC="$1"; DIR=$(dirname "$SRC"); B=$(basename "$SRC" .docx)
cp "$SRC" "$D/export.docx"
osascript <<OSA
tell application "Microsoft Word"
  open file name (POSIX file "$D/export.docx" as string)
  repeat 30 times
    if (name of documents) contains "export.docx" then exit repeat
    delay 1
  end repeat
  set d to document "export.docx"
  try
    update (table of contents 1 of d)
  end try
  save d
  save as d file name "$D/export.pdf" file format format PDF
  close d saving no
end tell
OSA
cp "$D/export.docx" "$SRC" && cp "$D/export.pdf" "$DIR/$B.pdf" && rm -rf $D
