#!/bin/zsh
# Rebuild the proposal .docx from proposal.md with IS5116 template styles, then add table borders.
cd "$(dirname "$0")"
OUT="${1:-Coreitera_Synthesia_Business_Proposal_v0.2.docx}"
pandoc proposal.md -o "$OUT" --reference-doc="../../NUS-Learning/IS5116/Team Quotient Business Proposal - FINAL.docx" --resource-path=. || exit 1
python3 - "$OUT" <<'PY'
import sys, zipfile, re, shutil
f=sys.argv[1]; tmp=f+".tmp"
zin=zipfile.ZipFile(f); zout=zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED)
b='<w:tblBorders>'+''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="B4B2A9"/>' for s in ["top","left","bottom","right","insideH","insideV"])+'</w:tblBorders>'
for it in zin.infolist():
    d=zin.read(it.filename)
    if it.filename=="word/document.xml":
        x=re.sub(r'<w:tblBorders>.*?</w:tblBorders>','',d.decode())
        x,n=re.subn(r'(<w:tblPr>)',r'\1'+b,x); print("tables bordered:",n); d=x.encode()
    zout.writestr(it,d)
zout.close(); shutil.move(tmp,f)
PY
