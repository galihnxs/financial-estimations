set -e
B="/Users/galihmujiagung/Documents/GitHub/core-itera/financial-estimations/BSN Pemetaan Karyawan/build"
P="$B/../Proposal Teman Journey - Pembaruan Akhir - BSN.pptx"; Q="$B/_qa"
S=/private/tmp/claude-501/-Users-galihmujiagung-Documents-GitHub-core-itera/5305cb00-179e-4211-a0ef-d93579aebf93/scratchpad
cd "$B"; PW=/Users/galihmujiagung/Documents/GitHub/core-itera/customer-facing-platform-FE/node_modules/@playwright/test PPTXGEN=$S/pptx/node_modules/pptxgenjs node html_to_pptx.cjs update_deck.html "$P"
zip -dq "$P" '*/' || true
rm -f "$Q"/deck.* "$Q"/p*.png; cp "$P" "$Q/deck.pptx"
osascript -e "with timeout of 300 seconds
tell application \"Microsoft PowerPoint\"
open POSIX file \"$Q/deck.pptx\"
delay 5
save active presentation in POSIX file \"$Q/deck.pdf\" as save as PDF
close active presentation saving no
end tell
end timeout"
cd "$Q"; /opt/homebrew/Caskroom/miniforge/base/bin/python -W ignore -c "
import fitz
from PIL import Image
d=fitz.open('deck.pdf')
for i,p in enumerate(d): p.get_pixmap(dpi=${DPI:-60}).save(f'p{i+1:02d}.png')
ims=[Image.open(f'p{i:02d}.png') for i in range(1,d.page_count+1)]
w,h=ims[0].size
for k in range(3):
  g=Image.new('RGB',(w*2,h*4),'white')
  for j,im in enumerate(ims[k*8:(k+1)*8]): g.paste(im,((j%2)*w,(j//2)*h))
  g.save(f'grid{k+1}.png')
print(d.page_count)"
