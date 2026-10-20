#!/bin/bash
# Usage: ./check_slides.sh            -> list slides that overflow (headless Chrome)
#        ./check_slides.sh shots a1 b1 -> PNG screenshots of those slides in ../output/shots/
CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
PAGE="$(cd "$(dirname "$0")/../proposal" && pwd)/theta-partnership-proposal-present.html"
OUT="$(cd "$(dirname "$0")/.." && pwd)/output/shots"
if [ "$1" = "shots" ]; then
  shift; mkdir -p "$OUT"
  for id in "$@"; do
    "$CHROME" --headless=new --disable-gpu --hide-scrollbars --force-device-scale-factor=1 --window-size=1280,720 \
      --virtual-time-budget=3000 --screenshot="$OUT/$id.png" "file://$PAGE?only=$id" >/dev/null 2>&1
    echo "$OUT/$id.png"
  done
else
  "$CHROME" --headless=new --disable-gpu --window-size=1440,900 --virtual-time-budget=3000 --dump-dom "file://$PAGE?check" 2>/dev/null \
    | python3 -I -c "import sys,re,html; m=re.search(r'<pre id=\"fitcheck\">(.*?)</pre>', sys.stdin.read(), re.S); print(html.unescape(m.group(1)) if m else 'no result')"
fi
