"""Photo wall page ("Jejak kegiatan kami") for Deck A, page 51.

Renders the design layer (navy bg, tag, title, 8 white polaroid frames) to PNG and writes
wall_slots.json: exact photo and caption geometry for placing real photos + editable captions in Canva.
Run: /opt/homebrew/bin/python3 -I deck_wall.py
"""
import json
import math
import os
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "deck_visuals")
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

PHOTOS = [  # media id, caption (editable in Canva)
    ("MAHaAuYi4bQ", "Program bersama Bank Indonesia"),
    ("MAHG7bJiC8c", "Pelatihan guru, Kemendikdasmen"),
    ("MAHZ9bxWZzk", "Lentera Journey, komunitas"),
    ("MAHCNLioOxs", "Art-based wellbeing"),
    ("MAHaAnIEMFU", "Workshop agen perubahan"),
    ("MAHaYzH9Ceg", "Fit &amp; Proper Test daring"),
    ("MAHG7RpTjCo", "Workshop karyawan"),
    ("MAHaIC-5Eig", "Ruang cerita komunitas"),
]
FW, FH, PAD, BOT = 400, 330, 14, 56
PW, PH = FW - 2 * PAD, FH - PAD - BOT
LEFTS = [108, 548, 988, 1428]
TOPS = [330, 690]
ROTS = [[-3, 2, -2, 3], [2, -3, 3, -2]]


def rot(cx, cy, ox, oy, deg):
    a = math.radians(deg)
    return cx + ox * math.cos(a) - oy * math.sin(a), cy + ox * math.sin(a) + oy * math.cos(a)


frames, slots = "", []
for i, (media, cap) in enumerate(PHOTOS):
    r, c = divmod(i, 4)
    fx, fy, deg = LEFTS[c], TOPS[r], ROTS[r][c]
    cx, cy = fx + FW / 2, fy + FH / 2
    pcx, pcy = rot(cx, cy, 0, (PAD + PH / 2) - FH / 2, deg)
    tcx, tcy = rot(cx, cy, 0, (FH - BOT / 2) - FH / 2, deg)
    slots.append({"media": media, "caption": cap.replace("&amp;", "&"),
                  "photo": [round(pcx - PW / 2, 2), round(pcy - PH / 2, 2), PW, PH, deg],
                  "caption_box": [round(tcx - PW / 2, 2), round(tcy - 17, 2), PW, deg]})
    tape = "" if i % 3 else f'<div style="position:absolute;left:{FW/2-60}px;top:-30px;width:120px;height:40px;background:rgba(246,227,182,.9);transform:rotate(-5deg)"></div>'
    frames += (f'<div style="position:absolute;left:{fx}px;top:{fy}px;width:{FW}px;height:{FH}px;background:#fff;transform:rotate({deg}deg);'
               f'box-shadow:0 26px 50px rgba(4,24,40,.38)"><div style="position:absolute;left:{PAD}px;top:{PAD}px;width:{PW}px;height:{PH}px;'
               f'background:repeating-linear-gradient(45deg,#9fb3c6 0 18px,#b8c7d6 18px 36px)"></div>{tape}</div>')

html = f"""<!doctype html><html><head><meta charset="utf-8"><style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&display=swap');
*{{box-sizing:border-box;margin:0;padding:0}} html,body{{width:1920px;height:1080px;overflow:hidden}}
body{{font-family:'Poppins',sans-serif;color:#fff;position:relative;background:linear-gradient(90deg,#104a70 0%,#4f5c74 52%,#628cab 100%)}}
</style></head><body>
<div style="position:absolute;width:760px;height:620px;left:-260px;top:560px;background:#e8d4ab;opacity:.10;border-radius:58% 42% 61% 39%/45% 55% 45% 55%"></div>
<div style="position:absolute;left:0;top:160px;height:80px;padding:0 66px 0 108px;display:flex;align-items:center;background:#e8d4ab;color:#104a70;font-weight:700;font-size:28px;
 clip-path:polygon(0 0,calc(100% - 40px) 0,100% 50%,calc(100% - 40px) 100%,0 100%)">Dokumentasi</div>
<div style="position:absolute;left:460px;top:150px;font-weight:800;font-size:76px;letter-spacing:-2px;line-height:1">Jejak kegiatan kami</div>
<div style="position:absolute;left:462px;top:240px;font-size:24px;font-weight:300;opacity:.9">Bersama instansi, sekolah, perusahaan, dan komunitas.</div>
<div style="position:absolute;left:1640px;top:150px;width:170px;height:170px;border-radius:50%;background:#f6e3b6;color:#104a70;display:flex;align-items:center;
 justify-content:center;text-align:center;font-weight:800;font-size:22px;line-height:1.15;transform:rotate(10deg);box-shadow:0 14px 30px rgba(0,0,0,.25)">Nyata,<br>bukan<br>stok foto</div>
{frames}
</body></html>"""
hp = os.path.join(OUT, "html", "p51_jejak.html")
open(hp, "w").write(html)
png = os.path.join(OUT, "p51_jejak.png")
subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                "--window-size=1920,1080", "--virtual-time-budget=4000", f"--screenshot={png}", "file://" + hp], check=True, capture_output=True)
json.dump(slots, open(os.path.join(OUT, "wall_slots.json"), "w"), indent=1)
print("rendered", png)
for s in slots:
    print(s)
