"""Re-titles the Teman Journey proposal cover for the BSN employee mapping offer.

Usage: python make_cover.py <cover_source_cbp.png> <cover_bsn.png>
Paints over the old date and title with the cover background, removes the old client logo
from the gold "prepared for" band (refilled with the band's own gradient), and writes the
client name as text. We have no BSN logo file, so no logo is drawn.
"""
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps

src, dst = sys.argv[1], sys.argv[2]
img = Image.open(src).convert("RGB")
bg = img.getpixel((700, 260))
draw = ImageDraw.Draw(img)
draw.rectangle([30, 200, 1060, 590], fill=bg)

# Remove the old client logo: refill each column with the band colour at its top edge.
for x in range(330, 1414):
    col = img.getpixel((x, 633))
    draw.line([(x, 636), (x, 744)], fill=col)

date_font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Georgia.ttf", 54)
title_font = ImageFont.truetype("/System/Library/Fonts/Avenir Next.ttc", 80, index=0)
client_font = ImageFont.truetype("/System/Library/Fonts/Avenir Next.ttc", 46, index=0)

shadow = Image.new("RGBA", img.size, (0, 0, 0, 0))
sd = ImageDraw.Draw(shadow)
lines = ["PROPOSAL -", "Pemetaan Karyawan", "dengan Tes BFI dan DISC"]
y0, step = 300, 92
for i, line in enumerate(lines):
    sd.text((60, y0 + i * step + 5), line, font=title_font, fill=(0, 0, 0, 90))
shadow = shadow.filter(ImageFilter.GaussianBlur(5))
img = Image.alpha_composite(img.convert("RGBA"), shadow)

draw = ImageDraw.Draw(img)
draw.text((84, 222), "November  2026", font=date_font, fill="white")
for i, line in enumerate(lines):
    draw.text((56, y0 + i * step), line, font=title_font, fill="white")
draw.text((352, 662), "Bank Syariah Nasional", font=client_font, fill=(71, 85, 113))

# Standard Teman Journey cover photos (Galih, 4 Nov 2026), in the template's three rounded frames.
# Frames: y 1215 to 1673, corner radius about 50 px. Left and right frames bleed off the page edge.
PHOTOS = Path(__file__).resolve().parents[2] / "_brand" / "cover-photos"
FRAMES = [  # photo file, frame box (x0, y0, x1, y1), mask box (corners off-canvas where the frame bleeds)
    ("1-tim-teman-journey.webp", (0, 1215, 424, 1674), (-80, 0, 424, 459)),
    ("2-pelatihan-korporat.webp", (473, 1215, 955, 1674), (0, 0, 482, 459)),
    ("3-pelatihan-kemendikdasmen.webp", (1005, 1215, 1414, 1674), (0, 0, 489, 459)),
]
img = img.convert("RGB")
for name, (x0, y0, x1, y1), mask_box in FRAMES:
    w, h = x1 - x0, y1 - y0
    photo = ImageOps.fit(Image.open(PHOTOS / name).convert("RGB"), (w, h), Image.LANCZOS, centering=(0.5, 0.4))
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle(mask_box, radius=50, fill=255)
    img.paste(photo, (x0, y0), mask)
img.save(dst)
print("saved", dst)
