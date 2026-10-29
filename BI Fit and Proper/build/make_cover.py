"""Re-titles the TJ proposal cover (from the CBP Rupiah proposal) for the Fit and Proper Test.

Usage: python make_cover.py <cover_in.png> <cover_out.png>
Paints over the old date and title block with the cover's own background colour.
"""
import sys

from PIL import Image, ImageDraw, ImageFilter, ImageFont

src, dst = sys.argv[1], sys.argv[2]
img = Image.open(src).convert("RGB")
bg = img.getpixel((700, 260))
draw = ImageDraw.Draw(img)
draw.rectangle([30, 200, 1060, 590], fill=bg)

date_font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Georgia.ttf", 54)
title_font = ImageFont.truetype("/System/Library/Fonts/Avenir Next.ttc", 80, index=0)

# Soft drop shadow like the original title
shadow = Image.new("RGBA", img.size, (0, 0, 0, 0))
sd = ImageDraw.Draw(shadow)
lines = ["PROPOSAL -", "Asesmen dalam Rangka", "Fit and Proper Test"]
y0, step = 300, 92
for i, line in enumerate(lines):
    sd.text((60, y0 + i * step + 5), line, font=title_font, fill=(0, 0, 0, 90))
shadow = shadow.filter(ImageFilter.GaussianBlur(5))
img = Image.alpha_composite(img.convert("RGBA"), shadow)

draw = ImageDraw.Draw(img)
draw.text((84, 222), "Oktober  2026", font=date_font, fill="white")
for i, line in enumerate(lines):
    draw.text((56, y0 + i * step), line, font=title_font, fill="white")
img.convert("RGB").save(dst)
print("saved", dst)
