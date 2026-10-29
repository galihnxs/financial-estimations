"""Re-titles the TJ proposal cover (from the CBP Rupiah proposal) for the JAMPURNAS 2026 jury proposal.

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
lines = ["PROPOSAL -", "Dewan Juri JAMPURNAS", "Apresiasi CBP Rupiah 2026"]
size = 80
while True:  # shrink until the longest line fits the painted block
    title_font = ImageFont.truetype("/System/Library/Fonts/Avenir Next.ttc", size, index=0)
    if max(title_font.getlength(t) for t in lines) <= 990 or size <= 50:
        break
    size -= 2

# Soft drop shadow like the original title
shadow = Image.new("RGBA", img.size, (0, 0, 0, 0))
sd = ImageDraw.Draw(shadow)
y0, step = 300, int(size * 1.15)
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
