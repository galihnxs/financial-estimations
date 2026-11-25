"""Writes update_deck_censored.html: every rupiah amount in the deck text becomes "Rp •••".
Data URIs (embedded images) are left untouched. Used for the price-censored PPTX version."""
import re
from pathlib import Path

HERE = Path(__file__).parent
src = (HERE / "update_deck.html").read_text()
AMOUNT = re.compile(r"Rp\s?\d[\d.,]*(?:\s?(?:jt|juta|rb|ribu))?")
parts = re.split(r"(data:[^\"')\s]+)", src)
out, n = [], 0
for i, part in enumerate(parts):
    if i % 2 == 0:
        part, k = AMOUNT.subn("Rp •••", part)
        n += k
    out.append(part)
(HERE / "update_deck_censored.html").write_text("".join(out))
print("censored amounts:", n)
