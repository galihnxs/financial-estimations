"""Write 01_Daftar_Produk_Final.md from a RECALCULATED copy of Master_Produk_dan_Harga.xlsx.

openpyxl files have no cached values, so first recalculate the workbook (Excel, or the `formulas` package), then run:
    /opt/homebrew/bin/python3 -I export_product_list.py <recalculated.xlsx>
"""
import os
import sys
import openpyxl

src = sys.argv[1]
out = os.path.join(os.path.dirname(__file__), "..", "01_Daftar_Produk_Final.md")
wb = openpyxl.load_workbook(src, data_only=True)
ws = wb["Katalog"] if "Katalog" in wb.sheetnames else wb["KATALOG"]
rows = [r for r in ws.iter_rows(min_row=6, values_only=True) if r[0]]
inp = wb["Input"] if "Input" in wb.sheetnames else wb["INPUT"]
seg = {r[0]: r[2] for r in inp.iter_rows(min_row=15, max_row=17, values_only=True) if r[0]}

# column positions in Katalog (0-based)
CODE, LINE, PROD, SCOPE, BUYER, UOM, STATUS, ROLE, RULE = 0, 2, 3, 4, 5, 6, 7, 8, 9
CUR, COST, CURM, BASE, EDU, MEDU, KORP, MKORP, REG, MREG, CHECK = 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 26


def num(v):
    return isinstance(v, (int, float))


def rp(v):
    if not num(v):
        return "-"
    v = round(v)
    if v == 0:
        return "Gratis"
    if v >= 1_000_000:
        s = f"{v / 1_000_000:,.2f}".rstrip("0").rstrip(".").replace(",", "X").replace(".", ",").replace("X", ".")
        return f"Rp{s} jt"
    if v >= 1000:
        return f"Rp{v / 1000:,.1f}".rstrip("0").rstrip(".").replace(".", ",") + "rb"
    return f"Rp{v:,}".replace(",", ".")


def pct(v):
    return f"{v * 100:.0f}%" if num(v) else "-"


def cell(price, margin):
    return "-" if not num(price) else f"**{rp(price)}** ({pct(margin)})"


L = []
L.append("# Daftar Produk dan Harga Final: Teman Journey + Coreitera\n")
L.append("Tanggal: 20 Nov 2026. Sumber: `Master_Produk_dan_Harga.xlsx` (dihitung ulang). Struktur dan aturan: "
         "`05_Struktur_Produk_dan_Segmen.md`. Semua harga Rupiah, belum termasuk PPN (Teman Journey dan Coreitera Non-PKP). "
         "Di luar Jakarta: perjalanan dan akomodasi ditagih sebesar biaya aktual. Angka dalam kurung = margin bersih.\n")
L.append("**Aturan harga**\n")
L.append(f"- **Satu biaya, harga per segmen.** Pendidikan dan komunitas: target margin bersih {pct(seg.get('SEG_EDU'))}. "
         f"Korporat dan pemerintah: {pct(seg.get('SEG_KORP'))}. Regulator: menuju {pct(seg.get('SEG_REG'))}, minimal 30%. "
         "Margin bersih setelah PPh 23 (2%) dan komisi 10% dari laba kotor.")
L.append("- **Pengaman.** Harga korporat dan regulator tidak pernah di bawah harga hari ini. Item yang sudah pernah dibeli regulator tetap memakai harga lama.")
L.append("- **Peran.** Pintu masuk (gratis bila biaya antar hampir nol), Inti (membawa laba, tidak pernah gratis), Retensi (kontrak berulang).")
L.append("- **Pendidikan:** minimal Rp10 jt per kontrak; bagi hasil sekolah maks. 10%.")
L.append("- **Pasar / Langganan / Harga pokok:** satu harga (konseling dan tes individu, event bertiket, Coreitera).")
L.append("- Harga kosong (-) = produk tidak dijual ke segmen itu.\n")

lines = []
for r in rows:
    if r[LINE] not in lines:
        lines.append(r[LINE])

for ln in lines:
    lr = [r for r in rows if r[LINE] == ln]
    L.append(f"## {ln}\n")
    seg_rows = [r for r in lr if r[RULE] == "Penuh"]
    if seg_rows:
        L.append("| Kode | Produk | Isi | Peran | Satuan | Pendidikan | Korporat | Regulator | Catatan |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        for r in seg_rows:
            note = []
            if num(r[CUR]) and r[CUR] > 0 and any(num(x) and abs(x - r[CUR]) > 1 for x in (r[EDU], r[KORP], r[REG])):
                note.append(f"harga lama {rp(r[CUR])}")
            if r[STATUS] and str(r[STATUS]).startswith("Baru"):
                note.append("baru")
            if r[CHECK] not in ("Lolos", "Info"):
                note.append(f"**{r[CHECK]}**")
            L.append(f"| {r[CODE]} | {r[PROD]} | {r[SCOPE]} | {r[ROLE]} | {r[UOM]} | {cell(r[EDU], r[MEDU])} | "
                     f"{cell(r[KORP], r[MKORP])} | {cell(r[REG], r[MREG])} | {'; '.join(note) or '-'} |")
        L.append("")
    b2c = [r for r in lr if str(r[CODE]).startswith("TJ-B2C")]
    if b2c:
        L.append("**Individu (harga pasar, per tier psikolog)**\n")
        L.append("| Layanan | Tier 1 (0-2 th) | Tier 2 (2-5 th) | Tier 3 (5-10 th) | Tier 4 (10+ th) | Margin bersih TJ |")
        L.append("|---|---|---|---|---|---|")
        groups = {}
        for r in b2c:
            groups.setdefault(r[PROD].rsplit(", Tier", 1)[0], []).append(r)
        for base, g in groups.items():
            ms = [x[CURM] for x in g]
            L.append(f"| {base} | " + " | ".join(rp(x[BASE]) for x in g) + f" | {pct(min(ms))}-{pct(max(ms))} |")
        L.append("\nMargin setelah fee psikolog dan biaya platform Rp50rb per sesi, sebelum biaya iklan.\n")
    one = [r for r in lr if r[RULE] != "Penuh" and not str(r[CODE]).startswith("TJ-B2C")]
    if one:
        L.append("**Satu harga (Pasar, Langganan, Harga pokok, Pintu masuk)**\n")
        L.append("| Kode | Produk | Isi | Peran | Aturan | Satuan | Harga | Margin bersih |")
        L.append("|---|---|---|---|---|---|---|---|")
        for r in one:
            m = r[CURM] if num(r[CUR]) and r[CUR] == r[BASE] else None
            L.append(f"| {r[CODE]} | {r[PROD]} | {r[SCOPE]} | {r[ROLE]} | {r[RULE]} | {r[UOM]} | **{rp(r[BASE])}** | {pct(m)} |")
        L.append("")

below = [r for r in rows if r[CHECK] not in ("Lolos", "Info")]
L.append("## Harga di bawah 30% (perlu persetujuan Direktur saat dijual)\n")
L.append("| Kode | Produk | Segmen | Harga | Margin bersih | Harga minimum 30% |")
L.append("|---|---|---|---|---|---|")
for r in below:
    L.append(f"| {r[CODE]} | {r[PROD]} | Regulator (harga lama dibekukan) | {rp(r[REG])} | {pct(r[MREG])} | {rp(r[BASE])} |")
L.append("")
open(out, "w").write("\n".join(L))
print("wrote", os.path.abspath(out), len(rows), "products")
