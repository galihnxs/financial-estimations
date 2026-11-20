"""Write 01_Daftar_Produk_Final.md from a RECALCULATED copy of Master_Produk_dan_Harga.xlsx.

openpyxl files have no cached values, so first open and save the workbook in Excel (or recalculate it with the
`formulas` package), then run:  python3 -I export_product_list.py <recalculated.xlsx>
"""
import os
import sys
import openpyxl

src = sys.argv[1]
out = os.path.join(os.path.dirname(__file__), "..", "01_Daftar_Produk_Final.md")
wb = openpyxl.load_workbook(src, data_only=True)
ws = wb["Katalog"] if "Katalog" in wb.sheetnames else wb["KATALOG"]
rows = [r for r in ws.iter_rows(min_row=6, values_only=True) if r[0]]


def rp(v):
    if v in (None, ""):
        return "-"
    v = round(v)
    if v >= 1_000_000:
        s = f"{v / 1_000_000:,.2f}".rstrip("0").rstrip(".").replace(",", "X").replace(".", ",").replace("X", ".")
        return f"Rp{s} jt"
    if v >= 1000:
        return f"Rp{v / 1000:,.1f}".rstrip("0").rstrip(".").replace(".", ",") + "rb"
    return f"Rp{v:,}".replace(",", ".")


def pct(v):
    return "-" if v in (None, "") else f"{v * 100:.0f}%"


L = []
L.append("# Daftar Produk dan Harga Final: Teman Journey + Coreitera\n")
L.append("Tanggal: 19 Nov 2026. Sumber: `Master_Produk_dan_Harga.xlsx` (dihitung ulang). Semua harga Rupiah, "
         "belum termasuk PPN (Teman Journey dan Coreitera Non-PKP). Di luar Jakarta: perjalanan dan akomodasi ditagih sebesar biaya aktual.\n")
L.append("**Aturan harga yang dipakai**\n")
L.append("- **Penuh**: margin bersih paling sedikit 30% setelah PPh 23 (2%) dan komisi 10% dari laba kotor. Tidak ada aturan kelipatan 3x.")
L.append("- **Harga tetap**: paket EAP memakai harga hari ini. Margin tetap dicek; yang di bawah 30% perlu persetujuan Direktur saat dijual.")
L.append("- **Pasar**: B2C dan event bertiket. Harga mengikuti pasar dan bagi hasil psikolog; tidak dinaikkan otomatis.")
L.append("- **Harga pokok**: dijual sekitar biaya (mesin tes white-label Coreitera).")
L.append("- Biaya psikolog: B2B (EAP, MHCU) flat Rp350rb per sesi; B2C bagi hasil per tier (60/65/65/70%); kerja dokumen Rp110rb/jam; tatap muka asesmen Rp350rb/jam.\n")

fams = []
for r in rows:
    if r[2] not in fams:
        fams.append(r[2])

for fam in fams:
    fr = [r for r in rows if r[2] == fam]
    L.append(f"## {fam}\n")
    if fam == "E. B2C Konseling":
        L.append("| Layanan | Tier 1 (0-2 th) | Tier 2 (2-5 th) | Tier 3 (5-10 th) | Tier 4 (10+ th) | Margin bersih TJ |")
        L.append("|---|---|---|---|---|---|")
        groups = {}
        for r in fr:
            base = r[3].rsplit(", Tier", 1)[0]
            groups.setdefault(base, []).append(r)
        for base, g in groups.items():
            ms = [x[22] for x in g]
            L.append(f"| {base} | " + " | ".join(rp(x[19]) for x in g) + f" | {pct(min(ms))}-{pct(max(ms))} |")
        L.append("\nMargin setelah fee psikolog dan biaya platform Rp50rb per sesi, sebelum biaya iklan.\n")
        continue
    L.append("| Kode | Produk | Isi | Pembeli | Satuan | Harga daftar | Harga lama | Margin bersih | Status |")
    L.append("|---|---|---|---|---|---|---|---|---|")
    for r in fr:
        code, prod, scope, buyer, uom, status = r[0], r[3], r[4], r[5], r[6], r[7]
        cur, new, chg, margin, check = r[10], r[19], r[20], r[22], r[24]
        old = "-" if (chg in (0, None) or (isinstance(chg, (int, float)) and abs(chg) < 1e-6)) else ("baru" if chg == "Baru" else rp(cur))
        flag = check if check not in ("Info",) else status
        if check in ("Lolos", "Lolos (margin)", "Info"):
            flag = status
        else:
            flag = f"{status}; **{check}**"
        L.append(f"| {code} | {prod} | {scope} | {buyer} | {uom} | **{rp(new)}** | {old} | {pct(margin)} | {flag} |")
    L.append("")

changed = [r for r in rows if isinstance(r[20], (int, float)) and r[20] > 1e-6]
L.append("## Harga yang berubah dari harga hari ini\n")
L.append("| Kode | Produk | Harga lama | Harga baru | Kenaikan | Alasan |")
L.append("|---|---|---|---|---|---|")
for r in changed:
    L.append(f"| {r[0]} | {r[3]} | {rp(r[10])} | {rp(r[19])} | +{r[20] * 100:.0f}% | {"Harga ditetapkan (keputusan 19 Nov 2026)" if len(r) > 26 and r[26] else "Margin bersih di bawah 30%"} |")
L.append("")
below = [r for r in rows if r[24] == "Di bawah margin"]
L.append("## Harga tetap di bawah margin 30% (perlu persetujuan Direktur saat dijual)\n")
L.append("| Kode | Produk | Harga | Margin bersih | Harga minimum untuk 30% |")
L.append("|---|---|---|---|---|")
for r in below:
    L.append(f"| {r[0]} | {r[3]} | {rp(r[19])} | {pct(r[22])} | {rp(r[17])} |")
L.append("")
open(out, "w").write("\n".join(L))
print("wrote", os.path.abspath(out), len(rows), "products")
