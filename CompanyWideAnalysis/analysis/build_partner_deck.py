"""Partner program deck (Bahasa Indonesia): price list, profit-sharing scheme and worked examples for strategic partners.

Reuses the design helpers of build_growth_deck.py (same colours, fonts and slide frame). All numbers mirror the partner
contract in SPKnProposalAnalysis/Kemitraan Strategic Partnership - Iza/ (Lampiran A, B, C). Change them in both places.
Run:  python analysis/build_partner_deck.py  ->  TemanJourney_Program_Mitra.pptx
"""
from pathlib import Path

_SRC = Path(__file__).with_name("build_growth_deck.py").read_text()
_g = {"__file__": str(Path(__file__).with_name("build_growth_deck.py")), "__name__": "helpers"}
exec(compile(_SRC[:_SRC.index('FULL = "Our model')], "build_growth_deck.py", "exec"), _g)   # helpers, colours, prs
globals().update({k: v for k, v in _g.items() if not k.startswith("__")})

OUT = Path(__file__).resolve().parent.parent / "TemanJourney_Program_Mitra.pptx"
TAGS.update({"ringkas": ("RINGKASAN", INK_D, WHITE), "strategi": ("STRATEGI", SKY, BLUE), "harga": ("DAFTAR HARGA", CREAM, TERRA),
             "skema": ("BAGI HASIL", RED_BG, TERRA), "contoh": ("CONTOH", SKY, BLUE), "aturan": ("ATURAN MAIN", SAND, MUTED)})
B = lambda t: (t, {"bold": True, "color": INK_D})
N = lambda t: (t, {})
SRC_C = "Perjanjian Kemitraan Strategic Partnership, Lampiran A, B dan C. Semua harga belum termasuk PPN."


def rpf(v):                                                    # full rupiah, Indonesian thousands separator
    return "Rp" + f"{v:,.0f}".replace(",", ".")


def src(s, t=SRC_C):
    text(s, 0.5, 7.02, 11.6, 0.32, "Sumber: " + t, size=9, color=MUTED, name="Source")


def bullets(items, color=TERRA):
    return [[("•  ", {"color": color, "bold": True})] + (it if isinstance(it, list) else [N(it)]) for it in items]


# ------------------------------------------------------------------ numbers (same as the contract)
PPH = 0.02
PRICES = [("Program pelatih regional (ToT)", "1 hari, hingga 30 peserta", 28.5e6, 8e6),
          ("Program pelatih kantor pusat", "3 hari, hingga 40 peserta", 100e6, 25e6),
          ("Program bank BUMN", "1 kelompok, 1 hari, hingga 40 peserta", 51e6, 12e6),
          ("Program bank swasta", "1 kelompok, 1 hari, hingga 40 peserta", 61e6, 12e6)]
profit = lambda p, c: p - c - PPH * p
REG_P, REG_C = PRICES[0][2], PRICES[0][3]
REG_PROF = profit(REG_P, REG_C)
EAP_Y, EAP_COST_Y = 44e6, 4 * 8.525e6
EAP_PROF = profit(EAP_Y, EAP_COST_Y)
ABOVE = 32e6
SEM = [("Program bank BUMN, 2 hari", 102e6, 0.10 * profit(102e6, 24e6), "10% dari profit Rp75,96 juta"),
       ("Pelatihan regional, 4 hari", 114e6, 0.10 * profit(114e6, 32e6), "10% dari profit Rp79,72 juta"),
       ("2 EAP Tahunan", 88e6, 3e6, "tetap 2 x Rp1,5 juta")]
SEM_V = sum(v for _, v, _, _ in SEM)
BONUS = [(50e6, 1e6), (100e6, 3e6), (250e6, 9e6), (400e6, 18e6)]
SEM_B = max([b for t, b in BONUS if SEM_V >= t] or [0])
SEM_T = sum(c for _, _, c, _ in SEM) + SEM_B

# ================================================================== 1. Cover
s = new_slide(None, dark=True)
text(s, 0.8, 0.85, 4.9, 0.38, "TEMAN JOURNEY  ·  PROGRAM MITRA  ·  2026-2027", size=11, bold=True, color=INK_D, fill=AMBER,
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.8, 1.5, 11.6, 1.9, [[("Tumbuh bersama: ", {}), ("Anda membuka pintu", {"color": AMBER}), (", kami mengantar layanan", {})]],
     size=44, bold=True, color=WHITE, font=HEAD, margin=0)
text(s, 0.8, 3.45, 10.8, 0.9, "Daftar harga, cara bagi hasil, dan contoh perhitungan untuk Mitra Strategic Partnership Teman Journey.",
     size=18, color=rgb("C9D2E3"), margin=0)
for i, (v, lab) in enumerate([("10% - 12,5%", "dari profit setiap deal yang Anda tutup"),
                              ("30%", "dari harga yang Anda jual di atas Daftar Harga"),
                              ("Rp18 juta", "bonus terbesar per semester")]):
    stat(s, 0.8 + i * 3.95, 4.7, 3.7, 1.35, v, lab, vcolor=AMBER, fill=DARK2, lcolor=rgb("C9D2E3"), vsize=30, lsize=12)
text(s, 0.8, 6.4, 11.5, 0.4, "Rp juta = Rp1.000.000. Semua harga belum termasuk PPN. Isi deck ini mengikuti Perjanjian Kemitraan.",
     size=12, color=rgb("9FAAC0"), margin=0)
notes(s, "Deck ini untuk Mitra Strategic Partnership. Isinya sama dengan Perjanjian Kemitraan: Pasal 4 (harga), Pasal 5-8 "
         "(komisi, bonus, perpanjangan) dan Lampiran A-C (contoh, biaya standar, daftar harga).")

# ================================================================== 2. Summary
s = new_slide("Mitra mendapat empat jenis penghasilan", accent="empat jenis penghasilan", tag="ringkas",
              sub="Semakin banyak dan semakin besar deal yang Anda tutup, semakin besar bagian Anda.")
PARTS = [("1", "Komisi layanan non-EAP", "10% - 12,5%", "dari profit setiap deal pelatihan, ToT dan program bank yang Anda tutup", INK_D),
         ("2", "Komisi EAP", "Rp1,5 juta", "tetap untuk setiap kontrak EAP tahunan baru (paket 3 bulan Rp500 ribu)", BLUE),
         ("3", "Bagi hasil di atas harga", "30%", "dari selisih bila Anda menjual di atas Daftar Harga", TERRA),
         ("4", "Bonus semester", "Rp18 juta", "bonus terbesar per 6 bulan. Mulai Rp1 juta di Rp50 juta: 1 EAP tahunan dan 1 hari pelatihan sudah cukup", rgb("6379A1"))]
for i, (n, h, v, lab, f) in enumerate(PARTS):
    x = 0.5 + i * 3.1
    card(s, x, 1.55, 2.95, 3.9, "")
    circle(s, x + 0.2, 1.72, 0.55, n, fill=f, size=18)
    text(s, x + 0.85, 1.65, 2.05, 0.7, h, size=14, bold=True, color=INK_D, font=HEAD, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    text(s, x + 0.2, 2.6, 2.6, 2.7, [[(v, {"font": HEAD, "size": 28, "bold": True, "color": TERRA})], [(lab, {"size": 12, "color": BODY})]],
         margin=0.02, space=6)
callout(s, 0.5, 5.7, 12.3, 1.15,
        [[B("Tiga prinsip: "), N("(1) buka setiap penawaran dengan Daftar Harga, tidak pernah di bawahnya tanpa persetujuan Direktur; "
                                 "(2) semua dibayar setelah klien membayar; (3) rekap perhitungan dikirim paling lambat 7 hari kerja, "
                                 "komisi ditransfer paling lambat 14 hari.")]], size=13)
src(s)
notes(s, "Empat sumber penghasilan Mitra. Komisi non-EAP dan EAP dibayar per deal; bagi hasil di atas harga dibayar bila harga jual "
         "lebih tinggi dari Daftar Harga; bonus semester dibayar 14 hari setelah semester berakhir.")

# ================================================================== 3. Strategy and process
s = new_slide("Strategi kami, dan di mana peran Mitra", accent="peran Mitra", tag="strategi",
              sub="Kami tumbuh dari satu kantor pusat ke banyak cabang. Mitra membuka klien baru di luar yang sudah kami tangani.")
STEPS = [("1  Mulai", "Program pelatih di kantor pusat", "Satu kesepakatan tahunan dengan kantor pusat", INK_D),
         ("2  Sebarkan", "Program di kantor cabang", "Pelatihan regional dan program bank per kelompok", BLUE),
         ("3  Pertahankan", "EAP untuk pegawai klien", "Konseling dan check-up setiap kuartal, klien kembali", TERRA)]
for i, (h, a, b, f) in enumerate(STEPS):
    x = 0.5 + i * 4.15
    text(s, x, 1.55, 3.95, 0.55, h, size=16, bold=True, color=WHITE, fill=f, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         anchor=MSO_ANCHOR.MIDDLE, margin=0.2, font=HEAD)
    card(s, x, 2.15, 3.95, 1.05, [[B(a)], [N(b)]], size=12, space=2)
card(s, 0.5, 3.4, 6.05, 1.75, [[("Target Mitra", {"font": HEAD, "bold": True, "size": 15, "color": INK_D})]] +
     bullets(["Bank BUMN dan bank swasta, perusahaan, institusi pendidikan, dan instansi",
              "Klien yang sudah aktif dengan Teman Journey (termasuk unit lain dari institusi yang sama) tetap ditangani tim kami"]),
     size=12, space=4)
card(s, 6.75, 3.4, 6.05, 1.75, [[("Peran masing-masing", {"font": HEAD, "bold": True, "size": 15, "color": INK_D})]] +
     bullets([[B("Mitra: "), N("membuka pintu, menyampaikan proposal, memimpin negosiasi harga")],
              [B("Teman Journey: "), N("materi, harga, pelaksanaan layanan oleh psikolog, laporan")]]), size=12, space=4)
FLOW = ["Daftar prospek", "Konfirmasi 3 hari kerja", "Proposal dengan Daftar Harga", "SPK", "Layanan", "Klien bayar",
        "Rekap 7 hari kerja", "Komisi 14 hari"]
for i, t in enumerate(FLOW):
    x = 0.5 + i * 1.55
    text(s, x, 5.45, 1.42, 0.85, t, size=11, bold=True, color=INK_D if i not in (3, 7) else WHITE,
         fill=CREAM if i not in (3, 7) else TERRA, shape=MSO_SHAPE.ROUNDED_RECTANGLE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, margin=0.05)
    if i < len(FLOW) - 1:
        text(s, x + 1.4, 5.65, 0.17, 0.45, "›", size=16, bold=True, color=MUTED, align=PP_ALIGN.CENTER, margin=0)
text(s, 0.5, 6.4, 12.3, 0.5, "Alur kerja dari prospek sampai komisi: kotak merah adalah dua titik penting, yaitu SPK ditandatangani dan komisi diterima.",
     size=11, color=MUTED, margin=0.02)
src(s, "Rencana pertumbuhan Teman Journey 2027 dan Perjanjian Kemitraan Pasal 4 dan Pasal 9.")
notes(s, "Strategi perusahaan: mulai di kantor pusat, sebarkan ke cabang, pertahankan dengan EAP. Mitra membuka klien baru; "
         "pendaftaran prospek dikonfirmasi 3 hari kerja dan berlaku 6 bulan.")

# ================================================================== 4. Price list A
s = new_slide("Daftar Harga: pelatihan dan program organisasi", accent="pelatihan dan program organisasi", tag="harga",
              sub="Harga dasar. Buka penawaran dengan harga ini; boleh lebih tinggi hingga 20%, tidak boleh lebih rendah.")
rows = [["Layanan", "Isi", "Satuan", "Harga Daftar"]]
for n, isi, p, _ in PRICES:
    rows.append([n, isi + ", termasuk asesmen dan laporan dampak", "per program" if "kantor pusat" in n else "per hari", rpf(p)])
rows += [["Bank Elite", "Program bank + lisensi pelatih + pendampingan 3 bulan", "per hari", "BUMN Rp66.300.000\nswasta Rp79.300.000"],
         ["Asesmen dan laporan dampak", "Peserta tambahan di luar paket", "per peserta", "Rp150.000"],
         ["Event publik", "Event psikoedukasi untuk umum", "per event", "Rp24.000.000"]]
table(s, 0.5, 1.5, 12.3, [3.2, 5.2, 1.5, 2.4], rows, size=12, row_h=0.56)
callout(s, 0.5, 6.15, 12.3, 0.72, [[B("Di luar Jakarta: "), N("biaya perjalanan dan akomodasi tim ditagihkan terpisah kepada klien sebesar biaya aktual. "
                                                              "Harga belum termasuk PPN.")]], size=12)
src(s, "Perjanjian Kemitraan, Lampiran C bagian A.")

# ================================================================== 5. Price list B (EAP)
s = new_slide("Daftar Harga: Employee Assistance Program (EAP)", accent="Employee Assistance Program (EAP)", tag="harga",
              sub="Paket kesejahteraan pegawai. Kontrak tahunan dibayar per kuartal; paket 3 bulan dibayar sekaligus.")
rows = [["Paket", "Isi", "Masa", "Harga Daftar"],
        ["Jeda", "15 sesi konseling online, 1 webinar psikoedukasi", "3 bulan", "Rp12.000.000"],
        ["Jeda Tahunan", "Per kuartal: 15 sesi konseling online, 1 putaran Mental Health Check-Up, laporan agregat anonim, 1 webinar",
         "12 bulan", "Rp11.000.000 per kuartal\n(Rp44.000.000 per tahun)"],
        ["Tumbuh", "24 sesi konseling online, 1 putaran Mental Health Check-Up, laporan agregat anonim, 1 webinar", "3 bulan", "Rp20.000.000"],
        ["Seperjalanan", "36 sesi konseling online, Mental Health Check-Up pre dan post, hasil individual untuk pegawai, laporan agregat dengan perbandingan, 1 seminar offline",
         "3 bulan", "Rp31.500.000"],
        ["Sesi tambahan", "Konseling online 60 menit", "per sesi", "Rp500.000"]]
table(s, 0.5, 1.5, 12.3, [2.1, 6.6, 1.3, 2.3], rows, size=12, row_h=0.7, fills={(2, j): CREAM for j in range(4)})
callout(s, 0.5, 6.0, 12.3, 0.87, [[B("Kenapa EAP penting: "), N("EAP membuat klien kembali setiap kuartal dan membuka jalan ke program lain. "
                                                               "Kerahasiaan: manajemen klien hanya menerima laporan agregat anonim.")]], size=12)
src(s, "Perjanjian Kemitraan, Lampiran C bagian B. Paket Jeda Tahunan disorot karena menjadi kontrak EAP utama.")

# ================================================================== 6. Scheme
s = new_slide("Cara bagi hasil: tiga bagian, satu aturan harga", accent="tiga bagian", tag="skema",
              sub="Non-EAP dihitung dari profit. EAP memakai komisi tetap. Harga di atas Daftar Harga dibagi 30% untuk Anda.")
text(s, 0.5, 1.5, 6.6, 0.4, "1. Layanan non-EAP: persen dari profit event", size=15, bold=True, color=INK_D, font=HEAD, margin=0.02)
table(s, 0.5, 1.95, 6.6, [1.95, 3.55, 1.1],
      [["Tingkat", "Kapan berlaku", "Komisi"],
       ["Penutupan", "Anda menyampaikan proposal, ikut negosiasi harga, SPK ditandatangani. Deal ke-1 s/d ke-4", "10%"],
       ["Penutupan Lanjutan", "Deal penutupan ke-5 dan seterusnya, dihitung dari semua klien Anda", "12,5%"],
       ["Perpanjangan", "Klien yang sama memperpanjang, Anda tetap aktif (lihat aturan main)", "2,5%"]],
      size=11, row_h=0.62)
text(s, 0.5, 4.5, 6.6, 1.4, [[B("Hanya memperkenalkan, tanpa ikut menutup: "), N("tidak ada komisi. Komisi dasar adalah untuk menutup deal.")],
                             [("Profit event = Harga Daftar - biaya standar (Lampiran B) - PPh yang dipotong klien. "
                               "Deal dihitung bila nilai SPK paling sedikit Rp10 juta.", {"color": MUTED})]], size=11, margin=0.02, space=6)
text(s, 7.4, 1.5, 5.4, 0.4, "2. EAP: komisi tetap", size=15, bold=True, color=INK_D, font=HEAD, margin=0.02)
table(s, 7.4, 1.95, 5.4, [3.6, 1.8],
      [["Situasi", "Komisi"], ["Menutup EAP tahunan baru", "Rp1.500.000"], ["Menutup paket EAP 3 bulan", "Rp500.000"],
       ["Perpanjangan 1 tahun", "Rp1.000.000"]], size=11, row_h=0.5)
card(s, 7.4, 4.45, 5.4, 1.45, [[("3. Di atas Daftar Harga", {"font": HEAD, "bold": True, "size": 15, "color": INK_D})],
                               [N("Anda menerima "), B("30% dari selisih"), N(" di atas Daftar Harga, untuk non-EAP maupun EAP. "
                                                                                   "Batas atas tanpa persetujuan: +20%.")]],
     size=12, fill=RED_BG, space=4)
callout(s, 0.5, 6.1, 12.3, 0.75, [[B("Pengaman: "), N("setiap deal tetap memberi Teman Journey margin paling sedikit 15% setelah komisi. "
                                                     "Bila tidak, komisi otomatis dikurangi sampai batas itu.")]], size=12)
src(s, "Perjanjian Kemitraan, Pasal 5, 6 dan 7 ayat (1).")

# ================================================================== 7. Semester bonus
s = new_slide("Bonus semester: mulai Rp50 juta, sampai Rp18 juta", accent="mulai Rp50 juta", tag="skema",
              sub="Paket kami bernilai besar, jadi satu atau dua deal sudah mencapai ambang pertama.")
bar(s, 0.5, 1.45, 5.6, 4.45, [f"Rp{t / 1e6:.0f} juta" for t, _ in BONUS], [("Bonus", [b / 1e6 for _, b in BONUS])],
    "Total bonus semester (Rp juta) menurut nilai deal dalam 6 bulan", point_colors=[GREY, LBLUE, BLUE, TERRA], fmt='"Rp"0" juta"', size=12)
text(s, 6.4, 1.5, 6.4, 0.4, "Seberapa dekat? Contoh deal yang mencapai setiap ambang", size=14, bold=True, color=INK_D, font=HEAD, margin=0.02)
table(s, 6.4, 1.95, 6.4, [1.25, 3.85, 1.3],
      [["Ambang", "Contoh deal dalam 6 bulan", "Bonus"],
       ["Rp50 juta", "1 EAP tahunan (Rp44 juta) + 1 hari pelatihan regional (Rp28,5 juta) = Rp72,5 juta", "Rp1 juta"],
       ["Rp100 juta", "1 program kantor pusat (Rp100 juta), atau 1 program bank BUMN 2 hari (Rp102 juta)", "Rp3 juta"],
       ["Rp250 juta", "1 program kantor pusat + 1 program bank swasta 2 hari + 1 EAP tahunan = Rp266 juta", "Rp9 juta"],
       ["Rp400 juta", "2 program kantor pusat + 2 program bank BUMN 2 hari = Rp404 juta", "Rp18 juta"]],
      size=11, row_h=0.72, fills={(1, j): CREAM for j in range(3)})
callout(s, 0.5, 6.0, 12.3, 0.87,
        [[B("Cara menghitung: "), N("hanya deal yang Anda tutup (bukan perpanjangan), nilai SPK minimal Rp10 juta, masuk ke semester saat klien "
                                    "membayar pertama kali. EAP tahunan dihitung nilai penuh. Bonus satu angka sesuai ambang tertinggi, "
                                    "dibayar 14 hari setelah semester berakhir.")]], size=12)
src(s, "Perjanjian Kemitraan, Pasal 7 ayat (2) sampai (5), dan Daftar Harga (Lampiran C).")

# ================================================================== 8. Per-deal examples
s = new_slide("Contoh per deal: harga daftar, di atas harga, dan EAP", accent="Contoh per deal", tag="contoh",
              sub="Tiga contoh paling umum, memakai Daftar Harga dan biaya standar dalam perjanjian.")
EX = [("A. Pelatihan regional, harga daftar", INK_D,
       [("Harga", rpf(REG_P)), ("Biaya standar", "- " + rpf(REG_C)), ("PPh dipotong klien (2%)", "- " + rpf(PPH * REG_P)),
        ("Profit event", rpf(REG_PROF)), ("Komisi penutupan 10%", rpf(0.1 * REG_PROF))], rpf(0.1 * REG_PROF)),
      ("B. Dijual di atas harga: Rp32 juta", TERRA,
       [("Komisi penutupan 10%", rpf(0.1 * REG_PROF)), ("Selisih di atas harga", rpf(ABOVE - REG_P)),
        ("Bagi hasil 30% dari selisih", rpf(0.3 * (ABOVE - REG_P))), ("Nilai untuk bonus", rpf(ABOVE)), ("", "")],
       rpf(0.1 * REG_PROF + 0.3 * (ABOVE - REG_P))),
      ("C. EAP tahunan, dibayar per kuartal", BLUE,
       [("Nilai kontrak", rpf(EAP_Y)), ("Dibayar", "4 x " + rpf(EAP_Y / 4)), ("Komisi tetap, kontrak baru", rpf(1.5e6)),
        ("Perpanjangan tahun berikutnya", rpf(1e6)), ("Nilai untuk bonus", rpf(EAP_Y))], rpf(1.5e6))]
for i, (h, f, lines, tot) in enumerate(EX):
    x = 0.5 + i * 4.15
    text(s, x, 1.55, 3.95, 0.55, h, size=14, bold=True, color=WHITE, fill=f, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         anchor=MSO_ANCHOR.MIDDLE, margin=0.15, font=HEAD)
    table(s, x, 2.2, 3.95, [2.35, 1.6], [["Komponen", "Nilai"]] + [[a, b] for a, b in lines if a], size=11, row_h=0.5,
          first_col_bold=False)
    stat(s, x, 5.2, 3.95, 0.85, tot, "diterima Mitra", vcolor=TERRA, vsize=22, lsize=11)
callout(s, 0.5, 6.2, 12.3, 0.68, [[B("Deal ke-5 dan seterusnya: "), N(f"komisi penutupan naik menjadi 12,5%. Untuk contoh A: {rpf(0.125 * REG_PROF)}.")]],
        size=12)
src(s, "Perjanjian Kemitraan, Lampiran A Contoh 1, 2 dan 3.")

# ================================================================== 9. Semester example
s = new_slide(f"Contoh satu semester: Anda menerima {rpf(SEM_T)}", accent=rpf(SEM_T), tag="contoh",
              sub="Tiga deal dalam 6 bulan. Komisi per deal ditambah bonus semester.")
rows = [["Deal dalam semester", "Nilai SPK", "Komisi", "Dasar"]]
for n, v, c, d in SEM:
    rows.append([n, rpf(v), rpf(c), d])
rows += [["Jumlah nilai deal", rpf(SEM_V), "", ""], ["Bonus semester", "", rpf(SEM_B), "melewati ambang Rp250 juta"],
         ["Total diterima Mitra", "", rpf(SEM_T), f"sekitar Rp{SEM_T / 6e6:.1f} juta per bulan".replace(".", ",")]]
table(s, 0.5, 1.5, 8.0, [2.8, 1.75, 1.6, 1.85], rows, size=12, row_h=0.6,
      fills={(6, j): CREAM for j in range(4)})
bar(s, 8.8, 1.45, 4.0, 4.4, ["Komisi", "Bonus"], [("Rp juta", [round((SEM_T - SEM_B) / 1e6, 1), SEM_B / 1e6])],
    "Asal penghasilan (Rp juta)", point_colors=[BLUE, TERRA], fmt='0.0', size=12)
callout(s, 0.5, 6.0, 12.3, 0.87, [[B("Agar lebih besar: "), N("jual di atas Daftar Harga (bagian 30%), capai ambang Rp400 juta (bonus Rp18 juta), "
                                                             "dan tutup deal ke-5 agar komisi naik ke 12,5%.")]], size=12)
src(s, "Perjanjian Kemitraan, Lampiran A Contoh 5.")

# ================================================================== 10. Rules
s = new_slide("Aturan main yang menjaga kita berdua", accent="menjaga kita berdua", tag="aturan",
              sub="Aturan ini ada di perjanjian. Tujuannya: harga terjaga, perhitungan jelas, dan tidak ada salah paham.")
RULES = [("Harga", ["Buka penawaran dengan Daftar Harga", "Tidak boleh di bawah Daftar Harga tanpa persetujuan tertulis Direktur",
                    "Di atas Daftar Harga boleh sampai +20%; lebih dari itu perlu persetujuan"]),
         ("Perpanjangan", ["Minimal 1 pertemuan atau panggilan dengan PIC klien per kuartal, dikonfirmasi tim kami",
                           "Ikut dalam proses perpanjangan; tingkat komisi ditulis saat SPK",
                           "Berakhir 24 bulan setelah deal terakhir Anda dengan klien itu; tidak dihitung untuk bonus"]),
         ("Pembayaran dan pajak", ["Rekap 7 hari kerja setelah klien membayar; komisi paling lambat 14 hari",
                                   "Klien bayar bertahap: komisi persen ikut bertahap; komisi EAP setelah pembayaran pertama",
                                   "Dipotong PPh 21 (bukan pegawai), bukti potong diberikan"])]
for i, (h, items) in enumerate(RULES):
    x = 0.5 + i * 4.15
    text(s, x, 1.55, 3.95, 0.55, h, size=16, bold=True, color=WHITE, fill=[INK_D, BLUE, TERRA][i], shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         anchor=MSO_ANCHOR.MIDDLE, margin=0.2, font=HEAD)
    card(s, x, 2.15, 3.95, 3.7, bullets(items), size=13, space=10)
callout(s, 0.5, 6.05, 12.3, 0.82, [[B("Denda dan keadaan kahar: "), N("pelanggaran kerahasiaan, pengalihan klien, atau harga di bawah Daftar Harga "
                                                                  "dikenai denda, ditambah ganti rugi bila kerugian lebih besar. Keadaan kahar diberitahukan "
                                                                  "paling lambat 2 hari; bila 21 hari tetap tidak bisa dipenuhi, dibatalkan bersama.")]], size=12)
src(s, "Perjanjian Kemitraan, Pasal 4, 8, 9, 10, 12, 15 dan 19.")

assert len(prs.slides) == 10, len(prs.slides)
prs.save(OUT)
print("wrote", OUT, len(prs.slides), "slides")
