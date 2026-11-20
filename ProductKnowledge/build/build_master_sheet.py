"""Build Master_Produk_dan_Harga.xlsx: one catalog, one cost table, one margin rule for TJ + Coreitera.

Run:  python3 -I build_master_sheet.py   (needs openpyxl)
Every price, cost and margin in the workbook is a formula. Inputs are blue. Edit inputs in Excel,
or edit the data below and rebuild.
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

OUT = os.path.join(os.path.dirname(__file__), "..", "Master_Produk_dan_Harga.xlsx")

F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F)
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
TITLE = Font(name=F, bold=True, size=14)
HDR_FONT = Font(name=F, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
YELLOW = PatternFill("solid", fgColor="FFFF00")
SECTION = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))
RP = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
MULT = '0.0"x";(0.0"x");"-"'

wb = Workbook()


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.alignment = HDR_FONT, HDR_FILL, Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.row_dimensions[row].height = 32


def put(ws, ref, value, font=BLACK, fmt=None, fill=None, align=None):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    return c


# ---------------------------------------------------------------------------------------------
# 1. Panduan
# ---------------------------------------------------------------------------------------------
ws = wb.active
ws.title = "Panduan"
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 110
put(ws, "A1", "Master Produk dan Harga: Teman Journey + Coreitera", TITLE)
put(ws, "A2", "Satu sumber harga untuk tim sales. Dibuat 19 Nov 2026 dari semua proposal, RAB, model 2027 v15, "
              "Daftar Harga Kemitraan (Lampiran B dan C) dan model Coreitera. Semua angka Rupiah, belum termasuk PPN "
              "(Teman Journey dan Coreitera Non-PKP).")
ws["A2"].alignment = WRAP
ws.merge_cells("A2:B2")
ws.row_dimensions[2].height = 45
rows = [
    ("CARA BACA", None),
    ("Teks biru", "Input. Boleh diubah. Semua angka lain mengikuti otomatis."),
    ("Teks hitam", "Rumus. Jangan ditimpa."),
    ("Teks hijau", "Mengambil angka dari sheet lain."),
    ("Isian kuning", "Estimasi yang perlu diganti dengan angka asli."),
    ("ATURAN HARGA (keputusan 19 Nov 2026)", None),
    ("Aturan harga", "Margin bersih paling sedikit 30% setelah PPh 23 yang dipotong klien dan komisi 10% dari laba kotor "
                     "(keputusan 19 Nov 2026: hanya aturan margin, tanpa aturan kelipatan 3x)."),
    ("Kelipatan (opsional)", "Kolom kelipatan atas biaya luar grup tetap ditampilkan sebagai informasi. Batas kelipatan di Input = 0, "
                             "artinya tidak dipakai. Isi angka (misal 2) hanya jika ingin menambah batas kedua."),
    ("Berlaku untuk", "Produk dengan Aturan harga = 'Penuh' (B2B Teman Journey dan produk berbayar Coreitera). "
                      "Batas margin bisa diatur per keluarga produk di sheet Input."),
    ("Pengecualian", "'Pasar' = harga mengikuti pasar dan bagi hasil psikolog (konseling B2C, tes B2C, event bertiket). "
                     "Margin tetap ditampilkan, harga tidak dinaikkan otomatis. 'Harga pokok' = dijual sekitar biaya "
                     "(mesin tes white-label Coreitera). 'Harga tetap' = harga hari ini dipertahankan (paket EAP, keputusan 19 Nov 2026); margin bersih tetap dicek dan ditandai bila di bawah batas."),
    ("Biaya psikolog", "B2C: harga x bagi hasil tier psikolog (60/65/65/70%). B2B (EAP, MHCU): flat Rp350rb per sesi 60 menit "
                       "untuk semua tier (keputusan 19 Nov 2026). Kerja dokumen Rp110rb/jam, tatap muka asesmen Rp350rb/jam."),
    ("Harga daftar baru", "Jika harga saat ini lolos kedua aturan, harga tidak berubah. Jika tidak lolos, harga naik ke "
                          "harga minimum aturan, dibulatkan ke atas (Rp1rb di bawah Rp100rb, Rp25rb di bawah Rp1 jt, Rp100rb di bawah Rp10 jt, "
                          "Rp500rb di atasnya). Harga saat ini tetap tampil sebagai pembanding. Kolom 'Harga ditetapkan' di Katalog (biru) mengganti harga hasil rumus bila diisi (MHCU Flash dan Deep)."),
    ("SHEET", None),
    ("Ringkasan", "Jumlah produk, berapa yang lolos aturan, dan dampak kenaikan harga per keluarga."),
    ("Katalog", "Semua produk: isi, pembeli, satuan, harga saat ini, biaya, harga minimum, harga daftar baru, margin."),
    ("Rincian_Biaya", "Biaya tiap produk per komponen: jumlah x tarif. Inilah sumber kolom Biaya di Katalog."),
    ("Tarif_Biaya", "Satu tarif untuk tiap sumber daya (jam psikolog, alat tes, webinar, server, AI)."),
    ("Input", "PPh 23, komisi, tier psikolog EAP, aturan per keluarga, harga dan bagi hasil tier, pembulatan."),
    ("Simulasi_Penawaran", "Pilih kode produk dan jumlah, isi diskon: keluar total harga, laba bersih, margin, dan siapa yang harus menyetujui."),
    ("Aturan_Penjualan", "Aturan diskon, pembayaran, perjalanan, dan persetujuan untuk sales dan mitra."),
    ("Sumber_Konflik", "Harga yang dulu berbeda antar dokumen, dan bagaimana master ini menyelesaikannya."),
]
r = 4
for a, b in rows:
    if b is None:
        put(ws, f"A{r}", a, BOLD, fill=SECTION)
        ws[f"B{r}"].fill = SECTION
    else:
        put(ws, f"A{r}", a, BOLD)
        put(ws, f"B{r}", b, align=WRAP)
    r += 1
ws["A5"].font = Font(name=F, bold=True, color="0000FF")
ws["A7"].font = Font(name=F, bold=True, color="008000")
ws["A8"].fill = YELLOW

# ---------------------------------------------------------------------------------------------
# 2. Input
# ---------------------------------------------------------------------------------------------
wi = wb.create_sheet("Input")
header(wi, 4, ["Kunci", "Parameter", "Nilai", "Satuan", "Status", "Sumber / catatan"], [16, 52, 16, 14, 14, 70])
put(wi, "A1", "Input: semua angka yang mengatur harga", TITLE)
put(wi, "A2", "Biru = input. Ubah di sini, Katalog dan Simulasi ikut berubah.")
params = [
    ("TARIF_PPH23", "PPh 23 yang dipotong klien B2B", 0.02, "% harga", "Pasti", "Semua RAB BI dan BSN. Dihitung sebagai biaya (konservatif).", PCT),
    ("KOMISI", "Komisi dari laba kotor (produk 'Penuh')", 0.10, "%", "Pasti", "Semua RAB internal. Mitra: lihat Aturan_Penjualan.", PCT),
    ("STEP0", "Pembulatan harga di bawah Rp100 ribu", 1000, "Rp", "Rencana", "Untuk produk software dan tes per peserta.", RP),
    ("STEP1", "Pembulatan harga Rp100 ribu - Rp1 juta", 25000, "Rp", "Rencana", "Sama dengan B2C Pricing.", RP),
    ("STEP2", "Pembulatan harga Rp1 juta - Rp10 juta", 100000, "Rp", "Rencana", "Sama dengan pembulatan paket B2C.", RP),
    ("STEP3", "Pembulatan harga di atas Rp10 juta", 500000, "Rp", "Rencana", "Estimasi wajar untuk harga program.", RP),
    ("TEST_SHARE", "Bagi hasil asesor untuk tes B2C", 0.5, "%", "Pasti", "Sheet ACC Kia: tes 50:50.", PCT),
]
r = 5
for k, p, v, u, s, src, fmt in params:
    put(wi, f"A{r}", k, BOLD)
    put(wi, f"B{r}", p)
    put(wi, f"C{r}", v, BLUE, fmt)
    put(wi, f"D{r}", u)
    put(wi, f"E{r}", s)
    put(wi, f"F{r}", src, align=WRAP)
    name(k, f"Input!$C${r}")
    r += 1

FAM_START = 15
put(wi, f"A{FAM_START-2}", "Aturan per keluarga produk", BOLD, fill=SECTION)
for col in "BCDEF":
    wi[f"{col}{FAM_START-2}"].fill = SECTION
for i, lab in enumerate(["Keluarga", "Kelipatan min (0 = tidak dipakai)", "Margin bersih min", "Catatan"], 1):
    c = wi.cell(row=FAM_START - 1, column=i, value=lab)
    c.font, c.fill = HDR_FONT, HDR_FILL
families = [
    ("A. Pelatihan", "Program pelatih dan program bank."),
    ("B. EAP", "Keputusan 19 Nov 2026: EAP memakai harga hari ini (aturan 'Harga tetap'). Margin bersih tetap dicek."),
    ("C. Asesmen", "Seleksi, pemetaan massal, MHCU. MHCU Flash Rp500rb dan Deep Rp750rb (keputusan 19 Nov 2026)."),
    ("D. Jasa ahli dan psikoedukasi", "Juri, webinar, seminar, event publik."),
    ("E. B2C Konseling", "Aturan 'Pasar': hanya untuk dibaca."),
    ("F. B2C Tes", "Aturan 'Pasar': hanya untuk dibaca."),
    ("G. Coreitera Software", "Theta, Synthesia, biaya transaksi."),
    ("H. Coreitera Asesmen", "Mesin tes white-label dan tes mandiri."),
]
for i, (fam, note) in enumerate(families):
    rr = FAM_START + i
    put(wi, f"A{rr}", fam, BOLD)
    put(wi, f"B{rr}", 0, BLUE, MULT)
    put(wi, f"C{rr}", 0.30, BLUE, PCT)
    put(wi, f"D{rr}", note)
FAM_END = FAM_START + len(families) - 1
name("FAM", f"Input!$A${FAM_START}:$A${FAM_END}")
name("FAM_MULT", f"Input!$B${FAM_START}:$B${FAM_END}")
name("FAM_MARG", f"Input!$C${FAM_START}:$C${FAM_END}")

TIER_START = FAM_END + 4
put(wi, f"A{TIER_START-2}", "Tier psikolog: harga online 60 menit dan bagi hasil", BOLD, fill=SECTION)
for col in "BCDEF":
    wi[f"{col}{TIER_START-2}"].fill = SECTION
for i, lab in enumerate(["Tier", "Pengalaman", "Harga konseling online 60 mnt", "Bagian psikolog", "Fee psikolog per sesi", "Sumber"], 1):
    c = wi.cell(row=TIER_START - 1, column=i, value=lab)
    c.font, c.fill = HDR_FONT, HDR_FILL
tiers = [(1, "0-2 tahun", 500000, 0.60), (2, "2 th 1 bln - 5 tahun", 500000, 0.65),
         (3, "5 th 1 bln - 10 tahun", 725000, 0.65), (4, "Di atas 10 tahun", 850000, 0.70)]
for i, (t, exp, price, share) in enumerate(tiers):
    rr = TIER_START + i
    put(wi, f"A{rr}", t, BOLD)
    put(wi, f"B{rr}", exp)
    put(wi, f"C{rr}", price, BLUE, RP)
    put(wi, f"D{rr}", share, BLUE, PCT)
    put(wi, f"E{rr}", f"=C{rr}*D{rr}", BLACK, RP)
    put(wi, f"F{rr}", "B2C Pricing (sheet Pricing_Counseling, Profit_Share), 13 Okt 2026")
TIER_END = TIER_START + 3
name("TIER_PRICE", f"Input!$C${TIER_START}:$C${TIER_END}")
name("TIER_SHARE", f"Input!$D${TIER_START}:$D${TIER_END}")
name("TIER_FEE", f"Input!$E${TIER_START}:$E${TIER_END}")
wi.freeze_panes = "A5"

# ---------------------------------------------------------------------------------------------
# 3. Tarif_Biaya
# ---------------------------------------------------------------------------------------------
wt = wb.create_sheet("Tarif_Biaya")
put(wt, "A1", "Tarif biaya: satu tarif untuk tiap sumber daya", TITLE)
put(wt, "A2", "Biru = input. Hijau = diambil dari Input. Kuning = estimasi, ganti dengan angka asli. Dipakai oleh Rincian_Biaya.")
header(wt, 4, ["Kunci", "Komponen", "Tarif (Rp)", "Satuan", "Status", "Sumber / catatan"], [16, 52, 16, 20, 12, 70])
# (key, label, value-or-formula, unit, status, source)
rates = [
    ("PSI_T1", "Fee psikolog Tier 1, sesi 60 mnt", "=INDEX(TIER_FEE,1)", "per sesi", "Pasti", "Input: harga online x bagi hasil"),
    ("PSI_T2", "Fee psikolog Tier 2, sesi 60 mnt", "=INDEX(TIER_FEE,2)", "per sesi", "Pasti", "Input"),
    ("PSI_T3", "Fee psikolog Tier 3, sesi 60 mnt", "=INDEX(TIER_FEE,3)", "per sesi", "Pasti", "Input"),
    ("PSI_T4", "Fee psikolog Tier 4, sesi 60 mnt", "=INDEX(TIER_FEE,4)", "per sesi", "Pasti", "Input"),
    ("PSI_EAP", "Fee psikolog sesi B2B 60 mnt (EAP, MHCU), tarif flat", 350000, "per sesi", "Pasti", "Keputusan 19 Nov 2026: flat Rp350rb per sesi untuk semua tier. Sama dengan Lampiran B dan Model BI Sibolga"),
    ("PSI_DOC", "Psikolog: kerja dokumen (rubrik, skoring, laporan)", 110000, "per jam", "Pasti", "Galih 28 Okt 2026, RAB BI Fit and Proper"),
    ("PSI_F2F", "Psikolog: tatap muka asesmen (LGD, wawancara, kalibrasi)", 350000, "per jam", "Pasti", "RAB BI Fit and Proper, BSN"),
    ("ADM_SESI", "Administrasi dan platform per sesi B2B", 35000, "per sesi", "Pasti", "Lampiran B Kemitraan"),
    ("B2C_PLAT", "Platform dan pembayaran per sesi B2C", 50000, "per sesi", "Estimasi", "B2C Pricing, default CFO"),
    ("MHCU_TOOL", "Alat ukur dan skoring MHCU", 7000, "per orang", "Pasti", "Galih 19 Nov 2026: maksimal Rp7rb per tes (Lampiran B lama menulis Rp70rb)"),
    ("AGG_REPORT", "Laporan agregat anonim", 750000, "per laporan", "Pasti", "Lampiran B"),
    ("WEBINAR", "Webinar online oleh psikolog", 1500000, "per sesi", "Pasti", "Lampiran B"),
    ("SEMINAR", "Seminar offline oleh psikolog (Jakarta)", 3000000, "per sesi", "Pasti", "Lampiran B"),
    ("KONTRAK", "Kontrak, meterai, invoice", 150000, "per SPK", "Pasti", "Lampiran B"),
    ("BFI_DISC", "Tes BFI + DISC (platform dan skoring)", 120000, "per peserta", "Pasti", "RAB BI Fit and Proper"),
    ("TRN_HQ", "Biaya pelaksanaan program pelatih kantor pusat, 3 hari", 25000000, "per program", "Estimasi", "Model v15 dan Lampiran B [CEK]"),
    ("TRN_REG", "Biaya pelaksanaan program pelatih regional, 1 hari", 8000000, "per hari", "Estimasi", "Model v15 dan Lampiran B [CEK]"),
    ("TRN_BANK", "Biaya pelaksanaan program bank, 1 kelompok", 12000000, "per hari", "Estimasi", "Model v15 dan Lampiran B [CEK]"),
    ("EVENT_PUB", "Biaya event publik", 8000000, "per event", "Estimasi", "Model v15 dan Lampiran B [CEK]"),
    ("CORE_FEE", "Fee asesmen Coreitera ke TJ (internal grup)", 75000, "per peserta", "Estimasi", "Model v15: Rp50-150rb, dipakai Rp75rb"),
    ("CORE_ADMIN", "Administrasi proyek ke Coreitera (internal grup)", 400000, "per proyek", "Pasti", "RAB BI Jampurnas"),
    ("PM_PROJ", "Project manager per proyek B2B", 2000000, "per proyek", "Pasti", "Model v15 Inputs: pm_fee"),
    ("JURI_FEE", "Fee juri (psikolog atau dosen), termasuk rubrik", 2500000, "per juri", "Pasti", "RAB BI Jampurnas"),
    ("TRANSPORT", "Transport psikolog di Jakarta", 150000, "per orang-hari", "Estimasi", "RAB BI Fit and Proper"),
    ("FOLLOWUP_LIC", "Materi lisensi pelatih (paket Elite)", 1000000, "per program", "Estimasi", "Belum pernah dihitung. Isi angka asli"),
    ("ENG_JAM", "Engineer Coreitera (internal grup)", 150000, "per jam", "Estimasi", "RAB BSN: belum ada tarif resmi"),
    ("ADM_JAM", "Admin / koordinator proyek TJ", 50000, "per jam", "Estimasi", "RAB BSN"),
    ("AI_LAPORAN", "Laporan AI per peserta (Gemini)", 1250, "per laporan", "Pasti", "Galih, RAB BSN"),
    ("AI_SINT", "Biaya AI per laporan Synthesia (draf interpretasi)", 8000, "per laporan", "Estimasi", "Model Coreitera: Rp8.000 per laporan"),
    ("AI_SEAT", "Biaya AI copilot per kursi psikolog", 120000, "per kursi/bulan", "Estimasi", "Model Coreitera: 15 laporan x Rp8.000"),
    ("THETA_SERVER", "Server, storage, email/WA per biro", 50000, "per biro/bulan", "Estimasi", "Model Coreitera: SERVER_PER_BUREAU"),
    ("THETA_AGENT", "Agen AI balas chat + template WA per biro", 100000, "per biro/bulan", "Estimasi", "Model Coreitera: AGENT_COST_PER_BUREAU"),
    ("THETA_ONB", "Onboarding biro baru (dibagi 12 bulan)", 300000, "per biro", "Estimasi", "Model Coreitera: ONBOARDING_PER_NEW_BUREAU"),
    ("TEST_SERVER", "Server dan storage per peserta tes", 1500, "per peserta", "Estimasi", "Estimasi dari RAB BSN (storage + email per peserta)"),
    ("EVENT_KIT", "Activity kit event komunitas", 30000, "per peserta", "Pasti", "RAB Lyfe at 30s"),
    ("EVENT_FNB", "Makan-minum venue (nett)", 50000, "per peserta", "Pasti", "Penawaran Kanteen n Co"),
    ("EVENT_ADS", "Iklan event komunitas (batas)", 300000, "per event", "Pasti", "Final proposal Lyfe at 30s"),
    ("EVENT_PSI", "Psikolog event komunitas", 150000, "per event", "Pasti", "RAB Lyfe at 30s"),
]
for i, (k, lab, v, u, s, src) in enumerate(rates):
    rr = 5 + i
    put(wt, f"A{rr}", k, BOLD)
    put(wt, f"B{rr}", lab)
    is_f = isinstance(v, str)
    put(wt, f"C{rr}", v, GREEN if is_f else BLUE, RP, fill=YELLOW if s == "Estimasi" and not is_f else None)
    put(wt, f"D{rr}", u)
    put(wt, f"E{rr}", s)
    put(wt, f"F{rr}", src, align=WRAP)
RATE_END = 5 + len(rates) - 1
name("RATE_KEY", f"Tarif_Biaya!$A$5:$A${RATE_END}")
name("RATE_VAL", f"Tarif_Biaya!$C$5:$C${RATE_END}")
wt.freeze_panes = "A5"

# ---------------------------------------------------------------------------------------------
# 4. Product data
# ---------------------------------------------------------------------------------------------
# line: (component, qty, unit, rate_key or None, direct_rate or None, internal "Ya"/"Tidak", note)
def L(comp, qty, unit, key=None, direct=None, internal="Tidak", note=""):
    return (comp, qty, unit, key, direct, internal, note)

TJ, CI = "Teman Journey", "Coreitera"
P = []  # (code, unit, family, product, scope, buyer, uom, status, rule, pph, current, source, lines)

def add(*a):
    P.append(a)

# A. Training
add("TJ-PLT-01", TJ, "A. Pelatihan", "Program Pelatih Kantor Pusat (Flagship ToT)",
    "3 hari, hingga 40 peserta, termasuk asesmen pre/post dan laporan dampak", "Regulator (BI, OJK, LPS) kantor pusat",
    "per program", "Aktif", "Penuh", "Ya", 100000000, "Rate_Card v15, Lampiran C",
    [L("Pelaksanaan (fasilitator, materi, perjalanan)", 1, "program", "TRN_HQ"),
     L("Asesmen pre/post Coreitera", 40, "peserta", "CORE_FEE", internal="Ya")])
add("TJ-PLT-02", TJ, "A. Pelatihan", "Program Pelatih Regional (ToT KPW/LPS)",
    "1 hari, hingga 30 peserta, termasuk asesmen dan laporan dampak", "Kantor perwakilan BI, LPS, OJK daerah",
    "per hari", "Aktif", "Penuh", "Ya", 28500000, "Rate_Card v15, Lampiran C",
    [L("Pelaksanaan 1 hari", 1, "hari", "TRN_REG"),
     L("Asesmen pre/post Coreitera", 30, "peserta", "CORE_FEE", internal="Ya")])
add("TJ-PLT-03", TJ, "A. Pelatihan", "Program Bank BUMN (Signature)",
    "1 kelompok, 1 hari, hingga 40 peserta, asesmen pre/post dan laporan dampak", "Bank BUMN (BRI, Mandiri, BNI, BTN)",
    "per hari", "Aktif", "Penuh", "Ya", 51000000, "Rate_Card v15, Lampiran C",
    [L("Pelaksanaan 1 hari", 1, "hari", "TRN_BANK"),
     L("Asesmen pre/post Coreitera", 40, "peserta", "CORE_FEE", internal="Ya")])
add("TJ-PLT-04", TJ, "A. Pelatihan", "Program Bank Swasta (Signature)",
    "1 kelompok, 1 hari, hingga 40 peserta, asesmen pre/post dan laporan dampak", "Bank swasta besar (BCA dan lainnya)",
    "per hari", "Aktif", "Penuh", "Ya", 61000000, "Rate_Card v15, Lampiran C",
    [L("Pelaksanaan 1 hari", 1, "hari", "TRN_BANK"),
     L("Asesmen pre/post Coreitera", 40, "peserta", "CORE_FEE", internal="Ya")])
for code, seg, price in (("TJ-PLT-05", "BUMN", 66300000), ("TJ-PLT-06", "Swasta", 79300000)):
    add(code, TJ, "A. Pelatihan", f"Program Bank Elite {seg}",
        "Signature + lisensi pelatih internal + pendampingan 3 bulan", f"Bank {seg}",
        "per hari", "Aktif", "Penuh", "Ya", price, "Rate_Card v15, Lampiran C",
        [L("Pelaksanaan 1 hari", 1, "hari", "TRN_BANK"),
         L("Asesmen pre/post Coreitera", 40, "peserta", "CORE_FEE", internal="Ya"),
         L("Pendampingan 3 bulan (psikolog tatap muka)", 6, "jam", "PSI_F2F", note="Estimasi: 2 jam per bulan"),
         L("Materi lisensi pelatih", 1, "program", "FOLLOWUP_LIC")])
add("TJ-PLT-07", TJ, "A. Pelatihan", "Asesmen dan laporan dampak, peserta tambahan",
    "Asesmen pre/post dan laporan dampak untuk peserta di luar kuota paket", "Klien program pelatihan",
    "per peserta", "Aktif", "Penuh", "Ya", 150000, "Rate_Card v15, Lampiran C",
    [L("Asesmen Coreitera", 1, "peserta", "CORE_FEE", internal="Ya"),
     L("Bagian laporan dampak (psikolog)", 0.1, "jam", "PSI_DOC")])

# B. EAP
def eap_lines(sessions, mhcu_people=0, rounds=0, webinar=0, seminar=0):
    lines = [L("Fee psikolog sesi konseling online", sessions, "sesi", "PSI_EAP"),
             L("Administrasi dan platform", sessions, "sesi", "ADM_SESI")]
    if rounds:
        lines += [L("Alat ukur MHCU", mhcu_people * rounds, "orang-putaran", "MHCU_TOOL"),
                  L("Laporan agregat anonim", 1, "laporan", "AGG_REPORT")]
    if webinar:
        lines.append(L("Webinar psikoedukasi", webinar, "sesi", "WEBINAR"))
    if seminar:
        lines.append(L("Seminar offline", seminar, "sesi", "SEMINAR"))
    lines.append(L("Kontrak, meterai, invoice", 1, "SPK", "KONTRAK"))
    return lines

add("TJ-EAP-01", TJ, "B. EAP", "EAP Jeda", "15 sesi konseling online, 1 webinar psikoedukasi; sekitar 5 pegawai",
    "HR / unit SDM instansi dan perusahaan", "per 3 bulan", "Aktif", "Harga tetap", "Ya", 12000000,
    "Model Biaya BI Sibolga, Lampiran C", eap_lines(15, webinar=1))
add("TJ-EAP-02", TJ, "B. EAP", "EAP Jeda Tahunan (per kuartal)",
    "Per kuartal: 15 sesi online, MHCU untuk 70 pegawai, laporan agregat, 1 webinar. Kontrak 12 bulan (4 kuartal)",
    "HR / unit SDM instansi dan perusahaan", "per kuartal", "Aktif", "Harga tetap", "Ya", 11000000,
    "Model Biaya BI Sibolga, Lampiran C", eap_lines(15, mhcu_people=70, rounds=1, webinar=1))
add("TJ-EAP-03", TJ, "B. EAP", "EAP Tumbuh", "24 sesi online, 1 putaran MHCU, laporan agregat, 1 webinar; sekitar 8 pegawai",
    "HR / unit SDM instansi dan perusahaan", "per 3 bulan", "Aktif", "Harga tetap", "Ya", 20000000,
    "Model Biaya BI Sibolga, Lampiran C", eap_lines(24, mhcu_people=8, rounds=1, webinar=1))
add("TJ-EAP-04", TJ, "B. EAP", "EAP Seperjalanan",
    "36 sesi online, 2 putaran MHCU (pre/post), hasil individual, laporan perbandingan, 1 seminar offline; sekitar 12 pegawai",
    "HR / unit SDM instansi dan perusahaan", "per 3 bulan", "Aktif", "Harga tetap", "Ya", 31500000,
    "Model Biaya BI Sibolga, Lampiran C", eap_lines(36, mhcu_people=12, rounds=2, seminar=1))
add("TJ-EAP-05", TJ, "B. EAP", "Sesi konseling tambahan EAP", "Konseling online 60 menit di luar kuota paket",
    "Klien EAP", "per sesi", "Aktif", "Harga tetap", "Ya", 500000, "Lampiran C, Model BI Sibolga",
    [L("Fee psikolog", 1, "sesi", "PSI_EAP"), L("Administrasi dan platform", 1, "sesi", "ADM_SESI")])

# C. Assessment
add("TJ-ASM-01", TJ, "C. Asesmen", "Asesmen Seleksi / Fit and Proper (6 peserta)",
    "Penyesuaian kompetensi dan rubrik, tes BFI + DISC, LGD, wawancara psikolog, laporan komprehensif per peserta",
    "Regulator, bank, BUMN: seleksi pejabat", "per proyek 6 peserta", "Aktif", "Penuh", "Ya", 15000000,
    "RAB BI Fit and Proper (Rp2,5 jt per peserta)",
    [L("Psikolog kerja dokumen: rubrik, studi kasus, skoring, laporan, integrasi", 34, "jam", "PSI_DOC"),
     L("Psikolog tatap muka: LGD (2 x 2 jam) dan wawancara (6 x 1 jam)", 10, "jam", "PSI_F2F"),
     L("Tes BFI + DISC", 6, "peserta", "BFI_DISC"),
     L("Admin dan koordinasi jadwal", 6, "peserta", None, 35000),
     L("Transport psikolog", 3, "orang-hari", "TRANSPORT"),
     L("Kontrak dan meterai", 1, "SPK", "KONTRAK")])

def bsn_base():
    return [L("Storage hasil dan laporan 12 bulan", 1, "proyek", None, 17789),
            L("Email undangan dan pengingat", 1, "paket", None, 500000),
            L("Setup akun, paket, branding, uji coba", 16, "jam", "ENG_JAM", internal="Ya"),
            L("Helpdesk peserta", 25, "jam", "ADM_JAM"),
            L("Laporan AI per karyawan (termasuk ulang 3%)", 2060, "laporan", "AI_LAPORAN"),
            L("Kontrak dan meterai", 1, "SPK", "KONTRAK")]

add("TJ-ASM-02", TJ, "C. Asesmen", "Profiling Massal Paket 1: Laporan Perusahaan",
    "Tes BFI + DISC online, laporan otomatis per karyawan (dibaca perusahaan), skor di dashboard HC dan Excel; basis 2.000 karyawan, jendela 7 hari",
    "HC perusahaan dan bank besar", "per proyek 2.000 peserta", "Aktif", "Penuh", "Ya", 50000000,
    "RAB BSN (Rp25rb per peserta)", bsn_base())
add("TJ-ASM-03", TJ, "C. Asesmen", "Profiling Massal Paket 2: Laporan Perusahaan dan Karyawan",
    "Paket 1 + karyawan membaca laporannya sendiri; basis 2.000 karyawan", "HC perusahaan dan bank besar",
    "per proyek 2.000 peserta", "Aktif", "Penuh", "Ya", 70000000, "RAB BSN (Rp35rb per peserta)", bsn_base())
for code, lvl, price, extra, desc in (
        ("TJ-ASM-04", "3A Pemetaan Umum", 101500000, 12350000, "Paket 2 + analisis profil 10 unit, laporan pemetaan, presentasi psikolog"),
        ("TJ-ASM-05", "3B Pemetaan + Talent Box", 139500000, 27150000, "3A + talent box per unit dengan data kinerja, kalibrasi bersama HC"),
        ("TJ-ASM-06", "3C Pemetaan Talenta Lengkap", 210500000, 55000000, "3B + analisis 15 rumpun jabatan dan skor kecocokan jabatan")):
    add(code, TJ, "C. Asesmen", f"Pemetaan Karyawan {lvl}", desc + "; basis 2.000 karyawan",
        "Direksi dan HC perusahaan", "per proyek 2.000 peserta", "Aktif", "Penuh", "Ya", price,
        "RAB BSN, sheet Opsi dan Biaya",
        bsn_base() + [L(f"Pekerjaan psikolog dan tim pemetaan level {lvl[:2]} (lihat RAB BSN)", 1, "proyek", None, extra)])
add("TJ-ASM-07", TJ, "C. Asesmen", "Tambahan pelaksanaan serentak",
    "Upgrade server, load test, engineer standby di hari tes serentak (2 gelombang)", "Klien profiling massal",
    "per proyek", "Aktif", "Penuh", "Ya", 30000000, "RAB BSN (+Rp15rb per peserta)",
    [L("Upgrade server sementara 1 bulan", 1, "bulan", None, 2145000),
     L("Load test dan tuning", 16, "jam", "ENG_JAM", internal="Ya"),
     L("Engineer standby hari H", 32, "jam", "ENG_JAM", internal="Ya")])
add("TJ-ASM-08", TJ, "C. Asesmen", "MHCU Flash", "5 instrumen klinis (DASS-21, CBI, BFI-10, SWLS, My Story), review psikolog, laporan personal, konseling 30 menit",
    "Komunitas, kampus, perusahaan", "per peserta", "Aktif", "Penuh", "Ya", 242000, "Proposal MHCU LPDP (Telering)",
    [L("Alat ukur MHCU", 1, "orang", "MHCU_TOOL"),
     L("Konseling 30 menit", 0.5, "sesi", "PSI_EAP"),
     L("Review manual dan laporan", 0.25, "jam", "PSI_DOC"),
     L("Administrasi dan platform", 1, "sesi", "ADM_SESI")])
add("TJ-ASM-09", TJ, "C. Asesmen", "MHCU Deep", "Sama dengan Flash, konseling 60 menit",
    "Komunitas, kampus, perusahaan", "per peserta", "Aktif", "Penuh", "Ya", 380000, "Proposal MHCU LPDP (Telering)",
    [L("Alat ukur MHCU", 1, "orang", "MHCU_TOOL"),
     L("Konseling 60 menit", 1, "sesi", "PSI_EAP"),
     L("Review manual dan laporan", 0.25, "jam", "PSI_DOC"),
     L("Administrasi dan platform", 1, "sesi", "ADM_SESI")])
add("TJ-ASM-10", TJ, "C. Asesmen", "Hasil MHCU individual untuk pegawai (add-on EAP)",
    "Laporan MHCU personal untuk tiap pegawai, di luar laporan agregat", "Klien EAP",
    "per pegawai", "Aktif", "Penuh", "Ya", 70000, "Model BI Sibolga",
    [L("Alat ukur MHCU", 1, "orang", "MHCU_TOOL"),
     L("Review dan laporan individual", 0.25, "jam", "PSI_DOC")])

# D. Expert services and psychoeducation
add("TJ-EDU-01", TJ, "D. Jasa ahli dan psikoedukasi", "Juri Ahli Kompetisi",
    "1 juri psikolog + 1 juri ahli bidang + project manager; rubrik dan penilaian hingga 46 karya, hadir di hari acara",
    "Panitia lomba instansi (BI, kementerian, kampus)", "per acara", "Aktif", "Penuh", "Ya", 15000000, "RAB BI Jampurnas",
    [L("Fee juri", 2, "juri", "JURI_FEE"), L("Project manager", 1, "proyek", "PM_PROJ"),
     L("Administrasi Coreitera", 1, "proyek", "CORE_ADMIN", internal="Ya")])
add("TJ-EDU-02", TJ, "D. Jasa ahli dan psikoedukasi", "Webinar Psikoedukasi", "Webinar online 60-90 menit oleh psikolog, materi sesuai kebutuhan",
    "HR, komunitas, kampus", "per sesi", "Aktif", "Penuh", "Ya", None, "Lampiran B (biaya). Belum pernah dijual terpisah",
    [L("Webinar oleh psikolog", 1, "sesi", "WEBINAR"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-EDU-03", TJ, "D. Jasa ahli dan psikoedukasi", "Seminar Psikoedukasi Offline", "Seminar tatap muka di Jakarta oleh psikolog",
    "HR, komunitas, kampus", "per sesi", "Aktif", "Penuh", "Ya", None, "Lampiran B (biaya). Belum pernah dijual terpisah",
    [L("Seminar oleh psikolog", 1, "sesi", "SEMINAR"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-EDU-04", TJ, "D. Jasa ahli dan psikoedukasi", "Event Psikoedukasi Publik", "Event untuk umum, dikelola TJ (venue, materi, fasilitator)",
    "Instansi sponsor, CSR", "per event", "Aktif", "Penuh", "Ya", 24000000, "Lampiran C, Model v15 (Rp20 jt di model lama)",
    [L("Biaya event", 1, "event", "EVENT_PUB"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-EDU-05", TJ, "D. Jasa ahli dan psikoedukasi", "Event Komunitas Bertiket (per peserta)",
    "Sesi cerita bersama psikolog, activity kit, makan-minum; basis 12 peserta", "Individu, lewat mitra komunitas",
    "per peserta", "Aktif", "Pasar", "Tidak", 259000, "Final proposal Lyfe at 30s x Kanteen n Co",
    [L("Activity kit", 1, "peserta", "EVENT_KIT"), L("Makan-minum", 1, "peserta", "EVENT_FNB"),
     L("Psikolog (dibagi 12 peserta)", 1 / 12, "event", "EVENT_PSI"), L("Iklan (dibagi 12 peserta)", 1 / 12, "event", "EVENT_ADS")])

# E. B2C counseling: price per tier from B2C Pricing catalog
b2c = [
    ("KON-ON", "Konseling Individual Online 60 menit", "Konseling 1:1 via video call", [500000, 500000, 725000, 850000]),
    ("KON-OFF", "Konseling Individual Offline 60 menit", "Konseling 1:1 tatap muka di kantor TJ", [580000, 580000, 800000, 925000]),
    ("PSG", "Konseling Pasangan Offline 90 menit", "Konseling pasangan tatap muka", [800000, 950000, 1100000, 1425000]),
    ("PMC", "Pre-Marriage Check Up (3 sesi)", "3 sesi pasangan 75-90 menit, diskon paket 10%", [2200000, 2600000, 2900000, 3800000]),
    ("MC", "Marriage Check Up (2 sesi)", "2 sesi pasangan 75-90 menit (harga estimasi)", [1400000, 1700000, 2000000, 2600000]),
    ("TRP-OFF", "Terapi Klinis Offline (EMDR / Hipnoterapi)", "60-90 menit, hanya psikolog bersertifikat", [750000, 900000, 1000000, 1325000]),
    ("TRP-ON", "Terapi Klinis Online (EMDR / Hipnoterapi)", "60-90 menit, hanya psikolog bersertifikat", [675000, 800000, 900000, 1200000]),
]
SESS = {"PMC": 3, "MC": 2}
for key, prod, scope, prices in b2c:
    for t in range(4):
        code = f"TJ-B2C-{key}-T{t+1}"
        add(code, TJ, "E. B2C Konseling", f"{prod}, Tier {t+1}", scope, "Individu dan pasangan", "per sesi" if key not in SESS else "per paket",
            "Aktif", "Pasar", "Tidak", prices[t], "B2C Pricing, sheet Catalog (13 Okt 2026)",
            [L(f"Fee psikolog Tier {t+1} (bagi hasil)", f"=INDEX(TIER_SHARE,{t+1})", "% harga", None, "PRICE"),
             L("Platform dan pembayaran", SESS.get(key, 1), "sesi", "B2C_PLAT")])

# F. B2C tests
tests = [("STK", "Skrining Tumbuh Kembang (0-5 th)", 625000, 0), ("IQ", "Tes IQ (estimasi harga)", 725000, 0),
         ("KS", "Tes Kesiapan Sekolah (4-6 th)", 800000, 0), ("MB", "Tes Minat Bakat (14 th ke atas)", 1625000, 0),
         ("PRF", "Tes Profiling (6-13 th)", 1225000, 0), ("MHCU", "Tes Kesehatan Mental Dewasa (MHCU)", 375000, 1)]
for key, prod, price, tool in tests:
    lines = [L("Fee asesor (bagi hasil)", "=TEST_SHARE", "% harga", None, "PRICE")]
    if tool:
        lines.append(L("Alat ukur MHCU", 1, "orang", "MHCU_TOOL"))
    add(f"TJ-TES-{key}", TJ, "F. B2C Tes", prod, "Tes dan laporan tertulis oleh psikolog", "Orang tua, pelajar, individu",
        "per tes", "Aktif", "Pasar", "Tidak", price, "B2C Pricing, sheet Pricing_Assessment", lines)

# G. Coreitera software
for code, plan, size, price in (("CI-THT-01", "Starter", "1-3 psikolog", 1000000), ("CI-THT-02", "Growth", "4-10 psikolog", 2900000),
                                ("CI-THT-03", "Pro", "11 psikolog ke atas", 5730000)):
    add(code, CI, "G. Coreitera Software", f"Theta Paket {plan}",
        f"Booking, pembayaran, jadwal, catatan klien, agen AI 24 jam untuk biro {size}", "Pemilik biro psikologi",
        "per biro per bulan", "Live di TJ 30 Nov 2026", "Penuh", "Tidak", price, "Model Coreitera inputs.py (Galih 19 Okt 2026)",
        [L("Server, storage, email/WA", 1, "bulan", "THETA_SERVER"), L("Agen AI dan template WA", 1, "bulan", "THETA_AGENT"),
         L("Onboarding dibagi 12 bulan", 1 / 12, "biro", "THETA_ONB")])
add("CI-THT-04", CI, "G. Coreitera Software", "Theta Biaya Transaksi", "Per transaksi berbayar di Theta; biaya Midtrans diteruskan sebesar biaya",
    "Pemilik biro psikologi", "per transaksi", "Rencana", "Pasar", "Tidak", 5000, "Model Coreitera inputs.py",
    [L("Tidak ada biaya langsung (Midtrans ditagih terpisah)", 0, "transaksi", None, 0)])
add("CI-SYN-01", CI, "G. Coreitera Software", "Synthesia Copilot (kursi psikolog)",
    "AI menyusun draf interpretasi sesuai model kompetensi; psikolog review, edit, dan tanda tangan", "Biro psikologi korporat",
    "per kursi per bulan", "Rencana 2028", "Penuh", "Tidak", 300000, "Model Coreitera inputs.py (COPILOT_SEAT)",
    [L("Biaya AI per kursi", 1, "kursi", "AI_SEAT")])
add("CI-SYN-02", CI, "G. Coreitera Software", "Synthesia Fee Laporan Bertanda Tangan",
    "Fee per laporan yang ditandatangani psikolog lewat Synthesia (harga belum ditetapkan)", "Biro psikologi korporat",
    "per laporan", "Rencana", "Penuh", "Tidak", None, "Proposal Synthesia v0.2: fee per laporan, harga belum ditetapkan",
    [L("Biaya AI per laporan", 1, "laporan", "AI_SINT")])

# H. Coreitera assessment
add("CI-ASM-01", CI, "H. Coreitera Asesmen", "Mesin Tes White-Label", "Tes online dengan merek biro, skor otomatis; dijual sekitar biaya",
    "Biro psikologi", "per peserta", "Rencana", "Harga pokok", "Tidak", None, "Proposal Synthesia v0.2: test engine at about cost",
    [L("Laporan AI", 1, "laporan", "AI_LAPORAN"), L("Server dan storage", 1, "peserta", "TEST_SERVER")])
add("CI-ASM-02", CI, "H. Coreitera Asesmen", "Tes Mandiri + Upgrade 'Ditandatangani Psikolog'",
    "Tes mandiri untuk HR dan UMKM dengan harga setara pasar; tombol upgrade ke interpretasi psikolog biro mitra",
    "HR perusahaan, UMKM", "per peserta", "Rencana (fase terakhir)", "Pasar", "Tidak", 25000,
    "Proposal Synthesia v0.2 (pembanding Scalify Rp15rb, Rekrutiva Rp45rb). Harga estimasi",
    [L("Laporan AI", 1, "laporan", "AI_LAPORAN"), L("Server dan storage", 1, "peserta", "TEST_SERVER")])

# ---------------------------------------------------------------------------------------------
# 5. Rincian_Biaya
# ---------------------------------------------------------------------------------------------
wr = wb.create_sheet("Rincian_Biaya")
put(wr, "A1", "Rincian biaya per produk: jumlah x tarif", TITLE)
put(wr, "A2", "Kunci tarif kosong = tarif diisi langsung di kolom F (biru). Internal grup 'Ya' = uang tetap di grup; tidak dihitung di kolom kelipatan atas biaya luar grup.")
header(wr, 4, ["Kode produk", "Komponen biaya", "Jumlah", "Satuan", "Kunci tarif", "Tarif langsung (Rp)",
               "Tarif dipakai (Rp)", "Total (Rp)", "Internal grup?", "Catatan"], [18, 58, 10, 14, 14, 16, 16, 16, 12, 40])
code_row = {}  # filled after katalog layout is known
KAT_FIRST = 6
for i, p in enumerate(P):
    code_row[p[0]] = KAT_FIRST + i
r = 5
for p in P:
    code = p[0]
    for comp, qty, unit, key, direct, internal, note in p[12]:
        put(wr, f"A{r}", code)
        put(wr, f"B{r}", comp)
        put(wr, f"C{r}", qty, BLACK if isinstance(qty, str) else BLUE, "#,##0.00")
        put(wr, f"D{r}", unit)
        put(wr, f"E{r}", key)
        if direct == "PRICE":
            put(wr, f"F{r}", f"=Katalog!$K${code_row[code]}", GREEN, RP)
        elif key is None:
            put(wr, f"F{r}", direct, BLUE, RP)
        put(wr, f"G{r}", f'=IF(E{r}="",F{r},INDEX(RATE_VAL,MATCH(E{r},RATE_KEY,0)))', BLACK, RP)
        put(wr, f"H{r}", f"=C{r}*G{r}", BLACK, RP)
        put(wr, f"I{r}", internal, BLUE)
        put(wr, f"J{r}", note)
        r += 1
BOM_END = r - 1
wr.freeze_panes = "C5"
wr.auto_filter.ref = f"A4:J{BOM_END}"
dv = DataValidation(type="list", formula1='"Ya,Tidak"', allow_blank=False)
wr.add_data_validation(dv)
dv.add(f"I5:I{BOM_END}")

# ---------------------------------------------------------------------------------------------
# 6. Katalog
# ---------------------------------------------------------------------------------------------
wk = wb.create_sheet("Katalog", 1)
put(wk, "A1", "Katalog produk dan harga: Teman Journey + Coreitera", TITLE)
put(wk, "A2", "Harga daftar baru = harga yang dipakai sales. Harga saat ini tetap tampil sebagai pembanding. "
              "Rupiah, belum termasuk PPN. Kolom biru boleh diubah; kolom lain rumus.")
put(wk, "A3", "Cek aturan: Lolos = margin bersih memenuhi batas keluarga di sheet Input. 'Harga tetap' (EAP) hanya dicek marginnya. Info = produk 'Pasar' atau 'Harga pokok'.")
cols = ["Kode", "Unit bisnis", "Keluarga", "Produk", "Isi / cakupan", "Pembeli utama", "Satuan", "Status produk",
        "Aturan harga", "Potong PPh 23?", "Harga saat ini (Rp)", "Biaya total (Rp)", "Biaya luar grup (Rp)",
        "Kelipatan min (opsional)", "Margin bersih min", "Harga min: kelipatan (Rp)", "Harga min: margin (Rp)", "Harga minimum aturan (Rp)",
        "Pembulatan (Rp)", "HARGA DAFTAR BARU (Rp)", "Perubahan vs saat ini", "Laba bersih per satuan (Rp)",
        "Margin bersih", "Kelipatan atas biaya luar", "Cek aturan", "Sumber", "Harga ditetapkan (Rp, opsional)"]
widths = [20, 14, 22, 40, 50, 30, 18, 16, 12, 10, 16, 15, 15, 9, 9, 15, 15, 15, 11, 17, 11, 15, 10, 10, 11, 42, 16]
OVERRIDE = {"TJ-ASM-08": 500000, "TJ-ASM-09": 750000}
header(wk, 5, cols, widths)
wk.row_dimensions[5].height = 45
B = "Rincian_Biaya!$A$5:$A$%d" % BOM_END
H = "Rincian_Biaya!$H$5:$H$%d" % BOM_END
I_ = "Rincian_Biaya!$I$5:$I$%d" % BOM_END
for i, p in enumerate(P):
    rr = KAT_FIRST + i
    code, unit, fam, prod, scope, buyer, uom, status, rule, pph, cur, src, _ = p
    for col, v in zip("ABCDEFGH", (code, unit, fam, prod, scope, buyer, uom, status)):
        put(wk, f"{col}{rr}", v, BOLD if col == "A" else BLACK, align=WRAP if col in "DEF" else None)
    put(wk, f"I{rr}", rule, BLUE)
    put(wk, f"J{rr}", pph, BLUE)
    put(wk, f"K{rr}", cur, BLUE, RP)
    tp = f'IF(J{rr}="Ya",TARIF_PPH23,0)'
    put(wk, f"L{rr}", f"=SUMIFS({H},{B},A{rr})", GREEN, RP)
    put(wk, f"M{rr}", f'=SUMIFS({H},{B},A{rr},{I_},"Tidak")', GREEN, RP)
    put(wk, f"N{rr}", f"=INDEX(FAM_MULT,MATCH(C{rr},FAM,0))", GREEN, MULT)
    put(wk, f"O{rr}", f"=INDEX(FAM_MARG,MATCH(C{rr},FAM,0))", GREEN, PCT)
    put(wk, f"P{rr}", f'=IF(OR(I{rr}="Penuh",I{rr}="Harga tetap"),M{rr}*N{rr},"")', BLACK, RP)
    put(wk, f"Q{rr}", f'=IF(OR(I{rr}="Penuh",I{rr}="Harga tetap"),L{rr}*(1-KOMISI)/((1-{tp})*(1-KOMISI)-O{rr}),"")', BLACK, RP)
    put(wk, f"R{rr}", f'=IF(I{rr}="Penuh",MAX(P{rr},Q{rr}),IF(I{rr}="Harga tetap",Q{rr},IF(I{rr}="Harga pokok",L{rr}/(1-{tp}),"")))', BLACK, RP)
    put(wk, f"S{rr}", f'=IF(R{rr}="","",IF(R{rr}<100000,STEP0,IF(R{rr}<1000000,STEP1,IF(R{rr}<10000000,STEP2,STEP3))))', BLACK, RP)
    put(wk, f"AA{rr}", OVERRIDE.get(code), BLUE, RP)
    put(wk, f"T{rr}", f'=IF(AA{rr}>0,AA{rr},IF(I{rr}="Penuh",MAX(K{rr},ROUNDUP(R{rr}/S{rr},0)*S{rr}),'
                      f'IF(I{rr}="Harga pokok",IF(K{rr}>0,K{rr},ROUNDUP(R{rr}/S{rr},0)*S{rr}),K{rr})))',
        Font(name=F, bold=True), RP, fill=PatternFill("solid", fgColor="E2EFDA"))
    put(wk, f"U{rr}", f'=IF(K{rr}>0,T{rr}/K{rr}-1,"Baru")', BLACK, PCT)
    gp = f"(T{rr}*(1-{tp})-L{rr})"
    put(wk, f"V{rr}", f'={gp}-IF(OR(I{rr}="Penuh",I{rr}="Harga tetap"),MAX(0,{gp}*KOMISI),0)', BLACK, RP)
    put(wk, f"W{rr}", f'=IF(T{rr}>0,V{rr}/T{rr},"")', BLACK, PCT)
    put(wk, f"X{rr}", f'=IF(M{rr}>0,T{rr}/M{rr},"")', BLACK, MULT)
    put(wk, f"Y{rr}", f'=IF(I{rr}="Harga tetap",IF(W{rr}>=O{rr}-0.0001,"Lolos (margin)","Di bawah margin"),IF(I{rr}<>"Penuh","Info",IF(AND(X{rr}>=N{rr}-0.001,W{rr}>=O{rr}-0.0001),"Lolos","Tidak lolos")))')
    put(wk, f"Z{rr}", src, align=WRAP)
KAT_LAST = KAT_FIRST + len(P) - 1
wk.freeze_panes = "E6"
wk.auto_filter.ref = f"A5:AA{KAT_LAST}"
green = PatternFill("solid", fgColor="C6EFCE")
red = PatternFill("solid", fgColor="FFC7CE")
wk.conditional_formatting.add(f"Y{KAT_FIRST}:Y{KAT_LAST}", CellIsRule(operator="equal", formula=['"Lolos"'], fill=green))
wk.conditional_formatting.add(f"Y{KAT_FIRST}:Y{KAT_LAST}", CellIsRule(operator="equal", formula=['"Tidak lolos"'], fill=red))
wk.conditional_formatting.add(f"Y{KAT_FIRST}:Y{KAT_LAST}", CellIsRule(operator="equal", formula=['"Lolos (margin)"'], fill=green))
wk.conditional_formatting.add(f"Y{KAT_FIRST}:Y{KAT_LAST}", CellIsRule(operator="equal", formula=['"Di bawah margin"'], fill=red))
wk.conditional_formatting.add(f"U{KAT_FIRST}:U{KAT_LAST}", FormulaRule(formula=[f'AND(ISNUMBER(U{KAT_FIRST}),U{KAT_FIRST}>0)'],
                                                                      fill=PatternFill("solid", fgColor="FCE4D6")))
for f1, opts in (("I", '"Penuh,Harga tetap,Pasar,Harga pokok"'), ("J", '"Ya,Tidak"')):
    d = DataValidation(type="list", formula1=opts)
    wk.add_data_validation(d)
    d.add(f"{f1}{KAT_FIRST}:{f1}{KAT_LAST}")
name("KAT_CODE", f"Katalog!$A${KAT_FIRST}:$A${KAT_LAST}")
name("KAT_PROD", f"Katalog!$D${KAT_FIRST}:$D${KAT_LAST}")
name("KAT_UOM", f"Katalog!$G${KAT_FIRST}:$G${KAT_LAST}")
name("KAT_RULE", f"Katalog!$I${KAT_FIRST}:$I${KAT_LAST}")
name("KAT_PPH", f"Katalog!$J${KAT_FIRST}:$J${KAT_LAST}")
name("KAT_COST", f"Katalog!$L${KAT_FIRST}:$L${KAT_LAST}")
name("KAT_EXT", f"Katalog!$M${KAT_FIRST}:$M${KAT_LAST}")
name("KAT_PRICE", f"Katalog!$T${KAT_FIRST}:$T${KAT_LAST}")
name("KAT_FAM", f"Katalog!$C${KAT_FIRST}:$C${KAT_LAST}")
name("KAT_CUR", f"Katalog!$K${KAT_FIRST}:$K${KAT_LAST}")
name("KAT_CHG", f"Katalog!$U${KAT_FIRST}:$U${KAT_LAST}")
name("KAT_CHECK", f"Katalog!$Y${KAT_FIRST}:$Y${KAT_LAST}")
name("KAT_MARGIN", f"Katalog!$W${KAT_FIRST}:$W${KAT_LAST}")

# ---------------------------------------------------------------------------------------------
# 7. Ringkasan
# ---------------------------------------------------------------------------------------------
ws = wb.create_sheet("Ringkasan", 1)
put(ws, "A1", "Ringkasan: produk, aturan, dan dampak harga baru", TITLE)
put(ws, "A2", "Semua angka rumus dari Katalog. Ubah batas margin per keluarga di Input untuk melihat dampaknya.")
header(ws, 4, ["Keluarga", "Jumlah produk", "Produk dengan aturan (Penuh / Harga tetap)", "Lolos dengan harga saat ini", "Naik harga", "Produk baru (belum ada harga)",
               "Kenaikan rata-rata (yang naik)", "Margin bersih rata-rata, harga baru", "Kelipatan min (opsional)", "Margin min", "Di bawah margin (harga tetap)"],
       [32, 12, 14, 14, 12, 14, 14, 16, 11, 11, 14])
for i, (fam, _) in enumerate(families):
    rr = 5 + i
    put(ws, f"A{rr}", fam, BOLD)
    put(ws, f"B{rr}", f"=COUNTIFS(KAT_FAM,A{rr})", GREEN)
    put(ws, f"C{rr}", f'=COUNTIFS(KAT_FAM,A{rr},KAT_RULE,"Penuh")+COUNTIFS(KAT_FAM,A{rr},KAT_RULE,"Harga tetap")', GREEN)
    put(ws, f"D{rr}", f'=COUNTIFS(KAT_FAM,A{rr},KAT_CHECK,"Lolos*",KAT_CHG,"<=0.0001")', GREEN)
    put(ws, f"E{rr}", f'=COUNTIFS(KAT_FAM,A{rr},KAT_CHG,">0.0001")', GREEN)
    put(ws, f"F{rr}", f'=COUNTIFS(KAT_FAM,A{rr},KAT_CHG,"Baru")', GREEN)
    put(ws, f"G{rr}", f'=IF(E{rr}=0,0,SUMIFS(KAT_CHG,KAT_FAM,A{rr},KAT_CHG,">0.0001")/E{rr})', GREEN, PCT)
    put(ws, f"H{rr}", f"=IFERROR(AVERAGEIFS(KAT_MARGIN,KAT_FAM,A{rr}),0)", GREEN, PCT)
    put(ws, f"I{rr}", f"=INDEX(FAM_MULT,MATCH(A{rr},FAM,0))", GREEN, MULT)
    put(ws, f"J{rr}", f"=INDEX(FAM_MARG,MATCH(A{rr},FAM,0))", GREEN, PCT)
    put(ws, f"K{rr}", f'=COUNTIFS(KAT_FAM,A{rr},KAT_CHECK,"Di bawah margin")', GREEN)
tot = 5 + len(families)
put(ws, f"A{tot}", "Total", BOLD)
for col in "BCDEFK":
    put(ws, f"{col}{tot}", f"=SUM({col}5:{col}{tot-1})", BOLD)
put(ws, f"A{tot+2}", "Cara baca: 'Naik harga' = produk yang harga saat ininya di bawah aturan dan dinaikkan di Katalog. "
                     "Jika kenaikan terlalu besar untuk pasar, ubah produk itu ke aturan 'Harga tetap' di Katalog "
                     "(harga hari ini dipertahankan, margin tetap dicek).", align=WRAP)
ws.merge_cells(f"A{tot+2}:J{tot+2}")
ws.row_dimensions[tot + 2].height = 45

# ---------------------------------------------------------------------------------------------
# 8. Simulasi_Penawaran
# ---------------------------------------------------------------------------------------------
wq = wb.create_sheet("Simulasi_Penawaran")
put(wq, "A1", "Simulasi penawaran", TITLE)
put(wq, "A2", "Pilih kode produk (daftar dari Katalog), isi jumlah dan diskon (biru). Contoh di baris pertama: 1 paket EAP Jeda.")
put(wq, "A3", "Klien:", BOLD)
put(wq, "B3", "(isi nama klien)", BLUE)
header(wq, 5, ["No", "Kode produk", "Produk", "Satuan", "Jumlah", "Harga daftar (Rp)", "Diskon", "Harga setelah diskon (Rp)",
               "Total harga (Rp)", "Total biaya (Rp)", "Biaya luar grup (Rp)", "Potong PPh 23?", "Aturan",
               "Harga minimum kelipatan, opsional (Rp)"],
       [5, 20, 44, 18, 9, 16, 9, 17, 17, 16, 16, 10, 11, 18])
Q1, QN = 6, 20
dvq = DataValidation(type="list", formula1=f"=Katalog!$A${KAT_FIRST}:$A${KAT_LAST}", allow_blank=True)
wq.add_data_validation(dvq)
dvq.add(f"B{Q1}:B{QN}")
for rr in range(Q1, QN + 1):
    put(wq, f"A{rr}", rr - Q1 + 1)
    put(wq, f"B{rr}", "TJ-EAP-01" if rr == Q1 else None, BLUE)
    put(wq, f"E{rr}", 1 if rr == Q1 else None, BLUE, "#,##0")
    put(wq, f"G{rr}", 0 if rr == Q1 else None, BLUE, PCT)
    look = lambda rng: f'=IF($B{rr}="","",INDEX({rng},MATCH($B{rr},KAT_CODE,0)))'
    put(wq, f"C{rr}", look("KAT_PROD"), GREEN)
    put(wq, f"D{rr}", look("KAT_UOM"), GREEN)
    put(wq, f"F{rr}", look("KAT_PRICE"), GREEN, RP)
    put(wq, f"H{rr}", f'=IF($B{rr}="",0,F{rr}*(1-G{rr}))', BLACK, RP)
    put(wq, f"I{rr}", f'=IF($B{rr}="",0,H{rr}*E{rr})', BLACK, RP)
    put(wq, f"J{rr}", f'=IF($B{rr}="",0,INDEX(KAT_COST,MATCH($B{rr},KAT_CODE,0))*E{rr})', GREEN, RP)
    put(wq, f"K{rr}", f'=IF($B{rr}="",0,INDEX(KAT_EXT,MATCH($B{rr},KAT_CODE,0))*E{rr})', GREEN, RP)
    put(wq, f"L{rr}", look("KAT_PPH"), GREEN)
    put(wq, f"M{rr}", look("KAT_RULE"), GREEN)
    put(wq, f"N{rr}", f'=IF(OR($B{rr}="",M{rr}<>"Penuh"),0,K{rr}*INDEX(FAM_MULT,MATCH(INDEX(KAT_FAM,MATCH($B{rr},KAT_CODE,0)),FAM,0)))', BLACK, RP)
s = QN + 2
summary = [
    ("Total harga ke klien (belum PPN)", f"=SUM(I{Q1}:I{QN})", RP),
    ("PPh 23 dipotong klien", f'=SUMPRODUCT((L{Q1}:L{QN}="Ya")*I{Q1}:I{QN})*TARIF_PPH23', RP),
    ("Total biaya", f"=SUM(J{Q1}:J{QN})", RP),
    ("Laba kotor", f"=C{s}-C{s+1}-C{s+2}", RP),
    ("Komisi (dari laba kotor, produk 'Penuh')", f'=MAX(0,SUMPRODUCT(((M{Q1}:M{QN}="Penuh")+(M{Q1}:M{QN}="Harga tetap"))*(I{Q1}:I{QN}*(1-TARIF_PPH23*(L{Q1}:L{QN}="Ya"))-J{Q1}:J{QN})))*KOMISI', RP),
    ("Laba bersih", f"=C{s+3}-C{s+4}", RP),
    ("Margin bersih", f"=IF(C{s}>0,C{s+5}/C{s},0)", PCT),
    ("Kelipatan atas biaya luar grup", f"=IF(SUM(K{Q1}:K{QN})>0,C{s}/SUM(K{Q1}:K{QN}),0)", MULT),
    ("Diskon terbesar", f"=MAX(G{Q1}:G{QN})", PCT),
    ("Lolos aturan margin bersih (minimal 30%)?", f'=IF(C{s+6}>=0.3,"Ya","Tidak: naikkan harga atau kurangi diskon")', None),
    ("Lolos batas kelipatan (opsional, 0 = tidak dipakai)?", f'=IF(C{s}>=SUM(N{Q1}:N{QN})-1,"Ya","Tidak: perlu persetujuan Direktur")', None),
    ("Persetujuan yang dibutuhkan", f'=IF(OR(C{s+9}<>"Ya",C{s+10}<>"Ya",C{s+8}>0.2),"Direktur",IF(C{s+8}>0.1,"Founder / CEO",IF(C{s+8}>0,"Lead sales","Tidak perlu")))', None),
]
for i, (lab, f, fmt) in enumerate(summary):
    put(wq, f"B{s+i}", lab, BOLD)
    put(wq, f"C{s+i}", f, Font(name=F, bold=True), fmt)
wq.freeze_panes = "A6"

# ---------------------------------------------------------------------------------------------
# 9. Aturan_Penjualan
# ---------------------------------------------------------------------------------------------
wa = wb.create_sheet("Aturan_Penjualan")
wa.column_dimensions["A"].width = 6
wa.column_dimensions["B"].width = 34
wa.column_dimensions["C"].width = 100
put(wa, "A1", "Aturan penjualan untuk tim sales dan mitra", TITLE)
header(wa, 3, ["No", "Aturan", "Isi"])
rules = [
    ("Harga daftar", "Buka setiap penawaran dengan Harga Daftar Baru di Katalog. Tidak pernah di bawahnya tanpa persetujuan."),
    ("Satuan harga", "Pelatihan per kelompok per hari; asesmen per peserta; EAP per paket atau per kuartal; software per biro atau per kursi per bulan. Tidak pernah satu angka borongan tanpa rincian."),
    ("Diskon pilot", "Paling besar 20%, sekali per klien, dicatat sebagai biaya penjualan, dan diikat ke kontrak lanjutan."),
    ("Diskon volume", "5% untuk 5 kelompok atau lebih setahun, 10% untuk 10 kelompok atau lebih. Tidak pernah lebih dari 10%."),
    ("Persetujuan diskon", "0%: tidak perlu. Sampai 10%: lead sales. Sampai 20%: founder / CEO. Di atas 20%, atau penawaran dengan margin bersih di bawah 30%: Direktur."),
    ("Harga di atas daftar (mitra)", "Mitra boleh menawarkan hingga 20% di atas Harga Daftar tanpa persetujuan dan mendapat 30% dari selisihnya."),
    ("Komisi mitra non-EAP", "10% dari profit bersih event untuk deal ke-1 sampai ke-4, 12,5% mulai deal ke-5, 2,5% untuk perpanjangan. Margin bersih setelah komisi tidak boleh di bawah 15%."),
    ("Komisi mitra EAP", "Rp1,5 jt per kontrak EAP tahunan baru, Rp500rb per paket 3 bulan, Rp1 jt per perpanjangan 1 tahun."),
    ("Perjalanan dan venue", "Di luar Jakarta: ditagihkan ke klien sebesar biaya aktual, atau biaya + 10% jika TJ yang mengatur."),
    ("Pembayaran", "EAP 3 bulan: DP 50% saat SPK, pelunasan setelah laporan akhir (atau opsi bulanan, total +20%). Kontrak bank di atas Rp100 jt: DP 30%. Regulator: mengikuti ketentuan mereka."),
    ("Pajak", "Teman Journey dan Coreitera Non-PKP: jangan pernah menulis 'termasuk PPN'. PPh 23 dipotong klien B2B."),
    ("Regulator", "Jangan menaikkan harga item lama. Tambahkan item baru (asesmen, lisensi, pendampingan) yang mendapat pagu baru."),
    ("Posisi harga", "Elite boutique: selalu jelas di bawah trainer besar, tidak pernah yang termurah. Bersaing dengan layanan lebih lengkap dan hasil terukur."),
    ("Coreitera ke biro", "Coreitera menjual lewat biro, tidak pernah menjual tes langsung ke korporat. Mesin tes dijual sekitar biaya; keuntungan dari interpretasi dan langganan."),
    ("Kerahasiaan", "Manajemen klien hanya menerima data agregat anonim. Hasil individual hanya untuk pegawai yang bersangkutan."),
]
for i, (a, b) in enumerate(rules):
    rr = 4 + i
    put(wa, f"A{rr}", i + 1)
    put(wa, f"B{rr}", a, BOLD, align=WRAP)
    put(wa, f"C{rr}", b, align=WRAP)
    wa.row_dimensions[rr].height = 32
put(wa, f"B{6 + len(rules)}", "Sumber: Rate_Card model v15, Perjanjian Kemitraan (Pasal 4-8, Lampiran C), Model Biaya BI Sibolga, Proposal Synthesia v0.2, keputusan 19 Nov 2026.", align=WRAP)

# ---------------------------------------------------------------------------------------------
# 10. Sumber_Konflik
# ---------------------------------------------------------------------------------------------
wc = wb.create_sheet("Sumber_Konflik")
put(wc, "A1", "Harga yang berbeda antar dokumen, dan keputusan di master ini", TITLE)
header(wc, 3, ["No", "Konflik", "Dokumen", "Keputusan di master"], [5, 55, 40, 70])
conf = [
    ("Harga konseling B2C: model 2027 memakai Rp1 jt/sesi dan psikolog 40%; daftar harga Rp500-850rb dan psikolog 60-70%",
     "Model v15 sheet B2C vs B2C Pricing", "Master memakai daftar harga B2C. Model 2027 perlu diperbarui (laba B2C di model terlalu tinggi)."),
    ("Dasar biaya psikolog: Rp350rb/jam vs bagi hasil % vs retainer Rp5 jt/bulan", "RAB BI, B2C Pricing, model v15",
     "B2C: fee per tier (bagi hasil). B2B (EAP, MHCU): flat Rp350rb per sesi. Tatap muka asesmen Rp350rb/jam, kerja dokumen Rp110rb/jam."),
    ("Aturan margin: 3x, 15-26%, 35-40%, 15% setelah komisi", "Rate_Card, Sibolga, RAB BI, Kemitraan",
     "Satu aturan: margin bersih minimal 30% setelah PPh 23 dan komisi (keputusan 19 Nov 2026). Aturan 3x tidak dipakai."),
    ("Harga asesmen per orang: Rp25-50rb, Rp150rb, Rp200rb", "BSN, Rate_Card, Fit and Proper",
     "Dipisah jadi produk berbeda: profiling massal otomatis (per proyek), add-on asesmen + laporan dampak (per peserta), seleksi (per proyek)."),
    ("MHCU dijual di sekitar atau di bawah biaya", "Model Sibolga, Proposal MHCU LPDP", "Biaya alat MHCU dikoreksi ke Rp7rb (Galih 19 Nov 2026). Flash Rp500rb dan Deep Rp750rb ditetapkan di kolom Harga ditetapkan; add-on individual tetap Rp70rb karena sudah di atas margin 30%."),
    ("Tulisan 'termasuk PPN 11%' padahal Non-PKP", "Proposal MHCU LPDP, penawaran BI 28 Okt", "Semua harga di master belum termasuk PPN dan tanpa PPN. Perbaiki templat proposal."),
    ("Event publik Rp20 jt vs Rp24 jt", "Model v15 vs Lampiran C", "Master memakai Rp24 jt (Lampiran C)."),
    ("EAP di model 2027 Rp15 jt/bulan margin 67% vs paket nyata margin 20-37%", "Model v15 vs Model Sibolga",
     "Master memakai paket nyata dengan harga hari ini (keputusan 19 Nov 2026). Jeda Tahunan dan sesi tambahan di bawah margin 30%: ditandai di Katalog. Model 2027 perlu memakai harga paket dari master."),
    ("Pemetaan 3A-3C di bawah 3x (sekitar 2,5x)", "RAB BSN", "Tidak masalah lagi: aturan 3x tidak dipakai, margin bersih 62-71%."),
    ("Komisi 10% dari laba kotor di RAB vs komisi mitra di Kemitraan", "Semua RAB vs Kemitraan",
     "Katalog memakai 10% laba kotor (konservatif). Komisi mitra mengikuti Aturan_Penjualan."),
]
for i, (a, b, c) in enumerate(conf):
    rr = 4 + i
    put(wc, f"A{rr}", i + 1)
    put(wc, f"B{rr}", a, align=WRAP)
    put(wc, f"C{rr}", b, align=WRAP)
    put(wc, f"D{rr}", c, align=WRAP)
    wc.row_dimensions[rr].height = 45

# Order sheets: Panduan, Ringkasan, Katalog, Simulasi, Aturan, Rincian, Tarif, Input, Konflik
order = ["Panduan", "Ringkasan", "Katalog", "Simulasi_Penawaran", "Aturan_Penjualan", "Rincian_Biaya", "Tarif_Biaya", "Input", "Sumber_Konflik"]
wb._sheets = [wb[n] for n in order]
for sh in wb.worksheets:
    sh.sheet_view.showGridLines = False
    for row in sh.iter_rows():
        for c in row:
            if c.font is None or c.font.name != F:
                c.font = Font(name=F, bold=c.font.bold if c.font else False, color=c.font.color if c.font else None)
wb.save(OUT)
print("saved", os.path.abspath(OUT), "products:", len(P), "cost lines:", BOM_END - 4)
