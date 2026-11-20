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
put(ws, "A2", "Satu sumber harga untuk tim sales. Dibuat 19 Nov 2026, struktur lini dan segmen 20 Nov 2026, dari semua proposal, RAB, model 2027 v15, "
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
    ("ATURAN HARGA (keputusan 20 Nov 2026)", None),
    ("Satu biaya, harga per segmen", "Setiap produk punya satu biaya dasar (Rincian_Biaya). Harga = biaya dinaikkan sampai target margin "
                                     "segmen: Pendidikan dan komunitas 30%, Korporat dan pemerintah 50%, Regulator menuju 50%. Margin bersih "
                                     "dihitung setelah PPh 23 yang dipotong klien dan komisi 10% dari laba kotor."),
    ("Pengaman harga", "Korporat dan Regulator: harga saat ini yang sudah lebih tinggi tidak diturunkan. Regulator: item yang sudah "
                       "pernah dibeli (kolom 'Regulator dibekukan' = Ya) tetap memakai harga saat ini. Kolom 'Harga regulator "
                       "ditetapkan' (biru) mengganti harga regulator bila diisi, misalnya untuk menyesuaikan pagu SBM."),
    ("Aturan harga per produk", "'Penuh' = harga per segmen dari rumus. 'Langganan' = satu harga untuk biro (Coreitera), minimal 30%. "
                                "'Pasar' = harga mengikuti pasar dan bagi hasil psikolog (konseling dan tes individu, event bertiket). "
                                "'Harga pokok' = dijual sekitar biaya (mesin tes white-label). 'Pintu masuk' = gratis, biaya antar hampir nol."),
    ("Peran produk", "Pintu masuk (membuka hubungan), Inti (membawa laba), Retensi (kontrak berulang). Pintu masuk hanya boleh gratis "
                     "kalau biaya antarnya hampir nol; semua yang memakai waktu psikolog per orang adalah Inti dan tidak pernah gratis."),
    ("Biaya psikolog", "Tatap muka (konseling, wawancara, fasilitasi) Rp350rb/jam; kerja dokumen Rp110rb/jam; pembicara per kegiatan 2 jam "
                       "menurut tier (Rp1,5 / 2,5 / 3 / 4 jt). B2C: harga x bagi hasil tier psikolog (60/65/65/70%)."),
    ("Pembulatan", "Harga dibulatkan ke atas: Rp1rb di bawah Rp100rb, Rp25rb di bawah Rp1 jt, Rp100rb di bawah Rp10 jt, Rp500rb di atasnya."),
    ("SHEET", None),
    ("Ringkasan", "Jumlah produk per lini, produk baru, dan margin rata-rata per segmen."),
    ("Katalog", "Semua produk: lini, peran, segmen, harga saat ini, biaya dasar, harga Pendidikan, Korporat, Regulator, dan margin masing-masing."),
    ("Rincian_Biaya", "Biaya tiap produk per komponen: jumlah x tarif. Inilah sumber kolom Biaya dasar di Katalog."),
    ("Tarif_Biaya", "Satu tarif untuk tiap sumber daya (jam psikolog, pembicara, co-fasilitator, alat tes, webinar, server, AI)."),
    ("Input", "PPh 23, komisi, target margin per segmen, lini produk, harga dan bagi hasil tier psikolog, pembulatan."),
    ("Simulasi_Penawaran", "Pilih kode produk, segmen, jumlah, dan diskon: keluar total harga, laba bersih, margin, status warna, dan siapa yang menyetujui."),
    ("Aturan_Penjualan", "Aturan segmen, pintu masuk, diskon, bagi hasil, pembayaran, perjalanan, dan persetujuan."),
    ("Sumber_Konflik", "Harga yang dulu berbeda antar dokumen, preseden harga khusus, dan bagaimana master ini menyelesaikannya."),
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

SEG_START = 15
put(wi, f"A{SEG_START-2}", "Target margin bersih per segmen klien (keputusan 20 Nov 2026)", BOLD, fill=SECTION)
for col in "BCDEF":
    wi[f"{col}{SEG_START-2}"].fill = SECTION
for i, lab in enumerate(["Kunci", "Segmen", "Target margin bersih", "Siapa", "Status", "Aturan"], 1):
    c = wi.cell(row=SEG_START - 1, column=i, value=lab)
    c.font, c.fill = HDR_FONT, HDR_FILL
segments = [
    ("SEG_EDU", "Pendidikan dan komunitas", 0.30, "Sekolah, kampus, yayasan, komunitas",
     "Batas bawah. Versi produk lebih ringan; minimal Rp10 jt per kontrak; bagi hasil sekolah maks. 10%."),
    ("SEG_KORP", "Korporat dan pemerintah", 0.50, "Bank, BUMN, swasta, kementerian, pemda",
     "Harga saat ini yang lebih tinggi tidak diturunkan."),
    ("SEG_REG", "Regulator", 0.50, "BI, OJK, LPS",
     "Target menuju 50%, tidak di bawah 30%; dijual sebagai program dan paket, acuan pagu SBM. Item lama dibekukan."),
]
for i, (k, seg, m, who, rule) in enumerate(segments):
    rr = SEG_START + i
    put(wi, f"A{rr}", k, BOLD)
    put(wi, f"B{rr}", seg)
    put(wi, f"C{rr}", m, BLUE, PCT)
    put(wi, f"D{rr}", who)
    put(wi, f"E{rr}", "Keputusan")
    put(wi, f"F{rr}", rule, align=WRAP)
    name(k, f"Input!$C${rr}")
SEG_END = SEG_START + len(segments) - 1

FAM_START = SEG_END + 4
put(wi, f"A{FAM_START-2}", "Lini produk (struktur 20 Nov 2026)", BOLD, fill=SECTION)
for col in "BCDEF":
    wi[f"{col}{FAM_START-2}"].fill = SECTION
for i, lab in enumerate(["Lini", "Pertanyaan klien", "Unit bisnis"], 1):
    c = wi.cell(row=FAM_START - 1, column=i, value=lab)
    c.font, c.fill = HDR_FONT, HDR_FILL
families = [
    ("1. Asesmen", "Siapa orang-orang ini, dan bagaimana kondisinya?", "Teman Journey"),
    ("2. Konseling dan pendampingan", "Siapa yang menemani orang kami saat tidak baik-baik saja?", "Teman Journey"),
    ("3. Pembelajaran dan psikoedukasi", "Bagaimana orang kami belajar dan berubah?", "Teman Journey"),
    ("4. Pengalaman dan komunitas", "Acara apa yang bermakna untuk komunitas kami?", "Teman Journey"),
    ("5. Konsultasi organisasi", "Bagaimana sistem dan program kami jadi lebih baik?", "Teman Journey"),
    ("6. Coreitera", "Bagaimana biro saya bekerja lebih cepat?", "Coreitera"),
]
for i, (fam, q, bu) in enumerate(families):
    rr = FAM_START + i
    put(wi, f"A{rr}", fam, BOLD)
    put(wi, f"B{rr}", q)
    put(wi, f"C{rr}", bu)
FAM_END = FAM_START + len(families) - 1
name("FAM", f"Input!$A${FAM_START}:$A${FAM_END}")

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
    ("EVENT_KIT", "Activity kit per peserta (juga dipakai untuk kit seni dan kit Teman Kecil)", 30000, "per peserta", "Pasti",
     "RAB Lyfe at 30s. Pengganti sementara harga kit seni / kit rumah yang belum ada (disetujui 20 Nov 2026)"),
    ("EVENT_FNB", "Makan-minum venue (nett)", 50000, "per peserta", "Pasti", "Penawaran Kanteen n Co"),
    ("EVENT_ADS", "Iklan event komunitas (batas)", 300000, "per event", "Pasti", "Final proposal Lyfe at 30s"),
    ("EVENT_PSI", "Psikolog event komunitas (sekitar 2 jam)", 500000, "per event", "Pasti",
     "RAB Lyfe at 30s Paket 1-3 (Okt 2026). Disetujui Galih 20 Nov 2026 (dulu Rp150rb)"),
    ("SPK_T1", "Fee psikolog pembicara Tier 1, kegiatan 2 jam", 1500000, "per kegiatan", "Pasti",
     "Ratecard internal pembicara (Canva, Okt 2026). Tambahan per jam = setengahnya. Disetujui 20 Nov 2026"),
    ("SPK_T2", "Fee psikolog pembicara Tier 2, kegiatan 2 jam", 2500000, "per kegiatan", "Pasti", "Sama"),
    ("SPK_T3", "Fee psikolog pembicara Tier 3, kegiatan 2 jam", 3000000, "per kegiatan", "Pasti", "Sama"),
    ("SPK_T4", "Fee psikolog pembicara Tier 4, kegiatan 2 jam", 4000000, "per kegiatan", "Pasti", "Sama"),
    ("COFAS", "Co-fasilitator / MC / admin acara", 300000, "per acara", "Pasti", "RAB Lyfe at 30s. Disetujui 20 Nov 2026"),
    ("DOKUM", "Dokumentasi acara", 150000, "per acara", "Pasti", "RAB Lyfe at 30s. Disetujui 20 Nov 2026"),
    ("WEB_PROD", "Produksi webinar oleh TJ (platform 500 orang, desain dan e-sertifikat, host)", 2300000, "per webinar", "Pasti",
     "Costing BI Sibolga v4: Rp1,3 jt + Rp0,5 jt + Rp0,5 jt. Disetujui 20 Nov 2026"),
    ("WORKSHEET", "Worksheet / lembar kerja cetak", 5500, "per peserta", "Pasti", "Budget OSH 'Friends of Friend'. Disetujui 20 Nov 2026"),
    ("MATERI", "Materi ringkas dan poster", 200000, "per sesi", "Pasti", "Costing BI Sibolga v3. Disetujui 20 Nov 2026"),
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
    "HR / unit SDM instansi dan perusahaan", "per 3 bulan", "Aktif", "Penuh", "Ya", 12000000,
    "Model Biaya BI Sibolga, Lampiran C", eap_lines(15, webinar=1))
add("TJ-EAP-02", TJ, "B. EAP", "EAP Jeda Tahunan (per kuartal)",
    "Per kuartal: 15 sesi online, MHCU untuk 70 pegawai, laporan agregat, 1 webinar. Kontrak 12 bulan (4 kuartal)",
    "HR / unit SDM instansi dan perusahaan", "per kuartal", "Aktif", "Penuh", "Ya", 11000000,
    "Model Biaya BI Sibolga, Lampiran C", eap_lines(15, mhcu_people=70, rounds=1, webinar=1))
add("TJ-EAP-03", TJ, "B. EAP", "EAP Tumbuh", "24 sesi online, 1 putaran MHCU, laporan agregat, 1 webinar; sekitar 8 pegawai",
    "HR / unit SDM instansi dan perusahaan", "per 3 bulan", "Aktif", "Penuh", "Ya", 20000000,
    "Model Biaya BI Sibolga, Lampiran C", eap_lines(24, mhcu_people=8, rounds=1, webinar=1))
add("TJ-EAP-04", TJ, "B. EAP", "EAP Seperjalanan",
    "36 sesi online, 2 putaran MHCU (pre/post), hasil individual, laporan perbandingan, 1 seminar offline; sekitar 12 pegawai",
    "HR / unit SDM instansi dan perusahaan", "per 3 bulan", "Aktif", "Penuh", "Ya", 31500000,
    "Model Biaya BI Sibolga, Lampiran C", eap_lines(36, mhcu_people=12, rounds=2, seminar=1))
add("TJ-EAP-05", TJ, "B. EAP", "Sesi konseling tambahan EAP", "Konseling online 60 menit di luar kuota paket",
    "Klien EAP", "per sesi", "Aktif", "Penuh", "Ya", 500000, "Lampiran C, Model BI Sibolga",
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
# 4b. New products and variants (Canva inventory + approved costs, 20 Nov 2026; see 08_Kalkulasi_Biaya_Layanan.md)
# ---------------------------------------------------------------------------------------------
A1, A2, A3, A4, A5, A6 = ("1. Asesmen", "2. Konseling dan pendampingan", "3. Pembelajaran dan psikoedukasi",
                          "4. Pengalaman dan komunitas", "5. Konsultasi organisasi", "6. Coreitera")
NEW = "Baru (20 Nov 2026)"

# Line 1. Asesmen
add("TJ-ASM-11", TJ, A1, "MHCU Skrining (L1)", "Skrining DASS-21 online, review psikolog SIPP, laporan individual maks. 48 jam; tanpa konseling",
    "Perusahaan, sekolah, komunitas", "per orang", NEW, "Penuh", "Ya", None, "Proposal Mr DIY (Canva, Nov 2026): Rp250rb",
    [L("Alat ukur MHCU", 1, "orang", "MHCU_TOOL"), L("Review dan laporan psikolog", 0.25, "jam", "PSI_DOC"),
     L("Administrasi dan platform", 1, "orang", "ADM_SESI")])
add("TJ-ASM-12", TJ, A1, "MHCU Ringkas 15 menit (L2)", "5 instrumen klinis, review psikolog, laporan personal, konseling 15 menit; khusus pendidikan dan komunitas",
    "Komunitas, kampus, sekolah", "per orang", NEW, "Penuh", "Ya", None, "Telering (Canva Okt 2026) Rp242rb = setara L2",
    [L("Alat ukur MHCU", 1, "orang", "MHCU_TOOL"), L("Konseling 15 menit", 0.25, "sesi", "PSI_EAP"),
     L("Review manual dan laporan", 0.25, "jam", "PSI_DOC"), L("Administrasi dan platform", 1, "sesi", "ADM_SESI")])
add("TJ-ASM-13", TJ, A1, "Asesmen Manajerial", "Tes BFI + DISC, wawancara psikolog 1 jam, laporan dan rekomendasi per peserta",
    "HC perusahaan dan bank", "per peserta", NEW, "Penuh", "Ya", None, "Disusun dari blok RAB BI Fit and Proper",
    [L("Tes BFI + DISC", 1, "peserta", "BFI_DISC"), L("Wawancara psikolog", 1, "jam", "PSI_F2F"),
     L("Skoring wawancara", 0.5, "jam", "PSI_DOC"), L("Laporan", 2, "jam", "PSI_DOC"), L("Review mutu", 0.5, "jam", "PSI_DOC"),
     L("Administrasi", 1, "peserta", "ADM_SESI")])
add("TJ-ASM-14", TJ, A1, "Asesmen Eksekutif", "Tes, wawancara mendalam 2 jam, simulasi bisnis dengan 2 psikolog, laporan senior",
    "Direksi dan komisaris", "per peserta", NEW, "Penuh", "Ya", None, "Disusun dari blok RAB BI Fit and Proper",
    [L("Tes BFI + DISC", 1, "peserta", "BFI_DISC"), L("Wawancara mendalam", 2, "jam", "PSI_F2F"),
     L("Simulasi bisnis: 2 psikolog x 2 jam", 4, "jam", "PSI_F2F"), L("Skoring", 2, "jam", "PSI_DOC"),
     L("Laporan senior", 3, "jam", "PSI_DOC"), L("Review mutu", 1, "jam", "PSI_DOC"),
     L("Transport 2 psikolog", 2, "orang-hari", "TRANSPORT"), L("Administrasi", 1, "peserta", "ADM_SESI")])
SCHOOL_BASE = [L("Administrasi klasikal: psikolog 2,5 jam x 2 kelas / 50 siswa", 0.1, "jam", "PSI_F2F"),
               L("Asisten pengawas 2 kelas / 50 siswa", 0.04, "acara", "COFAS"),
               L("Platform dan skoring Coreitera", 1, "siswa", "CORE_FEE", internal="Ya"),
               L("Review dan laporan psikolog", 0.25, "jam", "PSI_DOC"), L("Kontrak / 50 siswa", 0.02, "SPK", "KONTRAK")]
add("TJ-ASM-15", TJ, A1, "Paket Tes Sekolah Dasar", "Tes klasikal (IQ, minat, gaya belajar) + laporan per siswa; minimal 50 siswa",
    "Sekolah", "per siswa", NEW, "Penuh", "Ya", None, "Proposal Asisi (Canva, Nov 2026): IQ klasikal Rp250rb", SCHOOL_BASE)
add("TJ-ASM-16", TJ, A1, "Paket Tes Sekolah Standar", "Paket Dasar + laporan rekap dan rekomendasi untuk sekolah",
    "Sekolah", "per siswa", NEW, "Penuh", "Ya", None, "Proposal master sekolah (Canva)",
    SCHOOL_BASE + [L("Rekap dan rekomendasi sekolah 6 jam / 50 siswa", 0.12, "jam", "PSI_DOC")])
add("TJ-ASM-17", TJ, A1, "Paket Tes Sekolah Lengkap", "Paket Standar + konseling hasil 30 menit untuk orang tua",
    "Sekolah", "per siswa", NEW, "Penuh", "Ya", None, "Proposal master sekolah (Canva)",
    SCHOOL_BASE + [L("Rekap dan rekomendasi sekolah 6 jam / 50 siswa", 0.12, "jam", "PSI_DOC"),
                   L("Konseling hasil ke orang tua 30 menit", 0.5, "jam", "PSI_F2F"), L("Administrasi sesi", 1, "sesi", "ADM_SESI")])
add("TJ-ASM-18", TJ, A1, "Asesmen Perkembangan Anak (kelompok)", "Asesmen bersama orang tua, 1 asesor psikolog per 3 anak, 90 menit, 5-12 anak per sesi",
    "Sekolah, PAUD, komunitas keluarga", "per anak", NEW, "Penuh", "Ya", None, "Proposal TJxDH (Canva, Sep 2026): Rp250rb per anak",
    [L("Asesor psikolog 1,5 jam / 3 anak", 0.5, "jam", "PSI_F2F"), L("PIC psikolog pembuka 1,5 jam / 12 anak", 0.125, "jam", "PSI_F2F"),
     L("Laporan dan rekomendasi", 0.5, "jam", "PSI_DOC"), L("Administrasi", 1, "anak", "ADM_SESI")])
add("TJ-ASM-19", TJ, A1, "Snapshot Burnout untuk HR (L0)", "Tes mandiri online + laporan agregat otomatis untuk HR; pintu masuk, selalu diikuti proposal MHCU atau EAP",
    "HR perusahaan", "per orang", NEW, "Pintu masuk", "Tidak", None, "Keputusan 20 Nov 2026: pintu masuk gratis bila biaya hampir nol",
    [L("Laporan AI", 1, "laporan", "AI_LAPORAN", internal="Ya"), L("Server dan storage", 1, "orang", "TEST_SERVER", internal="Ya")])

# Line 2. Konseling dan pendampingan
add("TJ-EAP-06", TJ, A2, "EAP Korporat (retainer, 300 pegawai)",
    "Pool 10 sesi konseling per bulan, MHCU 2x setahun untuk semua pegawai, laporan agregat, 1 webinar per kuartal; harga per bulan untuk 300 pegawai",
    "HR bank, BUMN, perusahaan besar", "per bulan", NEW, "Penuh", "Ya", None, "Keputusan 20 Nov 2026; rencana 2027 Rp15 jt/bulan",
    [L("Sesi konseling (pool)", 10, "sesi", "PSI_EAP", note="Asumsi model 2027: 10 sesi/bulan"),
     L("Administrasi dan platform", 10, "sesi", "ADM_SESI"), L("Alat ukur MHCU, 300 x 2 / 12", 50, "orang", "MHCU_TOOL"),
     L("Laporan agregat 2x setahun", 2 / 12, "laporan", "AGG_REPORT"), L("Webinar per kuartal: pembicara Tier 2", 1 / 3, "kegiatan", "SPK_T2"),
     L("Webinar per kuartal: host/moderator", 1 / 3, "kegiatan", None, 500000, note="Bagian dari WEB_PROD"),
     L("Webinar per kuartal: materi dan poster", 1 / 3, "sesi", "MATERI"), L("Account management", 4, "jam", "ADM_JAM"),
     L("Kontrak dibagi 12 bulan", 1 / 12, "SPK", "KONTRAK")])
add("TJ-KON-01", TJ, A2, "Sesi Konseling B2B 30 menit", "Konseling online 30 menit di luar kuota paket",
    "Klien B2B", "per sesi", NEW, "Penuh", "Ya", None, "Proposal TFI (Canva, Agu 2026)",
    [L("Fee psikolog 30 menit", 0.5, "sesi", "PSI_EAP"), L("Administrasi dan platform", 1, "sesi", "ADM_SESI")])
add("TJ-KON-02", TJ, A2, "Paket 10 Sesi (60 menit) + MHCU", "10 sesi konseling online 60 menit, 10 MHCU, 10 laporan individual; 6 bulan",
    "Komunitas, sekolah, kantor kecil", "per paket", NEW, "Penuh", "Ya", None, "Proposal TFI Welcoming Rp5,05 jt (preseden)",
    [L("Fee psikolog", 10, "sesi", "PSI_EAP"), L("Administrasi dan platform", 10, "sesi", "ADM_SESI"),
     L("Alat ukur MHCU", 10, "orang", "MHCU_TOOL"), L("Laporan individual", 2.5, "jam", "PSI_DOC"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-KON-03", TJ, A2, "Paket 10 Sesi (30 menit) + MHCU", "10 sesi konseling online 30 menit, 10 MHCU, 10 laporan individual; 6 bulan",
    "Komunitas, sekolah, kantor kecil", "per paket", NEW, "Penuh", "Ya", None, "Proposal TFI Welcoming Rp3,3 jt (preseden)",
    [L("Fee psikolog 30 menit", 5, "sesi", "PSI_EAP"), L("Administrasi dan platform", 10, "sesi", "ADM_SESI"),
     L("Alat ukur MHCU", 10, "orang", "MHCU_TOOL"), L("Laporan individual", 2.5, "jam", "PSI_DOC"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-KON-04", TJ, A2, "Paket 20 Sesi + MHCU Pre/Post (18 bulan)", "Maks. 20 sesi online 60 menit, 2 putaran MHCU untuk 29 orang (58 laporan); 18 bulan",
    "Program komunitas dan yayasan", "per kontrak", NEW, "Penuh", "Ya", None, "Proposal TFI TWB Full Impact Rp19,5 jt",
    [L("Fee psikolog", 20, "sesi", "PSI_EAP"), L("Administrasi dan platform", 20, "sesi", "ADM_SESI"),
     L("Alat ukur MHCU", 58, "orang", "MHCU_TOOL"), L("Laporan individual", 14.5, "jam", "PSI_DOC"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-KON-05", TJ, A2, "Konseling Kelompok 90 menit", "Sesi kelompok 5-8 orang dengan 1 psikolog",
    "Komunitas, sekolah, perusahaan", "per sesi", NEW, "Penuh", "Ya", None, "Proposal TFI (Canva, Agu 2026)",
    [L("Psikolog tatap muka", 1.5, "jam", "PSI_F2F"), L("Catatan sesi", 0.5, "jam", "PSI_DOC"), L("Administrasi dan platform", 1, "sesi", "ADM_SESI")])
add("TJ-KON-06", TJ, A2, "In-house Psychologist (1 hari)", "Psikolog di lokasi: konseling, office hours, rujukan HR; 6 jam layanan",
    "Perusahaan, sekolah", "per hari", NEW, "Penuh", "Ya", None, "Proposal korporat dan sekolah (Canva)",
    [L("Psikolog tatap muka", 6, "jam", "PSI_F2F"), L("Catatan dan rujukan", 2, "jam", "PSI_DOC"), L("Transport", 1, "orang-hari", "TRANSPORT")])
add("TJ-KON-07", TJ, A2, "In-house Psychologist (setengah hari)", "Psikolog di lokasi 3 jam layanan; untuk sekolah",
    "Sekolah", "per setengah hari", NEW, "Penuh", "Ya", None, "Proposal sekolah (Canva)",
    [L("Psikolog tatap muka", 3, "jam", "PSI_F2F"), L("Catatan", 1, "jam", "PSI_DOC"), L("Transport", 1, "orang-hari", "TRANSPORT")])

# Line 3. Pembelajaran dan psikoedukasi
for code, seg, price in (("TJ-PLT-08", "BUMN", 45000000), ("TJ-PLT-09", "Swasta", 55000000)):
    add(code, TJ, A3, f"Program Bank Essential {seg}", "Hari pelatihan saja, tanpa asesmen dan laporan dampak", f"Bank {seg}",
        "per hari", "Aktif", "Penuh", "Ya", price, "Rate_Card model 2027 v15 (paket Essential)",
        [L("Pelaksanaan 1 hari", 1, "hari", "TRN_BANK")])
for t in (1, 2, 3, 4):
    add(f"TJ-SPK-0{t}", TJ, A3, f"Pembicara / Webinar Tier {t}, 2 jam",
        "Kegiatan 2 jam (webinar, seminar, kuliah umum) oleh psikolog; materi disiapkan TJ; platform milik klien",
        "Semua klien", "per kegiatan", NEW, "Penuh", "Ya", None, "Ratecard pembicara (Canva Okt 2026)",
        [L(f"Fee psikolog pembicara Tier {t}", 1, "kegiatan", f"SPK_T{t}"), L("Kontrak", 1, "SPK", "KONTRAK")])
for t in (1, 2, 3, 4):
    add(f"TJ-SPK-1{t}", TJ, A3, f"Tambahan durasi pembicara Tier {t}", "Per jam di atas 2 jam",
        "Semua klien", "per jam", NEW, "Penuh", "Ya", None, "Ratecard pembicara (Canva Okt 2026): 50% tarif per kegiatan",
        [L(f"Fee psikolog Tier {t} per jam", 0.5, "kegiatan", f"SPK_T{t}")])
add("TJ-SPK-20", TJ, A3, "Produksi Webinar oleh TJ (tambahan)", "Platform kapasitas 500 orang, desain poster dan e-sertifikat, host/moderator",
    "Semua klien", "per webinar", NEW, "Penuh", "Ya", None, "Costing BI Sibolga v4",
    [L("Produksi webinar", 1, "webinar", "WEB_PROD")])
add("TJ-WS-01", TJ, A3, "Workshop Korporat 2 jam (Healthy Working Environment)", "Workshop tatap muka Tier 2, 30 peserta, co-fasilitator, worksheet",
    "HR perusahaan", "per sesi", NEW, "Penuh", "Ya", None, "Proposal korporat (Canva)",
    [L("Fee pembicara Tier 2", 1, "kegiatan", "SPK_T2"), L("Co-fasilitator", 1, "acara", "COFAS"),
     L("Transport 2 orang", 2, "orang-hari", "TRANSPORT"), L("Worksheet", 30, "peserta", "WORKSHEET"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-WS-02", TJ, A3, "Art-Based Wellbeing 2 jam", "Sesi seni terpandu Tier 2 (kompetensi art therapy), 30 peserta, kit seni",
    "HR perusahaan, komunitas", "per sesi", NEW, "Penuh", "Ya", None, "Proposal korporat (Canva)",
    [L("Fee pembicara Tier 2", 1, "kegiatan", "SPK_T2"), L("Co-fasilitator", 1, "acara", "COFAS"),
     L("Transport 2 orang", 2, "orang-hari", "TRANSPORT"), L("Kit seni", 30, "peserta", "EVENT_KIT"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-WS-03", TJ, A3, "Team-Building 4 jam", "Aktivitas experiential Tier 2, 30 peserta, 2 co-fasilitator, kit permainan",
    "HR perusahaan", "per sesi", NEW, "Penuh", "Ya", None, "Proposal korporat (Canva)",
    [L("Fee pembicara Tier 2 (2 jam)", 1, "kegiatan", "SPK_T2"), L("Tambahan 2 jam Tier 2", 1, "kegiatan", "SPK_T2"),
     L("Co-fasilitator", 2, "acara", "COFAS"), L("Transport 3 orang", 3, "orang-hari", "TRANSPORT"),
     L("Kit permainan", 30, "peserta", "EVENT_KIT"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-WS-04", TJ, A3, "Workshop Sekolah dan Kampus 2 jam",
    "Career Readiness 101, Gadget Addiction, Sexual Education, UpStander, atau Parenting Class (Standar); Tier 1, 40 peserta",
    "Sekolah, kampus", "per sesi", NEW, "Penuh", "Ya", None, "Proposal sekolah dan kampus (Canva)",
    [L("Fee pembicara Tier 1", 1, "kegiatan", "SPK_T1"), L("Co-fasilitator", 1, "acara", "COFAS"),
     L("Transport 2 orang", 2, "orang-hari", "TRANSPORT"), L("Worksheet", 40, "peserta", "WORKSHEET"), L("Kontrak", 1, "SPK", "KONTRAK")])
add("TJ-WS-05", TJ, A3, "Tambahan Psikolog Kedua (versi Lengkap)", "Misalnya Sexual Education Lengkap: 1 psikolog perempuan + 1 laki-laki",
    "Sekolah, kampus", "per sesi", NEW, "Penuh", "Ya", None, "Proposal sekolah (Canva)",
    [L("Fee pembicara Tier 1", 1, "kegiatan", "SPK_T1")])

# Line 4. Pengalaman dan komunitas
TK = 20
add("TJ-EXP-01", TJ, A4, "Teman Kecil Paket Komplit", "3 aktivitas 2,5 jam untuk anak + orang tua, 1 psikolog utama, 2 co-fasilitator, kit rumah dan Lentera Tumbuh; basis 20 pasangan",
    "Perusahaan (family day), sekolah", "per pasangan", NEW, "Penuh", "Ya", None, "Proposal JLL (Canva, Nov 2026): Rp500-550rb",
    [L("Psikolog utama 2,5 jam / 20", 2.5 / TK, "jam", "PSI_F2F"), L("Co-fasilitator x2 / 20", 2 / TK, "acara", "COFAS"),
     L("Koordinator / 20", 1 / TK, "acara", "COFAS"), L("Dokumentasi / 20", 1 / TK, "acara", "DOKUM"),
     L("Transport 4 orang / 20", 4 / TK, "orang-hari", "TRANSPORT"), L("Dokumen refleksi 2 jam / 20", 2 / TK, "jam", "PSI_DOC"),
     L("Kit rumah + Lentera Tumbuh", 2, "kit", "EVENT_KIT"), L("Kontrak / 20", 1 / TK, "SPK", "KONTRAK")])
add("TJ-EXP-02", TJ, A4, "Teman Kecil Art Therapy Only", "1 aktivitas 90-100 menit, kit rumah dibawa pulang; basis 20 pasangan",
    "Perusahaan (family day), sekolah", "per pasangan", NEW, "Penuh", "Ya", None, "Proposal JLL (Canva, Nov 2026): Rp400-450rb",
    [L("Psikolog utama 1,75 jam / 20", 1.75 / TK, "jam", "PSI_F2F"), L("Co-fasilitator x2 / 20", 2 / TK, "acara", "COFAS"),
     L("Koordinator / 20", 1 / TK, "acara", "COFAS"), L("Transport 4 orang / 20", 4 / TK, "orang-hari", "TRANSPORT"),
     L("Dokumen refleksi 2 jam / 20", 2 / TK, "jam", "PSI_DOC"), L("Kit rumah", 1, "kit", "EVENT_KIT"), L("Kontrak / 20", 1 / TK, "SPK", "KONTRAK")])
add("TJ-EXP-03", TJ, A4, "Jasa Event untuk Mitra (12 peserta)", "TJ menjalankan event milik mitra: psikolog, fasilitasi, kit, MC dan registrasi, dokumentasi",
    "Kreator, penulis, komunitas", "per event", NEW, "Penuh", "Tidak", None, "RAB Lyfe at 30s Paket 3 (Rp2,5 jt untuk <=12 peserta)",
    [L("Psikolog event", 1, "event", "EVENT_PSI"), L("Dokumentasi", 1, "acara", "DOKUM"), L("Activity kit", 12, "peserta", "EVENT_KIT"),
     L("MC / admin dan registrasi", 1, "acara", "COFAS")])

# Line 5. Konsultasi organisasi
add("TJ-KNS-01", TJ, A5, "Hari Konsultan (HR Advisory / evaluasi program)", "4 jam tatap muka + 4 jam analisis dan penulisan oleh psikolog senior",
    "Direksi, HC, unit program", "per hari", NEW, "Penuh", "Ya", None, "Pola RAB BSN",
    [L("Psikolog senior tatap muka", 4, "jam", "PSI_F2F"), L("Analisis dan penulisan", 4, "jam", "PSI_DOC"), L("Transport", 1, "orang-hari", "TRANSPORT")])
add("TJ-KNS-02", TJ, A5, "FGD 2 jam (online)", "Fasilitator psikolog + notulen, olah transkrip dan ringkasan",
    "Regulator, perusahaan", "per FGD", NEW, "Penuh", "Ya", None, "Proposal BI DPU (Canva, Nov 2026)",
    [L("Fasilitator psikolog", 2, "jam", "PSI_F2F"), L("Notulen / co-fasilitator", 1, "acara", "COFAS"), L("Olah transkrip dan ringkasan", 3, "jam", "PSI_DOC")])

# Line, role and segment for every product: (line, role, Pendidikan?, Korporat?, Regulator?, regulator price frozen?)
LINE_OF_OLD = {"A. Pelatihan": A3, "B. EAP": A2, "C. Asesmen": A1, "E. B2C Konseling": A2, "F. B2C Tes": A1,
               "G. Coreitera Software": A6, "H. Coreitera Asesmen": A6}
LINE_OF_CODE = {"TJ-EDU-01": A5, "TJ-EDU-04": A4, "TJ-EDU-05": A4}
SEGS = {  # code: (role, edu, korp, reg, frozen)
    "TJ-PLT-01": ("Inti", 0, 0, 1, 1), "TJ-PLT-02": ("Inti", 0, 0, 1, 1), "TJ-PLT-03": ("Inti", 0, 1, 0, 0),
    "TJ-PLT-04": ("Inti", 0, 1, 0, 0), "TJ-PLT-05": ("Retensi", 0, 1, 0, 0), "TJ-PLT-06": ("Retensi", 0, 1, 0, 0),
    "TJ-PLT-07": ("Inti", 0, 1, 1, 1), "TJ-PLT-08": ("Inti", 0, 1, 0, 0), "TJ-PLT-09": ("Inti", 0, 1, 0, 0),
    "TJ-EAP-01": ("Retensi", 1, 1, 1, 1), "TJ-EAP-02": ("Retensi", 1, 1, 1, 1), "TJ-EAP-03": ("Retensi", 1, 1, 1, 1),
    "TJ-EAP-04": ("Retensi", 1, 1, 1, 1), "TJ-EAP-05": ("Inti", 1, 1, 1, 1), "TJ-EAP-06": ("Retensi", 0, 1, 1, 0),
    "TJ-ASM-01": ("Inti", 0, 1, 1, 1), "TJ-ASM-02": ("Inti", 0, 1, 1, 0), "TJ-ASM-03": ("Inti", 0, 1, 1, 0),
    "TJ-ASM-04": ("Inti", 0, 1, 1, 0), "TJ-ASM-05": ("Inti", 0, 1, 1, 0), "TJ-ASM-06": ("Inti", 0, 1, 1, 0),
    "TJ-ASM-07": ("Inti", 0, 1, 1, 0), "TJ-ASM-08": ("Inti", 1, 1, 1, 0), "TJ-ASM-09": ("Inti", 1, 1, 1, 0),
    "TJ-ASM-10": ("Retensi", 1, 1, 1, 1), "TJ-ASM-11": ("Inti", 1, 1, 1, 0), "TJ-ASM-12": ("Inti", 1, 0, 0, 0),
    "TJ-ASM-13": ("Inti", 0, 1, 1, 0), "TJ-ASM-14": ("Inti", 0, 1, 1, 0), "TJ-ASM-15": ("Inti", 1, 0, 0, 0),
    "TJ-ASM-16": ("Inti", 1, 0, 0, 0), "TJ-ASM-17": ("Inti", 1, 0, 0, 0), "TJ-ASM-18": ("Inti", 1, 0, 0, 0),
    "TJ-ASM-19": ("Pintu masuk", 0, 0, 0, 0),
    "TJ-KON-01": ("Inti", 1, 1, 1, 0), "TJ-KON-02": ("Inti", 1, 1, 0, 0), "TJ-KON-03": ("Inti", 1, 1, 0, 0),
    "TJ-KON-04": ("Retensi", 1, 1, 0, 0), "TJ-KON-05": ("Inti", 1, 1, 0, 0), "TJ-KON-06": ("Retensi", 1, 1, 0, 0),
    "TJ-KON-07": ("Retensi", 1, 0, 0, 0),
    "TJ-SPK-20": ("Inti", 1, 1, 1, 0), "TJ-WS-01": ("Inti", 0, 1, 1, 0), "TJ-WS-02": ("Inti", 1, 1, 1, 0),
    "TJ-WS-03": ("Inti", 0, 1, 1, 0), "TJ-WS-04": ("Inti", 1, 0, 0, 0), "TJ-WS-05": ("Inti", 1, 0, 0, 0),
    "TJ-EXP-01": ("Inti", 1, 1, 0, 0), "TJ-EXP-02": ("Inti", 1, 1, 0, 0), "TJ-EXP-03": ("Inti", 1, 0, 0, 0),
    "TJ-EDU-01": ("Inti", 1, 1, 1, 1), "TJ-EDU-04": ("Inti", 0, 1, 1, 0), "TJ-EDU-05": ("Inti", 0, 0, 0, 0),
    "TJ-KNS-01": ("Inti", 1, 1, 1, 0), "TJ-KNS-02": ("Inti", 1, 1, 1, 0),
    "CI-THT-01": ("Inti", 0, 0, 0, 0), "CI-THT-02": ("Inti", 0, 0, 0, 0), "CI-THT-03": ("Retensi", 0, 0, 0, 0),
    "CI-THT-04": ("Inti", 0, 0, 0, 0), "CI-SYN-01": ("Inti", 0, 0, 0, 0), "CI-SYN-02": ("Inti", 0, 0, 0, 0),
    "CI-ASM-01": ("Pintu masuk", 0, 0, 0, 0), "CI-ASM-02": ("Pintu masuk", 0, 0, 0, 0),
}
for t in (1, 2, 3, 4):
    SEGS[f"TJ-SPK-0{t}"] = ("Inti", 1, 1, 1, 0)
    SEGS[f"TJ-SPK-1{t}"] = ("Inti", 1, 1, 1, 0)
LANGGANAN = {"CI-THT-01", "CI-THT-02", "CI-THT-03", "CI-SYN-01", "CI-SYN-02"}
P2 = []
for p in P:
    p = list(p)
    code = p[0]
    p[2] = LINE_OF_CODE.get(code, LINE_OF_OLD.get(p[2], p[2]))
    if code in LANGGANAN:
        p[8] = "Langganan"
    if code.startswith("TJ-B2C") or code.startswith("TJ-TES"):
        role = ("Inti", 0, 0, 0, 0)
    else:
        role = SEGS[code]
    P2.append(tuple(p) + (role,))
LINE_ORDER = [f[0] for f in families]
P = sorted(P2, key=lambda x: LINE_ORDER.index(x[2]))

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
            put(wr, f"F{r}", f"=Katalog!$P${code_row[code]}", GREEN, RP)
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
put(wk, "A1", "Katalog produk dan harga per segmen: Teman Journey + Coreitera", TITLE)
put(wk, "A2", "Satu biaya dasar per produk; harga per segmen dari target margin di Input. Rupiah, belum termasuk PPN. "
              "Kolom biru boleh diubah; kolom lain rumus. Harga kosong = produk tidak dijual ke segmen itu.")
put(wk, "A3", "Cek: 'Lolos' = semua harga segmen minimal 30% margin bersih. 'Regulator (beku) di bawah 30%' = harga lama regulator "
              "dipertahankan; perlu persetujuan Direktur. 'Info' = aturan Pasar, Langganan, Harga pokok, atau Pintu masuk.")
cols = ["Kode", "Unit bisnis", "Lini", "Produk", "Isi / cakupan", "Pembeli utama", "Satuan", "Status",
        "Peran", "Aturan harga", "Potong PPh 23?", "Pendidikan?", "Korporat?", "Regulator?", "Regulator dibekukan?",
        "Harga saat ini (Rp)", "Biaya dasar (Rp)", "Margin bersih harga saat ini",
        "HARGA DASAR / LANTAI 30% (Rp)", "HARGA PENDIDIKAN (Rp)", "Margin Pendidikan",
        "HARGA KORPORAT (Rp)", "Margin Korporat", "HARGA REGULATOR (Rp)", "Margin Regulator",
        "Harga regulator ditetapkan (Rp, opsional)", "Cek aturan", "Sumber", "Harga daftar utama (Rp)",
        "Bantu: harga 30% sebelum bulat", "Bantu: harga 50% sebelum bulat", "Bantu: harga target regulator sebelum bulat"]
widths = [13, 13, 22, 36, 46, 26, 14, 14, 11, 11, 9, 9, 9, 9, 10, 14, 14, 10, 15, 15, 10, 15, 10, 15, 10, 14, 16, 36, 15, 13, 13, 13]
header(wk, 5, cols, widths)
wk.row_dimensions[5].height = 58
B = "Rincian_Biaya!$A$5:$A$%d" % BOM_END
H = "Rincian_Biaya!$H$5:$H$%d" % BOM_END
I_ = "Rincian_Biaya!$I$5:$I$%d" % BOM_END


def rnd(x):
    step = f"IF({x}<100000,STEP0,IF({x}<1000000,STEP1,IF({x}<10000000,STEP2,STEP3)))"
    return f"ROUNDUP({x}/{step},0)*{step}"


PRICE_FILL = PatternFill("solid", fgColor="E2EFDA")
for i, p in enumerate(P):
    rr = KAT_FIRST + i
    code, unit, fam, prod, scope, buyer, uom, status, rule, pph, cur, src, _, seg = p
    role, edu, korp, reg, frozen = seg
    for col, v in zip("ABCDEFGH", (code, unit, fam, prod, scope, buyer, uom, status)):
        put(wk, f"{col}{rr}", v, BOLD if col == "A" else BLACK, align=WRAP if col in "DEF" else None)
    put(wk, f"I{rr}", role, BLUE)
    put(wk, f"J{rr}", rule, BLUE)
    put(wk, f"K{rr}", pph, BLUE)
    for col, flag in zip("LMNO", (edu, korp, reg, frozen)):
        put(wk, f"{col}{rr}", "Ya" if flag else "Tidak", BLUE)
    put(wk, f"P{rr}", cur, BLUE, RP)
    put(wk, f"Q{rr}", f"=SUMIFS({H},{B},A{rr})", GREEN, RP)
    tp = f'IF(K{rr}="Ya",TARIF_PPH23,0)'
    kom = f'IF(OR(J{rr}="Penuh",J{rr}="Langganan"),KOMISI,0)'

    def net_margin(pc):
        gp = f"({pc}*(1-{tp})-Q{rr})"
        return f'=IF(AND(ISNUMBER({pc}),IF(ISNUMBER({pc}),{pc},0)>0),({gp}-MAX(0,{gp}*{kom}))/{pc},"")'

    def raw(m):
        return f"=IF(Q{rr}>0,Q{rr}*(1-KOMISI)/((1-{tp})*(1-KOMISI)-{m}),0)"

    put(wk, f"AD{rr}", raw("SEG_EDU"), BLACK, RP)
    put(wk, f"AE{rr}", raw("SEG_KORP"), BLACK, RP)
    put(wk, f"AF{rr}", raw("SEG_REG"), BLACK, RP)
    put(wk, f"R{rr}", net_margin(f"P{rr}"), BLACK, PCT)
    put(wk, f"S{rr}", f'=IF(J{rr}="Pintu masuk",0,IF(J{rr}="Pasar",P{rr},IF(J{rr}="Harga pokok",IF(IF(ISNUMBER(P{rr}),P{rr},0)>0,P{rr},'
                      f'{rnd(f"(Q{rr}/(1-{tp}))")}),IF(J{rr}="Langganan",MAX(IF(ISNUMBER(P{rr}),P{rr},0),{rnd(f"AD{rr}")}),{rnd(f"AD{rr}")}))))',
        Font(name=F, bold=True), RP)
    put(wk, f"T{rr}", f'=IF(AND(J{rr}="Penuh",L{rr}="Ya"),S{rr},"")', Font(name=F, bold=True), RP, fill=PRICE_FILL)
    put(wk, f"U{rr}", net_margin(f"T{rr}"), BLACK, PCT)
    put(wk, f"V{rr}", f'=IF(AND(J{rr}="Penuh",M{rr}="Ya"),MAX(IF(ISNUMBER(P{rr}),P{rr},0),{rnd(f"AE{rr}")}),"")', Font(name=F, bold=True), RP, fill=PRICE_FILL)
    put(wk, f"W{rr}", net_margin(f"V{rr}"), BLACK, PCT)
    put(wk, f"X{rr}", f'=IF(AND(J{rr}="Penuh",N{rr}="Ya"),IF(IF(ISNUMBER(Z{rr}),Z{rr},0)>0,Z{rr},IF(AND(O{rr}="Ya",IF(ISNUMBER(P{rr}),P{rr},0)>0),P{rr},'
                      f'MAX(IF(ISNUMBER(P{rr}),P{rr},0),{rnd(f"AF{rr}")}))),"")', Font(name=F, bold=True), RP, fill=PRICE_FILL)
    put(wk, f"Y{rr}", net_margin(f"X{rr}"), BLACK, PCT)
    put(wk, f"Z{rr}", None, BLUE, RP)
    put(wk, f"AA{rr}", f'=IF(J{rr}<>"Penuh","Info",IF(AND(ISNUMBER(Y{rr}),IF(ISNUMBER(Y{rr}),Y{rr},0)<0.2999),"Regulator (beku) di bawah 30%",'
                       f'IF(MIN(IF(ISNUMBER(U{rr}),U{rr},1),IF(ISNUMBER(W{rr}),W{rr},1))<0.2999,"Di bawah 30%","Lolos")))')
    put(wk, f"AB{rr}", src, align=WRAP)
    put(wk, f"AC{rr}", f'=IF(ISNUMBER(V{rr}),V{rr},IF(ISNUMBER(T{rr}),T{rr},IF(ISNUMBER(X{rr}),X{rr},S{rr})))', BOLD, RP)
KAT_LAST = KAT_FIRST + len(P) - 1
wk.freeze_panes = "E6"
wk.auto_filter.ref = f"A5:AF{KAT_LAST}"
for col in ("AD", "AE", "AF"):
    wk.column_dimensions[col].hidden = True
green = PatternFill("solid", fgColor="C6EFCE")
red = PatternFill("solid", fgColor="FFC7CE")
wk.conditional_formatting.add(f"AA{KAT_FIRST}:AA{KAT_LAST}", CellIsRule(operator="equal", formula=['"Lolos"'], fill=green))
wk.conditional_formatting.add(f"AA{KAT_FIRST}:AA{KAT_LAST}", FormulaRule(formula=[f'ISNUMBER(SEARCH("30%",AA{KAT_FIRST}))'], fill=red))
wk.conditional_formatting.add(f"AA{KAT_FIRST}:AA{KAT_LAST}", CellIsRule(operator="notEqual", formula=['"Lolos"'],
                                                                         fill=PatternFill("solid", fgColor="FFF2CC")))
for f1, opts, cols_ in (("I", '"Pintu masuk,Inti,Retensi"', "I"), ("J", '"Penuh,Langganan,Pasar,Harga pokok,Pintu masuk"', "J"),
                        ("K", '"Ya,Tidak"', "K"), ("L", '"Ya,Tidak"', "L"), ("M", '"Ya,Tidak"', "M"), ("N", '"Ya,Tidak"', "N"),
                        ("O", '"Ya,Tidak"', "O")):
    d = DataValidation(type="list", formula1=opts)
    wk.add_data_validation(d)
    d.add(f"{f1}{KAT_FIRST}:{f1}{KAT_LAST}")
for nm, col in (("KAT_CODE", "A"), ("KAT_FAM", "C"), ("KAT_PROD", "D"), ("KAT_UOM", "G"), ("KAT_ROLE", "I"), ("KAT_RULE", "J"),
                ("KAT_PPH", "K"), ("KAT_CUR", "P"), ("KAT_COST", "Q"), ("KAT_BASE", "S"), ("KAT_EDU", "T"), ("KAT_MEDU", "U"),
                ("KAT_KORP", "V"), ("KAT_MKORP", "W"), ("KAT_REG", "X"), ("KAT_MREG", "Y"), ("KAT_CHECK", "AA"), ("KAT_PRICE", "AC")):
    name(nm, f"Katalog!${col}${KAT_FIRST}:${col}${KAT_LAST}")

# ---------------------------------------------------------------------------------------------
# 7. Ringkasan
# ---------------------------------------------------------------------------------------------
ws = wb.create_sheet("Ringkasan", 1)
put(ws, "A1", "Ringkasan: produk per lini dan margin per segmen", TITLE)
put(ws, "A2", "Semua angka rumus dari Katalog. Ubah target margin per segmen di Input untuk melihat dampaknya.")
header(ws, 4, ["Lini", "Jumlah produk", "Produk baru (20 Nov 2026)", "Pintu masuk", "Inti", "Retensi",
               "Dijual ke Pendidikan", "Dijual ke Korporat", "Dijual ke Regulator",
               "Margin rata-rata Pendidikan", "Margin rata-rata Korporat", "Perlu persetujuan (di bawah 30%)"],
       [34, 11, 13, 11, 9, 9, 12, 12, 12, 13, 13, 15])
for i, (fam, _, _) in enumerate(families):
    rr = 5 + i
    put(ws, f"A{rr}", fam, BOLD)
    put(ws, f"B{rr}", f"=COUNTIFS(KAT_FAM,A{rr})", GREEN)
    put(ws, f"C{rr}", f'=COUNTIFS(KAT_FAM,A{rr},Katalog!$H${KAT_FIRST}:$H${KAT_LAST},"{NEW}")', GREEN)
    for col, rl in zip("DEF", ("Pintu masuk", "Inti", "Retensi")):
        put(ws, f"{col}{rr}", f'=COUNTIFS(KAT_FAM,A{rr},KAT_ROLE,"{rl}")', GREEN)
    for col, rng in zip("GHI", ("KAT_EDU", "KAT_KORP", "KAT_REG")):
        put(ws, f"{col}{rr}", f'=COUNTIFS(KAT_FAM,A{rr},{rng},">0")', GREEN)
    put(ws, f"J{rr}", f"=IFERROR(AVERAGEIFS(KAT_MEDU,KAT_FAM,A{rr}),0)", GREEN, PCT)
    put(ws, f"K{rr}", f"=IFERROR(AVERAGEIFS(KAT_MKORP,KAT_FAM,A{rr}),0)", GREEN, PCT)
    put(ws, f"L{rr}", f'=COUNTIFS(KAT_FAM,A{rr},KAT_CHECK,"*30%*")', GREEN)
tot = 5 + len(families)
put(ws, f"A{tot}", "Total", BOLD)
for col in "BCDEFGHIL":
    put(ws, f"{col}{tot}", f"=SUM({col}5:{col}{tot-1})", BOLD)
put(ws, f"A{tot+2}", "Target margin: Pendidikan", BOLD)
put(ws, f"C{tot+2}", "=SEG_EDU", GREEN, PCT)
put(ws, f"A{tot+3}", "Target margin: Korporat dan pemerintah", BOLD)
put(ws, f"C{tot+3}", "=SEG_KORP", GREEN, PCT)
put(ws, f"A{tot+4}", "Target margin: Regulator", BOLD)
put(ws, f"C{tot+4}", "=SEG_REG", GREEN, PCT)
put(ws, f"A{tot+6}", "Cara baca: margin rata-rata Korporat bisa di atas 50% karena harga saat ini yang lebih tinggi tidak diturunkan "
                     "(misalnya program bank, profiling, pemetaan). 'Perlu persetujuan' = harga regulator yang dibekukan di bawah 30%.",
    align=WRAP)
ws.merge_cells(f"A{tot+6}:L{tot+6}")
ws.row_dimensions[tot + 6].height = 40

# ---------------------------------------------------------------------------------------------
# 8. Simulasi_Penawaran
# ---------------------------------------------------------------------------------------------
wq = wb.create_sheet("Simulasi_Penawaran")
put(wq, "A1", "Simulasi penawaran", TITLE)
put(wq, "A2", "Pilih kode produk dan segmen (daftar), isi jumlah dan diskon (biru). Harga mengikuti segmen. "
              "Contoh: 30 MHCU Flash + 1 workshop korporat untuk perusahaan.")
put(wq, "A3", "Klien:", BOLD)
put(wq, "B3", "(isi nama klien)", BLUE)
put(wq, "D3", "Bagi hasil sekolah / mitra (% harga, maks. 10%):", BOLD)
put(wq, "H3", 0, BLUE, PCT)
header(wq, 5, ["No", "Kode produk", "Produk", "Satuan", "Segmen", "Jumlah", "Harga daftar segmen (Rp)", "Diskon",
               "Harga setelah diskon (Rp)", "Total harga (Rp)", "Total biaya (Rp)", "Potong PPh 23?", "Aturan",
               "Target margin segmen", "Catatan"],
       [5, 13, 40, 16, 13, 9, 16, 9, 16, 16, 16, 10, 11, 11, 30])
Q1, QN = 6, 20
dvq = DataValidation(type="list", formula1=f"=Katalog!$A${KAT_FIRST}:$A${KAT_LAST}", allow_blank=True)
wq.add_data_validation(dvq)
dvq.add(f"B{Q1}:B{QN}")
dvs = DataValidation(type="list", formula1='"Pendidikan,Korporat,Regulator,Umum"', allow_blank=True)
wq.add_data_validation(dvs)
dvs.add(f"E{Q1}:E{QN}")
example = {Q1: ("TJ-ASM-08", "Korporat", 30, 0), Q1 + 1: ("TJ-WS-01", "Korporat", 1, 0.05)}
for rr in range(Q1, QN + 1):
    ex = example.get(rr, (None, None, None, None))
    put(wq, f"A{rr}", rr - Q1 + 1)
    put(wq, f"B{rr}", ex[0], BLUE)
    put(wq, f"E{rr}", ex[1], BLUE)
    put(wq, f"F{rr}", ex[2], BLUE, "#,##0")
    put(wq, f"H{rr}", ex[3], BLUE, PCT)
    look = lambda rng: f'INDEX({rng},MATCH($B{rr},KAT_CODE,0))'
    put(wq, f"C{rr}", f'=IF($B{rr}="","",{look("KAT_PROD")})', GREEN)
    put(wq, f"D{rr}", f'=IF($B{rr}="","",{look("KAT_UOM")})', GREEN)
    put(wq, f"G{rr}", f'=IF($B{rr}="","",IF($E{rr}="Pendidikan",{look("KAT_EDU")},IF($E{rr}="Korporat",{look("KAT_KORP")},'
                      f'IF($E{rr}="Regulator",{look("KAT_REG")},{look("KAT_BASE")}))))', GREEN, RP)
    put(wq, f"I{rr}", f'=IF(ISNUMBER(G{rr}),G{rr}*(1-H{rr}),0)', BLACK, RP)
    put(wq, f"J{rr}", f'=IF($B{rr}="",0,I{rr}*F{rr})', BLACK, RP)
    put(wq, f"K{rr}", f'=IF($B{rr}="",0,{look("KAT_COST")}*F{rr})', GREEN, RP)
    put(wq, f"L{rr}", f'=IF($B{rr}="","",{look("KAT_PPH")})', GREEN)
    put(wq, f"M{rr}", f'=IF($B{rr}="","",{look("KAT_RULE")})', GREEN)
    put(wq, f"N{rr}", f'=IF($B{rr}="",0,IF($E{rr}="Pendidikan",SEG_EDU,IF($E{rr}="Korporat",SEG_KORP,IF($E{rr}="Regulator",SEG_REG,0))))', BLACK, PCT)
    put(wq, f"O{rr}", f'=IF($B{rr}="","",IF(ISNUMBER(G{rr}),"","Tidak dijual ke segmen ini: pilih segmen lain"))', BLACK)
s = QN + 2
summary = [
    ("Total harga ke klien (belum PPN)", f"=SUM(J{Q1}:J{QN})", RP),
    ("PPh 23 dipotong klien", f'=SUMPRODUCT((L{Q1}:L{QN}="Ya")*J{Q1}:J{QN})*TARIF_PPH23', RP),
    ("Total biaya", f"=SUM(K{Q1}:K{QN})", RP),
    ("Bagi hasil sekolah / mitra", f"=C{s}*H3", RP),
    ("Laba kotor", f"=C{s}-C{s+1}-C{s+2}-C{s+3}", RP),
    ("Komisi sales (10% dari laba kotor)", f"=MAX(0,C{s+4})*KOMISI", RP),
    ("Laba bersih", f"=C{s+4}-C{s+5}", RP),
    ("Margin bersih", f"=IF(C{s}>0,C{s+6}/C{s},0)", PCT),
    ("Target margin (rata-rata tertimbang segmen)", f"=IF(C{s}>0,SUMPRODUCT(J{Q1}:J{QN},N{Q1}:N{QN})/C{s},0)", PCT),
    ("Diskon terbesar", f"=MAX(H{Q1}:H{QN})", PCT),
    ("Status", f'=IF(C{s+7}>=C{s+8}-0.0001,"Hijau: sesuai target",IF(C{s+7}>=0.3,"Kuning: di bawah target, perlu persetujuan",'
               f'"Merah: di bawah 30%, tidak boleh tanpa Direktur"))', None),
    ("Kontrak pendidikan minimal Rp10 jt?", f'=IF(COUNTIF(E{Q1}:E{QN},"Pendidikan")=0,"Tidak berlaku",IF(C{s}>=10000000,"Ya",'
                                            f'"Tidak: tambah item ke paket"))', None),
    ("Persetujuan yang dibutuhkan", f'=IF(OR(C{s+7}<0.3,C{s+9}>0.2,H3>0.1),"Direktur",IF(OR(C{s+7}<C{s+8}-0.0001,C{s+9}>0.1),'
                                    f'"Founder / CEO",IF(C{s+9}>0,"Lead sales","Tidak perlu")))', None),
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
    ("Harga per segmen", "Pakai harga segmen di Katalog: Pendidikan dan komunitas (target 30%), Korporat dan pemerintah (target 50%), "
                         "Regulator (menuju 50%, minimal 30%). Harga kosong = produk tidak dijual ke segmen itu."),
    ("Satu metode, harga berbeda", "Instrumen, standar laporan, dan review psikolog sama untuk semua segmen. Yang berbeda hanya durasi, "
                                   "senioritas psikolog, dan harga. Jangan menjual 'versi murah'."),
    ("Pintu masuk", "Gratis atau di bawah pasar hanya untuk produk yang biaya antarnya hampir nol (tes mandiri, snapshot burnout, "
                    "webinar rekaman, konten). Selalu diikuti proposal produk inti. Subsidi maks. Rp6 jt per perusahaan, dicatat sebagai "
                    "biaya penjualan. Konversi dilacak; dihentikan bila di bawah 5% dalam 6 bulan. Tidak gratis untuk regulator dan bank: "
                    "untuk mereka, pintu masuknya pilot berbayar."),
    ("Produk inti tidak pernah gratis", "Semua yang memakai waktu psikolog per orang (MHCU L1-L4, workshop, sesi, asesmen) dijual minimal di harga segmen."),
    ("Pendidikan: minimal kontrak", "Kontrak segmen Pendidikan minimal Rp10 jt. Item tunggal hanya sebagai tambahan di dalam paket."),
    ("Bagi hasil sekolah / mitra", "Maksimal 10% dari harga, dicatat sebagai biaya di Simulasi_Penawaran. Di atas 10%: Direktur."),
    ("Regulator", "Jual sebagai program dan paket, bukan honor per jam. Acuan pagu SBM (honor narasumber sekitar Rp1,7 jt per jam). "
                  "Jangan menaikkan harga item yang sudah pernah dibeli; tambahkan item baru."),
    ("Satuan harga", "Pelatihan per kelompok per hari; asesmen per peserta; EAP per paket, per kuartal, atau per bulan; pembicara per 2 jam "
                     "+ per jam tambahan; software per biro atau per kursi per bulan. Tidak pernah satu angka borongan tanpa rincian."),
    ("Diskon pilot", "Paling besar 20%, sekali per klien, dicatat sebagai biaya penjualan, dan diikat ke kontrak lanjutan."),
    ("Diskon volume", "5% untuk 5 kelompok atau lebih setahun, 10% untuk 10 kelompok atau lebih. Tidak ada diskon volume untuk MHCU."),
    ("Persetujuan", "Margin di atas target segmen dan diskon 0%: tidak perlu. Diskon sampai 10%: lead sales. Margin di bawah target segmen "
                    "atau diskon sampai 20%: founder / CEO. Margin di bawah 30%, diskon di atas 20%, atau bagi hasil di atas 10%: Direktur."),
    ("Komisi sales", "10% dari laba kotor. Harga lebih tinggi = margin perusahaan lebih tinggi dan komisi sales lebih besar."),
    ("Harga di atas daftar (mitra)", "Mitra boleh menawarkan hingga 20% di atas harga segmen tanpa persetujuan dan mendapat 30% dari selisihnya."),
    ("Komisi mitra non-EAP", "10% dari profit bersih event untuk deal ke-1 sampai ke-4, 12,5% mulai deal ke-5, 2,5% untuk perpanjangan. Margin bersih setelah komisi tidak boleh di bawah 15%."),
    ("Komisi mitra EAP", "Rp1,5 jt per kontrak EAP tahunan baru, Rp500rb per paket 3 bulan, Rp1 jt per perpanjangan 1 tahun."),
    ("Perjalanan dan venue", "Di luar Jakarta: ditagihkan ke klien sebesar biaya aktual, atau biaya + 10% jika TJ yang mengatur."),
    ("Pembayaran", "Pendidikan dan komunitas: deposit penuh di muka sebelum pelaksanaan. EAP 3 bulan: DP 50% saat SPK, pelunasan setelah laporan akhir "
                   "(atau bayar bulanan, total +20%). Kontrak bank dan korporat di atas Rp100 jt: DP 30%. Regulator: mengikuti ketentuan mereka."),
    ("Pajak", "Teman Journey dan Coreitera Non-PKP: jangan pernah menulis 'termasuk PPN'. PPh 23 dipotong klien B2B."),
    ("Posisi harga", "Butik premium: tidak ikut perang harga untuk produk inti. Bersaing di interpretasi, laporan, dan program."),
    ("Coreitera ke biro", "Coreitera menjual lewat biro, tidak pernah menjual tes langsung ke korporat. Mesin tes dijual sekitar biaya; keuntungan dari interpretasi dan langganan."),
    ("Kerahasiaan", "Manajemen klien hanya menerima data agregat anonim. Hasil individual hanya untuk pegawai yang bersangkutan."),
]
for i, (a, b) in enumerate(rules):
    rr = 4 + i
    put(wa, f"A{rr}", i + 1)
    put(wa, f"B{rr}", a, BOLD, align=WRAP)
    put(wa, f"C{rr}", b, align=WRAP)
    wa.row_dimensions[rr].height = 32
put(wa, f"B{6 + len(rules)}", "Sumber: Rate_Card model v15, Perjanjian Kemitraan (Pasal 4-8, Lampiran C), Model Biaya BI Sibolga, Proposal Synthesia v0.2, keputusan 19 dan 20 Nov 2026 (05_Struktur_Produk_dan_Segmen.md).", align=WRAP)

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
    ("MHCU Telering: Flash Rp242rb, Deep Rp380rb (diskon volume 5% / 7,5%), di bawah biaya", "Proposal Telering (Canva, Okt 2026)",
     "Dihormati untuk Telering saja (preseden). Ke depan, komunitas seperti Telering memakai MHCU Ringkas 15 menit (TJ-ASM-12). Tidak ada diskon volume MHCU."),
    ("MHCU DASS-21 tanpa konseling Rp250rb (Mr DIY)", "Proposal Mr DIY (Canva, Nov 2026)", "Menjadi produk TJ-ASM-11 (MHCU Skrining) dengan harga per segmen."),
    ("Paket TFI: Welcoming Rp5,05 jt / Rp3,3 jt, add-on 5 sesi Rp1,35 jt (rugi)", "Proposal Teacher Wellbeing TFI (Canva, Agu 2026)",
     "Harga lama. Struktur baru (paket 10 dan 20 sesi) dihitung ulang: TJ-KON-02 sampai 04. Add-on sesi memakai harga sesi B2B."),
    ("Dua ratecard pembicara: internal Rp2,8/4,5/6/8 jt; kampus Rp3/4,5/7/8 jt dengan tambahan per 2 jam (rugi)", "Canva: ratecard internal dan proposal Teman Kampus",
     "Fee psikolog dari ratecard internal masuk Tarif_Biaya (SPK_T1-T4). Harga per segmen dari rumus; tambahan durasi dihitung per jam."),
    ("Psikolog event komunitas Rp150rb vs Rp500rb", "RAB Lyfe at 30s (dua versi)", "Rp500rb (disetujui 20 Nov 2026)."),
    ("EAP bank di rencana 2027 Rp15 jt/bulan vs paket kecil Rp4-10,5 jt/bulan", "Model 2027 v15 vs paket Jeda/Tumbuh/Seperjalanan",
     "EAP korporat menjadi produk sendiri (TJ-EAP-06), Rp13,5 jt per bulan untuk 300 pegawai pada target 50%."),
    ("Asesmen anak Deen Haus Rp250rb per anak, di bawah biaya bila asesor psikolog", "Proposal TJxDH (Canva, Sep 2026)",
     "Produk TJ-ASM-18 dengan harga Pendidikan dari rumus. Rp250rb hanya mungkin dengan asisten terlatih (bukan psikolog)."),
    ("Paket EAP dulu 'Harga tetap'", "Keputusan 19 Nov 2026", "Keputusan 20 Nov 2026: harga per segmen. Harga lama tetap untuk regulator (BI KPw Sibolga) lewat aturan dibekukan."),
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
