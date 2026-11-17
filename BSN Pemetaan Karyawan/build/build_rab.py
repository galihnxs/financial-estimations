"""Builds the internal cost and price model for the BSN (Bank Syariah Nasional)
employee mapping offer: BFI + DISC for 2,000 employees, three packages.

  Paket 1 Laporan untuk Perusahaan : tes BFI + DISC + laporan otomatis, hanya dibaca perusahaan
  Paket 2 Laporan untuk Perusahaan dan Karyawan: sama, karyawan juga membaca laporannya (Rp35k di 2.000, DIKUNCI)
  Paket 3 Pemetaan Talenta    : Paket 2 + analisis jabatan, kecocokan orang-jabatan, talent box,
                                kalibrasi per unit, laporan pemetaan, presentasi
  On demand (biaya terpisah)  : sesi penjelasan hasil, umpan balik individu, unit / rumpun jabatan tambahan

Run:  <python with openpyxl> build_rab.py   Output lands one folder up.
Blue = input. Yellow = estimate to confirm. Every other number is a formula.
Recalculate in Excel before reading values.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent.parent / "RAB - BSN Pemetaan Karyawan BFI DISC (Internal).xlsx"

RP = '"Rp"#,##0;[Red]-"Rp"#,##0'
PCT = "0.0%"
MULT = '0.0"x"'
F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F)
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
TITLE = Font(name=F, bold=True, size=12)
WHITE_BOLD = Font(name=F, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="475571")
EST_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="B7BFCF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")


def header(ws, row, labels):
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.font, c.fill, c.alignment, c.border = WHITE_BOLD, HEAD_FILL, WRAP, BOX


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def put(ws, ref, value, font=BLACK, fmt=None, fill=None):
    c = ws[ref]
    c.value, c.font = value, font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    c.border = BOX
    return c


wb = Workbook()
REF = {}


def r(key):
    return REF[key]


# ---------------------------------------------------------------- Asumsi
A = wb.active
A.title = "Asumsi"
A["A1"] = "ASUMSI - BSN Pemetaan Karyawan (INTERNAL). Biru = input. Kuning = estimasi, mohon dikonfirmasi."
A["A1"].font = TITLE
header(A, 2, ["Key", "Asumsi", "Nilai", "Satuan", "Status", "Sumber / catatan"])

ASSUMPTIONS = [
    ("N", "Jumlah peserta", 2000, "orang", "Known", "Permintaan BSN", "#,##0"),
    ("FX", "Kurs USD ke Rupiah", 16500, "Rp/USD", "Estimate", "Untuk biaya cloud dalam USD", "#,##0"),
    # Teknologi
    ("AI", "Biaya AI per laporan (Google Gemini)", 1250, "Rp/laporan", "Your number - confirm unit",
     "Galih: 'charged by Google 1250'. Semua paket (laporan otomatis)", RP),
    ("RERUN", "Laporan gagal dan harus diulang", 0.03, "%", "Estimate", "Worker laporan max_retries 0, ulang manual", PCT),
    ("MB", "Ukuran file per peserta (PDF + gambar)", 2, "MB", "Estimate", "130 laporan = 100 MB gambar + PDF", "0.0"),
    ("GBM", "Tarif storage per GB per bulan", 0.023, "USD", "Known", "Vercel Blob Pro", "0.000"),
    ("KEEP", "Lama simpan data", 12, "bulan", "Your plan - confirm", "Ikut kontrak BSN", "0"),
    ("MAIL", "Email undangan + 3 pengingat (1 bulan)", 500000, "Rp", "Estimate", "Mailtrap, sekitar 8.000 email", RP),
    ("ENG", "Tarif kerja engineer / tim teknis", 150000, "Rp/jam", "Estimate", "Belum ada tarif resmi", RP),
    ("SETUP", "Jam setup: impor karyawan, akun, paket, branding, uji coba", 16, "jam", "Estimate", "", "0"),
    ("ADM", "Tarif admin / koordinator proyek TJ", 50000, "Rp/jam", "Estimate", "", RP),
    ("TIX", "Peserta yang butuh bantuan", 0.05, "%", "Estimate", "Lupa link, email salah, putus koneksi", PCT),
    ("TIXMIN", "Menit per bantuan", 15, "menit", "Estimate", "", "0"),
    # Serentak
    ("VMUSD", "Upgrade server sementara (e2-standard-4, Jakarta) per bulan", 130, "USD", "Estimate",
     "Prod e2-small 2 GB, database di VM yang sama. Load test R7: Postgres jenuh 40-70 request/detik", "#,##0"),
    ("VMBLN", "Lama upgrade server", 1, "bulan", "Estimate", "", "0"),
    ("LOAD", "Jam load test 2.000 peserta + tuning worker laporan", 16, "jam", "Estimate", "Worker laporan 5 paralel", "0"),
    ("SBY_P", "Engineer standby di hari serentak", 2, "orang", "Estimate", "", "0"),
    ("SBY_H", "Jam standby per orang per hari", 8, "jam", "Estimate", "", "0"),
    ("SBY_D", "Jumlah hari serentak", 2, "hari", "Your plan - confirm", "Misal 2 gelombang x 1.000 peserta", "0"),
    # Tarif psikolog
    ("PDOC", "Honor psikolog kerja dokumen", 110000, "Rp/jam", "Known", "Tarif TJ (28 Okt 2026)", RP),
    ("PF2F", "Honor psikolog tatap muka (termasuk online)", 350000, "Rp/jam", "Known", "Biaya dasar konseling TJ", RP),
    # Paket 3: penggerak lingkup
    ("JF", "Jumlah rumpun jabatan yang dianalisis", 15, "rumpun", "Known",
     "PENGGERAK 1. Diputuskan Galih 4 Nov 2026. Misal teller, CS, RM, back office, manajer cabang, dst. Tanyakan struktur jabatan BSN", "0"),
    ("UNITS", "Jumlah unit yang dipetakan dan dikalibrasi (direktorat / wilayah)", 10, "unit", "Known",
     "PENGGERAK 2. Diputuskan Galih 4 Nov 2026", "0"),
    ("EDGE_B", "Level B: karyawan di kotak ujung talent box yang direview psikolog", 0.10, "% peserta", "Estimate",
     "Hanya talenta kunci", PCT),
    ("EDGE_C", "Level C: tambahan review kotak ujung", 0.10, "% peserta", "Estimate",
     "Ditambah karyawan yang perlu perhatian, total 20%", PCT),
    # Paket 3: jam kerja
    ("KO_P", "Kick-off: psikolog", 2, "orang", "Estimate", "", "0"),
    ("KO_H", "Kick-off: jam per psikolog (tatap muka)", 2, "jam", "Estimate", "Lingkup, rumpun jabatan, data kinerja", "0"),
    ("JA_F2F", "Analisis jabatan: wawancara narasumber per rumpun (tatap muka)", 1.5, "jam/rumpun", "Estimate",
     "1 psikolog dengan atasan / pemegang jabatan", "0.0"),
    ("JA_DOC", "Analisis jabatan: telaah jobdesk + tulis profil jabatan per rumpun", 3, "jam/rumpun", "Estimate",
     "Tugas, kompetensi, profil BFI dan DISC target", "0.0"),
    ("JA_RV", "Review profil jabatan oleh psikolog senior", 0.5, "jam/rumpun", "Estimate", "", "0.0"),
    ("DASH_H", "Level A: tim teknis siapkan tabel dan grafik agregat per unit", 8, "jam", "Estimate", "", "0"),
    ("TB_ENG", "Level B: tim teknis bangun talent box di platform", 8, "jam", "Estimate",
     "CEK: fitur talent box belum ada di platform, kemungkinan build baru", "0"),
    ("FIT_ENG", "Level C: tim teknis kode profil jabatan dan skor kecocokan", 8, "jam", "Estimate", "", "0"),
    ("PERF_ENG", "Tim teknis: impor data kinerja BSN", 4, "jam", "Estimate", "BSN harus kirim nilai kinerja", "0"),
    ("CAL_P_B", "Level B: psikolog di kalibrasi talent box per unit", 1, "orang", "Estimate", "", "0"),
    ("CAL_P_C", "Level C: tambahan psikolog di kalibrasi", 1, "orang", "Estimate", "Total 2 psikolog", "0"),
    ("CAL_H", "Kalibrasi: jam per psikolog per unit (tatap muka)", 2, "jam/unit", "Estimate", "Bersama HC dan atasan unit", "0"),
    ("EDGE_MIN", "Review psikolog per karyawan di kotak ujung", 15, "menit/orang", "Estimate", "", "0"),
    ("AN_U", "Analisis data per unit", 1.5, "jam/unit", "Estimate", "", "0.0"),
    ("WR_U", "Penulisan ringkasan per unit", 2, "jam/unit", "Estimate", "", "0.0"),
    ("OV_P", "Laporan gambaran umum BSN: psikolog", 2, "orang", "Estimate", "", "0"),
    ("OV_H", "Laporan gambaran umum: jam per psikolog", 8, "jam", "Estimate", "Pola lintas unit, saran HC", "0"),
    ("RV_H", "Review mutu laporan oleh psikolog senior", 4, "jam", "Estimate", "", "0"),
    ("PR_P", "Presentasi hasil ke direksi / HC BSN: psikolog", 2, "orang", "Estimate", "", "0"),
    ("PR_H", "Presentasi: jam per psikolog (tatap muka)", 3, "jam", "Estimate", "", "0"),
    ("TRP", "Transport psikolog per orang per hari", 150000, "Rp", "Estimate", "Jakarta. Kalibrasi online", RP),
    ("TRP_D", "Kehadiran psikolog (orang x hari)", 4, "orang-hari", "Estimate", "Kick-off + presentasi", "0"),
    ("PM_A", "Level A: koordinasi proyek TJ", 20, "jam", "Estimate", "", "0"),
    ("PM_B", "Level B: tambahan koordinasi (data kinerja, jadwal kalibrasi)", 10, "jam", "Estimate", "", "0"),
    ("PM_C", "Level C: tambahan koordinasi (narasumber analisis jabatan)", 10, "jam", "Estimate", "", "0"),
    # On demand
    ("BRF_P", "On demand: psikolog di sesi penjelasan hasil", 2, "orang", "Your plan - confirm", "Bukan gratis. Biaya terpisah", "0"),
    ("BRF_H", "On demand: durasi sesi penjelasan hasil (online)", 1.5, "jam", "Your plan - confirm", "", "0.0"),
    ("FB_H", "On demand: sesi umpan balik individu", 0.5, "jam", "Your plan - confirm", "", "0.0"),
    # Pajak dan lain-lain
    ("PPN", "Tarif PPN efektif (0 jika entitas Non-PKP)", 0, "%", "Known",
     "TJ dan Coreitera sama-sama Non-PKP (Galih, 4 Nov 2026), jadi 0", PCT),
    ("PPH23", "PPh 23 dipotong BSN", 0.02, "%", "Known", "Dihitung sebagai biaya (konservatif)", PCT),
    ("KOM", "Komisi (% dari profit kotor)", 0.10, "%", "Known", "Sama dengan RAB BI", PCT),
    ("CTR", "Kontrak dan materai", 150000, "Rp", "Known", "Rujukan RAB BI", RP),
    ("OPEX", "Biaya operasional bulanan TJ", 20000000, "Rp", "Known", "Rujukan RAB BI", RP),
]
for i, (key, label, value, unit, status, source, fmt) in enumerate(ASSUMPTIONS, start=3):
    put(A, f"A{i}", key)
    put(A, f"B{i}", label)
    put(A, f"C{i}", value, BLUE, fmt, EST_FILL if status != "Known" else None)
    put(A, f"D{i}", unit)
    put(A, f"E{i}", status)
    put(A, f"F{i}", source)
    for col in "BF":
        A[f"{col}{i}"].alignment = WRAP
    REF[key] = f"Asumsi!$C${i}"
widths(A, [9, 56, 12, 12, 22, 66])

# ---------------------------------------------------------------- Biaya
B = wb.create_sheet("Biaya")
B["A1"] = "BIAYA - per blok. Psikolog = orang x jam x tarif. Teknologi = per peserta atau per proyek."
B["A1"].font = TITLE
header(B, 2, ["No", "Blok", "Item biaya", "Qty", "Satuan", "Tarif", "Total", "Dipakai di"])

BLK_BASE, BLK_AI, BLK_SER = "1. Tes (semua paket)", "2. Laporan AI", "3. Tambahan serentak"
MA, MB, MC = "4A. Pemetaan umum", "4B. Tambahan talent box", "4C. Tambahan analisis jabatan"
COSTS = [
    (BLK_BASE, "Storage hasil dan laporan selama masa simpan", f"={r('N')}*{r('MB')}/1024*{r('KEEP')}", "GB-bulan", f"={r('GBM')}*{r('FX')}", "Paket 1-3"),
    (BLK_BASE, "Email undangan dan pengingat", "=1", "paket", f"={r('MAIL')}", "Paket 1-3"),
    (BLK_BASE, "Setup akun, paket, branding, uji coba", f"={r('SETUP')}", "jam", f"={r('ENG')}", "Paket 1-3"),
    (BLK_BASE, "Helpdesk peserta", f"={r('N')}*{r('TIX')}*{r('TIXMIN')}/60", "jam", f"={r('ADM')}", "Paket 1-3"),
    (BLK_BASE, "Kontrak dan materai", "=1", "deal", f"={r('CTR')}", "Paket 1-3"),
    (BLK_AI, "Laporan AI per karyawan (termasuk yang diulang)", f"={r('N')}*(1+{r('RERUN')})", "laporan", f"={r('AI')}", "Semua paket"),
    (BLK_SER, "Upgrade server sementara", f"={r('VMBLN')}", "bulan", f"={r('VMUSD')}*{r('FX')}", "Serentak"),
    (BLK_SER, "Load test dan tuning worker laporan", f"={r('LOAD')}", "jam", f"={r('ENG')}", "Serentak"),
    (BLK_SER, "Engineer standby hari H", f"={r('SBY_P')}*{r('SBY_H')}*{r('SBY_D')}", "jam", f"={r('ENG')}", "Serentak"),
    (MA, "Kick-off (tatap muka)", f"={r('KO_P')}*{r('KO_H')}", "jam", f"={r('PF2F')}", "Level A-C"),
    (MA, "Tabel dan grafik agregat per unit", f"={r('DASH_H')}", "jam", f"={r('ENG')}", "Level A-C"),
    (MA, "Analisis data per unit", f"={r('UNITS')}*{r('AN_U')}", "jam", f"={r('PDOC')}", "Level A-C"),
    (MA, "Ringkasan per unit", f"={r('UNITS')}*{r('WR_U')}", "jam", f"={r('PDOC')}", "Level A-C"),
    (MA, "Laporan gambaran umum BSN", f"={r('OV_P')}*{r('OV_H')}", "jam", f"={r('PDOC')}", "Level A-C"),
    (MA, "Review mutu laporan (senior)", f"={r('RV_H')}", "jam", f"={r('PDOC')}", "Level A-C"),
    (MA, "Presentasi hasil (tatap muka)", f"={r('PR_P')}*{r('PR_H')}", "jam", f"={r('PF2F')}", "Level A-C"),
    (MA, "Transport psikolog", f"={r('TRP_D')}", "orang-hari", f"={r('TRP')}", "Level A-C"),
    (MA, "Koordinasi proyek", f"={r('PM_A')}", "jam", f"={r('ADM')}", "Level A-C"),
    (MB, "Talent box di platform", f"={r('TB_ENG')}", "jam", f"={r('ENG')}", "Level B-C"),
    (MB, "Impor data kinerja BSN", f"={r('PERF_ENG')}", "jam", f"={r('ENG')}", "Level B-C"),
    (MB, "Kalibrasi talent box per unit, 1 psikolog (online)", f"={r('UNITS')}*{r('CAL_P_B')}*{r('CAL_H')}", "jam", f"={r('PF2F')}", "Level B-C"),
    (MB, "Review talenta kunci di talent box", f"={r('N')}*{r('EDGE_B')}*{r('EDGE_MIN')}/60", "jam", f"={r('PDOC')}", "Level B-C"),
    (MB, "Koordinasi tambahan", f"={r('PM_B')}", "jam", f"={r('ADM')}", "Level B-C"),
    (MC, "Analisis jabatan: wawancara narasumber", f"={r('JF')}*{r('JA_F2F')}", "jam", f"={r('PF2F')}", "Level C"),
    (MC, "Analisis jabatan: telaah jobdesk dan profil jabatan", f"={r('JF')}*{r('JA_DOC')}", "jam", f"={r('PDOC')}", "Level C"),
    (MC, "Review profil jabatan (senior)", f"={r('JF')}*{r('JA_RV')}", "jam", f"={r('PDOC')}", "Level C"),
    (MC, "Skor kecocokan orang-jabatan di platform", f"={r('FIT_ENG')}", "jam", f"={r('ENG')}", "Level C"),
    (MC, "Psikolog kedua di kalibrasi", f"={r('UNITS')}*{r('CAL_P_C')}*{r('CAL_H')}", "jam", f"={r('PF2F')}", "Level C"),
    (MC, "Review karyawan yang perlu perhatian", f"={r('N')}*{r('EDGE_C')}*{r('EDGE_MIN')}/60", "jam", f"={r('PDOC')}", "Level C"),
    (MC, "Koordinasi tambahan", f"={r('PM_C')}", "jam", f"={r('ADM')}", "Level C"),
]
first = 3
for i, (block, label, qty, unit, rate, used) in enumerate(COSTS, start=first):
    put(B, f"A{i}", i - first + 1)
    put(B, f"B{i}", block)
    put(B, f"C{i}", label)
    put(B, f"D{i}", qty, GREEN, "#,##0.0")
    put(B, f"E{i}", unit)
    put(B, f"F{i}", rate, GREEN, RP)
    put(B, f"G{i}", f"=D{i}*F{i}", BLACK, RP)
    put(B, f"H{i}", used)
last = first + len(COSTS) - 1
row = last + 2
SUBT = {}
for blk in [BLK_BASE, BLK_AI, BLK_SER, MA, MB, MC]:
    put(B, f"C{row}", f"Subtotal {blk}", BOLD)
    put(B, f"G{row}", f'=SUMIF($B${first}:$B${last},"{blk}",$G${first}:$G${last})', BOLD, RP)
    SUBT[blk] = f"Biaya!$G${row}"
    row += 1
LVL_COST = {1: f"{SUBT[MA]}", 2: f"({SUBT[MA]}+{SUBT[MB]})", 3: f"({SUBT[MA]}+{SUBT[MB]}+{SUBT[MC]})"}
for lvl, name, blks in [(1, "A", [MA]), (2, "B", [MA, MB]), (3, "C", [MA, MB, MC])]:
    put(B, f"C{row}", f"Biaya pemetaan level {name} (kumulatif)", BOLD)
    put(B, f"G{row}", "=" + LVL_COST[lvl], BOLD, RP)
    row += 1
    put(B, f"C{row}", f"Jam psikolog level {name}")
    crit = "+".join(f'($B${first}:$B${last}="{b_}")' for b_ in blks)
    put(B, f"G{row}", f'=SUMPRODUCT(({crit})*(($F${first}:$F${last}={r("PDOC")})+($F${first}:$F${last}={r("PF2F")}))*$D${first}:$D${last})', BLACK, "#,##0")
    row += 1
row += 2
put(B, f"C{row}", "BIAYA ON DEMAND (per unit layanan)", BOLD)
ONDEMAND_COST = {}
for key, label, formula in [
    ("C_BRF", "Sesi penjelasan hasil (per sesi)", f"={r('BRF_P')}*{r('BRF_H')}*{r('PF2F')}"),
    ("C_FB", "Sesi umpan balik individu (per orang)", f"={r('FB_H')}*{r('PF2F')}"),
    ("C_UNIT", "Unit tambahan di Paket 3 (per unit)", f"=({r('AN_U')}+{r('WR_U')})*{r('PDOC')}+({r('CAL_P_B')}+{r('CAL_P_C')})*{r('CAL_H')}*{r('PF2F')}"),
    ("C_JF", "Rumpun jabatan tambahan di Paket 3 (per rumpun)", f"={r('JA_F2F')}*{r('PF2F')}+({r('JA_DOC')}+{r('JA_RV')})*{r('PDOC')}"),
]:
    row += 1
    put(B, f"C{row}", label)
    put(B, f"G{row}", formula, BLACK, RP)
    ONDEMAND_COST[key] = f"Biaya!$G${row}"
widths(B, [5, 22, 50, 11, 11, 13, 15, 11])

# ---------------------------------------------------------------- Harga
H = wb.create_sheet("Harga")
H["A1"] = "HARGA JUAL (biru = input, hitam = rumus). Hanya harga di tab ini yang boleh keluar ke BSN."
H["A1"].font = TITLE
header(H, 2, ["Key", "Komponen harga", "Nilai", "Satuan", "Catatan"])


def tier_formula(prefix):
    n = r("N")
    return (f"=IF({n}>=2000,{r(prefix + '2000')},IF({n}>=1000,{r(prefix + '1000')},"
            f"IF({n}>=500,{r(prefix + '500')},IF({n}>=100,{r(prefix + '100')},{r(prefix + 'LIST')}))))")


PRICE_ROWS = [
    # Paket 1 tes saja
    ("SLIST", "Paket 1 Laporan untuk Perusahaan: harga normal (di bawah 100 peserta)", 50000, "Rp/peserta", "Harga coret", True, RP),
    ("S100", "Paket 1 tier mulai 100 peserta", 35000, "Rp/peserta", "", True, RP),
    ("S500", "Paket 1 tier mulai 500 peserta", 30000, "Rp/peserta", "", True, RP),
    ("S1000", "Paket 1 tier mulai 1.000 peserta", 27500, "Rp/peserta", "", True, RP),
    ("S2000", "Paket 1 tier mulai 2.000 peserta", 25000, "Rp/peserta", "Laporan otomatis hanya dibaca perusahaan. 20x biaya AI", True, RP),
    ("P_S", "Paket 1 harga untuk BSN", lambda: tier_formula("S"), "Rp/peserta", "Otomatis dari jumlah peserta", False, RP),
    # Paket 2 tes + laporan
    ("ALIST", "Paket 2 Laporan untuk Perusahaan dan Karyawan: harga normal", 75000, "Rp/peserta", "Harga coret", True, RP),
    ("A100", "Paket 2 tier mulai 100 peserta", 50000, "Rp/peserta", "Keputusan Galih", True, RP),
    ("A500", "Paket 2 tier mulai 500 peserta", 45000, "Rp/peserta", "Sama dengan Rekrutiva reguler", True, RP),
    ("A1000", "Paket 2 tier mulai 1.000 peserta", 40000, "Rp/peserta", "", True, RP),
    ("A2000", "Paket 2 tier mulai 2.000 peserta", 35000, "Rp/peserta", "DIKUNCI Galih 4 Nov 2026", True, RP),
    ("P_A", "Paket 2 harga untuk BSN", lambda: tier_formula("A"), "Rp/peserta", "Otomatis dari jumlah peserta", False, RP),
    # Serentak
    ("P_SER", "Tambahan mode serentak, semua paket", 15000, "Rp/peserta", "Tidak didiskon: biaya kas nyata", True, RP),
    # Paket 3
    ("MAP_M", "Multiple harga normal pemetaan terhadap biaya", 3, "x", "Aturan TJ minimal 3x", True, MULT),
    ("P_MAPL1", "Biaya pemetaan level A (Pemetaan Umum): harga normal",
     lambda: f"=ROUND({LVL_COST[1]}*{r('MAP_M')}/500000,0)*500000", "Rp/paket", "", False, RP),
    ("P_MAPL2", "Biaya pemetaan level B (Pemetaan + Talent Box): harga normal",
     lambda: f"=ROUND({LVL_COST[2]}*{r('MAP_M')}/500000,0)*500000", "Rp/paket", "", False, RP),
    ("P_MAPL3", "Biaya pemetaan level C (Pemetaan Talenta Lengkap): harga normal",
     lambda: f"=ROUND({LVL_COST[3]}*{r('MAP_M')}/500000,0)*500000", "Rp/paket", "", False, RP),
    ("D_MAP", "Diskon pesanan pertama pemetaan", 0.15, "%", "Kecil: honor psikolog kas keluar nyata", True, PCT),
    ("P_MAP1", "Biaya pemetaan level A: pesanan pertama",
     lambda: f"=ROUND({r('P_MAPL1')}*(1-{r('D_MAP')})/500000,0)*500000", "Rp/paket", "", False, RP),
    ("P_MAP2", "Biaya pemetaan level B: pesanan pertama",
     lambda: f"=ROUND({r('P_MAPL2')}*(1-{r('D_MAP')})/500000,0)*500000", "Rp/paket", "", False, RP),
    ("P_MAP3", "Biaya pemetaan level C: pesanan pertama",
     lambda: f"=ROUND({r('P_MAPL3')}*(1-{r('D_MAP')})/500000,0)*500000", "Rp/paket", "", False, RP),
    # On demand
    ("OD_M", "Multiple harga on demand terhadap biaya", 3, "x", "", True, MULT),
    ("P_BRF", "On demand: sesi penjelasan hasil 90 menit, 2 psikolog, online",
     lambda: f"=CEILING({ONDEMAND_COST['C_BRF']}*{r('OD_M')},500000)", "Rp/sesi", "Bukan gratis", False, RP),
    ("P_FB", "On demand: sesi umpan balik individu 30 menit", 550000, "Rp/orang", "", True, RP),
    ("P_UNIT", "On demand: unit tambahan di Paket 3",
     lambda: f"=CEILING({ONDEMAND_COST['C_UNIT']}*{r('OD_M')},500000)", "Rp/unit", "", False, RP),
    ("P_JF", "On demand: rumpun jabatan tambahan di Paket 3",
     lambda: f"=CEILING({ONDEMAND_COST['C_JF']}*{r('OD_M')},500000)", "Rp/rumpun", "", False, RP),
]
for i, (key, label, value, unit, note, is_input, fmt) in enumerate(PRICE_ROWS, start=3):
    put(H, f"A{i}", key)
    put(H, f"B{i}", label)
    put(H, f"C{i}", value() if callable(value) else value, BLUE if is_input else BLACK, fmt)
    put(H, f"D{i}", unit)
    put(H, f"E{i}", note)
    H[f"B{i}"].alignment = WRAP
    H[f"E{i}"].alignment = WRAP
    REF[key] = f"Harga!$C${i}"
widths(H, [9, 60, 15, 12, 56])

# ---------------------------------------------------------------- Opsi (3 paket x 2 mode)
O = wb.create_sheet("Opsi")
O["A1"] = "OPSI: 5 paket x 2 mode (jendela 7 hari / serentak). Pemetaan level 0 = tanpa, 1 = A, 2 = B, 3 = C."
O["A1"].font = TITLE
OPTS = [
    ("1. Laporan perusahaan, jendela", 0, 0, 0), ("1. Laporan perusahaan, serentak", 1, 0, 0),
    ("2. Laporan perusahaan + karyawan, jendela", 0, 1, 0), ("2. Laporan perusahaan + karyawan, serentak", 1, 1, 0),
    ("3A. Pemetaan Umum, jendela", 0, 1, 1), ("3A. Pemetaan Umum, serentak", 1, 1, 1),
    ("3B. + Talent Box, jendela", 0, 1, 2), ("3B. + Talent Box, serentak", 1, 1, 2),
    ("3C. Lengkap, jendela", 0, 1, 3), ("3C. Lengkap, serentak", 1, 1, 3),
]
header(O, 2, ["Baris"] + [o[0] for o in OPTS])
LINES = [
    "Serentak? (1 = ya)", "Karyawan ikut membaca laporan? (1 = ya)", "Level pemetaan (0-3)",
    "Harga total (DPP)", "Harga per peserta", "Harga normal", "Potongan (Rp)", "Potongan (%)",
    "PPN (0 jika Non-PKP)", "Ditagihkan ke BSN", "PPh 23 dipotong", "Diterima setelah pajak",
    "Biaya tes", "Biaya laporan AI", "Biaya tambahan serentak", "Biaya pemetaan talenta",
    "Jumlah biaya", "Profit kotor", "Komisi", "PROFIT NETT", "Margin nett terhadap harga",
    "Harga / biaya (multiple)", "% opex TJ 1 bulan tertutup",
]
for j, name in enumerate(LINES, start=3):
    put(O, f"A{j}", name, BOLD if name.isupper() else BLACK)
L = {name: j for j, name in enumerate(LINES, start=3)}
for k, (name, ser, ai, mp) in enumerate(OPTS):
    c = get_column_letter(k + 2)

    def cell(line):
        return f"{c}{L[line]}"

    put(O, cell("Serentak? (1 = ya)"), ser, BLUE, "0")
    put(O, cell("Karyawan ikut membaca laporan? (1 = ya)"), ai, BLUE, "0")
    put(O, cell("Level pemetaan (0-3)"), mp, BLUE, "0")
    S, AIc, M = cell("Serentak? (1 = ya)"), cell("Karyawan ikut membaca laporan? (1 = ya)"), cell("Level pemetaan (0-3)")
    map_price = f"CHOOSE({M}+1,0,{r('P_MAP1')},{r('P_MAP2')},{r('P_MAP3')})"
    map_list = f"CHOOSE({M}+1,0,{r('P_MAPL1')},{r('P_MAPL2')},{r('P_MAPL3')})"
    map_cost = f"CHOOSE({M}+1,0,{LVL_COST[1]},{LVL_COST[2]},{LVL_COST[3]})"
    test_price = f"IF({AIc}=1,{r('P_A')},{r('P_S')})"
    test_list = f"IF({AIc}=1,{r('ALIST')},{r('SLIST')})"
    put(O, cell("Harga total (DPP)"), f"={r('N')}*({test_price}+{S}*{r('P_SER')})+{map_price}", GREEN, RP)
    put(O, cell("Harga per peserta"), f"={cell('Harga total (DPP)')}/{r('N')}", BLACK, RP)
    put(O, cell("Harga normal"), f"={r('N')}*({test_list}+{S}*{r('P_SER')})+{map_list}", GREEN, RP)
    put(O, cell("Potongan (Rp)"), f"={cell('Harga normal')}-{cell('Harga total (DPP)')}", BLACK, RP)
    put(O, cell("Potongan (%)"), f"={cell('Potongan (Rp)')}/{cell('Harga normal')}", BLACK, PCT)
    put(O, cell("PPN (0 jika Non-PKP)"), f"={cell('Harga total (DPP)')}*{r('PPN')}", GREEN, RP)
    put(O, cell("Ditagihkan ke BSN"), f"={cell('Harga total (DPP)')}+{cell('PPN (0 jika Non-PKP)')}", BLACK, RP)
    put(O, cell("PPh 23 dipotong"), f"={cell('Harga total (DPP)')}*{r('PPH23')}", GREEN, RP)
    put(O, cell("Diterima setelah pajak"), f"={cell('Harga total (DPP)')}-{cell('PPh 23 dipotong')}", BLACK, RP)
    put(O, cell("Biaya tes"), f"={SUBT[BLK_BASE]}", GREEN, RP)
    put(O, cell("Biaya laporan AI"), f"={SUBT[BLK_AI]}", GREEN, RP)
    put(O, cell("Biaya tambahan serentak"), f"={S}*{SUBT[BLK_SER]}", GREEN, RP)
    put(O, cell("Biaya pemetaan talenta"), f"={map_cost}", GREEN, RP)
    put(O, cell("Jumlah biaya"), f"=SUM({cell('Biaya tes')}:{cell('Biaya pemetaan talenta')})", BOLD, RP)
    put(O, cell("Profit kotor"), f"={cell('Diterima setelah pajak')}-{cell('Jumlah biaya')}", BLACK, RP)
    put(O, cell("Komisi"), f"=MAX(0,{cell('Profit kotor')})*{r('KOM')}", GREEN, RP)
    put(O, cell("PROFIT NETT"), f"={cell('Profit kotor')}-{cell('Komisi')}", BOLD, RP)
    put(O, cell("Margin nett terhadap harga"), f"={cell('PROFIT NETT')}/{cell('Harga total (DPP)')}", BLACK, PCT)
    put(O, cell("Harga / biaya (multiple)"), f"={cell('Harga total (DPP)')}/{cell('Jumlah biaya')}", BLACK, MULT)
    put(O, cell("% opex TJ 1 bulan tertutup"), f"={cell('PROFIT NETT')}/{r('OPEX')}", BLACK, PCT)
widths(O, [30] + [18] * len(OPTS))

row = L["% opex TJ 1 bulan tertutup"] + 2
put(O, f"A{row}", "CEK PER KOMPONEN: harga / biaya", BOLD)
row += 1
header(O, row, ["Komponen", "Harga", "Biaya", "Multiple", "Aturan TJ >= 3x?"])
for name, p, cst in [
    ("Paket 1 (per proyek)", f"={r('N')}*{r('P_S')}", f"={SUBT[BLK_BASE]}+{SUBT[BLK_AI]}"),
    ("Paket 2 (per proyek)", f"={r('N')}*{r('P_A')}", f"={SUBT[BLK_BASE]}+{SUBT[BLK_AI]}"),
    ("Tambahan serentak", f"={r('N')}*{r('P_SER')}", f"={SUBT[BLK_SER]}"),
    ("Pemetaan A, pesanan pertama", f"={r('P_MAP1')}", f"={LVL_COST[1]}"),
    ("Pemetaan B, pesanan pertama", f"={r('P_MAP2')}", f"={LVL_COST[2]}"),
    ("Pemetaan C, pesanan pertama", f"={r('P_MAP3')}", f"={LVL_COST[3]}"),
    ("On demand: sesi penjelasan hasil", f"={r('P_BRF')}", f"={ONDEMAND_COST['C_BRF']}"),
    ("On demand: umpan balik individu", f"={r('P_FB')}", f"={ONDEMAND_COST['C_FB']}"),
    ("On demand: unit tambahan", f"={r('P_UNIT')}", f"={ONDEMAND_COST['C_UNIT']}"),
    ("On demand: rumpun jabatan tambahan", f"={r('P_JF')}", f"={ONDEMAND_COST['C_JF']}"),
]:
    row += 1
    put(O, f"A{row}", name)
    put(O, f"B{row}", p, GREEN, RP)
    put(O, f"C{row}", cst, GREEN, RP)
    put(O, f"D{row}", f"=B{row}/C{row}", BLACK, MULT)
    put(O, f"E{row}", f'=IF(D{row}>=3,"OK",IF(D{row}>=2,"Di bawah 3x, masih di atas 2x (diskon)","TERLALU RENDAH"))')

# ---------------------------------------------------------------- Tier
TR = wb.create_sheet("Tier")
TR["A1"] = "TIER HARGA (jendela 7 hari). Biaya konservatif: biaya setup dan helpdesk penuh di tiap tier."
TR["A1"].font = TITLE
header(TR, 2, ["Paket", "Peserta", "Harga / peserta", "Potongan dari normal", "Harga total", "Jumlah biaya", "Profit nett", "Margin nett"])
row = 3
tech_var = f"({SUBT[BLK_BASE]}/{r('N')})"  # conservative: whole base cost spread per head
for pkg, prefix, ai in [("Paket 1 Laporan perusahaan", "S", 1), ("Paket 2 Laporan perusahaan + karyawan", "A", 1)]:
    for qty in [100, 500, 1000, 2000]:
        put(TR, f"A{row}", pkg)
        put(TR, f"B{row}", qty, BLUE, "#,##0")
        put(TR, f"C{row}", f"={r(prefix + str(qty))}", GREEN, RP)
        put(TR, f"D{row}", f"=1-C{row}/{r(prefix + 'LIST')}", BLACK, PCT)
        put(TR, f"E{row}", f"=B{row}*C{row}", BLACK, RP)
        cost = (f"={SUBT[BLK_BASE]}-{r('N')}*{r('TIX')}*{r('TIXMIN')}/60*{r('ADM')}"
                f"+B{row}*{r('TIX')}*{r('TIXMIN')}/60*{r('ADM')}+{ai}*B{row}*(1+{r('RERUN')})*{r('AI')}")
        put(TR, f"F{row}", cost, GREEN, RP)
        put(TR, f"G{row}", f"=(E{row}*(1-{r('PPH23')})-F{row})*(1-{r('KOM')})", BOLD, RP)
        put(TR, f"H{row}", f"=G{row}/E{row}", BLACK, PCT)
        row += 1
widths(TR, [22, 10, 15, 18, 16, 15, 15, 11])

# ---------------------------------------------------------------- Kompetitor
K = wb.create_sheet("Kompetitor")
K["A1"] = "5 PESAING (harga publik, dicek 4 Nov 2026). Harga besar biasanya negosiasi, anggap batas atas."
K["A1"].font = TITLE
header(K, 2, ["Pesaing", "Jenis", "Harga publik", "Per peserta (perkiraan)", "BFI + DISC?", "Laporan", "Psikolog?", "Sumber"])
COMP = [
    ("Rekrutiva", "Platform psikotes online", "Rp45 ribu reguler, Rp75 ribu dengan proctoring", 45000,
     "Tidak disebut (20+ tes)", "Psikogram AI, PDF/Excel", "Tidak disebut", "rekrutiva.com/services/e-psychotest"),
    ("Taldio", "Platform + psikolog", "Mulai Rp50 ribu, ada paket unlimited", 50000,
     "Ya (Big Five OCEAN, DISC)", "Real time, bisa dikustom", "Ya, interpretasi psikolog", "taldio.com/blog/biaya-psikotes"),
    ("teskepribadian.id", "Alat tes untuk psikolog", "Rp10 juta untuk 500 responden", 20000,
     "Tidak disebut", "Analisis tipe kepribadian", "Tidak (alat untuk psikolog)", "teskepribadian.id/pricing"),
    ("PsikotesOnline.com", "Langganan platform", "Rp8 juta per tahun, semua tes", None,
     "DISC ya, Big Five tidak jelas", "Unduh laporan", "Tidak (dikerjakan sendiri)", "psikotesonline.com"),
    ("Biro psikologi (MindHarmony, Associe, Solutiva)", "Biro psikologi", "Rp200-250 ribu per orang / per aspek, assessment center mulai Rp1,5 juta", 200000,
     "Tergantung paket", "Laporan psikolog", "Ya", "mindharmonyindonesia.com, associe.co.id, solutiva.co.id"),
]
for i, row_ in enumerate(COMP, start=3):
    for j, v in enumerate(row_, start=1):
        c = put(K, f"{get_column_letter(j)}{i}", v, BLACK, RP if j == 4 and v else None)
        c.alignment = WRAP
widths(K, [26, 20, 30, 16, 22, 24, 22, 34])

# ---------------------------------------------------------------- To_Fill
T = wb.create_sheet("To_Fill")
T["A1"] = "TO FILL - semua estimasi. Isi kolom kuning dengan angka asli, lalu salin ke tab Asumsi."
T["A1"].font = TITLE
header(T, 2, ["Key", "Asumsi", "Nilai sekarang", "Kenapa penting", "Sel Asumsi", "Angka asli"])
WHY = {
    "AI": "Konfirmasi satuan (per laporan atau per tes)",
    "ENG": "Penggerak biaya setup, serentak dan talent box",
    "JF": "Penggerak terbesar harga Paket 3",
    "UNITS": "Penggerak kalibrasi dan laporan per unit",
    "EDGE_B": "Jam review individu di level B",
    "TB_ENG": "Talent box mungkin build baru di platform",
    "PPN": "Entitas mana yang kontrak (TJ Non-PKP atau Coreitera)",
}
row = 3
for key, label, value, unit, status, source, fmt in ASSUMPTIONS:
    if status == "Known":
        continue
    put(T, f"A{row}", key)
    put(T, f"B{row}", label)
    put(T, f"C{row}", f"={r(key)}", GREEN, fmt)
    put(T, f"D{row}", WHY.get(key, "Estimasi"))
    put(T, f"E{row}", r(key).replace("$", ""))
    put(T, f"F{row}", None, BLUE, fmt, EST_FILL)
    T[f"B{row}"].alignment = WRAP
    T[f"D{row}"].alignment = WRAP
    row += 1
widths(T, [9, 50, 14, 46, 14, 14])

wb.move_sheet("Opsi", offset=-3)
wb.active = 0
wb.save(OUT)
print(OUT)
