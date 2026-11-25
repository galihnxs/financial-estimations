# Adds sheet "Penawaran Akhir 24 Nov" to the CONFIDENTIAL BSN RAB. Live formulas; existing sheets untouched.
# Mirrors build_update_deck.py (24 Nov 2026, v3): pendampingan Rp75 jt with 4 deliverables, test option A (ours,
# Rp40rb x 3) vs option B (Faxtor BFI + FTPI + OPTI at 31% net after tax).
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

F = "[CONFIDENTIAL] RAB Teman Journey x BSN.xlsx"
NAME = "Penawaran Akhir 24 Nov"
wb = openpyxl.load_workbook(F)
if NAME in wb.sheetnames:
    del wb[NAME]
ws = wb.create_sheet(NAME, 0)
RP = '"Rp"#,##0'; PCT = '0.0%'
B = Font(name="Arial", bold=True); N = Font(name="Arial"); H = Font(name="Arial", bold=True, size=13, color="104A70")
BLUE = Font(name="Arial", color="0000FF"); NOTE = Font(name="Arial", italic=True, color="666666", size=9)
HEAD = PatternFill("solid", fgColor="E8D4AB"); TOT = PatternFill("solid", fgColor="F6EDE2"); INP = PatternFill("solid", fgColor="FFF2CC")
thin = Border(bottom=Side(style="thin", color="BBBBBB"))
for col, w in zip("ABCDEFGH", (5, 50, 16, 9, 17, 17, 12, 62)):
    ws.column_dimensions[col].width = w


def put(row, vals, font=N, fill=None, fmt=None):
    for i, v in enumerate(vals):
        if v is None:
            continue
        c = ws.cell(row=row, column=i + 1, value=v); c.font = font
        if fill:
            c.fill = fill
        if fmt and i in fmt:
            c.number_format = fmt[i]


def head(row, vals):
    put(row, vals, B, HEAD)
    for i in range(len(vals)):
        ws.cell(row=row, column=i + 1).border = thin


put(1, [None, "RAB PENAWARAN AKHIR BSN (24 Nov 2026, v3), INTERNAL"], H)
put(2, [None, "Mengikuti deck 'Pembaruan Akhir' 18 halaman. Sheet MR DIY MHCU dibiarkan sebagai versi 4 Nov. Sel biru = asumsi yang boleh diubah."], NOTE)

head(4, [None, "ASUMSI", "Nilai", None, None, None, None, "Keterangan"])
ASM = [
    ("Margin bersih korporat minimum (layanan TJ)", 0.50, PCT, "Aturan korporat: layanan TJ >= 50% net"),
    ("Margin bersih tes mitra (Faxtor)", 0.31, PCT, "Keputusan Galih 24 Nov: margin rendah di tes, untung dari pendampingan dan layanan lain"),
    ("PPh 23 dipotong klien", 0.02, PCT, "Diterima 98% dari harga"),
    ("Komisi partner (% profit)", 0.10, PCT, "Sama dengan H5 sheet lama"),
    ("Tarif psikolog tatap muka / PIO per jam", 350000, RP, "Sama dengan tarif konseling"),
    ("Tarif kerja dokumen per jam", 110000, RP, "Olah data, laporan"),
    ("Jam telaah psikolog PIO", 20, "0", "1 sampai 2 psikolog, total 20 jam"),
    ("Jumlah peserta", 300, "0", "Pemetaan hingga 300 karyawan"),
    ("Faxtor BFI per peserta", 30000, RP, "Pricelist Faxtor 2026, belum pajak"),
    ("Faxtor FTPI per peserta", 30000, RP, "Pricelist Faxtor 2026, belum pajak"),
    ("Faxtor OPTI per peserta", 50000, RP, "Pricelist Faxtor 2026, belum pajak"),
    ("Faxtor rekap skor Excel per peserta", 35000, RP, "Add-on, dari pesan Faxtor (tidak ada di pricelist)"),
    ("Pajak atas harga Faxtor", 0.11, PCT, "TJ Non-PKP, pajak masukan tidak bisa dikreditkan"),
    ("Harga tes kami per alat tes (Pilihan A)", 40000, RP, "Keputusan Galih 24 Nov. TJ / Coreitera Non-PKP: tidak ada PPN ditambahkan"),
    ("Jumlah alat tes kami", 3, "0", "DISC, BFI, gaya kepemimpinan"),
]
ref = {}
for i, (k, v, fmt, note) in enumerate(ASM):
    r = 5 + i
    put(r, [None, k, v, None, None, None, None, note], N, None, {2: fmt})
    ws.cell(row=r, column=3).font = BLUE; ws.cell(row=r, column=3).fill = INP
    ref[i] = f"$C${r}"
MARG, MMIT, PPH, KOM, RATE, DOC, JAM, PES, FBFI, FFTPI, FOPTI, FREC, FTAX, OURP, OURN = (ref[i] for i in range(15))
KEEP = f"(1-{PPH})*(1-{KOM})"
NETM = lambda cost, price: f"={KEEP}-(1-{KOM})*{cost}/{price}"
r = 5 + len(ASM)
put(r, [None, "Rumus harga: biaya x (1 - komisi) / ((1 - PPh23) x (1 - komisi) - margin). Margin bersih = (1 - PPh23)(1 - komisi) - (1 - komisi) x biaya / harga"], NOTE)
r += 2

# 1 Pendampingan
put(r, [None, "1. PENDAMPINGAN PEMETAAN TALENTA (untuk Pilihan A dan B, tidak menunggu lisensi)"], H); r += 1
head(r, ["No", "Item biaya", "Harga satuan", "Qty", "Jumlah", None, None, "Keterangan"]); r += 1
first = r
for i, (item, unit, qty, note) in enumerate([
    ("Group report agregat & HR advisory PIO 3 sesi", 15000000, 1, "Baris RAB lama (C5). Sesi alignment, pemantauan, finalisasi + group report"),
    ("Telaah setiap peserta oleh psikolog PIO", f"={RATE}", f"={JAM}", "Rekomendasi per peserta. 20 jam = 4 menit per peserta"),
    ("Project Coordinator", 1000000, 1, "Baris RAB lama (C3)"),
    ("Talent box: olah data kinerja BSN", f"={DOC}", 8, "Gabung potensi (hasil tes) dan kinerja (data BSN), 9 kotak"),
    ("Kalibrasi talent box bersama HC", f"={RATE}", 4, "1 sesi online 2 jam x 2 psikolog"),
    ("Talent analytics per kelompok", f"={DOC}", 10, "Pola per unit + rekomendasi tindak lanjut")]):
    put(r, [i + 1, item, unit, qty, f"=C{r}*D{r}", None, None, note], N, None, {2: RP, 4: RP}); r += 1
put(r, [None, "TOTAL BIAYA PENDAMPINGAN", None, None, f"=SUM(E{first}:E{r - 1})"], B, TOT, {4: RP}); cost = f"E{r}"; r += 1
put(r, [None, "Harga jual (keputusan Galih)", None, None, 75000000, None, None, "Dipilih Galih 24 Nov dari opsi 75 / 85 / 95 jt"], B, None, {4: RP})
ws.cell(row=r, column=5).font = Font(name="Arial", bold=True, color="0000FF"); ws.cell(row=r, column=5).fill = INP; price = f"E{r}"; r += 1
put(r, [None, "Margin bersih setelah PPh 23 dan komisi", None, None, NETM(cost, price)], B, None, {4: PCT}); r += 1
put(r, [None, "Lolos aturan 50%?", None, None, f'=IF(E{r - 1}>={MARG},"Ya","TIDAK")'], N); r += 1
put(r, [None, "Komisi partner", None, None, f"=({price}*(1-{PPH})-{cost})*{KOM}"], N, None, {4: RP}); r += 1
put(r, [None, "Profit nett", None, None, f"=({price}*(1-{PPH})-{cost})*(1-{KOM})"], B, None, {4: RP}); pend_profit = f"E{r}"; r += 1
put(r, [None, "Luaran: 1 mapping alignment, 2 rekomendasi + group report, 3 talent box, 4 talent analytics + tindak lanjut per kelompok. BSN menyiapkan data penilaian kinerja."], NOTE); r += 2

# 2 Pilihan A
put(r, [None, "2. PILIHAN A: TES DARI KAMI (DISC, BFI, gaya kepemimpinan), setelah lisensi"], H); r += 1
head(r, ["No", "Item", "Harga satuan", "Qty", "Jumlah", None, None, "Keterangan"]); r += 1
put(r, [1, "Harga per karyawan", f"={OURP}", f"={OURN}", f"=C{r}*D{r}", None, None, "Rp40rb per tes"], B, None, {2: RP, 4: RP}); a_price = f"E{r}"; r += 1
put(r, [2, "Total tes untuk peserta", None, f"={PES}", f"={a_price}*D{r}"], B, TOT, {4: RP}); a_total = f"E{r}"; r += 1
put(r, [None, "Biaya lisensi per peserta belum diketahui (negosiasi LPT UI / UGM). Margin Pilihan A belum bisa dihitung. Jangan sebut LPT UI / UGM ke klien."], NOTE); r += 2

# 3 Pilihan B
put(r, [None, "3. PILIHAN B: TES MELALUI FAXTOR (BFI, FTPI, OPTI), ditagih TJ"], H); r += 1
head(r, ["No", "Item", "Harga satuan", "Qty", "Jumlah", "Margin bersih", None, "Keterangan"]); r += 1
put(r, [1, "Tes BFI + FTPI + OPTI per peserta", f"={FBFI}+{FFTPI}+{FOPTI}", 1, f"=C{r}*D{r}", None, None, "Harga list Faxtor"], N, None, {2: RP, 4: RP}); f1 = r; r += 1
put(r, [2, "Rekap skor Excel per peserta", f"={FREC}", 1, f"=C{r}*D{r}", None, None, "Dibutuhkan untuk talent box dan analytics"], N, None, {2: RP, 4: RP}); r += 1
put(r, [3, "Pajak atas tagihan Faxtor", None, None, f"=(E{f1}+E{f1 + 1})*{FTAX}", None, None, "Tidak bisa dikreditkan (Non-PKP)"], N, None, {4: RP}); r += 1
put(r, [None, "Biaya per peserta", None, None, f"=SUM(E{f1}:E{r - 1})"], B, TOT, {4: RP}); b_cost = f"E{r}"; r += 1
put(r, [None, "Harga jual per peserta (dibulatkan ke atas Rp5rb)", None, None,
        f"=ROUNDUP({b_cost}*(1-{KOM})/({KEEP}-{MMIT})/5000,0)*5000", NETM(b_cost, f"E{r}"), None, "Yang tampil di deck"], B, None, {4: RP, 5: PCT}); b_price = f"E{r}"; r += 1
put(r, [None, "Total biaya tes untuk peserta", None, f"={PES}", f"={b_cost}*D{r}", None, None, "Dibayar TJ ke Faxtor"], N, None, {4: RP}); r += 1
put(r, [None, "Total harga tes untuk peserta", None, f"={PES}", f"={b_price}*D{r}", None, None, "Ditagih TJ ke BSN"], B, TOT, {4: RP}); b_total = f"E{r}"; r += 1
put(r, [None, "Profit nett tes", None, None, f"=({b_total}*(1-{PPH})-{b_cost}*{PES})*(1-{KOM})"], N, None, {4: RP}); b_profit = f"E{r}"; r += 1
put(r, [None, "Proktoring dan batas akses 23.59 WIB: belum tertulis di dokumen Faxtor, konfirmasi tertulis sebelum dijanjikan."], NOTE); r += 2

# 4 Tambahan kandidat
put(r, [None, "4. FIT AND PROPER TEST UNTUK KANDIDAT TERATAS (per 6 kandidat, harga keputusan Galih 24 Nov)"], H); r += 1
head(r, ["No", "Layanan", "Biaya", None, "Harga jual", "Margin bersih", None, "Keterangan"]); r += 1
for i, (item, c, p, note) in enumerate([
    ("Fit and Proper Test, 6 kandidat", 8770000 - 6 * 120000, 15000000, "Biaya RAB BI Fit and Proper Rp8,77 jt dikurangi tes BFI+DISC 6 x 120rb (pakai hasil tes pemetaan). DI BAWAH LANTAI 50%, atas keputusan Galih")]):
    put(r, [i + 1, item, c, None, p, NETM(f"C{r}", f"E{r}"), None, note], N, None, {2: RP, 4: RP, 5: PCT}); r += 1
r += 1

# 5 Katalog
put(r, [None, "5. LAYANAN KESEHATAN MENTAL (harga katalog saat ini, di luar lingkup pemetaan)"], H); r += 1
head(r, ["No", "Layanan", "Harga", None, "Satuan", None, None, "Isi"]); r += 1
for i, (item, p, unit, note) in enumerate([
    ("MHCU Skrining", 175000, "per orang", "DASS-21 online, ditinjau psikolog, laporan individu paling lambat 48 jam"),
    ("MHCU Flash", 600000, "per orang", "5 instrumen klinis, laporan personal, konseling 30 menit"),
    ("MHCU Deep", 1000000, "per orang", "5 instrumen klinis, laporan personal, konseling 60 menit"),
    ("Program bantuan karyawan (EAP)", 13500000, "per bulan", "300 karyawan: kuota 10 sesi per bulan, MHCU 2x setahun, laporan agregat, webinar per kuartal"),
    ("Paket 10 sesi konseling (60 menit)", 10500000, "per paket", "10 sesi online + 10 MHCU, 6 bulan"),
    ("Paket 10 sesi konseling (30 menit)", 6200000, "per paket", "Versi sesi 30 menit"),
    ("Workshop lingkungan kerja sehat", 8100000, "per sesi", "2 jam tatap muka, 30 peserta")]):
    put(r, [i + 1, item, p, None, unit, None, None, note], N, None, {2: RP}); r += 1
r += 1

# 6 Pilot (paid, no free offers)
put(r, [None, "6. LAYANAN PILOT (berbayar, tidak ada layanan gratis)"], H); r += 1
head(r, ["No", "Item", "Biaya", "Qty", "Harga", "Margin net", None, "Keterangan"]); r += 1
g0 = r
put(r, [1, "MHCU Flash untuk tim HC (per orang)", 244500, 5, 600000, NETM(f"C{r}", f"E{r}"), None, "Biaya dari ProductKnowledge 05; total 5 orang Rp3.000.000"], N, None, {2: RP, 4: RP, 5: PCT}); r += 1
put(r, [2, "Snapshot burnout satu unit kerja (sampai 50 orang)", "=6*110000+1*350000", 1, 2500000, NETM(f"C{r}", f"E{r}"), None, "Laporan agregat 6 jam kerja dokumen + paparan online 1 jam psikolog"], N, None, {2: RP, 4: RP, 5: PCT}); r += 1
put(r, [None, "TOTAL BILA KEDUANYA", None, None, f"=E{g0}*D{g0}+E{g0 + 1}*D{g0 + 1}"], B, TOT, {4: RP}); r += 2

# Ringkasan A vs B
put(r, [None, "RINGKASAN: PILIHAN A vs PILIHAN B (300 karyawan)"], H); r += 1
head(r, [None, "Item", None, None, "Pilihan A", "Pilihan B", None, "Keterangan"]); r += 1
put(r, [None, "Tes", None, None, f"={a_total}", f"={b_total}"], N, None, {4: RP, 5: RP}); t = r; r += 1
put(r, [None, "Pendampingan", None, None, f"={price}", f"={price}"], N, None, {4: RP, 5: RP}); r += 1
put(r, [None, "Total ke BSN", None, None, f"=E{t}+E{t + 1}", f"=F{t}+F{t + 1}"], B, TOT, {4: RP, 5: RP}); r += 1
put(r, [None, "Profit nett TJ", None, None, f"={pend_profit}", f"={pend_profit}+{b_profit}", None, "Pilihan A: profit tes belum dihitung (biaya lisensi belum diketahui)"], B, None, {4: RP, 5: RP}); r += 1
ws.freeze_panes = "A4"
wb.save(F)
print("rows", r)
