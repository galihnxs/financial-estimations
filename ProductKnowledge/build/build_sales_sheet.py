"""Build Kalkulator_Sales.xlsx for the sales team from Master_Produk_dan_Harga.xlsx.

Run:  /opt/homebrew/bin/python3 -I build_sales_sheet.py   (needs openpyxl and formulas)

The script recalculates the master workbook (formulas package), reads one cost total per product, the segment flags and
today's prices, and writes a sales workbook where every price, margin and commission is a formula. No cost breakdown or
rate table is copied: only one cost total per product, on a hidden, protected sheet.
"""
import os
import shutil
import tempfile

import formulas
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.protection import WorkbookProtection

HERE = os.path.dirname(os.path.abspath(__file__))
MASTER = os.path.join(HERE, "..", "Master_Produk_dan_Harga.xlsx")
OUT = os.path.join(HERE, "..", "Kalkulator_Sales.xlsx")
PROTECT_PW = "tj-sales-2026"  # sheet and structure protection; not a security control, only stops accidental edits

# ------------------------------------------------------------------ read the recalculated master
tmp = tempfile.mkdtemp()
xl = formulas.ExcelModel().loads(MASTER).finish()
xl.calculate()
xl.write(dirpath=tmp)
calc = os.path.join(tmp, os.listdir(tmp)[0])
mw = openpyxl.load_workbook(calc, data_only=True)
kat = mw["KATALOG"]
inp = mw["INPUT"]
params = {}
for r in inp.iter_rows(min_row=5, max_row=11, values_only=True):
    if r[0]:
        params[r[0]] = r[2]
for r in inp.iter_rows(min_row=15, max_row=17, values_only=True):
    if r[0]:
        params[r[0]] = r[2]
prods = []
for r in kat.iter_rows(min_row=6, values_only=True):
    if not r[0]:
        continue
    prods.append(dict(code=r[0], line=r[2], prod=r[3], scope=r[4], uom=r[6], role=r[8], rule=r[9], pph=r[10],
                      edu=r[11], korp=r[12], reg=r[13], frozen=r[14], cur=r[15], cost=r[16]))
shutil.rmtree(tmp, ignore_errors=True)
assert len(prods) > 50, "master workbook not read"

# ------------------------------------------------------------------ styles
F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F)
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
TITLE = Font(name=F, bold=True, size=14)
HDR_FONT = Font(name=F, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1C4587")
SAND = PatternFill("solid", fgColor="F3E9D2")
INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")
RP = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
G_FILL = PatternFill("solid", fgColor="C6EFCE")
Y_FILL = PatternFill("solid", fgColor="FFEB9C")
R_FILL = PatternFill("solid", fgColor="FFC7CE")

wb = Workbook()


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


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


def header(ws, row, labels, widths):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.row_dimensions[row].height = 42


def rnd(x):
    step = f"IF({x}<100000,STEP0,IF({x}<1000000,STEP1,IF({x}<10000000,STEP2,STEP3)))"
    return f"ROUNDUP({x}/{step},0)*{step}"


# ------------------------------------------------------------------ hidden settings and data
wp = wb.active
wp.title = "Pengaturan"
put(wp, "A1", "Pengaturan (dari Master_Produk_dan_Harga.xlsx). Jangan diubah di sini; ubah di master lalu bangun ulang.", BOLD)
settings = [("PPH", "PPh 23 dipotong klien", params["TARIF_PPH23"], PCT),
            ("KOMISI", "Komisi sales dari laba kotor", params["KOMISI"], PCT),
            ("T_EDU", "Target margin Pendidikan", params["SEG_EDU"], PCT),
            ("T_KORP", "Target margin Korporat", params["SEG_KORP"], PCT),
            ("T_REG", "Target margin Regulator", params["SEG_REG"], PCT),
            ("T_MIN", "Batas bawah margin semua segmen", 0.30, PCT),
            ("STEP0", "Pembulatan di bawah Rp100rb", params["STEP0"], RP),
            ("STEP1", "Pembulatan Rp100rb-1 jt", params["STEP1"], RP),
            ("STEP2", "Pembulatan Rp1-10 jt", params["STEP2"], RP),
            ("STEP3", "Pembulatan di atas Rp10 jt", params["STEP3"], RP),
            ("EDU_MIN", "Minimal kontrak Pendidikan", 10000000, RP),
            ("SHARE_MAX", "Bagi hasil sekolah / mitra maksimal", 0.10, PCT)]
for i, (k, lab, v, fmt) in enumerate(settings):
    rr = 3 + i
    put(wp, f"A{rr}", k, BOLD)
    put(wp, f"B{rr}", lab)
    put(wp, f"C{rr}", v, BLACK, fmt)
    name(k, f"Pengaturan!$C${rr}")
wp.column_dimensions["B"].width = 40

wd = wb.create_sheet("Data")
dcols = ["Kode", "Lini", "Produk", "Isi", "Satuan", "Peran", "Aturan", "Potong PPh?", "Pendidikan?", "Korporat?",
         "Regulator?", "Regulator dibekukan?", "Harga saat ini", "Biaya total per satuan"]
header(wd, 1, dcols, [13, 26, 36, 40, 14, 11, 11, 9, 9, 9, 9, 10, 14, 14])
D1 = 2
for i, p in enumerate(prods):
    rr = D1 + i
    vals = [p["code"], p["line"], p["prod"], p["scope"], p["uom"], p["role"], p["rule"], p["pph"], p["edu"], p["korp"],
            p["reg"], p["frozen"], p["cur"] if isinstance(p["cur"], (int, float)) else 0, round(p["cost"] or 0, 2)]
    for j, v in enumerate(vals):
        c = wd.cell(row=rr, column=j + 1, value=v)
        c.font = BLACK
        if j >= 12:
            c.number_format = RP
DN = D1 + len(prods) - 1
for nm, col in (("D_CODE", "A"), ("D_LINE", "B"), ("D_PROD", "C"), ("D_SCOPE", "D"), ("D_UOM", "E"), ("D_ROLE", "F"),
                ("D_RULE", "G"), ("D_PPH", "H"), ("D_EDU", "I"), ("D_KORP", "J"), ("D_REG", "K"), ("D_FROZEN", "L"),
                ("D_CUR", "M"), ("D_COST", "N")):
    name(nm, f"Data!${col}${D1}:${col}${DN}")

# ------------------------------------------------------------------ Harga (price list)
wh = wb.create_sheet("Harga", 0)
put(wh, "A1", "Daftar harga per segmen klien", TITLE)
put(wh, "A2", "Harga belum termasuk PPN. Pilih kolom sesuai jenis klien. Harga minimum = harga terendah yang masih memberi margin "
              "bersih 30% (setelah PPh 23 dan komisi). Di bawah harga minimum: tidak boleh tanpa Direktur. Kosong = tidak dijual ke segmen itu.")
wh.merge_cells("A2:N2")
wh["A2"].alignment = WRAP
wh.row_dimensions[2].height = 32
hcols = ["Kode", "Lini", "Produk", "Isi", "Satuan", "Peran", "Harga Pendidikan", "Harga Korporat", "Harga Regulator",
         "Harga umum (satu harga)", "Harga minimum (margin 30%)", "Margin di harga Pendidikan", "Margin di harga Korporat",
         "Status"]
header(wh, 4, hcols, [12, 24, 34, 46, 14, 11, 15, 15, 15, 15, 15, 12, 12, 26])
H1 = 5
for i in range(len(prods)):
    rr, dr = H1 + i, D1 + i
    put(wh, f"A{rr}", f"=Data!A{dr}", BOLD)
    for col, dcol in zip("BCDEF", "BCDEF"):
        put(wh, f"{col}{rr}", f"=Data!{dcol}{dr}", BLACK, align=WRAP if col in "CD" else None)
    cost, rule, tp = f"Data!N{dr}", f"Data!G{dr}", f'IF(Data!H{dr}="Ya",PPH,0)'
    cur = f"Data!M{dr}"
    kom = f'IF(OR({rule}="Penuh",{rule}="Langganan"),KOMISI,0)'

    def raw(m):
        return f"({cost}*(1-KOMISI)/((1-{tp})*(1-KOMISI)-{m}))"

    put(wh, f"G{rr}", f'=IF(AND({rule}="Penuh",Data!I{dr}="Ya"),{rnd(raw("T_EDU"))},"")', BOLD, RP, fill=SAND)
    put(wh, f"H{rr}", f'=IF(AND({rule}="Penuh",Data!J{dr}="Ya"),MAX({cur},{rnd(raw("T_KORP"))}),"")', BOLD, RP, fill=SAND)
    put(wh, f"I{rr}", f'=IF(AND({rule}="Penuh",Data!K{dr}="Ya"),IF(AND(Data!L{dr}="Ya",{cur}>0),{cur},MAX({cur},{rnd(raw("T_REG"))})),"")',
        BOLD, RP, fill=SAND)
    put(wh, f"J{rr}", f'=IF({rule}="Penuh","",IF({rule}="Pintu masuk",0,IF({rule}="Pasar",{cur},IF({rule}="Harga pokok",'
                      f'IF({cur}>0,{cur},{rnd(f"({cost}/(1-{tp}))")}),MAX({cur},{rnd(raw("T_MIN"))})))))', BOLD, RP, fill=SAND)
    put(wh, f"K{rr}", f'=IF(AND({rule}="Penuh",{cost}>0),{rnd(raw("T_MIN"))},"")', BLACK, RP)

    def marg(pc):
        gp = f"({pc}*(1-{tp})-{cost})"
        return f'=IF(ISNUMBER({pc}),IF({pc}>0,({gp}-MAX(0,{gp}*{kom}))/{pc},""),"")'

    put(wh, f"L{rr}", marg(f"G{rr}"), BLACK, PCT)
    put(wh, f"M{rr}", marg(f"H{rr}"), BLACK, PCT)
    put(wh, f"N{rr}", f'=IF(Data!F{dr}="Pintu masuk","Pintu masuk: boleh gratis",IF({rule}<>"Penuh","Satu harga",'
                      f'IF(AND(ISNUMBER(I{rr}),Data!L{dr}="Ya",{cur}>0,{cur}<K{rr}),"Regulator lama: perlu Direktur","Sesuai aturan")))')
HN = H1 + len(prods) - 1
wh.freeze_panes = "D5"
wh.auto_filter.ref = f"A4:N{HN}"

# ------------------------------------------------------------------ Penawaran (offer builder)
wq = wb.create_sheet("Penawaran", 1)
put(wq, "A1", "Penawaran: cek margin dan komisi Anda", TITLE)
put(wq, "A2", "Isi hanya sel biru. Pilih segmen dan kode produk, isi jumlah, lalu isi diskon ATAU harga usulan per satuan.")
put(wq, "A3", "Klien:", BOLD)
put(wq, "B3", "(nama klien)", BLUE, fill=INPUT_FILL)
put(wq, "D3", "Tanggal:", BOLD)
put(wq, "E3", "(tanggal)", BLUE, fill=INPUT_FILL)
put(wq, "G3", "Bagi hasil sekolah / mitra:", BOLD)
put(wq, "I3", 0, BLUE, PCT, fill=INPUT_FILL)
qcols = ["No", "Segmen", "Kode produk", "Produk", "Satuan", "Jumlah", "Harga daftar per satuan", "Harga minimum per satuan",
         "Diskon", "Harga usulan per satuan (opsional)", "Harga dipakai per satuan", "Total harga", "Margin bersih",
         "Status", "Komisi Anda", "Kalau diskon 10% lebih kecil, komisi Anda naik", "Catatan"]
header(wq, 5, qcols, [5, 13, 13, 34, 13, 8, 14, 14, 8, 15, 14, 15, 10, 30, 13, 17, 30])
wq.column_dimensions["R"].hidden = True
Q1, QN = 6, 25
dv_seg = DataValidation(type="list", formula1='"Pendidikan,Korporat,Regulator,Umum"', allow_blank=True)
dv_code = DataValidation(type="list", formula1=f"=Data!$A${D1}:$A${DN}", allow_blank=True)
for dv in (dv_seg, dv_code):
    wq.add_data_validation(dv)
dv_seg.add(f"B{Q1}:B{QN}")
dv_code.add(f"C{Q1}:C{QN}")
example = {Q1: ("Korporat", "TJ-ASM-08", 50, 0.10, None), Q1 + 1: ("Korporat", "TJ-WS-01", 1, 0, None)}
for rr in range(Q1, QN + 1):
    seg_, code_, qty_, disc_, prop_ = example.get(rr, (None, None, None, None, None))
    put(wq, f"A{rr}", rr - Q1 + 1)
    put(wq, f"B{rr}", seg_, BLUE, fill=INPUT_FILL)
    put(wq, f"C{rr}", code_, BLUE, fill=INPUT_FILL)
    put(wq, f"F{rr}", qty_, BLUE, "#,##0", fill=INPUT_FILL)
    put(wq, f"I{rr}", disc_, BLUE, PCT, fill=INPUT_FILL)
    put(wq, f"J{rr}", prop_, BLUE, RP, fill=INPUT_FILL)
    m = f"MATCH($C{rr},D_CODE,0)"
    hrow = f"INDEX(Harga!$G${H1}:$J${HN},{m},IF($B{rr}=\"Pendidikan\",1,IF($B{rr}=\"Korporat\",2,IF($B{rr}=\"Regulator\",3,4))))"
    put(wq, f"D{rr}", f'=IF($C{rr}="","",INDEX(D_PROD,{m}))', GREEN)
    put(wq, f"E{rr}", f'=IF($C{rr}="","",INDEX(D_UOM,{m}))', GREEN)
    put(wq, f"G{rr}", f'=IF(OR($C{rr}="",$B{rr}=""),"",{hrow})', GREEN, RP)
    put(wq, f"H{rr}", f'=IF($C{rr}="","",INDEX(Harga!$K${H1}:$K${HN},{m}))', GREEN, RP)
    put(wq, f"K{rr}", f'=IF(NOT(ISNUMBER(G{rr})),0,IF(ISNUMBER(J{rr}),J{rr},G{rr}*(1-IF(ISNUMBER(I{rr}),I{rr},0))))', BLACK, RP)
    put(wq, f"L{rr}", f"=K{rr}*IF(ISNUMBER(F{rr}),F{rr},0)", BLACK, RP)
    tp = f'IF(INDEX(D_PPH,{m})="Ya",PPH,0)'
    cost = f"INDEX(D_COST,{m})*IF(ISNUMBER(F{rr}),F{rr},0)"
    gp = f"(L{rr}*(1-{tp}-$I$3)-{cost})"
    put(wq, f"M{rr}", f'=IF(OR($C{rr}="",L{rr}<=0),"",({gp}-MAX(0,{gp})*KOMISI)/L{rr})', BLACK, PCT)
    tgt = f'IF($B{rr}="Pendidikan",T_EDU,IF($B{rr}="Korporat",T_KORP,IF($B{rr}="Regulator",T_REG,T_MIN)))'
    disc = f"IF(ISNUMBER(G{rr}),IF(G{rr}>0,1-K{rr}/G{rr},0),0)"
    put(wq, f"N{rr}", f'=IF($C{rr}="","",IF(NOT(ISNUMBER(G{rr})),"Tidak dijual ke segmen ini",IF(OR(M{rr}<T_MIN,{disc}>0.2),'
                      f'"Merah: tidak boleh (perlu Direktur)",IF(OR(M{rr}<{tgt}-0.0001,{disc}>0.1),"Kuning: perlu persetujuan",'
                      f'"Hijau: boleh"))))', BOLD)
    put(wq, f"O{rr}", f'=IF($C{rr}="",0,MAX(0,{gp})*KOMISI)', BOLD, RP)
    put(wq, f"P{rr}", f'=IF(OR($C{rr}="",NOT(ISNUMBER(G{rr}))),0,G{rr}*MIN({disc},0.1)*IF(ISNUMBER(F{rr}),F{rr},0)*(1-{tp})*KOMISI)', BLACK, RP)
    put(wq, f"R{rr}", f'=IF(ISNUMBER(M{rr}),M{rr}*L{rr},0)', BLACK, RP)
    put(wq, f"Q{rr}", f'=IF($C{rr}="","",IF(AND(ISNUMBER(H{rr}),K{rr}<H{rr}),"Di bawah harga minimum",""))')
for col, fill, test in (("N", G_FILL, "Hijau"), ("N", Y_FILL, "Kuning"), ("N", R_FILL, "Merah")):
    wq.conditional_formatting.add(f"{col}{Q1}:{col}{QN}", FormulaRule(formula=[f'LEFT({col}{Q1},{len(test)})="{test}"'], fill=fill))
wq.conditional_formatting.add(f"N{Q1}:N{QN}", FormulaRule(formula=[f'LEFT(N{Q1},5)="Tidak"'], fill=R_FILL))
s = QN + 2
summary = [
    ("Total harga ke klien (belum PPN)", f"=SUM(L{Q1}:L{QN})", RP),
    ("Margin bersih penawaran", f"=IF(G{s}>0,SUM(R{Q1}:R{QN})/G{s},0)", PCT),
    ("Komisi Anda (10% dari laba kotor)", f"=SUM(O{Q1}:O{QN})", RP),
    ("Kalau setiap diskon 10% lebih kecil, komisi Anda naik", f"=SUM(P{Q1}:P{QN})", RP),
    ("Status penawaran", f'=IF(COUNTIF(N{Q1}:N{QN},"Merah*")+COUNTIF(N{Q1}:N{QN},"Tidak*")>0,"Merah: ada baris yang tidak boleh",'
                         f'IF(COUNTIF(N{Q1}:N{QN},"Kuning*")>0,"Kuning: perlu persetujuan",IF(G{s}>0,"Hijau: boleh dikirim","")))', None),
    ("Kontrak Pendidikan minimal Rp10 jt?", f'=IF(COUNTIF(B{Q1}:B{QN},"Pendidikan")=0,"Tidak berlaku",IF(G{s}>=EDU_MIN,"Ya",'
                                            f'"Belum: tambah item ke paket"))', None),
    ("Bagi hasil sekolah / mitra di batas?", f'=IF(I3>SHARE_MAX,"Tidak: di atas 10%, perlu Direktur","Ya")', None),
    ("Siapa yang menyetujui", f'=IF(OR(LEFT(G{s+4},5)="Merah",I3>SHARE_MAX),"Direktur",IF(LEFT(G{s+4},6)="Kuning","Founder / CEO",'
                              f'IF(COUNTIF(I{Q1}:I{QN},">0")>0,"Lead sales","Tidak perlu")))', None),
]
for i, (lab, f, fmt) in enumerate(summary):
    put(wq, f"D{s+i}", lab, BOLD)
    put(wq, f"G{s+i}", f, Font(name=F, bold=True, size=12), fmt)
wq.conditional_formatting.add(f"G{s+4}", FormulaRule(formula=[f'LEFT(G{s+4},5)="Hijau"'], fill=G_FILL))
wq.conditional_formatting.add(f"G{s+4}", FormulaRule(formula=[f'LEFT(G{s+4},6)="Kuning"'], fill=Y_FILL))
wq.conditional_formatting.add(f"G{s+4}", FormulaRule(formula=[f'LEFT(G{s+4},5)="Merah"'], fill=R_FILL))
put(wq, f"D{s+9}", "Pesannya sederhana: harga lebih tinggi = margin perusahaan lebih tinggi = komisi Anda lebih besar. "
                   "Setiap diskon yang tidak perlu mengurangi komisi Anda sendiri.", Font(name=F, italic=True), align=WRAP)
wq.merge_cells(f"D{s+9}:P{s+9}")
wq.row_dimensions[s + 9].height = 30
wq.freeze_panes = "D6"

# ------------------------------------------------------------------ Panduan (one page)
wg = wb.create_sheet("Panduan", 0)
wg.column_dimensions["A"].width = 4
wg.column_dimensions["B"].width = 34
wg.column_dimensions["C"].width = 95
put(wg, "B1", "Kalkulator Sales Teman Journey + Coreitera", TITLE)
put(wg, "B2", "Satu halaman. Dibuat otomatis dari Master_Produk_dan_Harga.xlsx; jangan ubah angka harga di sini.")
rows = [
    ("SEL YANG BOLEH DIUBAH", None),
    ("Sel biru (latar biru muda)", "Hanya di sheet Penawaran: nama klien, tanggal, bagi hasil sekolah / mitra, segmen, kode produk, jumlah, diskon, harga usulan."),
    ("Sel lain", "Rumus. Jangan ditimpa."),
    ("LANGKAH", None),
    ("1. Pilih segmen", "Pendidikan (sekolah, kampus, komunitas), Korporat (bank, BUMN, swasta, kementerian, pemda), Regulator (BI, OJK, LPS), atau Umum (harga pasar)."),
    ("2. Pilih kode produk", "Daftar lengkap ada di sheet Harga. Kolom harga kosong = produk tidak dijual ke segmen itu."),
    ("3. Isi jumlah", "Sesuai satuan: per orang, per sesi, per paket, per hari, per bulan."),
    ("4. Isi diskon ATAU harga usulan", "Kosongkan keduanya untuk memakai harga daftar. Harga usulan mengalahkan diskon."),
    ("5. Baca warna", "Hijau: boleh. Kuning: perlu persetujuan founder / CEO (margin di bawah target segmen, atau diskon 10-20%). "
                      "Merah: tidak boleh tanpa Direktur (margin di bawah 30%, atau diskon di atas 20%)."),
    ("TARGET MARGIN", None),
    ("Pendidikan", "=TEXT(T_EDU,\"0%\")&\" margin bersih. Kontrak minimal Rp10 jt; bagi hasil sekolah maks. 10%.\""),
    ("Korporat", "=TEXT(T_KORP,\"0%\")&\" margin bersih.\""),
    ("Regulator", "=TEXT(T_REG,\"0%\")&\" margin bersih, minimal 30%. Item yang sudah pernah dibeli regulator memakai harga lama.\""),
    ("KOMISI ANDA", None),
    ("Rumus", "10% dari laba kotor. Harga lebih tinggi = margin lebih tinggi = komisi lebih besar."),
    ("Kolom 'diskon 10% lebih kecil'", "Menunjukkan berapa komisi Anda bertambah kalau diskon dikurangi 10 poin (misalnya dari 15% ke 5%)."),
    ("CONTOH (baris 1 sheet Penawaran)", None),
    ("Isian", "Segmen Korporat, TJ-ASM-08 MHCU Flash, 50 orang, diskon 10%."),
    ("Hasil", "=\"Harga dipakai Rp\"&TEXT(Penawaran!K6,\"#,##0\")&\" per orang, margin \"&TEXT(Penawaran!M6,\"0%\")&\", status: \"&Penawaran!N6&\". Komisi Anda Rp\"&TEXT(Penawaran!O6,\"#,##0\")&\"; tanpa diskon, komisi naik Rp\"&TEXT(Penawaran!P6,\"#,##0\")&\".\""),
    ("ATURAN LAIN", None),
    ("Pajak", "Harga belum termasuk PPN. Teman Journey dan Coreitera Non-PKP: jangan pernah menulis 'termasuk PPN'."),
    ("Pintu masuk", "Snapshot burnout dan tes mandiri boleh gratis. Produk yang memakai waktu psikolog tidak pernah gratis."),
    ("Perjalanan", "Di luar Jakarta: tiket, hotel, dan uang harian ditagih sebesar biaya aktual."),
]
r = 4
for a, b in rows:
    if b is None:
        put(wg, f"B{r}", a, Font(name=F, bold=True, color="FFFFFF"), fill=HDR_FILL)
        wg[f"C{r}"].fill = HDR_FILL
    else:
        put(wg, f"B{r}", a, BOLD, fill=INPUT_FILL if a.startswith("Sel biru") else None)
        put(wg, f"C{r}", b, BLACK, align=WRAP)
    r += 1
wg.sheet_view.showGridLines = False

# ------------------------------------------------------------------ protection, order, fonts
for sh in (wp, wd, wh):
    sh.protection.sheet = True
    sh.protection.password = PROTECT_PW
wg.protection.sheet = True
wg.protection.password = PROTECT_PW
wq.protection.sheet = True
wq.protection.password = PROTECT_PW
from openpyxl.styles import Protection
for rr in range(Q1, QN + 1):
    for col in "BCFIJ":
        wq[f"{col}{rr}"].protection = Protection(locked=False)
for ref in ("B3", "E3", "I3"):
    wq[ref].protection = Protection(locked=False)
wp.sheet_state = "hidden"
wd.sheet_state = "hidden"
wb.security = WorkbookProtection(workbookPassword=PROTECT_PW, lockStructure=True)
wb._sheets = [wg, wq, wh, wp, wd]
wb.active = 1
for sh in wb.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if c.font is None or c.font.name != F:
                c.font = Font(name=F, bold=c.font.bold if c.font else False, color=c.font.color if c.font else None)
wb.save(OUT)
print("saved", os.path.abspath(OUT), "products:", len(prods))
