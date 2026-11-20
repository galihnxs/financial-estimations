"""Build Katalog_Produk_TJ_Coreitera.pptx from the Employee Wellbeing proposal template.

Every slide is a template slide, kept or duplicated, with its text replaced. Prices come from
01_Daftar_Produk_Final.md / Master_Produk_dan_Harga.xlsx (19 Nov 2026). No cost or margin is shown.

Run: python3 -I build_product_deck.py   (needs python-pptx)
"""
import os
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Emu
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
TEMPLATE = os.path.join(ROOT, "[PROPOSAL TEMAN JOURNEY] Employee Wellbeing - Bank Indonesia.pptx")
OUT = os.path.join(ROOT, "ProductKnowledge", "Katalog_Produk_TJ_Coreitera.pptx")
DATE_OLD, DATE_NEW = "Oktober 2026", "November 2026"
TABLE_H = 7.3  # inches; table starts at 2.8", footnote sits at 10.3"

prs = Presentation(TEMPLATE)
T = {i: s for i, s in enumerate(prs.slides, 1)}  # template slides by original number


# ---------------------------------------------------------------- helpers
def dup(src):
    """Duplicate a slide (shapes, background, image relationships) at the end of the deck."""
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)
    bg = src._element.cSld.find(qn("p:bg"))
    if bg is not None:
        new._element.cSld.insert(0, deepcopy(bg))
    rid = {}
    for rel in src.part.rels.values():
        if "slideLayout" in rel.reltype or "notesSlide" in rel.reltype:
            continue
        if rel.is_external:
            rid[rel.rId] = new.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            rid[rel.rId] = new.part.relate_to(rel.target_part, rel.reltype)
    tree = new.shapes._spTree
    for el in src.shapes._spTree:
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        tree.append(deepcopy(el))
    for el in tree.iter():
        for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
            v = el.get(attr)
            if v in rid:
                el.set(attr, rid[v])
    return new


def all_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:
            yield from all_shapes(sh.shapes)


def shp(slide, sid):
    for sh in all_shapes(slide.shapes):
        if sh.shape_id == sid:
            return sh
    raise KeyError(f"shape {sid} not on slide")


def remove(slide, *sids):
    for sid in sids:
        el = shp(slide, sid)._element
        el.getparent().remove(el)


def set_text(slide, sid, paras, scale=None, size=None, bold=None):
    """Replace a text box's paragraphs; each item is a string or a list of (text, bold) runs.
    Keeps the first paragraph's format and the longest run's character format."""
    body = shp(slide, sid).text_frame._txBody
    ps = body.findall(qn("a:p"))
    runs = [r for p in ps for r in p.findall(qn("a:r"))]
    tmpl_r = max(runs, key=lambda r: len(r.findtext(qn("a:t")) or ""))
    tmpl_p = ps[0]
    for p in ps:
        body.remove(p)
    for item in paras:
        p = deepcopy(tmpl_p)
        for x in p.findall(qn("a:r")) + p.findall(qn("a:br")) + p.findall(qn("a:fld")):
            p.remove(x)
        end = p.find(qn("a:endParaRPr"))
        for text, b in (item if isinstance(item, list) else [(item, bold)]):
            r = deepcopy(tmpl_r)
            r.find(qn("a:t")).text = text
            rpr = r.find(qn("a:rPr"))
            if b is not None:
                rpr.set("b", "1" if b else "0")
            if size:
                rpr.set("sz", str(int(size * 100)))
            elif scale and rpr.get("sz"):
                rpr.set("sz", str(int(int(rpr.get("sz")) * scale)))
            if end is not None:
                end.addprevious(r)
            else:
                p.append(r)
        body.append(p)


def replace_runs(slide, old, new):
    for sh in all_shapes(slide.shapes):
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if old in r.text:
                        r.text = r.text.replace(old, new)


def _cell(tc, value, bold=None, align=None, size=None, pad=None):
    body = tc.find(qn("a:txBody"))
    ps = body.findall(qn("a:p"))
    for p in ps[1:]:
        body.remove(p)
    p = ps[0]
    runs = p.findall(qn("a:r"))
    for r in runs[1:]:
        p.remove(r)
    r = runs[0]
    r.find(qn("a:t")).text = value
    rpr = r.find(qn("a:rPr"))
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
    if size:
        rpr.set("sz", str(int(size * 100)))
        lns = p.find(qn("a:pPr")).find(qn("a:lnSpc"))
        if lns is not None:
            lns.find(qn("a:spcPts")).set("val", str(int(size * 130)))
    if align:
        p.find(qn("a:pPr")).set("algn", align)
    if pad is not None:
        tcpr = tc.find(qn("a:tcPr"))
        tcpr.set("marT", str(pad))
        tcpr.set("marB", str(pad))


def fill_table(slide, sid, header, rows, widths, aligns=None, size=None, row_h=None, pad=None, bold_cols=()):
    """Refill the template table: header row style from row 1, body style from row 3."""
    gf = shp(slide, sid)
    tbl = gf.table._tbl
    trs = tbl.findall(qn("a:tr"))
    hdr_t, body_t = deepcopy(trs[0]), deepcopy(trs[2])
    for tr in trs:
        tbl.remove(tr)
    n = len(header)
    grid = tbl.find(qn("a:tblGrid"))
    for c in grid.findall(qn("a:gridCol"))[n:]:
        grid.remove(c)
    for c, w in zip(grid.findall(qn("a:gridCol")), widths):
        c.set("w", str(int(gf.width * w / sum(widths))))
    aligns = aligns or ["ctr"] * n
    nrows = len(rows) + 1
    h = row_h or int(Inches(TABLE_H) / nrows)

    def add(tmpl, values, is_hdr):
        tr = deepcopy(tmpl)
        for tc in tr.findall(qn("a:tc"))[n:]:
            tr.remove(tc)
        tr.set("h", str(h))
        for j, (tc, v) in enumerate(zip(tr.findall(qn("a:tc")), values)):
            _cell(tc, v, bold=True if is_hdr else (j in bold_cols), align="ctr" if is_hdr else aligns[j],
                  size=size, pad=pad)
        tbl.append(tr)

    add(hdr_t, header, True)
    for row in rows:
        add(body_t, row, False)
    gf.height = Emu(h * nrows)


def set_title_table_slide(slide, title, note):
    set_text(slide, 17, [title])
    set_text(slide, 18, [note])
    fn = shp(slide, 18)
    fn.left, fn.top, fn.width, fn.height = Inches(1.12), Inches(10.3), Inches(17.5), Inches(0.4)


# ---------------------------------------------------------------- build the new slides (duplicates first)
D = {}
D["grup"] = dup(T[13])
D["peta"] = dup(T[24])
D["harga_latih"] = dup(T[24])
D["paket_latih"] = dup(T[18])
D["harga_asesmen"] = dup(T[24])
D["edukasi"] = dup(T[22])
D["b2c"] = dup(T[24])
D["tes"] = dup(T[24])
D["core"] = dup(T[13])
D["theta"] = dup(T[22])
D["aturan"] = dup(T[24])

# ---------------------------------------------------------------- 1. cover
s = T[1]
set_text(s, 27, ["Katalog Produk dan Layanan", "Teman Journey × Coreitera"])
set_text(s, 20, ["untuk tim sales"])
remove(s, 21)  # client (BI) logo

# ---------------------------------------------------------------- group slide (from 13)
s = D["grup"]
set_text(s, 17, ["Siapa kami - Satu Grup, Satu Solusi Psikologi"])
set_text(s, 16, ["Teman Journey"], scale=0.78)
set_text(s, 18, ["Coreitera"])
set_text(s, 19, ["Bersama"])
set_text(s, 20, ["Biro psikologi: pelatihan, EAP, asesmen, psikoedukasi, dan konseling oleh psikolog bersertifikat HIMPSI."])
set_text(s, 23, ["Platform teknologi: tes online, laporan otomatis berbasis AI, dan Theta untuk booking dan pembayaran biro psikologi."])
set_text(s, 24, ["Psikolog dan teknologi dalam satu kontrak: tes massal yang cepat, laporan konsisten, dan tindak lanjut oleh psikolog."])

# ---------------------------------------------------------------- product map (from 24)
s = D["peta"]
set_title_table_slide(s, "Peta Produk", "*harga belum termasuk PPN; di luar Jakarta, perjalanan ditagih sebesar biaya aktual")
fill_table(s, 16, ["Keluarga produk", "Untuk siapa", "Mulai dari"], [
    ["Pelatihan dan Program Pelatih", "Regulator (BI, OJK, LPS), bank", "Rp28,5 jt / hari"],
    ["Employee Assistance Program (EAP)", "HR instansi dan perusahaan", "Rp12 jt / 3 bulan"],
    ["Asesmen Seleksi, Profiling, MHCU", "HC, direksi, regulator", "Rp500rb / peserta"],
    ["Psikoedukasi dan Jasa Ahli", "HR, komunitas, panitia lomba", "Rp2,6 jt / webinar"],
    ["Konseling Individu dan Pasangan", "Individu, pasangan", "Rp500rb / sesi"],
    ["Tes Psikologi Individu", "Orang tua, pelajar, dewasa", "Rp375rb / tes"],
    ["Coreitera Theta", "Biro psikologi", "Rp1 jt / bulan"],
    ["Coreitera Synthesia dan Mesin Tes", "Biro psikologi korporat", "Segera hadir"],
], widths=[7, 6, 4], aligns=["l", "l", "ctr"], size=18, pad=76200, bold_cols=(2,))

# ---------------------------------------------------------------- training prices (from 24)
s = D["harga_latih"]
set_title_table_slide(s, "Harga Pelatihan", "*semua program termasuk asesmen pre/post dan laporan dampak; di luar Jakarta, perjalanan ditagih sebesar biaya aktual")
fill_table(s, 16, ["Program", "Isi", "Satuan", "Harga daftar"], [
    ["Program Pelatih Kantor Pusat", "3 hari, hingga 40 peserta", "per program", "Rp100.000.000"],
    ["Program Pelatih Regional", "1 hari, hingga 30 peserta", "per hari", "Rp28.500.000"],
    ["Program Bank BUMN (Signature)", "1 kelompok, hingga 40 peserta", "per hari", "Rp51.000.000"],
    ["Program Bank Swasta (Signature)", "1 kelompok, hingga 40 peserta", "per hari", "Rp61.000.000"],
    ["Bank Elite BUMN", "Signature + lisensi pelatih + pendampingan 3 bulan", "per hari", "Rp66.300.000"],
    ["Bank Elite Swasta", "Signature + lisensi pelatih + pendampingan 3 bulan", "per hari", "Rp79.300.000"],
    ["Asesmen dan laporan dampak tambahan", "peserta di luar kuota paket", "per peserta", "Rp150.000"],
], widths=[6, 7, 2.6, 3.6], aligns=["l", "l", "ctr", "ctr"], size=18, pad=76200, bold_cols=(3,))

# ---------------------------------------------------------------- training packages (from 18)
s = D["paket_latih"]
set_text(s, 16, ["Program Pelatihan untuk Regulator dan Bank"])
set_text(s, 15, ["Pelatih"])
set_text(s, 17, ["Signature"])
shp(s, 17).left, shp(s, 17).width = Inches(7.28), Inches(4.81)
set_text(s, 18, ["Elite"])
set_text(s, 29, ["Training of Trainers untuk regulator", "Kantor pusat: 3 hari, hingga 40 peserta", "Regional: 1 hari, hingga 30 peserta",
                 "Asesmen pre/post + laporan dampak", "Rp100 jt / program; Rp28,5 jt / hari"])
set_text(s, 28, ["Program bank 1 hari, 1 kelompok hingga 40 peserta", "Asesmen pre/post + laporan dampak",
                 "BUMN Rp51 jt / hari", "Swasta Rp61 jt / hari"])
set_text(s, 30, ["Semua isi Signature", "Lisensi pelatih internal bank", "Pendampingan 3 bulan",
                 "BUMN Rp66,3 jt / hari", "Swasta Rp79,3 jt / hari"])
for g in (19, 22, 25):
    shp(s, g).height = Inches(4.9)

# ---------------------------------------------------------------- assessment overview (8, edited in place)
s = T[8]
set_text(s, 8, ["Layanan asesmen Teman Journey mencakup pengukuran psikologis yang komprehensif, dalam tiga kelompok:"])
set_text(s, 12, ["Seleksi"])
set_text(s, 13, ["Fit and proper dan promosi: tes BFI dan DISC, LGD, wawancara psikolog, dan laporan rekomendasi per kandidat."])
set_text(s, 18, ["Profiling"])
set_text(s, 17, ["Tes BFI dan DISC online untuk ratusan sampai ribuan karyawan, laporan otomatis, sampai peta talenta per unit."])
set_text(s, 23, ["Kesehatan Mental"])
set_text(s, 22, ["Mental Health Check-Up (MHCU): 5 instrumen klinis, review psikolog, laporan personal, dan sesi konseling."])

# ---------------------------------------------------------------- assessment prices (from 24)
s = D["harga_asesmen"]
set_title_table_slide(s, "Harga Asesmen", "*profiling dan pemetaan dihitung untuk 2.000 karyawan, jendela tes 7 hari; MHCU dan add-on per orang")
fill_table(s, 16, ["Layanan", "Isi", "Satuan", "Harga daftar"], [
    ["Seleksi / Fit and Proper", "BFI + DISC, LGD, wawancara, laporan", "per 6 peserta", "Rp15.000.000"],
    ["Profiling Massal Paket 1", "Laporan otomatis untuk perusahaan", "per proyek", "Rp50.000.000"],
    ["Profiling Massal Paket 2", "Paket 1 + laporan untuk karyawan", "per proyek", "Rp70.000.000"],
    ["Pemetaan 3A Umum", "Paket 2 + analisis 10 unit + presentasi", "per proyek", "Rp101.500.000"],
    ["Pemetaan 3B Talent Box", "3A + talent box dan kalibrasi", "per proyek", "Rp139.500.000"],
    ["Pemetaan 3C Talenta Lengkap", "3B + analisis 15 rumpun jabatan", "per proyek", "Rp210.500.000"],
    ["Tambahan tes serentak", "Server + engineer standby hari H", "per proyek", "Rp30.000.000"],
    ["MHCU Flash", "5 instrumen + konseling 30 menit", "per peserta", "Rp500.000"],
    ["MHCU Deep", "5 instrumen + konseling 60 menit", "per peserta", "Rp750.000"],
    ["Hasil MHCU individual", "Add-on EAP, laporan per pegawai", "per pegawai", "Rp70.000"],
], widths=[5.6, 7.2, 2.8, 3.6], aligns=["l", "l", "ctr", "ctr"], size=17, pad=57150, bold_cols=(3,))

# ---------------------------------------------------------------- psychoeducation prices (from 22)
s = D["edukasi"]
set_text(s, 16, ["Psikoedukasi dan Jasa Ahli*"])
set_text(s, 15, ["Investasi Program"])
set_text(s, 28, ["Webinar"])
set_text(s, 29, ["Seminar"])
set_text(s, 30, ["Event Publik"])
set_text(s, 27, ["Rp2.600.000 / sesi"])
set_text(s, 38, ["Rp4.900.000 / sesi"])
set_text(s, 39, ["Rp24.000.000 / event"])
set_text(s, 32, ["Webinar online 60-90 menit oleh psikolog, topik sesuai kebutuhan"])
set_text(s, 33, ["Seminar tatap muka di Jakarta oleh psikolog"])
set_text(s, 34, ["Event psikoedukasi untuk umum, dikelola penuh oleh Teman Journey"])
set_text(s, 31, ["Juri ahli kompetisi Rp15.000.000 / acara  •  Event komunitas bertiket Rp259.000 / peserta"])
set_text(s, 40, ["*di luar Jakarta: perjalanan ditagih sebesar biaya aktual"])
shp(s, 40).width = Inches(12)
remove(s, 35)

# ---------------------------------------------------------------- B2C counseling prices (from 24)
s = D["b2c"]
set_title_table_slide(s, "Harga Layanan Individu", "*Tier = pengalaman psikolog: Tier 1 0-2 th, Tier 2 2-5 th, Tier 3 5-10 th, Tier 4 di atas 10 th")
fill_table(s, 16, ["Layanan", "Tier 1", "Tier 2", "Tier 3", "Tier 4"], [
    ["Konseling online 60 menit", "Rp500.000", "Rp500.000", "Rp725.000", "Rp850.000"],
    ["Konseling offline 60 menit", "Rp580.000", "Rp580.000", "Rp800.000", "Rp925.000"],
    ["Konseling pasangan 90 menit", "Rp800.000", "Rp950.000", "Rp1.100.000", "Rp1.425.000"],
    ["Pre-Marriage Check Up (3 sesi)", "Rp2.200.000", "Rp2.600.000", "Rp2.900.000", "Rp3.800.000"],
    ["Marriage Check Up (2 sesi)", "Rp1.400.000", "Rp1.700.000", "Rp2.000.000", "Rp2.600.000"],
    ["Terapi klinis offline (EMDR / hipnoterapi)", "Rp750.000", "Rp900.000", "Rp1.000.000", "Rp1.325.000"],
    ["Terapi klinis online (EMDR / hipnoterapi)", "Rp675.000", "Rp800.000", "Rp900.000", "Rp1.200.000"],
], widths=[7, 3, 3, 3, 3], aligns=["l", "ctr", "ctr", "ctr", "ctr"], size=18, pad=76200)

# ---------------------------------------------------------------- B2C tests (from 24)
s = D["tes"]
set_title_table_slide(s, "Tes Psikologi Individu", "*harga sudah termasuk laporan tertulis oleh psikolog")
fill_table(s, 16, ["Tes", "Usia", "Harga"], [
    ["Skrining Tumbuh Kembang", "0-5 tahun", "Rp625.000"],
    ["Tes IQ", "Anak sampai dewasa", "Rp725.000"],
    ["Tes Kesiapan Sekolah", "4-6 tahun", "Rp800.000"],
    ["Tes Profiling", "6-13 tahun", "Rp1.225.000"],
    ["Tes Minat Bakat", "14 tahun ke atas", "Rp1.625.000"],
    ["Tes Kesehatan Mental Dewasa (MHCU)", "Dewasa", "Rp375.000"],
], widths=[8, 5, 4], aligns=["l", "ctr", "ctr"], size=20, bold_cols=(2,))

# ---------------------------------------------------------------- Coreitera overview (from 13)
s = D["core"]
set_text(s, 17, ["Coreitera - Platform untuk Biro Psikologi"])
set_text(s, 16, ["Theta"])
set_text(s, 18, ["Synthesia"])
set_text(s, 19, ["Mesin Tes"])
set_text(s, 20, ["Booking, pembayaran, jadwal, catatan klien, dan agen AI 24 jam untuk biro. Mulai dipakai Teman Journey 30 November 2026."])
set_text(s, 23, ["AI menyusun draf interpretasi tes; psikolog biro mereview dan menandatangani. Laporan lebih cepat dan konsisten. Rencana 2028."])
set_text(s, 24, ["Tes online dengan merek biro sendiri dan skor otomatis, dijual sekitar harga pokok. Dalam pengembangan."])

# ---------------------------------------------------------------- Theta prices (from 22)
s = D["theta"]
set_text(s, 16, ["Coreitera Theta*"])
set_text(s, 15, ["Paket Langganan Biro"])
shp(s, 15).width = Inches(17)
set_text(s, 28, ["Starter"])
set_text(s, 29, ["Growth"])
set_text(s, 30, ["Pro"])
set_text(s, 27, ["Rp1.000.000 / bulan"])
set_text(s, 38, ["Rp2.900.000 / bulan"])
set_text(s, 39, ["Rp5.730.000 / bulan"])
set_text(s, 32, ["Biro dengan 1-3 psikolog"])
set_text(s, 33, ["Biro dengan 4-10 psikolog"])
set_text(s, 34, ["Biro dengan 11 psikolog atau lebih"])
set_text(s, 31, ["Biaya transaksi Rp5.000 + biaya Midtrans  •  Synthesia Rp300.000 / kursi psikolog / bulan (2028)"])
set_text(s, 40, ["*harga per biro per bulan, belum termasuk PPN"])
shp(s, 40).width = Inches(12)
remove(s, 35)

# ---------------------------------------------------------------- selling rules (from 24)
s = D["aturan"]
set_title_table_slide(s, "Aturan Penjualan (internal)", "*sembunyikan slide ini sebelum deck dikirim ke klien")
fill_table(s, 16, ["Aturan", "Isi"], [
    ["Harga daftar", "Buka penawaran dengan harga di katalog ini. Tidak boleh di bawahnya tanpa persetujuan."],
    ["Diskon", "Pilot maks. 20%, sekali per klien. Volume 5% (5+ kelompok) atau 10% (10+ kelompok)."],
    ["Persetujuan", "Sampai 10%: lead sales. Sampai 20%: CEO. Di atas 20% atau margin di bawah 30%: Direktur."],
    ["Selalu ke Direktur", "EAP Jeda Tahunan dan sesi tambahan EAP (margin di bawah 30%)."],
    ["Pembayaran", "EAP: DP 50% saat SPK. Kontrak bank di atas Rp100 jt: DP 30%. Regulator: ikuti ketentuan mereka."],
    ["Perjalanan", "Di luar Jakarta: perjalanan dan akomodasi ditagih sebesar biaya aktual."],
    ["Pajak", "Teman Journey dan Coreitera Non-PKP: jangan tulis “termasuk PPN”. PPh 23 dipotong klien."],
], widths=[4, 14], aligns=["l", "l"], size=18, pad=76200, bold_cols=(0,))

# ---------------------------------------------------------------- text fixes on kept slides
replace_runs(T[13], "Asessment", "Assessment")
set_text(T[23], 39, ["≈ Rp11.000.000 / term"])
replace_runs(T[14], "PhD Candidate of", "PhD Candidate")
for sl in prs.slides:
    replace_runs(sl, DATE_OLD, DATE_NEW)

# ---------------------------------------------------------------- order: drop template 9 and 10 (BI-only)
order = [T[1], T[2], T[3], T[4], T[5], D["grup"], D["peta"],
         T[6], D["harga_latih"], D["paket_latih"],
         T[11], T[12], T[13], T[18], T[17], T[22], T[23], T[24],
         T[8], D["harga_asesmen"], T[20],
         D["edukasi"], T[19],
         T[7], D["b2c"], D["tes"],
         D["core"], D["theta"],
         T[14], T[15], T[16], T[21],
         D["aturan"], T[25]]
sld_lst = prs.slides._sldIdLst
by_part = {}
for sid in list(sld_lst):
    by_part[prs.part.related_part(sid.rId)] = sid
    sld_lst.remove(sid)
keep = set()
for sl in order:
    sld_lst.append(by_part[sl.part])
    keep.add(sl.part)
for part, sid in by_part.items():
    if part not in keep:
        prs.part.drop_rel(sid.rId)
prs.save(OUT)
print("saved", OUT, len(order), "slides")
