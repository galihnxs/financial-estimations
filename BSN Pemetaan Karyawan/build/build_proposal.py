"""Builds the BSN (Bank Syariah Nasional) employee mapping proposal from the Teman Journey proposal template.

Usage: python build_proposal.py <template.docx> <out.docx>
Keeps the cover frame, profile, legality, products and attachments of the template,
replaces the program section, and reads service prices from the RAB workbook
(Harga and Opsi tabs), so the proposal never shows internal costs.
Run build_rab.py and recalc the workbook in Excel first.
"""
import copy
import sys
from pathlib import Path

import docx
import openpyxl
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

HERE = Path(__file__).resolve().parent
RAB = HERE.parent / "RAB - BSN Pemetaan Karyawan BFI DISC (Internal).xlsx"
COVER = HERE / "cover_bsn.png"
NAVY = RGBColor(0x1C, 0x45, 0x87)
HEAD_HEX, ZEBRA_HEX, STRIPE_HEX, RULE_HEX = "475571", "E4E8F0", "F5F7FB", "D5DBE6"
WHITE, MUTED = RGBColor(0xFF, 0xFF, 0xFF), RGBColor(0x6B, 0x7A, 0x99)

template, out = sys.argv[1], sys.argv[2]
d = docx.Document(template)
body = d.element.body


def find_par(start):
    for p in d.paragraphs:
        if p.text.strip().startswith(start):
            return p
    raise LookupError(start)


T = {
    "h1": find_par("Pelatihan Metode Belajar Kreatif"),
    "h2": find_par("Latar Belakang Program"),
    "sub": find_par("Tujuan Program"),
    "minor": find_par("Materi 1 - Kesiapan"),
    "body": find_par("Literasi keuangan merupakan"),
    "bullet": find_par("Keterlibatan aktif peserta"),
}
T = {k: copy.deepcopy(v._p) for k, v in T.items()}
anchor = find_par("Lampiran-lampiran Dokumen")._p

# Drop the old program section: from its H1 up to the attachments H1.
start = find_par("Pelatihan Metode Belajar Kreatif")._p
el = start
while el is not anchor:
    nxt = el.getnext()
    body.remove(el)
    el = nxt

# Fix a stray double comma in the template's product text.
for p in d.paragraphs:
    if "penyusunan asesmen, , serta" in p.text:
        for r in p.runs:
            r.text = r.text.replace("penyusunan asesmen, , serta", "penyusunan asesmen, serta")
        if "penyusunan asesmen, , serta" in p.text:  # comma split across runs
            full = p.text.replace("penyusunan asesmen, , serta", "penyusunan asesmen, serta").rstrip(" -")
            for r in p.runs[1:]:
                r.text = ""
            p.runs[0].text = full


def ppr_flag(ppr, tag):
    """Add an on/off paragraph property right after pStyle, where the schema expects it."""
    if ppr.find(qn(tag)) is not None:
        return
    style = ppr.find(qn("w:pStyle"))
    ppr.insert(0 if style is None else list(ppr).index(style) + 1, OxmlElement(tag))


def add(kind, *runs, page_break=False, keep_next=False):
    """Clone a template paragraph. runs: strings, or (text, bold) tuples."""
    p = copy.deepcopy(T[kind])
    rpr = p.find(".//" + qn("w:rPr"))
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    ppr = p.find(qn("w:pPr"))
    if page_break and ppr is not None:
        ppr_flag(ppr, "w:pageBreakBefore")
    if (keep_next or kind in ("h1", "h2", "sub", "minor")) and ppr is not None:
        ppr_flag(ppr, "w:keepNext")  # never strand a heading at the page bottom
    for item in runs:
        text, bold = (item, None) if isinstance(item, str) else item
        r = OxmlElement("w:r")
        if rpr is not None:
            new_rpr = copy.deepcopy(rpr)
            if bold is not None:
                for tag in ("w:b", "w:bCs"):
                    for old in new_rpr.findall(qn(tag)):
                        new_rpr.remove(old)
                    if bold:
                        new_rpr.append(OxmlElement(tag))
            r.append(new_rpr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        p.append(r)
    anchor.addprevious(p)
    return p


def shade(cell, hex_fill):
    xml_child(cell._tc.get_or_add_tcPr(), "w:shd", TCPR_ORDER, val="clear", color="auto", fill=hex_fill)


TBLPR_ORDER = ["w:tblStyle", "w:tblW", "w:jc", "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"]
TCPR_ORDER = ["w:tcW", "w:gridSpan", "w:vMerge", "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:vAlign"]


def xml_child(parent, tag, order=None, **attrs):
    """Add (or replace) a child, placed where the Word schema expects it when an order is given."""
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(f"w:{k}"), str(v))
    if order:
        for old in parent.findall(qn(tag)):
            parent.remove(old)
        rank = order.index(tag)
        later = [c for c in parent if c.tag in {qn(t) for t in order[rank + 1:]}]
        if later:
            later[0].addprevious(el)
            return el
    parent.append(el)
    return el


def cell_text(cell, value, *, color, bold=False, size=10.5, align=None):
    """value: a string, or (title, subtitle) for a bold title with a small grey line under it."""
    title, subtitle = (value, None) if isinstance(value, str) else value
    strike = title.startswith("~~")
    title = title.strip("~").strip("*")
    lines = [(title, bold or subtitle is not None, size, color)]
    if subtitle:
        lines.append((subtitle, False, 9, MUTED))
    for k, (text, is_bold, pt, rgb) in enumerate(lines):
        para = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
        fmt = para.paragraph_format
        fmt.space_before, fmt.space_after, fmt.line_spacing = Pt(0), Pt(1 if subtitle and k == 0 else 0), 1.0
        if align is not None:
            para.alignment = align
        run = para.add_run(text)
        run.font.name, run.font.size, run.bold = "Calibri", Pt(pt), is_bold
        run.font.color.rgb = rgb
        if strike and k == 0:
            run.font.strike = True
            run.font.color.rgb = MUTED
    return cell


def table(rows, col_widths_cm, footer=(), align_right=(), align_center=(), size=10.5):
    """rows[0] is the header. footer: styles for the last rows ("dark" or "light"); their
    label cells merge across every column except the last. Moves the table before the anchor."""
    tbl = d.add_table(rows=len(rows), cols=len(rows[0]))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = tbl._tbl.tblPr
    xml_child(tblPr, "w:tblW", TBLPR_ORDER, w=sum(int(w * 567) for w in col_widths_cm), type="dxa")
    xml_child(tblPr, "w:tblLayout", TBLPR_ORDER, type="fixed")
    # Only horizontal rules: a strong line under the table, hairlines between rows.
    borders = xml_child(tblPr, "w:tblBorders", TBLPR_ORDER)
    for edge, val, sz, color in [("top", "nil", 0, "auto"), ("left", "nil", 0, "auto"),
                                 ("bottom", "single", 12, HEAD_HEX), ("right", "nil", 0, "auto"),
                                 ("insideH", "single", 4, RULE_HEX), ("insideV", "nil", 0, "auto")]:
        attrs = {"val": val} if val == "nil" else {"val": val, "sz": sz, "space": 0, "color": color}
        xml_child(borders, f"w:{edge}", **attrs)
    margins = xml_child(tblPr, "w:tblCellMar", TBLPR_ORDER)
    for edge, w in [("top", 110), ("left", 130), ("bottom", 110), ("right", 130)]:
        xml_child(margins, f"w:{edge}", w=w, type="dxa")
    grid = tbl._tbl.tblGrid
    for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(col_widths_cm[j] * 567)))

    n_body_end = len(rows) - len(footer)
    for i, row in enumerate(rows):
        tr = tbl.rows[i]._tr
        trPr = tr.get_or_add_trPr()
        xml_child(trPr, "w:cantSplit")
        if i == 0:
            xml_child(trPr, "w:tblHeader")
        style = footer[i - n_body_end] if i >= n_body_end else None
        if style:
            label_end = len(row) - 2
            tbl.cell(i, 0).merge(tbl.cell(i, label_end))
        for j, value in enumerate(row):
            if style and 0 < j <= len(row) - 2:
                continue  # merged into the label cell
            cell = tbl.cell(i, j)
            cell.width = docx.shared.Cm(col_widths_cm[j] if not (style and j == 0) else sum(col_widths_cm[:-1]))
            xml_child(cell._tc.get_or_add_tcPr(), "w:vAlign", TCPR_ORDER, val="center")
            align = 2 if j in align_right else 1 if j in align_center else None
            if i == 0:
                shade(cell, HEAD_HEX)
                cell_text(cell, value, color=WHITE, bold=True, size=min(10, size), align=align)
            elif style == "dark":
                shade(cell, HEAD_HEX)
                cell_text(cell, value, color=WHITE, bold=True, size=11.5, align=2 if j else None)
            elif style == "light":
                shade(cell, ZEBRA_HEX)
                cell_text(cell, value, color=NAVY, bold=True, align=2 if j else None)
            else:
                if i % 2 == 0:
                    shade(cell, STRIPE_HEX)
                cell_text(cell, value, color=NAVY, align=align, size=size,
                          bold=isinstance(value, str) and value.startswith("**"))
            for para in cell.paragraphs:
                para.paragraph_format.keep_with_next = i < len(rows) - 1  # keep the table on one page
    anchor.addprevious(tbl._tbl)
    spacer = OxmlElement("w:p")
    anchor.addprevious(spacer)
    return tbl


def rupiah(n):
    return "Rp" + f"{round(n):,}".replace(",", ".")


# ---------------------------------------------------------------- prices from the RAB
wb = openpyxl.load_workbook(RAB, data_only=True)
H = {row[0]: row[2] for row in wb["Harga"].iter_rows(min_row=3, values_only=True) if row[0]}
if H.get("P_A") is None:
    raise SystemExit("RAB has no cached values. Recalculate it in Excel first.")
OP = wb["Opsi"]
op_rows = {OP.cell(i, 1).value: i for i in range(1, OP.max_row + 1)}


def opt(line, col):
    """col: 0-based option index in Opsi (0 = Paket 1 jendela, 1 = Paket 1 serentak, ...)."""
    return OP.cell(op_rows[line], col + 2).value


N = wb["Asumsi"]["C3"].value
assert N == 2000, N
PKG = [  # name, option column (window mode), need
    ("Paket 1: Laporan untuk Perusahaan", 0, "Laporan otomatis, hanya dibaca perusahaan"),
    ("Paket 2: Laporan untuk Perusahaan dan Karyawan", 2, "Laporan otomatis, dibaca perusahaan dan karyawan"),
    ("Paket 3A: Pemetaan Umum", 4, "Gambaran profil per unit oleh psikolog"),
    ("Paket 3B: Pemetaan + Talent Box", 6, "Peta talenta per unit"),
    ("Paket 3C: Pemetaan Talenta Lengkap", 8, "Peta talenta berdasarkan analisis jabatan"),
]
assert opt("Harga total (DPP)", 2) == N * H["A2000"] == 70_000_000  # the locked price


def rp(n):
    return rupiah(n)


# ---------------------------------------------------------------- content
add("h1", "Pemetaan Karyawan dengan Tes BFI dan DISC", page_break=True)

add("h2", "Latar Belakang")
add("body", "Bank Syariah Nasional (BSN) ingin memetakan profil kepribadian dan gaya kerja 2.000 karyawan. Pemetaan ini membantu HC "
    "melihat kekuatan setiap unit, menempatkan orang di jabatan yang paling cocok, dan menyusun program pengembangan yang tepat sasaran.")
add("body", "Teman Journey bersama platform Coreitera menyediakan tes online, laporan otomatis untuk setiap karyawan, dan pemetaan oleh "
    "psikolog. BSN bisa memilih paket sesuai kebutuhan, dari laporan otomatis saja sampai pemetaan talenta lengkap.")

add("h2", "Ringkasan Paket dan Harga")
rows = [["Paket", "Untuk kebutuhan", "Harga normal", "Harga untuk BSN", "Per peserta"]]
AFTER_MEETING = "Disampaikan setelah pertemuan"  # Galih, 4 Nov 2026: Paket 3 prices only after the scoping meeting
for name, col, need in PKG:
    if name.startswith("Paket 3"):
        rows.append([name, need, "~~" + rp(opt("Harga normal", col)), "**" + AFTER_MEETING, "Setelah pertemuan"])
    else:
        rows.append([name, need, "~~" + rp(opt("Harga normal", col)),
                     "**" + rp(opt("Harga total (DPP)", col)), rp(opt("Harga per peserta", col))])
table(rows, [4.6, 4.4, 2.6, 2.7, 2.2], align_right=(2, 3, 4), size=9.5)
add("body", f"Harga untuk {N:,}".replace(",", ".") + " peserta dengan pelaksanaan jendela 7 hari dan sudah termasuk pajak yang berlaku. "
    f"Untuk pelaksanaan serentak, tambah {rp(H['P_SER'])} per peserta.")
add("body", "Harga khusus Paket 3 untuk BSN kami sampaikan setelah pertemuan lingkup, karena harganya bergantung pada jumlah unit, "
    "rumpun jabatan, dan data kinerja yang akan dipakai. Harga normal di atas menjadi batas atas.")

add("h2", "Isi Setiap Paket")
add("minor", "Paket 1: Laporan untuk Perusahaan")
add("bullet", ("Tes BFI (Big Five Inventory)", True), ": mengukur lima dimensi kepribadian, yaitu keterbukaan, kehati-hatian, "
    "ekstraversi, keramahan, dan kestabilan emosi.")
add("bullet", ("Tes DISC", True), ": mengukur gaya kerja dan gaya komunikasi, yaitu Dominance, Influence, Steadiness, dan Compliance.")
add("bullet", ("Laporan otomatis per karyawan", True), ", keluar langsung setelah tes selesai, dengan bahasa yang mudah dipahami. "
    "Laporan hanya dibaca perusahaan.")
add("bullet", ("Skor setiap karyawan", True), " di dashboard HC dan dalam file Excel.")
add("bullet", ("Undangan dan pengingat lewat email", True), ", serta bantuan teknis untuk peserta selama periode tes.")

add("minor", "Paket 2: Laporan untuk Perusahaan dan Karyawan")
add("body", "Semua isi Paket 1, ditambah:")
add("bullet", ("Karyawan juga membaca laporannya sendiri", True), ", sehingga hasil tes bisa langsung dipakai untuk pengembangan diri.")
add("bullet", ("Akses laporan untuk setiap karyawan", True), " dan bantuan teknis untuk karyawan yang membuka laporannya.")

add("minor", "Paket 3: Pemetaan oleh Psikolog Teman Journey")
add("body", "Paket 3 punya tiga tingkat. Semua tingkat sudah termasuk isi Paket 2.")
add("body", ("3A. Pemetaan Umum", True))
add("bullet", ("Pertemuan awal", True), " dengan HC BSN untuk menyepakati lingkup dan unit yang dipetakan.")
add("bullet", ("Analisis profil BFI dan DISC per unit untuk 10 unit", True), " (direktorat atau wilayah), lengkap dengan tabel dan grafik.")
add("bullet", ("Laporan pemetaan", True), ": gambaran umum BSN, ringkasan per unit, kekuatan dan area pengembangan tiap unit.")
add("bullet", ("Presentasi hasil", True), " kepada direksi atau HC BSN oleh psikolog.")
add("body", ("3B. Pemetaan + Talent Box", True), ": semua isi 3A, ditambah:")
add("bullet", ("Talent box (9 kotak)", True), ": karyawan dipetakan berdasarkan kinerja (dari data BSN) dan potensi (dari hasil tes).")
add("bullet", ("Kalibrasi talent box untuk 10 unit", True), " bersama HC dan atasan unit, dipandu psikolog secara online.")
add("bullet", ("Review psikolog untuk talenta kunci", True), " (sekitar 10% peserta).")
add("body", ("3C. Pemetaan Talenta Lengkap", True), ": semua isi 3B, ditambah:")
add("bullet", ("Analisis jabatan untuk 15 rumpun jabatan", True), ". Psikolog menelaah jobdesk dan mewawancarai narasumber, lalu "
    "menyusun profil jabatan: tugas utama, kompetensi, dan profil kepribadian yang cocok.")
add("bullet", ("Skor kecocokan orang dan jabatan", True), ": setiap karyawan dibandingkan dengan profil jabatannya.")
add("bullet", ("Kalibrasi dipandu 2 psikolog", True), ", dan review psikolog diperluas ke karyawan yang perlu perhatian "
    "(total sekitar 20% peserta).")

add("h2", "Harga Tes Turun Sesuai Jumlah Peserta")
add("body", "Harga per peserta untuk pelaksanaan jendela 7 hari. Harga normal berlaku untuk kurang dari 100 peserta.", keep_next=True)
rows = [["Jumlah peserta", "Paket 1: Laporan untuk Perusahaan", "Paket 2: Laporan untuk Perusahaan dan Karyawan"],
        ["Harga normal", "~~" + rp(H["SLIST"]), "~~" + rp(H["ALIST"])]]
for qty, label in [(100, "Mulai 100"), (500, "Mulai 500"), (1000, "Mulai 1.000")]:
    rows.append([label, rp(H[f"S{qty}"]), rp(H[f"A{qty}"])])
rows.append(["**Mulai 2.000 (BSN)", "**" + rp(H["S2000"]), "**" + rp(H["A2000"])])
table(rows, [4.5, 6.0, 6.0], align_right=(1, 2))

add("body", "Paket 3 memakai harga Paket 2, ditambah biaya pemetaan berikut.", keep_next=True)
rows = [["Tingkat pemetaan", "Harga normal", "Harga untuk BSN"]]
for lvl, name in [(1, "3A. Pemetaan Umum"), (2, "3B. Pemetaan + Talent Box"), (3, "3C. Pemetaan Talenta Lengkap")]:
    rows.append([name, "~~" + rp(H[f"P_MAPL{lvl}"]), "**" + AFTER_MEETING])
table(rows, [7.5, 4.5, 4.5], align_right=(1, 2))

add("h2", "Layanan Tambahan Sesuai Permintaan")
add("body", "Layanan berikut tidak termasuk dalam paket dan bisa dipesan kapan saja.", keep_next=True)
table([
    ["Layanan", "Harga"],
    ["Sesi penjelasan hasil 90 menit dengan 2 psikolog (online), untuk HC atau atasan", rp(H["P_BRF"]) + " per sesi"],
    ["Sesi umpan balik individu 30 menit dengan psikolog (online)", rp(H["P_FB"]) + " per orang"],
    ["Unit tambahan di Paket 3 (di atas 10 unit)", AFTER_MEETING],
    ["Rumpun jabatan tambahan di Paket 3C (di atas 15 rumpun)", AFTER_MEETING],
], [11.5, 5.0], align_right=(1,))

add("h2", "Perbandingan dengan Pilihan Lain di Pasar")
table([
    ["", "Teman Journey x Coreitera", "Platform psikotes online biasa", "Alat tes mandiri untuk psikolog", "Biro psikologi"],
    ["Harga tes + laporan per peserta", "**" + rp(H["S2000"]) + " sampai " + rp(H["A2000"]), "Rp45.000 sampai Rp75.000",
     "Rp20.000 sampai Rp40.000 (tanpa laporan)", "Rp200.000 sampai Rp250.000"],
    ["Tes BFI dan DISC", "**Ya, keduanya", "Tidak selalu", "Tidak selalu", "Tergantung paket"],
    ["Laporan per karyawan", "**Instan, bahasa mudah dipahami", "Psikogram otomatis", "Tidak ada, ditulis psikolog sendiri",
     "Ya, tetapi butuh waktu lebih lama"],
    ["Siapa yang membaca laporan", "**Bisa pilih: hanya perusahaan (Paket 1), atau perusahaan dan karyawan (Paket 2)",
     "Biasanya hanya perusahaan", "Hanya psikolog", "Biasanya hanya perusahaan"],
    ["Analisis jabatan dan talent box", "**Ya (Paket 3B dan 3C)", "Tidak", "Tidak", "Ya, biasanya dengan biaya assessment center"],
    ["Sesi dengan psikolog", "**Tersedia sesuai permintaan", "Tidak", "Tidak", "Ya"],
    ["Siap untuk 2.000 peserta", "**Ya, jendela 7 hari atau serentak", "Ya", "Terbatas jumlah paket", "Sulit dalam waktu singkat"],
], [3.3, 3.7, 3.1, 3.2, 3.2], size=9)
add("body", "Kisaran harga pasar diambil dari harga publik penyedia di Indonesia (November 2026). Harga untuk jumlah besar bisa "
    "berbeda sesuai negosiasi.")

add("h2", "Dua Cara Pelaksanaan")
table([
    ["", "Jendela 7 hari", "Serentak"],
    ["Cara kerja", "Karyawan menerima email dan mengerjakan tes kapan saja dalam 7 hari",
     "Seluruh karyawan mengerjakan tes pada jadwal yang sama, misalnya 2 gelombang di 2 hari"],
    ["Cocok untuk", "Karyawan tersebar di banyak cabang dan shift", "BSN ingin tes diawasi atau dilakukan di waktu yang sama"],
    ["Yang kami siapkan", "Platform standar dan bantuan teknis",
     "Kapasitas server tambahan, uji beban sebelum hari H, dan tim teknis siaga"],
    ["Biaya tambahan", "Tidak ada", rp(H["P_SER"]) + " per peserta"],
], [3.5, 6.5, 6.5], size=10)

add("h2", "Jadwal Kerja (Perkiraan)")
table([
    ["Minggu", "Paket 1 dan 2", "Paket 3"],
    ["1", "Kontrak, data karyawan, penyiapan akun, uji coba", "Kontrak, pertemuan awal. 3C: mulai analisis jabatan"],
    ["2", "Pelaksanaan tes", "Pelaksanaan tes. 3C: analisis jabatan berlanjut"],
    ["3", "Penutupan tes, laporan lengkap", "3A: analisis per unit. 3B dan 3C: BSN mengirim data kinerja, talent box disusun"],
    ["4 sampai 5", "", "3A: laporan dan presentasi (selesai minggu 4). 3B dan 3C: kalibrasi per unit, review psikolog"],
    ["6", "", "3B dan 3C: laporan pemetaan dan presentasi"],
], [2.2, 6.0, 8.3], align_center=(0,), size=10)
add("body", "Tanggal pasti disepakati bersama BSN setelah kontrak ditandatangani.")

add("h2", "Tim Psikolog")
add("body", "Pemetaan dilaksanakan oleh psikolog Teman Journey yang tersertifikasi HIMPSI. Psikolog yang bertugas ditetapkan dari tim "
    "berikut dan dikonfirmasi kepada BSN sebelum pelaksanaan.", keep_next=True)
table([
    ["Nama", "Profesi", "S2", "S1"],
    ["Riskia Ramadhina Sukriananda", "Psikolog", "Universitas Indonesia", "Universitas Indonesia"],
    ["Amanda Nurshadrina", "Psikolog", "Universitas Indonesia", "Universitas Padjadjaran"],
    ["Laras Sita Persitia", "Psikolog Klinis Dewasa", "Universitas Indonesia", "Universitas Brawijaya"],
    ["Fasya Khalida Suffah", "Psikolog Klinis Dewasa", "Universitas Indonesia", "Universitas Indonesia"],
    ["Nadira Khairunnisa", "Psikolog Klinis Dewasa", "Universitas Indonesia", "Universitas Indonesia"],
    ["Ninette Putri Mustika", "Psikolog", "Universitas Indonesia", "Universitas Indonesia"],
    ["Rhajiv Nur Ilham", "Psikolog Klinis Dewasa", "Universitas Indonesia", "Universitas Indonesia"],
], [5.0, 3.8, 3.8, 3.9])

add("h2", "Kerahasiaan Data")
add("body", "Data dan hasil tes karyawan hanya digunakan untuk keperluan pemetaan ini dan hanya diserahkan kepada BSN. Setiap karyawan "
    "memberikan persetujuan sebelum mengerjakan tes, termasuk siapa yang dapat membaca hasilnya. Psikolog Teman Journey bekerja sesuai "
    "Kode Etik Psikologi Indonesia (HIMPSI), dan data pribadi dikelola sesuai UU No. 27 Tahun 2022 tentang Pelindungan Data Pribadi.")

add("h2", "Ketentuan")
for line in [
    "Harga tes dihitung dari jumlah peserta dalam satu pesanan. Jika jumlah akhir berubah, harga per peserta mengikuti tier yang sesuai.",
    "Harga khusus Paket 3 disampaikan setelah pertemuan lingkup dengan BSN dan berlaku untuk kontrak yang ditandatangani dalam masa berlaku penawaran.",
    "Paket 3B dan 3C membutuhkan data kinerja karyawan dari BSN untuk menyusun talent box.",
    "Pembayaran dalam dua termin: 50% setelah periode tes ditutup, 50% setelah laporan akhir diserahkan.",
    "Harga sudah termasuk pajak yang berlaku. Penawaran berlaku 30 hari sejak tanggal proposal.",
]:
    add("bullet", line)

# Banned characters in client text (house style): long dashes, semicolons, arrows. Template text is reported too.
BANNED = "\u2014\u2013;\u2192"
bad = [p.text for p in d.paragraphs if any(ch in p.text for ch in BANNED)]
bad += [c.text for t in d.tables for row in t.rows for c in row.cells if any(ch in c.text for ch in BANNED)]
for b in bad:
    print("WARNING banned character:", b[:90])

# Attachments heading starts on a fresh page.
apPr = anchor.find(qn("w:pPr"))
if apPr is not None:
    ppr_flag(apPr, "w:pageBreakBefore")

# ---------------------------------------------------------------- cover image
swapped = False
for section in d.sections:
    for hdr in (section.first_page_header, section.header):
        for rel in hdr.part.rels.values():
            if rel.reltype.endswith("/image") and rel.target_ref.endswith("image7.png"):
                rel.target_part._blob = COVER.read_bytes()
                swapped = True
assert swapped, "cover image not found"

# The longer TOC pushes the template's empty spacer paragraphs onto a blank page. Drop the spacers
# between the TOC block and the paragraph that holds the page break.
toc = next(el for el in body if el.tag == qn("w:sdt"))
el = toc.getnext()
while el is not None and el.tag == qn("w:p") and not "".join(t.text or "" for t in el.iter(qn("w:t"))).strip() \
        and el.find(".//" + qn("w:br")) is None:
    nxt = el.getnext()
    body.remove(el)
    el = nxt

# The template's TOC lists styles by name, which Word on Mac does not match. Use outline levels.
for instr in body.iter(qn("w:instrText")):
    if instr.text and instr.text.strip().startswith("TOC"):
        instr.text = ' TOC \\o "1-2" \\h \\z \\u '

d.core_properties.title ="Proposal Pemetaan Karyawan dengan Tes BFI dan DISC - Bank Syariah Nasional"
d.save(out)
print("saved", out)
