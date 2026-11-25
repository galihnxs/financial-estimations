"""Builds the BSN proposal UPDATE (24 Nov 2026): assessment licence status, SME fallback, rotation support.

Usage: python build_update.py <template.docx> <out.docx>
Same template and helpers as build_proposal.py. Service prices come from ProductKnowledge (20 Nov 2026 catalog, corporate column).
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
COVER = HERE / "cover_bsn_update.png"
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


# ---------------------------------------------------------------- content (update of 24 Nov 2026)
# Prices come from ProductKnowledge/01_Daftar_Produk_Final.md (Master_Produk_dan_Harga.xlsx, 20 Nov 2026),
# corporate column. BSN is a corporate client. Old BSN prices (P_BRF) still come from the BSN RAB.
from docx.shared import Cm

PK = HERE.parent.parent / "ProductKnowledge" / "deck_visuals"
wb = openpyxl.load_workbook(RAB, data_only=True)
H = {row[0]: row[2] for row in wb["Harga"].iter_rows(min_row=3, values_only=True) if row[0]}
rp = rupiah

CATALOG = {  # code: corporate price
    "TJ-ASM-11": 175_000,     # MHCU Skrining
    "TJ-ASM-08": 600_000,     # MHCU Flash, konseling 30 menit
    "TJ-ASM-09": 1_000_000,   # MHCU Deep, konseling 60 menit
    "TJ-KON-01": 500_000,     # konseling B2B 30 menit
    "TJ-EAP-05": 925_000,     # konseling 60 menit di luar kuota
    "TJ-EAP-06": 13_500_000,  # EAP Korporat per bulan, 300 pegawai
    "TJ-WS-01": 8_100_000,    # Workshop Healthy Working Environment 2 jam
    "TJ-KNS-01": 4_700_000,   # Hari Konsultan
    "TJ-KNS-02": 3_200_000,   # FGD 2 jam
    "TJ-ASM-13": 2_000_000,   # Asesmen Manajerial (wawancara psikolog 1 jam + laporan)
    "TJ-ASM-14": 7_600_000,   # Asesmen Eksekutif
}
C = {k: rp(v) for k, v in CATALOG.items()}


def picture(path, width_cm=16.0):
    p = d.add_paragraph()
    p.alignment = 1
    p.paragraph_format.keep_with_next = False
    p.add_run().add_picture(str(path), width=Cm(width_cm))
    anchor.addprevious(p._p)


add("h1", "Pembaruan Proposal untuk BSN", page_break=True)

add("h2", "Ringkasan Pembaruan")
add("body", "Terima kasih atas diskusi yang sudah berjalan bersama tim Human Capital BSN. Dokumen ini memperbarui proposal "
    "Pemetaan Karyawan tanggal 4 November 2026 dan menjadi pembaruan akhir dari kami. Ada tiga hal yang ingin kami sampaikan.")
add("bullet", ("Status layanan asesmen. ", True), "Kami sedang menyelesaikan lisensi validasi instrumen dengan penyedia kami, "
    "dengan target selesai pertengahan Januari 2027. Selama proses ini, ada kemungkinan kami belum bisa memberikan layanan "
    "asesmen seperti yang kami tawarkan sebelumnya.")
add("bullet", ("Tetap mendampingi BSN. ", True), "Jika BSN memilih penyedia asesmen lain, kami tetap siap bekerja sama sebagai "
    "pendamping ahli untuk pemetaan talenta dan wawancara psikolog bagi kandidat terpilih.")
add("bullet", ("Dukungan untuk karyawan. ", True), "Jika BSN menjalankan rotasi atau perubahan peran, kami menyediakan "
    "pemeriksaan kesehatan mental, konseling, dan program pendampingan karyawan (EAP).")

add("h2", "1. Status Layanan Asesmen")
add("body", "Kami ingin terbuka kepada BSN. Hasil asesmen akan dipakai untuk keputusan tentang orang, seperti penempatan, rotasi, "
    "dan pengembangan. Karena itu setiap instrumen yang kami pakai harus memiliki lisensi dan validasi yang sah untuk penggunaan "
    "di perusahaan. Saat ini lisensi validasi instrumen kami sedang kami selesaikan bersama penyedia instrumen. Kami tidak ingin "
    "BSN memakai hasil asesmen yang dasarnya belum lengkap, sehingga kami menyampaikan kondisi ini sejak awal.")
add("body", "Kebutuhan asesmen BSN sudah kita rumuskan bersama. Berikut posisi kami untuk setiap kebutuhan tersebut.", keep_next=True)
table([
    ["Kebutuhan asesmen BSN", "Posisi Teman Journey", "Target"],
    ["Kepribadian (Big Five Inventory)", "Instrumen tersedia. Lisensi validasi sedang diselesaikan", "Pertengahan Januari 2027"],
    ["Gaya kerja dan komunikasi (DISC)", "Instrumen tersedia. Lisensi validasi sedang diselesaikan", "Pertengahan Januari 2027"],
    ["Gaya kepemimpinan", "Belum tersedia. Kami sedang memperluas lisensi agar laporan gaya kepemimpinan juga bisa kami berikan",
     "Diproses bersama penyelesaian lisensi"],
], [5.2, 7.3, 4.0], size=10)
add("body", "Lingkup kebutuhan yang sudah kita sepakati tetap kami simpan. Begitu lisensi selesai, kami akan mengonfirmasi jadwal "
    "dan harga kepada BSN, sehingga pekerjaan bisa langsung dimulai tanpa mengulang diskusi dari awal.")

add("minor", "Mencoba layanan kami yang lain, tanpa biaya")
add("body", "Selama menunggu, kami ingin BSN tetap bisa merasakan cara kerja tim kami. Kami menawarkan contoh layanan berikut "
    "tanpa biaya, agar HC BSN bisa menilai sendiri kualitas layanan kami sebelum mengambil keputusan.")
add("bullet", ("Snapshot burnout untuk satu unit kerja. ", True), "Karyawan mengisi kuesioner singkat secara online, lalu HC "
    "menerima laporan agregat tentang tingkat kelelahan kerja di unit tersebut. Tidak ada data individu yang dibuka.")
add("bullet", ("Mental Health Check-Up Flash untuk 5 orang tim HC. ", True), "Lima instrumen klinis, laporan personal yang "
    "ditinjau psikolog, dan konseling 30 menit, agar tim HC mengalami sendiri layanan yang nanti diterima karyawan.")
add("bullet", ("Contoh laporan. ", True), "Contoh laporan individu dan laporan agregat (dengan data contoh, bukan data karyawan "
    "BSN), agar HC bisa melihat bentuk hasil yang akan diterima.")

add("h2", "2. Jika BSN Memilih Penyedia Asesmen Lain")
add("body", "Kami memahami jika BSN perlu berjalan sesuai jadwal internal dan memilih penyedia asesmen lain. Dalam kondisi itu, kami "
    "tetap siap bekerja sama sebagai pendamping ahli. Psikolog kami bisa membantu BSN membaca dan memakai hasil asesmen dari "
    "penyedia mana pun.", keep_next=True)
table([
    ["Layanan", "Isi", "Harga"],
    [("Pendamping ahli pemetaan talenta", "Hari Konsultan"),
     "Menyusun kerangka pemetaan dan profil jabatan, membaca hasil asesmen, memandu kalibrasi talent box bersama HC. "
     "Satu hari berisi 4 jam tatap muka dan 4 jam analisis oleh psikolog senior", C["TJ-KNS-01"] + " per hari"],
    [("Diskusi kelompok terarah", "FGD 2 jam, online"),
     "Menggali kebutuhan jabatan atau menguji hasil pemetaan bersama atasan unit, dipandu psikolog dan notulen",
     C["TJ-KNS-02"] + " per FGD"],
    [("Wawancara psikolog untuk kandidat terpilih", "Jabatan manajerial"),
     "Wawancara 1 jam oleh psikolog, lalu laporan dan rekomendasi per kandidat", C["TJ-ASM-13"] + " per kandidat"],
    [("Asesmen kandidat eksekutif", "Jabatan senior"),
     "Wawancara mendalam 2 jam dan simulasi bisnis dengan 2 psikolog, laporan untuk jabatan senior",
     C["TJ-ASM-14"] + " per kandidat"],
    [("Sesi penjelasan hasil", "90 menit, online"), "Dua psikolog menjelaskan hasil pemetaan kepada HC atau atasan",
     rp(H["P_BRF"]) + " per sesi"],
], [4.6, 8.4, 3.5], align_right=(2,), size=9.5)

add("h2", "3. Dukungan untuk Rotasi dan Perubahan Peran")
add("body", "Rotasi dan perubahan peran membawa tuntutan baru: tim baru, atasan baru, dan kadang kota baru. Masa ini sering menjadi "
    "masa yang berat bagi karyawan, walaupun jarang terlihat. Bersama Coreitera, kami menyediakan layanan yang bisa membantu "
    "karyawan BSN melewati masa transisi dengan sehat.")
picture(PK / "p10_mhcu.png", 15.5)
table([
    ["Layanan", "Isi", "Harga"],
    [("Mental Health Check-Up Skrining", "Untuk banyak orang sekaligus"),
     "Kuesioner DASS-21 online, ditinjau psikolog berizin, laporan individual paling lambat 48 jam", C["TJ-ASM-11"] + " per orang"],
    [("Mental Health Check-Up Flash", "Dengan konseling 30 menit"),
     "Lima instrumen klinis, laporan personal yang ditinjau psikolog, dan konseling 30 menit", C["TJ-ASM-08"] + " per orang"],
    [("Mental Health Check-Up Deep", "Dengan konseling 60 menit"),
     "Sama dengan Flash, dengan konseling 60 menit. Cocok untuk karyawan kunci dan kasus prioritas", C["TJ-ASM-09"] + " per orang"],
    [("Konseling online", "30 atau 60 menit"), "Sesi konseling dengan psikolog, sesuai kebutuhan karyawan",
     C["TJ-KON-01"] + " (30 menit), " + C["TJ-EAP-05"] + " (60 menit)"],
    [("Program Pendampingan Karyawan (EAP)", "Untuk 300 karyawan"),
     "10 sesi konseling per bulan untuk dipakai bersama, check-up 2 kali setahun untuk semua karyawan, laporan agregat untuk HC, "
     "dan 1 webinar per kuartal", C["TJ-EAP-06"] + " per bulan"],
    [("Workshop lingkungan kerja yang sehat", "2 jam, tatap muka"), "Untuk 30 peserta, dipandu psikolog dan co-fasilitator",
     C["TJ-WS-01"] + " per sesi"],
], [4.6, 8.4, 3.5], align_right=(2,), size=9.5)
add("body", "Untuk rotasi, kami menyarankan alur sederhana: skrining untuk semua karyawan yang berpindah, lalu Flash atau Deep untuk "
    "karyawan yang hasilnya perlu perhatian. Dengan cara ini biaya tetap terkendali dan karyawan yang paling membutuhkan "
    "mendapat dukungan lebih dulu. Hasil individu hanya dibaca karyawan dan psikolog. HC menerima laporan agregat.")

add("minor", "Satu mitra untuk kebutuhan lain")
picture(PK / "p07_peta_layanan.png", 15.5)

add("h2", "Langkah Berikutnya")
table([
    ["Waktu", "Langkah"],
    ["Minggu ini", "Diskusi singkat 30 menit dengan HC BSN untuk membahas pembaruan ini dan memilih contoh layanan"],
    ["Desember 2026", "Pelaksanaan contoh layanan tanpa biaya dan penyerahan contoh laporan"],
    ["Pertengahan Januari 2027", "Kami mengabarkan status akhir lisensi asesmen kepada BSN"],
    ["Setelah itu", "BSN memilih jalur kerja sama: asesmen oleh Teman Journey, pendamping ahli, dukungan rotasi, atau gabungan"],
], [4.5, 12.0], size=10)

add("h2", "Tim Psikolog")
add("body", "Layanan dilaksanakan oleh psikolog Teman Journey yang tersertifikasi HIMPSI. Psikolog yang bertugas ditetapkan dari tim "
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
add("body", "Data dan hasil layanan karyawan hanya digunakan untuk keperluan kerja sama ini. Setiap karyawan memberikan persetujuan "
    "sebelum mengikuti layanan, termasuk siapa yang dapat membaca hasilnya. Hasil pemeriksaan kesehatan mental dan isi konseling "
    "individu tidak dibagikan kepada perusahaan. Psikolog Teman Journey bekerja sesuai Kode Etik Psikologi Indonesia (HIMPSI), dan "
    "data pribadi dikelola sesuai UU No. 27 Tahun 2022 tentang Pelindungan Data Pribadi.")

add("h2", "Ketentuan")
for line in [
    "Harga untuk layanan di wilayah Jakarta dan sudah termasuk pajak yang berlaku. Di luar Jakarta, biaya perjalanan dan "
    "akomodasi ditagih sesuai biaya aktual.",
    "Contoh layanan tanpa biaya berlaku satu kali dan dijadwalkan bersama HC BSN.",
    "Harga asesmen pemetaan karyawan akan kami konfirmasi ulang setelah lisensi selesai.",
    "Penawaran ini berlaku 30 hari sejak tanggal dokumen.",
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

d.core_properties.title = "Pembaruan Proposal Pemetaan Karyawan - Bank Syariah Nasional"
d.save(out)
print("saved", out)
