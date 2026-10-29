"""Builds the BI Fit and Proper Test proposal from the CBP Rupiah proposal template.

Usage: python build_proposal.py <template.docx> <out.docx>
Keeps the cover frame, profile, legality, products and attachments of the template,
replaces the program section, and reads service prices from the RAB workbook
("Harga Proposal" tab), so the proposal never shows internal costs.
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
RAB = HERE.parent / "RAB - BI Fit and Proper Test (Internal).xlsx"
COVER = HERE / "cover_fit_proper.png"
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
    return cell


def table(rows, col_widths_cm, footer=(), align_right=(), align_center=()):
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
                cell_text(cell, value, color=WHITE, bold=True, size=10, align=align)
            elif style == "dark":
                shade(cell, HEAD_HEX)
                cell_text(cell, value, color=WHITE, bold=True, size=11.5, align=2 if j else None)
            elif style == "light":
                shade(cell, ZEBRA_HEX)
                cell_text(cell, value, color=NAVY, bold=True, align=2 if j else None)
            else:
                if i % 2 == 0:
                    shade(cell, STRIPE_HEX)
                cell_text(cell, value, color=NAVY, align=align)
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
hp = wb["Harga Proposal"]
services = []
for row in hp.iter_rows(min_row=5, max_row=9, values_only=True):
    no, name, vol, unit, price, total = row[:6]
    if total is None:
        raise SystemExit("RAB has no cached values. Recalculate it in Excel first.")
    services.append((no, name, vol, unit, price, total))
grand = sum(s[5] for s in services)
peserta = wb["Asumsi"]["B3"].value
agreed = wb["Asumsi"]["B4"].value
assert grand == agreed, (grand, agreed)

# ---------------------------------------------------------------- content
add("h1", "Asesmen dalam Rangka Fit and Proper Test", page_break=True)

add("h2", "Latar Belakang Program")
add("body", "Fit and proper test adalah proses untuk menilai apakah seorang kandidat layak dan tepat untuk menempati suatu jabatan. "
    "Penilaian ini tidak cukup hanya melihat riwayat pendidikan dan pengalaman kerja. Organisasi juga perlu memahami cara kandidat "
    "berpikir, mengambil keputusan, bekerja sama, dan menjaga integritas ketika berada di bawah tekanan.")
add("body", "Asesmen psikologis membantu Bank Indonesia mendapatkan gambaran tersebut secara objektif dan terukur. Dengan memadukan "
    "tes psikologi, diskusi kelompok, dan wawancara, setiap kesimpulan tentang kandidat didukung oleh lebih dari satu sumber bukti. "
    "Cara ini mengurangi bias penilaian dan membuat hasil asesmen lebih mudah dipertanggungjawabkan.")
add("body", "Teman Journey menawarkan program asesmen yang rubrik dan materinya disusun bersama Bank Indonesia. Kompetensi yang dinilai "
    "mengikuti kebutuhan jabatan yang akan diisi, bukan paket standar yang sama untuk semua organisasi.")

add("sub", "Landasan Pendekatan")
add("minor", "Metode Assessment Center")
add("body", "Assessment center adalah pendekatan asesmen yang menggunakan beberapa metode dan beberapa asesor untuk menilai kompetensi "
    "yang sama. Pedoman internasional assessment center menekankan bahwa kompetensi yang dinilai harus jelas sejak awal, perilaku "
    "kandidat diamati melalui simulasi, dan kesimpulan akhir diambil melalui integrasi penilaian para asesor. "
    "(Sumber: International Taskforce on Assessment Center Guidelines, 2015, Journal of Management, 41(4).)")
add("minor", "Wawancara Terstruktur Berbasis Perilaku")
add("body", "Wawancara terstruktur menggunakan pertanyaan dan rubrik yang sama untuk setiap kandidat. Kandidat diminta menceritakan "
    "situasi nyata yang pernah dihadapi, tindakan yang diambil, dan hasilnya. Penelitian meta-analisis menunjukkan wawancara "
    "terstruktur termasuk metode seleksi dengan daya prediksi kinerja yang tinggi. (Sumber: Schmidt & Hunter, 1998, Psychological Bulletin, 124(2).)")
add("minor", "Leaderless Group Discussion (LGD)")
add("body", "Dalam LGD, kandidat mendiskusikan sebuah studi kasus dalam kelompok tanpa ketua yang ditunjuk. Psikolog mengamati perilaku "
    "yang muncul secara alami, seperti inisiatif, cara menyampaikan pendapat, cara mendengarkan, dan cara kelompok mencapai keputusan.")
add("minor", "Big Five Inventory (BFI) dan DISC")
add("body", "BFI mengukur lima dimensi kepribadian yang paling banyak diteliti, yaitu keterbukaan, kehati-hatian, ekstraversi, keramahan, "
    "dan stabilitas emosi. DISC menggambarkan gaya perilaku kerja dan komunikasi. Hasil kedua tes menjadi data awal yang kemudian "
    "dikonfirmasi melalui LGD dan wawancara.")

add("sub", "Tujuan Program")
for line in [
    "Memberikan gambaran objektif tentang kompetensi, gaya kerja, dan kesiapan setiap kandidat untuk jabatan yang dituju.",
    "Menilai setiap kandidat dengan rubrik yang sama, sehingga hasil antarkandidat dapat dibandingkan secara adil.",
    "Memberikan rekomendasi tertulis dari psikolog sebagai bahan pertimbangan Bank Indonesia dalam mengambil keputusan.",
    "Memberikan saran pengembangan untuk setiap kandidat, termasuk kandidat yang direkomendasikan.",
]:
    add("bullet", line)

add("h2", "Target Peserta")
add("body", f"{peserta} (enam) kandidat yang ditunjuk oleh Bank Indonesia untuk mengikuti fit and proper test.")

add("h2", "Komponen Layanan")
add("body", "Program terdiri atas lima layanan. Setiap peserta mengikuti seluruh rangkaian.")

add("minor", "Layanan 1 - Penyesuaian Kebutuhan dan Penyusunan Rubrik Penilaian")
add("body", "Sebelum asesmen dimulai, psikolog Teman Journey bertemu dengan tim Bank Indonesia untuk menyepakati jabatan yang dituju, "
    "kompetensi yang akan dinilai, dan format laporan yang dibutuhkan. Dari pertemuan tersebut, kami menyusun:")
add("bullet", ("Rubrik penilaian wawancara", True), ", berisi kompetensi, indikator perilaku, dan skala nilai.")
add("bullet", ("Rubrik penilaian LGD", True), ", berisi indikator perilaku yang diamati selama diskusi.")
add("bullet", ("Studi kasus dan instruksi LGD", True), " yang relevan dengan konteks kerja Bank Indonesia.")
add("body", "Rubrik disusun oleh dua psikolog dan dikirim ke Bank Indonesia untuk disetujui sebelum digunakan.")

add("minor", "Layanan 2 - Tes Psikologi BFI dan DISC")
add("body", "Peserta mengerjakan tes BFI dan DISC secara online sebelum hari pelaksanaan, dengan durasi sekitar 30 sampai 45 menit. "
    "Hasil tes dibaca oleh psikolog dan menjadi bahan awal untuk wawancara.")

add("minor", "Layanan 3 - Leaderless Group Discussion (LGD)")
add("body", "Enam peserta mengikuti satu sesi LGD selama sekitar 2 jam, termasuk pengarahan dan penutupan. Sesi dipandu oleh dua "
    "psikolog, yaitu satu psikolog fasilitator dan satu psikolog asesor. Setiap psikolog mengamati tiga peserta, sehingga perilaku "
    "setiap peserta tercatat secara rinci pada rubrik.")

add("minor", "Layanan 4 - Wawancara oleh Psikolog Juri Utama")
add("body", "Setiap peserta mengikuti wawancara terstruktur selama 60 menit bersama psikolog juri utama Teman Journey. Pertanyaan disusun "
    "dari rubrik yang telah disetujui Bank Indonesia serta temuan dari tes dan LGD, sehingga wawancara dapat menggali hal yang belum "
    "terlihat jelas.")

add("minor", "Layanan 5 - Laporan Asesmen Komprehensif dan Rekomendasi")
add("body", "Para psikolog memadukan hasil tes, LGD, dan wawancara dalam rapat integrasi. Hasilnya dituangkan dalam laporan untuk "
    "setiap peserta yang berisi:")
for line in [
    "Profil kepribadian dan gaya kerja.",
    "Nilai setiap kompetensi beserta bukti perilaku dari LGD dan wawancara.",
    "Kekuatan dan area pengembangan.",
    "Rekomendasi akhir dalam tiga kategori: Disarankan, Disarankan dengan Catatan, atau Belum Disarankan.",
]:
    add("bullet", line)
add("body", "Setiap laporan direview oleh psikolog juri utama sebelum dikirim. Bank Indonesia juga menerima satu tabel ringkasan yang "
    "membandingkan hasil seluruh peserta.")

add("h2", "Alur Pelaksanaan")
table([
    ["Tahap", "Kegiatan", "Waktu"],
    ["1", "Penandatanganan SPK dan pertemuan penyesuaian kebutuhan", "Minggu ke-1"],
    ["2", "Penyusunan rubrik dan studi kasus, lalu persetujuan Bank Indonesia", "Minggu ke-1 sampai ke-2"],
    ["3", "Tes BFI dan DISC secara online", "Paling lambat 2 hari sebelum pelaksanaan"],
    ["4", "LGD dan wawancara", "Hari pelaksanaan (1 sampai 2 hari)"],
    ["5", "Rapat integrasi dan penulisan laporan", "Setelah pelaksanaan"],
    ["6", "Penyerahan laporan dan tabel ringkasan", "Paling lambat 7 hari kerja setelah pelaksanaan"],
], [1.6, 9.2, 5.7], align_center=(0,))
add("body", "Tanggal pasti disepakati bersama Bank Indonesia setelah SPK ditandatangani.")

add("h2", "Hasil yang Diterima Bank Indonesia")
for line in [
    "Rubrik penilaian wawancara dan LGD yang disesuaikan dengan kebutuhan Bank Indonesia.",
    f"{peserta} laporan asesmen komprehensif, satu untuk setiap peserta.",
    "1 tabel ringkasan hasil dan rekomendasi seluruh peserta.",
]:
    add("bullet", line)

add("h2", "Tim Psikolog")
add("body", "Asesmen dilaksanakan oleh psikolog Teman Journey yang tersertifikasi HIMPSI. Psikolog juri utama dan psikolog fasilitator "
    "ditetapkan dari tim berikut dan dikonfirmasi kepada Bank Indonesia sebelum pelaksanaan.", keep_next=True)
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
add("body", "Data dan hasil asesmen peserta hanya digunakan untuk keperluan fit and proper test ini dan hanya diserahkan kepada Bank "
    "Indonesia. Psikolog Teman Journey bekerja sesuai Kode Etik Psikologi Indonesia (HIMPSI), dan data pribadi peserta dikelola sesuai "
    "UU No. 27 Tahun 2022 tentang Pelindungan Data Pribadi.")

add("h2", "Proposal Anggaran", page_break=True)
add("body", "Rincian harga untuk setiap layanan adalah sebagai berikut.")
# One line under each service name so BI sees what the price buys.
WHAT = {
    1: "Pertemuan kebutuhan, rubrik wawancara, rubrik LGD, studi kasus LGD",
    2: "Tes online, hasil dibaca oleh psikolog",
    3: "1 kelompok, 2 psikolog (fasilitator dan asesor), sekitar 2 jam",
    4: "Wawancara terstruktur 60 menit untuk setiap peserta",
    5: "Laporan per peserta, rapat integrasi, tabel ringkasan",
}
rows = [["No", "Layanan", "Volume", "Harga Satuan", "Jumlah"]]
for no, name, vol, unit, price, total in services:
    rows.append([str(no), (name[0].upper() + name[1:], WHAT[no]), f"{vol} {unit}", rupiah(price), rupiah(total)])
rows.append([f"Total untuk {peserta} peserta*", "", "", "", rupiah(grand)])
rows.append(["Harga per peserta", "", "", "", rupiah(grand / peserta)])
table(rows, [0.9, 7.2, 2.3, 2.8, 3.3], footer=("dark", "light"), align_right=(3, 4), align_center=(0, 2))
add("body", "*Harga sudah termasuk pajak yang berlaku. Ketentuan pembayaran mengikuti SPK. Penawaran berlaku 30 hari sejak tanggal proposal.")

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

# The template's TOC lists styles by name, which Word on Mac does not match. Use outline levels.
for instr in body.iter(qn("w:instrText")):
    if instr.text and instr.text.strip().startswith("TOC"):
        instr.text = ' TOC \\o "1-2" \\h \\z \\u '

d.core_properties.title ="Proposal Asesmen dalam Rangka Fit and Proper Test - Bank Indonesia"
d.save(out)
print("saved", out)
