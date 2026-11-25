"""BSN proposal update (24 Nov 2026), designed version in the Katalog / deck_visuals look.

Writes update_deck.html next to this file and prints it to
"../Proposal Teman Journey - Pembaruan Akhir - BSN.pdf" with headless Chrome (16:9 pages).

Prices:
- Paket Asesmen BSN: cost from "[CONFIDENTIAL] RAB Teman Journey x BSN.xlsx" (sheet MR DIY MHCU, E16),
  priced with the corporate rule from ProductKnowledge/05: cost x 0,9 / (0,882 - 0,50), rounded up to Rp1 jt
  (Galih 24 Nov 2026: BSN is corporate, not the 35% row).
- Pendampingan (main highlight): RAB row 5 Rp15 jt + coordinator Rp1 jt, same rule. No day rates (Galih 24 Nov).
- Everything else: ProductKnowledge/01_Daftar_Produk_Final.md, corporate column (20 Nov 2026), only rows at 50% or more.
  PRICE_CHECK asserts the 50% floor for every price with a known cost.
- Sample report images: Customer FE fixture pages /packages/<pkg>/sample-report (fictional people), cropped to sample_*.jpg.
Run: /opt/homebrew/Caskroom/miniforge/base/bin/python build_update_deck.py
"""
import base64
import math
import subprocess
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
FE = ROOT.parent
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
OUT_PDF = ROOT / "Proposal Teman Journey - Pembaruan Akhir - BSN.pdf"

# ------------------------------------------------------------------ numbers
RAB_COST = 44_500_000          # E16 TOTAL BIAYA
PARTICIPANTS = 300
CORP_MARGIN = 0.50


def corp_price(cost):
    """Corporate rule (ProductKnowledge/05), rounded up to Rp1 jt."""
    return math.ceil(cost * 0.9 / (0.882 - CORP_MARGIN) / 1_000_000) * 1_000_000


def net_margin(cost, price):
    return 0.882 - 0.9 * cost / price


# Assessment tests: two options shown side by side (Galih 24 Nov 2026).
# Pilihan A, ours (Coreitera reports): DISC, BFI, gaya kepemimpinan at Rp40rb per test = Rp120rb per karyawan.
#   Available after the licence (target mid-January 2027). Licence cost is not known yet, so no margin check.
#   Galih said "120k before PPN": TJ and Coreitera are Non-PKP, so the deck adds no PPN (confirm with Galih).
OUR_TEST_PRICE, OUR_TESTS = 40_000, 3
OUR_PRICE = OUR_TEST_PRICE * OUR_TESTS
# Pilihan B, partner Faxtor (BFI, FTPI, OPTI), billed by TJ. Faxtor price list 2026, excl. tax: BFI 30k, FTPI 30k,
# OPTI 50k. Excel recap add-on 35k per participant (Faxtor's message). Tax 11% that TJ (Non-PKP) cannot reclaim.
# Lower margin on purpose (Galih 24 Nov 2026): 31% net after tax, profit comes from pendampingan and other services.
FAX_TESTS = {"BFI": 30_000, "FTPI": 30_000, "OPTI": 50_000}
FAX_RECAP, FAX_TAX, FAX_MARGIN = 35_000, 0.11, 0.31
FAX_COST = (sum(FAX_TESTS.values()) + FAX_RECAP) * (1 + FAX_TAX)                 # Rp160.950 per karyawan
FAX_PRICE = math.ceil(FAX_COST * 0.9 / (0.882 - FAX_MARGIN) / 5_000) * 5_000     # Rp255.000 per karyawan
assert net_margin(FAX_COST, FAX_PRICE) >= FAX_MARGIN

# Pendampingan (main highlight, for both test options). 4 deliverables (Galih 24 Nov 2026): mapping alignment,
# rekomendasi + group report, talent box, talent analytics + rekomendasi tindak lanjut per kelompok.
# Price Rp75 jt chosen by Galih (about 56% net). Never a day rate.
PIO_REVIEW_HOURS, PSY_RATE, DOC_RATE = 20, 350_000, 110_000
PEND_COST_LINES = {
    "Group report agregat dan 3 sesi HR advisory PIO (RAB lama)": 15_000_000,
    "Telaah peserta, 20 jam psikolog PIO": PIO_REVIEW_HOURS * PSY_RATE,
    "Koordinator proyek": 1_000_000,
    "Talent box: olah data kinerja BSN, 8 jam kerja dokumen": 8 * DOC_RATE,
    "Kalibrasi talent box bersama HC, 2 jam x 2 psikolog": 2 * 2 * PSY_RATE,
    "Talent analytics per kelompok, 10 jam kerja dokumen": 10 * DOC_RATE,
}
PEND_COST = sum(PEND_COST_LINES.values())          # Rp26,38 jt
PEND_PRICE = 75_000_000

# Fit and Proper Test, 6 candidates, Galih's price (24 Nov 2026). Cost = BI RAB Rp8,77 jt minus
# BFI+DISC Rp720rb (candidates reuse their mapping test). Below the 50% floor by decision: ~39.9%.
FP_PRICE, FP_PAX = 15_000_000, 6
FP_COST = 8_770_000 - FP_PAX * 120_000
BELOW_FLOOR_BY_DECISION = {"Fit and Proper Test 6 kandidat": (FP_COST, FP_PRICE)}

# Pilot services (no free offers): MHCU Flash at catalog price, burnout snapshot costed here.
FLASH_COST, FLASH_PRICE, FLASH_PAX = 244_500, 600_000, 5
SNAP_PAX = 50                                          # one work unit, up to 50 employees
SNAP_COST = 6 * DOC_RATE + 1 * PSY_RATE                # aggregate report 6 h + 1 h online briefing
SNAP_PRICE = 2_500_000

# Every TJ service price with a known cost must clear the 50% corporate net margin.
PRICE_CHECK = {
    "Pendampingan": (PEND_COST, PEND_PRICE),
    "Pilot MHCU Flash per orang": (FLASH_COST, FLASH_PRICE),
    "Pilot snapshot burnout per unit": (SNAP_COST, SNAP_PRICE),
}
for name, (c, p) in PRICE_CHECK.items():
    assert net_margin(c, p) >= CORP_MARGIN, f"{name} below 50%: {net_margin(c, p):.1%}"


def rp(n):
    return "Rp" + f"{round(n):,}".replace(",", ".")


def img64(path):
    ext = Path(path).suffix.lstrip(".").replace("jpg", "jpeg")
    return f"data:image/{ext};base64," + base64.b64encode(Path(path).read_bytes()).decode()


LOGO = img64(HERE / "tj_logo_white.png")
SAMPLE = {n: img64(HERE / f"sample_{n}.jpg") for n in ("t4d_a", "t4d_b", "mhcu_a", "mhcu_b")}  # fictional fixture data
# Faxtor FTPI report from Faxtor's own redacted sample (Scribd), name blacked out by Faxtor and masked again by us.
SAMPLE["ftpi"] = img64(HERE / "sample_ftpi.jpg")
# Visuals from the [MASTER 2026] client deck (ProductKnowledge): slide 13 F&P collage, slides 21 and 11 mental health.
SAMPLE["fp_collage"] = img64(HERE / "master_fp_collage.jpg")
SAMPLE["mh_wellbeing"] = img64(HERE / "master_mh_wellbeing.jpg")
SAMPLE["mh_tahapan"] = img64(HERE / "master_mh_tahapan.jpg")
PHOTOS = [img64(FE / "_brand" / "cover-photos" / n) for n in
          ("1-tim-teman-journey.webp", "2-pelatihan-korporat.webp", "3-pelatihan-kemendikdasmen.webp")]

NAVY, MID, LIGHT = "#104a70", "#4f5c74", "#628cab"
SAND, SAND2, CREAM, INK, CORAL = "#e8d4ab", "#f6e3b6", "#fdfaf1", "#163b42", "#c8553d"

CSS = f"""
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800;900&display=swap');
@page {{ size: 1920px 1080px; margin: 0 }}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Poppins',sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
.pg{{width:1920px;height:1080px;position:relative;overflow:hidden;page-break-after:always;color:#fff;
  background:linear-gradient(90deg,{NAVY} 0%,{MID} 52%,{LIGHT} 100%)}}
.pg.light{{background:{CREAM};color:{INK}}}
.abs{{position:absolute}}
.blob{{position:absolute;background:{SAND};opacity:.10;border-radius:58% 42% 61% 39%/45% 55% 45% 55%}}
.dots{{position:absolute;background-image:radial-gradient(rgba(246,227,182,.55) 2.5px,transparent 3px);background-size:26px 26px}}
.logo{{position:absolute;left:96px;top:60px;height:74px}}
.pill{{position:absolute;right:96px;top:72px;padding:10px 26px;border:2px solid rgba(255,255,255,.45);border-radius:999px;font-size:20px;font-weight:500}}
.tag{{position:absolute;left:0;top:170px;height:72px;padding:0 64px 0 96px;display:flex;align-items:center;background:{SAND};color:{NAVY};
  font-weight:700;font-size:26px;clip-path:polygon(0 0,calc(100% - 36px) 0,100% 50%,calc(100% - 36px) 100%,0 100%)}}
h1{{position:absolute;left:96px;top:270px;font-size:76px;font-weight:800;letter-spacing:-2.5px;line-height:1.04}}
.lead{{position:absolute;left:96px;font-size:27px;line-height:1.5;font-weight:400;opacity:.95}}
.card{{position:absolute;background:{CREAM};color:{INK};border-radius:30px;box-shadow:0 26px 60px rgba(6,28,48,.35);padding:40px 42px}}
.card.sand{{background:{SAND2}}}
.card.navy{{background:{NAVY};color:#fff;border:4px solid {SAND2}}}
.card h3{{font-size:36px;font-weight:800;letter-spacing:-1px;line-height:1.1}}
.card .sub{{font-size:21px;opacity:.75;margin-top:6px}}
.num{{font-size:22px;font-weight:700;opacity:.55}}
.row{{display:flex;justify-content:space-between;gap:20px;border-top:2px solid rgba(22,59,66,.12);padding:14px 0;font-size:21px}}
.navy .row{{border-color:rgba(255,255,255,.18)}}
.row b{{font-weight:700;text-align:right}}
.k{{letter-spacing:2.5px;font-size:16px;font-weight:600;opacity:.7;text-transform:uppercase}}
.price{{font-size:44px;font-weight:800;letter-spacing:-1px}}
.unit{{font-size:20px;opacity:.7}}
.foot{{position:absolute;left:96px;bottom:44px;font-size:18px;opacity:.75}}
.sticker{{position:absolute;border-radius:50%;background:{SAND2};color:{NAVY};display:flex;align-items:center;justify-content:center;
  text-align:center;font-weight:800;box-shadow:0 14px 30px rgba(0,0,0,.25);transform:rotate(-8deg)}}
.chip{{display:inline-block;padding:8px 20px;border-radius:999px;font-size:19px;font-weight:600}}
.ok{{background:#d9ead7;color:#24542c}}
.wait{{background:{SAND2};color:#7a5410}}
.no{{background:#f3d6cf;color:#8a2f1d}}
ul.b{{list-style:none}}
ul.b li{{position:relative;padding-left:34px;font-size:22px;line-height:1.45;margin-bottom:12px}}
ul.b li:before{{content:"";position:absolute;left:4px;top:12px;width:14px;height:14px;border-radius:50%;background:{CORAL}}}
.navy ul.b li:before{{background:{SAND2}}}
ul.b.sm li{{font-size:20px;margin-bottom:9px}}
.tbl{{width:100%;border-collapse:collapse;table-layout:fixed}}
.tbl th{{text-align:left;font-size:14px;letter-spacing:2px;text-transform:uppercase;opacity:.65;padding:8px 14px;border-bottom:2px solid rgba(22,59,66,.25)}}
.tbl td{{font-size:18.5px;padding:11px 14px;border-bottom:1px solid rgba(22,59,66,.12);vertical-align:top;line-height:1.38}}
.tbl td:first-child{{font-weight:700}}
.tbl tr:last-child td{{border-bottom:none}}
ul.b.sm li:before{{top:10px}}
.tbl.md td{{font-size:20px;padding:12px 16px}}
.tbl.lg td{{font-size:21px;padding:16px 16px}}
.tbl.lg th{{font-size:15px;padding:10px 16px}}
.photo{{position:absolute;border-radius:32px;background-size:cover;background-position:center 40%;box-shadow:0 26px 60px rgba(6,28,48,.35)}}
.step{{position:absolute;width:390px}}
.step .dot{{width:76px;height:76px;border-radius:50%;background:{SAND2};color:{NAVY};font-weight:800;font-size:30px;display:flex;align-items:center;justify-content:center;border:5px solid {NAVY}}}
.step h4{{font-size:28px;font-weight:700;margin:22px 0 8px}}
.step p{{font-size:21px;line-height:1.45;opacity:.9}}
.shot{{position:absolute;border-radius:26px;overflow:hidden;border:8px solid {CREAM};background:#fff;box-shadow:0 26px 60px rgba(6,28,48,.35)}}
.shot img{{display:block;width:100%}}
"""

SQ = ('<svg class="abs" style="left:{x}px;top:{y}px" width="{w}" height="60" viewBox="0 0 {w} 60">'
      '<path d="{d}" fill="none" stroke="#f6e3b6" stroke-width="7" stroke-linecap="round"/></svg>')


def squiggle(x, y, w=240):
    d, k = "M6 30", 0
    while (k + 1) * 40 < w - 6:
        d += f" q20 {-22 if k % 2 == 0 else 22} 40 0"
        k += 1
    return SQ.format(x=x, y=y, w=w, d=d)


def frame(body, tag=None, title=None, deco=True):
    head = f'<img class="logo" src="{LOGO}"><div class="pill">Pembaruan untuk BSN · November 2026</div>'
    d = ('<div class="blob" style="width:760px;height:620px;left:-280px;top:560px"></div>'
         '<div class="dots" style="width:260px;height:170px;right:70px;bottom:40px;opacity:.5"></div>') if deco else ""
    t = f'<div class="tag">{tag}</div>' if tag else ""
    h = f"<h1>{title}</h1>" if title else ""
    return f'<section class="pg">{d}{head}{t}{h}{body}</section>'


def rows(pairs):
    return "".join(f'<div class="row"><span>{a}</span><b>{b}</b></div>' for a, b in pairs)


def table(head, body_rows, widths, cls=""):
    cols = "".join(f'<col style="width:{w}%">' for w in widths)
    th = "".join(f"<th>{h}</th>" for h in head)
    tr = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in body_rows)
    return f'<table class="tbl {cls}"><colgroup>{cols}</colgroup><thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table>'


pages = []

# 1. Cover ---------------------------------------------------------------
pages.append(f"""
<section class="pg">
  <div class="blob" style="width:900px;height:760px;right:-300px;top:-260px"></div>
  <img class="logo" src="{LOGO}" style="height:90px">
  <div class="pill">November 2026</div>
  <div class="tag" style="top:200px">Pembaruan Proposal</div>
  <h1 style="top:300px;font-size:92px;letter-spacing:-3.5px">Pendampingan Pemetaan Talenta<br>dan Dukungan Karyawan</h1>
  {squiggle(100, 560, 300)}
  <div class="lead" style="top:640px;font-size:30px">Disiapkan untuk <b>Bank Syariah Nasional</b>, Human Capital</div>
  <div class="photo" style="left:1180px;top:640px;width:300px;height:360px;background-image:url({PHOTOS[1]})"></div>
  <div class="photo" style="left:1510px;top:560px;width:330px;height:440px;background-image:url({PHOTOS[0]})"></div>
  <div class="foot">Teman Journey x Coreitera · temanjourney.id · 0812 8184 5723</div>
</section>""")

FICT = '<div class="chip wait" style="position:absolute;right:96px;top:180px">Contoh dengan data fiktif</div>'

# 2. Ringkasan -------------------------------------------------------------
cards = [
    ("Bisa mulai sekarang", "Pendampingan pemetaan talenta", "Psikolog PIO kami menyusun peta talenta, talent box, dan rekomendasi untuk setiap peserta, apa pun pilihan tesnya.", "card navy", "ok"),
    ("Dua pilihan", "Tes asesmen", f"Tes dari kami {rp(OUR_PRICE)} per karyawan setelah perpanjangan lisensi, atau tes melalui mitra kami, Faxtor, {rp(FAX_PRICE)} per karyawan mulai sekarang.", "card", "wait"),
    ("Juga siap sekarang", "Dukungan kesehatan mental", "Pemeriksaan kesehatan mental, konseling, dan program bantuan karyawan yang dapat dipakai kapan pun dibutuhkan.", "card sand", "ok"),
]
body = "".join(
    f'<div class="{cls}" style="left:{96 + i * 590}px;top:580px;width:550px;height:400px">'
    f'<div class="chip {chip}">{n}</div><h3 style="margin:18px 0 16px">{h}</h3><p style="font-size:22px;line-height:1.5">{p}</p></div>'
    for i, (n, h, p, cls, chip) in enumerate(cards))
pages.append(frame(
    '<div class="lead" style="top:380px;width:1680px;font-size:25px">Terima kasih atas diskusi bersama tim Human Capital (HC) BSN. Ada satu perubahan penting dari proposal '
    '4 November 2026: kuota lisensi alat ukur kami untuk tahun ini sudah terpakai. Kami sedang memperpanjang lisensi, sehingga tes dari kami paling cepat tersedia awal Januari 2027. '
    'Karena itu, kami menyiapkan dua pilihan tes, dan layanan utama kami adalah pendampingan psikolog untuk pemetaan talenta BSN.</div>' + body,
    "Ringkasan Pembaruan", "Yang berubah dan yang kami tawarkan"))

# 3. Pendampingan (main highlight) ------------------------------------------
DELIV = [
    ("Mapping alignment.", "Kerangka pemetaan yang disepakati bersama HC dan berangkat dari job family BSN: tujuan, profil tiap job family, dan aspek yang diukur"),
    ("Rekomendasi dan group report.", "Rekomendasi psikolog untuk setiap peserta, termasuk kecocokan dengan job family, ditambah laporan kelompok"),
    ("Talent box.", "Peserta dipetakan dalam 9 kotak menurut potensi (hasil tes) dan kinerja (data BSN), lalu dikalibrasi bersama HC"),
    ("Talent analytics.", "Pola per job family, lengkap dengan rekomendasi tindak lanjut"),
]
pages.append(frame(
    f"""
    <div class="lead" style="top:360px;width:860px;font-size:23px">Psikolog PIO kami membaca hasil tes dari pilihan mana pun, lalu
    menyusun peta talenta dan rekomendasi bersama tim HC.</div>
    <div class="card" style="left:96px;top:500px;width:840px;height:520px;padding:32px 42px">
      <div class="k">Yang BSN terima</div>
      <ul class="b" style="margin-top:22px">
        {"".join(f"<li style='margin-bottom:18px'><b>{i + 1}. {a}</b> {b}</li>" for i, (a, b) in enumerate(DELIV))}
      </ul>
    </div>
    <div class="card navy" style="left:1010px;top:400px;width:820px;height:610px;padding:48px 56px">
      <div class="k">Investasi pendampingan</div>
      <div style="font-size:104px;font-weight:900;letter-spacing:-4px;line-height:1;margin:24px 0 10px">{rp(PEND_PRICE)}</div>
      <div style="font-size:30px;opacity:.85">untuk pemetaan hingga {PARTICIPANTS} karyawan</div>
      <div style="margin-top:34px">{rows([("Psikolog", "1 sampai 2 psikolog PIO"), ("Sesi bersama HC", "4 sesi online"),
                                          ("Pilihan tes", "Berlaku untuk A dan B"), ("Mulai", "Tidak menunggu lisensi")])}</div>
    </div>
    <div class="sticker" style="left:1700px;top:290px;width:190px;height:190px;font-size:28px;line-height:1.1">Bisa<br>mulai<br>sekarang</div>
    """,
    "1 · Layanan Utama", "Pendampingan Pemetaan Talenta"))

# 4-5. Contoh luaran pendampingan (fictional data, consistent across the two pages) ---------
# Talent box counts sum to 300; group report buckets are sums of talent box cells.
BOX = [  # potensi tinggi (top) to rendah (bottom); kinerja rendah, sedang, tinggi
    [("Berpotensi, belum tampil", 14, "#dbe7f0"), ("Berkembang pesat", 28, "#dbe7f0"), ("Bintang", 18, NAVY)],
    [("Perlu arahan", 22, SAND2), ("Andalan", 96, SAND2), ("Profesional kuat", 41, "#dbe7f0")],
    [("Perlu perhatian", 9, "#f3d6cf"), ("Pelaksana", 47, SAND2), ("Ahli teknis", 25, SAND2)],
]
assert sum(n for r in BOX for _, n, _ in r) == PARTICIPANTS
GROUP = [("Sangat sesuai", 120), ("Sesuai dengan pengembangan", 115), ("Lebih sesuai job family lain", 34), ("Perlu perhatian", 22 + 9)]
assert sum(n for _, n in GROUP) == PARTICIPANTS
UNITS = [("Operasional dan layanan cabang", 120, 48, 14, "Pelatihan pelayanan nasabah untuk 30 orang"),
         ("Pembiayaan dan risiko", 80, 32, 8, "Penguatan analisis risiko untuk 12 orang"),
         ("Pendanaan dan pemasaran", 60, 25, 6, "Coaching komunikasi untuk 10 orang"),
         ("Fungsi pendukung", 40, 15, 3, "Mentoring teknis di TI dan keuangan")]
assert sum(u[1] for u in UNITS) == PARTICIPANTS and sum(u[2] for u in UNITS) == GROUP[0][1] and sum(u[3] for u in UNITS) == GROUP[3][1]

align_rows = [("Kepribadian", "BFI", "Kecocokan dengan profil job family"),
              ("Gaya kerja", "DISC atau OPTI", "Cara kerja dalam tim"),
              ("Kepemimpinan", "Gaya kepemimpinan atau FTPI", "Kecenderungan memimpin"),
              ("Kinerja", "Data penilaian kinerja BSN", "Sumbu kinerja talent box")]
bars = "".join(
    f'<div style="display:flex;align-items:center;gap:16px;margin-top:12px;font-size:19px">'
    f'<div style="width:380px;white-space:nowrap">{lab}</div>'
    f'<div style="flex:1;height:20px;border-radius:10px;background:rgba(255,255,255,.16)"><div style="width:{n / max(n for _, n in GROUP) * 100:.0f}%;height:20px;border-radius:10px;background:{SAND2}"></div></div>'
    f'<div style="width:110px;text-align:right;font-weight:700">{n} orang</div></div>'
    for lab, n in GROUP)
pages.append(frame(
    '<div class="card" style="left:96px;top:290px;width:860px;height:740px;padding:36px 42px">'
    '<div class="k">Luaran 1 · Mapping alignment</div><h3 style="margin:14px 0 6px">Kerangka pemetaan BSN</h3>'
    '<div class="sub" style="margin-bottom:22px">Berangkat dari job family BSN, profil tiap job family disusun bersama HC</div>'
    + table(["Aspek", "Diukur dengan", "Dipakai untuk"], align_rows, [26, 38, 36], "lg") +
    '<div class="sub" style="margin-top:22px;font-size:19px">Disepakati pada sesi pertama bersama HC, sebelum tes dimulai.</div></div>'
    '<div class="card" style="left:1010px;top:290px;width:820px;height:430px;padding:30px 40px">'
    '<div class="k">Luaran 2 · Rekomendasi per peserta</div>'
    '<div style="font-size:30px;font-weight:800;margin-top:12px">Karyawan Contoh A</div><div class="sub" style="margin:0 0 10px">Job family operasional dan layanan cabang</div>'
    + rows([("Kekuatan", "Tenang di bawah tekanan, teliti"), ("Perlu dikembangkan", "Mendelegasikan tugas"),
            ("Kecocokan job family", "Sangat sesuai"), ("Rekomendasi", "Coaching delegasi tugas dalam 6 bulan")]) + '</div>'
    '<div class="card navy" style="left:1010px;top:745px;width:820px;height:285px;padding:24px 40px">'
    f'<div class="k">Luaran 2 · Group report</div><div style="font-size:28px;font-weight:800;margin-top:8px">Ringkasan {PARTICIPANTS} peserta</div>'
    + bars + '</div>' + FICT,
    "Contoh Luaran Pendampingan · 1 dan 2", None, deco=False))

AX = "font-size:16px;font-weight:600;opacity:.7;display:flex;align-items:center"
cells = ""
for r, lab in zip(BOX, ("Potensi tinggi", "Potensi sedang", "Potensi rendah")):
    cells += f'<div style="{AX};justify-content:flex-end;text-align:right;padding-right:6px">{lab}</div>'
    for name, n, col in r:
        fg = "#fff" if col == NAVY else INK
        cells += (f'<div style="background:{col};color:{fg};border-radius:16px;padding:14px 16px">'
                  f'<div style="font-size:17px;font-weight:600;line-height:1.25">{name}</div>'
                  f'<div style="font-size:38px;font-weight:800;margin-top:6px">{n}</div></div>')
cells += "<div></div>" + "".join(f'<div style="{AX};justify-content:center">Kinerja {k}</div>' for k in ("rendah", "sedang", "tinggi"))
unit_rows = [(u, str(n), str(b), str(c), t) for u, n, b, c, t in UNITS]
pages.append(frame(
    '<div class="card" style="left:96px;top:290px;width:800px;height:740px;padding:36px 40px">'
    '<div class="k">Luaran 3 · Talent box</div><h3 style="margin:14px 0 22px">Talent box seluruh peserta</h3>'
    '<div style="display:grid;grid-template-columns:120px 1fr 1fr 1fr;grid-template-rows:150px 150px 150px 40px;gap:12px">'
    + cells + '</div></div>'
    '<div class="card" style="left:940px;top:290px;width:890px;height:740px;padding:36px 40px">'
    '<div class="k">Luaran 4 · Talent analytics per job family</div><h3 style="margin:14px 0 22px">Pola dan tindak lanjut per job family</h3>'
    + table(["Job family", "Peserta", "Sangat sesuai", "Perlu perhatian", "Tindak lanjut"], unit_rows, [25, 12, 15, 15, 33]) +
    '<div class="sub" style="margin-top:20px;font-size:19px">Nama job family di atas hanya contoh. Pemetaan memakai job family yang dimiliki BSN, disepakati saat mapping alignment.</div></div>'
    + FICT,
    "Contoh Luaran Pendampingan · 3 dan 4", None, deco=False))

# 6. Pilihan tes: ours vs partner -------------------------------------------------------
TA, TB = OUR_PRICE * PARTICIPANTS, FAX_PRICE * PARTICIPANTS
opt_rows = [
    ("Alat tes", "DISC, Big Five Inventory (BFI), gaya kepemimpinan", "Big Five Inventory (BFI), FTPI, OPTI"),
    ("Bentuk hasil", "Laporan per peserta dalam bahasa yang mudah dipahami", "Laporan PDF per peserta dan rekap skor seluruh peserta"),
    ("Pembacaan hasil", "Psikolog PIO kami, lalu masuk ke pendampingan", "Psikolog PIO kami, lalu masuk ke pendampingan"),
    ("Mulai", "Setelah perpanjangan lisensi, paling cepat awal Januari 2027", "Bisa mulai sekarang"),
    ("Harga tes", f"{rp(OUR_PRICE)} per karyawan ({OUR_TESTS} tes x {rp(OUR_TEST_PRICE)})", f"{rp(FAX_PRICE)} per karyawan, termasuk rekap skor"),
    (f"Tes untuk {PARTICIPANTS} karyawan", rp(TA), rp(TB)),
    ("Pendampingan", rp(PEND_PRICE), rp(PEND_PRICE)),
    (f"<b>Total untuk {PARTICIPANTS} karyawan</b>", f"<b>{rp(TA + PEND_PRICE)}</b>", f"<b>{rp(TB + PEND_PRICE)}</b>"),
]
pages.append(frame(
    '<div class="lead" style="top:360px;width:1700px;font-size:24px">Pendampingan dan luarannya sama untuk kedua pilihan. Yang berbeda adalah alat tes, waktu mulai, dan harga tes.</div>'
    '<div class="card" style="left:96px;top:440px;width:1734px;height:590px;padding:28px 40px">'
    + table(["", "Pilihan A · Tes dari kami", "Pilihan B · Tes melalui mitra kami, Faxtor"], opt_rows, [22, 39, 39], "md") + '</div>',
    "2 · Pilihan Tes Asesmen", "Dua pilihan tes, satu pendampingan"))

# 7. Alat tes per aspek, side by side -----------------------------------------------------
tool_rows = [
    ("Kepribadian", "<b>Big Five Inventory (BFI)</b>. Lima dimensi kepribadian",
     "<b>Big Five Inventory (BFI)</b>. Lima dimensi kepribadian, dilengkapi skala kejujuran menjawab. Sekitar 40 menit",
     "Alat ukur yang sama"),
    ("Gaya kerja", "<b>DISC</b>. Gaya kerja dan komunikasi dalam empat tipe",
     "<b>OPTI</b>. Kepribadian di tempat kerja: growth mindset, ketekunan (grit), kemampuan interpersonal, dan kecerdasan emosional. Sekitar 50 menit",
     "OPTI tidak memakai empat tipe DISC. Bagian interpersonal dan keberanian sosial paling dekat dengan gaya komunikasi"),
    ("Kepemimpinan", "<b>Gaya kepemimpinan</b>. Profil gaya memimpin",
     "<b>FTPI</b>. 4 kelompok sifat dan 13 aspek, termasuk aspek kepemimpinan. Sekitar 20 menit",
     "FTPI mengukur kecenderungan memimpin, bukan gaya kepemimpinan"),
    ("Harga", f"{rp(OUR_TEST_PRICE)} per tes", f"{rp(FAX_PRICE)} untuk ketiga tes, termasuk rekap skor", ""),
]
pages.append(frame(
    '<div class="card" style="left:96px;top:290px;width:1734px;height:600px;padding:30px 40px">'
    + table(["Aspek", "Pilihan A · dari kami", "Pilihan B · melalui Faxtor", "Catatan"], tool_rows, [14, 25, 34, 27], "lg") + '</div>',
    "Perbandingan Alat Tes per Aspek", None, deco=False))

# 8. Contoh laporan A vs B ------------------------------------------------------------------
LBL = "position:absolute;top:262px;font-size:22px;font-weight:700;color:#f6e3b6"
pages.append(frame(
    f'<div style="{LBL};left:96px">Pilihan A · laporan dari kami</div>'
    f'<div class="shot" style="left:96px;top:310px;width:520px;height:720px"><img src="{SAMPLE["t4d_a"]}"></div>'
    f'<div style="{LBL};left:680px">Pilihan B · laporan FTPI dari Faxtor</div>'
    f'<div class="shot" style="left:680px;top:310px;width:540px;height:720px;background:#fff"><img src="{SAMPLE["ftpi"]}" style="object-fit:contain"></div>'
    '<div class="card navy" style="left:1290px;top:310px;width:540px;height:720px;padding:40px 44px">'
    '<h3 style="font-size:32px">Untuk kedua pilihan</h3><ul class="b" style="margin-top:24px">'
    '<li>Hasil tes setiap peserta dibaca psikolog PIO kami</li>'
    '<li>Rekomendasi per peserta disusun dalam format yang sama</li>'
    '<li>Hasil masuk ke talent box dan talent analytics</li>'
    '<li>Contohnya ada di halaman Contoh Luaran Pendampingan</li></ul></div>'
    + '<div class="chip wait" style="position:absolute;right:96px;top:180px">A data fiktif, B contoh Faxtor yang disamarkan</div>',
    "Contoh Laporan Tes", None, deco=False))

# 9. Fit and Proper Test as an optional extra for key positions (MASTER deck visual) ----
FP_LAYERS = ["Rubrik kompetensi disusun bersama HC",
             "Diskusi kelompok (LGD) dengan studi kasus",
             "Wawancara terstruktur oleh psikolog",
             "Laporan dan rekomendasi per kandidat"]
pages.append(frame(
    '<div class="lead" style="top:380px;width:960px;font-size:23px">Bila BSN membutuhkan asesmen yang lebih dalam untuk posisi kunci, kami menyediakan Fit and Proper Test, layanan yang sama dengan yang kami siapkan untuk Bank Indonesia.</div>'
    f'<div class="shot" style="left:96px;top:500px;width:800px;height:469px"><img src="{SAMPLE["fp_collage"]}"></div>'
    '<div class="card navy" style="left:1110px;top:380px;width:720px;height:590px;padding:40px 44px">'
    '<div class="k">Isi layanan</div><ul class="b" style="margin-top:20px">'
    + "".join(f"<li>{x}</li>" for x in FP_LAYERS) +
    '<li>Memakai hasil tes pemetaan, tanpa tes ulang</li></ul>'
    f'<div style="position:absolute;left:44px;bottom:40px"><div class="price" style="font-size:60px;color:#f6e3b6">{rp(FP_PRICE)}</div>'
    f'<div class="unit" style="color:#f6e3b6">per {FP_PAX} kandidat</div></div></div>'
    '<div class="foot">Gambar: Fit and Proper Test daring bersama Bank Indonesia. Jumlah kandidat ditentukan bersama HC BSN.</div>',
    "Layanan Tambahan", "Fit and Proper Test"))

# 10. Layanan pilot berbayar -------------------------------------------------
pilots = [
    ("Snapshot burnout", f"untuk satu unit kerja, sampai {SNAP_PAX} karyawan",
     "Karyawan mengisi kuesioner kelelahan kerja secara online. HC menerima laporan agregat dan paparan online 1 jam dari psikolog, tanpa membuka data individu.",
     rp(SNAP_PRICE), "per unit kerja", "card"),
    ("MHCU Flash", f"untuk {FLASH_PAX} orang tim HC",
     "Pemeriksaan kesehatan mental lengkap dengan konseling 30 menit, agar tim HC merasakan sendiri layanan yang nantinya diterima karyawan.",
     rp(FLASH_PRICE * FLASH_PAX), f"{rp(FLASH_PRICE)} per orang", "card sand"),
]
body = "".join(
    f'<div class="{cls}" style="left:{96 + i * 590}px;top:470px;width:550px;height:430px">'
    f'<h3>{h}</h3><div class="sub">{s_}</div>'
    f'<p style="font-size:21px;line-height:1.5;margin-top:20px;height:160px">{p}</p>'
    f'<div style="font-size:44px;font-weight:800;color:#104a70;margin-top:14px">{price}</div>'
    f'<div class="sub" style="margin-top:2px">{unit}</div></div>'
    for i, (h, s_, p, price, unit, cls) in enumerate(pilots))
body += (f'<div class="card navy" style="left:1276px;top:470px;width:550px;height:430px">'
         f'<h3>Mengapa pilot</h3>'
         f'<p style="font-size:22px;line-height:1.5;margin-top:20px">HC dapat memilih salah satu atau keduanya. Hasil pilot menjadi bahan HC untuk menilai kualitas tim kami sebelum program penuh dimulai.</p>'
         f'<div style="font-size:44px;font-weight:800;color:#f6e3b6;margin-top:26px">{rp(SNAP_PRICE + FLASH_PRICE * FLASH_PAX)}</div>'
         f'<div class="sub" style="color:#f6e3b6;margin-top:2px">bila keduanya diambil</div></div>')
pages.append(frame(
    '<div class="lead" style="top:370px;width:1500px;font-size:25px">Sebelum memutuskan, BSN dapat mencoba layanan kami dalam skala kecil.</div>'
    + body + '<div class="foot">MHCU Flash memakai harga katalog. Pilot dijadwalkan bersama HC BSN.</div>',
    "Layanan Pilot", "Coba dalam skala kecil"))


# 11. Contoh laporan MHCU (fictional fixture data from the Customer FE sample pages) ---
def sample_page(tag, a, b, a_w, right_x, heading, points):
    body = (f'<div class="shot" style="left:96px;top:290px;width:{a_w}px"><img src="{SAMPLE[a]}"></div>'
            f'<div class="shot" style="left:{right_x}px;top:290px;width:{1830 - right_x}px"><img src="{SAMPLE[b]}"></div>'
            f'<div class="card navy" style="left:{right_x}px;top:750px;width:{1830 - right_x}px;height:250px;padding:32px 40px">'
            f'<h3 style="font-size:30px">{heading}</h3><ul class="b" style="margin-top:16px">'
            + "".join(f"<li style='font-size:20px;margin-bottom:6px'>{p}</li>" for p in points) + "</ul></div>" + FICT)
    return frame(body, tag, None, deco=False)


pages.append(frame(
    f'<div class="shot" style="left:96px;top:290px;width:1060px;height:596px"><img src="{SAMPLE["mh_wellbeing"]}"></div>'
    f'<div class="shot" style="left:1190px;top:290px;width:640px;height:640px;background:#fff"><img src="{SAMPLE["mh_tahapan"]}"></div>'
    '<div style="position:absolute;left:96px;top:910px;width:1060px;font-size:22px;line-height:1.5;color:#fff">'
    'Sesi wellbeing memberi karyawan tempat aman untuk bercerita, di luar lingkungan internal. '
    'Hasil individu hanya untuk karyawan dan psikolog, sedangkan HC menerima laporan agregat.</div>',
    "Layanan Kesehatan Mental", None, deco=False))

# 7. MHCU (separate from the mapping scope) -------------------------------------------------------
tiers = [
    ("Skrining", "01", "Gambaran cepat untuk banyak orang", [("Instrumen", "DASS-21 online"), ("Ditinjau psikolog", "Ya"), ("Laporan", "Individu, 48 jam"), ("Konseling", "Tidak termasuk")], 175_000, "card"),
    ("Flash", "02", "Pemeriksaan lengkap dan konseling singkat", [("Instrumen", "5 instrumen klinis"), ("Ditinjau psikolog", "Ya"), ("Laporan", "Personal"), ("Konseling", "30 menit")], 600_000, "card sand"),
    ("Deep", "03", "Pemeriksaan lengkap dan konseling 1 jam", [("Instrumen", "5 instrumen klinis"), ("Ditinjau psikolog", "Ya"), ("Laporan", "Personal"), ("Konseling", "60 menit")], 1_000_000, "card navy"),
]
body = "".join(
    f'<div class="{cls}" style="left:{96 + i * 590}px;top:{460 - (30 if i == 1 else 0)}px;width:550px;height:{530 + (30 if i == 1 else 0)}px">'
    f'<h3 style="font-size:52px">{h} <span class="num">{n}</span></h3><div class="sub" style="margin-bottom:18px">{s}</div>{rows(r)}'
    f'<div style="position:absolute;left:42px;bottom:34px"><div class="price">{rp(v)}</div><div class="unit">per orang</div></div></div>'
    for i, (h, n, s, r, v, cls) in enumerate(tiers))
pages.append(frame(
    '<div class="lead" style="top:350px;width:1520px;font-size:24px">Pemeriksaan kesehatan mental (Mental Health Check-Up, MHCU) membantu HC melihat kondisi karyawan '
    'dan siapa yang perlu didukung. Layanan ini terpisah dari pemetaan dan dapat dipakai sekarang.</div>'
    + body + '<div class="sticker" style="left:1650px;top:200px;width:200px;height:200px;font-size:30px;line-height:1.1">Siap<br>sekarang</div>',
    "3 · Layanan Kesehatan Mental", "Pemeriksaan Kesehatan Mental"))

# 8. Alur + EAP --------------------------------------------------------------
steps = [("1", "Skrining", "Karyawan mengikuti skrining online."),
         ("2", "Tindak lanjut", "Karyawan dengan hasil yang perlu perhatian mendapat Flash atau Deep."),
         ("3", "Konseling", "Konseling lanjutan atau EAP sesuai kebutuhan karyawan."),
         ("4", "Laporan untuk HC", "HC menerima laporan agregat. Hasil individu hanya untuk karyawan dan psikolog.")]
body = "".join(
    f'<div class="step" style="left:{96 + i * 440}px;top:350px"><div class="dot">{n}</div><h4>{h}</h4><p>{p}</p></div>'
    for i, (n, h, p) in enumerate(steps))
body += '<div class="abs" style="left:172px;top:388px;width:1320px;border-top:4px dashed rgba(246,227,182,.6);z-index:0"></div>'
extra = [("Program bantuan karyawan (EAP)", "Employee Assistance Program untuk 300 karyawan: kuota 10 sesi konseling per bulan, MHCU 2 kali setahun, laporan agregat, dan 1 webinar per kuartal.", rp(13_500_000), "per bulan", "card navy"),
         ("Paket 10 sesi konseling", "Untuk karyawan tertentu di luar EAP: 10 sesi online 60 menit beserta 10 MHCU dan laporannya, berlaku 6 bulan.", rp(10_500_000), "per paket · " + rp(6_200_000) + " untuk sesi 30 menit", "card"),
         ("Workshop lingkungan kerja sehat", "2 jam tatap muka, 30 peserta, dipandu psikolog dan co-fasilitator.", rp(8_100_000), "per sesi", "card sand")]
body += "".join(
    f'<div class="{cls}" style="left:{96 + i * 590}px;top:640px;width:550px;height:370px;padding:34px 40px">'
    f'<h3 style="font-size:30px">{h}</h3><p style="font-size:20px;line-height:1.45;margin-top:12px">{p}</p>'
    f'<div style="position:absolute;left:40px;bottom:28px"><div class="price" style="font-size:38px">{v}</div><div class="unit">{u}</div></div></div>'
    for i, (h, p, v, u, cls) in enumerate(extra))
pages.append(frame(body, "Alur Dukungan Kesehatan Mental Karyawan", None))

# 9. Langkah berikutnya --------------------------------------------------------
nxt = [("Minggu ini", "Diskusi 30 menit dengan HC BSN untuk memilih pilihan tes dan layanan pilot."),
       ("Desember 2026", "Mapping alignment dimulai. Bila BSN memilih tes melalui Faxtor, tes dapat langsung berjalan."),
       ("Paling cepat awal Januari 2027", "Perpanjangan lisensi selesai, sehingga tes dari kami siap dipakai."),
       ("Setelah tes selesai", "Rekomendasi, group report, talent box, dan talent analytics disampaikan kepada HC.")]
body = "".join(
    f'<div class="step" style="left:{96 + i * 440}px;top:430px"><div class="dot">{i + 1}</div><h4>{h}</h4><p>{p}</p></div>'
    for i, (h, p) in enumerate(nxt))
body += '<div class="abs" style="left:172px;top:468px;width:1320px;border-top:4px dashed rgba(246,227,182,.6)"></div>'
body += ('<div class="card" style="left:96px;top:800px;width:1734px;height:180px;padding:30px 44px;display:flex;align-items:center;gap:60px">'
         '<div style="flex:1"><div class="k">Ketentuan</div><p style="font-size:19px;line-height:1.5;margin-top:8px">Harga berlaku untuk wilayah Jakarta dan tidak dikenai PPN, '
         'karena Teman Journey bukan Pengusaha Kena Pajak. Di luar Jakarta, perjalanan dan akomodasi ditagih sesuai biaya aktual. Rincian setiap layanan ada di Lampiran. Data karyawan dikelola sesuai Kode Etik Psikologi '
         'Indonesia (HIMPSI) dan UU No. 27 Tahun 2022 tentang Pelindungan Data Pribadi. Penawaran berlaku 30 hari sejak tanggal dokumen.</p></div>'
         '<div style="text-align:right;font-size:22px;line-height:1.6"><b>Teman Journey</b><br>temanjourney.id<br>0812 8184 5723</div></div>')
pages.append(frame(body, "Langkah Berikutnya", None))

# 10. Tim psikolog ------------------------------------------------------------
team = [("Riskia Ramadhina Sukriananda", "Psikolog"), ("Amanda Nurshadrina", "Psikolog"), ("Laras Sita Persitia", "Psikolog Klinis Dewasa"),
        ("Fasya Khalida Suffah", "Psikolog Klinis Dewasa"), ("Nadira Khairunnisa", "Psikolog Klinis Dewasa"), ("Ninette Putri Mustika", "Psikolog"),
        ("Rhajiv Nur Ilham", "Psikolog Klinis Dewasa")]
body = "".join(
    f'<div class="card{" sand" if i % 3 == 1 else ""}" style="left:{96 + (i % 4) * 440}px;top:{400 + (i // 4) * 220}px;width:400px;height:190px;padding:30px 34px">'
    f'<div style="font-size:25px;font-weight:700;line-height:1.25">{n}</div><div class="sub">{r}</div><div class="unit" style="margin-top:8px">S2 Universitas Indonesia</div></div>'
    for i, (n, r) in enumerate(team))
body += f'<div class="photo" style="left:1416px;top:620px;width:414px;height:390px;background-image:url({PHOTOS[0]})"></div>'
pages.append(frame('<div class="lead" style="top:340px;font-size:24px">Psikolog tersertifikasi HIMPSI. Psikolog yang bertugas dikonfirmasi kepada BSN sebelum pelaksanaan.</div>' + body,
                   "Tim Psikolog", None))

# Lampiran: rincian layanan -----------------------------------------------------
def appendix(tag, inner):
    return frame(inner, tag, None, deco=False)


pend_rows = [
    ("1. Mapping alignment", "Sesi online bersama HC. Menyepakati tujuan, job family BSN dan profilnya, aspek yang diukur, dan data kinerja yang dipakai", "Kerangka pemetaan yang disepakati"),
    ("Pengolahan hasil tes", "Psikolog membaca hasil tes setiap peserta dari pilihan A atau B", "Data hasil tes yang siap dipetakan"),
    ("2. Rekomendasi dan group report", f"1 sampai 2 psikolog PIO menelaah setiap peserta dengan total {PIO_REVIEW_HOURS} jam kerja, lalu merangkum per job family", "Rekomendasi untuk setiap peserta, termasuk kecocokan job family, dan 1 group report"),
    ("3. Talent box", "Potensi dari hasil tes digabung dengan data kinerja BSN, lalu dikalibrasi bersama HC dalam 1 sesi online", "Talent box 9 kotak untuk seluruh peserta"),
    ("4. Talent analytics", "Analisis pola per job family", "Rekomendasi tindak lanjut per job family"),
    ("Pemantauan dan finalisasi", "2 sesi online bersama HC untuk membahas temuan awal dan menyampaikan hasil akhir", "Peta talenta BSN per job family dan dasar program pengembangan"),
    ("Sepanjang program", "Koordinator proyek mengatur jadwal dan komunikasi", "Satu pintu komunikasi untuk HC"),
]
pages.append(appendix("Lampiran 1 · Rincian Pendampingan Pemetaan Talenta",
    '<div class="card" style="left:96px;top:280px;width:1734px;height:580px;padding:26px 36px">'
    + table(["Tahap", "Yang dilakukan", "Yang BSN terima"], pend_rows, [22, 50, 28]) + "</div>"
    '<div class="card sand" style="left:96px;top:880px;width:850px;height:160px;padding:22px 34px"><div class="k">Yang perlu disiapkan BSN</div>'
    '<p style="font-size:19px;line-height:1.5;margin-top:8px">Pilihan tes, daftar job family BSN, daftar peserta dengan unit dan jabatan, data penilaian kinerja terakhir untuk '
    'talent box, serta satu penanggung jawab dari tim HC.</p></div>'
    f'<div class="card navy" style="left:980px;top:880px;width:850px;height:160px;padding:22px 34px"><div class="k">Lingkup harga</div>'
    f'<p style="font-size:19px;line-height:1.5;margin-top:8px">{rp(PEND_PRICE)} untuk seluruh luaran di atas, hingga {PARTICIPANTS} karyawan, '
    'untuk pilihan tes A maupun B. Harga tes dihitung terpisah.</p></div>'))

cand_rows = [
    ("Fit and Proper Test", "Untuk posisi kunci bila dibutuhkan: rubrik kompetensi, LGD, wawancara terstruktur oleh psikolog, laporan dan rekomendasi per kandidat. Memakai hasil tes dari pemetaan", rp(FP_PRICE) + f" per {FP_PAX} kandidat"),
]
pkg_rows = [
    ("Alat tes", "DISC, BFI, gaya kepemimpinan", "BFI, FTPI, OPTI"),
    ("Pelaksanaan", "Tes online melalui platform kami", "Tes online melalui platform Faxtor, pendaftaran peserta melalui Teman Journey"),
    ("Hasil", "Laporan per peserta", "Laporan PDF per peserta dan rekap skor Excel seluruh peserta"),
    ("Mulai", "Setelah perpanjangan lisensi, paling cepat awal Januari 2027", "Bisa mulai sekarang"),
    ("Harga", f"{rp(OUR_PRICE)} per karyawan", f"{rp(FAX_PRICE)} per karyawan, ditagih oleh Teman Journey"),
]
pages.append(appendix("Lampiran 2 · Pilihan Tes dan Fit and Proper Test",
    '<div class="card" style="left:96px;top:280px;width:1734px;height:200px;padding:26px 36px"><div class="k" style="margin-bottom:8px">Fit and Proper Test (layanan tambahan)</div>'
    + table(["Layanan", "Isi", "Harga"], cand_rows, [24, 54, 22]) + "</div>"
    '<div class="card sand" style="left:96px;top:520px;width:1734px;height:400px;padding:26px 36px"><div class="k" style="margin-bottom:8px">Pilihan tes asesmen</div>'
    + table(["Komponen", "Pilihan A · dari kami", "Pilihan B · melalui Faxtor"], pkg_rows, [20, 36, 44]) + "</div>"))

sup_rows = [
    ("MHCU Skrining", "Kuesioner DASS-21 online tentang depresi, kecemasan, dan stres, ditinjau psikolog", "Laporan individu paling lambat 48 jam, tanpa konseling", rp(175_000) + " per orang"),
    ("MHCU Flash", "5 instrumen: DASS-21, kelelahan kerja (CBI), kepuasan hidup (SWLS), kepribadian singkat (BFI-10), dan cerita pribadi", "Laporan personal yang ditinjau psikolog, konseling 30 menit", rp(600_000) + " per orang"),
    ("MHCU Deep", "Instrumen sama dengan Flash", "Laporan personal yang ditinjau psikolog, konseling 60 menit", rp(1_000_000) + " per orang"),
    ("Program bantuan karyawan (EAP)", "Kuota 10 sesi konseling per bulan, MHCU 2 kali setahun untuk semua karyawan, 1 webinar per kuartal", "Laporan agregat untuk HC", rp(13_500_000) + " per bulan, 300 karyawan"),
    ("Paket 10 sesi konseling", "10 sesi online 60 menit dan 10 MHCU, berlaku 6 bulan", "10 laporan individu", rp(10_500_000) + " per paket, atau " + rp(6_200_000) + " untuk sesi 30 menit"),
    ("Workshop lingkungan kerja sehat", "2 jam tatap muka untuk 30 peserta, dipandu psikolog dan co-fasilitator", "Lembar kerja peserta", rp(8_100_000) + " per sesi"),
]
pages.append(appendix("Lampiran 3 · Rincian Dukungan Karyawan",
    '<div class="card" style="left:96px;top:280px;width:1734px;height:590px;padding:26px 36px">'
    + table(["Layanan", "Isi", "Yang diterima", "Harga"], sup_rows, [19, 40, 22, 19]) + "</div>"
    '<div class="card navy" style="left:96px;top:890px;width:1734px;height:140px;padding:24px 34px;display:flex;gap:40px;align-items:center">'
    '<div class="k" style="width:200px">Kerahasiaan</div><p style="font-size:19px;line-height:1.5;flex:1">Hasil individu MHCU dan konseling hanya dibaca karyawan '
    'yang bersangkutan dan psikolog. HC menerima laporan agregat. Semua psikolog yang meninjau hasil memiliki Surat Izin Praktik Psikologi (SIPP).</p></div>'))

html = f"<!doctype html><html lang='id'><head><meta charset='utf-8'><style>{CSS}</style></head><body>{''.join(pages)}</body></html>"
BANNED = "—–;→"
import re
text_only = re.sub(r"<[^>]*>", " ", html.split("</style>", 1)[1])  # visible text only, not inline styles
for ch in BANNED:
    assert ch not in text_only, f"banned character {ch!r} in client text"
src = HERE / "update_deck.html"
src.write_text(html, encoding="utf-8")
subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--virtual-time-budget=8000",
                f"--print-to-pdf={OUT_PDF}", src.as_uri()], check=True, capture_output=True, timeout=180)
print("pendampingan cost", rp(PEND_COST), "price", rp(PEND_PRICE), f"net {net_margin(PEND_COST, PEND_PRICE):.1%}",
      "| faxtor", rp(FAX_COST), "->", rp(FAX_PRICE), f"net {net_margin(FAX_COST, FAX_PRICE):.1%}",
      "| ours", rp(OUR_PRICE), "| pages", len(pages), "->", OUT_PDF.name)
