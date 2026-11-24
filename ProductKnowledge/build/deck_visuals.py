"""Render full-slide visuals (1920x1080 PNG) for the Teman Journey client deck.

Each slide is HTML in the deck's navy-gradient + sand look, rendered by headless Chrome.
The logo (top-left) and date pill (top-right) are left empty on purpose: they stay as
editable Canva elements on top of the PNG, together with any [placeholder] text.

Run:  /opt/homebrew/bin/python3 -I deck_visuals.py
Out:  ProductKnowledge/deck_visuals/*.png
"""
import os
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "deck_visuals")
SRC = os.path.join(OUT, "html")
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
os.makedirs(SRC, exist_ok=True)

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
html,body{width:1920px;height:1080px;overflow:hidden}
body{font-family:'Poppins',sans-serif;color:#fff;position:relative;
 background:linear-gradient(90deg,#104a70 0%,#4f5c74 52%,#628cab 100%)}
.ring{position:absolute;border-radius:50%;border:2px solid rgba(255,255,255,.08)}
.blob{position:absolute;border-radius:50%;background:radial-gradient(circle,rgba(255,255,255,.10),rgba(255,255,255,0) 70%)}
.sand{color:#f6e3b6}
.kicker{font-size:30px;font-weight:500;letter-spacing:.5px;color:#f6e3b6}
h1{font-size:96px;font-weight:700;line-height:1.02;letter-spacing:-2.5px}
.glass{background:rgba(255,255,255,.08);border:2px solid rgba(255,255,255,.22);border-radius:36px;backdrop-filter:blur(4px)}
.card{background:#fdfaf1;color:#163b42;border-radius:32px;box-shadow:0 24px 60px rgba(8,30,50,.28)}
.pill{display:inline-block;padding:8px 22px;border-radius:999px;font-weight:600;font-size:22px}
svg.i{flex-shrink:0;width:64px;height:64px;stroke:currentColor;fill:none;stroke-width:2.2;stroke-linecap:round;stroke-linejoin:round}
"""

ICON = {
    "search": '<circle cx="28" cy="28" r="16"/><path d="M40 40l14 14"/><path d="M21 28h14M28 21v14"/>',
    "heart": '<path d="M32 54S10 40 10 24a11 11 0 0 1 22-3 11 11 0 0 1 22 3c0 16-22 30-22 30z"/>',
    "grow": '<path d="M32 56V30"/><path d="M32 34c0-10 8-16 18-16 0 10-8 16-18 16z"/><path d="M32 40c0-8-6-13-14-13 0 8 6 13 14 13z"/><path d="M18 56h28"/>',
    "chat": '<path d="M12 14h40v28H30l-10 10V42h-8z"/><path d="M22 26h20M22 33h12"/>',
    "book": '<path d="M10 14h18a6 6 0 0 1 6 6v32a5 5 0 0 0-5-5H10z"/><path d="M54 14H36a6 6 0 0 0-6 6v32a5 5 0 0 1 5-5h19z"/>',
    "users": '<circle cx="24" cy="22" r="8"/><path d="M10 50c0-8 6-14 14-14s14 6 14 14"/><circle cx="44" cy="24" r="6"/><path d="M42 36c7 0 12 5 12 12"/>',
    "org": '<rect x="24" y="8" width="16" height="12" rx="3"/><rect x="8" y="44" width="16" height="12" rx="3"/><rect x="40" y="44" width="16" height="12" rx="3"/><path d="M32 20v12M16 44v-6h32v6"/>',
    "tired": '<circle cx="32" cy="32" r="22"/><path d="M22 28h6M36 28h6"/><path d="M24 42c5-3 11-3 16 0"/>',
    "focus": '<circle cx="32" cy="32" r="22"/><circle cx="32" cy="32" r="12"/><circle cx="32" cy="32" r="3"/>',
    "door": '<path d="M16 56V10h26v46"/><path d="M42 18l10 4v34"/><circle cx="35" cy="34" r="2"/><path d="M8 56h48"/>',
    "check": '<path d="M14 33l12 12 24-26"/>',
    "dash": '<path d="M18 32h28"/>',
    "chart": '<path d="M10 54h44"/><rect x="14" y="34" width="8" height="16"/><rect x="28" y="24" width="8" height="26"/><rect x="42" y="14" width="8" height="36"/>',
    "shield": '<path d="M32 8l20 8v14c0 13-9 22-20 26-11-4-20-13-20-26V16z"/><path d="M24 32l6 6 11-12"/>',
}


def icon(name, size=64, color="currentColor"):
    return f'<svg class="i" viewBox="0 0 64 64" style="width:{size}px;height:{size}px;color:{color}">{ICON[name]}</svg>'


def deco():
    return ('<div class="ring" style="width:1300px;height:1300px;left:-420px;top:-260px"></div>'
            '<div class="ring" style="width:900px;height:900px;left:-180px;top:120px"></div>'
            '<div class="blob" style="width:900px;height:900px;right:-300px;bottom:-420px"></div>')


def page(body):
    return f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head><body>{deco()}{body}</body></html>'


SLIDES = {}

# ---------- Section dividers ----------
DIV = [
    ("01", "Asesmen", "Ukur dulu, baru bertindak. Data yang tepat membuat setiap program berikutnya lebih terarah.",
     ["Mental Health Check-Up: Skrining, Flash, Deep", "Alur, jadwal, dan privasi data MHCU", "Seleksi dan pemetaan talenta",
      "Asesmen individu, kelompok, dan karier", "Skrining tumbuh kembang anak"], "search"),
    ("02", "Konseling &<br>Pendampingan", "Ruang aman untuk bercerita, kapan pun dibutuhkan. Ditangani psikolog berizin, data pribadi tetap rahasia.",
     ["Kapan pegawai perlu sesi wellbeing", "Paket EAP Jeda, Tumbuh, Seperjalanan", "EAP Korporat",
      "In-house Psychologist", "Konseling individu dan pasangan", "Program pendampingan kelompok"], "heart"),
    ("03", "Pembelajaran &<br>Psikoedukasi", "Membangun kapasitas lewat pelatihan, workshop, dan webinar yang interaktif dan sesuai konteks peserta.",
     ["Pelatihan dan seminar", "Healthy Working Environment", "Topik webinar",
      "Workshop sekolah: karier, gadget, edukasi seksual, anti-bullying, parenting", "Team building"], "book"),
    ("04", "Pengalaman &<br>Komunitas", "Ruang jeda yang hangat lewat seni, keluarga, dan komunitas.",
     ["Art-Based Wellbeing", "Teman Kecil untuk keluarga dan anak", "Event komunitas"], "users"),
    ("05", "Konsultasi<br>Organisasi", "Pendampingan sistem SDM dan evaluasi program, dari diagnosis sampai rekomendasi.",
     ["HR Advisory", "Linimasa kerja", "Hasil yang diterima"], "org"),
]
for num, title, lead, items, ic in DIV:
    rows = "".join(
        f'<div style="display:flex;gap:22px;align-items:flex-start;padding:18px 0;border-bottom:1px solid rgba(255,255,255,.14)">'
        f'<div style="min-width:44px;height:44px;border-radius:50%;background:#f6e3b6;color:#104a70;font-weight:700;font-size:20px;display:flex;align-items:center;justify-content:center">{k+1}</div>'
        f'<div style="font-size:30px;font-weight:500;line-height:1.35;padding-top:2px">{t}</div></div>'
        for k, t in enumerate(items))
    SLIDES[f"divider_{num}"] = page(f"""
<div style="position:absolute;left:108px;top:230px;width:820px">
  <div style="font-size:300px;font-weight:800;line-height:.9;color:transparent;-webkit-text-stroke:3px rgba(246,227,182,.85);letter-spacing:-8px">{num}</div>
  <div style="margin-top:28px;font-size:100px;font-weight:700;line-height:1.02;letter-spacing:-2.5px">{title}</div>
  <div style="margin-top:34px;width:120px;height:8px;border-radius:8px;background:#f6e3b6"></div>
</div>
<div class="glass" style="position:absolute;left:1000px;top:200px;width:812px;padding:54px 56px">
  <div style="display:flex;gap:22px;align-items:center">{icon(ic,72,'#f6e3b6')}
    <div style="font-size:32px;line-height:1.45;font-weight:400">{lead}</div></div>
  <div class="kicker" style="margin:40px 0 6px;font-size:24px;text-transform:uppercase;letter-spacing:3px">Isi modul ini</div>
  {rows}
</div>""")

# ---------- p7 service map ----------
LINES = [
    ("01", "Asesmen", "search", "MHCU, pemetaan talenta, tes IQ, minat bakat, tumbuh kembang"),
    ("02", "Konseling &amp; pendampingan", "heart", "EAP, in-house psychologist, konseling individu dan pasangan"),
    ("03", "Pembelajaran &amp; psikoedukasi", "book", "Pelatihan, workshop, webinar, team building"),
    ("04", "Pengalaman &amp; komunitas", "users", "Art-based wellbeing, Teman Kecil, event komunitas"),
    ("05", "Konsultasi organisasi", "org", "HR advisory, evaluasi program, riset kebutuhan"),
]
tiles = "".join(f"""
<div class="card" style="width:322px;height:470px;padding:40px 32px;display:flex;flex-direction:column">
  <div style="display:flex;justify-content:space-between;align-items:flex-start">
    <div style="width:96px;height:96px;border-radius:28px;background:linear-gradient(135deg,#104a70,#628cab);display:flex;align-items:center;justify-content:center">{icon(ic,56,'#f6e3b6')}</div>
    <div style="font-size:54px;font-weight:800;color:#e8d4ab;line-height:1">{n}</div></div>
  <div style="margin-top:34px;font-size:31px;font-weight:700;line-height:1.2;color:#104a70;letter-spacing:-.3px">{t}</div>
  <div style="margin-top:18px;font-size:23px;line-height:1.5;color:#4f5c74">{d}</div>
</div>""" for n, t, ic, d in LINES)
SLIDES["p07_peta_layanan"] = page(f"""
<div style="position:absolute;left:108px;top:200px">
  <div class="kicker">Layanan Teman Journey</div>
  <h1 style="margin-top:10px;font-size:84px">Satu mitra, lima lini layanan</h1>
  <div style="margin-top:18px;font-size:30px;font-weight:300;opacity:.92">Mengukur, mendampingi, dan membangun kapasitas, dalam satu standar psikolog berizin.</div>
</div>
<div style="position:absolute;left:108px;top:500px;display:flex;gap:24px">{tiles}</div>""")

# ---------- p5 Kenali / Dampingi / Kembangkan ----------
PILLARS = [
    ("Kenali", "search", "Asesmen untuk melihat kondisi dan potensi secara utuh: kesehatan mental (MHCU), seleksi, dan pemetaan talenta.", "Manajemen menerima data agregat"),
    ("Dampingi", "heart", "Konseling bersama psikolog berizin, daring, tatap muka, atau di lokasi. Ruang aman dan rahasia.", "Di luar lingkungan kantor dan HR"),
    ("Kembangkan", "grow", "Pelatihan, workshop, dan konsultasi berbasis psikologi terapan.", "Diukur sebelum dan sesudah"),
]
cols = ""
for k, (t, ic, d, tag) in enumerate(PILLARS):
    cols += f"""
<div class="card" style="width:500px;height:540px;padding:56px 44px 44px;position:relative;display:flex;flex-direction:column">
  <div style="position:absolute;top:-34px;left:44px;width:68px;height:68px;border-radius:50%;background:#f6e3b6;color:#104a70;font-weight:800;font-size:30px;display:flex;align-items:center;justify-content:center;box-shadow:0 8px 20px rgba(0,0,0,.18)">{k+1}</div>
  <div style="display:flex;align-items:center;gap:20px;color:#104a70">{icon(ic,70)}<div style="font-size:46px;font-weight:700;letter-spacing:-1px">{t}</div></div>
  <div style="margin-top:26px;font-size:27px;line-height:1.5;color:#4f5c74;flex:1">{d}</div>
  <div><span style="display:inline-block;background:#e8d4ab;color:#104a70;padding:12px 24px;border-radius:999px;font-weight:600;font-size:20px;white-space:nowrap">{tag}</span></div>
</div>"""
    if k < 2:
        cols += '<div style="align-self:center;font-size:64px;color:#f6e3b6;font-weight:300">&rarr;</div>'
SLIDES["p05_pendekatan"] = page(f"""
<div style="position:absolute;left:108px;top:190px">
  <div class="kicker">Pendekatan kami</div>
  <h1 style="margin-top:8px;font-size:66px">Satu mitra, dari memahami sampai mendampingi</h1>
</div>
<div style="position:absolute;left:108px;top:430px;display:flex;gap:30px">{cols}</div>""")

# ---------- p6 Perjalanan bersama kami ----------
STEPS = [
    ("Langkah 1", "Gambaran awal", "Snapshot singkat kondisi tim, dengan laporan agregat untuk HR. Tanpa biaya.", "chart"),
    ("Langkah 2", "Program inti", "Asesmen, pelatihan, atau workshop yang dirancang untuk organisasi Anda, dengan ukuran keberhasilan yang disepakati di awal.", "grow"),
    ("Langkah 3", "Jaga hasilnya", "Pendampingan rutin: EAP, MHCU berkala, atau psikolog di lokasi, supaya perubahan bertahan.", "shield"),
]
xs = [108, 690, 1272]
steps = ""
for (lbl, t, d, ic), x in zip(STEPS, xs):
    steps += f"""
<div style="position:absolute;left:{x}px;top:470px;width:540px">
  <div style="width:150px;height:150px;border-radius:50%;background:#f6e3b6;display:flex;align-items:center;justify-content:center;box-shadow:0 0 0 14px rgba(246,227,182,.18)">{icon(ic,76,'#104a70')}</div>
  <div class="kicker" style="margin-top:42px;font-size:24px;text-transform:uppercase;letter-spacing:3px">{lbl}</div>
  <div style="margin-top:6px;font-size:52px;font-weight:700;letter-spacing:-1px">{t}</div>
  <div style="margin-top:14px;font-size:27px;line-height:1.5;font-weight:300;width:500px">{d}</div>
</div>"""
SLIDES["p06_perjalanan"] = page(f"""
<div style="position:absolute;left:108px;top:200px">
  <div class="kicker">Cara kita bekerja sama</div>
  <h1 style="margin-top:8px">Perjalanan bersama kami</h1>
</div>
<div style="position:absolute;left:258px;top:543px;width:1170px;border-top:4px dashed rgba(246,227,182,.6)"></div>
{steps}""")

# ---------- p10 MHCU tiers ----------
ROWS = ["Instrumen", "Review psikolog SIPP", "Laporan", "Konseling"]
TIERS = [
    ("Skrining", "Gambaran cepat untuk banyak orang", ["DASS-21 online", "check", "Individual, 48 jam", "dash"], False),
    ("Flash", "Skrining plus sesi singkat", ["5 instrumen klinis", "check", "Laporan personal", "30 menit"], False),
    ("Deep", "Pemeriksaan paling lengkap", ["5 instrumen klinis", "check", "Laporan personal", "60 menit"], True),
]
def cellv(v, dark):
    c = "#f6e3b6" if dark else "#104a70"
    if v == "check":
        return icon("check", 44, c)
    if v == "dash":
        return f'<span style="font-size:24px;opacity:.55">Tidak termasuk</span>'
    return f'<span style="font-size:25px;font-weight:600">{v}</span>'
tcards = ""
for name, sub, vals, dark in TIERS:
    bg = "linear-gradient(160deg,#0d3a5a,#104a70)" if dark else "#fdfaf1"
    fg = "#fff" if dark else "#163b42"
    rows = "".join(f'<div style="height:78px;display:flex;justify-content:space-between;align-items:center;border-top:1px solid {"rgba(255,255,255,.18)" if dark else "#eadfc6"}">'
                   f'<div style="font-size:18px;letter-spacing:1.5px;text-transform:uppercase;opacity:.7;width:190px;line-height:1.3">{r}</div>'
                   f'<div style="text-align:right;display:flex;justify-content:flex-end">{cellv(v, dark)}</div></div>'
                   for r, v in zip(ROWS, vals))
    tcards += f"""
<div style="width:536px;padding:38px 40px 24px;border-radius:32px;background:{bg};color:{fg};box-shadow:0 24px 60px rgba(8,30,50,.3);{'border:3px solid #f6e3b6' if dark else ''}">
  <div style="font-size:56px;font-weight:700;letter-spacing:-1.5px;line-height:1.1;color:{'#f6e3b6' if dark else '#104a70'}">{name}</div>
  <div style="font-size:23px;margin:4px 0 18px;opacity:.8">{sub}</div>{rows}
</div>"""
SLIDES["p10_mhcu"] = page(f"""
<div style="position:absolute;left:108px;top:180px">
  <div class="kicker">Asesmen</div>
  <h1 style="margin-top:6px;font-size:84px">Mental Health Check-Up</h1>
  <div style="margin-top:12px;font-size:30px;font-weight:300">Tiga pilihan, satu standar: instrumen tervalidasi dan laporan yang ditinjau psikolog berizin.</div>
</div>
<div style="position:absolute;left:108px;top:430px;display:flex;gap:48px;align-items:flex-start">{tcards}</div>""")

# ---------- p3 Yang sering kami dengar (editable paragraph stays in Canva, box reserved bottom-right) ----------
SIG = [("tired", "Kelelahan"), ("focus", "Sulit fokus"), ("door", "Diam-diam bersiap pindah")]
sig = "".join(f'<div class="glass" style="flex:1;padding:34px 30px;display:flex;gap:22px;align-items:center">{icon(ic,64,"#f6e3b6")}'
              f'<div style="font-size:29px;font-weight:600;line-height:1.2;white-space:nowrap">{t}</div></div>' for ic, t in SIG)
SLIDES["p03_masalah"] = page(f"""
<div style="position:absolute;left:108px;top:190px;width:1704px">
  <div class="kicker">Yang sering kami dengar</div>
  <h1 style="margin-top:8px;font-size:80px">Masalahnya sering baru terlihat<br>setelah kinerja turun</h1>
  <div style="margin-top:44px;display:flex;gap:26px">{sig}</div>
</div>
<div style="position:absolute;left:108px;top:720px;width:780px;font-size:25px;line-height:1.55;font-weight:300">
  Tuntutan kerja tinggi tanpa sumber daya yang cukup menguras energi psikologis dan menaikkan risiko burnout
  <span class="sand" style="font-weight:500">(Bakker &amp; Demerouti, 2017)</span>. WHO pun sudah mengakui burnout sebagai fenomena okupasional
  <span class="sand" style="font-weight:500">(2019)</span>.
</div>
<div style="position:absolute;left:940px;top:700px;width:872px;height:250px;border-radius:28px;border:2px dashed rgba(246,227,182,.55)"></div>
""")

for name, html in SLIDES.items():
    hp = os.path.join(SRC, name + ".html")
    open(hp, "w").write(html)
    png = os.path.join(OUT, name + ".png")
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                    "--window-size=1920,1080", "--virtual-time-budget=4000", f"--screenshot={png}", "file://" + hp],
                   check=True, capture_output=True)
    print("rendered", os.path.relpath(png, os.path.join(HERE, "..")))
