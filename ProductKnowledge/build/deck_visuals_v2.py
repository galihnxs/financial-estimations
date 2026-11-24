"""Deck A visuals, v2: old-proposal look (navy gradient, sand, photos, chevron tags) but bolder.

Each slide = a PNG design layer (rendered by headless Chrome) + photo slots.
Photos are NOT baked in: they are placed in Canva as separate, replaceable images
at the exact slot geometry written to deck_visuals/slots.json.
Logo (top-left) and date pill (top-right) stay as Canva elements.

Run:  /opt/homebrew/bin/python3 -I deck_visuals_v2.py
"""
import json
import math
import os
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "deck_visuals")
SRC = os.path.join(OUT, "html")
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
os.makedirs(SRC, exist_ok=True)

NAVY, MID, LIGHT = "#104a70", "#4f5c74", "#628cab"
SAND, SAND2, CREAM, INK = "#e8d4ab", "#f6e3b6", "#fdfaf1", "#163b42"

CSS = f"""
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800;900&display=swap');
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{width:1920px;height:1080px;overflow:hidden}}
body{{font-family:'Poppins',sans-serif;color:#fff;position:relative;background:linear-gradient(90deg,{NAVY} 0%,{MID} 52%,{LIGHT} 100%)}}
.abs{{position:absolute}}
.blob{{position:absolute;background:{SAND};opacity:.10}}
.dots{{position:absolute;background-image:radial-gradient(rgba(246,227,182,.55) 2.5px,transparent 3px);background-size:26px 26px}}
.tag{{position:absolute;left:0;height:84px;padding:0 70px 0 108px;display:flex;align-items:center;background:{SAND};color:{NAVY};
  font-weight:700;font-size:30px;letter-spacing:.5px;clip-path:polygon(0 0,calc(100% - 42px) 0,100% 50%,calc(100% - 42px) 100%,0 100%)}}
.h{{font-weight:800;letter-spacing:-2.5px;line-height:1.02}}
.chip{{display:inline-flex;align-items:center;gap:10px;padding:10px 22px;border-radius:999px;border:2px solid rgba(255,255,255,.35);
  background:rgba(255,255,255,.08);font-size:22px;font-weight:500;margin:0 12px 14px 0;white-space:nowrap}}
.sticker{{position:absolute;border-radius:50%;background:{SAND2};color:{NAVY};display:flex;flex-direction:column;align-items:center;
  justify-content:center;text-align:center;font-weight:800;box-shadow:0 14px 30px rgba(0,0,0,.25)}}
.card{{position:absolute;background:{CREAM};color:{INK};border-radius:30px;box-shadow:0 26px 60px rgba(6,28,48,.35)}}
.slot{{position:absolute;background:repeating-linear-gradient(45deg,#9fb3c6 0 18px,#b8c7d6 18px 36px)}}
svg.i{{flex-shrink:0;stroke:currentColor;fill:none;stroke-width:2.4;stroke-linecap:round;stroke-linejoin:round}}
"""

ICON = {
    "search": '<circle cx="28" cy="28" r="16"/><path d="M40 40l14 14"/><path d="M21 28h14M28 21v14"/>',
    "heart": '<path d="M32 54S10 40 10 24a11 11 0 0 1 22-3 11 11 0 0 1 22 3c0 16-22 30-22 30z"/>',
    "grow": '<path d="M32 56V30"/><path d="M32 34c0-10 8-16 18-16 0 10-8 16-18 16z"/><path d="M32 40c0-8-6-13-14-13 0 8 6 13 14 13z"/><path d="M18 56h28"/>',
    "book": '<path d="M10 14h18a6 6 0 0 1 6 6v32a5 5 0 0 0-5-5H10z"/><path d="M54 14H36a6 6 0 0 0-6 6v32a5 5 0 0 1 5-5h19z"/>',
    "users": '<circle cx="24" cy="22" r="8"/><path d="M10 50c0-8 6-14 14-14s14 6 14 14"/><circle cx="44" cy="24" r="6"/><path d="M42 36c7 0 12 5 12 12"/>',
    "org": '<rect x="24" y="8" width="16" height="12" rx="3"/><rect x="8" y="44" width="16" height="12" rx="3"/><rect x="40" y="44" width="16" height="12" rx="3"/><path d="M32 20v12M16 44v-6h32v6"/>',
    "tired": '<circle cx="32" cy="32" r="22"/><path d="M22 28h6M36 28h6"/><path d="M24 42c5-3 11-3 16 0"/>',
    "focus": '<circle cx="32" cy="32" r="22"/><circle cx="32" cy="32" r="12"/><circle cx="32" cy="32" r="3"/>',
    "door": '<path d="M16 56V10h26v46"/><path d="M42 18l10 4v34"/><circle cx="35" cy="34" r="2"/><path d="M8 56h48"/>',
    "check": '<path d="M14 33l12 12 24-26"/>',
    "chart": '<path d="M10 54h44"/><rect x="14" y="34" width="8" height="16"/><rect x="28" y="24" width="8" height="26"/><rect x="42" y="14" width="8" height="36"/>',
    "shield": '<path d="M32 8l20 8v14c0 13-9 22-20 26-11-4-20-13-20-26V16z"/><path d="M24 32l6 6 11-12"/>',
}


def icon(name, size, color):
    return f'<svg class="i" viewBox="0 0 64 64" style="width:{size}px;height:{size}px;color:{color}">{ICON[name]}</svg>'


SQUIGGLE = ('<svg class="abs" style="left:{x}px;top:{y}px" width="{w}" height="60" viewBox="0 0 {w} 60">'
            '<path d="{d}" fill="none" stroke="#f6e3b6" stroke-width="7" stroke-linecap="round"/></svg>')


def squiggle(x, y, w=260):
    seg, d, k = 40, "M6 30", 0
    while (k + 1) * seg < w - 6:
        d += f" q{seg/2} {-22 if k % 2 == 0 else 22} {seg} 0"
        k += 1
    return SQUIGGLE.format(x=x, y=y, w=w, d=d)


def base_deco():
    return ('<div class="blob" style="width:760px;height:620px;left:-260px;top:520px;border-radius:58% 42% 61% 39%/45% 55% 45% 55%"></div>'
            '<div class="dots" style="width:260px;height:180px;right:60px;bottom:40px;opacity:.5"></div>')


SLOTS = {}


def polaroid(page, media, fx, fy, fw, fh, rot, pad=18, bottom=70, tape=True, caption=""):
    """White polaroid frame (in PNG) + photo slot (placed in Canva). Returns HTML."""
    pw, ph = fw - 2 * pad, fh - pad - bottom
    # photo centre = frame centre + R(rot) * offset
    ox, oy = 0, (pad + ph / 2) - fh / 2
    a = math.radians(rot)
    cx, cy = fx + fw / 2, fy + fh / 2
    pcx = cx + ox * math.cos(a) - oy * math.sin(a)
    pcy = cy + ox * math.sin(a) + oy * math.cos(a)
    SLOTS.setdefault(page, []).append({"media": media, "left": round(pcx - pw / 2, 2), "top": round(pcy - ph / 2, 2),
                                       "width": pw, "height": ph, "rotation": rot})
    t = (f'<div class="abs" style="left:{fw/2-70}px;top:-30px;width:140px;height:44px;background:rgba(246,227,182,.9);transform:rotate(-4deg)"></div>'
         if tape else "")
    cap = (f'<div class="abs" style="left:0;bottom:16px;width:100%;text-align:center;color:{NAVY};font-weight:600;font-size:24px">{caption}</div>'
           if caption else "")
    return (f'<div class="abs" style="left:{fx}px;top:{fy}px;width:{fw}px;height:{fh}px;transform:rotate({rot}deg);background:#fff;'
            f'box-shadow:0 30px 60px rgba(4,24,40,.38)"><div class="slot" style="left:{pad}px;top:{pad}px;width:{pw}px;height:{ph}px"></div>{t}{cap}</div>')


def rect_slot(page, media, x, y, w, h, rot=0):
    """Photo slot inside an element that itself is rotated by `rot` around (cx,cy) is handled by caller."""
    SLOTS.setdefault(page, []).append({"media": media, "left": x, "top": y, "width": w, "height": h, "rotation": rot})


def page(body):
    return f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head><body>{base_deco()}{body}</body></html>'


HTML = {}

# ================= Dividers =================
DIV = [
    ("01", "Asesmen", ["Mental Health Check-Up", "Alur &amp; privasi data", "Pemetaan talenta", "Asesmen individu &amp; karier", "Tumbuh kembang anak"],
     "MADatcJ6nD0", "Ukur dulu,<br>baru bertindak"),
    ("02", "Konseling &amp;<br>Pendampingan", ["Kapan perlu sesi", "Paket EAP", "EAP Korporat", "In-house Psychologist", "Konseling individu &amp; pasangan", "Pendampingan kelompok"],
     "MAG_V4SBhE0", "Ruang aman,<br>rahasia terjaga"),
    ("03", "Pembelajaran &amp;<br>Psikoedukasi", ["Pelatihan &amp; seminar", "Healthy Working Environment", "Topik webinar", "5 workshop sekolah", "Team building"],
     "MAHG7bJiC8c", "Interaktif &amp;<br>kontekstual"),
    ("04", "Pengalaman &amp;<br>Komunitas", ["Art-Based Wellbeing", "Teman Kecil", "Event komunitas"],
     "MAHCNAnKszc", "Jeda yang<br>hangat"),
    ("05", "Konsultasi<br>Organisasi", ["HR Advisory", "Linimasa kerja", "Hasil yang diterima"],
     "MAEJCqyVsIg", "Dari diagnosis<br>ke rekomendasi"),
]
for num, title, items, media, badge in DIV:
    key = f"divider_{num}"
    chips = "".join(f'<span class="chip"><span style="width:10px;height:10px;border-radius:50%;background:{SAND2}"></span>{t}</span>' for t in items)
    HTML[key] = page(f"""
<div class="abs h" style="left:80px;top:170px;font-size:380px;color:{SAND};opacity:.95;letter-spacing:-20px;line-height:.86">{num}</div>
{squiggle(500, 240, 280)}
<div class="tag" style="top:520px">Bagian {num}</div>
<div class="abs h" style="left:108px;top:628px;font-size:{88 if '<br>' in title else 110}px;width:1000px;text-shadow:0 6px 24px rgba(0,0,0,.25)">{title}</div>
<div class="abs" style="left:108px;top:{850 if '<br>' in title else 790}px;width:780px">{chips}</div>
{polaroid(key, media, 1180, 150, 600, 640, 4)}
<div class="sticker" style="left:905px;top:610px;width:230px;height:230px;font-size:30px;line-height:1.15;transform:rotate(-10deg)">{badge}</div>
""")

# ================= p7 service map =================
LINES = [("01", "Asesmen", "search", "MHCU, pemetaan talenta, tes IQ, minat bakat, tumbuh kembang"),
         ("02", "Konseling &amp; pendampingan", "heart", "EAP, in-house psychologist, konseling individu dan pasangan"),
         ("03", "Pembelajaran &amp; psikoedukasi", "book", "Pelatihan, workshop, webinar, team building"),
         ("04", "Pengalaman &amp; komunitas", "users", "Art-based wellbeing, Teman Kecil, event komunitas"),
         ("05", "Konsultasi organisasi", "org", "HR advisory, evaluasi program, riset kebutuhan")]
tiles = ""
for k, (n, t, ic, d) in enumerate(LINES):
    x = 108 + k * 344
    y = 470 if k % 2 == 0 else 530
    rot = [-2.5, 2, -1.5, 2.5, -2][k]
    bg = CREAM if k % 2 == 0 else SAND2
    tiles += f"""
<div class="card" style="left:{x}px;top:{y}px;width:320px;height:440px;padding:84px 30px 30px;background:{bg};transform:rotate({rot}deg)">
  <div style="color:{NAVY}">{icon(ic,60,NAVY)}</div>
  <div style="margin-top:18px;font-size:30px;font-weight:800;line-height:1.15;color:{NAVY};letter-spacing:-.5px">{t}</div>
  <div style="margin-top:14px;font-size:22px;line-height:1.5;color:{MID}">{d}</div>
  <div style="position:absolute;top:-38px;left:22px;width:96px;height:96px;border-radius:50%;background:{NAVY};color:{SAND2};
    display:flex;align-items:center;justify-content:center;font-weight:900;font-size:38px;border:5px solid {SAND2};box-shadow:0 10px 24px rgba(0,0,0,.25)">{n}</div>
</div>"""
HTML["p07_peta_layanan"] = page(f"""
<div class="tag" style="top:180px">Layanan Teman Journey</div>
<div class="abs h" style="left:108px;top:292px;font-size:88px">Satu mitra, lima lini layanan</div>
{squiggle(1365, 335, 170)}
<div class="sticker" style="left:1560px;top:190px;width:230px;height:230px;font-size:27px;line-height:1.2;transform:rotate(9deg)">Satu standar<br>psikolog<br>berizin</div>
{tiles}""")

# ================= p5 Kenali / Dampingi / Kembangkan =================
PIL = [("Kenali", "search", "Asesmen untuk melihat kondisi dan potensi secara utuh: MHCU, seleksi, dan pemetaan talenta.", "Manajemen menerima data agregat", "MADatcJ6nD0"),
       ("Dampingi", "heart", "Konseling bersama psikolog berizin: daring, tatap muka, atau di lokasi. Aman dan rahasia.", "Di luar lingkungan kantor dan HR", "MAEsg7Ei7t8"),
       ("Kembangkan", "grow", "Pelatihan, workshop, dan konsultasi berbasis psikologi terapan.", "Diukur sebelum dan sesudah", "MAHG7bJiC8c")]
cards = ""
for k, (t, ic, d, tag, media) in enumerate(PIL):
    x, y, w, h, rot = 108 + k * 590, 365 + (0, 30, 8)[k], 540, 630, (-2, 1.5, -1)[k]
    # photo slot inside rotated card: offset from card centre
    px, py, pw, ph = 24, 24, w - 48, 250
    a = math.radians(rot)
    ox, oy = (px + pw / 2) - w / 2, (py + ph / 2) - h / 2
    cx, cy = x + w / 2, y + h / 2
    pcx, pcy = cx + ox * math.cos(a) - oy * math.sin(a), cy + ox * math.sin(a) + oy * math.cos(a)
    rect_slot("p05_pendekatan", media, round(pcx - pw / 2, 2), round(pcy - ph / 2, 2), pw, ph, rot)
    cards += f"""
<div class="card" style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;transform:rotate({rot}deg)">
  <div class="slot" style="left:{px}px;top:{py}px;width:{pw}px;height:{ph}px;border-radius:0"></div>
  <div style="position:absolute;left:{w-70}px;top:{h-70}px;width:92px;height:92px;z-index:3;border-radius:50%;background:{SAND2};color:{NAVY};
    display:flex;align-items:center;justify-content:center;font-weight:900;font-size:40px;box-shadow:0 10px 24px rgba(0,0,0,.25)">{k+1}</div>
  <div style="position:absolute;left:36px;top:300px;right:36px">
    <div style="display:flex;align-items:center;gap:16px">{icon(ic,52,NAVY)}<div style="font-size:50px;font-weight:800;color:{NAVY};letter-spacing:-1.2px">{t}</div></div>
    <div style="margin-top:16px;font-size:25px;line-height:1.5;color:{MID}">{d}</div>
  </div>
  <div style="position:absolute;left:36px;bottom:32px"><span style="display:inline-block;background:{NAVY};color:{SAND2};padding:12px 22px;border-radius:999px;font-weight:600;font-size:20px;white-space:nowrap">{tag}</span></div>
</div>"""
HTML["p05_pendekatan"] = page(f"""
<div class="tag" style="top:160px">Pendekatan kami</div>
<div class="abs h" style="left:108px;top:262px;font-size:58px;letter-spacing:-1.5px">Satu mitra, dari memahami sampai mendampingi</div>
{cards}""")

# ================= p6 Perjalanan =================
STEPS = [("Langkah 1", "Gambaran awal", "Snapshot singkat kondisi tim, laporan agregat untuk HR. Tanpa biaya.", "MADatcJ6nD0"),
         ("Langkah 2", "Program inti", "Asesmen, pelatihan, atau workshop untuk organisasi Anda, dengan ukuran keberhasilan yang disepakati di awal.", "MAHG7RpTjCo"),
         ("Langkah 3", "Jaga hasilnya", "EAP, MHCU berkala, atau psikolog di lokasi, supaya perubahan bertahan.", "MAG_V4SBhE0")]
road = (f'<svg class="abs" style="left:0;top:0" width="1920" height="1080"><path d="M120 640 C 420 520, 520 760, 820 600 S 1260 420, 1560 520 S 1820 470, 1900 430" '
        f'fill="none" stroke="{SAND2}" stroke-width="8" stroke-dasharray="4 22" stroke-linecap="round" opacity=".85"/></svg>')
steps = ""
for k, (lbl, t, d, media) in enumerate(STEPS):
    cx = 108 + k * 600
    py = (480, 440, 400)[k]
    rot = (-5, 3, -3)[k]
    steps += polaroid("p06_perjalanan", media, cx + 40, py, 400, 300, rot, pad=16, bottom=50, tape=k != 1)
    steps += f"""
<div class="abs" style="left:{cx+10}px;top:{py-36}px;width:84px;height:84px;border-radius:50%;background:{SAND2};color:{NAVY};display:flex;
  align-items:center;justify-content:center;font-weight:900;font-size:40px;box-shadow:0 10px 24px rgba(0,0,0,.3);z-index:5">{k+1}</div>
<div class="abs" style="left:{cx+20}px;top:{py+335}px;width:520px">
  <div style="font-size:22px;font-weight:600;letter-spacing:3px;text-transform:uppercase;color:{SAND2}">{lbl}</div>
  <div class="h" style="margin-top:4px;font-size:48px;letter-spacing:-1px">{t}</div>
  <div style="margin-top:8px;font-size:23px;line-height:1.45;font-weight:300">{d}</div>
</div>"""
HTML["p06_perjalanan"] = page(f"""
{road}
<div class="tag" style="top:160px">Cara kita bekerja sama</div>
<div class="abs h" style="left:108px;top:262px;font-size:76px;letter-spacing:-2px">Perjalanan bersama kami</div>
{steps}""")

# ================= p3 Masalah (photo right + old-proposal wave placed in Canva) =================
rect_slot("p03_masalah", "MAGcwCIsObk", 1100, 0, 820, 1080, 0)
SIG = [("tired", "Kelelahan", -4), ("focus", "Sulit fokus", 3), ("door", "Diam-diam bersiap pindah", -2)]
sig = "".join(f'<div style="display:inline-flex;align-items:center;gap:14px;padding:16px 26px;margin:0 14px 16px 0;border-radius:999px;background:{SAND2};'
              f'color:{NAVY};font-weight:700;font-size:28px;transform:rotate({r}deg);box-shadow:0 10px 24px rgba(0,0,0,.22)">{icon(ic,40,NAVY)}{t}</div>'
              for ic, t, r in SIG)
HTML["p03_masalah"] = page(f"""
<div class="slot" style="left:1100px;top:0;width:820px;height:1080px"></div>
<div class="tag" style="top:180px">Yang sering kami dengar</div>
<div class="abs h" style="left:108px;top:296px;font-size:60px;width:960px;letter-spacing:-1.5px">Masalahnya sering baru terlihat setelah kinerja turun</div>
<div class="abs" style="left:108px;top:470px;width:960px">{sig}</div>
<div class="abs" style="left:108px;top:680px;width:900px;font-size:23px;line-height:1.55;font-weight:300">
  Tuntutan kerja tinggi tanpa sumber daya cukup menguras energi psikologis dan menaikkan risiko burnout
  <b style="color:{SAND2};font-weight:600">(Bakker &amp; Demerouti, 2017)</b>. WHO mengakui burnout sebagai fenomena okupasional <b style="color:{SAND2};font-weight:600">(2019)</b>.</div>
<div class="abs" style="left:108px;top:840px;width:920px;height:160px;border-radius:26px;border:3px dashed rgba(246,227,182,.7);background:rgba(255,255,255,.05)"></div>
<div class="abs" style="left:132px;top:822px;padding:4px 14px;background:{NAVY};color:{SAND2};font-size:18px;font-weight:600;letter-spacing:2px">SITUASI KLIEN</div>
""")

# ================= p10 MHCU (staircase tiers) =================
ROWS = ["Instrumen", "Review psikolog SIPP", "Laporan", "Konseling"]
FIT = ["Pemetaan cepat seluruh pegawai", "Tindak lanjut cepat per individu", "Pegawai kunci dan kasus prioritas"]
TIERS = [("Skrining", "Gambaran cepat untuk banyak orang", ["DASS-21 online", "check", "Individual, 48 jam", "Tidak termasuk"], 0),
         ("Flash", "Skrining plus sesi singkat", ["5 instrumen klinis", "check", "Laporan personal", "30 menit"], 1),
         ("Deep", "Pemeriksaan paling lengkap", ["5 instrumen klinis", "check", "Laporan personal", "60 menit"], 2)]
tc = ""
for name, sub, vals, k in TIERS:
    dark = k == 2
    x, w = 108 + k * 590, 550
    top = (490, 435, 380)[k]
    h = 1010 - top
    bg = f"linear-gradient(160deg,#0b3150,{NAVY})" if dark else (CREAM if k == 0 else SAND2)
    fg = "#fff" if dark else INK
    line = "rgba(255,255,255,.18)" if dark else "rgba(16,74,112,.15)"
    def v(x_, dark=dark):
        if x_ == "check":
            return icon("check", 40, SAND2 if dark else NAVY)
        op = ".55" if x_ == "Tidak termasuk" else "1"
        return f'<span style="font-size:24px;font-weight:700;opacity:{op}">{x_}</span>'
    rows = "".join(f'<div style="height:{56 if k==0 else 64 if k==1 else 76}px;display:flex;justify-content:space-between;align-items:center;border-top:2px solid {line}">'
                   f'<span style="font-size:17px;letter-spacing:1.5px;text-transform:uppercase;opacity:.7;width:180px">{r}</span>{v(val)}</div>'
                   for r, val in zip(ROWS, vals))
    tc += f"""
<div class="card" style="left:{x}px;top:{top}px;width:{w}px;height:{h}px;padding:34px 38px;background:{bg};color:{fg};{'border:4px solid '+SAND2 if dark else ''}">
  <div style="display:flex;align-items:baseline;gap:16px"><div class="h" style="font-size:58px;letter-spacing:-1.5px;color:{SAND2 if dark else NAVY}">{name}</div>
    <div style="font-size:22px;font-weight:700;opacity:.5">0{k+1}</div></div>
  <div style="font-size:22px;margin:2px 0 16px;opacity:.8">{sub}</div>{rows}
  <div style="position:absolute;left:38px;right:38px;bottom:30px;padding:14px 20px;border-radius:18px;background:{'rgba(246,227,182,.16)' if dark else 'rgba(16,74,112,.08)'}">
    <div style="font-size:16px;letter-spacing:2px;text-transform:uppercase;opacity:.7">Cocok untuk</div>
    <div style="font-size:22px;font-weight:700;margin-top:2px">{FIT[k]}</div></div>
</div>"""
HTML["p10_mhcu"] = page(f"""
<div class="tag" style="top:160px">Asesmen</div>
<div class="abs h" style="left:108px;top:262px;font-size:80px">Mental Health Check-Up</div>
<div class="abs" style="left:108px;top:362px;width:1040px;font-size:27px;font-weight:300">Tiga tingkat, satu standar: instrumen tervalidasi dan laporan yang ditinjau psikolog berizin.</div>
<div class="sticker" style="left:1530px;top:168px;width:200px;height:200px;font-size:24px;line-height:1.2;transform:rotate(10deg)">Paling<br>lengkap<br>&darr;</div>
{tc}""")

# ---------- render ----------
for name, html in HTML.items():
    hp = os.path.join(SRC, name + ".html")
    open(hp, "w").write(html)
    png = os.path.join(OUT, name + ".png")
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                    "--window-size=1920,1080", "--virtual-time-budget=4000", f"--screenshot={png}", "file://" + hp],
                   check=True, capture_output=True)
    print("rendered", name)
json.dump(SLOTS, open(os.path.join(OUT, "slots.json"), "w"), indent=1)
print("slots ->", sum(len(v) for v in SLOTS.values()), "photo slots")
