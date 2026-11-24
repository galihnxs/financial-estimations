"""Deck B (INTERNAL product & pricing guide) page images, 1920x1080.

All numbers come from 01_Daftar_Produk_Final.md / Master_Produk_dan_Harga.xlsx (20 Nov 2026).
When prices change: update the data below, re-run, re-upload. Logo (top-left) stays a Canva element.
Run: /opt/homebrew/bin/python3 -I deck_b_visuals.py
"""
import os
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "deck_b")
SRC = os.path.join(OUT, "html")
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
os.makedirs(SRC, exist_ok=True)

NAVY, MID, LIGHT, SAND, SAND2, CREAM, INK = "#104a70", "#4f5c74", "#628cab", "#e8d4ab", "#f6e3b6", "#fdfaf1", "#163b42"
RED, YEL, GRN = "#c8553d", "#e0a526", "#3f8f5a"

CSS = f"""
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap');
*{{box-sizing:border-box;margin:0;padding:0}} html,body{{width:1920px;height:1080px;overflow:hidden}}
body{{font-family:'Poppins',sans-serif;color:#fff;position:relative;background:linear-gradient(90deg,{NAVY} 0%,{MID} 55%,{LIGHT} 100%)}}
.abs{{position:absolute}}
.ribbon{{position:absolute;right:0;top:70px;height:66px;padding:0 60px 0 70px;background:{SAND};color:{NAVY};font-weight:800;font-size:22px;
  letter-spacing:2px;display:flex;align-items:center;gap:14px;clip-path:polygon(34px 0,100% 0,100% 100%,34px 100%,0 50%)}}
.ribbon i{{width:14px;height:14px;border-radius:50%;background:{RED};display:inline-block}}
.foot{{position:absolute;left:108px;bottom:34px;font-size:18px;opacity:.6;letter-spacing:1px}}
.kick{{font-size:24px;font-weight:600;letter-spacing:3px;text-transform:uppercase;color:{SAND2}}}
.h{{font-weight:800;letter-spacing:-2px;line-height:1.04}}
.card{{background:{CREAM};color:{INK};border-radius:24px;box-shadow:0 18px 40px rgba(6,28,48,.28)}}
.glass{{background:rgba(255,255,255,.08);border:2px solid rgba(255,255,255,.2);border-radius:24px}}
.lab{{font-size:17px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:{LIGHT}}}
.pill{{display:inline-block;padding:6px 16px;border-radius:999px;font-weight:700;font-size:18px}}
b{{font-weight:700}}
"""


def page(body, foot=""):
    return (f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head><body>'
            f'<div class="abs" style="width:700px;height:560px;left:-240px;top:600px;background:{SAND};opacity:.08;border-radius:58% 42% 61% 39%/45% 55% 45% 55%"></div>'
            f'<div class="ribbon"><i></i>INTERNAL · JANGAN DIKIRIM KE KLIEN</div>{body}'
            f'<div class="foot">Deck B · Panduan Produk &amp; Harga · 20 Nov 2026{(" · " + foot) if foot else ""}</div></body></html>')


def header(kick, title, sub="", top=170, size=68):
    s = f'<div class="abs" style="left:108px;top:{top + 36 + int(size * 1.08) + 16}px;width:1500px;font-size:26px;font-weight:300;opacity:.92">{sub}</div>' if sub else ""
    return (f'<div class="abs kick" style="left:108px;top:{top}px">{kick}</div>'
            f'<div class="abs h" style="left:108px;top:{top + 36}px;font-size:{size}px;width:1700px">{title}</div>{s}')


P = {}

# 1 Cover --------------------------------------------------------------------------------------
P["b01_cover"] = page(f"""
<div class="abs" style="left:108px;top:300px;width:1300px">
  <div class="kick">Deck B · untuk tim sales dan BD</div>
  <div class="h" style="font-size:120px;margin-top:20px">Panduan Produk<br>&amp; Harga</div>
  <div style="margin-top:34px;font-size:32px;font-weight:300;line-height:1.5;width:1100px">Apa yang kita jual, kepada siapa, kapan menawarkannya, berapa harganya, dan aturan yang harus diikuti.
  Teman Journey + Coreitera.</div>
</div>
<div class="abs card" style="left:1250px;top:330px;width:560px;padding:40px">
  <div class="lab">Sumber angka</div>
  <div style="margin-top:10px;font-size:24px;line-height:1.55"><b>Kalkulator_Sales.xlsx</b> untuk menyusun penawaran.<br><b>Master_Produk_dan_Harga.xlsx</b> untuk tim harga.</div>
  <div class="lab" style="margin-top:26px">Kalau berbeda</div>
  <div style="margin-top:10px;font-size:24px;line-height:1.55">Yang dipegang adalah Excel. Deck ini dicetak ulang setiap kali harga berubah.</div>
</div>""")

# 2 Cara memakai -------------------------------------------------------------------------------
RULES3 = [("1", "Jangan pernah dikirim ke klien", "Deck ini berisi target margin, aturan diskon, dan komisi. Untuk klien pakai Deck A."),
          ("2", "Harga selalu dari Kalkulator", "Masukkan segmen, kode produk, dan jumlah. Warna hijau, kuning, atau merah menunjukkan apakah penawaran perlu persetujuan."),
          ("3", "Semua harga belum PPN", "Teman Journey dan Coreitera bukan PKP. Jangan menulis \"termasuk PPN\". PPh 23 dipotong klien.")]
cards = "".join(f"""<div class="card" style="width:520px;padding:40px 38px">
  <div style="width:70px;height:70px;border-radius:50%;background:{NAVY};color:{SAND2};font-weight:800;font-size:34px;display:flex;align-items:center;justify-content:center">{n}</div>
  <div style="margin-top:22px;font-size:34px;font-weight:800;color:{NAVY};line-height:1.15">{t}</div>
  <div style="margin-top:14px;font-size:24px;line-height:1.5;color:{MID}">{d}</div></div>""" for n, t, d in RULES3)
P["b02_cara"] = page(header("Mulai di sini", "Cara memakai deck ini", "Pegang tiga aturan ini, lalu baca halaman lini yang sesuai dengan klien Anda.")
                     + f'<div class="abs" style="left:108px;top:430px;display:flex;gap:42px">{cards}</div>'
                     + f'<div class="abs" style="left:108px;top:930px;font-size:24px;opacity:.9">Urutan: peta lini (hal. 3) → segmen dan margin (4-6) → kartu jual dan harga per lini (7-18) → aturan (19-21) → Excel (22-23).</div>')

# 3 Peta 6 lini ---------------------------------------------------------------------------------
LINES = [("01", "Asesmen", "Kenali orang-orang Anda", "MHCU, seleksi, pemetaan talenta, tes sekolah"),
         ("02", "Konseling &amp; pendampingan", "Teman perjalanan untuk orang-orang Anda", "EAP, sesi B2B, in-house psychologist"),
         ("03", "Pembelajaran &amp; psikoedukasi", "Bantu orang-orang Anda belajar", "Program pelatih, pelatihan bank, pembicara, workshop"),
         ("04", "Pengalaman &amp; komunitas", "Momen bermakna untuk komunitas", "Event, Teman Kecil, jasa event mitra"),
         ("05", "Konsultasi organisasi", "Sistem dan program yang lebih baik", "Juri ahli, hari konsultan, FGD"),
         ("06", "Coreitera", "Teknologi untuk biro psikologi", "Theta, Synthesia, mesin tes white-label")]
tiles = ""
for k, (n, t, tag, ex) in enumerate(LINES):
    x, y = 108 + (k % 3) * 580, 400 + (k // 3) * 300
    dark = k == 5
    tiles += f"""<div class="abs {'glass' if dark else 'card'}" style="left:{x}px;top:{y}px;width:550px;height:270px;padding:30px 34px">
  <div style="display:flex;justify-content:space-between;align-items:center"><div style="font-size:30px;font-weight:800;color:{SAND2 if dark else NAVY}">{t}</div>
  <div style="font-size:44px;font-weight:800;color:{SAND if not dark else SAND2};opacity:.9">{n}</div></div>
  <div style="margin-top:8px;font-size:22px;font-style:italic;color:{SAND2 if dark else LIGHT}">"{tag}"</div>
  <div style="margin-top:16px;font-size:23px;line-height:1.45;color:{'#fff' if dark else MID}">{ex}</div></div>"""
P["b03_peta"] = page(header("Peta produk", "Enam lini, dua perusahaan", "Lini 01-05 dijual Teman Journey ke klien. Lini 06 dijual Coreitera ke biro psikologi, tidak pernah langsung ke korporat.") + tiles)

# 4 Segmen ---------------------------------------------------------------------------------------
SEG = [("Pendidikan &amp; komunitas", "Sekolah, kampus, yayasan, komunitas", "30%", "x 1,55", "Kontrak min. Rp10 jt. Bagi hasil sekolah maks. 10%. Tawarkan versi yang lebih ringan, jangan diskon."),
       ("Korporat &amp; pemerintah", "Bank, BUMN, swasta, kementerian, pemda", "50%", "x 2,36", "Harga hari ini yang lebih tinggi tetap dipakai, tidak diturunkan."),
       ("Regulator", "BI, OJK, LPS", "50% (min. 30%)", "s.d. x 2,36", "Dijual sebagai program dengan acuan pagu SBM. Item yang sudah pernah dibeli tetap memakai harga lama."),
       ("Individu", "Klien perorangan, pasangan", "Harga pasar", "-", "Harga per tier psikolog. Margin TJ 24-34% setelah fee psikolog."),
       ("Biro psikologi", "Pembeli Coreitera", "Langganan", "-", "Harga langganan per biro atau per kursi. Coreitera tidak menjual tes ke korporat.")]
rows = ""
for k, (s, who, m, mult, r) in enumerate(SEG):
    y = 380 + k * 128
    rows += f"""<div class="abs card" style="left:108px;top:{y}px;width:1704px;height:112px;padding:0 34px;display:flex;align-items:center;gap:30px">
  <div style="width:400px"><div style="font-size:28px;font-weight:800;color:{NAVY}">{s}</div><div style="font-size:19px;color:{MID}">{who}</div></div>
  <div style="width:260px;text-align:center"><div class="lab">Target margin</div><div style="font-size:{30 if len(m) < 12 else 23}px;font-weight:800;color:{NAVY};white-space:nowrap">{m}</div></div>
  <div style="width:150px;text-align:center"><div class="lab">Pengali biaya</div><div style="font-size:26px;font-weight:700;color:{MID}">{mult}</div></div>
  <div style="flex:1;font-size:21px;line-height:1.4;color:{INK}">{r}</div></div>"""
P["b04_segmen"] = page(header("Segmen", "Lima segmen, lima aturan harga", "Margin = margin bersih setelah PPh 23 (2%) dan komisi sales 10% dari laba kotor.") + rows)

# 5 Satu biaya ---------------------------------------------------------------------------------
P["b05_rumus"] = page(header("Cara harga dihitung", "Satu biaya, harga berbeda per segmen",
                             "Metode, instrumen, dan review psikolog sama. Yang berbeda: durasi, senioritas psikolog, dan harga.") + f"""
<div class="abs glass" style="left:108px;top:400px;width:760px;height:520px;padding:44px">
  <div class="lab" style="color:{SAND2}">Rumus</div>
  <div style="margin-top:18px;font-size:46px;font-weight:800;line-height:1.25">Harga = biaya × 0,9<br>÷ (0,882 − target)</div>
  <div style="margin-top:26px;font-size:24px;line-height:1.55;opacity:.92">Lalu dibulatkan ke atas: Rp1rb di bawah Rp100rb, Rp25rb di bawah Rp1 jt, Rp100rb di bawah Rp10 jt, Rp500rb di atasnya.
  Anda tidak perlu menghitungnya sendiri karena Excel sudah melakukannya.</div>
</div>
<div class="abs card" style="left:910px;top:400px;width:900px;height:520px;padding:44px">
  <div class="lab">Contoh: MHCU Flash (konseling 30 menit)</div>
  <div style="margin-top:14px;font-size:26px;color:{MID}">Biaya per orang <b style="color:{NAVY}">Rp244,5rb</b>, sama untuk semua klien.</div>
  <div style="display:flex;gap:24px;margin-top:30px">
    <div style="flex:1;padding:26px;border-radius:20px;background:{SAND2}"><div class="lab" style="color:{NAVY}">Pendidikan</div>
      <div style="font-size:52px;font-weight:800;color:{NAVY}">Rp400rb</div><div style="font-size:22px;color:{MID}">margin 33%</div></div>
    <div style="flex:1;padding:26px;border-radius:20px;background:{NAVY};color:#fff"><div class="lab" style="color:{SAND2}">Korporat &amp; regulator</div>
      <div style="font-size:52px;font-weight:800">Rp600rb</div><div style="font-size:22px;opacity:.8">margin 52%</div></div>
  </div>
  <div style="margin-top:28px;font-size:23px;line-height:1.5;color:{INK}">Kalau sekolah meminta harga korporat diturunkan, jangan memberi diskon. Tawarkan <b>MHCU Ringkas 15 menit</b> (Rp250rb) yang memang dibuat untuk pendidikan.</div>
</div>""")

# 6 Peran + tangga MHCU -----------------------------------------------------------------------------
ROLES = [("Pintu masuk", "Biaya layanannya hampir nol, misalnya tes mandiri, snapshot burnout, webinar rekaman, dan konten.", "Gratis atau di bawah pasar", GRN),
         ("Inti", "Semua yang memakai waktu psikolog per orang atau per acara.", "Harga segmen. Tidak pernah gratis.", NAVY),
         ("Retensi", "Kontrak berulang: EAP, in-house, Bank Elite, langganan Theta.", "Harga segmen, dalam paket", MID)]
rc = "".join(f"""<div class="card" style="width:540px;padding:30px 32px;border-top:10px solid {c}">
  <div style="font-size:32px;font-weight:800;color:{NAVY}">{t}</div><div style="margin-top:8px;font-size:21px;line-height:1.45;color:{MID};height:94px">{d}</div>
  <div style="margin-top:12px"><span class="pill" style="background:{c};color:#fff">{rule}</span></div></div>""" for t, d, rule, c in ROLES)
LAD = [("L0", "Snapshot", "Gratis"), ("L1", "Skrining", "Rp125-175rb"), ("L2", "Ringkas 15'", "Rp250rb (pend.)"), ("L3", "Flash 30'", "Rp400-600rb"),
       ("L4", "Deep 60'", "Rp650rb-1 jt"), ("L5", "Di dalam EAP", "paket")]
lad = "".join(f"""<div style="flex:1;margin-top:{(5 - i) * 18}px;padding:18px 14px;border-radius:18px;background:{'rgba(255,255,255,.1)' if i == 0 else (SAND2 if i < 5 else SAND)};
  color:{'#fff' if i == 0 else NAVY};text-align:center"><div style="font-size:30px;font-weight:800">{l}</div><div style="font-size:20px;font-weight:600">{n}</div>
  <div style="font-size:17px;opacity:.8">{p}</div></div>""" for i, (l, n, p) in enumerate(LAD))
P["b06_peran"] = page(header("Peran produk", "Pintu masuk, inti, retensi", "") + f'<div class="abs" style="left:108px;top:300px;display:flex;gap:42px">{rc}</div>'
                      + f"""<div class="abs" style="left:108px;top:640px;width:1704px">
  <div class="kick">Tangga MHCU: naik dari gratis ke paket</div>
  <div style="display:flex;gap:16px;align-items:flex-start;margin-top:14px">{lad}</div>
  <div style="margin-top:22px;font-size:21px;opacity:.92">Pengaman pintu masuk: selalu diikuti proposal produk inti; subsidi maks. Rp6 jt per perusahaan; dihentikan kalau konversi &lt; 5% dalam 6 bulan;
  <b>tidak gratis untuk regulator dan bank</b> (untuk mereka: pilot berbayar). MHCU Flash dan Deep termasuk produk inti.</div></div>""")


# Line sales cards --------------------------------------------------------------------------------
def line_card(num, name, tag, who, when, entry, inc, exc, objs):
    def box(x, y, w, h, lab, html, dark=False):
        return (f'<div class="abs {"glass" if dark else "card"}" style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;padding:24px 28px">'
                f'<div class="lab" style="color:{SAND2 if dark else LIGHT}">{lab}</div><div style="margin-top:8px;font-size:21px;line-height:1.45">{html}</div></div>')
    o = "".join(f'<div style="margin-top:12px"><b style="color:{SAND2}">"{q}"</b><br><span style="opacity:.92">{a}</span></div>' for q, a in objs)
    return page(f"""
<div class="abs kick" style="left:108px;top:170px">Kartu jual · Lini {num}</div>
<div class="abs h" style="left:108px;top:206px;font-size:64px">{name}</div>
<div class="abs" style="left:108px;top:286px;font-size:25px;font-style:italic;color:{SAND2}">"{tag}"</div>
{box(108, 350, 540, 250, "Siapa pembelinya", who)}
{box(668, 350, 540, 250, "Kapan ditawarkan", when)}
{box(108, 620, 540, 160, "Pintu masuk", entry)}
{box(668, 620, 540, 160, "Termasuk", inc)}
{box(108, 800, 1100, 170, "Tidak termasuk / jangan dijanjikan", exc)}
<div class="abs glass" style="left:1240px;top:350px;width:572px;height:620px;padding:28px 30px">
  <div class="lab" style="color:{SAND2}">Keberatan dan jawabannya</div><div style="font-size:20px;line-height:1.45">{o}</div></div>""", f"Lini {num}")


P["b07_asm_kartu"] = line_card("01", "Asesmen", "Kenali orang-orang Anda",
    "HR dan HC perusahaan, bank, regulator (seleksi, Fit &amp; Proper), sekolah (tes siswa), orang tua (tumbuh kembang).",
    "Sebelum program dimulai (baseline). Saat promosi, rotasi, atau rekrutmen. Awal tahun ajaran. Saat HR bertanya \"kondisi tim kami sebenarnya bagaimana?\".",
    "Snapshot Burnout L0 gratis untuk korporat swasta, selalu diikuti proposal MHCU atau EAP.",
    "Instrumen tervalidasi, review psikolog SIPP, laporan per orang; agregat anonim untuk manajemen.",
    "Diagnosis klinis. Laporan individual ke manajemen (MHCU). Konseling di MHCU Skrining. Perjalanan di luar Jakarta (ditagih sebesar biaya aktual). "
    "Tanyakan dulu ke HC keputusan apa yang akan diambil. Kalau untuk promosi, tawarkan seleksi atau pemetaan 3B/3C, jangan profiling biasa.",
    [("Mahal, di aplikasi ada tes Rp15 ribu.", "Aplikasi hanya memberi skor, sedangkan kita memberi interpretasi psikolog yang bisa dipakai untuk mengambil keputusan. Kalau butuh skor massal, Profiling Paket 1 setara Rp25rb per orang."),
     ("Data pegawai kami aman?", "Manajemen hanya melihat agregat anonim. Hasil individual hanya untuk pegawai. Draf AI tidak dipakai untuk melatih model dan selalu direview psikolog.")])


def price_page(key, num, title, sub, rows, note):
    head = f"""<div class="abs" style="left:108px;top:{360}px;width:1704px;display:flex;padding:0 28px;font-size:17px;font-weight:700;letter-spacing:2px;
  text-transform:uppercase;color:{SAND2}"><div style="width:560px">Produk</div><div style="width:200px">Satuan</div>
  <div style="width:250px;text-align:center">Pendidikan · 30%</div><div style="width:250px;text-align:center">Korporat · 50%</div>
  <div style="width:250px;text-align:center">Regulator</div></div>"""
    h = min(78, int(560 / max(len(rows), 1)) - 8)
    body = ""
    for k, (prod, unit, e, c, r, n) in enumerate(rows):
        y = 400 + k * (h + 8)
        def cell(v, strong):
            if v in ("-", ""):
                return f'<div style="width:250px;text-align:center;color:#b9b1a0;font-size:24px">tidak dijual</div>'
            warn = "⚠" in v
            v = v.replace("⚠", "").strip()
            col = RED if warn else (NAVY if strong else INK)
            return f'<div style="width:250px;text-align:center;font-size:{24 if len(v) < 14 else 20}px;font-weight:800;color:{col}">{v}{" ⚠" if warn else ""}</div>'
        body += f"""<div class="abs card" style="left:108px;top:{y}px;width:1704px;height:{h}px;padding:0 28px;display:flex;align-items:center;border-radius:16px">
  <div style="width:560px"><div style="font-size:22px;font-weight:700;color:{NAVY};line-height:1.15">{prod}</div>{f'<div style="font-size:16px;color:{MID}">{n}</div>' if n else ''}</div>
  <div style="width:200px;font-size:19px;color:{MID}">{unit}</div>{cell(e, False)}{cell(c, True)}{cell(r, False)}</div>"""
    P[key] = page(header(f"Harga · Lini {num}", title, sub) + head + body
                  + f'<div class="abs" style="left:108px;top:985px;width:1704px;font-size:19px;opacity:.9">{note}</div>', f"Lini {num}")


price_page("b08_asm_mhcu", "01", "MHCU: lima tingkat", "Harga per orang, belum PPN. Rp rb = ribu, jt = juta.", [
    ("L0 · Snapshot Burnout untuk HR", "per orang", "Gratis", "Gratis", "-", "Pintu masuk. Tidak untuk regulator dan bank."),
    ("L1 · Skrining (DASS-21, review psikolog)", "per orang", "Rp125rb", "Rp175rb", "Rp175rb", "Laporan maks. 48 jam, tanpa konseling"),
    ("L2 · Ringkas, konseling 15 menit", "per orang", "Rp250rb", "-", "-", "Khusus pendidikan dan komunitas"),
    ("L3 · Flash, konseling 30 menit", "per orang", "Rp400rb", "Rp600rb", "Rp600rb", "5 instrumen klinis + laporan personal"),
    ("L4 · Deep, konseling 60 menit", "per orang", "Rp650rb", "Rp1 jt", "Rp1 jt", ""),
    ("Hasil MHCU individual (add-on EAP)", "per pegawai", "Rp54rb", "Rp82rb", "Rp70rb", "Di luar laporan agregat"),
], "Tidak ada diskon volume untuk MHCU. Harga Telering (Flash Rp242rb, Deep Rp380rb) adalah pengecualian lama dan tidak dipakai sebagai acuan.")

price_page("b09_asm_talenta", "01", "Seleksi, talenta, dan sekolah", "Proyek profiling dan pemetaan berbasis 2.000 karyawan.", [
    ("Fit &amp; Proper / seleksi (6 peserta)", "per proyek", "-", "Rp21 jt", "Rp15 jt", "Regulator: harga lama"),
    ("Asesmen Manajerial / Eksekutif", "per peserta", "-", "Rp2 / 7,6 jt", "Rp2 / 7,6 jt", "Eksekutif: wawancara 2 jam + simulasi, 2 psikolog"),
    ("Profiling Massal Paket 1 / 2", "per proyek", "-", "Rp50 / 70 jt", "Rp50 / 70 jt", "Laporan otomatis; Paket 2 karyawan baca sendiri"),
    ("Pemetaan 3A / 3B / 3C", "per proyek", "-", "Rp101,5 / 139,5 / 210,5 jt", "sama", "3B + talent box; 3C + 15 rumpun jabatan"),
    ("Tambahan tes serentak", "per proyek", "-", "Rp30 jt", "Rp30 jt", "Server, load test, engineer standby"),
    ("Paket Tes Sekolah Dasar / Standar / Lengkap", "per siswa", "Rp250 / 275 / 600rb", "-", "-", "Minimal 50 siswa"),
    ("Asesmen Perkembangan Anak (kelompok)", "per anak", "Rp500rb", "-", "-", "5-12 anak per sesi"),
], "Tes individu untuk umum (IQ, minat bakat, tumbuh kembang, kesiapan sekolah): Rp375rb-1,62 jt, harga pasar, satu harga.")

P["b10_kon_kartu"] = line_card("02", "Konseling &amp; pendampingan", "Teman perjalanan untuk orang-orang Anda",
    "HR korporat dan bank, KPw regulator, sekolah (guru dan siswa), yayasan dan program fellowship.",
    "Setelah MHCU menunjukkan risiko. Saat restrukturisasi, target tinggi, atau mutasi. Saat pegawai butuh tempat bercerita di luar HR. Kontrak tahunan saat penyusunan anggaran.",
    "Snapshot Burnout L0 (swasta) atau MHCU Skrining berbayar (bank, regulator), lalu naik ke paket.",
    "Psikolog dari luar kantor, kerahasiaan, laporan agregat anonim untuk manajemen.",
    "Sesi di atas kuota (ditagih per sesi). Laporan individual ke manajemen. EAP korporat tidak dijual ke pendidikan. "
    "Paket Jeda/Tumbuh/Seperjalanan dirancang untuk tim kecil, jadi untuk 150+ pegawai tawarkan EAP Korporat.",
    [("Kami sudah punya konselor internal / BPJS.", "Banyak pegawai tidak mau bercerita ke orang dalam, apalagi soal atasan. Karena itu EAP kita dijalankan psikolog dari luar, dan manajemen hanya melihat agregat."),
     ("Anggarannya sudah habis.", "Tanyakan kapan anggaran berikutnya disusun, lalu pastikan kita masuk ke dalamnya. Sambil menunggu, tawarkan MHCU atau webinar yang biasanya masuk pos lain.")])

price_page("b11_kon_eap", "02", "Paket EAP", "⚠ = harga lama regulator di bawah margin 30%: perpanjangannya selalu perlu Direktur.", [
    ("EAP Jeda (15 sesi, ±5 pegawai)", "per 3 bulan", "Rp11,5 jt", "Rp17,5 jt", "Rp12 jt", "1 webinar psikoedukasi"),
    ("EAP Tumbuh (24 sesi, ±8 pegawai)", "per 3 bulan", "Rp18,5 jt", "Rp28 jt", "Rp20 jt", "+ 1 putaran MHCU, laporan agregat"),
    ("EAP Seperjalanan (36 sesi, ±12 pegawai)", "per 3 bulan", "Rp28 jt", "Rp42,5 jt", "Rp31,5 jt", "+ MHCU pre/post, seminar offline"),
    ("EAP Jeda Tahunan (kontrak 12 bulan)", "per kuartal", "Rp13,5 jt", "Rp20,5 jt", "Rp11 jt ⚠", "15 sesi + MHCU 70 pegawai per kuartal"),
    ("EAP Korporat (300 pegawai)", "per bulan", "-", "Rp13,5 jt", "Rp13,5 jt", "Pool 10 sesi/bulan, MHCU 2x setahun, webinar per kuartal"),
    ("Sesi tambahan EAP 60 menit", "per sesi", "Rp600rb", "Rp925rb", "Rp500rb ⚠", "Di luar kuota paket"),
], "Pembayaran EAP 3 bulan: DP 50% saat SPK, pelunasan setelah laporan akhir (atau bulanan, total +20%).")

price_page("b12_kon_sesi", "02", "Sesi dan paket konseling", "Untuk klien yang belum siap EAP, atau tambahan di luar paket.", [
    ("Sesi konseling B2B 30 menit", "per sesi", "Rp325rb", "Rp500rb", "Rp500rb", ""),
    ("Paket 10 sesi 60 menit + MHCU", "per paket", "Rp6,8 jt", "Rp10,5 jt", "-", "10 laporan individual, 6 bulan"),
    ("Paket 10 sesi 30 menit + MHCU", "per paket", "Rp4,1 jt", "Rp6,2 jt", "-", "6 bulan"),
    ("Paket 20 sesi + MHCU pre/post", "per kontrak", "Rp15,5 jt", "Rp23,5 jt", "-", "18 bulan, hingga 29 orang"),
    ("Konseling kelompok 90 menit", "per sesi", "Rp975rb", "Rp1,5 jt", "-", "5-8 orang, 1 psikolog"),
    ("In-house Psychologist", "per hari", "Rp3,9 jt", "Rp5,9 jt", "-", "6 jam layanan. Setengah hari (sekolah): Rp2,1 jt"),
], "Konseling individu dan pasangan untuk umum: Rp500rb-3,8 jt per sesi atau paket, mengikuti tier psikolog (harga pasar).")

P["b13_plt_kartu"] = line_card("03", "Pembelajaran &amp; psikoedukasi", "Bantu orang-orang Anda belajar dan berubah",
    "Regulator (program pelatih), bank BUMN dan swasta, korporat (HWE, team building), sekolah dan kampus, komunitas.",
    "Anggaran disusun Okt-Jan untuk kegiatan tahun berikutnya: ajukan satu proposal setahun penuh. Awal tahun ajaran untuk sekolah. Setelah MHCU sebagai tindak lanjut.",
    "Webinar rekaman gratis. Webinar langsung dengan harga pembicara (harga pasar).",
    "Signature dan program pelatih: asesmen pre/post dan laporan dampak. Materi disiapkan TJ.",
    "Essential tidak termasuk asesmen dan laporan dampak. Venue, konsumsi, dan perjalanan (ditagih terpisah). Platform webinar (milik klien, kecuali beli Produksi Webinar). "
    "Jual per kelompok per hari dan jangan dalam bentuk borongan.",
    [("Trainer langganan kami lebih dikenal.", "Harga kita lebih rendah daripada trainer besar, dan setiap program dilengkapi tes sebelum-sesudah serta laporan dampak. Tawarkan satu kelompok pilot dengan ukuran keberhasilan di awal."),
     ("Bisa honor per jam saja? (regulator)", "Jual dalam bentuk program dan paket dengan acuan pagu SBM, tidak sebagai honor per jam.")])

price_page("b14_plt_program", "03", "Program pelatih dan pelatihan bank", "Diskon volume pelatihan: 5% untuk 5+ kelompok setahun, 10% untuk 10+.", [
    ("Program Pelatih Kantor Pusat (Flagship ToT)", "per program", "-", "-", "Rp100 jt", "3 hari, hingga 40 peserta, asesmen pre/post + laporan dampak"),
    ("Program Pelatih Regional (KPw / LPS)", "per hari", "-", "-", "Rp28,5 jt", "1 hari, hingga 30 peserta"),
    ("Bank Essential BUMN / Swasta", "per hari", "-", "Rp45 / 55 jt", "-", "Hari pelatihan saja, tanpa asesmen"),
    ("Bank Signature BUMN / Swasta", "per hari", "-", "Rp51 / 61 jt", "-", "+ asesmen pre/post dan laporan dampak, 40 peserta"),
    ("Bank Elite BUMN / Swasta", "per hari", "-", "Rp66,3 / 79,3 jt", "-", "+ lisensi pelatih internal, pendampingan 3 bulan"),
    ("Asesmen + laporan dampak, peserta tambahan", "per peserta", "-", "Rp225rb", "Rp150rb", "Di luar kuota paket"),
], "Regulator sudah pernah membeli program pelatih, jadi harganya dibekukan. Jual per kelompok per hari dan jangan dalam bentuk borongan.")

price_page("b15_plt_pembicara", "03", "Pembicara dan workshop", "Pembicara per 2 jam; tier mengikuti senioritas psikolog.", [
    ("Pembicara / webinar Tier 1 / 2", "per kegiatan", "Rp2,6 / 4,1 jt", "Rp3,9 / 6,3 jt", "= Korporat", "Tambahan per jam: Rp1,2 / 2 jt (pend.), Rp1,8 / 3 jt (korp.)"),
    ("Pembicara / webinar Tier 3 / 4", "per kegiatan", "Rp4,9 / 6,5 jt", "Rp7,5 / 9,8 jt", "= Korporat", "Tambahan per jam: Rp2,4 / 3,1 jt (pend.), Rp3,6 / 4,8 jt (korp.)"),
    ("Workshop Healthy Working Environment", "per sesi", "-", "Rp8,1 jt", "Rp8,1 jt", "2 jam, Tier 2, 30 peserta, worksheet"),
    ("Team-Building 4 jam", "per sesi", "-", "Rp17 jt", "Rp17 jt", "2 co-fasilitator, kit permainan"),
    ("Art-Based Wellbeing 2 jam", "per sesi", "Rp6,5 jt", "Rp9,8 jt", "Rp9,8 jt", "Kompetensi art therapy, kit seni"),
    ("Workshop sekolah dan kampus 2 jam", "per sesi", "Rp3,9 jt", "-", "-", "Psikolog kedua (versi Lengkap) +Rp2,4 jt"),
    ("Produksi webinar oleh TJ", "per webinar", "Rp3,6 jt", "Rp5,5 jt", "Rp5,5 jt", "Platform 500 orang, poster, e-sertifikat, host"),
], "Webinar rekaman dan konten: gratis (pintu masuk).")

# 15 Lini 4 + 5 kartu ------------------------------------------------------------------------------
def mini(x, title, tag, items):
    li = "".join(f'<div style="margin-top:14px"><div class="lab">{a}</div><div style="font-size:21px;line-height:1.45;margin-top:2px">{b}</div></div>' for a, b in items)
    return (f'<div class="abs card" style="left:{x}px;top:300px;width:830px;height:680px;padding:30px 36px">'
            f'<div style="font-size:36px;font-weight:800;color:{NAVY}">{title}</div><div style="font-size:21px;font-style:italic;color:{LIGHT}">"{tag}"</div>{li}</div>')


P["b16_exp_kns_kartu"] = page('<div class="abs kick" style="left:108px;top:170px">Kartu jual · Lini 04 dan 05</div>'
    '<div class="abs h" style="left:108px;top:206px;font-size:58px">Pengalaman, komunitas, dan konsultasi</div>'
    + mini(108, "04 · Pengalaman &amp; komunitas", "Momen bermakna untuk komunitas Anda", [
        ("Siapa pembelinya", "Korporat (family day, CSR), sekolah, komunitas, mitra seperti kafe atau penulis."),
        ("Kapan ditawarkan", "Family day, akhir tahun, program CSR, peluncuran buku atau acara komunitas."),
        ("Pintu masuk", "Event komunitas bertiket untuk umum (Rp259rb per peserta)."),
        ("Termasuk / tidak", "Psikolog, fasilitasi, kit aktivitas. Venue dan konsumsi tidak termasuk kecuali disebut di paket."),
        ("Keberatan", "<b>\"Anggaran sudah habis.\"</b> Event kecil dan Teman Kecil per pasangan biasanya masuk pos anggaran lain.")])
    + mini(982, "05 · Konsultasi organisasi", "Sistem dan program yang lebih baik", [
        ("Siapa pembelinya", "HC dan direksi, regulator (evaluasi program), panitia kompetisi."),
        ("Kapan ditawarkan", "Setelah asesmen atau pemetaan; evaluasi akhir program; menjelang kompetisi atau penjurian."),
        ("Pintu masuk", "FGD 2 jam online, lalu hari konsultan."),
        ("Termasuk / tidak", "Psikolog senior, analisis, laporan rekomendasi. Implementasi penuh tidak termasuk: dihitung per hari konsultan tambahan."),
        ("Catatan", "HR Advisory dan evaluasi dijual per hari konsultan, jadi jangan memberi harga proyek borongan tanpa rincian harinya.")]),
    "Lini 04-05")

price_page("b17_exp_kns_harga", "04-05", "Pengalaman, komunitas, konsultasi", "", [
    ("Event psikoedukasi publik (sponsor / CSR)", "per event", "-", "Rp24 jt", "Rp24 jt", "Dikelola TJ: venue, materi, fasilitator"),
    ("Teman Kecil Paket Komplit", "per pasangan", "Rp325rb", "Rp500rb", "-", "3 aktivitas 2,5 jam, basis 20 pasangan"),
    ("Teman Kecil Art Therapy Only", "per pasangan", "Rp250rb", "Rp375rb", "-", "1 aktivitas 90-100 menit"),
    ("Jasa event untuk mitra (12 peserta)", "per event", "Rp2 jt", "-", "-", "Psikolog, fasilitasi, kit, MC, dokumentasi"),
    ("Juri ahli kompetisi", "per acara", "Rp11,5 jt", "Rp17,5 jt", "Rp15 jt", "Hingga 46 karya, hadir di hari acara"),
    ("Hari konsultan (HR Advisory / evaluasi)", "per hari", "Rp3,1 jt", "Rp4,7 jt", "Rp4,7 jt", "4 jam tatap muka + 4 jam analisis"),
    ("FGD 2 jam online", "per FGD", "Rp2,1 jt", "Rp3,2 jt", "Rp3,2 jt", "Fasilitator + notulen + ringkasan"),
], "Event komunitas bertiket untuk umum: Rp259rb per peserta (harga pasar, basis 12 peserta).")

# 17 Coreitera ---------------------------------------------------------------------------------------
CI = [("Theta Starter / Growth / Pro", "per biro per bulan", "Rp1 / 2,9 / 5,73 jt", "1-3, 4-10, 11+ psikolog"),
      ("Theta biaya transaksi", "per transaksi", "Rp5rb", "Biaya Midtrans diteruskan"),
      ("Synthesia Copilot", "per kursi per bulan", "Rp300rb", "Draf interpretasi AI, psikolog tanda tangan"),
      ("Synthesia laporan bertanda tangan", "per laporan", "Rp12rb", "Belum final"),
      ("Mesin tes white-label", "per peserta", "Rp3rb", "Pintu masuk, harga mendekati biaya pokok"),
      ("Tes mandiri + upgrade psikolog", "per peserta", "Rp25rb", "Pintu masuk untuk HR dan UMKM")]
ci = "".join(f"""<div class="card" style="padding:18px 26px;display:flex;align-items:center;border-radius:16px;margin-bottom:10px">
  <div style="flex:1"><div style="font-size:22px;font-weight:700;color:{NAVY}">{a}</div><div style="font-size:16px;color:{MID}">{b} · {d}</div></div>
  <div style="font-size:26px;font-weight:800;color:{NAVY}">{c}</div></div>""" for a, b, c, d in CI)
P["b18_coreitera"] = page(header("Lini 06 · Coreitera", "Teknologi untuk biro psikologi", "") + f"""
<div class="abs glass" style="left:108px;top:300px;width:700px;height:680px;padding:32px 36px;font-size:22px;line-height:1.5">
  <div class="lab" style="color:{SAND2}">Siapa pembelinya</div><div>Biro psikologi dengan 1 sampai 11+ psikolog. Korporat tidak termasuk.</div>
  <div class="lab" style="color:{SAND2};margin-top:18px">Kapan ditawarkan</div><div>Saat biro kewalahan mengatur jadwal, pembayaran, dan laporan; atau ingin menjual tes dengan mereknya sendiri.</div>
  <div class="lab" style="color:{SAND2};margin-top:18px">Aturan</div><div><b>Coreitera tidak pernah menjual tes langsung ke korporat.</b> Mesin tes dijual mendekati biaya pokok, sedangkan keuntungannya datang dari interpretasi dan langganan.</div>
  <div class="lab" style="color:{SAND2};margin-top:18px">Keberatan</div><div><b style="color:{SAND2}">"Kami pakai Google Form dan WhatsApp saja."</b> Hitung bersama mereka jam admin yang habis tiap minggu. Theta Starter Rp1 jt per bulan biasanya lebih murah daripada satu admin paruh waktu.</div>
</div>
<div class="abs" style="left:850px;top:300px;width:962px">{ci}</div>""", "Lini 06")

# 18 Diskon & persetujuan ------------------------------------------------------------------------------
TL = [(GRN, "Hijau", "Boleh dikirim", "Margin ≥ target segmen dan diskon 0-10%. Diskon sampai 10% cukup disetujui lead sales."),
      (YEL, "Kuning", "Founder / CEO", "Margin di bawah target segmen tapi ≥ 30%, atau diskon 10-20%."),
      (RED, "Merah", "Direktur", "Margin di bawah 30%, diskon di atas 20%, atau bagi hasil sekolah/mitra di atas 10%.")]
tl = "".join(f"""<div class="card" style="width:540px;padding:30px 32px;border-top:14px solid {c}">
  <div style="display:flex;justify-content:space-between;align-items:baseline"><div style="font-size:40px;font-weight:800;color:{c}">{n}</div>
  <div class="pill" style="background:{c};color:#fff">{who}</div></div><div style="margin-top:14px;font-size:23px;line-height:1.5;color:{INK}">{d}</div></div>""" for c, n, who, d in TL)
P["b19_diskon"] = page(header("Aturan", "Diskon dan persetujuan", "Kalkulator_Sales memberi warna otomatis di setiap baris penawaran.") + f'<div class="abs" style="left:108px;top:400px;display:flex;gap:42px">{tl}</div>'
    + f"""<div class="abs glass" style="left:108px;top:720px;width:1704px;padding:28px 36px;font-size:22px;line-height:1.55">
  <b style="color:{SAND2}">Diskon pilot</b> maks. 20%, sekali per klien, dicatat sebagai biaya penjualan, diikat ke kontrak lanjutan.
  <b style="color:{SAND2}">Diskon volume</b> pelatihan 5% (5+ kelompok) atau 10% (10+). <b style="color:{SAND2}">Tidak ada</b> diskon volume MHCU.
  <b style="color:{SAND2}">Pendidikan</b> minimal Rp10 jt per kontrak. <b style="color:{SAND2}">Regulator</b>: jangan menaikkan harga item yang sudah pernah dibeli, tambahkan item baru.
  <b style="color:{SAND2}">Mitra</b> boleh menawarkan hingga 20% di atas harga segmen tanpa persetujuan.</div>""")

# 19 Komisi ------------------------------------------------------------------------------------------
P["b20_komisi"] = page(header("Komisi", "Diskon mengurangi komisi Anda sendiri", "Komisi sales = 10% dari laba kotor. Harga lebih tinggi = margin lebih tinggi = komisi lebih besar.") + f"""
<div class="abs card" style="left:108px;top:400px;width:1100px;height:560px;padding:40px 44px">
  <div class="lab">Contoh dari Kalkulator_Sales</div>
  <div style="margin-top:8px;font-size:30px;font-weight:800;color:{NAVY}">MHCU Flash × 50 orang, korporat, diskon 10%</div>
  <div style="display:flex;gap:24px;margin-top:30px">
    <div style="flex:1;padding:24px;border-radius:18px;background:{SAND2}"><div class="lab" style="color:{NAVY}">Margin bersih</div><div style="font-size:48px;font-weight:800;color:{NAVY}">47,5%</div>
      <div class="pill" style="background:{YEL};color:#fff">Kuning</div></div>
    <div style="flex:1;padding:24px;border-radius:18px;background:{SAND2}"><div class="lab" style="color:{NAVY}">Komisi Anda</div><div style="font-size:48px;font-weight:800;color:{NAVY}">Rp1,42 jt</div></div>
    <div style="flex:1;padding:24px;border-radius:18px;background:{NAVY};color:#fff"><div class="lab" style="color:{SAND2}">Tanpa diskon</div><div style="font-size:48px;font-weight:800">+Rp294rb</div>
      <div style="font-size:19px;opacity:.85">komisi tambahan, dan hijau</div></div>
  </div>
  <div style="margin-top:30px;font-size:23px;line-height:1.5;color:{INK}">Kolom <b>"kalau diskon 10% lebih kecil"</b> di Kalkulator menunjukkan angka ini untuk setiap baris. Pakai sebelum memberi diskon.</div>
</div>
<div class="abs glass" style="left:1240px;top:400px;width:572px;height:560px;padding:34px;font-size:21px;line-height:1.55">
  <div class="lab" style="color:{SAND2}">Komisi mitra</div>
  <div style="margin-top:8px">Non-EAP: 10% profit bersih (deal 1-4), 12,5% mulai deal ke-5, 2,5% perpanjangan. Margin setelah komisi min. 15%.</div>
  <div style="margin-top:14px">EAP: Rp1,5 jt per kontrak tahunan baru, Rp500rb per paket 3 bulan, Rp1 jt per perpanjangan.</div>
  <div style="margin-top:14px">Harga di atas daftar: mitra mendapat 30% dari selisihnya.</div>
</div>""")

# 20 Pembayaran, pajak --------------------------------------------------------------------------------
PAY = [("Pembayaran", "Pendidikan dan komunitas: deposit penuh sebelum pelaksanaan. EAP 3 bulan: DP 50% saat SPK, lunas setelah laporan akhir. Bank dan korporat di atas Rp100 jt: DP 30%. Regulator: ikut ketentuan mereka."),
       ("Pajak", "Teman Journey dan Coreitera bukan PKP. Jangan pernah menulis \"termasuk PPN\". PPh 23 dipotong klien B2B; margin di Excel sudah memperhitungkannya."),
       ("Perjalanan dan venue", "Di luar Jakarta: tiket, hotel, uang harian ditagih sebesar biaya aktual, atau biaya + 10% kalau TJ yang mengatur."),
       ("Satuan harga", "Pelatihan per kelompok per hari; asesmen per peserta; EAP per paket, kuartal, atau bulan; pembicara per 2 jam + per jam tambahan. Tidak pernah satu angka borongan."),
       ("Kerahasiaan data", "Manajemen klien hanya menerima data agregat anonim. Hasil individual hanya untuk pegawai yang bersangkutan."),
       ("Waktu", "Regulator dan bank menyusun anggaran setahun sebelumnya. Kegiatan bulan Maret: penawaran masuk Okt-Jan sebelumnya.")]
pc = "".join(f'<div class="card" style="width:540px;height:270px;padding:26px 30px"><div style="font-size:28px;font-weight:800;color:{NAVY}">{a}</div>'
             f'<div style="margin-top:10px;font-size:20px;line-height:1.5;color:{INK}">{b}</div></div>' for a, b in PAY)
P["b21_pembayaran"] = page(header("Aturan", "Pembayaran, pajak, perjalanan", "") + f'<div class="abs" style="left:108px;top:300px;width:1704px;display:flex;flex-wrap:wrap;gap:30px 42px">{pc}</div>')


# 21-22 Excel -------------------------------------------------------------------------------------------
def sheet_mock(cols):
    head = "".join(f'<div style="flex:{w};padding:12px 14px;background:{NAVY};color:#fff;font-size:17px;font-weight:700;border-right:1px solid rgba(255,255,255,.2)">{c}</div>' for c, w, _ in cols)
    row = "".join(f'<div style="flex:{w};padding:12px 14px;background:{bg};color:{INK};font-size:17px;border-right:1px solid #e5dcc8">{"&nbsp;"}</div>' for c, w, bg in cols)
    return f'<div style="display:flex;border-radius:12px 12px 0 0;overflow:hidden">{head}</div><div style="display:flex">{row}</div><div style="display:flex">{row}</div>'


P["b22_excel_kalkulator"] = page(header("Cara membaca Excel · 1", "Kalkulator_Sales.xlsx: sheet Penawaran",
                                        "Satu-satunya file yang dipakai sales untuk menyusun harga. Biaya produk tersembunyi dan terkunci.") + f"""
<div class="abs card" style="left:108px;top:380px;width:1704px;padding:26px 30px">
  {sheet_mock([("Segmen", 2, "#dbe8f5"), ("Kode", 2, "#dbe8f5"), ("Produk", 4, "#fff"), ("Jumlah", 1.4, "#dbe8f5"), ("Diskon", 1.4, "#dbe8f5"), ("Harga usulan", 2, "#dbe8f5"),
               ("Harga daftar", 2, "#fff"), ("Total", 2, "#fff"), ("Margin", 1.4, "#fff"), ("Status", 2.4, "#fff"), ("Komisi", 2, "#fff"), ("Diskon −10%?", 2.2, "#fff")])}
  <div style="margin-top:14px;font-size:19px;color:{MID}"><span class="pill" style="background:#dbe8f5;color:{NAVY}">Biru muda</span> = boleh diisi. Putih = rumus, jangan ditimpa.</div>
</div>
<div class="abs" style="left:108px;top:700px;width:1704px;display:flex;gap:30px">
  {"".join(f'<div class="glass" style="flex:1;padding:24px 26px;font-size:20px;line-height:1.5"><div style="font-size:40px;font-weight:800;color:{SAND2}">{n}</div>{t}</div>' for n, t in [
      ("1", "Isi nama klien, tanggal, dan bagi hasil sekolah/mitra di atas (kalau ada)."),
      ("2", "Per baris: segmen, kode produk (daftar di sheet <b>Harga</b>), jumlah."),
      ("3", "Diskon <b>atau</b> harga usulan. Kosongkan keduanya untuk harga daftar."),
      ("4", "Baca <b>Status</b> per baris dan ringkasan di bawah: siapa yang harus menyetujui."),
      ("5", "Kolom <b>Diskon −10%?</b>: tambahan komisi kalau diskon dikurangi 10 poin.")])}
</div>""", "Excel")

P["b23_excel_master"] = page(header("Cara membaca Excel · 2", "Master_Produk_dan_Harga.xlsx",
                                    "File tim harga (berisi biaya). Sales cukup tahu cara membaca Katalog; jangan dikirim ke luar.") + f"""
<div class="abs" style="left:108px;top:380px;width:1704px;display:flex;flex-wrap:wrap;gap:24px">
  {"".join(f'<div class="card" style="width:404px;height:250px;padding:24px 26px"><div class="pill" style="background:{NAVY};color:{SAND2}">{c}</div>'
           f'<div style="margin-top:12px;font-size:25px;font-weight:800;color:{NAVY}">{t}</div><div style="margin-top:6px;font-size:19px;line-height:1.45;color:{MID}">{d}</div></div>' for c, t, d in [
      ("Katalog A-H", "Identitas produk", "Kode, lini, nama, isi, pembeli, satuan, status (baru atau lama)."),
      ("Katalog I-J", "Peran dan aturan", "Pintu masuk / inti / retensi; aturan harga: Penuh, Pasar, Langganan, Harga pokok."),
      ("Katalog T, V, X", "Harga per segmen", "Pendidikan, Korporat, Regulator. Kolom U, W, Y di sebelahnya = margin bersihnya."),
      ("Katalog AA", "Cek margin", "\"Lolos\" berarti aman. Selain itu margin di bawah 30% dan perlu persetujuan Direktur."),
      ("Simulasi_Penawaran", "Simulasi tim harga", "Versi lengkap Kalkulator, dengan biaya. Cek Rp10 jt pendidikan dan bagi hasil."),
      ("Aturan_Penjualan", "21 aturan", "Sumber resmi semua aturan di deck ini. Kalau ada perbedaan, yang dipegang adalah sheet ini."),
      ("Rincian &amp; Tarif_Biaya", "Biaya", "Rahasia dan hanya diubah oleh tim harga. Sales tidak perlu membukanya."),
      ("Ringkasan", "Gambaran cepat", "Jumlah produk per lini dan margin rata-rata per segmen.")])}
</div>""", "Excel")

# 23 Belum boleh dijanjikan ------------------------------------------------------------------------------
OPEN = [("Harga baru yang biayanya belum final", "Konseling kelompok, in-house psychologist, workshop (bahan), Teman Kecil, paket tes sekolah, asesmen manajerial/eksekutif, hari konsultan. Konfirmasi ke tim harga sebelum mengirim."),
        ("EAP Korporat", "Harga Rp13,5 jt/bulan untuk 300 pegawai masih asumsi pemakaian. Untuk jumlah pegawai lain, minta hitungan ulang."),
        ("Synthesia fee laporan", "Angka Rp12rb per laporan belum ditetapkan, jadi jangan dikutip ke biro."),
        ("Kontrak lama", "Sibolga, JAMPURNAS, Telering: cek status kontrak sebelum menawarkan perpanjangan."),
        ("Diagnosis dan krisis", "Kita tidak menjanjikan diagnosis klinis atau layanan darurat 24 jam. Kasus berat dirujuk."),
        ("Acuan SBM 2027", "Angka pagu SBM 2027 belum dikonfirmasi. Pakai harga regulator di Kalkulator.")]
oc = "".join(f'<div class="card" style="width:540px;height:262px;padding:24px 28px;border-left:10px solid {RED}"><div style="font-size:25px;font-weight:800;color:{NAVY}">{a}</div>'
             f'<div style="margin-top:10px;font-size:19px;line-height:1.5;color:{INK}">{b}</div></div>' for a, b in OPEN)
P["b24_belum"] = page(header("Hati-hati", "Yang belum boleh dijanjikan", "") + f'<div class="abs" style="left:108px;top:300px;width:1704px;display:flex;flex-wrap:wrap;gap:30px 42px">{oc}</div>')

# render ------------------------------------------------------------------------------------------------
for name, html in P.items():
    hp = os.path.join(SRC, name + ".html")
    open(hp, "w").write(html)
    png = os.path.join(OUT, name + ".png")
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                    "--window-size=1920,1080", "--virtual-time-budget=4000", f"--screenshot={png}", "file://" + hp], check=True, capture_output=True)
print("rendered", len(P), "pages")
