# Kalkulasi biaya layanan baru

Tanggal: 20 Nov 2026. **Status: disetujui Galih 20 Nov 2026 (tarif T11-T18 dan semua asumsi volume).** Setelah disetujui, tarif dan rincian ini masuk ke
`build/build_master_sheet.py` (sheet `Tarif_Biaya` dan `Rincian_Biaya`).

## 1. Cara menghitung

Satu rumus untuk semua layanan, sama dengan RAB yang sudah dipakai (BI Fit and Proper, BSN, BI Sibolga):

**Biaya dasar = (jam orang x tarif per jam) + (orang per acara x honor) + (bahan per peserta x jumlah peserta) + biaya tetap per kontrak.**

Biaya per acara dibagi ke jumlah peserta atau pasangan supaya keluar biaya per satuan jual.
Harga per segmen = biaya dasar x 0,9 / (0,882 - target margin), dibulatkan ke atas (Pendidikan 30%, Korporat 50%).
Kalau sekolah mendapat bagi hasil 10%, rumusnya menjadi biaya x 0,9 / (0,882 - 0,09 - target).

## 2. Tarif yang dipakai, dan sumbernya

| # | Tarif | Nilai | Sumber di repo | Sudah di `Tarif_Biaya`? |
|---|---|---|---|---|
| T1 | Psikolog tatap muka (konseling, wawancara, fasilitasi, termasuk online) | Rp350.000 / jam | RAB BI Fit and Proper, RAB BSN, Lampiran B Kemitraan | Ya (`PSI_F2F`, `PSI_EAP`) |
| T2 | Psikolog kerja dokumen (skoring, laporan, analisis) | Rp110.000 / jam | RAB BI Fit and Proper (Galih 28 Okt 2026), RAB BSN | Ya (`PSI_DOC`) |
| T3 | Admin dan platform per sesi | Rp35.000 / sesi | Lampiran B, BI Sibolga | Ya (`ADM_SESI`) |
| T4 | Admin / koordinator proyek | Rp50.000 / jam | RAB BSN | Ya (`ADM_JAM`) |
| T5 | Transport psikolog Jakarta | Rp150.000 / orang-hari | RAB BI Fit and Proper, RAB BSN | Ya (`TRANSPORT`) |
| T6 | Kontrak, meterai, invoice | Rp150.000 / kontrak | Lampiran B | Ya (`KONTRAK`) |
| T7 | Platform + skoring Coreitera per peserta | Rp75.000 | Model 2027 v15 (`core_fee`, estimasi) | Ya (`CORE_FEE`) |
| T8 | Alat ukur MHCU | Rp7.000 / orang | Galih 19 Nov 2026 | Ya (`MHCU_TOOL`) |
| T9 | Laporan agregat anonim | Rp750.000 | Lampiran B | Ya (`AGG_REPORT`) |
| T10 | Tes BFI + DISC | Rp120.000 / peserta | RAB BI Fit and Proper | Ya (`BFI_DISC`) |
| **T11** | **Fee psikolog pembicara, per kegiatan 2 jam: Tier 1 / 2 / 3 / 4** | **Rp1,5 / 2,5 / 3 / 4 jt**; tambahan per jam = setengahnya | Canva "INTERNAL - Ratecard & Profit Sharing Psikolog Pembicara" (Okt 2026) | **Belum** |
| **T12** | **Co-fasilitator / MC / admin acara** | **Rp300.000 / acara** | RAB Lyfe at 30s (Paket 2 dan 3) | **Belum** |
| **T13** | **Dokumentasi acara** | **Rp150.000 / acara** | RAB Lyfe at 30s | **Belum** |
| **T14** | **Psikolog event komunitas (sekitar 2 jam)** | **Rp500.000 / acara** (ganti Rp150rb) | RAB Lyfe at 30s Paket 1-3 (Okt 2026) | Ya, tetapi Rp150rb; diganti |
| **T15** | **Produksi webinar oleh TJ (platform 500 orang, desain dan e-sertifikat, host)** | **Rp2.300.000 / webinar** | Costing BI Sibolga v4 (Berbagi) | **Belum** |
| **T16** | **Worksheet / lembar kerja cetak** | **Rp5.500 / peserta** | Budget OSH "Friends of Friend" | **Belum** |
| **T17** | **Kit aktivitas per peserta (dipakai juga sebagai pengganti kit seni dan kit Teman Kecil)** | **Rp30.000 / peserta** | `EVENT_KIT`, RAB Lyfe at 30s | Ya, dipakai untuk produk baru juga |
| **T18** | **Materi ringkas dan poster** | **Rp200.000 / sesi** | Costing BI Sibolga v3 | **Belum** |

Tarif tebal (T11-T18) adalah yang perlu disetujui untuk masuk `Tarif_Biaya`.

## 3. Asumsi volume per layanan (perlu disetujui)

| Layanan | Asumsi | Kenapa |
|---|---|---|
| EAP korporat | 300 pegawai; **10 sesi konseling per bulan** dalam satu pool; MHCU 2x setahun untuk semua pegawai; 1 webinar per kuartal; 4 jam admin per bulan | 10 sesi per bulan = asumsi model 2027 (`eap_sessions`). Ini angka paling berisiko: kalau pemakaian dua kali lipat, biaya sesi naik Rp3,85 jt per bulan |
| Paket tes sekolah | 50 siswa, 2 kelas; tiap kelas diawasi 1 psikolog 2,5 jam + 1 asisten; skor dan laporan otomatis di Coreitera; review psikolog 15 menit per siswa | Pola klasikal Asisi (min. 50 siswa) |
| Asesmen perkembangan anak | 1 asesor psikolog per 3 anak, 90 menit; 12 anak per sesi | MoM Deen Haus 21 Sep 2026 |
| Konseling kelompok | 90 menit, 5-8 orang, 1 psikolog | Proposal TFI |
| In-house psychologist | 1 hari = 6 jam layanan + 2 jam catatan; setengah hari = 3 + 1 jam | Proposal sekolah: 2-3 hari per minggu |
| Teman Kecil | 20 pasangan; 1 psikolog utama, 2 co-fasilitator, 1 koordinator, dokumentasi | Proposal JLL |
| Asesmen manajerial / eksekutif | Disusun dari blok Fit and Proper per peserta | RAB BI Fit and Proper |
| Hari konsultan | 4 jam tatap muka + 4 jam analisis | Pola RAB BSN |

## 4. Ringkasan hasil

| Lini | Layanan | Satuan | Biaya dasar | Pendidikan 30% | Korporat 50% | Pembanding |
|---|---|---|---|---|---|---|
| 3 | Pembicara Tier 1 / 2 / 3 / 4, 2 jam | per kegiatan | Rp1,65 / 2,65 / 3,15 / 4,15 jt | Rp2,6 / 4,1 / 4,9 / 6,5 jt | Rp3,9 / 6,3 / 7,5 / 9,8 jt | Ratecard kampus Rp3 / 4,5 / 7 / 8 jt |
| 3 | Webinar diselenggarakan TJ (Tier 2) | per webinar | Rp4,95 jt | Rp7,7 jt | Rp12 jt | Katalog lama Rp2,6 jt (tanpa produksi) |
| 3 | Workshop korporat offline, Tier 2, 30 peserta | per sesi | Rp3,42 jt | Rp5,3 jt | Rp8,1 jt | - |
| 3 | Workshop sekolah offline, Tier 1, 40 siswa | per sesi | Rp2,47 jt | Rp3,9 jt | - | Kampus membayar Rp1-3 jt |
| 3 | Art-Based Wellbeing, Tier 2, 30 peserta | per sesi | Rp4,15 jt | Rp6,5 jt | Rp9,8 jt | - |
| 3 | Team-Building 4 jam, Tier 2, 30 peserta | per sesi | Rp7,1 jt | Rp11 jt | Rp17 jt | - |
| 2 | Konseling kelompok 90 menit | per sesi | Rp615rb | Rp975rb | Rp1,5 jt | TFI tidak ada harga terpisah |
| 2 | In-house psychologist | per hari | Rp2,47 jt | Rp3,9 jt | Rp5,9 jt | - |
| 2 | In-house psychologist, setengah hari | per setengah hari | Rp1,31 jt | Rp2,1 jt | Rp3,1 jt | - |
| 2 | **EAP korporat, 300 pegawai** | **per bulan** | **Rp5,6 jt** | - | **Rp13,5 jt** | **Rencana 2027: Rp15 jt (biaya Rp5 jt)** |
| 1 | Paket tes sekolah Dasar | per siswa | Rp152,5rb | Rp250rb (Rp300rb dengan bagi hasil 10%) | - | Asisi Rp250rb; pasar Rp50-250rb |
| 1 | Paket tes sekolah Standar (+ rekap sekolah) | per siswa | Rp165,7rb | Rp275rb | - | - |
| 1 | Paket tes sekolah Lengkap (+ konseling hasil ke orang tua) | per siswa | Rp375,7rb | Rp600rb | - | Biro Lestari Rp350rb |
| 1 | Asesmen perkembangan anak, kelompok | per anak | Rp308,75rb | Rp500rb | - | **Deen Haus Rp250rb: rugi** |
| 1 | Asesmen manajerial | per peserta | Rp835rb | - | Rp2 jt | Pasar Rp0,5-2,5 jt |
| 1 | Asesmen eksekutif | per peserta | Rp3,22 jt | - | Rp7,6 jt | Pasar Rp5-20 jt+ |
| 4 | Teman Kecil Paket Komplit | per pasangan | Rp204,75rb | Rp325rb | Rp500rb | JLL Rp500-550rb |
| 4 | Teman Kecil Art Therapy Only | per pasangan | Rp154,1rb | Rp250rb | Rp375rb | JLL Rp400-450rb |
| 5 | Hari konsultan (HR Advisory, evaluasi) | per hari | Rp1,99 jt | Rp3,1 jt | Rp4,7 jt | - |
| 5 | FGD online 2 jam | per FGD | Rp1,33 jt | Rp2,1 jt | Rp3,2 jt | - |
| 4 | Event komunitas, 12 peserta (psikolog Rp500rb) | per peserta | Rp146,7rb | - | - | Tiket Rp259rb: margin turun dari 55% ke sekitar 43% |

## 5. Yang terlihat

1. **EAP korporat Rp13,5 jt per bulan di target 50% hampir sama dengan rencana 2027 (Rp15 jt).** Biaya hasil hitung (Rp5,6 jt) juga dekat dengan angka model (Rp5 jt). Rencana 2027 tetap masuk akal, asal pemakaian tetap sekitar 10 sesi per bulan per 300 pegawai.
2. **Paket tes sekolah Rp250rb per siswa tepat di batas 30%.** Harga Asisi benar. Dengan bagi hasil sekolah 10%, harga harus Rp300rb.
3. **Teman Kecil untuk JLL (Rp550rb per pasangan) sudah setara harga Korporat.** Margin sekitar 55%.
4. **Asesmen anak Deen Haus Rp250rb per anak rugi** kalau asesornya psikolog dengan tarif Rp350rb per jam. Harga ini hanya masuk kalau asesornya asisten terlatih (bukan psikolog), atau satu asesor memegang lebih dari 3 anak.
5. **Workshop sekolah Rp3,9 jt per sesi masih di atas anggaran kampus (Rp1-3 jt).** Sesuai keputusan, workshop pendidikan dijual di dalam paket minimal Rp10 jt.
6. **Psikolog event Rp500rb** menurunkan margin event bertiket dari 55% ke sekitar 43%. Masih sehat untuk aturan 'Pasar'.

## 6. Yang belum bisa dihitung dari repo

| Biaya | Untuk | Sementara |
|---|---|---|
| Lisensi atau alat tes IQ per siswa (kalau tes kertas berlisensi) | Paket tes sekolah | Dianggap masuk Rp75rb platform Coreitera |
| Kit ASQ-3 per anak | Asesmen anak | Belum dihitung |
| Harga nyata kit seni dan kit rumah Teman Kecil | Art-Based, Teman Kecil | Memakai Rp30rb per peserta |
| Honor pakar eksternal (misalnya panel BI DPU) | Evaluasi program | Pakai SBM sekitar Rp1,7 jt per jam sebagai acuan, belum dimasukkan |
| Data pemakaian EAP nyata | EAP korporat | Memakai 10 sesi per bulan dari model 2027 |

## 7. Rincian per layanan

**Pembicara Tier 1, 2 jam, online di platform klien** (per kegiatan)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee psikolog pembicara | 1 | Rp1.500.000 | Rp1.500.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp1.650.000** |

Harga: Pendidikan (30%) **Rp2.600.000**, Korporat (50%) **Rp3.900.000**.

**Pembicara Tier 2, 2 jam, online di platform klien** (per kegiatan)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee psikolog pembicara | 1 | Rp2.500.000 | Rp2.500.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp2.650.000** |

Harga: Pendidikan (30%) **Rp4.100.000**, Korporat (50%) **Rp6.300.000**.

**Pembicara Tier 3, 2 jam, online di platform klien** (per kegiatan)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee psikolog pembicara | 1 | Rp3.000.000 | Rp3.000.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp3.150.000** |

Harga: Pendidikan (30%) **Rp4.900.000**, Korporat (50%) **Rp7.500.000**.

**Pembicara Tier 4, 2 jam, online di platform klien** (per kegiatan)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee psikolog pembicara | 1 | Rp4.000.000 | Rp4.000.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp4.150.000** |

Harga: Pendidikan (30%) **Rp6.500.000**, Korporat (50%) **Rp9.800.000**.

**Webinar Tier 2 diselenggarakan TJ (platform, desain, host)** (per webinar)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee pembicara T2 | 1 | Rp2.500.000 | Rp2.500.000 |
| Produksi webinar | 1 | Rp2.300.000 | Rp2.300.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp4.950.000** |

Harga: Pendidikan (30%) **Rp7.700.000**, Korporat (50%) **Rp12.000.000**.

**Workshop korporat offline Jakarta, Tier 2, 2 jam, 30 peserta** (per sesi)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee pembicara T2 | 1 | Rp2.500.000 | Rp2.500.000 |
| Co-fasilitator | 1 | Rp300.000 | Rp300.000 |
| Transport 2 orang | 2 | Rp150.000 | Rp300.000 |
| Worksheet | 30 | Rp5.500 | Rp165.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp3.415.000** |

Harga: Pendidikan (30%) **Rp5.300.000**, Korporat (50%) **Rp8.100.000**.

**Workshop sekolah offline, Tier 1, 2 jam, 40 siswa** (per sesi)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee pembicara T1 | 1 | Rp1.500.000 | Rp1.500.000 |
| Co-fasilitator | 1 | Rp300.000 | Rp300.000 |
| Transport 2 orang | 2 | Rp150.000 | Rp300.000 |
| Worksheet | 40 | Rp5.500 | Rp220.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp2.470.000** |

Harga: Pendidikan (30%) **Rp3.900.000**, Korporat (50%) **Rp5.900.000**.

**Art-Based Wellbeing offline, Tier 2, 2 jam, 30 peserta** (per sesi)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee pembicara T2 | 1 | Rp2.500.000 | Rp2.500.000 |
| Co-fasilitator | 1 | Rp300.000 | Rp300.000 |
| Transport 2 orang | 2 | Rp150.000 | Rp300.000 |
| Kit seni (proxy EVENT_KIT) | 30 | Rp30.000 | Rp900.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp4.150.000** |

Harga: Pendidikan (30%) **Rp6.500.000**, Korporat (50%) **Rp9.800.000**.

**Team-Building 4 jam, Tier 2, 30 peserta** (per sesi)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fee pembicara T2 (2 jam) | 1 | Rp2.500.000 | Rp2.500.000 |
| Tambahan 2 jam T2 | 2 | Rp1.250.000 | Rp2.500.000 |
| Co-fasilitator | 2 | Rp300.000 | Rp600.000 |
| Transport 3 orang | 3 | Rp150.000 | Rp450.000 |
| Kit permainan (proxy) | 30 | Rp30.000 | Rp900.000 |
| Kontrak | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp7.100.000** |

Harga: Pendidikan (30%) **Rp11.000.000**, Korporat (50%) **Rp17.000.000**.

**Konseling kelompok 90 menit, 5-8 orang** (per sesi)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Psikolog tatap muka | 1.5 | Rp350.000 | Rp525.000 |
| Catatan sesi | 0.5 | Rp110.000 | Rp55.000 |
| Admin & platform | 1 | Rp35.000 | Rp35.000 |
| **Biaya dasar** | | | **Rp615.000** |

Harga: Pendidikan (30%) **Rp975.000**, Korporat (50%) **Rp1.500.000**.

**In-house psychologist, 1 hari di lokasi** (per hari)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Psikolog tatap muka (konseling/office hours) | 6 | Rp350.000 | Rp2.100.000 |
| Catatan & rujukan | 2 | Rp110.000 | Rp220.000 |
| Transport | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp2.470.000** |

Harga: Pendidikan (30%) **Rp3.900.000**, Korporat (50%) **Rp5.900.000**.

**In-house psychologist, setengah hari (sekolah)** (per setengah hari)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Psikolog tatap muka | 3 | Rp350.000 | Rp1.050.000 |
| Catatan | 1 | Rp110.000 | Rp110.000 |
| Transport | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp1.310.000** |

Harga: Pendidikan (30%) **Rp2.100.000**, Korporat (50%) **Rp3.100.000**.

**EAP korporat 300 pegawai, per bulan** (per bulan)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Sesi konseling (pool, model v15: 10/bln) | 10 | Rp350.000 | Rp3.500.000 |
| Admin & platform per sesi | 10 | Rp35.000 | Rp350.000 |
| MHCU 2x/thn semua pegawai (alat) | 50 | Rp7.000 | Rp350.000 |
| Laporan agregat 2x/thn | 0.166667 | Rp750.000 | Rp125.000 |
| Webinar per kuartal (T2 + host + materi) | 0.333333 | Rp3.200.000 | Rp1.066.667 |
| Account management admin | 4 | Rp50.000 | Rp200.000 |
| Kontrak | 0.0833333 | Rp150.000 | Rp12.500 |
| **Biaya dasar** | | | **Rp5.604.167** |

Harga: Pendidikan (30%) **Rp8.700.000**, Korporat (50%) **Rp13.500.000**.

**Paket tes sekolah DASAR (IQ + minat + gaya belajar), per siswa, 50 siswa** (per siswa)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Administrasi klasikal: psikolog 2,5 jam x 2 kelas / 50 | 0.1 | Rp350.000 | Rp35.000 |
| Asisten pengawas 2 kelas / 50 | 0.04 | Rp300.000 | Rp12.000 |
| Platform + skoring Coreitera | 1 | Rp75.000 | Rp75.000 |
| Review & laporan psikolog | 0.25 | Rp110.000 | Rp27.500 |
| Kontrak / 50 | 0.02 | Rp150.000 | Rp3.000 |
| **Biaya dasar** | | | **Rp152.500** |

Harga: Pendidikan (30%) **Rp250.000**, Korporat (50%) **Rp375.000**.

**  ...sama, dengan bagi hasil sekolah 10%** (per siswa)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| biaya Dasar | 1 | Rp152.500 | Rp152.500 |
| **Biaya dasar** | | | **Rp152.500** |

Harga: Pendidikan (30%) **Rp300.000**, Korporat (50%) **Rp375.000**.

**Paket tes sekolah STANDAR (+ rekap & rekomendasi sekolah)** (per siswa)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| biaya Dasar | 1 | Rp152.500 | Rp152.500 |
| Rekap & rekomendasi sekolah 6 jam / 50 | 0.12 | Rp110.000 | Rp13.200 |
| **Biaya dasar** | | | **Rp165.700** |

Harga: Pendidikan (30%) **Rp275.000**, Korporat (50%) **Rp400.000**.

**Paket tes sekolah LENGKAP (+ konseling hasil ke orang tua 30 mnt)** (per siswa)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| biaya Dasar | 1 | Rp152.500 | Rp152.500 |
| Rekap 6 jam / 50 | 0.12 | Rp110.000 | Rp13.200 |
| Konseling hasil 30 mnt | 0.5 | Rp350.000 | Rp175.000 |
| Admin per sesi | 1 | Rp35.000 | Rp35.000 |
| **Biaya dasar** | | | **Rp375.700** |

Harga: Pendidikan (30%) **Rp600.000**, Korporat (50%) **Rp900.000**.

**Asesmen perkembangan anak kelompok (1 asesor/3 anak, 90 mnt)** (per anak) (alat ASQ-3 belum dihitung)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Asesor psikolog 1,5 jam / 3 anak | 0.5 | Rp350.000 | Rp175.000 |
| PIC psikolog pembuka 1,5 jam / 12 anak | 0.125 | Rp350.000 | Rp43.750 |
| Laporan & rekomendasi | 0.5 | Rp110.000 | Rp55.000 |
| Admin | 1 | Rp35.000 | Rp35.000 |
| **Biaya dasar** | | | **Rp308.750** |

Harga: Pendidikan (30%) **Rp500.000**, Korporat (50%) **Rp750.000**.

**Asesmen manajerial per peserta** (per peserta)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Tes BFI + DISC | 1 | Rp120.000 | Rp120.000 |
| Wawancara psikolog | 1 | Rp350.000 | Rp350.000 |
| Skoring wawancara | 0.5 | Rp110.000 | Rp55.000 |
| Laporan | 2 | Rp110.000 | Rp220.000 |
| Review mutu | 0.5 | Rp110.000 | Rp55.000 |
| Admin | 1 | Rp35.000 | Rp35.000 |
| **Biaya dasar** | | | **Rp835.000** |

Harga: Pendidikan (30%) **Rp1.300.000**, Korporat (50%) **Rp2.000.000**.

**Asesmen eksekutif per peserta** (per peserta)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Tes BFI + DISC | 1 | Rp120.000 | Rp120.000 |
| Wawancara mendalam | 2 | Rp350.000 | Rp700.000 |
| Simulasi bisnis: 2 psikolog x 2 jam | 4 | Rp350.000 | Rp1.400.000 |
| Skoring | 2 | Rp110.000 | Rp220.000 |
| Laporan senior | 3 | Rp110.000 | Rp330.000 |
| Review mutu | 1 | Rp110.000 | Rp110.000 |
| Transport 2 orang | 2 | Rp150.000 | Rp300.000 |
| Admin | 1 | Rp35.000 | Rp35.000 |
| **Biaya dasar** | | | **Rp3.215.000** |

Harga: Pendidikan (30%) **Rp5.000.000**, Korporat (50%) **Rp7.600.000**.

**Teman Kecil Paket Komplit, 2,5 jam, 20 pasangan** (per pasangan)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Psikolog utama 2,5 jam | 0.125 | Rp350.000 | Rp43.750 |
| Co-fasilitator x2 | 0.1 | Rp300.000 | Rp30.000 |
| Koordinator / MC | 0.05 | Rp300.000 | Rp15.000 |
| Dokumentasi | 0.05 | Rp150.000 | Rp7.500 |
| Transport 4 orang | 0.2 | Rp150.000 | Rp30.000 |
| Dokumen refleksi 2 jam | 0.1 | Rp110.000 | Rp11.000 |
| Kit rumah + Lentera (proxy 2 x EVENT_KIT) | 2 | Rp30.000 | Rp60.000 |
| Kontrak | 0.05 | Rp150.000 | Rp7.500 |
| **Biaya dasar** | | | **Rp204.750** |

Harga: Pendidikan (30%) **Rp325.000**, Korporat (50%) **Rp500.000**.

**Teman Kecil Art Therapy Only, 100 menit, 20 pasangan** (per pasangan)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Psikolog utama 1,75 jam | 0.0875 | Rp350.000 | Rp30.625 |
| Co-fasilitator x2 | 0.1 | Rp300.000 | Rp30.000 |
| Koordinator / MC | 0.05 | Rp300.000 | Rp15.000 |
| Transport 4 orang | 0.2 | Rp150.000 | Rp30.000 |
| Dokumen refleksi 2 jam | 0.1 | Rp110.000 | Rp11.000 |
| Kit rumah (proxy EVENT_KIT) | 1 | Rp30.000 | Rp30.000 |
| Kontrak | 0.05 | Rp150.000 | Rp7.500 |
| **Biaya dasar** | | | **Rp154.125** |

Harga: Pendidikan (30%) **Rp250.000**, Korporat (50%) **Rp375.000**.

**Hari konsultan (HR Advisory / evaluasi program)** (per hari)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Psikolog senior tatap muka | 4 | Rp350.000 | Rp1.400.000 |
| Analisis & penulisan | 4 | Rp110.000 | Rp440.000 |
| Transport | 1 | Rp150.000 | Rp150.000 |
| **Biaya dasar** | | | **Rp1.990.000** |

Harga: Pendidikan (30%) **Rp3.100.000**, Korporat (50%) **Rp4.700.000**.

**FGD 2 jam online (fasilitator + notulen)** (per FGD)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Fasilitator psikolog | 2 | Rp350.000 | Rp700.000 |
| Notulen / co-fasilitator | 1 | Rp300.000 | Rp300.000 |
| Olah transkrip & ringkasan | 3 | Rp110.000 | Rp330.000 |
| **Biaya dasar** | | | **Rp1.330.000** |

Harga: Pendidikan (30%) **Rp2.100.000**, Korporat (50%) **Rp3.200.000**.

**Event komunitas 12 peserta (psikolog Rp500rb RAB Lyfe)** (per peserta)

| Komponen | Jumlah | Tarif | Total |
|---|---|---|---|
| Psikolog event | 0.0833333 | Rp500.000 | Rp41.667 |
| Kit | 1 | Rp30.000 | Rp30.000 |
| Makan-minum | 1 | Rp50.000 | Rp50.000 |
| Iklan | 0.0833333 | Rp300.000 | Rp25.000 |
| **Biaya dasar** | | | **Rp146.667** |

Harga: Pendidikan (30%) **Rp250.000**, Korporat (50%) **Rp350.000**.
