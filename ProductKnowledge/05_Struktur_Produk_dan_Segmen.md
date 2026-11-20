# Struktur layanan, biaya dasar, dan target margin per segmen

Tanggal: 20 Nov 2026. **Status: disetujui Galih (struktur dan peran).** Harga di bawah adalah hasil rumus; harga final
masuk ke `build_master_sheet.py` di Fase 2. Latar belakang: `06_Strategi_ke_Produk_dan_Harga.md` (strategi) dan
`07_Benchmark_Kompetitor.md` (pasar).

## 1. Cara membaca

- **Biaya dasar** = biaya per satuan dari `Tarif_Biaya` dan `Rincian_Biaya` di `Master_Produk_dan_Harga.xlsx`. Biaya tidak berubah per klien.
- **Harga per segmen** = biaya x 0,9 / (0,882 - target margin), dibulatkan ke atas. Rumus ini sama dengan Katalog: margin bersih setelah PPh 23 2% dan komisi 10% dari laba kotor.
- **Pengaman:** untuk Korporat dan Regulator, harga yang sekarang sudah lebih tinggi **tidak diturunkan**. Untuk Regulator, item yang sudah pernah dibeli **tidak dinaikkan**.

## 2. Target margin per segmen

| Segmen | Siapa | Target margin bersih | Pengali biaya | Aturan |
|---|---|---|---|---|
| **Regulator** | BI, OJK, LPS | Setinggi mungkin, menuju 50%; minimal 30% | sampai x 2,36 | Dijual sebagai program dan paket, mengacu ke pagu SBM. Item lama dibekukan |
| **Korporat dan pemerintah** | Bank, BUMN, swasta, kementerian, pemda | **50%** | x 2,36 | Harga sekarang yang lebih tinggi tetap dipakai |
| **Pendidikan dan komunitas** | Sekolah, kampus, yayasan, komunitas | **30%** (batas bawah) | x 1,55 | Versi lebih ringan; minimal Rp10 jt per kontrak; bagi hasil sekolah maks. 10% |
| **Individu** | Klien perorangan, pasangan | Harga pasar (margin sekarang 24-34%) | - | Bagi hasil tier psikolog 60/65/65/70% |
| **Biro psikologi** | Pembeli Coreitera | Harga langganan (margin sekarang 54-87%) | - | Coreitera tidak menjual tes langsung ke korporat |

## 3. Peran produk dan aturan harganya

| Peran | Artinya | Aturan harga |
|---|---|---|
| **Pintu masuk** | Hampir tanpa biaya antar: tes mandiri, laporan otomatis, snapshot burnout untuk HR, webinar rekaman, konten | **Gratis atau di bawah pasar** (di sini kita boleh perang harga) |
| Pintu masuk dengan sedikit waktu psikolog | Debrief 1 jam, webinar langsung | **Harga pasar** |
| **Inti** | Semua yang memakai waktu psikolog per orang atau per acara | **Harga segmen, tidak pernah gratis** |
| **Retensi** | Kontrak berulang | Harga segmen, dalam paket |

Pengaman pintu masuk: (1) selalu diikuti proposal produk inti; (2) subsidi maksimal Rp6 jt per perusahaan (= biaya mendapatkan 1 klien B2B), dicatat sebagai biaya penjualan; (3) konversi dilacak, dipotong kalau di bawah 5% dalam 6 bulan; (4) tidak gratis untuk regulator dan bank. Untuk mereka, pintu masuknya pilot berbayar.

## 4. Kelompok layanan, biaya dasar, dan harga per segmen

Harga dalam Rupiah. "rb" = ribu, "jt" = juta. "Harga sekarang" = harga daftar di `01_Daftar_Produk_Final.md`.
**Tanda bintang (\*)** = biaya dari ratecard internal Okt 2026, belum masuk `Tarif_Biaya`, perlu persetujuan.

### Teman Journey

#### Lini 1. Asesmen: "Kenali orang-orang Anda"

| Layanan | Peran | Satuan | Biaya dasar | Pendidikan (30%) | Korporat (50%) | Harga sekarang (margin) |
|---|---|---|---|---|---|---|
| MHCU L0: tes mandiri + snapshot HR | Pintu masuk | per orang | sekitar Rp3rb | gratis | gratis | - |
| MHCU L1: Skrining, review psikolog | Inti | per orang | Rp69,5rb | Rp125rb | Rp175rb | (Mr DIY Rp250rb) |
| MHCU L2: Ringkas, konseling 15 mnt | Inti (pendidikan) | per orang | Rp157rb | Rp250rb | - | (Telering Rp242rb) |
| MHCU L3: Flash, konseling 30 mnt | Inti | per orang | Rp244,5rb | Rp400rb | Rp600rb | Rp500rb (44%) |
| MHCU L4: Deep, konseling 60 mnt | Inti | per orang | Rp419,5rb | Rp650rb | Rp1 jt | Rp750rb (38%) |
| Hasil MHCU individual (add-on EAP) | Retensi | per pegawai | Rp34,5rb | Rp54rb | Rp82rb | Rp70rb (44%) |
| Seleksi / Fit and Proper | Inti | per peserta | Rp1,46 jt | - | Rp3,5 jt | Rp2,5 jt (36%) |
| Asesmen manajerial / eksekutif *(baru)* | Inti | per peserta | belum ada | - | - | - |
| Profiling massal Paket 1 / 2 | Inti | proyek 2.000 orang | Rp6,89 jt | - | Rp16,5 jt | Rp50 / 70 jt (76-79%), tetap |
| Pemetaan 3A / 3B / 3C | Inti | proyek 2.000 orang | Rp19,2 / 34 / 61,9 jt | - | Rp45,5 / 80,5 / 146 jt | Rp101,5 / 139,5 / 210,5 jt (62-71%), tetap |
| Tambahan pelaksanaan serentak | Inti | per proyek | Rp9,35 jt | - | Rp22,5 jt | Rp30 jt (60%), tetap |
| Asesmen + laporan dampak, peserta tambahan | Inti | per peserta | Rp86rb | - | Rp225rb | Rp150rb (37%) |
| Paket tes sekolah Dasar / Standar / Lengkap *(baru)* | Inti (pendidikan) | per siswa | belum ada (fee asesor, alat tes) | - | - | (Asisi IQ Rp250rb) |
| Tes individu (IQ, minat bakat, dll.) | Inti (individu) | per tes | 50% harga | - | - | Rp375rb-1,62 jt (48-50%), harga pasar |

#### Lini 2. Konseling dan pendampingan: "Teman perjalanan untuk orang-orang Anda"

| Layanan | Peran | Satuan | Biaya dasar | Pendidikan (30%) | Korporat (50%) | Harga sekarang (margin) |
|---|---|---|---|---|---|---|
| EAP korporat *(baru)*: retainer per pegawai, kuota sesi, MHCU berkala | Retensi | per bulan | belum ada (perlu data pemakaian) | - | target sekitar Rp15 jt/bulan untuk 150-300+ pegawai | - |
| EAP Jeda | Retensi | per 3 bulan | Rp7,43 jt | Rp11,5 jt | Rp17,5 jt | Rp12 jt (33%) |
| EAP Jeda Tahunan | Retensi | per kuartal | Rp8,67 jt | Rp13,5 jt | Rp20,5 jt | Rp11 jt (17%) |
| EAP Tumbuh | Retensi | per 3 bulan | Rp11,7 jt | Rp18,5 jt | Rp28 jt | Rp20 jt (36%) |
| EAP Seperjalanan | Retensi | per 3 bulan | Rp17,93 jt | Rp28 jt | Rp42,5 jt | Rp31,5 jt (37%) |
| Sesi B2B 60 menit | Inti | per sesi | Rp385rb | Rp600rb | Rp925rb | Rp500rb (19%) |
| Sesi B2B 30 menit *(baru)* | Inti | per sesi | Rp210rb | Rp325rb | Rp500rb | - |
| Paket 10 sesi 60 mnt + 10 MHCU + laporan *(baru)* | Inti | per paket | Rp4,35 jt | Rp6,8 jt | Rp10,5 jt | (TFI Rp5,05 jt) |
| Paket 10 sesi 30 mnt + 10 MHCU + laporan *(baru)* | Inti | per paket | Rp2,6 jt | Rp4,1 jt | Rp6,2 jt | (TFI Rp3,3 jt) |
| Paket 20 sesi 60 mnt + 58 MHCU, 18 bulan *(baru)* | Retensi | per kontrak | Rp9,85 jt | Rp15,5 jt | Rp23,5 jt | (TFI Rp19,5 jt) |
| Konseling kelompok *(baru)* | Inti | per sesi | belum ada | - | - | - |
| In-house psychologist *(baru)* | Retensi | per hari | belum ada | - | - | - |
| Konseling individu, pasangan, pre-marriage, terapi | Inti (individu) | per sesi / paket | bagi hasil tier + Rp50rb platform | - | - | Rp500rb-3,8 jt (24-34%), harga pasar |

#### Lini 3. Pembelajaran dan psikoedukasi: "Bantu orang-orang Anda belajar dan berubah"

| Layanan | Peran | Satuan | Biaya dasar | Pendidikan (30%) | Korporat (50%) | Harga sekarang (margin) |
|---|---|---|---|---|---|---|
| Program Pelatih Kantor Pusat (Flagship) | Inti | per program | Rp28 jt | - | Rp66 jt | Rp100 jt (63%), regulator, dibekukan |
| Program Pelatih Regional | Inti | per hari | Rp10,25 jt | - | Rp24,5 jt | Rp28,5 jt (56%), regulator, dibekukan |
| Bank Essential *(tambah ke Katalog)* | Inti | per hari | Rp12 jt | - | Rp28,5 jt | Rate_Card: BUMN Rp45 jt / swasta Rp55 jt (64-69%), tetap |
| Bank Signature | Inti | per hari | Rp15 jt | - | Rp35,5 jt | BUMN Rp51 jt / swasta Rp61 jt (62-66%), tetap |
| Bank Elite | Retensi | per hari | Rp18,1 jt | - | Rp43 jt | BUMN Rp66,3 jt / swasta Rp79,3 jt (64-68%), tetap |
| Pembicara / webinar langsung Tier 1 | Inti | per 2 jam | Rp1,65 jt\* | Rp2,6 jt | Rp3,9 jt | Internal Rp2,8 jt; kampus Rp3 jt |
| Pembicara Tier 2 | Inti | per 2 jam | Rp2,65 jt\* | Rp4,1 jt | Rp6,3 jt | Rp4,5 jt |
| Pembicara Tier 3 | Inti | per 2 jam | Rp3,15 jt\* | Rp4,9 jt | Rp7,5 jt | Internal Rp6 jt; kampus Rp7 jt |
| Pembicara Tier 4 | Inti | per 2 jam | Rp4,15 jt\* | Rp6,5 jt | Rp9,8 jt | Rp8 jt |
| Tambahan durasi Tier 1 sampai Tier 4 | Inti | per jam | Rp0,75-2 jt\* | Rp1,2-3,1 jt | Rp1,8-4,8 jt | Rp1,4-4 jt per jam |
| Workshop korporat (HWE, Art-Based, Team-Building) *(baru)* | Inti | per sesi | tarif pembicara + bahan (bahan belum ada) | - | - | - |
| Workshop sekolah dan kampus (Career 101, Gadget, Sex-Ed, UpStander, Parenting) *(baru)* | Inti, dalam paket | per sesi | tarif pembicara + bahan (bahan belum ada) | - | - | - |
| Webinar rekaman, konten | Pintu masuk | - | sekitar nol | gratis | gratis | - |

Webinar Rp2,6 jt dan seminar Rp4,9 jt di Katalog lama diganti oleh ratecard pembicara per tier.

#### Lini 4. Pengalaman dan komunitas: "Momen bermakna untuk komunitas Anda"

| Layanan | Peran | Satuan | Biaya dasar | Pendidikan (30%) | Korporat (50%) | Harga sekarang (margin) |
|---|---|---|---|---|---|---|
| Event psikoedukasi publik (sponsor / CSR) | Inti | per event | Rp8,15 jt | Rp13 jt | Rp19,5 jt | Rp24 jt (58%), tetap |
| Event komunitas bertiket | Inti (individu) | per peserta | Rp117,5rb (psikolog Rp150rb, perlu dicek) | - | - | Rp259rb (55%), harga pasar |
| Teman Kecil: Komplit / Art Only *(baru)* | Inti | per pasangan | belum ada | - | - | (JLL Rp400-550rb) |
| Jasa event untuk mitra *(baru)* | Inti | per event | RAB Lyfe: Rp1,25 jt (12 peserta) | - | - | (RAB Rp2,5-3,5 jt; Canva Rp4,5-6 jt) |

#### Lini 5. Konsultasi organisasi: "Sistem dan program yang lebih baik"

| Layanan | Peran | Satuan | Biaya dasar | Pendidikan (30%) | Korporat (50%) | Harga sekarang (margin) |
|---|---|---|---|---|---|---|
| Juri ahli | Inti | per acara | Rp7,4 jt | Rp11,5 jt | Rp17,5 jt | Rp15 jt (44%) |
| HR Advisory *(baru)* | Inti | per hari / proyek | belum ada | - | - | - |
| Evaluasi program dan FGD *(baru)* | Inti | per proyek | belum ada | - | - | - |

### Coreitera

#### Lini 6. Teknologi (segmen Biro psikologi)

| Layanan | Peran | Satuan | Biaya dasar | Harga sekarang (margin) |
|---|---|---|---|---|
| Tes mandiri + upgrade "ditandatangani psikolog" | Pintu masuk | per peserta | Rp2,75rb | Rp25rb (89%) |
| Mesin tes white-label | Pintu masuk | per peserta | Rp2,75rb | Rp3rb (sekitar biaya) |
| Theta Starter / Growth / Pro | Inti, retensi | per biro per bulan | Rp175rb | Rp1 / 2,9 / 5,73 jt (74-87%) |
| Theta biaya transaksi | Inti | per transaksi | Rp0 | Rp5rb |
| Synthesia kursi / fee laporan | Inti (2028) | per kursi / per laporan | Rp120rb / Rp8rb | Rp300rb (54%) / Rp12rb (30%) |

## 5. Yang terlihat dari tabel ini

1. **Pelatihan, profiling, dan pemetaan sudah jauh di atas target 50%** (56-79%). Harga tetap; inilah mesin laba.
2. **Yang di bawah target Korporat:** MHCU Flash (44%) dan Deep (38%), Fit and Proper (36%), EAP (17-37%), sesi B2B (19%), juri (44%), asesmen tambahan (37%). Untuk klien korporat baru, harga naik ke kolom Korporat. Untuk pendidikan, harga turun ke kolom Pendidikan, hanya di dalam paket minimal Rp10 jt.
3. **EAP paket kecil tidak bisa menjadi produk korporat.** Bahkan dengan target 50%, Tumbuh hanya Rp28 jt per 3 bulan (Rp9,3 jt per bulan). Rencana 2027 butuh EAP korporat terpisah.
4. **Satu biaya untuk semua klien, harga berbeda per segmen.** Contoh MHCU Flash: biaya Rp244,5rb; pendidikan Rp400rb, korporat Rp600rb. Kualitas dan metodenya sama.

## 6. Biaya yang belum ada (untuk Fase 2)

| # | Biaya | Untuk layanan |
|---|---|---|
| 1 | Fee psikolog pembicara per tier (ratecard internal Rp1,5 / 2,5 / 3 / 4 jt per 2 jam): disetujui masuk `Tarif_Biaya`? | Pembicara, workshop |
| 2 | Bahan workshop per peserta (worksheet, kit seni) | Workshop |
| 3 | Data pemakaian EAP (pegawai, sesi per tahun) dan kuota sesi per pegawai | EAP korporat |
| 4 | Fee sesi konseling kelompok | Konseling kelompok |
| 5 | Tarif psikolog per hari di lokasi | In-house psychologist |
| 6 | Fee asesor dan biaya alat tes per siswa | Paket tes sekolah |
| 7 | Fee asesor untuk asesmen manajerial dan eksekutif | Seleksi |
| 8 | Psikolog utama, co-fasilitator, koordinator, kit per pasangan | Teman Kecil |
| 9 | Tarif konsultan dan pakar per hari | HR Advisory, evaluasi program |
| 10 | Psikolog event komunitas: Rp150rb atau Rp500rb (RAB Lyfe)? | Event bertiket, jasa event |

## 7. Keputusan yang sudah diambil (20 Nov 2026)

- 6 lini produk, 5 segmen, dan peran pintu masuk / inti / retensi disetujui.
- Pintu masuk boleh gratis atau di bawah pasar **hanya kalau biaya antarnya hampir nol**. MHCU Flash dan Deep adalah produk inti, bukan pintu masuk.
- Target: Korporat 50%, Pendidikan 30% (minimal Rp10 jt per kontrak), Regulator dimaksimalkan dengan acuan SBM.
- EAP korporat dipisah dari paket Jeda / Tumbuh / Seperjalanan.
- Telering dihormati sebagai pengecualian; ke depan masuk produk L2 Ringkas.
- Bagi hasil sekolah maksimal 10%.
- Harga lama TFI dihitung ulang hanya untuk struktur baru (paket 10 dan 20 sesi).
