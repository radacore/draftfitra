# Prompt Gambar untuk Proposal Skripsi

Proposal ini menggunakan **5 gambar programatik** yang sudah digenerate secara otomatis
(`gen_diagrams.py` → `figures/gambar1..5.png`). Berikut prompt AI-image untuk membuat
versi visual yang lebih estetis (rekomendasi **Ideogram** karena paling akurat dalam
merender teks berbahasa Indonesia; alternatif: Midjourney/Flux dengan penulisan manual
ulang teksnya setelah generasi).

> Catatan: Gambar programatik (matplotlib) sudah memenuhi syarat untuk proposal.
> Prompt berikut bersifat opsional, hanya jika ingin mengganti dengan versi AI-image.
> Prompt yang sama persis juga disertakan di dalam dokumen (bagian LAMPIRAN —
> PROMPT GENERASI GAMBAR).

---

## Gaya Visual Konsisten (Digunakan untuk Semua Gambar)

```
Clean flat vector diagram, white background, navy blue outlines and arrows
(#2B4D7F), rounded-rectangle boxes, capsule-shaped start and end boxes, thin
navy blue arrows, all labels in Indonesian, clean sans-serif font (Inter, Segoe
UI, or Poppins), dark gray text (#222222), vertical top-down layout, no
shadows, no gradients, no 3D effects, no watermark, portrait aspect 4:5.
```

## Gambar 1 — Alur Kerja SMOTE

```
Using the exact style described above, create a flowchart titled "Alur Kerja
SMOTE". Arrange 6 boxes vertically from top to bottom, connected by downward
arrows:
1. "Data Tidak Seimbang"
2. "Seleksi Sampel Kelas Minoritas"
3. "Perhitungan k Tetangga Terdekat (k-NN)"
4. "Interpolasi Sampel Sintetis" (box contains formula: X_new = X_i + λ (X_zi − X_i))
5. "Dataset Seimbang"
6. "Pelatihan Model Klasifikasi"
white background, navy blue outline (#2b4d7f), thin arrows, rounded rectangles,
minimalist academic style, no watermark, high contrast, text clearly readable
```

## Gambar 2 — Arsitektur Random Forest

```
Clean flat vector architecture diagram, Indonesian language text, top to bottom:
- top box: "Dataset Latih"
- middle box: "Bootstrap Sampling (Bagging)"
- three tree icons in a row labeled "Pohon Keputusan 1", "Pohon Keputusan 2",
  "Pohon Keputusan n"
- box: "Voting (Mayoritas Suara)"
- bottom box: "Hasil Klasifikasi"
arrows from bootstrap to each tree, then each tree to voting, then voting to result,
white background, navy blue outline, minimalist academic style, no watermark
```

## Gambar 3 — Kerangka Berpikir

```
Using the exact style described above, create a wide horizontal (landscape)
diagram titled "Kerangka Berpikir" for a study conducted at RSUD Labuang Baji,
Sulawesi Selatan. Arrange six wide rounded-rectangle boxes stacked vertically
from top to bottom, each spanning nearly the full image width, connected by
short downward arrows. Each box has a bold Indonesian label with smaller
supporting Indonesian text that must fit exactly two lines, as follows in order:
(1) "Masalah" — "Tingginya kegagalan pengobatan MDR-TB dan sulitnya deteksi
    dini karena data rekam medis RSUD Labuang Baji yang tidak seimbang.";
(2) "Tujuan" — "Mengembangkan model klasifikasi yang memprediksi keberhasilan
    pengobatan pasien MDR-TB secara akurat, khususnya kelas minoritas yang
    berisiko gagal.";
(3) "Data" — "Data rekam medis pasien MDR-TB RSUD Labuang Baji Sulawesi
    Selatan mencakup demografis, klinis, dan riwayat pengobatan.";
(4) "Solusi" — "Penerapan SMOTE untuk menyeimbangkan dataset dan Random Forest
    sebagai model klasifikasi.";
(5) "Metode" — "Data dipreprocessing, diterapkan SMOTE, dibagi data latih dan
    uji, lalu dilatih Random Forest dan dievaluasi.";
(6) "Hasil" — "Model akurat dan sensitif terhadap pasien berisiko gagal,
    berpotensi menjadi sistem pendukung keputusan tenaga medis."
Each supporting text must wrap to no more than two lines inside its wide box;
keep the boxes compact in height. Use the exact colors, shapes, and fonts from
the style block.
```

## Gambar 4 — Desain Metode Penelitian

```
Using the exact style described above, create a flowchart titled "Desain Metode
Penelitian". Arrange 10 boxes vertically from top to bottom, connected by
downward arrows. "Mulai" is the capsule-shaped start box and "Selesai" is the
capsule-shaped end box; the others are rounded rectangles. Each box must
contain exactly this Indonesian text, in order:
1. "Mulai"
2. "Pengumpulan Data Rekam Medis MDR-TB"
3. "Preprocessing Data"
4. "Pembagian Data Latih dan Data Uji"
5. "Penerapan SMOTE"
6. "Pelatihan Model Random Forest"
7. "Validasi Model"
8. "Evaluasi Model"
9. "Implementasi dan Simulasi"
10. "Selesai"
```

## Gambar 5 — Flowchart Metode SMOTE dan Random Forest

```
Using the exact style described above, create a flowchart titled "Flowchart
Metode SMOTE dan Random Forest". Arrange 10 boxes vertically from top to
bottom, connected by downward arrows. "Mulai" is the capsule-shaped start box
and "Selesai" is the capsule-shaped end box; the others are rounded rectangles.
Each box must contain exactly this Indonesian text, in order:
1. "Mulai"
2. "Input Data Rekam Medis"
3. "Preprocessing Data"
4. "Pembagian Data Latih dan Data Uji"
5. "Penerapan SMOTE pada Data Latih"
6. "Inisialisasi Parameter Random Forest"
7. "Pelatihan Model Random Forest"
8. "Pengujian Model"
9. "Evaluasi Model"
10. "Selesai"
```

---

## Cara Penggunaan

1. Generate gambar dengan prompt di atas (Ideogram paling disarankan).
2. Simpan hasilnya sebagai `figures/gambar1..5.png` (timpa file programatik
   atau ganti nama agar tidak konflik).
3. Jalankan ulang: `python3 build_proposal.py` lalu konversi ke PDF.

Jika hasil AI-image kurang pas secara akademik, gunakan kembali versi programatik
yang sudah ada di `figures/`.
