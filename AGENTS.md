# AGENTS.md — Proyek Fitra (Proposal Skripsi MDR-TB)

## Gambar (PENTING — JANGAN TERTIMPA)

- `build_proposal.py` men-embed gambar dari folder `figures/`.
- Gambar AI yang diganti user secara manual biasanya ditaruh di folder proyek
  (contoh: `alur kerja smote.png`, `kerangka pikir.png`) atau langsung di DOCX.
- **Sebelum menjalankan ulang `build_proposal.py`, WAJIB sync gambar manual ke
  `figures/` terlebih dahulu** dengan:
  ```
  python3 sync_images.py
  ```
  Skrip ini menyalin PNG di root proyek ke `figures/` dengan nama yang dipakai
  build. Tanpa sync, rebuild akan menimpa gambar manual dengan gambar
  programatik matplotlib lama di `figures/`.
- Jika gambar diganti langsung di DOCX (Word), ekstrak media dari DOCX lalu
  peta ke `figures/` sesuai `sync_images.py` sebelum rebuild, atau minta user
  menyimpan gambar di root proyek.

## Alur Build

Dua pass (nomor halaman Daftar Isi/Tabel/Gambar diambil dari PDF hasil render):
1. `python3 build_proposal.py` -> DOCX
2. `soffice --headless --convert-to pdf --outdir . Proposal_Skripsi_Zulfitra_MDR-TB.docx`
3. `python3 gen_pagemap.py Proposal_Skripsi_Zulfitra_MDR-TB.pdf` -> pagemap.json
4. `python3 build_proposal.py` (pass 2, isi nomor halaman riil)
5. `soffice --headless --convert-to pdf ...` (render final)
6. Verifikasi: `pdftotext`, `pdfimages -list`, `python3 office-skripsi/scripts/validate-docx.py`

## Regresi yang Dicek Setiap Rebuild

- 0 karakter `❑` (kotak dari OMML nary/rad/acc kosong; `repair_nary()` mencegahnya).
- `pdfimages -list`: 5 gambar; media DOCX di `word/media/` berukuran ±1 MB jika memakai
  gambar AI (di PDF ter-render menjadi JPEG, DPI bisa < 300 tergantung rasio gambar).
- Nomor halaman gambar/tabel di pagemap konsisten; DAFTAR PUSTAKA utuh.
- `validate-docx.py` 3/4 PASS ("KESIMPULAN MISSING" bukan bagian proposal).
