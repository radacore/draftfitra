#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_pagemap.py — mengekstrak nomor halaman riil dari PDF hasil render,
lalu menulis pagemap.json agar build_proposal.py bisa mengisi Daftar Isi,
Daftar Tabel, dan Daftar Gambar dengan nomor halaman yang benar.
Front matter: romawi = pdf_page - 1 (cover 1 halaman). Body: arab = pdf_page - p_bab1 + 1.
Deteksi heading memakai pencocokan "baris berdiri sendiri" agar tidak
terjebak oleh entri yang muncul di Daftar Isi.
"""

import json
import os
import re
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
OUTPAGEMAP = os.path.join(BASE, "pagemap.json")


def to_roman(n):
    vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
            (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
            (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def main():
    if len(sys.argv) < 2:
        print("usage: gen_pagemap.py <rendered.pdf>")
        sys.exit(1)
    pdf = sys.argv[1]
    txt = pdf.replace(".pdf", ".txt")
    subprocess.run(["pdftotext", "-layout", pdf, txt], check=True)
    with open(txt, encoding="utf-8") as f:
        raw = f.read()
    pages = raw.split("\f")

    def standalone_pages(key):
        hits = []
        for i, pg in enumerate(pages, start=1):
            for line in pg.splitlines():
                if line.strip() == key:
                    hits.append(i)
                    break
        return hits

    def first_standalone(key):
        h = standalone_pages(key)
        return h[0] if h else None

    def contain_pages(key):
        return [i + 1 for i, pg in enumerate(pages) if key in pg]

    def last_contain(key):
        h = contain_pages(key)
        return h[-1] if h else None

    # p_bab1: halaman body pertama (arab 1). Deteksi lewat baris "PENDAHULUAN"
    # yang berdiri sendiri pada halaman BAB I.
    p_bab1 = first_standalone("PENDAHULUAN")
    if p_bab1 is None:
        p_bab1 = last_contain("PENDAHULUAN")

    def display_front(pdf_page):
        if pdf_page is None:
            return "?"
        n = pdf_page - 1
        return to_roman(n) if n >= 1 else "?"

    def display_body(pdf_page):
        if pdf_page is None:
            return "?"
        return str(pdf_page - p_bab1 + 1)

    front_keys = {
        "KATA PENGANTAR": "KATA PENGANTAR",
        "DAFTAR ISI": "DAFTAR ISI",
        "DAFTAR TABEL": "DAFTAR TABEL",
        "DAFTAR GAMBAR": "DAFTAR GAMBAR",
    }
    body_keys = {
        "BAB I PENDAHULUAN": "PENDAHULUAN",
        "A. Latar Belakang": "A. Latar Belakang",
        "B. Rumusan Masalah": "B. Rumusan Masalah",
        "C. Tujuan Penelitian": "C. Tujuan Penelitian",
        "D. Manfaat Penelitian": "D. Manfaat Penelitian",
        "E. Ruang Lingkup Penelitian": "E. Ruang Lingkup Penelitian",
        "F. Sistematika Penulisan": "F. Sistematika Penulisan",
        "BAB II TINJAUAN PUSTAKA": "TINJAUAN PUSTAKA",
        "A. Landasan Teori": "A. Landasan Teori",
        "B. Penelitian Terkait": "B. Penelitian Terkait",
        "C. Kerangka Berpikir": "C. Kerangka Berpikir",
        "BAB III METODE PENELITIAN": "METODE PENELITIAN",
        "A. Tempat dan Waktu Penelitian": "A. Tempat dan Waktu Penelitian",
        "B. Alat dan Bahan": "B. Alat dan Bahan",
        "C. Perancangan Sistem": "C. Perancangan Sistem",
        "BAB IV HASIL DAN PEMBAHASAN": "HASIL DAN PEMBAHASAN",
        "A. Hasil Penerapan User Centered Design": "A. Hasil Penerapan User Centered Design",
        "B. Hasil Implementasi Sistem Akhir": "B. Hasil Implementasi Sistem Akhir",
        "C. Hasil Pengujian Black Box": "C. Hasil Pengujian Black Box",
        "D. Hasil Validasi System Usability Scale": "D. Hasil Validasi System Usability Scale",
        "F. Pembahasan": "F. Pembahasan",
        "BAB V PENUTUP": "PENUTUP",
        "A. KESIMPULAN": "A. KESIMPULAN",
        "B. SARAN": "B. SARAN",
        "DAFTAR PUSTAKA": "DAFTAR PUSTAKA",
    }
    tabel_keys = {
        "Tabel 1": "Simbol Flowchart",
        "Tabel 2": "Penelitian Terkait",
        "Tabel 3": "Jadwal Penelitian",
        "Tabel 4": "Hyperparameter Terpilih",
        "Tabel 5": "Confusion Matrix Data Uji",
        "Tabel 6": "Metrik Evaluasi Data Uji",
        "Tabel 7": "Hasil Cross Validation Lima Lipatan",
        "Tabel 8": "Perbandingan Model Tanpa dan Dengan SMOTE",
        "Tabel 9": "Skenario Variasi Proporsi Data",
        "Tabel 10": "Skenario Variasi Jumlah Pohon",
        "Tabel 11": "Hasil Pengujian Black Box",
        "Tabel 12": "Pernyataan Kuesioner SUS",
        "Tabel 13": "Rekapitulasi Jawaban SUS",
        "Tabel 14": "Rekapitulasi Penilaian TAM",
    }
    gambar_keys = {
        "Gambar 1": "Alur Kerja SMOTE",
        "Gambar 2": "Arsitektur Random Forest",
        "Gambar 3": "Kerangka Berpikir",
        "Gambar 4": "Lokasi RSUD Labuang Baji",
        "Gambar 5": "Desain Metode Penelitian",
        "Gambar 6": "Use Case Diagram Klasifikasi Keberhasilan Pengobatan",
        "Gambar 7": "Activity Diagram Klasifikasi Keberhasilan Pengobatan",
        "Gambar 8": "Class Diagram Klasifikasi Keberhasilan Pengobatan",
        "Gambar 9": "Flowchart Metode SMOTE dan Random Forest",
        "Gambar 10": "Halaman Login Sistem",
        "Gambar 11": "Dashboard Admin",
        "Gambar 12": "Dashboard Tenaga Medis",
        "Gambar 13": "Formulir Klasifikasi",
        "Gambar 14": "Hasil Klasifikasi dan Rekomendasi",
        "Gambar 15": "Riwayat Klasifikasi",
        "Gambar 16": "Evaluasi Model",
        "Gambar 17": "Pengujian Efektivitas SMOTE",
        "Gambar 18": "Skenario Pengujian",
        "Gambar 19": "Data Rekam Medis",
        "Gambar 20": "Pembersihan Data",
        "Gambar 21": "Manajemen Akun",
    }

    def resolve(keys, standalone=False):
        out = {}
        for disp, key in keys.items():
            pdfp = first_standalone(key) if standalone else last_contain(key)
            out[disp] = display_body(pdfp)
        return out

    def resolve_caption(keys):
        out = {}
        for disp, title in keys.items():
            num = disp.split()[-1]
            pat = re.compile(r"Gambar\s+" + num + r"\s*\.\s*" +
                             re.escape(title).replace(r"\ ", r"\s+"))
            hits = [i for i, pg in enumerate(pages, start=1) if pat.search(pg)]
            out[disp] = display_body(hits[0] if hits else None)
        return out

    result = {
        "front": {disp: display_front(first_standalone(key)) for disp, key in front_keys.items()},
        "body": resolve(body_keys, standalone=True),
        "tabel": resolve(tabel_keys),
        "gambar": resolve_caption(gambar_keys),
        "meta": {"p_bab1": p_bab1, "pages": len(pages)},
    }

    with open(OUTPAGEMAP, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print("pagemap.json written")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
