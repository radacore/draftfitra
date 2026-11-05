#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sync_images.py - Sinkronkan gambar AI hasil generasi manual ke figures/.
Peta nama file di folder proyek -> nama file yang dipakai build_proposal.py.
Jalankan SEBELUM `python3 build_proposal.py` agar gambar manual tidak tertimpa
oleh gambar programatik lama di figures/.

Usage:
    python3 sync_images.py
"""

import os
import shutil

BASE = os.path.dirname(os.path.abspath(__file__))
FIGDIR = os.path.join(BASE, "figures")

MAPPING = {
    "alur kerja smote.png": "gambar1_alur_smote.png",
    "arsitektur random forest.png": "gambar2_arsitektur_rf.png",
    "kerangka pikir.png": "gambar5_kerangka_berpikir.png",
    "desain metode penelitian.png": "gambar3_desain_metode.png",
    "flowchart gambar 9.png": "gambar4_flowchart_metode.png",
    "bab4_login.png": "gambar10_login.png",
    "bab4_dashboard_admin.png": "gambar11_dashboard_admin.png",
    "bab4_dashboard_nakes.png": "gambar12_dashboard_nakes.png",
    "bab4_form_klasifikasi.png": "gambar13_form_klasifikasi.png",
    "bab4_hasil_klasifikasi.png": "gambar14_hasil_klasifikasi.png",
    "bab4_riwayat.png": "gambar15_riwayat.png",
    "bab4_evaluasi_model.png": "gambar16_evaluasi_model.png",
    "bab4_pengujian_smote.png": "gambar17_pengujian_smote.png",
    "bab4_skenario_pengujian.png": "gambar18_skenario.png",
    "bab4_data_rekam_medis.png": "gambar19_data_medis.png",
    "bab4_pembersihan_data.png": "gambar20_pembersihan.png",
    "bab4_manajemen_akun.png": "gambar21_manajemen_akun.png",
}


def main():
    os.makedirs(FIGDIR, exist_ok=True)
    for src, dst in MAPPING.items():
        src_path = os.path.join(BASE, src)
        dst_path = os.path.join(FIGDIR, dst)
        if not os.path.exists(src_path):
            print(f"[SKIP] tidak ditemukan: {src}")
            continue
        shutil.copy2(src_path, dst_path)
        print(f"[OK] {src} -> figures/{dst}")


if __name__ == "__main__":
    main()
