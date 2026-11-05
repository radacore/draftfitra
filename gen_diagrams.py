#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_diagrams.py — Membuat 4 diagram penelitian (PNG 300 dpi) untuk proposal:
  Gambar 1: Alur Kerja SMOTE
  Gambar 2: Arsitektur Random Forest
  Gambar 3: Desain Metode Penelitian
  Gambar 4: Flowchart Metode SMOTE dan Random Forest
"""

import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "figures")
os.makedirs(OUT, exist_ok=True)

EDGE = "#2b4d7f"
FILL = "#ffffff"
ARROW = "#2b4d7f"
FONT = "DejaVu Sans"


def add_box(ax, x, y, w, h, text, fs=10.5, fc="#ffffff", ec=EDGE, lw=1.4, round_=True):
    style = "round,pad=0.02,rounding_size=0.05" if round_ else "round,pad=0.02,rounding_size=0.02"
    b = FancyBboxPatch((x - w / 2.0, y - h / 2.0), w, h,
                       boxstyle=style, linewidth=lw, edgecolor=ec, facecolor=fc)
    ax.add_patch(b)
    ax.text(x, y, text, ha="center", va="center", fontsize=fs, fontfamily=FONT)


def arrow(ax, x1, y1, x2, y2):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", lw=1.4, color=ARROW))


def vertical_flow(steps, filename, start_end=None, fig_w=9.5, fs=10.5, box_w=6.4):
    n = len(steps)
    if start_end is None:
        start_end = {0, n - 1}
    unit = 1.35
    ylim_top = n * unit
    fig, ax = plt.subplots(figsize=(fig_w, ylim_top / 1.4))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, ylim_top + 0.5)
    ax.axis("off")
    box_h = 0.95
    for i, s in enumerate(steps):
        y = ylim_top - 0.6 - i * unit
        if i in start_end:
            add_box(ax, 5, y, box_w - 0.6, box_h, s, fs=fs)
        else:
            add_box(ax, 5, y, box_w, box_h, s, fs=fs)
        if i < n - 1:
            y_next = y - unit
            arrow(ax, 5, y - box_h / 2.0, 5, y_next + box_h / 2.0)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, filename), dpi=300, bbox_inches="tight",
                facecolor="white")
    plt.close(fig)
    print("saved", filename)


def gambar1():
    steps = [
        "Data Tidak Seimbang",
        "Seleksi Sampel Kelas Minoritas",
        "Perhitungan k Tetangga Terdekat (k-NN)",
        "Interpolasi Sampel Sintetis\nX_new = X_i + \u03bb (X_zi \u2212 X_i)",
        "Dataset Seimbang",
        "Pelatihan Model Klasifikasi",
    ]
    vertical_flow(steps, "gambar1_alur_smote.png", fs=10.5)


def gambar3():
    steps = [
        "Mulai",
        "Pengumpulan Data Rekam Medis MDR-TB",
        "Preprocessing Data",
        "Pembagian Data Latih dan Data Uji",
        "Penerapan SMOTE",
        "Pelatihan Model Random Forest",
        "Validasi Model",
        "Evaluasi Model",
        "Implementasi dan Simulasi",
        "Selesai",
    ]
    vertical_flow(steps, "gambar3_desain_metode.png", fs=10.5)


def gambar4():
    steps = [
        "Mulai",
        "Input Data Rekam Medis",
        "Preprocessing Data",
        "Pembagian Data Latih dan Data Uji",
        "Penerapan SMOTE pada Data Latih",
        "Inisialisasi Parameter Random Forest",
        "Pelatihan Model Random Forest",
        "Pengujian Model",
        "Evaluasi Model",
        "Selesai",
    ]
    vertical_flow(steps, "gambar4_flowchart_metode.png", fs=10.5)


def gambar5():
    steps = [
        "Masalah",
        "Tujuan",
        "Data",
        "Solusi",
        "Metode",
        "Hasil",
    ]
    vertical_flow(steps, "gambar5_kerangka_berpikir.png",
                  start_end=set(), fs=12, box_w=5.2, fig_w=8.5)


def gambar2():
    fig, ax = plt.subplots(figsize=(10.5, 7.5))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis("off")

    # Dataset
    add_box(ax, 5, 9.1, 3.4, 0.9, "Dataset Latih", fs=11)
    # Bootstrap
    add_box(ax, 5, 7.6, 3.6, 0.9, "Bootstrap Sampling (Bagging)", fs=11)
    # Trees
    add_box(ax, 2.0, 5.9, 1.9, 1.0, "Pohon\nKeputusan 1", fs=10)
    add_box(ax, 5.0, 5.9, 1.9, 1.0, "Pohon\nKeputusan 2", fs=10)
    add_box(ax, 8.0, 5.9, 1.9, 1.0, "Pohon\nKeputusan n", fs=10)
    # Voting
    add_box(ax, 5, 4.0, 3.6, 0.9, "Voting (Mayoritas Suara)", fs=11)
    # Hasil
    add_box(ax, 5, 2.5, 3.0, 0.9, "Hasil Klasifikasi", fs=11)

    arrow(ax, 5, 8.65, 5, 8.05)
    arrow(ax, 5, 7.15, 2.0, 6.4)
    arrow(ax, 5, 7.15, 5.0, 6.4)
    arrow(ax, 5, 7.15, 8.0, 6.4)
    arrow(ax, 2.0, 5.4, 4.0, 4.45)
    arrow(ax, 5.0, 5.4, 5.0, 4.45)
    arrow(ax, 8.0, 5.4, 6.0, 4.45)
    arrow(ax, 5, 3.55, 5, 2.95)

    fig.tight_layout()
    fig.savefig(os.path.join(OUT, "gambar2_arsitektur_rf.png"), dpi=300,
                bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("saved gambar2_arsitektur_rf.png")


if __name__ == "__main__":
    gambar1()
    gambar2()
    gambar3()
    gambar4()
    gambar5()
