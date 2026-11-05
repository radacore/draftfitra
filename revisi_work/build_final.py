#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bangun DOCX lengkap (BAB I-V + DAFTAR PUSTAKA) dari master + ekstraksi PDF."""
import copy
import json
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from build_proposal import set_run

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "Proposal_Skripsi_Zulfitra_MDR-TB.docx")
OUT = os.path.join(BASE, "revisi_work", "Proposal_Skripsi_Zulfitra_MDR-TB_final.docx")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

full123 = json.load(open(os.path.join(BASE, "revisi_work", "full123.json")))
# pulihkan subscript yang hilang saat ekstraksi (ₖ pada penjelasan persamaan Gini/entropi)
for _x in full123["blocks"]:
    if _x["type"] != "body":
        continue
    for _r in _x["runs"]:
        _r["t"] = _r["t"].replace("jumlah kelas dan p adalah proporsi", "jumlah kelas dan p\u2096 adalah proporsi")
        _r["t"] = _r["t"].replace("dengan p adalah proporsi sampel", "dengan p\u2096 adalah proporsi sampel")
bab45_runs = json.load(open(os.path.join(BASE, "revisi_work", "bab45_runs.json")))
bab45 = json.load(open(os.path.join(BASE, "revisi_work", "bab45.json")))
refs = json.load(open(os.path.join(BASE, "revisi_work", "refs.json")))
img_dims = json.load(open(os.path.join(BASE, "revisi_work", "img_dims.json")))

# ---- selaraskan teks BAB IV dengan perubahan aplikasi (tanggal awal + status pengobatan) ----
def _patch_bab45(blocks, marker, old=None, new=None, new_full=None):
    for x in blocks:
        if "runs" not in x:
            continue
        if marker not in "".join(r["t"] for r in x["runs"]):
            continue
        if new_full is not None:
            x["runs"][0]["t"] = new_full
            for r in x["runs"][1:]:
                if r["t"] and not r["t"][0].isspace():
                    r["t"] = " " + r["t"]
            return True
        for r in x["runs"]:
            if old in r["t"]:
                r["t"] = r["t"].replace(old, new)
                return True
    return False


_b45 = bab45_runs["blocks"]
# Blok 27 — Formulir Klasifikasi
_patch_bab45(_b45, "Formulir klasifikasi memuat 15 variabel isian", new_full=(
    "Formulir klasifikasi memuat 16 variabel isian yang dikelompokkan menjadi identitas pasien, "
    "data demografis, kondisi fisik, riwayat kesehatan, dan data pengobatan. Lima belas variabel "
    "di antaranya menjadi masukan model, sedangkan Tanggal Awal Pengobatan dipakai untuk menurunkan "
    "durasi dan status pengobatan. Sistem juga menghitung empat informasi turunan secara otomatis, "
    "yaitu Kategori Usia dan IMT yang termasuk variabel model, serta Durasi Pengobatan dan Status "
    "Pengobatan yang bersifat informasi tampilan dan tidak digunakan sebagai variabel model. "
    "Penghitung di bagian bawah tombol menunjukkan jumlah variabel yang telah terisi dari 16 "
    "variabel wajib (Gambar 13)."))
# Blok 30 — Hasil Klasifikasi
_patch_bab45(_b45, "serta variabel paling berpengaruh. Pada contoh",
              old="berpengaruh. Pada contoh",
              new=("berpengaruh. Selain itu, ditampilkan pula kartu Informasi Pengobatan yang memuat "
                   "tanggal awal, durasi, dan status pengobatan sebagai informasi turunan yang tidak "
                   "digunakan sebagai variabel model. Pada contoh"))
# Blok 33 — Riwayat Klasifikasi
_patch_bab45(_b45, "Halaman riwayat mencatat seluruh hasil klasifikasi", new_full=(
    "Halaman riwayat mencatat seluruh hasil klasifikasi yang pernah dilakukan beserta pasien, hasil, "
    "keyakinan, pelaksana, dan waktu, dilengkapi penyaring berdasarkan hasil dan tombol rincian. "
    "Di bawah nama pasien ditampilkan tanggal awal, durasi, dan status pengobatan sebagai informasi "
    "turunan (Gambar 15)."))
# Blok 9 — Kebutuhan pengguna
_patch_bab45(_b45, "penginputan 17 variabel rekam medis dengan perhitungan otomatis",
              old="dengan perhitungan otomatis IMT dan kategori usia;",
              new=("dengan perhitungan otomatis IMT dan kategori usia serta pencatatan tanggal awal "
                   "pengobatan untuk menurunkan durasi dan status pengobatan;"))

REVISION = (" Jumlah rekam medis pasien MDR-TB yang dihimpun sebanyak 175 data, "
            "terdiri atas 121 pasien dengan status pengobatan berhasil dan 54 pasien "
            "dengan status pengobatan tidak berhasil.")

doc = Document(SRC)
body = list(doc.element.body)

# ---- perbarui KATA PENGANTAR (nama orang tua, pembimbing, tanggal) ----
def _replace_in_par(par, old, new):
    for r in par.runs:
        if old in r.text:
            r.text = r.text.replace(old, new)
            return True
    return False

def _set_run_text(par, contains, new_text):
    for r in par.runs:
        if contains in r.text:
            r.text = new_text
            return True
    return False

_replaced = _replace_in_par(doc.paragraphs[26], "[Nama Ayah] dan [Nama Ibu]",
                            "Bapak Endang Wijaya dan Ibu Warti, S.E.")
_set_run_text(doc.paragraphs[31], "[Nama Dosen Pembimbing I]",
              "Ibu Chyquitha Danuputri, S.Kom., M.Kom., selaku Dosen Pembimbing I dan "
              "Ibu Rizki Yusliana Bakti, S.T., M.T., selaku Dosen Pembimbing II, atas "
              "bimbingan, masukan ilmiah, dan koreksi yang sangat membantu penyempurnaan "
              "proposal penelitian ini.")
_set_run_text(doc.paragraphs[32], "Chyquitha Danuputri",
              "Bapak Fahrim Irhamna Rachman., S.Kom., M.T., Selaku Dosen Pembimbing "
              "Akademik Fakultas Teknik Universitas Muhammadiyah Makassar")
_replace_in_par(doc.paragraphs[42], "[tanggal bulan tahun]", "13 November 2026")

# ---- persamaan OMML (deepcopy) ----
EQ_IDX = {1: 152, 2: 155, 3: 166, 4: 168, 5: 171, 6: 179, 7: 181, 8: 183, 9: 185}
eq_elements = {n: copy.deepcopy(doc.paragraphs[i]._p) for n, i in EQ_IDX.items()}

front_anchor = body[74]
final_sectPr = body[404]

# --- rebuild sampul (sesuaikan dengan target PDF: 3 baris judul 14pt + logo) ---
cover_sect = body[23]  # paragraf kosong pembawa sectPr section 0
for el in body[0:23]:
    el.getparent().remove(el)


def cover_line(text, size=12, bold=False):
    p = doc.add_paragraph()
    p._p.getparent().remove(p._p)
    cover_sect.addprevious(p._p)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(text)
    set_run(r, size=size, bold=bold)
    return p


def cover_blank():
    cover_line("", 12)


for _ in range(2):
    cover_blank()
cover_line("PENERAPAN TEKNIK SMOTE DAN RANDOM FOREST", 14, True)
cover_line("DALAM KLASIFIKASI KEBERHASILAN PENGOBATAN", 14, True)
cover_line("PENYAKIT MDR-TB", 14, True)
cover_blank()
cover_blank()
cover_line("SKRIPSI", 14, True)
cover_blank()
cover_line("Diajukan Sebagai Salah Satu Syarat untuk Mendapat Gelar", 12, False)
cover_line("Sarjana Komputer (S.Kom) Program Studi Informatika", 12, False)
# logo
pl = doc.add_paragraph()
pl._p.getparent().remove(pl._p)
cover_sect.addprevious(pl._p)
pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
pl.paragraph_format.line_spacing = 1.0
pl.paragraph_format.space_before = Pt(0)
pl.paragraph_format.space_after = Pt(0)
pr = pl.add_run()
pr.add_picture(os.path.join(BASE, "revisi_work", "pdf_img", "img_16.png"), width=Cm(5.69))
cover_blank()
cover_line("Disusun Oleh :", 12, True)
cover_blank()
cover_line("ZULFITRA", 12, True)
cover_line("105841115022", 12, True)
cover_blank()
cover_blank()
cover_blank()
cover_line("PROGRAM STUDI INFORMATIKA", 12, True)
cover_line("FAKULTAS TEKNIK", 12, True)
cover_line("UNIVERSITAS MUHAMMADIYAH MAKASSAR", 12, True)
cover_line("2026", 12, True)

# hapus front matter lama (43..73) dan body lama (75..403)
for el in [body[i] for i in list(range(43, 74)) + list(range(75, 404))]:
    el.getparent().remove(el)


# ---------------------------------------------------------------- utilitas

def _make_par(anchor):
    p = doc.add_paragraph()
    p._p.getparent().remove(p._p)
    anchor.addprevious(p._p)
    return p


def _runs(p, spans, size=12):
    for s in spans:
        r = p.add_run(s["t"])
        set_run(r, size=size, bold=s.get("b", False), italic=s.get("i", False))
    return p


def _tab_right(p, pos="7824", dots=True):
    tabs = p._p.get_or_add_pPr().find(W + "tabs")
    if tabs is None:
        tabs = OxmlElement("w:tabs"); p._p.get_or_add_pPr().append(tabs)
    t = OxmlElement("w:tab")
    t.set(qn("w:val"), "right"); t.set(qn("w:pos"), pos)
    if dots:
        t.set(qn("w:leader"), "dot")
    tabs.append(t)


def _tab_left(p, pos):
    tabs = p._p.get_or_add_pPr().find(W + "tabs")
    if tabs is None:
        tabs = OxmlElement("w:tabs"); p._p.get_or_add_pPr().append(tabs)
    t = OxmlElement("w:tab"); t.set(qn("w:val"), "left"); t.set(qn("w:pos"), pos)
    tabs.append(t)


def _set_borders(tbl, kind):
    tblPr = tbl._tbl.tblPr
    old = tblPr.find(W + "tblBorders")
    if old is not None:
        tblPr.remove(old)
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + e)
        on = True if kind == "grid" else (e in ("top", "bottom", "insideH"))
        el.set(qn("w:val"), "single" if on else "nil")
        el.set(qn("w:sz"), "4"); el.set(qn("w:space"), "0"); el.set(qn("w:color"), "000000")
        b.append(el)
    tblPr.append(b)


def _add_table(anchor, rows, bounds_tw, kind, font_size, left_cols=()):
    ncols = len(bounds_tw) - 1
    tbl = doc.add_table(rows=len(rows), cols=ncols)
    tbl._tbl.getparent().remove(tbl._tbl)
    anchor.addprevious(tbl._tbl)
    tblPr = tbl._tbl.tblPr
    tw = OxmlElement("w:tblW"); tw.set(qn("w:type"), "dxa"); tw.set(qn("w:w"), str(bounds_tw[-1] - bounds_tw[0])); tblPr.append(tw)
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    _set_borders(tbl, kind)
    grid = tbl._tbl.find(W + "tblGrid")
    for gc in list(grid):
        grid.remove(gc)
    for i in range(ncols):
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(bounds_tw[i + 1] - bounds_tw[i])); grid.append(gc)
    left_cols = set(left_cols)
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = tbl.cell(ri, ci)
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = OxmlElement("w:tcW"); tcW.set(qn("w:type"), "dxa"); tcW.set(qn("w:w"), str(bounds_tw[ci + 1] - bounds_tw[ci])); tcPr.append(tcW)
            p = cell.paragraphs[0]
            pPr = p._p.get_or_add_pPr()
            sp = OxmlElement("w:spacing"); sp.set(qn("w:line"), "240"); sp.set(qn("w:lineRule"), "auto"); sp.set(qn("w:after"), "40"); pPr.append(sp)
            al = "center" if ri == 0 else ("left" if ci in left_cols else "center")
            j = OxmlElement("w:jc"); j.set(qn("w:val"), al); pPr.append(j)
            val = row[ci] if ci < len(row) else ""
            if isinstance(val, str):
                val = [{"t": val, "b": False, "i": False}]
            if ri == 0:
                val = [dict(x, b=True) for x in val]
            for item in val:
                if isinstance(item, dict) and "img" in item:
                    rr = p.add_run()
                    fn = os.path.basename(item["img"])
                    dim = img_dims.get(fn)
                    if dim:
                        pic = rr.add_picture(item["img"], width=Cm(dim["w"]))
                        pic.height = Cm(dim["h"])
                    else:
                        rr.add_picture(item["img"], width=Cm(1.2))
                else:
                    r = p.add_run(item.get("t", ""))
                    set_run(r, size=font_size, bold=item.get("b", False), italic=item.get("i", False))
    return tbl


# ---------------------------------------------------------------- FRONT MATTER

def front_heading(text):
    p = _make_par(front_anchor)
    p._p.get_or_add_pPr().append(OxmlElement("w:pageBreakBefore"))
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run(r, size=14, bold=True)
    return p


def toc_line(text, page, indent=0):
    p = _make_par(front_anchor)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    if indent:
        p.paragraph_format.left_indent = Pt(indent)
    _tab_right(p)
    r = p.add_run(text + "\t" + str(page))
    set_run(r, size=12)
    return p


TOC = [
    ("KATA PENGANTAR", "i", 0), ("DAFTAR ISI", "iii", 0), ("DAFTAR TABEL", "iv", 0),
    ("DAFTAR GAMBAR", "v", 0), ("DAFTAR SINGKATAN", "vi", 0),
    ("BAB I PENDAHULUAN", "1", 0),
    ("A. Latar Belakang", "1", 36), ("B. Rumusan Masalah", "5", 36),
    ("C. Tujuan Penelitian", "5", 36), ("D. Manfaat Penelitian", "6", 36),
    ("E. Ruang Lingkup Penelitian", "6", 36), ("F. Sistematika Penulisan", "7", 36),
    ("BAB II TINJAUAN PUSTAKA", "9", 0),
    ("A. Landasan Teori", "9", 36), ("B. Penelitian Terkait", "28", 36),
    ("C. Kerangka Berpikir", "30", 36),
    ("BAB III METODE PENELITIAN", "31", 0),
    ("A. Tempat dan Waktu Penelitian", "31", 36), ("B. Alat dan Bahan", "33", 36),
    ("C. Perancangan Sistem", "34", 36),
    ("BAB IV HASIL DAN PEMBAHASAN", "46", 0),
    ("A. Hasil Penerapan User Centered Design", "46", 36),
    ("B. Hasil Implementasi Sistem Akhir", "48", 36),
    ("C. Hasil Pengujian Black Box", "61", 36),
    ("D. Hasil Validasi System Usability Scale", "63", 36),
    ("F. Pembahasan", "64", 36),
    ("BAB V PENUTUP", "66", 0),
    ("A. KESIMPULAN", "66", 36), ("B. SARAN", "66", 36),
    ("DAFTAR PUSTAKA", "67", 0),
]

DAFTAR_TABEL = [
    ("Tabel 1 Simbol Flowchart", "20"), ("Tabel 2 Penelitian Terkait", "28"),
    ("Tabel 3 Jadwal Penelitian", "32"), ("Tabel 4 Hyperparameter Terpilih", "60"),
    ("Tabel 5 Confusion Matrix Data Uji", "60"), ("Tabel 6 Metrik Evaluasi Data Uji", "60"),
    ("Tabel 7 Hasil Cross Validation Lima Lipatan", "61"),
    ("Tabel 8 Perbandingan Model Tanpa dan Dengan SMOTE", "61"),
    ("Tabel 9 Skenario Variasi Proporsi Data", "61"),
    ("Tabel 10 Skenario Variasi Jumlah Pohon", "62"),
    ("Tabel 11 Hasil Pengujian Black Box", "62"),
    ("Tabel 12 Pernyataan Kuesioner SUS", "63"),
    ("Tabel 13 Rekapitulasi Jawaban SUS (Diisi Setelah Survei)", "63"),
    ("Tabel 14 Rekapitulasi Penilaian TAM (Diisi Setelah Survei)", "64"),
]

DAFTAR_GAMBAR = [
    ("Gambar 1 Alur Kerja SMOTE", "13"), ("Gambar 2 Arsitektur Random Forest", "15"),
    ("Gambar 3 Kerangka Berpikir", "30"), ("Gambar 4 Lokasi RSUD Labuang Baji", "32"),
    ("Gambar 5 Desain Metode Penelitian", "34"),
    ("Gambar 6 Use Case Diagram Klasifikasi Keberhasilan Pengobatan", "38"),
    ("Gambar 7 Activity Diagram Klasifikasi Keberhasilan Pengobatan", "39"),
    ("Gambar 8 Class Diagram Klasifikasi Keberhasilan Pengobatan", "39"),
    ("Gambar 9 Flowchart Metode SMOTE dan Random Forest", "40"),
    ("Gambar 10 Halaman Login Sistem", "48"), ("Gambar 11 Dashboard Admin", "49"),
    ("Gambar 12 Dashboard Tenaga Medis", "49"), ("Gambar 13 Formulir Klasifikasi", "51"),
    ("Gambar 14 Hasil Klasifikasi dan Rekomendasi", "53"), ("Gambar 15 Riwayat Klasifikasi", "54"),
    ("Gambar 16 Evaluasi Model", "55"), ("Gambar 17 Pengujian Efektivitas SMOTE", "56"),
    ("Gambar 18 Skenario Pengujian", "57"), ("Gambar 19 Data Rekam Medis", "58"),
    ("Gambar 20 Pembersihan Data", "58"), ("Gambar 21 Manajemen Akun", "59"),
]

front_heading("DAFTAR ISI")
for text, page, ind in TOC:
    toc_line(text, page, indent=ind)

front_heading("DAFTAR TABEL")
for text, page in DAFTAR_TABEL:
    toc_line(text, page)

front_heading("DAFTAR GAMBAR")
for text, page in DAFTAR_GAMBAR:
    toc_line(text, page)

front_heading("DAFTAR SINGKATAN")
_add_table(front_anchor, full123["sings"], [int(108 * 20), int(233.5 * 20), int(488.2 * 20)],
           "grid", 12, left_cols=[1])

# ---------------------------------------------------------------- BODY

body_anchor = final_sectPr
CUR_BODY = [None]  # paragraf body teks terakhir (untuk gabung lanjutan)


def append_runs(p, spans, size=12):
    if spans:
        r = p.add_run(" ")
        set_run(r, size=size)
    _runs(p, spans, size=size)
    return p


def emit_eq(anchor_text, anchor_el):
    m = re.search(r"persamaan\s*\((\d+)\)\s*berikut:\s*$", anchor_text)
    if m:
        n = int(m.group(1))
        if n in eq_elements:
            anchor_el.addnext(copy.deepcopy(eq_elements[n]))
    return m is not None


def body_par(spans, first_line=True):
    p = _make_par(body_anchor)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    if first_line:
        p.paragraph_format.first_line_indent = Pt(709 / 20.0)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _runs(p, spans, size=12)
    txt = "".join(s["t"] for s in spans)
    emit_eq(txt, p._p)
    CUR_BODY[0] = p
    return p


def chapter_heading(parts):
    for k, t in enumerate(parts):
        p = _make_par(body_anchor)
        p.paragraph_format.line_spacing = 1.5
        if k == 0:
            p.paragraph_format.page_break_before = True
            p.paragraph_format.space_before = Pt(0)
        else:
            p.paragraph_format.keep_with_next = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(t)
        set_run(r, size=14, bold=True)


def heading_letter(text):
    p = _make_par(body_anchor)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run(r, size=12, bold=True)
    return p


def heading_sub(text):
    p = _make_par(body_anchor)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run(r, size=12, bold=True)
    return p


def caption_from_runs(runs):
    p = _make_par(body_anchor)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    _runs(p, runs, size=12)
    return p


def image_from(path, width_cm):
    p = _make_par(body_anchor)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run()
    fn = os.path.basename(path)
    dims = img_dims.get(fn)
    if dims:
        width_cm = dims["w"]
        height_cm = dims["h"]
    else:
        height_cm = None
    pic = r.add_picture(os.path.join(BASE, path), width=Cm(width_cm))
    if height_cm is not None:
        pic.height = Cm(height_cm)
    return p


def numbered_list(spans, letter=False):
    p = _make_par(body_anchor)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    first = spans[0]
    ftxt = first["t"]
    if letter:
        m = re.match(r"^([a-z]\.)\s*", ftxt)
        num, rest = m.group(1), ftxt[m.end():]
        p.paragraph_format.left_indent = Pt(7)
        r = p.add_run(num + " ")
        set_run(r, size=12, italic=first.get("i", False))
        if rest:
            r2 = p.add_run(rest)
            set_run(r2, size=12, italic=first.get("i", False))
    else:
        m = re.match(r"^(\d+\.)\s*", ftxt)
        num, rest = m.group(1), ftxt[m.end():]
        p.paragraph_format.left_indent = Pt(18)
        p.paragraph_format.first_line_indent = Pt(-18)
        _tab_left(p, "360")
        r = p.add_run(num)
        set_run(r, size=12, italic=first.get("i", False))
        rtab = p.add_run("\t")
        set_run(rtab, size=12)
        if rest:
            r2 = p.add_run(rest)
            set_run(r2, size=12, italic=first.get("i", False))
    for s in spans[1:]:
        r = p.add_run(s["t"])
        set_run(r, size=12, bold=s.get("b", False), italic=s.get("i", False))
    return p


def sistematika_par(bab, title):
    p = _make_par(body_anchor)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(3.2)
    p.paragraph_format.first_line_indent = Cm(-3.2)
    _tab_left(p, "1814")
    r = p.add_run(bab); set_run(r, size=12)
    rt = p.add_run("\t"); set_run(rt, size=12)
    r2 = p.add_run(title); set_run(r2, size=12, bold=True)
    return p


def sistematika_desc(spans):
    p = _make_par(body_anchor)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(3.2)
    _runs(p, spans, size=12)
    return p


TB_BOUNDS = {
    1: [108.0, 207.2, 320.6, 504.9],
    2: [108.0, 183.0, 313.0, 381.0, 505.0],
    3: [108.0, 135.0, 402.0, 468.0, 516.7],
}
TB_KIND = {1: "grid", 2: "horizontal", 3: "horizontal"}
TB_FONT = {1: 12, 2: 10, 3: 10}
TB_LEFT = {1: [1, 2], 2: [0, 1, 2, 3], 3: [1]}

bab45_tables = [x for x in bab45 if x["type"] == "table"]
BAB45_LEFT = {4: [0], 5: [], 6: [0], 7: [0], 8: [0], 9: [], 10: [], 11: [1, 2, 3], 12: [2], 13: [], 14: [0]}


def build_bab45_table(table_no):
    tdata = bab45_tables[table_no - 4]
    left = 108.0
    acc = left
    bounds_pt = [left]
    for w in tdata["widths"]:
        acc += w * 567.0 / 20.0
        bounds_pt.append(acc)
    bounds_tw = [int(round(x * 20)) for x in bounds_pt]
    _add_table(body_anchor, tdata["rows"], bounds_tw, "horizontal", 10, left_cols=BAB45_LEFT[table_no])


def build123():
    blocks = full123["blocks"]
    i = 0
    while i < len(blocks):
        b = blocks[i]
        t = b["type"]
        if t == "chapter_line":
            txt = "".join(r["t"] for r in b["runs"])
            if txt == "DAFTAR SINGKATAN":
                i += 1
                continue
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            title = "".join(r["t"] for r in nxt["runs"]) if nxt and nxt["type"] == "chapter_line" else ""
            chapter_heading([txt, title])
            if nxt and nxt["type"] == "chapter_line":
                i += 1
        elif t == "h_a":
            heading_letter("".join(r["t"] for r in b["runs"]))
        elif t == "h_sub":
            heading_sub("".join(r["t"] for r in b["runs"]))
        elif t == "body":
            spans = b["runs"]
            txt = "".join(r["t"] for r in spans)
            if txt.startswith("Pengambilan data rekam medis pasien MDR-TB"):
                spans = spans + [{"t": REVISION, "b": False, "i": False}]
                body_par(spans)
            elif b.get("list"):
                numbered_list(spans, letter=bool(re.match(r"^[a-z]\.", spans[0]["t"])))
            elif b.get("x0") == 204.2:
                sistematika_desc(spans)
            elif b.get("x0") == 113.5:
                # lanjutan paragraf (kasus Gambar 9) -> gabung ke paragraf teks terakhir
                if CUR_BODY[0] is not None:
                    append_runs(CUR_BODY[0], spans)
                else:
                    body_par(spans, first_line=False)
            else:
                body_par(spans)
        elif t == "sistematika":
            sistematika_par(b["bab"], b["title"])
        elif t == "caption":
            caption_from_runs(b["runs"])
        elif t == "image":
            image_from(b["file"], b["width_cm"])
        elif t == "table":
            no = b["table_no"]
            _add_table(body_anchor, b["rows"], [int(round(x * 20)) for x in TB_BOUNDS[no]],
                       TB_KIND[no], TB_FONT[no], left_cols=TB_LEFT[no])
        i += 1


def build45():
    blocks = bab45_runs["blocks"]
    i = 0
    while i < len(blocks):
        b = blocks[i]
        t = b["type"]
        if t == "chapter_line":
            txt = "".join(r["t"] for r in b["runs"])
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            title = "".join(r["t"] for r in nxt["runs"]) if nxt and nxt["type"] == "chapter_line" else ""
            chapter_heading([txt, title])
            if nxt and nxt["type"] == "chapter_line":
                i += 1
        elif t == "h_a":
            heading_letter("".join(r["t"] for r in b["runs"]))
        elif t == "h_sub":
            heading_sub("".join(r["t"] for r in b["runs"]))
        elif t == "body":
            body_par(b["runs"])
        elif t == "caption":
            caption_from_runs(b["runs"])
        elif t == "image":
            image_from(b["file"], b["width_cm"])
        elif t == "table_slot":
            build_bab45_table(b["table_no"])
        i += 1


def build_refs():
    p = _make_par(body_anchor)
    p.paragraph_format.page_break_before = True
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("DAFTAR PUSTAKA")
    set_run(r, size=14, bold=True)
    for ent in refs:
        p = _make_par(body_anchor)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Pt(28.35)
        p.paragraph_format.first_line_indent = Pt(-28.35)
        for s in ent:
            r = p.add_run(s["t"])
            set_run(r, size=12, italic=s.get("i", False))


build123()
build45()
build_refs()

doc.save(OUT)
print("saved:", OUT)