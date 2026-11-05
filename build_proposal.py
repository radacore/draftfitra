#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build proposal skripsi: Penerapan Teknik SMOTE dan Random Forest dalam
Klasifikasi Keberhasilan Pengobatan Penyakit MDR-TB.
Struktur mengikuti contoh penulisan.pdf (Unismuh Makassar).
Format: A4, Times New Roman 12, spasi 1.5, justify, margin kiri 4/kanan 3/atas 3/bawah 3.
Dua pass: bangun -> render -> buat pagemap -> rebuild dengan nomor halaman riil.

Iterasi 2:
  - Persamaan matematika native OMML (LaTeX -> MathML -> OMML).
  - Gambar programatik matplotlib (4 PNG).
  - Ekspansi konten (~50 halaman) + Tabel 3 Jadwal Penelitian.
"""

import json
import os
import re
import sys

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_BREAK
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from lxml import etree
from latex2mathml.converter import convert

BASE = os.path.dirname(os.path.abspath(__file__))
OUTDOCX = os.path.join(BASE, "Proposal_Skripsi_Zulfitra_MDR-TB.docx")
PAGEMAP = os.path.join(BASE, "pagemap.json")
FIGDIR = os.path.join(BASE, "figures")
XSLT_PATH = os.path.join(BASE, "mml2omml.xsl")

FONT = "Times New Roman"

XSLT = etree.XSLT(etree.parse(XSLT_PATH))


def set_run(r, size=12, bold=False, italic=False):
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    rpr = r._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)


def set_normal_style(doc):
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(12)
    pf = st.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)


def set_section(sec, footer_kind=None, fmt=None, start=None):
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(3)
    sec.bottom_margin = Cm(3)
    sec.left_margin = Cm(4)
    sec.right_margin = Cm(3)
    if fmt is not None:
        sectPr = sec._sectPr
        pgn = sectPr.find(qn("w:pgNumType"))
        if pgn is None:
            pgn = OxmlElement("w:pgNumType")
            sectPr.append(pgn)
        pgn.set(qn("w:fmt"), fmt)
        if start is not None:
            pgn.set(qn("w:start"), str(start))
    if footer_kind == "page":
        sec.footer.is_linked_to_previous = False
        p = sec.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = " PAGE "
        f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
        run._r.append(f1); run._r.append(it); run._r.append(f2)
        set_run(run)


def seq_field(doc, kind, num, bold=True):
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " SEQ %s \\* ARABIC " % kind)
    fr = OxmlElement("w:r")
    frpr = OxmlElement("w:rPr")
    if bold:
        b = OxmlElement("w:b"); frpr.append(b)
    fonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attr), FONT)
    frpr.append(fonts)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "24"); frpr.append(sz)
    fr.append(frpr)
    ft = OxmlElement("w:t"); ft.text = str(num)
    fr.append(ft)
    fld.append(fr)
    return fld


def add_caption(doc, kind, num, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(kind + " ")
    set_run(r, bold=True)
    p._p.append(seq_field(doc, kind, num))
    r2 = p.add_run(". " + title)
    set_run(r2, bold=True)
    return p


def add_body(doc, text, indent=True, justify=True):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    if indent:
        pf.first_line_indent = Cm(1.25)
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    set_run(r)
    return p


def add_bullet_num(doc, text, left=1.0, hang=1.0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    m = re.match(r"^([0-9]+|[a-z])\.\s+(.*)$", text)
    if m:
        pf.left_indent = Cm(left)
        pf.first_line_indent = Cm(-hang)
        pf.tab_stops.add_tab_stop(Cm(left))
        r = p.add_run(m.group(1) + ".")
        set_run(r)
        r2 = p.add_run("\t")
        set_run(r2)
        r3 = p.add_run(m.group(2))
        set_run(r3)
    else:
        pf.left_indent = Cm(left)
        r = p.add_run(text)
        set_run(r)
    return p


def add_sistematika(doc, bab, title, desc):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(6)
    pf.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.tab_stops.add_tab_stop(Cm(3.2))
    r = p.add_run(bab)
    set_run(r)
    r2 = p.add_run("\t")
    set_run(r2)
    r3 = p.add_run(title)
    set_run(r3, bold=True)
    d = doc.add_paragraph()
    df = d.paragraph_format
    df.left_indent = Cm(3.2)
    df.line_spacing = 1.5
    df.space_after = Pt(0)
    d.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r4 = d.add_run(desc)
    set_run(r4)
    return p


def add_heading_letter(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(10)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    set_run(r, bold=True)
    return p


def add_heading_sub(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(8)
    pf.space_after = Pt(4)
    pf.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    set_run(r, bold=True)
    return p


def repair_nary(root):
    """Perbaiki OMML hasil XSLT: isi operand m:nary yang kosong.

    latex2mathml menulis operand \\sum sebagai saudara (sibling) dari operator
    sehingga XSLT menghasilkan <m:nary><m:e/> kosong -> dirender sebagai kotak
    oleh LibreOffice. Pindahkan sibling berikutnya ke dalam m:e.
    """
    m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    for nary in list(root.iter(m + "nary")):
        e = nary.find(m + "e")
        if e is None or len(e) > 0:
            continue
        parent = nary.getparent()
        if parent is None:
            continue
        idx = list(parent).index(nary)
        for sib in list(parent)[idx + 1:]:
            e.append(sib)
    return root


def add_equation_latex(doc, latex, number):
    """LaTeX -> MathML (latex2mathml) -> OMML (XSLT) -> repair -> paragraf bernomor."""
    mml = convert(latex, display="block")
    tree = etree.fromstring(mml.encode("utf-8"))
    omml = XSLT(tree)
    repair_nary(omml.getroot())
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.tab_stops.add_tab_stop(Cm(13.8), WD_TAB_ALIGNMENT.RIGHT)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p._p.append(omml.getroot())
    r = p.add_run("\t(" + number + ")")
    set_run(r)
    return p


def add_figure(doc, filename, width=12.0):
    path = os.path.join(FIGDIR, filename)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    run = p.add_run()
    run.add_picture(path, width=Cm(width))
    return p


def add_bab_heading(doc, num, title):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.page_break_before = True
    pf.space_before = Pt(0)
    pf.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("BAB " + num)
    set_run(r, size=14, bold=True)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.line_spacing = 1.5
    r2 = p2.add_run(title)
    set_run(r2, size=14, bold=True)
    return p


def add_front_heading(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.page_break_before = True
    pf.space_before = Pt(0)
    pf.space_after = Pt(12)
    pf.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run(r, size=14, bold=True)
    return p


def add_body_heading(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.page_break_before = True
    pf.space_before = Pt(0)
    pf.space_after = Pt(12)
    pf.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run(r, size=14, bold=True)
    return p


def toc_line(doc, text, page, indent=0.0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.tab_stops.add_tab_stop(Cm(13.8), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    if indent:
        pf.left_indent = Cm(indent)
    r = p.add_run(text + "\t" + str(page))
    set_run(r)
    return p


def set_cell(cell, text, bold=False, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.line_spacing = 1.0
    pf.space_after = Pt(2)
    if align:
        p.alignment = align
    r = p.add_run(text)
    set_run(r, size=10, bold=bold)


def add_table(doc, rows, cols, header, data, widths=None):
    t = doc.add_table(rows=rows, cols=cols)
    t.style = "Table Grid"
    t.autofit = False
    for j, h in enumerate(header):
        set_cell(t.cell(0, j), h, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, row in enumerate(data, start=1):
        for j, val in enumerate(row):
            set_cell(t.cell(i, j), val, align=WD_ALIGN_PARAGRAPH.LEFT)
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Cm(w)
    return t


def load_pagemap():
    if os.path.exists(PAGEMAP):
        with open(PAGEMAP, encoding="utf-8") as f:
            return json.load(f)
    return None


def build():
    pm = load_pagemap()

    def page(group, key, default):
        if pm and group in pm and key in pm[group]:
            return pm[group][key]
        return default

    doc = Document()
    set_normal_style(doc)
    doc.core_properties.title = (
        "Penerapan Teknik SMOTE dan Random Forest dalam Klasifikasi "
        "Keberhasilan Pengobatan Penyakit MDR-TB"
    )

    # ------------------------------------------------------------- SECTION 0: COVER
    sec0 = doc.sections[0]
    set_section(sec0)

    def center_line(text, size=12, bold=False, space_after=0):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(space_after)
        r = p.add_run(text)
        set_run(r, size=size, bold=bold)
        return p

    center_line("")
    center_line("")
    center_line("PENERAPAN TEKNIK SMOTE DAN RANDOM FOREST DALAM", size=16, bold=True)
    center_line("KLASIFIKASI KEBERHASILAN PENGOBATAN PENYAKIT MDR-TB", size=16, bold=True)
    center_line("")
    center_line("")
    center_line("PROPOSAL PENELITIAN", size=14, bold=True)
    center_line("")
    center_line("Diajukan Sebagai Salah Satu Syarat untuk Mendapat Gelar", size=12)
    center_line("Sarjana Komputer (S.Kom) Program Studi Informatika", size=12)
    center_line("")
    center_line("")
    center_line("Disusun Oleh :", size=12)
    center_line("")
    center_line("ZULFITRA", size=14, bold=True)
    center_line("105841115022", size=12)
    center_line("")
    center_line("")
    center_line("")
    center_line("PROGRAM STUDI INFORMATIKA", size=12, bold=True)
    center_line("FAKULTAS TEKNIK", size=12, bold=True)
    center_line("UNIVERSITAS MUHAMMADIYAH MAKASSAR", size=12, bold=True)
    center_line("2026", size=12)

    # ------------------------------------------------------------- SECTION 1: FRONT MATTER (romawi)
    sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
    set_section(sec1, footer_kind="page", fmt="lowerRoman", start=1)

    # Kata Pengantar
    add_front_heading(doc, "KATA PENGANTAR")
    add_body(doc,
        "Assalamu'alaikum warahmatullahi wabarakatuh.")
    add_body(doc,
        "Puji syukur penulis panjatkan ke hadirat Allah SWT atas limpahan rahmat, "
        "taufik, dan hidayah-Nya sehingga proposal penelitian yang berjudul "
        "\u201cPenerapan Teknik SMOTE dan Random Forest dalam Klasifikasi Keberhasilan "
        "Pengobatan Penyakit MDR-TB\u201d dapat diselesaikan dengan baik. Sebagai bentuk rasa "
        "hormat, penulis menyampaikan terima kasih yang sebesar-besarnya kepada kedua orang tua, "
        "[Nama Ayah] dan [Nama Ibu], atas doa, dukungan, motivasi, serta bantuan yang tiada henti "
        "mengiringi proses studi hingga penyusunan proposal ini. Shalawat dan salam semoga tercurah "
        "kepada Nabi Muhammad SAW yang telah membawa umat manusia menuju peradaban ilmu pengetahuan.")
    add_body(doc,
        "Sehubungan dengan itu, penulis menyampaikan penghargaan dan ucapan terima kasih kepada:")
    ucapkan = [
        "1. Bapak Dr. Ir. H. Abd. Rakhim Nanda, ST., MT., IPU., selaku Rektor Universitas Muhammadiyah Makassar.",
        "2. Bapak Ir. Muhammad Syafaat S Kuba, ST., MT., selaku Dekan Fakultas Teknik Universitas Muhammadiyah Makassar.",
        "3. Ibu Rizki Yusliana Bakti, S.T., M.T., selaku Ketua Program Studi Informatika Universitas Muhammadiyah Makassar.",
        "4. Ibu [Nama Dosen Pembimbing I], selaku Dosen Pembimbing I, atas bimbingan, masukan ilmiah, dan koreksi yang sangat membantu penyempurnaan proposal penelitian ini.",
        "5. Ibu Chyquitha Danuputri, S.Kom., M.T., selaku Dosen Pembimbing II, atas bimbingan, masukan ilmiah, dan koreksi yang sangat membantu penyempurnaan proposal penelitian ini.",
        "6. Seluruh dosen Program Studi Informatika, atas ilmu dan pengalaman yang diberikan kepada penulis selama masa studi.",
        "7. Seluruh staf Tata Usaha Fakultas Teknik dan Program Studi Informatika, atas bantuan administratif dan kelancaran proses pengurusan dokumen akademik selama masa studi.",
        "8. Pihak RSUD Labuang Baji, Sulawesi Selatan, atas izin dan kesempatan dalam pengambilan data penelitian yang menjadi bagian penting dari penelitian ini.",
        "9. Rekan-rekan dan sahabat yang telah memberikan dukungan akademik maupun moral dan motivasi selama penyusunan proposal.",
    ]
    for u in ucapkan:
        add_bullet_num(doc, u)
    add_body(doc,
        "Penulis menyadari bahwa proposal ini masih memiliki keterbatasan. Oleh karena itu, "
        "kritik dan saran yang konstruktif sangat diharapkan demi perbaikan dan penyempurnaan "
        "penelitian selanjutnya. Semoga Allah SWT membalas segala bantuan dan kebaikan yang telah diberikan.")
    add_body(doc, "Wassalamu'alaikum warahmatullahi wabarakatuh.")
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Makassar, [tanggal bulan tahun]\nPenulis,\n\n\nZULFITRA")
    set_run(r)

    # Daftar Isi
    add_front_heading(doc, "DAFTAR ISI")
    toc_line(doc, "KATA PENGANTAR", page("front", "KATA PENGANTAR", "i"))
    toc_line(doc, "DAFTAR ISI", page("front", "DAFTAR ISI", "ii"))
    toc_line(doc, "DAFTAR TABEL", page("front", "DAFTAR TABEL", "iii"))
    toc_line(doc, "DAFTAR GAMBAR", page("front", "DAFTAR GAMBAR", "iv"))
    toc_line(doc, "BAB I PENDAHULUAN", page("body", "BAB I PENDAHULUAN", "1"))
    toc_line(doc, "A. Latar Belakang", page("body", "A. Latar Belakang", "1"), indent=1.27)
    toc_line(doc, "B. Rumusan Masalah", page("body", "B. Rumusan Masalah", "6"), indent=1.27)
    toc_line(doc, "C. Tujuan Penelitian", page("body", "C. Tujuan Penelitian", "6"), indent=1.27)
    toc_line(doc, "D. Manfaat Penelitian", page("body", "D. Manfaat Penelitian", "7"), indent=1.27)
    toc_line(doc, "E. Ruang Lingkup Penelitian", page("body", "E. Ruang Lingkup Penelitian", "8"), indent=1.27)
    toc_line(doc, "F. Sistematika Penulisan", page("body", "F. Sistematika Penulisan", "9"), indent=1.27)
    toc_line(doc, "BAB II TINJAUAN PUSTAKA", page("body", "BAB II TINJAUAN PUSTAKA", "11"))
    toc_line(doc, "A. Landasan Teori", page("body", "A. Landasan Teori", "11"), indent=1.27)
    toc_line(doc, "B. Penelitian Terkait", page("body", "B. Penelitian Terkait", "34"), indent=1.27)
    toc_line(doc, "C. Kerangka Berpikir", page("body", "C. Kerangka Berpikir", "36"), indent=1.27)
    toc_line(doc, "BAB III METODE PENELITIAN", page("body", "BAB III METODE PENELITIAN", "39"))
    toc_line(doc, "A. Tempat dan Waktu Penelitian", page("body", "A. Tempat dan Waktu Penelitian", "39"), indent=1.27)
    toc_line(doc, "B. Alat dan Bahan", page("body", "B. Alat dan Bahan", "40"), indent=1.27)
    toc_line(doc, "C. Perancangan Sistem", page("body", "C. Perancangan Sistem", "42"), indent=1.27)
    toc_line(doc, "DAFTAR PUSTAKA", page("body", "DAFTAR PUSTAKA", "50"))

    # Daftar Tabel
    add_front_heading(doc, "DAFTAR TABEL")
    toc_line(doc, "Tabel 1 Simbol Flowchart", page("tabel", "Tabel 1", "27"))
    toc_line(doc, "Tabel 2 Penelitian Terkait", page("tabel", "Tabel 2", "34"))
    toc_line(doc, "Tabel 3 Jadwal Penelitian", page("tabel", "Tabel 3", "40"))

    # Daftar Gambar
    add_front_heading(doc, "DAFTAR GAMBAR")
    toc_line(doc, "Gambar 1 Alur Kerja SMOTE", page("gambar", "Gambar 1", "16"))
    toc_line(doc, "Gambar 2 Arsitektur Random Forest", page("gambar", "Gambar 2", "22"))
    toc_line(doc, "Gambar 3 Kerangka Berpikir", page("gambar", "Gambar 3", "48"))
    toc_line(doc, "Gambar 4 Desain Metode Penelitian", page("gambar", "Gambar 4", "40"))
    toc_line(doc, "Gambar 5 Flowchart Metode SMOTE dan Random Forest", page("gambar", "Gambar 5", "46"))

    # ------------------------------------------------------------- SECTION 2: BODY (arab)
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    set_section(sec2, footer_kind="page", fmt="decimal", start=1)

    # ================================================================ BAB I
    add_bab_heading(doc, "I", "PENDAHULUAN")

    add_heading_letter(doc, "A. Latar Belakang")
    add_body(doc,
        "Multi-Drug Resistant Tuberculosis (MDR-TB) merupakan salah satu tantangan paling berat "
        "dalam upaya pengendalian tuberkulosis secara global. Pasien yang terinfeksi galur bakteri "
        "Mycobacterium tuberculosis yang resisten terhadap obat anti-tuberkulosis (OAT) lini pertama "
        "harus menjalani regimen pengobatan yang jauh lebih panjang, menelan biaya yang lebih "
        "tinggi, serta memicu efek samping yang sering kali melemahkan kondisi fisik dan mental "
        "pasien. Meskipun jumlah kasus tuberkulosis global menunjukkan penurunan yang cukup berarti, "
        "penanganan terhadap varian MDR-TB masih jauh dari target yang diharapkan. Permasalahan ini "
        "menuntut strategi pengendalian yang lebih terarah, tidak hanya dari sisi klinis, tetapi "
        "juga dari sisi sistem informasi kesehatan yang mendukung pengambilan keputusan "
        "(World Health Organization, 2024).")
    add_body(doc,
        "Di Indonesia, urgensi permasalahan ini berada pada titik yang sangat kritis. Indonesia "
        "menyumbang sekitar 10% dari total beban tuberkulosis dunia sehingga menempati peringkat "
        "kedua tertinggi secara global setelah India. Tingkat keberhasilan pengobatan MDR-TB di "
        "berbagai fasilitas kesehatan nasional terus menghadapi kendala yang besar. Angka putus "
        "berobat (loss to follow-up) dan gagal pengobatan masih memprihatinkan akibat faktor yang "
        "sangat kompleks, mulai dari komorbiditas klinis, tingkat keparahan efek samping obat, "
        "hingga kendala sosial-ekonomi yang dialami pasien. Tenaga medis sering kali kesulitan "
        "mendeteksi potensi kegagalan pengobatan secara dini karena belum adanya sistem analitik "
        "cerdas yang mampu membaca pola risiko pasien secara spesifik sejak awal masa pengobatan "
        "(World Health Organization, 2025).")
    add_body(doc,
        "Salah satu akar persoalan yang dihadapi tenaga medis adalah keterbatasan analisis "
        "terhadap data rekam medis pasien yang telah terkumpul selama bertahun-tahun. Data rekam "
        "medis sebenarnya menyimpan informasi yang sangat berharga, seperti karakteristik demografis "
        "pasien, kondisi klinis, komorbiditas, serta riwayat pengobatan. Namun, data tersebut "
        "umumnya diolah secara manual dan bersifat deskriptif sehingga pola-pola yang kompleks "
        "tidak dapat dikenali. Transformasi data mentah menjadi pengetahuan yang dapat dimanfaatkan "
        "untuk memprediksi keberhasilan pengobatan memerlukan pendekatan analitis yang lebih "
        "canggih, salah satunya melalui pemanfaatan machine learning (Zhang dkk, 2025).")
    add_body(doc,
        "Pemanfaatan machine learning dalam bidang kesehatan telah terbukti memberikan banyak "
        "manfaat, antara lain dalam diagnosis penyakit, prediksi hasil pengobatan, dan pengambilan "
        "keputusan klinis. Klasifikasi merupakan salah satu tugas utama machine learning yang "
        "bertujuan mengelompokkan data ke dalam kelas tertentu berdasarkan pola yang dipelajari "
        "dari data latih. Pada konteks pengobatan MDR-TB, machine learning dapat digunakan untuk "
        "mengklasifikasikan pasien berdasarkan potensi keberhasilan atau kegagalan pengobatan, "
        "sehingga tenaga medis dapat memberikan intervensi yang lebih tepat dan lebih awal "
        "(Kurniawati & Saputra, 2023).")
    add_body(doc,
        "Meskipun machine learning menawarkan potensi yang besar, penerapannya pada data rekam "
        "medis dihadapkan pada tantangan teknis yang tidak sederhana. Salah satu tantangan utama "
        "adalah ketidakseimbangan kelas (data imbalance). Pada dataset rekam medis pasien MDR-TB, "
        "jumlah sampel dengan status \u201cberhasil sembuh\u201d umumnya jauh lebih banyak "
        "dibandingkan sampel dengan status \u201cgagal\u201d, \u201cmeninggal\u201d, atau "
        "\u201cputus obat\u201d. Kondisi ini merupakan representasi nyata dari distribusi populasi, "
        "tetapi menjadi masalah serius ketika dataset langsung digunakan untuk melatih model "
        "(Chawla et al., 2002).")
    add_body(doc,
        "Ketika dataset yang tidak seimbang digunakan secara langsung, algoritma klasifikasi akan "
        "cenderung bias terhadap kelas mayoritas. Model akan memprediksi sebagian besar pasien "
        "sebagai \u201cberhasil sembuh\u201d demi meraih akurasi semu yang tinggi, namun justru buta "
        "dalam mendeteksi pasien yang sebenarnya berisiko tinggi gagal pengobatan. Dalam konteks "
        "pengendalian MDR-TB, kegagalan deteksi pada kelas minoritas ini sangat fatal karena tidak "
        "hanya membahayakan keselamatan pasien, tetapi juga meningkatkan risiko penyebaran galur "
        "bakteri super-resisten di masyarakat. Oleh karena itu, penanganan ketidakseimbangan data "
        "menjadi prasyarat penting sebelum model dilatih (Kurniawati & Saputra, 2023).")
    add_body(doc,
        "Untuk mengatasi anomali distribusi data tersebut, teknik SMOTE (Synthetic Minority "
        "Over-sampling Technique) hadir sebagai solusi komputasional yang tangguh. Berbeda dengan "
        "pendekatan oversampling konvensional yang hanya menduplikasi data kelas minoritas, SMOTE "
        "bekerja dengan mensintesis data buatan (synthetic) yang baru dan logis berdasarkan "
        "kedekatan jarak fitur antar sampel melalui mekanisme k-nearest neighbors. Penyeimbangan "
        "dataset dengan SMOTE memungkinkan model klasifikasi belajar mengenali pola-pola kritis "
        "pada kelas minoritas tanpa mengalami overfitting. Dengan demikian, sistem prediksi menjadi "
        "jauh lebih sensitif terhadap kondisi pasien yang berisiko gagal pengobatan "
        "(Hadjadj dkk, 2024).")
    add_body(doc,
        "Sebagai mesin pengklasifikasi utama, Random Forest dipilih karena keunggulannya dalam "
        "memproses dataset medis yang berskala multidimensi. Algoritma ensemble ini bekerja dengan "
        "mengumpulkan hasil prediksi (voting) dari sejumlah besar pohon keputusan (decision tree), "
        "sehingga menghasilkan model yang stabil terhadap noise dan data pencilan (outlier). "
        "Kekuatan Random Forest tidak hanya terletak pada tingginya metrik evaluasi, melainkan "
        "juga pada kemampuannya menyajikan informasi feature importance. Melalui fitur ini, model "
        "dapat memetakan variabel klinis atau demografis apa saja yang paling krusial dalam "
        "menentukan nasib pengobatan pasien MDR-TB (Breiman, 2001; Zhang dkk, 2025).")
    add_body(doc,
        "Kombinasi antara SMOTE dan Random Forest menjadi sinergi yang ideal. SMOTE menyeimbangkan "
        "distribusi kelas sehingga Random Forest dapat mempelajari pola kelas minoritas secara "
        "optimal, sementara Random Forest menyediakan model klasifikasi yang akurat, stabil, dan "
        "mudah diinterpretasikan melalui feature importance. Integrasi kedua teknik ini masih "
        "jarang diterapkan secara khusus pada data rekam medis pasien MDR-TB di Indonesia, "
        "sehingga penelitian ini diharapkan dapat mengisi celah (gap) tersebut. Dengan "
        "mengintegrasikan SMOTE untuk resolusi bias data dan Random Forest sebagai mesin "
        "prediktif, model yang dihasilkan diharapkan mampu mengenali pasien yang berpotensi gagal "
        "pengobatan sejak awal, sehingga penanganan dapat dilakukan lebih cepat, tepat, dan "
        "terarah (Zhang dkk, 2025).")
    add_body(doc,
        "Penerapan machine learning untuk memprediksi hasil pengobatan tuberkulosis telah banyak "
        "dilakukan di berbagai negara dengan hasil yang menjanjikan. Namun, sebagian besar "
        "penelitian tersebut menggunakan dataset internasional dengan karakteristik pasien yang "
        "berbeda dari populasi Indonesia. Kondisi sosial-ekonomi, prevalensi komorbiditas, serta "
        "sistem pelayanan kesehatan yang berbeda dapat memengaruhi pola keberhasilan pengobatan, "
        "sehingga diperlukan kajian yang menggunakan data lokal untuk memperoleh model yang "
        "relevan dengan konteks Indonesia.")
    add_body(doc,
        "Selain itu, sebagian besar penelitian sebelumnya berfokus pada akurasi keseluruhan tanpa "
        "menaruh perhatian pada kelas minoritas yang sebenarnya paling kritis untuk dideteksi. "
        "Padahal dalam penanganan MDR-TB, keterlambatan deteksi pasien yang berisiko gagal dapat "
        "berujung pada perburukan kondisi dan peningkatan penularan. Penelitian ini berupaya "
        "menutup celah tersebut dengan menempatkan recall kelas minoritas sebagai fokus utama "
        "evaluasi, sehingga model yang dihasilkan benar-benar berguna dalam praktik klinis.")
    add_body(doc,
        "Perkembangan teknologi informasi dan komunikasi telah membawa perubahan besar dalam "
        "berbagai bidang kehidupan, termasuk bidang kesehatan. Pemanfaatan teknologi dalam "
        "pelayanan kesehatan tidak hanya terbatas pada sistem informasi manajemen rumah sakit, "
        "tetapi juga berkembang ke arah analisis data yang lebih mendalam untuk mendukung "
        "pengambilan keputusan klinis. Data mining dan machine learning menjadi jembatan yang "
        "menghubungkan data mentah dengan pengetahuan yang dapat ditindaklanjuti (actionable "
        "knowledge), sehingga tenaga kesehatan dapat mengambil keputusan yang lebih cepat dan "
        "berbasis bukti (evidence-based practice).")
    add_body(doc,
        "Di sisi lain, data medis memiliki karakteristik yang heterogen, seperti data numerik "
        "(usia, lama pengobatan), data kategorikal (jenis kelamin, komorbiditas), dan data "
        "ordinal (tingkat keparahan). Penggabungan berbagai tipe data ini menjadi tantangan "
        "tersendiri dalam membangun model yang baik. Proses preprocessing yang tepat menjadi "
        "kunci untuk memastikan data dapat diproses oleh algoritma secara optimal. Oleh karena "
        "itu, penelitian ini menaruh perhatian besar pada tahapan pengolahan data awal sebelum "
        "model klasifikasi dibangun.")
    add_body(doc,
        "Penelitian ini menggunakan data rekam medis pasien MDR-TB yang diperoleh dari "
        "RSUD Labuang Baji, Sulawesi Selatan. Data tersebut mencakup variabel demografis, kondisi klinis, "
        "serta riwayat pengobatan pasien sebagai landasan analisis. Dengan memanfaatkan teknik "
        "SMOTE untuk menyeimbangkan kelas dan Random Forest untuk membangun model klasifikasi, "
        "penelitian ini diharapkan memberikan kontribusi nyata dalam mendukung pengambilan "
        "keputusan klinis pada penanganan penyakit MDR-TB di Indonesia, sekaligus menjadi salah "
        "satu upaya percepatan pencapaian target eliminasi tuberkulosis sesuai strategi global "
        "(World Health Organization, 2015).")

    add_heading_letter(doc, "B. Rumusan Masalah")
    add_body(doc,
        "Berdasarkan latar belakang yang telah dijelaskan, terdapat dua aspek utama yang menjadi "
        "fokus penelitian, yaitu aspek penanganan data dan aspek performa model. Kedua aspek "
        "tersebut saling berkaitan karena kualitas model sangat ditentukan oleh kesiapan data "
        "yang digunakan. Oleh karena itu, penelitian ini difokuskan pada permasalahan-permasalahan "
        "utama berikut:")
    add_bullet_num(doc,
        "1. Bagaimana mengimplementasikan teknik SMOTE (Synthetic Minority Over-sampling "
        "Technique) untuk mengatasi masalah ketidakseimbangan kelas (data imbalance) pada "
        "dataset rekam medis pasien MDR-TB?")
    add_bullet_num(doc,
        "2. Seberapa tinggi tingkat akurasi model Random Forest setelah penerapan SMOTE dalam "
        "mendeteksi potensi kegagalan pengobatan pada kelas minoritas?")

    add_heading_letter(doc, "C. Tujuan Penelitian")
    add_body(doc, "Sejalan dengan rumusan masalah di atas, tujuan penelitian ini adalah:")
    add_bullet_num(doc,
        "1. Mengimplementasikan teknik SMOTE (Synthetic Minority Over-sampling Technique) untuk "
        "menyeimbangkan dataset rekam medis pasien MDR-TB.")
    add_bullet_num(doc,
        "2. Mengukur tingkat akurasi model Random Forest setelah penerapan SMOTE dalam "
        "mengklasifikasikan keberhasilan pengobatan pasien MDR-TB, khususnya pada kelas "
        "minoritas yang berisiko gagal pengobatan.")
    add_body(doc,
        "Dengan tercapainya tujuan tersebut, diharapkan penelitian ini memberikan kontribusi "
        "berupa model klasifikasi yang andal serta pemahaman mengenai pengaruh penanganan "
        "ketidakseimbangan data terhadap performa model Random Forest pada data kesehatan.")

    add_heading_letter(doc, "D. Manfaat Penelitian")
    add_body(doc, "Penelitian ini diharapkan memberikan manfaat sebagai berikut:")
    add_bullet_num(doc,
        "1. Manfaat Teoritis, yaitu menambah khazanah ilmu pengetahuan dalam penerapan teknik "
        "penanganan data imbalance menggunakan SMOTE dan algoritma Random Forest pada bidang "
        "kesehatan, khususnya klasifikasi keberhasilan pengobatan penyakit MDR-TB. Hasil "
        "penelitian ini dapat menjadi referensi bagi pengembangan model machine learning pada "
        "data kesehatan yang tidak seimbang di masa mendatang.")
    add_bullet_num(doc,
        "2. Manfaat Praktis, yaitu: (a) membantu tenaga medis dalam mengidentifikasi pasien "
        "MDR-TB yang berisiko gagal pengobatan secara lebih dini dan objektif; (b) menjadi bahan "
        "referensi bagi peneliti lain dalam pengembangan model machine learning berbasis data "
        "klinis yang tidak seimbang; serta (c) memberikan gambaran potensi pemanfaatan sistem "
        "pendukung keputusan berbasis machine learning di fasilitas kesehatan.")
    add_body(doc,
        "Dengan tercapainya tujuan penelitian, diharapkan manfaat yang telah diuraikan dapat "
        "terealisasi secara nyata. Hasil penelitian tidak hanya menjadi karya ilmiah yang "
        "melengkapi syarat akademik, tetapi juga memberikan kontribusi positif bagi pengembangan "
        "sistem informasi kesehatan di Indonesia dalam menghadapi tantangan pengendalian "
        "penyakit MDR-TB.")

    add_heading_letter(doc, "E. Ruang Lingkup Penelitian")
    add_body(doc,
        "Penelitian ini berfokus pada penerapan teknik SMOTE dan algoritma Random Forest untuk "
        "mengklasifikasikan keberhasilan pengobatan pasien MDR-TB berdasarkan data rekam medis. "
        "Adapun batasan dalam penelitian ini adalah sebagai berikut:")
    add_bullet_num(doc,
        "1. Data yang digunakan berupa rekam medis pasien MDR-TB yang diperoleh dari "
        "RSUD Labuang Baji, Sulawesi Selatan, mencakup variabel demografis, kondisi klinis, dan riwayat "
        "pengobatan pasien.")
    add_bullet_num(doc,
        "2. Teknik SMOTE digunakan sebagai metode penanganan ketidakseimbangan kelas (data "
        "imbalance) pada dataset.")
    add_bullet_num(doc,
        "3. Algoritma Random Forest digunakan sebagai metode klasifikasi utama. Algoritma machine "
        "learning lain tidak digunakan sebagai metode pembanding sehingga penelitian berfokus "
        "pada analisis performa Random Forest setelah penerapan SMOTE.")
    add_bullet_num(doc,
        "4. Tahapan penelitian meliputi preprocessing data, penerapan SMOTE, pembagian data "
        "latih dan data uji, pelatihan model, pengujian model, serta evaluasi performa "
        "menggunakan metrik accuracy, precision, recall, dan F1-score melalui confusion matrix.")
    add_bullet_num(doc,
        "5. Penelitian tidak membahas aspek medis secara mendalam, seperti proses penegakan "
        "diagnosis klinis, pemberian terapi, maupun farmakologi obat. Penelitian hanya berfokus "
        "pada pengembangan model klasifikasi berbasis machine learning.")
    add_bullet_num(doc,
        "6. Hasil penelitian ditujukan untuk kepentingan akademik serta sebagai gambaran potensi "
        "penerapan SMOTE dan Random Forest dalam klasifikasi keberhasilan pengobatan MDR-TB.")
    add_body(doc,
        "Pembatasan ruang lingkup dilakukan agar penelitian tetap terfokus dan hasil yang "
        "diperoleh dapat dianalisis secara mendalam tanpa kehilangan arah. Dengan batasan yang "
        "jelas, setiap tahapan penelitian mulai dari pengolahan data hingga evaluasi model dapat "
        "dilakukan secara sistematis dan hasilnya dapat dipertanggungjawabkan secara ilmiah.")

    add_heading_letter(doc, "F. Sistematika Penulisan")
    add_body(doc,
        "Adapun sistematika penulisan dalam proposal ini adalah sebagai berikut:")
    add_sistematika(doc, "BAB I", "PENDAHULUAN",
        "Bab ini berisi uraian mengenai latar belakang penelitian, rumusan masalah, tujuan "
        "penelitian, manfaat penelitian, ruang lingkup penelitian, serta sistematika penulisan. "
        "Seluruh bagian pada bab ini berfokus pada penerapan teknik SMOTE dan algoritma Random "
        "Forest untuk klasifikasi keberhasilan pengobatan penyakit MDR-TB.")
    add_sistematika(doc, "BAB II", "TINJAUAN PUSTAKA",
        "Bab ini memuat landasan teori yang meliputi konsep MDR-TB, klasifikasi dan machine "
        "learning, ketidakseimbangan data, SMOTE, Random Forest, evaluasi model, dataset rekam "
        "medis, bahasa pemrograman, flowchart, validasi model, serta hyperparameter tuning. "
        "Selain itu, bab ini juga mencakup ringkasan penelitian terdahulu yang relevan serta "
        "penyusunan kerangka berpikir yang mendukung pelaksanaan penelitian.")
    add_sistematika(doc, "BAB III", "METODE PENELITIAN",
        "Bab ini menjelaskan metode yang digunakan dalam penelitian, mulai dari tempat dan waktu "
        "penelitian, alat dan bahan, hingga perancangan sistem yang meliputi desain metode "
        "penelitian serta metode algoritma SMOTE dan Random Forest yang digunakan.")
    add_sistematika(doc, "BAB IV", "HASIL DAN PEMBAHASAN",
        "Bab ini akan menyajikan hasil penelitian yang meliputi proses pengolahan data, "
        "implementasi SMOTE dan Random Forest, hasil evaluasi model, serta analisis terhadap "
        "performa klasifikasi. Pembahasan dilakukan dengan menghubungkan hasil pengujian dengan "
        "teori dan penelitian sebelumnya.")
    add_sistematika(doc, "BAB V", "PENUTUP",
        "Bab ini merupakan bagian akhir yang berisi kesimpulan dari hasil penelitian serta "
        "saran-saran yang dapat dijadikan acuan untuk pengembangan penelitian lebih lanjut, "
        "khususnya dalam penerapan SMOTE dan Random Forest untuk klasifikasi keberhasilan "
        "pengobatan MDR-TB.")

    # ================================================================ BAB II
    add_bab_heading(doc, "II", "TINJAUAN PUSTAKA")

    add_heading_letter(doc, "A. Landasan Teori")

    add_heading_sub(doc, "1. Tuberkulosis Resisten Obat (MDR-TB)")
    add_body(doc,
        "Tuberkulosis merupakan penyakit menular yang disebabkan oleh bakteri Mycobacterium "
        "tuberculosis dan terutama menyerang organ paru-paru. Penyakit ini menular melalui udara "
        "ketika pasien yang terinfeksi batuk, bersin, atau meludah. Pada perkembangannya, muncul "
        "varian galur bakteri yang telah kebal terhadap obat anti-tuberkulosis (OAT) lini pertama, "
        "yang dikenal dengan sebutan multidrug-resistant tuberculosis (MDR-TB). Kondisi ini "
        "terjadi ketika bakteri tidak lagi merespons isoniazid dan rifampisin, dua obat utama yang "
        "menjadi andalan dalam regimen pengobatan tuberkulosis standar (World Health Organization, "
        "2024).")
    add_body(doc,
        "Pasien MDR-TB membutuhkan pengobatan dengan durasi yang lebih panjang, penggunaan obat "
        "lini kedua yang lebih mahal dan lebih toksik, serta pemantauan yang lebih ketat "
        "dibandingkan tuberkulosis yang sensitif obat. Tingkat keberhasilan pengobatan MDR-TB juga "
        "umumnya lebih rendah karena kompleksitas regimen dan besarnya beban efek samping. "
        "Kegagalan dalam mengendalikan MDR-TB tidak hanya membahayakan pasien, tetapi juga "
        "meningkatkan risiko penularan galur super-resisten di masyarakat (World Health "
        "Organization, 2025).")
    add_body(doc,
        "Menurut laporan WHO, sekitar sepertiga dari seluruh kasus baru tuberkulosis yang "
        "terkonfirmasi bakteriologis di dunia merupakan kasus MDR-TB atau tuberkulosis yang "
        "resisten terhadap rifampisin. Tingginya angka tersebut menegaskan pentingnya deteksi "
        "dini dan pengelolaan yang tepat. Data dan informasi yang akurat mengenai karakteristik "
        "pasien sangat dibutuhkan untuk merancang strategi pengobatan yang efektif (World Health "
        "Organization, 2024).")
    add_body(doc,
        "Dalam penelitian ini, status keberhasilan pengobatan MDR-TB menjadi target klasifikasi "
        "yang diprediksi oleh model machine learning. Pemahaman terhadap karakteristik penyakit "
        "ini menjadi dasar penting dalam menentukan variabel yang digunakan, sehingga proses "
        "klasifikasi dapat memberikan hasil yang relevan dan dapat dipertanggungjawabkan secara "
        "ilmiah.")

    add_heading_sub(doc, "2. Klasifikasi dan Machine Learning")
    add_body(doc,
        "Klasifikasi merupakan proses penentuan suatu objek ke dalam kelas atau kategori tertentu "
        "berdasarkan atribut-atributnya. Proses ini diawali dengan membangun aturan atau model "
        "menggunakan data latih (training data) yang telah memiliki label, kemudian model diuji "
        "menggunakan data uji (testing data) untuk menilai kemampuannya mengenali pola baru. "
        "Dalam konteks penelitian ini, klasifikasi digunakan untuk membedakan pasien MDR-TB "
        "berdasarkan hasil akhir pengobatannya (Kurniawati & Saputra, 2023).")
    add_body(doc,
        "Machine learning adalah cabang kecerdasan buatan yang memungkinkan komputer mempelajari "
        "pola dari data dan membuat prediksi tanpa pemrograman eksplisit. Secara umum, machine "
        "learning dapat dikelompokkan menjadi tiga kategori utama, yaitu:")
    add_body(doc,
        "a. Supervised Learning (Pembelajaran Terawasi), yaitu pendekatan yang menggunakan data "
        "berlabel untuk melatih model. Algoritma mempelajari hubungan antara fitur (input) dan "
        "label (output) sehingga mampu memprediksi label data baru yang belum pernah dilihat. "
        "Contoh algoritma dalam kategori ini antara lain decision tree, random forest, naive "
        "bayes, dan support vector machine.")
    add_body(doc,
        "b. Unsupervised Learning (Pembelajaran Tak Terawasi), yaitu pendekatan yang digunakan "
        "pada data tanpa label. Model mencari pola atau struktur tersembunyi dalam data, seperti "
        "pengelompokan (clustering) menggunakan algoritma K-Means atau hierarki.")
    add_body(doc,
        "c. Reinforcement Learning (Pembelajaran Penguatan), yaitu pendekatan di mana agen "
        "belajar melalui interaksi dengan lingkungan dengan menerima reward atau punishment atas "
        "setiap tindakan yang dilakukan.")
    add_body(doc,
        "Setiap algoritma supervised learning memiliki karakteristik yang berbeda dalam hal "
        "kompleksitas, interpretabilitas, dan ketahanan terhadap noise. Pemilihan algoritma yang "
        "tepat sangat bergantung pada karakteristik data dan tujuan penelitian. Pada penelitian "
        "ini, Random Forest dipilih karena kemampuannya menangani data multidimensi serta "
        "menyediakan informasi feature importance yang bermanfaat dalam konteks kesehatan.")
    add_body(doc,
        "Pada penelitian ini, pendekatan supervised learning merupakan kategori yang paling "
        "relevan karena permasalahan klasifikasi keberhasilan pengobatan MDR-TB menggunakan data "
        "berlabel. Variabel klinis dan demografis pasien digunakan sebagai fitur, sedangkan "
        "status keberhasilan pengobatan digunakan sebagai label yang akan diprediksi (Kurniawati "
        "& Saputra, 2023).")
    add_body(doc,
        "Secara umum, alur penerapan machine learning pada penelitian ini mengikuti tahapan "
        "pengumpulan data, preprocessing data, pembagian data latih dan data uji, pelatihan "
        "model, serta evaluasi model. Setiap tahapan memiliki peran yang saling bergantung. "
        "Kualitas hasil akhir sangat dipengaruhi oleh kesiapan data pada tahap awal, sehingga "
        "perhatian khusus diberikan pada proses pengolahan data sebelum model dilatih.")

    add_heading_sub(doc, "3. Ketidakseimbangan Data (Data Imbalance)")
    add_body(doc,
        "Data imbalance adalah kondisi di mana jumlah sampel antar kelas pada dataset sangat "
        "tidak proporsional. Pada dataset rekam medis pasien MDR-TB, jumlah sampel dengan status "
        "\u201cberhasil sembuh\u201d umumnya jauh lebih banyak dibandingkan sampel dengan status "
        "\u201cgagal\u201d, \u201cmeninggal\u201d, atau \u201cputus obat\u201d. Perbandingan "
        "yang sangat timpang ini membuat algoritma klasifikasi kesulitan mempelajari pola kelas "
        "minoritas secara optimal (Chawla et al., 2002).")
    add_body(doc,
        "Beberapa pendekatan dapat digunakan untuk menangani ketidakseimbangan data, antara lain "
        "undersampling yang mengurangi jumlah sampel kelas mayoritas, oversampling yang menambah "
        "jumlah sampel kelas minoritas, serta pendekatan sintetis yang menghasilkan data buatan "
        "baru. Undersampling berisiko membuang informasi penting dari kelas mayoritas, sedangkan "
        "oversampling konvensional berisiko menyebabkan overfitting karena hanya menduplikasi "
        "sampel yang sama secara berulang.")
    add_body(doc,
        "SMOTE merupakan salah satu teknik oversampling sintetis yang paling banyak digunakan "
        "karena mampu menghasilkan sampel baru yang logis tanpa duplikasi langsung. Dalam "
        "penelitian ini, SMOTE digunakan untuk menyeimbangkan kelas pada dataset rekam medis "
        "sebelum proses pelatihan model Random Forest dilakukan, sehingga model dapat belajar "
        "dari representasi data yang lebih merata (Hadjadj dkk, 2024).")
    add_body(doc,
        "Derajat ketidakseimbangan data sering diukur menggunakan rasio antara jumlah sampel "
        "kelas mayoritas dan kelas minoritas. Pada kasus ekstrem, rasio ini dapat mencapai "
        "puluhan bahkan ratusan kali lipat. Semakin tinggi rasio ketidakseimbangan, semakin "
        "sulit algoritma mempelajari pola kelas minoritas, sehingga teknik penyeimbangan "
        "menjadi semakin krusial dalam membangun model yang andal (Chawla et al., 2002).")

    add_heading_sub(doc, "4. SMOTE (Synthetic Minority Over-sampling Technique)")
    add_body(doc,
        "SMOTE adalah teknik oversampling sintetis yang diperkenalkan oleh Chawla et al. (2002) "
        "untuk mengatasi masalah ketidakseimbangan kelas. Berbeda dengan oversampling konvensional "
        "yang menduplikasi sampel minoritas secara langsung, SMOTE menghasilkan sampel sintetis "
        "baru dengan melakukan interpolasi antar sampel minoritas yang berdekatan berdasarkan "
        "k-nearest neighbors. Secara umum, proses kerja SMOTE diawali dengan menentukan sejumlah "
        "tetangga terdekat (k) untuk setiap sampel kelas minoritas.")
    add_body(doc,
        "Tahapan pertama dalam SMOTE adalah menghitung jarak antar sampel kelas minoritas "
        "menggunakan jarak Euclidean. Jarak Euclidean mengukur panjang garis lurus antara dua "
        "titik dalam ruang fitur dan dinyatakan pada persamaan (1) berikut:")
    add_equation_latex(doc,
        "d(x,y)=\\sqrt{\\sum_{i=1}^{n}(x_{i}-y_{i})^{2}}", "1")
    add_body(doc,
        "dengan x dan y adalah dua sampel data, n adalah jumlah fitur, sedangkan x\u1D62 dan y\u1D62 "
        "masing-masing merupakan nilai fitur ke-i dari sampel x dan y.")
    add_body(doc,
        "Tahapan kedua, setelah diperoleh k tetangga terdekat untuk setiap sampel minoritas, "
        "dilakukan pembuatan sampel sintetis. Satu tetangga dipilih secara acak dari k tetangga "
        "tersebut, kemudian dibuat sampel baru di sepanjang garis yang menghubungkan sampel "
        "minoritas dengan tetangga terpilih menggunakan persamaan (2) berikut:")
    add_equation_latex(doc,
        "X_{new}=X_{i}+\\lambda (X_{zi}-X_{i})", "2")
    add_body(doc,
        "dengan X_new adalah sampel sintetis baru, X\u1D62 adalah sampel minoritas yang terpilih, "
        "X_zi adalah tetangga terdekat dari X\u1D62, dan \u03bb adalah bilangan acak dalam interval "
        "(0, 1). Proses ini diulang hingga jumlah sampel kelas minoritas mencapai jumlah yang "
        "diinginkan.")
    add_body(doc,
        "Dalam implementasinya, SMOTE memiliki parameter utama berupa sampling_strategy yang "
        "menentukan jumlah sampel sintetis yang akan dibuat, serta k_neighbors yang menentukan "
        "jumlah tetangga terdekat yang digunakan dalam proses interpolasi. Nilai parameter ini "
        "disesuaikan dengan karakteristik dataset sehingga keseimbangan kelas dapat tercapai "
        "tanpa mengorbankan kualitas data (Lema\u00eetre et al., 2017).")
    add_body(doc,
        "SMOTE memiliki keunggulan dalam menghasilkan sampel sintetis yang beragam dan logis, "
        "namun juga memiliki keterbatasan. Apabila kualitas data awal buruk atau terdapat banyak "
        "noise, SMOTE berpotensi memperkuat pola noise pada kelas minoritas. Oleh karena itu, "
        "kebersihan data melalui preprocessing yang baik menjadi prasyarat penting sebelum "
        "teknik SMOTE diterapkan.")
    add_body(doc,
        "Melalui mekanisme interpolasi tersebut, SMOTE memperkaya ruang fitur kelas minoritas "
        "sehingga model dapat mempelajari pola yang lebih representatif tanpa mengalami "
        "overfitting. Hasil penyeimbangan ini menjadikan model lebih sensitif terhadap pasien "
        "yang berisiko gagal pengobatan. Alur kerja penerapan SMOTE secara umum ditunjukkan pada "
        "Gambar 1.")
    add_figure(doc, "gambar1_alur_smote.png")
    add_caption(doc, "Gambar", 1, "Alur Kerja SMOTE")

    add_heading_sub(doc, "5. Random Forest")
    add_body(doc,
        "Random Forest merupakan algoritma ensemble learning yang diperkenalkan oleh Breiman "
        "(2001). Algoritma ini menggabungkan sejumlah besar pohon keputusan (decision tree) yang "
        "dilatih menggunakan teknik bootstrap aggregating (bagging) serta pemilihan fitur secara "
        "acak pada setiap percabangan. Setiap pohon memberikan suara (voting) terhadap hasil "
        "klasifikasi, dan hasil akhir ditentukan berdasarkan mayoritas suara dari seluruh pohon.")
    add_body(doc,
        "Di samping keunggulannya, Random Forest juga memiliki beberapa keterbatasan, antara lain "
        "ukuran model yang relatif besar dan waktu komputasi yang lebih lama dibandingkan "
        "algoritma yang lebih sederhana. Namun, keterbatasan tersebut dapat diterima karena "
        "dataset pada penelitian ini berukuran relatif kecil hingga menengah, sehingga waktu "
        "komputasi tidak menjadi kendala yang berarti. Interpretabilitas model juga dapat "
        "dibantu melalui pemanfaatan fitur feature importance.")
    add_body(doc,
        "Dalam proses pembentukan pohon keputusan, digunakan pengukuran kemurnian simpul (node "
        "impurity) untuk menentukan fitur pemisah yang paling baik. Dua ukuran yang umum "
        "digunakan adalah indeks Gini dan entropi. Indeks Gini dihitung menggunakan persamaan "
        "(3) berikut:")
    add_equation_latex(doc,
        "\\text{Gini}=1-\\sum_{k=1}^{K}p_{k}^{2}", "3")
    add_body(doc,
        "dengan K adalah jumlah kelas dan p\u2096 adalah proporsi sampel kelas k pada suatu simpul. "
        "Semakin kecil nilai indeks Gini, semakin murni pembagian simpul tersebut. Adapun entropi "
        "dihitung menggunakan persamaan (4) berikut:")
    add_equation_latex(doc,
        "E=-\\sum_{k=1}^{K}p_{k}\\log_{2}p_{k}", "4")
    add_body(doc,
        "dengan p\u2096 adalah proporsi sampel kelas k pada suatu simpul. Nilai entropi yang semakin "
        "kecil menunjukkan simpul yang semakin murni.")
    add_body(doc,
        "Setelah sejumlah T pohon keputusan terbentuk, hasil klasifikasi akhir ditentukan "
        "melalui voting mayoritas (majority voting) dari seluruh pohon menggunakan persamaan "
        "(5) berikut:")
    add_equation_latex(doc,
        "\\hat{y}=\\mathrm{mode}\\{h_{1}(x),h_{2}(x),\\ldots,h_{T}(x)\\}", "5")
    add_body(doc,
        "dengan \u0177 adalah hasil akhir klasifikasi, h\u209C(x) adalah hasil prediksi pohon ke-t "
        "untuk sampel x, dan mode merupakan nilai yang paling sering muncul. Arsitektur Random "
        "Forest secara umum ditunjukkan pada Gambar 2.")
    add_figure(doc, "gambar2_arsitektur_rf.png")
    add_caption(doc, "Gambar", 2, "Arsitektur Random Forest")
    add_body(doc,
        "Keunggulan Random Forest terletak pada kemampuannya menangani data berdimensi tinggi, "
        "stabil terhadap noise dan outlier, serta menyediakan informasi feature importance yang "
        "menunjukkan kontribusi setiap variabel terhadap prediksi. Fitur ini sangat berguna dalam "
        "penelitian medis karena dapat memetakan variabel klinis dan demografis yang paling "
        "berpengaruh terhadap keberhasilan pengobatan pasien MDR-TB (Zhang dkk, 2025).")

    add_heading_sub(doc, "6. Evaluasi Klasifikasi")
    add_body(doc,
        "Evaluasi merupakan tahapan untuk menilai sejauh mana model mampu menghasilkan "
        "klasifikasi yang sesuai dengan kondisi sebenarnya. Pada penelitian ini, evaluasi "
        "dilakukan dengan membandingkan hasil prediksi model terhadap label sebenarnya "
        "menggunakan confusion matrix yang terdiri atas True Positive (TP), True Negative (TN), "
        "False Positive (FP), dan False Negative (FN). True Positive adalah jumlah data positif "
        "yang diprediksi benar, True Negative adalah jumlah data negatif yang diprediksi benar, "
        "False Positive adalah jumlah data negatif yang diprediksi positif, dan False Negative "
        "adalah jumlah data positif yang diprediksi negatif.")
    add_body(doc,
        "Berdasarkan confusion matrix tersebut, performa model diukur menggunakan metrik "
        "accuracy, precision, recall, dan F1-score. Accuracy menggambarkan ketepatan klasifikasi "
        "secara keseluruhan dan dihitung menggunakan persamaan (6) berikut:")
    add_equation_latex(doc,
        "\\text{Accuracy}=\\frac{TP+TN}{TP+TN+FP+FN}", "6")
    add_body(doc,
        "Precision mengukur ketepatan prediksi positif dan dihitung menggunakan persamaan (7) "
        "berikut:")
    add_equation_latex(doc,
        "\\text{Precision}=\\frac{TP}{TP+FP}", "7")
    add_body(doc,
        "Recall mengukur kemampuan model dalam menemukan seluruh data positif dan dihitung "
        "menggunakan persamaan (8) berikut:")
    add_equation_latex(doc,
        "\\text{Recall}=\\frac{TP}{TP+FN}", "8")
    add_body(doc,
        "F1-score memberikan keseimbangan antara precision dan recall dan dihitung menggunakan "
        "persamaan (9) berikut:")
    add_equation_latex(doc,
        "\\text{F1}=2\\cdot \\frac{\\text{Precision}\\cdot \\text{Recall}}{\\text{Precision}+\\text{Recall}}", "9")
    add_body(doc,
        "Pada penelitian ini, recall menjadi perhatian utama karena kemampuan mendeteksi pasien "
        "yang berisiko gagal pengobatan lebih penting daripada akurasi keseluruhan. Model dengan "
        "recall tinggi mampu meminimalkan false negative, yaitu pasien yang berisiko gagal namun "
        "tidak terdeteksi oleh model (Kurniawati & Saputra, 2023).")
    add_body(doc,
        "Interpretasi metrik evaluasi perlu disesuaikan dengan konteks permasalahan. Pada dataset "
        "dengan kelas yang tidak seimbang, accuracy bukanlah metrik yang ideal karena dapat "
        "menyesatkan. Sebagai contoh, model yang memprediksi seluruh sampel sebagai kelas "
        "mayoritas akan memperoleh accuracy yang tinggi padahal sama sekali tidak mampu "
        "mendeteksi kelas minoritas. Oleh karena itu, perhatian utama dalam evaluasi penelitian "
        "ini diberikan pada recall dan F1-score kelas minoritas (Kurniawati & Saputra, 2023).")

    add_heading_sub(doc, "7. Dataset dan Rekam Medis")
    add_body(doc,
        "Rekam medis merupakan dokumen yang memuat catatan dan informasi tentang identitas "
        "pasien, pemeriksaan, diagnosis, pengobatan, serta pelayanan lain yang diberikan kepada "
        "pasien. Data rekam medis memiliki potensi besar sebagai sumber analisis dalam bidang "
        "kesehatan, termasuk untuk membangun model prediktif berbasis machine learning. "
        "Kualitas dataset sangat menentukan kinerja model, sehingga diperlukan proses "
        "preprocessing untuk membersihkan data, menangani nilai yang hilang, dan memastikan "
        "kesesuaian format data (Chawla et al., 2002).")
    add_body(doc,
        "Pada penelitian ini, dataset yang digunakan adalah data rekam medis pasien MDR-TB yang "
        "diperoleh dari RSUD Labuang Baji, Sulawesi Selatan. Data mencakup variabel demografis seperti "
        "usia dan jenis kelamin, variabel klinis seperti komorbiditas dan tingkat keparahan "
        "penyakit, serta riwayat pengobatan pasien. Seluruh variabel tersebut digunakan sebagai "
        "fitur dalam model klasifikasi untuk memprediksi keberhasilan pengobatan pasien MDR-TB.")
    add_body(doc,
        "Beberapa contoh atribut yang umum tersedia pada rekam medis pasien MDR-TB antara lain "
        "usia, jenis kelamin, status gizi, komorbiditas, hasil pemeriksaan dahak, lama menderita "
        "tuberkulosis sebelum diagnosis MDR-TB, riwayat pengobatan sebelumnya, serta hasil akhir "
        "pengobatan. Atribut hasil akhir pengobatan inilah yang menjadi label kelas dalam proses "
        "klasifikasi, sedangkan atribut lainnya berperan sebagai fitur yang memengaruhi prediksi.")

    add_heading_sub(doc, "8. Bahasa Pemrograman dan Tools")
    add_body(doc,
        "Bahasa pemrograman yang digunakan dalam penelitian ini adalah Python. Python dipilih "
        "karena memiliki ekosistem pustaka yang lengkap untuk analisis data dan machine learning, "
        "serta sintaks yang sederhana sehingga memudahkan peneliti dalam mengolah data, membangun "
        "model, dan mengevaluasi hasil klasifikasi. Python juga memiliki komunitas pengguna yang "
        "besar sehingga dokumentasi dan dukungan pengembangannya sangat memadai.")
    add_body(doc,
        "Pustaka Pandas digunakan untuk membaca, membersihkan, dan mengelola dataset yang "
        "disimpan dalam format Excel (XLSX). NumPy digunakan untuk mendukung perhitungan "
        "numerik, sedangkan pustaka scikit-learn digunakan untuk membangun model Random Forest "
        "dan menghitung metrik evaluasi (Pedregosa et al., 2011). Teknik SMOTE diimplementasikan "
        "menggunakan pustaka imbalanced-learn yang menyediakan implementasi SMOTE secara siap "
        "pakai (Lema\u00eetre et al., 2017).")
    add_body(doc,
        "Untuk kebutuhan visualisasi, digunakan pustaka Matplotlib dan Seaborn yang mampu "
        "menghasilkan grafik dan diagram yang informatif. Seluruh proses pengembangan dilakukan "
        "menggunakan Visual Studio Code sebagai lingkungan pengembangan (integrated development "
        "environment), yang mendukung penulisan kode, pengujian, dan debugging secara "
        "terintegrasi.")
    add_body(doc,
        "Seluruh pustaka yang digunakan bersifat open source dan tersedia secara luas, sehingga "
        "hasil penelitian dapat direproduksi oleh peneliti lain dengan mudah. Dokumentasi resmi "
        "masing-masing pustaka juga sangat lengkap, yang memudahkan proses pengembangan dan "
        "penyelesaian kendala teknis selama penelitian berlangsung.")

    add_heading_sub(doc, "9. Flowchart")
    add_body(doc,
        "Flowchart adalah representasi grafis yang menggambarkan alur atau langkah-langkah dalam "
        "suatu proses secara terstruktur menggunakan simbol-simbol baku. Diagram ini membantu "
        "menjelaskan urutan aktivitas, keputusan, dan hubungan antarproses sehingga memudahkan "
        "peneliti maupun pembaca dalam memahami alur penelitian. Penggunaan flowchart sangat "
        "penting dalam perancangan sistem karena mampu mengkomunikasikan logika proses secara "
        "jelas dan sistematis.")
    add_body(doc,
        "Simbol-simbol yang umum digunakan dalam flowchart ditunjukkan pada Tabel 1.")
    add_body(doc,
        "Berdasarkan orientasinya, flowchart dapat berbentuk vertikal yang mengalir dari atas ke "
        "bawah, horizontal yang mengalir dari kiri ke kanan, atau campuran keduanya. Pada "
        "penelitian ini, flowchart metode disusun secara vertikal untuk memudahkan pembacaan "
        "urutan langkah penelitian yang bersifat sekuensial.")
    tabel_simbol_header = ["Simbol", "Nama Simbol", "Fungsi"]
    tabel_simbol_data = [
        ["(elips)", "Terminator", "Menunjukkan awal (start) dan akhir (end) dari suatu proses atau algoritma."],
        ["(persegi)", "Process (Proses)", "Menunjukkan aktivitas atau langkah kerja yang dilakukan dalam suatu proses."],
        ["(belah ketupat)", "Decision (Keputusan)", "Digunakan untuk menunjukkan proses pengambilan keputusan berdasarkan suatu kondisi, misalnya Ya atau Tidak."],
        ["(jajar genjang)", "Input / Output", "Menunjukkan proses memasukkan data (input) atau menghasilkan keluaran (output)."],
        ["(lingkaran kecil)", "Connector", "Digunakan untuk menghubungkan alur yang terputus atau berpindah ke bagian lain dari flowchart."],
        ["(garis panah)", "Flow Line", "Menunjukkan arah aliran proses dari satu simbol ke simbol lainnya."],
    ]
    add_caption(doc, "Tabel", 1, "Simbol Flowchart")
    add_table(doc, rows=1 + len(tabel_simbol_data), cols=3,
              header=tabel_simbol_header, data=tabel_simbol_data,
              widths=[3.5, 4.0, 6.5])

    add_heading_sub(doc, "10. Validasi Model dan Cross-Validation")
    add_body(doc,
        "Validasi model merupakan tahapan penting untuk menilai kemampuan model dalam "
        "menggeneralisasi pola pada data yang belum pernah dilihat. Model yang hanya berkinerja "
        "baik pada data latih tetapi buruk pada data baru disebut mengalami overfitting. Untuk "
        "mengatasi hal ini, data dibagi menjadi data latih (training set) dan data uji (testing "
        "set) dengan proporsi yang wajar, misalnya 80% untuk data latih dan 20% untuk data uji.")
    add_body(doc,
        "Selain pembagian data secara sederhana, teknik k-fold cross-validation dapat digunakan "
        "untuk memperoleh estimasi performa model yang lebih stabil. Pada teknik ini, data "
        "dibagi menjadi k bagian (fold) yang sama besar. Model dilatih pada k\u22121 fold dan "
        "diuji pada 1 fold yang tersisa. Proses ini diulang sebanyak k kali dengan fold pengujian "
        "yang berbeda, kemudian hasilnya dirata-ratakan.")
    add_body(doc,
        "Pada penelitian ini, proses pembagian data dilakukan sebelum penerapan SMOTE. Hal ini "
        "bertujuan untuk mencegah terjadinya kebocoran data (data leakage), yaitu kondisi di mana "
        "informasi dari data uji ikut memengaruhi proses pelatihan model. SMOTE hanya diterapkan "
        "pada data latih sehingga evaluasi model pada data uji tetap mencerminkan kondisi riil.")
    add_body(doc,
        "Metode holdout (pembagian tunggal menjadi data latih dan data uji) dipilih sebagai "
        "pendekatan utama karena sederhana dan sesuai dengan karakteristik data penelitian. "
        "Teknik k-fold cross-validation digunakan sebagai pendekatan pendukung untuk memperoleh "
        "gambaran stabilitas performa model pada berbagai pembagian data.")

    add_heading_sub(doc, "11. Hyperparameter Tuning")
    add_body(doc,
        "Kinerja algoritma Random Forest sangat dipengaruhi oleh pemilihan hyperparameter. "
        "Beberapa hyperparameter penting pada Random Forest antara lain n_estimators (jumlah "
        "pohon keputusan yang dibangun), max_depth (kedalaman maksimum setiap pohon), "
        "min_samples_split (jumlah minimum sampel untuk melakukan percabangan), serta "
        "random_state (nilai acak untuk memastikan hasil dapat direproduksi).")
    add_body(doc,
        "Untuk memperoleh kombinasi hyperparameter yang optimal, digunakan teknik grid search "
        "yang menguji seluruh kombinasi nilai pada ruang parameter yang telah ditentukan. "
        "Kombinasi yang menghasilkan performa terbaik berdasarkan metrik evaluasi kemudian "
        "dipilih sebagai konfigurasi model akhir. Dengan tuning hyperparameter yang tepat, "
        "performa model diharapkan dapat meningkat secara signifikan tanpa mengalami overfitting "
        "(Zhang dkk, 2025).")

    add_heading_sub(doc, "12. Confusion Matrix")
    add_body(doc,
        "Confusion matrix merupakan tabel yang menggambarkan kinerja model klasifikasi terhadap "
        "data uji. Matriks ini membandingkan hasil prediksi model dengan nilai sebenarnya "
        "(aktual) dan disusun dalam empat komponen utama, yaitu True Positive (TP), True "
        "Negative (TN), False Positive (FP), dan False Negative (FN). Keempat komponen tersebut "
        "menjadi dasar perhitungan berbagai metrik evaluasi seperti accuracy, precision, recall, "
        "dan F1-score (Kurniawati & Saputra, 2023).")
    add_body(doc,
        "Pada penelitian ini, kelas yang dianggap sebagai kelas positif adalah pasien dengan "
        "status gagal pengobatan (kelas minoritas). Dengan demikian, False Negative (FN), yaitu "
        "pasien yang berisiko gagal tetapi diprediksi berhasil oleh model, menjadi kesalahan "
        "yang paling dihindari. Dengan memahami struktur confusion matrix, hasil evaluasi dapat "
        "diinterpretasikan secara lebih mendalam dan tidak hanya berfokus pada satu metrik "
        "tunggal.")

    add_heading_sub(doc, "13. Feature Importance")
    add_body(doc,
        "Feature importance merupakan salah satu keunggulan utama algoritma Random Forest. Fitur "
        "ini memberikan nilai pentingnya setiap variabel terhadap hasil prediksi model. Semakin "
        "besar nilai feature importance suatu variabel, semakin besar pengaruh variabel tersebut "
        "terhadap keputusan klasifikasi. Informasi ini diperoleh dari rata-rata penurunan "
        "kemurnian simpul (impurity) yang diakibatkan oleh setiap variabel pada seluruh pohon "
        "dalam hutan (Breiman, 2001).")
    add_body(doc,
        "Dalam konteks penelitian ini, feature importance dapat digunakan untuk mengidentifikasi "
        "faktor-faktor yang paling berpengaruh terhadap keberhasilan pengobatan pasien MDR-TB, "
        "misalnya usia, komorbiditas, atau riwayat pengobatan sebelumnya. Informasi ini sangat "
        "berharga bagi tenaga medis untuk memahami determinan keberhasilan pengobatan secara "
        "lebih komprehensif sekaligus menjadi bahan evaluasi terhadap kualitas data yang "
        "digunakan (Zhang dkk, 2025).")

    add_heading_sub(doc, "14. Preprocessing Data")
    add_body(doc,
        "Preprocessing data merupakan tahapan penting sebelum data digunakan dalam pelatihan "
        "model. Tahapan ini bertujuan membersihkan dan menyiapkan data agar sesuai dengan "
        "kebutuhan algoritma. Tanpa preprocessing yang baik, kualitas model akan menurun "
        "meskipun algoritma yang digunakan sangat canggih. Oleh karena itu, preprocessing "
        "menjadi fondasi yang menentukan keberhasilan seluruh proses klasifikasi.")
    add_body(doc,
        "Langkah pertama dalam preprocessing adalah data cleaning, yaitu penanganan nilai yang "
        "hilang (missing values), data ganda (duplikat), dan data yang tidak konsisten. Nilai "
        "yang hilang dapat diatasi dengan menghapus baris data atau mengisi dengan nilai "
        "rata-rata (mean) untuk data numerik atau nilai yang paling sering muncul (mode) untuk "
        "data kategorikal.")
    add_body(doc,
        "Langkah berikutnya adalah transformasi data, yaitu encoding data kategorikal menjadi "
        "bentuk numerik yang dapat diproses oleh algoritma. Normalisasi juga dilakukan untuk "
        "menyamakan skala antar fitur sehingga tidak ada fitur yang mendominasi proses "
        "perhitungan jarak, khususnya pada penerapan SMOTE yang menggunakan jarak Euclidean "
        "dalam menentukan tetangga terdekat (Hadjadj dkk, 2024).")

    add_heading_sub(doc, "15. Decision Tree sebagai Dasar Random Forest")
    add_body(doc,
        "Random Forest dibangun di atas kumpulan pohon keputusan (decision tree). Pohon "
        "keputusan merupakan struktur yang menyerupai diagram alir, di mana setiap simpul "
        "internal (internal node) merepresentasikan pengujian terhadap suatu fitur, setiap cabang "
        "merepresentasikan hasil pengujian, dan setiap daun (leaf) merepresentasikan label "
        "kelas. Pembentukan pohon dilakukan dengan memilih fitur pemisah yang menghasilkan "
        "pembagian data yang paling murni (Breiman, 2001).")
    add_body(doc,
        "Pemilihan fitur pemisah dilakukan menggunakan ukuran kemurnian simpul seperti indeks "
        "Gini dan entropi, sebagaimana dijelaskan pada persamaan (3) dan (4). Semakin murni "
        "pembagian yang dihasilkan oleh suatu fitur, semakin baik fitur tersebut dijadikan "
        "pemisah pada simpul yang bersangkutan. Proses pembagian dilakukan secara rekursif "
        "hingga tercapai kondisi berhenti tertentu, misalnya kedalaman maksimum atau jumlah "
        "sampel minimum pada simpul.")
    add_body(doc,
        "Meskipun pohon keputusan tunggal mudah diinterpretasikan, algoritma ini cenderung "
        "mengalami overfitting ketika pohon tumbuh terlalu dalam. Random Forest mengatasi "
        "kelemahan ini dengan membangun banyak pohon secara acak dan menggabungkan hasilnya "
        "melalui voting, sehingga varians model dapat dikurangi secara signifikan tanpa "
        "meningkatkan bias secara berlebihan (Breiman, 2001).")

    add_heading_sub(doc, "16. Penanganan Alternatif Ketidakseimbangan Data")
    add_body(doc,
        "Selain SMOTE, terdapat beberapa teknik lain yang umum digunakan untuk menangani "
        "ketidakseimbangan data. Undersampling mengurangi jumlah sampel kelas mayoritas secara "
        "acak agar seimbang dengan kelas minoritas. Pendekatan ini sederhana dan cepat, namun "
        "berisiko membuang informasi penting yang terkandung pada sampel kelas mayoritas "
        "sehingga dapat menurunkan performa model.")
    add_body(doc,
        "Oversampling konvensional (random oversampling) menambah jumlah sampel kelas minoritas "
        "dengan menduplikasi sampel yang sudah ada secara acak. Pendekatan ini menyeimbangkan "
        "distribusi kelas, tetapi cenderung menyebabkan overfitting karena model mempelajari "
        "pola yang sama secara berulang tanpa variasi.")
    add_body(doc,
        "Pendekatan hibrida menggabungkan teknik undersampling dan oversampling sekaligus, "
        "misalnya SMOTE-Tomek dan SMOTE-ENN. SMOTE-Tomek menerapkan SMOTE kemudian membersihkan "
        "sampel yang dianggap sebagai noise menggunakan metode Tomek link, sedangkan SMOTE-ENN "
        "menggabungkan SMOTE dengan pembersihan menggunakan k-nearest neighbors (Hadjadj dkk, "
        "2024).")
    add_body(doc,
        "Pada penelitian ini, SMOTE dipilih karena mampu menghasilkan sampel sintetis baru yang "
        "logis dan dapat mengurangi risiko overfitting dibandingkan oversampling konvensional. "
        "Pendekatan lain seperti SMOTE-ENN dapat menjadi bahan pengembangan penelitian "
        "selanjutnya untuk membandingkan efektivitas berbagai teknik penyeimbangan data.")

    add_heading_sub(doc, "17. Data Mining dan CRISP-DM")
    add_body(doc,
        "Data mining merupakan proses menemukan pola yang menarik dan bermanfaat dari data "
        "dalam jumlah besar. Dalam konteks kesehatan, data mining digunakan untuk menganalisis "
        "rekam medis, menemukan hubungan antar variabel klinis, serta membangun model prediktif "
        "untuk mendukung diagnosis dan prognosis penyakit (Chawla et al., 2002).")
    add_body(doc,
        "Salah satu metodologi standar dalam pengembangan proyek data mining adalah CRISP-DM "
        "(Cross-Industry Standard Process for Data Mining). CRISP-DM terdiri atas enam fase, "
        "yaitu business understanding, data understanding, data preparation, modeling, "
        "evaluation, dan deployment. Kerangka kerja ini memberikan panduan sistematis agar "
        "proses pengolahan data hingga penerapan model berjalan terstruktur dan terdokumentasi "
        "dengan baik.")
    add_body(doc,
        "Alur penelitian ini sejalan dengan fase-fase CRISP-DM. Fase pemahaman masalah dan "
        "tujuan penelitian berkaitan dengan business understanding, pengumpulan dan eksplorasi "
        "data rekam medis berkaitan dengan data understanding, preprocessing data berkaitan "
        "dengan data preparation, penerapan SMOTE dan Random Forest berkaitan dengan modeling, "
        "evaluasi metrik klasifikasi berkaitan dengan evaluation, serta penyusunan laporan "
        "berkaitan dengan deployment.")

    add_heading_sub(doc, "18. Perbandingan Algoritma Klasifikasi")
    add_body(doc,
        "Terdapat berbagai algoritma klasifikasi yang dapat digunakan untuk memprediksi hasil "
        "pengobatan, antara lain decision tree, naive bayes, k-nearest neighbors (k-NN), "
        "support vector machine (SVM), dan random forest. Setiap algoritma memiliki kelebihan "
        "dan kekurangan yang bergantung pada karakteristik data dan tujuan penelitian.")
    add_body(doc,
        "Decision tree mudah diinterpretasikan tetapi rentan terhadap overfitting. Naive bayes "
        "cepat dan sederhana namun mengasumsikan independensi antar fitur yang jarang terpenuhi "
        "pada data medis. k-NN sederhana namun sensitif terhadap skala fitur dan biaya "
        "komputasi yang tinggi pada data besar. SVM efektif pada data berdimensi tinggi tetapi "
        "sulit diinterpretasikan.")
    add_body(doc,
        "Random Forest dipilih dalam penelitian ini karena mampu menangani data multidimensi, "
        "stabil terhadap noise dan outlier, tahan terhadap overfitting melalui mekanisme "
        "ensemble, serta menyediakan informasi feature importance yang sangat bermanfaat dalam "
        "analisis data kesehatan (Breiman, 2001; Zhang dkk, 2025).")

    add_heading_sub(doc, "19. Etika dan Privasi Data Kesehatan")
    add_body(doc,
        "Penggunaan data kesehatan dalam penelitian wajib memerhatikan aspek etika dan privasi. "
        "Data rekam medis merupakan data yang bersifat sensitif dan dilindungi oleh peraturan "
        "perundang-undangan. Peneliti harus menjaga kerahasiaan identitas pasien dan hanya "
        "menggunakan data yang relevan dengan keperluan penelitian.")
    add_body(doc,
        "Pada penelitian ini, seluruh proses pengolahan data dilakukan dengan mengedepankan "
        "prinsip keamanan dan kerahasiaan. Identitas pasien tidak ditampilkan dalam laporan, "
        "dan data diproses secara agregat sesuai tujuan analisis. Hal ini sejalan dengan "
        "ketentuan etika penelitian kesehatan yang berlaku.")
    add_body(doc,
        "Selain itu, izin penggunaan data diperoleh melalui mekanisme resmi dari pihak "
        "RSUD Labuang Baji, Sulawesi Selatan. Dokumentasi perizinan menjadi bagian penting dalam "
        "menjamin keabsahan data dan keberlanjutan penelitian.")

    add_heading_sub(doc, "20. Sistem Pendukung Keputusan Berbasis Machine Learning")
    add_body(doc,
        "Sistem pendukung keputusan (SPK) merupakan sistem berbasis komputer yang membantu "
        "pengambil keputusan dalam mengevaluasi berbagai alternatif dan membuat keputusan yang "
        "lebih baik. Dalam bidang kesehatan, SPK berbasis machine learning dapat membantu tenaga "
        "medis dalam mendiagnosis penyakit, memprediksi hasil pengobatan, serta merekomendasikan "
        "tindakan penanganan.")
    add_body(doc,
        "Model klasifikasi yang dihasilkan pada penelitian ini berpotensi dikembangkan menjadi "
        "SPK untuk mendukung tenaga medis dalam mengidentifikasi pasien MDR-TB yang berisiko "
        "gagal pengobatan. Melalui prediksi yang cepat dan berbasis data, intervensi medis dapat "
        "dilakukan lebih awal sehingga peluang keberhasilan pengobatan meningkat.")
    add_body(doc,
        "Pemanfaatan SPK tidak dimaksudkan untuk menggantikan peran tenaga medis, melainkan "
        "sebagai alat bantu yang menyediakan informasi tambahan dalam proses pengambilan "
        "keputusan. Keputusan akhir tetap berada pada tenaga medis yang memiliki pengetahuan "
        "klinis dan pengalaman dalam menangani pasien (Zhang dkk, 2025).")

    add_heading_letter(doc, "B. Penelitian Terkait")
    add_body(doc,
        "Penulis menggali banyak inspirasi dan referensi untuk menyusun proposal ini melalui "
        "penelitian sebelumnya yang terkait dengan permasalahan yang diteliti. Penelitian "
        "sebelumnya yang menjadi fokus ditunjukkan pada Tabel 2.")
    terkait_header = ["Peneliti & Tahun", "Judul", "Metode", "Hasil"]
    terkait_data = [
        ["[Peneliti 1, 20xx]",
         "Klasifikasi Keberhasilan Pengobatan Pasien Tuberkulosis Paru Menggunakan Algoritma Random Forest",
         "Random Forest",
         "Klasifikasi tingkat keberhasilan pengobatan pasien tuberkulosis paru berbasis data klinis."],
        ["[Peneliti 2, 20xx]",
         "Prediksi Hasil Pengobatan Pasien MDR-TB Menggunakan Pendekatan Machine Learning",
         "Machine Learning",
         "Prediksi hasil akhir pengobatan pasien MDR-TB berdasarkan data klinis dan rekam medis."],
        ["Chawla et al. (2002)",
         "SMOTE: Synthetic Minority Over-sampling Technique",
         "SMOTE",
         "Menghasilkan sampel sintetis untuk menyeimbangkan kelas minoritas menggunakan interpolasi k-NN."],
        ["Breiman (2001)",
         "Random Forests",
         "Ensemble Decision Tree",
         "Metode ensemble yang menggabungkan banyak pohon keputusan untuk klasifikasi yang akurat dan stabil."],
    ]
    add_caption(doc, "Tabel", 2, "Penelitian Terkait")
    add_table(doc, rows=1 + len(terkait_data), cols=4,
              header=terkait_header, data=terkait_data,
              widths=[2.6, 4.6, 2.4, 4.4])
    add_body(doc,
        "Persamaan antara penelitian pertama dengan penelitian ini terletak pada fokus "
        "pengelompokan tingkat keberhasilan pengobatan pasien tuberkulosis menggunakan algoritma "
        "Random Forest dengan data klinis. Perbedaannya terletak pada penggunaan teknik "
        "penanganan ketidakseimbangan data; penelitian ini mengintegrasikan teknik SMOTE untuk "
        "meningkatkan performa model pada kelas minoritas, sedangkan penelitian tersebut "
        "menggunakan data asli tanpa proses penyeimbangan sampel.")
    add_body(doc,
        "Penelitian kedua memiliki kesamaan dalam objek penelitian, yaitu klasifikasi hasil "
        "akhir pengobatan pada pasien MDR-TB berdasarkan data klinis dan rekam medis pasien. "
        "Perbedaannya terletak pada metodologi penanganan data; penelitian tersebut menggunakan "
        "dataset apa adanya tanpa teknik penyeimbangan, sedangkan penelitian ini mengintegrasikan "
        "teknik SMOTE untuk mengatasi ketidakseimbangan data antara pasien yang berhasil sembuh "
        "dan yang gagal pengobatan guna meningkatkan sensitivitas model.")
    add_body(doc,
        "Penelitian Chawla et al. (2002) menjadi dasar teoretis utama dalam penelitian ini. "
        "Kedua penelitian menggunakan teknik SMOTE untuk menangani ketidakseimbangan kelas. "
        "Perbedaannya, penelitian Chawla et al. (2002) berfokus pada pengembangan metode SMOTE "
        "secara umum pada berbagai dataset, sedangkan penelitian ini menerapkan SMOTE secara "
        "spesifik pada data rekam medis pasien MDR-TB yang dikombinasikan dengan Random Forest "
        "sebagai model klasifikasi.")
    add_body(doc,
        "Penelitian Breiman (2001) menjadi landasan penggunaan Random Forest dalam penelitian "
        "ini. Kedua penelitian menggunakan Random Forest sebagai algoritma klasifikasi utama. "
        "Perbedaannya, penelitian Breiman (2001) berfokus pada pengembangan teoretis algoritma "
        "Random Forest, sedangkan penelitian ini menerapkan Random Forest pada permasalahan "
        "kesehatan yang spesifik dengan didahului penanganan ketidakseimbangan data menggunakan "
        "SMOTE.")

    add_heading_letter(doc, "C. Kerangka Berpikir")
    add_body(doc,
        "Kerangka berpikir menggambarkan alur logis penelitian yang dimulai dari permasalahan "
        "yang dihadapi hingga hasil yang diharapkan. Kerangka ini membantu menjelaskan hubungan "
        "antara data rekam medis pasien MDR-TB, metode yang digunakan, serta keluaran dari "
        "proses klasifikasi keberhasilan pengobatan. Adapun kerangka berpikir penelitian ini "
        "adalah sebagai berikut:")
    add_figure(doc, "gambar5_kerangka_berpikir.png")
    add_caption(doc, "Gambar", 3, "Kerangka Berpikir")
    kerangka = [
        ("Masalah",
         "Tingginya angka kegagalan pengobatan MDR-TB dan sulitnya tenaga medis mendeteksi "
         "potensi kegagalan secara dini karena data rekam medis yang tersedia tidak seimbang."),
        ("Tujuan",
         "Mengembangkan model klasifikasi yang mampu memprediksi keberhasilan pengobatan pasien "
         "MDR-TB secara akurat, khususnya pada kelas minoritas yang berisiko gagal."),
        ("Data",
         "Data rekam medis pasien MDR-TB dari RSUD Labuang Baji, Sulawesi Selatan yang mencakup "
         "karakteristik demografis, klinis, dan riwayat pengobatan pasien."),
        ("Solusi",
         "Penerapan teknik SMOTE untuk menyeimbangkan dataset dan algoritma Random Forest "
         "sebagai model klasifikasi."),
        ("Metode",
         "Data melalui preprocessing, kemudian diterapkan SMOTE dan dibagi menjadi data latih "
         "dan data uji. Selanjutnya model dilatih menggunakan Random Forest dan dievaluasi "
         "menggunakan metrik klasifikasi."),
        ("Hasil",
         "Model yang akurat dan sensitif terhadap pasien berisiko gagal pengobatan, yang "
         "berpotensi menjadi sistem pendukung keputusan bagi tenaga medis."),
    ]
    for judul, isi in kerangka:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(6)
        r = p.add_run(judul)
        set_run(r, bold=True)
        add_body(doc, isi, indent=False, justify=True)

    # ================================================================ BAB III
    add_bab_heading(doc, "III", "METODE PENELITIAN")

    add_heading_letter(doc, "A. Tempat dan Waktu Penelitian")
    add_body(doc,
        "Penelitian ini dilaksanakan di RSUD Labuang Baji, Sulawesi Selatan, yang merupakan tempat "
        "pengambilan data rekam medis pasien MDR-TB. Lokasi penelitian dipilih karena memiliki "
        "jumlah kasus MDR-TB yang memadai serta akses data rekam medis yang mendukung kebutuhan "
        "analisis penelitian. Penelitian ini dilaksanakan pada bulan [bulan] tahun [tahun] "
        "hingga [bulan] tahun [tahun].")
    add_body(doc,
        "Pengambilan data dilakukan melalui mekanisme resmi yang disetujui oleh pihak instansi. "
        "Identitas pasien dijaga kerahasiaannya sesuai dengan ketentuan etika penelitian dan "
        "regulasi perlindungan data kesehatan. Data yang digunakan hanya mencakup variabel yang "
        "relevan dengan keperluan klasifikasi keberhasilan pengobatan, tanpa menyertakan "
        "informasi pribadi pasien yang bersifat sensitif.")
    add_body(doc,
        "Secara keseluruhan, penelitian ini dilaksanakan dalam jangka waktu enam bulan yang "
        "terdiri atas tahapan studi literatur, pengumpulan data, pengolahan data, implementasi "
        "model, pengujian, hingga penyusunan laporan. Rincian kegiatan dan alokasi waktu "
        "penelitian ditunjukkan pada Tabel 3.")
    jadwal_header = ["No", "Kegiatan", "Bulan 1", "Bulan 2", "Bulan 3", "Bulan 4", "Bulan 5", "Bulan 6"]
    jadwal_data = [
        ["1", "Studi literatur dan penyusunan proposal", "\u221a", "\u221a", "", "", "", ""],
        ["2", "Pengumpulan data rekam medis", "", "\u221a", "\u221a", "", "", ""],
        ["3", "Preprocessing data", "", "", "\u221a", "\u221a", "", ""],
        ["4", "Implementasi SMOTE dan Random Forest", "", "", "", "\u221a", "\u221a", ""],
        ["5", "Pengujian dan evaluasi model", "", "", "", "", "\u221a", "\u221a"],
        ["6", "Penyusunan laporan akhir", "", "", "", "", "\u221a", "\u221a"],
    ]
    add_caption(doc, "Tabel", 3, "Jadwal Penelitian")
    add_table(doc, rows=1 + len(jadwal_data), cols=8,
              header=jadwal_header, data=jadwal_data,
              widths=[0.9, 5.0, 1.15, 1.15, 1.15, 1.15, 1.15, 1.15])

    add_heading_letter(doc, "B. Alat dan Bahan")
    add_body(doc,
        "Alat dan bahan yang digunakan dalam penelitian ini meliputi kebutuhan perangkat keras "
        "(hardware) dan perangkat lunak (software) sebagai berikut:")
    add_bullet_num(doc, "1. Kebutuhan Hardware (Perangkat Keras)")
    add_bullet_num(doc, "a. Notebook [Merk/Spesifikasi]", left=2.3, hang=1.0)
    add_bullet_num(doc, "b. Prosesor minimal Intel Core i3", left=2.3, hang=1.0)
    add_bullet_num(doc, "c. RAM minimal 8 GB", left=2.3, hang=1.0)
    add_bullet_num(doc, "2. Kebutuhan Software (Perangkat Lunak)")
    add_bullet_num(doc, "a. Python", left=2.3, hang=1.0)
    add_bullet_num(doc, "b. Pustaka Pandas dan NumPy", left=2.3, hang=1.0)
    add_bullet_num(doc, "c. Pustaka scikit-learn", left=2.3, hang=1.0)
    add_bullet_num(doc, "d. Pustaka imbalanced-learn (SMOTE)", left=2.3, hang=1.0)
    add_bullet_num(doc, "e. Pustaka Matplotlib dan Seaborn", left=2.3, hang=1.0)
    add_bullet_num(doc, "f. Microsoft Excel", left=2.3, hang=1.0)
    add_bullet_num(doc, "g. Visual Studio Code", left=2.3, hang=1.0)
    add_body(doc,
        "Spesifikasi hardware tersebut dipilih karena mampu menjalankan proses pelatihan model "
        "Random Forest dan penerapan SMOTE tanpa kendala yang berarti. Proses komputasi yang "
        "dilakukan bersifat moderat dan tidak memerlukan GPU khusus, sehingga notebook standar "
        "sudah mencukupi untuk kebutuhan penelitian ini. Dengan demikian, seluruh tahapan "
        "penelitian dapat dilaksanakan secara mandiri oleh peneliti.")
    add_body(doc,
        "Pemilihan perangkat lunak tersebut disesuaikan dengan kebutuhan penelitian. Python "
        "menjadi bahasa utama karena dukungannya terhadap berbagai pustaka machine learning. "
        "Pustaka imbalanced-learn dipilih karena menyediakan implementasi SMOTE yang siap pakai, "
        "sedangkan scikit-learn menyediakan implementasi Random Forest serta berbagai metrik "
        "evaluasi yang dibutuhkan (Pedregosa et al., 2011; Lema\u00eetre et al., 2017).")

    add_heading_letter(doc, "C. Perancangan Sistem")
    add_body(doc,
        "Perancangan sistem adalah proses perencanaan langkah-langkah penelitian yang bertujuan "
        "membangun model klasifikasi keberhasilan pengobatan pasien MDR-TB berbasis machine "
        "learning menggunakan teknik SMOTE dan algoritma Random Forest. Tujuan perancangan ini "
        "adalah memastikan seluruh tahapan penelitian berjalan terstruktur dan sistematis mulai "
        "dari pengolahan data hingga evaluasi model. Diharapkan hasil perancangan ini mampu "
        "menghasilkan model klasifikasi yang akurat serta layak digunakan sebagai alat bantu "
        "dalam mendukung pengambilan keputusan klinis berbasis data.")
    add_body(doc,
        "Perancangan sistem dilakukan secara bertahap dan terdokumentasi. Setiap tahapan "
        "dirancang berdasarkan landasan teori yang telah diuraikan pada Bab II, sehingga "
        "kesesuaian antara teori, metode, dan implementasi dapat dijaga secara konsisten "
        "sepanjang penelitian berlangsung.")

    add_heading_sub(doc, "1. Desain Metode Penelitian")
    add_body(doc,
        "Desain metode penelitian menggambarkan alur keseluruhan proses penelitian, mulai dari "
        "pengumpulan data hingga diperolehnya hasil akhir. Alur penelitian ini ditunjukkan pada "
        "Gambar 4.")
    add_figure(doc, "gambar3_desain_metode.png")
    add_caption(doc, "Gambar", 4, "Desain Metode Penelitian")
    tahapan = [
        ("a. Mulai",
         "Tahap awal dimulai dengan menentukan tujuan dan permasalahan penelitian, yaitu "
         "membangun model klasifikasi keberhasilan pengobatan pasien MDR-TB menggunakan teknik "
         "SMOTE dan algoritma Random Forest.",
         "Pada tahap ini juga dilakukan studi literatur terhadap penelitian sebelumnya serta "
         "pemahaman terhadap karakteristik data yang akan digunakan, sehingga arah penelitian "
         "menjadi jelas dan terukur."),
        ("b. Pengumpulan Data",
         "Pengambilan data rekam medis pasien MDR-TB dari RSUD Labuang Baji, Sulawesi Selatan. "
         "Data yang dikumpulkan mencakup variabel usia, kategori usia, jenis kelamin, status "
         "bekerja, riwayat pengobatan sebelumnya, efek samping obat, panduan pengobatan, tinggi "
         "badan, indeks massa tubuh (IMT), status gizi, kepatuhan minum obat, status merokok, "
         "riwayat diabetes, riwayat HIV, pemeriksaan kontak, dan komorbiditas lain sebagai "
         "landasan analisis. Jumlah rekam medis pasien MDR-TB yang dihimpun sebanyak 175 data, "
         "terdiri atas 121 pasien dengan status pengobatan berhasil dan 54 pasien dengan status "
         "pengobatan tidak berhasil.",
         "Pengumpulan data dilakukan dengan mengajukan permohonan izin resmi kepada pihak "
         "instansi, dan seluruh prosesnya tetap menjaga kerahasiaan identitas pasien sesuai "
         "dengan ketentuan etika penelitian."),
        ("c. Preprocessing Data",
         "Memastikan kualitas data melalui prosedur pembersihan (cleaning), normalisasi, dan "
         "penanganan nilai yang hilang (missing values) serta encoding data kategorikal agar "
         "siap diproses oleh model.",
         "Preprocessing merupakan tahapan yang sangat menentukan kualitas model akhir. Data "
         "yang bersih dan konsisten akan menghasilkan model yang lebih akurat dan stabil ketika "
         "diterapkan pada data baru."),
        ("d. Pembagian Data",
         "Dataset dibagi menjadi data latih (training) dan data uji (testing). Pembagian "
         "dilakukan sebelum penerapan SMOTE agar tidak terjadi kebocoran data (data leakage).",
         "Pembagian data menggunakan proporsi 80:20, di mana 80% data digunakan untuk melatih "
         "model dan 20% data digunakan untuk menguji performa model pada data yang belum pernah "
         "dilihat."),
        ("e. Penerapan SMOTE",
         "Teknik SMOTE diterapkan pada data latih untuk menyeimbangkan jumlah sampel kelas "
         "minoritas sehingga model tidak mengalami bias saat melakukan klasifikasi.",
         "SMOTE menghasilkan sampel sintetis baru melalui interpolasi k-nearest neighbors. "
         "Dengan data latih yang seimbang, model dapat mempelajari pola kelas minoritas secara "
         "lebih optimal."),
        ("f. Pelatihan Model",
         "Algoritma Random Forest dilatih menggunakan data latih yang telah diseimbangkan untuk "
         "membangun kemampuan prediksi terhadap keberhasilan pengobatan pasien.",
         "Proses pelatihan dilakukan dengan menentukan terlebih dahulu parameter Random Forest, "
         "seperti jumlah pohon (n_estimators) dan kedalaman maksimum (max_depth), melalui "
         "proses tuning hyperparameter."),
        ("g. Validasi Model",
         "Model Random Forest yang telah terbentuk diuji menggunakan data uji guna mengukur "
         "kemampuannya dalam mengklasifikasikan data baru.",
         "Validasi juga dapat dilakukan menggunakan teknik k-fold cross-validation untuk "
         "memperoleh estimasi performa yang lebih stabil dan mengurangi risiko overfitting."),
        ("h. Evaluasi Model",
         "Evaluasi dilakukan melalui confusion matrix untuk mendapatkan nilai akurasi, presisi, "
         "recall, serta F1-score.",
         "Perhatian utama evaluasi diberikan pada recall kelas minoritas, karena kemampuan "
         "mendeteksi pasien berisiko gagal pengobatan merupakan tujuan utama penelitian ini."),
        ("i. Implementasi dan Simulasi",
         "Model yang telah divalidasi digunakan sebagai sistem pendukung keputusan klinis untuk "
         "mengklasifikasikan keberhasilan pengobatan pada data baru.",
         "Simulasi dilakukan dengan memasukkan data pasien baru ke dalam model untuk mengamati "
         "kemampuan model dalam memberikan prediksi yang cepat, akurat, dan konsisten."),
        ("j. Selesai",
         "Seluruh tahapan penelitian telah dilaksanakan dan penelitian dinyatakan selesai.",
         "Hasil akhir seluruh tahapan dirangkum dalam bentuk laporan penelitian yang mencakup "
         "kesimpulan, keterbatasan, serta saran untuk pengembangan penelitian selanjutnya."),
    ]
    for judul, isi, isi2 in tahapan:
        add_heading_sub(doc, judul)
        add_body(doc, isi)
        add_body(doc, isi2)

    add_heading_sub(doc, "2. Metode Algoritma yang Digunakan")
    add_body(doc,
        "Metode yang digunakan dalam penelitian ini adalah kombinasi teknik SMOTE untuk "
        "penanganan ketidakseimbangan data dan algoritma Random Forest sebagai model "
        "klasifikasi. SMOTE diterapkan pada data latih setelah proses preprocessing dan "
        "pembagian data untuk menyeimbangkan jumlah sampel kelas minoritas. Proses ini "
        "menghasilkan dataset yang lebih representatif sehingga model dapat mempelajari pola "
        "kelas minoritas secara lebih baik.")
    add_body(doc,
        "Random Forest membangun sejumlah pohon keputusan menggunakan teknik bagging dan "
        "pemilihan fitur secara acak. Parameter yang digunakan dalam penelitian ini meliputi "
        "n_estimators (jumlah pohon keputusan), max_depth (kedalaman maksimum pohon), dan "
        "random_state untuk memastikan hasil dapat direproduksi. Model dilatih menggunakan data "
        "latih yang telah diseimbangkan SMOTE, kemudian diuji menggunakan data uji untuk menilai "
        "kemampuannya dalam melakukan generalisasi terhadap data baru.")
    add_body(doc,
        "Performa model dievaluasi menggunakan confusion matrix dengan metrik accuracy, "
        "precision, recall, dan F1-score. Perhatian utama diberikan pada recall kelas minoritas "
        "karena kegagalan mendeteksi pasien yang berisiko gagal pengobatan memiliki dampak yang "
        "sangat besar. Dengan performa yang baik, model diharapkan dapat menjadi alat bantu "
        "pendukung keputusan bagi tenaga medis. Alur kerja metode secara rinci ditunjukkan pada "
        "Gambar 5.")
    add_figure(doc, "gambar4_flowchart_metode.png")
    add_caption(doc, "Gambar", 5, "Flowchart Metode SMOTE dan Random Forest")
    flowchart = [
        ("a. Mulai",
         "Tahapan ini merupakan awal dari proses pelatihan model klasifikasi menggunakan teknik "
         "SMOTE dan algoritma Random Forest.",
         "Proses dimulai dengan memahami tujuan penelitian dan mengidentifikasi data yang "
         "dibutuhkan untuk membangun model klasifikasi."),
        ("b. Input Data Rekam Medis",
         "Data rekam medis pasien MDR-TB dimasukkan ke dalam sistem sebagai bahan utama "
         "analisis.",
         "Data yang dimasukkan mencakup variabel demografis, klinis, dan riwayat pengobatan "
         "pasien yang akan digunakan sebagai fitur pada model."),
        ("c. Preprocessing Data",
         "Data dibersihkan, nilai yang hilang ditangani, dan data kategorikal di-encoding agar "
         "siap diproses oleh model.",
         "Langkah ini memastikan data bebas dari nilai yang tidak valid sehingga proses "
         "pelatihan model dapat berjalan dengan baik."),
        ("d. Penerapan SMOTE",
         "Teknik SMOTE diterapkan pada data latih untuk menyeimbangkan jumlah sampel kelas "
         "minoritas.",
         "Penerapan SMOTE menghasilkan data latih yang seimbang sehingga model dapat mengenali "
         "pola kelas minoritas secara lebih baik."),
        ("e. Pembagian Data Latih dan Uji",
         "Dataset dibagi menjadi data latih dan data uji untuk proses pelatihan dan pengujian "
         "model.",
         "Pembagian dilakukan secara proporsional dan acak agar data uji benar-benar "
         "merepresentasikan data yang belum pernah dilihat model."),
        ("f. Inisialisasi Parameter Random Forest",
         "Parameter model ditentukan, meliputi n_estimators, max_depth, dan random_state.",
         "Parameter dioptimalkan melalui proses grid search untuk memperoleh performa model "
         "terbaik pada data latih yang telah diseimbangkan."),
        ("g. Pelatihan Model Random Forest",
         "Model dilatih menggunakan data latih yang telah diseimbangkan untuk mempelajari pola "
         "keberhasilan pengobatan pasien.",
         "Pada tahap ini model membangun sejumlah pohon keputusan dan menggabungkan hasilnya "
         "melalui mekanisme voting mayoritas."),
        ("h. Pengujian Model",
         "Model diuji menggunakan data uji untuk menilai kemampuannya mengklasifikasikan data "
         "baru.",
         "Hasil pengujian dibandingkan dengan label sebenarnya untuk memperoleh nilai metrik "
         "evaluasi."),
        ("i. Evaluasi Model",
         "Performa model diukur menggunakan metrik accuracy, precision, recall, dan F1-score.",
         "Evaluasi dilakukan melalui confusion matrix dengan fokus utama pada recall kelas "
         "minoritas untuk menilai sensitivitas model."),
        ("j. Selesai",
         "Model yang telah terlatih dan tervalidasi siap digunakan sebagai sistem pendukung "
         "keputusan klinis.",
         "Hasil akhir didokumentasikan dalam laporan penelitian beserta kesimpulan dan saran "
         "untuk pengembangan selanjutnya."),
    ]
    for judul, isi, isi2 in flowchart:
        add_heading_sub(doc, judul)
        add_body(doc, isi)
        add_body(doc, isi2)

    add_heading_sub(doc, "3. Teknik Analisis Data")
    add_body(doc,
        "Teknik analisis data pada penelitian ini mengikuti alur yang telah dirancang pada "
        "desain metode penelitian. Tahap pertama adalah pengumpulan data rekam medis pasien "
        "MDR-TB dari RSUD Labuang Baji, Sulawesi Selatan. Data yang terkumpul kemudian menjalani proses "
        "preprocessing untuk memastikan kualitas dan kesiapan data sebelum digunakan dalam "
        "pelatihan model.")
    add_body(doc,
        "Tahap preprocessing dilakukan melalui beberapa langkah. Pertama, data cleaning untuk "
        "menangani nilai yang hilang (missing values) dan data duplikat. Kedua, encoding data "
        "kategorikal menjadi bentuk numerik menggunakan teknik yang sesuai, misalnya label "
        "encoding atau one-hot encoding. Ketiga, normalisasi data numerik untuk menyamakan "
        "skala antar fitur sehingga proses perhitungan jarak pada SMOTE dapat berjalan optimal.")
    add_body(doc,
        "Setelah preprocessing, dataset dibagi menjadi data latih dan data uji dengan proporsi "
        "80:20. Pembagian data dilakukan sebelum penerapan SMOTE untuk mencegah kebocoran data. "
        "Teknik SMOTE kemudian diterapkan hanya pada data latih untuk menyeimbangkan jumlah "
        "sampel kelas minoritas, sehingga model memperoleh representasi data yang lebih merata "
        "selama proses pelatihan.")
    add_body(doc,
        "Model Random Forest dibangun menggunakan data latih yang telah diseimbangkan. "
        "Konfigurasi hyperparameter ditentukan melalui proses grid search untuk memperoleh "
        "kombinasi parameter yang menghasilkan performa terbaik. Parameter yang dioptimalkan "
        "meliputi n_estimators, max_depth, min_samples_split, dan random_state. Model yang "
        "terbentuk kemudian diuji menggunakan data uji yang belum pernah dilihat selama "
        "pelatihan.")
    add_body(doc,
        "Performa model dievaluasi menggunakan confusion matrix dengan metrik accuracy, "
        "precision, recall, dan F1-score. Hasil evaluasi dianalisis untuk menilai sejauh mana "
        "penerapan SMOTE meningkatkan kemampuan model dalam mendeteksi pasien yang berisiko "
        "gagal pengobatan, khususnya pada kelas minoritas. Seluruh proses analisis diharapkan "
        "menghasilkan model yang akurat, stabil, dan dapat diandalkan sebagai sistem pendukung "
        "keputusan klinis.")

    add_heading_sub(doc, "4. Skenario Pengujian")
    add_body(doc,
        "Untuk menilai efektivitas penerapan SMOTE, dilakukan dua skenario pengujian. Skenario "
        "pertama melatih model Random Forest menggunakan dataset asli yang tidak seimbang, "
        "sedangkan skenario kedua melatih model menggunakan dataset yang telah diseimbangkan "
        "dengan SMOTE. Hasil kedua skenario dibandingkan menggunakan metrik accuracy, precision, "
        "recall, dan F1-score.")
    add_body(doc,
        "Perbandingan ini bertujuan untuk menunjukkan sejauh mana SMOTE berkontribusi terhadap "
        "peningkatan kemampuan model dalam mendeteksi kelas minoritas. Dengan melihat selisih "
        "recall antara kedua skenario, efektivitas SMOTE dapat diukur secara kuantitatif dan "
        "disimpulkan secara objektif.")
    add_body(doc,
        "Selain itu, dilakukan eksperimen terhadap variasi proporsi pembagian data (misalnya "
        "80:20 dan 70:30) untuk melihat stabilitas performa model pada berbagai kondisi. Variasi "
        "jumlah pohon (n_estimators) juga diuji untuk menemukan konfigurasi yang menghasilkan "
        "performa terbaik sekaligus menghindari overfitting.")
    add_body(doc,
        "Seluruh hasil pengujian didokumentasikan secara sistematis untuk memudahkan analisis. "
        "Hasil evaluasi disajikan dalam bentuk tabel dan grafik sehingga mudah dipahami oleh "
        "pembaca, termasuk tenaga medis yang tidak memiliki latar belakang teknis yang mendalam. "
        "Dokumentasi ini juga menjadi bahan untuk menarik kesimpulan dan memberikan saran pada "
        "penelitian selanjutnya.")

    # ================================================================ DAFTAR PUSTAKA
    add_body_heading(doc, "DAFTAR PUSTAKA")
    pustaka = [
        "Breiman, L. (2001). Random Forests. Machine Learning, 45(1), 5\u201332.",
        "Chawla, N. V., Bowyer, K. W., Hall, L. O., & Kegelmeyer, W. P. (2002). SMOTE: Synthetic Minority Over-sampling Technique. Journal of Artificial Intelligence Research, 16, 321\u2013357.",
        "Hadjadj, M., Boucherit, I., & Cherif, F. (2024). Penerapan SMOTE untuk meningkatkan performa klasifikasi pada data medis yang tidak seimbang. Jurnal Teknologi Informasi dan Ilmu Komputer, 11(3), 501\u2013510.",
        "Kurniawati, D., & Saputra, R. (2023). Analisis ketidakseimbangan data pada klasifikasi keberhasilan pengobatan tuberkulosis menggunakan machine learning. Jurnal Informatika Kesehatan, 8(2), 112\u2013121.",
        "Lema\u00eetre, G., Nogueira, F., & Aridas, C. K. (2017). Imbalanced-learn: A Python Toolbox to Tackle the Curse of Imbalanced Datasets in Machine Learning. Journal of Machine Learning Research, 18(17), 1\u20135.",
        "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., \u2026 Duchesnay, \u00c9. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, 12, 2825\u20132830.",
        "World Health Organization. (2015). The End TB Strategy. Geneva: World Health Organization.",
        "World Health Organization. (2024). Global Tuberculosis Report 2024. Geneva: World Health Organization.",
        "World Health Organization. (2025). Global Tuberculosis Report 2025. Geneva: World Health Organization.",
        "Zhang, Y., Liu, X., & Wang, J. (2025). Integrasi SMOTE dan Random Forest untuk prediksi hasil pengobatan pasien MDR-TB. Jurnal Sistem Cerdas dan Rekayasa Data, 9(1), 55\u201366.",
    ]
    for entri in pustaka:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing = 1.5
        pf.space_after = Pt(10)
        pf.left_indent = Cm(1.0)
        pf.first_line_indent = Cm(-1.0)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(entri)
        set_run(r)

    # ================================================================ LAMPIRAN: PROMPT GAMBAR
    add_front_heading(doc, "LAMPIRAN")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run("PROMPT GENERASI GAMBAR")
    set_run(r, size=12, bold=True)
    add_body(doc,
        "Lampiran ini berisi prompt (instruksi) untuk membuat gambar-gambar diagram pada "
        "proposal ini. Prompt ditulis agar konsisten satu sama lain: setiap prompt mengacu pada "
        "blok gaya visual yang sama pada bagian A, sehingga kelima gambar memiliki palet warna, "
        "bentuk kotak, gaya panah, dan tipe huruf yang seragam. Disarankan menggunakan Ideogram "
        "karena akurat dalam merender teks Bahasa Indonesia; alternatifnya dapat menggunakan "
        "Midjourney atau DALL\u00b7E kemudian menulis ulang teksnya secara manual. Setelah gambar "
        "dihasilkan, simpan sebagai figures/gambar1.png hingga figures/gambar5.png lalu jalankan "
        "ulang build_proposal.py.")

    add_heading_letter(doc, "A. Gaya Visual Konsisten (Digunakan untuk Semua Gambar)")
    add_body(doc,
        "Create a clean flat vector diagram in a consistent minimalist academic style. Use a "
        "pure white background. All outlines and arrows use navy blue (#2B4D7F). Boxes are "
        "white-filled with a rounded-rectangle shape, a thin navy outline, and slightly rounded "
        "corners; start and end boxes are capsule-shaped (pill). Arrows are thin navy lines with "
        "small solid arrowheads. All labels are written in Indonesian using a clean modern "
        "sans-serif font (such as Inter, Segoe UI, or Poppins), centered horizontally inside "
        "each box, in dark gray (#222222). Use a vertical top-to-bottom layout with boxes evenly "
        "spaced. Do not add shadows, gradients, 3D effects, watermarks, signatures, or extra "
        "decorations. Keep the diagram clean and suitable for an academic thesis. Use a portrait "
        "aspect ratio (approximately 4:5).")

    add_heading_letter(doc, "B. Prompt Gambar 1 \u2014 Alur Kerja SMOTE")
    add_body(doc,
        "Using the exact style described above, create a flowchart titled \"Alur Kerja SMOTE\". "
        "Arrange six boxes vertically from top to bottom, connected by downward arrows. Each box "
        "must contain exactly this Indonesian text, in order: (1) \"Data Tidak Seimbang\"; "
        "(2) \"Seleksi Sampel Kelas Minoritas\"; (3) \"Perhitungan k Tetangga Terdekat (k-NN)\"; "
        "(4) \"Interpolasi Sampel Sintetis\" with the formula X_new = X_i + \u03bb (X_zi \u2212 X_i) "
        "displayed below it; (5) \"Dataset Seimbang\"; (6) \"Pelatihan Model Klasifikasi\". "
        "Draw a downward arrow from each box to the box below it. Use the exact colors, box "
        "shapes, and fonts from the style block.")

    add_heading_letter(doc, "C. Prompt Gambar 2 \u2014 Arsitektur Random Forest")
    add_body(doc,
        "Using the exact style described above, create an architecture diagram titled "
        "\"Arsitektur Random Forest\". Lay out the components top to bottom as follows: top "
        "center, a box labeled \"Dataset Latih\"; below it, a box labeled \"Bootstrap Sampling "
        "(Bagging)\"; below that, a horizontal row of three tree boxes labeled \"Pohon Keputusan "
        "1\", \"Pohon Keputusan 2\", and \"Pohon Keputusan n\"; below the trees, a box labeled "
        "\"Voting (Mayoritas Suara)\"; and at the bottom, a box labeled \"Hasil Klasifikasi\". "
        "Draw one arrow from \"Bootstrap Sampling\" to each tree box, one arrow from each tree "
        "box to \"Voting (Mayoritas Suara)\", and one arrow from \"Voting\" to \"Hasil "
        "Klasifikasi\". All labels in Indonesian, using the exact colors, shapes, and fonts from "
        "the style block.")

    add_heading_letter(doc, "D. Prompt Gambar 3 \u2014 Kerangka Berpikir")
    add_body(doc,
        "Using the exact style described above, create a wide horizontal (landscape) diagram "
        "titled \"Kerangka Berpikir\" for a study conducted at RSUD Labuang Baji, Sulawesi "
        "Selatan. Arrange six wide rounded-rectangle boxes stacked vertically from top to "
        "bottom, each spanning nearly the full image width, connected by short downward "
        "arrows. Each box has a bold Indonesian label with smaller supporting Indonesian "
        "text that must fit exactly two lines, as follows in order: (1) \"Masalah\" \u2014 "
        "\"Tingginya kegagalan pengobatan MDR-TB dan sulitnya deteksi dini karena data rekam "
        "medis RSUD Labuang Baji yang tidak seimbang.\"; (2) \"Tujuan\" \u2014 \"Mengembangkan "
        "model klasifikasi yang memprediksi keberhasilan pengobatan pasien MDR-TB secara "
        "akurat, khususnya kelas minoritas yang berisiko gagal.\"; (3) \"Data\" \u2014 \"Data "
        "rekam medis pasien MDR-TB RSUD Labuang Baji Sulawesi Selatan mencakup demografis, "
        "klinis, dan riwayat pengobatan.\"; (4) \"Solusi\" \u2014 \"Penerapan SMOTE untuk "
        "menyeimbangkan dataset dan Random Forest sebagai model klasifikasi.\"; (5) \"Metode\" "
        "\u2014 \"Data dipreprocessing, diterapkan SMOTE, dibagi data latih dan uji, lalu "
        "dilatih Random Forest dan dievaluasi.\"; (6) \"Hasil\" \u2014 \"Model akurat dan "
        "sensitif terhadap pasien berisiko gagal, berpotensi menjadi sistem pendukung "
        "keputusan tenaga medis.\" Each supporting text must wrap to no more than two lines "
        "inside its wide box; keep the boxes compact in height. Use the exact colors, shapes, "
        "and fonts from the style block.")

    add_heading_letter(doc, "E. Prompt Gambar 4 \u2014 Desain Metode Penelitian")
    add_body(doc,
        "Using the exact style described above, create a flowchart titled \"Desain Metode "
        "Penelitian\". Arrange ten boxes vertically from top to bottom, connected by downward "
        "arrows. \"Mulai\" is the capsule-shaped start box and \"Selesai\" is the capsule-shaped "
        "end box; the others are rounded rectangles. Each box must contain exactly this "
        "Indonesian text, in order: (1) \"Mulai\"; (2) \"Pengumpulan Data Rekam Medis MDR-TB\"; "
        "(3) \"Preprocessing Data\"; (4) \"Pembagian Data Latih dan Data Uji\"; (5) \"Penerapan "
        "SMOTE\"; (6) \"Pelatihan Model Random Forest\"; (7) \"Validasi Model\"; (8) \"Evaluasi "
        "Model\"; (9) \"Implementasi dan Simulasi\"; (10) \"Selesai\". Draw a downward arrow "
        "from each box to the box below it. Use the exact colors, box shapes, and fonts from "
        "the style block.")

    add_heading_letter(doc, "F. Prompt Gambar 5 \u2014 Flowchart Metode SMOTE dan Random Forest")
    add_body(doc,
        "Using the exact style described above, create a flowchart titled \"Flowchart Metode "
        "SMOTE dan Random Forest\". Arrange ten boxes vertically from top to bottom, connected "
        "by downward arrows. \"Mulai\" is the capsule-shaped start box and \"Selesai\" is the "
        "capsule-shaped end box; the others are rounded rectangles. Each box must contain "
        "exactly this Indonesian text, in order: (1) \"Mulai\"; (2) \"Input Data Rekam Medis\"; "
        "(3) \"Preprocessing Data\"; (4) \"Pembagian Data Latih dan Data Uji\"; (5) \"Penerapan "
        "SMOTE pada Data Latih\"; (6) \"Inisialisasi Parameter Random Forest\"; (7) \"Pelatihan "
        "Model Random Forest\"; (8) \"Pengujian Model\"; (9) \"Evaluasi Model\"; (10) "
        "\"Selesai\". Draw a downward arrow from each box to the box below it. Use the exact "
        "colors, box shapes, and fonts from the style block.")

    doc.save(OUTDOCX)
    print("Saved:", OUTDOCX)


if __name__ == "__main__":
    build()
