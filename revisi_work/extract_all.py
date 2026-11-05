#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ekstrak DAFTAR SINGKATAN + BAB I-III dari PDF menjadi blok berurutan.

Output: revisi_work/full123.json
"""
import fitz, json, os, re

PDF = "Proposal_Skripsi_Zulfitra_MDR-TB.pdf"
OUT = "revisi_work/full123.json"
IMGDIR = "revisi_work/pdf_img"

doc = fitz.open(PDF)
os.makedirs(IMGDIR, exist_ok=True)

CAPTION = re.compile(r"^(Gambar|Tabel)\s+\d+\.")
H_A = re.compile(r"^[A-Z]\.\s")
H_SUB = re.compile(r"^(\d+|[a-z])\.\s")
SIST = re.compile(r"^BAB (I|II|III|IV|V)\s+(PENDAHULUAN|TINJAUAN PUSTAKA|METODE PENELITIAN|HASIL DAN PEMBAHASAN|PENUTUP)$")
EQ_FONTS = ("LiberationSerif", "OpenSymbol")

# batas kolom tiap tabel (poin, dari geometri PDF)
TABLE_COLS = {
    1: [108.0, 201.0, 319.0, 505.0],
    2: [108.0, 183.0, 313.0, 381.0, 505.0],
    3: [108.0, 135.0, 402.0, 468.0, 516.7],
}


def norm_runs(runs):
    out = []
    for r in runs:
        t = re.sub(r"[ \t]+", " ", r["t"])
        if not t:
            continue
        if out and out[-1].get("i") == r["i"] and out[-1].get("b") == r["b"]:
            out[-1]["t"] += t
        else:
            out.append({"t": t, "i": r["i"], "b": r["b"]})
    if not out:
        return out
    # rapikan spasi di batas run
    for k in range(1, len(out)):
        if out[k - 1]["t"].endswith(" ") and out[k]["t"].startswith(" "):
            out[k]["t"] = out[k]["t"].lstrip()
        out[k]["t"] = re.sub(r" {2,}", " ", out[k]["t"])
    out[0]["t"] = out[0]["t"].lstrip()
    # hyphenation lintas run
    for k in range(1, len(out)):
        if out[k - 1]["t"].endswith("-") and out[k]["t"].startswith(" "):
            out[k]["t"] = out[k]["t"].lstrip()
    merged = []
    for r in out:
        if not r["t"]:
            continue
        if merged and merged[-1]["i"] == r["i"] and merged[-1]["b"] == r["b"]:
            merged[-1]["t"] += r["t"]
        else:
            merged.append(dict(r))
    for k in range(1, len(merged)):
        if merged[k - 1]["t"].endswith(" ") and merged[k]["t"].startswith(" "):
            merged[k]["t"] = merged[k]["t"].lstrip()
        merged[k]["t"] = re.sub(r" {2,}", " ", merged[k]["t"])
    # squeeze spasi global sambil menjaga run
    squeezed = []
    prev_space = True
    for r in merged:
        t = re.sub(r" {2,}", " ", r["t"])
        if prev_space:
            t = t.lstrip()
        if not t:
            continue
        squeezed.append({"t": t, "i": r["i"], "b": r["b"]})
        prev_space = t.endswith(" ")
    if squeezed:
        squeezed[-1]["t"] = squeezed[-1]["t"].rstrip()
    # sambung kata yang terpotong tanda hubung di akhir baris: "Over- sampling"
    for r in squeezed:
        r["t"] = re.sub(r"(?<=[A-Za-z])-[ \t]+(?=[A-Za-z0-9])", "-", r["t"])
    for k in range(1, len(squeezed)):
        if squeezed[k - 1]["t"].endswith("-"):
            squeezed[k]["t"] = squeezed[k]["t"].lstrip()
    return squeezed


def runs_text(runs):
    return "".join(r["t"] for r in runs).strip()


def add_runs(dst, src, sep=" "):
    if not src:
        return
    if dst and sep:
        if dst[-1]["t"].endswith(" ") or src[0]["t"].startswith(" "):
            sep = ""
        if sep:
            dst.append({"t": sep, "i": dst[-1].get("i", False), "b": dst[-1].get("b", False)})
    for r in src:
        dst.append(dict(r))


def page_items(page):
    lines, images = [], []
    for info in page.get_image_info(xrefs=True):
        if not info.get("xref"):
            continue
        images.append({"kind": "img", "y": info["bbox"][1], "x": info["bbox"][0],
                       "bbox": info["bbox"], "xref": info["xref"],
                       "px": (info["width"], info["height"])})
    for b in page.get_text("dict")["blocks"]:
        if b["type"] == 1:
            continue
        for l in b["lines"]:
            if l["bbox"][1] > 780:
                continue
            txt = "".join(s["text"] for s in l["spans"]).strip()
            if not txt:
                continue
            spans = l["spans"]
            # nomor persamaan "(1)" di margin kanan bukan bagian teks paragraf
            if re.fullmatch(r"\(\d+\)", txt) and l["bbox"][0] > 450:
                continue
            # glif subscript yatim dari persamaan (DejaVuSans) bukan teks paragraf
            if all("DejaVuSans" in s["font"] for s in spans if s["text"].strip()):
                continue
            eq = any(any(f in s["font"] for f in EQ_FONTS) for s in spans)
            runs = [{"t": s["text"], "i": "Italic" in s["font"], "b": "Bold" in s["font"]}
                    for s in spans if s["text"]]
            lines.append({"kind": "line", "y": l["bbox"][1], "x0": l["bbox"][0],
                          "x1": l["bbox"][2], "text": txt, "runs": runs,
                          "size": round(spans[0]["size"], 1),
                          "bold": any("Bold" in s["font"] for s in spans),
                          "eq": eq})
    lines.sort(key=lambda z: (round(z["y"], 1), z["x0"]))
    return lines, images


def hline_zones(page, cap_ys=()):
    hs = []
    for d in page.get_drawings():
        r = d["rect"]
        if abs(r.y0 - r.y1) < 0.6 and (r.x1 - r.x0) > 100:
            hs.append((round(r.y0, 1), round(r.x0, 1), round(r.x1, 1)))
    groups = []
    for y, x0, x1 in sorted(hs):
        if groups and abs(groups[-1][-1][1] - x0) < 4 and abs(groups[-1][-1][2] - x1) < 4 \
                and y - groups[-1][-1][0] <= 200:
            groups[-1].append((y, x0, x1))
        else:
            groups.append([(y, x0, x1)])
    # pisahkan gabungan beberapa tabel pada satu halaman pada baris caption
    out = []
    for g in groups:
        seg = g
        for c in sorted(y for y in cap_ys if g[0][0] < y < g[-1][0]):
            left = [x for x in seg if x[0] < c]
            right = [x for x in seg if x[0] > c]
            if len(left) >= 2:
                out.append(left)
            seg = right
        if len(seg) >= 2:
            out.append(seg)
    return out


def merge_same_line(lines):
    merged = []
    for it in lines:
        if merged and abs(it["y"] - merged[-1]["y"]) < 2.2:
            m = merged[-1]
            add_runs(m["runs"], it["runs"], sep=" ")
            m["text"] = (m["text"] + " " + it["text"]).strip()
            m["x1"] = max(m["x1"], it["x1"])
            m["bold"] = m["bold"] or it["bold"]
        else:
            merged.append(dict(it, runs=[dict(r) for r in it["runs"]]))
    return merged


def build_table(zone, lines, images, bounds):
    ys = [z[0] for z in zone]

    def col_of(x):
        c = 0
        for b in bounds[1:-1]:
            if x >= b - 0.6:
                c += 1
        return c

    rows = []
    for a, b in zip(ys, ys[1:]):
        cells = [[] for _ in range(len(bounds) - 1)]
        content = [l for l in lines if a - 0.8 <= l["y"] < b - 0.8]
        content.sort(key=lambda z: (z["y"], z["x0"]))
        for l in content:
            c = col_of(l["x0"])
            if cells[c]:
                add_runs(cells[c], l["runs"], sep=" ")
            else:
                cells[c] = [dict(r) for r in l["runs"]]
        for im in images:
            cy = (im["bbox"][1] + im["bbox"][3]) / 2
            if a - 0.8 <= cy < b - 0.8:
                fn = save_img(im["xref"], im["bbox"], im.get("px", (9999, 9999))[0] < 320, tbl_page[0])
                cells[col_of(im["bbox"][0])] = [{"img": fn}]
        cleaned = []
        for c in cells:
            if c and "img" in c[0]:
                cleaned.append(c)
            else:
                cleaned.append(norm_runs(c))
        rows.append(cleaned)
    return [r for r in rows if any(c for c in r)]


def save_img(xref, bbox=None, small=False, page=None):
    fname = os.path.join(IMGDIR, f"img_{xref}.png")
    if not os.path.exists(fname):
        if small and page is not None and bbox is not None:
            zoom = 12
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(clip=fitz.Rect(bbox), matrix=mat, alpha=False)
        else:
            pix = fitz.Pixmap(doc, xref)
            if pix.n - pix.alpha >= 4:
                pix = fitz.Pixmap(fitz.csRGB, pix)
        pix.save(fname)
    return fname


def cell_text(cell):
    return "".join(r.get("t", "") for r in cell).strip()


def merge_table_fragments(blocks):
    """Gabungkan potongan tabel yang terpisah halaman (header berulang)."""
    out = []
    for b in blocks:
        if (b["type"] == "table" and out and out[-1]["type"] == "table"
                and out[-1]["table_no"] == b["table_no"]):
            tgt = out[-1]
            rows = b["rows"]
            if rows and cell_text(rows[0][0]).lower() == cell_text(tgt["rows"][0][0]).lower():
                rows = rows[1:]
            for r in rows:
                filled = [i for i, c in enumerate(r) if c]
                if len(filled) == 1 and filled[0] == len(r) - 1 and tgt["rows"]:
                    add_runs(tgt["rows"][-1][-1], r[-1])
                else:
                    tgt["rows"].append(r)
            continue
        out.append(b)
    return out


# ---------------------------------------------------------------- ekstraksi
sings_rows = None
blocks = []
buf = []
buf_x0 = [0.0]
table_counter = 0
last_caption_no = None
tbl_page = [None]
prev_page = [None]

for pno in range(6, 53):
    page = doc[pno]
    tbl_page[0] = page
    lines, images = page_items(page)
    cap_ys = [l["y"] for l in lines if l["bold"] and CAPTION.match(l["text"])]
    zones = hline_zones(page, cap_ys)

    def in_zone(y):
        return any(z[0][0] - 0.8 <= y < z[-1][0] - 0.8 for z in zones)

    body_lines = [l for l in lines if not in_zone(l["y"]) and not l["eq"]]
    table_imgs = [im for im in images if any(z[0][0] - 0.8 <= im["bbox"][1] < z[-1][0] for z in zones)]
    body_imgs = [im for im in images if im not in table_imgs]

    items = []
    for l in body_lines:
        items.append({"kind": "line", "y": l["y"], "x0": l["x0"], "x1": l["x1"],
                      "runs": l["runs"], "text": l["text"], "bold": l["bold"],
                      "size": l["size"]})
    for im in body_imgs:
        items.append({"kind": "img", "y": im["bbox"][1], "x0": im["bbox"][0],
                      "bbox": im["bbox"], "xref": im["xref"]})
    for z in zones:
        y0, y1 = z[0][0], z[-1][0]
        tl = [l for l in lines if y0 - 0.8 <= l["y"] < y1 - 0.8]
        ti = [im for im in table_imgs if y0 - 0.8 <= im["bbox"][1] < y1 - 0.8]
        items.append({"kind": "tbl", "y": y0, "zone": z, "lines": tl, "images": ti})

    items.sort(key=lambda z: (round(z["y"], 1), 0 if z["kind"] != "line" else z["x0"]))

    merged = merge_same_line([i for i in items if i["kind"] == "line"])
    seq = merged + [i for i in items if i["kind"] != "line"]
    seq.sort(key=lambda z: (round(z["y"], 1), z.get("x0", 0)))


    def flush():
        global buf
        if not buf:
            return
        runs = norm_runs(buf)
        if not runs:
            buf = []
            return
        txt = runs_text(runs)
        b = {"type": "body", "runs": runs, "x0": round(buf_x0[0], 1)}
        if SIST.match(txt):
            b = {"type": "sistematika", "bab": txt.split()[0] + " " + txt.split()[1],
                 "title": txt.split(" ", 2)[2]}
        elif re.match(r"^(\d+|[a-z])\.\s*\S", txt) and buf_x0[0] < 130:
            b["list"] = True
        blocks.append(b)
        buf = []

    i = 0
    while i < len(seq):
        it = seq[i]
        if it["kind"] == "img":
            flush()
            small = it.get("px", (9999, 9999))[0] < 320
            blocks.append({"type": "image",
                           "file": save_img(it["xref"], it["bbox"], small, page),
                           "width_cm": round((it["bbox"][2] - it["bbox"][0]) / 72 * 2.54, 2)})
            i += 1
            continue

        if it["kind"] == "tbl":
            flush()
            if pno == 6:
                rows = build_table(it["zone"], it["lines"], it["images"], [108.0, 233.5, 488.2])
                if sings_rows is None:
                    sings_rows = rows
                else:
                    sings_rows.extend(rows[1:] if rows and runs_text(rows[0][0]) == "Singkatan" else rows)
                i += 1
                continue
            bounds = TABLE_COLS.get(last_caption_no)
            if bounds is None:
                i += 1
                continue
            rows = build_table(it["zone"], it["lines"], it["images"], bounds)
            if os.environ.get("DBG"):
                print("DBG tbl", last_caption_no, "imgs", len(it["images"]), "rows", [[list(c) for c in r] for r in rows[:2]])
            blocks.append({"type": "table", "table_no": last_caption_no, "rows": rows})
            i += 1
            continue

        txt = it["text"].strip()
        runs = it["runs"]

        if it["bold"] and it["size"] >= 13:
            flush()
            blocks.append({"type": "chapter_line", "runs": runs})
            i += 1
            continue

        if it["bold"] and CAPTION.match(txt):
            flush()
            blocks.append({"type": "caption", "runs": runs})
            m = re.match(r"^Tabel\s+(\d+)\.", txt)
            if m:
                last_caption_no = int(m.group(1))
            i += 1
            continue

        if it["bold"] and H_A.match(txt) and len(txt) < 120:
            flush()
            blocks.append({"type": "h_a", "runs": runs})
            i += 1
            continue

        if it["bold"] and H_SUB.match(txt) and len(txt) < 120:
            flush()
            blocks.append({"type": "h_sub", "runs": runs})
            i += 1
            continue

        prev = seq[i - 1] if i > 0 else prev_page[0]
        newp = False
        if not buf:
            newp = True
        elif prev and prev["kind"] in ("img", "tbl"):
            newp = True
        elif prev and prev["kind"] == "line":
            # baris pertama paragraf menjorok (148.9 / 204.2 pt); lanjutan di 113.5 pt
            if it["x0"] >= 140 and prev["x0"] < 140:
                newp = True
            elif SIST.match(txt):
                newp = True
            elif it["x0"] < 140 and prev["x0"] > 180:
                # kembali ke margin setelah paragraf sistematika (indent 204.2)
                newp = True
            elif re.match(r"^(\d+|[a-z])\.(?=\S|\s)", txt):
                newp = True
            elif it["y"] - prev["y"] > 27:
                newp = True
        if newp:
            flush()
            buf_x0[0] = it["x0"]
        add_runs(buf, runs, sep=" " if buf else "")
        i += 1

    prev_page[0] = seq[-1] if seq else prev_page[0]

flush()
blocks = merge_table_fragments(blocks)

json.dump({"sings": sings_rows, "blocks": blocks}, open(OUT, "w"), indent=1, ensure_ascii=False)

from collections import Counter
print(Counter(b["type"] for b in blocks))
print("sings rows:", len(sings_rows) if sings_rows else None)
print("tabel:", [(b["table_no"], len(b["rows"]), max(len(r) for r in b["rows"])) for b in blocks if b["type"] == "table"])
for b in blocks[:14]:
    t = runs_text(b["runs"]) if "runs" in b else b.get("file", "")
    print("  ", b["type"], "|", t[:90])
