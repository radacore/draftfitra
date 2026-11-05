#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ekstrak DAFTAR PUSTAKA (PDF hal. 73-75) menjadi entri dengan run italic."""
import fitz, json, re

PDF = "Proposal_Skripsi_Zulfitra_MDR-TB.pdf"
OUT = "revisi_work/refs.json"
doc = fitz.open(PDF)

entries = []
cur = []
cur_x0 = None


def norm(runs):
    # 1. rapikan spasi di dalam tiap run
    seq = []
    for r in runs:
        t = re.sub(r"[ \t]+", " ", r["t"])
        if not t:
            continue
        seq.append({"t": t, "i": r["i"]})
    # 2. sambung antar-run: buang spasi ganda + gabung kata terpotong
    for k in range(1, len(seq)):
        prev, cur = seq[k - 1], seq[k]
        if prev["t"].endswith((" ", "-", "\u2013")) and cur["t"].startswith(" "):
            cur["t"] = cur["t"].lstrip()
        if prev["t"].endswith(("-", "\u2013")) and cur["t"]:
            cur["t"] = cur["t"].lstrip()
    # 3. gabung run dengan gaya sama
    merged = []
    for r in seq:
        if not r["t"]:
            continue
        if merged and merged[-1]["i"] == r["i"]:
            merged[-1]["t"] += r["t"]
        else:
            merged.append(dict(r))
    # 4. spasi ganda sisa (mis. "Sistem " + " Deteksi" yang sudah digabung)
    for r in merged:
        r["t"] = re.sub(r" {2,}", " ", r["t"])
    if merged:
        merged[0]["t"] = merged[0]["t"].lstrip()
        merged[-1]["t"] = merged[-1]["t"].rstrip()
    return merged


def flush():
    global cur
    if cur:
        runs = norm(cur)
        if runs:
            entries.append(runs)
    cur = []


for pno in (72, 73, 74):
    page = doc[pno]
    lines = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        for l in b["lines"]:
            if l["bbox"][1] > 780:
                continue
            txt = "".join(s["text"] for s in l["spans"]).strip()
            if not txt:
                continue
            lines.append(l)
    lines.sort(key=lambda l: (round(l["bbox"][1], 1), l["bbox"][0]))
    for l in lines:
        txt = "".join(s["text"] for s in l["spans"]).strip()
        if pno == 72 and l["bbox"][1] < 100 and "DAFTAR PUSTAKA" in txt:
            continue
        runs = [{"t": s["text"], "i": "Italic" in s["font"]}
                for s in l["spans"] if s["text"]]
        if l["bbox"][0] < 130:
            flush()
            cur = runs
        else:
            if cur:
                cur.append({"t": " ", "i": cur[-1]["i"]})
            cur.extend(runs)
flush()

json.dump(entries, open(OUT, "w"), indent=1, ensure_ascii=False)
print("entri:", len(entries))
for i, e in enumerate(entries):
    t = "".join(r["t"] for r in e)
    print(f"{i:2} [{len(t):3}] {t[:95]}")
