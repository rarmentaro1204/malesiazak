#!/usr/bin/env python3
"""Unisce tutti i file output/*.xlsx in un unico Excel (output/UNIONE_Malesia.xlsx).

Aggiunge le colonne stato_malesia e file_origine, deduplica per sito_web
(o ragione_sociale se il sito manca). Uso: python3 merge_output.py
"""
import glob, os, re
from openpyxl import Workbook, load_workbook

OUT = "output/UNIONE_Malesia.xlsx"
COLS = ["ragione_sociale", "sito_web", "provincia_area", "settore",
        "telefono", "email", "note", "stato_verifica"]
HEAD = ["stato_malesia"] + COLS + ["file_origine"]

rows, seen = [], set()
for path in sorted(glob.glob("output/*_Malesia_*.xlsx")):
    m = re.search(r"_Malesia_([^_]+)_", os.path.basename(path))
    stato = m.group(1) if m else ""
    ws = load_workbook(path, read_only=True).active
    it = ws.iter_rows(values_only=True)
    hdr = [str(h).strip() if h else "" for h in next(it, [])]
    for r in it:
        d = dict(zip(hdr, r))
        if not any(d.values()):
            continue
        key = str(d.get("sito_web") or d.get("ragione_sociale") or "").lower().strip().rstrip("/")
        if key and key in seen:
            continue
        seen.add(key)
        rows.append([stato] + [d.get(c) for c in COLS] + [os.path.basename(path)])

wb = Workbook()
ws = wb.active
ws.title = "Unione"
ws.append(HEAD)
for r in rows:
    ws.append(r)
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions
for i, h in enumerate(HEAD, 1):
    ws.column_dimensions[ws.cell(1, i).column_letter].width = max(16, len(h) + 4)
wb.save(OUT)
print(f"{len(rows)} righe unite -> {OUT}")
