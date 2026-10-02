"""Erzeugt aus Bedarf_Bestand_LST und LST_Zielzustand_gesamt eine Kopie mit Excel-Dashboard.

    python build_dashboard.py <Bedarf_Bestand_LST.xlsx> <LST_Zielzustand_gesamt.xlsx> <ausgabe.xlsx>

Ablauf:
  1. Bezirks-Blaetter pruefen (Kopfzeile 1 oder 2, Ueberschriften der verwendeten Spalten) und dem Zielzustand zuordnen
  2. build_sheets.py      - neue Blaetter (Uebersicht, Soll-Ist, Dashboard je Bezirk, Pruefliste, Zielzustand,
                            Einstellungen, Daten)
  3. merge_into_original  - Blaetter direkt in das .xlsx-Paket des Originals einsetzen (Original-Blaetter bleiben unveraendert)
  4. LibreOffice          - Kopie durchrechnen; keine Fehlerwerte in den neuen Blaettern
  5. inject_cache         - berechnete Werte in Zellen und Diagrammen speichern
Die Ergebnisse prueft verify.py gegen die unabhaengige Nachrechnung in reference.py.
Benoetigt: openpyxl, lxml, LibreOffice (soffice) mit Calc.
"""
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from build_sheets import HEADER_PREFIX, HEADERS, District, build
from inject_cache import ERRORS, inject
from merge_into_original import merge
from reference import compute

# Blatt in Bedarf_Bestand_LST -> Bezirk im Zielzustand (Zeile 1) -> erwartete Spalte "Zielzustand"
DISTRICTS = [
    ("Limburg (KRM)", "Limburg 1 (SFS KC)", "K"),
    ("Limburg 2", "Limburg 2", "O"),
    ("FFM West", "FFM West/Höchst", "S"),
    ("FFM Süd", "FFM Süd", "W"),
    ("FFM FA", "FFM Abstellbahnhof", "AA"),
    ("FFM Hbf", "FFM Hbf", "AE"),
    ("Aschaffenburg", "Aschaffenburg", "AI"),
    ("Hanau", "Hanau", "AM"),
    ("Wetzlar", "Wetzlar", "AQ"),
    ("Gießen", "Gießen", "AU"),
    ("Wetterau", "Wetterau", "AY"),
    ("Gelnhausen", "Gelnhausen", "BC"),
]

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def libreoffice_recalc(path, timeout=900):
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as prof:
        url = Path(prof).as_uri()
        subprocess.run(["soffice", "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                       env=env, capture_output=True, timeout=120)
        macro_dir = Path(prof) / "user" / "basic" / "Standard"
        (macro_dir / "Module1.xba").write_text(MACRO)
        before = os.stat(path).st_mtime_ns
        subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation={url}",
                        "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
                        str(Path(path).absolute())], env=env, capture_output=True, timeout=timeout)
        if os.stat(path).st_mtime_ns == before:
            raise RuntimeError("LibreOffice hat die Datei nicht neu berechnet.")


def check_districts(src_wb, zz_ws):
    """Blaetter und Ueberschriften pruefen, Zuordnung zum Zielzustand kontrollieren."""
    zz_col = {str(c.value).strip(): c.column_letter for c in zz_ws[1] if c.value}
    out = []
    for sheet, zz_name, zz_letter in DISTRICTS:
        if sheet not in src_wb.sheetnames:
            raise RuntimeError(f"Blatt „{sheet}“ fehlt in Bedarf_Bestand_LST.")
        ws = src_wb[sheet]
        hdr = next((r for r in (1, 2, 3) if str(ws[f"B{r}"].value or "").strip() == "Name"), None)
        if hdr is None:
            raise RuntimeError(f"„{sheet}“: keine Kopfzeile mit „Name“ in Spalte B gefunden.")
        wrong = [f"{c}{hdr}={ws[f'{c}{hdr}'].value!r} (erwartet {t!r})" for c, t in HEADERS.items()
                 if str(ws[f"{c}{hdr}"].value or "").strip() != t]
        wrong += [f"{c}{hdr}={ws[f'{c}{hdr}'].value!r} (erwartet „{t}…“)" for c, t in HEADER_PREFIX.items()
                  if not str(ws[f"{c}{hdr}"].value or "").startswith(t)]
        if wrong:
            raise RuntimeError(f"„{sheet}“: Spalten passen nicht: {'; '.join(wrong)}")
        if zz_name not in zz_col:
            raise RuntimeError(f"Bezirk „{zz_name}“ fehlt in Zeile 1 des Zielzustands.")
        idx = zz_ws[zz_col[zz_name] + "1"].column + 3
        letter = zz_ws.cell(2, idx).column_letter
        if letter != zz_letter or str(zz_ws.cell(2, idx).value).strip() != "Zielzustand":
            raise RuntimeError(f"Zielzustand „{zz_name}“: Spalte {letter} statt {zz_letter} oder keine Überschrift „Zielzustand“.")
        out.append(District(sheet, zz_name))
    return out


def chart_ymax(ref):
    """Gemeinsame Skala der kleinen Verlaufsdiagramme (Soll-Ist): hoechste Saeule (inkl. Teamleiter) bzw. Soll + Luft."""
    top = 0
    for d in ref["districts"].values():
        for row in d["counts"].values():
            top = max(top, sum(v for k, v in row.items()))
        top = max(top, sum(d["soll"].values()))
    return int(math.ceil(top * 1.15 / 5.0) * 5)


def main(src, zz_path, out, keep_calc=None):
    src_wb = load_workbook(src, data_only=True)
    zz_ws = load_workbook(zz_path).worksheets[0]
    zz_vals = load_workbook(zz_path, data_only=True).worksheets[0]
    districts = check_districts(src_wb, zz_vals)
    ref = compute(src, zz_path, [(d.sheet, d.zz_name) for d in districts], date.today().year)
    ymax = chart_ymax(ref)
    with tempfile.TemporaryDirectory(prefix="dashboard-") as tmp:
        new = os.path.join(tmp, "neue_blaetter.xlsx")
        merged = os.path.join(tmp, "merged.xlsx")
        calc = os.path.join(tmp, "calc.xlsx")
        zz_name = re.sub(r"^[0-9a-f]{8}-", "", Path(zz_path).name)    # Praefix hochgeladener Dateien entfernen
        sheet_names = build(districts, zz_ws, zz_name, new, ymax=ymax)
        merge(src, new, merged)
        shutil.copy(merged, calc)
        libreoffice_recalc(calc)
        wb = load_workbook(calc, data_only=True)
        errors = [f"{n}!{c.coordinate}" for n in sheet_names for row in wb[n].iter_rows() for c in row
                  if isinstance(c.value, str) and c.value in ERRORS]
        if errors:
            raise RuntimeError(f"Formelfehler in den neuen Blättern ({len(errors)}): {errors[:20]}")
        info = inject(merged, calc, out, sheet_names)
        if keep_calc:
            shutil.copy(calc, keep_calc)
    print(f"{len(districts)} Bezirke, {len(sheet_names)} neue Blätter, {info['cells']} Formelzellen, "
          f"{info['charts']} Diagramme -> {out}")
    return districts


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3], os.environ.get("KEEP_CALC"))
