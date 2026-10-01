"""Erzeugt aus Bedarf_Bestand_LST und LST_Zielzustand_gesamt eine Datei mit Excel-Dashboard.

    python build_dashboard.py <Bedarf_Bestand_LST.xlsx> <LST_Zielzustand_gesamt.xlsx> <ausgabe.xlsx>

Ablauf:
  1. Bezirks-Blaetter erkennen (Kopfzeile mit "Ziel-Qualifikation") und dem Zielzustand zuordnen
  2. build_sheets.py      - neue Blaetter (Uebersicht, Soll-Ist, Grafik je Bezirk, Daten, Zielzustand, Einstellungen)
  3. merge_into_original  - Blaetter in die Original-Datei einsetzen (Original-Blaetter bleiben unveraendert)
  4. LibreOffice          - Kopie durchrechnen, um die Ergebnisse zu pruefen
  5. inject_cache         - berechnete Werte in Zellen und Diagrammen speichern
Benoetigt: openpyxl, lxml, LibreOffice (soffice) mit Calc.
"""
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from build_sheets import SRC_FIRST_ROW, SRC_ROWS, District, build
from inject_cache import ERRORS, inject
from merge_into_original import merge

EXCLUDE = {"Bezirksleiter FBÜW"}
# Blattname in Bedarf_Bestand_LST -> Bezirk im Zielzustand (sonst gleicher Name)
ZZ_NAMES = {
    "Limburg (KRM)": "Limburg 1 (SFS KC)",   # Bezirk im Blatt: "Limburg SFS"
    "Limburg 2": "Limburg 2",
    "FFM West": "FFM West/Höchst",
    "FFM FA": "FFM Abstellbahnhof",
}

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


def find_districts(src_wb, zz_ws):
    """Bezirks-Blaetter = Blaetter mit "Ziel-Qualifikation" in Spalte G der ersten Zeilen."""
    zz_names = {str(c.value).strip() for c in zz_ws[1] if c.value}
    out = []
    for ws in src_wb.worksheets:
        if ws.title in EXCLUDE:
            continue
        if not any(str(ws.cell(r, 7).value or "").strip() == "Ziel-Qualifikation" for r in range(1, 4)):
            continue
        name = ZZ_NAMES.get(ws.title, ws.title)
        if name not in zz_names:
            print(f"Hinweis: Kein Bezirk „{name}“ im Zielzustand – Soll für „{ws.title}“ bleibt leer (Einstellungen).")
            name = ""
        out.append(District(ws.title, name))
    return out


def chart_ymax(src_wb, zz_ws, districts):
    """Gemeinsame Skala fuer die kleinen Verlaufsdiagramme: max(Mitarbeiter, Soll) je Bezirk + Luft."""
    col_of = {str(c.value).strip(): c.column for c in zz_ws[1] if c.value}
    labels = {str(zz_ws.cell(r, 1).value or "").strip(): r for r in range(1, zz_ws.max_row + 1)}
    top = 0
    for d in districts:
        ws = src_wb[d.sheet]
        staff = sum(1 for r in range(SRC_FIRST_ROW, SRC_FIRST_ROW + SRC_ROWS)
                    if ws.cell(r, 2).value and str(ws.cell(r, 2).value).strip() != "Name"
                    and (ws.cell(r, 7).value or ws.cell(r, 10).value))
        soll = 0
        if d.zz_name in col_of:
            c = col_of[d.zz_name] + 3
            for cat in ("Wmech", "SigMech", "SigMech RBEG", "Teamleiter"):
                v = zz_ws.cell(labels.get(cat, 0) or 1, c).value
                soll += v if isinstance(v, (int, float)) else 0
        top = max(top, staff, soll)
    return int(math.ceil(top * 1.15 / 5.0) * 5)


def main(src, zz_path, out):
    src_wb = load_workbook(src, data_only=True)
    zz_wb = load_workbook(zz_path)
    zz_ws = zz_wb.worksheets[0]
    zz_vals = load_workbook(zz_path, data_only=True).worksheets[0]
    districts = find_districts(src_wb, zz_ws)
    ymax = chart_ymax(src_wb, zz_vals, districts)
    with tempfile.TemporaryDirectory(prefix="dashboard-") as tmp:
        new = os.path.join(tmp, "neue_blaetter.xlsx")
        merged = os.path.join(tmp, "merged.xlsx")
        calc = os.path.join(tmp, "calc.xlsx")
        zz_name = re.sub(r"^[0-9a-f]{8}-", "", Path(zz_path).name)    # Praefix hochgeladener Dateien entfernen
        sheet_names = build(districts, zz_ws, zz_name, new, ymax=ymax)
        merge(src, new, merged)
        shutil.copy(merged, calc)
        libreoffice_recalc(calc)
        # Pruefen: keine Fehlerwerte in den neuen Blaettern
        wb = load_workbook(calc, data_only=True)
        errors = [f"{n}!{c.coordinate}" for n in sheet_names for row in wb[n].iter_rows() for c in row
                  if isinstance(c.value, str) and c.value in ERRORS]
        if errors:
            raise RuntimeError(f"Formelfehler in den neuen Blättern: {errors[:20]}")
        info = inject(merged, calc, out, sheet_names)
        if os.environ.get("KEEP_CALC"):
            shutil.copy(calc, os.environ["KEEP_CALC"])
    print(f"{len(districts)} Bezirke, {len(sheet_names)} neue Blätter, {info['cells']} Formelzellen, "
          f"{info['charts']} Diagramme -> {out}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
