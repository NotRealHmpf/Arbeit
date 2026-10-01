"""Erzeugt aus der NFP-Gesamtuebersicht eine Datei mit Excel-Dashboard.

    python build_dashboard.py <original.xlsx> <ausgabe.xlsx>

Ablauf:
  1. build_sheets.py      - neue Blaetter (Uebersicht, Grafik je Bezirk, Daten, Einstellungen) erzeugen
  2. merge_into_original  - Blaetter in die Original-Datei einsetzen (Original bleibt unveraendert)
  3. LibreOffice          - Kopie durchrechnen, um die Ergebnisse zu pruefen
  4. inject_cache         - berechnete Werte in Zellen und Diagrammen speichern
Benoetigt: openpyxl, lxml, LibreOffice (soffice) mit Calc.
"""
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from build_sheets import build
from inject_cache import ERRORS, inject
from merge_into_original import merge

EXCLUDE = {"Bezirksleiter FBÜW"}

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def libreoffice_recalc(path, timeout=600):
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


def main(src, out):
    districts = [n for n in load_workbook(src, read_only=True).sheetnames if n not in EXCLUDE]
    with tempfile.TemporaryDirectory(prefix="dashboard-") as tmp:
        new = os.path.join(tmp, "neue_blaetter.xlsx")
        merged = os.path.join(tmp, "merged.xlsx")
        calc = os.path.join(tmp, "calc.xlsx")
        sheet_names = build(districts, new)
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
    print(f"{len(districts)} Bezirke, {len(sheet_names)} neue Blätter, {info['cells']} Formelzellen, "
          f"{info['charts']} Diagramme -> {out}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
