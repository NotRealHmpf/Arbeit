"""Erzeugt aus Bedarf_Bestand_LST und LST_Zielzustand_gesamt eine Kopie mit Excel-Dashboard.

    python build_dashboard.py <Bedarf_Bestand_LST.xlsx> <LST_Zielzustand_gesamt.xlsx> <ausgabe.xlsx>

Ablauf:
  1. Bezirks-Blaetter erkennen, Spaltenueberschriften pruefen, dem Zielzustand zuordnen
  2. build_sheets.py      - neue Blaetter (Uebersicht, Soll-Ist, Pruefliste, Bezirk je Bezirk, Daten, Zielzustand, Einstellungen)
  3. merge_into_original  - Blaetter direkt in das .xlsx-Paket des Originals einsetzen (Original-Blaetter bleiben unveraendert)
  4. LibreOffice          - eine Kopie durchrechnen
  5. verify.py            - Ergebnisse mit der unabhaengigen Python-Nachrechnung vergleichen
  6. inject_cache         - berechnete Werte in Zellen und Diagrammen speichern
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

import reference as R
from build_sheets import SRC_FIRST_ROW, District, build
from inject_cache import ERRORS, inject
from merge_into_original import merge

# Blaetter, die keine Bezirks-Blaetter sind (werden ignoriert)
EXCLUDE = {"Bedarf_Bestand", "PLR2025", "Mifri", "Tabelle1", "Bezirksleiter FBÜW"}
# Blattname in Bedarf_Bestand_LST -> Bezirk im Zielzustand (Zeile 1); sonst gleicher Name
ZZ_NAMES = {
    "Limburg (KRM)": "Limburg 1 (SFS KC)",
    "Limburg 2": "Limburg 2",
    "FFM West": "FFM West/Höchst",
    "FFM Süd": "FFM Süd",
    "FFM FA": "FFM Abstellbahnhof",
    "FFM Hbf": "FFM Hbf",
}
# Erwartete Ueberschriften (Anfang des Textes, ohne Gross-/Kleinschreibung)
EXPECTED = {"B": "name", "C": "vorname", "G": "ziel-qualifikation", "J": "ist-qualifikation", "L": "azubi",
            "M": "arb lst", "N": "weichenmechaniker", "O": "signalmechaniker", "P": "signalmechaniker rbeg",
            "S": "teamleiter", "U": "örtl. verwendungsprüfung wmech", "V": "örtl. verw"}

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def strip_external_links(path):
    """Nur fuer die Rechenkopie: externe Verknuepfungen (SharePoint) entfernen, sonst versucht LibreOffice sie zu laden
    und haengt. Das Dashboard liest nur B, C, G, J, L-V der Bezirks-Blaetter; die haengen nicht davon ab."""
    import zipfile
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(path) as z:
        parts = {i.filename: z.read(i.filename) for i in z.infolist()}
    drop = {n for n in parts if n.startswith("xl/externalLinks/")}
    wb = etree.fromstring(parts["xl/workbook.xml"])
    for el in wb.findall(f"{{{ns}}}externalReferences"):
        wb.remove(el)
    parts["xl/workbook.xml"] = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)
    rels = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
    for rel in list(rels):
        if rel.get("Type", "").endswith("/externalLink"):
            rels.remove(rel)
    parts["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    ct = etree.fromstring(parts["[Content_Types].xml"])
    for ov in list(ct):
        if (ov.get("PartName") or "").startswith("/xl/externalLinks/"):
            ct.remove(ov)
    parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            if name not in drop:
                z.writestr(name, data)


def libreoffice_recalc(path, timeout=1200):
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


def header_row(ws):
    for r in (1, 2, 3):
        if str(ws.cell(r, 7).value or "").strip().lower() == "ziel-qualifikation":
            return r
    return None


def find_districts(src_wb, zz_ws):
    """Bezirks-Blaetter = Blaetter mit "Ziel-Qualifikation" in Spalte G (Zeile 1-3); Reihenfolge wie im Zielzustand."""
    zz_cols = {str(c.value).strip(): c.column for c in zz_ws[1] if c.value}
    out = []
    for ws in src_wb.worksheets:
        if ws.title.strip() in EXCLUDE:
            continue
        hr = header_row(ws)
        if hr is None:
            continue
        for col, exp in EXPECTED.items():
            got = re.sub(r"\s+", " ", str(ws[f"{col}{hr}"].value or "")).strip().lower()
            if not got.startswith(exp):
                raise RuntimeError(f"Blatt „{ws.title}“: Spalte {col} heißt „{got}“, erwartet „{exp}…“")
        if hr < SRC_FIRST_ROW - 1 or hr > SRC_FIRST_ROW:
            raise RuntimeError(f"Blatt „{ws.title}“: Kopfzeile in Zeile {hr} – erwartet Zeile 1 oder 2")
        name = ZZ_NAMES.get(ws.title, ws.title)
        if name not in zz_cols:
            print(f"Hinweis: Kein Bezirk „{name}“ im Zielzustand – Soll für „{ws.title}“ bleibt leer (Einstellungen).")
            name = ""
        out.append(District(ws.title, name))
    out.sort(key=lambda d: zz_cols.get(d.zz_name, 10 ** 6))
    return out


def chart_ymax(src_path, zz_path, districts, bj):
    """Gemeinsame Skala der kleinen Verlaufsdiagramme: hoechste Saeule (Stufen + Teamleiter) bzw. Ziel, plus Luft."""
    ref, _ = R.compute(src_path, zz_path, [(d.sheet, d.zz_name) for d in districts], bj)
    top = 0
    for res in ref.values():
        n = len(res["counts"]["Köpfe"])
        for k in range(n):
            top = max(top, res["counts"]["Köpfe"][k] + res["counts"]["Teamleiter"][k])
        top = max(top, sum(res["soll"].values()))
    return int(math.ceil(top * 1.15 / 5.0) * 5)


def main(src, zz_path, out):
    import datetime as dt
    import verify as V
    src_wb = load_workbook(src, data_only=True)
    zz_wb = load_workbook(zz_path)
    zz_ws = zz_wb.worksheets[0]
    districts = find_districts(src_wb, zz_ws)
    ymax = chart_ymax(src, zz_path, districts, dt.date.today().year)
    with tempfile.TemporaryDirectory(prefix="dashboard-") as tmp:
        new = os.path.join(tmp, "neue_blaetter.xlsx")
        merged = os.path.join(tmp, "merged.xlsx")
        calc = os.path.join(tmp, "calc.xlsx")
        zz_name = re.sub(r"^[0-9a-f]{8}-", "", Path(zz_path).name)    # Praefix hochgeladener Dateien entfernen
        sheet_names = build(districts, zz_ws, zz_name, new, ymax=ymax)
        merge(src, new, merged)
        shutil.copy(merged, calc)
        strip_external_links(calc)
        libreoffice_recalc(calc)
        if os.environ.get("KEEP_CALC"):
            shutil.copy(calc, os.environ["KEEP_CALC"])
        wb = load_workbook(calc, data_only=True)
        errors = [f"{n}!{c.coordinate}" for n in sheet_names for row in wb[n].iter_rows() for c in row
                  if isinstance(c.value, str) and c.value in ERRORS]
        if errors:
            raise RuntimeError(f"Formelfehler in den neuen Blättern: {errors[:20]}")
        diffs, checks, _, pruef, bj = V.verify(calc, src, zz_path, districts)
        print(f"Prüfung: {checks} Werte mit der Python-Nachrechnung verglichen, {len(diffs)} Abweichungen")
        if diffs:
            raise RuntimeError("Abweichungen:\n  " + "\n  ".join(diffs[:40]))
        info = inject(merged, calc, out, sheet_names)
    print(f"{len(districts)} Bezirke, {len(sheet_names)} neue Blätter, {info['cells']} Formelzellen, "
          f"{info['charts']} Diagramme -> {out}")
    return districts


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
