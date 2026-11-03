"""Erzeugt die vollautomatische Personal-, Schicht-, Bereitschafts-, Urlaubs-, Qualifikations- und
Ressourcenplanung (Excel) auf Basis der bestehenden PEP-Datei / des Bereitschaftskalenders.

    python build_personalplanung.py --quelle "<PEP_2026_OL_FH-FG.xlsx>" --out Personalplanung_2027.xlsx
                                    [--overrides seed_overrides.json] [--demo] [--no-recalc]

Ablauf:
  1. Teamaufstellung aus der Quelldatei lesen (Blatt "Bereitschaft 2027" bzw. "Kalender 2027 Hessen")
  2. Alle Blaetter mit openpyxl erzeugen (nur Formeln, keine festen Ergebnisse)
  3. Kopie mit LibreOffice durchrechnen und auf Formelfehler pruefen
  4. Berechnete Werte als Cache in die Datei schreiben (Vorschau/geschuetzte Ansicht zeigt Zahlen;
     Excel rechnet beim Oeffnen ohnehin neu)
--demo fuellt klar gekennzeichnete Testdaten (Abwesenheiten, Schulungsdaten, Vertretung, Buchungen) ein.
"""
import argparse
import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.properties import CalcProperties

from pp_base import Book
from pp_source import apply_overrides, override_notes, read_org_text, read_teams

SHEETS = ["Start", "Dashboard", "Einstellungen", "Mitarbeiterstamm", "Team- und Funktionsmatrix",
          "Qualifikationen", "Urlaub und Abwesenheiten", "Vertretungen", "Zeitbuchungen", "Schichtplan 2027",
          "PEP 2027", "Einsatzfähigkeitsprüfung", "Überstunden und JAZ", "Personalbedarfsplanung", "Statistiken",
          "Dienstverteilung", "Urlaubsstatistik", "Krankenstatistik", "Monatsbericht", "Jahresbericht",
          "Einsatzfähigkeitsbericht", "Druck Jahr", "Druck Monat", "Druck Quartal", "Druck Team",
          "Druck Bereitschaft", "Druck Qualifikation", "Hilfsberechnung"]

ERRORS = {"#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A", "Err:502", "Err:504", "Err:508"}

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def libreoffice_recalc(path, timeout=1800):
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as prof:
        url = Path(prof).as_uri()
        subprocess.run(["soffice", "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                       env=env, capture_output=True, timeout=180)
        macro_dir = Path(prof) / "user" / "basic" / "Standard"
        macro_dir.mkdir(parents=True, exist_ok=True)
        (macro_dir / "Module1.xba").write_text(MACRO)
        before = os.stat(path).st_mtime_ns
        subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation={url}",
                        "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
                        str(Path(path).absolute())], env=env, capture_output=True, timeout=timeout)
        if os.stat(path).st_mtime_ns == before:
            raise RuntimeError("LibreOffice hat die Datei nicht neu berechnet.")


def build(quelle, out, overrides=None, demo=False, only=None):
    import pp_input as I
    import pp_plan as P
    people = apply_overrides(read_teams(quelle), overrides)
    wb = Workbook()
    wb.remove(wb.active)
    for s in SHEETS:
        wb.create_sheet(s)
    book = Book(wb)
    demo_data = None
    if demo:
        import pp_demo
        demo_data = pp_demo.make(people)
    I.build_settings(book, wb["Einstellungen"])
    I.build_stamm(book, wb["Mitarbeiterstamm"], people, demo=demo)
    I.build_matrix(book, wb["Team- und Funktionsmatrix"], people)
    I.build_quali(book, wb["Qualifikationen"], people, demo_data["quali"] if demo else None)
    I.build_abwesenheit(book, wb["Urlaub und Abwesenheiten"], demo_data["abw"] if demo else None)
    I.build_vertretung(book, wb["Vertretungen"], demo_data["vert"] if demo else None)
    I.build_zeitbuchungen(book, wb["Zeitbuchungen"], demo_data["zb"] if demo else None)
    P.build_calc(book, wb["Hilfsberechnung"])
    P.build_plan(book, wb["Schichtplan 2027"])
    pep_sum_row = P.build_pep(book, wb["PEP 2027"])
    P.build_ef(book, wb["Einsatzfähigkeitsprüfung"])
    P.build_zeit(book, wb["Überstunden und JAZ"])
    P.build_bedarf(book, wb["Personalbedarfsplanung"])
    try:
        import pp_reports as R
        import pp_print as D
    except ImportError:
        R = D = None
    if R and only is None:
        notes = override_notes(people)
        if demo:
            notes.append("DEMO-DATEI: Abwesenheiten, Schulungsdaten, Vertretung, Zeitbuchungen und Personalnummern "
                         "sind frei erfundene Testdaten (Bemerkung „DEMO“) und dienen nur zur Veranschaulichung.")
        R.build_all(book, wb, people, demo, notes)
        D.build_all(book, wb, read_org_text(quelle))
    else:
        # Platzhalter fuer Namen aus noch fehlenden Blaettern (nur Entwicklung)
        ws = wb["Statistiken"]
        book.name_ws("ST_Urlaub", ws, "A6:A65")
        book.name_ws("Berichtsmonat", wb["Dashboard"], "C3")
        wb["Dashboard"]["C3"] = 1
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.active = 0
    wb.save(out)
    return wb, people


def check_errors(calc_path):
    wb = load_workbook(calc_path, data_only=True)
    errs = {}
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    errs.setdefault(c.value, []).append(f"{ws.title}!{c.coordinate}")
                    n += 1
    return n, errs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quelle", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--overrides")
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--no-recalc", action="store_true")
    ap.add_argument("--keep-calc", help="berechnete LibreOffice-Kopie zusaetzlich hier ablegen")
    a = ap.parse_args()
    t0 = time.time()
    with tempfile.TemporaryDirectory(prefix="pp-") as tmp:
        raw = os.path.join(tmp, "raw.xlsx")
        build(a.quelle, raw, a.overrides, a.demo)
        print(f"Arbeitsmappe erzeugt ({time.time() - t0:.0f}s)")
        if a.no_recalc:
            shutil.copy(raw, a.out)
            return
        calc = os.path.join(tmp, "calc.xlsx")
        shutil.copy(raw, calc)
        t1 = time.time()
        libreoffice_recalc(calc)
        print(f"LibreOffice-Neuberechnung ({time.time() - t1:.0f}s)")
        if a.keep_calc:
            shutil.copy(calc, a.keep_calc)
        n, errs = check_errors(calc)
        if n:
            for k, v in errs.items():
                print(f"  {k}: {len(v)} z. B. {v[:8]}")
            raise SystemExit(f"{n} Formelfehler – Datei nicht geschrieben.")
        from inject_cache import inject
        info = inject(raw, calc, a.out)
        print(f"OK: 0 Formelfehler, {info['cells']} Formelzellen mit Werten, {info['charts']} Diagramme -> {a.out}")


if __name__ == "__main__":
    main()
