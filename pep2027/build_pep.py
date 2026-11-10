"""Erzeugt das Personalplanungs- und Einsatzsteuerungssystem (PEP) als Excel-Datei mit Makros.

    python build_pep.py <original.xlsx> <ausgabe.xlsm> [--namen namen.csv] [--ohne-vba] [--schnell]

<original.xlsx>  Datei mit den Blaettern "PEP_2026" und "Bereitschaft 2027" (Teamzuordnung Ma1..Ma28)
--namen          optionale CSV (Semikolon): Platzhalter;Name;Personalnummer – wird nicht eingecheckt
--ohne-vba       nur .xlsx ohne Makros erzeugen
--schnell        ohne LibreOffice-Durchrechnung (keine gespeicherten Werte, keine Formelpruefung)
--vba-export DIR VBA-Quelltexte zusaetzlich als .bas/.txt ablegen (Rueckfallebene: manueller Import)

Ablauf: openpyxl baut alle Blaetter -> LibreOffice rechnet eine Kopie durch -> Pruefung auf Formelfehler ->
berechnete Werte werden als Cache in die Datei geschrieben -> VBA-Projekt wird eingebettet (.xlsm).
"""
import argparse
import os
import shutil
import sys
import tempfile
import time

from pep.workbook import build_workbook


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--namen")
    ap.add_argument("--ohne-vba", action="store_true")
    ap.add_argument("--schnell", action="store_true")
    ap.add_argument("--teil", default="all")
    ap.add_argument("--vba-export", help="VBA-Quelltexte zusätzlich in diesen Ordner schreiben")
    a = ap.parse_args()

    t0 = time.time()
    if a.vba_export:
        from pep.makros import export_sources
        print("VBA-Quelltexte:", ", ".join(export_sources(a.vba_export)))
    wb, ctx = build_workbook(a.src, a.namen, parts=a.teil)
    too_long = [f"{ws.title}!{c.coordinate} ({len(c.value)})" for ws in wb.worksheets for row in ws.iter_rows()
                for c in row if isinstance(c.value, str) and c.value.startswith("=") and len(c.value) > 8000]
    if too_long:
        sys.exit("Formeln länger als Excel erlaubt (8192 Zeichen): " + ", ".join(too_long[:10]))
    with tempfile.TemporaryDirectory(prefix="pep-") as tmp:
        raw = os.path.join(tmp, "raw.xlsx")
        wb.save(raw)
        print(f"Blätter erzeugt ({time.time() - t0:.0f} s)")
        if a.schnell:
            shutil.copy(raw, a.out)
            return
        from pep.lo import recalc
        from pep.package import check_errors, inject_values, make_xlsm
        calc = os.path.join(tmp, "calc.xlsx")
        shutil.copy(raw, calc)
        recalc(calc)
        print(f"LibreOffice-Durchrechnung fertig ({time.time() - t0:.0f} s)")
        errors = check_errors(calc, allow_na_cells={("Dashboard", c) for c in ("AB", "AC", "AD")})
        if errors:
            print("Formelfehler:", len(errors))
            for e in errors[:40]:
                print("  ", e)
            sys.exit(1)
        cached = os.path.join(tmp, "cached.xlsx")
        n = inject_values(raw, calc, cached)
        print(f"{n} Formelzellen mit Werten gespeichert")
        if a.ohne_vba or not a.out.lower().endswith(".xlsm"):
            shutil.copy(cached, a.out)
        else:
            from pep.makros import build_vba_project
            make_xlsm(cached, a.out, build_vba_project(ctx))
    print(f"fertig: {a.out} ({time.time() - t0:.0f} s)")


if __name__ == "__main__":
    main()
