"""Prueft, dass das Dashboard per Formel auf Aenderungen reagiert.

    python test_dynamic.py <Dashboard-Datei.xlsx>

Aendert in einer Testkopie (die Original-Datei bleibt unberuehrt) einige Eintraege in den Bezirks-Blaettern,
ein Soll im Blatt Zielzustand und das Bezugsjahr, rechnet mit LibreOffice neu und vergleicht alle Werte mit der
Python-Nachrechnung fuer die geaenderten Daten. Die Testkopie wird mit openpyxl gespeichert - nur fuer diesen Test.
"""
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from build_dashboard import DISTRICTS, libreoffice_recalc
from verify import verify


def first_row(ws, test):
    for r in range(2, 102):
        if str(ws[f"B{r}"].value or "").strip() not in ("", "Name") and test(ws, r):
            return r
    return None


def main(path):
    with tempfile.TemporaryDirectory(prefix="dyn-") as tmp:
        test = str(Path(tmp) / "test.xlsx")
        wb = load_workbook(path)
        changes = []
        s1, s2, s3 = DISTRICTS[1][0], DISTRICTS[7][0], DISTRICTS[8][0]
        ws = wb[s1]                       # geplantes RBEG-Jahr -> "x"
        r = first_row(ws, lambda w, r: str(w[f"P{r}"].value or "").strip().isdigit())
        ws[f"P{r}"] = "x"
        changes.append(f"{s1}!P{r} = x")
        ws = wb[s2]                       # Jahr in der Ziel-Spalte nachtragen / ueberschreiben
        r = first_row(ws, lambda w, r: str(w[f"G{r}"].value or "").strip() == "Weichmech" and not w[f"N{r}"].value)
        if r:
            ws[f"N{r}"] = 28
            changes.append(f"{s2}!N{r} = 28")
        ws = wb[s3]                       # neuer Mitarbeiter in der ersten freien Zeile
        r = next(r for r in range(3, 102) if not ws[f"B{r}"].value)
        for c, v in {"B": "Test", "C": "Neu", "G": "Sigmech", "L": "x", "M": 27, "O": "29", "V": "0"}.items():
            ws[f"{c}{r}"] = v
        changes.append(f"{s3}: neue Zeile {r}")
        zz = wb["Zielzustand"]            # Soll SigMech des zweiten Bezirks + 2
        col = next(c.column for c in zz[1] if str(c.value or "").strip() == DISTRICTS[1][1]) + 3
        row = next(r for r in range(1, 30) if zz.cell(r, 1).value == "SigMech")
        zz.cell(row, col).value = (zz.cell(row, col).value or 0) + 2
        changes.append(f"Zielzustand {zz.cell(row, col).coordinate} + 2")
        wb["Einstellungen"]["C4"] = 2027  # Bezugsjahr fest auf 2027
        changes.append("Bezugsjahr = 2027")
        wb.save(test)
        libreoffice_recalc(test)
        ck, _ = verify(test, test, test, 2027, zz_sheet="Zielzustand")
        print("Änderungen:", "; ".join(changes))
        print(f"{ck.n} Werte verglichen, {len(ck.bad)} Abweichungen")
        for b in ck.bad[:40]:
            print("  ", b)
        return 1 if ck.bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
