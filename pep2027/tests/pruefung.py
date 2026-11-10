"""Fachliche Pruefungen der erzeugten Arbeitsmappe (mit LibreOffice durchgerechnet).

    python tests/pruefung.py <original.xlsx> [--namen namen.csv]

Prueft u. a.: keine Formelfehler, Rotation = Farben des Bereitschaftskalenders, Feiertage/Arbeitstage Hessen,
Freistellungslogik inkl. manueller Anpassung, Prioritaeten, Einsatzfaehigkeit mit interner Vertretung,
Jahreswechsel (Planungsjahr 2028, Schaltjahr) ohne Fehler.
"""
import argparse
import datetime as dt
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from openpyxl import load_workbook  # noqa: E402

from pep import model as M  # noqa: E402
from pep.lo import recalc  # noqa: E402
from pep.package import check_errors  # noqa: E402
from pep.source import check_rotation, rotation_team  # noqa: E402
from pep.workbook import build_workbook  # noqa: E402

NA_OK = {("Dashboard", c) for c in ("AB", "AC", "AD")}
FEHLER = []


def ok(cond, text):
    print(("  ok    " if cond else "  FEHLER ") + text)
    if not cond:
        FEHLER.append(text)


def calc(wb, tmp, name):
    raw = os.path.join(tmp, name + "_raw.xlsx")
    out = os.path.join(tmp, name + ".xlsx")
    wb.save(raw)
    shutil.copy(raw, out)
    recalc(out)
    return out


def row_of(ws, header_row, header):
    for c in ws[header_row]:
        if c.value == header:
            return c.column
    raise KeyError(header)


def table_rows(ws, header_row, first, n):
    hdr = {c.value: c.column for c in ws[header_row] if c.value}
    out = []
    for r in range(first, first + n):
        out.append({h: ws.cell(r, col).value for h, col in hdr.items()})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("--namen")
    a = ap.parse_args()
    tmp = tempfile.mkdtemp(prefix="pep-test-")

    print("1) Rotation gegen Bereitschaftskalender")
    n, diffs = check_rotation(a.src)
    ok(n == 365 and not diffs, f"{n} Tage geprüft, Abweichungen: {len(diffs)}")

    print("2) Standardaufbau 2027")
    wb, ctx = build_workbook(a.src, a.namen)
    # Testszenario: Ma14 (Handwerker Gelb) ist zusaetzlich TVT-Fahrer -> interne Vertretung bei Krankheit von Ma13
    fm = wb["Funktionsmatrix"]
    tvt_col = next(c.column for c in fm[4] if c.value == "TVT-Fahrer")
    for r in range(5, fm.max_row + 1):
        if fm.cell(r, 2).value == "Ma14":
            fm.cell(r, tvt_col).value = "JA"
    f = calc(wb, tmp, "y2027")
    errs = check_errors(f, NA_OK)
    ok(not errs, f"Formelfehler: {len(errs)} {errs[:5]}")
    v = load_workbook(f, data_only=True)
    sch = table_rows(v["Schichtplanung"], 6, 7, 366)
    ok(sum(1 for d in sch if d["Arbeitstag"] == "Ja") == 255, "255 Arbeitstage in Hessen 2027")
    ok(sch[365]["Datum"] in (None, ""), "Tag 366 im Nicht-Schaltjahr leer")
    hol = {d["Datum"].date(): d["Feiertag"] for d in sch if d["Feiertag"]}
    ok(hol.get(dt.date(2027, 3, 26)) == "Karfreitag" and hol.get(dt.date(2027, 5, 27)) == "Fronleichnam",
       "Karfreitag 26.03. und Fronleichnam 27.05.2027")
    ok(len(hol) == 10, f"10 gesetzliche Feiertage (gefunden {len(hol)})")
    wrong = [d["Datum"] for d in sch[:365] if d["Nachtdienstteam"] != rotation_team(d["Datum"].date())]
    ok(not wrong, "Nachtdienstteam = Rotationsfolge an allen Tagen")
    by = {d["Datum"].date(): d for d in sch[:365]}
    ok(by[dt.date(2027, 1, 4)]["Freigestellt (Vorschlag)"] == "Rot", "Mo 04.01.: Team Rot frei (nach Fr–So)")
    ok(by[dt.date(2027, 1, 8)]["Freigestellt (Vorschlag)"] == "Gelb", "Fr 08.01.: Team Gelb frei (nach Mo–Do)")
    ok(by[dt.date(2027, 1, 5)]["Freigestellt (wirksam)"] in (None, ""), "Beispiel: Rot arbeitet Di 05.01.")
    ok("Rot" in (by[dt.date(2027, 1, 8)]["Freigestellt (wirksam)"] or ""), "Beispiel: Rot stattdessen Fr 08.01. frei")
    ber = v["Berechnung"]
    names = {ber.cell(r, 2).value: r for r in range(M.ROW0, M.ROWN + 1) if ber.cell(r, 1).value == "P"}

    def code(name, day):
        return ber.cell(names[name], M.DAY0 + (day - dt.date(2027, 1, 1)).days).value

    ok(code("Ma7", dt.date(2027, 1, 18)) == 3, "Urlaub wird als U (Code 3) dargestellt")
    ok(code("Ma7", dt.date(2027, 1, 22)) == 10, "Freistellung hat Vorrang vor Urlaub (kein Urlaubstag verbraucht)")
    ok(code("Ma13", dt.date(2027, 1, 8)) == 8, "Krankheit hat Vorrang vor Freistellung")
    ok(code("Ma12", dt.date(2027, 1, 18)) == 9, "Vertretung erscheint als Bereitschaft (grau)")
    ok(code("Ma6", dt.date(2027, 1, 5)) == 0, "aufgehobene Freistellung = regulärer Tagdienst")
    e = table_rows(v["Einsatzfähigkeit"], 7, 8, 366)
    e6 = e[5]
    ok(e6["OK"] == 1 and e6["Stufe"] == 1, "06.01.: TVT krank, Ma14 (TVT-qualifiziert) deckt intern ab → ✅ / 🟡")
    ok("intern vertretbar" in (e6["Fehlermeldung"] or ""), "Fehlermeldung nennt interne Vertretung")
    ok((e6["Ersatz TVT"] or "").startswith("Ma14"), f"Ersatzvorschlag gleiches Team: {e6['Ersatz TVT']}")
    ok(e[17]["Stufe"] == 0, "18.01.: AV-Urlaub durch eingetragene Vertretung gedeckt")
    k = v["Jahreskalender"]
    ok(k.cell(names["Ma7"], M.DAY0 + 17).value == "U", "Kalender zeigt U für Urlaub")
    blau1 = next(ber.cell(r, 2).value for r in range(M.ROW0, M.ROWN + 1)
                 if ber.cell(r, 1).value == "P" and ber.cell(r, 4).value == "Blau")
    ok(code(blau1, dt.date(2027, 1, 15)) == 9 and k.cell(names[blau1], M.DAY0 + 14).value in (None, ""),
       "Bereitschaft ohne Text (nur grau)")
    ok(k.cell(names["Ma9"], M.DAY0 + 26).value == "Arzt", "Besonderheit zeigt ihren Text")

    print("3) Jahreswechsel: Planungsjahr 2028 (Schaltjahr)")
    wb2, _ = build_workbook(a.src, a.namen)
    wb2["Einstellungen"]["C4"] = 2028
    f2 = calc(wb2, tmp, "y2028")
    errs = check_errors(f2, NA_OK)
    ok(not errs, f"Formelfehler 2028: {len(errs)} {errs[:5]}")
    v2 = load_workbook(f2, data_only=True)
    sch2 = table_rows(v2["Schichtplanung"], 6, 7, 366)
    ok(sch2[365]["Datum"] is not None and sch2[365]["Datum"].date() == dt.date(2028, 12, 31), "366 Tage in 2028")
    wrong = [d["Datum"] for d in sch2 if d["Nachtdienstteam"] != rotation_team(d["Datum"].date())]
    ok(not wrong, "Rotation läuft 2028 nahtlos weiter")
    hol2 = {d["Datum"].date(): d["Feiertag"] for d in sch2 if d["Feiertag"]}
    ok(hol2.get(dt.date(2028, 4, 14)) == "Karfreitag" and hol2.get(dt.date(2028, 6, 15)) == "Fronleichnam",
       "Feiertage 2028 (Ostern 16.04.2028) automatisch")
    print()
    print("Ergebnis:", "alle Prüfungen bestanden" if not FEHLER else f"{len(FEHLER)} Fehler")
    shutil.rmtree(tmp, ignore_errors=True)
    sys.exit(1 if FEHLER else 0)


if __name__ == "__main__":
    main()
