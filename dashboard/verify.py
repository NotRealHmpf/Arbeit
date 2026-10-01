"""Vergleicht die von LibreOffice berechneten Dashboard-Werte mit der unabhaengigen Python-Nachrechnung (reference.py).

    python verify.py <berechnete Datei.xlsx> <Bedarf_Bestand_LST.xlsx> <LST_Zielzustand_gesamt.xlsx>
"""
import sys

from openpyxl import load_workbook

import build_sheets as B
import reference as R


def gap_text(fehlt, mehr):
    if fehlt == 0 and mehr == 0:
        return "Soll erreicht"
    s = "es fehlt 1" if fehlt == 1 else (f"es fehlen {fehlt}" if fehlt > 1 else "")
    if fehlt > 0 and mehr > 0:
        s += " · "
    if mehr > 0:
        s += f"{mehr} mehr als Soll"
    return s


def num(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def verify(calc_path, src_path, zz_path, districts):
    wb = load_workbook(calc_path, data_only=True)
    bj = int(wb["Einstellungen"]["C4"].value)
    davon = int(wb["Einstellungen"]["C5"].value)
    ref, pruef = R.compute(src_path, zz_path, [(d.sheet, d.zz_name) for d in districts], bj, davon)
    errors = []
    checks = 0

    def eq(where, got, exp):
        nonlocal checks
        checks += 1
        got = num(got)
        if got is None:
            got = ""
        if got != exp:
            errors.append(f"{where}: Dashboard {got!r} ≠ Python {exp!r}")

    # ---- Daten je Mitarbeiter
    ws = wb["Daten"]
    DC = B.DC
    persons = {(p.sheet, p.row): p for d in districts for p in ref[d.sheet]["persons"]}
    seen = set()
    for r in range(2, ws.max_row + 1):
        if num(ws[f"{DC['ok']}{r}"].value) != 1:
            continue
        key = (ws[f"{DC['bez']}{r}"].value, num(ws[f"{DC['row']}{r}"].value))
        seen.add(key)
        p = persons.get(key)
        if p is None:
            errors.append(f"Daten Zeile {r}: {key} fehlt in der Python-Rechnung")
            continue
        st, zy, grp = R.status(p, bj)
        eq(f"Daten {key} Status", ws[f"{DC['stat']}{r}"].value, st)
        eq(f"Daten {key} Gruppe", ws[f"{DC['zgrp']}{r}"].value, grp)
        for k, t in enumerate(R.thresholds(bj)):
            eq(f"Daten {key} Stufe {k}", ws[f"{DC[f'lv{k}']}{r}"].value, p.level(t))
            eq(f"Daten {key} Teamleiter {k}", ws[f"{DC[f'tl{k}']}{r}"].value, p.tl(t))
        eq(f"Daten {key} Hinweise", ws[f"{DC['hint']}{r}"].value, R.hints(p, bj))
    for key in persons:
        if key not in seen:
            errors.append(f"Daten: {key} fehlt im Dashboard")

    # ---- Bezirksseiten
    for d in districts:
        res = ref[d.sheet]
        ws = wb[B.district_title(d)]
        rows = {key: B.B_SI_HDR + 1 + i for i, key in enumerate(B.SI_ROWS)}
        for key in ["Azubi"] + B.CMP + [B.NO_STAGE, "Köpfe"]:
            for k, col in enumerate(B.TIME_LETTERS):
                eq(f"{ws.title}!{col}{rows[key]} ({key})", ws[f"{col}{rows[key]}"].value, res["counts"][key][k])
        for key in B.CMP:
            eq(f"{ws.title} Soll {key}", ws[f"{B.T_SOLL}{rows[key]}"].value, res["soll"][key])
            for k, col in enumerate(B.TIME_LETTERS):
                eq(f"{ws.title} Abweichung {key} {k}", ws[f"{col}{B.B_GAP_HDR + 1 + B.CMP.index(key)}"].value,
                   gap_text(res["fehlt"][key][k], res["mehr"][key][k]))
            eq(f"{ws.title} Bewertung heute {key}", ws[f"N{rows[key]}"].value, gap_text(res["fehlt"][key][0], res["mehr"][key][0]))
            eq(f"{ws.title} Bewertung Plan {key}", ws[f"Q{rows[key]}"].value, gap_text(res["fehlt"][key][-1], res["mehr"][key][-1]))
        for k, col in enumerate(B.TIME_LETTERS):
            eq(f"{ws.title} Fehlende Stellen {k}", ws[f"{col}{rows['Fehlend']}"].value, res["fehlend"][k])
            eq(f"{ws.title} Mehr als Soll {k}", ws[f"{col}{rows['Mehr']}"].value, res["mehr_sum"][k])
        for i, g in enumerate(B.GROUPS):
            for code in range(1, 10):
                col = B.get_column_letter(B.ST_COL1 + code - 1)
                eq(f"{ws.title} Status {g}/{code}", ws[f"{col}{B.B_ST_FIRST + i}"].value, res["status"].get((g, code), 0))
        people = res["persons"]
        st_count = lambda c: sum(1 for p in people if R.status(p, bj)[0] == c)
        eq(f"{ws.title} KPI Mitarbeiter", ws["B6"].value, len(people))
        eq(f"{ws.title} KPI Fertig", ws["E6"].value, st_count(1))
        eq(f"{ws.title} KPI In Ausbildung", ws["H6"].value, sum(st_count(c) for c in range(2, 7)))
        eq(f"{ws.title} KPI davon", ws["H7"].value, f"davon Abschluss {davon}: {res['davon']}")
        eq(f"{ws.title} KPI Überfällig", ws["K6"].value, st_count(7))
        eq(f"{ws.title} KPI Fehlt", ws["N6"].value, st_count(8))
        eq(f"{ws.title} KPI Fehlende Stellen", ws["Q6"].value, res["fehlend"][-1])
        order = sorted(people, key=lambda p: (R.status(p, bj)[0], p.row))
        for k in range(B.LIST_ROWS):
            r = B.B_LIST_FIRST + k
            exp = order[k].name if k < len(order) else ""
            eq(f"{ws.title} Liste {k + 1}", ws[f"C{r}"].value, exp)
            if k < len(order):
                p = order[k]
                eq(f"{ws.title} Liste {k + 1} Stufe heute", ws[f"K{r}"].value,
                   R.stage_name(p.level(0)) + (" + TL" if p.tl(0) else ""))
                eq(f"{ws.title} Liste {k + 1} Stufe Plan", ws[f"M{r}"].value,
                   R.stage_name(p.level(9998)) + (" + TL" if p.tl(9998) else ""))

    # ---- Uebersicht
    ws = wb["Übersicht"]
    for i, d in enumerate(districts):
        r = B.OV_FIRST + i
        res = ref[d.sheet]
        for code in range(1, 10):
            col = B.get_column_letter(3 + code)
            eq(f"Übersicht {d.sheet} Status {code}", ws[f"{col}{r}"].value,
               sum(1 for p in res["persons"] if R.status(p, bj)[0] == code))
        eq(f"Übersicht {d.sheet} Summe", ws[f"M{r}"].value, len(res["persons"]))
        eq(f"Übersicht {d.sheet} davon", ws[f"O{r}"].value, res["davon"])
        eq(f"Übersicht {d.sheet} fehlend heute", ws[f"Q{r}"].value, res["fehlend"][0])
        eq(f"Übersicht {d.sheet} fehlend Plan", ws[f"R{r}"].value, res["fehlend"][-1])
        eq(f"Übersicht {d.sheet} mehr Plan", ws[f"S{r}"].value, res["mehr_sum"][-1])
    total = sum(len(ref[d.sheet]["persons"]) for d in districts)
    eq("Übersicht KPI Mitarbeiter", ws["B6"].value, total)
    eq("Übersicht KPI Fehlende Stellen", ws["Q6"].value, sum(ref[d.sheet]["fehlend"][-1] for d in districts))

    # ---- Soll-Ist (LST gesamt)
    ws = wb["Soll-Ist"]
    hdr = None
    for r in range(1, ws.max_row + 1):
        if ws[f"B{r}"].value == "Stufe" and ws[f"{B.T_SOLL}{r}"].value:
            hdr = r
            break
    if hdr is None:
        errors.append("Soll-Ist: Tabelle LST gesamt nicht gefunden")
    else:
        rows = {key: hdr + 1 + i for i, key in enumerate(B.SI_ROWS)}
        for key in ["Azubi"] + B.CMP + [B.NO_STAGE, "Köpfe"]:
            for k, col in enumerate(B.TIME_LETTERS):
                eq(f"Soll-Ist gesamt {key} {k}", ws[f"{col}{rows[key]}"].value,
                   sum(ref[d.sheet]["counts"][key][k] for d in districts))
        for key in B.CMP:
            f = sum(ref[d.sheet]["fehlt"][key][-1] for d in districts)
            m = sum(ref[d.sheet]["mehr"][key][-1] for d in districts)
            eq(f"Soll-Ist gesamt Bewertung {key}", ws[f"Q{rows[key]}"].value, gap_text(f, m))
        for k, col in enumerate(B.TIME_LETTERS):
            eq(f"Soll-Ist gesamt fehlend {k}", ws[f"{col}{rows['Fehlend']}"].value,
               sum(ref[d.sheet]["fehlend"][k] for d in districts))

    # ---- Pruefliste
    ws = wb["Prüfliste"]
    for k in range(B.PRUEF_ROWS):
        r = B.PF_HDR + 1 + k
        if k < len(pruef):
            sheet, row, name, vor, ziel, hint = pruef[k]
            eq(f"Prüfliste {k + 1} Bezirk", ws[f"C{r}"].value, sheet)
            eq(f"Prüfliste {k + 1} Name", ws[f"E{r}"].value, name)
            eq(f"Prüfliste {k + 1} Zeile", ws[f"K{r}"].value, row)
            eq(f"Prüfliste {k + 1} Hinweise", ws[f"L{r}"].value, hint)
        else:
            eq(f"Prüfliste {k + 1} leer", ws[f"C{r}"].value, "")
    return errors, checks, ref, pruef, bj


if __name__ == "__main__":
    from build_dashboard import find_districts
    calc, src, zz = sys.argv[1:4]
    districts = find_districts(load_workbook(src, data_only=True), load_workbook(zz).worksheets[0])
    errors, checks, ref, pruef, bj = verify(calc, src, zz, districts)
    print(f"{checks} Werte verglichen, {len(errors)} Abweichungen")
    for e in errors[:50]:
        print("  ", e)
