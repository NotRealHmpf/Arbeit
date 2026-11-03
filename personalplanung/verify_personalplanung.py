"""Unabhaengige Gegenpruefung der berechneten Arbeitsmappe.

    python verify_personalplanung.py --datei Personalplanung_2027.xlsx --quelle "<Bereitschaftskalender/PEP>.xlsx"

Die Fachlogik wird hier in Python nachgebaut (Rotation, Freistellung, Feiertage Hessen, Abwesenheiten,
Vertretungen, Dienstcodes, Qualifikations-Gueltigkeit, Einsatzfaehigkeit, Ersatzvorschlaege, Zeitkonten) und
mit den in der Datei gespeicherten Ergebnissen verglichen. Zusaetzlich: Rotation und Feiertage gegen die
Farben/Markierungen des Original-Bereitschaftskalenders.
"""
import argparse
import sys
from datetime import date, datetime, timedelta

from openpyxl import load_workbook

from pp_base import ABW_ARTEN, ABW_CODES, FUNKTIONEN, NMAX, ROTATION, TEAM_NAMES, U_NAMES
from pp_source import read_calendar, read_holidays

CODES = [c[0] for c in ABW_CODES]
ART2CODE = dict(ABW_ARTEN)
BIG = 2958465


def d(v):
    if isinstance(v, datetime):
        return v.date()
    return v


def easter(y):
    a, b, c = y % 19, y // 100, y % 100
    dd, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - dd - g + 15) % 30
    i, k = c // 4, c % 4
    l_ = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l_) // 451
    month = (h + l_ - 7 * m + 114) // 31
    day = (h + l_ - 7 * m + 114) % 31 + 1
    return date(y, month, day)


def holidays_he(y):
    o = easter(y)
    return {date(y, 1, 1): "Neujahr", o - timedelta(2): "Karfreitag", o: "Ostersonntag",
            o + timedelta(1): "Ostermontag", date(y, 5, 1): "Tag der Arbeit", o + timedelta(39): "Christi Himmelfahrt",
            o + timedelta(49): "Pfingstsonntag", o + timedelta(50): "Pfingstmontag", o + timedelta(60): "Fronleichnam",
            date(y, 10, 3): "Tag der Deutschen Einheit", date(y, 12, 25): "1. Weihnachtstag",
            date(y, 12, 26): "2. Weihnachtstag"}


def rotation(day, start=date(2027, 1, 1)):
    wd = day.weekday()  # 0 = Mo
    mon = day - timedelta(wd)
    mon0 = start - timedelta(start.weekday())
    k = 2 * (mon - mon0).days // 7 + (wd >= 4) - (start.weekday() >= 4)
    night = ROTATION[k % 10][1]
    freed = ROTATION[(k - 1) % 10][1] if (wd in (4, 0, 1) and k >= 1) else ""
    return k, night, freed


def edate(dt, months):
    y, m = divmod(dt.month - 1 + months, 12)
    y += dt.year
    m += 1
    import calendar
    return date(y, m, min(dt.day, calendar.monthrange(y, m)[1]))


class Check:
    def __init__(self):
        self.ok = 0
        self.fail = []

    def eq(self, what, got, exp):
        if got == exp or (isinstance(got, float) and isinstance(exp, (int, float)) and abs(got - exp) < 1e-6):
            self.ok += 1
        else:
            self.fail.append(f"{what}: Datei={got!r} erwartet={exp!r}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--datei", required=True)
    ap.add_argument("--quelle")
    a = ap.parse_args()
    wb = load_workbook(a.datei, data_only=True)
    C = Check()
    year = wb["Einstellungen"]["C5"].value
    stich = d(wb["Einstellungen"]["C9"].value)
    hol = holidays_he(year)
    days = [date(year, 1, 1) + timedelta(i) for i in range((date(year, 12, 31) - date(year, 1, 1)).days + 1)]

    # ---------------------------------------------------------------- Schichtplan
    sp = wb["Schichtplan 2027"]
    for i, day in enumerate(days):
        r = 5 + i
        k, night, freed = rotation(day)
        C.eq(f"Datum {r}", d(sp[f"A{r}"].value), day)
        C.eq(f"Nachtteam {day}", sp[f"E{r}"].value, night)
        C.eq(f"Frei {day}", sp[f"F{r}"].value or "", freed)
        C.eq(f"Feiertag {day}", sp[f"D{r}"].value or "", hol.get(day, ""))
        arb = 1 if day.weekday() < 5 and day not in hol else 0
        C.eq(f"Arbeitstag {day}", sp[f"K{r}"].value, arb)
        C.eq(f"KW {day}", sp[f"C{r}"].value, day.isocalendar()[1])
        if arb:
            exp_td = ", ".join(t for t in TEAM_NAMES if t not in (night, freed))
            C.eq(f"Tagdienstteams {day}", sp[f"G{r}"].value, exp_td)
    if a.quelle:
        cal = read_calendar(a.quelle, year)
        for day, team in cal.items():
            C.eq(f"Original-Kalender {day}", rotation(day)[1], team)
        src_h = read_holidays(a.quelle, year)
        for day in src_h:
            C.eq(f"Original-Feiertag {day}", day in hol, True)
        print(f"Original-Kalender: {len(cal)} Tage Farbe verglichen, {len(src_h)} rote Feiertage")

    # ---------------------------------------------------------------- Stammdaten aus der Datei
    st = wb["Mitarbeiterstamm"]
    staff = []
    for r in range(6, 66):
        n = st[f"B{r}"].value
        if n:
            staff.append({"name": n, "team": st[f"C{r}"].value, "aktiv": st[f"G{r}"].value,
                          "eintritt": d(st[f"D{r}"].value), "row": r})
    fm = wb["Team- und Funktionsmatrix"]
    fmd = {}
    for r in range(6, 66):
        n = fm[f"A{r}"].value
        if n:
            fmd[n] = {"tf": fm[f"C{r}"].value,
                      "fk": {f: fm.cell(r, 4 + j).value for j, f in enumerate(FUNKTIONEN)}}
    tf_order = FUNKTIONEN + ["Reserve"]
    for s in staff:
        s["tf"] = fmd.get(s["name"], {}).get("tf") or ""
        s["fk"] = fmd.get(s["name"], {}).get("fk") or {f: ("JA" if s["tf"] == f else "NEIN") for f in FUNKTIONEN}
        tidx = TEAM_NAMES.index(s["team"]) + 1 if s["team"] in TEAM_NAMES else 8
        if s["aktiv"] == "NEIN":
            tidx = 9
        fidx = tf_order.index(s["tf"]) + 1 if s["tf"] in tf_order else 6
        s["key"] = tidx * 100000 + fidx * 1000 + s["row"]
    staff.sort(key=lambda s: s["key"])
    pep = wb["PEP 2027"]
    for i, s in enumerate(staff):
        C.eq(f"Rasterzeile {i + 1}", pep[f"B{16 + i}"].value, s["name"])

    # Einstellungen
    es = wb["Einstellungen"]
    req = {f: {U_NAMES[j] for j in range(13) if es.cell(23 + i, 13 + j).value == "JA"} for i, f in enumerate(FUNKTIONEN)}
    months = {es[f"L{6 + j}"].value: es[f"M{6 + j}"].value for j in range(13)}
    strict = es["C41"].value == "JA"
    soll = {"Meister": es["C30"].value, "AV": es["C31"].value, "TVT-Fahrer": es["C32"].value,
            "Handwerker": es["C33"].value}
    std_tag = (es["C21"].value)
    std_nacht = es["C26"].value
    soll_std = es["C27"].value

    # Abwesenheiten
    aw = wb["Urlaub und Abwesenheiten"]
    absences = []
    for r in range(6, 506):
        n = aw[f"A{r}"].value
        if not n:
            continue
        art, von, bis, thema = aw[f"B{r}"].value, d(aw[f"C{r}"].value), d(aw[f"D{r}"].value), aw[f"E{r}"].value
        bis = bis or von
        code = ("FDL" if thema == "FDL-Einweisung" else "KAB" if thema == "Kabeleinweisung" else "UR") \
            if art == "Unterricht" else ART2CODE[art]
        absences.append((n, art, von, bis, thema, code))
        C.eq(f"Abw-Code Zeile {r}", aw[f"K{r}"].value, code)
    vt = wb["Vertretungen"]
    verts = []
    for r in range(19, 219):
        if vt[f"E{r}"].value and vt[f"A{r}"].value:
            verts.append((d(vt[f"A{r}"].value), d(vt[f"B{r}"].value) or d(vt[f"A{r}"].value), vt[f"C{r}"].value,
                          vt[f"D{r}"].value, vt[f"E{r}"].value))

    # Qualifikationen
    qs = wb["Qualifikationen"]
    qrec = {}
    for r in range(6, 506):
        n, q = qs[f"A{r}"].value, qs[f"B{r}"].value
        if not n or not q:
            continue
        last = d(qs[f"C{r}"].value)
        done = [x[3] for x in absences if x[0] == n and x[1] == "Unterricht" and x[4] == q and x[3] <= stich]
        cand = [x for x in [last] + done if x]
        manual = d(qs[f"I{r}"].value)
        exp = manual if manual else (edate(max(cand), months[q]) if cand else None)
        C.eq(f"Ablauf {n}/{q}", d(qs[f"D{r}"].value), exp)
        nxt = [x[2] for x in absences if x[0] == n and x[1] == "Unterricht" and x[4] == q and x[2] > stich]
        plan = (edate(min(nxt), months[q]) if (nxt and exp and min(nxt) <= exp) else exp)
        C.eq(f"Planung {n}/{q}", d(qs[f"O{r}"].value), plan)
        if exp:
            qrec[(n, q)] = max(qrec.get((n, q), date.min), plan)
            rest = (exp - stich).days
            nr = 4 if rest < 0 else 3 if rest < 90 else 2 if rest <= 180 else 1
        else:
            qrec.setdefault((n, q), None)
            nr = 5
        C.eq(f"Status {n}/{q}", qs[f"L{r}"].value, nr)

    def qual_until(s, f):
        if s["fk"].get(f) != "JA":
            return date.min
        vals = []
        for q in req[f]:
            e = qrec.get((s["name"], q))
            vals.append(e if e else (date.min if strict else date.max))
        return min(vals) if vals else date.max

    for s in staff:
        s["qb"] = {f: qual_until(s, f) for f in FUNKTIONEN}

    # ---------------------------------------------------------------- Dienstraster
    prio = {c: i + 1 for i, c in enumerate(CODES)}
    grid = {}
    for i, s in enumerate(staff):
        row = 16 + i
        for j, day in enumerate(days):
            k, night, freed = rotation(day)
            if s["aktiv"] == "NEIN" or (s["eintritt"] and day < s["eintritt"]):
                code, avail = "–", False
            else:
                hits = [x[5] for x in absences if x[0] == s["name"] and x[2] <= day <= x[3]]
                if hits:
                    code, avail = max(hits, key=lambda c: prio[c]), False
                else:
                    avail = True
                    if s["team"] == night:
                        code = "N"
                    elif any(v[4] == s["name"] and v[0] <= day <= v[1] for v in verts):
                        code = "V"
                    elif s["team"] == freed:
                        code = "F"
                    elif day.weekday() < 5 and day not in hol:
                        code = "T"
                    else:
                        code = ""
            grid[(s["name"], day)] = (code, avail)
            got = pep.cell(row, 6 + j).value or ""
            C.eq(f"PEP {s['name']} {day}", got, code)

    # ---------------------------------------------------------------- Einsatzfaehigkeit + Ersatz
    ef = wb["Einsatzfähigkeitsprüfung"]
    nights = {s["name"]: sum(1 for day in days if grid[(s["name"], day)][0] in ("N", "V")) for s in staff}
    idx = {s["name"]: i + 1 for i, s in enumerate(staff)}
    names_by_idx = {i + 1: s["name"] for i, s in enumerate(staff)}
    ex_checked = 0
    for j, day in enumerate(days):
        r = 6 + j
        k, night, freed = rotation(day)
        ist_by_team = {}
        for t_i, t in enumerate(TEAM_NAMES):
            ist = {}
            for f_i, f in enumerate(FUNKTIONEN):
                n = sum(1 for s in staff if s["team"] == t and s["tf"] == f and grid[(s["name"], day)][1]
                        and s["qb"][f] >= day)
                n += sum(1 for v in verts if v[2] == t and v[3] == f and v[0] <= day <= v[1])
                ist[f] = n
                C.eq(f"EF {t}/{f} {day}", ef.cell(r, 15 + 6 * t_i + f_i).value, n)
            ok = all(ist[f] >= soll[f] for f in FUNKTIONEN)
            C.eq(f"EF Status {t} {day}", ef.cell(r, 15 + 6 * t_i + 4).value,
                 "✅ Einsatzfähig" if ok else "❌ Nicht einsatzfähig")
            ist_by_team[t] = ist
        ist = ist_by_team[night]
        # Ersatz
        slots = [("Meister", ist["Meister"] < soll["Meister"]), ("AV", ist["AV"] < soll["AV"]),
                 ("TVT-Fahrer", ist["TVT-Fahrer"] < soll["TVT-Fahrer"]),
                 ("Handwerker", ist["Handwerker"] < soll["Handwerker"]),
                 ("Handwerker", ist["Handwerker"] < soll["Handwerker"] - 1),
                 ("Handwerker", ist["Handwerker"] < soll["Handwerker"] - 2)]
        used = []
        surplus = {f: ist[f] - soll[f] for f in FUNKTIONEN}
        surplus["Reserve"] = 1
        for s_i, (f, need) in enumerate(slots):
            cell = ef.cell(r, 8 + s_i).value or ""
            if not need:
                C.eq(f"Ersatz leer {day} {s_i}", cell, "")
                continue
            cands = []
            for s in staff:
                if not grid[(s["name"], day)][1] or s["qb"][f] < day or idx[s["name"]] in used:
                    continue
                if s["team"] == night:
                    g = 1 if (s["tf"] != f and surplus.get(s["tf"], 0) > 0) else 0
                elif s["team"] != freed:
                    g = 2 if s["tf"] == f else 3
                else:
                    g = 0
                if g:
                    cands.append((g, idx[s["name"]]))
            if not cands:
                C.eq(f"Ersatz {day} {f}", cell, "⚠ kein Ersatz")
                used.append(0)
            else:
                gmin = min(c[0] for c in cands)
                pool = sorted(c[1] for c in cands if c[0] == gmin)
                pick = pool[k % len(pool)]
                lab = ["P1 Team", "P2 anderes Team", "P3 Mehrfachfunktion"][gmin - 1]
                C.eq(f"Ersatz {day} {f}", cell, f"{names_by_idx[pick]} ({lab})")
                used.append(pick)
            ex_checked += 1

    # ---------------------------------------------------------------- Zeitkonten
    ue = wb["Überstunden und JAZ"]
    for i, s in enumerate(staff):
        r = 6 + i
        tot = 0
        for m in range(1, 13):
            md = [day for day in days if day.month == m]
            arb = [day for day in md if day.weekday() < 5 and day not in hol and grid[(s["name"], day)][0] != "–"]
            sol = len(arb) * soll_std
            t = sum(1 for day in md if grid[(s["name"], day)][0] == "T")
            nn = sum(1 for day in md if grid[(s["name"], day)][0] in ("N", "V"))
            absent = sum(1 for day in arb if grid[(s["name"], day)][0] in CODES)
            ist_h = t * std_tag + nn * std_nacht + absent * soll_std
            C.eq(f"Monatswert {s['name']} {m}", round(ue.cell(r, 17 + m).value, 6), round(ist_h - sol, 6))
            tot += ist_h - sol
        C.eq(f"Jahreswert {s['name']}", round(ue[f"H{r}"].value, 6), round(tot, 6))

    # ---------------------------------------------------------------- Statistiken
    stt = wb["Statistiken"]
    for i, s in enumerate(staff):
        r = 6 + i
        codes = [grid[(s["name"], day)][0] for day in days]
        arb = [day.weekday() < 5 and day not in hol for day in days]
        nv = [c in ("N", "V") for c in codes]
        C.eq(f"Stat N {s['name']}", stt[f"D{r}"].value, sum(nv))
        C.eq(f"Stat WE {s['name']}", stt[f"E{r}"].value, sum(1 for j, x in enumerate(nv) if x and days[j].weekday() >= 4))
        C.eq(f"Stat FT {s['name']}", stt[f"F{r}"].value, sum(1 for j, x in enumerate(nv) if x and days[j] in hol))
        blocks = sum(1 for j, x in enumerate(nv) if x and (j == 0 or days[j].weekday() in (0, 4) or not nv[j - 1]))
        C.eq(f"Stat Blöcke {s['name']}", stt[f"G{r}"].value, blocks)
        C.eq(f"Stat U {s['name']}", stt[f"I{r}"].value, sum(1 for j, c in enumerate(codes) if c == "U" and arb[j]))
        C.eq(f"Stat K {s['name']}", stt[f"J{r}"].value, sum(1 for j, c in enumerate(codes) if c == "K" and arb[j]))
    # ---------------------------------------------------------------- Personalbedarf Tagdienst
    pb = wb["Personalbedarfsplanung"]
    mintag, puffer = es["C35"].value, es["C36"].value
    for j, day in enumerate(days):
        r = 6 + j
        k, night, freed = rotation(day)
        if day.weekday() < 5 and day not in hol:
            plan = sum(1 for s in staff if grid[(s["name"], day)][0] != "–" and s["team"] not in (night, freed))
            ist = sum(1 for s in staff if grid[(s["name"], day)][0] == "T")
            C.eq(f"PB Plan {day}", pb[f"G{r}"].value, plan)
            C.eq(f"PB Ist {day}", pb[f"H{r}"].value, ist)
            C.eq(f"PB Ampel {day}", pb[f"N{r}"].value,
                 "Rot" if ist < mintag else ("Gelb" if ist < mintag + puffer else "Grün"))

    print(f"Geprüft: {C.ok} Werte in Ordnung, {len(C.fail)} Abweichungen, {ex_checked} Ersatzvorschläge nachgerechnet")
    for f in C.fail[:40]:
        print("  ✗", f)
    sys.exit(1 if C.fail else 0)


if __name__ == "__main__":
    main()
