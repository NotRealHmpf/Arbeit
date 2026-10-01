"""Unabhaengige Nachrechnung in Python (ohne Excel-Formeln) – zum Pruefen des Dashboards.

Liest die Bezirks-Blaetter und den Zielzustand direkt und wendet die Zaehlregeln an:
Status je Mitarbeiter, Stufe heute / je Jahresende / nach Plan, Teamleiter, Soll-Ist, Hinweise.
"""
import datetime as dt
import re

from openpyxl import load_workbook

SRC_FIRST_ROW, SRC_LAST_ROW = 2, 61
STAGES = ["Azubi", "Arbeiter LST", "Wmech", "SigMech", "SigMech RBEG"]
NO_STAGE = "keine Stufe"
CMP = ["Arbeiter LST", "Wmech", "SigMech", "SigMech RBEG", "Teamleiter"]
GROUP_OTHER = "Sonstige / unklar"
COLS = "LMNOPSUV"
COL_NO = {c: i + 1 for i, c in enumerate("LMNOPQRS")}
N_YEARS = 7

# Text in Spalte G/J -> (Spalte L..S, Verwendungspruefung, Gruppe) – gleiche Standardtabelle wie im Blatt Einstellungen
MAPPING = {
    "azubi": ("L", "", "Azubi"), "arb lst": ("M", "", "Arbeiter LST"), "arbeiter lst": ("M", "", "Arbeiter LST"),
    "weichmech": ("N", "U", "Wmech"), "weichenmechaniker": ("N", "U", "Wmech"), "wmech": ("N", "U", "Wmech"),
    "sigmech": ("O", "V", "SigMech"), "signalmechaniker": ("O", "V", "SigMech"),
    "sigmech rbeg": ("P", "", "SigMech RBEG"), "signalmechaniker rbeg": ("P", "", "SigMech RBEG"),
    "kennziffer 4": ("Q", "", GROUP_OTHER), "ihk-meister": ("R", "", GROUP_OTHER),
    "teamleiter": ("S", "", "Teamleiter"), "tl": ("S", "", "Teamleiter"),
    "senior expert lst": ("", "", GROUP_OTHER), "umschüler ebet": ("", "", GROUP_OTHER), "umschüler": ("", "", GROUP_OTHER),
    "quereinsteiger": ("", "", GROUP_OTHER),
}


def trim(v):
    """Wie TRIM(x&"") in Excel: Leerzeichen am Rand weg, mehrere innen zu einem."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "WAHR" if v else "FALSCH"
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r" +", " ", str(v)).strip(" ")


def text(v):
    """Wie TRIM(CLEAN(x&"")) in Excel (CLEAN entfernt Steuerzeichen, z. B. Zeilenumbrueche)."""
    return trim(re.sub(r"[\x00-\x1f]", "", trim(v)))


def year_of(v):
    """Jahr aus einem Eintrag (27 -> 2027, 2027, Datum) oder None."""
    if v is None:
        return None
    if isinstance(v, (dt.datetime, dt.date)):
        return v.year
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        x = float(v)
    else:
        s = text(v)
        if not s:
            return None
        try:
            x = float(s.replace(",", "."))
        except ValueError:
            return None
    if x > 2100:            # Datum als Seriennummer
        return (dt.date(1899, 12, 30) + dt.timedelta(days=int(x))).year
    if x >= 1900:
        return round(x)
    if 1 <= x <= 99:
        return 2000 + round(x)
    return None


def reached_from(v):
    """0 = "x", Jahr = geplant, 9999 = nichts / ungueltig."""
    if trim(v).lower() == "x":
        return 0
    y = year_of(v)
    return 9999 if y is None else y


class Person:
    def __init__(self, sheet, row, cells):
        self.sheet, self.row = sheet, row
        self.raw = cells
        self.name = text(cells["B"])
        self.vor = text(cells["C"])
        self.ziel = text(cells["G"])
        self.ist = text(cells["J"])
        self.e = {c: trim(cells[c]) for c in COLS}
        self.a = {c: reached_from(cells[c]) for c in COLS}
        self.n_eff = min(self.a["N"], 0 if self.e["U"].lower() == "x" else 9999)
        self.o_eff = min(self.a["O"], 0 if self.e["V"].lower() == "x" else 9999)

    def level(self, T):
        """Rechteste erreichte Stufe in L-P zum Zeitpunkt T (0 = heute, Jahr, 9998 = nach Plan)."""
        for nr, a in ((5, self.a["P"]), (4, self.o_eff), (3, self.n_eff), (2, self.a["M"]), (1, self.a["L"])):
            if a <= T:
                return nr
        return 0

    def tl(self, T):
        return 1 if self.a["S"] <= T else 0


def stage_name(nr):
    return ([NO_STAGE] + STAGES)[nr]


def read_people(src_wb, sheet):
    ws = src_wb[sheet]
    out = []
    for r in range(SRC_FIRST_ROW, SRC_LAST_ROW + 1):
        cells = {c: ws[f"{c}{r}"].value for c in "BCGJ" + COLS}
        name = text(cells["B"])
        if not name or name.lower() == "name":
            continue
        out.append(Person(sheet, r, cells))
    return out


def status(p, bj):
    """(Status-Nr, Jahr in Ziel-Spalte, Gruppe)."""
    m = MAPPING.get(p.ziel.lower()) if p.ziel else None
    grp = m[2] if m and m[2] in CMP else GROUP_OTHER
    if not m or not m[0]:
        return 9, None, grp
    col, vp, _ = m
    zent = trim(p.raw[col])
    vpent = trim(p.raw[vp]) if vp else ""
    zyear = year_of(p.raw[col]) if zent.lower() != "x" else None
    if zent.lower() == "x" or vpent.lower() == "x":
        return 1, zyear, grp
    if zyear is None:
        return 8, None, grp
    if zyear < bj:
        return 7, zyear, grp
    return min(6, 2 + zyear - bj), zyear, grp


def hints(p, bj):
    st, _, _ = status(p, bj)
    m = MAPPING.get(p.ziel.lower()) if p.ziel else None
    zcol = COL_NO.get(m[0], 0) if m and m[0] else 0
    out = []
    if st == 9:
        out.append("Ziel-Qualifikation (G) leer" if not p.ziel else f"Ziel „{p.ziel}“ hat keine Spalte in L–S")
    elif st == 8:
        right = any(zcol < i and p.e[c] for i, c in ((2, "M"), (3, "N"), (4, "O"), (5, "P")))
        out.append(f"Ziel „{p.ziel}“: in der Ziel-Spalte weder „x“ noch Jahr" + (" (Planung nur in Spalten weiter rechts)" if right else ""))
    past = []
    for col, a in (("L", p.a["L"]), ("M", p.a["M"]), ("N", p.n_eff), ("O", p.o_eff), ("P", p.a["P"]), ("S", p.a["S"]),
                   ("U", p.a["U"]), ("V", p.a["V"])):
        if 1900 <= a < bj:
            past.append(f"{col} {a}")
    if past:
        out.append("Jahr vorbei ohne „x“: " + ", ".join(past))
    parts = []
    bad = [f"{c} „{p.e[c]}“" for c in COLS if p.e[c] and p.a[c] == 9999]
    if bad:
        parts.append("kein gültiger Eintrag: " + ", ".join(bad))
    if p.level(9998) == 0:
        parts.append("in L–P nichts eingetragen (keine Stufe)")
    if p.e["V"] and not p.e["O"]:
        parts.append("Eintrag in V (örtl. VP SigMech), aber Spalte O leer")
    if p.e["U"] and not p.e["N"]:
        parts.append("Eintrag in U (örtl. VP Wmech), aber Spalte N leer")
    if p.e["P"] and not p.e["O"] and p.e["V"].lower() != "x":
        parts.append("RBEG (P) eingetragen, aber SigMech (O) leer")
    if parts:
        out.append("; ".join(parts))
    icol = 0
    if p.ist:
        if p.ist.lower().startswith("azubi"):
            icol = 1
        else:
            mi = MAPPING.get(p.ist.lower())
            icol = COL_NO.get(mi[0], 0) if mi and mi[0] else 0
    lv0 = p.level(0)
    if 1 <= icol <= 5 and icol != lv0:
        out.append(f"Ist-Qualifikation „{p.ist}“, laut L–P heute: {stage_name(lv0)}")
    elif icol == 8 and p.tl(0) == 0:
        out.append(f"Ist-Qualifikation „{p.ist}“, aber kein „x“ in S")
    return "; ".join(out)


def thresholds(bj):
    return [0] + [bj + k for k in range(N_YEARS)] + [9998]


def read_soll(zz_ws, zz_name):
    """Soll je Vergleichsstufe aus der Spalte "Zielzustand" (3 Spalten rechts vom Bezirksnamen in Zeile 1)."""
    col = None
    for c in zz_ws[1]:
        if text(c.value) == zz_name:
            col = c.column + 3
    soll = {}
    for cat in CMP:
        v = 0
        for r in range(1, 61):
            if text(zz_ws.cell(r, 1).value).lower() == cat.lower():
                x = zz_ws.cell(r, col).value if col else 0
                v = x if isinstance(x, (int, float)) else 0
                break
        soll[cat] = v
    return soll


def district_result(people, soll, bj):
    T = thresholds(bj)
    res = {"soll": soll, "counts": {}, "persons": people}
    for key in ["Azubi"] + CMP + [NO_STAGE, "Köpfe"]:
        row = []
        for t in T:
            if key == "Teamleiter":
                row.append(sum(p.tl(t) for p in people))
            elif key == "Köpfe":
                row.append(len(people))
            else:
                nr = 0 if key == NO_STAGE else STAGES.index(key) + 1
                row.append(sum(1 for p in people if p.level(t) == nr))
        res["counts"][key] = row
    res["fehlt"] = {c: [max(0, soll[c] - x) for x in res["counts"][c]] for c in CMP}
    res["mehr"] = {c: [max(0, x - soll[c]) for x in res["counts"][c]] for c in CMP}
    res["fehlend"] = [sum(res["fehlt"][c][k] for c in CMP) for k in range(len(T))]
    res["mehr_sum"] = [sum(res["mehr"][c][k] for c in CMP) for k in range(len(T))]
    res["status"] = {}
    for p in people:
        st, zy, grp = status(p, bj)
        res["status"][(grp, st)] = res["status"].get((grp, st), 0) + 1
    return res


def compute(src_path, zz_path, districts, bj, davon=None):
    """districts: Liste (Blattname, Name im Zielzustand). Ergebnis je Bezirk + Pruefliste."""
    src = load_workbook(src_path, data_only=True)
    zz = load_workbook(zz_path, data_only=True).worksheets[0]
    davon = davon or bj + 1
    out = {}
    for sheet, zz_name in districts:
        people = read_people(src, sheet)
        out[sheet] = district_result(people, read_soll(zz, zz_name), bj)
        out[sheet]["davon"] = sum(1 for p in people if 2 <= status(p, bj)[0] <= 6 and status(p, bj)[1] == davon)
    pruef = [(p.sheet, p.row, p.name, p.vor, p.ziel, hints(p, bj))
             for sheet, _ in districts for p in out[sheet]["persons"] if hints(p, bj)]
    return out, pruef
