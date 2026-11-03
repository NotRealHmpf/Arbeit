"""Testdaten fuer die Demo-Datei (klar als DEMO gekennzeichnet). Die Mitarbeiter werden ueber
Team + Funktion angesprochen, damit hier keine Namen stehen."""
from datetime import date, timedelta

from pp_base import PFLICHT, ROTATION, U_NAMES


def _who(people, team, fkt, k=0):
    hits = [p["name"] for p in people if p["team"] == team and p["funktion"] == fkt]
    return hits[k] if len(hits) > k else None


def night_blocks(team, start=date(2027, 1, 1)):
    """Liste (von, bis) der Nachtdienst-Bloecke eines Teams laut Rotation."""
    out = []
    d, k = start, 0
    while d.year == 2027:
        length = 3 if d.weekday() == 4 else 4
        if ROTATION[k % 10][1] == team:
            out.append((d, d + timedelta(days=length - 1)))
        d += timedelta(days=length)
        k += 1
    return out


def make(people):
    D = date
    w = lambda t, f, k=0: _who(people, t, f, k)
    abw = []

    def add(name, art, von, bis=None, thema=None, bem="DEMO"):
        if name:
            abw.append((name, art, von, bis, thema, bem))

    # Sommerurlaub gestaffelt
    summer = [("Blau", "Meister", D(2027, 7, 5), D(2027, 7, 23)), ("Blau", "Handwerker", D(2027, 8, 2), D(2027, 8, 20)),
              ("Rot", "AV", D(2027, 6, 28), D(2027, 7, 16)), ("Rot", "Handwerker", D(2027, 7, 19), D(2027, 8, 6)),
              ("Gelb", "TVT-Fahrer", D(2027, 7, 12), D(2027, 7, 30)), ("Gelb", "Meister", D(2027, 8, 9), D(2027, 8, 27)),
              ("Grün", "Handwerker", D(2027, 6, 28), D(2027, 7, 9)), ("Grün", "AV", D(2027, 7, 26), D(2027, 8, 13)),
              ("Orange", "Meister", D(2027, 7, 19), D(2027, 8, 6)), ("Orange", "Handwerker", D(2027, 8, 16), D(2027, 9, 3))]
    for t, f, a, b in summer:
        add(w(t, f), "Urlaub", a, b)
    # Oster- und Herbsturlaub
    add(w("Blau", "TVT-Fahrer"), "Urlaub", D(2027, 3, 22), D(2027, 4, 2))
    add(w("Gelb", "Handwerker", 1), "Urlaub", D(2027, 10, 4), D(2027, 10, 15))
    add(w("Grün", "Meister"), "Urlaub", D(2027, 10, 11), D(2027, 10, 22))
    add(w("Orange", "AV"), "Urlaub", D(2027, 12, 20), D(2027, 12, 31))
    add(w("Rot", "Meister"), "Urlaub", D(2027, 5, 18), D(2027, 5, 28))
    # Krankheit – u. a. Meister Rot waehrend des eigenen Nachtdienstes (loest Ersatz aus)
    rot_blocks = night_blocks("Rot")
    feb = [b for b in rot_blocks if b[0].month == 2 and b[0].weekday() == 0][0]
    add(w("Rot", "Meister"), "Krankheit", feb[0], feb[1] + timedelta(days=1))
    add(w("Grün", "TVT-Fahrer"), "Krankheit", D(2027, 3, 8), D(2027, 3, 10))
    add(w("Orange", "Handwerker", 1), "Krankheit", D(2027, 11, 15), D(2027, 11, 19))
    add(w("Blau", "Handwerker", 1), "Krankheit", D(2027, 1, 25), D(2027, 1, 29))
    add(w("Grün", "Handwerker"), "Krankheit", D(2027, 7, 5), D(2027, 7, 7), bem="DEMO – krank im Urlaub")
    # JAZ-Tage
    for t, f, d in [("Blau", "AV", D(2027, 5, 7)), ("Gelb", "AV", D(2027, 6, 4)), ("Orange", "TVT-Fahrer", D(2027, 9, 24)),
                    ("Rot", "TVT-Fahrer", D(2027, 11, 26)), ("Grün", "Meister", D(2027, 12, 3))]:
        add(w(t, f), "JAZ", d)
    # Unterrichte
    add(w("Gelb", "Meister"), "Unterricht", D(2027, 3, 3), D(2027, 3, 4), "FDL-Einweisung")
    add(w("Blau", "AV"), "Unterricht", D(2027, 2, 10), None, "Kabeleinweisung")
    add(w("Orange", "Meister"), "Unterricht", D(2027, 4, 13), D(2027, 4, 14), "Sakra & SiÜ")
    add(w("Rot", "Handwerker"), "Unterricht", D(2027, 9, 14), None, "Betrieb")
    add(w("Grün", "TVT-Fahrer"), "Unterricht", D(2027, 5, 11), D(2027, 5, 12), "Simulator")
    # Sonstiges
    add(w("Gelb", "Handwerker"), "Fortbildung", D(2027, 9, 20), D(2027, 9, 24))
    add(w("Blau", "Handwerker", 2), "Sonderurlaub", D(2027, 6, 11))
    add(w("Orange", "TVT-Fahrer"), "Freistellung", D(2027, 4, 30))

    # Schulungsdaten (letzte Schulung 2026, 12 Monate gueltig) und rechtzeitig geplante Auffrischungen 2027
    quali = {}
    months = {"Ersthelfer": 24, "Gefahrgut": 36}
    busy = {}
    for n, art, von, bis, thema, bem in abw:
        busy.setdefault(n, []).append((von, bis or von))

    def free_day(n, d0):
        d = d0
        while d.weekday() not in (1, 2) or any(a <= d <= b for a, b in busy.get(n, [])):
            d += timedelta(days=1)
        return d

    k = 0
    no_renewal = w("Gelb", "Handwerker")      # Auffrischung fehlt -> Qualifikation laeuft ab
    for p in people:
        fk = {p["funktion"]} | (set(p.get("zusatz", [])) & set(PFLICHT))
        req = set()
        for f in fk:
            req |= set(PFLICHT.get(f, []))
        req |= set(p.get("zusatz", [])) & set(U_NAMES)
        for u in U_NAMES:
            if u not in req:
                continue
            k += 1
            if k == 9:
                continue                      # Nachweis fehlt
            last = D(2026, 1 + (k * 5) % 12, 1 + (k * 7) % 27)
            if u == "Ersthelfer":
                last = D(2024, 10, 1)         # abgelaufen (keine Pflicht fuer die Funktion)
            quali[(p["name"], u)] = last
            exp = last.replace(year=last.year + months.get(u, 12) // 12)
            if exp.year == 2027 and u != "Ersthelfer" and not (p["name"] == no_renewal and u == "Betrieb"):
                day = free_day(p["name"], exp - timedelta(days=24))
                abw.append((p["name"], "Unterricht", day, None, u, "DEMO – Auffrischung"))
                busy.setdefault(p["name"], []).append((day, day))

    # Bestaetigte Vertretung fuer den kranken Meister Rot
    vert = [(feb[0], feb[1], "Rot", "Meister", w("Grün", "Meister"), w("Rot", "Meister"), None, "DEMO")]
    zb = []
    for t, f, v in [("Blau", "Meister", 14.5), ("Rot", "AV", 6), ("Gelb", "Handwerker", -3.25),
                    ("Orange", "TVT-Fahrer", 21)]:
        zb.append((w(t, f), D(2027, 1, 1), "Überstunden", "Vortrag Vorjahr", v, "DEMO"))
    for t, f in [("Blau", "AV"), ("Gelb", "AV"), ("Orange", "TVT-Fahrer"), ("Rot", "TVT-Fahrer"), ("Grün", "Meister")]:
        zb.append((w(t, f), D(2027, 1, 1), "JAZ", "Vortrag Vorjahr", 40, "DEMO"))
    zb.append((w("Rot", "AV"), D(2027, 6, 30), "Überstunden", "Auszahlung", -6, "DEMO"))
    zb.append((w("Gelb", "Meister"), D(2027, 1, 1), "Urlaub", "Vortrag Vorjahr", 3, "DEMO"))
    return {"abw": abw, "quali": quali, "vert": vert, "zb": zb}
