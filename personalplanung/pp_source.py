"""Liest die Teamaufstellung und den farbigen Bereitschaftskalender aus der
bestehenden Datei (PEP-Datei, Blatt "Bereitschaft 2027", oder Bereitschaftskalender,
Blatt "Kalender 2027 Hessen").

Aufbau des Quellblatts (wie im Original):
  Spalte U, Zeilen 3-8: Meister, AV, TVT Fahrer, Handwerker 1-3
  Spalten Y, AC, AG, AK, AO: Teams Blau, Rot, Gelb, Gruen, Orange (Namen, Teamfarbe)
  Zeilen 10-40: Kalender, je Monat 4 Spalten (Tag, Wochentag, Feiertag, KW);
  die Fuellfarbe der Tageszahl zeigt das Nachtteam.
"""
import json
from datetime import date

from openpyxl import load_workbook

TEAM_COLS = {"Y": "Blau", "AC": "Rot", "AG": "Gelb", "AK": "Grün", "AO": "Orange"}
ROLE_ROWS = {3: "Meister", 4: "AV", 5: "TVT-Fahrer", 6: "Handwerker", 7: "Handwerker", 8: "Handwerker"}
# Fuellfarben im Original -> Team (theme 4 = Akzent 1 = 4F81BD, theme 9 = Akzent 6 = F79646)
COLOR_TEAM = {"FF4F81BD": "Blau", "theme4": "Blau", "FFFF4B4B": "Rot", "FFFFFF00": "Gelb",
              "FF92D050": "Grün", "FFF79646": "Orange", "theme9": "Orange"}
SHEETS = ["Bereitschaft 2027", "Kalender 2027 Hessen"]


def _color_key(cell):
    if not cell.fill or not cell.fill.fill_type:
        return None
    c = cell.fill.fgColor
    if c.type == "rgb":
        return c.rgb
    if c.type == "theme":
        return f"theme{c.theme}"
    return None


def _sheet(wb):
    for s in wb.sheetnames:
        if s.strip() in SHEETS:
            return wb[s]
    raise ValueError(f"Kein Blatt {SHEETS} in der Quelldatei gefunden.")


def read_teams(path):
    """Liste der Mitarbeiter: dict(name, team, funktion) in Kalender-Reihenfolge."""
    ws = _sheet(load_workbook(path))
    people = []
    for col, team in TEAM_COLS.items():
        for row, fkt in ROLE_ROWS.items():
            v = ws[f"{col}{row}"].value
            if v is None or not str(v).strip():
                continue
            name = str(v).strip()
            # Team aus der Farbe der Zelle (Original: Name steht in Teamfarbe)
            t = COLOR_TEAM.get(_color_key(ws[f"{col}{row}"]), team)
            people.append({"name": name, "team": t, "funktion": fkt,
                           "quelle": f"{col}{row}"})
    return people


def read_calendar(path, year=2027):
    """{date: team} aus den Fuellfarben der Tageszahlen (zur Pruefung der Rotation)."""
    ws = _sheet(load_workbook(path))
    out = {}
    for m in range(12):
        col = 1 + 4 * m
        for r in range(10, 41):
            c = ws.cell(r, col)
            if isinstance(c.value, (int, float)):
                out[date(year, m + 1, int(c.value))] = COLOR_TEAM.get(_color_key(c))
    return out


def read_holidays(path, year=2027):
    """{date: name} der rot markierten Feiertage (Schrift CC0000) im Quellkalender."""
    ws = _sheet(load_workbook(path))
    out = {}
    for m in range(12):
        col = 1 + 4 * m
        for r in range(10, 41):
            c = ws.cell(r, col)
            if isinstance(c.value, (int, float)) and c.font and c.font.color is not None \
                    and c.font.color.type == "rgb" and c.font.color.rgb == "FFCC0000":
                out[date(year, m + 1, int(c.value))] = str(ws.cell(r, col + 2).value or "").replace("\n", "")
    return out


def apply_overrides(people, path):
    """Optionale Korrekturen (JSON, nicht eingecheckt):
    {"team": {"Name": "Blau"}, "funktion": {"Name": "Handwerker"},
     "zusatz": {"Name": ["AV", "TVT-Fahrer", "Handwerker", "Ersthelfer"]}}
    "zusatz" = weitere Funktionen/Qualifikationen (JA in der Funktionsmatrix)."""
    for p in people:
        p.setdefault("zusatz", [])
    if not path:
        return people
    with open(path, encoding="utf-8") as f:
        ov = json.load(f)
    by_name = {p["name"]: p for p in people}
    for n, t in ov.get("team", {}).items():
        if n in by_name:
            by_name[n]["team_quelle"] = by_name[n]["team"]
            by_name[n]["team"] = t
    for n, fk in ov.get("funktion", {}).items():
        if n in by_name:
            by_name[n]["funktion"] = fk
    for n, extra in ov.get("zusatz", {}).items():
        if n in by_name:
            by_name[n]["zusatz"] = list(extra)
    return people


def read_org_text(path):
    """Organisationszeile (Y2) des Bereitschaftskalenders."""
    ws = _sheet(load_workbook(path))
    return " ".join(str(ws["Y2"].value or "").split())


def override_notes(people):
    """Hinweise fuer die Startseite: Abweichungen zur Quelldatei."""
    out = []
    for p in people:
        if p.get("team_quelle") and p["team_quelle"] != p["team"]:
            out.append(f"Teamzuordnung laut Vorgabe: {p['name']} Team {p['team']} (im Bereitschaftskalender unter "
                       f"{p['team_quelle']} eingetragen) – bei Bedarf im Mitarbeiterstamm ändern.")
        extra = [x for x in p.get("zusatz", []) if x != p["funktion"]]
        if extra:
            out.append(f"Mehrfachfunktionen laut Vorgabe: {p['name']} – {', '.join(extra)}.")
    return out
