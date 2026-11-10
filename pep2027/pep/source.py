"""Liest die Teamzuordnung aus dem bestehenden Bereitschaftskalender und prueft die Rotation.

Im Blatt "Bereitschaft 2027" stehen die Funktionen in Spalte U (Zeilen 3-8) und je Team ein
4 Spalten breiter Block ab Y (Blau), AC (Rot), AG (Gelb), AK (Gruen), AO (Orange). Die Tagesfelder
(Spalten A, E, I, ... ab Zeile 10) sind in der Farbe des Nachtteams eingefaerbt.
"""
import csv
import datetime as dt

from openpyxl import load_workbook

from .model import ROTATION, TEAMS, TEAM_COLOR

TEAM_COLS = {"Blau": 25, "Rot": 29, "Gelb": 33, "Grün": 37, "Orange": 41}     # Y, AC, AG, AK, AO
ROLE_ROWS = {3: "Meister", 4: "AV", 5: "TVT-Fahrer", 6: "Handwerker", 7: "Handwerker", 8: "Handwerker"}
# Farbe der Tagesfelder -> Team (Bereitschaftskalender)
DAY_COLORS = {"FF4B4B": "Rot", "FFFF00": "Gelb", "92D050": "Grün", "F79646": "Orange", "4F81BD": "Blau"}


def find_sheet(wb, *needles):
    for ws in wb.worksheets:
        if all(n.lower() in ws.title.lower() for n in needles):
            return ws
    raise KeyError(f"Kein Blatt mit {needles} gefunden: {wb.sheetnames}")


def read_teams(path):
    """-> Liste von dicts (Name, Team, Funktion, Slot) in der Reihenfolge des Kalenders."""
    wb = load_workbook(path)
    ws = find_sheet(wb, "bereitschaft")
    people = []
    for team in TEAMS:
        c = TEAM_COLS[team]
        for r, role in ROLE_ROWS.items():
            label = str(ws.cell(r, 21).value or "").strip()          # Spalte U
            name = ws.cell(r, c).value
            if name is None or str(name).strip() == "":
                continue
            people.append({"Name": str(name).strip(), "Team": team, "Funktion": role, "Slot": label})
    return people


def read_rotation_colors(path, year=2027):
    """-> {datum: team} aus den eingefaerbten Tagesfeldern des Bereitschaftskalenders."""
    wb = load_workbook(path)
    ws = find_sheet(wb, "bereitschaft")
    out = {}
    for m in range(12):
        c = 1 + 4 * m
        for d in range(1, 32):
            cell = ws.cell(9 + d, c)
            if not isinstance(cell.value, (int, float)):
                continue
            try:
                date = dt.date(year, m + 1, int(cell.value))
            except ValueError:
                continue
            fg = cell.fill.fgColor
            rgb = fg.rgb[2:] if fg is not None and fg.type == "rgb" else None
            if rgb in DAY_COLORS:
                out[date] = DAY_COLORS[rgb]
    return out


def apply_names(people, csv_path):
    """Optionale Zuordnung Platzhalter -> Name (CSV mit Semikolon: Platzhalter;Name[;Personalnummer])."""
    if not csv_path:
        return people
    mapping = {}
    with open(csv_path, encoding="utf-8-sig") as f:
        for row in csv.reader(f, delimiter=";"):
            if len(row) >= 2 and row[0].strip() and not row[0].startswith("#"):
                mapping[row[0].strip()] = (row[1].strip(), row[2].strip() if len(row) > 2 else "")
    for p in people:
        if p["Name"] in mapping:
            p["Platzhalter"] = p["Name"]
            p["Name"], p["Personalnummer"] = mapping[p["Name"]]
    return people


def rotation_team(date, start=dt.date(2027, 1, 1)):
    n = (date - start).days
    k = 2 * (n // 7) + (1 if n % 7 >= 3 else 0)
    return ROTATION[k % 10][1]


def check_rotation(path, year=2027):
    """Vergleicht die Rotationsformel mit den Farben des Bereitschaftskalenders."""
    colors = read_rotation_colors(path, year)
    diffs = [(d, t, rotation_team(d)) for d, t in sorted(colors.items()) if t != rotation_team(d)]
    return len(colors), diffs
