"""Fachliche Konstanten und Layout des PEP-Planungssystems.

Alle Werte, die der Anwender spaeter aendern koennen soll (Jahr, Rotation, Sollbesetzung,
Kuerzel, Feiertage ...), landen im Blatt "Einstellungen". Hier stehen nur die Startwerte
und die festen Positionen der Raster.
"""
from openpyxl.utils import get_column_letter as col

YEAR = 2027
ROTATION_START = (2027, 1, 1)          # Freitag, Beginn der Rotation (Vorgabe)

TEAMS = ["Blau", "Rot", "Gelb", "Grün", "Orange"]
# Farben aus dem bestehenden Bereitschaftskalender 2027 (Tagesfelder der Rotation)
TEAM_COLOR = {"Blau": "4F81BD", "Rot": "FF4B4B", "Gelb": "FFFF00", "Grün": "92D050", "Orange": "F79646"}
TEAM_LIGHT = {"Blau": "C5D9F1", "Rot": "FFC7C7", "Gelb": "FFFF99", "Grün": "D8E4BC", "Orange": "FCD5B4"}
TEAM_FONT = {"Blau": "FFFFFF", "Rot": "FFFFFF", "Gelb": "000000", "Grün": "000000", "Orange": "000000"}
TEAM_SHORT = {"Blau": "BL", "Rot": "RT", "Gelb": "GE", "Grün": "GR", "Orange": "OR"}

GROUPS = TEAMS + ["Ohne Team", "Ressourcen"]
GROUP_TITLE = {"Blau": "Team Blau", "Rot": "Team Rot", "Gelb": "Team Gelb", "Grün": "Team Grün",
               "Orange": "Team Orange", "Ohne Team": "Ohne Team / Disponent",
               "Ressourcen": "Zusätzliche Ressourcen (Zusätzliche MA / AZUBIS)"}
GROUP_COLOR = dict(TEAM_COLOR, **{"Ohne Team": "A6A6A6", "Ressourcen": "B7DEE8"})
GROUP_FONT = dict(TEAM_FONT, **{"Ohne Team": "000000", "Ressourcen": "000000"})

# Rotationsfolge laut Vorgabe, Start 01.01.2027 (Freitag). Darf nicht automatisch veraendert werden.
ROTATION = [("Fr–So", "Rot"), ("Mo–Do", "Gelb"), ("Fr–So", "Grün"), ("Mo–Do", "Orange"),
            ("Fr–So", "Blau"), ("Mo–Do", "Rot"), ("Fr–So", "Gelb"), ("Mo–Do", "Grün"),
            ("Fr–So", "Orange"), ("Mo–Do", "Blau")]

FUNKTIONEN = ["Meister", "AV", "TVT-Fahrer", "Handwerker", "Disponent", "Zusätzliche MA", "AZUBI"]
ROLLEN = ["Meister", "AV", "TVT-Fahrer", "Handwerker"]           # Sollbesetzung Bereitschaft
SOLL = {"Meister": 1, "AV": 1, "TVT-Fahrer": 1, "Handwerker": 2}
ROLLE_KURZ = {"Meister": "Meister", "AV": "AV", "TVT-Fahrer": "TVT", "Handwerker": "Handwerker"}
QUALIS = ["Meister", "AV", "TVT-Fahrer", "Handwerker", "Ersthelfer", "SeSi", "Sakra", "SiÜ", "Gefahrgut",
          "Bühnenbediener", "S&R", "WBT", "Fahrwegmechaniker", "Kettensäge"]

SOLL_TAGDIENST = 20      # Annahme, im Blatt Einstellungen anpassbar
MINDESTRESERVE = 3       # Annahme, im Blatt Einstellungen anpassbar

# Status je Mitarbeiter und Tag (Berechnung).  Code -> (Bezeichnung, Kalendertext, Eingabekuerzel,
# Kategorie, Fuellfarbe, Schriftfarbe, fett, unterstrichen, kursiv, Beschreibung)
# Die Codes 1..8 sind zugleich die Prioritaet bei ueberlappenden Abwesenheiten (hoeher gewinnt).
STATUS = [
    (1, "Besonderheiten", "!", "", "Hinweis", "FFFF00", "000000", False, False, False,
     "Freitext im Kalender = Besonderheit (Text erscheint in der Zelle)"),
    (2, "JAZ + 99", "JAZ", "JAZ, J", "Abwesenheit", "C4D79B", "000000", False, False, False, "JAZ + 99"),
    (3, "Urlaub", "U", "U", "Abwesenheit", "00B050", "000000", True, False, False, "Urlaub"),
    (4, "Urlaub genehmigt eTime", "U", "UE, UG, U_", "Abwesenheit", "00B050", "000000", True, True, False,
     "Urlaub genehmigt in eTime (unterstrichenes U)"),
    (5, "Kabeleinweisungen", "Kabel", "KABEL, KAB, KA", "Abwesenheit", "7030A0", "FFFFFF", True, False, False,
     "Kabeleinweisung"),
    (6, "FDL-Einweisungen", "FDL", "FDL, F", "Abwesenheit", "FFC000", "000000", True, False, False,
     "FDL-Einweisung"),
    (7, "Unterrichte", "S", "S, UNT, UR", "Abwesenheit", "FF0000", "FFFFFF", True, False, False,
     "Unterricht / Schulung"),
    (8, "Krankheit", "K", "K", "Abwesenheit", "FFFF00", "FF0000", True, False, False,
     "Krank (gelber Hintergrund, rotes K)"),
    (9, "Bereitschaft / Nachtdienst", "", "V", "Dienst", "BFBFBF", "000000", False, False, False,
     "automatisch aus der Rotation; Eingabe V = Vertretung"),
    (10, "Freistellung", "FS", "FS, FREI", "Freistellung", None, "7F7F7F", True, False, True,
     "Vorschlag nach Nachtdienst, manuell änderbar; Eingabe - hebt auf"),
    (11, "Besonderheit in Bereitschaft", "!", "", "Hinweis", "FFFF00", "000000", False, False, False,
     "Besonderheit an einem Bereitschaftstag (zählt als Bereitschaft)"),
    (12, "Zusätzliche MA", "", "Z, +", "Ressource", "FFFF00", "000000", False, False, False,
     "zusätzliche Ressource, ersetzt keine Sollbesetzung"),
    (13, "AZUBIS", "", "A, AZ", "Ressource", "00B0F0", "000000", False, False, False,
     "Auszubildende, ersetzen keine Sollbesetzung"),
]
STATUS_BY_CODE = {s[0]: s for s in STATUS}
CODE = {s[1]: s[0] for s in STATUS}

# Auswahlliste "Art" im Blatt Abwesenheiten (Reihenfolge laut Vorgabe, Besonderheiten zuletzt)
ARTEN_AUSWAHL = ["Urlaub", "Urlaub genehmigt eTime", "Krankheit", "JAZ + 99", "Unterrichte",
                 "FDL-Einweisungen", "Kabeleinweisungen", "Besonderheiten"]
RES_ARTEN = ["Zusätzliche MA", "AZUBIS"]

WEEKEND_FILL = "F2AA84"       # PEP: Wochenendspalten (Akzent 2, heller 40 %)
HOLIDAY_FILL = "FFD9D9"       # Bereitschaftskalender: Feiertag
HEADER_FILL = "B7DEE8"        # PEP: Kopfzeilen Monat/Tag/Wochentag
TITLE_FILL = "E97132"         # PEP: Titelfeld "PEP 2027"
SAT_FILL, SUN_FILL = "FFFFCC", "FFCC99"   # Bereitschaftskalender: Sa / So

MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September",
          "Oktober", "November", "Dezember"]
MON_KURZ = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
WT = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]

# Feiertage Hessen: (Name, Formel relativ zu Ostern/Jahr, gesetzlich in Hessen)
FEIERTAGE = [
    ("Neujahr", "DATE(Jahr,1,1)", True),
    ("Karfreitag", "Ostern-2", True),
    ("Ostersonntag", "Ostern", False),
    ("Ostermontag", "Ostern+1", True),
    ("Tag der Arbeit", "DATE(Jahr,5,1)", True),
    ("Christi Himmelfahrt", "Ostern+39", True),
    ("Pfingstsonntag", "Ostern+49", False),
    ("Pfingstmontag", "Ostern+50", True),
    ("Fronleichnam", "Ostern+60", True),
    ("Tag der Deutschen Einheit", "DATE(Jahr,10,3)", True),
    ("Heiligabend", "DATE(Jahr,12,24)", False),
    ("1. Weihnachtstag", "DATE(Jahr,12,25)", True),
    ("2. Weihnachtstag", "DATE(Jahr,12,26)", True),
    ("Silvester", "DATE(Jahr,12,31)", False),
]

# ---------------------------------------------------------------- Raster (Jahreskalender/Berechnung)
NDAYS = 366                    # 366 Tagesspalten, im Nicht-Schaltjahr bleibt die letzte leer
DAY0 = 6                       # Spalte F = 1. Tag
DAYN = DAY0 + NDAYS - 1        # Spalte NG = 366. Tag
CAP = 50                       # Zeilen fuer Teamkoepfe + Mitarbeiter
ROW0 = 14                      # erste Personenzeile
ROWN = ROW0 + CAP - 1          # letzte Personenzeile (63)
AROW0 = ROW0 + 70              # Abwesenheits-Codes in Berechnung (84..133)
AROWN = AROW0 + CAP - 1
C_DAY0, C_DAYN = col(DAY0), col(DAYN)

# Jahreskalender: versteckte Hilfsspalten rechts vom sichtbaren Raster
K_GRP = DAYN + 2               # NH: Gruppe
K_TYP = DAYN + 3               # NI: Zeilentyp (K = Teamkopf, P = Person)
K_CODE0 = DAYN + 4             # NJ: Statuscode Tag 1 (Spiegel aus Berechnung)
K_OFF = K_CODE0 - DAY0         # Spaltenversatz sichtbares Feld -> Code

# Berechnung: Personenattribute rechts vom Raster
B_ATTR0 = DAYN + 2             # NH ...
B_ATTR = ["MA-Zeile", "k", "Gruppe-Nr", "j", "Q_Meister", "Q_AV", "Q_TVT", "Q_HW", "Bereitschaftstage", "Rang"]
B_COL = {name: col(B_ATTR0 + i) for i, name in enumerate(B_ATTR)}
B_GRP_COL0 = B_ATTR0 + len(B_ATTR) + 1     # Gruppentabelle
