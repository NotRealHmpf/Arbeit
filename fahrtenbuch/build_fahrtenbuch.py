"""Erzeugt ein Fahrtenbuch nach dem Musterfahrtenbuch (DIN A6-aehnliches Format).

    python build_fahrtenbuch.py [--jahr 2026] [--seiten 48] [--ausgabe ausgabe]

Ergebnis im Ausgabeordner:
  Fahrtenbuch.xlsx               Excel-Fahrtenbuch: Deckblatt mit Jahresuebersicht, je Monat ein
                                 Blatt (3 Druckseiten A4 quer), Hinweise, Abkuerzungen, Beispiel
  Fahrtenbuch_Druckvorlage.pdf   Vordruck zum handschriftlichen Fuehren: Deckblatt, Hinweise,
                                 Abkuerzungen, Beispiel und nummerierte Fahrtenbuchseiten

Ablauf: openpyxl baut beide Arbeitsmappen, LibreOffice rechnet sie durch (Pruefung gegen die
Mustereintraege, gespeicherte Werte fuer die geschuetzte Ansicht) und erzeugt das PDF.
Benoetigt: openpyxl, lxml, LibreOffice (soffice) mit Calc.
"""
import argparse
import datetime
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties

# ---------------------------------------------------------------- Konstanten
ZEILEN = 20            # Fahrten je Druckseite
SEITEN_MONAT = 3       # Druckseiten je Monatsblatt (60 Fahrten im Monat)
BLOCK = 6 + ZEILEN     # Titel, Abstand, 2 Kopfzeilen, Uebertrag, Spaltennummern, Fahrten
TITEL, KOPF1, KOPF2, UEBERTRAG, NUMMERN, DATEN = 0, 2, 3, 4, 5, 6
SKALIERUNG = 90        # Druckskalierung der Fahrtenbuchseiten (A4 quer) in Prozent

FONT = "Arial"
GRAU = "6B6B6B"
KOPF_BG = "F2F2F2"
UEBERTR_BG = "D9D9D9"  # wie im Muster
CALC_BG = "E8EEF6"     # berechnete Zellen (nur Excel)
INPUT_BG = "FFF2CC"    # Eingabefelder auf dem Deckblatt
LINK = "1F4E79"
WARN_BG, WARN_FG = "FCE4D6", "C55A11"
FEHLER_BG, FEHLER_FG = "F8CBAD", "C00000"

MONATE = [("Jan", "Januar"), ("Feb", "Februar"), ("Mär", "März"), ("Apr", "April"),
          ("Mai", "Mai"), ("Jun", "Juni"), ("Jul", "Juli"), ("Aug", "August"),
          ("Sep", "September"), ("Okt", "Oktober"), ("Nov", "November"), ("Dez", "Dezember")]

# Spalten 1-15 wie im Muster (Buchstabe -> Breite)
BREITEN = {"A": 4.6, "B": 4.6, "C": 6.6, "D": 31, "E": 14, "F": 20, "G": 6.6, "H": 8.6,
           "I": 8.6, "J": 6.6, "K": 6.6, "L": 6.6, "M": 6.6, "N": 14, "O": 14}
# Hilfsspalten rechts vom Druckbereich: Fahrtart, Pruefung, km-Stand-Kette, letzte Fahrtart
HILFE = {"P": 9.5, "Q": 48, "R": 9, "S": 9}

# Kopfzeilen: (erste Spalte, letzte Spalte, Zeilen, Text)
KOPF = [("A", "A", 2, "Lfd.\nNr."), ("B", "B", 2, "Tag\ntt"), ("C", "C", 2, "Abfahrt\nhh:mm"),
        ("D", "D", 2, "Fahrweg von - über - nach\nZiel: Ort, Straße"), ("E", "E", 2, "Zweck der Fahrt"),
        ("F", "F", 2, "besuchte Personen\nFirmen, Behörden, Objekte"), ("G", "G", 2, "Ankunft\nhh:mm"),
        ("H", "I", 1, "km-Zählerstände"), ("J", "L", 1, "gefahrene km"), ("M", "M", 2, "getankt\n[l]"),
        ("N", "N", 2, "Name\nFahrer"), ("O", "O", 2, "Unterschrift\nFahrer")]
KOPF_UNTEN = {"H": "Abfahrt", "I": "Ankunft", "J": "gesch.", "K": "WO/AO", "L": "privat"}
FORMATE = {"B": "00", "C": "hh:mm", "G": "hh:mm", "H": "0", "I": "0", "J": "0", "K": "0", "L": "0",
           "M": "0.0"}
LINKS = set("DEFNO")   # linksbuendig mit Umbruch, alle anderen zentriert

HINWEIS_FUSS = ("Beachten Sie bitte die gesetzlichen und die Herstellerfristen, "
                "prüfen Sie bitte regelmäßig den Ölstand und den Reifendruck.")

# Abkuerzungsvorschlaege (Seite 3 des Musters): (Zweck der Fahrt, besuchte Personen/Firmen, Abkuerzung)
ABKUERZUNGEN = [
    ("(Ruf-)Bereitschaft", "", "Ber"),
    ("Bauabnahme", "", "BA"),
    ("Bauüberwachung", "", "Büb"),
    ("Betriebsarztuntersuchung", "", "BArzt"),
    ("Betriebsratstätigkeit", "", "BRt"),
    ("Entstörung", "", "En"),
    ("Fahrten Wohn-/Arbeitsort (und umgekehrt)", "", "WO / AO bzw. AO / WO"),
    ("Fahrten Wohn-/Arbeitsort (und umgekehrt) im Zusammenhang mit Dienstreisen", "",
     "WO / AO DR bzw. AO / WO DR"),
    ("Fahrzeugrückgabe bei DEKRA", "", "FzR DEKRA"),
    ("Feinabstimmung", "", "FA"),
    ("Inspektion", "", "In"),
    ("Instandsetzung", "", "IS"),
    ("Kontrolle", "", "K"),
    ("Krankenbesuch", "MA … <Namen>", "KB <Name>"),
    ("Kundendienst", "", "KD"),
    ("Materialtransport", "", "MT"),
    ("Personentransport", "", "PT"),
    ("Sonderinspektionen", "", "Sin"),
    ("Überführungsfahrten (mit Anlass z. B. Werkstatt, HU, AU, etc.)", "", "ÜF"),
    ("Wartungsarbeiten", "", "Wa"),
    ("Weichen reinigen", "", "Wr"),
    ("Weichen schmieren", "", "Ws"),
    ("", "Bahnübergang / Ort", "BÜ / Ort"),
    ("", "Baustelle / Ort", "Bst / Ort"),
    ("", "Bundeseisenbahnvermögen", "BEV"),
    ("", "Eisenbahnbundesamt", "EBA"),
    ("", "Gesundheitszentrum", "GHZ"),
    ("", "Sammelpunkt / Ort", "SP / Ort"),
    ("", "Stellwerk / Ort", "Stw / Ort"),
    ("", "Strecke <A> und <B>", "S"),
    ("", "Stützpunkt / Ort", "Stp / Ort"),
    ("", "Werkstatt", "Wst"),
]
# Auswahlliste fuer Spalte 5 (im Blatt Abkuerzungen erweiterbar)
AUSWAHL_ZWECK = ["zurück", "WO/AO", "AO/WO", "WO/AO DR", "AO/WO DR", "PRIVAT", "Ber", "BA", "Büb",
                 "BArzt", "BRt", "En", "FzR DEKRA", "FA", "In", "IS", "K", "KB", "KD", "MT", "PT",
                 "Sin", "ÜF", "Wa", "Wr", "Ws", "Besprechung", "Inbetriebnahme"]
ZWECK_ERSTE, ZWECK_LETZTE = 5, 60   # Zeilen der Auswahlliste im Blatt Abkuerzungen

# Mustereintraege Seite 1 (Feb. 2013):
# Tag, Abfahrt, Fahrweg, Zweck, besucht, Ankunft, km Ankunft, getankt
BEISPIEL_UEBERTRAG = 6538
BEISPIEL = [
    (1, "07:00", "Stp Nürnberg - Pleinfeld", "BA", "BÜ, km 50,123", "08:15", 6589, None),
    (1, "15:30", "Pleinfeld - Stp Nürnberg", "zurück", "", "16:30", 6640, None),
    (3, "07:30", "Stp Nürnberg - Augsburg, Victoriastr. 3", "Besprechung", "Fr. Beispiel (I.NP-S)",
     "10:00", 6785, None),
    (3, "13:00", "Augsburg, Victoriastr. 3 - Stp Nürnberg", "zurück", "", "15:30", 6930, None),
    (4, "07:00", "Stp Nürnberg - Bamberg", "IS", "Str. 5102, km 61.850", "08:15", 6992, None),
    (4, "09:00", "Baustelle, km 61.850 - 62.850", "IS", "Bst Bamberg", "15:00", 6995, None),
    (4, "15:30", "Bamberg - Stp Nürnberg", "zurück", "", "16:45", 7057, 45.7),
    (9, "07:00", "Stp Nürnberg - Nürnberg, Fürther Str. 235", "ÜF Werkstatt", "NL Mercedes Benz",
     "07:15", 7062, None),
    (None, None, "Probefahrt Kundendienst", "", "", None, 7065, None),
    (9, "14:00", "Nürnberg, Fürther Str. 235 - Stp Nürnberg", "ÜF Werkstatt", "", "14:15", 7070, None),
    (9, "16:30", "Stp Nürnberg - Burgthann, Sandstr. 18", "AO/WO DR", "", "17:15", 7097, None),
    (10, "04:00", "Burgthann, Sandstr. 18 - Straßkirchen", "Inbetriebnahme", "ESTW", "05:30", 7240, None),
    (10, "16:00", "Straßkirchen - Burgthann, Sandstr. 18", "zurück", "", "17:30", 7383, 30.3),
    (11, "07:00", "Burgthann, Sandstr. 18 - Stp Nürnberg", "WO/AO DR", "", "07:45", 7410, None),
    (11, "16:00", "Stp Nürnberg - Deining, Schloßstr. 2", "AO/WO Ber", "", "17:00", 7472, None),
    (13, "14:00", "PRIVAT", "", "", "22:00", 7495, None),
    (15, "06:00", "Deining, Schloßstr. 2 - Stp Nürnberg", "WO/AO Ber", "", "07:30", 7557, None),
]
# Fahrer wie im Muster (Klammer = eine Unterschrift fuer mehrere Fahrten): erste, letzte Fahrt, Name
BEISPIEL_FAHRER = [(1, 1, "Müller, Mike"), (2, 2, "Müller, Mike"), (3, 4, "Mohr, Ingo"),
                   (5, 7, "Josef Kern"), (8, 8, "Max Schmidt"), (10, 10, "Max Schmidt"),
                   (11, 14, "Schulze, Tim"), (15, 15, "Kurt Meyer"), (16, 16, "Kurt Meyer"),
                   (17, 17, "Kurt Meyer")]
# Erwartete Werte laut Muster: Summen gesch., WO/AO, privat, getankt und Fahrtart je Fahrt
BEISPIEL_SUMMEN = (818, 178, 23, 76.0)
BEISPIEL_ARTEN = ["gesch."] * 10 + ["WO/AO", "gesch.", "gesch.", "WO/AO", "WO/AO", "privat", "WO/AO"]

# Deckblatt: Zeilen der Eingabefelder und der Jahresuebersicht
D_JAHR, D_KENNZ, D_KMSTART, D_LETZTE = 5, 6, 14, 16
D_FAHRER = range(20, 30)
D_UEBERSICHT = 7      # erste Monatszeile
FAHRZEUG = ["Amtliches Kennzeichen", "Hersteller / Typ", "Fahrzeug-Nr. / Inventar-Nr.",
            "Organisationseinheit / Bezirk", "Kostenstelle", "Standort / Stützpunkt",
            "Poolfahrzeugdisponent (Name, Tel.)", "Fahrtenbuch begonnen am", "km-Stand bei Beginn",
            "Fahrtenbuch abgeschlossen am"]

FEHLERWERTE = {"#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A"}

DUENN = Side(style="thin", color="000000")
DICK = Side(style="medium", color="000000")
HELL = Side(style="thin", color="BFBFBF")
MITTE = Alignment(horizontal="center", vertical="center", wrap_text=True)
LINKSB = Alignment(horizontal="left", vertical="center", wrap_text=True)
RECHTS = Alignment(horizontal="right", vertical="center")


# ---------------------------------------------------------------- Hilfsfunktionen
def font(size=9, bold=False, italic=False, color="000000", underline=None):
    return Font(name=FONT, size=size, bold=bold, italic=italic, color=color, underline=underline)


def fill(color):
    return PatternFill("solid", fgColor=color)


def q(blatt):
    """Blattname fuer Formeln (immer in Hochkommas, wegen Umlauten)."""
    return "'" + blatt.replace("'", "''") + "'"


def rahmen(ws, r1, r2, c1, c2, aussen=DICK, innen=DUENN):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = Border(left=aussen if c == c1 else innen,
                                          right=aussen if c == c2 else innen,
                                          top=aussen if r == r1 else innen,
                                          bottom=aussen if r == r2 else innen)


def schreibe(ws, ref, wert, **stil):
    c = ws[ref]
    c.value = wert
    for k, v in stil.items():
        setattr(c, k, v)
    return c


def spalten(ws, breiten):
    for col, w in breiten.items():
        ws.column_dimensions[col].width = w


def seite_einrichten(ws, bereich, quer=True, skalierung=None, hoch=1, fuss=None, fuss_links=None):
    ws.print_area = bereich
    ws.page_setup.orientation = "landscape" if quer else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    if skalierung:      # feste Skalierung, damit die manuellen Seitenumbrueche gelten
        ws.page_setup.scale = skalierung
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
    else:
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = hoch
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.31, right=0.31, top=0.39, bottom=0.55, header=0.2, footer=0.25)
    ws.print_options.horizontalCentered = True
    for teil, text in ((ws.oddFooter.center, fuss), (ws.oddFooter.left, fuss_links)):
        if text:
            teil.text = text
            teil.size = 7
            teil.font = "Arial,Italic"
    ws.sheet_view.showGridLines = False


def auswahl(ws, zelle):
    ws.sheet_view.selection[0].activeCell = zelle
    ws.sheet_view.selection[0].sqref = zelle


def regel(typ, bereiche, titel, text, stil="stop", **kw):
    dv = DataValidation(type=typ, allow_blank=True, showErrorMessage=stil is not None,
                        errorStyle=stil or "stop", errorTitle=titel, error=text, **kw)
    for b in bereiche.split():
        dv.add(b)
    return dv


# ---------------------------------------------------------------- Fahrtenbuchseite
def seitenblock(ws, top, *, excel, monat=None, kennz=None, seite="", nr_start=None, summe=None,
                zeilenhoehe=22.5):
    """Eine Druckseite wie im Muster ab Zeile top: Titel, Kopf, Uebertrag, Spaltennummern und
    ZEILEN Fahrten, optional mit Summenzeile. Gibt Zeilen der Fahrten, Uebertrag und Summe zurueck."""
    for off, h in ((TITEL, 24), (1, 6), (KOPF1, 16.5), (KOPF2, 16.5), (UEBERTRAG, 16.5), (NUMMERN, 12)):
        ws.row_dimensions[top + off].height = h
    zeilen = list(range(top + DATEN, top + DATEN + ZEILEN))

    t = top + TITEL
    unten = Alignment(horizontal="left", vertical="bottom")
    ws.merge_cells(f"A{t}:C{t}")
    schreibe(ws, f"A{t}", "Monat/Jahr:", font=font(10), alignment=unten)
    schreibe(ws, f"D{t}", monat, font=font(12, bold=True), alignment=unten, border=Border(bottom=DUENN))
    schreibe(ws, f"E{t}", "Kennzeichen:", font=font(8, color=GRAU),
             alignment=Alignment(horizontal="right", vertical="bottom"))
    schreibe(ws, f"F{t}", kennz, font=font(11, bold=True), alignment=unten, border=Border(bottom=DUENN))
    ws.merge_cells(f"H{t}:L{t}")
    schreibe(ws, f"H{t}", "Fahrtenbuch", font=font(12, bold=True, color=GRAU),
             alignment=Alignment(horizontal="center", vertical="bottom"))
    ws.merge_cells(f"N{t}:O{t}")
    schreibe(ws, f"N{t}", seite, font=font(8, color=GRAU),
             alignment=Alignment(horizontal="right", vertical="bottom"))

    k1, k2 = top + KOPF1, top + KOPF2
    for c1, c2, n, text in KOPF:
        if c1 != c2 or n > 1:
            ws.merge_cells(f"{c1}{k1}:{c2}{k1 + n - 1}")
        ws[f"{c1}{k1}"].value = text
    for col, text in KOPF_UNTEN.items():
        ws[f"{col}{k2}"].value = text
    for r in (k1, k2):
        for col in BREITEN:
            c = ws[f"{col}{r}"]
            c.font, c.alignment, c.fill = font(8, bold=True), MITTE, fill(KOPF_BG)

    u = top + UEBERTRAG
    schreibe(ws, f"H{u}", "Übertr. km", font=font(7), alignment=MITTE, fill=fill(UEBERTR_BG))
    schreibe(ws, f"I{u}", None, font=font(9, bold=True), alignment=MITTE, number_format="0")
    if excel:
        ws[f"I{u}"].fill = fill(CALC_BG)

    n = top + NUMMERN
    for i, col in enumerate(BREITEN, 1):
        schreibe(ws, f"{col}{n}", i, font=font(7), alignment=MITTE)

    for i, r in enumerate(zeilen):
        ws.row_dimensions[r].height = zeilenhoehe
        for col in BREITEN:
            c = ws[f"{col}{r}"]
            c.font = font(9)
            c.alignment = LINKSB if col in LINKS else MITTE
            c.number_format = FORMATE.get(col, "General")
            if excel and col in "HJKL":
                c.fill = fill(CALC_BG)
        if nr_start is not None:
            schreibe(ws, f"A{r}", nr_start + i, font=font(8, color=GRAU))
    rahmen(ws, k1, zeilen[-1], 1, len(BREITEN))
    # Uebertragszeile wie im Muster: links und rechts je ein leeres Feld (ohne Zellverbund, damit
    # Excel und LibreOffice den Rahmen gleich zeichnen)
    for col in "ABCDEFGJKLMNO":
        c = ws[f"{col}{u}"]
        c.border = Border(left=c.border.left if col in "AJ" else None,
                          right=c.border.right if col in "GO" else None, top=DUENN, bottom=DUENN)

    s = None
    if summe:
        s = zeilen[-1] + 1
        ws.row_dimensions[s].height = 21
        ws.merge_cells(f"A{s}:I{s}")
        schreibe(ws, f"A{s}", summe, font=font(9, bold=True), alignment=RECHTS)
        for col in "JKLMO":
            schreibe(ws, f"{col}{s}", None, font=font(9, bold=True), alignment=MITTE,
                     number_format="#,##0.0;-#,##0.0;;@" if col == "M" else "#,##0;-#,##0;;@")
            if excel:
                ws[f"{col}{s}"].fill = fill(CALC_BG)
        schreibe(ws, f"N{s}", "gesamt km", font=font(8, bold=True), alignment=RECHTS)
        rahmen(ws, s, s, 1, len(BREITEN))

    if excel:   # Hilfsspalten, werden nicht gedruckt
        for col, text in (("P", "Fahrtart\n(autom.)"), ("Q", "Prüfung (wird nicht gedruckt)")):
            ws.merge_cells(f"{col}{k1}:{col}{k2}")
            schreibe(ws, f"{col}{k1}", text, font=font(8, italic=True, color=GRAU), alignment=MITTE)
            for r in (k1, k2, *zeilen):
                ws[f"{col}{r}"].border = Border(left=HELL, right=HELL, top=HELL, bottom=HELL)
        for r in zeilen:
            schreibe(ws, f"P{r}", None, font=font(8), alignment=MITTE)
            schreibe(ws, f"Q{r}", None, font=font(8), alignment=LINKSB)
    return {"zeilen": zeilen, "uebertrag": f"I{u}", "summe": s}


def fahrt_formeln(ws, r, km_vor, art_vor):
    """Formeln einer Fahrtzeile. km_vor: Zelle mit dem letzten km-Stand davor (Uebertrag oder
    Kette in Spalte R), art_vor: Zelle mit der letzten Fahrtart davor (Spalte S) oder '""'."""
    ws[f"H{r}"] = f'=IF(COUNTA(B{r}:G{r},I{r})=0,"",IF(N({km_vor})=0,"",{km_vor}))'
    for col, art in (("J", "gesch."), ("K", "WO/AO"), ("L", "privat")):
        ws[f"{col}{r}"] = f'=IF(AND($P{r}="{art}",ISNUMBER(H{r}),ISNUMBER(I{r})),I{r}-H{r},"")'
    e = f'SUBSTITUTE(UPPER(E{r})," ","")'
    ws[f"P{r}"] = (f'=IF(AND(D{r}="",E{r}="",I{r}=""),"",'
                   f'IF(OR(LEFT(UPPER(TRIM(D{r})),6)="PRIVAT",LEFT(UPPER(TRIM(E{r})),6)="PRIVAT"),"privat",'
                   f'IF(OR(LEFT({e},5)="WO/AO",LEFT({e},5)="AO/WO"),"WO/AO",'
                   f'IF(AND(LEFT(LOWER(TRIM(E{r})),6)="zurück",{art_vor}<>""),{art_vor},"gesch."))))')
    ws[f"Q{r}"] = (f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r}),I{r}<H{r}),"km Ankunft ist kleiner als km Abfahrt",'
                   f'IF(AND(ISNUMBER(H{r}),N({km_vor})>0,H{r}<>N({km_vor})),'
                   f'"km-Stand weicht von der letzten Ankunft ("&N({km_vor})&") ab – Poolfahrzeugdisponent informieren",'
                   f'IF(AND(ISNUMBER(I{r}),D{r}="",E{r}=""),"Fahrweg bzw. Zweck fehlt","")))')
    ws[f"R{r}"] = f'=IF(ISNUMBER(I{r}),I{r},{km_vor})'
    ws[f"S{r}"] = f'=IF(P{r}<>"",P{r},{art_vor})'


def fahrten_kette(ws, bloecke, km_start, art_start):
    """Uebertrag je Seite und Formeln aller Fahrten; der km-Stand laeuft ueber alle Seiten weiter."""
    km_vor, art_vor = km_start, art_start
    for b in bloecke:
        if km_vor is not None:
            ws[b["uebertrag"]] = f'=IF(N({km_vor})=0,"",{km_vor})'
        km_vor = b["uebertrag"]
        for r in b["zeilen"]:
            fahrt_formeln(ws, r, km_vor, art_vor)
            km_vor, art_vor = f"R{r}", f"S{r}"


def summen(ws, s, bloecke):
    for col in "JKLM":
        ws[f"{col}{s}"] = "=SUM(" + ",".join(f"{col}{b['zeilen'][0]}:{col}{b['zeilen'][-1]}"
                                             for b in bloecke) + ")"
    ws[f"O{s}"] = f"=SUM(J{s}:L{s})"


def pruefregeln(ws, bloecke):
    """Gueltigkeitsregeln, Auswahllisten und Warnfarben fuer die Fahrtzeilen."""
    def b(*cols):
        return " ".join(f"{c}{x['zeilen'][0]}:{c}{x['zeilen'][-1]}" for x in bloecke for c in cols)
    erste = bloecke[0]["zeilen"][0]
    for dv in (
        regel("whole", b("B"), "Tag", "Bitte nur den Tag des Monats eintragen (1 bis 31).",
              operator="between", formula1="1", formula2="31"),
        regel("time", b("C", "G"), "Uhrzeit", "Bitte die Uhrzeit als hh:mm eintragen, z. B. 07:30.",
              operator="between", formula1="0", formula2="0.999305555555556"),
        regel("list", b("E"), "", "", stil=None, formula1="FB_ListeZweck"),
        regel("whole", b("H") + " " + " ".join(x["uebertrag"] for x in bloecke), "km-Stand",
              "Bitte den km-Stand als ganze Zahl eintragen.", operator="greaterThanOrEqual", formula1="0"),
        regel("custom", b("I"), "km-Stand bei Ankunft",
              "Bitte den km-Stand als ganze Zahl eintragen. Er darf nicht kleiner sein als der "
              "km-Stand bei Abfahrt.", formula1=f"AND(ISNUMBER(I{erste}),I{erste}>=N(H{erste}))"),
        regel("decimal", b("M"), "getankt", "Bitte die getankte Menge in Litern eintragen, z. B. 45,7.",
              operator="between", formula1="0", formula2="500"),
        regel("list", b("N"), "", "", stil=None, formula1="FB_ListeFahrer"),
        regel("list", b("P"), "Fahrtart", "Bitte gesch., WO/AO oder privat auswählen.",
              formula1='"gesch.,WO/AO,privat"'),
    ):
        ws.add_data_validation(dv)
    cf = ws.conditional_formatting
    cf.add(b("H"), FormulaRule(formula=[f'LEFT($Q{erste},8)="km-Stand"'], fill=fill(WARN_BG),
                               font=Font(color=WARN_FG, bold=True)))
    cf.add(b("I"), FormulaRule(formula=[f'LEFT($Q{erste},10)="km Ankunft"'], fill=fill(FEHLER_BG),
                               font=Font(color=FEHLER_FG, bold=True)))
    cf.add(b("D", "E"), FormulaRule(formula=[f'$Q{erste}="Fahrweg bzw. Zweck fehlt"'], fill=fill(WARN_BG)))
    cf.add(b("Q"), FormulaRule(formula=[f'$Q{erste}<>""'], font=Font(color=WARN_FG, bold=True)))


# ---------------------------------------------------------------- Blaetter
def monatsblatt(wb, idx):
    kurz, lang = MONATE[idx]
    ws = wb.create_sheet(kurz)
    spalten(ws, BREITEN | HILFE)
    ws.column_dimensions["R"].hidden = ws.column_dimensions["S"].hidden = True
    bloecke = []
    for k in range(SEITEN_MONAT):
        top = 1 + k * BLOCK
        if k:
            ws.row_breaks.append(Break(id=top - 1))
        bloecke.append(seitenblock(
            ws, top, excel=True, monat=f'=IF(FB_Jahr="","{lang} /","{lang} / "&FB_Jahr)',
            kennz='=IF(FB_Kennzeichen="","",FB_Kennzeichen)', seite=f"Seite {k + 1} von {SEITEN_MONAT}",
            nr_start=k * ZEILEN + 1, summe=f"Summe {lang}" if k == SEITEN_MONAT - 1 else None))
    letzte = bloecke[-1]["zeilen"][-1]
    if idx == 0:
        km_start, art_start = "FB_kmStart", '""'
    else:      # Uebertrag aus dem Vormonat
        vor = q(MONATE[idx - 1][0])
        km_start, art_start = f"{vor}!$R${letzte}", f"{vor}!$S${letzte}"
    fahrten_kette(ws, bloecke, km_start, art_start)
    summen(ws, bloecke[-1]["summe"], bloecke)
    pruefregeln(ws, bloecke)
    seite_einrichten(ws, f"A1:O{bloecke[-1]['summe']}", skalierung=SKALIERUNG, fuss=HINWEIS_FUSS)
    auswahl(ws, f"B{bloecke[0]['zeilen'][0]}")
    return bloecke


def beispielblatt(wb, excel):
    ws = wb.create_sheet("Beispiel")
    spalten(ws, BREITEN | HILFE)
    ws.column_dimensions["R"].hidden = ws.column_dimensions["S"].hidden = True
    b = seitenblock(ws, 1, excel=excel, monat="Februar / 2013", kennz="Beispiel",
                    seite="Ausfüllbeispiel aus dem Musterfahrtenbuch", nr_start=1, summe="Summe Februar")
    ws[b["uebertrag"]] = BEISPIEL_UEBERTRAG
    for r, (tag, ab, weg, zweck, besucht, an, km, liter) in zip(b["zeilen"], BEISPIEL):
        for col, wert in (("B", tag), ("D", weg), ("E", zweck), ("F", besucht), ("I", km), ("M", liter)):
            if wert not in (None, ""):
                ws[f"{col}{r}"] = wert
        for col, zeit in (("C", ab), ("G", an)):
            if zeit:
                ws[f"{col}{r}"] = datetime.time(*map(int, zeit.split(":")))
        if weg == "PRIVAT" or weg.startswith("Probefahrt"):
            ws[f"D{r}"].font = font(9, bold=True, italic=True)
    fahrten_kette(ws, [b], None, '""')
    summen(ws, b["summe"], [b])
    for von, bis, name in BEISPIEL_FAHRER:      # Klammer im Muster = verbundene Zellen
        r1, r2 = b["zeilen"][von - 1], b["zeilen"][bis - 1]
        if r2 > r1:
            ws.merge_cells(f"N{r1}:N{r2}")
            ws.merge_cells(f"O{r1}:O{r2}")
        ws[f"N{r1}"] = name
        schreibe(ws, f"O{r1}", name.split(",")[0].split()[-1], font=font(11, italic=True, color="1F3864"),
                 alignment=MITTE)
    if excel:
        pruefregeln(ws, [b])
    seite_einrichten(ws, f"A1:O{b['summe']}", skalierung=SKALIERUNG, fuss=HINWEIS_FUSS)
    ws.sheet_properties.tabColor = "A6A6A6"
    return b


def vordruckseiten(wb, anzahl):
    ws = wb.create_sheet("Fahrtenbuch")
    spalten(ws, BREITEN)
    for k in range(anzahl):
        top = 1 + k * BLOCK
        if k:
            ws.row_breaks.append(Break(id=top - 1))
        seitenblock(ws, top, excel=False, seite=f"Seite {k + 1}", zeilenhoehe=24)   # mehr Platz zum Schreiben
    seite_einrichten(ws, f"A1:O{anzahl * BLOCK}", skalierung=SKALIERUNG, fuss=HINWEIS_FUSS)


def deckblatt(ws, excel, jahr=None):
    spalten(ws, {"A": 2, "B": 28, "C": 30, "D": 3, "E": 12, "F": 8, "G": 10, "H": 10, "I": 9, "J": 9,
                 "K": 9, "L": 9, "M": 9})
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[3].height = 8
    for r in range(4, 32):
        ws.row_dimensions[r].height = 18
    schreibe(ws, "B1", "Fahrtenbuch", font=font(24, bold=True), alignment=Alignment(vertical="bottom"))
    schreibe(ws, "B2", ("Monatsblätter Januar bis Dezember" if excel else "Druckvorlage zum handschriftlichen Führen")
             + " · nach dem Musterfahrtenbuch im DIN A6-ähnlichen Format", font=font(10, color=GRAU))

    def abschnitt(ref, text):
        a, e = ref.split(":")
        ws.merge_cells(ref)
        schreibe(ws, a, text, font=font(11, bold=True), alignment=Alignment(vertical="bottom"))
        for c in ws[ref][0]:
            c.border = Border(bottom=DICK)

    abschnitt("B4:C4", "Fahrzeug")
    eingabe = Border(left=HELL, right=HELL, top=HELL, bottom=HELL) if excel else Border(bottom=DUENN)
    felder = [("Jahr", D_JAHR)] + [(t, D_KENNZ + i) for i, t in enumerate(FAHRZEUG)]
    felder.append(("km-Stand aktuell (automatisch)" if excel else "km-Stand bei Abschluss", D_LETZTE))
    for text, r in felder:
        schreibe(ws, f"B{r}", text, font=font(9), alignment=Alignment(vertical="center"))
        c = schreibe(ws, f"C{r}", None, font=font(10, bold=r in (D_JAHR, D_KENNZ)), border=eingabe,
                     alignment=Alignment(horizontal="left", vertical="center"))
        if excel:
            c.fill = fill(CALC_BG if r == D_LETZTE else INPUT_BG)
    if excel:
        ws[f"C{D_JAHR}"] = jahr
        kette = f"{q(MONATE[-1][0])}!$R${SEITEN_MONAT * BLOCK}"    # letzter km-Stand im Dezember
        ws[f"C{D_LETZTE}"] = f'=IF(N({kette})=0,"",{kette})'
    for r in (D_JAHR, D_KMSTART, D_LETZTE):
        ws[f"C{r}"].number_format = "0"
    for r in (D_KENNZ + 7, D_KENNZ + 9):          # Datumsfelder
        ws[f"C{r}"].number_format = "DD.MM.YYYY"

    abschnitt("B18:C18", "Berechtigte Fahrer" + (" (Auswahlliste für Spalte 14)" if excel else ""))
    schreibe(ws, "B19", "Name, Vorname", font=font(8, color=GRAU), alignment=Alignment(vertical="bottom"))
    schreibe(ws, "C19", "Bemerkung (z. B. Personal-Nr.)", font=font(8, color=GRAU),
             alignment=Alignment(vertical="bottom"))
    for r in D_FAHRER:
        for col in "BC":
            c = schreibe(ws, f"{col}{r}", None, font=font(9), border=eingabe,
                         alignment=Alignment(vertical="center"))
            if excel:
                c.fill = fill(INPUT_BG)

    # Jahresuebersicht
    abschnitt("E4:M4", "Jahresübersicht" + ("" if excel else " (am Monatsende aus den Seiten übertragen)"))
    for ref, text in (("E5:E6", "Monat"), ("F5:F6", "Fahrten"), ("G5:H5", "km-Zählerstand"),
                      ("I5:L5", "gefahrene km"), ("M5:M6", "getankt\n[l]")):
        ws.merge_cells(ref)
        ws[ref.split(":")[0]].value = text
    for col, text in zip("GHIJKL", ("Beginn", "Ende", "gesch.", "WO/AO", "privat", "gesamt")):
        ws[f"{col}6"].value = text
    for r in (5, 6):
        for col in "EFGHIJKLM":
            c = ws[f"{col}{r}"]
            c.font, c.alignment, c.fill = font(8, bold=True), MITTE, fill(KOPF_BG)
    erste, letzte = D_UEBERSICHT, D_UEBERSICHT + 11
    summe, anteil = letzte + 1, letzte + 2
    km = "#,##0;-#,##0;;@"
    for i, (kurz, lang) in enumerate(MONATE):
        r = erste + i
        c = schreibe(ws, f"E{r}", lang, font=font(9),
                     alignment=Alignment(horizontal="left", vertical="center", indent=1))
        if excel:
            s = q(kurz)
            daten = ",".join(f"{s}!$I${a}:$I${a + ZEILEN - 1}" for a in (1 + DATEN + k * BLOCK for k in range(SEITEN_MONAT)))
            ende, zsum = SEITEN_MONAT * BLOCK, SEITEN_MONAT * BLOCK + 1
            c.hyperlink = Hyperlink(ref=f"E{r}", location=f"{s}!A1", display=lang)
            c.font = font(9, color=LINK, underline="single")
            ws[f"F{r}"] = f"=COUNT({daten})"
            ws[f"G{r}"] = f'=IF(F{r}=0,"",{s}!$I${1 + UEBERTRAG})'
            ws[f"H{r}"] = f'=IF(F{r}=0,"",{s}!$R${ende})'
            for col, quelle in zip("IJK", "JKL"):
                ws[f"{col}{r}"] = f"={s}!${quelle}${zsum}"
            ws[f"L{r}"] = f"=SUM(I{r}:K{r})"
            ws[f"M{r}"] = f"={s}!$M${zsum}"
    eingerueckt = Alignment(horizontal="left", vertical="center", indent=1)
    schreibe(ws, f"E{summe}", "Summe", font=font(9, bold=True), alignment=eingerueckt)
    if excel:
        schreibe(ws, f"E{anteil}", "Anteil", font=font(8, color=GRAU), alignment=eingerueckt)
        ws[f"F{summe}"] = f"=SUM(F{erste}:F{letzte})"
        ws[f"G{summe}"] = f'=IF(COUNT(G{erste}:G{letzte})=0,"",MIN(G{erste}:G{letzte}))'
        ws[f"H{summe}"] = f'=IF(COUNT(H{erste}:H{letzte})=0,"",MAX(H{erste}:H{letzte}))'
        for col in "IJKLM":
            ws[f"{col}{summe}"] = f"=SUM({col}{erste}:{col}{letzte})"
        for col in "IJK":
            ws[f"{col}{anteil}"] = f'=IF($L${summe}=0,"",{col}{summe}/$L${summe})'
    for r in range(erste, (anteil if excel else summe) + 1):
        for col in "FGHIJKLM":
            c = ws[f"{col}{r}"]
            c.font = font(9, bold=r == summe, color=GRAU if r == anteil else "000000")
            c.alignment = MITTE
            c.number_format = {"F": "0;-0;;@", "G": "0", "H": "0", "M": "#,##0.0;-#,##0.0;;@"}.get(col, km)
            if r == anteil:
                c.number_format = "0.0%"
            if excel and r <= summe:
                c.fill = fill(CALC_BG)
    rahmen(ws, 5, summe, 5, 13)
    for col in "EFGHIJKLM":          # Summenzeile absetzen, Anteil ohne Rahmen
        c = ws[f"{col}{summe}"]
        c.border = Border(left=c.border.left, right=c.border.right, top=DICK, bottom=DICK)
    if excel:
        schreibe(ws, f"E{anteil + 2}", "Alle Werte rechnen sich aus den Monatsblättern; ein Klick auf "
                 "den Monat öffnet das Blatt.", font=font(8, italic=True, color=GRAU))
        ws.merge_cells("B31:M31")
        schreibe(ws, "B31", "Gelb hinterlegt = bitte ausfüllen · In den Monatsblättern die weißen Felder "
                 "ausfüllen, grau hinterlegte Felder rechnen automatisch · Erläuterungen im Blatt "
                 "„Hinweise“, Musterzeilen im Blatt „Beispiel“.", font=font(8, italic=True, color=GRAU),
                 alignment=Alignment(vertical="center", wrap_text=True))
        ws.row_dimensions[31].height = 24
        ws.add_data_validation(regel("whole", f"C{D_JAHR}", "Jahr", "Bitte das Jahr vierstellig eintragen.",
                                     operator="between", formula1="2000", formula2="2100"))
        ws.add_data_validation(regel("whole", f"C{D_KMSTART}", "km-Stand",
                                     "Bitte den km-Stand als ganze Zahl eintragen.",
                                     operator="greaterThanOrEqual", formula1="0"))
        ws.add_data_validation(regel("date", f"C{D_KENNZ + 7} C{D_KENNZ + 9}", "Datum",
                                     "Bitte ein Datum eintragen, z. B. 01.01.2027.", operator="between",
                                     formula1="36526", formula2="73415"))
        ws.sheet_properties.tabColor = LINK
        auswahl(ws, f"C{D_KENNZ}")
    else:
        ws.merge_cells("B31:M31")
        schreibe(ws, "B31", "Hinweise zum Ausfüllen, Abkürzungen und ein Ausfüllbeispiel stehen auf den "
                 "folgenden Seiten.", font=font(8, italic=True, color=GRAU), alignment=Alignment(vertical="center"))
    seite_einrichten(ws, "A1:M31")


def hinweise(ws, excel):
    spalten(ws, {"A": 2, "B": 13, "C": 128})
    schreibe(ws, "B1", "Hinweise zum Fahrtenbuch", font=font(16, bold=True), alignment=Alignment(vertical="bottom"))
    ws.row_dimensions[1].height = 30
    teile = [("Eintragungen je Fahrt (Spalten wie im Musterfahrtenbuch)", [
        ("Spalte 1–3", "Laufende Nummer, Tag (tt) und Abfahrtszeit (hh:mm)."),
        ("Spalte 4", "Fahrweg „von – über – nach“ mit Ziel (Ort, Straße). Örtlichkeiten nach Ril 100 "
                     "abkürzen, z. B. „Stp Nürnberg“, „Bst Bamberg“, „BÜ, km 50,123“. Privatfahrten nur "
                     "mit „PRIVAT“ kennzeichnen."),
        ("Spalte 5–6", "Zweck der Fahrt und besuchte Personen, Firmen, Behörden oder Objekte – "
                       "Abkürzungen siehe „Abkürzungen“. Für die Rückfahrt genügt „zurück“."),
        ("Spalte 7–9", "Ankunftszeit sowie km-Zählerstand bei Abfahrt und Ankunft. Der km-Stand bei "
                       "Abfahrt muss dem km-Stand der letzten Ankunft entsprechen; oben auf jeder Seite "
                       "steht der Übertrag."),
        ("Spalte 10–12", "Gefahrene km aufteilen: geschäftlich (gesch.), Fahrten Wohn-/Arbeitsort "
                         "(WO/AO bzw. AO/WO) oder privat."),
        ("Spalte 13", "Getankte Menge in Litern."),
        ("Spalte 14–15", "Name und Unterschrift des Fahrers. Aufeinanderfolgende Fahrten desselben "
                         "Fahrers können mit einer Klammer zusammengefasst und einmal unterschrieben "
                         "werden (siehe Beispiel)."),
    ]), ("Allgemein", [
        ("•", "Jede Fahrt sofort und vollständig eintragen – eine Zeile je Fahrt, Hin- und Rückfahrt getrennt."),
        ("•", "Einen abweichenden km-Stand sofort dem Poolfahrzeugdisponenten melden und im Fahrtenbuch vermerken."),
        ("•", "Weitere Abkürzungen können ergänzt oder in Tabellenform als Anlage fest mit dem "
              "Fahrtenbuch verbunden werden."),
        ("•", HINWEIS_FUSS),
    ])]
    if excel:
        teile.append(("Excel-Version", [
            ("Eingabe", "Deckblatt: gelbe Felder ausfüllen (Jahr, Fahrzeugdaten, km-Stand bei Beginn, "
                        "berechtigte Fahrer). Monatsblätter: weiße Felder ausfüllen; grau hinterlegte "
                        "Felder werden berechnet."),
            ("km-Stand", "Der km-Stand bei Abfahrt wird aus der letzten Ankunft übernommen – auch über "
                         "Seiten- und Monatsgrenzen. Zeigt der Tacho etwas anderes, den abgelesenen Wert "
                         "eintragen: Die Zelle wird orange und rechts neben der Tabelle erscheint ein Hinweis."),
            ("Fahrtart", "Die gefahrenen km werden aus den Zählerständen berechnet und nach der Fahrtart "
                         "(rechts neben der Tabelle, wird nicht gedruckt) auf die Spalten 10–12 verteilt: "
                         "„WO/AO“ oder „AO/WO“ im Zweck (auch mit DR oder Ber) → WO/AO; „PRIVAT“ im "
                         "Fahrweg oder Zweck → privat; „zurück“ → wie die vorherige Fahrt; sonst → gesch. "
                         "Eine andere Fahrtart kann in der Zeile ausgewählt werden."),
            ("Auswertung", "Übertrag, Monatssummen und die Jahresübersicht auf dem Deckblatt rechnen sich "
                           "automatisch. Neues Jahr: Datei kopieren, Jahr und km-Stand bei Beginn ändern, "
                           "Einträge löschen."),
            ("Drucken", f"Ein Monatsblatt ergibt {SEITEN_MONAT} Seiten A4 quer mit je {ZEILEN} Fahrten. "
                        "Ausgedruckt kann es unterschrieben werden. Das Blatt „Beispiel“ zeigt die "
                        "Mustereinträge aus dem Musterfahrtenbuch."),
            ("Wichtig", "Eine Excel-Datei lässt sich nachträglich unbemerkt ändern und erfüllt deshalb nicht "
                        "die steuerlichen Anforderungen an ein ordnungsgemäßes Fahrtenbuch. Wird das "
                        "Fahrtenbuch für das Finanzamt gebraucht (z. B. Fahrtenbuchmethode), die "
                        "Druckvorlage verwenden: vor Gebrauch fest binden und handschriftlich führen."),
        ]))
    else:
        teile.append(("Druckvorlage", [
            ("•", "Seiten vor Gebrauch fest binden (z. B. Klebe- oder Spiralbindung), der Reihe nach "
                  "verwenden und keine Seiten entfernen."),
            ("•", "Jede Seite beginnt mit Monat/Jahr, Kennzeichen und dem Übertrag des km-Stands der "
                  "vorherigen Seite. Ein neuer Monat beginnt auf einer neuen Seite."),
            ("•", "Die Jahresübersicht auf dem Deckblatt am Monatsende aus den Seiten übertragen."),
        ]))
    r = 2
    for titel, punkte in teile:
        r += 1
        ws.row_dimensions[r].height = 22
        ws.merge_cells(f"B{r}:C{r}")
        schreibe(ws, f"B{r}", titel, font=font(11, bold=True), alignment=Alignment(vertical="bottom"))
        for c in ws[f"B{r}:C{r}"][0]:
            c.border = Border(bottom=DICK)
        r += 1
        ws.row_dimensions[r].height = 5
        for label, text in punkte:
            r += 1
            wichtig = label == "Wichtig"
            schreibe(ws, f"B{r}", label, font=font(9, bold=True, color=FEHLER_FG if wichtig else "000000"),
                     alignment=Alignment(horizontal="right" if label == "•" else "left", vertical="top",
                                         indent=1 if label == "•" else 0))
            schreibe(ws, f"C{r}", text, font=font(9, bold=wichtig),
                     alignment=Alignment(vertical="top", wrap_text=True))
            ws.row_dimensions[r].height = 13 * (1 + len(text) // 150) + 4
        r += 1
    seite_einrichten(ws, f"A1:C{r}")
    ws.sheet_properties.tabColor = "A6A6A6"


def abkuerzungen(ws, excel):
    spalten(ws, {"A": 2, "B": 42, "C": 30, "D": 26, "E": 4, "F": 22, "G": 40})
    ws.merge_cells("B2:D2")
    schreibe(ws, "B2", "Abkürzungsvorschläge für Fahrtenbucheinträge", font=font(13, bold=True),
             alignment=Alignment(vertical="center"))
    ws.row_dimensions[2].height = 30
    for col, text in zip("BCD", ("Zweck der Fahrt", "besuchte Personen /\nFirmen / etc.", "Abkürzung")):
        schreibe(ws, f"{col}3", text, font=font(9, bold=True), alignment=LINKSB)
    ws.row_dimensions[3].height = 26
    fett, normal = InlineFont(rFont=FONT, sz=9, b=True), InlineFont(rFont=FONT, sz=9)
    for i, (zweck, besucht, abk) in enumerate(ABKUERZUNGEN):
        r = 4 + i
        schreibe(ws, f"B{r}", zweck, font=font(9), alignment=LINKSB)
        schreibe(ws, f"C{r}", besucht, font=font(9), alignment=LINKSB)
        if " bzw. " in abk:           # "bzw." wie im Muster nicht fett
            a, b = abk.split(" bzw. ")
            wert = CellRichText([TextBlock(fett, a), TextBlock(normal, " bzw. "), TextBlock(fett, b)])
        else:
            wert = abk
        schreibe(ws, f"D{r}", wert, font=font(9, bold=True), alignment=LINKSB)
        ws.row_dimensions[r].height = 26 if len(zweck) > 45 else 16
    ende = 3 + len(ABKUERZUNGEN)
    rahmen(ws, 2, ende, 2, 4, aussen=DUENN)
    ws.merge_cells(f"B{ende + 2}:D{ende + 2}")
    schreibe(ws, f"B{ende + 2}", "Weitere Einträge können ergänzt oder in Tabellenform als Anlage fest mit "
             "dem Fahrtenbuch verbunden werden.", font=font(9, bold=True, italic=True),
             alignment=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[ende + 2].height = 26
    if excel:      # Auswahlliste fuer Spalte 5, wird nicht gedruckt
        ws.merge_cells("F2:G2")
        schreibe(ws, "F2", "Auswahlliste „Zweck der Fahrt“ (Spalte 5)", font=font(10, bold=True),
                 alignment=Alignment(vertical="center"))
        ws.merge_cells("F3:G3")
        schreibe(ws, "F3", "Erweiterbar: neue Einträge ohne Leerzeile unten anfügen. Andere Texte "
                 "können in Spalte 5 trotzdem eingetragen werden. Wird nicht gedruckt.",
                 font=font(8, italic=True, color=GRAU), alignment=LINKSB)
        schreibe(ws, f"F{ZWECK_ERSTE - 1}", "Eintrag", font=font(8, bold=True), fill=fill(KOPF_BG))
        for i in range(ZWECK_LETZTE - ZWECK_ERSTE + 1):
            wert = AUSWAHL_ZWECK[i] if i < len(AUSWAHL_ZWECK) else None
            schreibe(ws, f"F{ZWECK_ERSTE + i}", wert, font=font(9), fill=fill(INPUT_BG),
                     border=Border(left=HELL, right=HELL, top=HELL, bottom=HELL))
    seite_einrichten(ws, f"A1:D{ende + 2}", quer=False, fuss_links="Abkürzung der Örtlichkeiten gem. Ril 100")
    ws.sheet_properties.tabColor = "A6A6A6"


def namen(wb):
    for name, ref in (
        ("FB_Jahr", f"Deckblatt!$C${D_JAHR}"),
        ("FB_Kennzeichen", f"Deckblatt!$C${D_KENNZ}"),
        ("FB_kmStart", f"Deckblatt!$C${D_KMSTART}"),
        ("FB_ListeFahrer", f"OFFSET(Deckblatt!$B${D_FAHRER[0]},0,0,"
                           f"MAX(1,COUNTA(Deckblatt!$B${D_FAHRER[0]}:$B${D_FAHRER[-1]})),1)"),
        ("FB_ListeZweck", f"OFFSET({q('Abkürzungen')}!$F${ZWECK_ERSTE},0,0,"
                          f"MAX(1,COUNTA({q('Abkürzungen')}!$F${ZWECK_ERSTE}:$F${ZWECK_LETZTE})),1)"),
    ):
        wb.defined_names[name] = DefinedName(name, attr_text=ref)


def baue_excel(pfad, jahr):
    wb = Workbook()
    deck = wb.active
    deck.title = "Deckblatt"
    deckblatt(deck, excel=True, jahr=jahr)
    for i in range(len(MONATE)):
        monatsblatt(wb, i)
    hinweise(wb.create_sheet("Hinweise"), excel=True)
    abkuerzungen(wb.create_sheet("Abkürzungen"), excel=True)
    beispielblatt(wb, excel=True)
    namen(wb)
    wb.active = 0
    wb.save(pfad)


def baue_vordruck(pfad, seiten):
    wb = Workbook()
    deck = wb.active
    deck.title = "Deckblatt"
    deckblatt(deck, excel=False)
    hinweise(wb.create_sheet("Hinweise"), excel=False)
    abkuerzungen(wb.create_sheet("Abkürzungen"), excel=False)
    beispielblatt(wb, excel=False)
    vordruckseiten(wb, seiten)
    wb.save(pfad)


# ---------------------------------------------------------------- LibreOffice
MAKRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def libreoffice(xlsx, pdf_ordner=None, timeout=600):
    """Rechnet die Datei mit LibreOffice durch (speichert in derselben Datei), optional mit PDF."""
    xlsx = Path(xlsx).absolute()
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    with tempfile.TemporaryDirectory(prefix="lo-profil-") as profil:
        basis = ["soffice", "--headless", "--norestore", f"-env:UserInstallation={Path(profil).as_uri()}"]
        subprocess.run(basis + ["--terminate_after_init"], env=env, capture_output=True, timeout=120)
        (Path(profil) / "user" / "basic" / "Standard" / "Module1.xba").write_text(MAKRO)
        # Gebietsschema Deutsch, damit das PDF Zahlen wie Excel zeigt (45,7 und 1.019)
        reg = Path(profil) / "user" / "registrymodifications.xcu"
        reg.write_text(reg.read_text(encoding="utf-8").replace("</oor:items>", (
            '<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" '
            'oor:op="fuse"><value>de-DE</value></prop></item>\n</oor:items>')), encoding="utf-8")
        vorher = os.stat(xlsx).st_mtime_ns
        subprocess.run(basis + ["vnd.sun.star.script:Standard.Module1.RecalculateAndSave"
                                "?language=Basic&location=application", str(xlsx)],
                       env=env, capture_output=True, timeout=timeout)
        if os.stat(xlsx).st_mtime_ns == vorher:
            raise RuntimeError("LibreOffice hat die Datei nicht neu berechnet.")
        if pdf_ordner:
            subprocess.run(basis + ["--convert-to", "pdf", "--outdir", str(pdf_ordner), str(xlsx)],
                           env=env, capture_output=True, timeout=timeout)


def pruefen(calc):
    """Keine Fehlerwerte; das Beispielblatt muss die Werte des Musters ergeben."""
    wb = load_workbook(calc, data_only=True)
    fehler = [f"{ws.title}!{c.coordinate}" for ws in wb for row in ws.iter_rows() for c in row
              if isinstance(c.value, str) and (c.value in FEHLERWERTE or c.value.startswith("Err:"))]
    if fehler:
        raise RuntimeError(f"Formelfehler: {fehler[:20]}")
    ws = wb["Beispiel"]
    zeilen = list(range(1 + DATEN, 1 + DATEN + len(BEISPIEL)))
    s = 1 + DATEN + ZEILEN
    ist = tuple(round(ws[f"{col}{s}"].value or 0, 1) for col in "JKLM")
    km_ab = [ws[f"H{r}"].value for r in zeilen]
    soll_ab = [BEISPIEL_UEBERTRAG] + [e[6] for e in BEISPIEL[:-1]]
    arten = [ws[f"P{r}"].value for r in zeilen]
    hinweise_ = [ws[f"Q{r}"].value for r in zeilen if ws[f"Q{r}"].value]
    if ist != BEISPIEL_SUMMEN or km_ab != soll_ab or arten != BEISPIEL_ARTEN or hinweise_:
        raise RuntimeError(f"Beispiel weicht vom Muster ab: Summen {ist}, km Abfahrt {km_ab}, "
                           f"Fahrtarten {arten}, Hinweise {hinweise_}")


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def werte_speichern(xlsx, calc):
    """Schreibt die von LibreOffice berechneten Werte in die Formelzellen, damit auch die
    geschuetzte Ansicht sie zeigt. Excel rechnet beim Oeffnen ohnehin neu (fullCalcOnLoad)."""
    werte = load_workbook(calc, data_only=True)
    with zipfile.ZipFile(xlsx) as z:
        teile = {i.filename: z.read(i.filename) for i in z.infolist()}
        reihenfolge = [i.filename for i in z.infolist()]
    mappe = etree.fromstring(teile["xl/workbook.xml"])
    ziele = {r.get("Id"): r.get("Target") for r in etree.fromstring(teile["xl/_rels/workbook.xml.rels"])}
    anzahl = 0
    for blatt in mappe.find(NS + "sheets"):
        ziel = ziele[blatt.get(RNS + "id")]
        teil = ziel.lstrip("/") if ziel.startswith("/") else "xl/" + ziel
        ws = werte[blatt.get("name")]
        root = etree.fromstring(teile[teil])
        for c in root.iter(NS + "c"):
            if c.find(NS + "f") is None:
                continue
            wert = ws[c.get("r")].value
            v = c.find(NS + "v")
            if v is None:
                v = etree.SubElement(c, NS + "v")
            if isinstance(wert, bool):
                c.set("t", "b")
                v.text = "1" if wert else "0"
            elif isinstance(wert, (int, float)):
                c.attrib.pop("t", None)
                v.text = str(int(wert)) if float(wert).is_integer() else repr(float(wert))
            else:
                c.set("t", "str")
                v.text = "" if wert is None else str(wert)
            anzahl += 1
        teile[teil] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(xlsx, "w", zipfile.ZIP_DEFLATED) as z:
        for name in reihenfolge:
            z.writestr(name, teile[name])
    return anzahl


# ---------------------------------------------------------------- Ablauf
def main():
    p = argparse.ArgumentParser(description="Fahrtenbuch nach dem Musterfahrtenbuch erzeugen")
    p.add_argument("--jahr", type=int, default=datetime.date.today().year, help="Jahr auf dem Deckblatt")
    p.add_argument("--seiten", type=int, default=48, help="Fahrtenbuchseiten in der Druckvorlage")
    p.add_argument("--ausgabe", default=str(Path(__file__).resolve().parent / "ausgabe"), help="Ausgabeordner")
    a = p.parse_args()
    aus = Path(a.ausgabe)
    aus.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="fahrtenbuch-") as tmp:
        tmp = Path(tmp)
        xlsx, calc = tmp / "Fahrtenbuch.xlsx", tmp / "calc" / "Fahrtenbuch.xlsx"
        baue_excel(xlsx, a.jahr)
        calc.parent.mkdir()
        shutil.copy(xlsx, calc)
        libreoffice(calc)
        pruefen(calc)
        zellen = werte_speichern(xlsx, calc)
        shutil.copy(xlsx, aus / "Fahrtenbuch.xlsx")

        vordruck = tmp / "Fahrtenbuch_Druckvorlage.xlsx"
        baue_vordruck(vordruck, a.seiten)
        libreoffice(vordruck, pdf_ordner=tmp)
        pdf = tmp / "Fahrtenbuch_Druckvorlage.pdf"
        if not pdf.exists():
            raise RuntimeError("LibreOffice hat kein PDF erzeugt.")
        shutil.copy(pdf, aus / pdf.name)
    print(f"Fahrtenbuch.xlsx ({a.jahr}, {zellen} Formelzellen, Beispiel geprüft) und "
          f"Fahrtenbuch_Druckvorlage.pdf ({a.seiten} Seiten) -> {aus}")


if __name__ == "__main__":
    main()
