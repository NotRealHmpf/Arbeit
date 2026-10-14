"""Erzeugt aus „UVV Überwachung.xlsx“ eine neue, übersichtliche Arbeitsmappe für die
Unterweisung 1.3334 (Arbeiten im Gleisbereich, DGUV Vorschrift 78 / DGUV Regel 101-024).

    python build_uvv.py "<UVV Überwachung.xlsx>" "<ausgabe.xlsx>"

Blätter der neuen Datei:
  Tabelle X  - Übersicht je Mitarbeiter: rot = abgelaufen (364 Tage), lila = nicht mehr in ZW-Daten
  Tabelle Y  - Namen und FK aus „ZW-Daten (Stichtag)“, Grundlage für Zu-/Abgänge und Führungskraft
  Anleitung  - Bedienung und Farben

Quelle: Blatt „Tabelle1“ (Mitarbeiter) und „ZW-Daten (Stichtag)“ (Spalte A Name, Spalte BP FK).
Mitarbeiter, die in den ZW-Daten stehen, aber noch nicht in Tabelle1, werden unten angehängt.
Benötigt: openpyxl, lxml, LibreOffice (soffice) mit Calc.
"""
import datetime as dt
import os
import re
import shutil
import sys
import tempfile
import warnings
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.datetime import from_excel, to_excel
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.properties import Outline, PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "dashboard"))
# openpyxl kann den mitwachsenden Druckbereich nur nicht selbst auswerten - fürs Lesen egal
warnings.filterwarnings("ignore", message="Print area cannot be set")
from build_dashboard import libreoffice_recalc  # noqa: E402
from inject_cache import ERRORS, N, RNS, dump, set_cell_cache  # noqa: E402

# ---------------------------------------------------------------- Quelle
SRC_SHEET = "Tabelle1"
ZW_SHEET = "ZW-Daten (Stichtag)"
ZW_FIRST_ROW = 5
ZW_NAME_COL = 1     # A  „Name komplett“
ZW_FK_COL = 68      # BP „FK“
COURSE = "1.3334 Unterweisung Arbeiten im Gleisbereich nach DGUV Vorschrift 78; DGUV Regel 101-024"

# ---------------------------------------------------------------- Aufbau
X, Y = "Tabelle X", "Tabelle Y"
X_HDR, X_FIRST, X_LAST = 5, 6, 2500      # Kopfzeile, erste Datenzeile, Formeln vorbereitet bis
Y_HDR, Y_FIRST, Y_LAST = 4, 5, 2004      # wie in „ZW-Daten (Stichtag)“: Kopf Zeile 4, Daten ab 5
GUELTIG_TAGE, VORWARNUNG_TAGE = 364, 30
ABGANG_LISTE = 100                       # Zeilen der Abgangsliste in Tabelle Y

ST_NEU = "Neu – Name festschreiben"
ST_ABGANG = "Nicht mehr in ZW-Daten"
ST_ENTFAELLT = "UVV entfällt"
ST_FEHLT = "UVV-Datum fehlt"
ST_PRUEFEN = "Datum prüfen"
ST_ABGELAUFEN = "Abgelaufen"
ST_BALD = "Bald fällig"
ST_OK = "Gültig"

FONT = "Arial"
INK = "1F2933"
MUTED = "6B7280"
HEAD_IN = "334155"      # Eingabespalten
HEAD_CALC = "5B7088"    # automatische Spalten
HEAD_HELP = "9CA3AF"    # Hilfsspalten (ausgeblendet)
INPUT_BG = "FFF2CC"
PASTE_BG = "FFFAEB"
NEUTRAL = "F3F4F6"
GREEN, GREEN_INK = "E3F1E6", "1E7B34"
AMBER = "FFF0C2"
RED = "F8D7DA"
PURPLE = "E4D9F2"
BLUE, BLUE_STRONG = "DCEBF7", "C5DDF3"
GREY = "EDEDED"
DUP_INK = "B42318"
LINE = "D7DCE2"

# Spalten Tabelle X: Buchstabe, Überschrift, Breite, Art (in = Eingabe, calc = Formel, help = Hilfsspalte)
X_COLS = [
    ("A", "Personennummer", 18, "in"),
    ("B", "Name", 30, "in"),
    ("C", "Abteilung", 34, "in"),
    ("D", "Verantw. Führungskraft", 25, "calc"),
    ("E", "Letzte UVV (gültig ab)", 15, "in"),
    ("F", "Gültig bis", 13, "calc"),
    ("G", "Resttage", 11, "calc"),
    ("H", "Status", 26, "calc"),
    ("I", "Bezeichnung", 40, "in"),
    ("J", "Personalakte", 34, "in"),
    ("K", "Bemerkung", 50, "in"),
    ("L", "Suchschlüssel", 24, "help"),
    ("M", "Zeile in Tabelle Y", 10, "help"),
]
HIDDEN_DATA = ("I", "J", "K")   # einblendbar über das „+“ (Gliederung)
HIDDEN_HELP = ("L", "M")

Q = "'"
YA = f"{Q}{Y}{Q}!$A${Y_FIRST}:$A${Y_LAST}"
YB = f"{Q}{Y}{Q}!$B${Y_FIRST}:$B${Y_LAST}"
YD = f"{Q}{Y}{Q}!$D${Y_FIRST}:$D${Y_LAST}"
XB = f"{Q}{X}{Q}!$B${X_FIRST}:$B${X_LAST}"
XC = f"{Q}{X}{Q}!$C${X_FIRST}:$C${X_LAST}"
XH = f"{Q}{X}{Q}!$H${X_FIRST}:$H${X_LAST}"
XL = f"{Q}{X}{Q}!$L${X_FIRST}:$L${X_LAST}"
GUELTIG_REF, VORWARN_REF, STICHTAG_REF = "$H$1", "$H$2", "$A$4"


def key_formula(cell):
    """Suchschlüssel: Name ohne Leerzeichen („Heß, Astrid“ = „Heß,Astrid“)."""
    return f'IF({cell}="","",SUBSTITUTE(SUBSTITUTE(TRIM({cell}&"")," ",""),CHAR(160),""))'


def name_key(name):
    """Wie key_formula, in Python (split() trennt auch an geschützten Leerzeichen)."""
    return "".join(str(name or "").split()).lower()


def pretty_name(name):
    """„Kirch,Ulf“ -> „Kirch, Ulf“ (Schreibweise wie in Tabelle1)."""
    return re.sub(r"\s*,\s*", ", ", str(name).strip())


# ---------------------------------------------------------------- Formeln je Zeile
def x_formulas(r):
    return {
        "D": f'=IF(M{r}="","",INDEX({YB},M{r})&"")',
        "F": f'=IF(ISNUMBER(E{r}),E{r}+{GUELTIG_REF},"")',
        "G": f'=IF(ISNUMBER(F{r}),F{r}-{STICHTAG_REF},"")',
        "H": (f'=IF(B{r}="","",IF(_xlfn.ISFORMULA(B{r}),"{ST_NEU}",IF(M{r}="","{ST_ABGANG}",'
              f'IF(E{r}="entfällt","{ST_ENTFAELLT}",IF(E{r}="","{ST_FEHLT}",IF(NOT(ISNUMBER(E{r})),"{ST_PRUEFEN}",'
              f'IF(G{r}<0,"{ST_ABGELAUFEN}",IF(G{r}<={VORWARN_REF},"{ST_BALD}","{ST_OK}"))))))))'),
        "L": "=" + key_formula(f"B{r}"),
        "M": f'=IF(L{r}="","",IFERROR(MATCH(L{r},{YD},0),""))',
    }


def auto_name_formula(r):
    """Hängt Namen aus Tabelle Y an, die oberhalb noch nicht vorkommen (je Zeile der nächste).
    Bezieht sich nur auf Zeilen darüber, daher kein Zirkelbezug. Stoppt nach der ersten leeren
    Formelzeile, damit nur so viele Zeilen rechnen, wie es Zugänge gibt (WENN wertet nur den
    benötigten Zweig aus). INDEX(...;0) rechnet die Liste ohne Matrixformel (kein Strg+Umschalt+Enter)."""
    p = r - 1
    return (f'=IF(AND(_xlfn.ISFORMULA(B{p}),B{p}=""),"",IFERROR(SUBSTITUTE(SUBSTITUTE(TRIM(INDEX({YA},'
            f'MATCH(0,INDEX(COUNTIF(L${X_FIRST}:L{p},{YD})+({YD}=""),0),0))),", ",","),",",", "),""))')


# ---------------------------------------------------------------- Stil-Helfer
def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def put(ws, ref, value, f=None, fl=None, al=None, nf=None, border=None):
    c = ws[ref]
    c.value = value
    c.font = f or font()
    if fl:
        c.fill = fill(fl)
    if al:
        c.alignment = al
    if nf:
        c.number_format = nf
    if border:
        c.border = border
    return c


def box():
    s = Side(style="thin", color=LINE)
    return Border(left=s, right=s, top=s, bottom=s)


def fit_width(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
DATE = "DD.MM.YYYY"


# ---------------------------------------------------------------- Quelle lesen
def read_source(src):
    wb = load_workbook(src, data_only=True)
    t, z = wb[SRC_SHEET], wb[ZW_SHEET]
    people = []
    for r in range(2, t.max_row + 1):
        v = [t.cell(r, c).value for c in range(1, 14)]
        if not (v[0] or v[1]):
            continue
        uvv = v[7]
        if isinstance(uvv, (int, float)) and not isinstance(uvv, bool):
            uvv = from_excel(uvv)       # Datum als Zahl eingetragen (z. B. 46038)
        people.append({
            "nr": str(v[0]).strip() if v[0] is not None else None,
            "name": str(v[1]).strip() if v[1] else None,
            "abt": v[2],
            "bez": v[6],
            "uvv": uvv,
            "akte": v[10].strip() if isinstance(v[10], str) else v[10],
            "bem": v[12].strip() if isinstance(v[12], str) else v[12],
        })
    zw = []
    for r in range(ZW_FIRST_ROW, z.max_row + 1):
        name = z.cell(r, ZW_NAME_COL).value
        if not name:
            continue
        fk = z.cell(r, ZW_FK_COL).value
        zw.append({"name": str(name).strip(), "fk": None if fk in ERRORS else fk})
    known = {name_key(p["name"]) for p in people}
    neu = []
    for p in zw:
        k = name_key(p["name"])
        if k not in known:
            neu.append({"nr": None, "name": pretty_name(p["name"]), "abt": None, "bez": None,
                        "uvv": None, "akte": None, "bem": None})
            known.add(k)
    return people, neu, zw


# ---------------------------------------------------------------- Tabelle X
def build_x(ws, rows):
    last_data = X_FIRST + len(rows) - 1
    ws.sheet_properties.tabColor = HEAD_IN
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    for col, _, width, _ in X_COLS:
        ws.column_dimensions[col].width = width
    for col in HIDDEN_DATA:
        ws.column_dimensions[col].outlineLevel = 1
        ws.column_dimensions[col].hidden = True
    for col in HIDDEN_HELP:
        ws.column_dimensions[col].hidden = True
    ws.sheet_properties.outlinePr = Outline(summaryRight=False)

    # Kopf
    put(ws, "A1", "UVV-Überwachung · Arbeiten im Gleisbereich", font(16, True))
    put(ws, "A2", COURSE, font(9, color=MUTED))
    put(ws, "G1", "Gültigkeit (Tage)", font(9, color=MUTED), al=RIGHT)
    put(ws, "G2", "Vorwarnung (Tage)", font(9, color=MUTED), al=RIGHT)
    put(ws, "H1", GUELTIG_TAGE, font(10, True), INPUT_BG, LEFT, "0", box())
    put(ws, "H2", VORWARNUNG_TAGE, font(10, True), INPUT_BG, LEFT, "0", box())
    ws["H1"].comment = Comment("Vorgabe: Die UVV gilt 364 Tage ab „Letzte UVV“. Bei Bedarf hier ändern.", "UVV")
    ws["H2"].comment = Comment("„Bald fällig“ (gelb), wenn die UVV in so vielen Tagen oder weniger abläuft.", "UVV")
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 16

    rng = f"$H${X_FIRST}:$H${X_LAST}"
    tiles = [
        ("A", "Stichtag (heute)", "=TODAY()", NEUTRAL, DATE),
        ("B", "Mitarbeiter gesamt", f'=COUNTIF($B${X_FIRST}:$B${X_LAST},"?*")', NEUTRAL, "0"),
        ("C", "Gültig", f'=COUNTIF({rng},"{ST_OK}")', GREEN, "0"),
        ("D", "Bald fällig", f'=COUNTIF({rng},"{ST_BALD}")', AMBER, "0"),
        ("E", "Abgelaufen", f'=COUNTIF({rng},"{ST_ABGELAUFEN}")', RED, "0"),
        ("F", "Nicht mehr in ZW-Daten", f'=COUNTIF({rng},"{ST_ABGANG}")', PURPLE, "0"),
        ("G", "Neu / Daten fehlen",
         f'=COUNTIF({rng},"Neu*")+COUNTIF({rng},"{ST_FEHLT}")+COUNTIF({rng},"{ST_PRUEFEN}")', BLUE, "0"),
        ("H", "Stand ZW-Daten (Tabelle Y)", f"=IF({Q}{Y}{Q}!$G$1=\"\",\"nicht eingetragen\",{Q}{Y}{Q}!$G$1)",
         NEUTRAL, DATE),
    ]
    for col, label, formula, color, nf in tiles:
        put(ws, f"{col}3", label, font(8, color=MUTED), color, CENTER_WRAP, border=box())
        put(ws, f"{col}4", formula, font(15, True), color,
            Alignment(horizontal="center", vertical="center"), nf, box())
    ws["H4"].font = font(11, True)
    ws.row_dimensions[3].height = 24
    ws.row_dimensions[4].height = 28

    # Kopfzeile der Tabelle
    notes = {
        "B": "Neue Namen aus Tabelle Y erscheinen automatisch unter der Tabelle (blau). "
             "Bitte festschreiben: markieren, kopieren, als Werte einfügen.",
        "D": "Automatisch: Führungskraft aus Tabelle Y (Spalte B), gesucht über den Namen.",
        "E": "Datum der letzten Unterweisung. „entfällt“ eintragen, wenn keine UVV nötig ist.",
        "F": "Automatisch: Letzte UVV + Gültigkeit (H1).",
        "G": "Automatisch: Tage bis „Gültig bis“, negativ = abgelaufen.",
        "H": "Automatisch. Lila = Name steht nicht mehr in Tabelle Y (ZW-Daten), rot = abgelaufen, "
             "gelb = bald fällig, blau = neu / Datum fehlt.",
        "L": "Hilfsspalte: Name ohne Leerzeichen für den Abgleich.",
        "M": "Hilfsspalte: Fundstelle des Namens in Tabelle Y.",
    }
    head_bg = {"in": HEAD_IN, "calc": HEAD_CALC, "help": HEAD_HELP}
    for col, title, _, kind in X_COLS:
        c = put(ws, f"{col}{X_HDR}", title, font(10, True, "FFFFFF"), head_bg[kind], CENTER_WRAP)
        if col in notes:
            c.comment = Comment(notes[col], "UVV")
            c.comment.width, c.comment.height = 260, 90
    ws.row_dimensions[X_HDR].height = 32

    # Daten (Konstanten)
    for i, p in enumerate(rows):
        r = X_FIRST + i
        ws[f"A{r}"] = p["nr"]
        ws[f"B{r}"] = p["name"]
        ws[f"C{r}"] = p["abt"]
        ws[f"E{r}"] = p["uvv"]
        ws[f"I{r}"] = p["bez"]
        ws[f"J{r}"] = p["akte"]
        ws[f"K{r}"] = p["bem"]
    # Formeln bis X_LAST, damit die Tabelle fortgeschrieben werden kann
    for r in range(X_FIRST, X_LAST + 1):
        for col, f in x_formulas(r).items():
            ws[f"{col}{r}"] = f
        if r > last_data:
            ws[f"B{r}"] = auto_name_formula(r)
        for col, _, _, kind in X_COLS:
            c = ws[f"{col}{r}"]
            c.font = font(10, bold=(col == "H"), color=MUTED if kind == "help" else INK)
            c.alignment = Alignment(vertical="center")
        ws[f"A{r}"].number_format = "@"
        ws[f"E{r}"].number_format = DATE
        ws[f"F{r}"].number_format = DATE
        ws[f"G{r}"].number_format = "0"
        ws[f"E{r}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"F{r}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"G{r}"].alignment = Alignment(horizontal="center", vertical="center")

    tab = Table(displayName="tblUVV", ref=f"A{X_HDR}:M{last_data}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)

    # Bedingte Formatierung (Reihenfolge = Priorität)
    area = f"A{X_FIRST}:K{X_LAST}"
    h = f"$H{X_FIRST}"
    ws.conditional_formatting.add(
        f"B{X_FIRST}:B{X_LAST}",
        FormulaRule(formula=[f'AND($L{X_FIRST}<>"",COUNTIF($L${X_FIRST}:$L${X_LAST},$L{X_FIRST})>1)'],
                    font=Font(bold=True, color=DUP_INK)))
    ws.conditional_formatting.add(
        f"A{X_FIRST}:A{X_LAST}",
        FormulaRule(formula=[f'AND($A{X_FIRST}<>"",COUNTIF($A${X_FIRST}:$A${X_LAST},$A{X_FIRST})>1)'],
                    font=Font(bold=True, color=DUP_INK)))
    rules = [
        (f'{h}="{ST_ABGANG}"', PURPLE, None),
        (f'{h}="{ST_ABGELAUFEN}"', RED, None),
        (f'{h}="{ST_BALD}"', AMBER, None),
        (f'{h}="{ST_NEU}"', BLUE_STRONG, Font(italic=True)),
        (f'OR({h}="{ST_FEHLT}",{h}="{ST_PRUEFEN}")', BLUE, None),
        (f'{h}="{ST_ENTFAELLT}"', GREY, Font(color=MUTED)),
    ]
    for formula, color, fnt in rules:
        ws.conditional_formatting.add(
            area, FormulaRule(formula=[formula], fill=fill(color), font=fnt, stopIfTrue=True))
    ws.conditional_formatting.add(
        f"H{X_FIRST}:H{X_LAST}", FormulaRule(formula=[f'{h}="{ST_OK}"'], font=Font(color=GREEN_INK)))

    ws.freeze_panes = f"A{X_FIRST}"
    ws.print_title_rows = f"{X_HDR}:{X_HDR}"
    fit_width(ws)
    # Druckbereich wächst mit: Kopf + alle Zeilen mit Namen, Spalten A-H (sonst druckt Excel bis Zeile 2500)
    area = f'OFFSET({Q}{X}{Q}!$A$1,0,0,COUNTIF({Q}{X}{Q}!$B${X_FIRST}:$B${X_LAST},"?*")+{X_HDR},8)'
    ws.defined_names["_xlnm.Print_Area"] = DefinedName("_xlnm.Print_Area", localSheetId=0, attr_text=area)
    return last_data


# ---------------------------------------------------------------- Tabelle Y
def build_y(ws, zw):
    ws.sheet_properties.tabColor = "B7791F"
    ws.sheet_view.showGridLines = False
    for col, width in (("A", 32), ("B", 26), ("C", 34), ("D", 24), ("E", 3), ("F", 44), ("G", 34)):
        ws.column_dimensions[col].width = width
    ws.column_dimensions["D"].hidden = True

    put(ws, "A1", "Tabelle Y · Abgleich mit ZW-Daten (Stichtag)", font(16, True))
    put(ws, "A2", "Namen aus „ZW-Daten (Stichtag)“ ab Zeile 5 einfügen (als Werte). "
                  "Der Abgleich mit Tabelle X läuft nur über den Namen.", font(9, color=MUTED))
    put(ws, "A3", "Spalte A ← ZW-Daten Spalte A oder BO („Name komplett“)   ·   "
                  "Spalte B ← ZW-Daten Spalte BP („FK“)", font(9, color=MUTED))
    ws.row_dimensions[1].height = 26
    for col, title, bg in (("A", "Name komplett", HEAD_IN), ("B", "Führungskraft (FK)", HEAD_IN),
                           ("C", "Abgleich mit Tabelle X", HEAD_CALC), ("D", "Suchschlüssel", HEAD_HELP)):
        put(ws, f"{col}{Y_HDR}", title, font(10, True, "FFFFFF"), bg, CENTER_WRAP)
    ws.row_dimensions[Y_HDR].height = 30

    for i, p in enumerate(zw):
        r = Y_FIRST + i
        ws[f"A{r}"] = p["name"]
        ws[f"B{r}"] = p["fk"]
    for r in range(Y_FIRST, Y_LAST + 1):
        ws[f"C{r}"] = (f'=IF(D{r}="","",IFERROR(IF(_xlfn.ISFORMULA(INDEX({XB},MATCH(D{r},{XL},0))),'
                       f'"Neu – unten in Tabelle X angehängt","in Tabelle X"),"fehlt in Tabelle X"))')
        ws[f"D{r}"] = "=" + key_formula(f"A{r}")
        for col in "AB":
            c = ws[f"{col}{r}"]
            c.font = font()
            c.fill = fill(PASTE_BG)
            c.alignment = Alignment(vertical="center")
        ws[f"C{r}"].font = font(10, color=MUTED)
        ws[f"D{r}"].font = font(10, color=MUTED)
    ws.conditional_formatting.add(
        f"A{Y_FIRST}:C{Y_LAST}",
        FormulaRule(formula=[f'LEFT($C{Y_FIRST},3)="Neu"'], fill=fill(BLUE_STRONG), font=Font(bold=True),
                    stopIfTrue=True))
    ws.conditional_formatting.add(
        f"A{Y_FIRST}:C{Y_LAST}",
        FormulaRule(formula=[f'$C{Y_FIRST}="fehlt in Tabelle X"'], fill=fill(AMBER), stopIfTrue=True))

    # Infobereich (Kopf in Zeile 4 wie die Liste, darunter gleich hohe Zeilen)
    lab = font(10, color=INK)
    put(ws, "F1", "Stand der eingefügten ZW-Daten", lab, al=LEFT)
    put(ws, "G1", None, font(10, True), INPUT_BG, LEFT, DATE, box())
    ws["G1"].comment = Comment("Optional: Datum (Stichtag) der eingefügten ZW-Daten. "
                               "Wird in Tabelle X oben rechts angezeigt.", "UVV")
    put(ws, f"F{Y_HDR}", "Auswertung", font(10, True, "FFFFFF"), HEAD_CALC, LEFT)
    put(ws, f"G{Y_HDR}", "Anzahl", font(10, True, "FFFFFF"), HEAD_CALC, CENTER_WRAP)
    kpis = [
        ("Namen in Tabelle Y", f'=COUNTIF($D${Y_FIRST}:$D${Y_LAST},"?*")'),
        ("davon schon in Tabelle X", f'=COUNTIF($C${Y_FIRST}:$C${Y_LAST},"in Tabelle X")'),
        ("Zugänge: unten in Tabelle X angehängt", f'=COUNTIF($C${Y_FIRST}:$C${Y_LAST},"Neu*")'),
        ("Abgänge: in Tabelle X lila markiert", f'=COUNTIF({XH},"{ST_ABGANG}")'),
        ("Namen ohne Führungskraft",
         f'=SUMPRODUCT(($D${Y_FIRST}:$D${Y_LAST}<>"")*($B${Y_FIRST}:$B${Y_LAST}=""))'),
    ]
    for i, (label, formula) in enumerate(kpis):
        r = Y_FIRST + i
        put(ws, f"F{r}", label, lab, NEUTRAL, LEFT, border=box())
        put(ws, f"G{r}", formula, font(10, True), NEUTRAL, Alignment(horizontal="center", vertical="center"),
            "0", box())
    r = Y_FIRST + len(kpis)
    put(ws, f"F{r}", f'=IF(SUMPRODUCT(($A${Y_FIRST}:$A${Y_LAST}="")*($B${Y_FIRST}:$B${Y_LAST}<>""))>0,'
                     f'"⚠ In Spalte B stehen FK ohne Namen. Spalte B ist vermutlich veraltet: A und B zusammen neu einfügen.",'
                     f'"")', font(9, True, DUP_INK), al=LEFT)

    top = r + 2                         # Abgangsliste
    put(ws, f"F{top}", "Abgänge: in Tabelle X, aber nicht mehr in Tabelle Y", font(10, True))
    put(ws, f"F{top + 1}", "Name", font(10, True, "FFFFFF"), "7E57C2", LEFT)
    put(ws, f"G{top + 1}", "Abteilung", font(10, True, "FFFFFF"), "7E57C2", LEFT)
    for k in range(1, ABGANG_LISTE + 1):
        r = top + 1 + k
        pos = (f'SMALL(IF({XH}="{ST_ABGANG}",ROW({XH})-ROW({Q}{X}{Q}!$H${X_FIRST})+1),{k})')
        ws[f"F{r}"] = ArrayFormula(f"F{r}", f'=IFERROR(INDEX({XB},{pos}),"")')
        ws[f"G{r}"] = ArrayFormula(f"G{r}", f'=IFERROR(INDEX({XC},{pos})&"","")')
        for col in "FG":
            ws[f"{col}{r}"].font = font(10)
            ws[f"{col}{r}"].fill = fill(PURPLE)
    first = top + 2
    ws.conditional_formatting.add(
        f"F{first}:G{first + ABGANG_LISTE - 1}",
        FormulaRule(formula=[f'$F{first}=""'], fill=fill("FFFFFF"), stopIfTrue=True))

    ws.freeze_panes = f"A{Y_FIRST}"
    fit_width(ws)
    ws.print_title_rows = f"{Y_HDR}:{Y_HDR}"


# ---------------------------------------------------------------- Anleitung
GUIDE = [
    ("h", "Anleitung · UVV-Überwachung"),
    ("t", COURSE),
    ("", ""),
    ("s", "Blätter"),
    ("b", "Tabelle X: Übersicht aller Mitarbeiter. Eingeben: Personennummer, Name, Abteilung, Letzte UVV "
          "(ausgeblendet: Bezeichnung, Personalakte, Bemerkung). Automatisch: Führungskraft, Gültig bis, "
          "Resttage, Status."),
    ("b", "Tabelle Y: Namen (und FK) aus „ZW-Daten (Stichtag)“. Daraus kommen der Abgleich auf Zu- und "
          "Abgänge und die Führungskraft in Tabelle X."),
    ("", ""),
    ("s", "Farben in Tabelle X"),
    ("c", RED, "Rot – abgelaufen: Letzte UVV ist älter als 364 Tage (Gültig bis liegt vor heute)."),
    ("c", AMBER, "Gelb – bald fällig: läuft in den nächsten 30 Tagen ab (Wert oben rechts in Tabelle X)."),
    ("c", PURPLE, "Lila – nicht mehr in ZW-Daten: Der Name steht nicht mehr in Tabelle Y (Abgang)."),
    ("c", BLUE_STRONG, "Blau (kursiv) – neu aus Tabelle Y: Name ist noch eine Formel und muss festgeschrieben werden."),
    ("c", BLUE, "Hellblau – UVV-Datum fehlt oder ist kein gültiges Datum."),
    ("c", GREY, "Grau – UVV entfällt: In „Letzte UVV“ steht „entfällt“."),
    ("b", "Fette rote Schrift bei Name oder Personennummer: doppelt eingetragen."),
    ("", ""),
    ("s", "Neue ZW-Daten einspielen (Zu- und Abgänge prüfen)"),
    ("n", "1.", "Tabelle Y: alte Einträge ab A5 in Spalte A und B markieren und mit Entf löschen."),
    ("n", "2.", "In „ZW-Daten (Stichtag)“ die Namen ab Zeile 5 kopieren: am besten Spalte BO und BP zusammen "
                "(„Name komplett“ und „FK“ liegen nebeneinander). Nur Spalte A geht auch, dann fehlt die FK."),
    ("n", "3.", "In Tabelle Y die Zelle A5 anklicken: Rechtsklick > Einfügeoptionen > Werte (123)."),
    ("n", "4.", "Optional: In Tabelle Y oben rechts (G1) das Datum der ZW-Daten eintragen."),
    ("n", "5.", "Tabelle X: Abgänge sind jetzt lila. Neue Mitarbeiter stehen automatisch unter der Tabelle "
                "(blau, Status „Neu – Name festschreiben“)."),
    ("n", "6.", "Neue Namen festschreiben: die blauen Namen in Spalte B markieren, Strg+C, Rechtsklick > Werte (123). "
                "Danach Personennummer, Abteilung und Letzte UVV von Hand ergänzen."),
    ("b", "Warum festschreiben? Bis dahin ist der Name eine Formel. Beim nächsten Einfügen in Tabelle Y könnte "
          "sich die Reihenfolge ändern, und die von Hand eingetragenen Daten stünden beim falschen Namen."),
    ("", ""),
    ("s", "Mitarbeiter von Hand hinzufügen"),
    ("b", "In die erste freie Zeile unter der Tabelle schreiben. Führungskraft, Gültig bis, Resttage und Status "
          "rechnen sofort. Formeln und Farben sind bis Zeile 2500 vorbereitet."),
    ("b", "Mitarbeiter ohne UVV-Pflicht nicht löschen, sondern „entfällt“ in „Letzte UVV“ eintragen. "
          "Sonst hängt Tabelle X den Namen aus Tabelle Y wieder unten an."),
    ("", ""),
    ("s", "Hinweise"),
    ("b", "Der Abgleich läuft nur über den Namen. Leerzeichen und Groß-/Kleinschreibung spielen keine Rolle "
          "(„Heß, Astrid“ = „Heß,Astrid“). Namensänderungen erscheinen als Abgang (alt) und Zugang (neu)."),
    ("b", "Sortieren und Filtern über die Pfeile in der Kopfzeile von Tabelle X. Vorher neue Namen festschreiben."),
    ("b", "Ausgeblendete Spalten (Bezeichnung, Personalakte, Bemerkung) über das „+“ über Spalte H/I einblenden."),
    ("b", "Gültigkeit (364 Tage) und Vorwarnung (30 Tage) stehen oben rechts in Tabelle X und können geändert werden."),
    ("b", "Die Datei enthält keine Verknüpfung zu anderen Dateien. Die Führungskraft kommt aus Tabelle Y."),
]


def build_guide(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 110
    r = 1
    for item in GUIDE:
        kind = item[0]
        if kind == "h":
            put(ws, f"A{r}", item[1], font(16, True))
            ws.row_dimensions[r].height = 26
        elif kind == "t":
            put(ws, f"A{r}", item[1], font(9, color=MUTED))
        elif kind == "s":
            put(ws, f"A{r}", item[1], font(12, True, HEAD_IN))
            ws.row_dimensions[r].height = 20
        elif kind == "b":
            put(ws, f"A{r}", "•", font(10, color=MUTED), al=Alignment(horizontal="center", vertical="top"))
            put(ws, f"B{r}", item[1], font(10), al=LEFT_WRAP)
        elif kind == "n":
            put(ws, f"A{r}", item[1], font(10, True), al=Alignment(horizontal="center", vertical="top"))
            put(ws, f"B{r}", item[2], font(10), al=LEFT_WRAP)
        elif kind == "c":
            put(ws, f"A{r}", None, fl=item[1], border=box())
            put(ws, f"B{r}", item[2], font(10), al=LEFT_WRAP)
        if kind in "bnc" and len(item[-1]) > 115:
            ws.row_dimensions[r].height = 27
        r += 1
    fit_width(ws, landscape=False)


# ---------------------------------------------------------------- Werte speichern
def inject_values(path, calc_path, out_path, sheet_names):
    """Speichert die von LibreOffice berechneten Ergebnisse in den Formelzellen (auch Datumswerte),
    damit die Datei auch in der geschützten Ansicht Zahlen zeigt. Excel rechnet beim Öffnen neu."""
    values = load_workbook(calc_path, data_only=True)
    zin = zipfile.ZipFile(path)
    parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    order = [i.filename for i in zin.infolist()]
    wb = etree.fromstring(parts["xl/workbook.xml"])
    rels = {r.get("Id"): r.get("Target") for r in etree.fromstring(parts["xl/_rels/workbook.xml.rels"])}
    n = 0
    for sh in wb.find(N + "sheets"):
        name = sh.get("name")
        if name not in sheet_names:
            continue
        target = rels[sh.get("{%s}id" % RNS)]
        part = target.lstrip("/") if target.startswith("/") else "xl/" + target
        ws = values[name]
        root = etree.fromstring(parts[part])
        for c in root.iter(N + "c"):
            if c.find(N + "f") is None:
                continue
            v = ws[c.get("r")].value
            if isinstance(v, (dt.datetime, dt.date)):
                v = to_excel(v)
            set_cell_cache(c, v)
            n += 1
        parts[part] = dump(root)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in order:
            zout.writestr(name, parts[name])
    return n


# ---------------------------------------------------------------- Prüfung
def expected_status(p, in_zw, today):
    if not in_zw:
        return ST_ABGANG
    u = p["uvv"]
    if u is None:
        return ST_FEHLT
    if not isinstance(u, (dt.datetime, dt.date)):
        return ST_ENTFAELLT if str(u).strip().lower() == "entfällt" else ST_PRUEFEN
    rest = (u + dt.timedelta(days=GUELTIG_TAGE) - dt.datetime.combine(today, dt.time())).days
    if rest < 0:
        return ST_ABGELAUFEN
    return ST_BALD if rest <= VORWARNUNG_TAGE else ST_OK


def check(calc_path, rows, zw):
    wb = load_workbook(calc_path, data_only=True)
    x, y = wb[X], wb[Y]
    errors = [f"{ws.title}!{c.coordinate}" for ws in (x, y) for row in ws.iter_rows() for c in row
              if isinstance(c.value, str) and c.value in ERRORS]
    if errors:
        raise RuntimeError(f"Formelfehler: {errors[:20]}")
    today = x["A4"].value.date() if isinstance(x["A4"].value, dt.datetime) else dt.date.today()
    zw_fk = {name_key(p["name"]): p["fk"] for p in zw}
    bad = []
    for i, p in enumerate(rows):
        r = X_FIRST + i
        k = name_key(p["name"])
        want = expected_status(p, k in zw_fk, today)
        got = x[f"H{r}"].value
        fk = x[f"D{r}"].value or None
        if got != want or fk != (zw_fk.get(k) or None):
            bad.append((r, p["name"], got, want, fk, zw_fk.get(k)))
    if bad:
        raise RuntimeError(f"Abweichungen in Tabelle X: {bad[:10]}")
    auto = [x[f"B{r}"].value for r in range(X_FIRST + len(rows), X_LAST + 1) if x[f"B{r}"].value]
    if auto:
        raise RuntimeError(f"Unerwartete Zugänge unter der Tabelle: {auto[:10]}")
    counts = {}
    for i in range(len(rows)):
        s = x[f"H{X_FIRST + i}"].value
        counts[s] = counts.get(s, 0) + 1
    return counts


def main(src, out):
    people, neu, zw = read_source(src)
    rows = people + neu
    wb = Workbook()
    ws_x = wb.active
    ws_x.title = X
    build_x(ws_x, rows)
    build_y(wb.create_sheet(Y), zw)
    build_guide(wb.create_sheet("Anleitung"))
    wb.calculation.fullCalcOnLoad = True
    with tempfile.TemporaryDirectory(prefix="uvv-") as tmp:
        raw = os.path.join(tmp, "uvv.xlsx")
        calc = os.path.join(tmp, "calc.xlsx")
        wb.save(raw)
        shutil.copy(raw, calc)
        libreoffice_recalc(calc)
        counts = check(calc, rows, zw)
        cells = inject_values(raw, calc, out, (X, Y))
    print(f"{len(people)} Mitarbeiter aus {SRC_SHEET}, {len(neu)} aus den ZW-Daten angehängt, "
          f"{len(zw)} Namen in Tabelle Y, {cells} Formelzellen -> {out}")
    for status, n in sorted(counts.items(), key=lambda kv: -kv[1]):
        print(f"  {status}: {n}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
