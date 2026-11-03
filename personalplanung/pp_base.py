"""Gemeinsame Konstanten, Farben und Hilfsfunktionen fuer den Generator der
Personal- und Einsatzplanung 2027.

Alle fachlichen Vorgaben (Teams, Rotationsfolge, Funktionen, Unterrichte,
Abwesenheitsarten, Farben) stehen hier zentral. Die Werte landen im Blatt
"Einstellungen" der erzeugten Arbeitsmappe und koennen dort weiter gepflegt werden.
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ------------------------------------------------------------------ Groessen
YEAR = 2027
NMAX = 60          # Plaetze im Planungsraster (Mitarbeiter)
NDAYS = 366        # Tagesspalten/-zeilen (Schaltjahr-faehig)
CAP_MA = 60        # Zeilen Mitarbeiterstamm
CAP_FM = 60        # Zeilen Funktionsmatrix
CAP_Q = 500        # Zeilen Qualifikationen
CAP_ABW = 500      # Zeilen Abwesenheiten
CAP_VERT = 200     # Zeilen Vertretungen
CAP_ZB = 300       # Zeilen Zeitbuchungen

# ------------------------------------------------------------------ Fachliche Listen
# Reihenfolge und Farben wie im Bereitschaftskalender 2027
TEAMS = [("Blau", "BL", "4F81BD"), ("Rot", "RO", "FF4B4B"), ("Gelb", "GE", "FFFF00"),
         ("Grün", "GR", "92D050"), ("Orange", "OR", "F79646")]
TEAM_NAMES = [t[0] for t in TEAMS]
TEAM_COLOR = {t[0]: t[2] for t in TEAMS}

# Feste Rotationsfolge ab Rotationsbeginn (darf nicht automatisch veraendert werden)
ROTATION = [("Fr-So", "Rot"), ("Mo-Do", "Gelb"), ("Fr-So", "Grün"), ("Mo-Do", "Orange"),
            ("Fr-So", "Blau"), ("Mo-Do", "Rot"), ("Fr-So", "Gelb"), ("Mo-Do", "Grün"),
            ("Fr-So", "Orange"), ("Mo-Do", "Blau")]

FUNKTIONEN = ["Meister", "AV", "TVT-Fahrer", "Handwerker"]
TEAMFUNKTIONEN = FUNKTIONEN + ["Reserve"]
FKT_KURZ = {"Meister": "M", "AV": "AV", "TVT-Fahrer": "TVT", "Handwerker": "HW", "Reserve": "RES"}
SOLL = {"Meister": 1, "AV": 1, "TVT-Fahrer": 1, "Handwerker": 3}

# Unterrichte/Qualifikationen in der Reihenfolge der Funktionsmatrix,
# Gueltigkeit in Monaten (Annahme, im Blatt Einstellungen anpassbar)
UNTERRICHTE = [("Ersthelfer", 24), ("SeSi", 12), ("Sakra", 12), ("Sakra & SiÜ", 12),
               ("Gefahrgut", 36), ("Simulator", 12), ("BB MISS", 12), ("Bühnenbediener", 12),
               ("Betrieb", 12), ("S&R", 12), ("WBT", 12), ("FDL-Einweisung", 12),
               ("Kabeleinweisung", 12)]
U_NAMES = [u[0] for u in UNTERRICHTE]

# Pflichtqualifikationen je Funktion (Annahme, im Blatt Einstellungen anpassbar)
PFLICHT = {
    "Meister": ["Sakra & SiÜ", "S&R", "Betrieb", "FDL-Einweisung"],
    "AV": ["Sakra", "S&R", "Betrieb", "FDL-Einweisung"],
    "TVT-Fahrer": ["Simulator", "Betrieb"],
    "Handwerker": ["SeSi", "Betrieb"],
}

# Abwesenheitsarten (Dropdown) -> Code
ABW_ARTEN = [("Urlaub", "U"), ("Krankheit", "K"), ("JAZ", "JAZ"), ("Unterricht", "UR"),
             ("Fortbildung", "FB"), ("Freistellung", "FS"), ("Sonderurlaub", "SU")]

# Status-Codes der Abwesenheiten, aufsteigende Prioritaet (hoechste gewinnt bei Ueberschneidung)
# (Code, Bedeutung, Fuellfarbe, Schriftfarbe, fett)
ABW_CODES = [
    ("FS", "Freistellung (manuell)", "BFBFBF", "000000", False),
    ("SU", "Sonderurlaub", "C6E0B4", "000000", False),
    ("JAZ", "JAZ (Freizeit aus JAZ-Konto)", "C4D79B", "000000", False),
    ("U", "Urlaub", "00B050", "000000", False),
    ("FB", "Fortbildung", "156082", "FFFFFF", False),
    ("UR", "Unterricht", "FF0000", "FFFFFF", False),
    ("FDL", "Unterricht FDL-Einweisung", "FFC000", "000000", False),
    ("KAB", "Unterricht Kabeleinweisung", "7030A0", "FFFFFF", False),
    ("K", "Krankheit", "FFFF00", "FF0000", True),
]
# Dienst-Codes (ergeben sich automatisch)
DUTY_CODES = [
    ("T", "Tagdienst 06:45–15:30", None, "000000", False),
    ("N", "Nachtdienst / Bereitschaft 15:30–06:45 (Arbeiten bei Nacht)", "FFFF00", "000000", True),
    ("V", "Vertretung im Nachtdienst (bestätigt)", "FFFF00", "7F6000", True),
    ("F", "Freistellung nach Nachtdienst", "D9D9D9", "000000", False),
    ("–", "nicht aktiv / noch nicht eingetreten", None, "A6A6A6", False),
]

# Schulferien Hessen 2027 (graue Markierung im Bereitschaftskalender 2027)
SCHULFERIEN = [("Weihnachtsferien 2026/27", (2027, 1, 1), (2027, 1, 12)),
               ("Osterferien", (2027, 3, 22), (2027, 4, 2)),
               ("Sommerferien", (2027, 6, 28), (2027, 8, 6)),
               ("Herbstferien", (2027, 10, 4), (2027, 10, 15)),
               ("Weihnachtsferien 2027/28", (2027, 12, 23), (2027, 12, 31))]

MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August",
          "September", "Oktober", "November", "Dezember"]
MON3 = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]

# ------------------------------------------------------------------ Farben
FONT = "Arial"
HEAD = "B7DEE8"        # Kopfzeilen PEP
HEAD_DARK = "31869B"
INPUT_HEAD = "FFE699"  # Spaltenkopf Eingabe
INPUT_BG = "FFF2CC"    # Eingabezelle
WEEKEND = "F2AA84"     # Wochenende im PEP (Akzent 2, heller 40 %)
SAT = "FFFFCC"
SUN = "FFCC99"
HOLIDAY = "FFD9D9"
HOLIDAY_FONT = "CC0000"
FERIEN = "BFBFBF"
GREY = "D9D9D9"
LIGHT = "F2F2F2"
INK = "000000"
MUTED = "7F7F7F"
GREEN, YELLOW, RED, DARKRED = "92D050", "FFFF00", "FF0000", "C00000"
AMPEL = {"Grün": (GREEN, "000000"), "Gelb": (YELLOW, "000000"), "Rot": (RED, "FFFFFF")}

DATE_FMT = "DD.MM.YYYY"

# ------------------------------------------------------------------ Stile
_side = Side(style="thin", color="A6A6A6")
_med = Side(style="medium", color="000000")
THIN = Border(left=_side, right=_side, top=_side, bottom=_side)
MEDIUM = Border(left=_med, right=_med, top=_med, bottom=_med)


def font(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def align(h=None, v="center", wrap=False, rot=None):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, text_rotation=rot or 0)


def L(c):
    return get_column_letter(c)


def ref(ws, a1):
    """Absoluter Bezug inkl. Blattname: ref(ws,'A1:B5') -> 'Blatt'!$A$1:$B$5"""
    parts = []
    for p in a1.split(":"):
        col = "".join(ch for ch in p if ch.isalpha())
        row = "".join(ch for ch in p if ch.isdigit())
        parts.append(f"${col}${row}" if row else f"${col}")
    return f"{quote_sheetname(ws.title)}!{':'.join(parts)}"


def style(cell, size=10, bold=False, color=INK, bg=None, h=None, v="center", wrap=False,
          border=None, fmt=None, italic=False):
    cell.font = font(size, bold, color, italic)
    cell.alignment = align(h, v, wrap)
    if bg:
        cell.fill = fill(bg)
    if border:
        cell.border = border
    if fmt:
        cell.number_format = fmt
    return cell


def put(ws, a1, value, **kw):
    c = ws[a1]
    c.value = value
    style(c, **kw)
    return c


class Book:
    """Wrapper um die Arbeitsmappe: definierte Namen, Tabellen, Dropdowns, bedingte Formate."""

    def __init__(self, wb):
        self.wb = wb
        self.names = {}

    def name(self, name, target):
        if name in self.names:
            raise ValueError(f"Name doppelt: {name}")
        self.names[name] = target
        self.wb.defined_names[name] = DefinedName(name, attr_text=target)

    def name_ws(self, name, ws, a1):
        self.name(name, ref(ws, a1))

    @staticmethod
    def table(ws, name, a1, style_name="TableStyleMedium6", stripes=True):
        t = Table(displayName=name, ref=a1)
        t.tableStyleInfo = TableStyleInfo(name=style_name, showRowStripes=stripes,
                                          showColumnStripes=False, showFirstColumn=False,
                                          showLastColumn=False)
        ws.add_table(t)
        return t

    @staticmethod
    def dv_list(ws, sqref, source, allow_blank=True, error=True, prompt=None):
        dv = DataValidation(type="list", formula1=source, allow_blank=allow_blank)
        dv.showErrorMessage = error
        dv.errorStyle = "stop" if error else "warning"
        dv.errorTitle = "Ungültiger Wert"
        dv.error = "Bitte einen Wert aus der Liste wählen."
        if prompt:
            dv.promptTitle = "Hinweis"
            dv.prompt = prompt
            dv.showInputMessage = True
        ws.add_data_validation(dv)
        for part in sqref.split():
            dv.add(part)
        return dv

    @staticmethod
    def dv_date(ws, sqref):
        dv = DataValidation(type="date", operator="between", formula1="DATE(2000,1,1)",
                            formula2="DATE(2100,12,31)", allow_blank=True)
        dv.showErrorMessage = True
        dv.errorTitle = "Datum"
        dv.error = "Bitte ein Datum im Format TT.MM.JJJJ eingeben."
        ws.add_data_validation(dv)
        for part in sqref.split():
            dv.add(part)
        return dv

    @staticmethod
    def dv_number(ws, sqref, lo, hi, decimal=True):
        dv = DataValidation(type="decimal" if decimal else "whole", operator="between",
                            formula1=str(lo), formula2=str(hi), allow_blank=True)
        dv.showErrorMessage = True
        dv.errorTitle = "Zahl"
        dv.error = f"Bitte eine Zahl zwischen {lo} und {hi} eingeben."
        ws.add_data_validation(dv)
        for part in sqref.split():
            dv.add(part)
        return dv

    @staticmethod
    def cf(ws, sqref, formula, bg=None, color=None, bold=None, italic=None, stop=False,
           border_left=False):
        f = Font(color=color, bold=bold, italic=italic) if (color or bold or italic) else None
        fl = PatternFill(start_color=bg, end_color=bg, fill_type="solid", bgColor=bg) if bg else None
        b = Border(left=Side(style="medium", color="000000")) if border_left else None
        dxf = DifferentialStyle(font=f, fill=fl, border=b)
        rule = FormulaRule(formula=[formula], stopIfTrue=stop)
        rule.dxf = dxf
        ws.conditional_formatting.add(sqref, rule)
        return rule


def col_widths(ws, widths):
    for k, w in widths.items():
        ws.column_dimensions[k].width = w


def title(ws, text, sub=None, width_to="H"):
    put(ws, "A1", text, size=16, bold=True)
    ws.row_dimensions[1].height = 26
    if sub:
        put(ws, "A2", sub, size=9, color=MUTED, italic=True)


def header_row(ws, row, col, labels, inputs=(), height=30, size=9):
    """Schreibt Spaltenkoepfe; Eingabespalten (Index in labels) gelb hinterlegt."""
    for i, lab in enumerate(labels):
        c = ws.cell(row, col + i, lab)
        is_in = i in inputs
        style(c, size=size, bold=True, color=INK if is_in else "FFFFFF",
              bg=INPUT_HEAD if is_in else HEAD_DARK, h="center", wrap=True, border=THIN)
    ws.row_dimensions[row].height = height


def landscape(ws, paper="A4", fit_width=True, fit_height=False, area=None, rows=None, cols=None):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3 if paper == "A3" else ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1 if fit_width else 0
    ws.page_setup.fitToHeight = 1 if fit_height else 0
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_options.horizontalCentered = True
    if area:
        ws.print_area = area
    if rows:
        ws.print_title_rows = rows
    if cols:
        ws.print_title_cols = cols
    ws.oddFooter.left.text = "&A"
    ws.oddFooter.left.size = 8
    ws.oddFooter.right.text = "Seite &P von &N – Stand &D"
    ws.oddFooter.right.size = 8


def weekday_formula(cell):
    return f'CHOOSE(WEEKDAY({cell},2),"Mo","Di","Mi","Do","Fr","Sa","So")'


def team_cf(book, ws, sqref, cell, by_fill=True):
    """Teamfarbe als Fuellung, wenn die Zelle (bzw. Bezugszelle) den Teamnamen enthaelt."""
    for t, _k, col in TEAMS:
        book.cf(ws, sqref, f'{cell}="{t}"', bg=col, color="000000")


def ampel_cf(book, ws, sqref, cell):
    for k, (bg, fc) in AMPEL.items():
        book.cf(ws, sqref, f'{cell}="{k}"', bg=bg, color=fc, bold=True)
