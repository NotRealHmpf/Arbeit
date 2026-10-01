"""Erzeugt die Dashboard-Blaetter als eigenstaendige Arbeitsmappe:

    Uebersicht, Soll-Ist, Grafik <Bezirk> (je Bezirk), Daten, Zielzustand, Einstellungen

Die Formeln verweisen auf die Bezirks-Blaetter der Original-Datei (Bedarf_Bestand_LST);
merge_into_original.py setzt die Blaetter anschliessend dort ein. Das Blatt Zielzustand
ist eine Kopie aus LST_Zielzustand_gesamt.xlsx und dient als Soll.
"""
import math
import re
from copy import copy
from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import DataPoint, Series, SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter, quote_sheetname
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break, RowBreak
from openpyxl.worksheet.properties import PageSetupProperties


@dataclass
class District:
    sheet: str      # Blattname in Bedarf_Bestand_LST
    zz_name: str    # Name des Bezirks im Zielzustand (Zeile 1)


# ---------------------------------------------------------------- Konstanten
SRC_FIRST_ROW = 2          # erste ausgewertete Zeile je Bezirks-Blatt (Kopfzeile wird erkannt und uebersprungen)
SRC_ROWS = 80              # Anzahl ausgewerteter Zeilen je Bezirks-Blatt (2..81)
LIST_ROWS = 30             # Zeilen der Mitarbeiterliste je Bezirksseite (passt auf eine Druckseite)
N_YEARS = 7                # Verlauf: Heute + Bezugsjahr .. Bezugsjahr+6
BEYOND_DEFAULT = "Nein"      # Planjahre ueber die Ziel-Qualifikation hinaus zaehlen (Einstellungen!C5)

# Spalten in den Bezirks-Blaettern (in allen zwoelf Blaettern gleich)
SRC_NAME, SRC_FIRST, SRC_ZIEL, SRC_IST = "B", "C", "G", "J"
SRC_QUALI = ("L", "S")     # Azubi .. Teamleiter
SRC_VP = ("U", "V")        # oertl. Verwendungspruefung Wmech / SigMech

FONT = "Arial"
INK = "0B0B0B"
INK2 = "52514E"
MUTED = "898781"
HAIR = "E1E0D9"
TILE_BG = "F4F4F2"
HEAD_BG = "EDEDEA"
INPUT_BG = "FFF2CC"
GOOD_BG, GOOD_INK = "E2F3E2", "0A6B0A"
BAD_BG, BAD_INK = "F9E0E0", "A12A2A"

# Farben Soll-Ist
C_SOLL = INK2
C_HEUTE = "86B6EF"
C_PLAN = "2A78D6"

# Status-Nr -> (Farbe, Textfarbe auf der Farbe). Bezeichnungen stehen im Blatt Einstellungen.
STATUS = [
    (1, "0CA30C", INK),       # Fertig
    (2, "104281", "FFFFFF"),  # Abschluss Bezugsjahr
    (3, "1C5CAB", "FFFFFF"),  # Abschluss Bezugsjahr+1
    (4, "2A78D6", "FFFFFF"),  # +2
    (5, "5598E7", INK),       # +3
    (6, "86B6EF", INK),       # ab +4
    (7, "EC835A", INK),       # Ueberfaellig
    (8, "D03B3B", "FFFFFF"),  # Fehlt
    (9, "898781", INK),       # Ziel nicht eingetragen
]
STATUS_TEXT = {
    1: ("Fertig", "Ziel-Qualifikation erreicht: „x“ in der Spalte der Ziel-Qualifikation (L–S) oder – bei Wmech/SigMech – „x“ bei der "
                  "örtlichen Verwendungsprüfung (U/V). Ebenfalls fertig, wenn Ist- und Ziel-Qualifikation gleich sind."),
    2: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr."),
    3: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr + 1."),
    4: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr + 2."),
    5: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr + 3."),
    6: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr + 4 oder später."),
    7: ("Überfällig", "Geplantes Jahr liegt vor dem Bezugsjahr, aber es steht noch kein „x“ – bitte prüfen/aktualisieren."),
    8: ("Fehlt (nichts geplant)", "Ziel-Qualifikation nicht erreicht und keine Ausbildung geplant (weder „x“ noch Jahr in der Ziel-Spalte bzw. Verwendungsprüfung)."),
    9: ("Ziel nicht eingetragen", "In der Spalte Ziel-Qualifikation steht nichts."),
}

# Kategorien = Zeilenbeschriftungen im Zielzustand (Spalte A)
CATS = ["Arbeiter LST", "Wmech", "SigMech", "SigMech RBEG", "Teamleiter"]
OTHER = "Sonstige"
GROUPS = CATS + [OTHER]

# Zuordnung Text in Ziel-/Ist-Spalte -> Spalte L..S, Verwendungspruefung U/V, Kategorie
MAPPING = [
    ("Azubi", "L", "", OTHER),
    ("Azubi 2023", "L", "", OTHER),
    ("Azubi 2024", "L", "", OTHER),
    ("Azubi 2025", "L", "", OTHER),
    ("Azubi 2026", "L", "", OTHER),
    ("Azubi 2027", "L", "", OTHER),
    ("Arb LST", "M", "", "Arbeiter LST"),
    ("Arbeiter LST", "M", "", "Arbeiter LST"),
    ("Weichmech", "N", "U", "Wmech"),
    ("Weichenmechaniker", "N", "U", "Wmech"),
    ("Wmech", "N", "U", "Wmech"),
    ("Sigmech", "O", "V", "SigMech"),
    ("Signalmechaniker", "O", "V", "SigMech"),
    ("Sigmech RBEG", "P", "", "SigMech RBEG"),
    ("Signalmechaniker RBEG", "P", "", "SigMech RBEG"),
    ("Kennziffer 4", "Q", "", OTHER),
    ("IHK-Meister", "R", "", OTHER),
    ("Teamleiter", "S", "", "Teamleiter"),
    ("TL", "S", "", "Teamleiter"),
    ("Umschüler EBET", "", "", OTHER),
    ("Umschüler", "", "", OTHER),
    ("Quereinsteiger", "", "", OTHER),
    ("Senior Expert LST", "", "", OTHER),
]

# Einstellungen: feste Zellpositionen
E_YEAR = "Einstellungen!$C$4"
E_BEYOND = "Einstellungen!$C$5"           # Planjahre ueber die Ziel-Qualifikation hinaus zaehlen (Ja/Nein)
E_STAT_FIRST, E_STAT_LAST = 8, 16
E_BEZ_HDR, E_BEZ_FIRST, E_BEZ_LAST = 20, 21, 40
E_MAP_HDR, E_MAP_FIRST, E_MAP_LAST = 44, 45, 84
E_STAT_LABEL = f"Einstellungen!$C${E_STAT_FIRST}:$C${E_STAT_LAST}"
E_BEZ_SHEET = f"Einstellungen!$B${E_BEZ_FIRST}:$B${E_BEZ_LAST}"
E_BEZ_ZZCOL = f"Einstellungen!$D${E_BEZ_FIRST}:$D${E_BEZ_LAST}"
E_MAP_TEXT = f"Einstellungen!$B${E_MAP_FIRST}:$B${E_MAP_LAST}"
E_MAP_COLNO = f"Einstellungen!$D${E_MAP_FIRST}:$D${E_MAP_LAST}"
E_MAP_VPNO = f"Einstellungen!$F${E_MAP_FIRST}:$F${E_MAP_LAST}"
E_MAP_CAT = f"Einstellungen!$G${E_MAP_FIRST}:$G${E_MAP_LAST}"
ZZ_AREA = "Zielzustand!$A$1:$ZZ$100"
ZZ_LABELS = "Zielzustand!$A$1:$A$100"


def stat_label_ref(code):
    return f"Einstellungen!$C${E_STAT_FIRST + code - 1}"


# ---------------------------------------------------------------- Stil-Helfer
def font(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


HAIR_SIDE = Side(style="thin", color=HAIR)
BOTTOM_HAIR = Border(bottom=HAIR_SIDE)
TOP_INK = Border(top=Side(style="thin", color=INK2))
LEFT = Alignment(horizontal="left", vertical="center")
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
NF_COUNT = '0;-0;"·"'
NF_DELTA = '+0;-0;"±0"'


DATA_LAST = 2              # letzte Zeile im Blatt Daten (wird in build() gesetzt)
_DATEN_COL = re.compile(r"Daten!\$([A-Z]+):\$\1(?![0-9$])")


def bound_daten(value):
    """Verweise auf ganze Spalten im Blatt Daten auf den Datenbereich begrenzen (rechnet deutlich schneller)."""
    if isinstance(value, str) and "Daten!$" in value:
        return _DATEN_COL.sub(lambda m: f"Daten!${m.group(1)}$2:${m.group(1)}${DATA_LAST}", value)
    return value


def put(ws, ref, value, f=None, fl=None, al=None, nf=None, border=None):
    c = ws[ref]
    c.value = bound_daten(value)
    c.font = f or font()
    if fl:
        c.fill = fl
    if al:
        c.alignment = al
    if nf:
        c.number_format = nf
    if border:
        c.border = border
    return c


def merge_put(ws, rng, value, **kw):
    first = rng.split(":")[0]
    c = put(ws, first, value, **kw)
    ws.merge_cells(rng)
    # Fuellung/Rahmen auf alle Zellen des Bereichs, damit es einheitlich aussieht
    if kw.get("fl") or kw.get("border"):
        for row in ws[rng]:
            for cell in row:
                if kw.get("fl"):
                    cell.fill = kw["fl"]
                if kw.get("border"):
                    cell.border = kw["border"]
    return c


def page_setup(ws, print_area, breaks=()):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_area = print_area
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    if breaks:
        ws.row_breaks = RowBreak()
        for r in breaks:
            ws.row_breaks.append(Break(id=r))


# ---------------------------------------------------------------- Diagramm-Helfer
def text_props(size=9, color=INK2, bold=False):
    cp = CharacterProperties(sz=int(size * 100), b=bold, solidFill=color,
                             latin=DFont(typeface=FONT), cs=DFont(typeface=FONT))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def chart_title(text, size=10):
    cp = CharacterProperties(sz=int(size * 100), b=True, solidFill=INK, latin=DFont(typeface=FONT), cs=DFont(typeface=FONT))
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def no_line():
    return LineProperties(noFill=True)


def data_labels(color, pos=None, numfmt=None, bold=False, size=9):
    dl = DataLabelList()
    dl.showVal = True
    dl.showLegendKey = False
    dl.showCatName = False
    dl.showSerName = False
    dl.showPercent = False
    dl.showBubbleSize = False
    if pos:
        dl.position = pos
    if numfmt:
        dl.numFmt = numfmt
    dl.txPr = text_props(size, color, bold)
    return dl


def style_chart_frame(chart):
    chart.graphical_properties = GraphicalProperties(ln=no_line())
    chart.roundedCorners = False


def cat_axis(ax, reverse=False):
    if reverse:
        ax.scaling.orientation = "maxMin"
    ax.delete = False
    ax.majorTickMark = "none"
    ax.txPr = text_props(9, INK2)
    ax.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="C3C2B7"))


def hide_value_axis(ax):
    ax.delete = True
    ax.majorGridlines = None
    ax.scaling.min = 0


def show_value_axis(ax, ymax=None, major=None):
    ax.delete = False
    ax.scaling.min = 0
    if ymax:
        ax.scaling.max = ymax
    if major:
        ax.majorUnit = major
    ax.numFmt = "0"
    ax.majorTickMark = "none"
    ax.txPr = text_props(8, MUTED)
    ax.graphicalProperties = GraphicalProperties(ln=no_line())
    ax.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=HAIR, w=6350)))


def ref_series(idx, values_ref, title_ref, cats_ref):
    """Serie aus Bereichs-Texten ('Blatt'!$A$1:$B$1): Werte, Titel-Zelle, Kategorien."""
    s = Series(idx=idx, order=idx)
    s.val = NumDataSource(numRef=NumRef(f=values_ref))
    s.tx = SeriesLabel(strRef=StrRef(f=title_ref))
    s.cat = AxDataSource(strRef=StrRef(f=cats_ref))
    return s


def fill_series(s, color):
    s.graphicalProperties = GraphicalProperties(solidFill=color, ln=no_line())
    s.invertIfNegative = False
    return s


def rng(ws_title, c1, r1, c2=None, r2=None):
    a = f"${c1}${r1}"
    b = f"${c2 or c1}${r2 or r1}"
    return f"{quote_sheetname(ws_title)}!{a}:{b}"


def status_bar_chart(ws, val_row, val_c1, val_c2, cat_row):
    """Ein Balken je Status-Kategorie, jeweils in der Status-Farbe."""
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "clustered"
    ch.style = 2
    ch.add_data(Reference(ws, min_col=val_c1, max_col=val_c2, min_row=val_row, max_row=val_row),
                from_rows=True, titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=val_c1, max_col=val_c2, min_row=cat_row, max_row=cat_row))
    s = ch.series[0]
    s.graphicalProperties = GraphicalProperties(solidFill=STATUS[0][1], ln=no_line())
    for i, (_, color, _) in enumerate(STATUS):
        pt = DataPoint(idx=i, invertIfNegative=False)
        pt.graphicalProperties = GraphicalProperties(solidFill=color, ln=no_line())
        s.dPt.append(pt)
    s.invertIfNegative = False
    s.dLbls = data_labels(INK, pos="outEnd", numfmt="0;-0;0", bold=True)
    ch.gapWidth = 35
    ch.legend = None
    cat_axis(ch.x_axis, reverse=True)
    hide_value_axis(ch.y_axis)
    style_chart_frame(ch)
    return ch


def stacked_status_chart(ws, hdr_row, first_row, last_row, cat_col, c1, c2):
    """Gestapelte Balken: je Kategorie (Zeile) die Anzahl je Status (Spalte)."""
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.style = 2
    ch.add_data(Reference(ws, min_col=c1, max_col=c2, min_row=hdr_row, max_row=last_row), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=cat_col, min_row=first_row, max_row=last_row))
    for s, (_, color, txt) in zip(ch.series, STATUS):
        s.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill="FFFFFF", w=12700))
        s.invertIfNegative = False
        s.dLbls = data_labels(txt, pos="ctr", numfmt="0;-0;;")   # Nullen ausblenden
    ch.gapWidth = 45
    ch.legend = Legend()
    ch.legend.position = "b"
    ch.legend.txPr = text_props(9, INK2)
    cat_axis(ch.x_axis, reverse=True)
    hide_value_axis(ch.y_axis)
    style_chart_frame(ch)
    return ch


def soll_ist_chart(cats_ref, series, horizontal=False, label_size=9):
    """Gruppierte Balken: Soll / Heute / Nach Plan. series: Liste (Werte-Bereich, Titel-Zelle, Farbe)."""
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.grouping = "clustered"
    ch.style = 2
    for i, (values_ref, title_ref, color) in enumerate(series):
        s = fill_series(ref_series(i, values_ref, title_ref, cats_ref), color)
        s.dLbls = data_labels(INK2, pos="outEnd", numfmt="0;-0;0", size=label_size)
        ch.series.append(s)
    ch.gapWidth = 60
    ch.overlap = -8
    ch.legend = Legend()
    ch.legend.position = "b"
    ch.legend.txPr = text_props(9, INK2)
    cat_axis(ch.x_axis, reverse=horizontal)
    hide_value_axis(ch.y_axis)
    style_chart_frame(ch)
    return ch


def composition_chart(cats_ref, series):
    """Gestapelte Saeulen: Zusammensetzung je Zeitpunkt. series: Liste (Werte, Titel-Zelle, Fuellfarbe, Textfarbe)."""
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.style = 2
    for i, (values_ref, title_ref, color, txt) in enumerate(series):
        s = ref_series(i, values_ref, title_ref, cats_ref)
        s.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill="FFFFFF", w=12700))
        s.invertIfNegative = False
        s.dLbls = data_labels(txt, pos="ctr", numfmt="0;-0;;", size=8)
        ch.series.append(s)
    ch.gapWidth = 50
    ch.legend = Legend()
    ch.legend.position = "b"
    ch.legend.txPr = text_props(8, INK2)
    cat_axis(ch.x_axis)
    hide_value_axis(ch.y_axis)
    style_chart_frame(ch)
    return ch


def verlauf_chart(cats_ref, values_ref, values_title_ref, soll_ref, soll_title_ref,
                  title=None, ymax=None, major=None, labels=True, legend=True, color=C_PLAN, axis_size=9):
    """Saeulen = Bestand je Zeitpunkt, gestrichelte Linie = Soll (beides auf derselben Achse)."""
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.style = 2
    s = fill_series(ref_series(0, values_ref, values_title_ref, cats_ref), color)
    if labels:
        s.dLbls = data_labels(INK2, pos="outEnd", numfmt="0;-0;0", size=8)
    bar.series.append(s)
    bar.gapWidth = 55

    line = LineChart()
    ls = ref_series(1, soll_ref, soll_title_ref, cats_ref)
    ls.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=INK, w=22225, prstDash="dash"))
    ls.marker = Marker(symbol="none")
    ls.smooth = False
    line.series.append(ls)

    cat_axis(bar.x_axis)
    bar.x_axis.txPr = text_props(axis_size, INK2)
    if labels and not ymax:
        hide_value_axis(bar.y_axis)
    else:
        show_value_axis(bar.y_axis, ymax=ymax, major=major)
    line.x_axis = bar.x_axis
    line.y_axis = bar.y_axis
    bar += line
    if legend:
        bar.legend = Legend()
        bar.legend.position = "b"
        bar.legend.txPr = text_props(9, INK2)
    else:
        bar.legend = None
    if title:
        bar.title = chart_title(title, 9)
    bar.visible_cells_only = False      # Soll-Linie steht in ausgeblendeten Hilfsspalten
    style_chart_frame(bar)
    return bar


def place(ws, chart, top_left, bottom_right):
    """Diagramm genau auf einen Zellbereich legen (von der linken oberen Ecke von top_left bis zur linken oberen
    Ecke von bottom_right). So passt es unabhaengig von Schriftart/Spaltenbreite in Excel und LibreOffice."""
    c1, r1 = coordinate_from_string(top_left)
    c2, r2 = coordinate_from_string(bottom_right)
    chart.anchor = TwoCellAnchor(_from=AnchorMarker(col=column_index_from_string(c1) - 1, row=r1 - 1),
                                 to=AnchorMarker(col=column_index_from_string(c2) - 1, row=r2 - 1))
    ws.add_chart(chart)


def fix_row_heights(ws, height=15):
    """Feste Zeilenhoehe fuer alle Zeilen ohne eigene Hoehe (die Original-Datei hat eine andere Standardschrift)."""
    for r in range(1, ws.max_row + 2):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = height


# ---------------------------------------------------------------- Bausteine
GRID_COLS = 16   # Spalten B..Q
GRID_WIDTH = 10.5


def setup_grid(ws, hidden=()):
    ws.column_dimensions["A"].width = 2
    for i in range(2, 2 + GRID_COLS):
        ws.column_dimensions[get_column_letter(i)].width = GRID_WIDTH
    ws.column_dimensions["R"].width = 2
    for col in hidden:
        ws.column_dimensions[col].hidden = True
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[4].height = 8


def nav_link(ws, rng_, target, text):
    link = "#" + quote_sheetname(target).replace('"', '""') + "!A1"
    merge_put(ws, rng_, f'=HYPERLINK("{link}","{text}")', f=font(10, color="1C5CAB", underline="single"), al=RIGHT)


def header_block(ws, title, subtitle_formula, links=()):
    merge_put(ws, "B2:L2", title, f=font(18, True), al=LEFT)
    ws.row_dimensions[2].height = 30
    if links:
        cells = ["M2:Q2"] if len(links) == 1 else ["M2:N2", "O2:Q2"]
        for c, (target, text) in zip(cells, links):
            nav_link(ws, c, target, text)
    merge_put(ws, "B3:Q3", subtitle_formula, f=font(9, color=INK2), al=LEFT)
    ws.row_dimensions[3].height = 16


def kpi_tiles(ws, tiles, row=5):
    """tiles: Liste (Titel, Wert-Formel, Unterzeile-Formel, Akzentfarbe); je 3 Spalten ab B."""
    ws.row_dimensions[row].height = 18
    ws.row_dimensions[row + 1].height = 34
    ws.row_dimensions[row + 2].height = 16
    for i, (label, value, sub, accent) in enumerate(tiles):
        c1 = 2 + 3 * i
        c3 = c1 + 2
        L1, L3 = get_column_letter(c1), get_column_letter(c3)
        merge_put(ws, f"{L1}{row}:{L3}{row}", label, f=font(10, True, INK2), al=Alignment(horizontal="left", vertical="bottom", indent=1))
        merge_put(ws, f"{L1}{row+1}:{L3}{row+1}", value, f=font(24, True, INK), al=Alignment(horizontal="left", vertical="center", indent=1), nf="0")
        merge_put(ws, f"{L1}{row+2}:{L3}{row+2}", sub, f=font(9, color=MUTED), al=Alignment(horizontal="left", vertical="top", indent=1))
        accent_side = Side(style="thick", color=accent)
        for r in range(row, row + 3):
            for c in range(c1, c3 + 1):
                cell = ws.cell(r, c)
                cell.fill = fill(TILE_BG)
                cell.border = Border(left=accent_side if c == c1 else None,
                                     right=Side(style="thick", color="FFFFFF") if c == c3 else None)


def section_title(ws, ref_range, text):
    merge_put(ws, ref_range, text, f=font(12, True), al=LEFT)
    r = int("".join(ch for ch in ref_range.split(":")[0] if ch.isdigit()))
    ws.row_dimensions[r].height = 20


def section_note(ws, ref_range, text):
    merge_put(ws, ref_range, text, f=font(9, color=MUTED), al=LEFT)


def status_header_cells(ws, row, first_col):
    """Status-Bezeichnungen (aus Einstellungen) als Spaltenkoepfe."""
    for code in range(1, 10):
        col = get_column_letter(first_col + code - 1)
        put(ws, f"{col}{row}", f"={stat_label_ref(code)}", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)


def pct_sub(count_ref, total_ref):
    return f'=IF({total_ref}=0,"–",TEXT({count_ref}/{total_ref},"0%")&" von "&{total_ref}&" Mitarbeitern")'


def timeline_label(k):
    """Spaltenkopf des Verlaufs: k=0 Heute, sonst Ende Bezugsjahr+k-1."""
    if k == 0:
        return "Heute"
    return f'=""&({E_YEAR}+{k - 1})'


def add_compare_cf(ws, cell_range, soll_col, first_row):
    """Gruen, wenn der Bestand das Soll erreicht, rot, wenn er darunter liegt."""
    first = cell_range.split(":")[0]
    col = "".join(ch for ch in first if ch.isalpha())
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{col}{first_row}<${soll_col}{first_row}"], fill=fill(BAD_BG), font=Font(color=BAD_INK)))
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"AND(${soll_col}{first_row}>0,{col}{first_row}>=${soll_col}{first_row})"], fill=fill(GOOD_BG), font=Font(color=GOOD_INK)))


# ---------------------------------------------------------------- Blatt: Zielzustand (Kopie)
def build_zielzustand(ws, src_ws, src_name):
    """Kopiert Tabelle1 aus LST_Zielzustand_gesamt.xlsx (Werte, Formeln, Formate)."""
    for row in src_ws.iter_rows():
        for c in row:
            if c.__class__.__name__ == "MergedCell":
                continue
            d = ws.cell(c.row, c.column)
            d.value = c.value
            if c.has_style:
                d.font = copy(c.font)
                d.fill = copy(c.fill)
                d.border = copy(c.border)
                d.alignment = copy(c.alignment)
                d.number_format = c.number_format
    for m in src_ws.merged_cells.ranges:
        ws.merge_cells(str(m))
    for k, dim in src_ws.column_dimensions.items():
        ws.column_dimensions[k].width = dim.width
        ws.column_dimensions[k].hidden = dim.hidden
    for k, dim in src_ws.row_dimensions.items():
        if dim.height:
            ws.row_dimensions[k].height = dim.height
    ws.freeze_panes = "B3"
    ws.page_setup.orientation = "landscape"
    put(ws, "A1", "Zielzustand (Soll)", f=font(11, True))
    r = src_ws.max_row + 2
    notes = [
        f"Quelle: {src_name}, Blatt „{src_ws.title}“ (unverändert übernommen).",
        "Für den Soll-Ist-Vergleich wird je Bezirk die Spalte „Zielzustand“ in den Zeilen Arbeiter LST, Wmech, SigMech, "
        "SigMech RBEG und Teamleiter verwendet.",
        "Zahlen hier ändern → Übersicht, Soll-Ist und Bezirksseiten passen sich automatisch an. "
        "Zuordnung Bezirks-Blatt → Bezirk im Zielzustand: Blatt „Einstellungen“.",
    ]
    for i, n in enumerate(notes):
        put(ws, f"A{r + i}", n, f=font(9, italic=True, color=INK2))


# ---------------------------------------------------------------- Blatt: Einstellungen
def build_settings(ws, districts):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 70
    put(ws, "B1", "Einstellungen & Erklärung", f=font(16, True))
    put(ws, "B2", "Gelb hinterlegte Zellen dürfen geändert werden. Alle Zahlen und Diagramme passen sich automatisch an.", f=font(10, color=INK2))

    put(ws, "B4", "Bezugsjahr (aktuelles Jahr)", f=font(10, True), al=LEFT)
    put(ws, "C4", "=YEAR(TODAY())", f=font(11, True, "0000FF"), fl=fill(INPUT_BG), al=CENTER, nf="0")
    put(ws, "D4", "Standard: =JAHR(HEUTE()) – kann mit einer festen Jahreszahl überschrieben werden (z. B. 2027). "
                  "Geplante Jahre davor gelten als „Überfällig“.", f=font(9, color=INK2), al=LEFT)
    put(ws, "B5", "Planjahre über das Ziel hinaus zählen", f=font(10, True), al=LEFT)
    put(ws, "C5", BEYOND_DEFAULT, f=font(11, True, "0000FF"), fl=fill(INPUT_BG), al=CENTER)
    put(ws, "D5", "Ja = der Verlauf folgt allen Jahreszahlen in L–S (z. B. Ziel SigMech, aber RBEG 28 eingetragen → ab 2028 RBEG). "
                  "Nein = jeder kommt höchstens bis zu seiner Ziel-Qualifikation (Spalte G).", f=font(9, color=INK2), al=LEFT)
    dv = DataValidation(type="list", formula1='"Ja,Nein"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C5")

    put(ws, "B6", "Status-Kategorien", f=font(12, True))
    for col, text in zip("BCDE", ["Nr.", "Bezeichnung", "Farbe", "Bedeutung"]):
        put(ws, f"{col}7", text, f=font(10, True, INK2), fl=fill(HEAD_BG), al=LEFT)
    for col in "FGH":
        put(ws, f"{col}7", None, fl=fill(HEAD_BG))
    for code, color, _ in STATUS:
        r = E_STAT_FIRST + code - 1
        label, meaning = STATUS_TEXT[code]
        if label is None:
            off = code - 2
            label = (f'="Abschluss "&({E_YEAR}+{off})' if code < 6 else f'="Abschluss ab "&({E_YEAR}+{off})')
        put(ws, f"B{r}", code, al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"C{r}", label, f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"D{r}", None, fl=fill(color), border=Border(bottom=Side(style="thin", color="FFFFFF")))
        merge_put(ws, f"E{r}:H{r}", meaning, f=font(9, color=INK2), al=LEFT_WRAP, border=BOTTOM_HAIR)
        ws.row_dimensions[r].height = 26 if len(meaning) > 110 else 16

    # Bezirke
    put(ws, f"B{E_BEZ_HDR - 2}", "Bezirke: Blatt → Bezirk im Zielzustand", f=font(12, True))
    put(ws, f"B{E_BEZ_HDR - 1}", "Name genau wie in Zeile 1 des Blatts „Zielzustand“. Die Spalte „Zielzustand“ liegt 3 Spalten rechts davon.",
        f=font(9, color=INK2))
    for col, text in zip("BCDE", ["Bezirks-Blatt", "Name im Zielzustand", "Spalte Soll (automatisch)", "Prüfung"]):
        put(ws, f"{col}{E_BEZ_HDR}", text, f=font(10, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    ws.row_dimensions[E_BEZ_HDR].height = 28
    for i in range(E_BEZ_LAST - E_BEZ_FIRST + 1):
        r = E_BEZ_FIRST + i
        d = districts[i] if i < len(districts) else None
        put(ws, f"B{r}", d.sheet if d else None, f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"C{r}", d.zz_name if d else None, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"D{r}", f'=IF(C{r}="","",IFERROR(MATCH(C{r},Zielzustand!$1:$1,0)+3,""))', al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"E{r}", f'=IF(B{r}="","",IF(D{r}="","nicht gefunden – kein Soll","ok"))', al=LEFT, border=BOTTOM_HAIR)

    # Qualifikationen
    put(ws, f"B{E_MAP_HDR - 2}", "Zuordnung Qualifikation → Spalte (L–S), Verwendungsprüfung (U/V) und Kategorie", f=font(12, True))
    put(ws, f"B{E_MAP_HDR - 1}", "Wird in der Spalte Ziel- oder Ist-Qualifikation ein neuer Begriff verwendet, hier eine Zeile ergänzen. "
                                 "Die Kategorie muss genau einer Zeile im Zielzustand entsprechen (sonst „Sonstige“).", f=font(9, color=INK2))
    for col, text in zip("BCDEFGH", ["Text in Ziel-/Ist-Spalte", "Spalte mit dem Stand (L–S)", "Nr. (auto)",
                                     "Spalte örtl. Verwendungs­prüfung (U/V)", "Nr. (auto)", "Kategorie (Zielzustand)", "Hinweis"]):
        put(ws, f"{col}{E_MAP_HDR}", text, f=font(10, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    ws.row_dimensions[E_MAP_HDR].height = 42
    hints = {"Azubi": "Azubis zählen erst nach Abschluss in der Ziel-Kategorie",
             "Kennziffer 4": "Zusatzqualifikation, i. d. R. kein Ziel",
             "Senior Expert LST": "keine eigene Spalte – gilt als fertig, wenn Ist = Ziel",
             "Weichmech": "fertig bei „x“ in Spalte N oder „x“ bei der örtl. Verwendungsprüfung Wmech (U)",
             "Sigmech": "fertig bei „x“ in Spalte O oder „x“ bei der örtl. Verwendungsprüfung SigMech (V)"}
    q1, q2 = (ord(c) for c in SRC_QUALI)
    v1, v2 = (ord(c) for c in SRC_VP)
    for i in range(E_MAP_LAST - E_MAP_FIRST + 1):
        r = E_MAP_FIRST + i
        text, col, vp, cat = MAPPING[i] if i < len(MAPPING) else (None, None, None, None)
        put(ws, f"B{r}", text, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"C{r}", col or None, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"D{r}", f'=IF(C{r}="","",IFERROR(IF(AND(CODE(UPPER(C{r}))>={q1},CODE(UPPER(C{r}))<={q2}),CODE(UPPER(C{r}))-{q1 - 1},""),""))',
            al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"E{r}", vp or None, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"F{r}", f'=IF(E{r}="","",IFERROR(IF(AND(CODE(UPPER(E{r}))>={v1},CODE(UPPER(E{r}))<={v2}),CODE(UPPER(E{r}))-{v1 - 1},""),""))',
            al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"G{r}", cat, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"H{r}", hints.get(text), f=font(9, color=INK2), al=LEFT, border=BOTTOM_HAIR)

    r = E_MAP_LAST + 3
    put(ws, f"B{r}", "So wird gezählt", f=font(12, True))
    notes = [
        "Ausgewertet werden die Bezirks-Blätter aus der Tabelle „Bezirke“ oben. Je Blatt werden die Zeilen "
        f"{SRC_FIRST_ROW}–{SRC_FIRST_ROW + SRC_ROWS - 1} gelesen. Eine Zeile zählt als Mitarbeiter, wenn ein Name (Spalte B) und eine Ziel- "
        "oder Ist-Qualifikation (Spalte G oder J) eingetragen ist.",
        "Für jeden Mitarbeiter wird die Spalte (L–S) gesucht, die zu seiner Ziel-Qualifikation gehört. Steht dort „x“, ist er fertig; "
        "steht dort eine Jahreszahl (z. B. 27 = 2027), ist das der geplante Abschluss.",
        "Weichenmechaniker und Signalmechaniker sind auch fertig, wenn die örtliche Verwendungsprüfung (U bzw. V) mit „x“ abgehakt ist. "
        "Steht dort ein Jahr, zählt das frühere der beiden Jahre (Ziel-Spalte oder Verwendungsprüfung) als geplanter Abschluss.",
        "Soll-Ist: Jeder Mitarbeiter zählt genau einmal – auf der Stufe, die er erreicht hat. Die Stufen laufen von links nach rechts: "
        "Azubi (L) → Arbeiter LST (M) → Wmech (N) → SigMech (O) → SigMech RBEG (P) → Teamleiter (S). Wer SigMech RBEG wird, zählt nicht "
        "mehr als SigMech. Heute = rechteste Spalte mit „x“ (Wmech/SigMech auch „x“ bei der Verwendungsprüfung U/V); ohne „x“ zählt "
        "die Ist-Qualifikation. Soll = Spalte „Zielzustand“ im Blatt Zielzustand.",
        "Verlauf: Zum Jahresende zählt die rechteste Spalte mit „x“ oder mit einem geplanten Jahr bis dahin. Ob dabei auch Jahre "
        "über die Ziel-Qualifikation hinaus zählen, steht oben in Zelle C5. Überfällige Jahre zählen ab dem Bezugsjahr. "
        "„Nach Plan“ = alle eingetragenen Jahre erreicht.",
        "Passend besetzt = Ziel-Stellen, auf denen jemand mit genau dieser Qualifikation sitzt (je Stufe höchstens so viele wie im "
        "Soll). Fehlende Ziel-Stellen = Soll − passend besetzt. Mehr als Soll = Mitarbeiter auf Stufen, die schon voll sind.",
        "Das Blatt „Daten“ enthält die Auswertung je Mitarbeiter. Es wird per Formel erzeugt und sollte nicht von Hand geändert werden.",
    ]
    for i, n in enumerate(notes):
        merge_put(ws, f"B{r + 1 + i}:H{r + 1 + i}", "• " + n, f=font(10, color=INK2), al=LEFT_WRAP)
        ws.row_dimensions[r + 1 + i].height = 30
    page_setup(ws, f"A1:H{r + len(notes)}")


# ---------------------------------------------------------------- Blatt: Daten
DATA_COLS = [
    ("A", "Bezirk", 18), ("B", "Quellzeile", 9), ("C", "Name", 18), ("D", "Vorname", 14),
    ("E", "Ziel-Qualifikation", 18), ("F", "Ist-Qualifikation", 18), ("G", "Gültig", 7),
    ("H", "Ziel-Nr", 7), ("I", "Ziel-Spalte", 8), ("J", "VP-Spalte", 8), ("K", "Ziel-Kategorie (für Status)", 14),
    ("L", "Ist-Nr", 7), ("M", "Ist-Spalte", 8), ("N", "Ist-Kategorie", 14),
    ("O", "Eintrag Ziel-Spalte", 10), ("P", "Eintrag Verwendungsprüfung", 12), ("Q", "Jahr Ziel-Spalte", 9),
    ("R", "Jahr Verwendungsprüfung", 12), ("S", "Planjahr", 9), ("T", "Fertig", 7), ("U", "Status-Nr", 8),
    ("V", "Status", 20), ("W", "Sortierung", 10), ("X", "Rang im Bezirk", 9), ("Y", "Schlüssel", 20),
]
# Stufen fuer Soll-Ist und Verlauf (aufsteigend). Jeder Mitarbeiter zaehlt in der hoechsten erreichten Stufe.
# (Kategorie wie im Zielzustand, Spalte im Bezirks-Blatt, Spalte oertl. Verwendungspruefung)
LEVELS = [("Arbeiter LST", "M", None), ("Wmech", "N", "U"), ("SigMech", "O", "V"),
          ("SigMech RBEG", "P", None), ("Teamleiter", "S", None)]
assert [lv[0] for lv in LEVELS] == CATS
ENTRY_COLS = [get_column_letter(26 + i) for i in range(len(LEVELS))]            # Z..AD: Eintraege M, N, O, P, S
VP_ENTRY = {"U": get_column_letter(26 + len(LEVELS)), "V": get_column_letter(27 + len(LEVELS))}   # AE, AF
ZR_COL = get_column_letter(28 + len(LEVELS))                                     # AG: Stufe der Ziel-Qualifikation
AB_COLS = [get_column_letter(29 + len(LEVELS) + i) for i in range(len(LEVELS))]   # AH..AL: Stufe vorhanden ab Jahr
TL_COLS = [get_column_letter(34 + len(LEVELS) + k) for k in range(N_YEARS + 1)]  # AM..AT: Heute, BJ..BJ+6
PLAN_COL = get_column_letter(35 + len(LEVELS) + N_YEARS)                         # AU: nach Plan


def year_formula(c):
    v = f"VALUE({c})"
    return (f'=IFERROR(IF({c}="","",IF({v}>2100,YEAR({v}),IF({v}>=1900,ROUND({v},0),'
            f'IF(AND({v}>=1,{v}<=99),2000+ROUND({v},0),"")))),"")')


def year9(c):
    """Jahr aus einem Eintrag (27 -> 2027, 2027, Datum); 9999, wenn kein Jahr."""
    v = f"VALUE({c})"
    return (f'IFERROR(IF({c}="",9999,IF({v}>2100,YEAR({v}),IF({v}>=1900,ROUND({v},0),'
            f'IF(AND({v}>=1,{v}<=99),2000+ROUND({v},0),9999)))),9999)')


def build_data(ws, districts):
    ws.sheet_view.showGridLines = True
    heads = list(DATA_COLS)
    heads += [(c, f"Eintrag {lv[0]} (Spalte {lv[1]})", 10) for c, lv in zip(ENTRY_COLS, LEVELS)]
    heads += [(VP_ENTRY[v], f"Eintrag Verwendungsprüfung (Spalte {v})", 12) for v in SRC_VP]
    heads += [(ZR_COL, "Stufe der Ziel-Qualifikation (1–5)", 10)]
    heads += [(c, f"{lv[0]} ab (0 = vorhanden, 9999 = nicht geplant)", 13) for c, lv in zip(AB_COLS, LEVELS)]
    for col, h, w in heads:
        ws.column_dimensions[col].width = w
        put(ws, f"{col}1", h, f=font(10, True, "FFFFFF"), fl=fill("52514E"), al=LEFT_WRAP)
    for k, col in enumerate(TL_COLS):
        ws.column_dimensions[col].width = 13
        head = "Höchste Qualifikation heute" if k == 0 else f'="Höchste Qualifikation Ende "&({E_YEAR}+{k - 1})'
        put(ws, f"{col}1", head, f=font(10, True, "FFFFFF"), fl=fill("52514E"), al=LEFT_WRAP)
    ws.column_dimensions[PLAN_COL].width = 13
    put(ws, f"{PLAN_COL}1", "Höchste Qualifikation nach Plan", f=font(10, True, "FFFFFF"), fl=fill("52514E"), al=LEFT_WRAP)
    ws.row_dimensions[1].height = 45
    ws.freeze_panes = "C2"
    last = 1 + len(districts) * SRC_ROWS
    ws.auto_filter.ref = f"A1:{PLAN_COL}{last}"
    r = 2
    plain = font(10)
    q1, q2 = SRC_QUALI
    v1, v2 = SRC_VP
    for d in districts:
        s = quote_sheetname(d.sheet)
        for k in range(SRC_ROWS):
            f = {
                "A": d.sheet,
                "B": SRC_FIRST_ROW + k,
                "C": f'=TRIM(INDEX({s}!${SRC_NAME}:${SRC_NAME},$B{r})&"")',
                "D": f'=TRIM(INDEX({s}!${SRC_FIRST}:${SRC_FIRST},$B{r})&"")',
                "E": f'=TRIM(INDEX({s}!${SRC_ZIEL}:${SRC_ZIEL},$B{r})&"")',
                "F": f'=TRIM(INDEX({s}!${SRC_IST}:${SRC_IST},$B{r})&"")',
                "G": f'=IF(AND(C{r}<>"",C{r}<>"Name",OR(E{r}<>"",F{r}<>"")),1,0)',
                "H": f'=IF(E{r}="",0,IFERROR(MATCH(E{r},{E_MAP_TEXT},0),0))',
                "I": f'=IF(H{r}=0,0,IFERROR(INDEX({E_MAP_COLNO},H{r})*1,0))',
                "J": f'=IF(H{r}=0,0,IFERROR(INDEX({E_MAP_VPNO},H{r})*1,0))',
                "K": f'=IF(H{r}=0,"{OTHER}",IF(INDEX({E_MAP_CAT},H{r})&""="","{OTHER}",INDEX({E_MAP_CAT},H{r})&""))',
                "L": f'=IF(F{r}="",0,IFERROR(MATCH(F{r},{E_MAP_TEXT},0),0))',
                "M": f'=IF(L{r}=0,0,IFERROR(INDEX({E_MAP_COLNO},L{r})*1,0))',
                "N": f'=IF(L{r}=0,"{OTHER}",IF(INDEX({E_MAP_CAT},L{r})&""="","{OTHER}",INDEX({E_MAP_CAT},L{r})&""))',
                "O": f'=IF(I{r}=0,"",TRIM(INDEX({s}!${q1}:${q2},$B{r},I{r})&""))',
                "P": f'=IF(J{r}=0,"",TRIM(INDEX({s}!${v1}:${v2},$B{r},J{r})&""))',
                "Q": year_formula(f"O{r}"),
                "R": year_formula(f"P{r}"),
                "S": f'=IF(AND(Q{r}="",R{r}=""),"",MIN(Q{r},R{r}))',
                # fertig: "x" in Ziel-Spalte oder Verwendungspruefung; Ist = Ziel nur, wenn dort gar nichts steht
                "T": (f'=IF(G{r}=0,0,IF(OR(LOWER(O{r})="x",LOWER(P{r})="x",AND(O{r}="",P{r}="",'
                      f'OR(AND(E{r}<>"",LOWER(E{r})=LOWER(F{r})),AND(I{r}>0,I{r}=M{r})))),1,0))'),
                "U": f'=IF(G{r}=0,"",IF(E{r}="",9,IF(T{r}=1,1,IF(S{r}="",8,IF(S{r}<{E_YEAR},7,MIN(6,2+S{r}-{E_YEAR}))))))',
                "V": f'=IF(U{r}="","",INDEX({E_STAT_LABEL},U{r}))',
                "W": f'=IF(U{r}="","",U{r}*1000+B{r})',
                "X": f'=IF(W{r}="","",COUNTIFS($A$2:$A${last},A{r},$W$2:$W${last},"<"&W{r})+1)',
                "Y": f'=IF(X{r}="","",A{r}&"|"&X{r})',
            }
            # Stufen: Eintrag je Spalte und "vorhanden ab" (0 = "x", sonst geplantes Jahr; Ueberfaellige ab Bezugsjahr)
            for ec, (cat, src, vp) in zip(ENTRY_COLS, LEVELS):
                f[ec] = f'=TRIM(INDEX({s}!${src}:${src},$B{r})&"")'
            for vp in SRC_VP:
                f[VP_ENTRY[vp]] = f'=TRIM(INDEX({s}!${vp}:${vp},$B{r})&"")'
            zr = "0"
            for i, (cat, _, _) in enumerate(LEVELS, start=1):
                zr = f'IF($K{r}="{cat}",{i},{zr})'
            f[ZR_COL] = f"={zr}"
            for rank, (ac, ec, (cat, src, vp)) in enumerate(zip(AB_COLS, ENTRY_COLS, LEVELS), start=1):
                e = f"{ec}{r}"
                plan = f"MAX(MIN({year9(e)},{year9(f'{VP_ENTRY[vp]}{r}')}),{E_YEAR})" if vp else f"MAX({year9(e)},{E_YEAR})"
                # geplante Stufen ueber der Ziel-Qualifikation nur, wenn in den Einstellungen "Ja"
                plan = f'IF(AND({E_BEYOND}="Nein",{rank}>${ZR_COL}{r}),9999,{plan})'
                done = f'OR(LOWER({e})="x",LOWER({VP_ENTRY[vp]}{r})="x")' if vp else f'LOWER({e})="x"'
                f[ac] = f'=IF($G{r}=0,"",IF({done},0,{plan}))'

            # Kategorie zu einem Zeitpunkt: hoechste Stufe, die bis dahin erreicht ist; sonst Ist-Qualifikation
            def category_at(thr):
                expr = f"$N{r}"
                for ac, (cat, _, _) in zip(AB_COLS, LEVELS):
                    expr = f'IF(${ac}{r}<={thr},"{cat}",{expr})'
                return f'=IF($G{r}=0,"",{expr})'
            for kk, col in enumerate(TL_COLS):
                f[col] = category_at("0" if kk == 0 else f"{E_YEAR}+{kk - 1}")
            f[PLAN_COL] = category_at("9998")
            for col, v in f.items():
                c = ws[f"{col}{r}"]
                c.value = v
                c.font = plain
            for col in "QRS":
                ws[f"{col}{r}"].number_format = "0"
            r += 1
    return last


# ---------------------------------------------------------------- Soll-Ist-Tabelle (Bezirk und LST gesamt)
# Spalten: B:C Kategorie, D Soll, E..L Verlauf (Heute, BJ..BJ+6), M Nach Plan, N Delta
T_SOLL, T_TL0, T_PLAN, T_DELTA = "D", 5, "M", "N"
TL_LETTERS = [get_column_letter(T_TL0 + k) for k in range(N_YEARS + 1)]   # E..L
HELP0 = 21                                                                 # Spalte U: Soll-Linie (Hilfswerte)
HELP_LETTERS = [get_column_letter(HELP0 + k) for k in range(N_YEARS + 1)]  # U..AB
COMP_LETTERS = [get_column_letter(HELP0 + N_YEARS + 1 + k) for k in range(N_YEARS + 2)]   # AC..AK: Zusammensetzung + Ziel


COUNT_NOTE = (
    "So wird gezählt: Jeder Mitarbeiter zählt genau einmal – auf der Stufe, die er erreicht hat. Die Stufen laufen von links nach "
    "rechts: Azubi → Arbeiter LST → Wmech → SigMech → SigMech RBEG → Teamleiter. Wer SigMech RBEG wird, zählt ab dann nicht mehr "
    "als SigMech. „Passend besetzt“ = Ziel-Stellen, auf denen jemand mit genau dieser Qualifikation sitzt.")
# Farben der Stufen (Zusammensetzung): hell = frueh im Ausbildungsweg, dunkel = spaet
STAGE_COLORS = {OTHER: ("E1E0D9", INK2), "Arbeiter LST": ("9EC5F4", INK), "Wmech": ("5598E7", INK),
                "SigMech": ("2A78D6", "FFFFFF"), "SigMech RBEG": ("1C5CAB", "FFFFFF"), "Teamleiter": ("0D366B", "FFFFFF")}
ROW_TOTAL, ROW_PASS = "Gesamt", "Passend"
ROW_LABEL = {ROW_TOTAL: "Gesamt (mit Qualifikation)", ROW_PASS: "davon passend besetzt",
             OTHER: "Azubis & Sonstige"}


def soll_ist_table(ws, hdr, soll_fn, count_fn, pass_fn=None):
    """Kopf, je Stufe eine Zeile, dann Gesamt, passend besetzt, noch ohne Abschluss.
    soll_fn(r, cat) / count_fn(r, daten_spalte, cat) liefern Formeln; pass_fn(tabellen_spalte) ersetzt die Formel der Zeile
    „passend besetzt“ (fuer LST gesamt: Summe der Bezirke). Gibt {Zeilenname: Zeile} zurueck."""
    merge_put(ws, f"B{hdr}:C{hdr}", "Qualifikation", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    put(ws, f"{T_SOLL}{hdr}", "Soll (Zielzustand)", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    for k, col in enumerate(TL_LETTERS):
        put(ws, f"{col}{hdr}", timeline_label(k), f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    put(ws, f"{T_PLAN}{hdr}", "Nach Plan", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    put(ws, f"{T_DELTA}{hdr}", "Nach Plan − Soll", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    merge_put(ws, f"O{hdr}:Q{hdr}", "Bewertung (nach Plan)", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    ws.row_dimensions[hdr].height = 30
    put(ws, f"{HELP_LETTERS[0]}{hdr}", "Soll-Linie (Hilfswerte für Diagramme)", f=font(8, color=MUTED))
    for k, col in enumerate(COMP_LETTERS[:-1]):
        put(ws, f"{col}{hdr}", f"={TL_LETTERS[k]}{hdr}", f=font(8, color=MUTED))
    put(ws, f"{COMP_LETTERS[-1]}{hdr}", "Ziel", f=font(8, color=MUTED))
    values = [T_SOLL] + TL_LETTERS + [T_PLAN]
    data_cols = TL_COLS + [PLAN_COL]
    rows = {}
    r = hdr + 1
    for cat in CATS:
        rows[cat] = r
        merge_put(ws, f"B{r}:C{r}", cat, f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"{T_SOLL}{r}", soll_fn(r, cat), f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
        for col, dc in zip(values[1:], data_cols):
            put(ws, f"{col}{r}", count_fn(r, dc, cat), f=font(10, col == T_PLAN), al=CENTER, nf="0", border=BOTTOM_HAIR)
        r += 1
    cat_rows = [rows[c] for c in CATS]
    rows[ROW_TOTAL], rows[ROW_PASS], rows[OTHER] = r, r + 1, r + 2
    rt, rp, ro = r, r + 1, r + 2
    merge_put(ws, f"B{rt}:C{rt}", ROW_LABEL[ROW_TOTAL], f=font(10, True), al=LEFT, border=TOP_INK)
    merge_put(ws, f"B{rp}:C{rp}", ROW_LABEL[ROW_PASS], f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
    merge_put(ws, f"B{ro}:C{ro}", ROW_LABEL[OTHER], f=font(9, color=MUTED), al=LEFT, border=BOTTOM_HAIR)
    for col in values:
        put(ws, f"{col}{rt}", "=" + "+".join(f"{col}{x}" for x in cat_rows), f=font(10, True), al=CENTER, nf="0", border=TOP_INK)
        if col == T_SOLL:
            formula = f"={T_SOLL}{rt}"
        elif pass_fn:
            formula = pass_fn(col)
        else:   # je Stufe hoechstens so viele, wie das Soll vorsieht
            formula = "=" + "+".join(f"MIN({col}{x},{T_SOLL}{x})" for x in cat_rows)
        put(ws, f"{col}{rp}", formula, f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
    put(ws, f"{T_SOLL}{ro}", "–", f=font(9, color=MUTED), al=CENTER, border=BOTTOM_HAIR)
    for col, dc in zip(values[1:], data_cols):
        put(ws, f"{col}{ro}", count_fn(ro, dc, OTHER), f=font(9, color=MUTED), al=CENTER, nf="0", border=BOTTOM_HAIR)
    for col in (T_DELTA, "O", "P", "Q"):
        put(ws, f"{col}{ro}", None, border=BOTTOM_HAIR)

    for r in cat_rows + [rt, rp]:
        S_, P_ = f"{T_SOLL}{r}", f"{T_PLAN}{r}"
        border = TOP_INK if r == rt else BOTTOM_HAIR
        put(ws, f"{T_DELTA}{r}", f"={P_}-{S_}", f=font(10, r in (rt, rp)), al=CENTER, nf=NF_DELTA, border=border)
        if r == rp:
            text = (f'=IF({P_}>={S_},"alle Ziel-Stellen passend besetzt",({S_}-{P_})&IF({S_}-{P_}=1," Ziel-Stelle"," Ziel-Stellen")'
                    f'&" nicht passend besetzt")')
        else:
            text = (f'=IF(AND({S_}=0,{P_}=0),"–",IF({P_}={S_},"Soll erreicht",IF({P_}>{S_},({P_}-{S_})&" mehr als Soll",'
                    f'IF({S_}-{P_}=1,"es fehlt 1","es fehlen "&({S_}-{P_})))))')
        merge_put(ws, f"O{r}:Q{r}", text, f=font(9, r in (rt, rp), INK2), al=LEFT, border=border)
        for col in HELP_LETTERS:
            put(ws, f"{col}{r}", f"=${T_SOLL}{r}", f=font(8, color=MUTED), nf="0")
    # Hilfswerte Zusammensetzung: Verlauf + Ziel (Soll) als letzte Saeule
    for r in [ro] + cat_rows:
        for k, col in enumerate(COMP_LETTERS[:-1]):
            put(ws, f"{col}{r}", f"={TL_LETTERS[k]}{r}", f=font(8, color=MUTED), nf="0")
        put(ws, f"{COMP_LETTERS[-1]}{r}", 0 if r == ro else f"={T_SOLL}{r}", f=font(8, color=MUTED), nf="0")
    first = cat_rows[0]
    add_compare_cf(ws, f"{TL_LETTERS[0]}{first}:{T_PLAN}{rp}", T_SOLL, first)
    merge_put(ws, f"B{ro + 1}:Q{ro + 1}", COUNT_NOTE, f=font(8, italic=True, color=MUTED), al=LEFT_WRAP)
    ws.row_dimensions[ro + 1].height = 24
    return rows


def comp_series(title, rows):
    """Serien fuer composition_chart aus den Hilfsspalten einer soll_ist_table (unten: noch ohne Abschluss)."""
    out = []
    for key in [OTHER] + CATS:
        r = rows[key]
        color, txt = STAGE_COLORS[key]
        out.append((rng(title, COMP_LETTERS[0], r, COMP_LETTERS[-1], r), rng(title, "B", r), color, txt))
    return out


# ---------------------------------------------------------------- Blatt: Grafik je Bezirk
R_SEC1, R_CH1 = 9, 10
R_SEC2, R_CH2 = 29, 31                  # Zeile 30: kurze Erklaerung
R_SEC3, R_CH3 = 47, 49
R_T1 = 61                               # Status nach Ziel-Qualifikation
MAT_HDR = R_T1 + 1
MAT_FIRST = MAT_HDR + 1
MAT_LAST = MAT_HDR + len(GROUPS)
MAT_SUM = MAT_LAST + 1
MAT_PCT = MAT_SUM + 1
R_T2 = MAT_PCT + 2                      # Soll-Ist und Verlauf
SI_HDR = R_T2 + 1
LIST_TITLE = SI_HDR + len(CATS) + 6
LIST_HDR = LIST_TITLE + 1
LIST_FIRST = LIST_HDR + 1
LIST_LAST = LIST_FIRST + LIST_ROWS - 1
STAT_COL1 = 4                           # Spalte D = Status 1 ... Spalte L = Status 9
SUM_COL = 13                            # Spalte M

# Diagrammgroessen (cm), passend zu 16 Spalten a 9,3 Zeichen (~65 px)


def build_district(ws, d, sheet_names):
    setup_grid(ws, hidden=["S", "T"] + HELP_LETTERS + COMP_LETTERS)
    title = ws.title
    put(ws, "S1", d.sheet)
    put(ws, "S2", f'=IFERROR(INDEX({E_BEZ_ZZCOL},MATCH($S$1,{E_BEZ_SHEET},0))*1,0)')
    header_block(ws, f"Ausbildungsstand – {d.sheet}",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   Quelle: Blatt „{d.sheet}“ (Ziel-Qualifikation G, Ist-Qualifikation J, Stand L–S, '
                 f'örtl. Verwendungsprüfung U–V)   ·   Soll: Zielzustand „"&IFERROR(INDEX(Einstellungen!$C${E_BEZ_FIRST}:$C${E_BEZ_LAST},'
                 f'MATCH($S$1,{E_BEZ_SHEET},0)),"–")&"“"',
                 links=[("Übersicht", "← Übersicht"), ("Soll-Ist", "Soll-Ist gesamt →")])

    sc = lambda code: f"{get_column_letter(STAT_COL1 + code - 1)}{MAT_SUM}"
    total = f"$M${MAT_SUM}"
    in_training = f"SUM({sc(2)}:{sc(6)})"
    rt, rp = SI_HDR + 1 + len(CATS), SI_HDR + 2 + len(CATS)     # Zeilen Gesamt / passend besetzt
    # Hilfswerte (ausgeblendete Spalte T): fehlende Ziel-Stellen nach Plan / heute, mehr als Soll nach Plan
    put(ws, "T1", f"={T_SOLL}{rp}-{T_PLAN}{rp}")
    put(ws, "T2", f"={T_SOLL}{rp}-{TL_LETTERS[0]}{rp}")
    put(ws, "T3", f"={T_PLAN}{rt}-{T_PLAN}{rp}")
    kpi_tiles(ws, [
        ("Mitarbeiter", f"={total}", f'=IF({sc(9)}>0,"davon "&{sc(9)}&" ohne Ziel-Qualifikation","im Bezirk")', INK2),
        ("Fertig", f"={sc(1)}", pct_sub(sc(1), total), STATUS[0][1]),
        ("In Ausbildung", f"={in_training}",
         f'="davon "&{sc(3)}&" mit Abschluss "&({E_YEAR}+1)&IF({sc(7)}>0," · "&{sc(7)}&" überfällig","")', "2A78D6"),
        ("Fehlt (nichts geplant)", f"={sc(8)}", pct_sub(sc(8), total), STATUS[7][1]),
        ("Fehlende Ziel-Stellen", "=$T$1", '="nach Plan · heute: "&$T$2&" · mehr als Soll: "&$T$3', BAD_INK),
    ])

    section_title(ws, f"B{R_SEC1}:I{R_SEC1}", "Mitarbeiter nach Status")
    section_title(ws, f"J{R_SEC1}:Q{R_SEC1}", "Status nach Ziel-Qualifikation")
    section_title(ws, f"B{R_SEC2}:I{R_SEC2}", "Soll-Ist je Qualifikation")
    section_title(ws, f"J{R_SEC2}:Q{R_SEC2}", "Verlauf: Mitarbeiter je Stufe – rechts der Zielzustand")
    section_note(ws, f"B{R_SEC2 + 1}:Q{R_SEC2 + 1}",
                 "Jeder zählt einmal – auf seiner erreichten Stufe: Azubi → Arbeiter LST → Wmech → SigMech → SigMech RBEG → Teamleiter. "
                 "Wer SigMech RBEG wird, zählt nicht mehr als SigMech.")
    section_title(ws, f"B{R_SEC3}:Q{R_SEC3}", "Verlauf je Qualifikation")
    section_note(ws, f"B{R_SEC3 + 1}:Q{R_SEC3 + 1}",
                 "Säulen = Mitarbeiter auf dieser Stufe zum Jahresende (laut Planung) · gestrichelte Linie = Soll (Zielzustand)")

    # Zahlentabelle 1: Status nach Ziel-Kategorie
    section_title(ws, f"B{R_T1}:Q{R_T1}", "Zahlen: Status nach Ziel-Qualifikation")
    merge_put(ws, f"B{MAT_HDR}:C{MAT_HDR}", "Ziel-Qualifikation", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    status_header_cells(ws, MAT_HDR, STAT_COL1)
    put(ws, f"M{MAT_HDR}", "Summe", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    ws.row_dimensions[MAT_HDR].height = 30
    for i, g in enumerate(GROUPS):
        r = MAT_FIRST + i
        merge_put(ws, f"B{r}:C{r}", g, al=LEFT, border=BOTTOM_HAIR)
        for code in range(1, 10):
            col = get_column_letter(STAT_COL1 + code - 1)
            put(ws, f"{col}{r}", f'=COUNTIFS(Daten!$A:$A,$S$1,Daten!$K:$K,$B{r},Daten!$U:$U,{code})',
                al=CENTER, nf=NF_COUNT, border=BOTTOM_HAIR)
        put(ws, f"M{r}", f"=SUM(D{r}:L{r})", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
    merge_put(ws, f"B{MAT_SUM}:C{MAT_SUM}", "Summe", f=font(10, True), al=LEFT)
    merge_put(ws, f"B{MAT_PCT}:C{MAT_PCT}", "Anteil", f=font(9, color=INK2), al=LEFT)
    for c in range(STAT_COL1, SUM_COL + 1):
        col = get_column_letter(c)
        put(ws, f"{col}{MAT_SUM}", f"=SUM({col}{MAT_FIRST}:{col}{MAT_LAST})", f=font(10, True), al=CENTER, nf="0", border=TOP_INK)
        put(ws, f"{col}{MAT_PCT}", f'=IF({total}=0,"",{col}{MAT_SUM}/{total})', f=font(9, color=INK2), al=CENTER, nf="0%")

    # Zahlentabelle 2: Soll-Ist und Verlauf
    section_title(ws, f"B{R_T2}:Q{R_T2}", "Zahlen: Soll-Ist und Verlauf je Qualifikation (Anzahl Mitarbeiter zum Jahresende)")
    soll = lambda r, cat: f'=IF($S$2=0,0,IFERROR(INDEX({ZZ_AREA},MATCH("{cat}",{ZZ_LABELS},0),$S$2)*1,0))'
    count = lambda r, col, cat: f'=COUNTIFS(Daten!$A:$A,$S$1,Daten!${col}:${col},"{cat}")'
    rows = soll_ist_table(ws, SI_HDR, soll, count)
    first = rows[CATS[0]]

    # Diagramme Reihe 1: Status
    ch_end = R_SEC2 - 1
    place(ws, status_bar_chart(ws, MAT_SUM, STAT_COL1, STAT_COL1 + 8, MAT_HDR), f"B{R_CH1}", f"J{ch_end}")
    place(ws, stacked_status_chart(ws, MAT_HDR, MAT_FIRST, MAT_LAST, 2, STAT_COL1, STAT_COL1 + 8), f"J{R_CH1}", f"R{ch_end}")
    # Reihe 2: Soll-Ist je Qualifikation + Zusammensetzung je Jahr
    cats = rng(title, "B", first, "B", first + len(CATS) - 1)
    soll_ist = soll_ist_chart(cats, [
        (rng(title, T_SOLL, first, T_SOLL, first + len(CATS) - 1), rng(title, T_SOLL, SI_HDR), C_SOLL),
        (rng(title, TL_LETTERS[0], first, TL_LETTERS[0], first + len(CATS) - 1), rng(title, TL_LETTERS[0], SI_HDR), C_HEUTE),
        (rng(title, T_PLAN, first, T_PLAN, first + len(CATS) - 1), rng(title, T_PLAN, SI_HDR), C_PLAN),
    ])
    place(ws, soll_ist, f"B{R_CH2}", f"J{R_SEC3 - 1}")
    comp_cats = rng(title, COMP_LETTERS[0], SI_HDR, COMP_LETTERS[-1], SI_HDR)
    place(ws, composition_chart(comp_cats, comp_series(title, rows)), f"J{R_CH2}", f"R{R_SEC3 - 1}")
    # Reihe 3: Verlauf je Qualifikation (kleine Diagramme)
    tl_cats = rng(title, TL_LETTERS[0], SI_HDR, TL_LETTERS[-1], SI_HDR)
    anchors = ["B", "F", "J", "N", "R"]
    for k, cat in enumerate(["Arbeiter LST", "Wmech", "SigMech", "SigMech RBEG"]):
        r = rows[cat]
        ch = verlauf_chart(tl_cats, rng(title, TL_LETTERS[0], r, TL_LETTERS[-1], r), rng(title, "B", r),
                           rng(title, HELP_LETTERS[0], r, HELP_LETTERS[-1], r), rng(title, T_SOLL, SI_HDR),
                           title=cat, labels=True, legend=False)
        place(ws, ch, f"{anchors[k]}{R_CH3}", f"{anchors[k + 1]}{R_T1 - 2}")

    # Mitarbeiterliste
    section_title(ws, f"B{LIST_TITLE}:Q{LIST_TITLE}", "Mitarbeiterliste (sortiert nach Status)")
    cols = [("B", "B", "Nr."), ("C", "D", "Name"), ("E", "F", "Vorname"), ("G", "H", "Ziel-Qualifikation"),
            ("I", "J", "Ist-Qualifikation"), ("K", "L", "Geplanter Abschluss"), ("M", "P", "Status")]
    for a, b, text in cols:
        if a == b:
            put(ws, f"{a}{LIST_HDR}", text, f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT)
        else:
            merge_put(ws, f"{a}{LIST_HDR}:{b}{LIST_HDR}", text, f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT)
    for k in range(1, LIST_ROWS + 1):
        r = LIST_FIRST + k - 1
        put(ws, f"S{r}", f'=IFERROR(MATCH($S$1&"|"&{k},Daten!$Y:$Y,0),"")')
        put(ws, f"T{r}", f'=IF($S{r}="","",INDEX(Daten!$U:$U,$S{r}))')
        put(ws, f"B{r}", f'=IF($S{r}="","",{k})', f=font(10, color=MUTED), al=LEFT)
        merge_put(ws, f"C{r}:D{r}", f'=IF($S{r}="","",INDEX(Daten!$C:$C,$S{r}))', f=font(10, True), al=LEFT)
        merge_put(ws, f"E{r}:F{r}", f'=IF($S{r}="","",INDEX(Daten!$D:$D,$S{r}))', al=LEFT)
        merge_put(ws, f"G{r}:H{r}", f'=IF($S{r}="","",INDEX(Daten!$E:$E,$S{r}))', al=LEFT)
        merge_put(ws, f"I{r}:J{r}", f'=IF($S{r}="","",INDEX(Daten!$F:$F,$S{r}))', al=LEFT)
        merge_put(ws, f"K{r}:L{r}", f'=IF($S{r}="","",IF(INDEX(Daten!$S:$S,$S{r})="","–",INDEX(Daten!$S:$S,$S{r})))', al=LEFT, nf="0")
        merge_put(ws, f"M{r}:P{r}", f'=IF($S{r}="","",INDEX(Daten!$V:$V,$S{r}))', f=font(10, True),
                  al=Alignment(horizontal="left", vertical="center", indent=1))
        ws.row_dimensions[r].height = 15
    over = LIST_LAST + 1
    merge_put(ws, f"B{over}:Q{over}",
              f'=IF({total}>{LIST_ROWS},"Hinweis: Der Bezirk hat mehr als {LIST_ROWS} Mitarbeiter – die Liste zeigt nur die ersten {LIST_ROWS}.","")',
              f=font(9, True, STATUS[7][1]), al=LEFT)

    # Bereiche ueberlappen nicht: LibreOffice wendet je Zelle nur einen Bereich an
    ws.conditional_formatting.add(f"B{LIST_FIRST}:L{LIST_LAST}",
                                  FormulaRule(formula=[f'$S{LIST_FIRST}<>""'], border=BOTTOM_HAIR))
    chip_border = Border(bottom=Side(style="thin", color="FFFFFF"))
    for code, color, txt in STATUS:
        ws.conditional_formatting.add(f"M{LIST_FIRST}:P{LIST_LAST}", FormulaRule(
            formula=[f"$T{LIST_FIRST}={code}"], fill=fill(color), font=Font(color=txt, bold=True), border=chip_border))
    # Druckseiten: Status | Soll-Ist + Verlauf | Zahlen | Mitarbeiterliste
    page_setup(ws, f"A1:Q{over}", breaks=[R_SEC2 - 1, R_T1 - 1, LIST_TITLE - 1])
    fix_row_heights(ws)
    ws.sheet_properties.tabColor = "86B6EF"
    return {"rows": rows}


# ---------------------------------------------------------------- Blatt: Uebersicht (Status)
OV_HDR = 34
OV_FIRST = OV_HDR + 1


def build_overview(ws, districts):
    setup_grid(ws, hidden=("S",))
    n = len(districts)
    ov_last = OV_FIRST + n - 1
    ov_sum = ov_last + 1
    ov_pct = ov_sum + 1
    header_block(ws, "Ausbildungsstand LST – alle Bezirke",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   {n} Bezirke   ·   Klick auf einen Bezirk in der Tabelle unten öffnet seine Seite   ·   '
                 f'Werte aktualisieren sich automatisch, wenn die Bezirks-Blätter geändert werden"',
                 links=[("Soll-Ist", "Soll-Ist-Vergleich →")])

    sc = lambda code: f"{get_column_letter(4 + code)}{ov_sum}"   # Status 1 -> Spalte E
    total = f"$N${ov_sum}"
    in_training = f"SUM({sc(2)}:{sc(6)})"
    kpi_tiles(ws, [
        ("Mitarbeiter gesamt", f"={total}", f'=IF({sc(9)}>0,"davon "&{sc(9)}&" ohne Ziel-Qualifikation","in allen Bezirken")', INK2),
        ("Fertig", f"={sc(1)}", pct_sub(sc(1), total), STATUS[0][1]),
        ("In Ausbildung", f"={in_training}",
         f'="davon "&{sc(3)}&" mit Abschluss "&({E_YEAR}+1)&IF({sc(7)}>0," · "&{sc(7)}&" überfällig","")', "2A78D6"),
        ("Fehlt (nichts geplant)", f"={sc(8)}", pct_sub(sc(8), total), STATUS[7][1]),
        ("Fehlende Ziel-Stellen", "='Soll-Ist'!$S$6", "='Soll-Ist'!$S$7", BAD_INK),
    ])
    section_title(ws, "B9:G9", "Alle Bezirke nach Status")
    section_title(ws, "H9:Q9", "Status je Bezirk")

    section_title(ws, f"B{OV_HDR - 1}:Q{OV_HDR - 1}", "Zahlen je Bezirk")
    merge_put(ws, f"B{OV_HDR}:D{OV_HDR}", "Bezirk (Klick öffnet die Seite)", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    status_header_cells(ws, OV_HDR, 5)
    put(ws, f"N{OV_HDR}", "Summe", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    merge_put(ws, f"O{OV_HDR}:P{OV_HDR}", "Anteil fertig", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    ws.row_dimensions[OV_HDR].height = 30
    for i, d in enumerate(districts):
        r = OV_FIRST + i
        put(ws, f"S{r}", d.sheet)
        link = "#" + quote_sheetname("Grafik " + d.sheet).replace('"', '""') + "!A1"
        merge_put(ws, f"B{r}:D{r}", f'=HYPERLINK("{link}","{d.sheet}")', f=font(10, True, "1C5CAB", underline="single"),
                  al=LEFT, border=BOTTOM_HAIR)
        for code in range(1, 10):
            col = get_column_letter(4 + code)
            put(ws, f"{col}{r}", f"=COUNTIFS(Daten!$A:$A,$S{r},Daten!$U:$U,{code})", al=CENTER, nf=NF_COUNT, border=BOTTOM_HAIR)
        put(ws, f"N{r}", f"=SUM(E{r}:M{r})", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
        merge_put(ws, f"O{r}:P{r}", f'=IF(N{r}=0,"",E{r}/N{r})', al=CENTER, nf="0%", border=BOTTOM_HAIR)
    merge_put(ws, f"B{ov_sum}:D{ov_sum}", "Summe", f=font(10, True), al=LEFT)
    merge_put(ws, f"B{ov_pct}:D{ov_pct}", "Anteil", f=font(9, color=INK2), al=LEFT)
    for c in range(5, 15):
        col = get_column_letter(c)
        put(ws, f"{col}{ov_sum}", f"=SUM({col}{OV_FIRST}:{col}{ov_last})", f=font(10, True), al=CENTER, nf="0", border=TOP_INK)
        put(ws, f"{col}{ov_pct}", f'=IF({total}=0,"",{col}{ov_sum}/{total})', f=font(9, color=INK2), al=CENTER, nf="0%")
    merge_put(ws, f"O{ov_sum}:P{ov_sum}", f'=IF({total}=0,"",E{ov_sum}/{total})', f=font(10, True), al=CENTER, nf="0%", border=TOP_INK)

    place(ws, status_bar_chart(ws, ov_sum, 5, 13, OV_HDR), "B10", f"H{OV_HDR - 3}")
    place(ws, stacked_status_chart(ws, OV_HDR, OV_FIRST, ov_last, 2, 5, 13), "H10", f"R{OV_HDR - 3}")

    # Legende / Erklaerung
    lg = ov_pct + 2
    section_title(ws, f"B{lg}:Q{lg}", "So wird gezählt")
    for code, color, _ in STATUS:
        r = lg + code
        put(ws, f"B{r}", None, fl=fill(color))
        merge_put(ws, f"C{r}:E{r}", f"={stat_label_ref(code)}", f=font(10, True), al=Alignment(horizontal="left", vertical="center", indent=1))
        merge_put(ws, f"F{r}:Q{r}", f"=Einstellungen!$E${E_STAT_FIRST + code - 1}", f=font(9, color=INK2), al=LEFT_WRAP)
        ws.row_dimensions[r].height = 26 if code == 1 else 16
    foot = lg + 11
    merge_put(ws, f"B{foot}:Q{foot}",
              "Bezugsjahr, Zuordnungen und Erklärungen: Blatt „Einstellungen“.   Soll: Blatt „Zielzustand“.   Auswertung je Mitarbeiter: Blatt „Daten“.",
              f=font(9, italic=True, color=MUTED), al=LEFT)
    page_setup(ws, f"A1:Q{foot}", breaks=[OV_HDR - 3])
    fix_row_heights(ws)
    ws.sheet_properties.tabColor = "104281"


# ---------------------------------------------------------------- Blatt: Soll-Ist (alle Bezirke)
SM_COLS = 4                 # kleine Diagramme je Zeile
SM_ROWS = 9


def build_soll_ist(ws, districts, dinfo, ymax):
    setup_grid(ws, hidden=["S", "T"] + HELP_LETTERS + COMP_LETTERS)
    title = ws.title
    n = len(districts)
    header_block(ws, "Soll-Ist-Vergleich LST – Zielzustand je Bezirk",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   Soll = Zielzustand (Blatt „Zielzustand“)   ·   jeder Mitarbeiter zählt einmal – auf seiner '
                 f'erreichten Stufe   ·   „Nach Plan“ = alle geplanten Ausbildungen abgeschlossen"',
                 links=[("Übersicht", "← Übersicht")])
    drows = dinfo["rows"]
    gsheets = ["Grafik " + d.sheet for d in districts]

    sm_bands = math.ceil(n / SM_COLS)
    R_SEC1, R_CH1 = 9, 10
    R_SEC2 = R_CH1 + 21
    R_SM = R_SEC2 + 2
    R_TG = R_SM + sm_bands * (SM_ROWS + 1) + 1          # Tabelle LST gesamt
    TG_HDR = R_TG + 1
    R_TB = TG_HDR + len(CATS) + 6                       # Tabelle je Bezirk (passend besetzt)
    TB_HDR = R_TB + 2
    TB_FIRST = TB_HDR + 1
    TB_LAST = TB_FIRST + n - 1
    TB_SUM = TB_LAST + 1
    R_TQ = TB_SUM + 3                                   # Tabelle je Bezirk und Qualifikation
    TQ_HDR = R_TQ + 2
    TQ_FIRST = TQ_HDR + 1
    TQ_LAST = TQ_FIRST + n - 1

    # Tabelle LST gesamt (Summe der Bezirke)
    section_title(ws, f"B{R_TG}:Q{R_TG}", "Zahlen: LST gesamt – Soll-Ist und Verlauf je Qualifikation")

    def g_sum(row, col):
        return "=" + "+".join(f"{quote_sheetname(g)}!${col}${row}" for g in gsheets)
    count_cols = {c: TL_LETTERS[k] for k, c in enumerate(TL_COLS)}
    count_cols[PLAN_COL] = T_PLAN
    soll = lambda r, cat: g_sum(drows[cat], T_SOLL)
    count = lambda r, col, cat: g_sum(drows[cat], count_cols[col])
    g_rows = soll_ist_table(ws, TG_HDR, soll, count, pass_fn=lambda col: g_sum(drows[ROW_PASS], col))
    g_rt, g_rp = g_rows[ROW_TOTAL], g_rows[ROW_PASS]

    # KPIs (fehlende Stellen als Summe der Bezirke - Bezirke gleichen sich nicht aus)
    tsum = lambda cell: "+".join(f"{quote_sheetname(g)}!${cell[0]}${cell[1:]}" for g in gsheets)
    put(ws, "S6", f"={tsum('T1')}")
    put(ws, "S7", f'="nach Plan · heute: "&({tsum("T2")})')
    put(ws, "S8", f"={tsum('T3')}")
    pct = lambda c: f'=IF({T_SOLL}{g_rp}=0,"–",TEXT({c}{g_rp}/{T_SOLL}{g_rp},"0%")&" des Solls")'
    kpi_tiles(ws, [
        ("Soll (Zielzustand)", f"={T_SOLL}{g_rt}", f'="Ziel-Stellen in "&{n}&" Bezirken"', INK2),
        ("Passend besetzt heute", f"={TL_LETTERS[0]}{g_rp}", pct(TL_LETTERS[0]), C_HEUTE),
        ("Passend besetzt nach Plan", f"={T_PLAN}{g_rp}", pct(T_PLAN), C_PLAN),
        ("Fehlende Ziel-Stellen", "=$S$6", "=$S$7", BAD_INK),
        ("Mehr als Soll (nach Plan)", "=$S$8", f'="davon Arbeiter LST: "&{T_PLAN}{g_rows["Arbeiter LST"]}', "EDA100"),
    ])

    # Tabelle je Bezirk: passend besetzte Ziel-Stellen
    section_title(ws, f"B{R_TB}:Q{R_TB}", "Zahlen: passend besetzte Ziel-Stellen je Bezirk")
    section_note(ws, f"B{R_TB + 1}:Q{R_TB + 1}",
                 "Passend besetzt = Ziel-Stellen, auf denen jemand mit genau dieser Qualifikation sitzt (je Stufe höchstens so viele wie im Soll)")
    merge_put(ws, f"B{TB_HDR}:C{TB_HDR}", "Bezirk (Klick öffnet die Seite)", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    put(ws, f"{T_SOLL}{TB_HDR}", "Soll (Zielzustand)", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    for k, col in enumerate(TL_LETTERS):
        put(ws, f"{col}{TB_HDR}", timeline_label(k), f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    put(ws, f"{T_PLAN}{TB_HDR}", "Nach Plan", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    put(ws, f"{T_DELTA}{TB_HDR}", "Nach Plan − Soll", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    merge_put(ws, f"O{TB_HDR}:Q{TB_HDR}", "Bewertung (nach Plan)", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    ws.row_dimensions[TB_HDR].height = 30
    rp = drows[ROW_PASS]
    for i, d in enumerate(districts):
        r = TB_FIRST + i
        g = quote_sheetname(gsheets[i])
        link = "#" + g.replace('"', '""') + "!A1"
        merge_put(ws, f"B{r}:C{r}", f'=HYPERLINK("{link}","{d.sheet}")', f=font(10, True, "1C5CAB", underline="single"),
                  al=LEFT, border=BOTTOM_HAIR)
        for col in [T_SOLL] + TL_LETTERS + [T_PLAN]:
            put(ws, f"{col}{r}", f"={g}!${col}${rp}", al=CENTER, nf="0", border=BOTTOM_HAIR,
                f=font(10, col in (T_SOLL, T_PLAN)))
        put(ws, f"{T_DELTA}{r}", f"={T_PLAN}{r}-{T_SOLL}{r}", al=CENTER, nf=NF_DELTA, border=BOTTOM_HAIR)
        merge_put(ws, f"O{r}:Q{r}", f'=IF({g}!$T$1=0,"alle Ziel-Stellen passend besetzt",{g}!$T$1&IF({g}!$T$1=1," Ziel-Stelle",'
                  f'" Ziel-Stellen")&" nicht passend besetzt")',
                  f=font(9, color=INK2), al=LEFT, border=BOTTOM_HAIR)
    merge_put(ws, f"B{TB_SUM}:C{TB_SUM}", "Summe", f=font(10, True), al=LEFT, border=TOP_INK)
    for col in [T_SOLL] + TL_LETTERS + [T_PLAN, T_DELTA]:
        put(ws, f"{col}{TB_SUM}", f"=SUM({col}{TB_FIRST}:{col}{TB_LAST})", f=font(10, True), al=CENTER,
            nf=NF_DELTA if col == T_DELTA else "0", border=TOP_INK)
    add_compare_cf(ws, f"{TL_LETTERS[0]}{TB_FIRST}:{T_PLAN}{TB_LAST}", T_SOLL, TB_FIRST)

    # Tabelle je Bezirk und Qualifikation (Soll / nach Plan)
    section_title(ws, f"B{R_TQ}:Q{R_TQ}", "Zahlen: Soll und Bestand nach Plan je Bezirk und Qualifikation")
    section_note(ws, f"B{R_TQ + 1}:Q{R_TQ + 1}", "grün = Soll erreicht · rot = unter Soll")
    merge_put(ws, f"B{TQ_HDR}:C{TQ_HDR}", "Bezirk", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    for j, cat in enumerate(CATS):
        c_s = get_column_letter(4 + 2 * j)
        c_p = get_column_letter(5 + 2 * j)
        put(ws, f"{c_s}{TQ_HDR}", f"{cat}\nSoll", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
        put(ws, f"{c_p}{TQ_HDR}", f"{cat}\nnach Plan", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    ws.row_dimensions[TQ_HDR].height = 42
    for i, d in enumerate(districts):
        r = TQ_FIRST + i
        g = quote_sheetname(gsheets[i])
        merge_put(ws, f"B{r}:C{r}", d.sheet, f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        for j, cat in enumerate(CATS):
            c_s = get_column_letter(4 + 2 * j)
            c_p = get_column_letter(5 + 2 * j)
            put(ws, f"{c_s}{r}", f"={g}!${T_SOLL}${drows[cat]}", al=CENTER, nf="0", border=BOTTOM_HAIR, f=font(10, color=INK2))
            put(ws, f"{c_p}{r}", f"={g}!${T_PLAN}${drows[cat]}", al=CENTER, nf="0", border=BOTTOM_HAIR, f=font(10, True))
    for j in range(len(CATS)):
        c_s = get_column_letter(4 + 2 * j)
        c_p = get_column_letter(5 + 2 * j)
        ws.conditional_formatting.add(f"{c_p}{TQ_FIRST}:{c_p}{TQ_LAST}", FormulaRule(
            formula=[f"{c_p}{TQ_FIRST}<{c_s}{TQ_FIRST}"], fill=fill(BAD_BG), font=Font(color=BAD_INK)))
        ws.conditional_formatting.add(f"{c_p}{TQ_FIRST}:{c_p}{TQ_LAST}", FormulaRule(
            formula=[f"AND({c_s}{TQ_FIRST}>0,{c_p}{TQ_FIRST}>={c_s}{TQ_FIRST})"], fill=fill(GOOD_BG), font=Font(color=GOOD_INK)))

    # Diagramme
    section_title(ws, f"B{R_SEC1}:H{R_SEC1}", "Ziel-Stellen je Bezirk: Soll und passend besetzt")
    section_title(ws, f"I{R_SEC1}:Q{R_SEC1}", "LST gesamt: Mitarbeiter je Stufe – rechts der Zielzustand")
    cats = rng(title, "B", TB_FIRST, "B", TB_LAST)
    si_chart = soll_ist_chart(cats, [
        (rng(title, T_SOLL, TB_FIRST, T_SOLL, TB_LAST), rng(title, T_SOLL, TB_HDR), C_SOLL),
        (rng(title, TL_LETTERS[0], TB_FIRST, TL_LETTERS[0], TB_LAST), rng(title, TL_LETTERS[0], TB_HDR), C_HEUTE),
        (rng(title, T_PLAN, TB_FIRST, T_PLAN, TB_LAST), rng(title, T_PLAN, TB_HDR), C_PLAN),
    ], horizontal=True, label_size=8)
    place(ws, si_chart, f"B{R_CH1}", f"I{R_SEC2 - 1}")
    comp_cats = rng(title, COMP_LETTERS[0], TG_HDR, COMP_LETTERS[-1], TG_HDR)
    place(ws, composition_chart(comp_cats, comp_series(title, g_rows)), f"I{R_CH1}", f"R{R_SEC2 - 1}")

    section_title(ws, f"B{R_SEC2}:Q{R_SEC2}", "Verlauf je Bezirk: passend besetzte Ziel-Stellen")
    section_note(ws, f"B{R_SEC2 + 1}:Q{R_SEC2 + 1}",
                 "Säulen = Ziel-Stellen, die zum Jahresende mit der richtigen Qualifikation besetzt sind (laut Planung) · "
                 "gestrichelte Linie = Soll · alle Diagramme mit gleicher Skala")
    for i, d in enumerate(districts):
        band, pos = divmod(i, SM_COLS)
        top = R_SM + band * (SM_ROWS + 1)
        c0 = 2 + 4 * pos
        g = gsheets[i]
        dc = rng(g, TL_LETTERS[0], SI_HDR, TL_LETTERS[-1], SI_HDR)
        ch = verlauf_chart(dc, rng(g, TL_LETTERS[0], rp, TL_LETTERS[-1], rp), rng(g, "B", rp),
                           rng(g, HELP_LETTERS[0], rp, HELP_LETTERS[-1], rp), rng(g, T_SOLL, SI_HDR),
                           title=d.sheet, ymax=ymax, major=10 if ymax > 20 else 5,
                           labels=False, legend=False, axis_size=7)
        place(ws, ch, f"{get_column_letter(c0)}{top}", f"{get_column_letter(c0 + 4)}{top + SM_ROWS}")

    # Druckseiten: Kennzahlen + Diagramme | Verlauf je Bezirk | Tabellen | Tabelle je Qualifikation
    page_setup(ws, f"A1:Q{TQ_LAST}", breaks=[R_SEC2 - 1, R_TG - 1, R_TQ - 1])
    fix_row_heights(ws)
    ws.sheet_properties.tabColor = "0CA30C"


# ---------------------------------------------------------------- Hauptfunktion
def build(districts, zz_ws, zz_name, out_path, ymax=30):
    global DATA_LAST
    DATA_LAST = 1 + len(districts) * SRC_ROWS
    wb = Workbook()
    ws_over = wb.active
    ws_over.title = "Übersicht"
    ws_si = wb.create_sheet("Soll-Ist")
    dinfo = None
    for d in districts:
        info = build_district(wb.create_sheet("Grafik " + d.sheet), d, None)
        dinfo = dinfo or info
    build_overview(ws_over, districts)
    build_soll_ist(ws_si, districts, dinfo, ymax)
    ws_data = wb.create_sheet("Daten")
    build_data(ws_data, districts)
    ws_data.sheet_properties.tabColor = "C3C2B7"
    ws_zz = wb.create_sheet("Zielzustand")
    build_zielzustand(ws_zz, zz_ws, zz_name)
    ws_zz.sheet_properties.tabColor = "FAB219"
    ws_set = wb.create_sheet("Einstellungen")
    build_settings(ws_set, districts)
    ws_set.sheet_properties.tabColor = "FAB219"
    wb.save(out_path)
    return [ws.title for ws in wb.worksheets]
