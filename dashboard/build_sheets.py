"""Erzeugt die Dashboard-Blaetter als eigenstaendige Arbeitsmappe:

    Uebersicht, Soll-Ist, Pruefliste, Bezirk <Name> (je Bezirk), Daten, Zielzustand, Einstellungen

Die Formeln verweisen auf die Bezirks-Blaetter der Original-Datei (Bedarf_Bestand_LST);
merge_into_original.py setzt die Blaetter anschliessend dort ein. Das Blatt Zielzustand
ist eine Kopie aus LST_Zielzustand_gesamt.xlsx und dient als Soll.

Zaehlregel Soll-Ist / Verlauf (nur Spalten L-P, nicht G/J):
  Stufen von links nach rechts: Azubi (L) -> Arb LST (M) -> Wmech (N) -> SigMech (O) -> SigMech RBEG (P).
  Jeder Mitarbeiter zaehlt genau einmal, auf der rechtesten erreichten Stufe.
  Heute: rechteste Spalte mit "x" ("x" in U = Wmech, "x" in V = SigMech).
  Ende eines Jahres: zusaetzlich jede Jahreszahl <= diesem Jahr (auch ueber das Ziel hinaus).
  Teamleiter ("x" in S) zaehlen zusaetzlich als Teamleiter (also z. B. bei RBEG und als Teamleiter).
Status je Mitarbeiter (gemessen an der Ziel-Qualifikation G):
  Fertig ("x" in der Ziel-Spalte, bei Wmech/SigMech auch "x" in U/V), Abschluss <Jahr>,
  Ueberfaellig (Jahr vorbei, kein "x"), Fehlt (nichts in der Ziel-Spalte), Ziel unklar (G leer/ohne Spalte).
"""
from copy import copy
from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart
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
from openpyxl.worksheet.pagebreak import Break, ColBreak, RowBreak
from openpyxl.worksheet.properties import PageSetupProperties


@dataclass
class District:
    sheet: str      # Blattname in Bedarf_Bestand_LST
    zz_name: str    # Name des Bezirks im Zielzustand (Zeile 1)


# ---------------------------------------------------------------- Konstanten
SRC_FIRST_ROW = 2          # erste gelesene Zeile je Bezirks-Blatt (Kopfzeile "Name" wird uebersprungen)
SRC_ROWS = 60              # gelesene Zeilen je Bezirks-Blatt (2..61)
SRC_LAST_ROW = SRC_FIRST_ROW + SRC_ROWS - 1
N_YEARS = 7                # Verlauf: Bezugsjahr .. Bezugsjahr+6
N_TIMES = N_YEARS + 2      # heute, 7 Jahresenden, nach Plan
LIST_ROWS = 26             # Zeilen der Mitarbeiterliste je Bezirksseite
PRUEF_ROWS = 120           # Zeilen der Pruefliste

FONT = "Aptos Narrow"      # Schrift der Original-Datei
INK = "0B0B0B"
INK2 = "52514E"
MUTED = "898781"
HAIR = "E1E0D9"
TILE_BG = "F4F4F2"
HEAD_BG = "EDEDEA"
INPUT_BG = "FFF2CC"
WHITE = "FFFFFF"
LINK = "1C5CAB"
GOOD_BG, GOOD_INK = "E2F3E2", "0A6B0A"
BAD_BG, BAD_INK = "F9E0E0", "A12A2A"
MORE_BG, MORE_INK = "FFF1D6", "8A5A00"
C_SOLL, C_HEUTE, C_PLAN = INK2, "86B6EF", "2A78D6"
C_GAP = "D03B3B"

# Stufen in L-P (Nr 1..5), 0 = keine Stufe
STAGES = ["Azubi", "Arbeiter LST", "Wmech", "SigMech", "SigMech RBEG"]
STAGE_COLS = ["L", "M", "N", "O", "P"]
NO_STAGE = "keine Stufe"
TL = "Teamleiter"
# Vergleich mit dem Zielzustand (Zeilenbeschriftungen in Spalte A dort)
CMP = ["Arbeiter LST", "Wmech", "SigMech", "SigMech RBEG", TL]
GROUP_OTHER = "Sonstige / unklar"
GROUPS = CMP + [GROUP_OTHER]                     # Zeilen der Status-Tabelle (Ziel-Qualifikation)
# Farben: Stufen als Blau-Rampe hell -> dunkel, Teamleiter als eigener Farbton
STAGE_COLORS = {NO_STAGE: ("C3C2B7", INK), "Azubi": ("86B6EF", INK), "Arbeiter LST": ("5598E7", INK),
                "Wmech": ("2A78D6", WHITE), "SigMech": ("1C5CAB", WHITE), "SigMech RBEG": ("104281", WHITE),
                TL: ("EB6834", INK)}
STACK = [NO_STAGE] + STAGES + [TL]               # Reihenfolge im gestapelten Diagramm (unten -> oben)

# Status-Nr -> (Farbe, Textfarbe). Bezeichnungen stehen im Blatt Einstellungen.
STATUS = [
    (1, "0CA30C", INK),       # Fertig
    (2, "104281", WHITE),     # Abschluss Bezugsjahr
    (3, "1C5CAB", WHITE),     # Abschluss +1
    (4, "2A78D6", WHITE),     # +2
    (5, "5598E7", INK),       # +3
    (6, "86B6EF", INK),       # ab +4
    (7, "EC835A", INK),       # Ueberfaellig
    (8, "D03B3B", WHITE),     # Fehlt
    (9, "898781", WHITE),     # Ziel unklar
]
STATUS_TEXT = {
    1: ("Fertig", "„x“ in der Spalte der Ziel-Qualifikation (L–S). Bei Wmech und SigMech reicht auch „x“ bei der örtlichen "
                  "Verwendungsprüfung (U bzw. V)."),
    2: (None, "In Ausbildung: In der Spalte der Ziel-Qualifikation steht das Bezugsjahr (Abschluss in diesem Jahr geplant)."),
    3: (None, "In Ausbildung: geplanter Abschluss im Bezugsjahr + 1."),
    4: (None, "In Ausbildung: geplanter Abschluss im Bezugsjahr + 2."),
    5: (None, "In Ausbildung: geplanter Abschluss im Bezugsjahr + 3."),
    6: (None, "In Ausbildung: geplanter Abschluss im Bezugsjahr + 4 oder später."),
    7: ("Überfällig", "Das Jahr in der Spalte der Ziel-Qualifikation ist vorbei, aber es steht noch kein „x“ – bitte prüfen."),
    8: ("Fehlt (nichts geplant)", "In der Spalte der Ziel-Qualifikation steht weder „x“ noch ein Jahr."),
    9: ("Ziel unklar", "Ziel-Qualifikation (G) ist leer oder hat keine Spalte in L–S (z. B. „Senior Expert LST“). "
                       "Zuordnung im Blatt Einstellungen ergänzen."),
}

# Zuordnung Text in Ziel-/Ist-Spalte -> Spalte L..S, Verwendungspruefung U/V, Gruppe (Zeile im Zielzustand)
MAPPING = [
    ("Azubi", "L", "", "Azubi"),
    ("Arb LST", "M", "", "Arbeiter LST"),
    ("Arbeiter LST", "M", "", "Arbeiter LST"),
    ("Weichmech", "N", "U", "Wmech"),
    ("Weichenmechaniker", "N", "U", "Wmech"),
    ("Wmech", "N", "U", "Wmech"),
    ("Sigmech", "O", "V", "SigMech"),
    ("Signalmechaniker", "O", "V", "SigMech"),
    ("Sigmech RBEG", "P", "", "SigMech RBEG"),
    ("Signalmechaniker RBEG", "P", "", "SigMech RBEG"),
    ("Kennziffer 4", "Q", "", GROUP_OTHER),
    ("IHK-Meister", "R", "", GROUP_OTHER),
    ("Teamleiter", "S", "", TL),
    ("TL", "S", "", TL),
    ("Senior Expert LST", "", "", GROUP_OTHER),
    ("Umschüler EBET", "", "", GROUP_OTHER),
    ("Umschüler", "", "", GROUP_OTHER),
    ("Quereinsteiger", "", "", GROUP_OTHER),
]

# Einstellungen: feste Zellpositionen
E_YEAR = "Einstellungen!$C$4"
E_DAVON = "Einstellungen!$C$5"
E_STAT_HDR, E_STAT_FIRST = 8, 9
E_STAT_LAST = E_STAT_FIRST + len(STATUS) - 1
E_BEZ_HDR, E_BEZ_FIRST, E_BEZ_LAST = 21, 22, 41
E_MAP_HDR, E_MAP_FIRST, E_MAP_LAST = 45, 46, 85
E_STAT_LABEL = f"Einstellungen!$C${E_STAT_FIRST}:$C${E_STAT_LAST}"
E_BEZ_SHEET = f"Einstellungen!$B${E_BEZ_FIRST}:$B${E_BEZ_LAST}"
E_BEZ_ZZ = f"Einstellungen!$C${E_BEZ_FIRST}:$C${E_BEZ_LAST}"
E_BEZ_ZZCOL = f"Einstellungen!$D${E_BEZ_FIRST}:$D${E_BEZ_LAST}"
E_MAP_TEXT = f"Einstellungen!$B${E_MAP_FIRST}:$B${E_MAP_LAST}"
E_MAP_COLNO = f"Einstellungen!$D${E_MAP_FIRST}:$D${E_MAP_LAST}"
E_MAP_VPNO = f"Einstellungen!$F${E_MAP_FIRST}:$F${E_MAP_LAST}"
E_MAP_GRP = f"Einstellungen!$G${E_MAP_FIRST}:$G${E_MAP_LAST}"
ZZ_AREA = "Zielzustand!$A$1:$CZ$60"
ZZ_LABELS = "Zielzustand!$A$1:$A$60"
ZZ_HEAD = "Zielzustand!$A$1:$CZ$1"


def stat_label_ref(code):
    return f"Einstellungen!$C${E_STAT_FIRST + code - 1}"


def stage_name_formula(nr_ref):
    names = ",".join(f'"{n}"' for n in [NO_STAGE] + STAGES)
    return f"CHOOSE({nr_ref}+1,{names})"


# Zeitpunkte: 0 = heute (nur "x"), 1..7 = Ende Bezugsjahr+k-1, 8 = nach Plan (alle Jahre)
def threshold(k):
    if k == 0:
        return "0"
    if k <= N_YEARS:
        return f"({E_YEAR}+{k - 1})"
    return "9998"


def time_label(k):
    if k == 0:
        return "heute"
    if k <= N_YEARS:
        return f'=""&({E_YEAR}+{k - 1})'
    return "nach Plan"


def time_short(k):
    if k == 0:
        return "heute"
    if k <= N_YEARS:
        return f'=RIGHT(""&({E_YEAR}+{k - 1}),2)'
    return "Plan"


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
TOP_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
NF_COUNT = '0;-0;"·"'


def put(ws, ref, value, f=None, fl=None, al=None, nf=None, border=None):
    c = ws[ref]
    c.value = value
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


def merge_put(ws, rng_, value, **kw):
    first = rng_.split(":")[0]
    c = put(ws, first, value, **kw)
    ws.merge_cells(rng_)
    if kw.get("fl") or kw.get("border"):
        for row in ws[rng_]:
            for cell in row:
                if kw.get("fl"):
                    cell.fill = kw["fl"]
                if kw.get("border"):
                    cell.border = kw["border"]
    return c


def page_setup(ws, print_area, breaks=(), title_rows=None):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_area = print_area
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.45
    ws.page_margins.header = ws.page_margins.footer = 0.2
    ws.oddFooter.right.text = "Seite &P von &N"
    ws.oddFooter.right.size = 8
    ws.oddFooter.left.text = "&A"
    ws.oddFooter.left.size = 8
    if title_rows:
        ws.print_title_rows = title_rows
    if breaks:
        ws.row_breaks = RowBreak()
        for r in breaks:
            ws.row_breaks.append(Break(id=r))


def heights(ws, mapping):
    for r, h in mapping.items():
        ws.row_dimensions[r].height = h


def fix_row_heights(ws, height=15):
    """Feste Zeilenhoehe fuer alle Zeilen ohne eigene Hoehe (die Original-Datei hat eine andere Standardschrift)."""
    for r in range(1, ws.max_row + 2):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = height


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


def style_frame(chart):
    chart.graphical_properties = GraphicalProperties(ln=no_line())
    chart.roundedCorners = False
    chart.visible_cells_only = False      # Hilfswerte stehen in ausgeblendeten Spalten


def cat_axis(ax, reverse=False, size=9):
    if reverse:
        ax.scaling.orientation = "maxMin"
    ax.delete = False
    ax.majorTickMark = "none"
    ax.txPr = text_props(size, INK2)
    ax.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="C3C2B7"))


def hide_value_axis(ax):
    ax.delete = True
    ax.majorGridlines = None
    ax.scaling.min = 0


def show_value_axis(ax, ymax=None, major=None, size=8):
    ax.delete = False
    ax.scaling.min = 0
    if ymax:
        ax.scaling.max = ymax
    if major:
        ax.majorUnit = major
    ax.numFmt = "0"
    ax.majorTickMark = "none"
    ax.txPr = text_props(size, MUTED)
    ax.graphicalProperties = GraphicalProperties(ln=no_line())
    ax.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=HAIR, w=6350)))


def legend(chart, size=9, pos="b"):
    chart.legend = Legend()
    chart.legend.position = pos
    chart.legend.txPr = text_props(size, INK2)


def ref_series(idx, values_ref, title_ref, cats_ref):
    s = Series(idx=idx, order=idx)
    s.val = NumDataSource(numRef=NumRef(f=values_ref))
    s.tx = SeriesLabel(strRef=StrRef(f=title_ref))
    s.cat = AxDataSource(strRef=StrRef(f=cats_ref))
    return s


def fill_series(s, color, gap_line=False):
    ln = LineProperties(solidFill=WHITE, w=12700) if gap_line else no_line()
    s.graphicalProperties = GraphicalProperties(solidFill=color, ln=ln)
    s.invertIfNegative = False
    return s


def rng(ws_title, c1, r1, c2=None, r2=None):
    return f"{quote_sheetname(ws_title)}!${c1}${r1}:${c2 or c1}${r2 or r1}"


def status_bar_chart(title, val_row, c1, c2, cat_row):
    """Ein Balken je Status, jeweils in der Status-Farbe (Bezeichnung steht an der Achse)."""
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "clustered"
    ch.style = 2
    L1, L2 = get_column_letter(c1), get_column_letter(c2)
    s = ref_series(0, rng(title, L1, val_row, L2, val_row), rng(title, "B", val_row), rng(title, L1, cat_row, L2, cat_row))
    fill_series(s, STATUS[0][1])
    for i, (_, color, _) in enumerate(STATUS):
        pt = DataPoint(idx=i, invertIfNegative=False)
        pt.graphicalProperties = GraphicalProperties(solidFill=color, ln=no_line())
        s.dPt.append(pt)
    s.dLbls = data_labels(INK, pos="outEnd", numfmt="0;-0;0", bold=True)
    ch.series.append(s)
    ch.gapWidth = 35
    ch.legend = None
    cat_axis(ch.x_axis, reverse=True)
    hide_value_axis(ch.y_axis)
    style_frame(ch)
    return ch


def stacked_status_chart(title, hdr_row, first_row, last_row, c1, c2):
    """Gestapelte Balken: je Zeile (Bezirk / Ziel) die Anzahl je Status."""
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.style = 2
    cats = rng(title, "B", first_row, "B", last_row)
    for i, (code, color, txt) in enumerate(STATUS):
        col = get_column_letter(c1 + i)
        s = fill_series(ref_series(i, rng(title, col, first_row, col, last_row), rng(title, col, hdr_row), cats), color, True)
        s.dLbls = data_labels(txt, pos="ctr", numfmt="0;-0;;", size=8)
        ch.series.append(s)
    ch.gapWidth = 45
    legend(ch, 8)
    cat_axis(ch.x_axis, reverse=True)
    hide_value_axis(ch.y_axis)
    style_frame(ch)
    return ch


def clustered_chart(cats_ref, series, horizontal=False, label_size=9, legend_size=9):
    """Gruppierte Saeulen/Balken. series: Liste (Werte-Bereich, Titel-Zelle, Farbe)."""
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
    legend(ch, legend_size)
    cat_axis(ch.x_axis, reverse=horizontal)
    hide_value_axis(ch.y_axis)
    style_frame(ch)
    return ch


def stacked_stage_chart(cats_ref, series, ymax=None, labels=True, show_legend=True, axis_size=9, title=None,
                        label_fmt="0;-0;;"):
    """Gestapelte Saeulen je Zeitpunkt (Mitarbeiter je Stufe) + Zielzustand. series: (Werte, Titel, Farbe, Textfarbe)."""
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.style = 2
    for i, (values_ref, title_ref, color, txt) in enumerate(series):
        s = fill_series(ref_series(i, values_ref, title_ref, cats_ref), color, True)
        if labels:
            s.dLbls = data_labels(txt, pos="ctr", numfmt=label_fmt, size=8)
        ch.series.append(s)
    ch.gapWidth = 45
    if show_legend:
        legend(ch, 8)
    else:
        ch.legend = None
    cat_axis(ch.x_axis, size=axis_size)
    if ymax:
        show_value_axis(ch.y_axis, ymax=ymax, major=10 if ymax > 30 else 5, size=7)
    else:
        hide_value_axis(ch.y_axis)
    if title:
        ch.title = chart_title(title, 9)
    style_frame(ch)
    return ch


def verlauf_chart(cats_ref, values_ref, values_title_ref, soll_ref=None, soll_title_ref=None, title=None, color=C_PLAN):
    """Saeulen = Anzahl je Zeitpunkt, gestrichelte Linie = Soll (beides auf derselben Achse)."""
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.style = 2
    s = fill_series(ref_series(0, values_ref, values_title_ref, cats_ref), color)
    s.dLbls = data_labels(INK2, pos="outEnd", numfmt="0;-0;0", size=7)
    bar.series.append(s)
    bar.gapWidth = 50
    cat_axis(bar.x_axis, size=7)
    hide_value_axis(bar.y_axis)
    if soll_ref:
        line = LineChart()
        ls = ref_series(1, soll_ref, soll_title_ref, cats_ref)
        ls.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=INK, w=19050, prstDash="dash"))
        ls.marker = Marker(symbol="none")
        ls.smooth = False
        line.series.append(ls)
        line.x_axis = bar.x_axis
        line.y_axis = bar.y_axis
        bar += line
    bar.legend = None
    if title:
        bar.title = chart_title(title, 9)
    style_frame(bar)
    return bar


def place(ws, chart, top_left, bottom_right):
    """Diagramm genau auf einen Zellbereich legen (linke obere Ecke von top_left bis linke obere Ecke von bottom_right)."""
    c1, r1 = coordinate_from_string(top_left)
    c2, r2 = coordinate_from_string(bottom_right)
    chart.anchor = TwoCellAnchor(_from=AnchorMarker(col=column_index_from_string(c1) - 1, row=r1 - 1),
                                 to=AnchorMarker(col=column_index_from_string(c2) - 1, row=r2 - 1))
    ws.add_chart(chart)


# ---------------------------------------------------------------- Bausteine (Raster B..S = 18 Spalten)
GRID_COLS = 18
GRID_WIDTH = 9.3
GRID_LAST = get_column_letter(1 + GRID_COLS)     # S
KEY_COL = "U"                                    # ausgeblendet: Schluessel (Bezirk) / Hilfswerte


def setup_grid(ws, hidden_from=None):
    ws.column_dimensions["A"].width = 2
    for i in range(2, 2 + GRID_COLS):
        ws.column_dimensions[get_column_letter(i)].width = GRID_WIDTH
    ws.column_dimensions["T"].width = 2
    if hidden_from:
        start = column_index_from_string(hidden_from)
        for i in range(start, start + 80):
            ws.column_dimensions[get_column_letter(i)].hidden = True
    heights(ws, {1: 8, 4: 8})


def nav_link(ws, rng_, target, text):
    link = "#" + quote_sheetname(target).replace('"', '""') + "!A1"
    merge_put(ws, rng_, f'=HYPERLINK("{link}","{text}")', f=font(10, color=LINK, underline="single"), al=RIGHT)


def header_block(ws, title, subtitle, links=()):
    merge_put(ws, "B2:L2", title, f=font(18, True), al=LEFT)
    ws.row_dimensions[2].height = 30
    cells = {1: ["P2:S2"], 2: ["M2:O2", "P2:S2"], 3: ["J2:L2", "M2:O2", "P2:S2"]}.get(len(links), [])
    for c, (target, text) in zip(cells, links):
        nav_link(ws, c, target, text)
    merge_put(ws, f"B3:{GRID_LAST}3", subtitle, f=font(9, color=INK2), al=LEFT)
    ws.row_dimensions[3].height = 16


def kpi_tiles(ws, tiles, row=5):
    """tiles: Liste (Titel, Wert-Formel, Unterzeile-Formel, Akzentfarbe); je 3 Spalten ab B (6 Kacheln = B..S)."""
    heights(ws, {row: 18, row + 1: 34, row + 2: 16})
    for i, (label, value, sub, accent) in enumerate(tiles):
        c1 = 2 + 3 * i
        c3 = c1 + 2
        L1, L3 = get_column_letter(c1), get_column_letter(c3)
        merge_put(ws, f"{L1}{row}:{L3}{row}", label, f=font(10, True, INK2),
                  al=Alignment(horizontal="left", vertical="bottom", indent=1))
        merge_put(ws, f"{L1}{row+1}:{L3}{row+1}", value, f=font(24, True, INK),
                  al=Alignment(horizontal="left", vertical="center", indent=1), nf="0")
        merge_put(ws, f"{L1}{row+2}:{L3}{row+2}", sub, f=font(9, color=MUTED),
                  al=Alignment(horizontal="left", vertical="top", indent=1))
        accent_side = Side(style="thick", color=accent)
        for r in range(row, row + 3):
            for c in range(c1, c3 + 1):
                cell = ws.cell(r, c)
                cell.fill = fill(TILE_BG)
                cell.border = Border(left=accent_side if c == c1 else None,
                                     right=Side(style="thick", color=WHITE) if c == c3 else None)


def section_title(ws, rng_, text, row_height=20):
    merge_put(ws, rng_, text, f=font(12, True), al=LEFT)
    r = int("".join(ch for ch in rng_.split(":")[0] if ch.isdigit()))
    ws.row_dimensions[r].height = row_height


def section_note(ws, rng_, text, height=None, size=9):
    merge_put(ws, rng_, text, f=font(size, color=MUTED), al=LEFT_WRAP)
    if height:
        r = int("".join(ch for ch in rng_.split(":")[0] if ch.isdigit()))
        ws.row_dimensions[r].height = height


def how_box(ws, row, text, height=30):
    """Kurze Erklaerung „So wird gezählt“ (hellgrau hinterlegt, ueber die ganze Breite)."""
    merge_put(ws, f"B{row}:{GRID_LAST}{row}", text, f=font(9, color=INK2), fl=fill("F7F7F5"),
              al=Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1))
    ws.row_dimensions[row].height = height


def head(ws, ref, text, align=CENTER_WRAP):
    if ":" in ref:
        return merge_put(ws, ref, text, f=font(9, True, INK2), fl=fill(HEAD_BG), al=align)
    return put(ws, ref, text, f=font(9, True, INK2), fl=fill(HEAD_BG), al=align)


def status_header_cells(ws, row, first_col):
    for code in range(1, len(STATUS) + 1):
        head(ws, f"{get_column_letter(first_col + code - 1)}{row}", f"={stat_label_ref(code)}")


def pct_sub(count_ref, total_ref, what="Mitarbeitern"):
    return f'=IF({total_ref}=0,"–",TEXT({count_ref}/{total_ref},"0%")&" von "&{total_ref}&" {what}")'


def gap_text(fehlt, mehr):
    """Bewertung: „Soll erreicht“ / „es fehlen X“ / „X mehr als Soll“ (bei Summen ueber Bezirke ggf. beides)."""
    return (f'=IF(AND({fehlt}=0,{mehr}=0),"Soll erreicht",IF({fehlt}=1,"es fehlt 1",IF({fehlt}>1,"es fehlen "&{fehlt},""))'
            f'&IF(AND({fehlt}>0,{mehr}>0)," · ","")&IF({mehr}>0,{mehr}&" mehr als Soll",""))')


def add_gap_cf(ws, cell_range, first_cell):
    """Text-Bewertung einfaerben: gruen = Soll erreicht, rot = es fehlen, gelb = nur mehr als Soll."""
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f'LEFT({first_cell},3)="es "'], fill=fill(BAD_BG), font=Font(color=BAD_INK)))
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f'ISNUMBER(SEARCH("mehr",{first_cell}))'], fill=fill(MORE_BG), font=Font(color=MORE_INK)))
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f'{first_cell}="Soll erreicht"'], fill=fill(GOOD_BG), font=Font(color=GOOD_INK)))


def add_compare_cf(ws, cell_range, first_cell, soll_cell):
    """Zahl gegen Soll: rot = darunter, gelb = darueber, gruen = gleich."""
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{first_cell}<{soll_cell}"], fill=fill(BAD_BG), font=Font(color=BAD_INK)))
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{first_cell}>{soll_cell}"], fill=fill(MORE_BG), font=Font(color=MORE_INK)))
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{first_cell}={soll_cell}"], fill=fill(GOOD_BG), font=Font(color=GOOD_INK)))


HOW_STATUS = ("So wird gezählt (Status): Maßgeblich ist die Ziel-Qualifikation (Spalte G) und deren Spalte in L–S. "
              "Fertig = dort steht „x“ (bei Wmech/SigMech reicht auch „x“ bei der örtlichen Verwendungsprüfung U/V). "
              "Abschluss <Jahr> = dort steht ein geplantes Jahr (27 = 2027). Überfällig = das Jahr ist vorbei, aber noch kein „x“. "
              "Fehlt = dort steht nichts. Ziel unklar = Ziel leer oder ohne Spalte.")
HOW_COUNT = ("So wird gezählt (Soll-Ist): Nur die Spalten L–P zählen. Jeder Mitarbeiter zählt einmal – auf der rechtesten erreichten "
             "Stufe: Azubi (L) → Arbeiter LST (M) → Wmech (N) → SigMech (O) → SigMech RBEG (P). Heute zählt nur „x“ "
             "(„x“ in U = Wmech, „x“ in V = SigMech). Zum Jahresende zählt zusätzlich jede Jahreszahl bis zu diesem Jahr, "
             "„nach Plan“ zählt alle Jahre. Teamleiter („x“ in S) zählen zusätzlich als Teamleiter. Soll = Spalte „Zielzustand“.")


HOW_SHORT = ("So wird gezählt – Status (gemessen an der Ziel-Qualifikation G): Fertig = „x“ in deren Spalte (Wmech/SigMech auch „x“ in "
             "U/V) · Abschluss <Jahr> = dort steht ein Jahr · Überfällig = Jahr vorbei ohne „x“ · Fehlt = dort steht nichts.   "
             "Soll-Ist (nur L–P): jeder zählt einmal auf seiner rechtesten erreichten Stufe – heute nur „x“, zum Jahresende auch "
             "Jahre bis dahin; Teamleiter („x“ in S) zusätzlich. Soll = Zielzustand.")


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
    put(ws, "A1", "Zielzustand (Soll)", f=font(11, True))
    # Druck: A4 quer, Spalte A auf jeder Seite, je Seite LST gesamt bzw. 3 Bezirke (Umbruch vor jedem 3. Bezirk)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
    ws.page_setup.scale = 85
    ws.print_title_cols = "A:A"
    starts = [c.column for c in src_ws[1] if c.value and c.column > 1]
    ws.col_breaks = ColBreak()
    for i, c in enumerate(starts):
        if i == 1 or (i > 1 and (i - 1) % 3 == 0):
            ws.col_breaks.append(Break(id=c - 1))
    r = src_ws.max_row + 2
    notes = [
        f"Quelle: {src_name}, Blatt „{src_ws.title}“ (unverändert übernommen). Hier kann das Soll geändert werden.",
        "Verwendet wird je Bezirk die Spalte „Zielzustand“, Zeilen Arbeiter LST, Wmech, SigMech, SigMech RBEG und Teamleiter. "
        "FBÜB und Azubi-Zeilen werden nicht verglichen.",
        "Zahlen hier ändern → Übersicht, Soll-Ist und Bezirksseiten passen sich automatisch an. "
        "Zuordnung Bezirks-Blatt → Bezirk im Zielzustand: Blatt „Einstellungen“.",
    ]
    for i, n in enumerate(notes):
        put(ws, f"A{r + i}", n, f=font(9, italic=True, color=INK2))
    ws.print_area = f"A1:{get_column_letter(src_ws.max_column)}{r + len(notes) - 1}"


# ---------------------------------------------------------------- Blatt: Einstellungen
def build_settings(ws, districts):
    ws.sheet_view.showGridLines = False
    widths = {"A": 2, "B": 28, "C": 24, "D": 12, "E": 14, "F": 12, "G": 18, "H": 60}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    put(ws, "B1", "Einstellungen & Erklärung", f=font(16, True))
    put(ws, "B2", "Gelb hinterlegte Zellen dürfen geändert werden. Alle Zahlen und Diagramme passen sich automatisch an.",
        f=font(10, color=INK2))

    put(ws, "B4", "Bezugsjahr", f=font(10, True), al=LEFT)
    put(ws, "C4", "=YEAR(TODAY())", f=font(11, True, "0000FF"), fl=fill(INPUT_BG), al=CENTER, nf="0")
    merge_put(ws, "D4:H4", "Standard: aktuelles Jahr (=JAHR(HEUTE())). Kann mit einer festen Zahl überschrieben werden, z. B. 2027. "
                           "Der Verlauf zeigt das Bezugsjahr und die 6 folgenden Jahre; geplante Jahre davor gelten als „Überfällig“.",
              f=font(9, color=INK2), al=LEFT_WRAP)
    put(ws, "B5", "Jahr für „In Ausbildung, davon …“", f=font(10, True), al=LEFT)
    put(ws, "C5", "=C4+1", f=font(11, True, "0000FF"), fl=fill(INPUT_BG), al=CENTER, nf="0")
    merge_put(ws, "D5:H5", "Kennzahl „In Ausbildung“: wie viele davon ihren Abschluss in diesem Jahr geplant haben. "
                           "Standard: Bezugsjahr + 1.", f=font(9, color=INK2), al=LEFT_WRAP)
    heights(ws, {4: 28, 5: 28})

    put(ws, f"B{E_STAT_HDR - 1}", "Status je Mitarbeiter", f=font(12, True))
    for col, text in zip("BCDE", ["Nr.", "Bezeichnung", "Farbe", "Bedeutung"]):
        put(ws, f"{col}{E_STAT_HDR}", text, f=font(10, True, INK2), fl=fill(HEAD_BG), al=LEFT)
    for col in "FGH":
        put(ws, f"{col}{E_STAT_HDR}", None, fl=fill(HEAD_BG))
    for code, color, _ in STATUS:
        r = E_STAT_FIRST + code - 1
        label, meaning = STATUS_TEXT[code]
        if label is None:
            off = code - 2
            label = f'="Abschluss "&({E_YEAR}+{off})' if code < 6 else f'="Abschluss ab "&({E_YEAR}+{off})'
        put(ws, f"B{r}", code, al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"C{r}", label, f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"D{r}", None, fl=fill(color), border=Border(bottom=Side(style="thin", color=WHITE)))
        merge_put(ws, f"E{r}:H{r}", meaning, f=font(9, color=INK2), al=LEFT_WRAP, border=BOTTOM_HAIR)
        ws.row_dimensions[r].height = 26 if len(meaning) > 100 else 16

    put(ws, f"B{E_BEZ_HDR - 2}", "Bezirke: Blatt → Bezirk im Zielzustand", f=font(12, True))
    put(ws, f"B{E_BEZ_HDR - 1}", "Name genau wie in Zeile 1 des Blatts „Zielzustand“. Das Soll steht in der Spalte „Zielzustand“, "
                                 "3 Spalten rechts vom Namen.", f=font(9, color=INK2))
    for col, text in zip("BCDEF", ["Bezirks-Blatt", "Name im Zielzustand", "Spalte Nr. (auto)", "Spalte Soll (auto)", "Prüfung"]):
        put(ws, f"{col}{E_BEZ_HDR}", text, f=font(10, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    ws.row_dimensions[E_BEZ_HDR].height = 28
    for i in range(E_BEZ_LAST - E_BEZ_FIRST + 1):
        r = E_BEZ_FIRST + i
        d = districts[i] if i < len(districts) else None
        put(ws, f"B{r}", d.sheet if d else None, f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"C{r}", d.zz_name if d else None, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"D{r}", f'=IF(C{r}="","",IFERROR(MATCH(C{r},{ZZ_HEAD},0)+3,""))', al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"E{r}", f'=IF(D{r}="","",SUBSTITUTE(ADDRESS(1,D{r},4),"1",""))', al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"F{r}", f'=IF(B{r}="","",IF(D{r}="","nicht gefunden – kein Soll","ok"))', al=LEFT, border=BOTTOM_HAIR)

    put(ws, f"B{E_MAP_HDR - 2}", "Zuordnung Qualifikation → Spalte (L–S), örtl. Verwendungsprüfung (U/V) und Gruppe", f=font(12, True))
    put(ws, f"B{E_MAP_HDR - 1}", "Steht in Spalte G (Ziel) oder J (Ist) ein neuer Begriff, hier eine Zeile ergänzen. Ohne Spalte "
                                 "bekommt der Mitarbeiter den Status „Ziel unklar“ (z. B. Senior Expert LST – bei Bedarf hier "
                                 "eine Spalte eintragen).", f=font(9, color=INK2))
    for col, text in zip("BCDEFGH", ["Text in Spalte G / J", "Spalte mit dem Stand (L–S)", "Nr. (auto)",
                                     "Spalte örtl. Verwendungs­prüfung (U/V)", "Nr. (auto)", "Gruppe (Zeile im Zielzustand)",
                                     "Hinweis"]):
        put(ws, f"{col}{E_MAP_HDR}", text, f=font(10, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    ws.row_dimensions[E_MAP_HDR].height = 42
    hints = {"Azubi": "Texte, die mit „Azubi“ beginnen (z. B. „Azubi 2025“), gelten automatisch als Azubi",
             "Kennziffer 4": "Zusatzqualifikation, i. d. R. kein Ziel",
             "Senior Expert LST": "keine eigene Spalte → Status „Ziel unklar“",
             "Weichmech": "fertig bei „x“ in N oder „x“ bei der örtl. Verwendungsprüfung Wmech (U)",
             "Sigmech": "fertig bei „x“ in O oder „x“ bei der örtl. Verwendungsprüfung SigMech (V)"}
    for i in range(E_MAP_LAST - E_MAP_FIRST + 1):
        r = E_MAP_FIRST + i
        text, col, vp, grp = MAPPING[i] if i < len(MAPPING) else (None, None, None, None)
        put(ws, f"B{r}", text, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"C{r}", col or None, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"D{r}", f'=IF(C{r}="","",IFERROR(MATCH(UPPER(C{r}),{{"L","M","N","O","P","Q","R","S"}},0),""))',
            al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"E{r}", vp or None, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"F{r}", f'=IF(E{r}="","",IFERROR(MATCH(UPPER(E{r}),{{"U","V"}},0),""))', al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"G{r}", grp, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"H{r}", hints.get(text), f=font(9, color=INK2), al=LEFT, border=BOTTOM_HAIR)

    r = E_MAP_LAST + 3
    put(ws, f"B{r}", "So wird gezählt", f=font(12, True))
    notes = [
        f"Gelesen werden die Bezirks-Blätter aus der Tabelle „Bezirke“ oben, jeweils Zeile {SRC_FIRST_ROW}–{SRC_LAST_ROW}. "
        "Jede Zeile mit einem Namen in Spalte B zählt als Mitarbeiter (die Kopfzeile „Name“ wird übersprungen).",
        "Einträge in L–S und U/V: „x“ = erledigt, Zahl = geplantes Jahr (27 = 2027, auch 2027 oder ein Datum). "
        "Alles andere (z. B. „0“) gilt als ungültig und steht in der Prüfliste.",
        "Status je Mitarbeiter: gemessen an der Ziel-Qualifikation (G) und ihrer Spalte in L–S. Fertig = „x“ dort, bei Wmech/SigMech "
        "auch „x“ bei der örtl. Verwendungsprüfung (U/V). Abschluss <Jahr> = Jahr dort. Überfällig = Jahr vor dem Bezugsjahr ohne „x“. "
        "Fehlt = dort steht nichts. Ziel unklar = G leer oder ohne Spalte.",
        "Soll-Ist und Verlauf: Nur die Spalten L–P zählen (G und J nicht). Jeder Mitarbeiter zählt genau einmal – auf der rechtesten "
        "erreichten Stufe: Azubi (L) → Arbeiter LST (M) → Wmech (N) → SigMech (O) → SigMech RBEG (P). Wer RBEG ist, ist kein SigMech mehr.",
        "Heute = rechteste Spalte in L–P mit „x“; „x“ in U zählt als Wmech erreicht, „x“ in V als SigMech erreicht. Ende eines Jahres = "
        "rechteste Spalte mit „x“ oder mit einer Jahreszahl bis zu diesem Jahr (auch über die Ziel-Qualifikation hinaus; überfällige "
        "Jahre zählen dabei schon im ersten Jahr). „Nach Plan“ = alle eingetragenen Jahre erreicht.",
        "Teamleiter werden zusätzlich über Spalte S gezählt (heute „x“, zum Jahresende auch ein Jahr bis dahin). Ein Teamleiter mit "
        "„x“ in P zählt also bei SigMech RBEG und zusätzlich als Teamleiter – auch in den gestapelten Säulen.",
        "Soll = Spalte „Zielzustand“ im Blatt Zielzustand (Zeilen Arbeiter LST, Wmech, SigMech, SigMech RBEG, Teamleiter). "
        "Je Stufe: „Soll erreicht“, „es fehlen X“ oder „X mehr als Soll“. Fehlende Stellen zum Soll = Summe der fehlenden "
        "Mitarbeiter über alle Stufen (für LST gesamt: Summe über die Bezirke – ein Überschuss in einem Bezirk gleicht "
        "keine Lücke in einem anderen aus).",
        "Das Blatt „Daten“ enthält die Auswertung je Mitarbeiter. Es wird per Formel erzeugt und sollte nicht von Hand geändert werden.",
    ]
    for i, n in enumerate(notes):
        merge_put(ws, f"B{r + 1 + i}:H{r + 1 + i}", "• " + n, f=font(10, color=INK2), al=LEFT_WRAP)
        ws.row_dimensions[r + 1 + i].height = 44
    page_setup(ws, f"A1:H{r + len(notes)}", breaks=[E_BEZ_HDR - 3, E_MAP_HDR - 3])


# ---------------------------------------------------------------- Blatt: Daten
DATA_SPEC = [
    ("bez", "Bezirk", 16), ("row", "Zeile im Bezirks-Blatt", 8), ("name", "Name", 16), ("vor", "Vorname", 14),
    ("ziel", "Ziel-Qualifikation (G)", 16), ("ist", "Ist-Qualifikation (J)", 16), ("ok", "Mitarbeiter (1 = Name vorhanden)", 10),
    ("zmap", "Ziel: Zeile in Zuordnung", 9), ("zcol", "Ziel-Spalte (1 = L … 8 = S)", 9),
    ("zvp", "Verwendungs­prüfung (1 = U, 2 = V)", 10), ("zgrp", "Ziel-Gruppe", 14),
    ("zent", "Eintrag Ziel-Spalte", 9), ("vpent", "Eintrag Verwendungs­prüfung", 11), ("zyear", "Jahr in Ziel-Spalte", 9),
    ("done", "Fertig (1/0)", 8), ("stat", "Status-Nr", 8), ("stxt", "Status", 18), ("sort", "Sortierung", 9),
    ("rank", "Rang im Bezirk", 8), ("key", "Schlüssel Liste", 16),
] + [(f"e{c}", f"Eintrag {c}", 7) for c in "LMNOPSUV"] \
  + [(f"a{c}", f"{c} erreicht ab (0 = „x“, 9999 = nichts)", 10) for c in "LMNOPSUV"] \
  + [("nEff", "Wmech erreicht ab (N oder „x“ in U)", 10), ("oEff", "SigMech erreicht ab (O oder „x“ in V)", 10)] \
  + [(f"lv{k}", None, 11) for k in range(N_TIMES)] \
  + [(f"tl{k}", None, 11) for k in range(N_TIMES)] \
  + [("lvtxt0", "Stufe heute", 18), ("lvtxt8", "Stufe nach Plan", 18), ("icol", "Ist: Spalte (1 = L … 8 = S)", 9),
     ("h2raw", "Hilfe: Jahre vorbei", 14), ("h3raw", "Hilfe: ungültige Einträge", 14),
     ("h1", "Hinweis Ziel", 30), ("h2", "Hinweis Jahre", 30), ("h3", "Hinweis Einträge", 30), ("h4", "Hinweis Ist-Qualifikation", 30),
     ("hint", "Hinweise (alle)", 60), ("pf", "Prüfen (1/0)", 8), ("pnr", "Prüf-Nr", 8)]
DC = {key: get_column_letter(i + 1) for i, (key, _, _) in enumerate(DATA_SPEC)}
DATA_LAST = 2      # wird in build() gesetzt


def drng(key):
    c = DC[key]
    return f"Daten!${c}$2:${c}${DATA_LAST}"


def year9(c):
    """Jahr aus einem Eintrag (27 -> 2027, 2027, Datum); 9999, wenn kein Jahr."""
    v = f"VALUE({c})"
    return (f'IFERROR(IF({c}="",9999,IF({v}>2100,YEAR({v}),IF({v}>=1900,ROUND({v},0),'
            f'IF(AND({v}>=1,{v}<=99),2000+ROUND({v},0),9999)))),9999)')


def build_data(ws, districts):
    ws.sheet_view.showGridLines = True
    for key, h, w in DATA_SPEC:
        col = DC[key]
        if key.startswith("lv") and key[2:].isdigit():
            k = int(key[2:])
            h = ("Stufe heute (0 = keine … 5 = RBEG)" if k == 0 else "Stufe nach Plan" if k == N_TIMES - 1
                 else f'="Stufe Ende "&({E_YEAR}+{k - 1})')
        if key.startswith("tl") and key[2:].isdigit():
            k = int(key[2:])
            h = ("Teamleiter heute" if k == 0 else "Teamleiter nach Plan" if k == N_TIMES - 1
                 else f'="Teamleiter Ende "&({E_YEAR}+{k - 1})')
        ws.column_dimensions[col].width = w
        put(ws, f"{col}1", h, f=font(9, True, WHITE), fl=fill(INK2), al=LEFT_WRAP)
    ws.row_dimensions[1].height = 48
    ws.freeze_panes = "C2"
    last = 1 + len(districts) * SRC_ROWS
    ws.auto_filter.ref = f"A1:{DC['pnr']}{last}"
    plain = font(9)
    r = 2
    for d in districts:
        s = quote_sheetname(d.sheet)
        for k in range(SRC_ROWS):
            c = {key: f"{DC[key]}{r}" for key in DC}
            idx = f"${DC['row']}{r}-{SRC_FIRST_ROW - 1}"

            def src(col):
                # IFERROR: werden im Bezirks-Blatt Zeilen geloescht, wird der Bereich kuerzer -> leer statt Fehler
                return f'IFERROR(INDEX({s}!${col}${SRC_FIRST_ROW}:${col}${SRC_LAST_ROW},{idx}),"")'
            ok = c["ok"]
            f = {
                "bez": d.sheet,
                "row": SRC_FIRST_ROW + k,
                "name": f'=TRIM(CLEAN({src("B")}&""))',
                "vor": f'=TRIM(CLEAN({src("C")}&""))',
                "ziel": f'=TRIM(CLEAN({src("G")}&""))',
                "ist": f'=TRIM(CLEAN({src("J")}&""))',
                "ok": f'=IF(AND({c["name"]}<>"",LOWER({c["name"]})<>"name"),1,0)',
                "zmap": f'=IF(OR({ok}=0,{c["ziel"]}=""),0,IFERROR(MATCH({c["ziel"]},{E_MAP_TEXT},0),0))',
                "zcol": f'=IF({c["zmap"]}=0,0,IFERROR(INDEX({E_MAP_COLNO},{c["zmap"]})*1,0))',
                "zvp": f'=IF({c["zmap"]}=0,0,IFERROR(INDEX({E_MAP_VPNO},{c["zmap"]})*1,0))',
                "zgrp": (f'=IF({c["zmap"]}=0,"{GROUP_OTHER}",IF(ISNUMBER(MATCH(INDEX({E_MAP_GRP},{c["zmap"]})&"",'
                         f'{{{",".join(chr(34) + g + chr(34) for g in CMP)}}},0)),INDEX({E_MAP_GRP},{c["zmap"]})&"","{GROUP_OTHER}"))'),
                "zent": (f'=IF({c["zcol"]}=0,"",TRIM(IFERROR(INDEX({s}!$L${SRC_FIRST_ROW}:$S${SRC_LAST_ROW},{idx},'
                         f'{c["zcol"]}),"")&""))'),
                "vpent": (f'=IF({c["zvp"]}=0,"",TRIM(IFERROR(INDEX({s}!$U${SRC_FIRST_ROW}:$V${SRC_LAST_ROW},{idx},'
                          f'{c["zvp"]}),"")&""))'),
                "zyear": f'=IF({year9(c["zent"])}=9999,"",{year9(c["zent"])})',
                "done": f'=IF(OR(LOWER({c["zent"]})="x",LOWER({c["vpent"]})="x"),1,0)',
                "stat": (f'=IF({ok}=0,"",IF({c["zcol"]}=0,9,IF({c["done"]}=1,1,IF({c["zyear"]}="",8,'
                         f'IF({c["zyear"]}<{E_YEAR},7,MIN(6,2+{c["zyear"]}-{E_YEAR}))))))'),
                "stxt": f'=IF({c["stat"]}="","",INDEX({E_STAT_LABEL},{c["stat"]}))',
                "sort": f'=IF({c["stat"]}="","",{c["stat"]}*1000+{c["row"]})',
                "rank": (f'=IF({c["sort"]}="","",COUNTIFS($A$2:$A${last},{c["bez"]},'
                         f'${DC["sort"]}$2:${DC["sort"]}${last},"<"&{c["sort"]})+1)'),
                "key": f'=IF({c["rank"]}="","",{c["bez"]}&"|"&{c["rank"]})',
            }
            for col in "LMNOPSUV":
                f[f"e{col}"] = f'=TRIM({src(col)}&"")'
                f[f"a{col}"] = f'=IF(LOWER({c["e" + col]})="x",0,{year9(c["e" + col])})'
            f["nEff"] = f'=MIN({c["aN"]},IF(LOWER({c["eU"]})="x",0,9999))'
            f["oEff"] = f'=MIN({c["aO"]},IF(LOWER({c["eV"]})="x",0,9999))'
            for t in range(N_TIMES):
                T = threshold(t)
                f[f"lv{t}"] = (f'=IF({ok}=0,"",IF({c["aP"]}<={T},5,IF({c["oEff"]}<={T},4,IF({c["nEff"]}<={T},3,'
                               f'IF({c["aM"]}<={T},2,IF({c["aL"]}<={T},1,0))))))')
                f[f"tl{t}"] = f'=IF({ok}=0,"",IF({c["aS"]}<={T},1,0))'
            last_t = N_TIMES - 1
            f["lvtxt0"] = f'=IF({ok}=0,"",{stage_name_formula(c["lv0"])}&IF({c["tl0"]}=1," + TL",""))'
            f["lvtxt8"] = (f'=IF({ok}=0,"",{stage_name_formula(c[f"lv{last_t}"])}'
                           f'&IF({c[f"tl{last_t}"]}=1," + TL",""))')
            f["icol"] = (f'=IF(OR({ok}=0,{c["ist"]}=""),0,IF(LEFT(LOWER({c["ist"]}),5)="azubi",1,'
                         f'IFERROR(INDEX({E_MAP_COLNO},MATCH({c["ist"]},{E_MAP_TEXT},0))*1,0)))')
            # Jahre vorbei ohne "x" (Wmech/SigMech: erledigt auch bei "x" in U/V)
            past = []
            for col, key in (("L", "aL"), ("M", "aM"), ("N", "nEff"), ("O", "oEff"), ("P", "aP"), ("S", "aS"),
                             ("U", "aU"), ("V", "aV")):
                past.append(f'IF(AND({c[key]}>=1900,{c[key]}<{E_YEAR}),", {col} "&{c[key]},"")')
            f["h2raw"] = f'=IF({ok}=0,"",' + "&".join(past) + ")"
            bad = [f'IF(AND({c["e" + col]}<>"",{c["a" + col]}=9999),", {col} „"&{c["e" + col]}&"“","")' for col in "LMNOPSUV"]
            f["h3raw"] = f'=IF({ok}=0,"",' + "&".join(bad) + ")"
            right = "+".join(f'(({c["zcol"]}<{i})*({c["e" + col]}<>""))' for i, col in ((2, "M"), (3, "N"), (4, "O"), (5, "P")))
            f["h1"] = (f'=IF({ok}=0,"",IF({c["stat"]}=9,IF({c["ziel"]}="","Ziel-Qualifikation (G) leer",'
                       f'"Ziel „"&{c["ziel"]}&"“ hat keine Spalte in L–S"),IF({c["stat"]}=8,"Ziel „"&{c["ziel"]}'
                       f'&"“: in der Ziel-Spalte weder „x“ noch Jahr"&IF({right}>0," (Planung nur in Spalten weiter rechts)",""),"")))')
            f["h2"] = f'=IF({c["h2raw"]}="","","Jahr vorbei ohne „x“: "&MID({c["h2raw"]},3,200))'
            lv_last = c[f"lv{last_t}"]
            f["h3"] = (f'=IF({ok}=0,"",MID(IF({c["h3raw"]}<>"","; kein gültiger Eintrag: "&MID({c["h3raw"]},3,200),"")'
                       f'&IF({lv_last}=0,"; in L–P nichts eingetragen (keine Stufe)","")'
                       f'&IF(AND({c["eV"]}<>"",{c["eO"]}=""),"; Eintrag in V (örtl. VP SigMech), aber Spalte O leer","")'
                       f'&IF(AND({c["eU"]}<>"",{c["eN"]}=""),"; Eintrag in U (örtl. VP Wmech), aber Spalte N leer","")'
                       f'&IF(AND({c["eP"]}<>"",{c["eO"]}="",LOWER({c["eV"]})<>"x"),"; RBEG (P) eingetragen, aber SigMech (O) leer",""),3,400))')
            f["h4"] = (f'=IF({ok}=0,"",IF(AND({c["icol"]}>=1,{c["icol"]}<=5,{c["icol"]}<>{c["lv0"]}),'
                       f'"Ist-Qualifikation „"&{c["ist"]}&"“, laut L–P heute: "&{stage_name_formula(c["lv0"])},'
                       f'IF(AND({c["icol"]}=8,{c["tl0"]}=0),"Ist-Qualifikation „"&{c["ist"]}&"“, aber kein „x“ in S","")))')
            f["hint"] = (f'=IF({ok}=0,"",MID(IF({c["h1"]}<>"","; "&{c["h1"]},"")&IF({c["h2"]}<>"","; "&{c["h2"]},"")'
                         f'&IF({c["h3"]}<>"","; "&{c["h3"]},"")&IF({c["h4"]}<>"","; "&{c["h4"]},""),3,600))')
            f["pf"] = f'=IF({c["hint"]}="",0,1)'
            f["pnr"] = f'=IF({c["pf"]}=1,COUNTIF(${DC["pf"]}$2:{c["pf"]},1),"")'
            for key, v in f.items():
                cell = ws[c[key]]
                cell.value = v
                cell.font = plain
            r += 1
    return last


# ---------------------------------------------------------------- Soll-Ist-Tabelle (Bezirk und LST gesamt)
# Spalten: B:C Stufe, D Soll, E..M Zeitpunkte (heute, 7 Jahre, nach Plan), N:P Bewertung heute, Q:S Bewertung nach Plan
T_SOLL = "D"
TIME_LETTERS = [get_column_letter(5 + k) for k in range(N_TIMES)]          # E..M
HELP0 = column_index_from_string("W")
COMP_LETTERS = [get_column_letter(HELP0 + k) for k in range(N_TIMES + 1)]           # W..AF: Zeitpunkte + Ziel
SOLL_LETTERS = [get_column_letter(HELP0 + N_TIMES + 2 + k) for k in range(N_TIMES)]   # Soll-Linie
FEHLT_LETTERS = [get_column_letter(HELP0 + 2 * N_TIMES + 3 + k) for k in range(N_TIMES)]
MEHR_LETTERS = [get_column_letter(HELP0 + 3 * N_TIMES + 4 + k) for k in range(N_TIMES)]
SI_ROWS = ["Azubi"] + CMP + [NO_STAGE, "Köpfe", "Fehlend", "Mehr"]
SI_LABEL = {"Azubi": "Azubi", NO_STAGE: "keine Stufe (L–P leer)", "Köpfe": "Mitarbeiter (Köpfe)",
            "Fehlend": "Fehlende Stellen", "Mehr": "Mehr als Soll", TL: "Teamleiter (zusätzlich)"}


def soll_ist_table(ws, hdr, soll_fn, count_fn, helper_fn=None):
    """Zeilen: Azubi, 5 Vergleichsstufen, keine Stufe, Koepfe, fehlende Stellen, mehr als Soll.
    soll_fn(cat) / count_fn(cat, k) liefern Formeln. helper_fn(kind, cat, k) ersetzt die Fehlt/Mehr-Hilfswerte
    (LST gesamt: Summe der Bezirke). Gibt {Zeilenname: Zeile} zurueck."""
    head(ws, f"B{hdr}:C{hdr}", "Stufe", LEFT_WRAP)
    head(ws, f"{T_SOLL}{hdr}", "Soll (Ziel­zustand)")
    for k, col in enumerate(TIME_LETTERS):
        head(ws, f"{col}{hdr}", time_label(k))
    head(ws, f"N{hdr}:P{hdr}", "Bewertung heute", LEFT_WRAP)
    head(ws, f"Q{hdr}:S{hdr}", "Bewertung nach Plan", LEFT_WRAP)
    ws.row_dimensions[hdr].height = 30
    # Hilfszeile: Kategorien fuer Diagramme (lang = Zeile hdr, kurz = Zeile hdr-1)
    for k in range(N_TIMES):
        put(ws, f"{COMP_LETTERS[k]}{hdr}", f"={TIME_LETTERS[k]}{hdr}", f=font(8, color=MUTED))
        put(ws, f"{COMP_LETTERS[k]}{hdr - 1}", time_short(k), f=font(8, color=MUTED))
        put(ws, f"{SOLL_LETTERS[k]}{hdr}", f"={TIME_LETTERS[k]}{hdr}", f=font(8, color=MUTED))
        put(ws, f"{SOLL_LETTERS[k]}{hdr - 1}", time_short(k), f=font(8, color=MUTED))
        put(ws, f"{FEHLT_LETTERS[k]}{hdr}", f'="fehlt "&{TIME_LETTERS[k]}{hdr}', f=font(8, color=MUTED))
        put(ws, f"{MEHR_LETTERS[k]}{hdr}", f'="mehr "&{TIME_LETTERS[k]}{hdr}', f=font(8, color=MUTED))
    put(ws, f"{COMP_LETTERS[-1]}{hdr}", "Ziel", f=font(8, color=MUTED))
    put(ws, f"{COMP_LETTERS[-1]}{hdr - 1}", "Ziel", f=font(8, color=MUTED))

    rows = {}
    r = hdr + 1
    for key in SI_ROWS:
        rows[key] = r
        r += 1
    cmp_rows = [rows[c] for c in CMP]
    for key in SI_ROWS:
        r = rows[key]
        info = key in ("Azubi", NO_STAGE)
        bold = key in ("Köpfe", "Fehlend")
        border = TOP_INK if key == "Köpfe" else BOTTOM_HAIR
        merge_put(ws, f"B{r}:C{r}", SI_LABEL.get(key, key), f=font(9 if info else 10, bold or not info, MUTED if info else INK),
                  al=LEFT, border=border)
        if key in CMP:
            put(ws, f"{T_SOLL}{r}", soll_fn(key), f=font(10, True), al=CENTER, nf="0", border=border)
            for k, col in enumerate(TIME_LETTERS):
                put(ws, f"{col}{r}", count_fn(key, k), f=font(10, k == N_TIMES - 1), al=CENTER, nf="0", border=border)
                fl, ml = FEHLT_LETTERS[k], MEHR_LETTERS[k]
                if helper_fn:
                    put(ws, f"{fl}{r}", helper_fn("fehlt", key, k), f=font(8, color=MUTED), nf="0")
                    put(ws, f"{ml}{r}", helper_fn("mehr", key, k), f=font(8, color=MUTED), nf="0")
                else:
                    put(ws, f"{fl}{r}", f"=MAX(0,${T_SOLL}{r}-{col}{r})", f=font(8, color=MUTED), nf="0")
                    put(ws, f"{ml}{r}", f"=MAX(0,{col}{r}-${T_SOLL}{r})", f=font(8, color=MUTED), nf="0")
            merge_put(ws, f"N{r}:P{r}", gap_text(f"{FEHLT_LETTERS[0]}{r}", f"{MEHR_LETTERS[0]}{r}"), f=font(9, color=INK2),
                      al=LEFT, border=border)
            merge_put(ws, f"Q{r}:S{r}", gap_text(f"{FEHLT_LETTERS[-1]}{r}", f"{MEHR_LETTERS[-1]}{r}"), f=font(9, True, INK2),
                      al=LEFT, border=border)
        elif key in ("Azubi", NO_STAGE):
            put(ws, f"{T_SOLL}{r}", "–", f=font(9, color=MUTED), al=CENTER, border=border)
            for k, col in enumerate(TIME_LETTERS):
                put(ws, f"{col}{r}", count_fn(key, k), f=font(9, color=MUTED), al=CENTER, nf="0", border=border)
            merge_put(ws, f"N{r}:P{r}", "nicht im Soll", f=font(9, color=MUTED), al=LEFT, border=border)
            merge_put(ws, f"Q{r}:S{r}", None, border=border)
        elif key == "Köpfe":
            put(ws, f"{T_SOLL}{r}", "=" + "+".join(f"{T_SOLL}{x}" for x in cmp_rows), f=font(10, True), al=CENTER, nf="0",
                border=border)
            for k, col in enumerate(TIME_LETTERS):
                put(ws, f"{col}{r}", count_fn(key, k), f=font(10, True), al=CENTER, nf="0", border=border)
            merge_put(ws, f"N{r}:P{r}", f'="Ist "&{TIME_LETTERS[0]}{r}&" · Soll "&{T_SOLL}{r}&" Köpfe"',
                      f=font(9, color=INK2), al=LEFT, border=border)
            merge_put(ws, f"Q{r}:S{r}", "Teamleiter sind hier nur einmal gezählt", f=font(9, color=MUTED), al=LEFT, border=border)
        else:   # Fehlend / Mehr: Summe der Hilfswerte ueber die Vergleichsstufen
            put(ws, f"{T_SOLL}{r}", None, border=border)
            src_letters = FEHLT_LETTERS if key == "Fehlend" else MEHR_LETTERS
            for k, col in enumerate(TIME_LETTERS):
                h = src_letters[k]
                put(ws, f"{col}{r}", f"=SUM({h}{cmp_rows[0]}:{h}{cmp_rows[-1]})", f=font(10, key == "Fehlend"),
                    al=CENTER, nf="0", border=border)
            merge_put(ws, f"N{r}:P{r}", None, border=border)
            merge_put(ws, f"Q{r}:S{r}", None, border=border)
    # Hilfswerte fuer Diagramme: gestapelte Saeulen (Zeitpunkte + Ziel), Soll-Linie
    for key in ["Azubi"] + CMP + [NO_STAGE]:
        r = rows[key]
        for k in range(N_TIMES):
            put(ws, f"{COMP_LETTERS[k]}{r}", f"={TIME_LETTERS[k]}{r}", f=font(8, color=MUTED), nf="0")
            if key in CMP:
                put(ws, f"{SOLL_LETTERS[k]}{r}", f"=${T_SOLL}{r}", f=font(8, color=MUTED), nf="0")
        put(ws, f"{COMP_LETTERS[-1]}{r}", f"={T_SOLL}{r}" if key in CMP else 0, f=font(8, color=MUTED), nf="0")
    first, last = cmp_rows[0], cmp_rows[-1]
    add_compare_cf(ws, f"{TIME_LETTERS[0]}{first}:{TIME_LETTERS[-1]}{last}", f"{TIME_LETTERS[0]}{first}", f"${T_SOLL}{first}")
    add_gap_cf(ws, f"N{first}:S{last}", f"N{first}")
    for k, col in enumerate(TIME_LETTERS):
        ws.conditional_formatting.add(f"{col}{rows['Fehlend']}", FormulaRule(
            formula=[f"{col}{rows['Fehlend']}>0"], font=Font(color=BAD_INK, bold=True)))
    return rows


def gap_table(ws, hdr, rows, title_rows_label="Stufe", row_height=26):
    """Abweichung je Stufe und Zeitpunkt als Text (aus den Fehlt/Mehr-Hilfswerten der Soll-Ist-Tabelle)."""
    head(ws, f"B{hdr}:C{hdr}", title_rows_label, LEFT_WRAP)
    head(ws, f"{T_SOLL}{hdr}", "Soll (Ziel­zustand)")
    for k, col in enumerate(TIME_LETTERS):
        head(ws, f"{col}{hdr}", time_label(k))
    ws.row_dimensions[hdr].height = 30
    out = {}
    for i, key in enumerate(CMP):
        r = hdr + 1 + i
        out[key] = r
        src = rows[key]
        merge_put(ws, f"B{r}:C{r}", SI_LABEL.get(key, key), f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"{T_SOLL}{r}", f"={T_SOLL}{src}", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
        for k, col in enumerate(TIME_LETTERS):
            put(ws, f"{col}{r}", gap_text(f"{FEHLT_LETTERS[k]}{src}", f"{MEHR_LETTERS[k]}{src}"), f=font(8, color=INK2),
                al=CENTER_WRAP, border=BOTTOM_HAIR)
        ws.row_dimensions[r].height = row_height
    first, last = hdr + 1, hdr + len(CMP)
    add_gap_cf(ws, f"{TIME_LETTERS[0]}{first}:{TIME_LETTERS[-1]}{last}", f"{TIME_LETTERS[0]}{first}")
    return out


def stack_series(title, rows):
    """Serien fuer das gestapelte Stufen-Diagramm aus den Hilfsspalten (unten: keine Stufe, oben: Teamleiter)."""
    out = []
    for key in STACK:
        r = rows[key]
        color, txt = STAGE_COLORS[key]
        out.append((rng(title, COMP_LETTERS[0], r, COMP_LETTERS[-1], r), rng(title, "B", r), color, txt))
    return out


# ---------------------------------------------------------------- Blatt: Bezirk
# Seite 1: Kennzahlen + Status + Soll/heute/Plan | Seite 2: Verlauf | Seite 3: Soll-Ist-Zahlen
# Seite 4: Status-Zahlen | Seite 5: Mitarbeiterliste
B_SEC1, B_CH1, B_CH1_END = 10, 11, 31
B_SEC2, B_CH2, B_CH2_END = 32, 34, 51
B_SEC3, B_CH3, B_CH3_END = 52, 53, 65
B_T1 = 66                                   # Soll-Ist Zahlen
B_SI_HDR = B_T1 + 1
B_SI_NOTE = B_SI_HDR + len(SI_ROWS) + 1
B_T2 = B_SI_NOTE + 2                        # Abweichung als Text
B_GAP_HDR = B_T2 + 1
B_T3 = B_GAP_HDR + len(CMP) + 2             # Status nach Ziel
B_ST_HDR = B_T3 + 1
B_ST_FIRST = B_ST_HDR + 1
B_ST_LAST = B_ST_HDR + len(GROUPS)
B_ST_SUM = B_ST_LAST + 1
B_ST_PCT = B_ST_SUM + 1
B_LEG = B_ST_PCT + 2                        # Legende Status
B_LIST = B_LEG + len(STATUS) + 2            # Mitarbeiterliste
B_LIST_HDR = B_LIST + 1
B_LIST_FIRST = B_LIST_HDR + 1
B_LIST_LAST = B_LIST_FIRST + LIST_ROWS - 1
ST_COL1 = 4                                 # Spalte D = Status 1 ... L = Status 9
ST_SUM_COL = "M"


def district_title(d):
    return "Bezirk " + d.sheet


def status_legend(ws, row):
    """Farbe, Bezeichnung und Bedeutung je Status."""
    section_title(ws, f"B{row}:{GRID_LAST}{row}", "Bedeutung der Status")
    for code, color, _ in STATUS:
        r = row + code
        put(ws, f"B{r}", None, fl=fill(color))
        merge_put(ws, f"C{r}:E{r}", f"={stat_label_ref(code)}", f=font(10, True),
                  al=Alignment(horizontal="left", vertical="center", indent=1))
        merge_put(ws, f"F{r}:{GRID_LAST}{r}", f"=Einstellungen!$E${E_STAT_FIRST + code - 1}", f=font(9, color=INK2), al=LEFT_WRAP)
        ws.row_dimensions[r].height = 16


def build_district(ws, d):
    setup_grid(ws, hidden_from=KEY_COL)
    title = ws.title
    K = f"${KEY_COL}$1"                     # Blattname des Bezirks (Schluessel in Daten)
    put(ws, f"{KEY_COL}1", d.sheet)
    put(ws, f"{KEY_COL}2", f'=IFERROR(INDEX({E_BEZ_ZZCOL},MATCH({K},{E_BEZ_SHEET},0))*1,0)')
    header_block(ws, f"Ausbildungsstand – {d.sheet}",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   Quelle: Blatt „{d.sheet}“ (Ziel-Qualifikation G, Stand L–S, örtl. Verwendungs'
                 f'prüfung U/V)   ·   Soll: Zielzustand „"&IFERROR(INDEX({E_BEZ_ZZ},MATCH({K},{E_BEZ_SHEET},0)),"–")&"“"',
                 links=[("Übersicht", "← Übersicht"), ("Soll-Ist", "Soll-Ist alle Bezirke →")])

    # ---- Zahlen: Status nach Ziel-Qualifikation (wird von Kennzahlen und Diagramm genutzt)
    section_title(ws, f"B{B_T3}:{GRID_LAST}{B_T3}", "Zahlen: Status je Mitarbeiter nach Ziel-Qualifikation")
    head(ws, f"B{B_ST_HDR}:C{B_ST_HDR}", "Ziel-Qualifikation", LEFT_WRAP)
    status_header_cells(ws, B_ST_HDR, ST_COL1)
    head(ws, f"{ST_SUM_COL}{B_ST_HDR}", "Summe")
    ws.row_dimensions[B_ST_HDR].height = 30
    for i, g in enumerate(GROUPS):
        r = B_ST_FIRST + i
        merge_put(ws, f"B{r}:C{r}", g, al=LEFT, border=BOTTOM_HAIR)
        for code in range(1, len(STATUS) + 1):
            col = get_column_letter(ST_COL1 + code - 1)
            put(ws, f"{col}{r}", f'=COUNTIFS({drng("bez")},{K},{drng("zgrp")},$B{r},{drng("stat")},{code})',
                al=CENTER, nf=NF_COUNT, border=BOTTOM_HAIR)
        put(ws, f"{ST_SUM_COL}{r}", f"=SUM(D{r}:L{r})", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
    merge_put(ws, f"B{B_ST_SUM}:C{B_ST_SUM}", "Summe", f=font(10, True), al=LEFT)
    merge_put(ws, f"B{B_ST_PCT}:C{B_ST_PCT}", "Anteil", f=font(9, color=INK2), al=LEFT)
    total = f"${ST_SUM_COL}${B_ST_SUM}"
    for c in range(ST_COL1, column_index_from_string(ST_SUM_COL) + 1):
        col = get_column_letter(c)
        put(ws, f"{col}{B_ST_SUM}", f"=SUM({col}{B_ST_FIRST}:{col}{B_ST_LAST})", f=font(10, True), al=CENTER, nf="0", border=TOP_INK)
        put(ws, f"{col}{B_ST_PCT}", f'=IF({total}=0,"",{col}{B_ST_SUM}/{total})', f=font(9, color=INK2), al=CENTER, nf="0%")
    status_legend(ws, B_LEG)

    # ---- Zahlen: Soll-Ist je Stufe
    section_title(ws, f"B{B_T1}:{GRID_LAST}{B_T1}", "Zahlen: Soll-Ist je Stufe (Anzahl Mitarbeiter heute und jeweils am Jahresende)")
    soll = lambda cat: f'=IF(${KEY_COL}$2=0,0,IFERROR(INDEX({ZZ_AREA},MATCH("{cat}",{ZZ_LABELS},0),${KEY_COL}$2)*1,0))'

    def count(cat, k):
        if cat == TL:
            return f'=COUNTIFS({drng("bez")},{K},{drng(f"tl{k}")},1)'
        if cat == "Köpfe":
            return f'=COUNTIFS({drng("bez")},{K},{drng("ok")},1)'
        nr = 0 if cat == NO_STAGE else STAGES.index(cat) + 1
        return f'=COUNTIFS({drng("bez")},{K},{drng(f"lv{k}")},{nr})'
    rows = soll_ist_table(ws, B_SI_HDR, soll, count)
    merge_put(ws, f"B{B_SI_NOTE}:{GRID_LAST}{B_SI_NOTE}",
              "Spalten: Azubi = L · Arbeiter LST = M · Wmech = N (oder „x“ in U) · SigMech = O (oder „x“ in V) · SigMech RBEG = P · "
              "Teamleiter = S (zählt zusätzlich, also auch bei seiner Stufe). Grün = Soll erreicht, rot = es fehlen, gelb = mehr als Soll.",
              f=font(8, italic=True, color=MUTED), al=LEFT_WRAP)
    ws.row_dimensions[B_SI_NOTE].height = 24
    section_title(ws, f"B{B_T2}:{GRID_LAST}{B_T2}", "Zahlen: Abweichung vom Soll je Stufe")
    gap_table(ws, B_GAP_HDR, rows)

    # ---- Kennzahlen
    sc = lambda code: f"{get_column_letter(ST_COL1 + code - 1)}{B_ST_SUM}"
    davon = (f'COUNTIFS({drng("bez")},{K},{drng("stat")},">=2",{drng("stat")},"<=6",{drng("zyear")},{E_DAVON})')
    rf, rm = rows["Fehlend"], rows["Mehr"]
    kpi_tiles(ws, [
        ("Mitarbeiter", f"={total}", f'=IF({sc(9)}>0,"davon "&{sc(9)}&" mit unklarem Ziel","im Bezirk")', INK2),
        ("Fertig", f"={sc(1)}", pct_sub(sc(1), total), STATUS[0][1]),
        ("In Ausbildung", f"=SUM({sc(2)}:{sc(6)})", f'="davon Abschluss "&{E_DAVON}&": "&{davon}', "2A78D6"),
        ("Überfällig", f"={sc(7)}", '="Jahr vorbei, noch kein „x“"', STATUS[6][1]),
        ("Fehlt (nichts geplant)", f"={sc(8)}", pct_sub(sc(8), total), STATUS[7][1]),
        ("Fehlende Stellen zum Soll", f"={TIME_LETTERS[-1]}{rf}",
         f'="nach Plan · heute: "&{TIME_LETTERS[0]}{rf}&" · mehr als Soll: "&{TIME_LETTERS[-1]}{rm}', BAD_INK),
    ])
    how_box(ws, 8, HOW_SHORT, height=34)

    # ---- Diagramme Seite 1
    section_title(ws, f"B{B_SEC1}:I{B_SEC1}", "Status je Mitarbeiter (gemessen an der Ziel-Qualifikation)")
    section_title(ws, f"J{B_SEC1}:{GRID_LAST}{B_SEC1}", "Soll, heute und nach Plan je Stufe")
    place(ws, status_bar_chart(title, B_ST_SUM, ST_COL1, ST_COL1 + len(STATUS) - 1, B_ST_HDR), f"B{B_CH1}", f"J{B_CH1_END}")
    cf, cl = rows[CMP[0]], rows[CMP[-1]]
    cats = rng(title, "B", cf, "B", cl)
    place(ws, clustered_chart(cats, [
        (rng(title, T_SOLL, cf, T_SOLL, cl), rng(title, T_SOLL, B_SI_HDR), C_SOLL),
        (rng(title, TIME_LETTERS[0], cf, TIME_LETTERS[0], cl), rng(title, TIME_LETTERS[0], B_SI_HDR), C_HEUTE),
        (rng(title, TIME_LETTERS[-1], cf, TIME_LETTERS[-1], cl), rng(title, TIME_LETTERS[-1], B_SI_HDR), C_PLAN),
    ]), f"J{B_CH1}", f"T{B_CH1_END}")

    # ---- Seite 2: Verlauf
    section_title(ws, f"B{B_SEC2}:{GRID_LAST}{B_SEC2}", "Verlauf: Mitarbeiter je Stufe heute und am Jahresende – rechts der Zielzustand")
    section_note(ws, f"B{B_SEC2 + 1}:{GRID_LAST}{B_SEC2 + 1}",
                 "Jede Säule = Mitarbeiter je Stufe (L–P, jeder einmal). Teamleiter (orange) zählen zusätzlich obendrauf. "
                 "Die Säule „Ziel“ zeigt den Zielzustand (Soll).", height=16)
    comp_cats = rng(title, COMP_LETTERS[0], B_SI_HDR, COMP_LETTERS[-1], B_SI_HDR)
    place(ws, stacked_stage_chart(comp_cats, stack_series(title, rows)), f"B{B_CH2}", f"T{B_CH2_END}")

    section_title(ws, f"B{B_SEC3}:{GRID_LAST}{B_SEC3}", "Verlauf je Stufe – Säulen = Mitarbeiter, gestrichelte Linie = Soll")
    short = rng(title, SOLL_LETTERS[0], B_SI_HDR - 1, SOLL_LETTERS[-1], B_SI_HDR - 1)
    anchors = ["B", "E", "H", "K", "N", "Q", "T"]
    for i, cat in enumerate(CMP):
        r = rows[cat]
        ch = verlauf_chart(short, rng(title, TIME_LETTERS[0], r, TIME_LETTERS[-1], r), rng(title, "B", r),
                           rng(title, SOLL_LETTERS[0], r, SOLL_LETTERS[-1], r), rng(title, T_SOLL, B_SI_HDR),
                           title=cat, color=STAGE_COLORS[cat][0])
        place(ws, ch, f"{anchors[i]}{B_CH3}", f"{anchors[i + 1]}{B_CH3_END}")
    rf_ = rows["Fehlend"]
    place(ws, verlauf_chart(short, rng(title, TIME_LETTERS[0], rf_, TIME_LETTERS[-1], rf_), rng(title, "B", rf_),
                            title="Fehlende Stellen zum Soll", color=C_GAP),
          f"{anchors[5]}{B_CH3}", f"{anchors[6]}{B_CH3_END}")

    # ---- Mitarbeiterliste
    section_title(ws, f"B{B_LIST}:{GRID_LAST}{B_LIST}", "Mitarbeiterliste (sortiert nach Status)")
    cols = [("B", "B", "Nr."), ("C", "D", "Name"), ("E", "F", "Vorname"), ("G", "H", "Ziel-Qualifikation (G)"),
            ("I", "J", "Ist-Qualifikation (J)"), ("K", "L", "Stufe heute (L–P)"), ("M", "N", "Stufe nach Plan"),
            ("O", "O", "Jahr Ziel"), ("P", GRID_LAST, "Status")]
    for a, b, text in cols:
        head(ws, f"{a}{B_LIST_HDR}" if a == b else f"{a}{B_LIST_HDR}:{b}{B_LIST_HDR}", text, LEFT_WRAP)
    ws.row_dimensions[B_LIST_HDR].height = 28
    m, st = KEY_COL, "V"
    for k in range(1, LIST_ROWS + 1):
        r = B_LIST_FIRST + k - 1
        put(ws, f"{m}{r}", f'=IFERROR(MATCH({K}&"|"&{k},{drng("key")},0),"")')
        put(ws, f"{st}{r}", f'=IF(${m}{r}="","",INDEX({drng("stat")},${m}{r}))')
        get = lambda key: f'=IF(${m}{r}="","",INDEX({drng(key)},${m}{r}))'
        put(ws, f"B{r}", f'=IF(${m}{r}="","",{k})', f=font(9, color=MUTED), al=LEFT)
        merge_put(ws, f"C{r}:D{r}", get("name"), f=font(10, True), al=LEFT)
        merge_put(ws, f"E{r}:F{r}", get("vor"), al=LEFT)
        merge_put(ws, f"G{r}:H{r}", get("ziel"), f=font(9), al=LEFT)
        merge_put(ws, f"I{r}:J{r}", get("ist"), f=font(9), al=LEFT)
        merge_put(ws, f"K{r}:L{r}", get("lvtxt0"), f=font(9), al=LEFT)
        merge_put(ws, f"M{r}:N{r}", get("lvtxt8"), f=font(9), al=LEFT)
        put(ws, f"O{r}", f'=IF(${m}{r}="","",IF(INDEX({drng("zyear")},${m}{r})="","–",INDEX({drng("zyear")},${m}{r})))',
            al=CENTER, nf="0")
        merge_put(ws, f"P{r}:{GRID_LAST}{r}", get("stxt"), f=font(10, True), al=Alignment(horizontal="left", vertical="center", indent=1))
        ws.row_dimensions[r].height = 15
    over = B_LIST_LAST + 1
    merge_put(ws, f"B{over}:{GRID_LAST}{over}",
              f'=IF({total}>{LIST_ROWS},"Hinweis: Der Bezirk hat mehr als {LIST_ROWS} Mitarbeiter – die Liste zeigt nur die ersten {LIST_ROWS}.","")',
              f=font(9, True, STATUS[7][1]), al=LEFT)
    ws.conditional_formatting.add(f"B{B_LIST_FIRST}:O{B_LIST_LAST}",
                                  FormulaRule(formula=[f'${m}{B_LIST_FIRST}<>""'], border=BOTTOM_HAIR))
    chip_border = Border(bottom=Side(style="thin", color=WHITE))
    for code, color, txt in STATUS:
        ws.conditional_formatting.add(f"P{B_LIST_FIRST}:{GRID_LAST}{B_LIST_LAST}", FormulaRule(
            formula=[f"${st}{B_LIST_FIRST}={code}"], fill=fill(color), font=Font(color=txt, bold=True), border=chip_border))

    page_setup(ws, f"A1:{GRID_LAST}{over}", breaks=[B_SEC2 - 1, B_T1 - 1, B_T3 - 1, B_LIST - 1])
    heights(ws, {9: 8, B_CH2_END: 8, B_T2 - 1: 8, B_T3 - 1: 8, B_ST_PCT + 1: 8, B_LIST - 1: 8})
    fix_row_heights(ws)
    ws.sheet_properties.tabColor = "86B6EF"
    return rows


# ---------------------------------------------------------------- Blatt: Uebersicht (Status)
OV_SEC1, OV_CH1, OV_CH1_END = 10, 11, 31
OV_T = 32
OV_HDR = OV_T + 1
OV_FIRST = OV_HDR + 1


def build_overview(ws, districts, drows):
    setup_grid(ws, hidden_from=KEY_COL)
    title = ws.title
    n = len(districts)
    ov_last = OV_FIRST + n - 1
    ov_sum, ov_pct = ov_last + 1, ov_last + 2
    header_block(ws, "Ausbildungsstand LST – alle Bezirke",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   {n} Bezirke   ·   Klick auf einen Bezirk in der Tabelle öffnet seine Seite   ·   '
                 f'Alles rechnet mit Formeln und passt sich an, wenn die Bezirks-Blätter oder der Zielzustand geändert werden"',
                 links=[("Soll-Ist", "Soll-Ist-Vergleich →"), ("Prüfliste", "Prüfliste →")])
    sc = lambda code: f"{get_column_letter(3 + code)}{ov_sum}"      # Status 1 -> Spalte D
    total = f"$M${ov_sum}"
    gsheets = [district_title(d) for d in districts]
    gsum = lambda ref: "+".join(f"{quote_sheetname(g)}!{ref}" for g in gsheets)
    rf = drows["Fehlend"]
    davon = f'COUNTIFS({drng("ok")},1,{drng("stat")},">=2",{drng("stat")},"<=6",{drng("zyear")},{E_DAVON})'
    kpi_tiles(ws, [
        ("Mitarbeiter gesamt", f"={total}", f'=IF({sc(9)}>0,"davon "&{sc(9)}&" mit unklarem Ziel","in allen Bezirken")', INK2),
        ("Fertig", f"={sc(1)}", pct_sub(sc(1), total), STATUS[0][1]),
        ("In Ausbildung", f"=SUM({sc(2)}:{sc(6)})", f'="davon Abschluss "&{E_DAVON}&": "&{davon}', "2A78D6"),
        ("Überfällig", f"={sc(7)}", '="Jahr vorbei, noch kein „x“"', STATUS[6][1]),
        ("Fehlt (nichts geplant)", f"={sc(8)}", pct_sub(sc(8), total), STATUS[7][1]),
        ("Fehlende Stellen zum Soll", f"={gsum(f'${TIME_LETTERS[-1]}${rf}')}",
         f'="nach Plan · heute: "&({gsum(f"${TIME_LETTERS[0]}${rf}")})', BAD_INK),
    ])
    how_box(ws, 8, HOW_STATUS, height=38)

    section_title(ws, f"B{OV_SEC1}:H{OV_SEC1}", "Alle Bezirke nach Status")
    section_title(ws, f"I{OV_SEC1}:{GRID_LAST}{OV_SEC1}", "Status je Bezirk")
    section_title(ws, f"B{OV_T}:{GRID_LAST}{OV_T}", "Zahlen je Bezirk (Klick auf den Bezirk öffnet seine Seite)")
    head(ws, f"B{OV_HDR}:C{OV_HDR}", "Bezirk", LEFT_WRAP)
    status_header_cells(ws, OV_HDR, 4)
    head(ws, f"M{OV_HDR}", "Summe")
    head(ws, f"N{OV_HDR}", "Anteil fertig")
    head(ws, f"O{OV_HDR}:P{OV_HDR}", f'="In Ausbildung, davon Abschluss "&{E_DAVON}')
    head(ws, f"Q{OV_HDR}", "Fehlende Stellen heute")
    head(ws, f"R{OV_HDR}", "Fehlende Stellen nach Plan")
    head(ws, f"S{OV_HDR}", "Mehr als Soll nach Plan")
    ws.row_dimensions[OV_HDR].height = 42
    for i, d in enumerate(districts):
        r = OV_FIRST + i
        g = quote_sheetname(gsheets[i])
        put(ws, f"{KEY_COL}{r}", d.sheet)
        link = "#" + g.replace('"', '""') + "!A1"
        merge_put(ws, f"B{r}:C{r}", f'=HYPERLINK("{link}","{d.sheet}")', f=font(10, True, LINK, underline="single"),
                  al=LEFT, border=BOTTOM_HAIR)
        for code in range(1, len(STATUS) + 1):
            col = get_column_letter(3 + code)
            put(ws, f"{col}{r}", f'=COUNTIFS({drng("bez")},${KEY_COL}{r},{drng("stat")},{code})', al=CENTER, nf=NF_COUNT,
                border=BOTTOM_HAIR)
        put(ws, f"M{r}", f"=SUM(D{r}:L{r})", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
        put(ws, f"N{r}", f'=IF(M{r}=0,"",D{r}/M{r})', al=CENTER, nf="0%", border=BOTTOM_HAIR)
        merge_put(ws, f"O{r}:P{r}", f'=COUNTIFS({drng("bez")},${KEY_COL}{r},{drng("stat")},">=2",{drng("stat")},"<=6",'
                                    f'{drng("zyear")},{E_DAVON})', al=CENTER, nf="0", border=BOTTOM_HAIR)
        put(ws, f"Q{r}", f"={g}!${TIME_LETTERS[0]}${rf}", al=CENTER, nf="0", border=BOTTOM_HAIR)
        put(ws, f"R{r}", f"={g}!${TIME_LETTERS[-1]}${rf}", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
        put(ws, f"S{r}", f"={g}!${TIME_LETTERS[-1]}${drows['Mehr']}", al=CENTER, nf="0", border=BOTTOM_HAIR)
    merge_put(ws, f"B{ov_sum}:C{ov_sum}", "Summe", f=font(10, True), al=LEFT)
    merge_put(ws, f"B{ov_pct}:C{ov_pct}", "Anteil", f=font(9, color=INK2), al=LEFT)
    for c in range(4, 20):
        col = get_column_letter(c)
        if col == "P":
            continue
        if col == "N":
            put(ws, f"N{ov_sum}", f'=IF({total}=0,"",D{ov_sum}/{total})', f=font(10, True), al=CENTER, nf="0%", border=TOP_INK)
            continue
        if col == "O":
            merge_put(ws, f"O{ov_sum}:P{ov_sum}", f"=SUM(O{OV_FIRST}:O{ov_last})", f=font(10, True), al=CENTER, nf="0",
                      border=TOP_INK)
            continue
        put(ws, f"{col}{ov_sum}", f"=SUM({col}{OV_FIRST}:{col}{ov_last})", f=font(10, True), al=CENTER, nf="0", border=TOP_INK)
        if c <= 13:
            put(ws, f"{col}{ov_pct}", f'=IF({total}=0,"",{col}{ov_sum}/{total})', f=font(9, color=INK2), al=CENTER, nf="0%")
    for col in ("Q", "R"):
        ws.conditional_formatting.add(f"{col}{OV_FIRST}:{col}{ov_last}", FormulaRule(
            formula=[f"{col}{OV_FIRST}>0"], font=Font(color=BAD_INK, bold=True)))

    place(ws, status_bar_chart(title, ov_sum, 4, 4 + len(STATUS) - 1, OV_HDR), f"B{OV_CH1}", f"I{OV_CH1_END}")
    place(ws, stacked_status_chart(title, OV_HDR, OV_FIRST, ov_last, 4, 4 + len(STATUS) - 1), f"I{OV_CH1}", f"T{OV_CH1_END}")

    lg = ov_pct + 2
    status_legend(ws, lg)
    foot = lg + len(STATUS) + 2
    merge_put(ws, f"B{foot}:{GRID_LAST}{foot}",
              "Bezugsjahr und Zuordnungen: Blatt „Einstellungen“  ·  Soll: Blatt „Zielzustand“  ·  Auffällige Einträge: Blatt "
              "„Prüfliste“  ·  Auswertung je Mitarbeiter: Blatt „Daten“.", f=font(9, italic=True, color=MUTED), al=LEFT)
    page_setup(ws, f"A1:{GRID_LAST}{foot}", breaks=[OV_T - 1])
    heights(ws, {9: 8, ov_pct + 1: 8})
    fix_row_heights(ws)
    ws.sheet_properties.tabColor = "104281"


# ---------------------------------------------------------------- Blatt: Soll-Ist (alle Bezirke)
SM_PER_ROW = 3
SM_ROWS = 15                # Zeilen je kleinem Diagramm
SM_PER_PAGE = 6


def build_soll_ist(ws, districts, drows, ymax):
    setup_grid(ws, hidden_from=KEY_COL)
    title = ws.title
    n = len(districts)
    gsheets = [district_title(d) for d in districts]
    header_block(ws, "Soll-Ist-Vergleich LST – Zielzustand je Bezirk",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   Soll = Spalte „Zielzustand“ (Blatt Zielzustand)   ·   Ist = Stufe laut L–P, '
                 f'jeder Mitarbeiter einmal, Teamleiter zusätzlich"',
                 links=[("Übersicht", "← Übersicht"), ("Prüfliste", "Prüfliste →")])

    # Zeilenplan
    R_SEC1, R_CH1, R_CH1_END = 10, 11, 31
    R_SM = 32                                            # Seite 2/3: Verlauf je Bezirk
    pages = (n + SM_PER_PAGE - 1) // SM_PER_PAGE
    band = SM_ROWS + 1
    sm_page_rows = 2 + 2 * band                          # Titel + Legende + 2 Reihen
    R_TG = R_SM + pages * sm_page_rows                  # Seite 4: LST gesamt + fehlende Stellen je Bezirk
    TG_HDR = R_TG + 2
    TG_NOTE = TG_HDR + len(SI_ROWS) + 1
    R_TB = TG_NOTE + 2
    TB_HDR = R_TB + 1
    TB_FIRST = TB_HDR + 1
    TB_LAST = TB_FIRST + n - 1
    TB_SUM = TB_LAST + 1
    R_TQ = TB_SUM + 2                                    # Seite 5: je Bezirk und Stufe
    TQ_HDR = R_TQ + 2
    TQ_FIRST = TQ_HDR + 1
    TQ_LAST = TQ_FIRST + n - 1
    R_TQ2 = TQ_LAST + 2                                  # Abweichung LST gesamt als Text
    TQ2_HDR = R_TQ2 + 1

    # ---- Tabelle LST gesamt (Summe der Bezirke)
    section_title(ws, f"B{R_TG}:{GRID_LAST}{R_TG}", "Zahlen: LST gesamt – Soll-Ist je Stufe (Summe der Bezirke)")
    section_note(ws, f"B{R_TG + 1}:{GRID_LAST}{R_TG + 1}",
                 "Bewertung als Summe über die Bezirke: ein Überschuss in einem Bezirk gleicht keine Lücke in einem anderen aus.")

    def g_sum(row, col):
        return "=" + "+".join(f"{quote_sheetname(g)}!${col}${row}" for g in gsheets)
    soll = lambda cat: g_sum(drows[cat], T_SOLL)
    count = lambda cat, k: g_sum(drows[cat], TIME_LETTERS[k])

    def helper(kind, cat, k):
        letters = FEHLT_LETTERS if kind == "fehlt" else MEHR_LETTERS
        return g_sum(drows[cat], letters[k])
    g_rows = soll_ist_table(ws, TG_HDR, soll, count, helper_fn=helper)
    merge_put(ws, f"B{TG_NOTE}:{GRID_LAST}{TG_NOTE}", HOW_COUNT, f=font(8, italic=True, color=MUTED), al=LEFT_WRAP)
    ws.row_dimensions[TG_NOTE].height = 34

    # ---- Tabelle: fehlende Stellen je Bezirk und Zeitpunkt
    section_title(ws, f"B{R_TB}:{GRID_LAST}{R_TB}", "Zahlen: Fehlende Stellen zum Soll je Bezirk (Summe über alle Stufen)")
    head(ws, f"B{TB_HDR}:C{TB_HDR}", "Bezirk (Klick öffnet die Seite)", LEFT_WRAP)
    head(ws, f"{T_SOLL}{TB_HDR}", "Soll (Köpfe)")
    for k, col in enumerate(TIME_LETTERS):
        head(ws, f"{col}{TB_HDR}", time_label(k))
    head(ws, f"N{TB_HDR}", "Mitarbeiter heute (Köpfe)")
    head(ws, f"O{TB_HDR}", "Mehr als Soll nach Plan")
    head(ws, f"P{TB_HDR}:S{TB_HDR}", "Bewertung nach Plan", LEFT_WRAP)
    ws.row_dimensions[TB_HDR].height = 42
    rf, rm, rk = drows["Fehlend"], drows["Mehr"], drows["Köpfe"]
    for i, d in enumerate(districts):
        r = TB_FIRST + i
        g = quote_sheetname(gsheets[i])
        link = "#" + g.replace('"', '""') + "!A1"
        merge_put(ws, f"B{r}:C{r}", f'=HYPERLINK("{link}","{d.sheet}")', f=font(10, True, LINK, underline="single"),
                  al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"{T_SOLL}{r}", f"={g}!${T_SOLL}${rk}", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
        for k, col in enumerate(TIME_LETTERS):
            put(ws, f"{col}{r}", f"={g}!${col}${rf}", f=font(10, k == N_TIMES - 1), al=CENTER, nf="0", border=BOTTOM_HAIR)
        put(ws, f"N{r}", f"={g}!${TIME_LETTERS[0]}${rk}", al=CENTER, nf="0", border=BOTTOM_HAIR)
        put(ws, f"O{r}", f"={g}!${TIME_LETTERS[-1]}${rm}", al=CENTER, nf="0", border=BOTTOM_HAIR)
        merge_put(ws, f"P{r}:S{r}", gap_text(f"{TIME_LETTERS[-1]}{r}", f"O{r}"), f=font(9, color=INK2), al=LEFT, border=BOTTOM_HAIR)
    merge_put(ws, f"B{TB_SUM}:C{TB_SUM}", "LST gesamt", f=font(10, True), al=LEFT, border=TOP_INK)
    for col in [T_SOLL] + TIME_LETTERS + ["N", "O"]:
        put(ws, f"{col}{TB_SUM}", f"=SUM({col}{TB_FIRST}:{col}{TB_LAST})", f=font(10, True), al=CENTER, nf="0", border=TOP_INK)
    merge_put(ws, f"P{TB_SUM}:S{TB_SUM}", None, border=TOP_INK)
    ws.conditional_formatting.add(f"{TIME_LETTERS[0]}{TB_FIRST}:{TIME_LETTERS[-1]}{TB_LAST}", FormulaRule(
        formula=[f"{TIME_LETTERS[0]}{TB_FIRST}>0"], fill=fill(BAD_BG), font=Font(color=BAD_INK)))
    ws.conditional_formatting.add(f"{TIME_LETTERS[0]}{TB_FIRST}:{TIME_LETTERS[-1]}{TB_LAST}", FormulaRule(
        formula=[f"{TIME_LETTERS[0]}{TB_FIRST}=0"], fill=fill(GOOD_BG), font=Font(color=GOOD_INK)))
    add_gap_cf(ws, f"P{TB_FIRST}:P{TB_LAST}", f"P{TB_FIRST}")

    # ---- Tabelle je Bezirk und Stufe: Soll / heute / nach Plan
    section_title(ws, f"B{R_TQ}:{GRID_LAST}{R_TQ}", "Zahlen: Soll, heute und nach Plan je Bezirk und Stufe")
    section_note(ws, f"B{R_TQ + 1}:{GRID_LAST}{R_TQ + 1}",
                 "Grün = Soll erreicht · rot = es fehlen · gelb = mehr als Soll (Teamleiter zählen zusätzlich auch bei ihrer Stufe)")
    head(ws, f"B{TQ_HDR}:C{TQ_HDR}", "Bezirk", LEFT_WRAP)
    for j, cat in enumerate(CMP):
        for m, lab in enumerate(["Soll", "heute", "nach Plan"]):
            head(ws, f"{get_column_letter(4 + 3 * j + m)}{TQ_HDR}", f"{cat}\n{lab}")
    ws.row_dimensions[TQ_HDR].height = 42
    for i, d in enumerate(districts):
        r = TQ_FIRST + i
        g = quote_sheetname(gsheets[i])
        merge_put(ws, f"B{r}:C{r}", d.sheet, f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        for j, cat in enumerate(CMP):
            src = drows[cat]
            for m, col in enumerate([T_SOLL, TIME_LETTERS[0], TIME_LETTERS[-1]]):
                put(ws, f"{get_column_letter(4 + 3 * j + m)}{r}", f"={g}!${col}${src}", al=CENTER, nf="0", border=BOTTOM_HAIR,
                    f=font(10, m == 0, INK2 if m == 0 else INK))
    for j in range(len(CMP)):
        cs = get_column_letter(4 + 3 * j)
        for m in (1, 2):
            cp = get_column_letter(4 + 3 * j + m)
            add_compare_cf(ws, f"{cp}{TQ_FIRST}:{cp}{TQ_LAST}", f"{cp}{TQ_FIRST}", f"${cs}{TQ_FIRST}")
    section_title(ws, f"B{R_TQ2}:{GRID_LAST}{R_TQ2}", "Zahlen: Abweichung vom Soll je Stufe – LST gesamt (Summe der Bezirke)")
    gap_table(ws, TQ2_HDR, g_rows, row_height=38)
    last_row = TQ2_HDR + len(CMP)

    # ---- Kennzahlen
    gr = g_rows
    davon_k = 2                                  # zweite Jahresspalte = Ende Bezugsjahr + 1
    kpi_tiles(ws, [
        ("Soll (Zielzustand)", f"={T_SOLL}{gr['Köpfe']}", f'="Köpfe in "&{n}&" Bezirken"', INK2),
        ("Mitarbeiter heute", f"={TIME_LETTERS[0]}{gr['Köpfe']}", '="Köpfe (Teamleiter einmal gezählt)"', C_HEUTE),
        ("Fehlende Stellen heute", f"={TIME_LETTERS[0]}{TB_SUM}", '="Summe über Stufen und Bezirke"', BAD_INK),
        ("Fehlende Stellen", f"={TIME_LETTERS[davon_k]}{TB_SUM}", f'="am Ende von "&({E_YEAR}+{davon_k - 1})', BAD_INK),
        ("Fehlende Stellen nach Plan", f"={TIME_LETTERS[-1]}{TB_SUM}", '="wenn alle Planjahre erreicht sind"', BAD_INK),
        ("Mehr als Soll nach Plan", f"=O{TB_SUM}", f'="davon Arbeiter LST: "&{TIME_LETTERS[-1]}{gr["Arbeiter LST"]}', "EDA100"),
    ])
    how_box(ws, 8, HOW_COUNT, height=44)

    # ---- Diagramme Seite 1
    section_title(ws, f"B{R_SEC1}:J{R_SEC1}", "LST gesamt: Mitarbeiter je Stufe – rechts der Zielzustand")
    section_title(ws, f"K{R_SEC1}:{GRID_LAST}{R_SEC1}", "Fehlende Stellen zum Soll je Bezirk")
    comp_cats = rng(title, COMP_LETTERS[0], TG_HDR, COMP_LETTERS[-1], TG_HDR)
    place(ws, stacked_stage_chart(comp_cats, stack_series(title, g_rows), label_fmt='[<3]"";0'), f"B{R_CH1}", f"K{R_CH1_END}")
    cats = rng(title, "B", TB_FIRST, "B", TB_LAST)
    place(ws, clustered_chart(cats, [
        (rng(title, TIME_LETTERS[0], TB_FIRST, TIME_LETTERS[0], TB_LAST), rng(title, TIME_LETTERS[0], TB_HDR), C_HEUTE),
        (rng(title, TIME_LETTERS[-1], TB_FIRST, TIME_LETTERS[-1], TB_LAST), rng(title, TIME_LETTERS[-1], TB_HDR), C_PLAN),
    ], horizontal=True, label_size=8), f"K{R_CH1}", f"T{R_CH1_END}")

    # ---- Seite 2/3: Verlauf je Bezirk (kleine Diagramme, gleiche Skala)
    breaks = [R_SM - 1]
    for p in range(pages):
        top = R_SM + p * sm_page_rows
        part = f" ({p + 1}/{pages})" if pages > 1 else ""
        section_title(ws, f"B{top}:{GRID_LAST}{top}",
                      f"Verlauf je Bezirk: Mitarbeiter je Stufe – rechts der Zielzustand{part}")
        # Legende (Farbfeld + Text) in einer Zeile
        lr = top + 1
        for i, key in enumerate(STACK):
            col = 2 + 2 * i
            put(ws, f"{get_column_letter(col)}{lr}", None, fl=fill(STAGE_COLORS[key][0]))
            put(ws, f"{get_column_letter(col + 1)}{lr}", key, f=font(8, color=INK2), al=LEFT)
        merge_put(ws, f"Q{lr}:{GRID_LAST}{lr}", "Teamleiter zusätzlich · gleiche Skala", f=font(8, italic=True, color=MUTED), al=RIGHT)
        ws.row_dimensions[lr].height = 14
        for j in range(SM_PER_PAGE):
            i = p * SM_PER_PAGE + j
            if i >= n:
                break
            brow, pos = divmod(j, SM_PER_ROW)
            r0 = top + 2 + brow * band
            c0 = 2 + 6 * pos
            g = gsheets[i]
            dc = rng(g, COMP_LETTERS[0], B_SI_HDR - 1, COMP_LETTERS[-1], B_SI_HDR - 1)
            ch = stacked_stage_chart(dc, stack_series(g, drows), ymax=ymax, labels=False, show_legend=False, axis_size=7,
                                     title=districts[i].sheet)
            place(ws, ch, f"{get_column_letter(c0)}{r0}", f"{get_column_letter(c0 + 6)}{r0 + SM_ROWS}")
        breaks.append(top + sm_page_rows - 1)

    page_setup(ws, f"A1:{GRID_LAST}{last_row}", breaks=breaks + [R_TQ - 1])
    heights(ws, {9: 8, TG_NOTE + 1: 8, TB_SUM + 1: 8})
    fix_row_heights(ws)
    ws.sheet_properties.tabColor = "0CA30C"


# ---------------------------------------------------------------- Blatt: Pruefliste
PF_HDR = 12


def build_pruefliste(ws):
    setup_grid(ws, hidden_from=KEY_COL)
    header_block(ws, "Prüfliste – auffällige Einträge",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   "&COUNTIF({drng("pf")},1)&" Mitarbeiter mit Hinweisen   ·   '
                 f'aktualisiert sich automatisch"',
                 links=[("Übersicht", "← Übersicht"), ("Soll-Ist", "Soll-Ist →")])
    notes = [
        ("Jahr vorbei ohne „x“", "In L–P, S, U oder V steht ein Jahr vor dem Bezugsjahr, aber noch kein „x“ (bei N/O nicht, wenn U/V "
                                 "abgehakt ist). Im Verlauf zählt so ein Jahr schon zum ersten Jahresende."),
        ("Ziel ohne passende Planung", "In der Spalte der Ziel-Qualifikation steht weder „x“ noch Jahr (Status „Fehlt“), ggf. ist nur "
                                       "weiter rechts etwas geplant."),
        ("Ziel unklar", "Ziel-Qualifikation leer oder ohne Spalte (Zuordnung im Blatt Einstellungen)."),
        ("Einträge", "Ungültiger Eintrag (weder „x“ noch Jahr, z. B. „0“), nichts in L–P, Verwendungsprüfung ohne passende Spalte, "
                     "RBEG ohne SigMech."),
        ("Ist-Qualifikation", "Spalte J passt nicht zum Stand heute laut L–P (nur zur Info – gezählt wird mit L–P)."),
    ]
    for i, (k, t) in enumerate(notes):
        r = 5 + i
        merge_put(ws, f"B{r}:D{r}", k, f=font(9, True, INK2), al=LEFT)
        merge_put(ws, f"E{r}:{GRID_LAST}{r}", t, f=font(9, color=INK2), al=LEFT_WRAP)
        ws.row_dimensions[r].height = 24
    section_title(ws, f"B{PF_HDR - 1}:{GRID_LAST}{PF_HDR - 1}", "Auffällige Einträge (Reihenfolge wie in den Bezirks-Blättern)")
    cols = [("B", "B", "Nr."), ("C", "D", "Bezirk"), ("E", "F", "Name"), ("G", "H", "Vorname"), ("I", "J", "Ziel-Qualifikation"),
            ("K", "K", "Zeile"), ("L", GRID_LAST, "Hinweise")]
    for a, b, text in cols:
        head(ws, f"{a}{PF_HDR}" if a == b else f"{a}{PF_HDR}:{b}{PF_HDR}", text, LEFT_WRAP)
    ws.row_dimensions[PF_HDR].height = 20
    m = KEY_COL
    for k in range(1, PRUEF_ROWS + 1):
        r = PF_HDR + k
        put(ws, f"{m}{r}", f'=IFERROR(MATCH({k},{drng("pnr")},0),"")')
        get = lambda key: f'=IF(${m}{r}="","",INDEX({drng(key)},${m}{r}))'
        put(ws, f"B{r}", f'=IF(${m}{r}="","",{k})', f=font(9, color=MUTED), al=Alignment(horizontal="left", vertical="top"))
        merge_put(ws, f"C{r}:D{r}", get("bez"), f=font(9), al=TOP_WRAP)
        merge_put(ws, f"E{r}:F{r}", get("name"), f=font(9, True), al=TOP_WRAP)
        merge_put(ws, f"G{r}:H{r}", get("vor"), f=font(9), al=TOP_WRAP)
        merge_put(ws, f"I{r}:J{r}", get("ziel"), f=font(9), al=TOP_WRAP)
        put(ws, f"K{r}", get("row"), f=font(9, color=MUTED), al=Alignment(horizontal="center", vertical="top"))
        merge_put(ws, f"L{r}:{GRID_LAST}{r}", get("hint"), f=font(9), al=TOP_WRAP)
        ws.row_dimensions[r].height = 26
    ws.conditional_formatting.add(f"B{PF_HDR + 1}:{GRID_LAST}{PF_HDR + PRUEF_ROWS}",
                                  FormulaRule(formula=[f'${m}{PF_HDR + 1}<>""'], border=BOTTOM_HAIR))
    over = PF_HDR + PRUEF_ROWS + 1
    merge_put(ws, f"B{over}:{GRID_LAST}{over}",
              f'=IF(COUNTIF({drng("pf")},1)>{PRUEF_ROWS},"Hinweis: mehr als {PRUEF_ROWS} Einträge – vollständig im Blatt Daten '
              f'(Filter „Prüfen“ = 1).","")', f=font(9, True, STATUS[7][1]), al=LEFT)
    page_setup(ws, f"A1:{GRID_LAST}{over}", title_rows=f"{PF_HDR}:{PF_HDR}")
    fix_row_heights(ws)
    ws.sheet_properties.tabColor = "EC835A"


# ---------------------------------------------------------------- Hauptfunktion
def build(districts, zz_ws, zz_name, out_path, ymax=30):
    global DATA_LAST
    DATA_LAST = 1 + len(districts) * SRC_ROWS
    wb = Workbook()
    ws_over = wb.active
    ws_over.title = "Übersicht"
    ws_si = wb.create_sheet("Soll-Ist")
    ws_pf = wb.create_sheet("Prüfliste")
    drows = None
    for d in districts:
        drows = build_district(wb.create_sheet(district_title(d)), d)
    build_overview(ws_over, districts, drows)
    build_soll_ist(ws_si, districts, drows, ymax)
    build_pruefliste(ws_pf)
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
