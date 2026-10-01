"""Erzeugt die Dashboard-Blaetter (Uebersicht, Grafik je Bezirk, Daten, Einstellungen)
als eigenstaendige Arbeitsmappe. Die Formeln verweisen auf die Bezirks-Blaetter der
Original-Datei; merge_into_original.py setzt die Blaetter anschliessend dort ein.
"""
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import Legend
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.worksheet.properties import PageSetupProperties

# ---------------------------------------------------------------- Konstanten
SRC_FIRST_ROW = 3          # erste ausgewertete Zeile je Bezirks-Blatt
SRC_ROWS = 120             # Anzahl ausgewerteter Zeilen je Bezirks-Blatt (3..122)
LIST_ROWS = 60             # Zeilen der Mitarbeiterliste je Bezirksseite

FONT = "Arial"
INK = "0B0B0B"
INK2 = "52514E"
MUTED = "898781"
HAIR = "E1E0D9"
TILE_BG = "F4F4F2"
HEAD_BG = "EDEDEA"
INPUT_BG = "FFF2CC"

# Status-Nr -> (Farbe, Textfarbe auf der Farbe). Die Bezeichnungen stehen im Blatt
# Einstellungen (Jahres-Kategorien rechnen sich aus dem Bezugsjahr).
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
    1: ("Fertig", "Ziel-Qualifikation erreicht: In der Spalte der Ziel-Qualifikation steht ein „x“ (oder Ist-Qualifikation = Ziel-Qualifikation)."),
    2: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr."),
    3: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr + 1."),
    4: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr + 2."),
    5: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr + 3."),
    6: (None, "In Ausbildung, geplanter Abschluss im Bezugsjahr + 4 oder später."),
    7: ("Überfällig", "Geplantes Jahr liegt vor dem Bezugsjahr, aber es steht noch kein „x“ – bitte prüfen/aktualisieren."),
    8: ("Fehlt", "Ziel-Qualifikation nicht erreicht und keine Ausbildung geplant (weder „x“ noch Jahr in der Ziel-Spalte)."),
    9: ("Ziel nicht eingetragen", "In Spalte F (Ziel-Qualifikation) steht nichts."),
}

# Zuordnung Text in Spalte F/H -> Spalte J..Q und Gruppe fuer das Diagramm
MAPPING = [
    ("Azubi", "J", "Sonstige"),
    ("Arb LST", "K", "Arb LST"),
    ("Weichmech", "L", "Weichmech"),
    ("Weichenmechaniker", "L", "Weichmech"),
    ("Sigmech", "M", "Sigmech"),
    ("Signalmechaniker", "M", "Sigmech"),
    ("Sigmech nur für Weichen", "M", "Sigmech"),
    ("Sigmech RBEG", "N", "Sigmech RBEG"),
    ("Signalmechaniker RBEG", "N", "Sigmech RBEG"),
    ("Kennziffer 4", "O", "Sonstige"),
    ("IHK-Meister", "P", "IHK-Meister"),
    ("TL", "Q", "Teamleiter"),
    ("Teamleiter", "Q", "Teamleiter"),
    ("Senior Expert LST", "", "Sonstige"),
]
GROUPS = ["Sigmech RBEG", "Sigmech", "Weichmech", "Arb LST", "IHK-Meister", "Teamleiter", "Sonstige"]

# Einstellungen: feste Zellpositionen
E_YEAR = "Einstellungen!$C$4"
E_STAT_FIRST, E_STAT_LAST = 8, 16           # Status-Tabelle (Nr in B, Bezeichnung in C)
E_MAP_FIRST, E_MAP_LAST = 21, 50            # Zuordnungstabelle
E_GRP_FIRST = 54                            # Diagramm-Gruppen (7 Zeilen)
E_STAT_LABEL = f"Einstellungen!$C${E_STAT_FIRST}:$C${E_STAT_LAST}"
E_MAP_TEXT = f"Einstellungen!$B${E_MAP_FIRST}:$B${E_MAP_LAST}"
E_MAP_COLNO = f"Einstellungen!$D${E_MAP_FIRST}:$D${E_MAP_LAST}"
E_MAP_GROUP = f"Einstellungen!$E${E_MAP_FIRST}:$E${E_MAP_LAST}"


def stat_label_ref(code):
    return f"Einstellungen!$C${E_STAT_FIRST + code - 1}"


# ---------------------------------------------------------------- Stil-Helfer
def font(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


HAIR_SIDE = Side(style="thin", color=HAIR)
BOTTOM_HAIR = Border(bottom=HAIR_SIDE)
LEFT = Alignment(horizontal="left", vertical="center")
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


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


def page_setup(ws, print_area):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_area = print_area
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5


def text_props(size=9, color=INK2, bold=False):
    cp = CharacterProperties(sz=int(size * 100), b=bold, solidFill=color,
                             latin=DFont(typeface=FONT), cs=DFont(typeface=FONT))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def no_line():
    return LineProperties(noFill=True)


def data_labels(color, pos=None, numfmt=None, bold=False):
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
    dl.txPr = text_props(9, color, bold)
    return dl


def style_chart_frame(chart):
    chart.graphical_properties = GraphicalProperties(ln=no_line())
    chart.roundedCorners = False


def status_bar_chart(ws, val_row, val_c1, val_c2, cat_row, width, height):
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
    ch.x_axis.scaling.orientation = "maxMin"     # erste Kategorie oben
    ch.x_axis.delete = False
    ch.x_axis.majorTickMark = "none"
    ch.x_axis.txPr = text_props(9, INK2)
    ch.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="C3C2B7"))
    ch.y_axis.delete = True
    ch.y_axis.majorGridlines = None
    ch.y_axis.scaling.min = 0
    ch.width, ch.height = width, height
    style_chart_frame(ch)
    return ch


def stacked_status_chart(ws, hdr_row, first_row, last_row, cat_col, c1, c2, width, height):
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
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.delete = False
    ch.x_axis.majorTickMark = "none"
    ch.x_axis.txPr = text_props(9, INK2)
    ch.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="C3C2B7"))
    ch.y_axis.delete = True
    ch.y_axis.majorGridlines = None
    ch.y_axis.scaling.min = 0
    ch.width, ch.height = width, height
    style_chart_frame(ch)
    return ch


# ---------------------------------------------------------------- Bausteine
GRID_COLS = 16   # Spalten B..Q
GRID_WIDTH = 9.3


def setup_grid(ws, helper_cols=("S", "T")):
    ws.column_dimensions["A"].width = 2
    for i in range(2, 2 + GRID_COLS):
        ws.column_dimensions[get_column_letter(i)].width = GRID_WIDTH
    ws.column_dimensions["R"].width = 2
    for col in helper_cols:
        ws.column_dimensions[col].hidden = True
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[4].height = 8


def header_block(ws, title, subtitle_formula, back_link=True):
    merge_put(ws, "B2:M2", title, f=font(18, True), al=LEFT)
    ws.row_dimensions[2].height = 30
    if back_link:
        merge_put(ws, "N2:Q2", '=HYPERLINK("#\'Übersicht\'!A1","← Zur Gesamtübersicht")',
                  f=font(10, color="1C5CAB", underline="single"), al=RIGHT)
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


def status_header_cells(ws, row, first_col):
    """Status-Bezeichnungen (aus Einstellungen) als Spaltenkoepfe."""
    for code in range(1, 10):
        col = get_column_letter(first_col + code - 1)
        put(ws, f"{col}{row}", f"={stat_label_ref(code)}", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)


def pct_sub(count_ref, total_ref):
    return f'=IF({total_ref}=0,"–",TEXT({count_ref}/{total_ref},"0%")&" von "&{total_ref}&" Mitarbeitern")'


# ---------------------------------------------------------------- Blatt: Einstellungen
def build_settings(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 90
    put(ws, "B1", "Einstellungen & Erklärung", f=font(16, True))
    put(ws, "B2", "Gelb hinterlegte Zellen dürfen geändert werden. Alle Diagramme passen sich automatisch an.", f=font(10, color=INK2))

    put(ws, "B4", "Bezugsjahr (aktuelles Jahr)", f=font(10, True), al=LEFT)
    put(ws, "C4", "=YEAR(TODAY())", f=font(11, True, "0000FF"), fl=fill(INPUT_BG), al=CENTER, nf="0")
    put(ws, "D4", "Standard: =JAHR(HEUTE()) – kann mit einer festen Jahreszahl überschrieben werden (z. B. 2027). "
                  "Geplante Jahre davor gelten als „Überfällig“.", f=font(9, color=INK2), al=LEFT)

    put(ws, "B6", "Status-Kategorien", f=font(12, True))
    for col, text in zip("BCDF", ["Nr.", "Bezeichnung", "Farbe", "Bedeutung"]):
        put(ws, f"{col}7", text, f=font(10, True, INK2), fl=fill(HEAD_BG), al=LEFT)
    put(ws, "E7", None, fl=fill(HEAD_BG))
    for code, color, _ in STATUS:
        r = E_STAT_FIRST + code - 1
        label, meaning = STATUS_TEXT[code]
        if label is None:
            off = code - 2
            label = (f'="Abschluss "&({E_YEAR}+{off})' if code < 6 else f'="Abschluss ab "&({E_YEAR}+{off})')
        put(ws, f"B{r}", code, al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"C{r}", label, f=font(10, True), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"D{r}", None, fl=fill(color), border=Border(bottom=Side(style="thin", color="FFFFFF")))
        put(ws, f"E{r}", None, border=BOTTOM_HAIR)
        put(ws, f"F{r}", meaning, f=font(10, color=INK2), al=LEFT, border=BOTTOM_HAIR)

    put(ws, "B19", "Zuordnung Qualifikation → Spalte (J–Q)", f=font(12, True))
    put(ws, "F19", "Wird in Spalte F (Ziel) oder H (Ist) ein neuer Begriff verwendet, hier eine Zeile ergänzen.", f=font(9, color=INK2))
    for col, text in zip("BCDEF", ["Text in Spalte F / H", "Spalte mit dem Stand (J–Q)", "Spalten-Nr. (automatisch)",
                                   "Gruppe im Diagramm", "Hinweis"]):
        put(ws, f"{col}20", text, f=font(10, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    ws.row_dimensions[20].height = 28
    hints = {"Sigmech nur für Weichen": "Annahme: wird wie Sigmech (Spalte M) gewertet",
             "Senior Expert LST": "keine eigene Spalte – gilt als fertig, wenn Ist = Ziel",
             "Kennziffer 4": "Zusatzqualifikation, i. d. R. kein Ziel"}
    for i in range(E_MAP_LAST - E_MAP_FIRST + 1):
        r = E_MAP_FIRST + i
        text, col, grp = MAPPING[i] if i < len(MAPPING) else (None, None, None)
        put(ws, f"B{r}", text, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"C{r}", col or None, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"D{r}", f'=IF(C{r}="","",IFERROR(IF(AND(CODE(UPPER(C{r}))>=74,CODE(UPPER(C{r}))<=81),CODE(UPPER(C{r}))-73,""),""))',
            al=CENTER, border=BOTTOM_HAIR)
        put(ws, f"E{r}", grp, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
        put(ws, f"F{r}", hints.get(text), f=font(9, color=INK2), al=LEFT, border=BOTTOM_HAIR)

    put(ws, f"B{E_GRP_FIRST - 1}", "Gruppen im Diagramm „Status nach Ziel-Qualifikation“ (Reihenfolge von oben nach unten)", f=font(12, True))
    for i, g in enumerate(GROUPS):
        put(ws, f"B{E_GRP_FIRST + i}", g, f=font(10, color="0000FF"), fl=fill(INPUT_BG), al=LEFT, border=BOTTOM_HAIR)
    put(ws, f"C{E_GRP_FIRST + len(GROUPS) - 1}", "← alles, was keiner Gruppe zugeordnet ist, landet hier", f=font(9, color=INK2))

    r = E_GRP_FIRST + len(GROUPS) + 2
    put(ws, f"B{r}", "So wird gezählt", f=font(12, True))
    notes = [
        "Ausgewertet werden die zwölf Bezirks-Blätter (FFM Hbf … Limburg FUB & R WEW). Das Blatt „Bezirksleiter FBÜW“ wird nicht berücksichtigt.",
        f"Je Blatt werden die Zeilen {SRC_FIRST_ROW}–{SRC_FIRST_ROW + SRC_ROWS - 1} gelesen. Eine Zeile zählt als Mitarbeiter, wenn ein Name (Spalte A) "
        "und eine Ziel- oder Ist-Qualifikation (Spalte F oder H) eingetragen ist.",
        "Für jeden Mitarbeiter wird die Spalte (J–Q) gesucht, die zu seiner Ziel-Qualifikation gehört (Tabelle oben). Steht dort „x“, ist er fertig; "
        "steht dort eine Jahreszahl (z. B. 27 = 2027), ist das der geplante Abschluss.",
        "Neue Mitarbeiter einfach im jeweiligen Bezirks-Blatt eintragen (auch per Zeile einfügen) – Übersicht und Bezirksseiten aktualisieren sich automatisch.",
        "Das Blatt „Daten“ enthält die Auswertung je Mitarbeiter als Tabelle. Es wird per Formel erzeugt und sollte nicht von Hand geändert werden "
        "(eignet sich aber z. B. als Quelle für Power BI).",
    ]
    for i, n in enumerate(notes):
        merge_put(ws, f"B{r + 1 + i}:F{r + 1 + i}", "• " + n, f=font(10, color=INK2), al=LEFT_WRAP)
        ws.row_dimensions[r + 1 + i].height = 28
    page_setup(ws, f"A1:F{r + len(notes)}")


# ---------------------------------------------------------------- Blatt: Daten
DATA_HEADERS = ["Bezirk", "Quellzeile", "Name", "Vorname", "Ziel-Qualifikation", "Ist-Qualifikation", "Gültig",
                "Ziel-Nr", "Ziel-Spalte", "Ziel-Gruppe", "Ist-Nr", "Ist-Spalte", "Eintrag Ziel-Spalte",
                "Planjahr", "Fertig", "Status-Nr", "Status", "Sortierung", "Rang im Bezirk", "Schlüssel"]


def build_data(ws, districts):
    ws.sheet_view.showGridLines = True
    widths = [20, 10, 20, 16, 22, 22, 8, 8, 10, 14, 8, 10, 14, 10, 8, 10, 22, 11, 11, 24]
    for i, (h, w) in enumerate(zip(DATA_HEADERS, widths), start=1):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = w
        put(ws, f"{col}1", h, f=font(10, True, "FFFFFF"), fl=fill("52514E"), al=LEFT_WRAP)
    ws.freeze_panes = "C2"
    last = 1 + len(districts) * SRC_ROWS
    ws.auto_filter.ref = f"A1:T{last}"
    r = 2
    plain = font(10)
    for d in districts:
        s = quote_sheetname(d)
        for k in range(SRC_ROWS):
            src = SRC_FIRST_ROW + k
            f = {
                "A": d,
                "B": src,
                "C": f'=TRIM(INDEX({s}!$A:$A,$B{r})&"")',
                "D": f'=TRIM(INDEX({s}!$B:$B,$B{r})&"")',
                "E": f'=TRIM(INDEX({s}!$F:$F,$B{r})&"")',
                "F": f'=TRIM(INDEX({s}!$H:$H,$B{r})&"")',
                "G": f'=IF(AND(C{r}<>"",C{r}<>"Name",OR(E{r}<>"",F{r}<>"")),1,0)',
                "H": f'=IF(E{r}="",0,IFERROR(MATCH(E{r},{E_MAP_TEXT},0),0))',
                "I": f'=IF(H{r}=0,0,IFERROR(INDEX({E_MAP_COLNO},H{r})*1,0))',
                "J": f'=IF(H{r}=0,"Sonstige",IF(INDEX({E_MAP_GROUP},H{r})&""="","Sonstige",INDEX({E_MAP_GROUP},H{r})&""))',
                "K": f'=IF(F{r}="",0,IFERROR(MATCH(F{r},{E_MAP_TEXT},0),0))',
                "L": f'=IF(K{r}=0,0,IFERROR(INDEX({E_MAP_COLNO},K{r})*1,0))',
                "M": f'=IF(I{r}=0,"",IF(INDEX({s}!$J:$Q,$B{r},I{r})&""="","",INDEX({s}!$J:$Q,$B{r},I{r})))',
                "N": f'=IF(M{r}="","",IFERROR(IF(M{r}*1>2100,YEAR(M{r}*1),IF(M{r}*1>=1900,ROUND(M{r}*1,0),2000+ROUND(M{r}*1,0))),""))',
                "O": f'=IF(G{r}=0,0,IF(OR(LOWER(TRIM(M{r}&""))="x",AND(E{r}<>"",LOWER(E{r})=LOWER(F{r})),AND(I{r}>0,I{r}=L{r})),1,0))',
                "P": f'=IF(G{r}=0,"",IF(E{r}="",9,IF(O{r}=1,1,IF(N{r}="",8,IF(N{r}<{E_YEAR},7,MIN(6,2+N{r}-{E_YEAR}))))))',
                "Q": f'=IF(P{r}="","",INDEX({E_STAT_LABEL},P{r}))',
                "R": f'=IF(P{r}="","",P{r}*1000+B{r})',
                "S": f'=IF(R{r}="","",COUNTIFS($A$2:$A${last},A{r},$R$2:$R${last},"<"&R{r})+1)',
                "T": f'=IF(S{r}="","",A{r}&"|"&S{r})',
            }
            for col, v in f.items():
                c = ws[f"{col}{r}"]
                c.value = v
                c.font = plain
            ws[f"N{r}"].number_format = "0"
            r += 1
    return last


# ---------------------------------------------------------------- Blatt: Grafik je Bezirk
MAT_HDR = 30                         # Kopfzeile der Zahlentabelle
MAT_FIRST = MAT_HDR + 1              # erste Gruppe
MAT_LAST = MAT_HDR + len(GROUPS)     # letzte Gruppe
MAT_SUM = MAT_LAST + 1               # Summenzeile (= Quelle Diagramm 1)
MAT_PCT = MAT_SUM + 1
LIST_TITLE = MAT_PCT + 2
LIST_HDR = LIST_TITLE + 1
LIST_FIRST = LIST_HDR + 1
LIST_LAST = LIST_FIRST + LIST_ROWS - 1
STAT_COL1 = 4                        # Spalte D = Status 1 ... Spalte L = Status 9
SUM_COL = 13                         # Spalte M


def build_district(ws, d):
    setup_grid(ws)
    put(ws, "S1", d)
    header_block(ws, f"Ausbildungsstand – {d}",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   Quelle: Blatt „{d}“ (Ziel-Qualifikation Spalte F, Ist-Qualifikation Spalte H, Stand Spalten J–Q)'
                 f'   ·   aktualisiert sich automatisch"')

    sc = lambda code: f"{get_column_letter(STAT_COL1 + code - 1)}{MAT_SUM}"
    total = f"$M${MAT_SUM}"
    in_training = f"SUM({sc(2)}:{sc(6)})"
    kpi_tiles(ws, [
        ("Mitarbeiter", f"={total}", f'=IF({sc(9)}>0,"davon "&{sc(9)}&" ohne Ziel-Qualifikation","im Bezirk")', INK2),
        ("Fertig", f"={sc(1)}", pct_sub(sc(1), total), STATUS[0][1]),
        ("In Ausbildung", f"={in_training}", pct_sub(in_training, total), "2A78D6"),
        ("Überfällig", f"={sc(7)}", pct_sub(sc(7), total), STATUS[6][1]),
        ("Fehlt (nichts geplant)", f"={sc(8)}", pct_sub(sc(8), total), STATUS[7][1]),
    ])

    section_title(ws, "B9:I9", "Mitarbeiter nach Status")
    section_title(ws, "J9:Q9", "Status nach Ziel-Qualifikation")
    ws.row_dimensions[9].height = 20

    # Zahlentabelle
    section_title(ws, f"B{MAT_HDR - 1}:Q{MAT_HDR - 1}", "Zahlen zu den Diagrammen")
    merge_put(ws, f"B{MAT_HDR}:C{MAT_HDR}", "Ziel-Qualifikation", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    status_header_cells(ws, MAT_HDR, STAT_COL1)
    put(ws, f"M{MAT_HDR}", "Summe", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    ws.row_dimensions[MAT_HDR].height = 30
    for i in range(len(GROUPS)):
        r = MAT_FIRST + i
        merge_put(ws, f"B{r}:C{r}", f"=Einstellungen!$B${E_GRP_FIRST + i}", al=LEFT, border=BOTTOM_HAIR)
        for code in range(1, 10):
            col = get_column_letter(STAT_COL1 + code - 1)
            put(ws, f"{col}{r}", f'=COUNTIFS(Daten!$A:$A,$S$1,Daten!$J:$J,$B{r},Daten!$P:$P,{code})',
                al=CENTER, nf="0;-0;\"·\"", border=BOTTOM_HAIR)
        put(ws, f"M{r}", f"=SUM(D{r}:L{r})", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
    merge_put(ws, f"B{MAT_SUM}:C{MAT_SUM}", "Summe", f=font(10, True), al=LEFT)
    merge_put(ws, f"B{MAT_PCT}:C{MAT_PCT}", "Anteil", f=font(9, color=INK2), al=LEFT)
    for c in range(STAT_COL1, SUM_COL + 1):
        col = get_column_letter(c)
        put(ws, f"{col}{MAT_SUM}", f"=SUM({col}{MAT_FIRST}:{col}{MAT_LAST})", f=font(10, True), al=CENTER, nf="0",
            border=Border(top=Side(style="thin", color=INK2)))
        put(ws, f"{col}{MAT_PCT}", f'=IF({total}=0,"",{col}{MAT_SUM}/{total})', f=font(9, color=INK2), al=CENTER, nf="0%")

    # Diagramme
    ch1 = status_bar_chart(ws, MAT_SUM, STAT_COL1, STAT_COL1 + 8, MAT_HDR, width=15.6, height=9.6)
    ws.add_chart(ch1, "B10")
    ch2 = stacked_status_chart(ws, MAT_HDR, MAT_FIRST, MAT_LAST, 2, STAT_COL1, STAT_COL1 + 8, width=15.6, height=9.6)
    ws.add_chart(ch2, "J10")

    # Mitarbeiterliste
    section_title(ws, f"B{LIST_TITLE}:Q{LIST_TITLE}", "Mitarbeiterliste (sortiert nach Status)")
    cols = [("B", "B", "Nr."), ("C", "D", "Name"), ("E", "F", "Vorname"), ("G", "H", "Ziel-Qualifikation"),
            ("I", "J", "Ist-Qualifikation"), ("K", "L", "Geplanter Abschluss"), ("M", "P", "Status")]
    for a, b, text in cols:
        rng = f"{a}{LIST_HDR}:{b}{LIST_HDR}"
        if a == b:
            put(ws, f"{a}{LIST_HDR}", text, f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT)
        else:
            merge_put(ws, rng, text, f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT)
    for k in range(1, LIST_ROWS + 1):
        r = LIST_FIRST + k - 1
        put(ws, f"S{r}", f'=IFERROR(MATCH($S$1&"|"&{k},Daten!$T:$T,0),"")')
        put(ws, f"T{r}", f'=IF($S{r}="","",INDEX(Daten!$P:$P,$S{r}))')
        put(ws, f"B{r}", f'=IF($S{r}="","",{k})', f=font(10, color=MUTED), al=LEFT)
        merge_put(ws, f"C{r}:D{r}", f'=IF($S{r}="","",INDEX(Daten!$C:$C,$S{r}))', f=font(10, True), al=LEFT)
        merge_put(ws, f"E{r}:F{r}", f'=IF($S{r}="","",INDEX(Daten!$D:$D,$S{r}))', al=LEFT)
        merge_put(ws, f"G{r}:H{r}", f'=IF($S{r}="","",INDEX(Daten!$E:$E,$S{r}))', al=LEFT)
        merge_put(ws, f"I{r}:J{r}", f'=IF($S{r}="","",INDEX(Daten!$F:$F,$S{r}))', al=LEFT)
        merge_put(ws, f"K{r}:L{r}", f'=IF($S{r}="","",IF(INDEX(Daten!$N:$N,$S{r})="","–",INDEX(Daten!$N:$N,$S{r})))', al=LEFT, nf="0")
        merge_put(ws, f"M{r}:P{r}", f'=IF($S{r}="","",INDEX(Daten!$Q:$Q,$S{r}))', f=font(10, True), al=Alignment(horizontal="left", vertical="center", indent=1))
        ws.row_dimensions[r].height = 17
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
    page_setup(ws, f"A1:Q{over}")
    ws.sheet_properties.tabColor = "86B6EF"


# ---------------------------------------------------------------- Blatt: Uebersicht
OV_HDR = 36
OV_FIRST = OV_HDR + 1


def build_overview(ws, districts):
    setup_grid(ws, helper_cols=("S",))
    n = len(districts)
    ov_last = OV_FIRST + n - 1
    ov_sum = ov_last + 1
    ov_pct = ov_sum + 1
    header_block(ws, "Ausbildungsstand LST – Gesamtübersicht aller Bezirke",
                 f'="Bezugsjahr "&{E_YEAR}&"   ·   {n} Bezirke   ·   Klick auf einen Bezirk in der Tabelle unten öffnet seine Seite   ·   '
                 f'Werte aktualisieren sich automatisch, wenn die Bezirks-Blätter geändert werden"',
                 back_link=False)

    sc = lambda code: f"{get_column_letter(4 + code)}{ov_sum}"   # Status 1 -> Spalte E
    total = f"$N${ov_sum}"
    in_training = f"SUM({sc(2)}:{sc(6)})"
    kpi_tiles(ws, [
        ("Mitarbeiter gesamt", f"={total}", f'=IF({sc(9)}>0,"davon "&{sc(9)}&" ohne Ziel-Qualifikation","in allen Bezirken")', INK2),
        ("Fertig", f"={sc(1)}", pct_sub(sc(1), total), STATUS[0][1]),
        ("In Ausbildung", f"={in_training}", pct_sub(in_training, total), "2A78D6"),
        ("Überfällig", f"={sc(7)}", pct_sub(sc(7), total), STATUS[6][1]),
        ("Fehlt (nichts geplant)", f"={sc(8)}", pct_sub(sc(8), total), STATUS[7][1]),
    ])
    section_title(ws, "B9:G9", "Alle Bezirke nach Status")
    section_title(ws, "H9:Q9", "Status je Bezirk")
    ws.row_dimensions[9].height = 20

    section_title(ws, f"B{OV_HDR - 1}:Q{OV_HDR - 1}", "Zahlen je Bezirk")
    merge_put(ws, f"B{OV_HDR}:D{OV_HDR}", "Bezirk (Klick öffnet die Seite)", f=font(9, True, INK2), fl=fill(HEAD_BG), al=LEFT_WRAP)
    status_header_cells(ws, OV_HDR, 5)
    put(ws, f"N{OV_HDR}", "Summe", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    merge_put(ws, f"O{OV_HDR}:P{OV_HDR}", "Anteil fertig", f=font(9, True, INK2), fl=fill(HEAD_BG), al=CENTER_WRAP)
    ws.row_dimensions[OV_HDR].height = 30
    for i, d in enumerate(districts):
        r = OV_FIRST + i
        put(ws, f"S{r}", d)
        link = "#" + quote_sheetname("Grafik " + d).replace('"', '""') + "!A1"
        merge_put(ws, f"B{r}:D{r}", f'=HYPERLINK("{link}","{d}")', f=font(10, True, "1C5CAB", underline="single"),
                  al=LEFT, border=BOTTOM_HAIR)
        for code in range(1, 10):
            col = get_column_letter(4 + code)
            put(ws, f"{col}{r}", f"=COUNTIFS(Daten!$A:$A,$S{r},Daten!$P:$P,{code})", al=CENTER, nf="0;-0;\"·\"", border=BOTTOM_HAIR)
        put(ws, f"N{r}", f"=SUM(E{r}:M{r})", f=font(10, True), al=CENTER, nf="0", border=BOTTOM_HAIR)
        merge_put(ws, f"O{r}:P{r}", f'=IF(N{r}=0,"",E{r}/N{r})', al=CENTER, nf="0%", border=BOTTOM_HAIR)
    merge_put(ws, f"B{ov_sum}:D{ov_sum}", "Summe", f=font(10, True), al=LEFT)
    merge_put(ws, f"B{ov_pct}:D{ov_pct}", "Anteil", f=font(9, color=INK2), al=LEFT)
    top = Border(top=Side(style="thin", color=INK2))
    for c in range(5, 15):
        col = get_column_letter(c)
        put(ws, f"{col}{ov_sum}", f"=SUM({col}{OV_FIRST}:{col}{ov_last})", f=font(10, True), al=CENTER, nf="0", border=top)
        put(ws, f"{col}{ov_pct}", f'=IF({total}=0,"",{col}{ov_sum}/{total})', f=font(9, color=INK2), al=CENTER, nf="0%")
    merge_put(ws, f"O{ov_sum}:P{ov_sum}", f'=IF({total}=0,"",E{ov_sum}/{total})', f=font(10, True), al=CENTER, nf="0%", border=top)

    ch1 = status_bar_chart(ws, ov_sum, 5, 13, OV_HDR, width=11.8, height=12.2)
    ws.add_chart(ch1, "B10")
    ch2 = stacked_status_chart(ws, OV_HDR, OV_FIRST, ov_last, 2, 5, 13, width=19.6, height=12.2)
    ws.add_chart(ch2, "H10")

    # Legende / Erklaerung
    lg = ov_pct + 2
    section_title(ws, f"B{lg}:Q{lg}", "So wird gezählt")
    for code, color, _ in STATUS:
        r = lg + code
        put(ws, f"B{r}", None, fl=fill(color))
        merge_put(ws, f"C{r}:E{r}", f"={stat_label_ref(code)}", f=font(10, True), al=Alignment(horizontal="left", vertical="center", indent=1))
        merge_put(ws, f"F{r}:Q{r}", f"=Einstellungen!$F${E_STAT_FIRST + code - 1}", f=font(9, color=INK2), al=LEFT)
    foot = lg + 11
    merge_put(ws, f"B{foot}:Q{foot}",
              "Bezugsjahr, Zuordnung der Qualifikationen und Erklärungen: Blatt „Einstellungen“.   Auswertung je Mitarbeiter: Blatt „Daten“.",
              f=font(9, italic=True, color=MUTED), al=LEFT)
    page_setup(ws, f"A1:Q{foot}")
    ws.sheet_properties.tabColor = "104281"


# ---------------------------------------------------------------- Hauptfunktion
def build(districts, out_path):
    wb = Workbook()
    ws_over = wb.active
    ws_over.title = "Übersicht"
    build_overview(ws_over, districts)
    for d in districts:
        build_district(wb.create_sheet("Grafik " + d), d)
    ws_data = wb.create_sheet("Daten")
    build_data(ws_data, districts)
    ws_data.sheet_properties.tabColor = "C3C2B7"
    ws_set = wb.create_sheet("Einstellungen")
    build_settings(ws_set)
    ws_set.sheet_properties.tabColor = "FAB219"
    wb.save(out_path)
    return [ws.title for ws in wb.worksheets]


if __name__ == "__main__":
    import sys
    from openpyxl import load_workbook
    src, out = sys.argv[1], sys.argv[2]
    names = [n for n in load_workbook(src, read_only=True).sheetnames if n != "Bezirksleiter FBÜW"]
    print(build(names, out))
