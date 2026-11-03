"""Auswertungen: Statistiken, Dienstverteilung, Urlaubs-/Krankenstatistik, Monats-/Jahresbericht,
Einsatzfaehigkeitsbericht, Dashboard und Startseite."""
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import DataBarRule
from openpyxl.worksheet.hyperlink import Hyperlink

from pp_base import (ABW_CODES, AMPEL, DATE_FMT, DUTY_CODES, GREY, HEAD, HEAD_DARK, HOLIDAY, HOLIDAY_FONT,
                     INPUT_BG, INPUT_HEAD, MON3, MONATE, MUTED, NDAYS, NMAX, TEAMS, THIN, WEEKEND, L, ampel_cf,
                     col_widths, fill, font, header_row, landscape, put, style, team_cf, title)
from pp_plan import C_FIRST, C_MON0, EF_FIRST, EF_LAST, NOK, OK_, PEP_D0, PEP_FIRST, dcol

X = "_xlfn."
P_FIRST = 6                     # erste Personenzeile in den Auswertungsblaettern
P_LAST = P_FIRST + NMAX - 1
SERIES1, SERIES2 = "2A78D6", "EB6834"      # validierte Kategoriefarben (Slot 1/2)
ST_GOOD, ST_WARN, ST_BAD, ST_CRIT = "1BAF7A", "EDA100", "E34948", "A61B1B"


def prow(i):
    r = PEP_FIRST + i
    return f"'PEP 2027'!$F{r}:${dcol(NDAYS - 1, PEP_D0)}{r}"


def prev_row(i):
    r = PEP_FIRST + i
    return f"'PEP 2027'!$E{r}:${dcol(NDAYS - 2, PEP_D0)}{r}"


def person_cols(ws, i, r, cols="ABC"):
    cr = C_FIRST + i
    src = {"A": "D", "B": "E", "C": "F"}
    for c in cols:
        ws[f"{c}{r}"] = f"=Hilfsberechnung!${src[c]}{cr}"


def calc_mon_col(m, k):
    """Spalte in der Hilfsberechnung: k=0 Soll, 1 Ist, 2 JAZ fuer Monat m (1..12)."""
    return L(C_MON0 + 3 * (m - 1) + k)


def calc_mon_sum(m, k):
    c = calc_mon_col(m, k)
    return f"SUM(Hilfsberechnung!${c}${C_FIRST}:${c}${C_FIRST + NMAX - 1})"


def body(ws, r, c0, c1, size=9, center_from=3, fmt=None):
    for c in range(c0, c1 + 1):
        style(ws.cell(r, c), size=size, h="center" if c >= center_from else None, border=THIN, fmt=fmt)


def team_summary(book, ws, r0, title_text, cols, name_col_rng, team_rng):
    """Teamtabelle unter Personenlisten: cols = [(Ueberschrift, Bereich, 'sum'|'avg')]"""
    put(ws, f"A{r0}", title_text, size=12, bold=True)
    header_row(ws, r0 + 1, 1, ["Team", "Mitarbeiter"] + [c[0] for c in cols], height=30, size=8)
    for k, (t, kk, col) in enumerate(TEAMS + [("Gesamt", "", GREY)]):
        r = r0 + 2 + k
        put(ws, f"A{r}", t, bold=True, bg=col, border=THIN)
        crit = f'{team_rng},"{t}"' if t != "Gesamt" else f'{name_col_rng},"?*"'
        put(ws, f"B{r}", f"=COUNTIFS({crit})", h="center", border=THIN)
        for j, (h, rng, mode) in enumerate(cols):
            if mode == "sum":
                f = f"=SUMIFS({rng},{crit})"
            else:
                f = f"=IF(B{r}=0,0,SUMIFS({rng},{crit})/B{r})"
            put(ws, f"{L(3 + j)}{r}", f, h="center", border=THIN, fmt="0.0" if mode == "avg" else None)
    return r0 + 2 + len(TEAMS)


# ====================================================================== Statistiken
def build_statistik(book, ws):
    title(ws, "Statistiken je Mitarbeiter (Dienst-, Abwesenheits- und Nachtdienstverteilung)",
          "Alle Werte automatisch aus PEP 2027. Urlaub/Krank/Unterricht = Arbeitstage. Wochenend-Nachtdienst = "
          "Fr–So-Block. Bereitschaften = Anzahl Bereitschaftsblöcke (Mo–Do bzw. Fr–So) mit mindestens einem Einsatztag.")
    labels = ["Name", "Team", "Teamfunktion", "Nachtdienste", "davon Wochenende (Fr–So)", "Feiertagsdienste",
              "Bereitschaften (Blöcke)", "Tagdienste", "Urlaubstage", "Krankheitstage", "Unterrichtstage",
              "JAZ-Tage", "Fortbildung", "Sonderurlaub", "Freistellung (manuell)", "Freistellung nach Nacht",
              "Vertretungen (V)", "Abweichung Nachtdienste vom Ø", "Verteilung"]
    hr = P_FIRST - 1
    header_row(ws, hr, 1, labels, height=42, size=8)
    for i in range(NMAX):
        r = P_FIRST + i
        pr, pv = prow(i), prev_row(i)
        person_cols(ws, i, r)
        e = f'IF($A{r}="","",'
        ws[f"D{r}"] = f'={e}COUNTIF({pr},"N")+COUNTIF({pr},"V"))'
        ws[f"E{r}"] = f'={e}COUNTIFS({pr},"N",PEP_FrSo,1)+COUNTIFS({pr},"V",PEP_FrSo,1))'
        ws[f"F{r}"] = f'={e}COUNTIFS({pr},"N",PEP_FT,1)+COUNTIFS({pr},"V",PEP_FT,1))'
        ws[f"G{r}"] = (f'={e}SUMPRODUCT((({pr}="N")+({pr}="V"))*(((PEP_BS=1)+({pv}<>"N")*({pv}<>"V"))>0)))')
        ws[f"H{r}"] = f'={e}COUNTIF({pr},"T"))'
        ws[f"I{r}"] = f'={e}COUNTIFS({pr},"U",PEP_Arb,1))'
        ws[f"J{r}"] = f'={e}COUNTIFS({pr},"K",PEP_Arb,1))'
        ws[f"K{r}"] = f'={e}COUNTIFS({pr},"UR",PEP_Arb,1)+COUNTIFS({pr},"FDL",PEP_Arb,1)+COUNTIFS({pr},"KAB",PEP_Arb,1))'
        ws[f"L{r}"] = f'={e}COUNTIFS({pr},"JAZ",PEP_Arb,1))'
        ws[f"M{r}"] = f'={e}COUNTIFS({pr},"FB",PEP_Arb,1))'
        ws[f"N{r}"] = f'={e}COUNTIFS({pr},"SU",PEP_Arb,1))'
        ws[f"O{r}"] = f'={e}COUNTIFS({pr},"FS",PEP_Arb,1))'
        ws[f"P{r}"] = f'={e}COUNTIF({pr},"F"))'
        ws[f"Q{r}"] = f'={e}COUNTIF({pr},"V"))'
        ws[f"R{r}"] = (f'={e}D{r}-AVERAGEIFS($D${P_FIRST}:$D${P_LAST},$A${P_FIRST}:$A${P_LAST},"?*",'
                       f'Hilfsberechnung!$G${C_FIRST}:$G${C_FIRST + NMAX - 1},"<>NEIN"))')
        ws[f"S{r}"] = f'={e}IF(ABS(R{r})<=2,"ausgeglichen",IF(R{r}>0,"über Ø","unter Ø")))'
        body(ws, r, 1, 19)
        ws[f"R{r}"].number_format = "+0.0;-0.0;0.0"
    book.table(ws, "tbl_Stat", f"A{hr}:S{P_LAST}")
    rng = lambda c: f"{c}{P_FIRST}:{c}{P_LAST}"
    for n, c in [("ST_Name", "A"), ("ST_Team", "B"), ("ST_N", "D"), ("ST_WE", "E"), ("ST_FT", "F"), ("ST_BB", "G"),
                 ("ST_T", "H"), ("ST_Urlaub", "I"), ("ST_Krank", "J"), ("ST_UR", "K"), ("ST_JAZ", "L"), ("ST_V", "Q")]:
        book.name_ws(n, ws, rng(c))
    f = P_FIRST
    team_cf(book, ws, rng("B"), f"B{f}")
    ws.conditional_formatting.add(rng("D"), DataBarRule(start_type="num", start_value=0, end_type="max",
                                                         color=SERIES1, showValue=True))
    book.cf(ws, rng("S"), f'S{f}="ausgeglichen"', bg="C6EFCE", color="006100")
    book.cf(ws, rng("S"), f'S{f}="über Ø"', bg="FFEB9C", color="9C5700")
    book.cf(ws, rng("S"), f'S{f}="unter Ø"', bg="DDEBF7", color="1F4E78")
    team_summary(book, ws, P_LAST + 3, "Teamübersicht",
                 [("Σ Nachtdienste", "ST_N", "sum"), ("Ø Nachtdienste je MA", "ST_N", "avg"),
                  ("Σ Wochenende", "ST_WE", "sum"), ("Σ Feiertagsdienste", "ST_FT", "sum"),
                  ("Σ Bereitschaften", "ST_BB", "sum"), ("Σ Urlaubstage", "ST_Urlaub", "sum"),
                  ("Σ Krankheitstage", "ST_Krank", "sum"), ("Σ Unterrichtstage", "ST_UR", "sum"),
                  ("Σ JAZ-Tage", "ST_JAZ", "sum")], "ST_Name", "ST_Team")
    col_widths(ws, {"A": 18, "B": 9, "C": 12, "S": 13})
    for c in range(4, 19):
        ws.column_dimensions[L(c)].width = 10
    ws.freeze_panes = f"B{P_FIRST}"
    landscape(ws, "A4", rows=f"{hr}:{hr}")
    ws.sheet_properties.tabColor = "00B050"


# ====================================================================== Dienstverteilung
def build_dienstverteilung(book, ws):
    title(ws, "Dienstverteilung und Nachtdienstverteilung je Mitarbeiter und Monat",
          "Nachtdienste = N + bestätigte Vertretungen (V). Tagdienste = T. Ziel: möglichst gleichmäßige Verteilung "
          "(Abweichungen entstehen nur durch Abwesenheiten und Vertretungen – die Rotation selbst ist gleich verteilt).")
    labels = ["Name", "Team"] + [f"Nacht {m}" for m in MON3] + ["Nacht Σ"] + [f"Tag {m}" for m in MON3] + ["Tag Σ"]
    hr = P_FIRST - 1
    header_row(ws, hr, 1, labels, height=30, size=8)
    put(ws, "C4", "Nachtdienstverteilung", size=10, bold=True, h="center", bg=HEAD, border=THIN)
    ws.merge_cells("C4:O4")
    put(ws, "P4", "Tagdienstverteilung", size=10, bold=True, h="center", bg=HEAD, border=THIN)
    ws.merge_cells("P4:AB4")
    for i in range(NMAX):
        r = P_FIRST + i
        pr = prow(i)
        person_cols(ws, i, r, "AB")
        e = f'IF($A{r}="","",'
        for m in range(12):
            ws[f"{L(3 + m)}{r}"] = f'={e}COUNTIFS({pr},"N",PEP_Mon,{m + 1})+COUNTIFS({pr},"V",PEP_Mon,{m + 1}))'
            ws[f"{L(16 + m)}{r}"] = f'={e}COUNTIFS({pr},"T",PEP_Mon,{m + 1}))'
        ws[f"O{r}"] = f'={e}SUM(C{r}:N{r}))'
        ws[f"AB{r}"] = f'={e}SUM(P{r}:AA{r}))'
        body(ws, r, 1, 28)
        ws[f"O{r}"].font = font(9, bold=True)
        ws[f"AB{r}"].font = font(9, bold=True)
    book.table(ws, "tbl_Dienst", f"A{hr}:AB{P_LAST}")
    f = P_FIRST
    team_cf(book, ws, f"B{f}:B{P_LAST}", f"B{f}")
    for rg in (f"O{f}:O{P_LAST}", f"AB{f}:AB{P_LAST}"):
        ws.conditional_formatting.add(rg, DataBarRule(start_type="num", start_value=0, end_type="max",
                                                      color=SERIES1, showValue=True))
    # Teamebene
    s = P_LAST + 3
    put(ws, f"A{s}", "Nachtdienst-Tage je Team und Monat (Rotation)", size=12, bold=True)
    header_row(ws, s + 1, 1, ["Team", "Mitglieder"] + MON3 + ["Σ", "Ø Nachtdienste je MA (Personen)"],
               height=30, size=8)
    for k, (t, kk, col) in enumerate(TEAMS):
        r = s + 2 + k
        put(ws, f"A{r}", t, bold=True, bg=col, border=THIN)
        put(ws, f"B{r}", f'=COUNTIFS(C_Team,"{t}",C_Aktiv,"<>NEIN")', h="center", border=THIN)
        for m in range(12):
            put(ws, f"{L(3 + m)}{r}", f'=COUNTIFS(Plan_Nacht,"{t}",Plan_Monat,{m + 1})', h="center", border=THIN)
        put(ws, f"O{r}", f"=SUM(C{r}:N{r})", h="center", border=THIN, bold=True)
        put(ws, f"P{r}", f'=IF(B{r}=0,0,SUMIFS($O${P_FIRST}:$O${P_LAST},$B${P_FIRST}:$B${P_LAST},"{t}")/B{r})',
            h="center", border=THIN, fmt="0.0")
    book.name_ws("DV_TeamNacht", ws, f"O{s + 2}:O{s + 6}")
    col_widths(ws, {"A": 18, "B": 9})
    for c in range(3, 29):
        ws.column_dimensions[L(c)].width = 7.5
    ws.column_dimensions["P"].width = 12
    ws.freeze_panes = f"C{P_FIRST}"
    landscape(ws, "A4", rows=f"{hr}:{hr}")
    ws.sheet_properties.tabColor = "00B050"


# ====================================================================== Urlaubsstatistik
def build_urlaub(book, ws):
    title(ws, "Urlaubsstatistik (Arbeitstage)",
          "Anspruch = Jahresurlaub (Mitarbeiterstamm) + Übertrag (Zeitbuchungen, Konto Urlaub). Genommen/geplant laut "
          "PEP 2027; Krankheit während des Urlaubs zählt nicht als Urlaub.")
    labels = ["Name", "Team", "Jahresurlaub", "Übertrag", "Anspruch gesamt"] + MON3 + \
             ["Urlaub gesamt", "genommen bis Stichtag", "geplant ab Stichtag", "Resturlaub", "Urlaubsquote", "Hinweis"]
    hr = P_FIRST - 1
    header_row(ws, hr, 1, labels, height=36, size=8)
    for i in range(NMAX):
        r = P_FIRST + i
        pr = prow(i)
        person_cols(ws, i, r, "AB")
        e = f'IF($A{r}="","",'
        ws[f"C{r}"] = f'={e}N(IFERROR(INDEX(MA_Urlaub,MATCH($A{r},MA_Name,0)),0)))'
        ws[f"D{r}"] = f'={e}SUMIFS(ZB_Wert,ZB_MA,$A{r},ZB_Konto,"Urlaub"))'
        ws[f"E{r}"] = f'={e}C{r}+D{r})'
        for m in range(12):
            ws[f"{L(6 + m)}{r}"] = f'={e}COUNTIFS({pr},"U",PEP_Arb,1,PEP_Mon,{m + 1}))'
        ws[f"R{r}"] = f'={e}SUM(F{r}:Q{r}))'
        ws[f"S{r}"] = f'={e}COUNTIFS({pr},"U",PEP_Arb,1,PEP_Datum,"<="&Stichtag))'
        ws[f"T{r}"] = f'={e}R{r}-S{r})'
        ws[f"U{r}"] = f'={e}E{r}-R{r})'
        ws[f"V{r}"] = f'={e}IFERROR(R{r}/(COUNTIF(PEP_Arb,1)-COUNTIFS({pr},"–",PEP_Arb,1)),0))'
        ws[f"W{r}"] = f'={e}IF(U{r}<0,"⚠ mehr Urlaub geplant als Anspruch",IF(U{r}>=10,"Resturlaub einplanen","")))'
        body(ws, r, 1, 23)
        ws[f"V{r}"].number_format = "0.0%"
        ws[f"U{r}"].font = font(9, bold=True)
    book.table(ws, "tbl_Urlaub", f"A{hr}:W{P_LAST}")
    f = P_FIRST
    rng = lambda c: f"{c}{f}:{c}{P_LAST}"
    book.name_ws("US_Mon", ws, f"F{f}:Q{P_LAST}")
    book.name_ws("US_Rest", ws, rng("U"))
    book.name_ws("US_Team", ws, rng("B"))
    book.name_ws("US_Name", ws, rng("A"))
    team_cf(book, ws, rng("B"), f"B{f}")
    book.cf(ws, rng("U"), f'AND(ISNUMBER(U{f}),U{f}<0)', bg="FFC7CE", color="9C0006")
    book.cf(ws, rng("W"), f'W{f}<>""', bg="FFEB9C", color="9C5700")
    book.cf(ws, f"F{f}:Q{P_LAST}", f'AND(ISNUMBER(F{f}),F{f}>0)', bg="C6EFCE")
    s = team_summary(book, ws, P_LAST + 3, "Urlaub je Team",
                     [("Anspruch", rng("E"), "sum"), ("Urlaub gesamt", rng("R"), "sum"),
                      ("genommen bis Stichtag", rng("S"), "sum"), ("geplant", rng("T"), "sum"),
                      ("Resturlaub", rng("U"), "sum"), ("Ø Resturlaub je MA", rng("U"), "avg")], "US_Name", "US_Team")
    m0 = s + 3
    put(ws, f"A{m0}", "Urlaub je Monat (alle Mitarbeiter)", size=12, bold=True)
    header_row(ws, m0 + 1, 1, ["Kennzahl", ""] + MON3 + ["Jahr"], height=20, size=8)
    put(ws, f"A{m0 + 2}", "Urlaubstage", bold=True, border=THIN)
    put(ws, f"A{m0 + 3}", "Urlaubsquote", bold=True, border=THIN)
    for m in range(12):
        c = L(6 + m)
        put(ws, f"{L(3 + m)}{m0 + 2}", f"=SUM({c}{f}:{c}{P_LAST})", h="center", border=THIN)
        put(ws, f"{L(3 + m)}{m0 + 3}", f"=IFERROR({L(3 + m)}{m0 + 2}/({calc_mon_sum(m + 1, 0)}/SollStdTag),0)",
            h="center", border=THIN, fmt="0.0%")
    put(ws, f"O{m0 + 2}", f"=SUM(C{m0 + 2}:N{m0 + 2})", h="center", border=THIN, bold=True)
    put(ws, f"O{m0 + 3}", f"=IFERROR(O{m0 + 2}/(SUM(UE_Soll)/SollStdTag),0)", h="center", border=THIN, bold=True,
        fmt="0.0%")
    col_widths(ws, {"A": 18, "B": 9, "W": 30})
    for c in range(3, 23):
        ws.column_dimensions[L(c)].width = 8.5
    ws.freeze_panes = f"C{P_FIRST}"
    landscape(ws, "A4", rows=f"{hr}:{hr}")
    ws.sheet_properties.tabColor = "00B050"


# ====================================================================== Krankenstatistik
def build_krank(book, ws):
    title(ws, "Krankenstatistik",
          "Krankheitstage je Monat in Arbeitstagen, zusätzlich Kalendertage und Anzahl Krankmeldungen. "
          "Krankenquote = Krankheitstage / Soll-Arbeitstage. Hinweis BEM ab 42 Kalendertagen.")
    labels = ["Name", "Team"] + MON3 + ["Krank Σ (Arbeitstage)", "Kalendertage krank", "Krankmeldungen",
                                        "Krankenquote", "Hinweis"]
    hr = P_FIRST - 1
    header_row(ws, hr, 1, labels, height=36, size=8)
    for i in range(NMAX):
        r = P_FIRST + i
        pr = prow(i)
        person_cols(ws, i, r, "AB")
        e = f'IF($A{r}="","",'
        for m in range(12):
            ws[f"{L(3 + m)}{r}"] = f'={e}COUNTIFS({pr},"K",PEP_Arb,1,PEP_Mon,{m + 1}))'
        ws[f"O{r}"] = f'={e}SUM(C{r}:N{r}))'
        ws[f"P{r}"] = f'={e}COUNTIF({pr},"K"))'
        ws[f"Q{r}"] = f'={e}COUNTIFS(Abw_MA,$A{r},Abw_Art,"Krankheit"))'
        ws[f"R{r}"] = f'={e}IFERROR(O{r}/(COUNTIF(PEP_Arb,1)-COUNTIFS({pr},"–",PEP_Arb,1)),0))'
        ws[f"S{r}"] = f'={e}IF(P{r}>=42,"BEM prüfen (≥ 6 Wochen)",""))'
        body(ws, r, 1, 19)
        ws[f"R{r}"].number_format = "0.0%"
    book.table(ws, "tbl_Krank", f"A{hr}:S{P_LAST}")
    f = P_FIRST
    rng = lambda c: f"{c}{f}:{c}{P_LAST}"
    book.name_ws("KS_Mon", ws, f"C{f}:N{P_LAST}")
    book.name_ws("KS_Team", ws, rng("B"))
    book.name_ws("KS_Name", ws, rng("A"))
    team_cf(book, ws, rng("B"), f"B{f}")
    book.cf(ws, f"C{f}:N{P_LAST}", f'AND(ISNUMBER(C{f}),C{f}>0)', bg="FFEB9C")
    book.cf(ws, rng("S"), f'S{f}<>""', bg="FFC7CE", color="9C0006")
    s = team_summary(book, ws, P_LAST + 3, "Krankheit je Team",
                     [("Krankheitstage (AT)", rng("O"), "sum"), ("Kalendertage", rng("P"), "sum"),
                      ("Krankmeldungen", rng("Q"), "sum"), ("Ø Krankheitstage je MA", rng("O"), "avg")],
                     "KS_Name", "KS_Team")
    m0 = s + 3
    put(ws, f"A{m0}", "Krankheit je Monat (alle Mitarbeiter)", size=12, bold=True)
    header_row(ws, m0 + 1, 1, ["Kennzahl", ""] + MON3 + ["Jahr"], height=20, size=8)
    put(ws, f"A{m0 + 2}", "Krankheitstage", bold=True, border=THIN)
    put(ws, f"A{m0 + 3}", "Krankenquote", bold=True, border=THIN)
    for m in range(12):
        c = L(3 + m)
        put(ws, f"{c}{m0 + 2}", f"=SUM({c}{f}:{c}{P_LAST})", h="center", border=THIN)
        put(ws, f"{c}{m0 + 3}", f"=IFERROR({c}{m0 + 2}/({calc_mon_sum(m + 1, 0)}/SollStdTag),0)", h="center",
            border=THIN, fmt="0.0%")
    put(ws, f"O{m0 + 2}", f"=SUM(C{m0 + 2}:N{m0 + 2})", h="center", border=THIN, bold=True)
    put(ws, f"O{m0 + 3}", f"=IFERROR(O{m0 + 2}/(SUM(UE_Soll)/SollStdTag),0)", h="center", border=THIN, bold=True,
        fmt="0.0%")
    col_widths(ws, {"A": 18, "B": 9, "S": 24})
    for c in range(3, 19):
        ws.column_dimensions[L(c)].width = 8.5
    ws.freeze_panes = f"C{P_FIRST}"
    landscape(ws, "A4", rows=f"{hr}:{hr}")
    ws.sheet_properties.tabColor = "00B050"


# ====================================================================== Monatsbericht
def kpi_tiles(ws, row, items, c0=2, width=2, label_size=8, value_size=16):
    """items: [(Beschriftung, Formel, Zahlenformat)] -> Kacheln in row (Label) / row+1 (Wert)."""
    for k, (lab, f, fmt) in enumerate(items):
        c = c0 + k * width
        a, b = L(c), L(c + width - 1)
        put(ws, f"{a}{row}", lab, size=label_size, color="595959", h="center", wrap=True, bg="F2F2F2")
        ws.merge_cells(f"{a}{row}:{b}{row}")
        put(ws, f"{a}{row + 1}", f, size=value_size, bold=True, h="center", fmt=fmt, bg="F2F2F2")
        ws.merge_cells(f"{a}{row + 1}:{b}{row + 1}")
    ws.row_dimensions[row].height = 26
    ws.row_dimensions[row + 1].height = 30


def build_monatsbericht(book, ws):
    put(ws, "A1", '="Monatsbericht "&INDEX(MonName,$C$3)&" "&Jahr', size=16, bold=True)
    ws.row_dimensions[1].height = 26
    put(ws, "A2", "Monat im gelben Feld wählen. Alle Werte automatisch.", size=9, color=MUTED, italic=True)
    put(ws, "B3", "Monat:", bold=True, h="right")
    put(ws, "C3", "=MONTH(Stichtag)", bold=True, bg=INPUT_BG, h="center", border=THIN)
    book.dv_list(ws, "C3", "=MonNr")
    put(ws, "D3", "=INDEX(MonName,$C$3)", bold=True)
    m = "$C$3"
    soll_m = f"SUM(INDEX(C_MonAll,0,3*{m}-2))"
    ist_m = f"SUM(INDEX(C_MonAll,0,3*{m}-1))"
    kpi_tiles(ws, 5, [
        ("Arbeitstage", f"=COUNTIFS(Plan_Arb,1,Plan_Monat,{m})", "0"),
        ("Ø Tagdienst anwesend", f"=IFERROR(AVERAGEIFS(PB_TDIst,PB_Monat,{m},PB_Arb,1),0)", "0.0"),
        ("Urlaubstage", f"=SUM(INDEX(US_Mon,0,{m}))", "0"),
        ("Krankheitstage", f"=SUM(INDEX(KS_Mon,0,{m}))", "0"),
        ("Urlaubsquote", f"=IFERROR(F6/({soll_m}/SollStdTag),0)", "0.0%"),
        ("Krankenquote", f"=IFERROR(H6/({soll_m}/SollStdTag),0)", "0.0%"),
        ("Überstunden Monat (Σ Std.)", f"=SUM(INDEX(UE_Monate,0,{m}))", "+0.0;-0.0;0.0"),
        ("Auslastung (Ist/Soll)", f"=IFERROR({ist_m}/{soll_m},0)", "0%"),
    ])
    kpi_tiles(ws, 8, [
        ("Nachtdienst einsatzfähig (Tage)", f'=COUNTIFS(EF_Ampel,"Grün",Plan_Monat,{m})', "0"),
        ("Lücke mit Ersatz (Gelb)", f'=COUNTIFS(EF_Ampel,"Gelb",Plan_Monat,{m})', "0"),
        ("Lücke ohne Ersatz (Rot)", f'=COUNTIFS(EF_Ampel,"Rot",Plan_Monat,{m})', "0"),
        ("Unterbesetzung Tagdienst (Tage)", f'=COUNTIFS(PB_Ampel,"Rot",PB_Monat,{m})', "0"),
        ("Unterrichtstage", f"=SUM(J22:J{21 + NMAX})", "0"),
        ("JAZ-Tage", f"=SUM(K22:K{21 + NMAX})", "0"),
        ("Bestätigte Vertretungen (Tage)", f"=SUM(L22:L{21 + NMAX})", "0"),
        ("Ampel Monat", f"=INDEX(PB_MonAmpel,{m})", None),
    ])
    ampel_cf(book, ws, "P9", "P9")
    # Teamtabelle
    put(ws, "A11", "Teams im Monat", size=12, bold=True)
    tl = ["Team", "MA aktiv", "Nachtdiensttage", "davon einsatzfähig", "Lücke mit Ersatz", "Lücke ohne Ersatz",
          "Tage Team einsatzfähig", "Urlaubstage", "Krankheitstage", "Überstunden (Σ Std.)"]
    header_row(ws, 12, 1, tl, height=30, size=8)
    for k, (t, kk, col) in enumerate(TEAMS):
        r = 13 + k
        put(ws, f"A{r}", t, bold=True, bg=col, border=THIN)
        vals = [f'=COUNTIFS(C_Team,"{t}",C_Aktiv,"<>NEIN")', f'=COUNTIFS(Plan_Nacht,"{t}",Plan_Monat,{m})',
                f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Grün",Plan_Monat,{m})',
                f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Gelb",Plan_Monat,{m})',
                f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Rot",Plan_Monat,{m})',
                f'=COUNTIFS(EF_St{k + 1},"✅*",Plan_Monat,{m})',
                f'=SUMIFS($H$22:$H${21 + NMAX},$B$22:$B${21 + NMAX},"{t}")',
                f'=SUMIFS($I$22:$I${21 + NMAX},$B$22:$B${21 + NMAX},"{t}")',
                f'=SUMIFS($Q$22:$Q${21 + NMAX},$B$22:$B${21 + NMAX},"{t}")']
        for j, v in enumerate(vals):
            put(ws, f"{L(2 + j)}{r}", v, h="center", border=THIN, fmt="0.0" if j == 8 else None)
    # Mitarbeitertabelle
    put(ws, "A20", "Mitarbeiter im Monat", size=12, bold=True)
    ml = ["Name", "Team", "Fkt.", "Tagdienste", "Nachtdienste", "davon Wochenende", "davon Feiertag", "Urlaub",
          "Krank", "Unterricht", "JAZ", "Vertretung (V)", "Sonstige Abw.", "Freistellung", "Soll-Std.", "Ist-Std.",
          "Monatswert"]
    header_row(ws, 21, 1, ml, height=30, size=8)
    for i in range(NMAX):
        r = 22 + i
        pr = prow(i)
        cr = C_FIRST + i
        person_cols(ws, i, r)
        e = f'IF($A{r}="","",'
        cnt = lambda code, extra="": f'COUNTIFS({pr},"{code}",PEP_Mon,{m}{extra})'
        ws[f"D{r}"] = f"={e}{cnt('T')})"
        ws[f"E{r}"] = f"={e}{cnt('N')}+{cnt('V')})"
        ws[f"F{r}"] = f"={e}{cnt('N', ',PEP_FrSo,1')}+{cnt('V', ',PEP_FrSo,1')})"
        ws[f"G{r}"] = f"={e}{cnt('N', ',PEP_FT,1')}+{cnt('V', ',PEP_FT,1')})"
        ws[f"H{r}"] = f"={e}{cnt('U', ',PEP_Arb,1')})"
        ws[f"I{r}"] = f"={e}{cnt('K', ',PEP_Arb,1')})"
        ws[f"J{r}"] = f"={e}{cnt('UR', ',PEP_Arb,1')}+{cnt('FDL', ',PEP_Arb,1')}+{cnt('KAB', ',PEP_Arb,1')})"
        ws[f"K{r}"] = f"={e}{cnt('JAZ', ',PEP_Arb,1')})"
        ws[f"L{r}"] = f"={e}{cnt('V')})"
        ws[f"M{r}"] = f"={e}{cnt('FB', ',PEP_Arb,1')}+{cnt('SU', ',PEP_Arb,1')}+{cnt('FS', ',PEP_Arb,1')})"
        ws[f"N{r}"] = f"={e}{cnt('F')})"
        ws[f"O{r}"] = f"={e}INDEX(Hilfsberechnung!${L(C_MON0)}{cr}:${L(C_MON0 + 35)}{cr},3*{m}-2))"
        ws[f"P{r}"] = f"={e}INDEX(Hilfsberechnung!${L(C_MON0)}{cr}:${L(C_MON0 + 35)}{cr},3*{m}-1))"
        ws[f"Q{r}"] = f"={e}P{r}-O{r})"
        body(ws, r, 1, 17)
        for c in "OPQ":
            ws[f"{c}{r}"].number_format = "0.0"
    team_cf(book, ws, f"B22:B{21 + NMAX}", "B22")
    book.cf(ws, f"Q22:Q{21 + NMAX}", 'AND(ISNUMBER(Q22),Q22<0)', color="C00000", bold=True)
    book.cf(ws, f"Q22:Q{21 + NMAX}", 'AND(ISNUMBER(Q22),Q22>0)', color="006100", bold=True)
    col_widths(ws, {"A": 18})
    for c in range(2, 18):
        ws.column_dimensions[L(c)].width = 9.5
    landscape(ws, "A4", area=f"A1:Q{21 + NMAX}", rows="21:21")
    ws.sheet_properties.tabColor = "00B050"


# ====================================================================== Jahresbericht
def build_jahresbericht(book, ws):
    put(ws, "A1", '="Jahresbericht "&Jahr', size=16, bold=True)
    ws.row_dimensions[1].height = 26
    put(ws, "A2", "Monatsübersicht und Teams für das gesamte Planungsjahr – vollständig automatisch.",
        size=9, color=MUTED, italic=True)
    lab = ["Monat", "Arbeitstage", "Ø Tagdienst anwesend", "Urlaubstage", "Krankheitstage", "Urlaubsquote",
           "Krankenquote", "Nacht Grün", "Nacht Gelb", "Nacht Rot", "Unterbesetzung TD (Tage)", "Soll-Std.",
           "Ist-Std.", "Saldo (Std.)", "Auslastung", "Ampel"]
    header_row(ws, 4, 1, lab, height=36, size=8)
    for m in range(13):
        r = 5 + m
        yr = m == 12
        put(ws, f"A{r}", "Jahr" if yr else MONATE[m], bold=True, border=THIN, bg=GREY if yr else None)
        if yr:
            vals = ["=SUM(B5:B16)", "=IFERROR(AVERAGEIFS(PB_TDIst,PB_Arb,1),0)", "=SUM(D5:D16)", "=SUM(E5:E16)",
                    "=IFERROR(D17/(L17/SollStdTag),0)", "=IFERROR(E17/(L17/SollStdTag),0)", "=SUM(H5:H16)",
                    "=SUM(I5:I16)", "=SUM(J5:J16)", "=SUM(K5:K16)", "=SUM(L5:L16)", "=SUM(M5:M16)", "=M17-L17",
                    "=IFERROR(M17/L17,0)", "=PB_JahrAmpel"]
        else:
            mm = m + 1
            vals = [f"=COUNTIFS(Plan_Arb,1,Plan_Monat,{mm})",
                    f"=IFERROR(AVERAGEIFS(PB_TDIst,PB_Monat,{mm},PB_Arb,1),0)",
                    f"=SUM(INDEX(US_Mon,0,{mm}))", f"=SUM(INDEX(KS_Mon,0,{mm}))",
                    f"=IFERROR(D{r}/(L{r}/SollStdTag),0)", f"=IFERROR(E{r}/(L{r}/SollStdTag),0)",
                    f"=INDEX(PB_MonNGruen,{mm})", f"=INDEX(PB_MonNGelb,{mm})", f"=INDEX(PB_MonNRot,{mm})",
                    f'=COUNTIFS(PB_Ampel,"Rot",PB_Monat,{mm})', f"={calc_mon_sum(mm, 0)}", f"={calc_mon_sum(mm, 1)}",
                    f"=M{r}-L{r}", f"=IFERROR(M{r}/L{r},0)", f"=INDEX(PB_MonAmpel,{mm})"]
        for j, v in enumerate(vals):
            fmt = {4: "0.0%", 5: "0.0%", 1: "0.0", 10: "0", 11: "0", 12: "+0.0;-0.0;0.0", 13: "0%"}.get(j)
            put(ws, f"{L(2 + j)}{r}", v, h="center", border=THIN, fmt=fmt, bold=yr)
    ampel_cf(book, ws, "P5:P17", "P5")
    book.name_ws("JB_Urlaub", ws, "D5:D16")
    book.name_ws("JB_Krank", ws, "E5:E16")
    book.name_ws("JB_UQ", ws, "F5:F16")
    book.name_ws("JB_KQ", ws, "G5:G16")
    book.name_ws("JB_Ausl", ws, "O5:O16")
    # Teams
    put(ws, "A20", "Teams im Jahr", size=12, bold=True)
    tl = ["Team", "MA aktiv", "Nachtdiensttage", "davon Grün", "Gelb (Ersatz)", "Rot (ohne Ersatz)",
          "Quote einsatzfähig", "Urlaubstage", "Krankheitstage", "Saldo Überstunden", "JAZ-Konto",
          "Personalbedarf (strukturell)", "Ampel Struktur"]
    header_row(ws, 21, 1, tl, height=36, size=8)
    for k, (t, kk, col) in enumerate(TEAMS):
        r = 22 + k
        put(ws, f"A{r}", t, bold=True, bg=col, border=THIN)
        vals = [f'=COUNTIFS(C_Team,"{t}",C_Aktiv,"<>NEIN")', f'=COUNTIF(Plan_Nacht,"{t}")',
                f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Grün")', f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Gelb")',
                f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Rot")', f"=IFERROR(D{r}/C{r},0)",
                f'=SUMIFS(ST_Urlaub,ST_Team,"{t}")', f'=SUMIFS(ST_Krank,ST_Team,"{t}")',
                f'=SUMIFS(UE_SaldoUe,UE_Team,"{t}")', f'=SUMIFS(UE_SaldoJAZ,UE_Team,"{t}")',
                f"=INDEX(PB_StrBedarf,{k + 1})", f"=INDEX(PB_StrAmpel,{k + 1})"]
        for j, v in enumerate(vals):
            put(ws, f"{L(2 + j)}{r}", v, h="center", border=THIN,
                fmt="0%" if j == 5 else ("0.0" if j in (8, 9) else None))
    ampel_cf(book, ws, "M22:M26", "M22")
    book.name_ws("JB_TeamQuote", ws, "G22:G26")
    put(ws, "A28", "Kernaussagen", size=12, bold=True)
    notes = [
        '="Nachtdienste: "&COUNTIF(EF_Ampel,"Grün")&" Tage voll einsatzfähig, "&COUNTIF(EF_Ampel,"Gelb")&'
        '" Tage mit Lücke (Ersatz vorhanden), "&COUNTIF(EF_Ampel,"Rot")&" Tage mit Lücke ohne Ersatz."',
        '="Strukturell fehlende Stammbesetzung: "&PB_Bedarf&" Funktion(en) – Details im Blatt Personalbedarfsplanung."',
        '="Urlaubsquote "&FIXED(F17*100,1)&" % · Krankenquote "&FIXED(G17*100,1)&" % · Auslastung "&FIXED(O17*100,0)&" %."',
        '="Qualifikationen: "&COUNTIF(Q_StatusNr,4)&" abgelaufen, "&COUNTIF(Q_StatusNr,3)&" laufen in < "&QualWarnTage&'
        '" Tagen ab, "&COUNTIF(Q_StatusNr,5)&" Nachweise fehlen."',
    ]
    for k, n in enumerate(notes):
        put(ws, f"A{29 + k}", n, size=10)
        ws.merge_cells(f"A{29 + k}:P{29 + k}")
    col_widths(ws, {"A": 16})
    for c in range(2, 17):
        ws.column_dimensions[L(c)].width = 11
    landscape(ws, "A4", area="A1:P33")
    ws.sheet_properties.tabColor = "00B050"


# ====================================================================== Einsatzfaehigkeitsbericht
def build_efbericht(book, ws):
    put(ws, "A1", '="Einsatzfähigkeitsbericht "&Jahr', size=16, bold=True)
    ws.row_dimensions[1].height = 26
    put(ws, "A2", "Bezogen auf die eigenen Nachtdienst-/Bereitschaftstage je Team. Grün = voll einsatzfähig, "
                  "Gelb = Lücke mit Ersatzvorschlag, Rot = Lücke ohne Ersatz.", size=9, color=MUTED, italic=True)
    put(ws, "A4", "Nachtdiensttage voll einsatzfähig je Team und Monat", size=12, bold=True)
    header_row(ws, 5, 1, ["Team"] + MON3 + ["Jahr", "von Nachtdiensttagen", "Quote"], height=20, size=8)
    for k, (t, kk, col) in enumerate(TEAMS):
        r = 6 + k
        put(ws, f"A{r}", t, bold=True, bg=col, border=THIN)
        for m in range(12):
            put(ws, f"{L(2 + m)}{r}", f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Grün",Plan_Monat,{m + 1})&"/"&'
                                      f'COUNTIFS(Plan_Nacht,"{t}",Plan_Monat,{m + 1})', h="center", border=THIN)
        put(ws, f"N{r}", f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Grün")', h="center", border=THIN, bold=True)
        put(ws, f"O{r}", f'=COUNTIF(Plan_Nacht,"{t}")', h="center", border=THIN)
        put(ws, f"P{r}", f"=IFERROR(N{r}/O{r},0)", h="center", border=THIN, fmt="0%", bold=True)
    book.cf(ws, "P6:P10", "P6>=0.95", bg="C6EFCE", color="006100")
    book.cf(ws, "P6:P10", "AND(P6<0.95,P6>=0.8)", bg="FFEB9C", color="9C5700")
    book.cf(ws, "P6:P10", "P6<0.8", bg="FFC7CE", color="9C0006")
    put(ws, "A13", "Ursachen der Lücken während der eigenen Nachtdienste (Anzahl Tage)", size=12, bold=True)
    reasons = [("Meister fehlt", "*Meister fehlt*"), ("AV fehlt", "*AV fehlt*"), ("TVT fehlt", "*TVT fehlt*"),
               ("Zu wenig Handwerker", "*Zu wenig Handwerker*"), ("Qualifikation abgelaufen", "*Qualifikation abgelaufen*")]
    header_row(ws, 14, 1, ["Team"] + [x[0] for x in reasons] + ["Tage Gelb", "Tage Rot",
                                                                "Tage Team einsatzfähig (alle Tage)", "Quote alle Tage"],
               height=36, size=8)
    for k, (t, kk, col) in enumerate(TEAMS):
        r = 15 + k
        put(ws, f"A{r}", t, bold=True, bg=col, border=THIN)
        for j, (lab, pat) in enumerate(reasons):
            put(ws, f"{L(2 + j)}{r}", f'=COUNTIFS(EF_Nacht,"{t}",EF_Fehlend,"{pat}")', h="center", border=THIN)
        put(ws, f"G{r}", f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Gelb")', h="center", border=THIN)
        put(ws, f"H{r}", f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Rot")', h="center", border=THIN)
        put(ws, f"I{r}", f'=COUNTIF(EF_St{k + 1},"✅*")', h="center", border=THIN)
        put(ws, f"J{r}", f"=IFERROR(I{r}/TageJahr,0)", h="center", border=THIN, fmt="0%")
    book.cf(ws, "H15:H19", "H15>0", bg="FFC7CE", color="9C0006", bold=True)
    book.cf(ws, "B15:F19", "B15>0", bg="FFEB9C")
    # Naechste kritische Nachtdienste
    put(ws, "A22", "Nächste Nachtdienste mit Lücke ab Stichtag (Gelb = Ersatz vorhanden, Rot = ohne Ersatz)",
        size=12, bold=True)
    header_row(ws, 23, 1, ["Nr", "Datum", "Wt", "Nachtteam", "Ampel", "Fehlende Funktionen", "", "", "",
                           "Ersatzvorschläge", "", "", "", ""], height=20, size=8)
    ws.merge_cells("F23:I23")
    ws.merge_cells("J23:N23")
    crit = '((EF_Ampel="Rot")+(EF_Ampel="Gelb"))*(Plan_N>=Stichtag-Jahresbeginn+1)'
    for k in range(20):
        r = 24 + k
        put(ws, f"R{r}", f"=SUMPRODUCT(SMALL({crit}*Plan_N+(1-{crit})*1E+9,{k + 1}))", size=7, color=MUTED)
        put(ws, f"A{r}", k + 1, h="center", border=THIN, size=9)
        put(ws, f"B{r}", f'=IF(R{r}>=1E+9,"",INDEX(Plan_Datum,R{r}))', h="center", border=THIN, fmt=DATE_FMT, size=9)
        put(ws, f"C{r}", f'=IF(R{r}>=1E+9,"",INDEX(Plan_WT,R{r}))', h="center", border=THIN, size=9)
        put(ws, f"D{r}", f'=IF(R{r}>=1E+9,"",INDEX(Plan_Nacht,R{r}))', h="center", border=THIN, size=9)
        put(ws, f"E{r}", f'=IF(R{r}>=1E+9,"",INDEX(EF_Ampel,R{r}))', h="center", border=THIN, size=9, bold=True)
        put(ws, f"F{r}", f'=IF(R{r}>=1E+9,"",INDEX(EF_Fehlend,R{r}))', border=THIN, size=8)
        ws.merge_cells(f"F{r}:I{r}")
        put(ws, f"J{r}", f'=IF(R{r}>=1E+9,"",INDEX(EF_ErsatzText,R{r}))', border=THIN, size=8)
        ws.merge_cells(f"J{r}:N{r}")
    ampel_cf(book, ws, "E24:E43", "E24")
    team_cf(book, ws, "D24:D43", "D24")
    ws.column_dimensions["R"].hidden = True
    col_widths(ws, {"A": 14})
    for c in "BCDEFGHIJKLMNOP":
        ws.column_dimensions[c].width = 10
    landscape(ws, "A4", area="A1:P43")
    ws.sheet_properties.tabColor = "00B050"


# ====================================================================== Dashboard
def _bar(title_text, y_title=None, stacked=False, horizontal=False):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = title_text
    ch.style = 10
    ch.height, ch.width = 7.2, 15.5
    ch.gapWidth = 60
    if stacked:
        ch.grouping = "stacked"
        ch.overlap = 100
    ch.y_axis.majorGridlines.spPr = None
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    if y_title:
        ch.y_axis.title = y_title
    ch.legend.position = "b"
    return ch


def _labels(fmt):
    d = DataLabelList()
    d.showVal = True
    d.showSerName = False
    d.showCatName = False
    d.showLegendKey = False
    d.showPercent = False
    d.numFmt = fmt
    return d


def _color_series(s, hex_):
    s.graphicalProperties.solidFill = hex_
    s.graphicalProperties.line.solidFill = "FFFFFF"


def build_dashboard(book, ws):
    ws.sheet_view.showGridLines = False
    put(ws, "B1", '="Management-Dashboard Personal- und Einsatzplanung "&Jahr', size=18, bold=True)
    ws.row_dimensions[1].height = 30
    put(ws, "B2", '="Stichtag "&TEXT(DAY(Stichtag),"00")&"."&TEXT(MONTH(Stichtag),"00")&"."&YEAR(Stichtag)&" · Nachtteam heute: "&INDEX(Plan_Nacht,Stichtag-Jahresbeginn+1)'
                  '&" ("&INDEX(EF_StatusNacht,Stichtag-Jahresbeginn+1)&")"', size=10, color="595959")
    put(ws, "M2", "Berichtsmonat:", bold=True, h="right")
    put(ws, "N2", "=MONTH(Stichtag)", bold=True, bg=INPUT_BG, h="center", border=THIN)
    put(ws, "O2", "=INDEX(MonName,$N$2)", bold=True)
    book.dv_list(ws, "N2", "=MonNr")
    book.name_ws("Berichtsmonat", ws, "N2")
    n0 = "(Stichtag-Jahresbeginn+1)"
    m = "Berichtsmonat"

    def section(r, text):
        put(ws, f"B{r}", text, size=11, bold=True, color="FFFFFF", bg=HEAD_DARK)
        for c in range(3, 16):
            ws.cell(r, c).fill = fill(HEAD_DARK)

    section(4, "Heute (Stichtag)")
    kpi_tiles(ws, 5, [
        ("Personalbestand (aktiv)", '=COUNTIFS(C_Name,"?*",C_Aktiv,"<>NEIN")', "0"),
        ("Verfügbare Mitarbeiter", f"=COUNTIF(INDEX(C_Grid,0,{n0}),0)", "0"),
        ("Urlauber", f'=COUNTIF(INDEX(PEP_Grid,0,{n0}),"U")', "0"),
        ("Kranke Mitarbeiter", f'=COUNTIF(INDEX(PEP_Grid,0,{n0}),"K")', "0"),
        ("Einsatzfähige Teams", '=COUNTIF($E$19:$E$23,"✅*")&" / 5"', None),
        ("Kritische Teams", '=COUNTIF($E$19:$E$23,"❌*")', "0"),
        ("Personalbedarf (fehlende Stammfunktionen)", "=PB_Bedarf", "0"),
    ])
    section(8, "Zeitkonten, Quoten und Auslastung")
    soll_m = f"SUM(INDEX(C_MonAll,0,3*{m}-2))"
    ist_m = f"SUM(INDEX(C_MonAll,0,3*{m}-1))"
    kpi_tiles(ws, 9, [
        ("Überstunden-Saldo (Σ Std.)", "=SUM(UE_SaldoUe)", "+0.0;-0.0;0.0"),
        ("JAZ-Konto (Σ Std.)", "=SUM(UE_SaldoJAZ)", "0.0"),
        ("Urlaubsquote Monat", f"=INDEX(JB_UQ,{m})", "0.0%"),
        ("Krankenquote Monat", f"=INDEX(JB_KQ,{m})", "0.0%"),
        ("Monatsauslastung (Ist/Soll)", f"=IFERROR({ist_m}/{soll_m},0)", "0%"),
        ("Jahresauslastung (Ist/Soll)", "=IFERROR(SUM(UE_Ist)/SUM(UE_Soll),0)", "0%"),
        ("Urlaubs- / Krankenquote Jahr", "=FIXED(Jahresbericht!F17*100,1)&\" % / \"&FIXED(Jahresbericht!G17*100,1)&\" %\"",
         None),
    ])
    section(12, "Qualifikationsstatus und Einsatzfähigkeit im Jahr")
    kpi_tiles(ws, 13, [
        ("Qualifikation gültig (> 180 T.)", "=COUNTIF(Q_StatusNr,1)", "0"),
        ("Vorwarnung (90–180 T.)", "=COUNTIF(Q_StatusNr,2)", "0"),
        ("Kritisch (< 90 T.)", "=COUNTIF(Q_StatusNr,3)", "0"),
        ("Abgelaufen", "=COUNTIF(Q_StatusNr,4)", "0"),
        ("Nachweis fehlt", "=COUNTIF(Q_StatusNr,5)", "0"),
        ("Nachtdienste mit Lücke ohne Ersatz (Tage)", '=COUNTIF(EF_Ampel,"Rot")', "0"),
        ("Ampel Jahr", "=PB_JahrAmpel", None),
    ])
    for c, (bg, fc) in {"D14": ("FFEB9C", "9C5700"), "F14": ("FFC7CE", "9C0006"), "H14": ("C00000", "FFFFFF"),
                        "J14": ("D9D9D9", "000000")}.items():
        book.cf(ws, c, f"{c}>0", bg=bg, color=fc)
    ampel_cf(book, ws, "N14", "N14")
    book.cf(ws, "L14", "L14>0", bg="FFC7CE", color="9C0006")
    book.cf(ws, "L6", "L6>0", bg="FFC7CE", color="9C0006")
    book.cf(ws, "N6", "N6>0", bg="FFEB9C", color="9C5700")
    # Teamtabelle
    section(16, "Teams")
    tl = ["Team", "Mitglieder", "verfügbar heute", "Status heute", "Fehlend heute", "Nachtdiensttage Jahr",
          "davon einsatzfähig", "Quote", "Urlaubstage", "Krankheitstage", "Überstunden-Saldo", "Struktur-Ampel"]
    header_row(ws, 18, 2, tl, height=30, size=8)
    for k, (t, kk, col) in enumerate(TEAMS):
        r = 19 + k
        put(ws, f"B{r}", t, bold=True, bg=col, border=THIN)
        vals = [f'=COUNTIFS(C_Team,"{t}",C_Aktiv,"<>NEIN")',
                f'=COUNTIFS(C_Team,"{t}",INDEX(C_Grid,0,{n0}),0)',
                f"=INDEX(EF_St{k + 1},{n0})", f"=INDEX(EF_Fe{k + 1},{n0})",
                f'=COUNTIF(Plan_Nacht,"{t}")', f'=COUNTIFS(EF_Nacht,"{t}",EF_Ampel,"Grün")',
                f"=IFERROR(H{r}/G{r},0)", f'=SUMIFS(ST_Urlaub,ST_Team,"{t}")', f'=SUMIFS(ST_Krank,ST_Team,"{t}")',
                f'=SUMIFS(UE_SaldoUe,UE_Team,"{t}")', f"=INDEX(PB_StrAmpel,{k + 1})"]
        for j, v in enumerate(vals):
            put(ws, f"{L(3 + j)}{r}", v, h="center" if j != 3 else None, border=THIN, size=9,
                fmt="0%" if j == 6 else ("+0.0;-0.0;0.0" if j == 9 else None), wrap=j == 3)
        ws.row_dimensions[r].height = 24
    book.cf(ws, "E19:E23", 'LEFT(E19,1)="✅"', bg="C6EFCE", color="006100", bold=True)
    book.cf(ws, "E19:E23", 'LEFT(E19,1)="❌"', bg="FFC7CE", color="9C0006", bold=True)
    ampel_cf(book, ws, "M19:M23", "M19")
    # Monatsdaten fuer Diagramme
    section(25, "Monatswerte (Datenbasis der Diagramme)")
    ml = ["Monat", "Urlaubstage", "Krankheitstage", "Nacht Grün", "Nacht Gelb", "Nacht Rot", "Auslastung",
          "Urlaubsquote", "Krankenquote", "Unterbesetzung TD (Tage)"]
    header_row(ws, 26, 2, ml, height=30, size=8)
    for mm in range(12):
        r = 27 + mm
        put(ws, f"B{r}", MON3[mm], bold=True, border=THIN)
        vals = [f"=INDEX(JB_Urlaub,{mm + 1})", f"=INDEX(JB_Krank,{mm + 1})", f"=INDEX(PB_MonNGruen,{mm + 1})",
                f"=INDEX(PB_MonNGelb,{mm + 1})", f"=INDEX(PB_MonNRot,{mm + 1})", f"=INDEX(JB_Ausl,{mm + 1})",
                f"=INDEX(JB_UQ,{mm + 1})", f"=INDEX(JB_KQ,{mm + 1})", f'=COUNTIFS(PB_Ampel,"Rot",PB_Monat,{mm + 1})']
        for j, v in enumerate(vals):
            put(ws, f"{L(3 + j)}{r}", v, h="center", border=THIN, size=9,
                fmt="0%" if j == 5 else ("0.0%" if j in (6, 7) else None))
    book.cf(ws, "B27:K38", f"ROW(B27)-26={m}", bg="DDEBF7")
    # Qualifikations- und Teamdaten fuer Diagramme
    put(ws, "M26", "Status", size=8, bold=True, bg=HEAD, border=THIN)
    put(ws, "N26", "Anzahl", size=8, bold=True, bg=HEAD, border=THIN)
    qlab = ["gültig", "Vorwarnung", "< 90 Tage", "abgelaufen", "Nachweis fehlt"]
    for k, lab in enumerate(qlab):
        put(ws, f"M{27 + k}", lab, size=9, border=THIN)
        put(ws, f"N{27 + k}", f"=COUNTIF(Q_StatusNr,{k + 1})", size=9, h="center", border=THIN)

    # Diagramme
    def anchor_chart(ch, cell):
        ws.add_chart(ch, cell)

    c1 = _bar("Urlaubs- und Krankheitstage je Monat", "Arbeitstage")
    data = Reference(ws, min_col=3, max_col=4, min_row=26, max_row=38)
    c1.add_data(data, titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=2, min_row=27, max_row=38))
    _color_series(c1.series[0], SERIES1)
    _color_series(c1.series[1], SERIES2)
    anchor_chart(c1, "B41")

    c2 = _bar("Nachtdienst-Status je Monat (Tage)", "Tage", stacked=True)
    c2.add_data(Reference(ws, min_col=5, max_col=7, min_row=26, max_row=38), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=2, min_row=27, max_row=38))
    for s, col in zip(c2.series, [ST_GOOD, ST_WARN, ST_BAD]):
        _color_series(s, col)
    c2.gapWidth = 40
    anchor_chart(c2, "I41")

    c3 = LineChart()
    c3.title = "Auslastung je Monat (Ist-/Sollstunden)"
    c3.style = 10
    c3.height, c3.width = 7.2, 15.5
    c3.add_data(Reference(ws, min_col=8, min_row=26, max_row=38), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=2, min_row=27, max_row=38))
    s = c3.series[0]
    s.graphicalProperties.line.solidFill = SERIES1
    s.graphicalProperties.line.width = 25400
    s.marker.symbol = "circle"
    s.marker.size = 7
    s.marker.graphicalProperties.solidFill = SERIES1
    s.marker.graphicalProperties.line.solidFill = "FFFFFF"
    c3.y_axis.number_format = "0%"
    c3.y_axis.scaling.min = 0
    c3.y_axis.scaling.max = 1.2
    c3.y_axis.majorGridlines.spPr = None
    c3.x_axis.delete = False
    c3.y_axis.delete = False
    c3.legend = None
    anchor_chart(c3, "B57")

    c4 = _bar("Nachtdienste voll einsatzfähig je Team (Quote)", horizontal=False)
    c4.add_data(Reference(ws, min_col=9, min_row=18, max_row=23), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=2, min_row=19, max_row=23))
    c4.legend = None
    c4.y_axis.number_format = "0%"
    c4.y_axis.scaling.min = 0
    c4.y_axis.scaling.max = 1
    ser = c4.series[0]
    for k, (t, kk, col) in enumerate(TEAMS):
        pt = DataPoint(idx=k)
        pt.graphicalProperties.solidFill = col
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        ser.dPt.append(pt)
    ser.dLbls = _labels("0%")
    anchor_chart(c4, "I57")

    c5 = _bar("Qualifikationsstatus (Anzahl Nachweise)")
    c5.add_data(Reference(ws, min_col=14, min_row=26, max_row=31), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=27, max_row=31))
    c5.legend = None
    ser = c5.series[0]
    for k, col in enumerate([ST_GOOD, ST_WARN, ST_BAD, ST_CRIT, "A6A6A6"]):
        pt = DataPoint(idx=k)
        pt.graphicalProperties.solidFill = col
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        ser.dPt.append(pt)
    ser.dLbls = _labels("0")
    anchor_chart(c5, "B73")

    put(ws, "I73", "Lesehilfe", size=10, bold=True)
    tips = ["Ampel Grün = Sollbesetzung erreicht · Gelb = kritisch (Lücke mit Ersatz / keine Reserve) · Rot = Unterbesetzung.",
            "Alle Kennzahlen rechnen automatisch aus Mitarbeiterstamm, Funktionsmatrix, Qualifikationen, "
            "Abwesenheiten, Vertretungen und Zeitbuchungen.",
            "Stichtag im Blatt Einstellungen (Standard: heute innerhalb des Planungsjahres), Berichtsmonat oben rechts.",
            "Details: Einsatzfähigkeitsprüfung (je Tag), Personalbedarfsplanung (Tag/Woche/Monat/Jahr), Statistiken."]
    for k, t in enumerate(tips):
        put(ws, f"I{74 + k}", t, size=9, color="595959", wrap=True)
        ws.merge_cells(f"I{74 + k}:O{74 + k}")
        ws.row_dimensions[74 + k].height = 28
    ws.column_dimensions["A"].width = 2
    for c in range(2, 17):
        ws.column_dimensions[L(c)].width = 11.5
    ws.column_dimensions["F"].width = 26
    landscape(ws, "A3", fit_height=True, area="A1:P88")
    ws.sheet_properties.tabColor = "C00000"


# ====================================================================== Startseite
NAV = [
    ("Eingaben", "FFC000", [
        ("Einstellungen", "Planungsjahr, Stichtag, Rotation, Dienstzeiten, Sollbesetzung, Ampel, Listen, Feiertage"),
        ("Mitarbeiterstamm", "Blatt 1 – Stammdaten, Urlaubsanspruch, Resturlaub, Überstunden, JAZ"),
        ("Team- und Funktionsmatrix", "Blatt 2 – Teamfunktion und Mehrfachfunktionen (JA/NEIN)"),
        ("Qualifikationen", "Blatt 3 – Schulungsnachweise, Ablaufdaten, Restlaufzeit, Ampel"),
        ("Urlaub und Abwesenheiten", "Blatt 4 – Urlaub, Krankheit, JAZ, Unterricht, Fortbildung, Freistellung, Sonderurlaub"),
        ("Vertretungen", "Ersatzsuche und bestätigte Vertretungen"),
        ("Zeitbuchungen", "Vorträge, Auszahlungen, Korrekturen für Überstunden-, JAZ- und Urlaubskonto")]),
    ("Planung (automatisch)", "4F81BD", [
        ("Schichtplan 2027", "Blatt 5 – Nachtteam, Freistellung, Tagdienstteams je Tag"),
        ("PEP 2027", "Personaleinsatzplan je Mitarbeiter und Tag (Layout wie PEP)"),
        ("Einsatzfähigkeitsprüfung", "Blatt 6 – Prüfung je Team und Tag, fehlende Funktionen, Ersatzvorschläge"),
        ("Überstunden und JAZ", "Blatt 7 – Soll/Ist, Monats- und Jahreswerte, Salden, JAZ-Konto"),
        ("Personalbedarfsplanung", "Soll/Ist, Verfügbarkeit, Reserve, Unter-/Überbesetzung – Tag, Woche, Monat, Jahr")]),
    ("Auswertungen", "00B050", [
        ("Dashboard", "Management-Dashboard"),
        ("Statistiken", "Nachtdienste, Wochenend-/Feiertagsdienste, Bereitschaften, Abwesenheiten je MA"),
        ("Dienstverteilung", "Dienst- und Nachtdienstverteilung je Monat"),
        ("Urlaubsstatistik", "Urlaub je Monat, Resturlaub, Urlaubsquote"),
        ("Krankenstatistik", "Krankheitstage, Krankmeldungen, Krankenquote"),
        ("Monatsbericht", "Bericht für einen wählbaren Monat"),
        ("Jahresbericht", "Monatsübersicht, Teams, Kernaussagen"),
        ("Einsatzfähigkeitsbericht", "Einsatzfähigkeit je Team, Ursachen, nächste kritische Nachtdienste")]),
    ("Druckansichten (A4/A3 quer)", "7F7F7F", [
        ("Druck Jahr", "Jahresübersicht (A3 quer)"),
        ("Druck Monat", "Monatsübersicht je Mitarbeiter (A4 quer, Monat wählbar)"),
        ("Druck Quartal", "Quartalsübersicht je Mitarbeiter (A3 quer, Quartal wählbar)"),
        ("Druck Team", "Teamübersicht (A4 quer, Team wählbar)"),
        ("Druck Bereitschaft", "Bereitschaftsübersicht im Layout des Bereitschaftskalenders (A3 quer)"),
        ("Druck Qualifikation", "Qualifikationsübersicht (A4/A3 quer)")]),
]


def build_start(book, ws, notes):
    ws.sheet_view.showGridLines = False
    put(ws, "B1", '="Personal- und Einsatzplanung "&Jahr', size=20, bold=True)
    ws.row_dimensions[1].height = 34
    put(ws, "B2", "Vollautomatische Personal-, Schicht-, Bereitschafts-, Urlaubs-, Qualifikations- und "
                  "Ressourcenplanung auf Basis der PEP-Planung und des Bereitschaftskalenders.", size=10, color="595959")
    r = 4
    put(ws, f"B{r}", "So wird gearbeitet", size=12, bold=True)
    steps = ["1. Nur in gelb markierten Spalten/Zellen eingeben (Eingabeblätter mit gelbem Register).",
             "2. Mitarbeiter im Mitarbeiterstamm und in der Team- und Funktionsmatrix pflegen.",
             "3. Schulungsdaten im Blatt Qualifikationen, Abwesenheiten im Blatt Urlaub und Abwesenheiten erfassen.",
             "4. Ersatzvorschläge im Blatt Einsatzfähigkeitsprüfung bzw. Vertretungen prüfen und bestätigte "
             "Vertretungen eintragen.",
             "5. Alles andere (Plan, Prüfungen, Zeitkonten, Statistiken, Dashboard, Druck) rechnet automatisch.",
             "Excel-Tabellen nicht sortieren/filtern, während Formeln neu berechnet werden; Blätter bitte nicht "
             "umbenennen (Formeln nutzen die Blattnamen)."]
    for k, s in enumerate(steps):
        put(ws, f"B{r + 1 + k}", s, size=10)
    r = 12
    for grp, col, items in NAV:
        put(ws, f"B{r}", grp, size=11, bold=True, color="FFFFFF", bg=col)
        ws.merge_cells(f"B{r}:D{r}")
        r += 1
        for sheet, desc in items:
            c = put(ws, f"B{r}", sheet, size=10, color="0563C1")
            c.hyperlink = Hyperlink(ref=f"B{r}", location=f"'{sheet}'!A1", display=sheet)
            c.font = font(10, color="0563C1", underline="single")
            put(ws, f"C{r}", desc, size=9, color="595959")
            ws.merge_cells(f"C{r}:D{r}")
            r += 1
        r += 1
    # Legende
    lg = 12
    put(ws, f"F{lg}", "Legende Dienst- und Abwesenheitscodes", size=11, bold=True)
    items = [(c, t, bg, fc) for c, t, bg, fc, b in DUTY_CODES] + [(c, t, bg, fc) for c, t, bg, fc, b in ABW_CODES]
    for k, (code, text, bg, fc) in enumerate(items):
        put(ws, f"F{lg + 1 + k}", code, bold=True, h="center", bg=bg, color=fc, border=THIN)
        put(ws, f"G{lg + 1 + k}", text, size=9)
    k0 = lg + 2 + len(items)
    put(ws, f"F{k0}", "Teamfarben", size=11, bold=True)
    for k, (t, kk, col) in enumerate(TEAMS):
        put(ws, f"F{k0 + 1 + k}", kk, bold=True, h="center", bg=col, border=THIN)
        put(ws, f"G{k0 + 1 + k}", t, size=9)
    a0 = k0 + 7
    put(ws, f"F{a0}", "Ampel", size=11, bold=True)
    for k, (t, txt) in enumerate([("Grün", "Sollbesetzung erreicht"),
                                  ("Gelb", "kritische Besetzung (keine Reserve bzw. Lücke mit Ersatz)"),
                                  ("Rot", "Unterbesetzung (Lücke ohne Ersatz)")]):
        bg, fc = AMPEL[t]
        put(ws, f"F{a0 + 1 + k}", t, bold=True, h="center", bg=bg, color=fc, border=THIN)
        put(ws, f"G{a0 + 1 + k}", txt, size=9)
    # Hinweise / Annahmen
    h0 = max(r, a0 + 5) + 1
    put(ws, f"B{h0}", "Annahmen und Hinweise (bitte fachlich prüfen – alle Werte im Blatt Einstellungen änderbar)",
        size=11, bold=True)
    base_notes = [
        "Ganze Kalendertage, keine Stundenplanung. Nachtteam hat Bereitschaft/Nachtdienst statt Tagdienst; "
        "alle anderen Teams arbeiten Mo–Fr (ohne Feiertag) im Tagdienst.",
        "Freistellung: nach Mo–Do-Nachtdienst Freitag, nach Fr–So-Nachtdienst Montag und Dienstag. Vor dem "
        "Rotationsbeginn gibt es keine Freistellung.",
        "Einsatzfähigkeit zählt Mitarbeiter mit passender Teamfunktion; Mehrfachfunktionen wirken über die "
        "Ersatzlogik (P1 Reserve/Überdeckung im Team, P2 anderes Team gleiche Funktion, P3 anderes Team mit "
        "Mehrfachfunktion; nicht aus dem freigestellten Team). Innerhalb einer Priorität rotieren die Vorschläge "
        "blockweise gleichmäßig über alle geeigneten Mitarbeiter; die manuelle Ersatzsuche (Blatt Vertretungen) "
        "sortiert nach wenigsten Nachtdiensten.",
        "Zeitkonten: Sollstunden 8,0 je Arbeitstag, Tagdienst 06:45–15:30 abzgl. 45 min Pause, Bereitschaft "
        "15:30–06:45 zu 60 % angerechnet (gleicht Freistellungstage aus), Freistellung nach Nacht 0 Std., JAZ-Tage "
        "werden vom JAZ-Konto abgebucht.",
        "Mindestbesetzung Tagdienst 12 Personen, gewünschte Reserve 2 Personen; Pflichtqualifikationen je Funktion "
        "und Gültigkeitsdauern der Unterrichte sind Annahmen.",
        "Fehlende Schulungsnachweise blockieren erst, wenn in den Einstellungen „Fehlender Nachweis gilt als "
        "ungültig = JA“ gesetzt wird.",
    ] + notes
    for k, t in enumerate(base_notes):
        put(ws, f"B{h0 + 1 + k}", "• " + t, size=9, wrap=True, v="top")
        ws.merge_cells(f"B{h0 + 1 + k}:J{h0 + 1 + k}")
        ws.row_dimensions[h0 + 1 + k].height = 26
    col_widths(ws, {"A": 2, "B": 30, "C": 40, "D": 30, "E": 3, "F": 8, "G": 48, "H": 10, "I": 10, "J": 10})
    landscape(ws, "A4", fit_height=True)
    ws.sheet_properties.tabColor = "000000"


def build_all(book, wb, people, demo, notes=()):
    build_statistik(book, wb["Statistiken"])
    build_dienstverteilung(book, wb["Dienstverteilung"])
    build_urlaub(book, wb["Urlaubsstatistik"])
    build_krank(book, wb["Krankenstatistik"])
    build_monatsbericht(book, wb["Monatsbericht"])
    build_jahresbericht(book, wb["Jahresbericht"])
    build_efbericht(book, wb["Einsatzfähigkeitsbericht"])
    build_dashboard(book, wb["Dashboard"])
    build_start(book, wb["Start"], list(notes))
