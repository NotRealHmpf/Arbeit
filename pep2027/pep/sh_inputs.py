"""Eingabetabellen: Abwesenheiten (Blatt 3), Ressourcen, Freistellungen, Vertretungen."""
import datetime as dt

from openpyxl.utils import get_column_letter as gl

from . import model as M
from .ctx import S_ABW, S_FS, S_RES, S_VER
from .sh_master import tr
from .style import (BOX, INPUT_FILL, add_table, cf, dv_date, dv_list, dxf, fill, note, put, title, widths)

DATE_NF = "DDD DD.MM.YYYY"


def _table(ws, name, hdr, rows, formulas, first_row, input_cols, extra_rows=200, style="TableStyleMedium2",
           date_cols=(), center_cols=()):
    """Tabelle mit Kopf in first_row-1 schreiben. rows: Liste von dicts (nur Eingabespalten)."""
    hr = first_row - 1
    for i, h in enumerate(hdr):
        put(ws, f"{gl(2 + i)}{hr}", h, bold=True, wrap=True)
    n = max(len(rows), 1)
    for k in range(n):
        r = first_row + k
        data = rows[k] if k < len(rows) else {}
        for i, h in enumerate(hdr):
            c = f"{gl(2 + i)}{r}"
            if h in formulas:
                put(ws, c, "=" + formulas[h], color="404040", h="center" if h not in ("Prüfung",) else None,
                    nf=DATE_NF if h in date_cols else None)
            else:
                v = data.get(h)
                put(ws, c, v, nf=DATE_NF if h in date_cols else None, h="center" if h in center_cols else None)
                if h in input_cols:
                    ws[c].fill = fill(INPUT_FILL)
    last = first_row + n - 1
    add_table(ws, name, f"B{hr}:{gl(1 + len(hdr))}{last}", formulas=formulas, style=style)
    ws.row_dimensions[hr].height = 30
    return last


def col_of(hdr, h):
    return gl(2 + hdr.index(h))


def build_abwesenheiten(ctx, examples):
    ws = ctx.wb[S_ABW]
    ws.sheet_properties.tabColor = "00B050"
    title(ws, "B1", "Blatt 3 – Abwesenheiten (zentrale Datenquelle)")
    note(ws, "B2", "Jede Zeile = ein Zeitraum. Einträge erscheinen automatisch im Jahreskalender, in der "
         "Einsatzfähigkeitsprüfung und in allen Auswertungen. Direkteingaben im Kalender (mit Makros) werden "
         "hierher zurückgeschrieben (Quelle = Kalender).")
    note(ws, "B3", "Überschneidungen sind erlaubt; es gilt die höhere Priorität: Krankheit > Unterricht > FDL > "
         "Kabel > Freistellung > Urlaub eTime > Urlaub > JAZ > Besonderheit.")
    T = "tblAbwesenheiten"
    hdr = ["Mitarbeiter", "Von", "Bis", "Art", "Bemerkung", "Code", "Team", "Kalendertage", "Arbeitstage",
           "Prüfung", "Quelle", "Nr"]
    MA, VON, BIS, ART = tr(T, "Mitarbeiter"), tr(T, "Von"), tr(T, "Bis"), tr(T, "Art")
    f = {
        "Code": f"IFERROR(INDEX(tblArten[Code],MATCH({ART},tblArten[Bezeichnung],0)),0)",
        "Team": f'IFERROR(INDEX(MA_Team,MATCH({MA},MA_Name,0))&"","")',
        "Kalendertage": f'IF(AND(ISNUMBER({VON}),ISNUMBER({BIS})),{BIS}-{VON}+1,"")',
        "Arbeitstage": f'IF(AND(ISNUMBER({VON}),ISNUMBER({BIS})),NETWORKDAYS({VON},{BIS},FT_Aktiv),"")',
        "Prüfung": (f'IF({MA}="","",IF(COUNTIF(MA_Name,{MA})=0,"⚠ Mitarbeiter unbekannt",'
                    f'IF(OR(NOT(ISNUMBER({VON})),NOT(ISNUMBER({BIS}))),"⚠ Datum fehlt",'
                    f'IF({BIS}<{VON},"⚠ Bis liegt vor Von",IF({tr(T, "Code")}=0,"⚠ Art fehlt",'
                    f'IF({tr(T, "Code")}>8,"⚠ Art gehört nicht in diese Tabelle",'
                    f'IF(COUNTIFS({T}[Mitarbeiter],{MA},{T}[Von],"<="&{BIS},{T}[Bis],">="&{VON})>1,'
                    f'"ℹ Überschneidung – höhere Priorität gilt","✓")))))))'),
        "Nr": f"ROW()-ROW({T}[#Headers])",
    }
    rows = [dict(zip(["Mitarbeiter", "Von", "Bis", "Art", "Bemerkung"], e)) for e in examples]
    last = _table(ws, T, hdr, rows, f, 6, ["Mitarbeiter", "Von", "Bis", "Art", "Bemerkung", "Quelle"],
                  date_cols=("Von", "Bis"), center_cols=("Art",))
    widths(ws, {"A": 2, "B": 22, "C": 16, "D": 16, "E": 24, "F": 34, "G": 7, "H": 10, "I": 12, "J": 12,
                "K": 36, "L": 11, "M": 6})
    end = last + 500
    dv_list(ws, f"B6:B{end}", "lstMA", prompt="Mitarbeiter aus dem Stamm wählen")
    dv_date(ws, f"C6:D{end}")
    dv_list(ws, f"E6:E{end}", "lstArtenAuswahl", prompt="Art der Abwesenheit wählen")
    for code in range(1, 9):
        s = M.STATUS_BY_CODE[code]
        cf(ws, f"E6:E{end}", f'$G6={code}', dxf(bg=s[5], fg=s[6], bold=s[7], underline="single" if s[8] else None))
    cf(ws, f"K6:K{end}", 'LEFT($K6,1)="⚠"', dxf(bg="FFC7CE", fg="9C0006"))
    cf(ws, f"K6:K{end}", 'LEFT($K6,1)="ℹ"', dxf(fg="9C5700"))
    for nm, c in [("Abw_MA", "Mitarbeiter"), ("Abw_Von", "Von"), ("Abw_Bis", "Bis"), ("Abw_Code", "Code"),
                  ("Abw_Nr", "Nr"), ("Abw_Bem", "Bemerkung"), ("Abw_Art", "Art")]:
        ctx.name(nm, f"{T}[{c}]")
    # kleine Uebersicht
    put(ws, "O5", "Übersicht", bold=True, bg=M.HEADER_FILL, border=BOX)
    put(ws, "P5", "Einträge", bold=True, bg=M.HEADER_FILL, border=BOX, h="center")
    put(ws, "Q5", "Arbeitstage", bold=True, bg=M.HEADER_FILL, border=BOX, h="center")
    for i, art in enumerate(M.ARTEN_AUSWAHL):
        r = 6 + i
        code = M.CODE[art]
        s = M.STATUS_BY_CODE[code]
        put(ws, f"O{r}", art, border=BOX, bg=s[5], color=s[6], bold=s[7], underline="single" if s[8] else None)
        put(ws, f"P{r}", f'=COUNTIF(Abw_Code,{code})', border=BOX, h="center")
        put(ws, f"Q{r}", f'=SUMIFS({T}[Arbeitstage],Abw_Code,{code})', border=BOX, h="center")
    r = 6 + len(M.ARTEN_AUSWAHL)
    put(ws, f"O{r}", "Prüfhinweise (⚠)", border=BOX, bold=True)
    put(ws, f"P{r}", f'=COUNTIF({T}[Prüfung],"⚠*")', border=BOX, h="center", bold=True)
    note(ws, f"O{r + 1}", "Arbeitstage = Mo–Fr ohne Feiertage (Hessen).")
    widths(ws, {"N": 3, "O": 26, "P": 10, "Q": 12})
    ws.freeze_panes = "C6"
    return ws


def build_ressourcen(ctx, examples):
    ws = ctx.wb[S_RES]
    ws.sheet_properties.tabColor = "00B0F0"
    title(ws, "B1", "Ressourcenplanung – Zusätzliche MA und AZUBIS")
    note(ws, "B2", "Zusätzliche Ressourcen sind keine Abwesenheiten. Sie werden im Kalender gelb (Zusätzliche MA) "
         "bzw. hellblau (AZUBIS) dargestellt und in allen Auswertungen getrennt ausgewiesen.")
    note(ws, "B3", "Sie ersetzen nie die Sollbesetzung (Tagdienst und Bereitschaft werden ohne sie geprüft).")
    T = "tblRessourcen"
    hdr = ["Ressource", "Art", "Von", "Bis", "Team", "Bemerkung", "Code", "Im Stamm", "Kalendertage",
           "Arbeitstage", "Prüfung", "Quelle"]
    NM, ART, VON, BIS = tr(T, "Ressource"), tr(T, "Art"), tr(T, "Von"), tr(T, "Bis")
    f = {
        "Code": f'IF({ART}="Zusätzliche MA",12,IF({ART}="AZUBIS",13,0))',
        "Im Stamm": f'IF(COUNTIF(MA_Name,{NM})>0,1,0)',
        "Kalendertage": f'IF(AND(ISNUMBER({VON}),ISNUMBER({BIS})),{BIS}-{VON}+1,"")',
        "Arbeitstage": f'IF(AND(ISNUMBER({VON}),ISNUMBER({BIS})),NETWORKDAYS({VON},{BIS},FT_Aktiv),"")',
        "Prüfung": (f'IF({NM}="","",IF({tr(T, "Code")}=0,"⚠ Art fehlt",'
                    f'IF(OR(NOT(ISNUMBER({VON})),NOT(ISNUMBER({BIS}))),"⚠ Datum fehlt",'
                    f'IF({BIS}<{VON},"⚠ Bis liegt vor Von",IF({tr(T, "Im Stamm")}=0,'
                    f'"ℹ nicht im Mitarbeiterstamm – zählt nur in den Summen","✓")))))'),
    }
    rows = [dict(zip(["Ressource", "Art", "Von", "Bis", "Team", "Bemerkung"], e)) for e in examples]
    last = _table(ws, T, hdr, rows, f, 6, ["Ressource", "Art", "Von", "Bis", "Team", "Bemerkung", "Quelle"],
                  date_cols=("Von", "Bis"), center_cols=("Art", "Team"))
    widths(ws, {"A": 2, "B": 22, "C": 16, "D": 16, "E": 16, "F": 10, "G": 30, "H": 7, "I": 9, "J": 12,
                "K": 12, "L": 44, "M": 11})
    end = last + 500
    dv_list(ws, f"B6:B{end}", "lstMA", error=False, prompt="Ressource aus dem Stamm wählen oder frei eingeben")
    dv_list(ws, f"C6:C{end}", "lstResArten")
    dv_date(ws, f"D6:E{end}")
    dv_list(ws, f"F6:F{end}", "lstTeams")
    cf(ws, f"C6:C{end}", '$C6="Zusätzliche MA"', dxf(bg="FFFF00"))
    cf(ws, f"C6:C{end}", '$C6="AZUBIS"', dxf(bg="00B0F0"))
    cf(ws, f"L6:L{end}", 'LEFT($L6,1)="⚠"', dxf(bg="FFC7CE", fg="9C0006"))
    for nm, c in [("R_Name", "Ressource"), ("R_Art", "Art"), ("R_Von", "Von"), ("R_Bis", "Bis"),
                  ("R_Code", "Code"), ("R_Stamm", "Im Stamm")]:
        ctx.name(nm, f"{T}[{c}]")
    ws.freeze_panes = "C6"
    return ws


def build_freistellungen(ctx, examples):
    ws = ctx.wb[S_FS]
    ws.sheet_properties.tabColor = "7F7F7F"
    title(ws, "B1", "Freistellungen – manuelle Anpassungen")
    note(ws, "B2", "Automatischer Vorschlag: nach Mo–Do-Nachtdienst Freitag frei, nach Fr–So-Nachtdienst Montag "
         "und Dienstag frei. Einträge hier haben immer Vorrang und werden nie automatisch überschrieben.")
    note(ws, "B3", "Beispiel „Dienstag arbeiten, stattdessen Freitag frei“: eine Zeile mit „Freistellung entfällt am“ = "
         "Dienstag und „Stattdessen frei am“ = Freitag. Ohne Mitarbeiter gilt die Zeile für das ganze Team.")
    T = "tblFreistellungen"
    hdr = ["Team", "Mitarbeiter", "Freistellung entfällt am", "Stattdessen frei am", "Bemerkung", "Gilt für",
           "Prüfung", "Quelle"]
    TEAM, MA = tr(T, "Team"), tr(T, "Mitarbeiter")
    ENT, STA = tr(T, "Freistellung entfällt am"), tr(T, "Stattdessen frei am")
    grp = f'IF({MA}<>"",IFERROR(INDEX(MA_Gruppe,MATCH({MA},MA_Name,0)),""),{TEAM})'
    auto = f'IFERROR(INDEX(Schicht_FSauto,{ENT}-DATE(Jahr,1,1)+1),"")'
    f = {
        "Gilt für": f'IF({MA}<>"",{MA},IF({TEAM}<>"","Team "&{TEAM},""))',
        "Prüfung": (f'IF(AND({TEAM}="",{MA}="",{ENT}="",{STA}=""),"",IF(AND({TEAM}="",{MA}=""),'
                    f'"⚠ Team oder Mitarbeiter angeben",IF(AND({ENT}="",{STA}=""),"⚠ Datum fehlt",'
                    f'IF(AND({MA}<>"",COUNTIF(MA_Name,{MA})=0),"⚠ Mitarbeiter unbekannt",'
                    f'IF(AND(ISNUMBER({ENT}),{auto}<>{grp}),"ℹ an diesem Tag keine automatische Freistellung",'
                    f'"✓")))))'),
    }
    rows = [dict(zip(["Team", "Mitarbeiter", "Freistellung entfällt am", "Stattdessen frei am", "Bemerkung"], e))
            for e in examples]
    last = _table(ws, T, hdr, rows, f, 6, ["Team", "Mitarbeiter", "Freistellung entfällt am",
                                           "Stattdessen frei am", "Bemerkung", "Quelle"],
                  date_cols=("Freistellung entfällt am", "Stattdessen frei am"), center_cols=("Team",))
    widths(ws, {"A": 2, "B": 10, "C": 20, "D": 18, "E": 18, "F": 36, "G": 18, "H": 44, "I": 11})
    end = last + 500
    dv_list(ws, f"B6:B{end}", "lstTeams")
    dv_list(ws, f"C6:C{end}", "lstMA")
    dv_date(ws, f"D6:E{end}")
    for t in M.TEAMS:
        cf(ws, f"B6:B{end}", f'$B6="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
    cf(ws, f"H6:H{end}", 'LEFT($H6,1)="⚠"', dxf(bg="FFC7CE", fg="9C0006"))
    cf(ws, f"H6:H{end}", 'LEFT($H6,1)="ℹ"', dxf(fg="9C5700"))
    for nm, c in [("FS_Team", "Team"), ("FS_MA", "Mitarbeiter"), ("FS_Entf", "Freistellung entfällt am"),
                  ("FS_Statt", "Stattdessen frei am")]:
        ctx.name(nm, f"{T}[{c}]")
    ws.freeze_panes = "B6"
    return ws


def build_vertretungen(ctx, examples):
    ws = ctx.wb[S_VER]
    ws.sheet_properties.tabColor = "BFBFBF"
    title(ws, "B1", "Vertretungen in der Bereitschaft")
    note(ws, "B2", "Übernommene Ersatzvorschläge hier eintragen (oder im Kalender „V“ eingeben). Der Vertreter "
         "erscheint im Kalender grau (Bereitschaft) und zählt in der Einsatzfähigkeitsprüfung für die Funktion.")
    T = "tblVertretungen"
    hdr = ["Von", "Bis", "Vertreter", "Funktion", "Für (vertretener MA)", "Bemerkung", "Team Vertreter",
           "Bereitschaftsteam", "Qualifiziert", "Prüfung", "Quelle"]
    VON, BIS, VT, FN = tr(T, "Von"), tr(T, "Bis"), tr(T, "Vertreter"), tr(T, "Funktion")
    team_of = f'MOD(2*INT(({VON}-RotStart)/7)+(MOD({VON}-RotStart,7)>=3),10)+1'
    f = {
        "Team Vertreter": f'IFERROR(INDEX(MA_Team,MATCH({VT},MA_Name,0))&"","")',
        "Bereitschaftsteam": f'IF(ISNUMBER({VON}),INDEX(RotTeam,{team_of}),"")',
        "Qualifiziert": (f'IF(OR({VT}="",{FN}=""),"",IF(IFERROR(INDEX(MA_Funktion,MATCH({VT},MA_Name,0)),"")={FN},"JA",'
                         f'IF(IFERROR(INDEX(tblFunktionsmatrix[[Meister]:[Handwerker]],MATCH({VT},FM_Name,0),'
                         f'MATCH({FN},lstRollen,0)),"")="JA","JA","NEIN")))'),
        "Prüfung": (f'IF({VT}="","",IF(COUNTIF(MA_Name,{VT})=0,"⚠ Vertreter unbekannt",'
                    f'IF(OR(NOT(ISNUMBER({VON})),NOT(ISNUMBER({BIS}))),"⚠ Datum fehlt",IF({BIS}<{VON},'
                    f'"⚠ Bis liegt vor Von",IF({tr(T, "Qualifiziert")}<>"JA","⚠ Funktion nicht qualifiziert",'
                    f'IF(COUNTIFS(Abw_MA,{VT},Abw_Von,"<="&{BIS},Abw_Bis,">="&{VON},Abw_Code,">=2")>0,'
                    f'"⚠ Vertreter im Zeitraum abwesend",IF({tr(T, "Team Vertreter")}={tr(T, "Bereitschaftsteam")},'
                    f'"ℹ gehört selbst zum Bereitschaftsteam","✓")))))))'),
    }
    rows = [dict(zip(["Von", "Bis", "Vertreter", "Funktion", "Für (vertretener MA)", "Bemerkung"], e))
            for e in examples]
    last = _table(ws, T, hdr, rows, f, 6, ["Von", "Bis", "Vertreter", "Funktion", "Für (vertretener MA)",
                                           "Bemerkung", "Quelle"],
                  date_cols=("Von", "Bis"), center_cols=("Funktion",))
    widths(ws, {"A": 2, "B": 16, "C": 16, "D": 20, "E": 13, "F": 20, "G": 30, "H": 12, "I": 14, "J": 11,
                "K": 36, "L": 11})
    end = last + 500
    dv_date(ws, f"B6:C{end}")
    dv_list(ws, f"D6:D{end}", "lstMA")
    dv_list(ws, f"E6:E{end}", "lstVertrFunktion")
    dv_list(ws, f"F6:F{end}", "lstMA")
    for t in M.TEAMS:
        cf(ws, f"H6:I{end}", f'H6="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
    cf(ws, f"K6:K{end}", 'LEFT($K6,1)="⚠"', dxf(bg="FFC7CE", fg="9C0006"))
    for nm, c in [("V_Von", "Von"), ("V_Bis", "Bis"), ("V_MA", "Vertreter"), ("V_Funktion", "Funktion"),
                  ("V_Fuer", "Für (vertretener MA)")]:
        ctx.name(nm, f"{T}[{c}]")
    ws.freeze_panes = "B6"
    return ws
