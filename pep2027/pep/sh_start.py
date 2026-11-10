"""Startseite (Navigation, Anleitung, Legende) und Management-Dashboard."""
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

from . import model as M
from .ctx import (DESCR, ORDER, S_ABW, S_BER, S_DASH, S_FS, S_JAHR, S_KAL, S_MA, S_RES, S_SET, S_START, S_VER)
from .style import (BOX, INPUT_FILL, cf, dv_list, dxf, fill, font, link, note, put, q, title, widths)

AMPEL = '"🟢","🟡","🔴"'


def build_start(ctx):
    ws = ctx.wb[S_START]
    ws.sheet_properties.tabColor = M.TITLE_FILL
    widths(ws, {"A": 2, "B": 30, "C": 62, "D": 3, "E": 26, "F": 12, "G": 3, "H": 30, "I": 40})
    ws.merge_cells("B1:I1")
    put(ws, "B1", '="PEP "&Jahr&" – Personal- und Einsatzplanung (Schicht, Bereitschaft, Urlaub, Abwesenheit)"',
        size=18, bold=True, bg=M.TITLE_FILL, color="FFFFFF")
    ws.row_dimensions[1].height = 34
    note(ws, "B2", "Vollständig verknüpft: Jede Eingabe in Mitarbeiter, Funktionsmatrix, Abwesenheiten, Ressourcen, "
         "Freistellungen oder Vertretungen wirkt sofort auf Kalender, Prüfungen, Auswertungen, Dashboard und "
         "Druckansichten.", wrap=True)
    ws.merge_cells("B2:I2")
    ws.row_dimensions[2].height = 28

    # Status
    put(ws, "B4", "Status", bold=True, size=12)
    ws.merge_cells("B5:C5")
    put(ws, "B5", '=IFERROR(IF(PEPMakrosAktiv(),"Makros aktiv: Doppelklick (ein Tag) oder Rechtsklick (markierte Tage) im Jahreskalender öffnet die Eingabe – sie wird automatisch in die Tabellen geschrieben.",""),"⚠ Makros nicht aktiv: Eingaben über die Tabellen (z. B. Abwesenheiten) – der Kalender zeigt sie automatisch an.")',
        bold=True, color="C00000", size=9, wrap=True)
    cf(ws, "B5", 'LEFT(B5,6)="Makros"', dxf(fg="006100"))
    ws.row_dimensions[5].height = 26
    ctx.name("Makrostatus", f"{S_START}!$B$5")
    stat = [
        ("Planungsjahr", "=Jahr"),
        ("Mitarbeiter in Teams / ohne Team / Ressourcen",
         '=COUNTIF(MA_Kat,"Team")&" / "&COUNTIF(MA_Kat,"Ohne Team")&" / "&COUNTIF(MA_Kat,"Ressource")'),
        ("Kalenderzeilen belegt / Kapazität",
         '=KalZeilenBelegt&" / "&KalKapazitaet&IF(KalZeilenBelegt>KalKapazitaet,"  ⚠ Kapazität überschritten","")'),
        ("Prüfhinweise (⚠) in den Eingabetabellen",
         '=COUNTIF(tblMitarbeiter[Prüfung],"⚠*")+COUNTIF(tblFunktionsmatrix[Prüfung],"⚠*")+'
         'COUNTIF(tblAbwesenheiten[Prüfung],"⚠*")+COUNTIF(tblRessourcen[Prüfung],"⚠*")+'
         'COUNTIF(tblFreistellungen[Prüfung],"⚠*")+COUNTIF(tblVertretungen[Prüfung],"⚠*")'),
        ("Tage Bereitschaft nicht einsatzfähig (🔴)", '=COUNTIF(Einsatz_Stufe,2)'),
        ("Rotationsstart", '=RIGHT("0"&DAY(RotStart),2)&"."&RIGHT("0"&MONTH(RotStart),2)&"."&YEAR(RotStart)'
                           '&" – "&Einstellungen!C8'),
    ]
    for i, (lab, f_) in enumerate(stat):
        r = 6 + i
        put(ws, f"B{r}", lab, border=BOX, size=9)
        put(ws, f"C{r}", f_, border=BOX, size=9, bold=True)
    cf(ws, "C9", "C9>0", dxf(bg="FFC7CE", fg="9C0006"))

    # Navigation
    r0 = 13
    put(ws, f"B{r0}", "Navigation", bold=True, size=12)
    for i, name in enumerate([n for n in ORDER if n not in (S_START, S_BER)]):
        r = r0 + 1 + i
        link(ws, f"B{r}", name, name, cell="A1", size=10, bold=True)
        put(ws, f"C{r}", DESCR.get(name, ""), size=9)
        ws[f"B{r}"].border = BOX
        ws[f"C{r}"].border = BOX
    # Kurzanleitung
    put(ws, "E4", "Kurzanleitung", bold=True, size=12)
    steps = [
        "1. Mitarbeiter: Platzhalter Ma6 … Ma28 durch echte Namen ersetzen, Team und Funktion wählen.",
        "2. Funktionsmatrix: Qualifikationen je Mitarbeiter mit JA/NEIN pflegen (Mehrfachfunktionen erlaubt).",
        "3. Abwesenheiten: Zeiträume erfassen – oder mit Makros direkt im Jahreskalender (Doppelklick/Rechtsklick).",
        "4. Freistellungen: Vorschläge bei Bedarf verschieben (z. B. Dienstag arbeiten, Freitag frei).",
        "5. Einsatzfähigkeit prüfen; übernommene Ersatzvorschläge im Blatt Vertretungen eintragen.",
        "6. Ressourcen: Zusätzliche MA und AZUBIS mit Zeitraum erfassen.",
        "7. Dashboard, Berichte und Druckansichten aktualisieren sich automatisch.",
        "8. Neues Jahr: Planungsjahr in Einstellungen ändern – Rotation läuft weiter, Feiertage werden neu berechnet.",
        "Beispielzeilen (BEISPIEL – bitte löschen) zeigen das Format und können gelöscht werden.",
    ]
    for i, t in enumerate(steps):
        ws.merge_cells(f"E{5 + i}:I{5 + i}")
        put(ws, f"E{5 + i}", t, size=9, wrap=True, v="top")
        ws.row_dimensions[5 + i].height = 15
    # Legende
    l0 = r0
    put(ws, f"E{l0}", "Legende (verbindlich in allen Blättern)", bold=True, size=12)
    put(ws, f"E{l0 + 1}", "Eintrag", bold=True, bg=M.HEADER_FILL, border=BOX, size=9)
    put(ws, f"F{l0 + 1}", "Darstellung", bold=True, bg=M.HEADER_FILL, border=BOX, size=9, h="center")
    put(ws, f"H{l0 + 1}", "Direkteingabe im Kalender (mit Makros)", bold=True, bg=M.HEADER_FILL, border=BOX, size=9)
    put(ws, f"I{l0 + 1}", "Wirkung", bold=True, bg=M.HEADER_FILL, border=BOX, size=9)
    legend = [(3, "Urlaub"), (4, "Urlaub genehmigt eTime"), (2, "JAZ + 99"), (8, "Krank"), (7, "Unterrichte"),
              (6, "FDL-Einweisungen"), (5, "Kabeleinweisungen"), (1, "Besonderheiten"), (12, "Zusätzliche MA"),
              (13, "AZUBIS"), (9, "Bereitschaft / Nachtdienst"), (10, "Freistellung (Vorschlag/angepasst)")]
    inputs = {3: ("U", "Urlaub"), 4: ("UE", "Urlaub genehmigt eTime"), 2: ("JAZ", "JAZ + 99"), 8: ("K", "Krankheit"),
              7: ("S", "Unterrichte"), 6: ("FDL", "FDL-Einweisungen"), 5: ("Kabel", "Kabeleinweisungen"),
              1: ("beliebiger Text", "Besonderheit mit diesem Text"), 12: ("Z", "Zusätzliche MA anwesend (Ressourcenzeile)"),
              13: ("A", "AZUBI anwesend (Ressourcenzeile)"), 9: ("V", "Vertretung in der Bereitschaft"),
              10: ("FS", "Freistellung setzen")}
    for i, (code, lab) in enumerate(legend):
        r = l0 + 2 + i
        s = M.STATUS_BY_CODE[code]
        put(ws, f"E{r}", lab, size=9, border=BOX)
        put(ws, f"F{r}", s[2] if s[2] and s[2] != "!" else ("Text" if code == 1 else None), size=9, bg=s[5],
            color=s[6], bold=s[7], italic=s[9], underline="single" if s[8] else None, h="center", border=BOX)
        k, w = inputs[code]
        put(ws, f"H{r}", k, size=9, border=BOX, h="center", bold=True)
        put(ws, f"I{r}", w, size=9, border=BOX)
    r = l0 + 2 + len(legend)
    for lab, bg, k, w in [("Regulärer Tagdienst", None, "-", "Abwesenheit/Freistellung aufheben (Tagdienst)"),
                          ("Wochenende", M.WEEKEND_FILL, "(leer)", "Eintrag des Tages löschen"),
                          ("Feiertag Hessen", M.HOLIDAY_FILL, "", "")]:
        put(ws, f"E{r}", lab, size=9, border=BOX)
        put(ws, f"F{r}", None, bg=bg, border=BOX)
        put(ws, f"H{r}", k or None, size=9, border=BOX, h="center", bold=True)
        put(ws, f"I{r}", w or None, size=9, border=BOX)
        r += 1
    note(ws, f"E{r + 1}", "Grundprinzip: leere/weiße Felder = regulärer Tagdienst, hellgrau = Bereitschaft/Nachtdienst. "
         "Es werden nur Abweichungen eingetragen, keine Dienstkürzel.", wrap=True)
    ws.merge_cells(f"E{r + 1}:I{r + 2}")
    # Makro-Hinweise
    h0 = r + 4
    put(ws, f"E{h0}", "Makros und Synchronisierung", bold=True, size=12)
    hints = [
        "Mit aktivierten Makros: Doppelklick auf einen Tag bzw. Rechtsklick auf markierte Tage im Jahreskalender öffnet "
        "die Eingabe. Sie wird sofort in Abwesenheiten, Freistellungen, Vertretungen bzw. Ressourcen geschrieben.",
        "Das Kalenderraster enthält immer nur Formeln und bleibt geschützt – dadurch gibt es nie zwei Datenstände "
        "(auch ohne Makros, in Excel Online und mit AutoSave). Ohne Makros erfolgt die Eingabe in den Tabellen.",
        "Datei aus E-Mail/Internet: Rechtsklick → Eigenschaften → „Zulassen“ (Mark of the Web), dann Inhalte aktivieren.",
        "Blattschutz Jahreskalender: Kennwort „pep“ (nur für Layoutänderungen nötig).",
        "Benötigt Excel 2019 oder Microsoft 365 (MAXIFS, TEXTJOIN).",
    ]
    for i, t in enumerate(hints):
        ws.merge_cells(f"E{h0 + 1 + i}:I{h0 + 1 + i}")
        put(ws, f"E{h0 + 1 + i}", "• " + t, size=9, wrap=True, v="top")
        ws.row_dimensions[h0 + 1 + i].height = 26
    ws.sheet_view.showGridLines = False
    from .style import landscape
    landscape(ws, "A4", 1, 0, area=f"A1:I{h0 + len(hints)}")
    return ws


# =========================================================================== Dashboard
def tile(ws, col, row, label, value, sub, nf=None, size=22):
    a, b = gl(col), gl(col + 2)
    for r in range(row, row + 3):
        ws.merge_cells(f"{a}{r}:{b}{r}")
    put(ws, f"{a}{row}", label, size=9, bold=True, color="FFFFFF", bg="404040", h="center", wrap=True)
    put(ws, f"{a}{row + 1}", value, size=size, bold=True, h="center", bg="F2F2F2", nf=nf)
    put(ws, f"{a}{row + 2}", sub, size=8, color="404040", h="center", bg="F2F2F2", wrap=True)
    white = Side(style="medium", color="FFFFFF")
    for r in range(row, row + 3):
        for ci in range(col, col + 3):
            ws.cell(r, ci).border = Border(left=white if ci == col else None, right=white if ci == col + 2 else None,
                                           top=white if r == row else None, bottom=white if r == row + 2 else None)
    ws.row_dimensions[row].height = 24
    ws.row_dimensions[row + 1].height = 34
    ws.row_dimensions[row + 2].height = 24
    return f"{a}{row + 1}"


def mput(ws, rng, value, **kw):
    ws.merge_cells(rng)
    first = rng.split(":")[0]
    put(ws, first, value, **kw)
    if kw.get("border") is not None:
        for row in ws[rng]:
            for c in row:
                c.border = kw["border"]


def build_dashboard(ctx):
    ws = ctx.wb[S_DASH]
    ws.sheet_properties.tabColor = M.TITLE_FILL
    ws.sheet_view.showGridLines = False
    cols = {"A": 2}
    for ci in range(2, 20):
        cols[gl(ci)] = 8.6
    widths(ws, cols)
    mput(ws, "B1:S1", '="PEP "&Jahr&" – Management-Dashboard"', size=18, bold=True, bg=M.TITLE_FILL, color="FFFFFF")
    ws.row_dimensions[1].height = 32
    put(ws, "B3", "Stichtag:", bold=True, h="right")
    mput(ws, "C3:D3", "=IF(YEAR(TODAY())=Jahr,TODAY(),WORKDAY(DATE(Jahr,1,1)-1,1,FT_Aktiv))", bg=INPUT_FILL, bold=True,
         nf="DD.MM.YYYY", border=BOX, h="center")
    put(ws, "E3", "Monat:", bold=True, h="right")
    mput(ws, "F3:G3", "=INDEX(lstMonate,MONTH(Stichtag))", bg=INPUT_FILL, bold=True, border=BOX, h="center")
    dv_list(ws, "F3", "lstMonate", prompt="Monat für die Monatskennzahlen wählen")
    put(ws, "H3", "Jahr:", bold=True, h="right")
    put(ws, "I3", "=Jahr", bold=True, h="left")
    mput(ws, "K3:S3", "=Makrostatus", size=8, color="C00000", wrap=True)
    ws.row_dimensions[3].height = 24
    ctx.name("Stichtag", f"{S_DASH}!$C$3")
    put(ws, "Z3", "=MATCH($F$3,lstMonate,0)", color="FFFFFF")
    put(ws, "Z4", "=MAX(1,MIN(DATE(Jahr,12,31)-DATE(Jahr,1,1)+1,Stichtag-DATE(Jahr,1,1)+1))", color="FFFFFF")
    ctx.name("DB_Monat", f"{S_DASH}!$Z$3")
    ctx.name("StIdx", f"{S_DASH}!$Z$4")
    note(ws, "B4", "Stichtag und Monat sind änderbar (gelb). Ohne Eingabe gilt heute bzw. der erste Arbeitstag des "
         "Planungsjahres.")
    G = "INDEX(Ber_Grid,0,StIdx)"
    avail = f"(({G}<2)+({G}>8))"   # einsatzbereit: nicht abwesend (Freistellung zaehlt als einsatzbereit)
    jb = ctx.jb
    jc = jb["cols"]
    JB = q(S_JAHR)
    jcol = lambda h: f"{JB}!${jc[h]}${jb['r_first']}:${jc[h]}${jb['r_last']}"
    jyear = lambda h: f"{JB}!${jc[h]}${jb['r_year']}"
    r1, r2 = 6, 10
    T = {}
    T["bestand"] = tile(ws, 2, r1, "Personalbestand (Teams)", "=Bestand",
                        '="+ "&COUNTIF(MA_Kat,"Ressource")&" Ressourcen · "&COUNTIF(MA_Kat,"Ohne Team")&" ohne Team"')
    T["verf"] = tile(ws, 5, r1, "Verfügbare MA (Stichtag)",
                     '=IF(INDEX(Bedarf_AT,StIdx)=1,INDEX(Bedarf_Ist,StIdx),"–")',
                     '=IF(INDEX(Bedarf_AT,StIdx)=1,"Soll "&SollTag&" · "&INDEX(tblBedarf[Ampel Tagdienst],StIdx),'
                     '"kein Arbeitstag")')
    T["urlaub"] = tile(ws, 8, r1, "Urlauber (Stichtag)", f"=SUMPRODUCT((({G}=3)+({G}=4))*(Ber_Typ=\"P\"))",
                       f'="davon eTime: "&COUNTIF({G},4)')
    T["krank"] = tile(ws, 11, r1, "Kranke Mitarbeiter (Stichtag)", f"=COUNTIF({G},8)",
                      f'="Unterricht/Einweisung: "&(COUNTIF({G},5)+COUNTIF({G},6)+COUNTIF({G},7))')
    # Teamtabelle (fuer Kacheln 5/6)
    t0 = 16
    put(ws, f"K{t0 - 1}", "Teams am Stichtag (einsatzfähig, falls in Bereitschaft)", bold=True, size=11)
    th = ["Team", "Bestand", "ohne Abw.", "Meister", "AV", "TVT", "Handw.", "Ampel", "heute"]
    tc = [gl(11 + i) for i in range(9)]
    for c, h in zip(tc, th):
        put(ws, f"{c}{t0}", h, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=8, wrap=True)
    hc = {"fM": "AA", "fAV": "AB", "fTVT": "AC", "fHW": "AD", "rM": "AE", "rAV": "AF", "rTVT": "AG", "rHW": "AH",
          "iM": "AI", "iAV": "AJ", "iTVT": "AK", "iHW": "AL", "st": "AM"}
    rk = {"Meister": "M", "AV": "AV", "TVT-Fahrer": "TVT", "Handwerker": "HW"}
    qn = {"Meister": "Ber_QMeister", "AV": "Ber_QAV", "TVT-Fahrer": "Ber_QTVT", "Handwerker": "Ber_QHW"}
    soll = {"Meister": "SollM", "AV": "SollAV", "TVT-Fahrer": "SollTVT", "Handwerker": "SollHW"}
    role_cols = tc[3:7]
    for i, t in enumerate(M.TEAMS):
        r = t0 + 1 + i
        put(ws, f"{tc[0]}{r}", t, bold=True, bg=M.TEAM_COLOR[t], color=M.TEAM_FONT[t], border=BOX, h="center", size=9)
        put(ws, f"{tc[1]}{r}", f'=COUNTIFS(Ber_Gruppe,"{t}",Ber_Typ,"P")', border=BOX, h="center", size=9)
        put(ws, f"{tc[2]}{r}", f'=SUMPRODUCT((Ber_Gruppe="{t}")*(Ber_Typ="P")*{avail})', border=BOX, h="center", size=9)
        for c, rolle in zip(role_cols, M.ROLLEN):
            put(ws, f"{c}{r}", f'=SUMPRODUCT((Ber_Gruppe="{t}")*(Ber_Funktion="{rolle}")*{avail})', border=BOX,
                h="center", size=9)
            k_ = rk[rolle]
            ws[f"{hc['f' + k_]}{r}"] = f"=MAX(0,{soll[rolle]}-{c}{r})"
            ws[f"{hc['r' + k_]}{r}"] = f"=MAX(0,{c}{r}-{soll[rolle]})"
        for rolle in M.ROLLEN:
            k_ = rk[rolle]
            terms = []
            for other in M.ROLLEN:
                if other != rolle:
                    cnt = f'SUMPRODUCT((Ber_Gruppe="{t}")*(Ber_Funktion="{other}")*{avail}*{qn[rolle]})'
                    terms.append(f"MIN({hc['r' + rk[other]]}{r},{cnt})")
            ws[f"{hc['i' + k_]}{r}"] = "=" + "+".join(terms)
        fs = "+".join(f"{hc['f' + rk[x]]}{r}" for x in M.ROLLEN)
        rs = "+".join(f"{hc['r' + rk[x]]}{r}" for x in M.ROLLEN)
        conds = ",".join(f"{hc['f' + rk[x]]}{r}<={hc['i' + rk[x]]}{r}" for x in M.ROLLEN)
        ws[f"{hc['st']}{r}"] = f"=IF({fs}=0,0,IF(AND({conds},{fs}<={rs}),1,2))"
        put(ws, f"{tc[7]}{r}", f"=CHOOSE({hc['st']}{r}+1,{AMPEL})", border=BOX, h="center", size=11)
        put(ws, f"{tc[8]}{r}", (f'=IF(INDEX(Schicht_Nacht,StIdx)="{t}","Bereitschaft",IF(ISNUMBER(SEARCH("{t}",'
                                f'INDEX(Schicht_FSwirksam,StIdx))),"freigestellt",IF(INDEX(Bedarf_AT,StIdx)=1,'
                                f'"Tagdienst","frei")))'), border=BOX, h="center", size=8)
    for c, rolle in zip(role_cols, M.ROLLEN):
        cf(ws, f"{c}{t0 + 1}:{c}{t0 + 5}", f"{c}{t0 + 1}<{soll[rolle]}", dxf(bg="FFC7CE", fg="9C0006", bold=True))
    hcol = tc[8]
    cf(ws, f"{hcol}{t0 + 1}:{hcol}{t0 + 5}", f'{hcol}{t0 + 1}="Bereitschaft"', dxf(bg="BFBFBF", bold=True))
    cf(ws, f"{hcol}{t0 + 1}:{hcol}{t0 + 5}", f'{hcol}{t0 + 1}="freigestellt"', dxf(fg="7F7F7F", italic=True))
    st = f"{hc['st']}{t0 + 1}:{hc['st']}{t0 + 5}"
    T["einsatz"] = tile(ws, 14, r1, "Einsatzfähige Teams (Stichtag)", f'=COUNTIF({st},"<2")&" / 5"',
                        f'="davon kritisch 🟡: "&COUNTIF({st},1)')
    crit = ",".join(f'IF({hc["st"]}{t0 + 1 + i}>0,"{t}","")' for i, t in enumerate(M.TEAMS))
    T["krit"] = tile(ws, 17, r1, "Kritische Teams (Stichtag)",
                     f'=IF(COUNTIF({st},">0")=0,"keine",_xlfn.TEXTJOIN(", ",TRUE,{crit}))',
                     '="🟡 kritisch · 🔴 unterbesetzt"', size=14)
    m = "DB_Monat"
    T["uq"] = tile(ws, 2, r2, "Urlaubsquote (Monat)", f"=INDEX({jcol('Urlaubsquote')},{m})",
                   f'="Jahr: "&ROUND({jyear("Urlaubsquote")}*100,1)&" %"', nf="0.0%")
    T["kq"] = tile(ws, 5, r2, "Krankenquote (Monat)", f"=INDEX({jcol('Krankenquote')},{m})",
                   f'="Jahr: "&ROUND({jyear("Krankenquote")}*100,1)&" %"', nf="0.0%")
    T["ma"] = tile(ws, 8, r2, "Monatsauslastung", f"=INDEX({jcol('Auslastung')},{m})",
                   f'="Ø Ist "&ROUND(INDEX({jcol("Ø Ist Tagdienst")},{m}),1)&" / Soll "&SollTag', nf="0%")
    T["ja"] = tile(ws, 11, r2, "Jahresauslastung", f"={jyear('Auslastung')}",
                   f'="Ø Ist "&ROUND({jyear("Ø Ist Tagdienst")},1)&" / Soll "&SollTag', nf="0%")
    T["bedarf"] = tile(ws, 14, r2, "Personalbedarf (fehlende Personentage im Monat)",
                       f"=INDEX({jcol('Fehlende Personentage')},{m})",
                       f'="Jahr: "&{jyear("Fehlende Personentage")}&" PT · Tage unterbesetzt: "&{jyear("Tage Unterbesetzung")}')
    T["ber"] = tile(ws, 17, r2, "Bereitschaft am Stichtag", "=INDEX(Schicht_Nacht,StIdx)",
                    "=INDEX(Einsatz_Ampel,StIdx)", size=18)
    for t in M.TEAMS:
        cf(ws, T["ber"], f'{T["ber"]}="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t]))
    cf(ws, T["verf"], f'AND(ISNUMBER({T["verf"]}),{T["verf"]}<SollTag)', dxf(bg="FFC7CE", fg="9C0006"))
    cf(ws, T["verf"], f'AND(ISNUMBER({T["verf"]}),{T["verf"]}<SollTag+Reserve)', dxf(bg="FFEB9C", fg="9C5700"))
    cf(ws, T["verf"], f'ISNUMBER({T["verf"]})', dxf(bg="C6EFCE", fg="006100"))
    cf(ws, T["krit"], f'{T["krit"]}<>"keine"', dxf(bg="FFEB9C", fg="9C5700"))

    # Ampeltabelle
    a0 = 15
    put(ws, f"B{a0}", "Ampelsystem", bold=True, size=11)
    for rng, h in (("B{r}:C{r}", "Bereich"), ("D{r}:E{r}", "Stichtag"), ("F{r}:G{r}", "Monat"), ("H{r}:I{r}", "Jahr")):
        mput(ws, rng.format(r=a0 + 1), h, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
    rows = [
        ("Tagdienst",
         "=INDEX(tblBedarf[Ampel Tagdienst],StIdx)",
         f'=IF(COUNTIFS(Bedarf_Monat,{m},Bedarf_AT,1,Bedarf_Unter,">0")>0,"🔴",IF(COUNTIFS(Bedarf_Monat,{m},Bedarf_AT,1,'
         f'Bedarf_Reserve,"<"&Reserve)>0,"🟡","🟢"))',
         '=IF(COUNTIFS(Bedarf_AT,1,Bedarf_Unter,">0")>0,"🔴",IF(COUNTIFS(Bedarf_AT,1,Bedarf_Reserve,"<"&Reserve)>0,'
         '"🟡","🟢"))'),
        ("Bereitschaft",
         f"=CHOOSE(INDEX(Einsatz_Stufe,StIdx)+1,{AMPEL})",
         f"=CHOOSE(_xlfn.MAXIFS(Bedarf_BerStufe,Bedarf_Monat,{m})+1,{AMPEL})",
         f"=CHOOSE(MAX(Bedarf_BerStufe)+1,{AMPEL})"),
    ]
    for i, (lab, a, b, c) in enumerate(rows):
        r = a0 + 2 + i
        mput(ws, f"B{r}:C{r}", lab, bold=True, border=BOX, size=9)
        mput(ws, f"D{r}:E{r}", a, border=BOX, h="center", size=14)
        mput(ws, f"F{r}:G{r}", b, border=BOX, h="center", size=14)
        mput(ws, f"H{r}:I{r}", c, border=BOX, h="center", size=14)
        ws.row_dimensions[r].height = 22
    ws.merge_cells(f"B{a0 + 4}:I{a0 + 5}")
    note(ws, f"B{a0 + 4}", "🟢 Sollbesetzung erreicht · 🟡 Kritische Besetzung · 🔴 Unterbesetzung "
         "(Monat/Jahr: schlechtester Tag)", wrap=True)

    # Diagrammdaten (versteckt): Tage des gewaehlten Monats und Teamwerte
    d0 = 50
    for c, h in (("AA", "Tag"), ("AB", "Ist Tagdienst"), ("AC", "Soll"), ("AD", "Abwesend")):
        put(ws, f"{c}{d0}", h, bold=True)
    for j in range(1, 32):
        r = d0 + j
        dd = f"DATE(Jahr,{m},{j})"
        idx = f"({dd}-DATE(Jahr,1,1)+1)"
        ok = f"{j}<=DAY(DATE(Jahr,{m}+1,0))"
        ws[f"AA{r}"] = f'=IF({ok},{j}&" "&INDEX(lstWT,WEEKDAY({dd},2)),"")'
        ws[f"AB{r}"] = f'=IF({ok},IF(INDEX(Bedarf_AT,{idx})=1,INDEX(Bedarf_Ist,{idx}),NA()),NA())'
        ws[f"AC{r}"] = f'=IF({ok},IF(INDEX(Bedarf_AT,{idx})=1,SollTag,NA()),NA())'
        ws[f"AD{r}"] = f'=IF({ok},IF(INDEX(Bedarf_AT,{idx})=1,INDEX(Bedarf_Abw,{idx}),NA()),NA())'
    put(ws, f"AF{d0}", "Team", bold=True)
    put(ws, f"AG{d0}", "Bereitschaftstage", bold=True)
    put(ws, f"AH{d0}", "Freistellungstage", bold=True)
    for i, t in enumerate(M.TEAMS):
        r = d0 + 1 + i
        ws[f"AF{r}"] = t
        ws[f"AG{r}"] = f'=COUNTIF(Schicht_Nacht,"{t}")'
        ws[f"AH{r}"] = f'=COUNTIF(Schicht_FSauto,"{t}")'
    for ci in range(26, 41):
        ws.column_dimensions[gl(ci)].hidden = True

    # Diagramme
    W, H = 15.2, 7.6
    c1 = BarChart()
    c1.type, c1.grouping, c1.overlap = "col", "stacked", 100
    c1.title = "Abwesenheitstage je Monat (Teams, Arbeitstage)"
    c1.y_axis.title = "Tage"
    cats = Reference(ctx.wb[S_JAHR], min_col=2, min_row=jb["r_first"], max_row=jb["r_last"])
    colors = {"Urlaubstage": "00B050", "Krankheitstage": "FFFF00", "Unterrichtstage": "FF0000", "JAZ-Tage": "C4D79B",
              "Einweisungen": "FFC000"}
    for h, colr in colors.items():
        ci = ctx.wb[S_JAHR][f"{jc[h]}1"].column
        c1.add_data(Reference(ctx.wb[S_JAHR], min_col=ci, min_row=jb["r_first"] - 1, max_row=jb["r_last"]),
                    titles_from_data=True)
        s_ = c1.series[-1]
        s_.graphicalProperties = GraphicalProperties(solidFill=colr)
        s_.graphicalProperties.line = LineProperties(solidFill="7F7F7F")
    c1.set_categories(cats)
    c1.height, c1.width = H, W
    c1.legend.position = "b"
    ws.add_chart(c1, "B24")

    c2 = LineChart()
    c2.title = "Ist- und Sollbesetzung Tagdienst – gewählter Monat"
    c2.y_axis.title = "MA"
    c2.add_data(Reference(ws, min_col=28, min_row=d0, max_col=30, max_row=d0 + 31), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=27, min_row=d0 + 1, max_row=d0 + 31))
    for s_, colr, dash in zip(c2.series, ["1F4E78", "C00000", "00B050"], [None, "dash", None]):
        s_.graphicalProperties.line.solidFill = colr
        s_.graphicalProperties.line.width = 22000
        if dash:
            s_.graphicalProperties.line.dashStyle = dash
        s_.smooth = False
    c2.height, c2.width = H, W
    c2.legend.position = "b"
    c2.visible_cells_only = False
    ws.add_chart(c2, "K24")

    c3 = BarChart()
    c3.type, c3.grouping, c3.overlap = "col", "stacked", 100
    c3.title = "Bereitschaft je Monat (Tage 🟢 / 🟡 / 🔴)"
    for h, colr in [("Bereitschaft 🟢 (Tage)", "70AD47"), ("Bereitschaft 🟡 (Tage)", "FFC000"),
                    ("Bereitschaft 🔴 (Tage)", "C00000")]:
        ci = ctx.wb[S_JAHR][f"{jc[h]}1"].column
        c3.add_data(Reference(ctx.wb[S_JAHR], min_col=ci, min_row=jb["r_first"] - 1, max_row=jb["r_last"]),
                    titles_from_data=True)
        c3.series[-1].graphicalProperties = GraphicalProperties(solidFill=colr)
    c3.set_categories(cats)
    c3.height, c3.width = H, W
    c3.legend.position = "b"
    ws.add_chart(c3, "B40")

    c4 = BarChart()
    c4.type = "bar"
    c4.title = "Bereitschafts- und Freistellungstage je Team"
    c4.add_data(Reference(ws, min_col=33, min_row=d0, max_col=34, max_row=d0 + 5), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=32, min_row=d0 + 1, max_row=d0 + 5))
    c4.series[0].graphicalProperties = GraphicalProperties(solidFill="BFBFBF")
    c4.series[1].graphicalProperties = GraphicalProperties(solidFill="9DC3E6")
    c4.height, c4.width = H, W
    c4.legend.position = "b"
    c4.visible_cells_only = False
    ws.add_chart(c4, "K40")

    # Ausblick 14 Tage
    o0 = 57
    put(ws, f"B{o0}", "Ausblick: 14 Tage ab Stichtag", bold=True, size=11)
    spec = [("B{r}:C{r}", "Datum"), ("D{r}:D{r}", "WT"), ("E{r}:F{r}", "Feiertag"), ("G{r}:H{r}", "Nachtteam"),
            ("I{r}:J{r}", "Freigestellt"), ("K{r}:K{r}", "Ist / Soll"), ("L{r}:L{r}", "Tagdienst"),
            ("M{r}:O{r}", "Bereitschaft"), ("P{r}:S{r}", "Hinweis")]
    for rng, h in spec:
        r_ = rng.format(r=o0 + 1)
        if r_.split(":")[0] != r_.split(":")[1]:
            mput(ws, r_, h, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
        else:
            put(ws, r_.split(":")[0], h, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
    for k in range(14):
        r = o0 + 2 + k
        idx = f"MIN(DATE(Jahr,12,31)-DATE(Jahr,1,1)+1,StIdx+{k})"
        vals = [("B{r}:C{r}", f"=INDEX(Einsatz_Datum,{idx})", dict(nf="DD.MM.YYYY", h="center")),
                ("D{r}:D{r}", f"=INDEX(lstWT,WEEKDAY(B{r},2))", dict(h="center")),
                ("E{r}:F{r}", f"=INDEX(Schicht_Feiertag,{idx})", dict(size=8)),
                ("G{r}:H{r}", f"=INDEX(Schicht_Nacht,{idx})", dict(h="center")),
                ("I{r}:J{r}", f"=INDEX(Schicht_FSwirksam,{idx})", dict(h="center")),
                ("K{r}:K{r}", f'=IF(INDEX(Bedarf_AT,{idx})=1,INDEX(Bedarf_Ist,{idx})&" / "&SollTag,"–")', dict(h="center")),
                ("L{r}:L{r}", f"=INDEX(tblBedarf[Ampel Tagdienst],{idx})", dict(h="center", size=10)),
                ("M{r}:O{r}", f"=INDEX(Einsatz_Ampel,{idx})", dict(size=8)),
                ("P{r}:S{r}", f"=INDEX(Einsatz_Fehler,{idx})", dict(size=8))]
        for rng, f_, kw in vals:
            r_ = rng.format(r=r)
            kw = dict(kw)
            kw.setdefault("size", 9)
            if r_.split(":")[0] != r_.split(":")[1]:
                mput(ws, r_, f_, border=BOX, **kw)
            else:
                put(ws, r_.split(":")[0], f_, border=BOX, **kw)
    for t in M.TEAMS:
        cf(ws, f"G{o0 + 2}:H{o0 + 15}", f'$G{o0 + 2}="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
        cf(ws, f"I{o0 + 2}:J{o0 + 15}", f'ISNUMBER(SEARCH("{t}",$I{o0 + 2}))', dxf(bg=M.TEAM_LIGHT[t]))
    cf(ws, f"B{o0 + 2}:F{o0 + 15}", f'$E{o0 + 2}<>""', dxf(bg=M.HOLIDAY_FILL, fg="CC0000"))
    cf(ws, f"B{o0 + 2}:F{o0 + 15}", f'WEEKDAY($B{o0 + 2},2)>=6', dxf(bg=M.WEEKEND_FILL))
    from .style import landscape
    landscape(ws, "A4", 1, 0, area=f"A1:S{o0 + 15}")
    return ws
