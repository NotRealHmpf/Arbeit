"""Planungskern: Hilfsberechnung, Schichtplan 2027, PEP 2027 (Personen x Tage),
Einsatzfaehigkeitspruefung mit Ersatzlogik, Ueberstunden und JAZ, Personalbedarfsplanung."""
from pp_base import (ABW_CODES, AMPEL, DATE_FMT, DUTY_CODES, FERIEN, FKT_KURZ, FUNKTIONEN, GREY, HEAD,
                     HEAD_DARK, HOLIDAY, HOLIDAY_FONT, INPUT_BG, MON3, MONATE, MUTED, NDAYS, NMAX,
                     SAT, SUN, TEAMS, THIN, U_NAMES, WEEKEND, L, ampel_cf, col_widths, fill, font,
                     header_row, landscape, put, style, team_cf, title, weekday_formula)

X = "_xlfn."
OK_ = "✅ Einsatzfähig"
NOK = "❌ Nicht einsatzfähig"

# ---------------------------------------------------------------- Layout-Konstanten
C_FIRST = 5                     # erste Personenzeile Hilfsberechnung
C_LAST = C_FIRST + NMAX - 1
C_GRID0 = 34                    # Spalte AH = Tag 1
C_MON0 = C_GRID0 + NDAYS + 1    # Monatsblock (Soll/Ist/JAZ je Monat)
PEP_FIRST = 16                  # erste Personenzeile PEP 2027
PEP_LAST = PEP_FIRST + NMAX - 1
PEP_D0 = 6                      # Spalte F = Tag 1
PLAN_FIRST = 5
PLAN_LAST = PLAN_FIRST + NDAYS - 1
EF_FIRST = 6
EF_LAST = EF_FIRST + NDAYS - 1
EF_TEAM0 = 15                   # Spalte O: erster Teamblock (6 Spalten je Team)
UE_FIRST = 6
UE_LAST = UE_FIRST + NMAX - 1
PB_FIRST = 6
PB_LAST = PB_FIRST + NDAYS - 1


def dcol(i, base):
    """Spaltenbuchstabe fuer Tag i (0-basiert)."""
    return L(base + i)


def pep_row(r):
    """Zeilenbereich einer Person im PEP-Raster."""
    return f"'PEP 2027'!$F{r}:${dcol(NDAYS - 1, PEP_D0)}{r}"


# ====================================================================== Hilfsberechnung
def build_calc(book, ws):
    title(ws, "Hilfsberechnung (technisch – nicht bearbeiten)",
          "Sortierte Mitarbeiterliste, Qualifikations-Gültigkeit je Funktion, Abwesenheits-Priorität je Tag "
          "(-1 = nicht aktiv, 0 = anwesend, >0 = Code-Prio), Soll/Ist-Stunden je Monat.")
    heads = ["i", "Schlüssel", "Stammzeile", "Name", "Team", "Teamfunktion", "Aktiv", "Eintritt", "Pers.-Nr."] + \
            FUNKTIONEN + U_NAMES + ["Gültig bis Meister", "Gültig bis AV", "Gültig bis TVT", "Gültig bis HW",
                                    "Nachtdienste", "Hat Vertretung"]
    for j, h in enumerate(heads):
        c = ws.cell(4, 1 + j, h)
        style(c, size=8, bold=True, bg=HEAD, wrap=True, h="center", border=THIN)
    for j in range(13):
        ws.cell(4, 14 + j).value = f"=INDEX(Unterrichte,{j + 1})"
    ws.row_dimensions[4].height = 40
    gl = dcol(NDAYS - 1, C_GRID0)
    put(ws, f"{L(C_GRID0 - 1)}2", "Tag-Nr", size=8, bold=True)
    put(ws, f"{L(C_GRID0 - 1)}3", "Datum", size=8, bold=True)
    for i in range(NDAYS):
        c = dcol(i, C_GRID0)
        ws[f"{c}2"] = i + 1
        ws[f"{c}3"] = f"=IF({c}2>TageJahr,\"\",Jahresbeginn+{c}2-1)"
        ws[f"{c}3"].number_format = "DD.MM."
        ws.column_dimensions[c].width = 6
    for m in range(12):
        for k, lab in enumerate(["Soll", "Ist", "JAZ"]):
            c = L(C_MON0 + 3 * m + k)
            ws[f"{c}3"] = m + 1
            ws[f"{c}4"] = lab
            style(ws[f"{c}4"], size=8, bold=True, bg=HEAD, h="center")
    put(ws, f"{L(C_MON0)}2", "Soll-/Ist-Stunden und JAZ-Tage je Monat", size=8, bold=True)
    req = ["Req_Meister", "Req_AV", "Req_TVT", "Req_HW"]
    for i in range(NMAX):
        r = C_FIRST + i
        pr = PEP_FIRST + i
        prow = pep_row(pr)
        ws[f"A{r}"] = i + 1
        ws[f"B{r}"] = f'=IFERROR(SMALL(MA_Sort,A{r}),"")'
        ws[f"C{r}"] = f'=IF(B{r}="","",MATCH(B{r},MA_Sort,0))'
        ws[f"D{r}"] = f'=IF(C{r}="","",INDEX(MA_Name,C{r})&"")'
        ws[f"E{r}"] = f'=IF(C{r}="","",INDEX(MA_Team,C{r})&"")'
        ws[f"F{r}"] = f'=IF(C{r}="","",INDEX(MA_TF,C{r})&"")'
        ws[f"G{r}"] = f'=IF(C{r}="","",INDEX(MA_Aktiv,C{r})&"")'
        ws[f"H{r}"] = f'=IF(C{r}="","",IF(INDEX(MA_Eintritt,C{r})="","",INDEX(MA_Eintritt,C{r})))'
        ws[f"I{r}"] = f'=IF(C{r}="","",INDEX(MA_PNr,C{r})&"")'
        for j, fm in enumerate(["FM_M", "FM_AV", "FM_TVT", "FM_HW"]):
            c = L(10 + j)
            ws[f"{c}{r}"] = (f'=IF($D{r}="","",IFERROR(INDEX({fm},MATCH($D{r},FM_Name,0))&"",'
                             f'IF($F{r}={c}$4,"JA","NEIN")))')
        for j in range(13):
            c = L(14 + j)
            mx = f'{X}MAXIFS(Q_PlanBis,Q_MA,$D{r},Q_Qual,{c}$4)'
            ws[f"{c}{r}"] = f'=IF($D{r}="","",IF({mx}=0,IF(StrengeQual="JA",0,2958465),{mx}))'
        for j in range(4):
            c = L(27 + j)
            flag = L(10 + j)
            ws[f"{c}{r}"] = (f'=IF($D{r}="","",IF({flag}{r}<>"JA",0,IF(COUNTIF({req[j]},"JA")=0,2958465,'
                             f'{X}MINIFS($N{r}:$Z{r},{req[j]},"JA"))))')
        ws[f"AE{r}"] = f'=IF($D{r}="",0,COUNTIF({prow},"N")+COUNTIF({prow},"V"))'
        ws[f"AF{r}"] = f'=IF($D{r}="",0,IF(COUNTIF(Vert_Vertreter,$D{r})>0,1,0))'
        for d in range(NDAYS):
            c = dcol(d, C_GRID0)
            ws[f"{c}{r}"] = (f'=IF($D{r}="","",IF({c}$3="","",IF(OR($G{r}="NEIN",AND($H{r}<>"",{c}$3<$H{r})),-1,'
                             f'{X}MAXIFS(Abw_Prio,Abw_MA,$D{r},Abw_Von,"<="&{c}$3,Abw_BisW,">="&{c}$3))))')
        for m in range(12):
            cs, ci, cj = (L(C_MON0 + 3 * m + k) for k in range(3))
            mm = f"{cs}$3"
            arb = f"COUNTIFS(PEP_Arb,1,PEP_Mon,{mm})-COUNTIFS({prow},\"–\",PEP_Arb,1,PEP_Mon,{mm})"
            ws[f"{cs}{r}"] = f'=IF($D{r}="","",SollStdTag*({arb}))'
            cnt = lambda code, a=False: (f'COUNTIFS({prow},"{code}",PEP_Arb,1,PEP_Mon,{mm})' if a
                                         else f'COUNTIFS({prow},"{code}",PEP_Mon,{mm})')
            ws[f"{ci}{r}"] = (f'=IF($D{r}="","",StdTag*{cnt("T")}+StdNacht*({cnt("N")}+{cnt("V")})'
                              f'+SollStdTag*({arb}-{cnt("T")}-{cnt("N", True)}-{cnt("V", True)}-{cnt("F", True)}))')
            ws[f"{cj}{r}"] = f'=IF($D{r}="","",{cnt("JAZ", True)})'
    rng = lambda c: f"{c}{C_FIRST}:{c}{C_LAST}"
    for n, c in [("C_Idx", "A"), ("C_Name", "D"), ("C_Team", "E"), ("C_Fkt", "F"), ("C_Aktiv", "G"),
                 ("C_PNr", "I"), ("C_FM", "J"), ("C_FAV", "K"), ("C_FT", "L"), ("C_FH", "M"),
                 ("C_QM", "AA"), ("C_QA", "AB"), ("C_QT", "AC"), ("C_QH", "AD"), ("C_N", "AE"),
                 ("C_Vert", "AF")]:
        book.name_ws(n, ws, rng(c))
    book.name_ws("MA_Liste", ws, rng("D"))
    book.name_ws("C_Grid", ws, f"{L(C_GRID0)}{C_FIRST}:{gl}{C_LAST}")
    ml = L(C_MON0 + 35)
    book.name_ws("C_MonNr", ws, f"{L(C_MON0)}3:{ml}3")
    book.name_ws("C_MonLab", ws, f"{L(C_MON0)}4:{ml}4")
    book.name_ws("C_MonAll", ws, f"{L(C_MON0)}{C_FIRST}:{ml}{C_LAST}")
    for c in range(1, 33):
        ws.column_dimensions[L(c)].width = 10
    ws.column_dimensions["D"].width = 18
    for r in range(C_FIRST, C_LAST + 1):
        for c in range(1, 33):
            ws.cell(r, c).font = font(8)
            if 14 <= c <= 30:
                ws.cell(r, c).number_format = DATE_FMT
    ws.freeze_panes = f"E{C_FIRST}"
    ws.sheet_state = "hidden"


# ====================================================================== Schichtplan 2027
PLAN_COLS = ["Datum", "Wochentag", "Kalenderwoche", "Feiertag", "Nachtteam", "Freigestelltes Team",
             "Tagdienstteams", "Bemerkung", "Blocktyp", "Block-Nr", "Arbeitstag", "Monat", "Quartal",
             "Schulferien", "Status Nachtteam", "Hinweis", "Tag-Nr", "Blockstart"]


def build_plan(book, ws):
    title(ws, "Blatt 5 – Schichtplan 2027 (Nachtdienst-Rotation, Freistellung, Tagdienst)",
          "Rotation ab Rotationsbeginn: Fr–So Rot · Mo–Do Gelb · Fr–So Grün · Mo–Do Orange · Fr–So Blau · Mo–Do Rot · "
          "Fr–So Gelb · Mo–Do Grün · Fr–So Orange · Mo–Do Blau – fortlaufend. Freistellung: nach Mo–Do Freitag, "
          "nach Fr–So Montag + Dienstag.")
    put(ws, "A3", "=\"Planungsjahr \"&Jahr&\" · Rotationsbeginn \"&TEXT(DAY(RotStart),\"00\")&\".\"&TEXT(MONTH(RotStart),\"00\")&\".\"&YEAR(RotStart)&\" · \"&RotCheck",
        size=9, bold=True)
    header_row(ws, 4, 1, PLAN_COLS, inputs=(7,))
    teams = [f"INDEX(Teams,{k})" for k in range(1, 6)]
    for i in range(NDAYS):
        r = PLAN_FIRST + i
        ws[f"Q{r}"] = i + 1
        ws[f"A{r}"] = f'=IF(Q{r}>TageJahr,"",Jahresbeginn+Q{r}-1)'
        ws[f"B{r}"] = f'=IF(A{r}="","",{weekday_formula(f"A{r}")})'
        ws[f"C{r}"] = f'=IF(A{r}="","",{X}ISOWEEKNUM(A{r}))'
        ws[f"D{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(FT_Name,MATCH(A{r},FT_Datum,0)),""))'
        ws[f"J{r}"] = (f'=IF(A{r}="","",IF(A{r}<RotStart,"",2*((A{r}-WEEKDAY(A{r},2))-(RotStart-WEEKDAY(RotStart,2)))/7'
                       f'+(WEEKDAY(A{r},2)>=5)-(WEEKDAY(RotStart,2)>=5)))')
        ws[f"I{r}"] = f'=IF(A{r}="","",IF(WEEKDAY(A{r},2)>=5,"Fr-So","Mo-Do"))'
        ws[f"E{r}"] = f'=IF(OR(A{r}="",J{r}=""),"",INDEX(RotTeams,MOD(J{r},10)+1))'
        ws[f"F{r}"] = (f'=IF(OR(A{r}="",J{r}=""),"",IF(AND(OR(WEEKDAY(A{r},2)=5,WEEKDAY(A{r},2)<=2),J{r}>=1),'
                       f'INDEX(RotTeams,MOD(J{r}-1,10)+1),""))')
        ws[f"K{r}"] = f'=IF(A{r}="","",IF(AND(WEEKDAY(A{r},2)<=5,D{r}=""),1,0))'
        parts = ",".join(f'IF(OR({t}=E{r},{t}=F{r}),"",{t})' for t in teams)
        ws[f"G{r}"] = (f'=IF(A{r}="","",IF(K{r}=0,IF(D{r}<>"","– (Feiertag)","– (Wochenende)"),'
                       f'{X}TEXTJOIN(", ",TRUE,{parts})))')
        ws[f"L{r}"] = f'=IF(A{r}="","",MONTH(A{r}))'
        ws[f"M{r}"] = f'=IF(A{r}="","",ROUNDUP(MONTH(A{r})/3,0))'
        cond = f"(SF_Von<=A{r})*(SF_Bis>=A{r})"
        ws[f"N{r}"] = (f'=IF(A{r}="","",IF(SUMPRODUCT({cond})=0,"",INDEX(SF_Name,SUMPRODUCT(MAX({cond}*ROW(SF_Von)))'
                       f'-ROW(INDEX(SF_Von,1))+1)))')
        ws[f"O{r}"] = f'=IF(A{r}="","",INDEX(EF_StatusNacht,Q{r}))'
        ws[f"P{r}"] = (f'=IF(A{r}="","",{X}TEXTJOIN(" | ",TRUE,IFERROR(INDEX(SO_Name,MATCH(A{r},SO_Datum,0)),""),'
                       f'IF(INDEX(EF_Fehlend,Q{r})="","","Nachtteam: "&INDEX(EF_Fehlend,Q{r})),'
                       f'IF(INDEX(EF_ErsatzText,Q{r})="","","Ersatz: "&INDEX(EF_ErsatzText,Q{r}))))')
        ws[f"R{r}"] = f'=IF(A{r}="",0,IF(OR(WEEKDAY(A{r},2)=1,WEEKDAY(A{r},2)=5,Q{r}=1),1,0))'
        for c in range(1, 19):
            cell = ws.cell(r, c)
            style(cell, size=9, h="center" if c not in (4, 7, 8, 14, 16) else None)
        ws[f"A{r}"].number_format = DATE_FMT
    book.name_ws("RotCheck", book.wb["Einstellungen"], "C14")
    book.table(ws, "tbl_Plan", f"A4:R{PLAN_LAST}")
    rng = lambda c: f"{c}{PLAN_FIRST}:{c}{PLAN_LAST}"
    for n, c in [("Plan_Datum", "A"), ("Plan_WT", "B"), ("Plan_KW", "C"), ("Plan_FT", "D"), ("Plan_Nacht", "E"),
                 ("Plan_Frei", "F"), ("Plan_TD", "G"), ("Plan_Typ", "I"), ("Plan_Block", "J"), ("Plan_Arb", "K"),
                 ("Plan_Monat", "L"), ("Plan_Quartal", "M"), ("Plan_SF", "N"), ("Plan_N", "Q"), ("Plan_BS", "R")]:
        book.name_ws(n, ws, rng(c))
    f = PLAN_FIRST
    team_cf(book, ws, rng("E"), f"E{f}")
    team_cf(book, ws, rng("F"), f"F{f}")
    book.cf(ws, f"A{f}:D{PLAN_LAST}", f'$D{f}<>""', bg=HOLIDAY, color=HOLIDAY_FONT, bold=True)
    book.cf(ws, f"A{f}:D{PLAN_LAST}", f'$B{f}="So"', bg=SUN, bold=True)
    book.cf(ws, f"A{f}:D{PLAN_LAST}", f'$B{f}="Sa"', bg=SAT, bold=True)
    book.cf(ws, rng("N"), f'N{f}<>""', bg=FERIEN)
    book.cf(ws, rng("O"), f'O{f}="{OK_}"', bg="C6EFCE", color="006100")
    book.cf(ws, rng("O"), f'O{f}="{NOK}"', bg="FFC7CE", color="9C0006", bold=True)
    col_widths(ws, {"A": 11, "B": 6, "C": 6, "D": 20, "E": 10, "F": 11, "G": 26, "H": 24, "I": 8, "J": 7,
                    "K": 8, "L": 7, "M": 7, "N": 20, "O": 20, "P": 60, "Q": 6, "R": 7})
    ws.freeze_panes = f"B{PLAN_FIRST}"
    landscape(ws, "A4", area=f"A1:P{PLAN_LAST}", rows="4:4")
    ws.sheet_properties.tabColor = "4F81BD"


# ====================================================================== PEP 2027
def build_pep(book, ws):
    lastc = dcol(NDAYS - 1, PEP_D0)
    put(ws, "A1", "PEP 2027 – Personaleinsatzplanung (automatisch aus Schichtplan, Abwesenheiten und Vertretungen)",
        size=14, bold=True)
    ws.row_dimensions[1].height = 30
    put(ws, "B2", "=\"PEP \"&Jahr", size=12, bold=True, h="center", bg=HEAD, border=THIN)
    labels = {3: "Tag", 4: "Name", 5: "KW", 6: "Nachtdienst (Team)", 7: "Freistellung (Team)"}
    for r, t in labels.items():
        put(ws, f"B{r}", t, size=9, bold=True, bg=HEAD, border=THIN)
    put(ws, "A4", "Nr", size=9, bold=True, bg=HEAD, border=THIN, h="center")
    put(ws, "C4", "Team", size=9, bold=True, bg=HEAD, border=THIN, h="center")
    put(ws, "D4", "Fkt.", size=9, bold=True, bg=HEAD, border=THIN, h="center")
    hidden = {8: "Tag-Nr", 9: "Arbeitstag", 10: "Monat", 11: "Fr-So-Block", 12: "Feiertag", 13: "Nachtteam",
              14: "Frei-Team", 15: "Blockstart"}
    for r, t in hidden.items():
        put(ws, f"B{r}", t, size=8, color=MUTED)
        ws.row_dimensions[r].hidden = True
    for i in range(NDAYS):
        c = dcol(i, PEP_D0)
        ws[f"{c}8"] = i + 1
        ws[f"{c}3"] = f'=IF({c}8>TageJahr,"",Jahresbeginn+{c}8-1)'
        ws[f"{c}2"] = f'=IF({c}3="","",IF(DAY({c}3)=1,CHOOSE(MONTH({c}3),{",".join(chr(34) + m + chr(34) for m in MON3)}),""))'
        ws[f"{c}4"] = f'=IF({c}3="","",{weekday_formula(f"{c}3")})'
        ws[f"{c}5"] = f'=IF({c}3="","",IF(OR(WEEKDAY({c}3,2)=1,{c}8=1),{X}ISOWEEKNUM({c}3),""))'
        ws[f"{c}13"] = f'=IF({c}3="","",INDEX(Plan_Nacht,{c}8))'
        ws[f"{c}14"] = f'=IF({c}3="","",INDEX(Plan_Frei,{c}8)&"")'
        ws[f"{c}6"] = f'=IF({c}13="","",IFERROR(INDEX(TeamKurz,MATCH({c}13,Teams,0)),""))'
        ws[f"{c}7"] = f'=IF({c}14="","",IFERROR(INDEX(TeamKurz,MATCH({c}14,Teams,0)),""))'
        ws[f"{c}9"] = f'=IF({c}3="",0,INDEX(Plan_Arb,{c}8))'
        ws[f"{c}10"] = f'=IF({c}3="",0,MONTH({c}3))'
        ws[f"{c}11"] = f'=IF({c}3="",0,IF(WEEKDAY({c}3,2)>=5,1,0))'
        ws[f"{c}12"] = f'=IF({c}3="",0,IF(INDEX(Plan_FT,{c}8)<>"",1,0))'
        ws[f"{c}15"] = f'=IF({c}3="",0,INDEX(Plan_BS,{c}8))'
        style(ws[f"{c}2"], size=9, bold=True, bg=HEAD)
        style(ws[f"{c}3"], size=9, bold=True, bg=HEAD, h="center", fmt="D", border=THIN)
        style(ws[f"{c}4"], size=8, bg=HEAD, h="center", border=THIN)
        style(ws[f"{c}5"], size=7, color=MUTED, h="center")
        style(ws[f"{c}6"], size=7, bold=True, h="center", border=THIN)
        style(ws[f"{c}7"], size=7, h="center", border=THIN, color="595959")
        for r in range(9, 16):
            ws[f"{c}{r}"].font = font(7, color=MUTED)
        ws.column_dimensions[c].width = 4.3
    for i in range(NMAX):
        r = PEP_FIRST + i
        cr = C_FIRST + i
        ws[f"A{r}"] = f'=IF(Hilfsberechnung!$D{cr}="","",Hilfsberechnung!$A{cr})'
        ws[f"B{r}"] = f'=Hilfsberechnung!$D{cr}'
        ws[f"C{r}"] = f'=Hilfsberechnung!$E{cr}'
        ws[f"D{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(FktKurz,MATCH(Hilfsberechnung!$F{cr},Teamfunktionen,0)),""))'
        style(ws[f"A{r}"], size=8, color=MUTED, h="center", border=THIN)
        style(ws[f"B{r}"], size=10, bold=True, border=THIN)
        style(ws[f"C{r}"], size=8, h="center", border=THIN)
        style(ws[f"D{r}"], size=8, h="center", border=THIN)
        for d in range(NDAYS):
            c = dcol(d, PEP_D0)
            g = f"Hilfsberechnung!{dcol(d, C_GRID0)}{cr}"
            vert = (f'IF(IF(Hilfsberechnung!$AF{cr}=1,COUNTIFS(Vert_Vertreter,$B{r},Vert_Von,"<="&{c}$3,'
                    f'Vert_BisW,">="&{c}$3),0)>0,"V",')
            ws[f"{c}{r}"] = (f'=IF($B{r}="","",IF({c}$3="","",IF({g}<0,"–",IF({g}>0,INDEX(AbwCodes,{g}),'
                             f'IF($C{r}={c}$13,"N",{vert}IF($C{r}={c}$14,"F",IF({c}$9=1,"T",""))))))))')
            cell = ws[f"{c}{r}"]
            style(cell, size=8, h="center", border=THIN)
    # Summenzeilen
    sums = [("Tagdienst anwesend (T)", ['"T"']), ("Nachtdienst / Bereitschaft (N+V)", ['"N"', '"V"']),
            ("Freistellung nach Nacht (F)", ['"F"']), ("Urlaub (U)", ['"U"']), ("Krank (K)", ['"K"']),
            ("JAZ", ['"JAZ"']), ("Unterricht (UR/FDL/KAB)", ['"UR"', '"FDL"', '"KAB"']),
            ("Sonstige (FB/SU/FS)", ['"FB"', '"SU"', '"FS"'])]
    s0 = PEP_LAST + 2
    for k, (lab, codes) in enumerate(sums):
        r = s0 + k
        put(ws, f"B{r}", lab, size=9, bold=True, bg=HEAD, border=THIN)
        for d in range(NDAYS):
            c = dcol(d, PEP_D0)
            rg = f"{c}${PEP_FIRST}:{c}${PEP_LAST}"
            ws[f"{c}{r}"] = f'=IF({c}$3="","",{"+".join(f"COUNTIF({rg},{x})" for x in codes)})'
            style(ws[f"{c}{r}"], size=8, h="center", border=THIN)
    rs = s0 + len(sums)
    put(ws, f"B{rs}", "Nachtteam einsatzfähig", size=9, bold=True, bg=HEAD, border=THIN)
    put(ws, f"B{rs + 1}", "Ampel Tagdienst", size=9, bold=True, bg=HEAD, border=THIN)
    for d in range(NDAYS):
        c = dcol(d, PEP_D0)
        ws[f"{c}{rs}"] = (f'=IF({c}$3="","",CHOOSE(MATCH(INDEX(EF_Ampel,{c}$8),{{"Grün","Gelb","Rot"}},0),'
                          f'"✔","!","✖"))')
        ws[f"{c}{rs + 1}"] = (f'=IF({c}$3="","",IF(INDEX(PB_Ampel,{c}$8)="","",CHOOSE(MATCH(INDEX(PB_Ampel,{c}$8),'
                              f'{{"Grün","Gelb","Rot"}},0),"✔","!","✖")))')
        for r in (rs, rs + 1):
            style(ws[f"{c}{r}"], size=8, bold=True, h="center", border=THIN)
    # Bedingte Formate
    grid = f"F{PEP_FIRST}:{lastc}{PEP_LAST}"
    f0 = f"F{PEP_FIRST}"
    for code, _t, bg, fc, b in DUTY_CODES:
        if bg or fc != "000000":
            book.cf(ws, grid, f'{f0}="{code}"', bg=bg, color=fc, bold=b or None, italic=(code == "V") or None,
                    stop=True)
    for code, _t, bg, fc, b in ABW_CODES:
        book.cf(ws, grid, f'{f0}="{code}"', bg=bg, color=fc, bold=b or None, stop=True)
    head = f"F3:{lastc}7 {grid}"
    book.cf(ws, head, "F$12=1", bg=HOLIDAY, color=HOLIDAY_FONT)
    book.cf(ws, head, "AND(F$3<>\"\",WEEKDAY(F$3,2)>=6)", bg=WEEKEND)
    book.cf(ws, f"F2:{lastc}{rs + 1}", 'AND(F$3<>"",DAY(F$3)=1)', border_left=True)
    for t, k, col in TEAMS:
        book.cf(ws, f"F6:{lastc}6", f'F$13="{t}"', bg=col, color="000000")
    book.cf(ws, f"F7:{lastc}7", 'F$14<>""', bg=GREY)
    team_cf(book, ws, f"C{PEP_FIRST}:C{PEP_LAST}", f"C{PEP_FIRST}")
    sym = f"F{rs}:{lastc}{rs + 1}"
    book.cf(ws, sym, f'F{rs}="✔"', bg="C6EFCE", color="006100")
    book.cf(ws, sym, f'F{rs}="!"', bg="FFEB9C", color="9C5700")
    book.cf(ws, sym, f'F{rs}="✖"', bg="FFC7CE", color="9C0006")
    # Legende wie im Original
    lg = rs + 3
    put(ws, f"B{lg}", "Legende", size=11, bold=True)
    items = [(c, t, bg, fc) for c, t, bg, fc, b in DUTY_CODES] + [(c, t, bg, fc) for c, t, bg, fc, b in ABW_CODES]
    items += [("", "Wochenende", WEEKEND, "000000"), ("", "Feiertag Hessen", HOLIDAY, HOLIDAY_FONT)]
    for k, (code, text, bg, fc) in enumerate(items):
        r = lg + 1 + k
        put(ws, f"D{r}", code, size=9, bold=True, h="center", bg=bg, color=fc, border=THIN)
        put(ws, f"B{r}", text, size=9)
    rng_row = lambda r: f"F{r}:{lastc}{r}"
    book.name_ws("PEP_Grid", ws, grid)
    book.name_ws("PEP_Names", ws, f"B{PEP_FIRST}:B{PEP_LAST}")
    book.name_ws("PEP_Datum", ws, rng_row(3))
    book.name_ws("PEP_Arb", ws, rng_row(9))
    book.name_ws("PEP_Mon", ws, rng_row(10))
    book.name_ws("PEP_FrSo", ws, rng_row(11))
    book.name_ws("PEP_FT", ws, rng_row(12))
    book.name_ws("PEP_BS", ws, rng_row(15))
    col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 5, "E": 1.7})
    for r in range(2, 8):
        ws.row_dimensions[r].height = 15
    ws.freeze_panes = f"F{PEP_FIRST}"
    landscape(ws, "A3", fit_width=False, fit_height=True, rows="2:7", cols="A:D")
    ws.sheet_properties.tabColor = "4F81BD"
    return rs


# ====================================================================== Einsatzfaehigkeit
def ef_team_cols(i):
    base = EF_TEAM0 + 6 * i
    return [L(base + k) for k in range(6)]


def build_ef(book, ws):
    title(ws, "Blatt 6 – Einsatzfähigkeitsprüfung je Team und Tag (mit automatischer Ersatzlogik)",
          "Soll je Team: 1 Meister · 1 AV · 1 TVT-Fahrer · 3 Handwerker (Einstellungen). Gezählt werden Mitarbeiter "
          "mit dieser Teamfunktion, die nicht abwesend sind (Urlaub, Krank, JAZ, Unterricht …) und deren "
          "Pflichtqualifikationen gültig sind, plus bestätigte Vertretungen.")
    put(ws, "A3", "Ersatz für das Nachtteam: P1 gleiches Team (Reserve/Überdeckung) · P2 anderes Team mit gleicher "
                  "Funktion · P3 anderes Team mit Mehrfachfunktion; jeweils mit gültiger Qualifikation, nicht freigestellt; "
                  "innerhalb der Priorität rotieren die Vorschläge blockweise gleichmäßig über alle Kandidaten. Ampel: Grün = einsatzfähig · Gelb = Lücke, aber Ersatz vorhanden · "
                  "Rot = Lücke ohne Ersatz.", size=8, color=MUTED, italic=True)
    hr = EF_FIRST - 1
    labels = ["Datum", "Wt", "KW", "Feiertag", "Nachtteam", "Status Nachtteam", "Fehlende Funktionen (Nachtteam)",
              "Ersatz Meister", "Ersatz AV", "Ersatz TVT", "Ersatz HW 1", "Ersatz HW 2", "Ersatz HW 3", "Ampel Nacht"]
    for i, (t, k, col) in enumerate(TEAMS):
        cols = ef_team_cols(i)
        labels += [f"{t}: Meister", f"{t}: AV", f"{t}: TVT", f"{t}: HW", f"{t}: Status", f"{t}: Fehlend"]
        put(ws, f"{cols[0]}4", f"=INDEX(Teams,{i + 1})", size=11, bold=True, h="center", bg=col, border=THIN)
        ws.merge_cells(f"{cols[0]}4:{cols[5]}4")
    SL = ["M", "AV", "TVT", "HW1", "HW2", "HW3"]
    labels += ["Tag-Nr", "Frei-Team", "Block-Nr", "Ist Meister (Nacht)", "Ist AV (Nacht)", "Ist TVT (Nacht)",
               "Ist HW (Nacht)", "Lücken ohne Ersatz"] + [f"Prio {x}" for x in SL] + [f"Kandidaten {x}" for x in SL] + \
              [f"Idx {x}" for x in SL] + ["Ersatz (Text)"]
    header_row(ws, hr, 1, labels, height=40, size=8)
    put(ws, "A4", "Nachtteam (Bereitschaft 15:30–06:45) und Ersatzvorschläge", size=11, bold=True, bg=HEAD)
    ws.merge_cells("A4:N4")
    # Hilfsspalten-Positionen
    HK = (["n", "frei", "blk", "iM", "iA", "iT", "iH", "luecke"] + [f"g{x}" for x in SL] + [f"c{x}" for x in SL] +
          [f"x{x}" for x in SL] + ["txt"])
    H = {n: L(45 + k) for k, n in enumerate(HK)}
    HL = 45 + len(HK) - 1
    fk_q = [("Meister", "C_QM", "SollM", "Meister fehlt"), ("AV", "C_QA", "SollA", "AV fehlt"),
            ("TVT-Fahrer", "C_QT", "SollT", "TVT fehlt"), ("Handwerker", "C_QH", "SollH", None)]
    for i in range(NDAYS):
        r = EF_FIRST + i
        n = f"${H['n']}{r}"
        ws[f"{H['n']}{r}"] = i + 1
        ws[f"A{r}"] = f'=IF({n}>TageJahr,"",INDEX(Plan_Datum,{n}))'
        ws[f"B{r}"] = f'=IF($A{r}="","",INDEX(Plan_WT,{n}))'
        ws[f"C{r}"] = f'=IF($A{r}="","",INDEX(Plan_KW,{n}))'
        ws[f"D{r}"] = f'=IF($A{r}="","",INDEX(Plan_FT,{n}))'
        ws[f"E{r}"] = f'=IF($A{r}="","",INDEX(Plan_Nacht,{n}))'
        ws[f"{H['frei']}{r}"] = f'=IF($A{r}="","",INDEX(Plan_Frei,{n})&"")'
        ws[f"{H['blk']}{r}"] = f'=IF($A{r}="",0,N(INDEX(Plan_Block,{n})))'
        avail = f"INDEX(C_Grid,0,{n}),0"
        status_cells, fehl_cells, ist_cells = [], [], []
        for t_i, (t, k, col) in enumerate(TEAMS):
            cols = ef_team_cols(t_i)
            tn = f"${cols[0]}$4"
            parts = []
            for j, (fk, qn, soll, msg) in enumerate(fk_q):
                c = cols[j]
                ws[f"{c}{r}"] = (f'=IF($A{r}="","",COUNTIFS(C_Team,{tn},C_Fkt,"{fk}",{avail},{qn},">="&$A{r})'
                                 f'+COUNTIFS(Vert_Team,{tn},Vert_Fkt,"{fk}",Vert_Von,"<="&$A{r},Vert_BisW,">="&$A{r}))')
                anw = f'COUNTIFS(C_Team,{tn},C_Fkt,"{fk}",{avail})'
                qok = f'COUNTIFS(C_Team,{tn},C_Fkt,"{fk}",{avail},{qn},">="&$A{r})'
                if msg:
                    parts.append(f'IF({c}{r}>={soll},"",IF({anw}>{qok},"Qualifikation abgelaufen ({fk})","{msg}"))')
                else:
                    parts.append(f'IF({c}{r}>={soll},"",IF({anw}>{qok},"Qualifikation abgelaufen (Handwerker), ","")'
                                 f'&"Zu wenig Handwerker ("&{c}{r}&"/"&{soll}&")")')
            st, fe = cols[4], cols[5]
            ws[f"{st}{r}"] = (f'=IF($A{r}="","",IF(AND({cols[0]}{r}>=SollM,{cols[1]}{r}>=SollA,{cols[2]}{r}>=SollT,'
                              f'{cols[3]}{r}>=SollH),"{OK_}","{NOK}"))')
            ws[f"{fe}{r}"] = f'=IF($A{r}="","",{X}TEXTJOIN(", ",TRUE,{",".join(parts)}))'
            status_cells.append(f"{st}{r}")
            fehl_cells.append(f"{fe}{r}")
            ist_cells.append(cols[:4])
            for c in cols:
                style(ws[f"{c}{r}"], size=8, h="center" if c not in (fe,) else None, border=THIN)
        tm = f'MATCH($E{r},Teams,0)'
        ws[f"F{r}"] = f'=IF($E{r}="","",CHOOSE({tm},{",".join(status_cells)}))'
        ws[f"G{r}"] = f'=IF($E{r}="","",CHOOSE({tm},{",".join(fehl_cells)}))'
        for j, key in enumerate(["iM", "iA", "iT", "iH"]):
            ws[f"{H[key]}{r}"] = f'=IF($E{r}="","",CHOOSE({tm},{",".join(c[j] + str(r) for c in ist_cells)}))'
        # Ersatz-Scores
        surplus = (f'((C_Fkt="Reserve")+(C_Fkt="Meister")*(${H["iM"]}{r}-SollM)+(C_Fkt="AV")*(${H["iA"]}{r}-SollA)'
                   f'+(C_Fkt="TVT-Fahrer")*(${H["iT"]}{r}-SollT)+(C_Fkt="Handwerker")*(${H["iH"]}{r}-SollH))')
        slots = [("M", "Meister", "C_QM", f'${H["iM"]}{r}>=SollM'),
                 ("AV", "AV", "C_QA", f'${H["iA"]}{r}>=SollA'),
                 ("TVT", "TVT-Fahrer", "C_QT", f'${H["iT"]}{r}>=SollT'),
                 ("HW1", "Handwerker", "C_QH", f'${H["iH"]}{r}>=SollH'),
                 ("HW2", "Handwerker", "C_QH", f'${H["iH"]}{r}>=SollH-1'),
                 ("HW3", "Handwerker", "C_QH", f'${H["iH"]}{r}>=SollH-2')]
        prev = []
        disp_cols = "HIJKLM"
        blk = f'${H["blk"]}{r}'
        for s_i, (sk, fk, qn, full) in enumerate(slots):
            gc, cc, xc = H[f"g{sk}"], H[f"c{sk}"], H[f"x{sk}"]
            excl = "*".join(f"(C_Idx<>${H[p]}{r})" for p in prev) or "1"
            g = (f'((C_Name<>"")*(INDEX(C_Grid,0,{n})=0)*({qn}>=$A{r})*{excl}*((C_Team=$E{r})*(C_Fkt<>"{fk}")*'
                 f'({surplus}>0)+(C_Team<>$E{r})*(C_Team<>${H["frei"]}{r})*((C_Fkt="{fk}")*2+(C_Fkt<>"{fk}")*3)))')
            # beste Prioritaetsgruppe, Anzahl Kandidaten darin, blockweise rotierende Auswahl
            ws[f"{gc}{r}"] = f'=IF(OR($E{r}="",{full}),"",SUMPRODUCT(MIN({g}+({g}=0)*9)))'
            ws[f"{cc}{r}"] = f'=IF(OR({gc}{r}="",{gc}{r}=9),0,SUMPRODUCT(({g}={gc}{r})*1))'
            ws[f"{xc}{r}"] = (f'=IF({gc}{r}="","",IF({gc}{r}=9,0,SUMPRODUCT(SMALL(({g}={gc}{r})*C_Idx'
                              f'+({g}<>{gc}{r})*1E+9,MOD({blk},{cc}{r})+1))))')
            dc = disp_cols[s_i]
            ws[f"{dc}{r}"] = (f'=IF({xc}{r}="","",IF({xc}{r}=0,"⚠ kein Ersatz",INDEX(C_Name,{xc}{r})&" ("&'
                              f'CHOOSE({gc}{r},"P1 Team","P2 anderes Team","P3 Mehrfachfunktion")&")"))')
            prev.append(f"x{sk}")
        ws[f"{H['luecke']}{r}"] = f'=IF($E{r}="","",COUNTIF(H{r}:M{r},"⚠*"))'
        ws[f"N{r}"] = f'=IF($E{r}="","",IF(F{r}="{OK_}","Grün",IF({H["luecke"]}{r}=0,"Gelb","Rot")))'
        lab = ["Meister", "AV", "TVT", "HW", "HW", "HW"]
        ws[f"{H['txt']}{r}"] = (f'=IF($E{r}="","",{X}TEXTJOIN("; ",TRUE,' +
                                ",".join(f'IF({dc}{r}="","","{lab[k]} → "&{dc}{r})' for k, dc in enumerate(disp_cols)) +
                                "))")
        for c in range(1, 15):
            style(ws.cell(r, c), size=8, h="center" if c not in (4, 7, 8, 9, 10, 11, 12, 13) else None, border=THIN)
        ws[f"A{r}"].number_format = DATE_FMT
        for c in range(45, HL + 1):
            ws.cell(r, c).font = font(7, color=MUTED)
    lastcol = H["txt"]
    book.table(ws, "tbl_EF", f"A{hr}:{lastcol}{EF_LAST}")
    rng = lambda c: f"{c}{EF_FIRST}:{c}{EF_LAST}"
    book.name_ws("EF_StatusNacht", ws, rng("F"))
    book.name_ws("EF_Fehlend", ws, rng("G"))
    book.name_ws("EF_Ampel", ws, rng("N"))
    book.name_ws("EF_Nacht", ws, rng("E"))
    book.name_ws("EF_Luecken", ws, rng(H["luecke"]))
    book.name_ws("EF_ErsatzText", ws, rng(H["txt"]))
    for k, key in enumerate(["iM", "iA", "iT", "iH"]):
        book.name_ws(["EF_IstM", "EF_IstA", "EF_IstT", "EF_IstH"][k], ws, rng(H[key]))
    for t_i in range(5):
        cols = ef_team_cols(t_i)
        book.name_ws(f"EF_St{t_i + 1}", ws, rng(cols[4]))
        book.name_ws(f"EF_Fe{t_i + 1}", ws, rng(cols[5]))
    f = EF_FIRST
    team_cf(book, ws, rng("E"), f"E{f}")
    all_status = f"F{f}:F{EF_LAST} " + " ".join(rng(ef_team_cols(t)[4]) for t in range(5))
    book.cf(ws, all_status, f'LEFT(F{f},1)="✅"', bg="C6EFCE", color="006100", bold=True)
    book.cf(ws, all_status, f'LEFT(F{f},1)="❌"', bg="FFC7CE", color="9C0006", bold=True)
    ampel_cf(book, ws, rng("N"), f"N{f}")
    book.cf(ws, f"H{f}:M{EF_LAST}", f'LEFT(H{f},1)="⚠"', bg="FFC7CE", color="9C0006", bold=True)
    book.cf(ws, f"H{f}:M{EF_LAST}", f'H{f}<>""', bg="FFEB9C")
    book.cf(ws, f"A{f}:D{EF_LAST}", f'$D{f}<>""', bg=HOLIDAY, color=HOLIDAY_FONT)
    book.cf(ws, f"A{f}:D{EF_LAST}", f'OR($B{f}="Sa",$B{f}="So")', bg=SAT)
    widths = {"A": 10, "B": 4, "C": 4, "D": 12, "E": 8, "F": 17, "G": 34, "N": 7}
    for c in "HIJKLM":
        widths[c] = 22
    col_widths(ws, widths)
    for t_i in range(5):
        cols = ef_team_cols(t_i)
        for c in cols[:4]:
            ws.column_dimensions[c].width = 5.5
        ws.column_dimensions[cols[4]].width = 17
        ws.column_dimensions[cols[5]].width = 30
    for c in range(45, HL + 1):
        ws.column_dimensions[L(c)].hidden = True
    ws.freeze_panes = f"F{EF_FIRST}"
    landscape(ws, "A3", area=f"A1:N{EF_LAST}", rows=f"{hr}:{hr}")
    ws.sheet_properties.tabColor = "4F81BD"


# ====================================================================== Ueberstunden und JAZ
def build_zeit(book, ws):
    title(ws, "Blatt 7 – Überstunden und JAZ (automatisch aus PEP 2027 und Zeitbuchungen)",
          "Soll = Arbeitstage × Sollstunden. Ist = Tagdienste × Std. Tagdienst + Nachtdienste × angerechnete Std. + "
          "Abwesenheiten an Arbeitstagen × Sollstunden. Freistellung nach Nacht = 0 Std. JAZ-Tage werden vom "
          "JAZ-Konto abgebucht (Parameter im Blatt Einstellungen).")
    labels = ["Name", "Team", "Vortrag Überstunden", "Soll-Std. Jahr", "Ist-Std. Jahr", "Überstunden",
              "Minderstunden", "Jahreswert", "Buchungen Überstunden", "Saldo Überstunden", "Vortrag JAZ",
              "JAZ-Tage", "JAZ-Abbau (Std.)", "Buchungen JAZ", "JAZ-Konto", "Saldo gesamt", "Monatswert Berichtsmonat"]
    labels += [f"Monat {m}" for m in MON3] + [f"Saldo Ü {m}" for m in MON3] + [f"JAZ-Konto {m}" for m in MON3]
    hr = UE_FIRST - 1
    header_row(ws, hr, 1, labels, height=42, size=8)
    groups = [(18, "Monatswerte (Ist − Soll, Std.)"), (30, "Saldo Überstunden kumuliert (Monatsende)"),
              (42, "JAZ-Konto (Monatsende)")]
    for c0, t in groups:
        put(ws, f"{L(c0)}4", t, size=9, bold=True, h="center", bg=HEAD, border=THIN)
        ws.merge_cells(f"{L(c0)}4:{L(c0 + 11)}4")
    put(ws, "C4", "Überstundenkonto", size=9, bold=True, h="center", bg=HEAD, border=THIN)
    ws.merge_cells("C4:J4")
    put(ws, "K4", "JAZ-Konto", size=9, bold=True, h="center", bg=HEAD, border=THIN)
    ws.merge_cells("K4:O4")
    lab_rng = f"Hilfsberechnung!${L(C_MON0)}$4:${L(C_MON0 + 35)}$4"
    for i in range(NMAX):
        r = UE_FIRST + i
        cr = C_FIRST + i
        crow = f"Hilfsberechnung!${L(C_MON0)}{cr}:${L(C_MON0 + 35)}{cr}"
        ws[f"A{r}"] = f"=Hilfsberechnung!$D{cr}"
        ws[f"B{r}"] = f"=Hilfsberechnung!$E{cr}"
        e = f'IF($A{r}="","",'
        ws[f"C{r}"] = f'={e}SUMIFS(ZB_Wert,ZB_MA,$A{r},ZB_Konto,"Überstunden",ZB_Art,"Vortrag Vorjahr"))'
        ws[f"D{r}"] = f'={e}SUMIF({lab_rng},"Soll",{crow}))'
        ws[f"E{r}"] = f'={e}SUMIF({lab_rng},"Ist",{crow}))'
        ws[f"F{r}"] = f'={e}SUMIF(R{r}:AC{r},">0"))'
        ws[f"G{r}"] = f'={e}-SUMIF(R{r}:AC{r},"<0"))'
        ws[f"H{r}"] = f'={e}SUM(R{r}:AC{r}))'
        ws[f"I{r}"] = f'={e}SUMIFS(ZB_Wert,ZB_MA,$A{r},ZB_Konto,"Überstunden")-C{r})'
        ws[f"J{r}"] = f'={e}C{r}+H{r}+I{r})'
        ws[f"K{r}"] = f'={e}SUMIFS(ZB_Wert,ZB_MA,$A{r},ZB_Konto,"JAZ",ZB_Art,"Vortrag Vorjahr"))'
        ws[f"L{r}"] = f'={e}SUMIF({lab_rng},"JAZ",{crow}))'
        ws[f"M{r}"] = f'={e}L{r}*SollStdTag)'
        ws[f"N{r}"] = f'={e}SUMIFS(ZB_Wert,ZB_MA,$A{r},ZB_Konto,"JAZ")-K{r})'
        ws[f"O{r}"] = f'={e}K{r}-M{r}+N{r})'
        ws[f"P{r}"] = f'={e}J{r}+O{r})'
        ws[f"Q{r}"] = f'={e}INDEX(R{r}:AC{r},Berichtsmonat))'
        for m in range(12):
            soll = f"Hilfsberechnung!{L(C_MON0 + 3 * m)}{cr}"
            ist = f"Hilfsberechnung!{L(C_MON0 + 3 * m + 1)}{cr}"
            ws[f"{L(18 + m)}{r}"] = f'={e}{ist}-{soll})'
            ws[f"{L(30 + m)}{r}"] = (f'={e}$C{r}+SUM($R{r}:{L(18 + m)}{r})+SUMIFS(ZB_Wert,ZB_MA,$A{r},ZB_Konto,'
                                     f'"Überstunden",ZB_Art,"<>Vortrag Vorjahr",ZB_Monat,"<="&{m + 1}))')
            ws[f"{L(42 + m)}{r}"] = (f'={e}$K{r}-SollStdTag*SUMIFS({crow},C_MonLab,"JAZ",C_MonNr,"<="&{m + 1})'
                                     f'+SUMIFS(ZB_Wert,ZB_MA,$A{r},ZB_Konto,"JAZ",ZB_Art,"<>Vortrag Vorjahr",'
                                     f'ZB_Monat,"<="&{m + 1}))')
        for c in range(1, 54):
            cell = ws.cell(r, c)
            style(cell, size=9, h="center" if c > 2 else None, fmt="0.00;-0.00;0.00" if c > 2 else None)
        ws[f"L{r}"].number_format = "0"
    last = L(53)
    book.table(ws, "tbl_Zeit", f"A{hr}:{last}{UE_LAST}")
    rng = lambda c: f"{c}{UE_FIRST}:{c}{UE_LAST}"
    for n, c in [("UE_Name", "A"), ("UE_Team", "B"), ("UE_Soll", "D"), ("UE_Ist", "E"), ("UE_Ue", "F"),
                 ("UE_Minder", "G"), ("UE_Jahr", "H"), ("UE_SaldoUe", "J"), ("UE_JAZTage", "L"),
                 ("UE_SaldoJAZ", "O"), ("UE_Gesamt", "P"), ("UE_MonatSel", "Q")]:
        book.name_ws(n, ws, rng(c))
    book.name_ws("UE_Monate", ws, f"R{UE_FIRST}:AC{UE_LAST}")
    f = UE_FIRST
    for c in ["H", "J", "O", "P", "Q"]:
        book.cf(ws, rng(c), f'AND(ISNUMBER({c}{f}),{c}{f}<0)', color="C00000", bold=True)
        book.cf(ws, rng(c), f'AND(ISNUMBER({c}{f}),{c}{f}>0)', color="006100", bold=True)
    mr = f"R{f}:AC{UE_LAST}"
    book.cf(ws, mr, f'AND(ISNUMBER(R{f}),R{f}<0)', bg="FFC7CE")
    book.cf(ws, mr, f'AND(ISNUMBER(R{f}),R{f}>0)', bg="C6EFCE")
    team_cf(book, ws, rng("B"), f"B{f}")
    # Statistik je Team
    s = UE_LAST + 3
    put(ws, f"A{s}", "Überstundenstatistik und JAZ-Statistik je Team", size=12, bold=True)
    header_row(ws, s + 1, 1, ["Team", "Mitarbeiter", "Soll-Std.", "Ist-Std.", "Überstunden", "Minderstunden",
                              "Jahreswert", "Saldo Überstunden", "JAZ-Tage", "JAZ-Konto", "Saldo gesamt",
                              "Ø Saldo je MA"], height=30, size=8)
    for k, (t, kk, col) in enumerate(TEAMS + [("Gesamt", "", "D9D9D9")]):
        r = s + 2 + k
        put(ws, f"A{r}", t, bold=True, bg=col, border=THIN)
        crit = f'UE_Team,"{t}"' if t != "Gesamt" else 'UE_Name,"?*"'
        put(ws, f"B{r}", f'=COUNTIFS({crit})', h="center", border=THIN)
        for j, n in enumerate(["UE_Soll", "UE_Ist", "UE_Ue", "UE_Minder", "UE_Jahr", "UE_SaldoUe", "UE_JAZTage",
                               "UE_SaldoJAZ", "UE_Gesamt"]):
            put(ws, f"{L(3 + j)}{r}", f'=SUMIFS({n},{crit})', h="center", border=THIN, fmt="0.00")
        put(ws, f"L{r}", f'=IF(B{r}=0,0,K{r}/B{r})', h="center", border=THIN, fmt="0.00")
    ms = s + 10
    put(ws, f"A{ms}", "Monatsstatistik (alle Mitarbeiter)", size=12, bold=True)
    header_row(ws, ms + 1, 1, ["Kennzahl"] + MONATE + ["Jahr"], height=20, size=8)
    rows = [("Σ Monatswert (Std.)", "SUM"), ("Σ Überstunden (Std.)", ">0"), ("Σ Minderstunden (Std.)", "<0"),
            ("Σ JAZ-Konto Monatsende", "JAZ")]
    for k, (lab, mode) in enumerate(rows):
        r = ms + 2 + k
        put(ws, f"A{r}", lab, bold=True, border=THIN)
        for m in range(12):
            col = L(18 + m)
            rg = f"{col}{UE_FIRST}:{col}{UE_LAST}"
            if mode == "SUM":
                fml = f"=SUM({rg})"
            elif mode == ">0":
                fml = f'=SUMIF({rg},">0")'
            elif mode == "<0":
                fml = f'=-SUMIF({rg},"<0")'
            else:
                jc = L(42 + m)
                fml = f"=SUM({jc}{UE_FIRST}:{jc}{UE_LAST})"
            put(ws, f"{L(2 + m)}{r}", fml, h="center", border=THIN, fmt="0.00")
        put(ws, f"N{r}", f"=SUM(B{r}:M{r})" if mode != "JAZ" else f"=M{r}", h="center", border=THIN,
            fmt="0.00", bold=True)
    book.name_ws("UE_MonSum", ws, f"B{ms + 2}:M{ms + 2}")
    col_widths(ws, {"A": 18, "B": 9})
    for c in range(3, 54):
        ws.column_dimensions[L(c)].width = 10
    ws.freeze_panes = f"C{UE_FIRST}"
    landscape(ws, "A3", area=f"A1:Q{ms + 6}", rows=f"{hr}:{hr}")
    ws.sheet_properties.tabColor = "4F81BD"


# ====================================================================== Personalbedarfsplanung
def build_bedarf(book, ws):
    title(ws, "Personalbedarfsplanung – Soll/Ist-Besetzung, Verfügbarkeit, Ausfallreserve",
          "Tagdienst: Plan = aktive Mitarbeiter der Tagdienstteams, Ist = tatsächlich im Tagdienst (T), Soll = "
          "Mindestbesetzung. Nacht: Soll = 6 Funktionen, Ist = besetzte Funktionen des Nachtteams. "
          "Ampel: Grün = Soll erreicht · Gelb = kritisch (Reserve < Puffer bzw. Lücke mit Ersatz) · Rot = Unterbesetzung.")
    labels = ["Datum", "Wt", "KW", "Monat", "Arbeitstag", "Nachtteam", "TD Plan", "TD Ist", "TD Soll",
              "Verfügbarkeit", "Ausfallreserve", "Unterbesetzung", "Überbesetzung", "Ampel Tagdienst",
              "Nacht Soll", "Nacht Ist", "Nacht Lücken", "Lücken ohne Ersatz", "Ampel Nacht", "Abwesend gesamt",
              "Aktive MA", "Tag-Nr"]
    hr = PB_FIRST - 1
    header_row(ws, hr, 1, labels, height=36, size=8)
    put(ws, "A4", "Tagesbasis", size=11, bold=True, bg=HEAD)
    ws.merge_cells("A4:V4")
    for i in range(NDAYS):
        r = PB_FIRST + i
        n = f"$V{r}"
        ws[f"V{r}"] = i + 1
        ws[f"A{r}"] = f'=IF({n}>TageJahr,"",INDEX(Plan_Datum,{n}))'
        ws[f"B{r}"] = f'=IF($A{r}="","",INDEX(Plan_WT,{n}))'
        ws[f"C{r}"] = f'=IF($A{r}="","",INDEX(Plan_KW,{n}))'
        ws[f"D{r}"] = f'=IF($A{r}="",0,INDEX(Plan_Monat,{n}))'
        ws[f"E{r}"] = f'=IF($A{r}="",0,INDEX(Plan_Arb,{n}))'
        ws[f"F{r}"] = f'=IF($A{r}="","",INDEX(Plan_Nacht,{n}))'
        ws[f"G{r}"] = (f'=IF(OR($A{r}="",$E{r}=0),0,COUNTIFS(INDEX(C_Grid,0,{n}),">=0",C_Team,"<>"&$F{r},'
                       f'C_Team,"<>"&INDEX(Plan_Frei,{n})))')
        ws[f"H{r}"] = f'=IF(OR($A{r}="",$E{r}=0),0,COUNTIF(INDEX(PEP_Grid,0,{n}),"T"))'
        ws[f"I{r}"] = f'=IF(OR($A{r}="",$E{r}=0),0,MinTag)'
        ws[f"J{r}"] = f'=IF(G{r}=0,"",H{r}/G{r})'
        ws[f"K{r}"] = f'=IF($E{r}<>1,"",MAX(0,H{r}-I{r}))'
        ws[f"L{r}"] = f'=IF($E{r}<>1,"",MAX(0,I{r}-H{r}))'
        ws[f"M{r}"] = f'=IF($E{r}<>1,"",MAX(0,H{r}-I{r}-PufferTag))'
        ws[f"N{r}"] = f'=IF($E{r}<>1,"",IF(H{r}<I{r},"Rot",IF(H{r}<I{r}+PufferTag,"Gelb","Grün")))'
        ws[f"O{r}"] = f'=IF($A{r}="","",SollNacht)'
        ws[f"P{r}"] = (f'=IF($A{r}="","",MIN(INDEX(EF_IstM,{n}),SollM)+MIN(INDEX(EF_IstA,{n}),SollA)'
                       f'+MIN(INDEX(EF_IstT,{n}),SollT)+MIN(INDEX(EF_IstH,{n}),SollH))')
        ws[f"Q{r}"] = f'=IF($A{r}="","",O{r}-P{r})'
        ws[f"R{r}"] = f'=IF($A{r}="","",INDEX(EF_Luecken,{n}))'
        ws[f"S{r}"] = f'=IF($A{r}="","",INDEX(EF_Ampel,{n}))'
        ws[f"T{r}"] = f'=IF($A{r}="","",COUNTIF(INDEX(C_Grid,0,{n}),">0"))'
        ws[f"U{r}"] = f'=IF($A{r}="","",COUNTIF(INDEX(C_Grid,0,{n}),">=0"))'
        for c in range(1, 23):
            style(ws.cell(r, c), size=8, h="center", border=THIN)
        ws[f"A{r}"].number_format = DATE_FMT
        ws[f"J{r}"].number_format = "0%"
        ws[f"V{r}"].font = font(7, color=MUTED)
    book.table(ws, "tbl_PB", f"A{hr}:V{PB_LAST}")
    rng = lambda c: f"{c}{PB_FIRST}:{c}{PB_LAST}"
    for nme, c in [("PB_Datum", "A"), ("PB_Monat", "D"), ("PB_Arb", "E"), ("PB_TDPlan", "G"), ("PB_TDIst", "H"),
                   ("PB_TDSoll", "I"), ("PB_Reserve", "K"), ("PB_Unter", "L"), ("PB_Ueber", "M"), ("PB_Ampel", "N"),
                   ("PB_NSoll", "O"), ("PB_NIst", "P"), ("PB_NLuecken", "Q"), ("PB_LoE", "R"), ("PB_AmpelN", "S"),
                   ("PB_Abw", "T"), ("PB_Aktiv", "U")]:
        book.name_ws(nme, ws, rng(c))
    f = PB_FIRST
    ampel_cf(book, ws, rng("N"), f"N{f}")
    ampel_cf(book, ws, rng("S"), f"S{f}")
    team_cf(book, ws, rng("F"), f"F{f}")
    book.cf(ws, f"A{f}:E{PB_LAST}", f'$E{f}=0', bg=LIGHT_ROW)

    # ---------------------------------------------------------------- rechte Seite: Aggregationen
    X0 = 24  # Spalte X
    xc = lambda k: L(X0 + k)
    put(ws, f"{xc(0)}4", "Jahresbasis", size=11, bold=True, bg=HEAD)
    ws.merge_cells(f"{xc(0)}4:{xc(15)}4")
    yl = ["Arbeitstage", "Plan-Personentage TD", "Ist-Personentage TD", "Verfügbarkeit", "Ø Ausfallreserve",
          "Unterbesetzung (PT)", "Überbesetzung (PT)", "TD Tage Grün", "TD Tage Gelb", "TD Tage Rot",
          "Nacht Tage Grün", "Nacht Tage Gelb", "Nacht Tage Rot", "Lücken ohne Ersatz", "Personalbedarf (strukturell)",
          "Ampel Jahr"]
    header_row(ws, 5, X0, yl, height=36, size=8)
    yv = ["=SUM(PB_Arb)", "=SUM(PB_TDPlan)", "=SUM(PB_TDIst)",
          f"=IF({xc(1)}6=0,0,{xc(2)}6/{xc(1)}6)", '=IFERROR(AVERAGEIFS(PB_Reserve,PB_Arb,1),0)',
          "=SUM(PB_Unter)", "=SUM(PB_Ueber)", '=COUNTIF(PB_Ampel,"Grün")', '=COUNTIF(PB_Ampel,"Gelb")',
          '=COUNTIF(PB_Ampel,"Rot")', '=COUNTIF(PB_AmpelN,"Grün")', '=COUNTIF(PB_AmpelN,"Gelb")',
          '=COUNTIF(PB_AmpelN,"Rot")', "=SUM(PB_LoE)", f"={xc(12)}15",
          f'=IF({xc(9)}6+{xc(12)}6>0,"Rot",IF({xc(8)}6+{xc(11)}6>0,"Gelb","Grün"))']
    for k, v in enumerate(yv):
        put(ws, f"{xc(k)}6", v, bold=True, h="center", border=THIN, size=10,
            fmt="0%" if k == 3 else ("0.0" if k == 4 else None))
    ampel_cf(book, ws, f"{xc(15)}6", f"{xc(15)}6")
    book.name_ws("PB_JahrAmpel", ws, f"{xc(15)}6")
    book.name_ws("PB_Bedarf", ws, f"{xc(14)}6")
    book.name_ws("PB_JahrVerf", ws, f"{xc(3)}6")

    # Strukturelle Besetzung
    put(ws, f"{xc(0)}8", "Strukturelle Besetzung je Team (Stammbesetzung laut Funktionsmatrix, nur aktive MA)",
        size=11, bold=True, bg=HEAD)
    ws.merge_cells(f"{xc(0)}8:{xc(13)}8")
    sl = ["Team", "Soll Meister", "Ist Meister", "Soll AV", "Ist AV", "Soll TVT", "Ist TVT", "Soll HW", "Ist HW",
          "Reserve", "Soll gesamt", "Ist gesamt", "Personalbedarf", "Ampel"]
    header_row(ws, 9, X0, sl, height=30, size=8)
    sollnames = ["SollM", "SollA", "SollT", "SollH"]
    for k, (t, kk, col) in enumerate(TEAMS):
        r = 10 + k
        put(ws, f"{xc(0)}{r}", t, bold=True, bg=col, border=THIN)
        for j, fk in enumerate(FUNKTIONEN):
            put(ws, f"{xc(1 + 2 * j)}{r}", f"={sollnames[j]}", h="center", border=THIN)
            put(ws, f"{xc(2 + 2 * j)}{r}", f'=COUNTIFS(C_Team,"{t}",C_Fkt,"{fk}",C_Aktiv,"<>NEIN")', h="center",
                border=THIN, bold=True)
        put(ws, f"{xc(9)}{r}", f'=COUNTIFS(C_Team,"{t}",C_Fkt,"Reserve",C_Aktiv,"<>NEIN")', h="center", border=THIN)
        put(ws, f"{xc(10)}{r}", "=SollNacht", h="center", border=THIN)
        put(ws, f"{xc(11)}{r}", f"={xc(2)}{r}+{xc(4)}{r}+{xc(6)}{r}+{xc(8)}{r}+{xc(9)}{r}", h="center", border=THIN)
        need = "+".join(f"MAX(0,{xc(1 + 2 * j)}{r}-{xc(2 + 2 * j)}{r})" for j in range(4))
        put(ws, f"{xc(12)}{r}", f"={need}", h="center", border=THIN, bold=True)
        put(ws, f"{xc(13)}{r}", f'=IF({xc(12)}{r}>0,"Rot",IF({xc(9)}{r}=0,"Gelb","Grün"))', h="center",
            border=THIN, bold=True)
    r = 15
    put(ws, f"{xc(0)}{r}", "Gesamt", bold=True, bg=GREY, border=THIN)
    for j in range(1, 13):
        put(ws, f"{xc(j)}{r}", f"=SUM({xc(j)}10:{xc(j)}14)", h="center", border=THIN, bold=True)
    put(ws, f"{xc(13)}{r}", f'=IF({xc(12)}15>0,"Rot",IF({xc(9)}15=0,"Gelb","Grün"))', h="center", border=THIN,
        bold=True)
    ampel_cf(book, ws, f"{xc(13)}10:{xc(13)}15", f"{xc(13)}10")
    book.name_ws("PB_StrTeam", ws, f"{xc(0)}10:{xc(0)}14")
    book.name_ws("PB_StrBedarf", ws, f"{xc(12)}10:{xc(12)}14")
    book.name_ws("PB_StrAmpel", ws, f"{xc(13)}10:{xc(13)}14")
    put(ws, f"{xc(0)}16", "Gelb = Soll erreicht, aber keine Reserve im Team (jeder Ausfall erfordert Ersatz). "
                          "Personalbedarf = fehlende Stammbesetzung je Funktion.", size=8, color=MUTED, italic=True)

    # Monatsbasis
    put(ws, f"{xc(0)}18", "Monatsbasis", size=11, bold=True, bg=HEAD)
    ws.merge_cells(f"{xc(0)}18:{xc(15)}18")
    ml = ["Monat", "Arbeitstage", "Plan-PT TD", "Ist-PT TD", "Verfügbarkeit", "Ø Ausfallreserve",
          "Unterbesetzung (PT)", "Überbesetzung (PT)", "TD Grün", "TD Gelb", "TD Rot", "Nacht Grün", "Nacht Gelb",
          "Nacht Rot", "Lücken o. Ersatz", "Ampel"]
    header_row(ws, 19, X0, ml, height=30, size=8)
    for m in range(13):
        r = 20 + m
        is_year = m == 12
        put(ws, f"{xc(0)}{r}", "Jahr" if is_year else MONATE[m], bold=True, border=THIN,
            bg=GREY if is_year else None)
        crit = "" if is_year else f",PB_Monat,{m + 1}"
        vals = [f"=COUNTIFS(PB_Arb,1{crit})", f"=SUMIFS(PB_TDPlan,PB_Arb,1{crit})",
                f"=SUMIFS(PB_TDIst,PB_Arb,1{crit})", f"=IF({xc(2)}{r}=0,0,{xc(3)}{r}/{xc(2)}{r})",
                f"=IFERROR(AVERAGEIFS(PB_Reserve,PB_Arb,1{crit}),0)", f"=SUMIFS(PB_Unter,PB_Arb,1{crit})",
                f"=SUMIFS(PB_Ueber,PB_Arb,1{crit})", f'=COUNTIFS(PB_Ampel,"Grün"{crit})',
                f'=COUNTIFS(PB_Ampel,"Gelb"{crit})', f'=COUNTIFS(PB_Ampel,"Rot"{crit})',
                f'=COUNTIFS(PB_AmpelN,"Grün"{crit})', f'=COUNTIFS(PB_AmpelN,"Gelb"{crit})',
                f'=COUNTIFS(PB_AmpelN,"Rot"{crit})', f"=SUMIFS(PB_LoE{crit})",
                f'=IF({xc(10)}{r}+{xc(13)}{r}>0,"Rot",IF({xc(9)}{r}+{xc(12)}{r}>0,"Gelb","Grün"))']
        if is_year:
            vals[13] = "=SUM(PB_LoE)"
        for k, v in enumerate(vals):
            put(ws, f"{xc(1 + k)}{r}", v, h="center", border=THIN, bold=is_year,
                fmt="0%" if k == 3 else ("0.0" if k == 4 else None))
    ampel_cf(book, ws, f"{xc(15)}20:{xc(15)}32", f"{xc(15)}20")
    book.name_ws("PB_MonAmpel", ws, f"{xc(15)}20:{xc(15)}31")
    book.name_ws("PB_MonVerf", ws, f"{xc(4)}20:{xc(4)}31")
    book.name_ws("PB_MonUnter", ws, f"{xc(6)}20:{xc(6)}31")
    book.name_ws("PB_MonNRot", ws, f"{xc(13)}20:{xc(13)}31")
    book.name_ws("PB_MonNGelb", ws, f"{xc(12)}20:{xc(12)}31")
    book.name_ws("PB_MonNGruen", ws, f"{xc(11)}20:{xc(11)}31")

    # Wochenbasis
    put(ws, f"{xc(0)}35", "Wochenbasis", size=11, bold=True, bg=HEAD)
    ws.merge_cells(f"{xc(0)}35:{xc(15)}35")
    wl = ["Wochenbeginn (Mo)", "KW", "Arbeitstage", "Plan-PT TD", "Ist-PT TD", "Verfügbarkeit",
          "Min. Ist TD", "Unterbesetzung (PT)", "TD Gelb", "TD Rot", "Nacht Gelb", "Nacht Rot", "Lücken o. Ersatz",
          "Nachtteams", "Ampel"]
    header_row(ws, 36, X0, wl, height=30, size=8)
    for w in range(53):
        r = 37 + w
        a = xc(0)
        put(ws, f"{a}{r}", f'=IF(Jahresbeginn-WEEKDAY(Jahresbeginn,2)+1+7*{w}>Jahresende,"",'
                           f'Jahresbeginn-WEEKDAY(Jahresbeginn,2)+1+7*{w})', h="center", border=THIN, fmt=DATE_FMT)
        crit = f'PB_Datum,">="&{a}{r},PB_Datum,"<="&({a}{r}+6)'
        vals = [f'=IF({a}{r}="","",{X}ISOWEEKNUM({a}{r}))', f'=IF({a}{r}="","",COUNTIFS(PB_Arb,1,{crit}))',
                f'=IF({a}{r}="","",SUMIFS(PB_TDPlan,{crit}))', f'=IF({a}{r}="","",SUMIFS(PB_TDIst,{crit}))',
                f'=IF({a}{r}="","",IF({xc(3)}{r}=0,"",{xc(4)}{r}/{xc(3)}{r}))',
                f'=IF({a}{r}="","",IF({xc(2)}{r}=0,"",{X}MINIFS(PB_TDIst,PB_Arb,1,{crit})))',
                f'=IF({a}{r}="","",SUMIFS(PB_Unter,PB_Arb,1,{crit}))',
                f'=IF({a}{r}="","",COUNTIFS(PB_Ampel,"Gelb",{crit}))',
                f'=IF({a}{r}="","",COUNTIFS(PB_Ampel,"Rot",{crit}))',
                f'=IF({a}{r}="","",COUNTIFS(PB_AmpelN,"Gelb",{crit}))',
                f'=IF({a}{r}="","",COUNTIFS(PB_AmpelN,"Rot",{crit}))',
                f'=IF({a}{r}="","",SUMIFS(PB_LoE,{crit}))',
                f'=IF({a}{r}="","",IFERROR(INDEX(Plan_Nacht,MAX(1,{a}{r}-Jahresbeginn+1))&" / "&'
                f'INDEX(Plan_Nacht,MIN(TageJahr,{a}{r}+4-Jahresbeginn+1)),""))',
                f'=IF({a}{r}="","",IF({xc(9)}{r}+{xc(11)}{r}>0,"Rot",IF({xc(8)}{r}+{xc(10)}{r}>0,"Gelb","Grün")))']
        for k, v in enumerate(vals):
            put(ws, f"{xc(1 + k)}{r}", v, h="center", border=THIN, size=9, fmt="0%" if k == 4 else None)
    ampel_cf(book, ws, f"{xc(14)}37:{xc(14)}89", f"{xc(14)}37")
    widths = {"A": 10, "B": 4, "C": 4, "D": 6, "E": 8, "F": 8, "N": 9, "S": 9}
    for c in "GHIJKLMOPQRTUV":
        widths[c] = 8
    col_widths(ws, widths)
    ws.column_dimensions["W"].width = 3
    for k in range(16):
        ws.column_dimensions[xc(k)].width = 11
    ws.column_dimensions[xc(0)].width = 14
    ws.freeze_panes = f"B{PB_FIRST}"
    landscape(ws, "A3", area=f"{xc(0)}1:{xc(15)}89")
    ws.sheet_properties.tabColor = "4F81BD"


LIGHT_ROW = "F2F2F2"
