"""Druckansichten fuer DIN A4/A3 quer: Jahresuebersicht, Monats- und Quartalsuebersicht, Teamuebersicht,
Bereitschaftsuebersicht (Layout des Bereitschaftskalenders) und Qualifikationsuebersicht."""
from openpyxl.workbook.defined_name import DefinedName

from pp_base import (ABW_CODES, AMPEL, DATE_FMT, DUTY_CODES, FERIEN, FUNKTIONEN, GREY, HEAD, HEAD_DARK, HOLIDAY, fill,
                     HOLIDAY_FONT, INPUT_BG, MONATE, MUTED, NMAX, SAT, SUN, TEAMS, THIN, U_NAMES, WEEKEND, L,
                     Border, Side, ampel_cf, col_widths, font, header_row, landscape, put, style, team_cf,
                     weekday_formula)
from pp_plan import C_FIRST, PEP_FIRST

X = "_xlfn."
SYM = 'CHOOSE(MATCH({a},{{"Grün","Gelb","Rot"}},0),"✔","!","✖")'


def dtxt(expr):
    """Datum als Text TT.MM.JJJJ ohne sprachabhaengige TEXT-Formatcodes."""
    return f'TEXT(DAY({expr}),"00")&"."&TEXT(MONTH({expr}),"00")&"."&YEAR({expr})'


def code_cf(book, ws, rng, first):
    for code, _t, bg, fc, b in DUTY_CODES + ABW_CODES:
        if bg or fc != "000000":
            book.cf(ws, rng, f'{first}="{code}"', bg=bg, color=fc, bold=b or None, stop=True)


def sym_cf(book, ws, rng, first):
    book.cf(ws, rng, f'{first}="✔"', bg="C6EFCE", color="006100")
    book.cf(ws, rng, f'{first}="!"', bg="FFEB9C", color="9C5700")
    book.cf(ws, rng, f'{first}="✖"', bg="FFC7CE", color="9C0006")


def dyn_print_area(wb, ws, rows_before, ncols):
    """Druckbereich passt sich der Anzahl Mitarbeiter an (OFFSET)."""
    idx = wb.sheetnames.index(ws.title)
    txt = (f"OFFSET('{ws.title}'!$A$1,0,0,{rows_before}+MAX(1,COUNTIF(PEP_Names,\"?*\")),{ncols})")
    ws.defined_names["_xlnm.Print_Area"] = DefinedName("_xlnm.Print_Area", localSheetId=idx, attr_text=txt)


# ====================================================================== Monat / Quartal (PEP-Ausschnitt)
def build_grid_print(book, wb, ws, mode):
    ndays = 31 if mode == "Monat" else 92
    D0 = 5                                   # Spalte E = erster Tag
    lastd = L(D0 + ndays - 1)
    cnt0 = D0 + ndays + 1                    # Zaehlspalten
    icol = L(cnt0 + 6)                       # Hilfsspalte i
    if mode == "Monat":
        put(ws, "A1", '="Monatsübersicht "&INDEX(MonName,$C$2)&" "&Jahr', size=16, bold=True)
        put(ws, "B2", "Monat:", bold=True, h="right")
        put(ws, "C2", "=MONTH(Stichtag)", bold=True, bg=INPUT_BG, h="center", border=THIN)
        book.dv_list(ws, "C2", "=MonNr")
        start = "DATE(Jahr,$C$2,1)"
        length = "DAY(EOMONTH(DATE(Jahr,$C$2,1),0))"
        paper = "A4"
    else:
        put(ws, "A1", '="Quartalsübersicht Q"&$C$2&" "&Jahr', size=16, bold=True)
        put(ws, "B2", "Quartal:", bold=True, h="right")
        put(ws, "C2", "=ROUNDUP(MONTH(Stichtag)/3,0)", bold=True, bg=INPUT_BG, h="center", border=THIN)
        book.dv_list(ws, "C2", "=QuNr")
        start = "DATE(Jahr,3*$C$2-2,1)"
        length = "(DATE(Jahr,3*$C$2+1,1)-DATE(Jahr,3*$C$2-2,1))"
        paper = "A3"
    ws.row_dimensions[1].height = 26
    # Legende in Zeile 2
    items = [(c, bg, fc) for c, t, bg, fc, b in DUTY_CODES + ABW_CODES if c != "–"]
    for k, (code, bg, fc) in enumerate(items):
        put(ws, f"{L(D0 + 2 * k)}2", code, size=7, bold=True, h="center", bg=bg, color=fc, border=THIN)
    put(ws, f"{L(D0 + 2 * len(items))}2", f"Legende: Blatt Start · optimiert für DIN {paper} quer (anderes Format: "
                                           f"Seite einrichten → Papierformat; Skalierung passt sich an)", size=7,
        color=MUTED)
    labels = {3: "Monat", 4: "Tag", 5: "Wochentag", 6: "KW", 7: "Nachtdienst (Team)",
              10: "Tagdienst anwesend", 11: "Abwesend", 12: "Nachtteam einsatzfähig"}
    for r, t in labels.items():
        put(ws, f"B{r}", t, size=8, bold=True, bg=HEAD, border=THIN)
    for r in (8, 9):
        ws.row_dimensions[r].hidden = True
    put(ws, "A13", "Nr", size=8, bold=True, bg=HEAD, border=THIN, h="center")
    put(ws, "B13", "Name", size=8, bold=True, bg=HEAD, border=THIN)
    put(ws, "C13", "Team", size=8, bold=True, bg=HEAD, border=THIN, h="center")
    put(ws, "D13", "Fkt.", size=8, bold=True, bg=HEAD, border=THIN, h="center")
    first, last = 14, 14 + NMAX - 1
    for j in range(ndays):
        c = L(D0 + j)
        ws[f"{c}4"] = f'=IF({j + 1}>{length},"",{start}+{j})'
        ws[f"{c}8"] = f'=IF({c}4="","",{c}4-Jahresbeginn+1)'
        ws[f"{c}9"] = f'=IF({c}8="","",INDEX(Plan_Nacht,{c}8))'
        ws[f"{c}3"] = (f'=IF({c}4="","",IF(DAY({c}4)=1,INDEX(MonName,MONTH({c}4)),""))' if mode != "Monat" else "")
        ws[f"{c}5"] = f'=IF({c}4="","",{weekday_formula(f"{c}4")})'
        ws[f"{c}6"] = f'=IF({c}4="","",IF(OR(WEEKDAY({c}4,2)=1,{j}=0),{X}ISOWEEKNUM({c}4),""))'
        ws[f"{c}7"] = f'=IF({c}9="","",IFERROR(INDEX(TeamKurz,MATCH({c}9,Teams,0)),""))'
        ws[f"{c}10"] = f'=IF({c}8="","",COUNTIF({c}${first}:{c}${last},"T"))'
        ws[f"{c}11"] = f'=IF({c}8="","",SUMPRODUCT(COUNTIF({c}${first}:{c}${last},AbwCodes)))'
        ws[f"{c}12"] = f'=IF({c}8="","",{SYM.format(a=f"INDEX(EF_Ampel,{c}8)")})'
        style(ws[f"{c}3"], size=8, bold=True)
        style(ws[f"{c}4"], size=8, bold=True, bg=HEAD, h="center", fmt="D", border=THIN)
        for r in (5, 6, 7, 10, 11, 12):
            style(ws[f"{c}{r}"], size=7, h="center", border=THIN, bold=r in (7, 12))
        for i in range(NMAX):
            r = first + i
            ws[f"{c}{r}"] = f'=IF(OR($B{r}="",{c}$8=""),"",INDEX(PEP_Grid,${icol}{r},{c}$8)&"")'
            style(ws[f"{c}{r}"], size=7, h="center", border=THIN)
        ws.column_dimensions[c].width = 3.9 if mode == "Monat" else 3.4
    cl = ["Tag", "Nacht", "Urlaub", "Krank", "Sonst."]
    for k, lab in enumerate(cl):
        put(ws, f"{L(cnt0 + k)}13", lab, size=7, bold=True, bg=HEAD, border=THIN, h="center")
        ws.column_dimensions[L(cnt0 + k)].width = 5.5
    ws.column_dimensions[L(cnt0 - 1)].width = 1.5
    rg = lambda r: f"$E{r}:${lastd}{r}"
    for i in range(NMAX):
        r = first + i
        pr = PEP_FIRST + i
        ws[f"{icol}{r}"] = i + 1
        put(ws, f"A{r}", f"=IF('PEP 2027'!$B{pr}=\"\",\"\",{i + 1})", size=7, color=MUTED, h="center", border=THIN)
        put(ws, f"B{r}", f"='PEP 2027'!$B{pr}", size=8, bold=True, border=THIN)
        put(ws, f"C{r}", f"='PEP 2027'!$C{pr}", size=7, h="center", border=THIN)
        put(ws, f"D{r}", f"='PEP 2027'!$D{pr}", size=7, h="center", border=THIN)
        vals = [f'COUNTIF({rg(r)},"T")', f'COUNTIF({rg(r)},"N")+COUNTIF({rg(r)},"V")', f'COUNTIF({rg(r)},"U")',
                f'COUNTIF({rg(r)},"K")',
                "+".join(f'COUNTIF({rg(r)},"{c}")' for c in ["FS", "SU", "JAZ", "FB", "UR", "FDL", "KAB"])]
        for k, v in enumerate(vals):
            put(ws, f"{L(cnt0 + k)}{r}", f'=IF($B{r}="","",{v})', size=7, h="center", border=THIN)
    ws.column_dimensions[icol].hidden = True
    grid = f"E{first}:{lastd}{last}"
    code_cf(book, ws, grid, f"E{first}")
    hd = f"E4:{lastd}7 {grid}"
    book.cf(ws, hd, 'AND(E$8<>"",INDEX(Plan_FT,N(E$8)+(E$8=""))<>"")', bg=HOLIDAY, color=HOLIDAY_FONT)
    book.cf(ws, hd, 'AND(E$4<>"",WEEKDAY(E$4,2)>=6)', bg=WEEKEND)
    for t, k, col in TEAMS:
        book.cf(ws, f"E7:{lastd}7", f'E$9="{t}"', bg=col)
    sym_cf(book, ws, f"E12:{lastd}12", "E12")
    book.cf(ws, f"E3:{lastd}{last}", 'AND(E$4<>"",DAY(E$4)=1)', border_left=True)
    team_cf(book, ws, f"C{first}:C{last}", f"C{first}")
    col_widths(ws, {"A": 3.5, "B": 16, "C": 6.5, "D": 4.5})
    ws.freeze_panes = f"E{first}"
    landscape(ws, paper, rows="3:13")
    dyn_print_area(wb, ws, 13, cnt0 + 4)
    if mode == "Monat":
        ws.row_dimensions[3].hidden = True
    ws.sheet_properties.tabColor = "7F7F7F"


# ====================================================================== Jahresuebersicht
def build_jahr(book, ws):
    put(ws, "A1", '="Jahresübersicht "&Jahr&" – Nachtdienst-Rotation, Abwesenheiten und Einsatzfähigkeit"',
        size=16, bold=True)
    ws.row_dimensions[1].height = 26
    put(ws, "A2", '="Nachtdienste: "&COUNTIF(EF_Ampel,"Grün")&" Tage einsatzfähig · "&COUNTIF(EF_Ampel,"Gelb")&'
                  '" Tage Lücke mit Ersatz · "&COUNTIF(EF_Ampel,"Rot")&" Tage Lücke ohne Ersatz · Urlaubstage "&'
                  'SUM(ST_Urlaub)&" · Krankheitstage "&SUM(ST_Krank)', size=10, color="595959")
    put(ws, "A3", "Je Monat: Tag (Farbe = Nachtteam) · Wochentag · Nachtteam · Anzahl Abwesende · Status Nachtteam "
                  "(✔ einsatzfähig, ! Lücke mit Ersatz, ✖ Lücke ohne Ersatz). Feiertage rot, Wochenende farbig, "
                  "Schulferien grau.", size=8, color=MUTED, italic=True)
    for m in range(12):
        c0 = 1 + 5 * m
        cs = [L(c0 + k) for k in range(5)]
        put(ws, f"{cs[0]}4", MONATE[m], size=11, bold=True, h="center", bg=HEAD, border=THIN)
        ws.merge_cells(f"{cs[0]}4:{cs[4]}4")
        for k, lab in enumerate(["Tag", "Wt", "Team", "Abw.", "St."]):
            put(ws, f"{cs[k]}5", lab, size=7, bold=True, h="center", bg=HEAD, border=THIN)
        for d in range(31):
            r = 6 + d
            a = f"{cs[0]}{r}"
            ws[a] = f'=IF({d + 1}>DAY(EOMONTH(DATE(Jahr,{m + 1},1),0)),"",DATE(Jahr,{m + 1},{d + 1}))'
            n = f"({a}-Jahresbeginn+1)"
            ws[f"{cs[1]}{r}"] = f'=IF({a}="","",{weekday_formula(a)})'
            ws[f"{cs[2]}{r}"] = f'=IF({a}="","",IFERROR(INDEX(TeamKurz,MATCH(INDEX(Plan_Nacht,{n}),Teams,0)),""))'
            ws[f"{cs[3]}{r}"] = f'=IF({a}="","",COUNTIF(INDEX(C_Grid,0,{n}),">0"))'
            ws[f"{cs[4]}{r}"] = f'=IF({a}="","",{SYM.format(a=f"INDEX(EF_Ampel,{n})")})'
            style(ws[a], size=8, bold=True, h="center", fmt="D", border=THIN)
            for k in range(1, 5):
                style(ws[f"{cs[k]}{r}"], size=7, h="center", border=THIN)
        rng = f"{cs[0]}6:{cs[0]}36"
        for t, k, col in TEAMS:
            book.cf(ws, rng, f'AND({cs[0]}6<>"",INDEX(Plan_Nacht,{cs[0]}6-Jahresbeginn+1)="{t}")', bg=col)
        wd = f"{cs[1]}6:{cs[2]}36"
        book.cf(ws, wd, f'AND(${cs[0]}6<>"",COUNTIF(FT_Datum,${cs[0]}6)>0)', bg=HOLIDAY, color=HOLIDAY_FONT, bold=True)
        book.cf(ws, wd, f'AND(${cs[0]}6<>"",WEEKDAY(${cs[0]}6,2)=7)', bg=SUN, bold=True)
        book.cf(ws, wd, f'AND(${cs[0]}6<>"",WEEKDAY(${cs[0]}6,2)=6)', bg=SAT, bold=True)
        book.cf(ws, wd, f'AND(${cs[0]}6<>"",SUMPRODUCT((SF_Von<=${cs[0]}6)*(SF_Bis>=${cs[0]}6))>0)', bg=FERIEN)
        sym_cf(book, ws, f"{cs[4]}6:{cs[4]}36", f"{cs[4]}6")
        book.cf(ws, f"{cs[3]}6:{cs[3]}36", f'AND(ISNUMBER({cs[3]}6),{cs[3]}6>=5)', bg="FCE4D6", bold=True)
        widths = [4, 3.6, 4, 3.6, 3.2]
        for k in range(5):
            ws.column_dimensions[cs[k]].width = widths[k]
    for r in range(6, 37):
        ws.row_dimensions[r].height = 15
    # Legende
    put(ws, "A38", "Teams:", size=8, bold=True)
    for k, (t, kk, col) in enumerate(TEAMS):
        put(ws, f"{L(3 + 3 * k)}38", t, size=8, bold=True, bg=col, h="center", border=THIN)
        ws.merge_cells(f"{L(3 + 3 * k)}38:{L(4 + 3 * k)}38")
    landscape(ws, "A3", fit_height=True, area="A1:BH38")
    ws.sheet_properties.tabColor = "7F7F7F"


# ====================================================================== Bereitschaftsuebersicht (Original-Layout)
def build_bereitschaft(book, ws, org_text):
    widths = [3.3, 3.7, 5.7, 2.7]
    for c in range(48):
        ws.column_dimensions[L(c + 1)].width = widths[c % 4]
    for r in range(1, 9):
        ws.row_dimensions[r].height = 12
    ws.row_dimensions[9].height = 27
    for r in range(10, 41):
        ws.row_dimensions[r].height = 21
    med = Side(style="medium", color="000000")
    thin = Side(style="thin", color="000000")
    put(ws, "C2", "=Jahr", size=24, bold=True, h="center")
    ws.merge_cells("C2:F3")
    put(ws, "Y2", org_text, size=10, bold=True, h="center")
    ws.merge_cells("Y2:AR2")
    for c in range(25, 45):
        ws.cell(2, c).border = Border(top=med, bottom=med, left=med if c == 25 else None, right=med if c == 44 else None)
    rows = [("Meister", "Meister", 1), ("AV", "AV", 1), ("TVT Fahrer", "TVT-Fahrer", 1),
            ("Handwerker 1", "Handwerker", 1), ("Handwerker 2", "Handwerker", 2), ("Handwerker 3", "Handwerker", 3)]
    team_cols = ["Y", "AC", "AG", "AK", "AO"]
    for i, (lab, fk, k) in enumerate(rows):
        r = 3 + i
        put(ws, f"U{r}", lab, size=10, bold=True, h="center")
        ws.merge_cells(f"U{r}:X{r}")
        for c in range(21, 25):
            ws.cell(r, c).border = Border(left=med if c == 21 else None, right=med if c == 24 else None,
                                          top=med if i == 0 else thin, bottom=med if i == 5 else thin)
        for t_i, (t, kk, col) in enumerate(TEAMS):
            a = team_cols[t_i]
            f = (f'=IFERROR(INDEX(C_Name,SUMPRODUCT(SMALL((C_Team="{t}")*(C_Fkt="{fk}")*C_Idx'
                 f'+((C_Team<>"{t}")+(C_Fkt<>"{fk}")>0)*999,{k}))),"")')
            put(ws, f"{a}{r}", f, size=10, bold=True, bg=col)
            c0 = ws[f"{a}{r}"].column
            ws.merge_cells(f"{a}{r}:{L(c0 + 3)}{r}")
            for c in range(c0, c0 + 4):
                ws.cell(r, c).border = Border(left=med if c == c0 else None, right=med if c == c0 + 3 else None,
                                              top=med if i == 0 else thin, bottom=med if i == 5 else thin)
                ws.cell(r, c).fill = fill(col)
    for m in range(12):
        c0 = 1 + 4 * m
        a, b, cc, dd = (L(c0 + k) for k in range(4))
        put(ws, f"{a}9", MONATE[m], size=14, bold=True, h="center", border=THIN)
        ws.merge_cells(f"{a}9:{dd}9")
        for k in range(4):
            ws.cell(9, c0 + k).border = THIN
        for d in range(31):
            r = 10 + d
            ws[f"{a}{r}"] = f'=IF({d + 1}>DAY(EOMONTH(DATE(Jahr,{m + 1},1),0)),"",DATE(Jahr,{m + 1},{d + 1}))'
            ws[f"{b}{r}"] = f'=IF({a}{r}="","",{weekday_formula(f"{a}{r}")})'
            ws[f"{cc}{r}"] = (f'=IF({a}{r}="","",IFERROR(INDEX(FT_Kurz,MATCH({a}{r},FT_Datum,0)),'
                              f'IFERROR(INDEX(SO_Name,MATCH({a}{r},SO_Datum,0)),"")))')
            ws[f"{dd}{r}"] = f'=IF({a}{r}="","",IF(WEEKDAY({a}{r},2)=1,{X}ISOWEEKNUM({a}{r}),""))'
            style(ws[f"{a}{r}"], size=11, h="center", fmt="D")
            style(ws[f"{b}{r}"], size=11)
            style(ws[f"{cc}{r}"], size=7, h="left", wrap=True)
            style(ws[f"{dd}{r}"], size=8, h="right")
            ws[f"{a}{r}"].border = Border(left=thin, top=thin, bottom=thin)
            ws[f"{b}{r}"].border = Border(top=thin, bottom=thin)
            ws[f"{cc}{r}"].border = Border(top=thin, bottom=thin)
            ws[f"{dd}{r}"].border = Border(right=thin, top=thin, bottom=thin)
        day = f"{a}10:{a}40"
        for t, k, col in TEAMS:
            book.cf(ws, day, f'AND({a}10<>"",INDEX(Plan_Nacht,{a}10-Jahresbeginn+1)="{t}")', bg=col)
        allc = f"{a}10:{dd}40"
        book.cf(ws, allc, f'AND(${a}10<>"",COUNTIF(FT_Datum,${a}10)>0)', color=HOLIDAY_FONT, bold=True)
        rest = f"{b}10:{dd}40"
        book.cf(ws, rest, f'AND(${a}10<>"",COUNTIF(FT_Datum,${a}10)>0)', bg=HOLIDAY, color=HOLIDAY_FONT, bold=True)
        book.cf(ws, rest, f'AND(${a}10<>"",WEEKDAY(${a}10,2)=7)', bg=SUN, bold=True)
        book.cf(ws, rest, f'AND(${a}10<>"",WEEKDAY(${a}10,2)=6)', bg=SAT, bold=True)
        book.cf(ws, rest, f'AND(${a}10<>"",SUMPRODUCT((SF_Von<=${a}10)*(SF_Bis>=${a}10))>0)', bg=FERIEN)
        book.cf(ws, day, f'AND({a}10<>"",WEEKDAY({a}10,2)>=6)', bold=True)
    put(ws, "A41", "Automatisch erzeugt: Nachtteam = Füllfarbe der Tageszahl (Rotation laut Einstellungen) · "
                   "Feiertage Hessen automatisch · Schulferien grau", size=8)
    put(ws, "AL41", '="Stand: "&' + dtxt("TODAY()"), size=8, h="right")
    ws.merge_cells("AL41:AV41")
    landscape(ws, "A3", fit_height=True, area="A1:AV41")
    ws.print_options.horizontalCentered = True
    ws.sheet_properties.tabColor = "7F7F7F"


# ====================================================================== Teamuebersicht
def build_team(book, ws):
    ws.sheet_view.showGridLines = False
    put(ws, "A1", '="Teamübersicht – Team "&$C$2&" – "&Jahr', size=16, bold=True)
    ws.row_dimensions[1].height = 26
    put(ws, "B2", "Team:", bold=True, h="right")
    put(ws, "C2", "Blau", bold=True, bg=INPUT_BG, h="center", border=THIN)
    book.dv_list(ws, "C2", "=Teams")
    team_cf(book, ws, "C2", "C2")
    put(ws, "D2", '="Stichtag "&' + dtxt("Stichtag") + '&" · Nachtteam heute: "&INDEX(Plan_Nacht,Stichtag-Jahresbeginn+1)',
        size=9, color="595959")
    n0 = "(Stichtag-Jahresbeginn+1)"
    put(ws, "A4", "Mitglieder", size=12, bold=True)
    hl = ["Nr", "Name", "Teamfunktion", "Funktionen (inkl. Mehrfachfunktion)", "Qualifikation (schlechtester Status)",
          "Nachtdienste", "Urlaubstage", "Krankheitstage", "Überstunden-Saldo", "JAZ-Konto", "Status heute"]
    header_row(ws, 5, 1, hl, height=30, size=8)
    first = "MATCH($C$2,C_Team,0)"
    cnt = 'COUNTIFS(C_Team,$C$2,C_Aktiv,"<>NEIN")'
    for k in range(12):
        r = 6 + k
        put(ws, f"N{r}", f"=IF(ISNA({first}),\"\",IF({k + 1}>{cnt},\"\",{first}+{k}))", size=7, color=MUTED)
        e = f'IF($N{r}="","",'
        vals = [f"={e}{k + 1})", f"={e}INDEX(C_Name,$N{r}))", f"={e}INDEX(C_Fkt,$N{r}))",
                f'={e}{X}TEXTJOIN(", ",TRUE,IF(INDEX(C_FM,$N{r})="JA","Meister",""),IF(INDEX(C_FAV,$N{r})="JA","AV",""),'
                f'IF(INDEX(C_FT,$N{r})="JA","TVT-Fahrer",""),IF(INDEX(C_FH,$N{r})="JA","Handwerker","")))',
                f'={e}IF(COUNTIF(Q_MA,B{r})=0,"–",CHOOSE({X}MAXIFS(Q_StatusNr,Q_MA,B{r}),"gültig","Vorwarnung",'
                f'"< 90 Tage","abgelaufen","Nachweis fehlt")))',
                f"={e}INDEX(ST_N,$N{r}))", f"={e}INDEX(ST_Urlaub,$N{r}))", f"={e}INDEX(ST_Krank,$N{r}))",
                f"={e}INDEX(UE_SaldoUe,$N{r}))", f"={e}INDEX(UE_SaldoJAZ,$N{r}))",
                f'={e}IF(INDEX(PEP_Grid,$N{r},{n0})="","frei",INDEX(PEP_Grid,$N{r},{n0})&""))']
        for j, v in enumerate(vals):
            put(ws, f"{L(1 + j)}{r}", v, size=9, border=THIN, h="center" if j not in (1, 3) else None,
                fmt="+0.0;-0.0;0.0" if j in (8, 9) else None, bold=j == 1)
    ws.column_dimensions["N"].hidden = True
    for st, bg, fc in [("gültig", "92D050", "000000"), ("Vorwarnung", "FFFF00", "000000"),
                       ("< 90 Tage", "FF0000", "FFFFFF"), ("abgelaufen", "C00000", "FFFFFF"),
                       ("Nachweis fehlt", "D9D9D9", "000000")]:
        book.cf(ws, "E6:E17", f'E6="{st}"', bg=bg, color=fc, bold=True)
    code_cf(book, ws, "K6:K17", "K6")
    # Kennzahlen
    put(ws, "A19", "Besetzung und Einsatzfähigkeit", size=12, bold=True)
    kl = ["Soll Meister", "Ist Meister", "Soll AV", "Ist AV", "Soll TVT", "Ist TVT", "Soll HW", "Ist HW", "Reserve",
          "Personalbedarf", "Ampel Struktur"]
    header_row(ws, 20, 1, kl, height=30, size=8)
    ti = "MATCH($C$2,Teams,0)"
    struct_cols = ["Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AJ", "AK"]  # Strukturtabelle Personalbedarf
    for j in range(11):
        col = struct_cols[j]
        put(ws, f"{L(1 + j)}21", f"=INDEX('Personalbedarfsplanung'!${col}$10:${col}$14,{ti})", h="center",
            border=THIN, bold=True)
    ampel_cf(book, ws, "K21", "K21")
    kl2 = ["Nachtdiensttage Jahr", "davon einsatzfähig", "Lücke mit Ersatz", "Lücke ohne Ersatz", "Quote",
           "Tage Team einsatzfähig (alle Tage)", "Status heute"]
    header_row(ws, 23, 1, kl2, height=30, size=8)
    vals = ['=COUNTIF(Plan_Nacht,$C$2)', '=COUNTIFS(EF_Nacht,$C$2,EF_Ampel,"Grün")',
            '=COUNTIFS(EF_Nacht,$C$2,EF_Ampel,"Gelb")', '=COUNTIFS(EF_Nacht,$C$2,EF_Ampel,"Rot")',
            "=IFERROR(B24/A24,0)", f'=COUNTIF(CHOOSE({ti},EF_St1,EF_St2,EF_St3,EF_St4,EF_St5),"✅*")',
            f"=INDEX(CHOOSE({ti},EF_St1,EF_St2,EF_St3,EF_St4,EF_St5),{n0})"]
    for j, v in enumerate(vals):
        put(ws, f"{L(1 + j)}24", v, h="center", border=THIN, bold=True, fmt="0%" if j == 4 else None)
    ws.merge_cells("G24:H24")
    book.cf(ws, "G24", 'LEFT(G24,1)="✅"', bg="C6EFCE", color="006100")
    book.cf(ws, "G24", 'LEFT(G24,1)="❌"', bg="FFC7CE", color="9C0006")
    # Monate
    put(ws, "A26", "Monatsübersicht", size=12, bold=True)
    header_row(ws, 27, 1, ["Monat", "Nachtdiensttage", "Grün", "Gelb", "Rot", "Urlaubstage", "Krankheitstage"],
               height=20, size=8)
    for m in range(12):
        r = 28 + m
        put(ws, f"A{r}", MONATE[m], border=THIN, bold=True, size=9)
        vv = [f"=COUNTIFS(Plan_Nacht,$C$2,Plan_Monat,{m + 1})",
              f'=COUNTIFS(EF_Nacht,$C$2,EF_Ampel,"Grün",Plan_Monat,{m + 1})',
              f'=COUNTIFS(EF_Nacht,$C$2,EF_Ampel,"Gelb",Plan_Monat,{m + 1})',
              f'=COUNTIFS(EF_Nacht,$C$2,EF_Ampel,"Rot",Plan_Monat,{m + 1})',
              f"=SUMIFS(INDEX(US_Mon,0,{m + 1}),US_Team,$C$2)",
              f"=SUMIFS(INDEX(KS_Mon,0,{m + 1}),KS_Team,$C$2)"]
        for j, v in enumerate(vv):
            put(ws, f"{L(2 + j)}{r}", v, h="center", border=THIN, size=9)
    book.cf(ws, "D28:D39", "D28>0", bg="FFEB9C")
    book.cf(ws, "E28:E39", "E28>0", bg="FFC7CE", color="9C0006", bold=True)
    # Naechste Bloecke
    put(ws, "H26", "Nächste Bereitschaftsblöcke ab Stichtag", size=12, bold=True)
    header_row(ws, 27, 8, ["Beginn", "Ende", "Block", "Ampel (schlechtester Tag)", "Fehlend am 1. Tag"], height=20,
               size=8)
    crit = f"(Plan_Nacht=$C$2)*(Plan_BS=1)*(Plan_N>={n0})"
    for k in range(12):
        r = 28 + k
        put(ws, f"O{r}", f"=SUMPRODUCT(SMALL({crit}*Plan_N+(1-{crit})*1E+9,{k + 1}))", size=7, color=MUTED)
        e = f'IF($O{r}>=1E+9,"",'
        put(ws, f"H{r}", f"={e}INDEX(Plan_Datum,$O{r}))", h="center", border=THIN, size=9, fmt=DATE_FMT)
        put(ws, f"I{r}", f'={e}MIN(Jahresende,H{r}+IF(INDEX(Plan_Typ,$O{r})="Mo-Do",3,2)))', h="center",
            border=THIN, size=9, fmt=DATE_FMT)
        put(ws, f"J{r}", f"={e}INDEX(Plan_Typ,$O{r}))", h="center", border=THIN, size=9)
        put(ws, f"K{r}", f'={e}IF(COUNTIFS(EF_Ampel,"Rot",Plan_N,">="&$O{r},Plan_N,"<="&($O{r}+I{r}-H{r}))>0,"Rot",'
                         f'IF(COUNTIFS(EF_Ampel,"Gelb",Plan_N,">="&$O{r},Plan_N,"<="&($O{r}+I{r}-H{r}))>0,"Gelb","Grün")))',
            h="center", border=THIN, size=9, bold=True)
        put(ws, f"L{r}", f"={e}INDEX(EF_Fehlend,$O{r}))", border=THIN, size=8)
        ws.merge_cells(f"L{r}:M{r}")
    ampel_cf(book, ws, "K28:K39", "K28")
    ws.column_dimensions["O"].hidden = True
    col_widths(ws, {"A": 12, "B": 17, "C": 13, "D": 30, "E": 18, "F": 11, "G": 11, "H": 11, "I": 11, "J": 9,
                    "K": 13, "L": 20, "M": 20})
    landscape(ws, "A4", fit_height=True, area="A1:M40")
    ws.sheet_properties.tabColor = "7F7F7F"


# ====================================================================== Qualifikationsuebersicht
def build_quali_print(book, ws):
    put(ws, "A1", '="Qualifikationsübersicht – Stichtag "&' + dtxt("Stichtag"), size=16, bold=True)
    ws.row_dimensions[1].height = 26
    put(ws, "A2", "Zelle = Ablaufdatum des neuesten Nachweises. Grün > 180 Tage · Gelb 90–180 Tage · Rot < 90 Tage · "
                  "Dunkelrot abgelaufen · Grau Nachweis fehlt („Pflicht ✗“ = Pflichtqualifikation der Teamfunktion "
                  "ohne Eintrag).", size=8, color=MUTED, italic=True)
    hl = ["Name", "Team", "Teamfunktion"] + U_NAMES + ["abgelaufen", "< 90 Tage", "fehlt"]
    header_row(ws, 5, 1, hl, height=48, size=8)
    for j in range(13):
        ws.cell(5, 4 + j).value = f"=INDEX(Unterrichte,{j + 1})"
    S0 = 21  # Spalte U: Status-Spiegel
    for i in range(NMAX):
        r = 6 + i
        cr = C_FIRST + i
        put(ws, f"A{r}", f"=Hilfsberechnung!$D{cr}", size=9, bold=True, border=THIN)
        put(ws, f"B{r}", f"=Hilfsberechnung!$E{cr}", size=8, h="center", border=THIN)
        put(ws, f"C{r}", f"=Hilfsberechnung!$F{cr}", size=8, h="center", border=THIN)
        for j in range(13):
            c = L(4 + j)
            mx = f"{X}MAXIFS(Q_Ablauf,Q_MA,$A{r},Q_Qual,{c}$5)"
            put(ws, f"{c}{r}", f'=IF($A{r}="","",IF(COUNTIFS(Q_MA,$A{r},Q_Qual,{c}$5)=0,'
                               f'IF(IFERROR(INDEX(ReqMatrix,MATCH($C{r},Funktionen4,0),{j + 1}),"")="JA","Pflicht ✗",""),'
                               f'IF({mx}=0,"fehlt",{mx})))', size=8, h="center", border=THIN, fmt=DATE_FMT)
            sc = L(S0 + j)
            ws[f"{sc}{r}"] = (f'=IF({c}{r}="","",IF(OR({c}{r}="Pflicht ✗",{c}{r}="fehlt"),5,'
                              f'{X}MINIFS(Q_StatusNr,Q_MA,$A{r},Q_Qual,{c}$5)))')
            ws[f"{sc}{r}"].font = font(7, color=MUTED)
        srng = f"${L(S0)}{r}:${L(S0 + 12)}{r}"
        for k, v in enumerate([4, 3, 5]):
            put(ws, f"{L(17 + k)}{r}", f'=IF($A{r}="","",COUNTIF({srng},{v}))', size=8, h="center", border=THIN,
                bold=True)
        ws.column_dimensions[L(S0 + j)].hidden = True
    for j in range(13):
        ws.column_dimensions[L(S0 + j)].hidden = True
    grid = "D6:P65"
    sc0 = L(S0)
    for nr, bg, fc in [(1, "92D050", "000000"), (2, "FFFF00", "000000"), (3, "FF0000", "FFFFFF"),
                       (4, "C00000", "FFFFFF"), (5, "D9D9D9", "000000")]:
        book.cf(ws, grid, f"{sc0}6={nr}", bg=bg, color=fc, bold=nr in (3, 4))
    book.cf(ws, "Q6:Q65", "AND(ISNUMBER(Q6),Q6>0)", bg="C00000", color="FFFFFF")
    book.cf(ws, "R6:R65", "AND(ISNUMBER(R6),R6>0)", bg="FF0000", color="FFFFFF")
    book.cf(ws, "S6:S65", "AND(ISNUMBER(S6),S6>0)", bg="D9D9D9")
    team_cf(book, ws, "B6:B65", "B6")
    put(ws, "A67", "Summe je Unterricht", size=11, bold=True)
    for k, (lab, bg, fc) in enumerate([("gültig", "92D050", "000000"), ("Vorwarnung", "FFFF00", "000000"),
                                       ("< 90 Tage", "FF0000", "FFFFFF"), ("abgelaufen", "C00000", "FFFFFF"),
                                       ("fehlt", "D9D9D9", "000000")]):
        r = 68 + k
        put(ws, f"C{r}", lab, size=8, bold=True, bg=bg, color=fc, border=THIN)
        for j in range(13):
            sc = L(S0 + j)
            put(ws, f"{L(4 + j)}{r}", f"=COUNTIF({sc}$6:{sc}$65,{k + 1})", size=8, h="center", border=THIN)
    col_widths(ws, {"A": 16, "B": 7, "C": 11, "Q": 9, "R": 9, "S": 7})
    for j in range(13):
        ws.column_dimensions[L(4 + j)].width = 10
    ws.freeze_panes = "D6"
    landscape(ws, "A4", rows="5:5")
    ws.sheet_properties.tabColor = "7F7F7F"


def build_all(book, wb, org_text=""):
    build_jahr(book, wb["Druck Jahr"])
    build_grid_print(book, wb, wb["Druck Monat"], "Monat")
    build_grid_print(book, wb, wb["Druck Quartal"], "Quartal")
    build_team(book, wb["Druck Team"])
    build_bereitschaft(book, wb["Druck Bereitschaft"], org_text)
    build_quali_print(book, wb["Druck Qualifikation"])
    wb["Druck Qualifikation"].print_area = "A1:S72"
