"""Druckansichten: Jahresuebersicht (A3), Monatsuebersicht (A4), Teamuebersicht (A4), Bereitschaftsuebersicht."""
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

from . import model as M
from .ctx import S_DB, S_DJ, S_DM, S_DT, S_KAL
from .style import (BOX, INPUT_FILL, NOBORDER, cf, dv_list, dxf, fill, font, landscape, note, put, q, title, widths)

KAL = q(S_KAL)
KGRID = f"{KAL}!${M.C_DAY0}${M.ROW0}:${M.C_DAYN}${M.ROWN}"
LEFT = Border(left=Side(style="thin", color="000000"))
CODES = (8, 4, 3, 2, 7, 5, 6, 10, 1, 11, 9, 12, 13)


def status_cf(ws, rng, ref, typ_ref=None):
    """Farben je Statuscode. ref = relative Zelle mit dem Code (oben links des Bereichs)."""
    for code in CODES:
        s = M.STATUS_BY_CODE[code]
        cf(ws, rng, f"{ref}={code}",
           dxf(bg=s[5], fg=s[6], bold=s[7], underline="single" if s[8] else None, italic=s[9] or None))


LEGEND_ITEMS = [(3, "Urlaub"), (4, "Urlaub eTime"), (2, "JAZ + 99"), (8, "Krank"), (7, "Unterricht"),
                (6, "FDL-Einw."), (5, "Kabeleinw."), (1, "Besonderheit"), (12, "Zusätzl. MA"), (13, "AZUBIS"),
                (9, "Bereitschaft"), (10, "Freistellung")]


def legend_row(ws, row, col0, per_row=6, step=5, size=7, swatch=1):
    """Legende in Zeilen zu je per_row Eintraegen; je Eintrag Farbfeld (swatch Spalten) + Text."""
    put(ws, f"{gl(col0)}{row}", "Legende:", bold=True, size=size)
    items = [(M.STATUS_BY_CODE[c], lab) for c, lab in LEGEND_ITEMS]
    items += [((None, None, None, None, None, M.WEEKEND_FILL, "000000", False, False, False), "Wochenende"),
              ((None, None, None, None, None, M.HOLIDAY_FILL, "000000", False, False, False), "Feiertag"),
              ((None, None, None, None, None, None, "000000", False, False, False), "leer = Tagdienst")]
    start = col0 + 2
    for i, (s_, lab) in enumerate(items):
        r = row + i // per_row
        c = start + (i % per_row) * step
        txt = s_[2] if s_[2] and s_[2] != "!" else ("Text" if s_[0] == 1 else None)
        cell = f"{gl(c)}{r}"
        if swatch > 1:
            ws.merge_cells(f"{cell}:{gl(c + swatch - 1)}{r}")
        put(ws, cell, txt, size=size, bg=s_[5], color=s_[6], bold=s_[7], italic=s_[9],
            underline="single" if s_[8] else None, h="center", border=BOX)
        put(ws, f"{gl(c + swatch)}{r}", lab, size=size)
    return row + (len(items) - 1) // per_row


# =========================================================================== Monatsuebersicht
def month_grid(ws, ctx, first_row, rows_spec, month_ref, label_cols=("B", "C"), title_text=None):
    """Tage des Monats in Spalten D..AH, Codes versteckt in AK..BO.
    rows_spec: Liste (Blattzeile, Index in Ber_Grid als Formeltext, Bedingung oder None)."""
    D0, CODE0 = 4, 37
    hr = first_row - 6     # Kopfzeilen: hr = Tagindex (versteckt), hr+1 Tag, +2 WT, +3 KW, +4 Feiertag, +5 Nachtteam
    for j in range(1, 32):
        c = gl(D0 + j - 1)
        ws[f"{c}{hr}"] = (f'=IF({j}>DAY(DATE(Jahr,{month_ref}+1,0)),"",DATE(Jahr,{month_ref},{j})-DATE(Jahr,1,1)+1)')
        ws[f"{c}{hr}"].font = font(6, color="FFFFFF")
        ws[f"{c}{hr + 1}"] = f'=IF({c}${hr}="","",DATE(Jahr,{month_ref},{j}))'
        ws[f"{c}{hr + 1}"].number_format = "D"
        ws[f"{c}{hr + 2}"] = f'=IF({c}${hr}="","",INDEX(lstWT,WEEKDAY({c}${hr + 1},2)))'
        ws[f"{c}{hr + 3}"] = f'=IF({c}${hr}="","",IF(WEEKDAY({c}${hr + 1},2)=1,_xlfn.ISOWEEKNUM({c}${hr + 1}),""))'
        ws[f"{c}{hr + 4}"] = f'=IF({c}${hr}="","",INDEX(Schicht_Feiertag,{c}${hr}))'
        ws[f"{c}{hr + 5}"] = f'=IF({c}${hr}="","",INDEX(Schicht_Nacht,{c}${hr}))'
        for k, (sz, b, rot) in enumerate([(9, True, 0), (8, False, 0), (7, False, 0), (6, True, 90), (6, True, 0)]):
            cell = ws[f"{c}{hr + 1 + k}"]
            cell.font = font(sz, b, "CC0000" if k == 3 else "000000")
            cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=rot, wrap_text=k == 3,
                                       shrink_to_fit=k == 4)
            cell.border = BOX
            if k < 3:
                cell.fill = fill(M.HEADER_FILL)
        cc = gl(CODE0 + j - 1)
        for (r, idx, cond) in rows_spec:
            ok = f"AND({c}${hr}<>\"\",{cond})" if cond else f'{c}${hr}<>""'
            ws[f"{c}{r}"] = f'=IF({ok},INDEX({KGRID},{idx},{c}${hr}),"")'
            ws[f"{c}{r}"].font = font(7, True)
            ws[f"{c}{r}"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"{c}{r}"].border = BOX
            ws[f"{cc}{r}"] = f"=IF({ok},INDEX(Ber_Grid,{idx},{c}${hr}),0)"
            ws[f"{cc}{r}"].font = font(6, color="FFFFFF")
        ws.column_dimensions[c].width = 3.7
    for lab, r in zip(["Tag", "Wochentag", "KW", "Feiertag", "Nachtteam"], range(hr + 1, hr + 6)):
        put(ws, f"{label_cols[0]}{r}", lab, size=8, bold=True, h="right", bg=M.HEADER_FILL, border=BOX)
        put(ws, f"{label_cols[1]}{r}", None, bg=M.HEADER_FILL, border=BOX)
    ws.row_dimensions[hr + 4].height = 44
    rows = [r for r, _, _ in rows_spec]
    r1, r2 = min(rows), max(rows)
    grid = f"D{r1}:AH{r2}"
    cf(ws, f"D{hr + 1}:AH{r2}", f"AND(ISNUMBER(D${hr + 1}),DAY(D${hr + 1})=1)", dxf(border=LEFT), stop=False)
    return hr, grid, r1, r2, CODE0 - D0


def build_druck_monat(ctx):
    ws = ctx.wb[S_DM]
    ws.sheet_properties.tabColor = "808080"
    ws.merge_cells("B1:AH1")
    put(ws, "B1", '="Monatsübersicht "&$C$2&" "&Jahr', size=14, bold=True, bg=M.TITLE_FILL, color="FFFFFF")
    put(ws, "B2", "Monat:", bold=True, h="right", size=9)
    put(ws, "C2", "Januar", bg=INPUT_FILL, border=BOX, bold=True, h="center", size=9)
    dv_list(ws, "C2", "lstMonate", prompt="Monat für den Druck wählen")
    ws["D2"] = "=MATCH($C$2,lstMonate,0)"
    ws["D2"].font = font(6, color="FFFFFF")
    first = 9
    spec = [(first + i, i + 1, None) for i in range(M.CAP)]
    hr, grid, r1, r2, off = month_grid(ws, ctx, first, spec, "$D$2")
    typ, grp = "BQ", "BR"
    for r, i, _ in spec:
        br = M.ROW0 + i - 1
        ws[f"B{r}"] = f"=INDEX(Ber_Name,{i})"
        ws[f"C{r}"] = f"=INDEX(Ber_Funktion,{i})"
        ws[f"{typ}{r}"] = f"=INDEX(Ber_Typ,{i})"
        ws[f"{grp}{r}"] = f"=INDEX(Ber_Gruppe,{i})"
        for c, sz in (("B", 8), ("C", 7)):
            ws[f"{c}{r}"].font = font(sz)
            ws[f"{c}{r}"].border = BOX
        ws.row_dimensions[r].height = 12.5
    _grid_cf(ws, grid, r1, typ, grp, off, hr)
    lend = legend_row(ws, r2 + 2, 2)
    widths(ws, {"A": 1.5, "B": 18, "C": 10})
    for ci in range(35, 71):
        ws.column_dimensions[gl(ci)].hidden = True
    landscape(ws, "A4", 1, 1, area=f"A1:AH{lend}")
    ws.freeze_panes = f"D{first}"
    return ws


def _grid_cf(ws, grid, r1, typ, grp, off, hr):
    code = f"{gl(4 + off)}{r1}"
    for g in M.GROUPS:
        cf(ws, f"A{r1}:AH{int(grid.split(':')[1][2:])}", f'AND(${typ}{r1}="K",${grp}{r1}="{g}")',
           dxf(bg=M.GROUP_COLOR[g], fg=M.GROUP_FONT[g], bold=True))
        cf(ws, f"A{r1}:A{int(grid.split(':')[1][2:])}", f'AND(${typ}{r1}="P",${grp}{r1}="{g}")', dxf(bg=M.GROUP_COLOR[g]))
    status_cf(ws, grid, code)
    cf(ws, grid, f'AND(${typ}{r1}="P",D${hr + 4}<>"")', dxf(bg=M.HOLIDAY_FILL))
    cf(ws, grid, f'AND(${typ}{r1}="P",OR(D${hr + 2}="Sa",D${hr + 2}="So"))', dxf(bg=M.WEEKEND_FILL))
    hdr = f"D{hr + 1}:AH{hr + 3}"
    cf(ws, hdr, f'D${hr + 4}<>""', dxf(bg=M.HOLIDAY_FILL, fg="CC0000", bold=True))
    cf(ws, hdr, f'OR(D${hr + 2}="Sa",D${hr + 2}="So")', dxf(bg=M.WEEKEND_FILL, bold=True))
    cf(ws, f"D{hr + 4}:AH{hr + 4}", f'D{hr + 4}<>""', dxf(bg=M.HOLIDAY_FILL))
    for t in M.TEAMS:
        cf(ws, f"D{hr + 5}:AH{hr + 5}", f'D{hr + 5}="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))


# =========================================================================== Teamuebersicht (Druck)
def build_druck_team(ctx):
    ws = ctx.wb[S_DT]
    ws.sheet_properties.tabColor = "808080"
    ws.merge_cells("B1:AH1")
    put(ws, "B1", '="Teamübersicht Team "&$C$2&" – "&$E$2&" "&Jahr', size=14, bold=True, bg=M.TITLE_FILL, color="FFFFFF")
    put(ws, "B2", "Team:", bold=True, h="right", size=9)
    put(ws, "C2", "Blau", bg=INPUT_FILL, border=BOX, bold=True, h="center", size=9)
    dv_list(ws, "C2", "lstTeams", prompt="Team wählen")
    put(ws, "D2", None)
    ws.merge_cells("E2:H2")
    put(ws, "E2", "Januar", bg=INPUT_FILL, border=BOX, bold=True, h="center", size=9)
    dv_list(ws, "E2", "lstMonate", prompt="Monat wählen")
    ws["I2"] = "=MATCH($E$2,lstMonate,0)"
    ws["I2"].font = font(6, color="FFFFFF")
    ws["J2"] = "=MATCH($C$2,lstGruppen,0)"
    ws["J2"].font = font(6, color="FFFFFF")
    first = 9
    spec = []
    for j in range(1, 10):
        idx = f"(INDEX(Grp_Start,$J$2)+{j})"
        spec.append((first + j - 1, idx, f"{j}<=INDEX(Grp_Anzahl,$J$2)"))
    hr, grid, r1, r2, off = month_grid(ws, ctx, first, spec, "$I$2")
    typ, grp = "BQ", "BR"
    for j, (r, idx, cond) in enumerate(spec, start=1):
        ws[f"B{r}"] = f'=IF({cond},INDEX(Ber_Name,{idx}),"")'
        ws[f"C{r}"] = f'=IF({cond},INDEX(Ber_Funktion,{idx}),"")'
        ws[f"{typ}{r}"] = f'=IF({cond},"P","")'
        ws[f"{grp}{r}"] = "=$C$2"
        for c, sz in (("B", 9), ("C", 8)):
            ws[f"{c}{r}"].font = font(sz)
            ws[f"{c}{r}"].border = BOX
        ws.row_dimensions[r].height = 16
    _grid_cf(ws, grid, r1, typ, grp, off, hr)
    # Team-Kennzahlen im Monat
    k0 = r2 + 2
    put(ws, f"B{k0}", "Kennzahlen im Monat", bold=True, size=10)
    m = "$I$2"
    kp = [
        ("Bereitschaftstage", f'=COUNTIFS(Schicht_Nacht,$C$2,Bedarf_Monat,{m})'),
        ("davon Wochenende/Feiertag", f'=COUNTIFS(Schicht_Nacht,$C$2,Bedarf_Monat,{m},Bedarf_AT,0)'),
        ("Freistellungstage (Vorschlag)", f'=COUNTIFS(Schicht_FSauto,$C$2,Bedarf_Monat,{m})'),
        ("Tage 🔴 nicht einsatzfähig", f'=COUNTIFS(Einsatz_Team,$C$2,Bedarf_Monat,{m},Einsatz_Stufe,2)'),
        ("Tage 🟡 kritisch", f'=COUNTIFS(Einsatz_Team,$C$2,Bedarf_Monat,{m},Einsatz_Stufe,1)'),
        ("Abwesenheitstage (Arbeitstage)", f'=SUMPRODUCT((Ber_Grid>=2)*(Ber_Grid<=8)*(Ber_Gruppe=$C$2)*(Ber_Monat={m})*(Ber_AT=1))'),
    ]
    for i, (lab, f_) in enumerate(kp):
        put(ws, f"B{k0 + 1 + i}", lab, size=8, border=BOX)
        put(ws, f"C{k0 + 1 + i}", f_, size=9, bold=True, border=BOX, h="center")
    # Qualifikationen der Mitglieder
    put(ws, f"E{k0}", "Qualifikationen (Funktionsmatrix)", bold=True, size=10)
    quals = M.QUALIS
    for j, (r, idx, cond) in enumerate(spec, start=1):
        rr = k0 + j
        nm = f"B{r}"
        parts = ",".join(f'IF(IFERROR(INDEX(tblFunktionsmatrix[{qn}],MATCH({nm},FM_Name,0)),"")="JA","{qn}","")'
                         for qn in quals)
        ws[f"E{rr}"] = f'=IF({nm}="","",{nm}&": "&_xlfn.TEXTJOIN(", ",TRUE,{parts}))'
        ws[f"E{rr}"].font = font(8)
    lend = legend_row(ws, k0 + 11, 2)
    widths(ws, {"A": 1.5, "B": 24, "C": 11})
    for ci in range(35, 71):
        ws.column_dimensions[gl(ci)].hidden = True
    landscape(ws, "A4", 1, 1, area=f"A1:AH{lend}")
    return ws


# =========================================================================== Jahresuebersicht (A3)
def build_druck_jahr(ctx):
    ws = ctx.wb[S_DJ]
    ws.sheet_properties.tabColor = "808080"
    put(ws, "B1", '="Jahresübersicht "&Jahr&" – Abwesenheiten, Bereitschaft und Ressourcen"', size=16, bold=True)
    note(ws, "B2", "Je Zeile ein Mitarbeiter, je Spalte ein Tag (Farben wie im Jahreskalender). Erstes Halbjahr oben, "
         "zweites Halbjahr unten.")
    NC = 184
    C0 = 3
    typ, grp = gl(C0 + NC + 2), gl(C0 + NC + 3)
    blocks = [(4, "0"), (4 + M.CAP + 6, "DATE(Jahr,7,1)-DATE(Jahr,1,1)")]
    for b, (top, offset) in enumerate(blocks):
        hr = top            # versteckter Tagindex
        put(ws, f"B{hr + 1}", "1. Halbjahr" if b == 0 else "2. Halbjahr", bold=True, size=11)
        for j in range(1, NC + 1):
            c = gl(C0 + j - 1)
            lim = "DATE(Jahr,7,1)-DATE(Jahr,1,1)" if b == 0 else "DATE(Jahr,12,31)-DATE(Jahr,1,1)+1"
            ws[f"{c}{hr}"] = f'=IF({offset}+{j}>{lim},"",{offset}+{j})'
            ws[f"{c}{hr}"].font = font(5, color="FFFFFF")
            ws[f"{c}{hr}"].number_format = ";;;"
            ws[f"{c}{hr + 2}"] = f'=IF({c}${hr}="","",DATE(Jahr,1,1)+{c}${hr}-1)'
            ws[f"{c}{hr + 2}"].number_format = ";;;"
            for i in range(M.CAP):
                r = hr + 3 + i
                cell = ws[f"{c}{r}"]
                cell.value = f'=IF({c}${hr}="",0,INDEX(Ber_Grid,{i + 1},{c}${hr}))'
                cell.number_format = ";;;"
                cell.font = font(5)
            ws.column_dimensions[c].width = 0.9
        lens = [31, 28, 31, 30, 31, 30] if b == 0 else [31, 31, 30, 31, 30, 31]
        cstart = C0
        for k, ln in enumerate(lens):
            mi = k + (0 if b == 0 else 6)
            a_, z_ = gl(cstart), gl(cstart + ln - 1)
            ws.merge_cells(f"{a_}{hr + 1}:{z_}{hr + 1}")
            put(ws, f"{a_}{hr + 1}", M.MONATE[mi], size=9, bold=True, h="center", bg=M.HEADER_FILL, border=BOX)
            cstart += ln
        for i in range(M.CAP):
            r = hr + 3 + i
            ws[f"B{r}"] = f"=INDEX(Ber_Name,{i + 1})"
            ws[f"B{r}"].font = font(7)
            ws[f"{typ}{r}"] = f"=INDEX(Ber_Typ,{i + 1})"
            ws[f"{grp}{r}"] = f"=INDEX(Ber_Gruppe,{i + 1})"
            ws.row_dimensions[r].height = 10
        r1, r2 = hr + 3, hr + 2 + M.CAP
        grid = f"{gl(C0)}{r1}:{gl(C0 + NC - 1)}{r2}"
        dref = f"{gl(C0)}${hr + 2}"
        cf(ws, f"{gl(C0)}{hr + 1}:{gl(C0 + NC - 1)}{r2}", f"AND(ISNUMBER({dref}),DAY({dref})=1)", dxf(border=LEFT), stop=False)
        for g in M.GROUPS:
            cf(ws, f"A{r1}:{gl(C0 + NC - 1)}{r2}", f'AND(${typ}{r1}="K",${grp}{r1}="{g}")',
               dxf(bg=M.GROUP_COLOR[g], fg=M.GROUP_FONT[g], bold=True))
        status_cf(ws, grid, f"{gl(C0)}{r1}")
        cf(ws, grid, f'AND(${typ}{r1}="P",ISNUMBER({dref}),COUNTIF(FT_Aktiv,{dref})>0)', dxf(bg=M.HOLIDAY_FILL))
        cf(ws, grid, f'AND(${typ}{r1}="P",ISNUMBER({dref}),WEEKDAY({dref},2)>=6)', dxf(bg=M.WEEKEND_FILL))
    end = blocks[1][0] + 2 + M.CAP
    lend = legend_row(ws, end + 2, 2, per_row=8, step=22, size=8, swatch=6)
    widths(ws, {"A": 1, "B": 16})
    ws.column_dimensions[typ].hidden = True
    ws.column_dimensions[grp].hidden = True
    landscape(ws, "A3", 1, 1, area=f"A1:{gl(C0 + NC - 1)}{lend}", margins=0.3)
    return ws


# =========================================================================== Bereitschaftsuebersicht
def build_druck_bereitschaft(ctx):
    """Layout wie der bestehende Bereitschaftskalender 2027 (Kalenderpedia), aber vollstaendig per Formel."""
    ws = ctx.wb[S_DB]
    ws.sheet_properties.tabColor = "808080"
    for i in range(12):
        for k, w in enumerate([3.29, 3.71, 5.71, 2.71]):
            ws.column_dimensions[gl(1 + 4 * i + k)].width = w
    for r in range(1, 9):
        ws.row_dimensions[r].height = 12
    ws.row_dimensions[9].height = 27
    for r in range(10, 41):
        ws.row_dimensions[r].height = 21
    ws.merge_cells("C2:F3")
    put(ws, "C2", "=Jahr", size=24, bold=True, h="center")
    put(ws, "H2", "Bereitschaftskalender", size=12, bold=True)
    put(ws, "H3", "Nachtdienst/Bereitschaft je Tag in Teamfarbe", size=8, italic=True, color="595959")
    roles = [("Meister", "Meister", 1), ("AV", "AV", 1), ("TVT Fahrer", "TVT-Fahrer", 1),
             ("Handwerker 1", "Handwerker", 1), ("Handwerker 2", "Handwerker", 2), ("Handwerker 3", "Handwerker", 3)]
    team_cols = {"Blau": 25, "Rot": 29, "Gelb": 33, "Grün": 37, "Orange": 41}
    for i, (lab, rolle, k) in enumerate(roles):
        r = 3 + i
        ws.merge_cells(f"U{r}:X{r}")
        put(ws, f"U{r}", lab, bold=True, h="center", border=BOX)
        for t, c in team_cols.items():
            a, b = gl(c), gl(c + 3)
            ws.merge_cells(f"{a}{r}:{b}{r}")
            nth = (f'SUMPRODUCT(SMALL(Ber_Rang+((Ber_Gruppe<>"{t}")+(Ber_Funktion<>"{rolle}")+(Ber_Typ<>"P")>0)*1000,{k}))')
            put(ws, f"{a}{r}", f'=IF({nth}>=1000,"",INDEX(Ber_Name,{nth}))', bold=True, v="center",
                bg=M.TEAM_COLOR[t], color=M.TEAM_FONT[t], border=BOX)
            for cc in range(c + 1, c + 4):
                ws.cell(r, cc).fill = fill(M.TEAM_COLOR[t])
                ws.cell(r, cc).border = BOX
    for t, c in team_cols.items():
        ws.merge_cells(f"{gl(c)}2:{gl(c + 3)}2")
        put(ws, f"{gl(c)}2", f"Team {t}", bold=True, size=8, h="center", color="404040")
    for mi in range(12):
        c0 = 1 + 4 * mi
        A, B, C, D = (gl(c0 + k) for k in range(4))
        ws.merge_cells(f"{A}9:{D}9")
        put(ws, f"{A}9", M.MONATE[mi], size=14, bold=True, h="center", border=BOX)
        for k in range(1, 4):
            ws.cell(9, c0 + k).border = BOX
        mnum = mi + 1
        for d in range(1, 32):
            r = 9 + d
            ok = f"{d}<=DAY(DATE(Jahr,{mnum}+1,0))"
            ws[f"{A}{r}"] = f'=IF({ok},DATE(Jahr,{mnum},{d}),"")'
            ws[f"{A}{r}"].number_format = "D"
            ws[f"{B}{r}"] = f'=IF(ISNUMBER({A}{r}),INDEX(lstWT,WEEKDAY({A}{r},2)),"")'
            ws[f"{C}{r}"] = f'=IF(ISNUMBER({A}{r}),IFERROR(INDEX(FT_Name,MATCH({A}{r},FT_Datum,0)),""),"")'
            ws[f"{D}{r}"] = f'=IF(ISNUMBER({A}{r}),IF(WEEKDAY({A}{r},2)=1,_xlfn.ISOWEEKNUM({A}{r}),""),"")'
            ws[f"{A}{r}"].font = font(11)
            ws[f"{A}{r}"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"{B}{r}"].font = font(11)
            ws[f"{B}{r}"].alignment = Alignment(vertical="center")
            ws[f"{C}{r}"].font = font(7)
            ws[f"{C}{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws[f"{D}{r}"].font = font(8)
            ws[f"{D}{r}"].alignment = Alignment(horizontal="right", vertical="center")
            thin = Side(style="thin", color="000000")
            ws[f"{A}{r}"].border = Border(left=thin, top=thin, bottom=thin)
            ws[f"{B}{r}"].border = Border(top=thin, bottom=thin)
            ws[f"{C}{r}"].border = Border(top=thin, bottom=thin)
            ws[f"{D}{r}"].border = Border(right=thin, top=thin, bottom=thin)
        rng_day = f"{A}10:{A}40"
        rng_rest = f"{B}10:{D}40"
        team = f"INDEX(RotTeam,MOD(2*INT(({A}10-RotStart)/7)+(MOD({A}10-RotStart,7)>=3),10)+1)"
        hol = f"ISNUMBER(MATCH({A}10,FT_Aktiv,0))"
        cf(ws, f"{A}10:{D}40", f'NOT(ISNUMBER(${A}10))', dxf(bg="C0C0C0"))
        for t in M.TEAMS:
            cf(ws, rng_day, f'AND(ISNUMBER({A}10),{team}="{t}")',
               dxf(bg=M.TEAM_COLOR[t]), stop=False)
        cf(ws, f"{A}10:{B}40", f"AND(ISNUMBER(${A}10),{hol.replace(A + '10', '$' + A + '10')})",
           dxf(fg="CC0000", bold=True), stop=False)
        cf(ws, f"{A}10:{B}40", f"AND(ISNUMBER(${A}10),WEEKDAY(${A}10,2)>=6)", dxf(bold=True), stop=False)
        cf(ws, rng_rest, f"AND(ISNUMBER(${A}10),{hol.replace(A + '10', '$' + A + '10')})", dxf(bg=M.HOLIDAY_FILL))
        cf(ws, rng_rest, f"AND(ISNUMBER(${A}10),WEEKDAY(${A}10,2)=6)", dxf(bg=M.SAT_FILL))
        cf(ws, rng_rest, f"AND(ISNUMBER(${A}10),WEEKDAY(${A}10,2)=7)", dxf(bg=M.SUN_FILL))
    note(ws, "A41", "Layout wie Bereitschaftskalender 2027 · Daten automatisch aus dem PEP-System (Rotation, Feiertage Hessen, Teams)")
    landscape(ws, "A4", 1, 1, area="A1:AV41", margins=0.3)
    return ws


def build_all(ctx):
    build_druck_jahr(ctx)
    build_druck_monat(ctx)
    build_druck_team(ctx)
    build_druck_bereitschaft(ctx)
