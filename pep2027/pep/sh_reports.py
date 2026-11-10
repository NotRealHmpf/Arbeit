"""Auswertungen: Statistik je MA, Urlaubs-/Krankenstatistik, Monatsbericht, Jahresbericht, Teamuebersicht."""
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as gl

from . import model as M
from .ctx import S_BER, S_JAHR, S_MON, S_STA, S_TEAM, S_UK
from .style import (BOX, INPUT_FILL, cf, dv_list, dxf, fill, font, note, put, q, title, widths)

BER = q(S_BER)
C0, CN = M.C_DAY0, M.C_DAYN
ON = '">=9",{r},"<=11",{r},"<>10"'


def brow(br):
    return f"{BER}!${C0}${br}:${CN}${br}"


def on_cnt(br, extra=""):
    r = brow(br)
    return f'COUNTIFS({r},">=9",{r},"<=11",{r},"<>10"{extra})'


def code_cnt(br, code, extra=",Ber_AT,1"):
    return f"COUNTIFS({brow(br)},{code}{extra})"


def mirror(ws, s0, name_col="B", grp_col="C", fn_col="D", typ_col=None, size=9):
    """Personenzeilen aus Berechnung spiegeln (inkl. Teamkopfzeilen). -> Liste (Blattzeile, Berechnungszeile)."""
    rows = []
    for i in range(M.CAP):
        r, br = s0 + i, M.ROW0 + i
        ws[f"{name_col}{r}"] = f"={BER}!$B{br}"
        ws[f"{grp_col}{r}"] = f'=IF({BER}!$A{br}="P",{BER}!$D{br},"")'
        if fn_col:
            ws[f"{fn_col}{r}"] = f"={BER}!$C{br}"
        if typ_col:
            ws[f"{typ_col}{r}"] = f"={BER}!$A{br}"
            ws[f"{typ_col}{r}"].font = font(7, color="FFFFFF")
            gc = gl(ws[f"{typ_col}{r}"].column + 1)
            ws[f"{gc}{r}"] = f"={BER}!$D{br}"
            ws[f"{gc}{r}"].font = font(7, color="FFFFFF")
        for c in (name_col, grp_col, fn_col):
            if c:
                ws[f"{c}{r}"].font = font(size)
                ws[f"{c}{r}"].border = BOX
        rows.append((r, br))
    return rows


def group_cf(ws, rng, typ_col, first_row):
    """Teamkopfzeilen in Teamfarbe (Typ K); Gruppe steht in der Spalte rechts neben der Typspalte."""
    gc = gl(ws[f"{typ_col}1"].column + 1)
    for g in M.GROUPS:
        cf(ws, rng, f'AND(${typ_col}{first_row}="K",${gc}{first_row}="{g}")',
           dxf(bg=M.GROUP_COLOR[g], fg=M.GROUP_FONT[g], bold=True))


# =========================================================================== Statistik
STA_COLS = [("Nachtdienste", "on"), ("davon Wochenende (Fr–So)", "we"), ("Feiertagsdienste", "ft"),
            ("Bereitschaften (Blöcke)", "blk"), ("Vertretungstage", "vt"), ("Urlaubstage", "u"),
            ("davon eTime", "ue"), ("Krankheitstage", "k"), ("Unterrichtstage", "s"), ("JAZ-Tage", "j"),
            ("FDL-Einweisungen", "fdl"), ("Kabeleinweisungen", "kab"), ("Freistellungstage", "fs"),
            ("Besonderheiten", "bes"), ("Abwesenheitstage gesamt", "abw"), ("Anwesenheitsquote", "quote")]


def sta_formula(key, br, r, cols):
    r_ = brow(br)
    nx = lambda c: f"{BER}!${gl(M.DAY0 + 1)}${br}:${CN}${br}" if c else None
    if key == "on":
        return on_cnt(br)
    if key == "we":
        return on_cnt(br, ",Ber_WE,1")
    if key == "ft":
        return on_cnt(br, ",Ber_FT,1")
    if key == "blk":
        d1 = f"{BER}!${C0}${br}"
        cur = f"{BER}!${gl(M.DAY0 + 1)}${br}:${CN}${br}"
        prev = f"{BER}!${C0}${br}:${gl(M.DAYN - 1)}${br}"
        bs = f"{BER}!${gl(M.DAY0 + 1)}$1:${CN}$1"
        oncur = f"({cur}>=9)*({cur}<=11)*({cur}<>10)"
        onprev = f"({prev}>=9)*({prev}<=11)*({prev}<>10)"
        return f"({d1}>=9)*({d1}<=11)*({d1}<>10)+SUMPRODUCT({oncur}*(((1-{onprev})+{bs})>0))"
    if key == "vt":
        return on_cnt(br, f',Ber_Nacht,"<>"&{BER}!$D{br}')
    if key == "u":
        return f"{code_cnt(br, 3)}+{code_cnt(br, 4)}"
    if key == "ue":
        return code_cnt(br, 4)
    m = {"k": 8, "s": 7, "j": 2, "fdl": 6, "kab": 5}
    if key in m:
        return code_cnt(br, m[key])
    if key == "fs":
        return f"COUNTIF({r_},10)"
    if key == "bes":
        return f"COUNTIF({r_},1)+COUNTIF({r_},11)"
    if key == "abw":
        return "+".join(f"{cols[k]}{r}" for k in ("u", "k", "s", "j", "fdl", "kab"))
    if key == "quote":
        return f'IF(SUM(Ber_AT)=0,"",MAX(0,SUM(Ber_AT)-{cols["abw"]}{r}-COUNTIFS({r_},10,Ber_AT,1))/SUM(Ber_AT))'
    raise KeyError(key)


def build_statistik(ctx):
    ws = ctx.wb[S_STA]
    ws.sheet_properties.tabColor = "7030A0"
    title(ws, "B1", "Statistik je Mitarbeiter")
    put(ws, "F1", '="Planungsjahr "&Jahr', bold=True, size=12)
    note(ws, "B2", "Nachtdienste = Bereitschaftsnächte (15:30–06:45, inkl. Wochenende/Feiertag). Bereitschaften = "
         "Anzahl Einsätze (Blöcke Mo–Do bzw. Fr–So). Abwesenheiten in Arbeitstagen (Mo–Fr ohne Feiertage). "
         "Teamzeilen zeigen die Summe des Teams.")
    s0 = 5
    hdr = ["Name", "Gruppe", "Funktion"] + [h for h, _ in STA_COLS]
    for i, h in enumerate(hdr):
        put(ws, f"{gl(2 + i)}{s0 - 1}", h, bold=True, wrap=True, h="center", bg=M.HEADER_FILL, border=BOX, size=9)
    ws.row_dimensions[s0 - 1].height = 42
    cols = {k: gl(5 + i) for i, (_, k) in enumerate(STA_COLS)}
    typ = gl(5 + len(STA_COLS) + 1)
    rows = mirror(ws, s0, "B", "C", "D", typ)
    last = s0 + M.CAP - 1
    for r, br in rows:
        for _, k in STA_COLS:
            c = cols[k]
            per = sta_formula(k, br, r, cols)
            if k == "quote":
                grp = f'IFERROR(AVERAGEIFS({c}${s0}:{c}${last},$C${s0}:$C${last},{BER}!$D{br},${typ}${s0}:${typ}${last},"P"),"")'
            else:
                grp = f'SUMIFS({c}${s0}:{c}${last},$C${s0}:$C${last},{BER}!$D{br},${typ}${s0}:${typ}${last},"P")'
            ws[f"{c}{r}"] = f'=IF(${typ}{r}="P",{per},IF(${typ}{r}="K",{grp},""))'
            ws[f"{c}{r}"].font = font(9)
            ws[f"{c}{r}"].alignment = Alignment(horizontal="center")
            ws[f"{c}{r}"].border = BOX
            ws[f"{c}{r}"].number_format = "0%" if k == "quote" else "0;-0;;@"
    rt = last + 1
    put(ws, f"B{rt}", "Summe Teams", bold=True, border=BOX)
    for _, k in STA_COLS:
        c = cols[k]
        f_ = (f'=IFERROR(AVERAGEIFS({c}{s0}:{c}{last},${typ}{s0}:${typ}{last},"P"),"")' if k == "quote"
              else f'=SUMIFS({c}{s0}:{c}{last},${typ}{s0}:${typ}{last},"P",$C{s0}:$C{last},"<>Ressourcen")')
        put(ws, f"{c}{rt}", f_, bold=True, border=BOX, h="center", nf="0%" if k == "quote" else "0")
    widths(ws, {"A": 2, "B": 24, "C": 11, "D": 13})
    for c in cols.values():
        ws.column_dimensions[c].width = 10
    ws.column_dimensions[typ].width = 3
    rng = f"B{s0}:{cols['quote']}{last}"
    group_cf(ws, rng, typ, s0)
    ctx.sta = {"s0": s0, "cols": cols, "typ": typ}
    ws.freeze_panes = f"E{s0}"
    return ws


# =========================================================================== Urlaub-Krank
def build_urlaub_krank(ctx):
    ws = ctx.wb[S_UK]
    ws.sheet_properties.tabColor = "00B050"
    title(ws, "B1", "Urlaubs- und Krankenstatistik")
    put(ws, "F1", '="Planungsjahr "&Jahr', bold=True, size=12)
    note(ws, "B2", "Arbeitstage (Mo–Fr ohne Feiertage) je Monat. Urlaub = Urlaub + Urlaub genehmigt eTime. "
         "Quote = Tage / Arbeitstage des Jahres. Freistellungstage werden nicht als Urlaub gezählt.")
    s0 = 6
    typ = "AH"
    put(ws, "E4", "Urlaubstage", bold=True, size=11, bg="00B050", h="center")
    put(ws, "T4", "Krankheitstage", bold=True, size=11, bg="FFFF00", color="FF0000", h="center")
    for c in [gl(i) for i in range(6, 19)]:
        ws[f"{c}4"].fill = fill("00B050")
    for c in [gl(i) for i in range(21, 34)]:
        ws[f"{c}4"].fill = fill("FFFF00")
    hdr = {"B": "Name", "C": "Gruppe", "D": "Funktion"}
    for i, mk in enumerate(M.MON_KURZ):
        hdr[gl(5 + i)] = mk
        hdr[gl(20 + i)] = mk
    hdr["Q"], hdr["R"] = "Summe", "Quote"
    hdr["AF"], hdr["AG"] = "Summe", "Quote"
    for c, h in hdr.items():
        put(ws, f"{c}{s0 - 1}", h, bold=True, h="center", bg=M.HEADER_FILL, border=BOX, size=9)
    rows = mirror(ws, s0, "B", "C", "D", typ)
    last = s0 + M.CAP - 1
    for r, br in rows:
        for i in range(12):
            m = i + 1
            cu, ck = gl(5 + i), gl(20 + i)
            u = f"{code_cnt(br, 3, f',Ber_AT,1,Ber_Monat,{m}')}+{code_cnt(br, 4, f',Ber_AT,1,Ber_Monat,{m}')}"
            k = code_cnt(br, 8, f",Ber_AT,1,Ber_Monat,{m}")
            for c, per in ((cu, u), (ck, k)):
                grp = f'SUMIFS({c}${s0}:{c}${last},$C${s0}:$C${last},{BER}!$D{br},${typ}${s0}:${typ}${last},"P")'
                ws[f"{c}{r}"] = f'=IF(${typ}{r}="P",{per},IF(${typ}{r}="K",{grp},""))'
        for a, b, sq in (("E", "P", "Q"), ("T", "AE", "AF")):
            ws[f"{sq}{r}"] = f'=IF(${typ}{r}="","",SUM({a}{r}:{b}{r}))'
            qc = gl(ws[f"{sq}1"].column + 1)
            ws[f"{qc}{r}"] = (f'=IF(${typ}{r}="P",IF(SUM(Ber_AT)=0,"",{sq}{r}/SUM(Ber_AT)),IF(${typ}{r}="K",'
                              f'IFERROR({sq}{r}/(SUM(Ber_AT)*COUNTIFS($C${s0}:$C${last},{BER}!$D{br},'
                              f'${typ}${s0}:${typ}${last},"P")),""),""))')
            ws[f"{qc}{r}"].number_format = "0.0%"
        for ci in list(range(5, 19)) + list(range(20, 34)):
            cell = ws.cell(r, ci)
            cell.font = font(9)
            cell.alignment = Alignment(horizontal="center")
            cell.border = BOX
            if cell.number_format == "General":
                cell.number_format = "0;-0;;@"
    rt = last + 1
    put(ws, f"B{rt}", "Summe Teams", bold=True, border=BOX)
    for ci in list(range(5, 18)) + list(range(20, 33)):
        c = gl(ci)
        put(ws, f"{c}{rt}", f'=SUMIFS({c}{s0}:{c}{last},${typ}{s0}:${typ}{last},"P",$C{s0}:$C{last},"<>Ressourcen")',
            bold=True, border=BOX, h="center", size=9)
    for qc, sq in (("R", "Q"), ("AG", "AF")):
        put(ws, f"{qc}{rt}", f'=IFERROR({sq}{rt}/(SUM(Ber_AT)*Bestand),"")', bold=True, border=BOX, h="center",
            nf="0.0%", size=9)
    widths(ws, {"A": 2, "B": 22, "C": 10, "D": 12, "S": 2, typ: 3})
    for ci in list(range(5, 18)) + list(range(20, 33)):
        ws.column_dimensions[gl(ci)].width = 5.5
    ws.column_dimensions["Q"].width = ws.column_dimensions["AF"].width = 7
    ws.column_dimensions["R"].width = ws.column_dimensions["AG"].width = 7
    rng = f"B{s0}:AG{last}"
    group_cf(ws, rng, typ, s0)
    cf(ws, f"E{s0}:P{last}", f'AND(${typ}{s0}="P",E{s0}>0)', dxf(bg="C6EFCE", fg="006100"))
    cf(ws, f"T{s0}:AE{last}", f'AND(${typ}{s0}="P",T{s0}>0)', dxf(bg="FFFF99", fg="FF0000"))
    ws.freeze_panes = f"E{s0}"
    return ws


# =========================================================================== Monatsbericht
def build_monat(ctx):
    ws = ctx.wb[S_MON]
    ws.sheet_properties.tabColor = "E97132"
    title(ws, "B1", "Monatsbericht")
    put(ws, "B3", "Monat:", bold=True, h="right")
    put(ws, "C3", "Januar", bg=INPUT_FILL, bold=True, border=BOX, h="center")
    dv_list(ws, "C3", "lstMonate", prompt="Monat wählen")
    put(ws, "D3", "=MATCH(C3,lstMonate,0)", color="FFFFFF")
    put(ws, "E3", '="Planungsjahr "&Jahr', bold=True)
    ctx.name("MB_Monat", f"{S_MON}!$D$3")
    m = "MB_Monat"
    at = f"Bedarf_Monat,{m},Bedarf_AT,1"
    team_cells = "(Ber_Kat=\"Team\")"
    sp = lambda codes: ("SUMPRODUCT((" + "+".join(f"(Ber_Grid={c})" for c in codes) +
                        f")*(Ber_Monat={m})*(Ber_AT=1)*{team_cells})")
    kpis = [
        ("Arbeitstage", f"=COUNTIFS({at})", "0"),
        ("Ø Ist Tagdienst (MA)", f'=IFERROR(AVERAGEIFS(Bedarf_Ist,{at}),"")', "0.0"),
        ("Sollbesetzung Tagdienst (MA)", "=SollTag", "0"),
        ("Tage mit Unterbesetzung", f'=COUNTIFS({at},Bedarf_Unter,">0")', "0"),
        ("Fehlende Personentage", f"=SUMIFS(Bedarf_Unter,{at})", "0"),
        ("Bereitschaft: Tage kritisch 🟡", f"=COUNTIFS(Bedarf_Monat,{m},Bedarf_BerStufe,1)", "0"),
        ("Bereitschaft: Tage nicht einsatzfähig 🔴", f"=COUNTIFS(Bedarf_Monat,{m},Bedarf_BerStufe,2)", "0"),
        ("Urlaubstage (Teams)", "=" + sp([3, 4]), "0"),
        ("Krankheitstage (Teams)", "=" + sp([8]), "0"),
        ("Unterrichtstage (Teams)", "=" + sp([7]), "0"),
        ("JAZ-Tage (Teams)", "=" + sp([2]), "0"),
        ("Einweisungen FDL/Kabel (Teams)", "=" + sp([5, 6]), "0"),
        ("Urlaubsquote", "=IFERROR(C13/(C5*Bestand),\"\")", "0.0%"),
        ("Krankenquote", "=IFERROR(C14/(C5*Bestand),\"\")", "0.0%"),
        ("Auslastung (Soll-/Ist-Personentage)", f'=IFERROR(SUMIFS(Bedarf_Soll,{at})/SUMIFS(Bedarf_Ist,{at}),"")', "0%"),
        ("Zusätzliche MA (Personentage)", f"=SUMIFS(Bedarf_Zus,Bedarf_Monat,{m})", "0"),
        ("AZUBIS (Personentage)", f"=SUMIFS(Bedarf_Azubi,Bedarf_Monat,{m})", "0"),
    ]
    put(ws, "B4", "Kennzahlen", bold=True, size=11)
    for i, (lab, f_, nf) in enumerate(kpis):
        r = 5 + i
        put(ws, f"B{r}", lab, border=BOX, size=9)
        put(ws, f"C{r}", f_, border=BOX, h="center", nf=nf, bold=True, size=9)
    ctx.mb_kpi = {lab: f"C{5 + i}" for i, (lab, _, _) in enumerate(kpis)}

    # Tagesuebersicht
    t0 = 4
    cols = ["Datum", "WT", "Feiertag", "Nachtteam", "Freigestellt", "Abwesend", "Ist Tagdienst", "Soll",
            "Ampel", "Bereitschaft", "Fehlermeldung / Ersatzbedarf", "Zus. MA", "AZUBIS"]
    L = [gl(5 + i) for i in range(len(cols))]
    put(ws, f"{L[0]}{t0}", "Tagesübersicht", bold=True, size=11)
    for c, h in zip(L, cols):
        put(ws, f"{c}{t0 + 1}", h, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9, wrap=True)
    for j in range(1, 32):
        r = t0 + 1 + j
        d = f"${L[0]}{r}"
        idx = f"({d}-DATE(Jahr,1,1)+1)"
        ws[f"{L[0]}{r}"] = f'=IF({j}>DAY(DATE(Jahr,{m}+1,0)),"",DATE(Jahr,{m},{j}))'
        ws[f"{L[1]}{r}"] = f'=IF({d}="","",INDEX(lstWT,WEEKDAY({d},2)))'
        ws[f"{L[2]}{r}"] = f'=IF({d}="","",INDEX(Schicht_Feiertag,{idx}))'
        ws[f"{L[3]}{r}"] = f'=IF({d}="","",INDEX(Schicht_Nacht,{idx}))'
        ws[f"{L[4]}{r}"] = f'=IF({d}="","",INDEX(Schicht_FSwirksam,{idx}))'
        ws[f"{L[5]}{r}"] = f'=IF({d}="","",INDEX(Bedarf_Abw,{idx}))'
        ws[f"{L[6]}{r}"] = f'=IF({d}="","",IF(INDEX(Bedarf_AT,{idx})=1,INDEX(Bedarf_Ist,{idx}),""))'
        ws[f"{L[7]}{r}"] = f'=IF({d}="","",IF(INDEX(Bedarf_AT,{idx})=1,SollTag,""))'
        ws[f"{L[8]}{r}"] = f'=IF({d}="","",INDEX(tblBedarf[Ampel Tagdienst],{idx}))'
        ws[f"{L[9]}{r}"] = f'=IF({d}="","",INDEX(Einsatz_Ampel,{idx}))'
        ws[f"{L[10]}{r}"] = (f'=IF({d}="","",INDEX(Einsatz_Fehler,{idx})&IF(INDEX(tblEinsatz[Ersatz TVT],{idx})&'
                             f'INDEX(tblEinsatz[Ersatz AV],{idx})&INDEX(tblEinsatz[Ersatz Meister],{idx})&'
                             f'INDEX(tblEinsatz[Ersatz Handwerker],{idx})="",""," → Ersatz: "&_xlfn.TEXTJOIN(", ",TRUE,'
                             f'INDEX(tblEinsatz[Ersatz Meister],{idx}),INDEX(tblEinsatz[Ersatz AV],{idx}),'
                             f'INDEX(tblEinsatz[Ersatz TVT],{idx}),INDEX(tblEinsatz[Ersatz Handwerker],{idx}))))')
        ws[f"{L[11]}{r}"] = f'=IF({d}="","",INDEX(Bedarf_Zus,{idx}))'
        ws[f"{L[12]}{r}"] = f'=IF({d}="","",INDEX(Bedarf_Azubi,{idx}))'
        for c in L:
            ws[f"{c}{r}"].font = font(9)
            ws[f"{c}{r}"].border = BOX
            ws[f"{c}{r}"].alignment = Alignment(horizontal="left" if c == L[10] else "center", vertical="center")
        ws[f"{L[0]}{r}"].number_format = "DD.MM."
        for c in (L[11], L[12]):
            ws[f"{c}{r}"].number_format = "0;-0;;@"
    tl = t0 + 32
    tr_ = f"{L[0]}{t0 + 2}:{L[-1]}{tl}"
    cf(ws, f"{L[0]}{t0 + 2}:{L[2]}{tl}", f'${L[2]}{t0 + 2}<>""', dxf(bg=M.HOLIDAY_FILL, fg="CC0000", bold=True))
    cf(ws, f"{L[0]}{t0 + 2}:{L[2]}{tl}", f'OR(${L[1]}{t0 + 2}="Sa",${L[1]}{t0 + 2}="So")', dxf(bg=M.WEEKEND_FILL))
    for t in M.TEAMS:
        cf(ws, f"{L[3]}{t0 + 2}:{L[3]}{tl}", f'{L[3]}{t0 + 2}="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
        cf(ws, f"{L[4]}{t0 + 2}:{L[4]}{tl}", f'ISNUMBER(SEARCH("{t}",{L[4]}{t0 + 2}))', dxf(bg=M.TEAM_LIGHT[t]))
    cf(ws, f"{L[9]}{t0 + 2}:{L[9]}{tl}", f'LEFT({L[9]}{t0 + 2},2)="🔴"', dxf(bg="FFC7CE"))
    cf(ws, f"{L[9]}{t0 + 2}:{L[9]}{tl}", f'LEFT({L[9]}{t0 + 2},2)="🟡"', dxf(bg="FFEB9C"))
    cf(ws, f"{L[11]}{t0 + 2}:{L[11]}{tl}", f'AND(ISNUMBER({L[11]}{t0 + 2}),{L[11]}{t0 + 2}>0)', dxf(bg="FFFF00"))
    cf(ws, f"{L[12]}{t0 + 2}:{L[12]}{tl}", f'AND(ISNUMBER({L[12]}{t0 + 2}),{L[12]}{t0 + 2}>0)', dxf(bg="00B0F0"))
    ctx.mb_daily = {"first": t0 + 2, "last": tl, "date": L[0], "ist": L[6], "soll": L[7]}

    # Teams im Monat
    r0 = 24
    put(ws, f"B{r0}", "Teams im Monat", bold=True, size=11)
    put(ws, f"B{r0 + 1}", "Team", bold=True, bg=M.HEADER_FILL, border=BOX, size=9)
    put(ws, f"C{r0 + 1}", "Bereitschaftstage / 🔴", bold=True, bg=M.HEADER_FILL, border=BOX, size=9, wrap=True)
    for i, t in enumerate(M.TEAMS):
        r = r0 + 2 + i
        put(ws, f"B{r}", t, bold=True, border=BOX, bg=M.TEAM_COLOR[t], color=M.TEAM_FONT[t], size=9)
        put(ws, f"C{r}", f'=COUNTIFS(tblEinsatz[Bereitschaftsteam],"{t}",Bedarf_Monat,{m})&" / "&'
                         f'COUNTIFS(tblEinsatz[Bereitschaftsteam],"{t}",Bedarf_Monat,{m},Einsatz_Stufe,2)',
            border=BOX, h="center", size=9)
    ws.row_dimensions[r0 + 1].height = 26

    # Mitarbeiter im Monat
    p0 = 40
    put(ws, f"B{p0}", "Mitarbeiter im Monat (Arbeitstage)", bold=True, size=11)
    ph = ["Name", "Gruppe", "Urlaub", "Krank", "Unterr.", "JAZ", "FDL/Kabel", "Nächte", "Freist."]
    pc = ["B", "C", "D", "E", "F", "G", "H", "I", "J"]
    for c, h in zip(pc, ph):
        put(ws, f"{c}{p0 + 1}", h, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
    typ = "K"
    s0 = p0 + 2
    rows = mirror(ws, s0, "B", "C", None, typ, size=8)
    for r, br in rows:
        ex = f",Ber_AT,1,Ber_Monat,{m}"
        vals = {"D": f"{code_cnt(br, 3, ex)}+{code_cnt(br, 4, ex)}", "E": code_cnt(br, 8, ex),
                "F": code_cnt(br, 7, ex), "G": code_cnt(br, 2, ex),
                "H": f"{code_cnt(br, 5, ex)}+{code_cnt(br, 6, ex)}",
                "I": on_cnt(br, f",Ber_Monat,{m}"), "J": code_cnt(br, 10, f",Ber_Monat,{m}")}
        for c, f_ in vals.items():
            ws[f"{c}{r}"] = f'=IF(${typ}{r}="P",{f_},"")'
            ws[f"{c}{r}"].font = font(8)
            ws[f"{c}{r}"].alignment = Alignment(horizontal="center")
            ws[f"{c}{r}"].border = BOX
            ws[f"{c}{r}"].number_format = "0;-0;;@"
    last = s0 + M.CAP - 1
    group_cf(ws, f"B{s0}:J{last}", typ, s0)
    widths(ws, {"A": 2, "B": 34, "C": 11, "D": 8, "E": 8, "F": 8, "G": 8, "H": 9, "I": 9, "J": 8, typ: 2})
    for c, w in zip(L, [8, 5, 16, 10, 12, 9, 9, 6, 7, 24, 60, 7, 7]):
        ws.column_dimensions[c].width = max(ws.column_dimensions[c].width or 0, w)
    from .style import landscape
    landscape(ws, "A4", 1, 0, area=f"A1:{L[-1]}{last}")
    return ws


# =========================================================================== Jahresbericht
def build_jahr(ctx):
    ws = ctx.wb[S_JAHR]
    ws.sheet_properties.tabColor = "E97132"
    title(ws, "B1", "Jahresbericht")
    put(ws, "E1", '="Planungsjahr "&Jahr', bold=True, size=12)
    note(ws, "B2", "Alle Werte aktualisieren sich automatisch. Quoten bezogen auf Personalbestand der Teams × Arbeitstage. "
         "Auslastung = Soll-Personentage / Ist-Personentage im Tagdienst (über 100 % = Unterbesetzung).")
    hdr = ["Monat", "Arbeitstage", "Ø Ist Tagdienst", "Min. Ist", "Tage Unterbesetzung", "Fehlende Personentage",
           "Bereitschaft 🟢 (Tage)", "Bereitschaft 🟡 (Tage)", "Bereitschaft 🔴 (Tage)", "Urlaubstage", "Krankheitstage",
           "Unterrichtstage", "JAZ-Tage", "Einweisungen", "Urlaubsquote", "Krankenquote", "Auslastung",
           "Zusätzliche MA (PT)", "AZUBIS (PT)"]
    L = [gl(2 + i) for i in range(len(hdr))]
    r0 = 4
    for c, h in zip(L, hdr):
        put(ws, f"{c}{r0}", h, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", wrap=True, size=9)
    ws.row_dimensions[r0].height = 42
    sp = lambda codes, mm: ("SUMPRODUCT((" + "+".join(f"(Ber_Grid={c})" for c in codes) +
                            f")*({mm})*(Ber_AT=1)*(Ber_Kat=\"Team\"))")
    for i in range(13):
        r = r0 + 1 + i
        if i < 12:
            mnum = i + 1
            ws[f"B{r}"] = f"=INDEX(lstMonate,{mnum})"
            at = f"Bedarf_Monat,{mnum},Bedarf_AT,1"
            cr = f"Bedarf_Monat,{mnum}"
            mm = f"Ber_Monat={mnum}"
        else:
            ws[f"B{r}"] = "Jahr"
            at = "Bedarf_AT,1"
            cr = 'Bedarf_Datum,">0"'
            mm = "Ber_Monat>0"
        F = {
            "Arbeitstage": f"=COUNTIFS({at})",
            "Ø Ist Tagdienst": f'=IFERROR(AVERAGEIFS(Bedarf_Ist,{at}),"")',
            "Min. Ist": f'=IF(COUNTIFS({at})=0,"",_xlfn.MINIFS(Bedarf_Ist,{at}))',
            "Tage Unterbesetzung": f'=COUNTIFS({at},Bedarf_Unter,">0")',
            "Fehlende Personentage": f"=SUMIFS(Bedarf_Unter,{at})",
            "Bereitschaft 🟢 (Tage)": f"=COUNTIFS({cr},Bedarf_BerStufe,0)",
            "Bereitschaft 🟡 (Tage)": f"=COUNTIFS({cr},Bedarf_BerStufe,1)",
            "Bereitschaft 🔴 (Tage)": f"=COUNTIFS({cr},Bedarf_BerStufe,2)",
            "Urlaubstage": "=" + sp([3, 4], mm),
            "Krankheitstage": "=" + sp([8], mm),
            "Unterrichtstage": "=" + sp([7], mm),
            "JAZ-Tage": "=" + sp([2], mm),
            "Einweisungen": "=" + sp([5, 6], mm),
            "Urlaubsquote": f'=IFERROR({L[9]}{r}/({L[1]}{r}*Bestand),"")',
            "Krankenquote": f'=IFERROR({L[10]}{r}/({L[1]}{r}*Bestand),"")',
            "Auslastung": f'=IFERROR(SUMIFS(Bedarf_Soll,{at})/SUMIFS(Bedarf_Ist,{at}),"")',
            "Zusätzliche MA (PT)": f"=SUMIFS(Bedarf_Zus,{cr})",
            "AZUBIS (PT)": f"=SUMIFS(Bedarf_Azubi,{cr})",
        }
        for c, h in zip(L[1:], hdr[1:]):
            nf = {"Ø Ist Tagdienst": "0.0", "Urlaubsquote": "0.0%", "Krankenquote": "0.0%", "Auslastung": "0%"}.get(h, "0")
            put(ws, f"{c}{r}", F[h], border=BOX, h="center", nf=nf, size=9, bold=i == 12)
        put(ws, f"B{r}", None, border=BOX, bold=True, size=9)
    ctx.jb = {"r_first": r0 + 1, "r_last": r0 + 12, "r_year": r0 + 13, "cols": dict(zip(hdr, L))}
    widths(ws, {"A": 2, "B": 12})
    for c in L[1:]:
        ws.column_dimensions[c].width = 10.5

    # Dienstverteilung je Team x Monat
    d0 = r0 + 16
    put(ws, f"B{d0}", "Dienstverteilung je Team und Monat", bold=True, size=11)
    blocks = [("Bereitschaftstage (Nachtdienst)", "B"), ("Freistellungstage (Vorschlag)", "F"),
              ("Tagdienst-Personentage", "T")]
    r = d0 + 1
    for lab, kind in blocks:
        put(ws, f"B{r}", lab, bold=True, size=10, color="404040")
        r += 1
        put(ws, f"B{r}", "Team", bold=True, bg=M.HEADER_FILL, border=BOX, size=9)
        for i, mk in enumerate(M.MON_KURZ):
            put(ws, f"{gl(3 + i)}{r}", mk, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
        put(ws, f"{gl(15)}{r}", "Jahr", bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
        r += 1
        for t in M.TEAMS:
            put(ws, f"B{r}", t, bold=True, bg=M.TEAM_COLOR[t], color=M.TEAM_FONT[t], border=BOX, size=9)
            for i in range(12):
                mnum = i + 1
                if kind == "B":
                    f_ = f'=COUNTIFS(Schicht_Nacht,"{t}",Bedarf_Monat,{mnum})'
                elif kind == "F":
                    f_ = f'=COUNTIFS(Schicht_FSauto,"{t}",Bedarf_Monat,{mnum})'
                else:
                    f_ = (f'=SUMPRODUCT(((Ber_Grid=0)+(Ber_Grid=1)+(Ber_Grid=9)+(Ber_Grid=11))*(Ber_Gruppe="{t}")'
                          f'*(Ber_Typ="P")*(Ber_Monat={mnum})*(Ber_AT=1))')
                put(ws, f"{gl(3 + i)}{r}", f_, border=BOX, h="center", size=9)
            put(ws, f"{gl(15)}{r}", f"=SUM(C{r}:N{r})", border=BOX, h="center", bold=True, size=9)
            r += 1
        r += 1
    ctx.jb["dv_first"] = d0 + 3

    # Nachtdienstverteilung je MA x Monat
    n0 = r + 1
    put(ws, f"B{n0}", "Nachtdienstverteilung je Mitarbeiter (Bereitschaftsnächte)", bold=True, size=11)
    put(ws, f"B{n0 + 1}", "Name", bold=True, bg=M.HEADER_FILL, border=BOX, size=9)
    for i, mk in enumerate(M.MON_KURZ):
        put(ws, f"{gl(3 + i)}{n0 + 1}", mk, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
    put(ws, f"O{n0 + 1}", "Jahr", bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
    put(ws, f"P{n0 + 1}", "Wochenende", bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
    put(ws, f"Q{n0 + 1}", "Feiertage", bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
    typ = "S"
    s0 = n0 + 2
    for i in range(M.CAP):
        rr, br = s0 + i, M.ROW0 + i
        ws[f"B{rr}"] = f"={BER}!$B{br}"
        ws[f"{typ}{rr}"] = f"={BER}!$A{br}"
        ws[f"{typ}{rr}"].font = font(7, color="FFFFFF")
        ws[f"T{rr}"] = f"={BER}!$D{br}"
        ws[f"T{rr}"].font = font(7, color="FFFFFF")
        ws[f"B{rr}"].font = font(8)
        ws[f"B{rr}"].border = BOX
        for k in range(12):
            c = gl(3 + k)
            ws[f"{c}{rr}"] = f'=IF(${typ}{rr}="P",{on_cnt(br, f",Ber_Monat,{k + 1}")},"")'
        ws[f"O{rr}"] = f'=IF(${typ}{rr}="P",SUM(C{rr}:N{rr}),"")'
        ws[f"P{rr}"] = f'=IF(${typ}{rr}="P",{on_cnt(br, ",Ber_WE,1")},"")'
        ws[f"Q{rr}"] = f'=IF(${typ}{rr}="P",{on_cnt(br, ",Ber_FT,1")},"")'
        for ci in range(3, 18):
            cell = ws.cell(rr, ci)
            cell.font = font(8, ci == 15)
            cell.alignment = Alignment(horizontal="center")
            cell.border = BOX
            cell.number_format = "0;-0;;@"
    last = s0 + M.CAP - 1
    group_cf(ws, f"B{s0}:Q{last}", typ, s0)
    ws.column_dimensions[typ].width = 2
    ws.column_dimensions["T"].width = 2
    from .style import landscape
    landscape(ws, "A4", 1, 0, area=f"A1:Q{last}")
    ws.freeze_panes = "C5"
    return ws


# =========================================================================== Teamuebersicht
def build_team(ctx):
    ws = ctx.wb[S_TEAM]
    ws.sheet_properties.tabColor = "4F81BD"
    title(ws, "B1", "Teamübersicht und Einsatzfähigkeitsbericht")
    put(ws, "F1", '="Planungsjahr "&Jahr', bold=True, size=12)
    note(ws, "B2", "Zusammensetzung aus dem Mitarbeiterstamm, Qualifikationen aus der Funktionsmatrix. "
         "Kennzahlen aus Schichtplanung, Einsatzfähigkeitsprüfung und Kalender.")
    width = 5
    quals = M.QUALIS[4:]
    for ti, t in enumerate(M.TEAMS):
        c0 = 2 + ti * (width + 1)
        cs = [gl(c0 + i) for i in range(width)]
        r = 4
        put(ws, f"{cs[0]}{r}", f"Team {t}", bold=True, size=12, bg=M.TEAM_COLOR[t], color=M.TEAM_FONT[t])
        for c in cs[1:]:
            ws[f"{c}{r}"].fill = fill(M.TEAM_COLOR[t])
        r += 1
        for c, h in zip(cs, ["Name", "Funktion", "Nächte", "Urlaub", "Krank"]):
            put(ws, f"{c}{r}", h, bold=True, bg=M.HEADER_FILL, border=BOX, size=8, h="center")
        gi = M.GROUPS.index(t) + 1
        for j in range(1, 9):
            rr = r + j
            idx = f"INDEX(Grp_Start,{gi})+{j}"
            ok = f"{j}<=INDEX(Grp_Anzahl,{gi})"
            ws[f"{cs[0]}{rr}"] = f'=IF({ok},INDEX(Ber_Name,{idx}),"")'
            ws[f"{cs[1]}{rr}"] = f'=IF({ok},INDEX(Ber_Funktion,{idx}),"")'
            row = f"INDEX(Ber_Grid,{idx},0)"
            ws[f"{cs[2]}{rr}"] = f'=IF({ok},COUNTIFS({row},">=9",{row},"<=11",{row},"<>10"),"")'
            ws[f"{cs[3]}{rr}"] = f'=IF({ok},COUNTIFS({row},3,Ber_AT,1)+COUNTIFS({row},4,Ber_AT,1),"")'
            ws[f"{cs[4]}{rr}"] = f'=IF({ok},COUNTIFS({row},8,Ber_AT,1),"")'
            for c in cs:
                ws[f"{c}{rr}"].font = font(8)
                ws[f"{c}{rr}"].border = BOX
                ws[f"{c}{rr}"].alignment = Alignment(horizontal="center" if c not in cs[:2] else None)
        r = r + 10
        kp = [
            ("Bestand", f"=INDEX(Grp_Anzahl,{gi})"),
            ("Bereitschaftstage", f'=COUNTIF(Schicht_Nacht,"{t}")'),
            ("Wochenend-Blöcke (Fr–So)", f'=COUNTIFS(Schicht_Nacht,"{t}",Schicht_WT,"Fr")'),
            ("Feiertagsdienste", f'=SUMPRODUCT((Schicht_Nacht="{t}")*(Schicht_Feiertag<>""))'),
            ("Freistellungstage (Vorschlag)", f'=COUNTIF(Schicht_FSauto,"{t}")'),
            ("Tage 🟢 Sollbesetzung", f'=COUNTIFS(Einsatz_Team,"{t}",Einsatz_Stufe,0)'),
            ("Tage 🟡 kritisch", f'=COUNTIFS(Einsatz_Team,"{t}",Einsatz_Stufe,1)'),
            ("Tage 🔴 nicht einsatzfähig", f'=COUNTIFS(Einsatz_Team,"{t}",Einsatz_Stufe,2)'),
            ("Vertretungstage durch andere", f'=SUMPRODUCT((Ber_Grid>=9)*(Ber_Grid<=11)*(Ber_Grid<>10)*(Ber_Gruppe<>"{t}")'
                                             f'*(Ber_Nacht="{t}"))'),
            ("Abwesenheitstage (Arbeitstage)", f'=SUMPRODUCT((Ber_Grid>=2)*(Ber_Grid<=8)*(Ber_Gruppe="{t}")*(Ber_AT=1))'),
        ]
        for i, (lab, f_) in enumerate(kp):
            rr = r + i
            put(ws, f"{cs[0]}{rr}", lab, size=8, border=BOX)
            ws.merge_cells(f"{cs[0]}{rr}:{cs[2]}{rr}")
            put(ws, f"{cs[3]}{rr}", f_, size=9, bold=True, border=BOX, h="center")
            ws.merge_cells(f"{cs[3]}{rr}:{cs[4]}{rr}")
        r = r + len(kp) + 1
        put(ws, f"{cs[0]}{r}", "Qualifikationen im Team", bold=True, size=8)
        for i, qn in enumerate(quals):
            rr = r + 1 + i
            col_ = f"tblFunktionsmatrix[{qn}]"
            put(ws, f"{cs[0]}{rr}", qn, size=8, border=BOX)
            ws.merge_cells(f"{cs[0]}{rr}:{cs[2]}{rr}")
            put(ws, f"{cs[3]}{rr}", f'=COUNTIFS(tblFunktionsmatrix[Team],"{t}",{col_},"JA")', size=8, border=BOX,
                h="center")
            ws.merge_cells(f"{cs[3]}{rr}:{cs[4]}{rr}")
        for c, w in zip(cs, [16, 11, 7, 7, 7]):
            ws.column_dimensions[c].width = w
        ws.column_dimensions[gl(c0 + width)].width = 2
    # Einsatzfaehigkeitsbericht je Monat
    e0 = 4 + 1 + 9 + len(kp) + 2 + len(quals) + 3
    put(ws, f"B{e0}", "Einsatzfähigkeitsbericht je Monat (Bereitschaftstage 🟢 / 🟡 / 🔴)", bold=True, size=11)
    put(ws, f"B{e0 + 1}", "Team", bold=True, bg=M.HEADER_FILL, border=BOX, size=9)
    for i, mk in enumerate(M.MON_KURZ):
        put(ws, f"{gl(3 + i)}{e0 + 1}", mk, bold=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
    for ti, t in enumerate(M.TEAMS):
        rr = e0 + 2 + ti
        put(ws, f"B{rr}", t, bold=True, bg=M.TEAM_COLOR[t], color=M.TEAM_FONT[t], border=BOX, size=9)
        for i in range(12):
            mnum = i + 1
            f_ = (f'=COUNTIFS(Einsatz_Team,"{t}",Bedarf_Monat,{mnum},Einsatz_Stufe,0)&" / "&'
                  f'COUNTIFS(Einsatz_Team,"{t}",Bedarf_Monat,{mnum},Einsatz_Stufe,1)&" / "&'
                  f'COUNTIFS(Einsatz_Team,"{t}",Bedarf_Monat,{mnum},Einsatz_Stufe,2)')
            put(ws, f"{gl(3 + i)}{rr}", f_, border=BOX, h="center", size=8)
            cf(ws, f"{gl(3 + i)}{rr}", f'COUNTIFS(Einsatz_Team,"{t}",Bedarf_Monat,{mnum},Einsatz_Stufe,2)>0',
               dxf(bg="FFC7CE", fg="9C0006"))
    # Liste kritischer Tage
    l0 = e0 + 9
    put(ws, f"B{l0}", "Kritische Tage (🔴 nicht einsatzfähig) – die ersten 40 im Jahr", bold=True, size=11)
    lh = ["Nr", "Datum", "Team", "Fehlermeldung", "Ersatzvorschläge"]
    lc = ["B", "C", "D", "E", "H"]
    for c, h in zip(lc, lh):
        put(ws, f"{c}{l0 + 1}", h, bold=True, bg=M.HEADER_FILL, border=BOX, size=9)
    for k in range(1, 41):
        rr = l0 + 1 + k
        idx = f"SUMPRODUCT(SMALL(ROW(Einsatz_Stufe)-ROW(tblEinsatz[#Headers])+(Einsatz_Stufe<>2)*100000,{k}))"
        put(ws, f"B{rr}", k, size=8, border=BOX, h="center")
        ws[f"C{rr}"] = f'=IF({idx}>=100000,"",INDEX(Einsatz_Datum,{idx}))'
        ws[f"D{rr}"] = f'=IF(C{rr}="","",INDEX(Einsatz_Team,C{rr}-DATE(Jahr,1,1)+1))'
        ws[f"E{rr}"] = f'=IF(C{rr}="","",INDEX(Einsatz_Fehler,C{rr}-DATE(Jahr,1,1)+1))'
        i_ = f"C{rr}-DATE(Jahr,1,1)+1"
        ws[f"H{rr}"] = (f'=IF(C{rr}="","",_xlfn.TEXTJOIN("; ",TRUE,INDEX(tblEinsatz[Ersatz Meister],{i_}),'
                        f'INDEX(tblEinsatz[Ersatz AV],{i_}),INDEX(tblEinsatz[Ersatz TVT],{i_}),'
                        f'INDEX(tblEinsatz[Ersatz Handwerker],{i_})))')
        ws[f"C{rr}"].number_format = "DDD DD.MM.YYYY"
        for c in "CDEH":
            ws[f"{c}{rr}"].font = font(8)
            ws[f"{c}{rr}"].border = BOX
    for t in M.TEAMS:
        cf(ws, f"D{l0 + 2}:D{l0 + 41}", f'D{l0 + 2}="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
    from .style import landscape
    landscape(ws, "A4", 1, 0, area=f"A1:{gl(2 + 5 * (width + 1))}{l0 + 41}")
    return ws


def build_all(ctx):
    build_statistik(ctx)
    build_urlaub_krank(ctx)
    build_monat(ctx)
    build_jahr(ctx)
    build_team(ctx)
