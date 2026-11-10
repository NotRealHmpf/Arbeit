"""Kern: Schichtplanung (Blatt 5), verstecktes Berechnungsraster, Jahreskalender (Blatt 4)."""
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

from . import model as M
from .ctx import S_BER, S_KAL, S_SCH
from .style import (BOX, INPUT_FILL, NOBORDER, THIN, add_table, cf, dv_list, dxf, fill, font, note, put, q,
                    title, widths)

C0, CN = M.C_DAY0, M.C_DAYN
R0, RN = M.ROW0, M.ROWN
BER = q(S_BER)
KAL = q(S_KAL)


def team_of(x):
    """Formel: Nachtteam fuer Datum x (Rotation ab RotStart, 10 Bloecke)."""
    return f"INDEX(RotTeam,MOD(2*INT(({x}-RotStart)/7)+(MOD({x}-RotStart,7)>=3),10)+1)"


# ======================================================================== Schichtplanung (Blatt 5)
SCH_HDR = ["Nr", "Datum", "Wochentag", "KW", "Monat", "Feiertag", "Arbeitstag", "Block", "Block-Nr",
           "Nachtdienstteam", "Bereitschaftsteam", "Bereitschaftsart", "Freigestellt (Vorschlag)",
           "Freigestellt (wirksam)", "Individuelle Anpassungen", "Tagdienstteams", "Tagdienst (Kurz)",
           "Wochenendrotation", "Feiertagsrotation", "Bemerkung", "Hinweis"]
SCH_ROW0 = 7


def sch_col(h):
    return gl(2 + SCH_HDR.index(h))


def build_schicht(ctx):
    ws = ctx.wb[S_SCH]
    ws.sheet_properties.tabColor = "F79646"
    title(ws, "B1", "Blatt 5 – Schicht- und Bereitschaftsplanung")
    note(ws, "B2", "Automatisch aus Rotation, Kalender und Feiertagen Hessen. Tagdienst Mo–Fr 06:45–15:30, "
         "Bereitschaft/Nachtdienst 15:30–06:45 (an Wochenenden und Feiertagen ganztägig). Planung in ganzen Tagen.")
    note(ws, "B3", "Freistellung (Vorschlag): nach Mo–Do → Freitag, nach Fr–So → Montag + Dienstag. "
         "„wirksam“ berücksichtigt Team-Anpassungen aus dem Blatt Freistellungen; Einzelanpassungen je MA "
         "stehen im Jahreskalender.")
    hr = SCH_ROW0 - 1
    for i, h in enumerate(SCH_HDR):
        put(ws, f"{gl(2 + i)}{hr}", h, bold=True, wrap=True)
    c = {h: sch_col(h) for h in SCH_HDR}
    teams = M.TEAMS
    for k in range(1, M.NDAYS + 1):
        r = SCH_ROW0 + k - 1
        D = f"${c['Datum']}{r}"
        wd = f"WEEKDAY({D},2)"
        F = {}
        F["Nr"] = k
        F["Datum"] = f'=IF($B{r}>DATE(Jahr,12,31)-DATE(Jahr,1,1)+1,"",DATE(Jahr,1,1)+$B{r}-1)'
        F["Wochentag"] = f'=IF({D}="","",INDEX(lstWT,{wd}))'
        F["KW"] = f'=IF({D}="","",_xlfn.ISOWEEKNUM({D}))'
        F["Monat"] = f'=IF({D}="","",INDEX(lstMonate,MONTH({D})))'
        F["Feiertag"] = f'=IF({D}="","",IFERROR(INDEX(FT_Name,MATCH({D},FT_Aktiv,0)),""))'
        ft = f"${c['Feiertag']}{r}"
        F["Arbeitstag"] = f'=IF({D}="","",IF(AND({wd}<=5,{ft}=""),"Ja","Nein"))'
        F["Block"] = f'=IF({D}="","",IF(MOD({D}-RotStart,7)>=3,"Mo–Do","Fr–So"))'
        F["Block-Nr"] = f'=IF({D}="","",MOD(2*INT(({D}-RotStart)/7)+(MOD({D}-RotStart,7)>=3),10)+1)'
        F["Nachtdienstteam"] = f'=IF({D}="","",INDEX(RotTeam,${c["Block-Nr"]}{r}))'
        F["Bereitschaftsteam"] = f'=IF({D}="","",${c["Nachtdienstteam"]}{r})'
        F["Bereitschaftsart"] = (f'=IF({D}="","",IF({ft}<>"","Feiertag ganztägig + Nacht",IF({wd}>=6,'
                                 f'"Wochenende ganztägig + Nacht","Nacht 15:30–06:45")))')
        at = f"${c['Arbeitstag']}{r}"
        F["Freigestellt (Vorschlag)"] = (f'=IF({D}="","",IF({at}<>"Ja","",IF(OR({wd}=1,{wd}=5),{team_of(D + "-1")},'
                                         f'IF({wd}=2,{team_of(D + "-2")},""))))')
        auto = f"${c['Freigestellt (Vorschlag)']}{r}"
        parts = [f'IF(AND({auto}<>"",COUNTIFS(FS_Team,{auto},FS_MA,"",FS_Entf,{D})=0),{auto},"")']
        for t in teams:
            parts.append(f'IF(AND({auto}<>"{t}",COUNTIFS(FS_Team,"{t}",FS_MA,"",FS_Statt,{D})>0),"{t}","")')
        F["Freigestellt (wirksam)"] = f'=IF({D}="","",_xlfn.TEXTJOIN(", ",TRUE,{",".join(parts)}))'
        n_ind = f'COUNTIFS(FS_MA,"<>",FS_Entf,{D})+COUNTIFS(FS_MA,"<>",FS_Statt,{D})'
        F["Individuelle Anpassungen"] = f'=IF({D}="","",IF({n_ind}=0,"",{n_ind}&" MA"))'
        eff = f"${c['Freigestellt (wirksam)']}{r}"
        tparts = ",".join(f'IF(ISNUMBER(SEARCH("{t}",{eff})),"","{t}")' for t in teams)
        F["Tagdienstteams"] = f'=IF({D}="","",IF({at}<>"Ja","– kein Tagdienst –",_xlfn.TEXTJOIN(", ",TRUE,{tparts})))'
        kparts = ",".join(f'IF(ISNUMBER(SEARCH("{t}",{eff})),"",INDEX(TeamKurz,{i + 1}))' for i, t in enumerate(teams))
        F["Tagdienst (Kurz)"] = f'=IF({D}="","",IF({at}<>"Ja","–",_xlfn.TEXTJOIN(" ",TRUE,{kparts})))'
        F["Wochenendrotation"] = f'=IF({D}="","",IF({wd}>=6,${c["Nachtdienstteam"]}{r},""))'
        F["Feiertagsrotation"] = f'=IF({D}="","",IF({ft}<>"",${c["Nachtdienstteam"]}{r},""))'
        F["Bemerkung"] = f'=IF({D}="","",INDEX({KAL}!${C0}$10:${CN}$10,$B{r})&"")'
        F["Hinweis"] = (f'=IF({D}="","",IF(AND(OR({wd}=1,{wd}=2,{wd}=5),{ft}<>""),'
                        f'"Freistellungstag fällt auf Feiertag – ggf. manuell verlegen",""))')
        for h in SCH_HDR:
            nf = {"Datum": "DDD DD.MM.YYYY"}.get(h)
            ws[f"{c[h]}{r}"] = F[h]
            ws[f"{c[h]}{r}"].font = font(9)
            if nf:
                ws[f"{c[h]}{r}"].number_format = nf
            ws[f"{c[h]}{r}"].alignment = Alignment(horizontal="left" if h in ("Bemerkung", "Hinweis",
                                                                                "Tagdienstteams") else "center")
    last = SCH_ROW0 + M.NDAYS - 1
    add_table(ws, "tblSchicht", f"B{hr}:{gl(1 + len(SCH_HDR))}{last}", style="TableStyleLight15")
    ws.row_dimensions[hr].height = 42
    wd_ = {"Nr": 5, "Datum": 15, "Wochentag": 6, "KW": 5, "Monat": 10, "Feiertag": 20, "Arbeitstag": 8,
           "Block": 8, "Block-Nr": 6, "Nachtdienstteam": 11, "Bereitschaftsteam": 11, "Bereitschaftsart": 24,
           "Freigestellt (Vorschlag)": 12, "Freigestellt (wirksam)": 13, "Individuelle Anpassungen": 11,
           "Tagdienstteams": 26, "Tagdienst (Kurz)": 13, "Wochenendrotation": 11, "Feiertagsrotation": 11,
           "Bemerkung": 24, "Hinweis": 44}
    for h, w in wd_.items():
        ws.column_dimensions[c[h]].width = w
    ws.column_dimensions["A"].width = 2
    rng = f"B{SCH_ROW0}:{gl(1 + len(SCH_HDR))}{last}"
    # Farben: Teams, Wochenende, Feiertage
    for h in ("Nachtdienstteam", "Bereitschaftsteam", "Wochenendrotation", "Feiertagsrotation",
              "Freigestellt (Vorschlag)"):
        cc = c[h]
        for t in teams:
            light = h == "Freigestellt (Vorschlag)"
            cf(ws, f"{cc}{SCH_ROW0}:{cc}{last}", f'{cc}{SCH_ROW0}="{t}"',
               dxf(bg=M.TEAM_LIGHT[t] if light else M.TEAM_COLOR[t], fg="000000" if light else M.TEAM_FONT[t],
                   bold=True))
    ce = c["Freigestellt (wirksam)"]
    for t in teams:
        cf(ws, f"{ce}{SCH_ROW0}:{ce}{last}", f'ISNUMBER(SEARCH("{t}",{ce}{SCH_ROW0}))', dxf(bg=M.TEAM_LIGHT[t], bold=True))
    cf(ws, f"{c['Hinweis']}{SCH_ROW0}:{c['Hinweis']}{last}", f'{c["Hinweis"]}{SCH_ROW0}<>""', dxf(fg="9C5700", bold=True))
    cf(ws, f"B{SCH_ROW0}:{c['Arbeitstag']}{last}", f'${c["Feiertag"]}{SCH_ROW0}<>""',
       dxf(bg=M.HOLIDAY_FILL, fg="CC0000", bold=True))
    cf(ws, f"B{SCH_ROW0}:{c['Arbeitstag']}{last}", f'WEEKDAY(${c["Datum"]}{SCH_ROW0},2)>=6', dxf(bg=M.WEEKEND_FILL))
    for nm, h in [("Schicht_Datum", "Datum"), ("Schicht_WT", "Wochentag"), ("Schicht_KW", "KW"),
                  ("Schicht_Monat", "Monat"), ("Schicht_Feiertag", "Feiertag"), ("Schicht_AT", "Arbeitstag"),
                  ("Schicht_Block", "Block"), ("Schicht_Nacht", "Nachtdienstteam"),
                  ("Schicht_FSauto", "Freigestellt (Vorschlag)"), ("Schicht_FSwirksam", "Freigestellt (wirksam)"),
                  ("Schicht_TagKurz", "Tagdienst (Kurz)"), ("Schicht_Tag", "Tagdienstteams"),
                  ("Schicht_Bem", "Bemerkung")]:
        ctx.name(nm, f"tblSchicht[{h}]")

    # Rotationsuebersicht je Team (rechts)
    x0 = 2 + len(SCH_HDR) + 1
    cols = [gl(x0 + i) for i in range(7)]
    put(ws, f"{cols[0]}{hr - 1}", "Rotationsübersicht im Planungsjahr", bold=True, size=11)
    labels = ["Team", "Mo–Do-Blöcke", "Fr–So-Blöcke (Wochenenden)", "Bereitschaftstage", "davon Sa/So",
              "davon Feiertage", "Freistellungstage (Vorschlag)"]
    for cc, lab in zip(cols, labels):
        put(ws, f"{cc}{hr}", lab, bold=True, bg=M.HEADER_FILL, border=BOX, wrap=True, h="center")
    for i, t in enumerate(teams):
        r = SCH_ROW0 + i
        put(ws, f"{cols[0]}{r}", t, bold=True, border=BOX, bg=M.TEAM_COLOR[t], color=M.TEAM_FONT[t], h="center")
        put(ws, f"{cols[1]}{r}", f'=COUNTIFS(Schicht_Nacht,"{t}",Schicht_WT,"Mo")', border=BOX, h="center")
        put(ws, f"{cols[2]}{r}", f'=COUNTIFS(Schicht_Nacht,"{t}",Schicht_WT,"Fr")', border=BOX, h="center")
        put(ws, f"{cols[3]}{r}", f'=COUNTIF(Schicht_Nacht,"{t}")', border=BOX, h="center")
        put(ws, f"{cols[4]}{r}", f'=COUNTIFS(Schicht_Nacht,"{t}",Schicht_WT,"Sa")+COUNTIFS(Schicht_Nacht,"{t}",Schicht_WT,"So")',
            border=BOX, h="center")
        put(ws, f"{cols[5]}{r}", f'=SUMPRODUCT((Schicht_Nacht="{t}")*(Schicht_Feiertag<>""))', border=BOX, h="center")
        put(ws, f"{cols[6]}{r}", f'=COUNTIF(Schicht_FSauto,"{t}")', border=BOX, h="center")
    r = SCH_ROW0 + len(teams)
    put(ws, f"{cols[0]}{r}", "Summe", bold=True, border=BOX, h="center")
    for cc in cols[1:]:
        put(ws, f"{cc}{r}", f"=SUM({cc}{SCH_ROW0}:{cc}{r - 1})", bold=True, border=BOX, h="center")
    for cc, w in zip(cols, [10, 10, 13, 12, 9, 10, 13]):
        ws.column_dimensions[cc].width = w
    ws.column_dimensions[gl(x0 - 1)].width = 3
    note(ws, f"{cols[0]}{r + 2}", "Die Folge Rot → Gelb → Grün → Orange → Blau wechselt je Block zwischen Fr–So und Mo–Do;")
    note(ws, f"{cols[0]}{r + 3}", "dadurch übernimmt jedes Team abwechselnd Wochenend- und Wochenblöcke (Zyklus 5 Wochen).")
    ws.freeze_panes = f"D{SCH_ROW0}"
    return ws


# ======================================================================== Berechnung (versteckt)
def build_berechnung(ctx):
    ws = ctx.wb[S_BER]
    ws.sheet_state = "hidden"
    ws.sheet_properties.tabColor = "000000"
    put(ws, "A1", "Berechnung – nicht bearbeiten", bold=True, color="C00000")
    labels = {1: "Blockstart (Mo/Fr)", 2: "Tag-Nr", 3: "Datum", 4: "Wochentag 1-7", 5: "Arbeitstag",
              6: "Feiertag", 7: "Nachtteam", 8: "Freistellung (Vorschlag)", 9: "Anz. FS-Anpassungen",
              10: "Anz. Vertretungen", 11: "Anz. Ressourcen", 12: "Fr/Sa/So", 13: "Monat"}
    for r, lab in labels.items():
        put(ws, f"E{r}", lab, size=8, h="right", color="595959")
    for k in range(1, M.NDAYS + 1):
        c = gl(M.DAY0 + k - 1)
        ws[f"{c}1"] = f'=IF({c}$3="",0,IF(OR({c}$4=1,{c}$4=5),1,0))'
        ws[f"{c}2"] = k
        ws[f"{c}3"] = f'=IF({c}$2>DATE(Jahr,12,31)-DATE(Jahr,1,1)+1,"",DATE(Jahr,1,1)+{c}$2-1)'
        ws[f"{c}3"].number_format = "DD.MM."
        ws[f"{c}4"] = f'=IF({c}$3="",0,WEEKDAY({c}$3,2))'
        ws[f"{c}5"] = f'=IF({c}$3="",0,IF({c}$4>5,0,IF(COUNTIF(FT_Aktiv,{c}$3)>0,0,1)))'
        ws[f"{c}6"] = f'=IF({c}$3="",0,IF(COUNTIF(FT_Aktiv,{c}$3)>0,1,0))'
        ws[f"{c}7"] = f'=IF({c}$3="","",INDEX(Schicht_Nacht,{c}$2))'
        ws[f"{c}8"] = f'=IF({c}$3="","",INDEX(Schicht_FSauto,{c}$2))'
        ws[f"{c}9"] = f'=IF({c}$3="",0,COUNTIF(FS_Entf,{c}$3)+COUNTIF(FS_Statt,{c}$3))'
        ws[f"{c}10"] = f'=IF({c}$3="",0,COUNTIFS(V_Von,"<="&{c}$3,V_Bis,">="&{c}$3))'
        ws[f"{c}11"] = f'=IF({c}$3="",0,COUNTIFS(R_Von,"<="&{c}$3,R_Bis,">="&{c}$3))'
        ws[f"{c}12"] = f'=IF({c}$3="",0,IF({c}$4>=5,1,0))'
        ws[f"{c}13"] = f'=IF({c}$3="",0,MONTH({c}$3))'

    A = M.B_COL
    g0 = M.B_GRP_COL0
    G = {k: gl(g0 + i) for i, k in enumerate(["Gruppe", "Anzahl", "Kopf", "Start", "Davor", "Titel"])}
    # Gruppentabelle (Zeilen 14-20) + Gesamtzeilen
    for k, lab in G.items():
        put(ws, f"{lab}13", k, bold=True, size=8)
    for i in range(len(M.GROUPS)):
        r = 14 + i
        ws[f"{G['Gruppe']}{r}"] = f"=INDEX(lstGruppen,{i + 1})"
        ws[f"{G['Anzahl']}{r}"] = f"=COUNTIF(MA_Gruppe,{G['Gruppe']}{r})"
        ws[f"{G['Kopf']}{r}"] = f"=IF({G['Anzahl']}{r}>0,1,0)"
        ws[f"{G['Start']}{r}"] = 1 if i == 0 else f"={G['Start']}{r - 1}+{G['Kopf']}{r - 1}+{G['Anzahl']}{r - 1}"
        ws[f"{G['Davor']}{r}"] = 0 if i == 0 else f"={G['Davor']}{r - 1}+{G['Anzahl']}{r - 1}"
        ws[f"{G['Titel']}{r}"] = f"=INDEX(GruppenTitel,{i + 1})"
    rl = 14 + len(M.GROUPS) - 1
    put(ws, f"{G['Gruppe']}{rl + 1}", "Zeilen belegt", size=8)
    ws[f"{G['Start']}{rl + 1}"] = f"={G['Start']}{rl}+{G['Kopf']}{rl}+{G['Anzahl']}{rl}-1"
    put(ws, f"{G['Gruppe']}{rl + 2}", "Kapazität", size=8)
    ws[f"{G['Start']}{rl + 2}"] = M.CAP
    total = f"${G['Start']}${rl + 1}"
    ctx.name("KalZeilenBelegt", f"{S_BER}!{total}")
    ctx.name("Grp_Start", f"{S_BER}!${G['Start']}$14:${G['Start']}${rl}")
    ctx.name("Grp_Anzahl", f"{S_BER}!${G['Anzahl']}$14:${G['Anzahl']}${rl}")
    ctx.name("KalKapazitaet", f"{S_BER}!${G['Start']}${rl + 2}")
    gs = lambda k: f"${G[k]}$14:${G[k]}${rl}"

    for lab, cc in A.items():
        put(ws, f"{cc}13", lab, bold=True, size=8)
    for col_, lab in zip("ABCDE", ["Typ", "Name", "Funktion", "Gruppe", "Kategorie"]):
        put(ws, f"{col_}13", lab, bold=True, size=8)
    for r in range(R0, RN + 1):
        kk = r - R0 + 1
        ws[f"{A['k']}{r}"] = kk
        ws[f"{A['Rang']}{r}"] = kk
        ws[f"{A['Gruppe-Nr']}{r}"] = f"=IF({A['k']}{r}>{total},0,MATCH({A['k']}{r},{gs('Start')},1))"
        ws[f"{A['j']}{r}"] = f"=IF({A['Gruppe-Nr']}{r}=0,0,{A['k']}{r}-INDEX({gs('Start')},{A['Gruppe-Nr']}{r}))"
        ws[f"A{r}"] = f'=IF({A["Gruppe-Nr"]}{r}=0,"",IF({A["j"]}{r}=0,"K","P"))'
        ws[f"{A['MA-Zeile']}{r}"] = (f'=IF($A{r}<>"P","",MATCH(SMALL(MA_Sort,INDEX({gs("Davor")},'
                                     f'{A["Gruppe-Nr"]}{r})+{A["j"]}{r}),MA_Sort,0))')
        ws[f"B{r}"] = (f'=IF($A{r}="","",IF($A{r}="K",INDEX({gs("Titel")},{A["Gruppe-Nr"]}{r}),'
                       f'INDEX(MA_Name,{A["MA-Zeile"]}{r})))')
        ws[f"C{r}"] = f'=IF($A{r}<>"P","",INDEX(MA_Funktion,{A["MA-Zeile"]}{r})&"")'
        ws[f"D{r}"] = f'=IF($A{r}="","",INDEX({gs("Gruppe")},{A["Gruppe-Nr"]}{r}))'
        ws[f"E{r}"] = f'=IF($A{r}<>"P","",INDEX(MA_Kat,{A["MA-Zeile"]}{r}))'
        for qn, fm, rolle in [("Q_Meister", "FM_Meister", "Meister"), ("Q_AV", "FM_AV", "AV"),
                              ("Q_TVT", "FM_TVT", "TVT-Fahrer"), ("Q_HW", "FM_HW", "Handwerker")]:
            ws[f"{A[qn]}{r}"] = (f'=IF($A{r}<>"P",0,IF(OR($C{r}="{rolle}",IFERROR(INDEX({fm},MATCH($B{r},FM_Name,0)),"")'
                                 f'="JA"),1,0))')
        ws[f"{A['Bereitschaftstage']}{r}"] = (f'=IF($A{r}<>"P",0,MIN(99,COUNTIFS(${C0}{r}:${CN}{r},">=9",'
                                              f'${C0}{r}:${CN}{r},"<=11",${C0}{r}:${CN}{r},"<>10")))')
        ra = r + (M.AROW0 - M.ROW0)
        for k in range(1, M.NDAYS + 1):
            c = gl(M.DAY0 + k - 1)
            d = f"{c}$3"
            ws[f"{c}{ra}"] = (f'=IF(OR($A{r}<>"P",{d}=""),0,_xlfn.MAXIFS(Abw_Code,Abw_MA,$B{r},Abw_Von,"<="&{d},'
                              f'Abw_Bis,">="&{d},Abw_Code,"<=8"))')
            a = f"{c}{ra}"
            team_fs = (f'IF({c}$9=0,AND($E{r}="Team",{c}$5=1,$D{r}={c}$8),'
                       f'OR(COUNTIFS(FS_MA,$B{r},FS_Statt,{d})+COUNTIFS(FS_Team,$D{r},FS_MA,"",FS_Statt,{d})>0,'
                       f'AND($E{r}="Team",{c}$5=1,$D{r}={c}$8,'
                       f'COUNTIFS(FS_MA,$B{r},FS_Entf,{d})+COUNTIFS(FS_Team,$D{r},FS_MA,"",FS_Entf,{d})=0)))')
            bereit = f'IF($D{r}={c}$7,TRUE,IF({c}$10=0,FALSE,COUNTIFS(V_MA,$B{r},V_Von,"<="&{d},V_Bis,">="&{d})>0))'
            res = (f'IF(AND($E{r}="Ressource",{c}$11>0),_xlfn.MAXIFS(R_Code,R_Name,$B{r},R_Von,"<="&{d},'
                   f'R_Bis,">="&{d}),0)')
            ws[f"{c}{r}"] = (f'=IF(OR($A{r}<>"P",{d}=""),0,IF({a}>=5,{a},IF({team_fs},10,IF({a}>=2,{a},'
                             f'IF({bereit},IF({a}=1,11,9),IF({a}=1,1,{res}))))))')
    ctx.name("Ber_Typ", f"{S_BER}!$A${R0}:$A${RN}")
    ctx.name("Ber_Name", f"{S_BER}!$B${R0}:$B${RN}")
    ctx.name("Ber_Funktion", f"{S_BER}!$C${R0}:$C${RN}")
    ctx.name("Ber_Gruppe", f"{S_BER}!$D${R0}:$D${RN}")
    ctx.name("Ber_Kat", f"{S_BER}!$E${R0}:$E${RN}")
    ctx.name("Ber_Grid", f"{S_BER}!${C0}${R0}:${CN}${RN}")
    for nm, r in [("Ber_Blockstart", 1), ("Ber_Tag", 2), ("Ber_Datum", 3), ("Ber_WT", 4), ("Ber_AT", 5),
                  ("Ber_FT", 6), ("Ber_Nacht", 7), ("Ber_WE", 12), ("Ber_Monat", 13)]:
        ctx.name(nm, f"{S_BER}!${C0}${r}:${CN}${r}")
    for qn in ["Q_Meister", "Q_AV", "Q_TVT", "Q_HW", "Bereitschaftstage", "Rang"]:
        nm = "Ber_" + qn.replace("Q_", "Q")
        ctx.name(nm, f"{S_BER}!${A[qn]}${R0}:${A[qn]}${RN}")
    ws.freeze_panes = f"{C0}{R0}"
    widths(ws, {"A": 4, "B": 22, "C": 12, "D": 10, "E": 10})
    return ws


# ======================================================================== Jahreskalender (Blatt 4)
HEADER_ROWS = {2: "Monat", 3: "Datum", 4: "Wochentag", 5: "KW", 6: "Feiertag", 7: "Nachtteam",
               8: "Freigestelltes Team", 9: "Tagdienstteams", 10: "Bemerkung", 11: "Bereitschaft einsatzfähig",
               12: "MA im Tagdienst"}
PLAN_ROWS = {RN + 2: ("Zusätzliche MA (Anzahl)", 12), RN + 3: ("AZUBIS (Anzahl)", 13),
             RN + 4: ("Mecka", None), RN + 5: ("TVT Frist", None), RN + 6: ("Arbeiten bei Tag", None),
             RN + 7: ("Arbeiten bei Nacht", None)}


def kal_code_ref(c_letter_index, r):
    return f"{gl(c_letter_index + M.K_OFF)}{r}"


def kal_grid_formula(ci, r):
    c = gl(ci)
    code = kal_code_ref(ci, r)
    bes = (f'IFERROR(INDEX(Abw_Bem,MATCH(_xlfn.MAXIFS(Abw_Nr,Abw_MA,$B{r},Abw_Von,"<="&{c}$3,Abw_Bis,">="&{c}$3,'
           f'Abw_Code,1),Abw_Nr,0))&"","")')
    return (f'=IF({code}=0,"",IF(OR({code}=1,{code}=11),IF({bes}="",INDEX(StatusText,1)&"",{bes}),'
            f'INDEX(StatusText,{code})&""))')


def kal_grid_formula_r1c1():
    off = M.K_OFF
    code = f"RC[{off}]"
    bes = ('IFERROR(INDEX(Abw_Bem,MATCH(MAXIFS(Abw_Nr,Abw_MA,RC2,Abw_Von,"<="&R3C,Abw_Bis,">="&R3C,'
           'Abw_Code,1),Abw_Nr,0))&"","")')
    return (f'=IF({code}=0,"",IF(OR({code}=1,{code}=11),IF({bes}="",INDEX(StatusText,1)&"",{bes}),'
            f'INDEX(StatusText,{code})&""))')


def build_kalender(ctx):
    ws = ctx.wb[S_KAL]
    ws.sheet_properties.tabColor = M.TITLE_FILL
    widths(ws, {"A": 2.5, "B": 24, "C": 12, "D": 5.5, "E": 1.5})
    for ci in range(M.DAY0, M.DAYN + 1):
        ws.column_dimensions[gl(ci)].width = 5.3
    # Kopf links
    put(ws, "B1", '="PEP "&Jahr&" – Jahreskalender"', size=13, bold=True, bg=M.TITLE_FILL, color="FFFFFF")
    put(ws, "C1", "Januar", bg=INPUT_FILL, h="center", border=BOX, size=9)
    ws["D1"] = (f'=HYPERLINK("#{KAL}!"&ADDRESS({R0},{M.DAY0 - 1}+DATE(Jahr,MATCH($C$1,lstMonate,0),1)'
                f'-DATE(Jahr,1,1)+1),"▶ los")')
    ws["D1"].font = font(8, True, "0563C1", underline="single")
    dv_list(ws, "C1", "lstMonate", prompt="Monat wählen und rechts auf „▶ los“ klicken")
    ws.row_dimensions[1].height = 26
    put(ws, "F1", ("Leere/weiße Felder = regulärer Tagdienst · Hellgrau = Bereitschaft/Nachtdienst · nur Abweichungen "
                   "werden eingetragen · Eingabe mit Makros: Doppelklick auf einen Tag oder Rechtsklick auf markierte "
                   "Tage (U, UE, K, JAZ, S, FDL, Kabel, FS, -, V, Z, A, Freitext = Besonderheit, leer = löschen) · "
                   "ohne Makros: Blatt „Abwesenheiten“"), size=9, italic=True, color="404040")
    ws["AW1"] = "=Makrostatus"
    ws["AW1"].font = font(9, True, "C00000")
    for r, lab in HEADER_ROWS.items():
        put(ws, f"B{r}", lab, bold=True, size=9, h="right", bg=M.HEADER_FILL, border=BOX)
        put(ws, f"C{r}", None, bg=M.HEADER_FILL, border=BOX)
        put(ws, f"D{r}", None, bg=M.HEADER_FILL, border=BOX)
    put(ws, "B10", "Bemerkung (Eingabe)", bold=True, size=9, h="right", bg=INPUT_FILL, border=BOX)
    put(ws, f"B13", "Name", bold=True, bg=M.HEADER_FILL, border=BOX)
    put(ws, f"C13", "Funktion", bold=True, bg=M.HEADER_FILL, border=BOX)
    put(ws, f"D13", "JAZ", bold=True, bg="00B0F0", border=BOX, h="center", size=9)
    heights = {2: 13, 3: 14, 4: 13, 5: 12, 6: 66, 7: 13, 8: 22, 9: 64, 10: 58, 11: 15, 12: 14, 13: 15}
    for r, h in heights.items():
        ws.row_dimensions[r].height = h
    for r in range(6, 13):
        ws.row_dimensions[r].outlineLevel = 1

    for ci in range(M.DAY0, M.DAYN + 1):
        c = gl(ci)
        ws[f"{c}2"] = f'=IF({c}$3="","",INDEX(lstMonKurz,MONTH({c}$3)))'
        ws[f"{c}3"] = f"={BER}!{c}3"
        ws[f"{c}4"] = f'=IF({c}$3="","",INDEX(lstWT,WEEKDAY({c}$3,2)))'
        ws[f"{c}5"] = f'=IF({c}$3="","",IF(OR(WEEKDAY({c}$3,2)=1,{BER}!{c}$2=1),_xlfn.ISOWEEKNUM({c}$3),""))'
        ws[f"{c}6"] = f'=IF({c}$3="","",IFERROR(INDEX(FT_Name,MATCH({c}$3,FT_Aktiv,0)),""))'
        ws[f"{c}7"] = f"={BER}!{c}7"
        ws[f"{c}8"] = f'=IF({c}$3="","",INDEX(Schicht_FSwirksam,{BER}!{c}$2))'
        ws[f"{c}9"] = f'=IF({c}$3="","",INDEX(Schicht_TagKurz,{BER}!{c}$2))'
        ws[f"{c}11"] = f'=IF({c}$3="","",INDEX(Einsatz_Kurz,{BER}!{c}$2))'
        ws[f"{c}12"] = f'=IF({c}$3="","",IF({BER}!{c}$5=1,INDEX(Bedarf_Ist,{BER}!{c}$2),""))'
        styles = {2: (8, False, 0), 3: (9, True, 0), 4: (8, False, 0), 5: (7, False, 0), 6: (7, True, 90),
                  7: (7, True, 0), 8: (7, True, 0), 9: (7, False, 90), 10: (7, False, 90), 11: (8, False, 0),
                  12: (8, True, 0)}
        for r, (sz, b, rot) in styles.items():
            cell = ws[f"{c}{r}"]
            cell.font = font(sz, b, "CC0000" if r == 6 else "000000")
            cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=rot,
                                       wrap_text=r in (6, 8, 10), shrink_to_fit=r in (7,))
            cell.fill = fill(INPUT_FILL if r == 10 else M.HEADER_FILL if r <= 5 else None)
            cell.border = BOX
        ws[f"{c}3"].number_format = "D"
        ws[f"{c}12"].number_format = "0;-0;;@"
        # Personenraster
        for r in range(R0, RN + 1):
            cell = ws[f"{c}{r}"]
            cell.value = kal_grid_formula(ci, r)
            cell.font = font(8, True)
            cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=False)
            cell.border = BOX
        # Ressourcen-/Planungszeilen
        for r, (lab, code) in PLAN_ROWS.items():
            cell = ws[f"{c}{r}"]
            if code:
                art = "Zusätzliche MA" if code == 12 else "AZUBIS"
                cell.value = (f'=IF({c}$3="","",IF({BER}!{c}$11=0,0,COUNTIF({BER}!{c}${R0}:{c}${RN},{code})'
                              f'+COUNTIFS(R_Art,"{art}",R_Stamm,0,R_Von,"<="&{c}$3,R_Bis,">="&{c}$3)))')
                cell.number_format = "0;-0;;@"
                cell.font = font(8, True)
            else:
                cell.font = font(7)
                cell.fill = fill(INPUT_FILL)
            cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=code is None)
            cell.border = BOX
    # Labels Personen
    for r in range(R0, RN + 1):
        ws[f"B{r}"] = f"={BER}!B{r}"
        ws[f"C{r}"] = f"={BER}!C{r}"
        ws[f"D{r}"] = (f'=IF({BER}!$A{r}="P",COUNTIFS({BER}!${C0}{r}:${CN}{r},2,{BER}!${C0}$5:${CN}$5,1),"")')
        ws[f"D{r}"].number_format = "0;-0;;@"
        for cc, sz, b in (("A", 8, False), ("B", 9, False), ("C", 8, False), ("D", 8, False)):
            ws[f"{cc}{r}"].font = font(sz, b)
            ws[f"{cc}{r}"].border = BOX
            ws[f"{cc}{r}"].alignment = Alignment(vertical="center", horizontal="center" if cc == "D" else None)
        ws.row_dimensions[r].height = 15
        ws[f"{gl(M.K_GRP)}{r}"] = f"={BER}!$D{r}"
        ws[f"{gl(M.K_TYP)}{r}"] = f"={BER}!$A{r}"
        for ci in range(M.DAY0, M.DAYN + 1):
            ws[f"{gl(ci + M.K_OFF)}{r}"] = f"={BER}!{gl(ci)}{r}"
    for r, (lab, code) in PLAN_ROWS.items():
        bg = {"Mecka": "000000", "Arbeiten bei Nacht": "FFFF00"}.get(lab)
        bg = "FFFF00" if code == 12 else "00B0F0" if code == 13 else bg
        put(ws, f"B{r}", lab, size=9, bold=True, bg=bg, color="FFFFFF" if lab == "Mecka" else "000000", border=BOX)
        put(ws, f"C{r}", "Eingabe" if code is None else "automatisch", size=7, italic=True, color="595959", border=BOX)
        if lab == "Arbeiten bei Tag":
            ws[f"A{r}"].fill = fill("C4BD97")
        if lab == "Mecka":
            ws[f"A{r}"].fill = fill("000000")
    # Hilfsspalten verstecken
    for ci in range(M.DAYN + 1, M.K_CODE0 + M.NDAYS):
        ws.column_dimensions[gl(ci)].hidden = True
    ws.column_dimensions.group(gl(M.DAYN + 1), gl(M.K_CODE0 + M.NDAYS - 1), hidden=True)

    grid = f"{C0}{R0}:{CN}{RN}"
    allrows = f"{C0}2:{CN}{RN + 7}"
    TYP, GRP = gl(M.K_TYP), gl(M.K_GRP)
    # Monatsanfang: duenne schwarze Linie links (ohne Stopp, hoechste Prioritaet)
    left = Border(left=Side(style="thin", color="000000"))
    cf(ws, allrows, f'AND(ISNUMBER({C0}$3),DAY({C0}$3)=1)', dxf(border=left), stop=False)
    # Teamkopfzeilen
    for g in M.GROUPS:
        cf(ws, f"A{R0}:{CN}{RN}", f'AND(${TYP}{R0}="K",${GRP}{R0}="{g}")',
           dxf(bg=M.GROUP_COLOR[g], fg=M.GROUP_FONT[g], bold=True))
        cf(ws, f"A{R0}:A{RN}", f'AND(${TYP}{R0}="P",${GRP}{R0}="{g}")', dxf(bg=M.GROUP_COLOR[g]))
    # Statuscodes
    code0 = kal_code_ref(M.DAY0, R0)
    for code in (8, 4, 3, 2, 7, 5, 6, 10, 1, 11, 9, 12, 13):
        s = M.STATUS_BY_CODE[code]
        cf(ws, grid, f"{code0}={code}",
           dxf(bg=s[5], fg=s[6], bold=s[7], underline="single" if s[8] else None, italic=s[9] or None))
    cf(ws, grid, f'AND(${TYP}{R0}="P",{C0}$6<>"")', dxf(bg=M.HOLIDAY_FILL))
    cf(ws, grid, f'AND(${TYP}{R0}="P",OR({C0}$4="Sa",{C0}$4="So"))', dxf(bg=M.WEEKEND_FILL))
    # Kopfzeilen
    hdr = f"{C0}2:{CN}5"
    cf(ws, hdr, f'{C0}$6<>""', dxf(bg=M.HOLIDAY_FILL, fg="CC0000", bold=True))
    cf(ws, hdr, f'OR({C0}$4="Sa",{C0}$4="So")', dxf(bg=M.WEEKEND_FILL, bold=True))
    cf(ws, f"{C0}6:{CN}6", f'{C0}6<>""', dxf(bg=M.HOLIDAY_FILL))
    for t in M.TEAMS:
        cf(ws, f"{C0}7:{CN}7", f'{C0}7="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
        cf(ws, f"{C0}8:{CN}8", f'ISNUMBER(SEARCH("{t}",{C0}8))', dxf(bg=M.TEAM_LIGHT[t], bold=True))
    cf(ws, f"{C0}11:{CN}11", f'{C0}11="✅"', dxf(bg="C6EFCE"))
    cf(ws, f"{C0}11:{CN}11", f'{C0}11="❌"', dxf(bg="FFC7CE"))
    cf(ws, f"{C0}12:{CN}12", f'AND(ISNUMBER({C0}12),{C0}12<SollTag)', dxf(bg="FFC7CE", fg="9C0006"))
    cf(ws, f"{C0}12:{CN}12", f'AND(ISNUMBER({C0}12),{C0}12<SollTag+Reserve)', dxf(bg="FFEB9C", fg="9C5700"))
    cf(ws, f"{C0}12:{CN}12", f'ISNUMBER({C0}12)', dxf(bg="C6EFCE", fg="006100"))
    pr = RN + 2
    cf(ws, f"{C0}{pr}:{CN}{pr}", f'AND(ISNUMBER({C0}{pr}),{C0}{pr}>0)', dxf(bg="FFFF00"))
    cf(ws, f"{C0}{pr + 1}:{CN}{pr + 1}", f'AND(ISNUMBER({C0}{pr + 1}),{C0}{pr + 1}>0)', dxf(bg="00B0F0"))
    cf(ws, f"{C0}{pr + 2}:{CN}{pr + 7}", f'AND(ISNUMBER({C0}$3),OR({C0}$4="Sa",{C0}$4="So"))', dxf(bg=M.WEEKEND_FILL))

    # Legende unten
    l0 = RN + 10
    put(ws, f"B{l0}", "Legende", bold=True, size=11)
    put(ws, f"C{l0}", "Kalender", bold=True, size=9)
    legend = [(3, "Urlaub"), (4, "Urlaub genehmigt eTime"), (2, "JAZ + 99"), (8, "Krank"), (7, "Unterrichte"),
              (6, "FDL-Einweisungen"), (5, "Kabeleinweisungen"), (1, "Besonderheiten (Text)"),
              (12, "Zusätzliche MA"), (13, "AZUBIS"), (9, "Bereitschaft / Nachtdienst"), (10, "Freistellung")]
    for i, (code, lab) in enumerate(legend):
        r = l0 + 1 + i
        s = M.STATUS_BY_CODE[code]
        put(ws, f"B{r}", lab, size=9, border=BOX)
        txt = s[2] if code not in (1,) else "Text"
        put(ws, f"C{r}", txt or None, size=9, bg=s[5], color=s[6], bold=s[7], italic=s[9],
            underline="single" if s[8] else None, h="center", border=BOX)
    r = l0 + 1 + len(legend)
    put(ws, f"B{r}", "Regulärer Tagdienst", size=9, border=BOX)
    put(ws, f"C{r}", None, border=BOX)
    put(ws, f"B{r + 1}", "Wochenende", size=9, border=BOX)
    put(ws, f"C{r + 1}", None, bg=M.WEEKEND_FILL, border=BOX)
    put(ws, f"B{r + 2}", "Feiertag (Hessen)", size=9, border=BOX)
    put(ws, f"C{r + 2}", None, bg=M.HOLIDAY_FILL, border=BOX)
    note(ws, f"B{r + 4}", "Eingabe im Kalender (mit Makros): Doppelklick auf einen Tag oder Rechtsklick auf markierte "
         "Tage. Kürzel: U = Urlaub, UE = Urlaub eTime, K = Krank, JAZ, S = Unterricht, FDL, Kabel, FS = Freistellung, "
         "- = regulärer Tagdienst, V = Vertretung, Z = Zusätzliche MA, A = AZUBI, leer = Eintrag löschen, sonstiger "
         "Text = Besonderheit. Die Eingabe wird in die Tabellen geschrieben; das Raster selbst bleibt geschützt.")
    ws.freeze_panes = f"{C0}{R0}"
    ws.sheet_view.zoomScale = 90
    return ws
