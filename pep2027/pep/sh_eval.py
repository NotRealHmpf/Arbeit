"""Blatt 6 Einsatzfaehigkeit (mit Ersatzlogik) und Personalbedarfsplanung."""
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as gl

from . import model as M
from .ctx import S_BED, S_EIN
from .style import BOX, add_table, cf, dxf, font, note, put, title, widths

ROW0 = 8                    # erste Datenzeile (Kopf in Zeile 7)
LAST = ROW0 + M.NDAYS - 1
RK = {"Meister": "M", "AV": "AV", "TVT-Fahrer": "TVT", "Handwerker": "HW"}
QN = {"Meister": "Ber_QMeister", "AV": "Ber_QAV", "TVT-Fahrer": "Ber_QTVT", "Handwerker": "Ber_QHW"}
SOLLN = {"Meister": "SollM", "AV": "SollAV", "TVT-Fahrer": "SollTVT", "Handwerker": "SollHW"}

EIN_HDR = ["Nr", "Datum", "Wochentag", "Feiertag", "Bereitschaftsteam", "Bereitschaftsart", "Soll", "Verfügbar",
           "Meister", "AV", "TVT", "Handwerker", "Vertretung", "Status", "Ampel", "Fehlermeldung",
           "Ersatz Meister", "Ersatz AV", "Ersatz TVT", "Ersatz Handwerker", "Kurz",
           "Fehlt M", "Fehlt AV", "Fehlt TVT", "Fehlt HW", "Res M", "Res AV", "Res TVT", "Res HW",
           "Intern M", "Intern AV", "Intern TVT", "Intern HW", "OK", "Stufe",
           "Best M", "Best AV", "Best TVT", "Best HW1", "Best HW2"]


def ec(h):
    return gl(2 + EIN_HDR.index(h))


def build_einsatz(ctx):
    ws = ctx.wb[S_EIN]
    ws.sheet_properties.tabColor = "C00000"
    title(ws, "B1", "Blatt 6 – Einsatzfähigkeitsprüfung Bereitschaft")
    note(ws, "B2", "Täglich automatisch für das Bereitschaftsteam. Erforderlich: 1 Meister, 1 AV, 1 TVT-Fahrer, "
         "2 Handwerker (Blatt Einstellungen). Berücksichtigt Urlaub, Krankheit, JAZ, Unterrichte, FDL- und "
         "Kabeleinweisungen, Freistellungen und eingetragene Vertretungen.")
    note(ws, "B3", "Ersatzvorschlag je fehlender Funktion – Priorität: 1. gleiches Team (Mehrfachqualifikation, ohne "
         "neue Lücke), 2. anderes Team (Stammfunktion), 3. Mitarbeiter mit Mehrfachfunktion. Bei Gleichstand: "
         "weniger Bereitschaftstage zuerst. Übernommene Vorschläge im Blatt Vertretungen eintragen.")
    note(ws, "B4", "Ampel: 🟢 Sollbesetzung erreicht · 🟡 Kritische Besetzung (nur durch interne Vertretung) · "
         "🔴 Unterbesetzung (nicht einsatzfähig).")
    c = {h: ec(h) for h in EIN_HDR}
    hr = ROW0 - 1
    for h in EIN_HDR:
        put(ws, f"{c[h]}{hr}", h, bold=True, wrap=True, h="center")
    for k in range(1, M.NDAYS + 1):
        r = ROW0 + k - 1
        d = f"$C{r}"
        G = f"INDEX(Ber_Grid,0,$B{r})"
        T = f"$F{r}"
        on = f'{G},">=9",{G},"<=11",{G},"<>10"'
        F = {"Nr": k}
        F["Datum"] = f'=IF($B{r}>DATE(Jahr,12,31)-DATE(Jahr,1,1)+1,"",DATE(Jahr,1,1)+$B{r}-1)'
        F["Wochentag"] = f'=IF({d}="","",INDEX(lstWT,WEEKDAY({d},2)))'
        F["Feiertag"] = f'=IF({d}="","",INDEX(Schicht_Feiertag,$B{r}))'
        F["Bereitschaftsteam"] = f'=IF({d}="","",INDEX(Schicht_Nacht,$B{r}))'
        F["Bereitschaftsart"] = f'=IF({d}="","",INDEX(tblSchicht[Bereitschaftsart],$B{r}))'
        F["Soll"] = f'=IF({d}="","",SollBereitschaft)'
        F["Verfügbar"] = f'=IF({d}="","",COUNTIFS({on}))'
        for rolle in M.ROLLEN:
            h = M.ROLLE_KURZ[rolle]
            F[h] = (f'=IF({d}="","",COUNTIFS({on},Ber_Gruppe,{T},Ber_Funktion,"{rolle}")'
                    f'+COUNTIFS(V_Funktion,"{rolle}",V_Von,"<="&{d},V_Bis,">="&{d}))')
        nv = f'COUNTIFS(V_Von,"<="&{d},V_Bis,">="&{d})'
        first = (f'INDEX(V_MA,SUMPRODUCT(SMALL(ROW(V_MA)-ROW(tblVertretungen[#Headers])'
                 f'+((V_Von<={d})*(V_Bis>={d})=0)*1000000,1)))')
        F["Vertretung"] = f'=IF({d}="","",IF({nv}=0,"",{first}&IF({nv}>1," +"&({nv}-1),"")))'
        for rolle in M.ROLLEN:
            k_ = RK[rolle]
            have = f"{c[M.ROLLE_KURZ[rolle]]}{r}"
            F[f"Fehlt {k_}"] = f'=IF({d}="",0,MAX(0,{SOLLN[rolle]}-{have}))'
            F[f"Res {k_}"] = f'=IF({d}="",0,MAX(0,{have}-{SOLLN[rolle]}))'
        for rolle in M.ROLLEN:
            k_ = RK[rolle]
            terms = []
            for other in M.ROLLEN:
                if other == rolle:
                    continue
                cnt = f'COUNTIFS({on},Ber_Gruppe,{T},Ber_Funktion,"{other}",{QN[rolle]},1)'
                terms.append(f'MIN({c["Res " + RK[other]]}{r},{cnt})')
            F[f"Intern {k_}"] = f'=IF({d}="",0,{"+".join(terms)})'
        fehlt = [f'{c["Fehlt " + RK[x]]}{r}' for x in M.ROLLEN]
        res = [f'{c["Res " + RK[x]]}{r}' for x in M.ROLLEN]
        intern = [f'{c["Intern " + RK[x]]}{r}' for x in M.ROLLEN]
        conds = ",".join(f"{a}<={b}" for a, b in zip(fehlt, intern))
        F["OK"] = f'=IF({d}="",0,IF(AND({conds},{"+".join(fehlt)}<={"+".join(res)}),1,0))'
        F["Stufe"] = f'=IF({d}="","",IF({"+".join(fehlt)}=0,0,IF({c["OK"]}{r}=1,1,2)))'
        st = f'{c["Stufe"]}{r}'
        F["Status"] = f'=IF({d}="","",IF({c["OK"]}{r}=1,"✅ Einsatzfähig","❌ Nicht einsatzfähig"))'
        F["Ampel"] = (f'=IF({d}="","",CHOOSE({st}+1,"🟢 Sollbesetzung erreicht","🟡 Kritische Besetzung",'
                      f'"🔴 Unterbesetzung"))')
        F["Kurz"] = f'=IF({d}="","",IF({c["OK"]}{r}=1,"✅","❌"))'
        msgs = {"Meister": "Meister fehlt", "AV": "AV fehlt", "TVT-Fahrer": "TVT fehlt",
                "Handwerker": "Zu wenig Handwerker"}
        parts = []
        for x in M.ROLLEN:
            fz, iz = f'{c["Fehlt " + RK[x]]}{r}', f'{c["Intern " + RK[x]]}{r}'
            parts.append(f'IF({fz}>0,"{msgs[x]}"&IF({fz}<={iz}," (intern vertretbar)",""),"")')
        F["Fehlermeldung"] = f'=IF({d}="","",_xlfn.TEXTJOIN("; ",TRUE,{",".join(parts)}))'
        # Ersatzvorschlag
        resown = "+".join(f'(Ber_Funktion="{x}")*{c["Res " + RK[x]]}{r}' for x in M.ROLLEN)
        for rolle in M.ROLLEN:
            k_ = RK[rolle]
            fz = f'{c["Fehlt " + k_]}{r}'
            q_ = QN[rolle]
            onarr = f"({G}>=9)*({G}<=11)*({G}<>10)"
            free = f"(({G}=0)+({G}=1))"
            same = f"(Ber_Gruppe={T})"
            stamm = f'(Ber_Funktion="{rolle}")'
            p1 = f"{same}*{onarr}*{q_}*(1-{stamm})*(({resown})>0)"
            p2 = f'(1-{same})*(Ber_Kat="Team")*{free}*{stamm}'
            p3 = f'(1-{same})*(Ber_Kat<>"Ressource")*{free}*{q_}*(1-{stamm})'
            base = "(Ber_Bereitschaftstage*1000+Ber_Rang)"
            v = (f"999999-{p1}*(999999-(100000+{base}))-{p2}*(999999-(200000+{base}))"
                 f"-{p3}*(999999-(300000+{base}))")

            def best(n):
                return f'=IF(OR({d}="",{fz}<{n}),"",SUMPRODUCT(SMALL({v},{n})))'

            def show(ref):
                return (f'IF({ref}="","",IF({ref}>=999999,"⚠ kein Ersatz verfügbar",INDEX(Ber_Name,MOD({ref},1000))'
                        f'&" ("&INDEX(Ber_Gruppe,MOD({ref},1000))&" – "&CHOOSE(INT({ref}/100000),"gleiches Team",'
                        f'"anderes Team","Mehrfachfunktion")&")"))')
            h = "Ersatz " + M.ROLLE_KURZ[rolle]
            if rolle == "Handwerker":
                F["Best HW1"], F["Best HW2"] = best(1), best(2)
                b1, b2 = f'{c["Best HW1"]}{r}', f'{c["Best HW2"]}{r}'
                F[h] = f'=IF({b1}="","",{show(b1)}&IF({b2}="","","; "&{show(b2)}))'
            else:
                F["Best " + k_] = best(1)
                F[h] = "=" + show(f'{c["Best " + k_]}{r}')
        for h in EIN_HDR:
            cell = ws[f"{c[h]}{r}"]
            cell.value = F[h]
            cell.font = font(9)
            cell.alignment = Alignment(horizontal="left" if h in ("Fehlermeldung", "Vertretung") or
                                       h.startswith("Ersatz") or h in ("Status", "Ampel", "Bereitschaftsart") else "center",
                                       vertical="center")
        ws[f"{c['Datum']}{r}"].number_format = "DDD DD.MM.YYYY"
    last_col = c[EIN_HDR[-1]]
    add_table(ws, "tblEinsatz", f"B{hr}:{last_col}{LAST}", style="TableStyleLight9")
    ws.row_dimensions[hr].height = 30
    wd = {"Nr": 5, "Datum": 15, "Wochentag": 6, "Feiertag": 14, "Bereitschaftsteam": 11, "Bereitschaftsart": 27,
          "Soll": 6, "Verfügbar": 8, "Meister": 8, "AV": 6, "TVT": 6, "Handwerker": 10, "Vertretung": 14,
          "Status": 19, "Ampel": 24, "Fehlermeldung": 40, "Ersatz Meister": 30, "Ersatz AV": 30, "Ersatz TVT": 30,
          "Ersatz Handwerker": 46, "Kurz": 6}
    for h in EIN_HDR:
        ws.column_dimensions[c[h]].width = wd.get(h, 7)
        if h not in wd:
            ws.column_dimensions[c[h]].outlineLevel = 1
    ws.column_dimensions["A"].width = 2
    rng = lambda h: f"{c[h]}{ROW0}:{c[h]}{LAST}"
    for t in M.TEAMS:
        cf(ws, rng("Bereitschaftsteam"), f'{c["Bereitschaftsteam"]}{ROW0}="{t}"',
           dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
    cf(ws, rng("Status"), f'{c["OK"]}{ROW0}=1', dxf(bg="C6EFCE", fg="006100", bold=True))
    cf(ws, rng("Status"), f'AND({c["Datum"]}{ROW0}<>"",{c["OK"]}{ROW0}=0)', dxf(bg="FFC7CE", fg="9C0006", bold=True))
    cf(ws, rng("Ampel"), f'{c["Stufe"]}{ROW0}=0', dxf(bg="C6EFCE", fg="006100"))
    cf(ws, rng("Ampel"), f'{c["Stufe"]}{ROW0}=1', dxf(bg="FFEB9C", fg="9C5700"))
    cf(ws, rng("Ampel"), f'{c["Stufe"]}{ROW0}=2', dxf(bg="FFC7CE", fg="9C0006"))
    for h in ("Meister", "AV", "TVT", "Handwerker"):
        rolle = {"TVT": "TVT-Fahrer"}.get(h, h)
        cf(ws, rng(h), f'AND({c["Datum"]}{ROW0}<>"",{c[h]}{ROW0}<{SOLLN[rolle]})', dxf(bg="FFC7CE", fg="9C0006", bold=True))
    cf(ws, f"{c['Datum']}{ROW0}:{c['Wochentag']}{LAST}", f'${c["Feiertag"]}{ROW0}<>""',
       dxf(bg=M.HOLIDAY_FILL, fg="CC0000", bold=True))
    cf(ws, f"{c['Datum']}{ROW0}:{c['Wochentag']}{LAST}", f'OR(${c["Wochentag"]}{ROW0}="Sa",${c["Wochentag"]}{ROW0}="So")',
       dxf(bg=M.WEEKEND_FILL))
    for h in ("Ersatz Meister", "Ersatz AV", "Ersatz TVT", "Ersatz Handwerker"):
        cf(ws, rng(h), f'LEFT({c[h]}{ROW0},1)="⚠"', dxf(fg="9C0006", bold=True))
    # Zusammenfassung
    put(ws, f"{c['Status']}5", "Tage einsatzfähig", bold=True, size=9)
    put(ws, f"{c['Ampel']}5", f'=COUNTIFS(tblEinsatz[OK],1,tblEinsatz[Datum],"<>")', bold=True, h="left")
    put(ws, f"{c['Status']}6", "Tage nicht einsatzfähig", bold=True, size=9)
    put(ws, f"{c['Ampel']}6", f'=SUMPRODUCT((tblEinsatz[Stufe]=2)*1)', bold=True, h="left", color="9C0006")
    put(ws, f"{c['Fehlermeldung']}5", '="🟢 "&COUNTIF(tblEinsatz[Stufe],0)&"   🟡 "&COUNTIF(tblEinsatz[Stufe],1)&"   🔴 "'
        '&COUNTIF(tblEinsatz[Stufe],2)', bold=True)
    for nm, h in [("Einsatz_Datum", "Datum"), ("Einsatz_Team", "Bereitschaftsteam"), ("Einsatz_Verfuegbar", "Verfügbar"),
                  ("Einsatz_Status", "Status"), ("Einsatz_Ampel", "Ampel"), ("Einsatz_Kurz", "Kurz"),
                  ("Einsatz_Stufe", "Stufe"), ("Einsatz_Fehler", "Fehlermeldung"), ("Einsatz_OK", "OK")]:
        ctx.name(nm, f"tblEinsatz[{h}]")
    ws.freeze_panes = f"D{ROW0}"
    return ws


# ======================================================================== Personalbedarf
BED_HDR = ["Nr", "Datum", "Wochentag", "KW", "Monat", "Arbeitstag", "Bestand", "Abwesend", "Freigestellt",
           "Ist Tagdienst", "Soll Tagdienst", "Verfügbarkeit", "Ausfallreserve", "Unterbesetzung", "Überbesetzung",
           "Ampel Tagdienst", "Bereitschaft Soll", "Bereitschaft Ist", "Bereitschaft Ampel", "Bereitschaft Stufe",
           "Zusätzliche MA", "AZUBIS"]


def bc(h):
    return gl(2 + BED_HDR.index(h))


def build_bedarf(ctx):
    ws = ctx.wb[S_BED]
    ws.sheet_properties.tabColor = "7030A0"
    title(ws, "B1", "Personalbedarfsplanung")
    note(ws, "B2", "Tagdienst an Arbeitstagen (Mo–Fr ohne Feiertage): Ist = Bestand der Teams − Abwesende − "
         "Freigestellte. Soll = Mindestbesetzung Tagdienst, Ausfallreserve = Ist − Soll, Unterbesetzung = Soll − Ist, "
         "Überbesetzung = Ist − Soll − Mindestreserve (Blatt Einstellungen).")
    note(ws, "B3", "Zusätzliche MA und AZUBIS werden getrennt ausgewiesen und nicht in Ist/Soll eingerechnet. "
         "Bereitschaft: Soll und Ist aus der Einsatzfähigkeitsprüfung.")
    put(ws, "B4", "Personalbestand (Teams)", bold=True)
    put(ws, "F4", '=COUNTIF(MA_Kat,"Team")', bold=True, h="center", border=BOX)
    ctx.name("Bestand", f"{S_BED}!$F$4")
    c = {h: bc(h) for h in BED_HDR}
    hr = ROW0 - 1
    for h in BED_HDR:
        put(ws, f"{c[h]}{hr}", h, bold=True, wrap=True, h="center")
    for k in range(1, M.NDAYS + 1):
        r = ROW0 + k - 1
        d = f"$C{r}"
        G = f"INDEX(Ber_Grid,0,$B{r})"
        at = f"${c['Arbeitstag']}{r}"
        ist, soll = f"${c['Ist Tagdienst']}{r}", f"${c['Soll Tagdienst']}{r}"
        F = {"Nr": k,
             "Datum": f'=IF($B{r}>DATE(Jahr,12,31)-DATE(Jahr,1,1)+1,"",DATE(Jahr,1,1)+$B{r}-1)',
             "Wochentag": f'=IF({d}="","",INDEX(lstWT,WEEKDAY({d},2)))',
             "KW": f'=IF({d}="","",_xlfn.ISOWEEKNUM({d}))',
             "Monat": f'=IF({d}="","",MONTH({d}))',
             "Arbeitstag": f'=IF({d}="",0,INDEX(Ber_AT,1,$B{r}))',
             "Bestand": f'=IF({d}="","",Bestand)',
             "Abwesend": f'=IF({d}="","",COUNTIFS({G},">=2",{G},"<=8",Ber_Kat,"Team"))',
             "Freigestellt": f'=IF({d}="","",COUNTIFS({G},10,Ber_Kat,"Team"))',
             "Ist Tagdienst": f'=IF({d}="","",IF({at}=1,${c["Bestand"]}{r}-${c["Abwesend"]}{r}-${c["Freigestellt"]}{r},0))',
             "Soll Tagdienst": f'=IF({d}="","",IF({at}=1,SollTag,0))',
             "Verfügbarkeit": f'=IF(OR({d}="",{at}=0),"",IF(Bestand=0,0,{ist}/Bestand))',
             "Ausfallreserve": f'=IF(OR({d}="",{at}=0),"",MAX(0,{ist}-{soll}))',
             "Unterbesetzung": f'=IF(OR({d}="",{at}=0),"",MAX(0,{soll}-{ist}))',
             "Überbesetzung": f'=IF(OR({d}="",{at}=0),"",MAX(0,{ist}-{soll}-Reserve))',
             "Ampel Tagdienst": f'=IF({d}="","",IF({at}=0,"–",IF({ist}<{soll},"🔴",IF({ist}<{soll}+Reserve,"🟡","🟢"))))',
             "Bereitschaft Soll": f'=IF({d}="","",SollBereitschaft)',
             "Bereitschaft Ist": f'=IF({d}="","",INDEX(Einsatz_Verfuegbar,$B{r}))',
             "Bereitschaft Stufe": f'=IF({d}="","",INDEX(Einsatz_Stufe,$B{r}))',
             "Bereitschaft Ampel": f'=IF({d}="","",CHOOSE(${c["Bereitschaft Stufe"]}{r}+1,"🟢","🟡","🔴"))',
             "Zusätzliche MA": (f'=IF({d}="","",COUNTIF({G},12)+COUNTIFS(R_Art,"Zusätzliche MA",R_Stamm,0,'
                                f'R_Von,"<="&{d},R_Bis,">="&{d}))'),
             "AZUBIS": (f'=IF({d}="","",COUNTIF({G},13)+COUNTIFS(R_Art,"AZUBIS",R_Stamm,0,'
                        f'R_Von,"<="&{d},R_Bis,">="&{d}))')}
        for h in BED_HDR:
            cell = ws[f"{c[h]}{r}"]
            cell.value = F[h]
            cell.font = font(9)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{c['Datum']}{r}"].number_format = "DDD DD.MM.YYYY"
        ws[f"{c['Verfügbarkeit']}{r}"].number_format = "0%"
    last_col = c[BED_HDR[-1]]
    add_table(ws, "tblBedarf", f"B{hr}:{last_col}{LAST}", style="TableStyleLight11")
    ws.row_dimensions[hr].height = 30
    for h in BED_HDR:
        ws.column_dimensions[c[h]].width = {"Datum": 15, "Nr": 5, "Wochentag": 6, "KW": 5, "Monat": 6}.get(h, 10)
    ws.column_dimensions["A"].width = 2
    rng = lambda h: f"{c[h]}{ROW0}:{c[h]}{LAST}"
    cf(ws, rng("Unterbesetzung"), f'AND(ISNUMBER({c["Unterbesetzung"]}{ROW0}),{c["Unterbesetzung"]}{ROW0}>0)',
       dxf(bg="FFC7CE", fg="9C0006", bold=True))
    cf(ws, rng("Überbesetzung"), f'AND(ISNUMBER({c["Überbesetzung"]}{ROW0}),{c["Überbesetzung"]}{ROW0}>0)',
       dxf(bg="DDEBF7", fg="1F4E78"))
    cf(ws, rng("Zusätzliche MA"), f'AND(ISNUMBER({c["Zusätzliche MA"]}{ROW0}),{c["Zusätzliche MA"]}{ROW0}>0)', dxf(bg="FFFF00"))
    cf(ws, rng("AZUBIS"), f'AND(ISNUMBER({c["AZUBIS"]}{ROW0}),{c["AZUBIS"]}{ROW0}>0)', dxf(bg="00B0F0"))
    cf(ws, f"{c['Datum']}{ROW0}:{c['Arbeitstag']}{LAST}", f'AND(${c["Datum"]}{ROW0}<>"",${c["Arbeitstag"]}{ROW0}=0)',
       dxf(bg=M.WEEKEND_FILL))
    for nm, h in [("Bedarf_Datum", "Datum"), ("Bedarf_KW", "KW"), ("Bedarf_Monat", "Monat"),
                  ("Bedarf_AT", "Arbeitstag"), ("Bedarf_Abw", "Abwesend"), ("Bedarf_FS", "Freigestellt"),
                  ("Bedarf_Ist", "Ist Tagdienst"), ("Bedarf_Soll", "Soll Tagdienst"), ("Bedarf_Verf", "Verfügbarkeit"),
                  ("Bedarf_Reserve", "Ausfallreserve"), ("Bedarf_Unter", "Unterbesetzung"),
                  ("Bedarf_Ueber", "Überbesetzung"), ("Bedarf_BerIst", "Bereitschaft Ist"),
                  ("Bedarf_BerStufe", "Bereitschaft Stufe"), ("Bedarf_Zus", "Zusätzliche MA"), ("Bedarf_Azubi", "AZUBIS")]:
        ctx.name(nm, f"tblBedarf[{h}]")

    # ------------------------------------------------ Wochen / Monate / Jahr (rechts)
    x0 = 2 + len(BED_HDR) + 1
    agg_hdr = ["Zeitraum", "Von", "Bis", "Arbeitstage", "Ø Ist Tagdienst", "Min. Ist", "Ø Verfügbarkeit",
               "Ø Ausfallreserve", "Tage Unterbesetzung", "Fehlende Personentage", "Ø fehlende MA",
               "Tage Überbesetzung", "Bereitschaft 🟡 (Tage)", "Bereitschaft 🔴 (Tage)", "Zusätzliche MA (PT)",
               "AZUBIS (PT)"]
    A = [gl(x0 + i) for i in range(len(agg_hdr))]
    ws.column_dimensions[gl(x0 - 1)].width = 3

    def agg_block(top, label, crit_col, keys, key_label):
        put(ws, f"{A[0]}{top}", label, bold=True, size=11)
        for cc, h in zip(A, agg_hdr):
            put(ws, f"{cc}{top + 1}", h, bold=True, wrap=True, bg=M.HEADER_FILL, border=BOX, h="center", size=9)
        ws.row_dimensions[top + 1].height = 30
        for i, key in enumerate(keys):
            r = top + 2 + i
            crit = f"{crit_col},{A[0]}{r}" if crit_col else None
            put(ws, f"{A[0]}{r}", key if not key_label else key_label(key), border=BOX, h="center", bold=True)
            if crit_col is None:
                cr = 'Bedarf_Datum,">0"'
                at = 'Bedarf_AT,1'
            else:
                cr = f"{crit_col},{key}"
                at = f"{crit_col},{key},Bedarf_AT,1"
            f_ = {
                "Von": f'=IFERROR(1/(1/_xlfn.MINIFS(Bedarf_Datum,{cr})),"")',
                "Bis": f'=IFERROR(1/(1/_xlfn.MAXIFS(Bedarf_Datum,{cr})),"")',
                "Arbeitstage": f'=COUNTIFS({at})',
                "Ø Ist Tagdienst": f'=IFERROR(AVERAGEIFS(Bedarf_Ist,{at}),"")',
                "Min. Ist": f'=IF(COUNTIFS({at})=0,"",_xlfn.MINIFS(Bedarf_Ist,{at}))',
                "Ø Verfügbarkeit": f'=IFERROR(AVERAGEIFS(Bedarf_Verf,{at}),"")',
                "Ø Ausfallreserve": f'=IFERROR(AVERAGEIFS(Bedarf_Reserve,{at}),"")',
                "Tage Unterbesetzung": f'=COUNTIFS({at},Bedarf_Unter,">0")',
                "Fehlende Personentage": f'=SUMIFS(Bedarf_Unter,{at})',
                "Ø fehlende MA": f'=IFERROR(SUMIFS(Bedarf_Unter,{at})/COUNTIFS({at}),"")',
                "Tage Überbesetzung": f'=COUNTIFS({at},Bedarf_Ueber,">0")',
                "Bereitschaft 🟡 (Tage)": f'=COUNTIFS({cr},Bedarf_BerStufe,1)',
                "Bereitschaft 🔴 (Tage)": f'=COUNTIFS({cr},Bedarf_BerStufe,2)',
                "Zusätzliche MA (PT)": f'=SUMIFS(Bedarf_Zus,{cr})',
                "AZUBIS (PT)": f'=SUMIFS(Bedarf_Azubi,{cr})',
            }
            for cc, h in zip(A[1:], agg_hdr[1:]):
                nf = {"Von": "DD.MM.YY", "Bis": "DD.MM.YY", "Ø Verfügbarkeit": "0%", "Ø Ist Tagdienst": "0.0",
                      "Ø Ausfallreserve": "0.0", "Ø fehlende MA": "0.0"}.get(h)
                put(ws, f"{cc}{r}", f_[h], border=BOX, h="center", nf=nf, size=9)
        return top + 2 + len(keys)

    put(ws, f"{A[0]}4", "Auswertung", bold=True, size=12)
    end = agg_block(5, "Jahresbasis", None, ['=Jahr'], None)
    ctx.name("BedarfJahr_Unter", f"{S_BED}!${A[9]}$7")
    mstart = end + 1
    end = agg_block(mstart, "Monatsbasis", "Bedarf_Monat", list(range(1, 13)), None)
    for i in range(12):
        ws[f"{A[0]}{mstart + 2 + i}"] = f"=INDEX(lstMonKurz,{i + 1})"
        # Kriterium bleibt die Monatszahl -> Formeln oben nutzen den Schluessel direkt
    wstart = end + 1
    agg_block(wstart, "Wochenbasis (Kalenderwoche ISO)", "Bedarf_KW", list(range(1, 54)), None)
    for cc, w in zip(A, [11, 9, 9, 10, 10, 8, 11, 11, 12, 12, 10, 12, 12, 12, 11, 10]):
        ws.column_dimensions[cc].width = w
    ctx.bedarf_agg = {"cols": A, "hdr": agg_hdr, "year_row": 7, "month_row0": mstart + 2, "week_row0": wstart + 2}
    ws.freeze_panes = f"D{ROW0}"
    return ws
