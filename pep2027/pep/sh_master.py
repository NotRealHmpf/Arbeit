"""Blatt 1 Mitarbeiter (Stammdaten) und Blatt 2 Funktionsmatrix."""
from . import model as M
from .ctx import S_FM, S_MA
from .style import (BOX, INPUT_FILL, add_table, cf, dv_list, dxf, fill, note, put, title, widths)


def tr(tbl, col):
    """Strukturierter Verweis auf die aktuelle Zeile."""
    return f"{tbl}[[#This Row],[{col}]]"


def ma_formulas():
    T = "tblMitarbeiter"
    name, team, funk = tr(T, "Name"), tr(T, "Team"), tr(T, "Funktion")
    kat = f'IFERROR(INDEX(FunktionKat,MATCH({funk},lstFunktionen,0)),"")'
    return {
        "Gruppe": (f'IF({name}="","",IF({kat}="Ressource","Ressourcen",IF({kat}="Ohne Team","Ohne Team",'
                   f'IF(ISNUMBER(MATCH({team},lstTeams,0)),{team},"Ohne Team"))))'),
        "Kategorie": (f'IF({name}="","",IF({tr(T, "Gruppe")}="Ressourcen","Ressource",'
                      f'IF({tr(T, "Gruppe")}="Ohne Team","Ohne Team","Team")))'),
        "Sortierschlüssel": (f'IF({name}="",999999999,MATCH({tr(T, "Gruppe")},lstGruppen,0)*1000000'
                             f'+IFERROR(MATCH({funk},lstFunktionen,0),9)*10000+ROW()-ROW({T}[#Headers]))'),
        "Prüfung": (f'IF({name}="","",IF(COUNTIF({T}[Name],{name})>1,"⚠ Name doppelt",'
                    f'IF(ISNA(MATCH({funk},lstFunktionen,0)),"⚠ Funktion fehlt/ungültig",'
                    f'IF(AND({team}<>"",ISNA(MATCH({team},lstTeams,0))),"⚠ Team ungültig",'
                    f'IF(AND({tr(T, "Kategorie")}="Team",{team}=""),"⚠ Team fehlt",'
                    f'IF(COUNTIF(FM_Name,{name})=0,"ℹ fehlt in Funktionsmatrix","✓"))))))'),
    }


def fm_formulas():
    T = "tblFunktionsmatrix"
    name = tr(T, "Name")
    return {
        "Team": f'IFERROR(INDEX(tblMitarbeiter[Team],MATCH({name},tblMitarbeiter[Name],0))&"","")',
        "Funktion": f'IFERROR(INDEX(tblMitarbeiter[Funktion],MATCH({name},tblMitarbeiter[Name],0))&"","")',
        "Anzahl Qualifikationen": f'COUNTIF({T}[[#This Row],[Meister]:[Kettensäge]],"JA")',
        "Mehrfachfunktion": f'IF(COUNTIF({T}[[#This Row],[Meister]:[Handwerker]],"JA")>=2,"JA","NEIN")',
        "Prüfung": (f'IF({name}="","",IF(COUNTIF(lstMA,{name})=0,"⚠ nicht im Mitarbeiterstamm",'
                    f'IF(COUNTIF({T}[Name],{name})>1,"⚠ doppelt",'
                    f'IF(ISNUMBER(MATCH({tr(T, "Funktion")},lstRollen,0)),'
                    f'IF(INDEX({T}[[#This Row],[Meister]:[Handwerker]],1,MATCH({tr(T, "Funktion")},lstRollen,0))="JA",'
                    f'"✓","ℹ Stammfunktion nicht als JA markiert"),"✓"))))'),
    }


def build_mitarbeiter(ctx, people):
    ws = ctx.wb[S_MA]
    ws.sheet_properties.tabColor = "4F81BD"
    widths(ws, {"A": 2, "B": 16, "C": 24, "D": 12, "E": 16, "F": 36, "G": 13, "H": 12, "I": 13, "J": 26})
    title(ws, "B1", "Blatt 1 – Mitarbeiterstammdaten")
    note(ws, "B2", "Eingabe in den Spalten Personalnummer bis Bemerkung (gelb hinterlegt). Neue Mitarbeiter einfach "
         "unten an die Tabelle anfügen – Kalender, Matrix und Auswertungen übernehmen sie automatisch.")
    note(ws, "B3", "Der Name ist der Schlüssel für alle anderen Blätter und muss eindeutig sein. "
         "Platzhalter Ma6 … Ma28 stammen aus dem Bereitschaftskalender 2027 – bitte durch echte Namen ersetzen.")
    hdr = ["Personalnummer", "Name", "Team", "Funktion", "Bemerkung", "Gruppe", "Kategorie",
           "Sortierschlüssel", "Prüfung"]
    for i, h in enumerate(hdr):
        put(ws, f"{chr(66 + i)}5", h, bold=True)
    f = ma_formulas()
    for i, p in enumerate(people):
        r = 6 + i
        put(ws, f"B{r}", p.get("Personalnummer") or None, h="center")
        put(ws, f"C{r}", p["Name"])
        put(ws, f"D{r}", p.get("Team") or None, h="center")
        put(ws, f"E{r}", p["Funktion"])
        put(ws, f"F{r}", p.get("Bemerkung") or None)
        for c, key in zip("GHIJ", ["Gruppe", "Kategorie", "Sortierschlüssel", "Prüfung"]):
            put(ws, f"{c}{r}", "=" + f[key], color="404040", h="center" if c != "J" else None)
    last = 5 + len(people)
    add_table(ws, "tblMitarbeiter", f"B5:J{last}", formulas=f)
    for c in "BCDEF":
        for r in range(6, last + 1):
            ws[f"{c}{r}"].fill = fill(INPUT_FILL)
    dv_list(ws, f"D6:D{last + 200}", "lstTeams", prompt="Team wählen (leer bei Disponent/Ressourcen)")
    dv_list(ws, f"E6:E{last + 200}", "lstFunktionen", prompt="Stammfunktion wählen")
    cf(ws, f"J6:J{last + 200}", 'LEFT($J6,1)="⚠"', dxf(bg="FFC7CE", fg="9C0006"))
    cf(ws, f"C6:C{last + 200}", 'AND($C6<>"",COUNTIF($C$6:$C$999,$C6)>1)', dxf(bg="FFC7CE", fg="9C0006"))
    for t in M.TEAMS:
        cf(ws, f"D6:D{last + 200}", f'$D6="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
    ctx.name("lstMA", "tblMitarbeiter[Name]")
    ctx.name("MA_Name", "tblMitarbeiter[Name]")
    ctx.name("MA_Team", "tblMitarbeiter[Team]")
    ctx.name("MA_Funktion", "tblMitarbeiter[Funktion]")
    ctx.name("MA_Gruppe", "tblMitarbeiter[Gruppe]")
    ctx.name("MA_Kat", "tblMitarbeiter[Kategorie]")
    ctx.name("MA_Sort", "tblMitarbeiter[Sortierschlüssel]")
    ctx.name("MA_PNr", "tblMitarbeiter[Personalnummer]")
    # Kennzahlen rechts
    put(ws, "L5", "Übersicht", bold=True, bg=M.HEADER_FILL, border=BOX)
    put(ws, "M5", "Anzahl", bold=True, bg=M.HEADER_FILL, border=BOX, h="center")
    rows = [("Mitarbeiter gesamt", '=COUNTIF(MA_Name,"?*")')]
    rows += [(f"Team {t}", f'=COUNTIFS(MA_Gruppe,"{t}")') for t in M.TEAMS]
    rows += [("Ohne Team / Disponent", '=COUNTIFS(MA_Gruppe,"Ohne Team")'),
             ("Ressourcen (Zusätzliche MA / AZUBIS)", '=COUNTIFS(MA_Gruppe,"Ressourcen")'),
             ("Prüfhinweise (⚠)", '=COUNTIF(tblMitarbeiter[Prüfung],"⚠*")')]
    for i, (lab, fo) in enumerate(rows):
        put(ws, f"L{6 + i}", lab, border=BOX)
        put(ws, f"M{6 + i}", fo, border=BOX, h="center")
    ws.column_dimensions["K"].width = 3
    ws.column_dimensions["L"].width = 36
    ws.column_dimensions["M"].width = 10
    ws.freeze_panes = "C6"
    return ws


def build_matrix(ctx, people):
    ws = ctx.wb[S_FM]
    ws.sheet_properties.tabColor = "4F81BD"
    title(ws, "B1", "Blatt 2 – Team- und Funktionsmatrix")
    note(ws, "B2", "JA/NEIN je Qualifikation. Mehrfachqualifikationen sind möglich und werden von der "
         "Einsatzfähigkeitsprüfung für Vertretungen genutzt. Neue Mitarbeiter: Namen unten anfügen.")
    hdr = ["Name", "Team", "Funktion"] + M.QUALIS + ["Anzahl Qualifikationen", "Mehrfachfunktion", "Prüfung"]
    from openpyxl.utils import get_column_letter as gl
    for i, h in enumerate(hdr):
        c = gl(2 + i)
        rot = 90 if 3 <= i < 3 + len(M.QUALIS) else 0
        put(ws, f"{c}4", h, bold=True, rot=rot, h="center", wrap=rot == 0)
    ws.row_dimensions[4].height = 95
    f = fm_formulas()
    for i, p in enumerate(people):
        r = 5 + i
        put(ws, f"B{r}", p["Name"])
        put(ws, f"C{r}", "=" + f["Team"], h="center", color="404040")
        put(ws, f"D{r}", "=" + f["Funktion"], color="404040")
        for j, qn in enumerate(M.QUALIS):
            c = gl(5 + j)
            put(ws, f"{c}{r}", "JA" if qn == p["Funktion"] else "NEIN", h="center")
        base = 5 + len(M.QUALIS)
        put(ws, f"{gl(base)}{r}", "=" + f["Anzahl Qualifikationen"], h="center", color="404040")
        put(ws, f"{gl(base + 1)}{r}", "=" + f["Mehrfachfunktion"], h="center", color="404040")
        put(ws, f"{gl(base + 2)}{r}", "=" + f["Prüfung"], color="404040")
    last = 4 + len(people)
    add_table(ws, "tblFunktionsmatrix", f"B4:{gl(1 + len(hdr))}{last}", formulas=f)
    q0, q1 = gl(5), gl(4 + len(M.QUALIS))
    dv_list(ws, f"{q0}5:{q1}{last + 200}", "lstJaNein")
    dv_list(ws, f"B5:B{last + 200}", "lstMA", prompt="Mitarbeiter aus dem Stamm wählen")
    cf(ws, f"{q0}5:{q1}{last + 200}", f'{q0}5="JA"', dxf(bg="C6EFCE", fg="006100", bold=True))
    cf(ws, f"{q0}5:{q1}{last + 200}", f'{q0}5="NEIN"', dxf(fg="A6A6A6"))
    pc = gl(1 + len(hdr))
    cf(ws, f"{pc}5:{pc}{last + 200}", f'LEFT(${pc}5,1)="⚠"', dxf(bg="FFC7CE", fg="9C0006"))
    for t in M.TEAMS:
        cf(ws, f"C5:C{last + 200}", f'$C5="{t}"', dxf(bg=M.TEAM_COLOR[t], fg=M.TEAM_FONT[t], bold=True))
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 9
    ws.column_dimensions["D"].width = 14
    for j in range(len(M.QUALIS)):
        ws.column_dimensions[gl(5 + j)].width = 6.5
    ws.column_dimensions[gl(5 + len(M.QUALIS))].width = 13
    ws.column_dimensions[gl(6 + len(M.QUALIS))].width = 13
    ws.column_dimensions[gl(7 + len(M.QUALIS))].width = 34
    ctx.name("FM_Name", "tblFunktionsmatrix[Name]")
    ctx.name("FM_Meister", "tblFunktionsmatrix[Meister]")
    ctx.name("FM_AV", "tblFunktionsmatrix[AV]")
    ctx.name("FM_TVT", "tblFunktionsmatrix[TVT-Fahrer]")
    ctx.name("FM_HW", "tblFunktionsmatrix[Handwerker]")
    ws.freeze_panes = "C5"
    return ws
